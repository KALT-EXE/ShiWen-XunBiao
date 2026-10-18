"""本地 Excel/CSV 读取与智能查找核心。

不依赖 openpyxl 或 Office，直接解析 xlsx/xlsm 的 Open XML 包。
"""

from __future__ import annotations

import csv
import io
import posixpath
import re
import time
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Iterable
import xml.etree.ElementTree as ET


class ExcelReadError(Exception):
    """可以直接展示给用户的表格读取错误。"""


# ---------------------------------------------------------------------------
# 基础数据模型
# ---------------------------------------------------------------------------


@dataclass(slots=True)
class CellRecord:
    row: int
    col: int
    address: str
    value: str
    norm: str


@dataclass(slots=True)
class SheetData:
    name: str
    cells: list[CellRecord]
    rows: dict[int, list[str]]
    headers: dict[int, str] = field(default_factory=dict)
    header_row: int | None = None

    @property
    def data_row_count(self) -> int:
        count = len(self.rows)
        if self.header_row is not None and self.header_row in self.rows:
            count -= 1
        return max(0, count)

    @property
    def max_col(self) -> int:
        return max((len(row) for row in self.rows.values()), default=0)


@dataclass(slots=True)
class WorkbookData:
    path: Path
    sheets: list[SheetData]
    warnings: list[str]
    file_size: int
    mtime: float
    load_ms: float

    @property
    def cell_count(self) -> int:
        return sum(len(sheet.cells) for sheet in self.sheets)

    @property
    def data_row_count(self) -> int:
        return sum(sheet.data_row_count for sheet in self.sheets)

    @property
    def sheet_count(self) -> int:
        return len(self.sheets)


@dataclass(slots=True)
class SearchHit:
    score: float
    match_type: str
    query_term: str
    sheet: str
    row: int
    col: int
    address: str
    cell_value: str
    row_text: str


@dataclass(slots=True)
class SearchResult:
    hits: list[SearchHit]
    total: int
    elapsed_ms: float
    query_terms: list[str]
    truncated: bool
    collect_truncated: bool


# ---------------------------------------------------------------------------
# 文本归一化
# ---------------------------------------------------------------------------

_PUNCT_RE = re.compile(
    r"[\s\u3000\u00a0\-—–_·•,，.。;；:：、|/\\()（）\[\]【】{}<>《》"
    r"\"'“”‘’!！?？~～]+"
)

COMMON_LABELS = (
    "姓名",
    "名字",
    "学号",
    "工号",
    "编号",
    "序号",
    "班级",
    "专业",
    "学院",
    "年级",
    "性别",
    "年龄",
    "电话",
    "手机号",
    "手机",
    "邮箱",
    "邮件",
    "地址",
    "部门",
    "公司",
    "职位",
    "岗位",
    "身份证号",
    "身份证",
    "学籍号",
    "考生号",
    "准考证号",
    "宿舍",
    "房号",
    "床位",
    "成绩",
    "分数",
    "备注",
    "日期",
    "时间",
)

_LABEL_PREFIX_RE = re.compile(
    r"^(?:" + "|".join(sorted(COMMON_LABELS, key=len, reverse=True)) + r")"
    r"\s*[：:\-—]?\s*(.+)$"
)


def normalize_for_search(value: object) -> str:
    """用于比较的归一化：全半角统一、忽略大小写和分隔符。"""

    text = unicodedata.normalize("NFKC", str(value))
    text = text.casefold()
    return _PUNCT_RE.sub("", text)


def tidy_ocr_line(line: str) -> str:
    """清理 Windows OCR 在中文之间插入的多余空格。"""

    line = line.replace("\u00a0", " ").replace("\u3000", " ")
    # 中文与中文、中文与数字/字母之间的空格通常是 OCR 误插。
    line = re.sub(r"(?<=[\u3400-\u9fff])\s+(?=[\u3400-\u9fff])", "", line)
    line = re.sub(r"(?<=[\u3400-\u9fff])\s+(?=[A-Za-z0-9])", "", line)
    line = re.sub(r"(?<=[A-Za-z0-9])\s+(?=[\u3400-\u9fff])", "", line)
    line = re.sub(r"\s*([，,。.；;：:、！!？?])\s*", r"\1", line)
    line = re.sub(r"[ \t]+", " ", line)
    return line.strip()


def split_ocr_lines(raw_text: str) -> list[str]:
    lines: list[str] = []
    for raw in str(raw_text).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = tidy_ocr_line(raw)
        if line:
            lines.append(line)
    return lines


def expand_query_terms(query_text: str) -> list[tuple[str, str]]:
    """把一行检索内容扩展为完整文本、标签后的值等变体。"""

    variants: list[tuple[str, str]] = []
    seen: set[str] = set()

    def add(display: str) -> None:
        display = tidy_ocr_line(str(display)).strip()
        if not display:
            return
        norm = normalize_for_search(display)
        if len(norm) < 1 or norm in seen:
            return
        seen.add(norm)
        variants.append((display, norm))

    raw_lines = [
        line.strip()
        for line in str(query_text).replace("\r", "\n").split("\n")
        if line.strip()
    ]
    if not raw_lines:
        return []

    for raw in raw_lines[:20]:
        line = tidy_ocr_line(raw)
        add(line)

        # “姓名：张三” -> “张三”
        m = re.match(r"^(.{1,16}?)[：:]\s*(.+)$", line)
        if m and m.group(2).strip():
            add(m.group(2))

        # “姓名 张三” / “姓名张三” -> “张三”
        m = _LABEL_PREFIX_RE.match(line)
        if m and m.group(1).strip() and m.group(1).strip() != line:
            add(m.group(1))

        # 三个以上的空格通常表示视觉上的分栏。
        parts = re.split(r"\s{3,}", line)
        if len(parts) > 1:
            for part in parts:
                add(part)
                m = _LABEL_PREFIX_RE.match(part)
                if m:
                    add(m.group(1))

    # 优先使用更具体的值（较短的标签值），但保留整行。
    variants.sort(key=lambda item: (len(item[1]), item[0]))
    return variants[:12]


def suggest_query(lines: Iterable[str]) -> str:
    """为多行 OCR 结果挑选一个最适合直接查找的默认值。"""

    cleaned = [tidy_ocr_line(line) for line in lines if tidy_ocr_line(line)]
    if not cleaned:
        return ""
    for line in cleaned:
        m = re.match(r"^.{1,16}?[：:]\s*(.+)$", line)
        if m and len(m.group(1).strip()) >= 1:
            return line
    for line in cleaned:
        if len(normalize_for_search(line)) >= 2 and not line.endswith(("：", ":")):
            return line
    return cleaned[0]


# ---------------------------------------------------------------------------
# 表格读取
# ---------------------------------------------------------------------------

_DATE_FMT_IDS = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(
    range(50, 59)
)
_BUILTIN_PERCENT_FMT_IDS = {9, 10}


def column_letter(col: int) -> str:
    if col < 1:
        return "?"
    letters: list[str] = []
    while col:
        col, rem = divmod(col - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _first_child(elem: ET.Element, name: str) -> ET.Element | None:
    for child in elem:
        if _local(child.tag) == name:
            return child
    return None


def _all_text(elem: ET.Element) -> str:
    return "".join(node.text or "" for node in elem.iter() if _local(node.tag) == "t")


def _is_date_format_code(code: str) -> bool:
    if not code:
        return False
    # 去掉引号字面量、转义字符和颜色/条件段，避免把普通格式误判为日期。
    clean = re.sub(r'"[^"]*"', "", code)
    clean = re.sub(r"\\.", "", clean)
    clean = re.sub(r"\[[^\]]*\]", "", clean)
    return bool(re.search(r"(?i)(y{2,}|d{1,4}|h{1,2}|s{1,2}|m{1,2})", clean))


def _parse_styles(zf: zipfile.ZipFile) -> tuple[list[bool], list[bool]]:
    date_styles: list[bool] = []
    percent_styles: list[bool] = []
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except KeyError:
        return date_styles, percent_styles
    except ET.ParseError:
        return date_styles, percent_styles

    custom: dict[int, str] = {}
    num_fmts = root.find(".//{*}numFmts")
    if num_fmts is not None:
        for fmt in num_fmts:
            if _local(fmt.tag) != "numFmt":
                continue
            try:
                fmt_id = int(fmt.attrib.get("numFmtId", "-1"))
            except ValueError:
                continue
            custom[fmt_id] = fmt.attrib.get("formatCode", "")

    cell_xfs = root.find(".//{*}cellXfs")
    if cell_xfs is None:
        return date_styles, percent_styles
    for xf in cell_xfs:
        if _local(xf.tag) != "xf":
            continue
        try:
            fmt_id = int(xf.attrib.get("numFmtId", "0"))
        except ValueError:
            fmt_id = 0
        code = custom.get(fmt_id, "")
        is_date = fmt_id in _DATE_FMT_IDS or _is_date_format_code(code)
        is_percent = fmt_id in _BUILTIN_PERCENT_FMT_IDS or "%" in re.sub(
            r'"[^"]*"', "", code
        )
        date_styles.append(is_date)
        percent_styles.append(is_percent)
    return date_styles, percent_styles


def _format_number(raw: str, style_index: int, styles: tuple[list[bool], list[bool]]) -> str:
    try:
        number = float(raw)
    except (TypeError, ValueError):
        return raw.strip()

    date_styles, percent_styles = styles
    is_date = 0 <= style_index < len(date_styles) and date_styles[style_index]
    is_percent = 0 <= style_index < len(percent_styles) and percent_styles[style_index]

    if is_date:
        return _format_excel_serial(number)

    if is_percent:
        pct = number * 100.0
        text = f"{pct:.4f}".rstrip("0").rstrip(".")
        return f"{text}%"

    if raw.strip().lstrip("+-").isdigit():
        return raw.strip()
    if number.is_integer() and abs(number) < 1e15:
        return str(int(number))
    if abs(number) >= 1e15 or (0 < abs(number) < 1e-9):
        return raw.strip()
    return f"{number:.15g}"


def _format_excel_serial(serial: float) -> str:
    # Excel 把 1900 年当闰年，60 之前和之后需要不同基准。
    base = datetime(1899, 12, 31) if serial < 60 else datetime(1899, 12, 30)
    try:
        value = base + timedelta(days=float(serial))
    except (OverflowError, ValueError):
        return str(serial)
    value = value.replace(microsecond=0)
    if 0 <= serial < 1:
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return "00:00"
        return value.strftime("%H:%M:%S" if value.second else "%H:%M")
    if value.hour == 0 and value.minute == 0 and value.second == 0:
        return value.strftime("%Y-%m-%d")
    if value.second:
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return value.strftime("%Y-%m-%d %H:%M")


def _parse_iso_date(raw: str) -> str:
    text = raw.strip()
    for candidate in (
        text,
        text.replace("Z", "+00:00"),
    ):
        try:
            dt = datetime.fromisoformat(candidate)
            if dt.tzinfo is not None:
                dt = dt.astimezone().replace(tzinfo=None)
            if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
                return dt.strftime("%Y-%m-%d")
            return dt.strftime("%Y-%m-%d %H:%M:%S" if dt.second else "%Y-%m-%d %H:%M")
        except ValueError:
            continue
    return text


def _parse_cell_value(
    cell: ET.Element,
    shared_strings: list[str],
    styles: tuple[list[bool], list[bool]],
) -> str:
    cell_type = cell.attrib.get("t", "n")
    style_index = 0
    try:
        style_index = int(cell.attrib.get("s", "0"))
    except ValueError:
        pass

    v = _first_child(cell, "v")
    raw = (v.text or "") if v is not None else ""

    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except (ValueError, IndexError):
            return raw
    if cell_type == "inlineStr":
        is_node = _first_child(cell, "is")
        return _all_text(is_node) if is_node is not None else raw
    if cell_type == "str":
        return raw
    if cell_type == "b":
        return "TRUE" if raw.strip() in ("1", "true", "TRUE") else "FALSE"
    if cell_type == "e":
        return raw
    if cell_type == "d":
        return _parse_iso_date(raw)
    if not raw:
        return ""
    return _format_number(raw, style_index, styles)


def _parse_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        with zf.open("xl/sharedStrings.xml") as stream:
            result: list[str] = []
            for _, elem in ET.iterparse(stream, events=("end",)):
                if _local(elem.tag) == "si":
                    result.append(_all_text(elem))
                    elem.clear()
            return result
    except KeyError:
        return []
    except ET.ParseError as exc:
        raise ExcelReadError(f"共享文本表损坏：{exc}") from exc


def _parse_sheet_xml(
    stream,
    shared_strings: list[str],
    styles: tuple[list[bool], list[bool]],
) -> dict[int, list[str]]:
    rows: dict[int, list[str]] = {}
    row_index = 0
    for _, elem in ET.iterparse(stream, events=("end",)):
        if _local(elem.tag) != "row":
            continue
        try:
            row_index = int(elem.attrib.get("r", row_index + 1))
        except ValueError:
            row_index += 1
        values: dict[int, str] = {}
        last_col = 0
        for cell in elem:
            if _local(cell.tag) != "c":
                continue
            ref = cell.attrib.get("r", "")
            m = re.match(r"^([A-Za-z]+)(\d+)$", ref)
            if m:
                col = 0
                for ch in m.group(1).upper():
                    col = col * 26 + (ord(ch) - 64)
                cell_row = int(m.group(2))
                if cell_row != row_index:
                    # 极少数生成器会漏写 row 的 r 属性；以单元格坐标为准。
                    pass
            else:
                col = last_col + 1
            if col < 1:
                col = 1
            last_col = col
            try:
                text = _parse_cell_value(cell, shared_strings, styles)
            except Exception:
                continue
            if text is None:
                continue
            text = str(text)
            if text.strip():
                values[col] = text.strip()
        if values:
            max_col = max(values)
            rows[row_index] = [values.get(col, "") for col in range(1, max_col + 1)]
        elem.clear()
    return rows


def _header_from_rows(rows: dict[int, list[str]]) -> tuple[int | None, dict[int, str]]:
    if not rows:
        return None, {}
    first_row = min(rows)
    values = rows[first_row]
    non_empty = [(idx + 1, value) for idx, value in enumerate(values) if value.strip()]
    if len(non_empty) < 2:
        return None, {}

    keywords = {
        "姓名",
        "名字",
        "学号",
        "工号",
        "编号",
        "序号",
        "班级",
        "专业",
        "学院",
        "部门",
        "名称",
        "日期",
        "时间",
        "电话",
        "手机",
        "手机号",
        "邮箱",
        "成绩",
        "分数",
        "性别",
        "年龄",
        "地址",
        "备注",
        "公司",
        "职位",
        "岗位",
        "课程",
        "科目",
        "数量",
        "金额",
        "状态",
    }
    has_keyword = any(
        any(word in value for word in keywords) for _, value in non_empty
    )
    text_like = sum(
        1
        for _, value in non_empty
        if not re.fullmatch(r"[-+]?\d+(?:\.\d+)?%?", value.strip())
    )
    looks_like_header = has_keyword or (
        len(non_empty) >= 3 and text_like >= len(non_empty) * 0.6
    )
    if not looks_like_header:
        return None, {}
    return first_row, {col: value for col, value in non_empty}


def _build_sheet(name: str, rows: dict[int, list[str]]) -> SheetData:
    header_row, headers = _header_from_rows(rows)
    cells: list[CellRecord] = []
    for row_num in sorted(rows):
        values = rows[row_num]
        for col, value in enumerate(values, 1):
            if not value:
                continue
            cells.append(
                CellRecord(
                    row=row_num,
                    col=col,
                    address=f"{column_letter(col)}{row_num}",
                    value=value,
                    norm=normalize_for_search(value),
                )
            )
    return SheetData(
        name=name,
        cells=cells,
        rows=rows,
        headers=headers,
        header_row=header_row,
    )


def _load_xlsx(
    path: Path, progress: Callable[[str], None] | None
) -> tuple[list[SheetData], list[str]]:
    warnings: list[str] = []
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise ExcelReadError(
            "无法打开这个 Excel 文件。请确认它是有效的 .xlsx/.xlsm 文件，"
            "而不是被重命名的 .xls 文件。"
        ) from exc
    except PermissionError as exc:
        raise ExcelReadError("文件正被占用或没有读取权限，请关闭 Excel 后重试。") from exc

    with zf:
        try:
            workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
        except KeyError as exc:
            raise ExcelReadError("这个文件缺少 Excel 工作簿结构，可能已损坏。") from exc
        except ET.ParseError as exc:
            raise ExcelReadError(f"工作簿结构无法解析：{exc}") from exc

        rels: dict[str, str] = {}
        try:
            rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            for rel in rel_root:
                if _local(rel.tag) != "Relationship":
                    continue
                if rel.attrib.get("TargetMode") == "External":
                    continue
                rels[rel.attrib.get("Id", "")] = rel.attrib.get("Target", "")
        except (KeyError, ET.ParseError):
            warnings.append("未能读取工作簿关系文件，部分工作表可能被跳过。")

        shared_strings = _parse_shared_strings(zf)
        styles = _parse_styles(zf)

        sheets: list[SheetData] = []
        sheets_node = workbook_root.find(".//{*}sheets")
        if sheets_node is None:
            raise ExcelReadError("工作簿中没有找到工作表。")
        sheet_entries = [
            elem for elem in sheets_node if _local(elem.tag) == "sheet"
        ]
        for sheet_elem in sheet_entries:
            name = sheet_elem.attrib.get("name", "未命名工作表")
            state = sheet_elem.attrib.get("state", "visible")
            if state in ("hidden", "veryHidden"):
                continue
            rel_id = next(
                (value for key, value in sheet_elem.attrib.items() if _local(key) == "id"),
                "",
            )
            target = rels.get(rel_id, "")
            if not target:
                warnings.append(f"工作表“{name}”缺少文件路径，已跳过。")
                continue
            if target.startswith("/"):
                sheet_path = target.lstrip("/")
            else:
                sheet_path = posixpath.normpath(posixpath.join("xl", target))
            if progress:
                progress(f"正在读取工作表：{name}")
            try:
                with zf.open(sheet_path) as stream:
                    rows = _parse_sheet_xml(stream, shared_strings, styles)
            except KeyError:
                warnings.append(f"工作表“{name}”的数据文件不存在，已跳过。")
                continue
            except (ET.ParseError, OSError) as exc:
                warnings.append(f"工作表“{name}”读取失败：{exc}")
                continue
            sheets.append(_build_sheet(name, rows))

        if not sheets:
            raise ExcelReadError("没有读取到任何可见工作表。")
        return sheets, warnings


def _decode_csv(raw: bytes) -> str:
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    for encoding in ("utf-8-sig", "gb18030", "utf-8"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _load_csv(path: Path) -> tuple[list[SheetData], list[str]]:
    try:
        raw = path.read_bytes()
    except PermissionError as exc:
        raise ExcelReadError("CSV 文件正被占用或没有读取权限。") from exc
    if not raw:
        raise ExcelReadError("CSV 文件是空的。")
    text = _decode_csv(raw)
    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
    except csv.Error:
        dialect = csv.excel
    rows: dict[int, list[str]] = {}
    reader = csv.reader(io.StringIO(text, newline=""), dialect)
    try:
        for row_index, row in enumerate(reader, 1):
            values = [str(value).strip() for value in row]
            while values and not values[-1]:
                values.pop()
            if any(values):
                rows[row_index] = values
    except csv.Error as exc:
        raise ExcelReadError(f"CSV 内容格式有误：{exc}") from exc
    if not rows:
        raise ExcelReadError("CSV 文件中没有可读取的数据。")
    return [_build_sheet(path.stem, rows)], []


def load_workbook(
    path: str | Path, progress: Callable[[str], None] | None = None
) -> WorkbookData:
    """读取 Excel/CSV，返回可搜索的内存索引。"""

    file_path = Path(path).expanduser()
    if not file_path.exists():
        raise ExcelReadError(f"找不到文件：{file_path}")
    if not file_path.is_file():
        raise ExcelReadError("请选择一个文件，而不是文件夹。")
    suffix = file_path.suffix.lower()
    if suffix not in (".xlsx", ".xlsm", ".csv"):
        if suffix == ".xls":
            raise ExcelReadError(
                "暂不支持旧版 .xls。请在 Excel 中“另存为” .xlsx 后再试。"
            )
        if suffix == ".xlsb":
            raise ExcelReadError("暂不支持 .xlsb。请在 Excel 中另存为 .xlsx 后再试。")
        raise ExcelReadError("仅支持 .xlsx、.xlsm 和 .csv 文件。")

    started = time.perf_counter()
    try:
        stat = file_path.stat()
    except OSError as exc:
        raise ExcelReadError(f"无法读取文件信息：{exc}") from exc

    if progress:
        progress("正在打开表格…")
    if suffix == ".csv":
        sheets, warnings = _load_csv(file_path)
    else:
        sheets, warnings = _load_xlsx(file_path, progress)

    elapsed = (time.perf_counter() - started) * 1000
    return WorkbookData(
        path=file_path,
        sheets=sheets,
        warnings=warnings,
        file_size=stat.st_size,
        mtime=stat.st_mtime,
        load_ms=elapsed,
    )


# ---------------------------------------------------------------------------
# 检索
# ---------------------------------------------------------------------------

_MODE_NAMES = {
    "smart": "智能匹配",
    "exact": "精确匹配",
    "contains": "包含匹配",
}


def mode_label(mode: str) -> str:
    return _MODE_NAMES.get(mode, "智能匹配")


def _bounded_levenshtein(a: str, b: str, max_distance: int) -> int:
    if abs(len(a) - len(b)) > max_distance:
        return max_distance + 1
    previous = list(range(len(b) + 1))
    for i, char_a in enumerate(a, 1):
        current = [i] + [0] * len(b)
        row_min = current[0]
        for j, char_b in enumerate(b, 1):
            cost = 0 if char_a == char_b else 1
            current[j] = min(
                previous[j] + 1,
                current[j - 1] + 1,
                previous[j - 1] + cost,
            )
            if current[j] < row_min:
                row_min = current[j]
        if row_min > max_distance:
            return max_distance + 1
        previous = current
    return previous[-1]


def _score_match(
    cell_norm: str, query_norm: str, mode: str
) -> tuple[float, str] | None:
    if not cell_norm or not query_norm:
        return None
    if cell_norm == query_norm:
        return 1.0, "完全匹配"
    if mode == "exact":
        return None

    if cell_norm.startswith(query_norm):
        return 0.96, "开头匹配"
    if query_norm in cell_norm:
        return 0.90, "包含匹配"
    if mode != "smart":
        return None

    # 近似匹配：只接受长度接近、且首/尾至少一端相同的文本，减少误报。
    q_len = len(query_norm)
    c_len = len(cell_norm)
    if q_len < 3 or q_len > 40 or c_len < 3:
        return None
    digits = sum(1 for ch in query_norm if ch.isdigit())
    if digits * 2 >= q_len:
        return None
    max_distance = 1 if q_len <= 8 else 2
    if abs(q_len - c_len) > max_distance:
        return None
    if cell_norm[0] != query_norm[0] and cell_norm[-1] != query_norm[-1]:
        return None
    distance = _bounded_levenshtein(query_norm, cell_norm, max_distance)
    if distance > max_distance:
        return None
    ratio = 1.0 - distance / max(q_len, c_len)
    threshold = 0.66 if q_len <= 4 else 0.72
    if ratio < threshold:
        return None
    return 0.60 + 0.20 * ratio, "近似匹配"


def _row_context(sheet: SheetData, row_num: int) -> str:
    """返回整行全部非空字段，不做截断，供结果详情直接输出。"""

    values = sheet.rows.get(row_num, [])
    if not values:
        return ""
    available = [col for col in range(1, len(values) + 1) if values[col - 1]]
    if not available:
        return ""
    parts: list[str] = []
    for col in available:
        header = sheet.headers.get(col)
        label = header if header else column_letter(col)
        parts.append(f"{label}：{values[col - 1]}")
    return "  |  ".join(parts)


def search_workbook(
    workbook: WorkbookData,
    query_text: str,
    mode: str = "smart",
    display_limit: int = 300,
    collect_limit: int = 5000,
) -> SearchResult:
    """在全部可见工作表中查找，返回按相关度排序的结果。"""

    if mode not in _MODE_NAMES:
        mode = "smart"
    variants = expand_query_terms(query_text)
    if not variants:
        raise ValueError("请输入要查找的内容。")
    started = time.perf_counter()
    hits: list[SearchHit] = []
    collect_truncated = False

    for sheet_index, sheet in enumerate(workbook.sheets):
        for cell in sheet.cells:
            if sheet.header_row is not None and cell.row == sheet.header_row:
                continue
            if not cell.norm:
                continue
            best: tuple[float, str, str] | None = None
            for display, norm in variants:
                match = _score_match(cell.norm, norm, mode)
                if match is None:
                    continue
                score, kind = match
                if best is None or score > best[0]:
                    best = (score, kind, display)
            if best is None:
                continue
            score, kind, term = best
            hits.append(
                SearchHit(
                    score=score,
                    match_type=kind,
                    query_term=term,
                    sheet=sheet.name,
                    row=cell.row,
                    col=cell.col,
                    address=cell.address,
                    cell_value=cell.value,
                    row_text="",
                )
            )
            if len(hits) >= collect_limit:
                collect_truncated = True
                break
        if collect_truncated:
            break

    hits.sort(key=lambda hit: (-hit.score, hit.sheet, hit.row, hit.col))
    total = len(hits)
    for hit in hits[:display_limit]:
        sheet = next((s for s in workbook.sheets if s.name == hit.sheet), None)
        if sheet is not None:
            hit.row_text = _row_context(sheet, hit.row)
    display_hits = hits[:display_limit]
    elapsed = (time.perf_counter() - started) * 1000
    return SearchResult(
        hits=display_hits,
        total=total,
        elapsed_ms=elapsed,
        query_terms=[display for display, _ in variants],
        truncated=total > len(display_hits),
        collect_truncated=collect_truncated,
    )


def format_file_size(size: int) -> str:
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.1f} MB"
