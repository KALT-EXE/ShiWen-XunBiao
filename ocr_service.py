"""调用 Windows 10/11 自带 OCR 的轻量封装。"""

from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image, ImageEnhance, ImageFilter, ImageOps

from excel_index import split_ocr_lines


class OcrError(Exception):
    """可以直接展示给用户的 OCR 错误。"""


@dataclass(slots=True)
class OcrResult:
    lines: list[str]
    text: str
    raw_text: str
    language: str
    available_languages: list[str]
    elapsed_ms: float
    variant: str = "normal"
    warnings: list[str] = field(default_factory=list)

    @property
    def char_count(self) -> int:
        return len(self.text.replace("\n", ""))


def _writable_temp_root() -> Path:
    candidates: list[Path] = []
    local = os.environ.get("LOCALAPPDATA")
    if local:
        candidates.append(Path(local) / "ShiWenXunBiao" / "temp")
    candidates.append(Path(__file__).resolve().parent / ".cache")
    candidates.append(Path(tempfile.gettempdir()) / "ShiWenXunBiao")
    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            probe = candidate / ".write_test"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink(missing_ok=True)
            return candidate
        except OSError:
            continue
    raise OcrError("无法创建临时文件目录，请检查磁盘权限。")


class OcrService:
    def __init__(self, script_path: str | Path) -> None:
        self.script_path = Path(script_path)
        self.powershell = (
            Path(os.environ.get("SystemRoot", r"C:\Windows"))
            / "System32"
            / "WindowsPowerShell"
            / "v1.0"
            / "powershell.exe"
        )
        self.last_language = ""

    # ------------------------------------------------------------------
    # 公共入口
    # ------------------------------------------------------------------

    def recognize(self, image: Image.Image) -> OcrResult:
        if image.width < 4 or image.height < 4:
            raise OcrError("截图区域太小，无法识别。")
        if not self.script_path.exists():
            raise OcrError(f"缺少 OCR 组件：{self.script_path.name}")
        if not self.powershell.exists():
            raise OcrError("找不到 Windows PowerShell，无法调用系统 OCR。")

        started = time.perf_counter()
        prepared = self._prepare(image, invert=False)
        first = self._recognize_prepared(prepared, variant="normal")
        best = first

        # 深色背景或第一次没有结果时，再试一次反色，兼容深色模式截图。
        if not first.lines or self._average_luma(image) < 105:
            inverted = self._prepare(image, invert=True)
            second = self._recognize_prepared(inverted, variant="inverted")
            if self._result_quality(second) > self._result_quality(best):
                best = second

        best.elapsed_ms = (time.perf_counter() - started) * 1000
        if not best.lines:
            raise OcrError(
                "没有识别到文字。请框选更大、更清晰、对比度更高的文字区域后重试。"
            )
        self.last_language = best.language
        return best

    # ------------------------------------------------------------------
    # 图像预处理
    # ------------------------------------------------------------------

    @staticmethod
    def _average_luma(image: Image.Image) -> float:
        small = image.convert("L")
        if small.width > 160 or small.height > 160:
            small.thumbnail((160, 160))
        data = list(small.getdata())
        return sum(data) / max(1, len(data))

    def _prepare(self, image: Image.Image, invert: bool) -> Image.Image:
        if image.mode in ("RGBA", "LA"):
            background = Image.new("RGB", image.size, "white")
            alpha = image.getchannel("A")
            background.paste(image.convert("RGB"), mask=alpha)
            work = background
        else:
            work = image.convert("RGB")

        max_dim = max(work.size)
        if max_dim <= 1500:
            scale = 2.0
        elif max_dim <= 2600:
            scale = 1.5
        else:
            scale = 1.0
        if min(work.size) < 70:
            scale = min(3.0, scale * 1.5)
        scale = min(scale, 4096 / max(1, max_dim))
        if scale > 1.02:
            work = work.resize(
                (max(1, int(work.width * scale)), max(1, int(work.height * scale))),
                Image.Resampling.LANCZOS,
            )

        if invert:
            work = ImageOps.invert(work)
        work = ImageEnhance.Contrast(work).enhance(1.06)
        work = work.filter(
            ImageFilter.UnsharpMask(radius=1.0, percent=55, threshold=4)
        )

        padding = 18
        padded = Image.new(
            "RGB", (work.width + padding * 2, work.height + padding * 2), "white"
        )
        padded.paste(work, (padding, padding))
        if padded.width > 4096 or padded.height > 4096:
            padded.thumbnail((4096, 4096), Image.Resampling.LANCZOS)
        return padded

    # ------------------------------------------------------------------
    # PowerShell 调用
    # ------------------------------------------------------------------

    def _recognize_prepared(self, image: Image.Image, variant: str) -> OcrResult:
        temp_root = _writable_temp_root()
        token = uuid.uuid4().hex
        image_path = temp_root / f"capture_{token}.png"
        output_path = temp_root / f"result_{token}.json"
        try:
            image.save(image_path, "PNG", optimize=True)
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            startupinfo = None
            if os.name == "nt":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
            command = [
                str(self.powershell),
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(self.script_path),
                "-ImagePath",
                str(image_path),
                "-OutputPath",
                str(output_path),
            ]
            try:
                completed = subprocess.run(
                    command,
                    capture_output=True,
                    timeout=45,
                    creationflags=creationflags,
                    startupinfo=startupinfo,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                )
            except subprocess.TimeoutExpired as exc:
                raise OcrError("OCR 识别超时，请缩小框选区域后重试。") from exc
            except OSError as exc:
                raise OcrError(f"无法启动 Windows OCR：{exc}") from exc

            data = None
            if output_path.exists():
                try:
                    data = json.loads(output_path.read_text(encoding="utf-8-sig"))
                except (OSError, json.JSONDecodeError):
                    data = None
            if not isinstance(data, dict):
                detail = (completed.stderr or "").strip()
                if len(detail) > 400:
                    detail = detail[-400:]
                raise OcrError(
                    "OCR 组件没有返回有效结果。"
                    + (f"\n{detail}" if detail else "请确认系统已安装中文 OCR。")
                )
            if not data.get("ok"):
                raise OcrError(str(data.get("error") or "OCR 识别失败。"))

            lines = [
                str(line).strip()
                for line in (data.get("lines") or [])
                if str(line).strip()
            ]
            raw_text = str(data.get("text") or "")
            if not lines and raw_text:
                lines = split_ocr_lines(raw_text)
            cleaned = split_ocr_lines("\n".join(lines))
            return OcrResult(
                lines=cleaned,
                text="\n".join(cleaned),
                raw_text=raw_text,
                language=str(data.get("language") or ""),
                available_languages=[
                    str(item) for item in (data.get("available_languages") or [])
                ],
                elapsed_ms=0.0,
                variant=variant,
            )
        finally:
            for path in (image_path, output_path):
                try:
                    path.unlink(missing_ok=True)
                except OSError:
                    pass

    @staticmethod
    def _result_quality(result: OcrResult) -> tuple[int, int]:
        return (len(result.text.strip()), result.char_count * 2 + len(result.lines))
