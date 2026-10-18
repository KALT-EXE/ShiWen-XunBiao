"""Headless tests for the workbook parser and search engine."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from excel_index import (  # noqa: E402
    ExcelReadError,
    load_workbook,
    normalize_for_search,
    search_workbook,
)


class WorkbookSearchTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.workbook = load_workbook(ROOT / "示例数据.xlsx")

    def test_sample_workbook_shape(self) -> None:
        self.assertEqual(self.workbook.sheet_count, 2)
        self.assertEqual(self.workbook.data_row_count, 40)
        self.assertEqual(self.workbook.sheets[0].name, "学生名单")
        self.assertEqual(self.workbook.sheets[0].headers[2], "姓名")
        self.assertEqual(self.workbook.sheets[0].header_row, 1)

    def test_dates_are_formatted(self) -> None:
        dates = [
            cell.value
            for cell in self.workbook.sheets[0].cells
            if cell.address == "G2"
        ]
        self.assertEqual(dates, ["2026-09-01"])

    def test_label_query_expands_to_name(self) -> None:
        result = search_workbook(self.workbook, "姓名：张三")
        self.assertGreaterEqual(result.total, 3)
        self.assertEqual(result.hits[0].cell_value, "张三")

    def test_full_row_contains_every_populated_field(self) -> None:
        result = search_workbook(self.workbook, "张三", mode="exact")
        self.assertGreaterEqual(result.total, 1)
        row = result.hits[0].row_text
        for field in (
            "学号：",
            "姓名：",
            "班级：",
            "专业：",
            "宿舍：",
            "手机号：",
            "报到日期：",
            "状态：",
            "备注：",
        ):
            self.assertIn(field, row)
        self.assertNotIn("…", row)

    def test_exact_mode_only_returns_exact_cells(self) -> None:
        result = search_workbook(self.workbook, "张三", mode="exact")
        self.assertTrue(result.hits)
        self.assertTrue(all(hit.cell_value == "张三" for hit in result.hits))

    def test_contains_mode(self) -> None:
        result = search_workbook(self.workbook, "园艺2026", mode="contains")
        self.assertGreaterEqual(result.total, 6)

    def test_search_ignores_empty_query(self) -> None:
        with self.assertRaises(ValueError):
            search_workbook(self.workbook, "   ")


class CsvTests(unittest.TestCase):
    def test_csv_is_read_and_searched(self) -> None:
        workbook = load_workbook(ROOT / "tests" / "fixtures" / "sample.csv")
        self.assertEqual(workbook.sheet_count, 1)
        self.assertEqual(workbook.data_row_count, 2)
        self.assertEqual(workbook.sheets[0].headers[1], "姓名")

        result = search_workbook(workbook, "张三")
        self.assertEqual(result.total, 1)
        self.assertEqual(result.hits[0].address, "A2")
        self.assertIn("学号：2025010101", result.hits[0].row_text)

    def test_unsupported_old_format(self) -> None:
        with self.assertRaises(ExcelReadError):
            load_workbook(ROOT / "tests" / "fixtures" / "legacy.xls")


class TextHelperTests(unittest.TestCase):
    def test_normalize_ignores_spaces_and_full_width(self) -> None:
        self.assertEqual(
            normalize_for_search("园艺 ２０２６级-1班"),
            "园艺2026级1班",
        )

    def test_normalize_is_case_insensitive(self) -> None:
        self.assertEqual(normalize_for_search("ABC-123"), "abc123")


if __name__ == "__main__":
    unittest.main()
