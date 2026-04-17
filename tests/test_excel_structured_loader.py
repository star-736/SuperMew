from __future__ import annotations

import unittest
import uuid
from pathlib import Path

from openpyxl import Workbook

from backend.document_loader import DocumentLoader


class ExcelStructuredLoaderTests(unittest.TestCase):
    def setUp(self) -> None:
        self.loader = DocumentLoader()

    def _build_workbook(self, path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "商品表"
        ws.append(["产品", "价格", "价格", None, "备注"])
        ws.append(["苹果", 5, 6, "第一行\n第二行", "## markdown-like"])
        ws.append(["香蕉", None, 7, "", ""])
        ws.append([None, None, None, None, None])

        ws2 = wb.create_sheet("季度汇总")
        ws2.append(["区域", "Q1", "Q2"])
        ws2.append(["华东", 120, 150])
        ws2.append(["华南", 98, "=B3+12"])
        wb.save(path)

    def test_excel_loader_preserves_structure_for_rows_and_sheets(self) -> None:
        artifacts_dir = Path(__file__).resolve().parent / ".artifacts"
        artifacts_dir.mkdir(exist_ok=True)
        file_path = artifacts_dir / f"sample-{uuid.uuid4().hex}.xlsx"
        try:
            self._build_workbook(file_path)

            workbook = self.loader.load_excel_workbook(str(file_path), "sample.xlsx")

            self.assertEqual(len(workbook["sheets"]), 2)
            self.assertEqual(len(workbook["sheet_docs"]), 2)
            self.assertEqual(len(workbook["rows"]), 4)
            self.assertEqual(len(workbook["row_docs"]), 4)

            first_row = workbook["rows"][0]
            self.assertEqual(first_row["sheet_name"], "商品表")
            self.assertIn("column_4", first_row["headers"])
            self.assertIn("价格_2", first_row["headers"])
            self.assertEqual(first_row["row_obj"]["产品"], "苹果")
            self.assertEqual(first_row["row_obj"]["column_4"], "第一行\n第二行")
            self.assertEqual(first_row["row_obj"]["备注"], "## markdown-like")

            first_row_doc = workbook["row_docs"][0]
            self.assertEqual(first_row_doc["record_type"], "excel_row")
            self.assertIn("Sheet: 商品表", first_row_doc["text"])
            self.assertIn("备注=## markdown-like", first_row_doc["text"])

            first_sheet = workbook["sheets"][0]
            self.assertEqual(first_sheet["sheet_name"], "商品表")
            self.assertIn("<table>", first_sheet["sheet_html"])
            self.assertIn("## markdown-like", first_sheet["sheet_html"])
            self.assertIn("示例行1:", first_sheet["sheet_text"])

            second_sheet_row = workbook["rows"][-1]
            self.assertEqual(second_sheet_row["sheet_name"], "季度汇总")
            self.assertEqual(second_sheet_row["row_obj"]["Q2"], "=B3+12")
        finally:
            if file_path.exists():
                file_path.unlink()


if __name__ == "__main__":
    unittest.main()
