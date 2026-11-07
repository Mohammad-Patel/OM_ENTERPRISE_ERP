
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from app.services.excel_service import (
    SCHEMAS,
    append_row,
    get_sheet_rows,
    initialize_workbook,
    workbook_status,
)


class ExcelServiceTests(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)

        self.workbook_path = (
            Path(self.temp_dir.name) / "test_workbook.xlsx"
        )

        initialize_workbook(self.workbook_path)

    def test_workbook_has_all_required_sheets(self):
        status = workbook_status(self.workbook_path)

        self.assertTrue(status["exists"])
        self.assertEqual(
            set(status["sheets"]),
            set(SCHEMAS.keys()),
        )

    def test_required_headers_exist(self):
        wb = load_workbook(
            self.workbook_path,
            read_only=True,
            data_only=True,
        )
        self.addCleanup(wb.close)

        for sheet_name, expected_headers in SCHEMAS.items():
            ws = wb[sheet_name]
            actual_headers = [
                ws.cell(row=1, column=column).value
                for column in range(1, ws.max_column + 1)
            ]

            for header in expected_headers:
                self.assertIn(
                    header,
                    actual_headers,
                    f"Missing {header} in {sheet_name}",
                )

    def test_append_and_read_row(self):
        append_row(
            "Customers",
            {
                "Party_ID": "TEST-P001",
                "Party_Name": "Test Party",
                "Party_Type": "WITHOUT_GST",
                "GSTIN": None,
                "Is_Active": True,
            },
            self.workbook_path,
        )

        rows = get_sheet_rows(
            "Customers",
            self.workbook_path,
        )

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Party_ID"], "TEST-P001")
        self.assertEqual(rows[0]["Party_Name"], "Test Party")

    def test_unknown_sheet_is_rejected(self):
        with self.assertRaises(Exception):
            get_sheet_rows(
                "Not_A_Real_Sheet",
                self.workbook_path,
            )

    def test_unknown_fields_are_rejected(self):
        with self.assertRaises(Exception):
            append_row(
                "Customers",
                {"Not_A_Real_Field": "value"},
                self.workbook_path,
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
