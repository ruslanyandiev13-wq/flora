"""Раскладка и оформление Invoice total по правкам закупщика 02.10.2026.

Проверки используют синтетические данные, без клиентских документов.
"""
import datetime
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import xlrd  # noqa: E402

from import_xls_writer import build_combined_factura_xls  # noqa: E402


class InvoiceTotalOctober2LayoutTest(unittest.TestCase):
    def _workbook(self, item_count=6):
        boxes = []
        for index in range(item_count):
            price = 0.91 if index % 2 == 0 else 0.81
            boxes.append({
                "farm": "TEST FARM",
                "product": "HYDRANGEAS",
                "box_size": 0.25,
                "items": [{"variety": "PINK", "grade_text": "PR", "stems": 30,
                           "price": price, "total": round(30 * price, 2)}],
            })
        info = {
            "boxes": boxes,
            "total_stems": item_count * 30,
            "total_fob": round(sum(box["items"][0]["total"] for box in boxes), 2),
            "total_full_boxes": item_count * 0.25,
            "invoice_date": "2026-09-30",
            "destination": "Russia",
            "forwarder": "TEST FORWARDER",
            "airline": "TEST AIRLINE",
            "hawb_number": "TEST-HAWB",
            "awb": {"total_awb": 93.52, "chargeable_weight": 28, "gross_weight": 28},
        }
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "invoice.xls")
            build_combined_factura_xls(
                path, {"TEST": info}, {"awb_no": "000-0000 0000"},
                consignee="TEST MOS", delivery_date=datetime.date(2026, 10, 4),
            )
            return xlrd.open_workbook(path, formatting_info=True)

    @staticmethod
    def _value(sheet, row, column):
        # После переноса ставки файл может заканчиваться раньше V/W.
        if row >= sheet.nrows or column >= sheet.ncols:
            return ""
        return sheet.cell_value(row, column)

    def _assert_style(self, workbook, sheet, row, column, *, bold, fill=None):
        xf = workbook.xf_list[sheet.cell_xf_index(row, column)]
        self.assertEqual(bool(workbook.font_list[xf.font_index].bold), bold)
        if fill is None:
            self.assertEqual(xf.background.fill_pattern, 0)
        else:
            self.assertEqual(xf.background.fill_pattern, 1)
            self.assertEqual(xf.background.pattern_colour_index, fill)

    def test_header_moves_left_and_rate_moves_above_table(self):
        workbook = self._workbook()
        sheet = workbook.sheet_by_name("factura")
        for row, label in {
            1: "DATE", 2: "CARGO AGENCY", 4: "AIRLINE", 5: "M.A.W.B",
            6: "H.A.W.B", 7: "TOTAL FULL BOXES",
        }.items():
            with self.subTest(row=row + 1):
                self.assertEqual(sheet.cell_value(row, 10), label)  # K
                self.assertEqual(sheet.cell_value(row, 11), "")     # L
        self.assertEqual(
            sorted(column for column, info in sheet.colinfo_map.items() if info.hidden),
            [4, 8, 9, 11, 13],  # E, I, J, L, N
        )
        self.assertAlmostEqual(sheet.cell_value(3, 16), 3.34)  # Q4
        self.assertEqual(sheet.cell_value(3, 17), "ставка за кг")  # R4
        self.assertEqual(self._value(sheet, 9, 21), "")  # V10
        self.assertEqual(self._value(sheet, 9, 22), "")  # W10
        self.assertEqual(sheet.cell_value(5, 16), "000-0000 0000")  # Q6

    def _assert_totals(self, workbook, item_count):
        sheet = workbook.sheet_by_name("factura")
        subtotal = 10 + item_count
        stems_row = subtotal + 2
        fob_row = subtotal + 3
        chargeable_row = subtotal + 6
        awb_row = subtotal + 10
        usd_row = subtotal + 11
        for row, label in [
            (stems_row, "TOTAL STEMS"), (fob_row, "TOTAL FLOWERS FOB USD"),
            (chargeable_row, "TOTAL CHARGEABLE WEIGHT(Kg)"),
            (awb_row, "TOTAL AWB"), (usd_row, "TOTAL USD"),
        ]:
            self.assertIn(label, sheet.row_values(row))

        self.assertEqual(sheet.cell_value(subtotal, 12), item_count * 30)
        self._assert_style(workbook, sheet, subtotal, 12, bold=False)
        for column in (5, 6, 15):  # F/G/P сохраняют жирный итог.
            self._assert_style(workbook, sheet, subtotal, column, bold=True)

        for row, value, fill, bold in [
            (stems_row, item_count * 30, 13, True),
            (chargeable_row, 28, 50, False),
        ]:
            self.assertIn((row, row + 1, 13, 16), sheet.merged_cells)  # N:P
            self.assertEqual(sheet.cell_value(row, 13), value)
            for column in (13, 14, 15):
                self._assert_style(workbook, sheet, row, column, bold=bold, fill=fill)

        expected_fob = round(sum(27.3 if index % 2 == 0 else 24.3
                                 for index in range(item_count)), 2)
        self.assertAlmostEqual(sheet.cell_value(fob_row, 19), expected_fob)
        self._assert_style(workbook, sheet, fob_row, 19, bold=True)
        self.assertAlmostEqual(sheet.cell_value(awb_row, 19), 93.52)
        self._assert_style(workbook, sheet, awb_row, 19, bold=False, fill=50)
        self.assertAlmostEqual(sheet.cell_value(usd_row, 19), expected_fob + 93.52)
        self._assert_style(workbook, sheet, usd_row, 19, bold=True, fill=13)

    def test_totals_match_six_row_reference(self):
        # Итог строки 17; блоки итогов в строках 19, 20, 23, 27 и 28.
        self._assert_totals(self._workbook(6), 6)

    def test_totals_follow_table_length(self):
        for item_count in (2, 9):
            with self.subTest(item_count=item_count):
                self._assert_totals(self._workbook(item_count), item_count)

    def test_red_review_annotations_are_not_exported(self):
        workbook = self._workbook()
        for sheet in workbook.sheets():
            for row in range(sheet.nrows):
                for column in range(sheet.ncols):
                    xf = workbook.xf_list[sheet.cell_xf_index(row, column)]
                    font = workbook.font_list[xf.font_index]
                    self.assertNotEqual(font.colour_index, 10)
                    self.assertFalse(xf.background.fill_pattern
                                     and xf.background.pattern_colour_index == 10)


if __name__ == "__main__":
    unittest.main()
