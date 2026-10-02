"""Сортность PR применяется только к гортензии, во всех копиях итоговой таблицы."""
import copy
import io
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import xlrd

from import_combine import combine_by_mark
from import_parser import parse_invoice_file
from import_xls_writer import build_combined_factura_xls


def invoice(product="HYDRANGEAS", grade="PREMIUN", length=None):
    return {"filename": "synthetic.xls", "template": "broker_xls", "data": {
        "supplier": "TEST FARM", "mark": "AGATA", "boxes": [{
            "farm": "TEST FARM", "product": product, "box_size": 0.25,
            "items": [{"variety": "PINK", "length_cm": length, "grade_text": grade,
                       "stems": 30, "price": 0.91, "total": 27.3}],
        }],
    }}


class HydrangeasGradeTest(unittest.TestCase):
    def test_pr_prefix_is_normalized_without_mutating_source_or_amounts(self):
        for grade in ("PREMIUM", "PREMIUN", "PREM", "PR", " premium "):
            with self.subTest(grade=grade):
                inv = invoice(product=" hydrangeas ", grade=grade)
                original = copy.deepcopy(inv)
                combined = combine_by_mark([inv])["AGATA"]
                item = combined["boxes"][0]["items"][0]
                self.assertEqual(item["grade_text"], "PR")
                self.assertIsNone(item["length_cm"])
                self.assertEqual((item["stems"], item["price"], item["total"]), (30, 0.91, 27.3))
                self.assertEqual((combined["total_stems"], combined["total_fob"]), (30, 27.3))
                self.assertEqual(inv, original)

    def test_other_cultures_grades_and_numeric_lengths_are_unchanged(self):
        cases = [("CARNATION", "PREMIUM", None), ("ROSES test", "PREMIUN", None),
                 ("HYDRANGEAS", "SELECT", None), ("HYDRANGEAS", "FANCY", None),
                 ("HYDRANGEAS", None, 60), ("HYDRANGEAS", "PREMIUM", 60)]
        for product, grade, length in cases:
            with self.subTest(product=product, grade=grade, length=length):
                combined = combine_by_mark([invoice(product, grade, length)])
                item = combined["AGATA"]["boxes"][0]["items"][0]
                self.assertEqual((item["grade_text"], item["length_cm"]), (grade, length))

    def test_factura_warehouse_and_other_flowers_use_same_grade(self):
        combined = combine_by_mark([invoice()])
        buf = io.BytesIO()
        build_combined_factura_xls(buf, combined)
        wb = xlrd.open_workbook(file_contents=buf.getvalue())
        # GRADE и STEMS: factura K/M, склад G/H, другие цветы F/G.
        for sheet, row, grade_col, stems_col in [
                ("factura", 10, 10, 12), ("Склад приёмка", 7, 6, 7), ("другие цветы", 4, 5, 6)]:
            with self.subTest(sheet=sheet):
                ws = wb.sheet_by_name(sheet)
                self.assertEqual(ws.cell_value(row, grade_col), "PR")
                self.assertEqual(ws.cell_value(row, stems_col), 30)

    @unittest.skipUnless(os.path.isfile(os.path.join(ROOT, "Правки 0110-2", "INV-16010.xls")),
                         "нет локального INV-16010.xls")
    def test_real_broker_hydrangeas_grades(self):
        path = os.path.join(ROOT, "Правки 0110-2", "INV-16010.xls")
        data, template = parse_invoice_file(path)
        original = copy.deepcopy(data)
        combined = combine_by_mark([{"filename": os.path.basename(path), "template": template,
                                     "data": data}])["AGATA"]
        items = [it for box in combined["boxes"] for it in box["items"]]
        self.assertEqual(len(items), 6)
        self.assertTrue(all(it["product"] == "HYDRANGEAS" and it["grade_text"] == "PR" for it in items))
        self.assertTrue(all(it["stems"] == 30 for it in items))
        self.assertEqual((combined["total_stems"], combined["total_fob"]), (180, 154.8))
        self.assertEqual(data, original)
