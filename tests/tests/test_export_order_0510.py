"""Порядок коробок и сводка roses по правкам закупщика 05.10.2026."""
import copy
import os
import sys
import tempfile
import unittest
from collections import Counter

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import xlrd  # noqa: E402

from import_xls_writer import build_combined_factura_xls  # noqa: E402


def item(variety, stems, grade=60, product=None):
    result = {"variety": variety, "stems": stems, "price": 0.5, "total": stems * 0.5,
              "merged_from": ["SOURCE ONE", "SOURCE TWO"]}
    if isinstance(grade, (int, float)):
        result["length_cm"] = grade
    else:
        result["grade_text"] = grade
    if product:
        result["product"] = product
    return result


def box(farm, *items, product="ROSES", size=0.25):
    return {"farm": farm, "product": product, "box_size": size, "items": list(items)}


def batch(**marks):
    return {mark: {
        "boxes": boxes,
        "total_stems": sum(it["stems"] for b in boxes for it in b["items"]),
        "total_fob": sum(it["total"] for b in boxes for it in b["items"]),
        "total_full_boxes": sum(b["box_size"] for b in boxes),
    } for mark, boxes in marks.items()}


class ExportOrderTest(unittest.TestCase):
    def _write(self, data):
        original = copy.deepcopy(data)
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "invoice.xls")
            build_combined_factura_xls(path, data)
            workbook = xlrd.open_workbook(path, formatting_info=True)
        self.assertEqual(data, original, "Экспорт не должен менять исходные коробки")
        return workbook

    @staticmethod
    def _rose_rows(workbook):
        sheet = workbook.sheet_by_name("roses")
        return [sheet.row_values(r) for r in range(3, sheet.nrows)
                if sheet.cell_value(r, 0)]

    def test_boxes_group_across_marks_and_all_sheets_share_numbers(self):
        data = batch(
            A=[box("Star", item("gotcha", 10), item("pink floyd", 15)),
               box("Other", item("don pedro", 30), product="CARNATION"),
               box("Tessa", item("playa blanca", 20))],
            B=[box(" STAR ", item("mix", 25, product="SPRAY ROSES"), product="MIXED"),
               box("Other 2", item("pink", 40), product="HYDRANGEAS"),
               box("tessa", item("full monty", 20)),
               box("matiz", item("explorer", 25))],
        )
        workbook = self._write(data)
        factura = workbook.sheet_by_name("factura")
        warehouse = workbook.sheet_by_name("Склад приёмка")
        explanation = workbook.sheet_by_name("объяснение")
        expected = [
            ("Other", 1, "DON PEDRO"), ("Other 2", 2, "PINK"),
            ("Star", 3, "GOTCHA"), ("Star", 3, "PINK FLOYD"),
            (" STAR ", 4, "MIX"), ("Tessa", 5, "PLAYA BLANCA"),
            ("tessa", 6, "FULL MONTY"), ("matiz", 7, "EXPLORER"),
        ]
        actual = [(factura.cell_value(r, 0), factura.cell_value(r, 1),
                   factura.cell_value(r, 7)) for r in range(10, 18)]
        self.assertEqual(actual, expected)
        self.assertEqual(
            [(warehouse.cell_value(r, 0), warehouse.cell_value(r, 1),
              warehouse.cell_value(r, 5)) for r in range(7, 15)], expected,
        )
        self.assertEqual(
            [(explanation.cell_value(r, 2), explanation.cell_value(r, 1))
             for r in range(4, explanation.nrows)],
            [(farm, number) for farm, number, _variety in expected],
        )
        original_rows = Counter(
            (b["farm"], mark.lower(), it.get("product") or b["product"],
             it["variety"].upper(), it["length_cm"], it["stems"], it["price"], it["total"])
            for mark, info in data.items() for b in info["boxes"] for it in b["items"]
        )
        self.assertEqual(
            Counter(tuple(factura.cell_value(r, c) for c in (0, 2, 3, 7, 10, 12, 14, 15))
                    for r in range(10, 18)), original_rows,
        )
        self.assertEqual([factura.cell_value(r, 5) for r in range(10, 18)],
                         [1, 1, 1, "", 1, 1, 1, 1])
        self.assertEqual([factura.cell_value(r, 6) for r in range(10, 18)],
                         [0.25, 0.25, 0.25, "", 0.25, 0.25, 0.25, 0.25])
        self.assertEqual(factura.cell_value(18, 5), 7)
        self.assertEqual(factura.cell_value(18, 6), 1.75)
        self.assertEqual(factura.cell_value(18, 12), 185)
        self.assertEqual(factura.cell_value(18, 15), 92.5)
        self.assertEqual(warehouse.cell_value(3, 2), 7)
        self.assertEqual(warehouse.cell_value(15, 7), 185)
        other = workbook.sheet_by_name("другие цветы")
        self.assertEqual([other.cell_value(r, 4) for r in (4, 5)], ["DON PEDRO", "PINK"])
        self.assertEqual(other.cell_value(other.nrows - 1, 6), 70)
        self.assertEqual(sum(row[3] for row in self._rose_rows(workbook)), 115)

    def test_roses_aggregate_across_boxes_and_marks_in_requested_order(self):
        data = batch(
            A=[box("matiz", item("esperance", 10, 70)),
               box("tessa", item("full monty", 20, 50)),
               box("corazon", item("nina", 30, 80)),
               box("star", item("gotcha", 40, 60)),
               box("tessa", item("FULL MONTY", 5, 50))],
            B=[box("monterosas", item("queens crown", 50, 50)),
               box("tessa", item("playa blanca", 60, 60)),
               box("matiz", item("explorer", 70, 50)),
               box("tessa", item("mix", 80, 70, product="SPRAY ROSES tessa")),
               box("TESSA", item("full monty", 90, 50))],
        )
        workbook = self._write(data)
        rows = self._rose_rows(workbook)
        self.assertEqual([(row[0], row[1], row[2]) for row in rows], [
            ("tessa", "MIX", 70), ("matiz", "EXPLORER", 50),
            ("tessa", "FULL MONTY", 50), ("monterosas", "QUEENS CROWN", 50),
            ("star", "GOTCHA", 60), ("tessa", "PLAYA BLANCA", 60),
            ("matiz", "ESPERANCE", 70), ("corazon", "NINA", 80),
        ])
        self.assertEqual(rows[2][3:], [115, "A, B"])
        self.assertEqual(sum(row[3] for row in rows), 455)
        sheet = workbook.sheet_by_name("roses")
        self.assertEqual(sheet.cell_value(sheet.nrows - 1, 3), 455)
        # Агрегация справочного листа не объединяет физические коробки factura.
        factura = workbook.sheet_by_name("factura")
        self.assertEqual(factura.cell_value(20, 5), 10)
        self.assertEqual(factura.cell_value(20, 12), 455)
        self.assertEqual(factura.cell_value(20, 15), 227.5)

    def test_item_product_separates_roses_and_preserves_type_farm_and_grade(self):
        data = batch(A=[
            box("tessa", item("mix", 10, 50), item("mix", 20, 60),
                item("mix", 30, 60, product="SPRAY ROSES"),
                item("don pedro", 40, 60, product="CARNATION")),
            box("tessa", item("mix", 50, 60, product="SPRAY ROSES"), product="OTHER"),
            box("tessa", item("mix", 60, 60)),
            box("other farm", item("mix", 70, 60)),
            box("tessa", item("mix", 5, "SELECT")),
        ])
        workbook = self._write(data)
        self.assertEqual(self._rose_rows(workbook), [
            ["tessa", "MIX", 60, 80, "A"],  # spray, отдельная от normal
            ["tessa", "MIX", 50, 10, "A"],
            ["other farm", "MIX", 60, 70, "A"],
            ["tessa", "MIX", 60, 80, "A"],
            ["tessa", "MIX", "SELECT", 5, "A"],
        ])
        other = workbook.sheet_by_name("другие цветы")
        self.assertEqual(other.row_values(4)[1:],
                         ["tessa", 1, "CARNATION", "DON PEDRO", 60, 40, "A"])
        self.assertEqual(other.cell_value(other.nrows - 1, 6), 40)
        factura = workbook.sheet_by_name("factura")
        self.assertEqual([factura.cell_value(r, 1) for r in range(10, 14)], [1, 1, 1, 1])
        self.assertEqual(factura.cell_value(18, 5), 5)
        self.assertEqual(factura.cell_value(18, 12), 285)

    def test_explanation_no_longer_requires_carnation_or_moon_mix(self):
        workbook = self._write(batch(A=[box("farm", item("don pedro", 25), product="CARNATION")]))
        text = workbook.sheet_by_name("объяснение").cell_value(1, 0)
        self.assertIn("гвоздика, включая серию Moon, сохраняется посортово", text)
        self.assertNotIn("гвоздика - всегда MIX", text)
        self.assertNotIn("MOON MIX", text)


if __name__ == "__main__":
    unittest.main()
