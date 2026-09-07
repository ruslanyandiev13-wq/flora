"""
Тесты на структуру генерируемого .xls - в первую очередь на то, что строка
"Палета,$"/"КурсUSD" (нужна 1С для расчёта логистики) пишется строго на одну
строку выше шапки таблицы товаров, и что шапка/коробки при этом не сдвигаются.
Если кто-то в будущем поменяет раскладку строк в xls_writer.py, эти тесты
должны немедленно об этом сообщить, а не дать 1С молча получить битый файл.
"""
import datetime
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import xlrd

import xls_writer
from xls_writer import build_xls, HEADER_ROW


def _sample_boxes():
    return [{"box_no": 1, "fust": "AAA", "items": [
        {"aantal": 20, "omschrijving": "Test Flower", "prijs": 1.5,
         "lengte": 70, "gew": None, "bedrag": 30.0},
    ]}]


def _sample_header():
    return {
        "petia": "PETIA", "paklijst_title": "Paklijst Export Unie Flora",
        "week_no": 33, "date": datetime.date(2026, 8, 18), "blad_no": 1,
        "debnr": "IRIS", "naam": 'LLC "Flora Cargo"',
    }


class TestXlsStructure(unittest.TestCase):

    def _build(self, path, **kwargs):
        build_xls(path, _sample_boxes(), _sample_header(), 30.0, **kwargs)
        return xlrd.open_workbook(path, formatting_info=False).sheet_by_index(0)

    def test_header_row_position_unchanged(self):
        ws = self._build("/tmp/_test_header_row.xls",
                          pallet_cost_usd=1650.0, usd_rate=91.245)
        header_row = ws.row_values(HEADER_ROW)
        self.assertIn("Omschrijving", header_row)
        self.assertIn("Kolli", header_row)
        self.assertIn("Prijs", header_row)
        # первая коробка должна начинаться через одну пустую строку после шапки
        self.assertEqual(ws.cell_value(HEADER_ROW + 2, 1), 1)  # box_no
        self.assertEqual(ws.cell_value(HEADER_ROW + 2, 2), "AAA")  # fust

    def test_logistics_row_written_with_values(self):
        ws = self._build("/tmp/_test_logistics_values.xls",
                          pallet_cost_usd=1650.0, usd_rate=91.245)
        row = ws.row_values(HEADER_ROW - 1)
        self.assertEqual(row[1], "Палета,$")
        self.assertEqual(row[2], 1650.0)
        self.assertEqual(row[4], "КурсUSD")
        self.assertEqual(row[5], 91.245)

    def test_logistics_values_optional_when_not_set(self):
        # Если курс/стоимость палеты ещё не заданы - файл всё равно генерируется,
        # подписи остаются, числовые ячейки просто пустые (не 0, не падение).
        ws = self._build("/tmp/_test_logistics_empty.xls")
        row = ws.row_values(HEADER_ROW - 1)
        self.assertEqual(row[1], "Палета,$")
        self.assertEqual(row[2], "")
        self.assertEqual(row[4], "КурсUSD")
        self.assertEqual(row[5], "")

    def test_other_rows_unaffected_by_logistics_row(self):
        without = self._build("/tmp/_test_regress_without.xls")
        with_vals = self._build("/tmp/_test_regress_with.xls",
                                 pallet_cost_usd=1650.0, usd_rate=91.245)
        # всё начиная с шапки (row HEADER_ROW) должно быть побайтово одинаковым
        for r in range(HEADER_ROW, without.nrows):
            self.assertEqual(
                without.row_values(r), with_vals.row_values(r),
                f"Строка {r} изменилась из-за добавления логистики - так быть не должно",
            )

    def test_raises_if_logistics_row_already_occupied(self):
        # Симулируем ситуацию "структуру выше поменяли, и строка над шапкой
        # больше не пустая" - подмена HEADER_ROW так, чтобы строка логистики
        # (HEADER_ROW-1) совпала со строкой "Debnr."/"Naam" (row 4), которая
        # уже точно занята.
        original = xls_writer.HEADER_ROW
        xls_writer.HEADER_ROW = 5
        try:
            with self.assertRaises(RuntimeError):
                build_xls("/tmp/_test_guard.xls", _sample_boxes(), _sample_header(),
                          30.0, pallet_cost_usd=1650.0, usd_rate=91.245)
        finally:
            xls_writer.HEADER_ROW = original


if __name__ == "__main__":
    unittest.main()
