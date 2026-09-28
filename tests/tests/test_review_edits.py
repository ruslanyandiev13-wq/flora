"""
Тесты ручных правок позиций голландского модуля (review_edits.py): правки
накладываются поверх разобранных данных, сумма пересчитывается как
кол-во × цена, а итог товара сдвигается ровно на разницу.
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import review_edits as re_  # noqa: E402


def _boxes():
    return [{"box_no": 1, "fust": "AAA", "items": [
        {"aantal": 40, "omschrijving": "Rosa A", "prijs": 0.75, "lengte": 70, "gew": None, "bedrag": 30.0},
        {"aantal": 20, "omschrijving": "Rosa B", "prijs": 1.0, "lengte": 60, "gew": None, "bedrag": 20.0},
    ]}]


class ReviewEditsTest(unittest.TestCase):
    def test_quantity_edit_recalculates_amount(self):
        edits = re_.set_edit({}, 1, "aantal", 100, item_idx=0)
        boxes, delta = re_.apply_edits(_boxes(), edits)
        it = boxes[0]["items"][0]
        self.assertEqual(it["bedrag"], 75.0)
        self.assertEqual(it["original"], {"aantal": 40, "bedrag": 30.0})
        self.assertEqual(delta, 45.0)

    def test_manual_amount_is_not_overwritten(self):
        edits = re_.set_edit({}, 1, "aantal", 100, item_idx=0)
        re_.set_edit(edits, 1, "bedrag", 70, item_idx=0)
        boxes, delta = re_.apply_edits(_boxes(), edits)
        self.assertEqual(boxes[0]["items"][0]["bedrag"], 70)
        self.assertEqual(delta, 40.0)

    def test_box_fust_and_source_untouched(self):
        source = _boxes()
        boxes, delta = re_.apply_edits(source, re_.set_edit({}, 1, "fust", "PPS"))
        self.assertEqual(boxes[0]["fust"], "PPS")
        self.assertEqual(source[0]["fust"], "AAA")
        self.assertEqual(delta, 0)

    def test_parse_value(self):
        self.assertEqual(re_.parse_value("prijs", "0,85"), 0.85)
        self.assertEqual(re_.parse_value("aantal", "50"), 50)
        self.assertIsNone(re_.parse_value("gew", ""))
        for field, raw in [("aantal", ""), ("aantal", "2.5"), ("prijs", "abc"), ("prijs", "-1")]:
            with self.assertRaises(ValueError):
                re_.parse_value(field, raw)

    def test_count_edits(self):
        edits = re_.set_edit({}, 1, "fust", "PPS")
        re_.set_edit(edits, 1, "aantal", 10, item_idx=0)
        re_.set_edit(edits, 1, "prijs", 1, item_idx=1)
        self.assertEqual(re_.count_edits(edits), 3)


if __name__ == "__main__":
    unittest.main()
