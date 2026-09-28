"""
Регрессия на правки закупщика по партии AWB 369-1151 1964 (папка
"Правки2 28.09", его файл "Invoice total 30.09 VADIM.xls"):

1. SHIPPING INVOICE Rosaprima Cia. Ltda. (RU_968408) не разбирался вообще -
   пропадала 93-я коробка.
2. Вес метки берётся из её отдельной накладной внутри AWB (house AWB), а не
   делится пропорционально местам; "ставка за кг" = итог AWB / платный вес.
3. Коробка TESSA с кодом SB - 1/16 полной (так в AWB), а не 1/6 из "Number
   in Fulls" инвойса.

Реальные документы в репозитории не хранятся - без них тесты пропускаются.
"""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import import_parser  # noqa: E402

DIR = os.path.join(ROOT, "Правки2 28.09")


def _path(name):
    return os.path.join(DIR, name)


class BoxSizeCodeTest(unittest.TestCase):
    def test_small_box(self):
        self.assertEqual(import_parser._box_size_from_code("SB"), 0.0625)
        self.assertEqual(import_parser._box_size_from_code("QB"), 0.25)
        self.assertEqual(import_parser._box_size_from_code("HB"), 0.5)


@unittest.skipUnless(os.path.isdir(DIR), "нет документов партии 369-1151 1964")
class Batch3009Test(unittest.TestCase):
    def test_rosaprima_shipping_invoice(self):
        data, template = import_parser.parse_invoice_file(_path("RU_968408.pdf"))
        self.assertEqual(template, "rosaprima_ec")
        self.assertEqual(data["mark"], "POLINA")
        self.assertEqual(len(data["boxes"]), 1)
        box = data["boxes"][0]
        self.assertEqual((box["box_type"], box["box_size"]), ("JL", 0.5))
        self.assertEqual([(it["variety"], it["length_cm"], it["stems"], it["price"], it["total"])
                          for it in box["items"]],
                         [("MONDIAL", 70, 200, 0.69, 138.0), ("ESPERANCE", 70, 100, 0.55, 55.0)])
        self.assertEqual(data["totals"]["total_fob"], 193.0)

    def test_tessa_small_box(self):
        data, _ = import_parser.parse_invoice_file(_path("90832514.pdf"))
        self.assertEqual([b["box_size"] for b in data["boxes"]], [0.0625])

    def test_awb_house_weights(self):
        awb = import_parser.parse_awb_pdf(_path("AWB_369-1151 1964.pdf"))
        houses = awb["houses"]
        self.assertEqual({m: (h["gross_weight"], h["chargeable_weight"]) for m, h in houses.items()},
                         {"DAMIR": (406, 415), "POLINA": (1024, 1024), "VADIM": (618, 621)})
        self.assertEqual(round(sum(h["total_awb"] for h in houses.values()), 2), awb["total_awb"])

    def test_all_boxes_found(self):
        import glob
        boxes = 0
        for path in glob.glob(os.path.join(DIR, "*.pdf")):
            if import_parser.parse_awb_pdf(path):
                continue
            boxes += len(import_parser.parse_invoice_file(path)[0]["boxes"])
        self.assertEqual(boxes, 93)


if __name__ == "__main__":
    unittest.main()
