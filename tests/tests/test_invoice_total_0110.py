"""
Правки закупщика 2026-10-01 (папка "Правки 0110"): раскладка Invoice total,
листы «Склад приёмка»/roses/«другие цветы», MATHIOLA, стебли Rosaprima и
скан HAWB IFC (HAWB 6476 BESST), который раньше вводили руками.

Реальные документы в репозитории не хранятся - без них тесты на них
пропускаются.
"""
import datetime
import glob
import os
import shutil
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import xlrd  # noqa: E402

import import_parser  # noqa: E402
from import_combine import box_farm_product, combine_by_mark  # noqa: E402
from import_xls_writer import build_combined_factura_xls  # noqa: E402

DIR = os.path.join(ROOT, "Правки 0110")
BATCH_DIR = os.path.join(ROOT, "ТЕСТ3009")
SCAN = os.path.join(DIR, "HAWB 6476 BESST.pdf")

# Так tesseract (--psm 4) читает нужные строки скана IFC.
OCR_SAMPLE = """IFC -000380445 IFC-000380445
Accounting Information
MARK BESST
BOG - AEROPUERTO INTERNACIONAL EL DORADO
DOH QATAR SVO |QR USD! PX
Handling Information PIEZAS: 10 - FULL: 2.5 TRM $3, 349.63 - AWB: 157-58076476 **KEEP
10 PCS
10 82 82 AS AGREE
FRESH CUT FLOWERS
CONDOR ANDINO S.A.S 1 4
FLORES DE SERREZUELA 8.A.S 1.5 6
10 82
AS AGREED : DUE AGENT 10.00
DUE CARRIER 12.00
09/29/2026 BOGOTA - COLOMBIA
"""


class ScannedHawbTest(unittest.TestCase):
    def test_parse_ocr_text(self):
        h = import_parser.parse_scanned_hawb(OCR_SAMPLE)
        self.assertEqual((h["mark"], h["mawb"], h["hawb"]), ("BESST", "157-5807 6476", "IFC-000380445"))
        self.assertEqual((h["pieces"], h["gross_weight"], h["chargeable_weight"], h["total_full"]),
                         (10, 82.0, 82.0, 2.5))
        self.assertEqual((h["origin"], h["origin_country"], h["airport"], h["flight_date"]),
                         ("BOG", "colombia", "SVO", "2026-09-29"))
        self.assertEqual(h["airline"], "QATAR AIRWAYS")
        self.assertEqual([(g["name"], g["full_boxes"], g["pieces"]) for g in h["growers"]],
                         [("CONDOR ANDINO S.A.S", 1.0, 4), ("FLORES DE SERREZUELA S.A.S", 1.5, 6)])

    def test_not_a_hawb(self):
        self.assertIsNone(import_parser.parse_scanned_hawb("INVOICE 123\nTOTAL 10"))

    def test_pieces_label_split_by_ocr(self):
        h = import_parser.parse_scanned_hawb(OCR_SAMPLE.replace("PIEZAS", "PIE ZAS"))
        self.assertEqual((h["pieces"], h["total_full"], h["gross_weight"], h["chargeable_weight"]),
                         (10, 2.5, 82.0, 82.0))

    @unittest.skipUnless(os.path.isfile(os.path.join(ROOT, "tests", "test1009_columb",
                                                     "HAWB AS AGREED AWB 6502 AGATA.pdf"))
                         and shutil.which("tesseract"), "нет скана AGATA или tesseract")
    def test_real_scan_with_split_pieces_label(self):
        h = import_parser.parse_forwarder_hawb_pdf(os.path.join(
            ROOT, "tests", "test1009_columb", "HAWB AS AGREED AWB 6502 AGATA.pdf"))
        self.assertEqual((h["mark"], h["pieces"], h["gross_weight"], h["chargeable_weight"], h["total_full"]),
                         ("AGATA", 5, 62.0, 62.0, 1.25))

    @unittest.skipUnless(os.path.exists(SCAN) and shutil.which("tesseract"), "нет скана или tesseract")
    def test_real_scan(self):
        h = import_parser.parse_forwarder_hawb_pdf(SCAN)
        self.assertEqual((h["mark"], h["mawb"], h["pieces"], h["gross_weight"]),
                         ("BESST", "157-5807 6476", 10, 82.0))
        self.assertEqual(len(h["growers"]), 2)

    def test_import_awb_doc_and_moscow_date(self):
        import import_app
        h = import_parser.parse_scanned_hawb(OCR_SAMPLE)
        with patch.object(import_app.db, "get_setting", return_value="8"):
            awb_doc = import_app._awb_doc_from_hawb(h)
        self.assertEqual(import_app._awb_from_doc("BESST", awb_doc)["chargeable_weight"], 82.0)
        # Вылет вторник 29.09 в Москву - поставка в воскресенье 04.10.
        self.assertEqual(import_app._delivery_date({"awb_doc": awb_doc, "invoices": []}),
                         datetime.date(2026, 10, 4))


class MathiolaTest(unittest.TestCase):
    def test_alias(self):
        inv = {"template": "x", "data": {"supplier": "NINTANGA"}}
        self.assertEqual(box_farm_product(inv, {"product": "S MATHIOLAS"})[1], "MATHIOLA")


@unittest.skipUnless(os.path.exists(os.path.join(DIR, "RU_962563.pdf")), "нет инвойса RU_962563")
class RosaprimaStemsTest(unittest.TestCase):
    def test_bunches_times_stems(self):
        data, _ = import_parser.parse_invoice_file(os.path.join(DIR, "RU_962563.pdf"))
        self.assertEqual([(it["variety"], it["stems"]) for b in data["boxes"] for it in b["items"]],
                         [("EXPLORER", 300), ("ESPERANCE", 100)])


@unittest.skipUnless(os.path.isdir(BATCH_DIR), "нет документов ТЕСТ3009")
class InvoiceTotalLayoutTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        invs = []
        for path in sorted(glob.glob(os.path.join(BATCH_DIR, "*"))):
            data, template = import_parser.parse_invoice_file(path)
            invs.append({"filename": os.path.basename(path), "data": data, "template": template})
        cls._tmp = tempfile.TemporaryDirectory()
        path = os.path.join(cls._tmp.name, "it.xls")
        build_combined_factura_xls(path, combine_by_mark(invs), {"awb_no": "235-7841 9854"},
                                   consignee="BESST MOS 2", delivery_date=datetime.date(2026, 10, 4))
        cls.wb = xlrd.open_workbook(path, formatting_info=True)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def test_sheets(self):
        self.assertEqual(self.wb.sheet_names(),
                         ["factura", "Склад приёмка", "roses", "другие цветы", "объяснение"])

    def test_factura_header(self):
        ws = self.wb.sheet_by_name("factura")
        self.assertEqual(ws.cell_value(1, 1), "BESST MOS 2")      # B2
        self.assertEqual(ws.cell_value(6, 3), "мест")             # D7
        self.assertEqual(ws.cell_value(8, 0), "awb")              # A9
        date = xlrd.xldate_as_datetime(ws.cell_value(7, 7), self.wb.datemode).date()
        self.assertEqual(date, datetime.date(2026, 10, 4))        # H8
        hidden = sorted(c for c, info in ws.colinfo_map.items() if info.hidden)
        self.assertEqual(hidden, [4, 8, 9, 13])                   # E, I, J, N

    def test_totals_merged(self):
        ws = self.wb.sheet_by_name("factura")
        row = next(r for r in range(ws.nrows) if ws.cell_value(r, 9) == "TOTAL STEMS")
        self.assertIn((row, row + 1, 9, 13), ws.merged_cells)    # J:M
        self.assertIn((row, row + 1, 13, 16), ws.merged_cells)   # N:P
        usd = next(r for r in range(ws.nrows) if ws.cell_value(r, 9) == "TOTAL USD")
        self.assertIn((usd, usd + 1, 9, 12), ws.merged_cells)    # J:L

    def test_roses_and_other_sheets(self):
        roses = self.wb.sheet_by_name("roses")
        self.assertEqual(roses.row_values(2), ["Farm", "Variety", "Длина", "Стеблей", "Метка"])
        other = self.wb.sheet_by_name("другие цветы")
        products = {other.cell_value(r, 3) for r in range(4, other.nrows)}
        self.assertIn("MATHIOLA", products)
        self.assertFalse(any("ROSES" in str(p) for p in products))


if __name__ == "__main__":
    unittest.main()
