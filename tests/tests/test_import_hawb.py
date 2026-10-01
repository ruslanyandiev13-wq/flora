"""Загрузка HAWB должна сохранять веса, метку и московскую дату поставки."""
import datetime
import io
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

from flask import Flask

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import import_app


class ImportHawbUploadTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.pending = {}
        state = patch.multiple(import_app, PENDING=self.pending, PENDING_DIR=self.tmp.name,
                               UPLOAD_DIR=self.tmp.name)
        state.start()
        self.addCleanup(state.stop)
        rate = patch.object(import_app.db, "get_setting", side_effect=lambda key, default: str(default))
        rate.start()
        self.addCleanup(rate.stop)
        invoice = patch.object(import_app, "parse_invoice_file", return_value=(
            {"supplier": "Test farm", "mark": "VIKA", "boxes": []}, "test"))
        invoice.start()
        self.addCleanup(invoice.stop)
        app = Flask(__name__)
        app.secret_key = "test"
        app.config["TESTING"] = True
        app.register_blueprint(import_app.import_bp)
        self.client = app.test_client()

    def upload(self, pdf):
        response = self.client.post("/import/upload", data={"invoices": [
            (io.BytesIO(pdf), "airwaybill.pdf"), (io.BytesIO(b"invoice"), "invoice.xls")
        ]})
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.location.endswith("/import/review"))
        self.assertEqual(len(self.pending), 1)
        return next(iter(self.pending.values()))

    def test_forwarder_hawb_takes_priority_over_generic_awb(self):
        hawb = {"mawb": "157-0401 2400", "hawb": "306 0594 4328", "mark": "VIKA",
                "origin_country": "ecuador", "airport": "SVO", "flight_date": "2026-09-22",
                "pieces": 9, "gross_weight": 221.0, "chargeable_weight": 221.0, "total_full": 4.25}
        with patch.object(import_app, "parse_forwarder_hawb_pdf", return_value=hawb), \
                patch.object(import_app, "parse_awb_pdf", return_value={"marks": {}}) as generic:
            pending = self.upload(b"pdf")
        generic.assert_not_called()
        awb = pending["awb_doc"]
        self.assertEqual((awb["pieces"], awb["gross_weight"], awb["rate_per_kg"]), (9, 221.0, 8.1))
        self.assertEqual(awb["marks"], {"VIKA": {"pieces": 9, "full_boxes": 4.25}})
        self.assertEqual(import_app._delivery_date(pending), datetime.date(2026, 9, 27))

    def test_regular_awb_keeps_document_rate_and_amsterdam_date(self):
        awb = {"awb_no": "369-1151 1964", "pieces": 93, "gross_weight": 2048.0,
               "chargeable_weight": 2060.0, "rate_per_kg": 3.25, "flight_date": "2026-09-23",
               "marks": {"VIKA": {"pieces": 93, "full_boxes": 41.56}}}
        with patch.object(import_app, "parse_forwarder_hawb_pdf", return_value=None), \
                patch.object(import_app, "parse_awb_pdf", return_value=awb):
            pending = self.upload(b"pdf")
        self.assertEqual(pending["awb_doc"]["rate_per_kg"], 3.25)
        self.assertEqual(pending["awb_doc"]["gross_weight"], 2048.0)
        self.assertEqual(import_app._delivery_date(pending), datetime.date(2026, 9, 30))

    @unittest.skipUnless(os.path.isfile(os.path.join(ROOT, "Тестируем Москву", "HAWB 2400_VIKA.pdf")),
                         "нет HAWB партии VIKA")
    def test_real_text_hawb_upload(self):
        with open(os.path.join(ROOT, "Тестируем Москву", "HAWB 2400_VIKA.pdf"), "rb") as f:
            pending = self.upload(f.read())
        awb = pending["awb_doc"]
        self.assertEqual((awb["awb_no"], awb["pieces"], awb["gross_weight"], awb["rate_per_kg"]),
                         ("157-0401 2400", 9, 221.0, 8.1))
        self.assertEqual(import_app._delivery_date(pending), datetime.date(2026, 9, 27))

    @unittest.skipUnless(os.path.isfile(os.path.join(ROOT, "Правки2 28.09", "AWB_369-1151 1964.pdf")),
                         "нет AWB партии 369-1151 1964")
    def test_real_regular_awb_upload(self):
        with open(os.path.join(ROOT, "Правки2 28.09", "AWB_369-1151 1964.pdf"), "rb") as f:
            pending = self.upload(f.read())
        awb = pending["awb_doc"]
        self.assertEqual((awb["pieces"], awb["gross_weight"], awb["rate_per_kg"]), (93, 2048.0, 3.25))
        self.assertEqual(set(awb["houses"]), {"DAMIR", "POLINA", "VADIM"})
        self.assertEqual(import_app._delivery_date(pending), datetime.date(2026, 9, 30))
