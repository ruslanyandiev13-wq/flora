"""Правки 01.10-2: CONSIGNEE брокера и дата поставки в имени Invoice total.

Основные проверки используют синтетический XLS; исходный документ клиента
проверяется отдельно, если он есть локально, и не входит в репозиторий.
"""
import datetime
import io
import os
import sys
import unittest
from unittest.mock import patch

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import xlrd  # noqa: E402
import xlwt  # noqa: E402
from flask import Flask  # noqa: E402
from werkzeug.http import parse_options_header  # noqa: E402

import import_app  # noqa: E402
import import_parser  # noqa: E402
from import_combine import combine_by_mark  # noqa: E402

REAL_INVOICE = os.path.join(ROOT, "Правки 0110-2", "INV-16010.xls")


def broker_invoice(consignee="AGATA", invoice_date=datetime.date(2026, 9, 30)):
    """Создаёт шапку с DATE справа и OBS на колонку левее её заголовка."""
    wb = xlwt.Workbook()
    sh = wb.add_sheet("factura")
    if consignee is not None:
        sh.write(4, 0, "CONSIGNEE ")
        sh.write(4, 2, consignee)
    sh.write(4, 12, "DATE")
    if invoice_date:
        sh.write(4, 17, invoice_date, xlwt.easyxf(num_format_str="dd/mm/yyyy"))
    sh.write(5, 12, "CARGO AGENCY")
    sh.write(5, 17, "TEST FORWARDER")
    columns = {"FARM": 0, "FARM INVOICE": 1, "PRODUCT": 4, "BOX": 6,
               "BOX SIZE": 7, "VARIETY": 8, "GRADE": 11, "TOTAL STEMS": 13,
               "UNIT PRICE": 15, "TOTAL USD": 16, "OBS": 19}
    for name, col in columns.items():
        sh.write(12, col, name)
    rows = [(2, "PINK", 60, 0.5, 30, "AGATA"),
            (1, "BLUE", 40, 0.75, 30, "BESST"),
            (1, "WHITE", 50, 1.0, 50, "BESST")]
    for r, (pieces, variety, stems, price, total, obs) in enumerate(rows, 13):
        values = {"FARM": "TEST FARM", "FARM INVOICE": "1001", "PRODUCT": "HYDRANGEAS",
                  "BOX": pieces, "BOX SIZE": 0.25, "VARIETY": variety,
                  "GRADE": "PREMIUM", "TOTAL STEMS": stems,
                  "UNIT PRICE": price, "TOTAL USD": total}
        for name, value in values.items():
            sh.write(r, columns[name], value)
        sh.write(r, columns["OBS"] - 1, obs)
    buf = io.BytesIO()
    wb.save(buf)
    return import_parser.parse_broker_xls(xlrd.open_workbook(file_contents=buf.getvalue()))


def invoice_entry(data):
    return {"filename": "synthetic-broker.xls", "template": "broker_xls", "data": data}


class BrokerConsigneeTest(unittest.TestCase):
    def test_consignee_overrides_conflicting_obs_without_changing_items(self):
        data = broker_invoice()
        self.assertEqual(data["mark"], "AGATA")
        self.assertEqual(data["invoice_date"], "2026-09-30")
        self.assertEqual(len(data["boxes"]), 4)
        items = [it for box in data["boxes"] for it in box["items"]]
        self.assertEqual([(it["variety"], it["stems"], it["price"], it["total"]) for it in items],
                         [("PINK", 30, 0.5, 15), ("PINK", 30, 0.5, 15),
                          ("BLUE", 40, 0.75, 30), ("WHITE", 50, 1.0, 50)])
        combined = combine_by_mark([invoice_entry(data)])
        self.assertEqual(list(combined), ["AGATA"])
        self.assertEqual((len(combined["AGATA"]["boxes"]), combined["AGATA"]["total_stems"],
                          combined["AGATA"]["total_fob"]), (4, 150, 110))

    def test_obs_fallback_when_consignee_missing_or_empty(self):
        for consignee in (None, "", "   "):
            with self.subTest(consignee=consignee):
                data = broker_invoice(consignee=consignee)
                self.assertEqual(data["mark"], "BESST")
                self.assertEqual(data["invoice_date"], "2026-09-30")
                self.assertEqual(data["totals"], {"total_stems": 150, "total_fob": 110})

    @unittest.skipUnless(os.path.isfile(REAL_INVOICE), "нет локального INV-16010.xls")
    def test_real_invoice_stays_wholly_agata(self):
        data, template = import_parser.parse_invoice_file(REAL_INVOICE)
        self.assertEqual((template, data["mark"], data["invoice_date"]),
                         ("broker_xls", "AGATA", "2026-09-30"))
        self.assertEqual((len(data["boxes"]), data["totals"]),
                         (6, {"total_stems": 180, "total_fob": 154.8}))
        combined = combine_by_mark([invoice_entry(data)])
        self.assertEqual(list(combined), ["AGATA"])
        self.assertEqual(combined["AGATA"]["total_full_boxes"], 1.5)


class BrokerDownloadTest(unittest.TestCase):
    def setUp(self):
        app = Flask(__name__)
        app.config.update(TESTING=True, SECRET_KEY="broker-download-test")
        app.register_blueprint(import_app.import_bp)
        self.client = app.test_client()
        with self.client.session_transaction() as session:
            session["import_token"] = "broker-download-test"

    def check_downloads(self, data, delivery_date, awb_doc=None):
        pending = {"invoices": [invoice_entry(data)], "awb": {}, "awb_doc": awb_doc}
        for route, label in [("/import/factura/download", (awb_doc or {}).get("awb_no") or "AGATA"),
                             ("/import/factura/AGATA/download", "AGATA")]:
            with self.subTest(route=route, date=delivery_date, awb=bool(awb_doc)):
                with patch.object(import_app, "_load_pending", return_value=pending), \
                        patch.object(import_app, "_charge_once") as charge:
                    response = self.client.get(route)
                self.assertEqual(response.status_code, 200)
                _, disposition = parse_options_header(response.headers["Content-Disposition"])
                date_prefix = delivery_date.strftime("%d.%m ") if delivery_date else ""
                filename = f"Invoice total {date_prefix}{label}.xls"
                self.assertEqual(disposition["filename"], filename)
                self.assertEqual(charge.call_args.args[-1], filename)
                wb = xlrd.open_workbook(file_contents=response.data)
                sh = wb.sheet_by_name("factura")
                self.assertEqual(sh.cell_value(1, 1), "AGATA")
                self.assertEqual([sh.cell_value(r, 2) for r in range(10, 14)], ["agata"] * 4)
                self.assertEqual(sum(sh.cell_value(r, 12) for r in range(10, 14)), 150)
                self.assertAlmostEqual(sum(sh.cell_value(r, 15) for r in range(10, 14)), 110)
                if delivery_date:
                    actual_date = xlrd.xldate_as_datetime(sh.cell_value(7, 7), wb.datemode).date()
                    self.assertEqual(actual_date, delivery_date)
                else:
                    self.assertEqual(sh.cell_value(7, 7), "")
                if awb_doc:
                    self.assertEqual(sh.cell_value(5, 16), awb_doc["awb_no"])
                response.close()

    def test_both_downloads_include_delivery_date(self):
        self.check_downloads(broker_invoice(), datetime.date(2026, 10, 7))

    def test_awb_identifier_and_flight_delivery_date_are_preserved(self):
        self.check_downloads(broker_invoice(), datetime.date(2026, 10, 4),
                             {"awb_no": "999-1234 5678", "flight_date": "2026-09-29", "airport": "SVO"})

    def test_unknown_date_is_not_invented(self):
        self.check_downloads(broker_invoice(invoice_date=None), None)


if __name__ == "__main__":
    unittest.main()
