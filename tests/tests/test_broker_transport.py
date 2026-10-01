"""Логистика брокера: напечатанный итог AWB сохраняется без потери копеек."""
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

import import_app  # noqa: E402
import import_parser  # noqa: E402
from import_combine import combine_by_mark  # noqa: E402
from import_xls_writer import build_combined_factura_xls  # noqa: E402

REAL_INVOICE = os.path.join(ROOT, "Правки 0110-2", "INV-16010.xls")


def broker_with_transport(cost=100.01, gross=300, chargeable=333,
                          printed_stems=60, printed_fob=150):
    wb = xlwt.Workbook()
    sh = wb.add_sheet("factura")
    sh.write(4, 0, "CONSIGNEE")
    sh.write(4, 2, "AGATA")
    sh.write(4, 12, "DATE")
    sh.write(4, 17, datetime.date(2026, 9, 30), xlwt.easyxf(num_format_str="dd/mm/yyyy"))
    headers = {0: "FARM", 1: "FARM INVOICE", 4: "PRODUCT", 6: "BOX",
               7: "BOX SIZE", 8: "VARIETY", 11: "GRADE", 13: "TOTAL STEMS",
               15: "UNIT PRICE", 16: "TOTAL USD", 19: "OBS"}
    for col, value in headers.items():
        sh.write(12, col, value)
    row = {0: "TEST FARM", 1: "1001", 4: "HYDRANGEAS", 6: 1, 7: 0.25,
           8: "PINK", 11: "PREMIUM", 13: 60, 15: 2.5, 16: 150, 18: "AGATA"}
    for col, value in row.items():
        sh.write(13, col, value)
    totals = [("TOTAL STEMS", printed_stems), ("TOTAL FLOWERS FOB USD", printed_fob),
              ("TOTAL CHARGEABLE WEIGHT(Kg)", chargeable), ("GROSS WEIGHT", gross),
              ("TOTAL AWB", cost), ("TOTAL USD", printed_fob + cost if cost is not None else None)]
    for r, (label, value) in enumerate(totals, 17):
        sh.write(r, 10, label)
        if value is not None:
            # Номер столбца меняется между брокерскими шаблонами: искать
            # нужно число справа от подписи, а не жёстко U или O.
            sh.write(r, 20 if r % 2 else 14, value)
    output = io.BytesIO()
    wb.save(output)
    return import_parser.parse_broker_xls(xlrd.open_workbook(file_contents=output.getvalue()))


def combine(data):
    return combine_by_mark([{"filename": "broker.xls", "template": "broker_xls", "data": data}])


def xls_totals(by_mark, awb_doc=None):
    output = io.BytesIO()
    build_combined_factura_xls(output, by_mark, awb_doc)
    sh = xlrd.open_workbook(file_contents=output.getvalue()).sheet_by_name("factura")
    result = {}
    for r in range(sh.nrows):
        label = sh.cell_value(r, 9)
        if label in {"TOTAL FLOWERS FOB USD", "TOTAL AWB", "TOTAL USD",
                     "TOTAL CHARGEABLE WEIGHT(Kg)", "GROSS WEIGHT"}:
            result[label] = next(sh.cell_value(r, c) for c in range(10, sh.ncols)
                                 if sh.cell_type(r, c) == xlrd.XL_CELL_NUMBER)
    return result


class BrokerTransportTest(unittest.TestCase):
    def test_transport_adds_invoices_without_merging_distinct_weights(self):
        data = broker_with_transport()
        invoices = [{"filename": name, "template": "broker_xls", "data": data}
                    for name in ("one.xls", "two.xls")]
        by_mark = combine_by_mark(invoices)
        import_app._apply_awb(by_mark, {})
        awb = by_mark["AGATA"]["awb"]
        self.assertEqual((awb["total_awb"], awb["gross_weight"], awb["chargeable_weight"]),
                         (200.02, 600, 666))

    def test_partial_invoice_transport_does_not_replace_full_awb(self):
        invoices = [{"filename": name, "template": "broker_xls", "data": data}
                    for name, data in (("one.xls", broker_with_transport()),
                                       ("two.xls", broker_with_transport(cost=None, gross=None,
                                                                         chargeable=None)))]
        by_mark = combine_by_mark(invoices)
        import_app._apply_awb(by_mark, {}, {"pieces": 2, "marks": {"AGATA": {"pieces": 2}},
                                          "gross_weight": 600, "chargeable_weight": 666,
                                          "rate_per_kg": 1})
        self.assertEqual(by_mark["AGATA"]["awb"]["total_awb"], 666)
        self.assertEqual(by_mark["AGATA"]["awb"]["source"], "awb")

    def test_legacy_astoria_transport_still_uses_its_printed_cost(self):
        data = broker_with_transport()
        data["transport"] = {"cost_usd": 100.01, "weight_kg": 333}
        by_mark = combine(data)
        import_app._apply_awb(by_mark, {})
        awb = by_mark["AGATA"]["awb"]
        self.assertEqual((awb["total_awb"], awb["gross_weight"], awb["chargeable_weight"]),
                         (100.01, 333, 333))

    def test_interface_shows_prefilled_cost_and_optional_manual_form(self):
        from flask import Flask
        from jinja2 import ChoiceLoader, DictLoader
        from bs4 import BeautifulSoup

        app = Flask(__name__, template_folder=os.path.join(ROOT, "templates"))
        app.config.update(TESTING=True, SECRET_KEY="broker-transport-test")
        app.jinja_loader = ChoiceLoader([DictLoader({"base.html": "{% block content %}{% endblock %}"}),
                                        app.jinja_loader])
        app.register_blueprint(import_app.import_bp)
        client = app.test_client()
        with client.session_transaction() as session:
            session["import_token"] = "broker-transport-test"
        for cost in (100.01, 0, None):
            with self.subTest(cost=cost):
                pending = {"invoices": [{"filename": "broker.xls", "template": "broker_xls",
                                          "data": broker_with_transport(cost=cost)}]}
                with patch.object(import_app, "_load_pending", return_value=pending):
                    response = client.get("/import/factura")
                self.assertEqual(response.status_code, 200)
                soup = BeautifulSoup(response.data, "html.parser")
                self.assertEqual(soup.find("details").has_attr("open"), cost is None)
                if cost is not None:
                    self.assertIn(f"AWB: {cost:.2f} $", soup.get_text())
                    self.assertIn(f"Итого: {150 + cost:.2f} $", soup.get_text())
                    self.assertEqual(float(soup.find("input", {"name": "chargeable_weight"})["value"]),
                                     333)
                response.close()

    def test_parser_reads_separate_weights_and_printed_totals(self):
        data = broker_with_transport(printed_stems=61, printed_fob=149.99)
        transport = data["transport"]
        self.assertEqual((transport["cost_usd"], transport["weight_kg"],
                          transport["gross_weight"], transport["chargeable_weight"]),
                         (100.01, 333, 300, 333))
        self.assertEqual(data["totals"], {"total_stems": 61, "total_fob": 149.99})
        # Печатные контрольные итоги не переписывают исходные строки.
        self.assertEqual(data["boxes"][0]["items"][0]["total"], 150)

    def test_exact_cost_survives_rounded_display_rate(self):
        by_mark = combine(broker_with_transport())
        import_app._apply_awb(by_mark, {})
        awb = by_mark["AGATA"]["awb"]
        self.assertEqual(awb["source"], "invoice")
        self.assertEqual((awb["pieces"], awb["gross_weight"], awb["chargeable_weight"]), (1, 300, 333))
        self.assertEqual(awb["total_awb"], 100.01)
        self.assertNotEqual(round(awb["rate_per_kg"] * awb["chargeable_weight"], 2), 100.01)
        self.assertEqual(xls_totals(by_mark)["TOTAL USD"], 250.01)

    def test_unchanged_prefilled_form_preserves_exact_cost(self):
        for chargeable in (333, None):
            with self.subTest(chargeable=chargeable):
                by_mark = combine(broker_with_transport(gross=333, chargeable=chargeable))
                import_app._apply_awb(by_mark, {})
                awb = by_mark["AGATA"]["awb"]
                submitted = {key: awb[key] for key in ("pieces", "gross_weight", "chargeable_weight", "rate_per_kg")}
                import_app._apply_awb(by_mark, {"AGATA": submitted})
                self.assertEqual(by_mark["AGATA"]["awb"]["total_awb"], 100.01)
                self.assertEqual(by_mark["AGATA"]["awb"]["source"], "invoice")

    def test_adding_missing_weight_preserves_printed_cost(self):
        by_mark = combine(broker_with_transport(gross=None, chargeable=None))
        import_app._apply_awb(by_mark, {"AGATA": {"gross_weight": 333}})
        self.assertEqual(by_mark["AGATA"]["awb"]["total_awb"], 100.01)
        self.assertEqual(by_mark["AGATA"]["awb"]["rate_per_kg"], 0.3003)

    def test_manual_rate_or_chargeable_weight_changes_recalculate(self):
        for submitted, expected in [({"rate_per_kg": 2}, 666),
                                    ({"chargeable_weight": 400}, 120.12)]:
            with self.subTest(submitted=submitted):
                by_mark = combine(broker_with_transport())
                import_app._apply_awb(by_mark, {"AGATA": submitted})
                self.assertEqual(by_mark["AGATA"]["awb"]["total_awb"], expected)

    def test_invoice_total_has_priority_over_additional_awb_document(self):
        doc = {"pieces": 1, "marks": {"AGATA": {"pieces": 1}},
               "gross_weight": 350, "chargeable_weight": 400, "rate_per_kg": 2,
               "other_charges": 100, "total_awb": 900}
        by_mark = combine(broker_with_transport())
        import_app._apply_awb(by_mark, {}, doc)
        self.assertEqual(by_mark["AGATA"]["awb"]["total_awb"], 100.01)
        self.assertEqual(xls_totals(by_mark, doc)["TOTAL AWB"], 100.01)

    def test_partial_transport_and_zero_cost_are_not_lost(self):
        for cost, gross, chargeable in [(100.01, None, None), (None, 300, 333), (0, 300, 333)]:
            with self.subTest(cost=cost, gross=gross, chargeable=chargeable):
                data = broker_with_transport(cost=cost, gross=gross, chargeable=chargeable)
                self.assertEqual(data["transport"]["cost_usd"], cost)
                by_mark = combine(data)
                import_app._apply_awb(by_mark, {})
                self.assertEqual(by_mark["AGATA"]["awb"]["total_awb"], cost)
                if cost == 0:
                    self.assertEqual(xls_totals(by_mark)["TOTAL AWB"], 0)

    @unittest.skipUnless(os.path.isfile(REAL_INVOICE), "нет локального INV-16010.xls")
    def test_real_invoice_and_export_include_printed_transport(self):
        data, _ = import_parser.parse_invoice_file(REAL_INVOICE)
        transport = data["transport"]
        self.assertEqual((transport["cost_usd"], transport["gross_weight"], transport["chargeable_weight"]),
                         (93.52, 28, 28))
        self.assertEqual(data["totals"], {"total_stems": 180, "total_fob": 154.8})
        by_mark = combine(data)
        import_app._apply_awb(by_mark, {})
        self.assertEqual(by_mark["AGATA"]["awb"]["source"], "invoice")
        self.assertEqual(by_mark["AGATA"]["awb"]["total_awb"], 93.52)
        self.assertEqual(xls_totals(by_mark), {"TOTAL FLOWERS FOB USD": 154.8,
                                              "TOTAL CHARGEABLE WEIGHT(Kg)": 28,
                                              "GROSS WEIGHT": 28,
                                              "TOTAL AWB": 93.52, "TOTAL USD": 248.32})


if __name__ == "__main__":
    unittest.main()
