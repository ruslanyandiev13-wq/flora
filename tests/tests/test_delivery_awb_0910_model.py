"""House AWB через AMS: деньги, дата поставки и шапка отдельного самолёта.

Синтетические инвойсы и временная БД; клиентские PDF для тестов не нужны.
"""
import datetime
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

import xlrd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import db  # noqa: E402
import deliveries  # noqa: E402


class DeliveryHouseAwbTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        db_path = patch.object(db, "DB_PATH", os.path.join(tmp.name, "test.db"))
        db_path.start()
        self.addCleanup(db_path.stop)
        db.init_db()

    def make_batch(self, mark="DAMIR", **overrides):
        polina = mark == "POLINA"
        sizes = [0.5] * 53 + [0.25] * 12 if polina else [0.5] * 10
        data = {
            "mark": mark, "mawb": "416-5236 2811", "hawb": "0003" if polina else "0001",
            "flight_date": "2026-10-07", "airport": "AMS", "origin_country": "ecuador",
            "destination": "AMSTERDAM HOLLAND", "forwarder": "SAFTEC S.A.",
            "airline": "NATIONAL AIR CARGO / NCR", "pieces": len(sizes),
            "gross_weight": 1211 if polina else 278, "chargeable_weight": 1273 if polina else 278,
            "rate_per_kg": 3.45, "other_charges": 20,
            "total_awb": 4411.85 if polina else 979.10,
            "master_total_awb": 5410.95, "master_houses_total_awb": 5390.95,
            "master_awb_difference": 20,
            "growers": [{"name": "MATIZ", "full_boxes": sum(sizes)}],
        }
        data.update(overrides)
        batch_id = db.create_delivery_batch(mark, "2026-10-09T10:00:00")
        self.save_hawb(batch_id, data)
        invoice = {
            "mark": mark, "supplier": "MATIZ", "invoice_no": f"TEST-{mark}",
            "invoice_date": "2026-10-06",
            "boxes": [{"box_size": size, "items": [
                {"variety": "EXPLORER", "length_cm": 60, "stems": 100, "price": 0.4, "total": 40},
            ]} for size in sizes],
        }
        db.save_delivery_doc(batch_id, mark, "invoice", f"TEST-{mark}", "synthetic.json",
                             "matiz_roses", json.dumps(invoice), "", "test")
        return batch_id

    @staticmethod
    def save_hawb(batch_id, data):
        db.save_delivery_doc(batch_id, data["mark"], "hawb", f"{data['mawb']}:{data['hawb']}",
                             "synthetic-house.json", "forwarder_hawb", json.dumps(data), "", "test")

    @staticmethod
    def workbook(model):
        delivery = model["deliveries"][0]
        hawb = delivery["hawbs"][0]
        content = deliveries.invoice_total_xls(deliveries.plane_delivery(delivery, hawb),
                                                consignee=hawb["plane_label"]).getvalue()
        return xlrd.open_workbook(file_contents=content)

    def test_each_house_keeps_its_amount_and_ams_header(self):
        for mark, total, weight, gross, pieces in (
                ("DAMIR", 979.10, 278, 278, 10), ("POLINA", 4411.85, 1273, 1211, 65)):
            with self.subTest(mark=mark):
                model = deliveries.build(self.make_batch(mark))
                delivery, hawb = model["deliveries"][0], model["hawbs"][0]
                self.assertEqual((delivery["arrival"], delivery["id"]),
                                 ("2026-10-14", "2026-10-14_AMS"))
                self.assertEqual((hawb["rate"], hawb["other_charges"], hawb["cost"]), (3.45, 20, total))
                self.assertEqual(hawb["cost_source"], "document")
                self.assertEqual(delivery["cost"], total)
                self.assertEqual(delivery["boxes_found"], pieces)
                self.assertEqual(hawb["plane_label"], f"{mark} AMS")
                self.assertEqual(deliveries.invoice_total_name(delivery, hawb),
                                 f"Invoice total 14.10 {mark} AMS (AI)")
                wb = self.workbook(model)
                ws = wb.sheet_by_name("factura")
                self.assertEqual(ws.cell_value(1, 1), f"{mark} AMS")
                self.assertEqual(ws.cell_value(1, 16), "07.10.2026")
                self.assertEqual(ws.cell_value(2, 1), "AMSTERDAM HOLLAND")
                self.assertEqual(ws.cell_value(2, 16), "SAFTEC S.A.")
                self.assertEqual(ws.cell_value(4, 16), "NATIONAL AIR CARGO / NCR")
                self.assertEqual(ws.cell_value(5, 16), "416-5236 2811")
                self.assertEqual(ws.cell_value(6, 16), hawb["data"]["hawb"])
                self.assertEqual(xlrd.xldate_as_datetime(ws.cell_value(7, 7), wb.datemode).date(),
                                 datetime.date(2026, 10, 14))
                totals = {ws.cell_value(r, 9): r for r in range(ws.nrows) if ws.cell_value(r, 9)}
                self.assertEqual(ws.cell_value(totals["TOTAL AWB"], 19), total)
                self.assertEqual(ws.cell_value(totals["TOTAL CHARGEABLE WEIGHT(Kg)"], 13), weight)
                self.assertEqual(ws.cell_value(totals["GROSS WEIGHT"], 13), gross)

    def test_printed_total_wins_over_rate_and_master_difference(self):
        model = deliveries.build(self.make_batch(total_awb=950, rate_per_kg=None))
        hawb = model["hawbs"][0]
        self.assertEqual(hawb["rate"], 8.1)
        self.assertEqual(hawb["cost"], 950)
        self.assertEqual(hawb["data"]["master_awb_difference"], 20)

    def test_without_printed_total_calculates_printed_rate_plus_charges(self):
        hawb = deliveries.build(self.make_batch(total_awb=None))["hawbs"][0]
        self.assertEqual((hawb["rate"], hawb["cost"], hawb["cost_source"]),
                         (3.45, 979.10, "calculated"))

    def test_zero_printed_values_are_preserved(self):
        hawb = deliveries.build(self.make_batch(total_awb=0, rate_per_kg=0))["hawbs"][0]
        self.assertEqual((hawb["rate"], hawb["cost"], hawb["cost_source"]), (0, 0, "document"))

    def test_legacy_moscow_keeps_date_id_rate_and_headers(self):
        model = deliveries.build(self.make_batch(
            airport="SVO", rate_per_kg=None, total_awb=None, other_charges=None,
            destination=None, forwarder=None))
        delivery, hawb = model["deliveries"][0], model["hawbs"][0]
        self.assertEqual(delivery["id"], "2026-10-11")
        self.assertEqual((hawb["rate"], hawb["cost"]), (8.1, 2251.80))
        self.assertEqual(hawb["plane_label"], "DAMIR MOS")
        ws = self.workbook(model).sheet_by_name("factura")
        self.assertEqual(ws.cell_value(2, 1), "Russia")
        self.assertEqual(ws.cell_value(2, 16), "IFC")

    def test_different_routes_same_date_stay_separate(self):
        batch_id = self.make_batch(airport="SVO")
        data = dict(deliveries.build(batch_id)["hawbs"][0]["data"], airport="LED", hawb="0002")
        self.save_hawb(batch_id, data)
        model = deliveries.build(batch_id)
        self.assertEqual([d["id"] for d in model["deliveries"]],
                         ["2026-10-11", "2026-10-11_LED"])
        self.assertEqual([d["hawbs"][0]["plane_label"] for d in model["deliveries"]],
                         ["DAMIR MOS", "DAMIR LED"])

    def make_ambiguous_model(self, two_houses):
        db.upsert_grower_alias("EC BLOOMS", "STAR ROSES")
        batch_id = db.create_delivery_batch("DAMIR", "2026-10-09T10:00:00")
        growers = [{"name": name, "full_boxes": 0.5}
                   for name in ("EC BLOOMS S.A.S.", "EL CAMPANARIO")]
        for index, lines in enumerate(([g] for g in growers) if two_houses else [growers]):
            self.save_hawb(batch_id, {
                "mark": "DAMIR", "mawb": "416-5236 2811", "hawb": f"000{index + 1}",
                "flight_date": "2026-10-07", "airport": "AMS", "pieces": len(lines),
                "chargeable_weight": 50, "growers": lines,
            })
        invoice = {"mark": "DAMIR", "supplier": "STAR ROSES", "invoice_date": "2026-10-06",
                   "boxes": [{"box_size": 0.5, "items": [
                       {"variety": variety, "length_cm": 60, "stems": 100, "price": 0.4, "total": 40},
                   ]} for variety in ("EXPLORER", "FREEDOM")]}
        db.save_delivery_doc(batch_id, "DAMIR", "invoice", "ambiguity-test", "synthetic.json",
                             "synthetic", json.dumps(invoice), "", "test")
        return deliveries.build(batch_id)

    def test_aliases_in_same_house_do_not_warn_after_both_boxes_are_assigned(self):
        model = self.make_ambiguous_model(two_houses=False)
        self.assertEqual(len(model["hawbs"][0]["boxes"]), 2)
        self.assertEqual(model["missing"], [])
        self.assertEqual(model["ambiguous"], [])
        self.assertEqual(model["deliveries"][0]["ambiguous"], [])

    def test_boxes_on_different_houses_still_warn_about_ambiguous_assignment(self):
        model = self.make_ambiguous_model(two_houses=True)
        self.assertEqual(model["missing"], [])
        self.assertEqual(len(model["ambiguous"]), 1)
        pair = model["ambiguous"][0][2]["ambiguous"][0]
        self.assertNotEqual(pair[0]["hawb_id"], pair[1]["hawb_id"])


if __name__ == "__main__":
    unittest.main()
