"""Один PDF с house AWB загружается по меткам, повтор не списывает токены."""
import copy
import io
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

from flask import Flask
from jinja2 import ChoiceLoader, DictLoader

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import db
import deliveries_app
import timeutil
from billing import db as billing_db
from billing.cost_calc import COST_DELIVERY_HAWB

SAMPLE = os.path.expanduser("~/Downloads/AWB-2_261008_153814.pdf")


def houses(mawb="416-5236 2811"):
    return [dict(mark=mark, hawb=number, doc_key=f"{mawb.replace(' ', '')}:{number}",
                 mawb=mawb, airport="AMS", destination="AMSTERDAM HOLLAND",
                 flight_date="2026-10-07", origin_country="ecuador", origin="UIO",
                 airline="NATIONAL AIR CARGO / NCR", forwarder="SAFTEC S.A.",
                 pieces=pieces, gross_weight=gross, chargeable_weight=weight,
                 rate_per_kg=3.45, total_awb=cost, other_charges=20,
                 total_full=full, growers=[{"name": "TEST FARM", "full_boxes": full}],
                 master_awb_difference=20)
            for mark, number, pieces, gross, weight, cost, full in [
                ("DAMIR", "0001", 10, 278, 278, 979.10, 5),
                ("POLINA", "0003", 65, 1211, 1273, 4411.85, 29.5)]]


class DeliveryHouseUploadTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        state = patch.object(db, "DB_PATH", os.path.join(self.tmp.name, "app.db"))
        state.start()
        self.addCleanup(state.stop)
        upload_dir = patch.object(deliveries_app, "UPLOAD_DIR", self.tmp.name)
        upload_dir.start()
        self.addCleanup(upload_dir.stop)
        errors = patch.object(deliveries_app.support_db, "log_error")
        self.log_error = errors.start()
        self.addCleanup(errors.stop)
        db.init_db()
        billing_db.init_db()
        self.balance = billing_db.get_balance()["current_balance"]
        app = Flask(__name__, template_folder=os.path.join(ROOT, "templates"))
        app.config.update(TESTING=True, SECRET_KEY="house-upload-test")
        app.jinja_env.filters["dt"] = timeutil.format_dt
        app.jinja_loader = ChoiceLoader([
            DictLoader({"base.html": "{% block content %}{% endblock %}"}), app.jinja_loader])
        app.add_url_rule("/dictionaries", "dictionaries", lambda: "")
        app.register_blueprint(deliveries_app.deliveries_bp)
        self.client = app.test_client()
        with self.client.session_transaction() as session:
            session.update(user_id=1, username="test-admin", role="admin")

    def upload(self, parsed=None, content=b"synthetic", name="airwaybill.pdf"):
        with patch.object(deliveries_app, "parse_delivery_hawbs_pdf", return_value=parsed), \
                patch.object(deliveries_app, "_page_count", return_value=3):
            return self.client.post("/deliveries/upload", data={"docs": (io.BytesIO(content), name)})

    def documents(self):
        return [doc for batch in db.get_delivery_batches() for doc in db.get_delivery_docs(batch["id"])]

    def test_multi_house_upload_once_and_repeat_without_charge(self):
        response = self.upload(houses())
        self.assertEqual(response.status_code, 302)
        first = self.documents()
        self.assertEqual({d["mark"] for d in first}, {"DAMIR", "POLINA"})
        self.assertEqual(len(first), 2)
        self.assertEqual(billing_db.get_balance()["current_balance"], self.balance - COST_DELIVERY_HAWB)
        self.assertEqual(len(billing_db.get_ledger()), 1)
        self.assertEqual(billing_db.get_history()[0]["awb_count"], 2)
        with self.client.session_transaction() as session:
            self.assertTrue(any(category == "warning" and "20.00" in text
                                for category, text in session["_flashes"]))
        self.upload(houses(), name="renamed.pdf")
        self.assertEqual({d["id"] for d in self.documents()}, {d["id"] for d in first})
        self.assertEqual(len(billing_db.get_ledger()), 1)
        self.assertEqual(billing_db.get_balance()["current_balance"], self.balance - COST_DELIVERY_HAWB)
        self.log_error.assert_not_called()

    def test_same_short_house_number_on_another_master_is_new_document(self):
        self.upload(houses())
        self.upload(houses("416-1234 5678"))
        self.assertEqual(len(self.documents()), 4)
        self.assertEqual(len(billing_db.get_ledger()), 2)

    def test_invalid_house_rejects_whole_file_before_any_save_or_charge(self):
        invalid = copy.deepcopy(houses())
        invalid[1]["growers"] = []
        self.upload(invalid)
        self.assertEqual(self.documents(), [])
        self.assertEqual(db.get_delivery_batches(), [])
        self.assertEqual(billing_db.get_ledger(), [])
        self.log_error.assert_called_once()

    def test_master_without_house_remains_rejected_without_charge(self):
        with patch.object(deliveries_app, "parse_awb_pdf", return_value={"awb_no": "416-5236 2811"}):
            self.upload([])
        self.assertEqual(self.documents(), [])
        self.assertEqual(billing_db.get_ledger(), [])
        self.log_error.assert_called_once()

    @unittest.skipUnless(os.path.isfile(SAMPLE), "нет пользовательской AWB от 08.10")
    def test_real_pdf_uploads_both_houses_and_renders_correct_cost(self):
        with open(SAMPLE, "rb") as source:
            content = source.read()
        response = self.client.post("/deliveries/upload", data={"docs": (io.BytesIO(content), "awb.pdf")})
        self.assertEqual(response.status_code, 302)
        docs = self.documents()
        self.assertEqual(len(docs), 2)
        data = {d["mark"]: json.loads(d["data"]) for d in docs}
        self.assertEqual(data["DAMIR"]["total_awb"], 979.1)
        self.assertEqual(data["POLINA"]["total_awb"], 4411.85)
        self.assertEqual(len(billing_db.get_ledger()), 1)
        damir = next(d for d in docs if d["mark"] == "DAMIR")
        page = self.client.get(f"/deliveries/b/{damir['batch_id']}")
        self.assertEqual(page.status_code, 200)
        html = page.get_data(as_text=True)
        self.assertIn("DAMIR AMS", html)
        self.assertIn("итог по накладной 979.10", html)
        self.assertIn("включая сборы 20.00", html)
        self.assertNotIn("кг × 3.45 $ = 979.10", html)
        self.log_error.assert_not_called()


if __name__ == "__main__":
    unittest.main()
