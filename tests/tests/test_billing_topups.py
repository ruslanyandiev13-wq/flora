"""Минимум ручного пополнения проверяется в форме, роуте и слое БД."""
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
from billing import db as billing_db
from billing import routes


class ManualTopupTest(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        database = patch.object(db, "DB_PATH", os.path.join(temporary.name, "billing-test.db"))
        database.start()
        self.addCleanup(database.stop)
        app = Flask(__name__, template_folder=os.path.join(ROOT, "templates"))
        app.config.update(TESTING=True, SECRET_KEY="topup-test")
        # Рендерим настоящую форму, не подключая остальное приложение и его БД.
        app.jinja_loader = ChoiceLoader([
            DictLoader({"base.html": "{% block content %}{% endblock %}"}), app.jinja_loader,
        ])
        routes.init_app(app)
        self.client = app.test_client()
        with self.client.session_transaction() as session:
            session.update(user_id=1, username="test-admin", role="admin")
        self.initial_balance = billing_db.get_balance()["current_balance"]

    def test_form_displays_and_enforces_minimum_100(self):
        response = self.client.get("/billing/admin/topup")
        self.assertEqual(response.status_code, 200)
        self.assertIn("минимум 100", response.get_data(as_text=True))
        self.assertIn('name="tokens" min="100"', response.get_data(as_text=True))

    def test_database_accepts_100_and_400(self):
        expected = self.initial_balance
        for tokens in (100, 400):
            with self.subTest(tokens=tokens):
                expected += tokens
                self.assertEqual(billing_db.manual_topup(tokens, "test invoice"), expected)
                self.assertEqual(billing_db.get_balance()["current_balance"], expected)
                latest = max(billing_db.get_ledger(), key=lambda row: row["id"])
                self.assertEqual((latest["delta"], latest["balance_after"]), (tokens, expected))
                self.assertEqual(latest["reason"], "manual_adjustment: test invoice")

    def test_database_rejects_99_without_changing_balance_or_history(self):
        with self.assertRaisesRegex(ValueError, "100 токенов"):
            billing_db.manual_topup(99, "test invoice")
        self.assertEqual(billing_db.get_balance()["current_balance"], self.initial_balance)
        self.assertEqual(billing_db.get_ledger(), [])

    def test_admin_route_accepts_100_and_400(self):
        expected = self.initial_balance
        for tokens in (100, 400):
            with self.subTest(tokens=tokens):
                response = self.client.post("/billing/admin/topup", data={
                    "tokens": str(tokens), "comment": "test invoice",
                })
                expected += tokens
                self.assertEqual(response.status_code, 302)
                self.assertEqual(billing_db.get_balance()["current_balance"], expected)
                latest = max(billing_db.get_ledger(), key=lambda row: row["id"])
                self.assertEqual(latest["delta"], tokens)

    def test_admin_route_rejects_99_without_changing_balance_or_history(self):
        response = self.client.post("/billing/admin/topup", data={
            "tokens": "99", "comment": "test invoice",
        })
        self.assertEqual(response.status_code, 302)
        self.assertEqual(billing_db.get_balance()["current_balance"], self.initial_balance)
        self.assertEqual(billing_db.get_ledger(), [])
        with self.client.session_transaction() as session:
            self.assertIn(("error", "Минимальный объём докупки - 100 токенов"), session["_flashes"])
