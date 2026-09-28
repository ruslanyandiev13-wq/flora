"""
Раздел «Поставки»: раскладка коробок из инвойсов по фактическим рейсам по
HAWB форвардера. Эталон - партия VIKA 21-25.09 (папка "Тестируем Москву"),
разобранная вручную 2026-09-28: 5 HAWB, 43 места, 3 доставки, всё сходится.

Реальные документы в репозитории не хранятся - без них тесты на партию
пропускаются, юнит-тесты правил работают всегда.
"""
import glob
import json
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import db  # noqa: E402
import deliveries  # noqa: E402
import import_parser  # noqa: E402

VIKA_DIR = os.path.join(ROOT, "Тестируем Москву")


class GrowerNamesTest(unittest.TestCase):
    def setUp(self):
        self.growers = deliveries.Growers([{"alias": a, "grower": g}
                                           for a, g in db.DEFAULT_GROWER_ALIASES])

    def test_tessa_codes(self):
        self.assertEqual(self.growers.resolve("TESSA-PS2"), "POSITANO")
        self.assertEqual(self.growers.resolve("TESSA-P"), "POSITANO")
        # Одиночная буква кода - не форма юрлица "S.A.".
        self.assertEqual(self.growers.resolve("TESSA-S"), "SOLERA")

    def test_hawb_spellings(self):
        self.assertEqual(self.growers.resolve("POSITANO FARMS S.A.S."), "POSITANO")
        self.assertEqual(self.growers.resolve("INVERSIONES PONTE"), "PONTE TRESA")
        self.assertEqual(self.growers.resolve("EL CAMPANARIO DE"), "STAR ROSES")
        self.assertEqual(self.growers.resolve("QUALISA SERVICE S.A."), "QUALITY SERVICE")


class SubsetTest(unittest.TestCase):
    def test_earliest_exact(self):
        # 0.5 = 8/16: берём самую раннюю коробку 8, а не 4+4.
        self.assertEqual(deliveries._earliest_subset([8, 4, 4], 8), [0])
        self.assertEqual(deliveries._earliest_subset([8, 4, 4], 4), [1])
        self.assertIsNone(deliveries._earliest_subset([8, 8], 4))

    def test_best_when_not_enough(self):
        self.assertEqual(deliveries._best_subset([8], 16), [0])
        self.assertEqual(deliveries._best_subset([], 8), [])


@unittest.skipUnless(os.path.isdir(VIKA_DIR), "нет документов партии VIKA")
class VikaBatchTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._old_path = db.DB_PATH
        cls._tmp = tempfile.TemporaryDirectory()
        db.DB_PATH = os.path.join(cls._tmp.name, "test.db")
        db.init_db()
        for path in sorted(glob.glob(os.path.join(VIKA_DIR, "*"))):
            name = os.path.basename(path)
            if name.endswith((".jpeg", ".jpg")) or "VADIM" in name:
                continue
            hawb = import_parser.parse_forwarder_hawb_pdf(path) if name.lower().endswith(".pdf") else None
            if hawb:
                db.save_delivery_doc(hawb["mark"], "hawb", hawb["hawb"], name, "forwarder_hawb",
                                     json.dumps(hawb), "", "test")
                continue
            data, template = import_parser.parse_invoice_file(path)
            db.save_delivery_doc(data["mark"], "invoice", f"{template}:{data.get('invoice_no')}",
                                 name, template, json.dumps(data), "", "test")
        cls.model = deliveries.build("VIKA")

    @classmethod
    def tearDownClass(cls):
        db.DB_PATH = cls._old_path
        cls._tmp.cleanup()

    def test_three_deliveries_all_boxes_found(self):
        got = [(d["id"], d["pieces"], d["boxes_found"], d["cost"]) for d in self.model["deliveries"]]
        self.assertEqual(got, [("2026-09-22_SVO", 13, 13, 2640.6),
                               ("2026-09-23_VKO", 7, 7, 996.3),
                               ("2026-09-25_SVO", 23, 23, 3815.1)])
        self.assertEqual(self.model["missing"], [])
        self.assertEqual(self.model["not_shipped"], [])

    def test_invoice_split_across_three_flights(self):
        mawb = {h["id"]: h["data"]["mawb"] for h in self.model["hawbs"]}
        flights = {(b["box"]["farm_code"], mawb[b["hawb_id"]]) for b in self.model["boxes"]
                   if b["invoice"]["filename"].startswith("90832511")}
        self.assertEqual(flights, {("TESSA-E2", "235-7841 9821"), ("TESSA-P", "157-0401 2411"),
                                   ("TESSA-R1", "157-0401 2400"), ("TESSA-R2", "157-0401 2400")})

    def test_broker_boxes_matched(self):
        growers = {b["grower"]: mawb for b in self.model["boxes"]
                   for mawb in [next(h["data"]["mawb"] for h in self.model["hawbs"] if h["id"] == b["hawb_id"])]}
        self.assertEqual(growers["QUALITY SERVICE"], "157-0401 2433")
        self.assertEqual(growers["ECOROSES"], "157-0401 2466")


if __name__ == "__main__":
    unittest.main()
