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


ASTORIA_BESST = os.path.join(ROOT, "ТЕСТ3009", "001E-20260928 BESST.xls")
ASTORIA_MIXED = os.path.join(ROOT, "samples_files_0109", "001C-20260828 (1).xls")


class AstoriaHandlerMarkTest(unittest.TestCase):
    """У Astoria метка - в колонке Handler коробки, CLIENT = счёт клиента
    (SIRI). Правка закупщика 2026-09-30: инвойс BESST уходил в поставку SIRI."""

    @unittest.skipUnless(os.path.exists(ASTORIA_BESST), "нет инвойса 001E-20260928")
    def test_mark_from_handler(self):
        data, _ = import_parser.parse_invoice_file(ASTORIA_BESST)
        self.assertEqual(data["mark"], "BESST")

    @unittest.skipUnless(os.path.exists(ASTORIA_MIXED), "нет инвойса 001C-20260828")
    def test_mixed_invoice_split_by_mark(self):
        from import_combine import combine_by_mark, split_invoice_by_mark
        data, template = import_parser.parse_invoice_file(ASTORIA_MIXED)
        inv = {"filename": "x", "data": data, "template": template}
        self.assertEqual([p["data"]["mark"] for p in split_invoice_by_mark(inv)], ["SIRI", "BESST", "AGATA"])
        self.assertEqual({m: len(i["boxes"]) for m, i in combine_by_mark([inv]).items()},
                         {"SIRI": 4, "BESST": 4, "AGATA": 2})


@unittest.skipUnless(os.path.isdir(VIKA_DIR), "нет документов партии VIKA")
class VikaBatchTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._old_path = db.DB_PATH
        cls._tmp = tempfile.TemporaryDirectory()
        db.DB_PATH = os.path.join(cls._tmp.name, "test.db")
        db.init_db()
        cls.batch_id = db.create_delivery_batch("VIKA", "2026-09-28T10:00:00")
        for path in sorted(glob.glob(os.path.join(VIKA_DIR, "*"))):
            name = os.path.basename(path)
            if name.endswith((".jpeg", ".jpg")) or "VADIM" in name:
                continue
            hawb = import_parser.parse_forwarder_hawb_pdf(path) if name.lower().endswith(".pdf") else None
            if hawb:
                db.save_delivery_doc(cls.batch_id, hawb["mark"], "hawb", hawb["hawb"], name,
                                     "forwarder_hawb", json.dumps(hawb), "", "test")
                continue
            data, template = import_parser.parse_invoice_file(path)
            db.save_delivery_doc(cls.batch_id, data["mark"], "invoice",
                                 f"{template}:{data.get('invoice_no')}", name, template,
                                 json.dumps(data), "", "test")
        cls.model = deliveries.build(cls.batch_id)

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

    def test_batch_label_and_separate_batches(self):
        self.assertEqual(self.model["label"], "VIKA · 28.09.2026")
        # Вторая поставка той же метки - свои документы, первая не меняется.
        other = db.create_delivery_batch("VIKA", "2026-10-05T09:00:00")
        self.assertEqual(deliveries.build(other)["docs"], [])
        self.assertEqual(len(deliveries.build(self.batch_id)["docs"]), 23)
        same_day = db.create_delivery_batch("VIKA", "2026-09-28T18:00:00")
        batches = db.get_delivery_batches()
        batches.append({"id": same_day, "mark": "VIKA", "created_at": "2026-09-28T18:00:00"})
        self.assertEqual(deliveries.batch_label(batches[-1], batches), "VIKA · 28.09.2026 (2)")

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
