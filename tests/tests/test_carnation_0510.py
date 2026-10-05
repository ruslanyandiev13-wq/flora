"""Сорта и сортность гвоздик сохраняются в итоговом файле без сведения в MIX."""
import copy
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from import_combine import combine_by_mark
from import_app import _edited_invoices
from import_parser import parse_invoice_file


def item(variety="Carnation crimea select", length=70, grade=None, stems=400):
    return {"variety": variety, "length_cm": length, "grade_text": grade,
            "stems": stems, "price": 0.25, "total": stems * 0.25}


def invoice(items, product=None, no_merge=False, template="tessa"):
    return {"filename": "synthetic.pdf", "template": template, "data": {
        "supplier": "TEST FARM", "mark": "AGATA", "boxes": [{
            "farm": "TEST FARM", "product": product, "box_size": 0.25,
            "no_merge": no_merge, "items": items,
        }],
    }}


def combined_items(inv):
    return combine_by_mark([inv])["AGATA"]["boxes"][0]["items"]


class CarnationVarietyTest(unittest.TestCase):
    def test_tessa_description_becomes_variety_and_grade_without_changing_source(self):
        inv = invoice([item()])
        original = copy.deepcopy(inv)
        combined = combine_by_mark([inv])["AGATA"]
        result = combined["boxes"][0]["items"][0]
        self.assertEqual((result["product"], result["variety"], result["grade_text"], result["length_cm"]),
                         ("CARNATION", "CRIMEA", "SELECT", None))
        self.assertEqual((result["stems"], result["price"], result["total"]), (400, 0.25, 100))
        self.assertEqual((combined["total_stems"], combined["total_fob"]), (400, 100))
        self.assertEqual(inv, original)

    def test_distinct_varieties_are_not_mixed_including_moon(self):
        for no_merge in (False, True):
            with self.subTest(no_merge=no_merge):
                inv = invoice([item("Carnation crimea select"), item("Carnation nobbio babel select"),
                               item("Moonlite", None, "FANCY"), item("Moonaqua", None, "FANCY")],
                              product="Carnation", no_merge=no_merge, template="astoria_export")
                self.assertEqual([it["variety"] for it in combined_items(inv)],
                                 ["CRIMEA", "NOBBIO BABEL", "MOONLITE", "MOONAQUA"])

    def test_grade_is_part_of_duplicate_key(self):
        result = combined_items(invoice([
            item("Carnation crimea select", stems=100), item("Carnation crimea fancy", stems=200),
            item("Carnation crimea select", stems=300),
        ]))
        self.assertEqual([(it["variety"], it["grade_text"], it["stems"], it["total"]) for it in result],
                         [("CRIMEA", "SELECT", 400, 100), ("CRIMEA", "FANCY", 200, 50)])

    def test_explicit_grade_wins_and_unknown_suffix_is_not_guessed(self):
        inv = invoice([item("Carnation crimea select", grade="fancy"),
                       item("Carnation crimea special"), item("Crimea", grade="SELECT")],
                      product="CARNATION")
        result = combined_items(inv)
        self.assertEqual((result[0]["variety"], result[0]["grade_text"], result[0]["length_cm"]),
                         ("CRIMEA", "FANCY", None))
        self.assertEqual((result[1]["variety"], result[1]["grade_text"], result[1]["length_cm"]),
                         ("CRIMEA SPECIAL", None, 70))
        self.assertEqual((result[2]["variety"], result[2]["grade_text"], result[2]["length_cm"]),
                         ("CRIMEA", "SELECT", None))

    def test_manual_grade_and_reset_survive_review_to_combined_pipeline(self):
        for value, expected_grade, expected_length in [(60, None, 60), (None, None, None),
                                                      ("FANCY", "FANCY", None)]:
            with self.subTest(manual_grade=value):
                source = invoice([item()])
                original = copy.deepcopy(source)
                pending = {"invoices": [source], "edits": {"1": {"items": {"0": {"grade": value}}}}}
                result = combine_by_mark(_edited_invoices(pending))["AGATA"]["boxes"][0]["items"][0]
                self.assertEqual((result["variety"], result["grade_text"], result["length_cm"]),
                                 ("CRIMEA", expected_grade, expected_length))
                self.assertEqual((result["stems"], result["price"], result["total"]), (400, 0.25, 100))
                self.assertEqual(source, original)
                pending.pop("edits")
                reset = combine_by_mark(_edited_invoices(pending))["AGATA"]["boxes"][0]["items"][0]
                self.assertEqual((reset["variety"], reset["grade_text"], reset["length_cm"]),
                                 ("CRIMEA", "SELECT", None))
                self.assertEqual(source, original)

    def test_ready_mixes_from_document_are_preserved(self):
        result = combined_items(invoice([item("mix", None, "SELECT"),
                                         item("moon mix", None, "FANCY")], product="CARNATION"))
        self.assertEqual([it["variety"] for it in result], ["MIX", "MOON MIX"])

    def test_no_merge_broker_and_astoria_normalize_variety_case(self):
        for template in ("broker_xls", "astoria_export"):
            with self.subTest(template=template):
                inv = invoice([item("Crimea", None, "SELECT"), item("Nobbio Babel", None, "FANCY")],
                              product="CARNATION", no_merge=True, template=template)
                result = combined_items(inv)
                self.assertEqual([(it["variety"], it["grade_text"]) for it in result],
                                 [("CRIMEA", "SELECT"), ("NOBBIO BABEL", "FANCY")])

    def test_other_cultures_are_uppercase_without_carnation_cleanup(self):
        for no_merge in (False, True):
            with self.subTest(no_merge=no_merge):
                inv = invoice([item("Pink fancy", 60)], product="ROSES test", no_merge=no_merge)
                result = combined_items(inv)[0]
                self.assertEqual((result["variety"], result["length_cm"], result["grade_text"]),
                                 ("PINK FANCY", 60, None))

    def test_spray_mix_keeps_product_and_does_not_merge_with_regular_roses(self):
        items = [item(f"Rose {index}", 60, stems=25) for index in range(5)]
        items += [item(f"Spray rose {index}", 60, stems=25) for index in range(5)]
        inv = invoice(items, product="ROSES test")
        original = copy.deepcopy(inv)
        result = combined_items(inv)
        self.assertEqual([(it["product"], it["variety"], it["stems"], it["total"]) for it in result],
                         [("ROSES test", "MIX", 125, 31.25),
                          ("SPRAY ROSES test", "MIX", 125, 31.25)])
        self.assertEqual(inv, original)

    def test_corazon_sp_prefix_keeps_spray_product_after_mix(self):
        inv = invoice([item(f"SPVARIETY{index}", 60, stems=25) for index in range(5)],
                      product="ROSES corazon", template="rosas_corazon")
        result = combined_items(inv)
        self.assertEqual([(it["product"], it["variety"], it["stems"]) for it in result],
                         [("SPRAY ROSES corazon", "MIX", 125)])

    def test_explicit_spray_product_uses_same_length_and_variety_thresholds_as_roses(self):
        for length, count, expected_rows in [(70, 2, 2), (70, 5, 5), (60, 2, 2), (60, 5, 1)]:
            with self.subTest(length=length, count=count):
                inv = invoice([item(f"Variety {index}", length, stems=25) for index in range(count)],
                              product="SPRAY ROSES TESSA")
                result = combined_items(inv)
                self.assertEqual(len(result), expected_rows)
                self.assertTrue(all(it["product"] == "SPRAY ROSES TESSA" for it in result))
                self.assertEqual(sum(it["stems"] for it in result), count * 25)
                self.assertEqual(sum(it["total"] for it in result), count * 6.25)
                self.assertEqual([it["variety"] for it in result],
                                 ["MIX"] if expected_rows == 1 else
                                 [f"VARIETY {index}" for index in range(count)])

    @unittest.skipUnless(os.path.isfile(os.path.join(ROOT, "ТЕСТ3009", "90836820 b.pdf")),
                         "нет локального TESSA 90836820")
    def test_real_tessa_carnation_varieties(self):
        path = os.path.join(ROOT, "ТЕСТ3009", "90836820 b.pdf")
        data, template = parse_invoice_file(path)
        original = copy.deepcopy(data)
        result = combine_by_mark([{"filename": os.path.basename(path), "template": template,
                                   "data": data}])[data["mark"]]
        rows = [it for box in result["boxes"] for it in box["items"]]
        raw = [it for box in data["boxes"] for it in box["items"]]
        self.assertEqual(len(rows), len(raw))
        self.assertTrue(all(it["grade_text"] == "SELECT" and it["length_cm"] is None for it in rows))
        self.assertTrue(all(it["variety"] == it["variety"].upper() and it["variety"] != "MIX" for it in rows))
        self.assertTrue(any(it["variety"] == "CRIMEA" and it["stems"] == 200 for it in rows))
        self.assertEqual(result["total_stems"], sum(it["stems"] for it in raw))
        self.assertAlmostEqual(result["total_fob"], sum(it["total"] for it in raw))
        self.assertEqual(data, original)
