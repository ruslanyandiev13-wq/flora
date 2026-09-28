"""
Регрессия на правки бухгалтера от 2026-09-15 (файлы в tests/test_holland):

1. Тележка, перенесённая на следующую страницу паклиста, не должна
   разваливаться на две - из-за этого съезжала вся нумерация в файле для 1С.
2. Имя плантации, напечатанное MH Flowers в конце колонки Omschrijving, не
   должно попадать в название сорта.
3. Приведение названий к справочнику ассортимента не должно съедать ростовку
   и калибр ("Li Ot Zambesi 5+", "Cymb T Toledo Decorum 80cm").

Тесты на PDF пропускаются, если реальных паклистов нет рядом (они не
хранятся в репозитории - чужие коммерческие документы).
"""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import assortment
import pdf_parser

SIRI_PDF = os.path.join(ROOT, "tests", "test_holland",
                        "08-09-26 17-09-38-202604952   SIRI.PDF")


class StripKwekerTest(unittest.TestCase):
    def test_removes_grower_repeated_in_description(self):
        self.assertEqual(
            pdf_parser._strip_kweker_from_omschrijving("R Tr Green Glow Flora Ola", "Flora Ola Ltd."),
            "R Tr Green Glow")
        self.assertEqual(
            pdf_parser._strip_kweker_from_omschrijving("Glad Gr Careless Apg Van Den Berg",
                                                       "A.P.G. van den Berg"),
            "Glad Gr Careless")

    def test_splits_glued_grower_word(self):
        self.assertEqual(
            pdf_parser._strip_kweker_from_omschrijving("Cymb T Mrs Sugar Lee Decorum 80cmKwekerij",
                                                       "Koningshof"),
            "Cymb T Mrs Sugar Lee Decorum 80cm")

    def test_keeps_variety_that_is_not_the_grower(self):
        # "Groove" - и сорт, и название плантации: без совпадения с колонкой
        # Kweker этой строки резать нельзя.
        self.assertEqual(
            pdf_parser._strip_kweker_from_omschrijving("R Tr Groove", "Braam Roses"),
            "R Tr Groove")


class CanonicalNameTest(unittest.TestCase):
    lookup = {
        "names": {"r tr summer dance": "R Tr Summer Dance",
                  "r tr yellow babe flora ola": "R Tr Yellow Babe Flora Ola",
                  "li ot zambesi": "Li Ot Zambesi",
                  "cymb t toledo": "Cymb T Toledo"},
        "aliases": {"blushing bride 4-6": "Serruria Blushing Bride"},
    }

    def name(self, printed, fallback=None):
        return assortment.canonical_name(printed, fallback, self.lookup)

    def test_strips_grower_tail_when_short_name_is_in_assortment(self):
        self.assertEqual(self.name("R Tr Summer Dance Water"), "R Tr Summer Dance")

    def test_keeps_name_that_includes_the_grower(self):
        self.assertEqual(self.name("R Tr Yellow Babe Flora Ola", "R Tr Yellow Babe"),
                         "R Tr Yellow Babe Flora Ola")

    def test_keeps_size_and_grade_suffixes(self):
        self.assertEqual(self.name("Li Ot Zambesi 5+"), "Li Ot Zambesi 5+")
        self.assertEqual(self.name("Cymb T Toledo Decorum 80cm"), "Cymb T Toledo Decorum 80cm")

    def test_alias_wins(self):
        self.assertEqual(self.name("Blushing Bride 4-6"), "Serruria Blushing Bride")

    def test_unknown_name_falls_back_to_cleaned(self):
        self.assertEqual(self.name("R Tr Green Glow Flora Ola", "R Tr Green Glow"),
                         "R Tr Green Glow")


@unittest.skipUnless(os.path.exists(SIRI_PDF), "нет реального паклиста SIRI")
class SiriPaklijstTest(unittest.TestCase):
    """Эталон - ручной файл бухгалтера "Голландия MH Siri & Chrys на 13.09"."""

    @classmethod
    def setUpClass(cls):
        cls.data, _ = pdf_parser.parse_invoice_pdf(SIRI_PDF)

    def test_trolley_is_not_split_on_page_break(self):
        self.assertEqual(len(self.data["boxes"]), 53)
        numbers = [b["box_no"] for b in self.data["boxes"]]
        self.assertEqual(len(numbers), len(set(numbers)), "номера тележек не должны повторяться")

    def test_trolley_10_keeps_all_five_positions(self):
        box = next(b for b in self.data["boxes"] if b["box_no"] == 10)
        self.assertEqual([it["omschrijving"] for it in box["items"]],
                         ["R Tr Green Glow", "R Tr Mix Novelties", "R Tr Mimi Eden",
                          "Glad Gr Careless", "Chr S Rossi Cream"])

    def test_totals(self):
        self.assertEqual(self.data["totals"]["subtotaal"], 9801.01)
        # 7731 стебель в SIRI + 3120 в CHRYS = 10851 - столько же, сколько в
        # ручном файле бухгалтера за 13.09.
        self.assertEqual(sum(it["aantal"] for b in self.data["boxes"] for it in b["items"]), 7731)


if __name__ == "__main__":
    unittest.main()
