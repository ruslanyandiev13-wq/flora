"""Master + house AWB SAFTEC: раздельные метки, печатные суммы и номера."""
import os
import sys
import unittest
from unittest.mock import patch

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import import_parser
import pdfplumber


MASTER = """UIO
416-5236 2811 416-5236 2811
Shipper's Name and Address
SAFTEC S.A. AIR WAYBILL
NATIONAL AIR CARGO / NCR
Airport of Destination Request Flight/Date Amount of Insurance
AMSTERDAM HOLLAND NIL
Handling Information
DAMIR - IRIS FLOWERS / TRUCK
0.50 = FARM A
4.50 = FARM B
BXS: 5.00 PCS: 10
POLINA - IRIS FLOWERS / TRUCK
29.50 = FARM C
BXS: 29.50 PCS: 65
No Of Gross Kg Rate Class Chargeable Rate
75 1489 K 1551 3.45 5,350.95 FRESH CUT FLOWERS
TOTAL IN FULL: 34.50
Prepaid Weight Charge Collect AWC: 60.00
Total Prepaid Total Collect
5,410.95 PREPAID
OCT-07-2026 QUITO ECUADOR
"""


def house(mark="DAMIR", number="0001", mawb="416-5236 2811"):
    is_damir = mark == "DAMIR"
    growers = "0.50 = FARM A\n4.50 = FARM B" if is_damir else "29.50 = FARM C"
    weights = "10 278 K 278 3.45 959.10" if is_damir else "65 1211 K 1273 3.45 4,391.85"
    total_full = "5.00" if is_damir else "29.50"
    total_cost = "979.10" if is_damir else "4,411.85"
    return f"""UIO
{mawb} {number}
Shipper's Name and Address
SAFTEC S.A. AIR WAYBILL
NATIONAL AIR CARGO / NCR
BOX LABEL: {mark}
Airport of Destination Request Flight/Date Amount of Insurance
AMSTERDAM HOLLAND NIL
Handling Information
{growers}
No Of Gross Kg Rate Class Chargeable Rate
{weights} FRESH FLOWERS
TOTAL IN FULL: {total_full}
Prepaid Weight Charge Collect AWC: 20.00
Total Prepaid Total Collect
{total_cost} PREPAID
OCT-07-2026 QUITO ECUADOR
"""


class TextPage:
    def __init__(self, text):
        self.text = text

    def extract_text(self):
        return self.text


class TextPdf:
    def __init__(self, texts):
        self.pages = [TextPage(text) for text in texts]

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass


def parse_pages(*texts):
    with patch.object(import_parser.pdfplumber, "open", return_value=TextPdf(texts)):
        return import_parser.parse_delivery_hawbs_pdf("synthetic.pdf")


class DeliveryAwbParserTest(unittest.TestCase):
    def test_master_creates_only_two_house_shipments_with_own_financials(self):
        houses = parse_pages(MASTER, house(), house("POLINA", "0003"))
        self.assertEqual([h["mark"] for h in houses], ["DAMIR", "POLINA"])
        self.assertEqual([h["hawb"] for h in houses], ["0001", "0003"])
        self.assertEqual([h["doc_key"] for h in houses], ["416-52362811:0001", "416-52362811:0003"])
        self.assertEqual([(h["pieces"], h["gross_weight"], h["chargeable_weight"], h["total_full"])
                          for h in houses], [(10, 278, 278, 5), (65, 1211, 1273, 29.5)])
        self.assertEqual([(h["rate_per_kg"], h["weight_charge"], h["other_charges"], h["total_awb"])
                          for h in houses], [(3.45, 959.1, 20, 979.1), (3.45, 4391.85, 20, 4411.85)])
        self.assertEqual(houses[0]["growers"], [{"name": "FARM A", "full_boxes": 0.5},
                                              {"name": "FARM B", "full_boxes": 4.5}])
        for h in houses:
            self.assertEqual((h["origin"], h["origin_country"], h["airport"], h["destination"]),
                             ("UIO", "ecuador", "AMS", "AMSTERDAM HOLLAND"))
            self.assertEqual((h["flight_date"], h["airline"], h["forwarder"]),
                             ("2026-10-07", "NATIONAL AIR CARGO / NCR", "SAFTEC S.A."))
            self.assertEqual((h["master_total_awb"], h["master_houses_total_awb"], h["master_awb_difference"]),
                             (5410.95, 5390.95, 20))

    def test_master_alone_and_invoice_are_not_house_shipments(self):
        self.assertEqual(parse_pages(MASTER), [])
        self.assertEqual(parse_pages(MASTER + "416-5236 2811\nDestination\n"), [])
        self.assertEqual(parse_pages("INVOICE 123\nTOTAL 100\nBOX LABEL: DAMIR"), [])

    def test_single_house_page_and_number_scope(self):
        first = parse_pages(house())[0]
        second = parse_pages(house(mawb="416-5236 9999"))[0]
        self.assertEqual(first["hawb"], second["hawb"])
        self.assertNotEqual(first["doc_key"], second["doc_key"])
        self.assertIsNone(first["master_total_awb"])
        self.assertNotIn("master_awb_difference", first)

    def test_origin_and_house_number_can_share_the_header_line(self):
        result = parse_pages(house().replace("UIO\n416-5236 2811", "UIO 416-5236 2811"))[0]
        self.assertEqual((result["origin"], result["origin_country"], result["hawb"]),
                         ("UIO", "ecuador", "0001"))

    def test_printed_full_box_rounding_is_not_a_missing_grower(self):
        text = house().replace("0.50 = FARM A", "0.06 = FARM A").replace(
            "4.50 = FARM B", "4.25 = FARM B").replace("TOTAL IN FULL: 5.00", "TOTAL IN FULL: 4.313")
        self.assertEqual(parse_pages(text)[0]["total_full"], 4.313)

    def test_malformed_second_house_rejects_the_whole_document(self):
        original = house("POLINA", "0003")
        for broken in [original.replace("BOX LABEL: POLINA\n", ""),
                       original.replace("416-5236 2811 0003", "unreadable house number"),
                       original.replace("65 1211 K 1273 3.45 4,391.85", "unreadable weights"),
                       original.replace("29.50 = FARM C", "unreadable growers"),
                       original.replace("29.50 = FARM C", "1.00 = FARM C"),
                       original.replace("AIR WAYBILL", "unreadable title")]:
            with self.subTest(broken=broken):
                with self.assertRaisesRegex(ValueError, "Страница 3"):
                    parse_pages(MASTER, house(), broken)

    def test_duplicate_house_number_is_not_double_counted(self):
        with self.assertRaisesRegex(ValueError, "повторяется номер house"):
            parse_pages(MASTER, house(), house())

    def test_master_number_cannot_be_used_as_house_number(self):
        text = house(mawb="416-52362811", number="416-52362811")
        with self.assertRaisesRegex(ValueError, "номер master AWB"):
            parse_pages(text)

    def test_house_prepaid_is_used_without_recomputing_or_adding_master_fees(self):
        result = parse_pages(house().replace("979.10 PREPAID", "981.25 PREPAID"))[0]
        self.assertEqual(result["total_awb"], 981.25)
        self.assertEqual((result["weight_charge"], result["other_charges"]), (959.1, 20))

    def test_legacy_forwarder_and_ocr_contract_is_wrapped_unchanged(self):
        for ocr in (False, True):
            with self.subTest(ocr=ocr):
                previous = {"mark": "VIKA", "hawb": "306 0594 4328", "ocr": ocr}
                with patch.object(import_parser, "parse_forwarder_hawb_pdf", return_value=previous):
                    self.assertEqual(import_parser.parse_delivery_hawbs_pdf("legacy.pdf"), [previous])

    @unittest.skipUnless(os.path.isfile(os.path.expanduser("~/Downloads/AWB-2_261008_153814.pdf")),
                         "нет локального AWB-2 от 08.10")
    def test_real_awb_matches_existing_house_weights_and_costs(self):
        path = os.path.expanduser("~/Downloads/AWB-2_261008_153814.pdf")
        houses = import_parser.parse_delivery_hawbs_pdf(path)
        with pdfplumber.open(path) as pdf:
            prior = import_parser._awb_house_pages(pdf, ("DAMIR", "POLINA"))
        self.assertEqual(len(houses), 2)
        self.assertEqual([len(h["growers"]) for h in houses], [2, 9])
        for h in houses:
            for field, value in prior[h["mark"]].items():
                self.assertEqual(h[field], value)
            self.assertEqual(h["airport"], "AMS")
            self.assertEqual(h["master_awb_difference"], 20)
