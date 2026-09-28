"""
Парсер PDF/XLS-инвойсов импортных поставок (не-Голландия): Эквадор, Колумбия,
Кения и т.д. Отдельный, параллельный модуль голландскому pdf_parser.py -
голландскую логику не трогаем.

Архитектура - тот же принцип "шаблон на поставщика", что и в pdf_parser.py:
detect_*() + parse_*(), регистрация в TEMPLATES. Каждый шаблон возвращает
единую промежуточную структуру:

    {
        "source_filename": str,
        "supplier": str,           # человекочитаемое имя поставщика
        "mark": str | None,        # метка (склад/направление), если есть в самом файле
        "invoice_no": str | None,
        "invoice_date": str | None,
        "awb": str | None,
        "hawb": str | None,
        "forwarder": str | None,
        "airline": str | None,
        "destination": str | None,
        "boxes": [
            {"box_no": str, "box_type": str | None, "box_size": float | None,
             "items": [
                {"variety": str, "length_cm": float, "stems": int,
                 "price": float, "total": float},
                ...
             ]},
            ...
        ],
        "totals": {"total_stems": float | None, "total_fob": float | None},
    }

Длины (length_cm) парсер отдаёт ровно такими, как напечатаны в инвойсе.
Правило MIX "считать по верхней ростовке" здесь БОЛЬШЕ НЕ ПРИМЕНЯЕТСЯ (правка
закупщика 2026-09-11): оно относится только к альстромерии и живёт теперь в
import_combine.py, где уже известна культура коробки. Раньше оно стояло во всех
шаблонах и портило розы - 50 см превращались в 60/70/80 см.
"""
import re
from collections import Counter
import pdfplumber
import xlrd

# Доля коробки по первой букве кода типа коробки (F=Full/H=Half/Q=Quarter/E=Eighth).
BOX_SIZE_TABLE = {"F": 1.0, "H": 0.5, "Q": 0.25, "E": 0.125}

# Коды, которые по первой букве не определить. SB (small box) = 1/16:
# подтверждено авианакладной 369-1151 1964 (у POLINA "TOTAL IN FULL 22.813"
# сходится только при SB = 0.0625) и файлом закупщика за 30.09. TESSA при
# этом печатает в "Number in Fulls" 0.1667 - этой цифре для SB не верим.
BOX_SIZE_CODES = {"SB": 0.0625}


def _box_size_from_code(box_type_code):
    if not box_type_code:
        return None
    code = box_type_code.strip().upper()
    if code in BOX_SIZE_CODES:
        return BOX_SIZE_CODES[code]
    return BOX_SIZE_TABLE.get(code[:1])


NUM_RE = re.compile(r"^-?\d+([.,]\d+)?$")


def _to_float(s):
    if s is None:
        return None
    s = str(s).strip().replace(" ", "")
    if not s:
        return None
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _to_int(s):
    f = _to_float(s)
    return int(f) if f is not None else None


def _cluster_rows(words, y_tol=3.0):
    words = sorted(words, key=lambda w: w["top"])
    rows, current, current_top = [], [], None
    for w in words:
        if current_top is None or abs(w["top"] - current_top) <= y_tol:
            current.append(w)
            current_top = w["top"] if current_top is None else current_top
        else:
            rows.append(current)
            current, current_top = [w], w["top"]
    if current:
        rows.append(current)
    return rows


def _col_for(x0, columns):
    for name, a, b in columns:
        if a <= x0 < b:
            return name
    return "other"


# ---------------------------------------------------------------------------
# Шаблон "Matiz Roses" (Эквадор)
# ---------------------------------------------------------------------------

def detect_matiz(pdf):
    text = pdf.pages[0].extract_text() or ""
    return "MATIZ ROSES" in text


# В таблице позиций Matiz колонка VARIETIES не имеет фиксированной ширины -
# она начинается сразу после текста BOX TYPE, а тот варьируется по длине
# ("HB-1000" / "HB S-100" / ...), из-за чего пиксельные координаты колонок
# съезжают от инвойса к инвойсу. Разбираем построчно текстом: у каждой строки
# позиции всегда есть хвост из 6 чисел (CM, BUNCH STEMS, BUNCH BOX, STEMS BOX,
# UNIT PRICE, TOTAL PRICE) - остаток слева однозначно восстанавливается.
_MATIZ_TAIL_RE = re.compile(
    r"^(.*?)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([\d.]+)\s+([\d.]+)$"
)
# Шапка коробки: "<order1> - <order2> [<box_code...>] <bx> <box_type> <variety>"
# box_code (если есть) - два токена без цифро-буквенного кода тары в начале
# (напр. "mix40 250"); отличаем от <bx> тем, что сразу после <bx> идёт код
# тары, начинающийся с буквы вместимости (H/Q/F/E).
_MATIZ_HEADER_RE = re.compile(
    r"^(?P<order1>\d+)\s*-\s*(?P<order2>\d+)\s+"
    r"(?:(?!\d+\s+[HQFE])(?P<box_code>\S+\s+\S+)\s+)?"
    r"(?P<bx>\d+)\s+"
    # Код тары бывает из двух частей: "HB S-100" / "HB L-120" / "HB M-110"
    # (вторая часть - размер бутона). Если её не захватить, она утекает в
    # название сорта (правка закупщика 2026-09-10: в VARIETY попадало
    # "L-120 EXPLORER" вместо "EXPLORER").
    r"(?P<box_type>[HQFE]\S*(?:\s+[A-Z]{1,2}-\d+)?)\s+"
    r"(?P<variety>.+)$"
)


def parse_matiz(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"Invoice #:\s*([0-9A-Za-z]+)", full_text)
    date_m = re.search(r"Date:\s*([0-9\-]+)", full_text)
    to_m = re.search(r"To:\s*([A-Za-z ]+)/\s*([A-Za-z0-9]+)", full_text)
    awb_m = re.search(r"AWB:\s*([0-9\-]+)", full_text)
    hawb_m = re.search(r"HAWB:\s*([0-9A-Za-z\-]+)", full_text)
    airline_m = re.search(r"Airline:\s*(.+)", full_text)
    forwarder_m = re.search(r"Freigh Forward:\s*([A-Za-z]+)", full_text)
    country_m = re.search(r"Country:\s*([A-Za-z]+)", full_text)
    total_fob_m = re.search(r"TOTAL FOB\s+([\d,\.]+)\s+([\d\.]+)\s+([\d\.]+)", full_text)

    mark = to_m.group(2).strip() if to_m else None

    boxes = []
    current_box = None
    in_table = False

    for page in pdf.pages:
        text = page.extract_text() or ""
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            if line.startswith(("ORDER", "TYPE", "BOX BUNCH")) or "VARIETIES" in line:
                in_table = True
                continue
            if not in_table:
                continue
            if line.startswith("TOTAL FOB") or line.startswith("TOTAL:") \
                    or line.startswith("Box Type") or "Forwarder:" in line:
                in_table = False
                current_box = None
                continue

            m = _MATIZ_TAIL_RE.match(line)
            if not m:
                continue
            prefix, cm, bunch_stems, bunch_box, stems_box, unit_price, total_price = m.groups()
            cm = _to_float(cm)
            stems_box = _to_int(stems_box)
            unit_price = _to_float(unit_price)
            total_price = _to_float(total_price)

            hm = _MATIZ_HEADER_RE.match(prefix)
            if hm:
                box_type = hm.group("box_type")
                current_box = {
                    "box_no": f"{hm.group('order1')}-{hm.group('order2')}",
                    "box_type": box_type,
                    "box_size": _box_size_from_code(box_type),
                    "items": [],
                }
                boxes.append(current_box)
                variety = hm.group("variety")
            else:
                variety = prefix

            if current_box is not None:
                current_box["items"].append({
                    "variety": variety, "length_cm": cm, "stems": stems_box,
                    "price": unit_price, "total": total_price,
                })


    total_stems = _to_float(total_fob_m.group(1)) if total_fob_m else None
    total_fob = _to_float(total_fob_m.group(3)) if total_fob_m else None

    return {
        "source_filename": source_filename,
        "supplier": "Matiz Roses",
        "mark": mark,
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1) if awb_m else None,
        "hawb": hawb_m.group(1) if hawb_m else None,
        "forwarder": forwarder_m.group(1) if forwarder_m else None,
        "airline": airline_m.group(1).strip() if airline_m else None,
        "destination": country_m.group(1) if country_m else None,
        "boxes": boxes,
        "totals": {"total_stems": total_stems, "total_fob": total_fob},
    }


# ---------------------------------------------------------------------------
# Шаблон "Gardaexport S.A." (Эквадор)
# ---------------------------------------------------------------------------

def detect_gardaexport(pdf):
    text = pdf.pages[0].extract_text() or ""
    return "GARDAEXPORT" in text.upper()


# Колонки данных Gardaexport - PRODUCT NAME (статичный текст "ALSTROEMERIAS
# (ALSTROEMERIA AUREA)") переносится на 3 отдельные строки-обёртки поверх
# самой строки с данными, поэтому строим колонки не из шапки, а из реально
# наблюдаемых x0 в строке данных (TB/COLOR/VARIETY/LENGTH/BUNCHES/STEMS/
# PRICE/TOTAL); BOX N° стоит ещё на одну строку выше и разбирается отдельно.
_GARDA_COLUMNS = [
    # До 66, а не до 50: номер коробки бывает диапазоном из трёхзначных
    # номеров ("001 - 002"), и его хвост вылезал в колонку TB.
    ("box_no", -1, 66),
    ("tb", 66, 144.5),
    ("color", 144.5, 242.5),
    ("variety", 242.5, 334.05),
    ("length", 334.05, 418.85),
    ("bunches", 418.85, 451.85),
    ("stems", 451.85, 485.25),
    ("price", 485.25, 526.9),
    ("total", 526.9, 9995),
]


def parse_gardaexport(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"INVOICE #\s*([0-9]+)", full_text)
    date_m = re.search(r"Date\s*:\s*([0-9/]+)", full_text)
    label_m = re.search(r"Label\s*:\s*([A-Za-z0-9]+)", full_text)
    awb_m = re.search(r"A\.W\.B\. N°\s*:\s*([0-9]+)", full_text)
    hawb_m = re.search(r"H\.A\.W\.B\. N°\s*:\s*([0-9A-Za-z]+)", full_text)
    airline_m = re.search(r"Airline\.\s*:\s*(.+)", full_text)
    forwarder_m = re.search(r"Shipper\s*:\s*(.+)", full_text)
    country_m = re.search(r"Country:\s*([A-Za-z]+)", full_text)
    total_m = re.search(r"TOTAL\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$", full_text, re.M)

    boxes = []
    current_box = None
    pending_box_no = None

    for page in pdf.pages:
        words = page.extract_words()
        rows = _cluster_rows(words, y_tol=1.5)
        for row_words in rows:
            row_words = sorted(row_words, key=lambda w: w["x0"])
            rowdict = {}
            for w in row_words:
                c = _col_for(w["x0"], _GARDA_COLUMNS)
                rowdict.setdefault(c, []).append(w["text"])
            rowdict = {k: " ".join(v) for k, v in rowdict.items()}

            # Номер коробки бывает и диапазоном: "01 - 02" значит ДВЕ
            # физические коробки одинакового содержимого, а количества в
            # строках даны на всю группу (проверено: сумма строк совпадает с
            # итогом инвойса). Раньше такой инвойс терялся целиком - regex
            # ждал ровно две цифры, коробка не создавалась и все позиции
            # молча отбрасывались (реальный случай 2026-09-10).
            box_no_tok = rowdict.get("box_no", "").strip()
            range_m = re.match(r"^(\d{2,3})\s*-\s*(\d{2,3})$", box_no_tok)
            if range_m:
                start, end = int(range_m.group(1)), int(range_m.group(2))
                pending_box_no = [str(n).zfill(len(range_m.group(1)))
                                   for n in range(start, end + 1)] or [range_m.group(1)]
            elif re.match(r"^\d{2,3}$", box_no_tok):
                pending_box_no = [box_no_tok]

            tb_raw = rowdict.get("tb", "").strip()
            tb = tb_raw.split()[0] if tb_raw else ""
            if tb not in BOX_SIZE_TABLE:
                continue  # не строка с товаром (шапка/обёртка PRODUCT NAME/итоги)

            length_cm = _to_float(rowdict.get("length"))
            bunches = _to_int(rowdict.get("bunches"))
            stems = _to_int(rowdict.get("stems"))
            price = _to_float(rowdict.get("price"))
            total = _to_float(rowdict.get("total"))
            variety = rowdict.get("variety", "").strip()
            color = rowdict.get("color", "").strip()
            if not variety or stems is None:
                continue

            if pending_box_no is not None:
                # Для диапазона заводим по коробке на номер; позиции в них
                # положим ниже, поделив количества поровну.
                current_box = {
                    "box_no": pending_box_no[0],
                    "box_type": tb,
                    "box_size": BOX_SIZE_TABLE.get(tb),
                    "items": [],
                    "_siblings": pending_box_no[1:],
                }
                boxes.append(current_box)
                pending_box_no = None

            if current_box is not None:
                current_box["items"].append({
                    "variety": f"{color} {variety}".strip(), "length_cm": length_cm,
                    "stems": stems, "price": price, "total": total,
                    "bunches": bunches,
                })

    # Разворачиваем диапазоны: количества в строках даны на всю группу коробок,
    # делим их поровну между физическими коробками группы.
    expanded = []
    for box in boxes:
        siblings = box.pop("_siblings", [])
        count = 1 + len(siblings)
        if count == 1:
            expanded.append(box)
            continue
        for box_no in [box["box_no"]] + siblings:
            expanded.append({
                "box_no": box_no,
                "box_type": box["box_type"],
                "box_size": box["box_size"],
                "items": [{**it,
                            "stems": (it["stems"] / count if it["stems"] % count
                                      else it["stems"] // count),
                            "total": round(it["total"] / count, 2) if it.get("total") else it.get("total"),
                            } for it in box["items"]],
            })
    boxes = expanded


    total_stems = _to_float(total_m.group(2)) if total_m else None
    total_fob = _to_float(total_m.group(3)) if total_m else None

    return {
        "source_filename": source_filename,
        "supplier": "Gardaexport S.A.",
        "mark": label_m.group(1) if label_m else None,
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1) if awb_m else None,
        "hawb": hawb_m.group(1) if hawb_m else None,
        "forwarder": forwarder_m.group(1).strip() if forwarder_m else None,
        "airline": airline_m.group(1).strip() if airline_m else None,
        "destination": country_m.group(1) if country_m else None,
        "boxes": boxes,
        "totals": {"total_stems": total_stems, "total_fob": total_fob},
    }


# ---------------------------------------------------------------------------
# Шаблон "Florsani" (Эквадор, аэропорт-код ECFLOSANMAL в подвале инвойса)
# ---------------------------------------------------------------------------

def detect_florsani(pdf):
    text = pdf.pages[0].extract_text() or ""
    return "ECFLOSANMAL" in text


# Строка позиции - "<Pcs> <BoxType> <Description...> <Color> <Weigth> <CM>
# <Bunch/Box> <Stems/Bunch> <Price> <TotalStems> <TotalPrice>". Pcs - число
# ФИЗИЧЕСКИХ коробок этой строки (сумма Pcs по всем строкам = TOTAL PIECES);
# TotalStems/TotalPrice - агрегат по всем Pcs коробкам сразу, поэтому каждую
# строку разворачиваем на Pcs одинаковых коробок с stems/total, поделёнными
# на Pcs (см. ТЗ: "закупщик разворачивает на N одинаковых строк").
_FLORSANI_ITEM_RE = re.compile(
    r"^(?P<pcs>\d+)\s+(?P<box_type>[A-Z]{2,4})\s+(?P<desc>.+?)\s+(?P<color>\S+)\s+"
    r"(?P<weigth>\d+)\s+(?P<cm>\d+)\s+(?P<bunch_box>\d+)\s+(?P<stems_bunch>\d+)\s+"
    r"(?P<price>[\d.]+)\s+(?P<total_stems>[\d,]+)\s+(?P<total_price>[\d.]+)$"
)


def parse_florsani(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"INVOICE N°\s*(\d+)", full_text)
    date_m = re.search(r"DATE:\s*([\d\-]+)", full_text)
    mark_m = re.search(r"CONSIGNEE\s*:\s*(\S+)", full_text)
    destination_m = re.search(r"DESTINATION:\s*(\S+)", full_text)
    awb_m = re.search(r"AWB#:(\S+)", full_text)
    hawb_m = re.search(r"HAWB:(\S+)", full_text)
    flight_m = re.search(r"FLIGHT:\s*(\S+)", full_text)
    forwarder_m = re.search(r"FORWARDER:\s*(.+?)\s*PO:", full_text)

    boxes = []
    box_counter = 0
    in_table = False

    for page in pdf.pages:
        text = page.extract_text() or ""
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            if line.startswith("PcsBox"):
                in_table = True
                continue
            if not in_table:
                continue
            if line.startswith(("TOTAL PIECES", "Single Flowers")):
                in_table = False
                continue

            m = _FLORSANI_ITEM_RE.match(line)
            if not m:
                continue  # шапка-продолжение ("Type Box Bunch...") или строка субтотала

            pcs = _to_int(m.group("pcs"))
            box_type = m.group("box_type")
            desc = m.group("desc").strip()
            color = m.group("color").strip()
            cm = _to_float(m.group("cm"))
            price = _to_float(m.group("price"))
            total_stems = float(m.group("total_stems").replace(",", ""))
            total_price = _to_float(m.group("total_price"))
            if not pcs:
                continue

            stems_per_box = total_stems / pcs
            if stems_per_box == int(stems_per_box):
                stems_per_box = int(stems_per_box)
            price_per_box = round(total_price / pcs, 2)
            variety = desc if color.upper() == "NINGUNO" else f"{color} {desc}"

            for _ in range(pcs):
                box_counter += 1
                boxes.append({
                    "box_no": str(box_counter),
                    "box_type": box_type,
                    "box_size": _box_size_from_code(box_type),
                    "items": [{
                        "variety": variety, "length_cm": cm, "stems": stems_per_box,
                        "price": price, "total": price_per_box,
                    }],
                })


    # Итоги сверяем по блоку "Single Flowers" (Nro Pieces Stems Price Value).
    summary_m = re.search(r"Nro Pieces Stems Price Value\n(.*?)\nCliente:", full_text, re.S)
    total_stems = total_fob = None
    if summary_m:
        total_stems, total_fob = 0.0, 0.0
        for line in summary_m.group(1).splitlines():
            row_m = re.match(r"^.+\s(\d+)\s+([\d,]+)\s+([\d.]+)\s+([\d.]+)$", line.strip())
            if row_m:
                # "," здесь - разделитель тысяч (US-формат: "2,250"), не десятичных.
                total_stems += float(row_m.group(2).replace(",", ""))
                total_fob += _to_float(row_m.group(4))

    return {
        "source_filename": source_filename,
        "supplier": "Florsani",
        "mark": mark_m.group(1) if mark_m else None,
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1) if awb_m else None,
        "hawb": hawb_m.group(1) if hawb_m else None,
        "forwarder": forwarder_m.group(1).strip() if forwarder_m else None,
        "airline": flight_m.group(1) if flight_m else None,
        "destination": destination_m.group(1) if destination_m else None,
        "boxes": boxes,
        "totals": {"total_stems": total_stems, "total_fob": total_fob},
    }


# ---------------------------------------------------------------------------
# Шаблон "Rosas del Corazon" (Эквадор, форвардер Saftec, псевдо-табличный
# формат с "I" вместо вертикальных линий таблицы)
# ---------------------------------------------------------------------------

def detect_rosas_corazon(pdf):
    text = (pdf.pages[0].extract_text() or "").replace(" ", "")
    return "ROSASDELCORAZON" in text.upper()


# Коды коробок Rosas del Corazon не подчиняются общему правилу "первая буква
# кода = размер" (см. BOX_SIZE_TABLE) - у них свой закрытый каталог из 6 кодов
# (3 формата H, 2 формата Q, 1 формат E), полученный от закупщика напрямую.
# Проверено на реальных данных: сумма по этой таблице для F_SIRI_2227531.pdf
# даёт TOTALFULL=9.750 - ровно как напечатано в самом инвойсе.
_ROSAS_BOX_SIZE = {
    "XJHB": 0.5,   # Extra Jumbo
    "JB2": 0.5,    # Jumbo
    "HB": 0.5,     # HB - EU
    "LQB": 0.25,   # QBL (буквы иногда в другом порядке в разных инвойсах)
    "QBL": 0.25,
    "QBM": 0.25,
    "EB": 0.125,
}


def _rosas_box_size(box_type_code):
    return _ROSAS_BOX_SIZE.get((box_type_code or "").strip().upper())


# Строка позиции: "[<Pcs> ]I<BoxType> I <Units>I<Mark> I <Variety> <NN>CMI
# $<UnitPrice>I $<Total>". Pcs указан только в СТРОКЕ, открывающей новую
# физическую коробку; последующие строки без Pcs - дополнительные позиции
# (другой сорт/длина), упакованные в ТУ ЖЕ коробку - это и есть MIX-случай
# из ТЗ. Пустого Pcs без предшествующей "новой" строки в реальных файлах
# не встречалось.
_ROSAS_ITEM_RE = re.compile(
    r"^(?P<pcs>\d+)?\s*I(?P<box_type>\S+)\s+I\s*(?P<units>\d+)I(?P<mark>\S+)\s+I\s*"
    r"(?P<variety>.+?)\s*(?P<length>\d+)CMI\s*\$(?P<price>[\d.]+)I\s*\$(?P<total>[\d,.]+)$"
)


def parse_rosas_corazon(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"COMERCIALINVOICE\s*(\d+)", full_text)
    date_m = re.search(r"DATE:\s*([\d/]+)", full_text)
    awb_m = re.search(r"MAWB#:\s*(\S+)", full_text)
    hawb_m = re.search(r"HAWB#:\s*(\S+)", full_text)
    forwarder_m = re.search(r"FORWARDER:(\S+)", full_text)
    mark_m = re.search(r"CONSIGNEE:\s*I.*\n(\S+)\s+I", full_text)
    destination_m = re.search(r"^([A-Z]+)\s+I\s+[A-Z]+\s*\nPHONE:", full_text, re.M)

    boxes = []
    box_counter = 0
    current_group = None  # список item-строк текущей физической коробки/группы
    # Метка печатается в колонке CODE каждой строки ("VIKA", "DAMIR"). Шапка
    # CONSIGNEE ненадёжна: часто там "FLOWERS IRIS", и регекс брал
    # "FLOWERSIRIS" (реальные случаи: F_DAMIR_2250103, F_VIKA_2251588).
    code_marks = Counter()

    def _flush_group(group):
        nonlocal box_counter
        if not group:
            return
        pcs, items = group
        for _ in range(pcs):
            box_counter += 1
            box_items = []
            for it in items:
                stems_per_box = it["stems"] / pcs
                if stems_per_box == int(stems_per_box):
                    stems_per_box = int(stems_per_box)
                box_items.append({
                    "variety": it["variety"], "length_cm": it["length_cm"],
                    "stems": stems_per_box, "price": it["price"],
                    "total": round(it["total"] / pcs, 2),
                })
            boxes.append({
                "box_no": str(box_counter),
                "box_type": items[0]["box_type"],
                "box_size": _rosas_box_size(items[0]["box_type"]),
                "items": box_items,
            })

    for page in pdf.pages:
        text = page.extract_text() or ""
        for line in text.splitlines():
            m = _ROSAS_ITEM_RE.match(line.strip())
            if not m:
                continue
            code_marks[m.group("mark")] += 1
            item = {
                "box_type": m.group("box_type"),
                "variety": m.group("variety").strip(),
                "length_cm": _to_float(m.group("length")),
                "stems": float(m.group("units").replace(",", "")),
                "price": _to_float(m.group("price")),
                "total": float(m.group("total").replace(",", "")),
            }
            pcs = m.group("pcs")
            if pcs is not None:
                _flush_group(current_group)
                current_group = (int(pcs), [item])
            elif current_group is not None:
                current_group[1].append(item)
            # строка без Pcs до первой "новой" строки в реальных файлах не
            # встречалась - если встретится, будет молча потеряна; лучше
            # сообщить явно, чем придумать поведение.
    _flush_group(current_group)


    footer_m = re.search(
        r"^([\d.]+)\s*I\s*I\s*([\d,]+)I\s*TOTALF\.O\.B\.VALUE:I\s*\$([\d,.]+)$",
        full_text, re.M)
    total_stems = float(footer_m.group(2).replace(",", "")) if footer_m else None
    total_fob = float(footer_m.group(3).replace(",", "")) if footer_m else None

    return {
        "source_filename": source_filename,
        "supplier": "Rosas del Corazon (Saftec)",
        "mark": (code_marks.most_common(1)[0][0] if code_marks
                 else (mark_m.group(1) if mark_m else None)),
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1) if awb_m else None,
        "hawb": hawb_m.group(1) if hawb_m else None,
        "forwarder": forwarder_m.group(1) if forwarder_m else None,
        "airline": None,
        "destination": destination_m.group(1) if destination_m else None,
        "boxes": boxes,
        "totals": {"total_stems": total_stems, "total_fob": total_fob},
    }


# ---------------------------------------------------------------------------
# Шаблон "Monterosas" (Эквадор)
# ---------------------------------------------------------------------------

def detect_monterosas(pdf):
    text = (pdf.pages[0].extract_text() or "").replace(" ", "").upper()
    return "MONTEROSAS" in text


# Длина (CM) закодирована не отдельной колонкой, а тем, В КАКОЙ из трёх
# колонок-значений (50/60/70) стоит число бунчей - остальные две пустые.
# "Order" - номер физической коробки: несколько строк с одинаковым Order
# упакованы в одну коробку (MIX-случай, см. ТЗ).
_MONTE_COLUMNS = [
    ("box_type", -1, 70),
    ("pieces", 70, 100),
    ("order", 100, 145),
    ("mark", 145, 198),
    ("product", 198, 260),
    ("stems_bunch", 260, 283),
    ("len50", 283, 302),
    ("len60", 302, 322),
    ("len70", 322, 342),
    ("t_bunch", 342, 385),
    ("t_stems", 385, 429),
    ("price", 429, 474),
    ("total", 474, 9995),
]


def parse_monterosas(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"INVOICE\s+(\d+)", full_text)
    date_m = re.search(r"DATE:\s*([\d\-]+)", full_text)
    mark_m = re.search(r"CONSIGNEE\s+(\S+)", full_text)
    country_m = re.search(r"COUNTRY:\s*(\S+)", full_text)
    awb_m = re.search(r"M\s*A\s*W\s*B\s*:\s*(\S[\d\- ]*\d)", full_text)
    hawb_m = re.search(r"H\s*A\s*W\s*B\s*:\s*(\S+)", full_text)
    airline_m = re.search(r"AIR\s*LINE:\s*(.+)", full_text)
    forwarder_m = re.search(r"FREIGHT FORWARDER\s*\n(\S+)", full_text)

    groups = {}  # order_no -> список item-словарей (сохраняем порядок вставки)
    order_seq = []

    for page in pdf.pages:
        words = page.extract_words()
        rows = _cluster_rows(words, y_tol=2.0)
        for row_words in rows:
            rowdict = {}
            for w in sorted(row_words, key=lambda w: w["x0"]):
                c = _col_for(w["x0"], _MONTE_COLUMNS)
                rowdict.setdefault(c, []).append(w["text"])
            rowdict = {k: " ".join(v) for k, v in rowdict.items()}

            box_type = rowdict.get("box_type", "").strip()
            if _box_size_from_code(box_type) is None:
                continue  # не строка позиции (шапка/итоги/пусто)

            order_no = rowdict.get("order", "").strip()
            pieces = _to_int(rowdict.get("pieces")) or 1
            variety = rowdict.get("product", "").strip()
            stems_bunch = _to_float(rowdict.get("stems_bunch"))
            t_stems = _to_float(rowdict.get("t_stems"))
            price = _to_float(rowdict.get("price"))
            total = _to_float(rowdict.get("total"))
            if not variety or t_stems is None:
                continue

            length_cm = None
            for col in ("len50", "len60", "len70"):
                if rowdict.get(col, "").strip():
                    length_cm = float(col[3:])
                    break

            item = {"variety": variety, "length_cm": length_cm, "stems": t_stems,
                     "price": price, "total": total, "box_type": box_type,
                     "pieces": pieces}
            if order_no not in groups:
                groups[order_no] = []
                order_seq.append(order_no)
            groups[order_no].append(item)

    boxes = []
    box_counter = 0
    for order_no in order_seq:
        items = groups[order_no]
        pieces = items[0]["pieces"]
        for _ in range(pieces):
            box_counter += 1
            box_items = []
            for it in items:
                stems_per_box = it["stems"] / pieces
                if stems_per_box == int(stems_per_box):
                    stems_per_box = int(stems_per_box)
                box_items.append({
                    "variety": it["variety"], "length_cm": it["length_cm"],
                    "stems": stems_per_box, "price": it["price"],
                    "total": round(it["total"] / pieces, 2),
                })
            boxes.append({
                "box_no": str(box_counter),
                "box_type": items[0]["box_type"],
                "box_size": _box_size_from_code(items[0]["box_type"]),
                "items": box_items,
            })


    total_stems_m = re.search(r"Total Stems\s+([\d,]+)", full_text)
    total_fob_m = re.search(r"^Total\s+([\d,.]+)$", full_text, re.M)
    total_stems = float(total_stems_m.group(1).replace(",", "")) if total_stems_m else None
    total_fob = float(total_fob_m.group(1).replace(",", "")) if total_fob_m else None

    return {
        "source_filename": source_filename,
        "supplier": "Monterosas",
        "mark": mark_m.group(1) if mark_m else None,
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1).strip() if awb_m else None,
        "hawb": hawb_m.group(1) if hawb_m else None,
        "forwarder": forwarder_m.group(1) if forwarder_m else None,
        "airline": airline_m.group(1).strip() if airline_m else None,
        "destination": country_m.group(1) if country_m else None,
        "boxes": boxes,
        "totals": {"total_stems": total_stems, "total_fob": total_fob},
    }


# ---------------------------------------------------------------------------
# Шаблон "star roses" (EL CAMPANARIO DE SANTA ANITA SCC, staroses.com)
# ---------------------------------------------------------------------------
#
# Тот самый поставщик, который долго был отложен ("нет образца"). Образцы
# пришли 2026-09-09. Формат простой и ровный, одна строка = одна позиция:
#   "BOX N° | TB | BOX CODE | VARIETY | LENGTH (CM) | TOTAL BUNCHE |
#    TOTAL STEMS | UNIT PRICE | TOTAL FOB $"
#   "1 H POLINA EXPLORER 50 16 400 0,35 140,00"
# Номер коробки печатается только у ПЕРВОЙ позиции коробки (у остальных
# позиций той же коробки - пусто, это MIX). TB - код типа коробки (H/Q/F/E),
# BOX CODE - метка. Числа в европейском формате (запятая - десятичный
# разделитель). Внизу есть сводка "Total Full F H Q E" - используем её для
# сверки суммы долей коробок.
# FARM/PRODUCT для этого поставщика ПОДТВЕРЖДЕНЫ реальным примером factura
# ("Invoice total 30.08 SIRI MOS.xls": FARM "star roses", PRODUCT
# "ROSES star roses") - см. import_combine.py.
_STAR_ITEM_RE = re.compile(
    r"^(?P<box_no>\d+)?\s*(?P<tb>[FHQE])\s+(?P<code>\S+)\s+(?P<variety>.+?)\s+"
    r"(?P<len>\d+)\s+(?P<bunches>\d+)\s+(?P<stems>\d+)\s+"
    r"(?P<price>[\d.,]+)\s+(?P<total>[\d.,]+)$"
)


def _eu_float(s):
    """Числа star roses в европейском формате: "1.234,56" / "140,00"."""
    if s is None:
        return None
    s = str(s).strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def detect_star_roses(pdf):
    text = (pdf.pages[0].extract_text() or "").upper()
    return "STAROSES.COM" in text or "EL CAMPANARIO DE SANTA ANITA" in text


def parse_star_roses(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"INVOICE #\s*(\S+)", full_text)
    date_m = re.search(r"Date\s*:\s*([\d/]+)", full_text)
    awb_m = re.search(r"A\.W\.B\.\s*N°\s*:\s*(\S+)", full_text)
    hawb_m = re.search(r"H\.A\.W\.B\.\s*N°\s*:\s*(\S+)", full_text)
    airline_m = re.search(r"Airline\.?\s*:\s*(.+)", full_text)
    forwarder_m = re.search(r"Shipper\s*:\s*(\S+)", full_text)
    totals_m = re.search(r"^TOTAL\s+(\d+)\s+(\d+)\s+([\d.,]+)\s*$", full_text, re.M)
    # "Country:" встречается дважды - у самой фермы (ECUADOR) и у получателя;
    # нужен последний (страна назначения).
    countries = re.findall(r"Country:\s*(\S+)", full_text)

    boxes = []
    current_box = None
    marks = []
    for line in full_text.splitlines():
        m = _STAR_ITEM_RE.match(line.strip())
        if not m:
            continue
        marks.append(m.group("code"))
        if m.group("box_no") or current_box is None:
            box_type = m.group("tb")
            current_box = {
                "box_no": m.group("box_no") or str(len(boxes) + 1),
                "box_type": box_type,
                "box_size": _box_size_from_code(box_type),
                "items": [],
            }
            boxes.append(current_box)
        current_box["items"].append({
            "variety": m.group("variety").strip(),
            "length_cm": _to_float(m.group("len")),
            "stems": _to_int(m.group("stems")),
            "price": _eu_float(m.group("price")),
            "total": _eu_float(m.group("total")),
        })


    # Метка - значение колонки BOX CODE (одинаковое во всех строках инвойса);
    # дублируется в блоке "Label :" вверху, но там оно склеено с соседней
    # колонкой H.A.W.B., поэтому берём из таблицы.
    mark = max(set(marks), key=marks.count) if marks else None

    return {
        "source_filename": source_filename,
        "supplier": "EL CAMPANARIO DE SANTA ANITA SCC (star roses)",
        "mark": mark,
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1) if awb_m else None,
        "hawb": hawb_m.group(1) if hawb_m else None,
        "forwarder": forwarder_m.group(1) if forwarder_m else None,
        "airline": airline_m.group(1).strip() if airline_m else None,
        "destination": countries[-1] if countries else None,
        "boxes": boxes,
        "totals": {
            "total_stems": _to_float(totals_m.group(2)) if totals_m else None,
            "total_fob": _eu_float(totals_m.group(3)) if totals_m else None,
        },
    }


# ---------------------------------------------------------------------------
# Шаблон "Monterosas", ВТОРОЙ макет (встретился 2026-09-09)
# ---------------------------------------------------------------------------
#
# Та же ферма, но полностью другая вёрстка таблицы: длина вынесена в
# отдельную колонку CM (в старом макете длина кодировалась тем, в какой из
# трёх колонок 50/60/70 стоит число бунчей), другой порядок и другие
# X-координаты колонок. Старый парсер на таком файле молча выдавал 0 коробок,
# поэтому это отдельный шаблон, а не правка старого.
#
# Строка-заголовок коробки: "1 - 1 1 HB 4 ALTAMIRA 50 25 4 100 0.40 40.00"
#   (order "1 - 1", BX 1, тип коробки "HB 4", дальше как у обычной позиции)
# Последующие позиции той же коробки - без префикса: "MANDALA 50 25 4 100 ...".
# Пока есть ровно один реальный образец - структурные сюрпризы во втором
# файле этого макета ожидаемы (та же оговорка, что и для новых поставщиков).
_MONTE2_ITEM_RE = re.compile(
    r"^(?P<variety>.+?)\s+(?P<cm>\d+)\s+(?P<bunch_stems>\d+)\s+(?P<bunches>\d+)\s+"
    r"(?P<stems>\d+)\s+(?P<price>[\d.]+)\s+(?P<total>[\d,.]+)$"
)
_MONTE2_BOX_RE = re.compile(
    # Код тары может быть из двух частей ("HB 4", "HB L-120") - вторую часть
    # тоже забираем в btype, иначе она утечёт в название сорта.
    r"^(?P<order>\d+\s*-\s*\d+)\s+(?P<bx>\d+)\s+"
    r"(?P<btype>[A-Z]{1,3}(?:\s+(?:\d+|[A-Z]{1,2}-\d+))?)\s+(?P<rest>\S.*)$"
)


def detect_monterosas_v2(pdf):
    text = (pdf.pages[0].extract_text() or "")
    return "MONTEROSAS" in text.replace(" ", "").upper() and "VARIETIES" in text.upper()


def parse_monterosas_v2(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"Invoice #:\s*(\S+)", full_text)
    date_m = re.search(r"Date:\s*([\d\-]+)", full_text)
    mark_m = re.search(r"To:\s*(\S+)", full_text)
    country_m = re.search(r"Country:\s*(\S+)", full_text)
    awb_m = re.search(r"AWB:\s*(\S+)", full_text)
    hawb_m = re.search(r"HAWB:\s*(\S+)", full_text)
    airline_m = re.search(r"Airline:\s*(.+)", full_text)
    forwarder_m = re.search(r"Freigh\w*\s+Forward:\s*(.+)", full_text)
    totals_m = re.search(r"TOTAL FCA\s+(\d+)\s+[\d.]+\s+([\d,.]+)", full_text)

    boxes = []
    current_box = None
    for line in full_text.splitlines():
        line = line.strip()
        box_m = _MONTE2_BOX_RE.match(line)
        item_m = _MONTE2_ITEM_RE.match(box_m.group("rest") if box_m else line)
        if not item_m:
            continue
        if box_m:
            box_type = box_m.group("btype").split()[0]
            current_box = {
                "box_no": box_m.group("order").replace(" ", ""),
                "box_type": box_type,
                "box_size": _box_size_from_code(box_type),
                "items": [],
            }
            boxes.append(current_box)
        if current_box is None:
            continue
        current_box["items"].append({
            "variety": item_m.group("variety").strip(),
            "length_cm": _to_float(item_m.group("cm")),
            "stems": _to_int(item_m.group("stems")),
            "price": _to_float(item_m.group("price")),
            "total": _to_float(item_m.group("total").replace(",", "")),
        })


    return {
        "source_filename": source_filename,
        "supplier": "Monterosas",
        "mark": mark_m.group(1) if mark_m else None,
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1) if awb_m else None,
        "hawb": hawb_m.group(1) if hawb_m else None,
        "forwarder": forwarder_m.group(1).strip() if forwarder_m else None,
        "airline": airline_m.group(1).strip() if airline_m else None,
        "destination": country_m.group(1) if country_m else None,
        "boxes": boxes,
        "totals": {
            "total_stems": _to_float(totals_m.group(1)) if totals_m else None,
            "total_fob": _to_float(totals_m.group(2).replace(",", "")) if totals_m else None,
        },
    }


# ---------------------------------------------------------------------------
# Шаблон "TESSA CORP." (Эквадор)
# ---------------------------------------------------------------------------
#
# Настоящая координатная таблица со строкой заголовков колонок ("Boxes Order
# BoxT. Loc. Description Len Bun/Box Stems Price Total Label"), но описание
# сорта (Description) и код коробки (BoxT.) при этом свободно переносятся на
# соседние текстовые строки (сколько слов уместилось по ширине) - поэтому
# разбираем не текстовые строки, а слова с координатами (как для CeresFarms/
# Gardaexport/Monterosas), группируя их по X-диапазону колонки, а не по
# порядку строк. Числовой "хвост" позиции (Len/Bun/Stems/$Price/$Total) всегда
# на одной строке ("якорь"); Boxes/Order печатаются только у первой позиции
# коробки. Слова из зон BoxT./Description на строках ДО и ПОСЛЕ якоря
# относятся к тому якорю, к которому они ближе по вертикали (см. правки
# закупщика 2026-09-08 - до этой правки код брал ровно одну соседнюю
# текстовую строку целиком, из-за чего в описание сорта попадали обрывки
# кода коробки/локации, а тип коробки на многокоробочных инвойсах путался
# между соседними коробками).
_TESSA_TAIL_RE = re.compile(
    r"(?P<len>\d+)\s+(?P<bun>\d+)\s+(?P<stems>\d+)\s+\$(?P<price>[\d.]+)\s+\$(?P<total>[\d,.]+)$"
)
# Код типа коробки - 1-3 заглавные буквы, начинающиеся с F/H/Q/E (см.
# BOX_SIZE_TABLE), или SB (BOX_SIZE_CODES). Именно по этому шаблону, а не "первое слово в зоне",
# отличаем настоящий код коробки (QB/HB) от случайных слов сорта/локации,
# которые тоже иногда попадают в ту же X-зону из-за смещения при переносе.
_TESSA_BOXCODE_TOKEN_RE = re.compile(r"^(?:[FHQE][A-Z]{0,2}|SB)$")

# Колонка "Loc." - код фермы внутри TESSA ("TESSA-" и под ним "P", "E2",
# "PS2"...). Она печатается в правой части зоны boxcode (x ~198), левее -
# сам код коробки ("HB XL 1"). По этому коду коробка относится к реальной
# ферме (TESSA-P -> Positano) - нужно разделу "Поставки": форвардер пишет в
# HAWB именно ферму, а не TESSA.
_TESSA_LOC_MIN_X = 193

_TESSA_COLUMNS = [
    ("boxes", -1, 100), ("order", 100, 145), ("boxcode", 145, 212), ("desc", 212, 305),
    ("nums", 305, 515), ("label", 515, 9995),
]


def _tessa_mark(pdf):
    """Метка = значение колонки SHIP CUSTOMER. В строке под шапкой
    "TO BILL CUSTOMER ... SHIP CUSTOMER" стоят СРАЗУ ДВА значения
    ("Flowers IRIS" слева - это клиент, "DAMIR" справа - это метка), поэтому
    делим строку по X: всё правее середины между заголовками колонок.
    Раньше метка бралась регексом "второе слово строки" и всегда получалось
    "IRIS" (часть названия клиента) - из-за чего инвойсы РАЗНЫХ меток
    сваливались в одну несуществующую метку (найдено на реальной партии
    2026-09-09: там были DAMIR и POLINA)."""
    rows = _cluster_rows(pdf.pages[0].extract_words(), y_tol=2.5)
    for i, row_words in enumerate(rows):
        texts = [w["text"] for w in row_words]
        if "SHIP" not in texts or "CUSTOMER" not in texts or i + 1 >= len(rows):
            continue
        ship_x = min(w["x0"] for w in row_words if w["text"] == "SHIP")
        bill_x = min(w["x0"] for w in row_words)
        boundary = (bill_x + ship_x) / 2
        value = [w["text"] for w in sorted(rows[i + 1], key=lambda w: w["x0"])
                  if w["x0"] >= boundary]
        if value:
            return " ".join(value)
    return None


def detect_tessa(pdf):
    text = pdf.pages[0].extract_text() or ""
    return "TESSA CORP" in text.upper()


def parse_tessa(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"Invoice Number\s+(\S+)", full_text)
    date_m = re.search(r"Invoice Date\s+([\d/]+)", full_text)
    mark = _tessa_mark(pdf)
    awb_m = re.search(r"AWB\s+(\S[\d\- ]*\d)", full_text)
    hawb_m = re.search(r"HAWB\s+(\S+)", full_text)
    airline_m = re.search(r"Airline\s+(.+)", full_text)
    forwarder_m = re.search(r"Cargo Agency\s+(\S+)", full_text)
    country_m = re.search(r"Country Final Destination\s+(\S+)", full_text)
    total_m = re.search(r"TOTALS\s+(\d+)\s+(\d+)\s+\$([\d,.]+)", full_text)

    boxes = []
    current_box = None

    # Считаем все страницы одним непрерывным потоком строк (со смещением
    # "top" на накопленную высоту предыдущих страниц) - иначе позиция, чьё
    # описание сорта переносится через разрыв страницы (последнее слово
    # оказывается в начале следующей страницы), делится на два несвязанных
    # куска. Важно: фильтровать шапку/итоги/AWB и считать смещение нужно ДО
    # объединения страниц и на основе только оставшихся строк - иначе futer/
    # юридический текст внизу страницы (который тоже входит в "все слова
    # страницы") раздувает смещение на сотни pt, и порог MAX_ASSIGN_DIST ниже
    # перестаёт "склеивать" перенос через разрыв страницы.
    row_info = []
    y_offset = 0.0
    for page in pdf.pages:
        rows = _cluster_rows(page.extract_words(), y_tol=2.5)
        page_rows = []
        for row_words in rows:
            row_words = sorted(row_words, key=lambda w: w["x0"])
            text = " ".join(w["text"] for w in row_words)
            if "Description" in text and "Stems" in text:
                continue  # строка заголовка колонок, а не данные
            if "TOTALS" in text or "AWB" in text:
                continue  # строка итогов/AWB в самом низу страницы, а не данные
            cols = {}
            for w in row_words:
                cols.setdefault(_col_for(w["x0"], _TESSA_COLUMNS), []).append(w)
            # Числовой хвост ищем БЕЗ колонки Label: она печатается правее
            # суммы (напр. "50-60" - разбивка по длинам) и, попав в строку,
            # ломала бы привязку регекса к концу строки (позиция терялась
            # целиком вместе со своей коробкой - реальный случай, инвойс
            # 90823483).
            text_wo_label = " ".join(w["text"] for w in row_words
                                      if _col_for(w["x0"], _TESSA_COLUMNS) != "label")
            tail = _TESSA_TAIL_RE.search(text_wo_label)
            if not tail and not any(k in cols for k in ("boxes", "order", "boxcode", "desc")):
                # Вне определённых колонок и не якорь (например, счётчик
                # страниц вида "1 of 3" или Label-колонка справа) - не влияет
                # на данные, но испортил бы расчёт межстраничного смещения.
                continue
            page_rows.append({"top": row_words[0]["top"], "cols": cols, "tail": tail})
        if not page_rows:
            continue
        min_top = min(r["top"] for r in page_rows)
        max_top = max(r["top"] for r in page_rows)
        for r in page_rows:
            row_info.append({**r, "top": r["top"] - min_top + y_offset})
        y_offset += (max_top - min_top) + 15  # обычный межстрочный интервал

    anchor_idxs = [i for i, r in enumerate(row_info) if r["tail"]]

    # Каждую НЕ-якорную строку (описание сорта/код коробки, перенесённые
    # на соседнюю строку) относим к ближайшему по вертикали якорю.
    assigned_desc = {i: [] for i in anchor_idxs}
    assigned_boxcode = {i: [] for i in anchor_idxs}
    assigned_loc = {i: [] for i in anchor_idxs}
    # Порог отсечения: настоящие переносы описания/кода коробки лежат в
    # пределах пары строк (~5-15pt) от своего якоря. Более далёкие
    # строки - это шапка/подвал страницы (адрес, TOTALS, юр. текст),
    # их не привязываем ни к какому якорю вообще.
    MAX_ASSIGN_DIST = 20
    for j, rj in enumerate(row_info):
        if j in assigned_desc:
            continue
        if not anchor_idxs:
            continue
        nearest = min(anchor_idxs, key=lambda ai: abs(row_info[ai]["top"] - rj["top"]))
        if abs(row_info[nearest]["top"] - rj["top"]) > MAX_ASSIGN_DIST:
            continue
        assigned_desc[nearest].append((rj["top"], [w["text"] for w in rj["cols"].get("desc", [])]))
        assigned_boxcode[nearest].extend(w["text"] for w in rj["cols"].get("boxcode", []))
        assigned_loc[nearest].extend((rj["top"], w["text"]) for w in rj["cols"].get("boxcode", [])
                                     if w["x0"] >= _TESSA_LOC_MIN_X)

    for idx in anchor_idxs:
        r = row_info[idx]
        m = r["tail"]
        is_new_box = bool(r["cols"].get("order"))

        contributions = [(r["top"], [w["text"] for w in r["cols"].get("desc", [])])]
        contributions += assigned_desc[idx]
        contributions.sort(key=lambda t: t[0])
        variety = " ".join(w for _, words_ in contributions for w in words_)

        length_cm = _to_float(m.group("len"))
        stems = _to_int(m.group("stems"))
        price = _to_float(m.group("price"))
        total = float(m.group("total").replace(",", ""))

        if is_new_box:
            boxcode_words = [w["text"] for w in r["cols"].get("boxcode", [])] + assigned_boxcode[idx]
            box_type = next((w for w in boxcode_words if _TESSA_BOXCODE_TOKEN_RE.match(w)), None)
            loc_words = [(r["top"], w["text"]) for w in r["cols"].get("boxcode", [])
                         if w["x0"] >= _TESSA_LOC_MIN_X] + assigned_loc[idx]
            loc = "".join(t for _, t in sorted(loc_words, key=lambda t: t[0]))
            current_box = {
                "box_no": [w["text"] for w in r["cols"].get("order", [])][0],
                "box_type": box_type,
                "box_size": _box_size_from_code(box_type),
                "farm_code": loc if loc.startswith("TESSA") else None,
                "items": [],
            }
            boxes.append(current_box)

        if current_box is not None:
            current_box["items"].append({
                "variety": variety, "length_cm": length_cm, "stems": stems,
                "price": price, "total": total,
            })


    # Если код коробки не распознан, TESSA печатает внизу "Number in Fulls" -
    # полный объём инвойса в полных коробках; недостающее распределяем
    # поровну между коробками без кода. (Инвойс 90823437 раньше шёл этим
    # путём как "1/6", но там код SB = 1/16 - см. BOX_SIZE_CODES.) Это не догадка: цифра берётся из самого документа и по коробкам
    # с известным кодом сходится (проверено на 90821511: 8xQB + 4xHB = 4.0).
    fulls_m = re.search(r"Number in Fulls\s+([\d.]+)", full_text)
    unsized = [b for b in boxes if b["box_size"] is None]
    if fulls_m and unsized:
        known = sum(b["box_size"] or 0 for b in boxes)
        remainder = _to_float(fulls_m.group(1)) - known
        if remainder > 0:
            for b in unsized:
                b["box_size"] = round(remainder / len(unsized), 4)

    total_stems = _to_float(total_m.group(2)) if total_m else None
    total_fob = float(total_m.group(3).replace(",", "")) if total_m else None

    return {
        "source_filename": source_filename,
        "supplier": "TESSA CORP.",
        "mark": mark,
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1).strip() if awb_m else None,
        "hawb": hawb_m.group(1) if hawb_m else None,
        "forwarder": forwarder_m.group(1) if forwarder_m else None,
        "airline": airline_m.group(1).strip() if airline_m else None,
        "destination": country_m.group(1) if country_m else None,
        "boxes": boxes,
        "totals": {"total_stems": total_stems, "total_fob": total_fob},
    }


# ---------------------------------------------------------------------------
# Шаблон "Rosaprima International" (Эквадор/США-инвойсинг, mark = PO #)
# ---------------------------------------------------------------------------
#
# Два вида строк-позиций:
# 1) "прямая" коробка - шапка сама и есть единственная позиция:
#    "<code> <color> <Variety...> <len> x <units> Stem (<cubes> cubes) <mark>
#     <boxes> <boxcode> <units> $<price> $<amount>".
# 2) MIX-коробка - шапка содержит буквально "Mix x <units> Stem (...)" без
#    длины, дальше идёт строка "Vendor:Rosaprima" и под ней N строк реальных
#    сортов: "<code> <color> <Variety...> <len> <bunches> Bun. <st/bun> St/Bun
#    at $<price>" - до следующей шапки. Название сорта в этих строках -
#    ТОЛЬКО variety-часть (без ведущего 3-буквенного кода семейства и цвета) -
#    подтверждено сверкой с примером xlsx от бухгалтера (там код/цвет
#    отброшены в 9 из 10 строк - расхождение по одной строке сочли опечаткой
#    бухгалтера, длину/цифры доверяем PDF, а не ручной табличке).
_ROSAPRIMA_HEADER_RE = re.compile(
    r"^(?P<code>[A-Z]{2,4})\s+(?P<color>[A-Z]{2,4})\s+(?P<name>Mix|.+?)\s+"
    r"(?:(?P<len>\d+)\s+)?x\s+(?P<units>\d+)\s+Stem\s+\([\d.]+\s+cubes\)\s+"
    r"(?P<mark>\S+)\s+(?P<boxes>\d+)\s+(?P<boxcode>[A-Z]{2,3})\s+(?P<units2>\d+)\s+"
    r"\$(?P<price>[\d.]+)\s+\$(?P<amount>[\d.,]+)$"
)
_ROSAPRIMA_ITEM_RE = re.compile(
    r"^(?P<code>[A-Z]{2,4})\s+(?P<color>[A-Z]{2,4})\s+(?P<name>.+?)\s+(?P<len>\d+)\s+"
    r"(?P<bunches>\d+)\s+Bun\.\s+(?P<stbun>\d+)\s+St/Bun\s+at\s+\$(?P<price>[\d.]+)$"
)


def detect_rosaprima(pdf):
    text = pdf.pages[0].extract_text() or ""
    return "ROSAPRIMA" in text.upper()


def parse_rosaprima(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"Invoice #\s*(\d+)", full_text)
    date_m = re.search(r"Invoice Date\s*([\d/]+)", full_text)
    # Метка (mark) - по PO#, подтверждено сверкой с именами файлов, которые
    # прислал закупщик (розаприма ... siri.xlsx <-> "PO # SIRI").
    mark_m = re.search(r"PO #\s*(\S+)", full_text)
    awb_m = re.search(r"Way Bill / Ref #\s*(\S+)", full_text)
    forwarder_m = re.search(r"Ship To\s*Carrier\n.*\n(.+)", full_text)
    total_stems_m = re.search(r"Total stems:\s*(\d+)", full_text)
    total_fob_m = re.search(r"Totals\s+\$([\d,.]+)", full_text)

    boxes = []
    box_counter = 0
    pending_boxes = None  # (pcs, box_type) шапки, ждущей позиций MIX-группы
    is_mix = False

    def _flush_direct(box_type, pcs, variety, length_cm, stems, price, total):
        nonlocal box_counter
        for _ in range(pcs):
            box_counter += 1
            boxes.append({
                "box_no": str(box_counter), "box_type": box_type,
                "box_size": _box_size_from_code(box_type),
                "items": [{"variety": variety, "length_cm": length_cm,
                           "stems": round(stems / pcs, 2), "price": price,
                           "total": round(total / pcs, 2)}],
            })

    def _flush_mix(box_type, pcs, items):
        nonlocal box_counter
        if not items:
            return
        for _ in range(pcs):
            box_counter += 1
            box_items = []
            for it in items:
                stems_per_box = it["stems"] / pcs
                if stems_per_box == int(stems_per_box):
                    stems_per_box = int(stems_per_box)
                box_items.append({
                    "variety": it["variety"], "length_cm": it["length_cm"],
                    "stems": stems_per_box, "price": it["price"],
                    "total": round(it["total"] / pcs, 2),
                })
            boxes.append({
                "box_no": str(box_counter), "box_type": box_type,
                "box_size": _box_size_from_code(box_type), "items": box_items,
            })

    mix_items = []
    mix_box_type = None
    mix_pcs = 1

    for page in pdf.pages:
        text = page.extract_text() or ""
        for line in text.splitlines():
            line = line.strip()
            hm = _ROSAPRIMA_HEADER_RE.match(line)
            if hm:
                if is_mix:
                    _flush_mix(mix_box_type, mix_pcs, mix_items)
                is_mix = (hm.group("name").strip().upper() == "MIX")
                pcs = _to_int(hm.group("boxes")) or 1
                box_type = hm.group("boxcode")
                if is_mix:
                    mix_items, mix_box_type, mix_pcs = [], box_type, pcs
                else:
                    variety = hm.group("name").strip()
                    length_cm = _to_float(hm.group("len"))
                    units = _to_float(hm.group("units"))
                    price = _to_float(hm.group("price"))
                    amount = _to_float(hm.group("amount"))
                    _flush_direct(box_type, pcs, variety, length_cm, units, price, amount)
                continue
            if is_mix:
                im = _ROSAPRIMA_ITEM_RE.match(line)
                if im:
                    stems = (_to_int(im.group("bunches")) or 0) * (_to_int(im.group("stbun")) or 0)
                    mix_items.append({
                        "variety": im.group("name").strip(),
                        "length_cm": _to_float(im.group("len")),
                        "stems": stems, "price": _to_float(im.group("price")),
                        "total": round(stems * _to_float(im.group("price")), 2),
                    })
    if is_mix:
        _flush_mix(mix_box_type, mix_pcs, mix_items)


    return {
        "source_filename": source_filename,
        "supplier": "Rosaprima International",
        "mark": mark_m.group(1) if mark_m else None,
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1) if awb_m else None,
        "hawb": None,
        "forwarder": forwarder_m.group(1).strip() if forwarder_m else None,
        "airline": None,
        "destination": None,
        "boxes": boxes,
        "totals": {
            "total_stems": _to_float(total_stems_m.group(1)) if total_stems_m else None,
            "total_fob": _to_float(total_fob_m.group(1)) if total_fob_m else None,
        },
    }


# ---------------------------------------------------------------------------
# Шаблон "Rosaprima Cia. Ltda." - SHIPPING INVOICE самой эквадорской фермы
# (не путать с "Rosaprima International" выше: другой документ и вёрстка).
# Первый образец - RU_968408 (партия AWB 369-1151 1964, 2026-09-21).
#
# Строка коробки: "<всего> <с> <по> <тип> ASSORTED <пачек> ...", под ней
# строки сортов: "<Сорт> <длина> <пачек> <ст/пачка> <стеблей> ... <сумма>".
# Жирный шрифт в PDF сделан двойной печатью символов со сдвигом 1-2 pt, а
# колонки цен перекрывают друг друга - поэтому строки позиций читаем после
# dedupe_chars, а цену за стебель считаем как сумма / стебли (колонка "Unit
# Price" в тексте нечитаема). Сумма сверяется с GRAND TOTAL.
# ---------------------------------------------------------------------------

# Размер коробки по коду. JL = 0.5 подтверждено AWB (у POLINA "TOTAL IN FULL
# 22.813") и файлом закупщика за 30.09. HB/QB - стандартные половина/четверть.
# Прочие коды из подвала инвойса (JX, JB, OB, JS, QS, QN, TH) пока не
# встречались - размер не угадываем, на странице проверки будет "?".
# JB = 0.5: в HAWB 157-0008 3845 / 235-7841 9795 (метка VIKA, 15-16.09)
# коробка JB Rosaprima стоит строкой "ROSA PRIMA CIA. LTDA. 0.50".
_ROSAPRIMA_EC_BOX_SIZE = {"JL": 0.5, "JB": 0.5, "HB": 0.5, "QB": 0.25}

_ROSAPRIMA_EC_BOX_RE = re.compile(
    r"^(?P<total>\d+)\s+(?P<from>\d+)\s+(?P<to>\d+)\s+(?P<type>[A-Z]{2})\s+\D"
)
_ROSAPRIMA_EC_ITEM_RE = re.compile(
    r"^(?P<name>[A-Za-z][A-Za-z .'&/-]*?)\s+(?P<len>\d+)\s+(?P<bun>\d+)\s+(?P<stbun>\d+)\s+"
    r"(?P<stems>\d+)\s+.*?(?P<ext>[\d,]+\.\d{2})$"
)


def _whole(x):
    return int(x) if x == int(x) else round(x, 2)


def detect_rosaprima_ec(pdf):
    text = (pdf.pages[0].extract_text() or "").upper()
    return "ROSAPRIMA CIA" in text and "SHIPPING INVOICE" in text


def parse_rosaprima_ec(pdf, source_filename=""):
    raw = "\n".join((p.extract_text() or "") for p in pdf.pages)
    # Двойные буквы ("Carrier") dedupe склеивает - шапку читаем из обычного
    # текста, dedupe только для таблицы позиций.
    clean = "\n".join((p.dedupe_chars(tolerance=3).extract_text() or "") for p in pdf.pages)

    mark_m = re.search(r"BOXES MARKED AS\s*:\s*(\S+)", raw)
    date_m = re.search(r"([A-Z][a-z]{2}/\d{1,2}/\d{4})\s+(\d+)", raw)
    carrier_m = re.search(r"Carrier\s*:.*\n(\S+)\s+(\d{6,})\s+(\S+)\s+(\S+)", raw)
    totals_m = re.search(r"Total Stems=(\d+).*Total Invoice=([\d,.]+)", clean)

    boxes = []
    group = None  # текущая строка коробки: {"pcs", "type", "items"}

    def flush():
        if not group or not group["items"]:
            return
        pcs = group["pcs"]
        for _ in range(pcs):
            boxes.append({
                "box_no": str(len(boxes) + 1), "box_type": group["type"],
                "box_size": _ROSAPRIMA_EC_BOX_SIZE.get(group["type"]),
                # Несколько одинаковых коробок одной строкой - делим поровну
                # (как у Rosaprima International).
                "items": [dict(it, stems=_whole(it["stems"] / pcs),
                               total=round(it["total"] / pcs, 2)) for it in group["items"]],
            })

    for line in clean.splitlines():
        line = line.strip()
        bm = _ROSAPRIMA_EC_BOX_RE.match(line)
        if bm:
            flush()
            group = {"pcs": _to_int(bm.group("total")) or 1, "type": bm.group("type"), "items": []}
            continue
        if line.startswith("Total ") or group is None:
            continue
        im = _ROSAPRIMA_EC_ITEM_RE.match(line)
        if im:
            stems = _to_int(im.group("stems"))
            total = float(im.group("ext").replace(",", ""))
            group["items"].append({
                "variety": im.group("name").strip().upper(),
                "length_cm": _to_float(im.group("len")),
                "stems": stems,
                "price": round(total / stems, 4) if stems else None,
                "total": total,
            })
    flush()

    return {
        "source_filename": source_filename,
        "supplier": "Rosaprima Cia. Ltda.",
        "mark": mark_m.group(1) if mark_m else None,
        "invoice_no": date_m.group(2) if date_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        # "AWB M" в этом инвойсе - внутренний номер, с накладной партии не
        # совпадает; номер AWB берётся из самой авианакладной.
        "awb": None,
        "hawb": carrier_m.group(3) if carrier_m else None,
        "forwarder": carrier_m.group(4) if carrier_m else None,
        "airline": carrier_m.group(1) if carrier_m else None,
        "destination": None,
        "boxes": boxes,
        "totals": {
            "total_stems": _to_float(totals_m.group(1)) if totals_m else None,
            "total_fob": _to_float(totals_m.group(2).replace(",", "")) if totals_m else None,
        },
    }


# ---------------------------------------------------------------------------
# Шаблон "CeresFarms" (Эквадор) - табличная сетка длин 30/40/.../90 см:
# в строке позиции количество бунчей стоит в ОДНОЙ из колонок-длин, остальные
# пустые (кроме служебной колонки "30", где в реальных файлах печатается
# литеральный "0" даже когда позиция пуста - её игнорируем, если есть другое
# непустое значение в строке). Длина определяется по x0 самого числа:
# берём максимальную подпись колонки "<=" координате значения (не midpoint -
# midpoint даёт неверный столбец, проверено на реальном примере, где значения
# у левого края следующей колонки всё ещё принадлежат текущей).
# ---------------------------------------------------------------------------

_CERES_COLUMNS = [
    ("box_no", -1, 26),
    ("box_type", 26, 65),
    ("variety", 65, 184),
    ("unxb", 184, 222),
    ("grade", 222, 360),
    ("stems", 360, 417),
    ("price", 417, 448),
    ("total", 448, 9995),
]
_CERES_GRADE_LABELS = [(228, 30), (244, 40), (261, 50), (279, 60), (297, 70), (315, 80), (332, 90)]


def _ceres_grade_length(x0):
    length = None
    for lx, val in _CERES_GRADE_LABELS:
        if x0 >= lx:
            length = val
        else:
            break
    return length


def detect_ceresfarms(pdf):
    text = (pdf.pages[0].extract_text() or "").upper()
    return "CERESFARMS" in text


def parse_ceresfarms(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"INVOICE:\d+\s+(\d+)", full_text)
    date_m = re.search(r"DATE(\d{2}/\d{2}/\d{4})", full_text)
    mark_m = re.search(r"CONSIGNEE(\S+)", full_text)
    awb_m = re.search(r"M\.A\.W\.B\s+(\S+)", full_text)
    hawb_m = re.search(r"H\.A\.W\.B\s+(\S+)", full_text)
    forwarder_m = re.search(r"CARRIER\s+(\S+)", full_text)

    boxes_by_no = {}
    box_seq = []

    for page in pdf.pages:
        words = page.extract_words()
        rows = _cluster_rows(words, y_tol=2.0)
        for row_words in rows:
            rowdict = {}
            grade_words = []
            for w in sorted(row_words, key=lambda w: w["x0"]):
                c = _col_for(w["x0"], _CERES_COLUMNS)
                if c == "grade":
                    grade_words.append((w["x0"], w["text"]))
                else:
                    rowdict.setdefault(c, []).append(w["text"])
            rowdict = {k: " ".join(v) for k, v in rowdict.items()}

            box_no = rowdict.get("box_no", "").strip()
            box_type = rowdict.get("box_type", "").strip()
            variety = rowdict.get("variety", "").strip()
            stems = _to_float(rowdict.get("stems"))
            if not box_no.isdigit() or not variety or stems is None:
                continue  # шапка/итоги/пустая строка

            grade_words = [(x0, t) for x0, t in grade_words if t.strip() != "0"]
            if not grade_words:
                continue
            length_cm = _ceres_grade_length(grade_words[0][0])

            price = _to_float(rowdict.get("price"))
            total = _to_float(rowdict.get("total"))

            if box_no not in boxes_by_no:
                boxes_by_no[box_no] = {
                    "box_no": box_no, "box_type": box_type,
                    "box_size": _box_size_from_code(box_type), "items": [],
                }
                box_seq.append(box_no)
            boxes_by_no[box_no]["items"].append({
                "variety": variety, "length_cm": length_cm, "stems": stems,
                "price": price, "total": total,
            })

    boxes = [boxes_by_no[n] for n in box_seq]

    total_stems_m = re.search(r"TOT\.\s*STEMS\s+([\d,]+)", full_text)
    total_fob_m = re.search(r"TOTAL\s+USD\s*\n?([\d.]+)", full_text)

    return {
        "source_filename": source_filename,
        "supplier": "CeresFarms",
        "mark": mark_m.group(1) if mark_m else None,
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1) if awb_m else None,
        "hawb": hawb_m.group(1) if hawb_m else None,
        "forwarder": forwarder_m.group(1) if forwarder_m else None,
        "airline": None,
        "destination": None,
        "boxes": boxes,
        "totals": {
            "total_stems": float(total_stems_m.group(1).replace(",", "")) if total_stems_m else None,
            "total_fob": _to_float(total_fob_m.group(1)) if total_fob_m else None,
        },
    }


# ---------------------------------------------------------------------------
# Шаблон "Utopia Farms" (Эквадор, гипсофила) - "Grade" тут не длина в см, а
# вес пучка в граммах (суффикс "GR"), поэтому length_cm сознательно не
# заполняем (в отличие от cm-форматов) - подставить туда граммы было бы
# неверно, а сверенного правила пересчёта в см нет (единственный имеющийся
# пример от бухгалтера расходится с PDF и по этой причине не может быть
# ориентиром - см. историю разработки). Каждая позиция состоит из ДВУХ строк:
# "якорной" (с Boxes/box_type/F.B.E.) и "чистой" строки-повтора под ней с
# тем же Grade/Units/Stems/Price/Total, но БЕЗ Boxes - берём название сорта
# и цифры из чистой строки как более короткое и однозначное, если она есть.
# ---------------------------------------------------------------------------

_UTOPIA_ANCHOR_RE = re.compile(
    r"^(?P<desc>.+?)\s+\d+\s*(?:GR|CM)\s+(?P<boxes>\d+)\s+(?P<boxtype>[HQFE])\s+[\d.]+\s+"
    r"\d+\s*BC\s+(?P<stems>\d+)\s+(?P<price>[\d.]+)\s+(?P<total>[\d,.]+)$"
)
_UTOPIA_CLEAN_RE = re.compile(
    r"^(?P<name>.+?)\s+\d+\s*(?:GR|CM)\s+\d+\s*BC\s+(?P<stems>\d+)\s+(?P<price>[\d.]+)\s+(?P<total>[\d,.]+)$"
)


def detect_utopia(pdf):
    text = (pdf.pages[0].extract_text() or "").upper()
    return "UTOPIA FARMS" in text


def parse_utopia(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"\n(\d+)\s+\d{1,2}/\d{1,2}/\d{4}\n", full_text)
    date_m = re.search(r"\n\d+\s+(\d{1,2}/\d{1,2}/\d{4})\n", full_text)
    mark_m = re.search(r"\*\*Box Marking:\s*(\S+)", full_text) \
        or re.search(r"Ship To:.*\n(\S+)", full_text)
    awb_m = re.search(r"Air Waybill:\s*([\d\- ]+\d)", full_text)
    hawb_m = re.search(r"HAWB:\s*(\S+)", full_text)
    airline_m = re.search(r"Airline:\s*(\S+)", full_text)
    forwarder_m = re.search(r"Agent:\s*(\S+)", full_text)
    total_m = re.search(r"Total Stems:\s*(\d+)\s+USD:\s*([\d.]+)", full_text)

    boxes = []
    box_counter = 0
    lines = full_text.splitlines()

    for i, line in enumerate(lines):
        line = line.strip()
        am = _UTOPIA_ANCHOR_RE.match(line)
        if not am:
            continue
        pcs = _to_int(am.group("boxes")) or 1
        box_type = am.group("boxtype")
        variety = am.group("desc").strip()
        stems = _to_float(am.group("stems"))
        price = _to_float(am.group("price"))
        total = _to_float(am.group("total").replace(",", ""))

        next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
        cm = _UTOPIA_CLEAN_RE.match(next_line)
        if cm:
            variety = cm.group("name").strip()
            stems = _to_float(cm.group("stems"))
            price = _to_float(cm.group("price"))
            total = _to_float(cm.group("total").replace(",", ""))

        for _ in range(pcs):
            box_counter += 1
            stems_per_box = stems / pcs
            if stems_per_box == int(stems_per_box):
                stems_per_box = int(stems_per_box)
            boxes.append({
                "box_no": str(box_counter), "box_type": box_type,
                "box_size": _box_size_from_code(box_type),
                "items": [{"variety": variety, "length_cm": None,
                           "stems": stems_per_box, "price": price,
                           "total": round(total / pcs, 2)}],
            })


    return {
        "source_filename": source_filename,
        "supplier": "Utopia Farms",
        "mark": mark_m.group(1) if mark_m else None,
        "invoice_no": invoice_no_m.group(1) if invoice_no_m else None,
        "invoice_date": date_m.group(1) if date_m else None,
        "awb": awb_m.group(1).strip() if awb_m else None,
        "hawb": hawb_m.group(1) if hawb_m else None,
        "forwarder": forwarder_m.group(1) if forwarder_m else None,
        "airline": airline_m.group(1) if airline_m else None,
        "destination": None,
        "boxes": boxes,
        "totals": {
            "total_stems": _to_float(total_m.group(1)) if total_m else None,
            "total_fob": _to_float(total_m.group(2)) if total_m else None,
        },
    }


TEMPLATES = [
    ("matiz_roses", detect_matiz, parse_matiz),
    ("gardaexport", detect_gardaexport, parse_gardaexport),
    ("florsani", detect_florsani, parse_florsani),
    ("rosas_corazon", detect_rosas_corazon, parse_rosas_corazon),
    # v2 проверяем ПЕРЕД старым: detect_monterosas ловит только слово
    # "MONTEROSAS" и матчит оба макета, а v2 требует ещё колонку VARIETIES.
    ("star_roses", detect_star_roses, parse_star_roses),
    ("monterosas_v2", detect_monterosas_v2, parse_monterosas_v2),
    ("monterosas", detect_monterosas, parse_monterosas),
    ("tessa", detect_tessa, parse_tessa),
    # Эквадорский SHIPPING INVOICE тоже содержит слово "Rosaprima" -
    # проверяем его раньше шаблона Rosaprima International.
    ("rosaprima_ec", detect_rosaprima_ec, parse_rosaprima_ec),
    ("rosaprima", detect_rosaprima, parse_rosaprima),
    ("ceresfarms", detect_ceresfarms, parse_ceresfarms),
    ("utopia", detect_utopia, parse_utopia),
]


def parse_invoice_pdf(path):
    with pdfplumber.open(path) as pdf:
        for name, detect, parse in TEMPLATES:
            if detect(pdf):
                return parse(pdf, source_filename=path), name
        if not any((page.extract_text() or "").strip() for page in pdf.pages):
            # Скан/фотография без текстового слоя (реальный случай: HAWB,
            # присланный картинкой). Разобрать нечего - но сообщение должно
            # объяснять причину, а не говорить про "неизвестный шаблон".
            raise ValueError(
                "В этом PDF нет текста - похоже, это скан или фотография. "
                "Нужен файл с текстовым слоем (как его выгружает ферма), "
                "либо данные придётся ввести вручную."
            )
        raise ValueError(
            "Не удалось распознать шаблон поставщика в этом импортном PDF. "
            "Нужно добавить новый шаблон разбора (см. import_parser.py -> TEMPLATES)."
        )


# ---------------------------------------------------------------------------
# Авианакладная (AWB) - ОДИН документ на всю партию, а не инвойс фермы
# ---------------------------------------------------------------------------
#
# Отсюда берутся вес и ставка, которых нет ни в одном фермерском инвойсе
# (см. историю: раньше их вводили руками). Реальный образец 2026-09-09
# (LAN CARGO, AWB 145-9999 8334):
#   строка веса:  "42 841 K 910 3.25 2,957.50 CONSOLIDATION FLOWERS"
#                  места брутто    платный ставка сумма
#   блок Handling Information - разбивка по меткам:
#     "DAMIR - IRIS FLOWERS / TRUCK" (шапка метки)
#     "1.00 = ECUANROS finca 1" (ферма = полных коробок)
#     "BXS: 3.25 PCS: 7" (итог метки: полных коробок / мест)
# Сверено с разбором инвойсов: по метке DAMIR совпало точно (7 коробок, 3.25
# полных). Брутто-вес даётся ТОЛЬКО общий на всю накладную - по меткам его
# делим пропорционально числу мест (решение пользователя 2026-09-09; это же
# согласуется с методом закупщика "средний вес коробки = брутто / места").
_AWB_WEIGHT_RE = re.compile(
    r"^(?P<pieces>\d+)\s+(?P<gross>[\d,]+)\s+K\s+(?P<chargeable>[\d,]+)\s+"
    r"(?P<rate>[\d.]+)\s+(?P<total>[\d,.]+)\s", re.M
)
_AWB_NO_RE = re.compile(r"\b(\d{3}-\d{4}\s?\d{4})\b")
# Итог к оплате - именно "Total Prepaid" (фрахт ПЛЮС "Total Other Charges Due
# Carrier"), а не одна строка Prepaid Weight Charge (правка закупщика
# 2026-09-10). В образце: 2 957.50 фрахт + 43.00 прочие = 3 000.50.
_AWB_TOTAL_PREPAID_RE = re.compile(r"Total Prepaid.*\n\s*([\d,.]+)")
_AWB_OTHER_CARRIER_RE = re.compile(r"Total Other Charges Due Carrier\s*\n\s*([\d,.]+)")
_AWB_MARK_HEADER_RE = re.compile(r"^(?P<mark>[A-Z][A-Z0-9 .]*?)\s+-\s+\S.*$")
_AWB_MARK_TOTAL_RE = re.compile(r"^BXS:\s*(?P<bxs>[\d.]+)\s+PCS:\s*(?P<pcs>\d+)\s*$")


_AWB_BOX_LABEL_RE = re.compile(r"BOX LABEL:\s*(.+)")
_AWB_AWC_RE = re.compile(r"AWC:\s*([\d,.]+)")
_AWB_TOTAL_FULL_RE = re.compile(r"TOTAL IN FULL:\s*([\d.]+)")


def _awb_house_pages(pdf, mark_names):
    """Накладные по меткам (house AWB) на страницах после первой: у каждой
    метки свои места, брутто и платный вес, тариф и сбор AWC. Закупщик
    считает вес метки именно по ним (правка 2026-09-28, AWB 369-1151 1964:
    DAMIR 415 кг, POLINA 1024, VADIM 621 - а не 2060 кг поровну по местам).

    mark_names - метки из блока Handling Information первой страницы; подпись
    на house-странице бывает длиннее ("VADIM- IRIS FLOWERS")."""
    houses = {}
    for page in pdf.pages[1:]:
        text = page.extract_text() or ""
        label_m = _AWB_BOX_LABEL_RE.search(text)
        weight_m = _AWB_WEIGHT_RE.search(text)
        if not label_m or not weight_m:
            continue
        label = label_m.group(1).strip().upper()
        mark = next((m for m in sorted(mark_names, key=len, reverse=True)
                     if label.startswith(m.upper())), None)
        if mark is None:
            mark = re.match(r"[A-Z0-9]+", label).group(0) if re.match(r"[A-Z0-9]+", label) else label
        awc_m = _AWB_AWC_RE.search(text)
        full_m = _AWB_TOTAL_FULL_RE.search(text)
        weight_charge = _to_float(weight_m.group("total").replace(",", ""))
        awc = _to_float(awc_m.group(1).replace(",", "")) if awc_m else 0.0
        houses[mark] = {
            "pieces": _to_int(weight_m.group("pieces")),
            "gross_weight": _to_float(weight_m.group("gross").replace(",", "")),
            "chargeable_weight": _to_float(weight_m.group("chargeable").replace(",", "")),
            "rate_per_kg": _to_float(weight_m.group("rate")),
            "weight_charge": weight_charge,
            "other_charges": awc or None,
            "total_awb": round((weight_charge or 0) + (awc or 0), 2),
            "full_boxes": _to_float(full_m.group(1)) if full_m else None,
        }
    return houses


def detect_awb(pdf):
    text = (pdf.pages[0].extract_text() or "").upper()
    return "AIR WAYBILL" in text and "SHIPPER" in text


def _awb_handling_lines(pdf):
    """Строки блока "Handling Information" в правильном порядке чтения.

    Блок бывает свёрстан В ДВЕ КОЛОНКИ, и тогда обычный extract_text()
    склеивает левую и правую колонку в одну строку: например шапка метки
    "DAMIR - IRIS FLOWERS / TRUCK" оказывается в одной строке с "1.00 =
    SOLERA FARMS", а итог другой метки "BXS: 32.00 PCS: 70" - в одной строке
    с фермой из первой. Из-за этого метки читались неверно (реальный случай
    2026-09-10: из трёх меток находилась одна). Поэтому режем блок на колонки
    по X и читаем их подряд: сначала левую сверху вниз, потом правую.
    """
    page = pdf.pages[0]
    words = page.extract_words()
    rows = _cluster_rows(words, y_tol=2.5)

    start_top = end_top = None
    for row in rows:
        texts = [w["text"] for w in row]
        if start_top is None and "Handling" in texts:
            start_top = row[0]["top"]
        elif start_top is not None and end_top is None and "Gross" in texts:
            # Шапка блока веса ("No Of Gross Kg Rate Class ...") - конец
            # блока Handling Information. Важно закончить именно на ней:
            # её широкая вёрстка заполняет промежутки между колонками и
            # ломает определение границ колонок ниже.
            end_top = row[0]["top"]
    if start_top is None:
        return []
    block = [w for w in words if w["top"] > start_top
             and (end_top is None or w["top"] < end_top)]
    if not block:
        return []

    # Границы колонок ищем как вертикальные "просветы" - диапазоны X, которые
    # НЕ пересекает ни одно слово блока (по разрывам между началами слов
    # надёжно не выходит: внутри строки разрывы бывают шире, чем между
    # колонками).
    left = int(min(w["x0"] for w in block))
    right = int(max(w["x1"] for w in block)) + 1
    occupied = [False] * (right - left + 1)
    for w in block:
        for x in range(int(w["x0"]) - left, min(int(w["x1"]) - left + 1, len(occupied))):
            occupied[x] = True

    boundaries = []
    run_start = None
    for i, busy in enumerate(occupied):
        if not busy:
            run_start = i if run_start is None else run_start
        elif run_start is not None:
            # 4pt хватает: просвет должен быть пустым по ВСЕЙ высоте
            # блока, поэтому случайный разрыв внутри строки сюда не попадёт.
            if i - run_start >= 4:
                boundaries.append(left + (run_start + i) / 2)
            run_start = None

    edges = [left - 1] + boundaries + [right + 1]
    lines = []
    for start, end in zip(edges, edges[1:]):
        col_words = [w for w in block if start <= w["x0"] < end]
        for row in _cluster_rows(col_words, y_tol=2.5):
            lines.append(" ".join(w["text"] for w in sorted(row, key=lambda w: w["x0"])))
    return lines


def parse_awb(pdf, source_filename=""):
    text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    weight_m = _AWB_WEIGHT_RE.search(text)
    awb_no_m = _AWB_NO_RE.search(text)
    prepaid_m = _AWB_TOTAL_PREPAID_RE.search(text)
    other_m = _AWB_OTHER_CARRIER_RE.search(text)

    marks = {}
    last_header = None
    for line in _awb_handling_lines(pdf):
        line = line.strip()
        total_m = _AWB_MARK_TOTAL_RE.match(line)
        if total_m and last_header:
            marks[last_header] = {
                "pieces": _to_int(total_m.group("pcs")),
                "full_boxes": _to_float(total_m.group("bxs")),
            }
            last_header = None
            continue
        header_m = _AWB_MARK_HEADER_RE.match(line)
        # Строки ферм внутри блока ("1.00 = ECUANROS") тоже начинаются с
        # цифры/заглавных - отсекаем их по знаку "=" и по ведущей цифре.
        if header_m and "=" not in line and not line[0].isdigit():
            last_header = header_m.group("mark").strip()

    return {
        "source_filename": source_filename,
        "awb_no": awb_no_m.group(1) if awb_no_m else None,
        "pieces": _to_int(weight_m.group("pieces")) if weight_m else None,
        "gross_weight": _to_float(weight_m.group("gross").replace(",", "")) if weight_m else None,
        "chargeable_weight": _to_float(weight_m.group("chargeable").replace(",", "")) if weight_m else None,
        "rate_per_kg": _to_float(weight_m.group("rate")) if weight_m else None,
        "weight_charge": _to_float(weight_m.group("total").replace(",", "")) if weight_m else None,
        "other_charges": _to_float(other_m.group(1).replace(",", "")) if other_m else None,
        # Итоговая сумма к оплате по накладной (фрахт + прочие сборы).
        "total_awb": (_to_float(prepaid_m.group(1).replace(",", "")) if prepaid_m
                       else (_to_float(weight_m.group("total").replace(",", "")) if weight_m else None)),
        "marks": marks,
        "houses": _awb_house_pages(pdf, marks.keys()),
    }


# ---------------------------------------------------------------------------
# HAWB форвардера (Fresh Solutions Cargo) - накладная на одну метку одним
# рейсом. В Handling Information - сколько полных коробок какой фермы летит
# ("2.00 = POSITANO FARMS S.A.S."); по этим строкам раздел "Поставки"
# раскладывает коробки из инвойсов по фактическим рейсам. Сумма перевозки
# в HAWB не печатается ("AS AGREED") - считается по весу и ставке за кг.
# Первые образцы: HAWB 2400/2411/2433/2466/9821, метка VIKA, 22-25.09.
# ---------------------------------------------------------------------------

_AIRPORTS = {"SHEREMETYEVO": "SVO", "VNUKOVO": "VKO", "DOMODEDOVO": "DME", "PULKOVO": "LED"}
_ORIGIN_COUNTRY = {"UIO": "ecuador", "GYE": "ecuador", "BOG": "colombia", "MDE": "colombia"}
_HAWB_NUMBERS_RE = re.compile(r"^(\d{3}-\d{4}\s?\d{4})\s+(\d{3}\s?\d{4}\s?\d{4})\s*$", re.M)
_HAWB_LINE_RE = re.compile(r"^(\d+(?:\.\d+)?)\s*=\s*(.+?)\s*$", re.M)
_HAWB_WEIGHT_RE = re.compile(r"^(\d+)\s+([\d,.]+)\s+K\s+([\d,.]+)\s", re.M)
_HAWB_DATE_RE = re.compile(r"(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)-(\d{2})-(\d{4})\b")
_MONTHS = {m: i + 1 for i, m in enumerate(
    ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"])}


def detect_forwarder_hawb(pdf):
    text = (pdf.pages[0].extract_text() or "").upper()
    return "AIR WAYBILL" in text and "NOTIFY TO: MARK" in text and "HANDLING INFORMATION" in text


def parse_forwarder_hawb(pdf, source_filename=""):
    text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    lines = [l.strip() for l in text.splitlines()]

    numbers_m = _HAWB_NUMBERS_RE.search(text)
    mark_m = re.search(r"NOTIFY TO:\s*MARK\s+(\S+)", text)
    weight_m = _HAWB_WEIGHT_RE.search(text)
    full_m = re.search(r"TOTAL IN FULL:\s*([\d.]+)", text)
    date_m = _HAWB_DATE_RE.search(text)
    airline_m = re.search(r"(QATAR AIRWAYS|TURKISH AIRLINES|[A-Z]+ AIR(?:WAYS|LINES))", text)

    airport = None
    for i, line in enumerate(lines):
        if line.startswith("Airport of Destination") and i + 1 < len(lines):
            dest = lines[i + 1].upper()
            code_m = re.search(r"\(([A-Z]{3})\)", dest)
            airport = code_m.group(1) if code_m else next(
                (code for name, code in _AIRPORTS.items() if name in dest), dest.split()[0] if dest else None)
            break

    # Строки ферм - между "Handling Information" и шапкой веса.
    growers = []
    try:
        start = next(i for i, l in enumerate(lines) if l.startswith("Handling Information"))
        end = next(i for i, l in enumerate(lines) if i > start and l.startswith("No Of"))
    except StopIteration:
        start, end = 0, 0
    for line in lines[start + 1:end]:
        m = _HAWB_LINE_RE.match(line)
        if m:
            growers.append({"name": m.group(2), "full_boxes": float(m.group(1))})

    origin = lines[0].upper() if lines and re.fullmatch(r"[A-Z]{3}", lines[0].upper()) else None
    flight_date = None
    if date_m:
        flight_date = "%s-%02d-%s" % (date_m.group(3), _MONTHS.get(date_m.group(1), 0), date_m.group(2))

    return {
        "source_filename": source_filename,
        "mawb": numbers_m.group(1) if numbers_m else None,
        "hawb": numbers_m.group(2) if numbers_m else None,
        "mark": mark_m.group(1) if mark_m else None,
        "airline": airline_m.group(1) if airline_m else None,
        "origin": origin,
        "origin_country": _ORIGIN_COUNTRY.get(origin),
        "airport": airport,
        "flight_date": flight_date,
        "pieces": _to_int(weight_m.group(1)) if weight_m else None,
        "gross_weight": _to_float(weight_m.group(2).replace(",", "")) if weight_m else None,
        "chargeable_weight": _to_float(weight_m.group(3).replace(",", "")) if weight_m else None,
        "total_full": _to_float(full_m.group(1)) if full_m else None,
        "growers": growers,
    }


def parse_forwarder_hawb_pdf(path):
    """HAWB форвардера или None, если это другой документ."""
    with pdfplumber.open(path) as pdf:
        if not detect_forwarder_hawb(pdf):
            return None
        return parse_forwarder_hawb(pdf, source_filename=path)


def parse_awb_pdf(path):
    """Возвращает разобранную авианакладную или None, если это не AWB."""
    with pdfplumber.open(path) as pdf:
        if not detect_awb(pdf):
            return None
        return parse_awb(pdf, source_filename=path)


# ---------------------------------------------------------------------------
# Шаблон "Astoria Export" (.xls) - форвардер-консолидатор, собирающий коробки
# НЕСКОЛЬКИХ плантаций (Plantation) в одном рейсе; строка-шапка коробки несёт
# Full Boxes/Pieces/Packing x FB, за ней 1+ строк-сортов. Если плантация сама
# уже прислала "MIX" одной строкой (без разбивки по сортам) - берём её как есть,
# MIX-правило это не меняет (уже применено источником).
# ---------------------------------------------------------------------------

def _xls_find_label(sh, label, max_rows=60):
    label = label.strip().lower()
    for r in range(min(max_rows, sh.nrows)):
        for c in range(sh.ncols):
            v = sh.cell_value(r, c)
            if isinstance(v, str) and v.strip().lower() == label:
                return r, c
    return None, None


def _xls_scan(sh, row, anchor_col, width=1):
    """Значение может лежать на 1-2 колонки левее/правее подписи шапки
    из-за объединённых ячеек - берём первую непустую в окне."""
    if anchor_col is None or row is None:
        return ""
    lo, hi = max(0, anchor_col - width), min(sh.ncols - 1, anchor_col + width)
    for c in range(lo, hi + 1):
        v = sh.cell_value(row, c)
        if v != "":
            return v
    return ""


def _xls_value_right_of_label(sh, row, label_col, max_gap=20):
    """Для полей шапки инвойса (CLIENT/INVOICE No/Date) значение может стоять
    далеко правее подписи из-за широкой объединённой ячейки подписи (или на
    строку ниже, как у CLIENT) - ищем первую непустую ячейку правее по той же
    строке, а если не нашли - на строке ниже."""
    if row is None:
        return ""
    for search_row in (row, row + 1):
        if search_row >= sh.nrows:
            continue
        for c in range(label_col + 1, min(sh.ncols, label_col + 1 + max_gap)):
            v = sh.cell_value(search_row, c)
            if v != "":
                return v
    return ""


def detect_astoria_xls(wb):
    sh = wb.sheet_by_index(0)
    for r in range(min(15, sh.nrows)):
        for c in range(sh.ncols):
            v = sh.cell_value(r, c)
            if isinstance(v, str) and "ASTORIA EXPORT" in v.upper():
                return True
    return False


def parse_astoria_xls(wb, source_filename=""):
    sh = wb.sheet_by_index(0)

    _, plantation_c = _xls_find_label(sh, "Plantation")
    header_r, type_c = _xls_find_label(sh, "Type")
    _, variety_c = _xls_find_label(sh, "Variety")
    _, length_c = _xls_find_label(sh, "Length")
    _, fb_c = _xls_find_label(sh, "Full Boxes")
    _, pieces_c = _xls_find_label(sh, "Pieces")
    _, price_c = _xls_find_label(sh, "Price per Items (US$)")
    total_r, total_c = _xls_find_label(sh, "Total (US$)")
    stems_r, stems_c = _xls_find_label(sh, "Stems")
    invoice_no_r, invoice_no_c = _xls_find_label(sh, "INVOICE No")
    client_r, client_c = _xls_find_label(sh, "CLIENT")
    date_r, date_c = _xls_find_label(sh, "Date :")

    data_start = max(header_r or 0, total_r or 0, stems_r or 0) + 1

    # Страна отправления пишется в шапке отдельной ячейкой. Для КОЛУМБИИ
    # действуют свои правила (закупщик, 2026-09-10): позиции не сводятся в
    # MIX - всё выводится построчно, и коробку набивают одним сортом.
    origin_country = None
    for r in range(min(data_start, sh.nrows)):
        for c in range(sh.ncols):
            value = str(sh.cell_value(r, c)).strip().upper()
            if value in ("COLOMBIA", "ECUADOR", "KENYA"):
                origin_country = value
                break
        if origin_country:
            break
    is_colombia = origin_country == "COLOMBIA"

    boxes = []
    box_counter = 0
    pieces = None
    box_size = None  # Full Boxes / Pieces - вместимость ОДНОЙ физической коробки
    current_items = None  # позиции текущей группы коробок (до раздачи по штукам)

    def _box_from(items, box_size):
        nonlocal box_counter
        box_counter += 1
        return {"box_no": str(box_counter), "box_type": None, "box_size": box_size,
                "farm": items[0].get("farm"), "product": items[0].get("product"),
                "no_merge": is_colombia, "items": items}

    def _flush(items, pieces, box_size):
        if not items or not pieces:
            return
        if is_colombia:
            # Колумбия: коробку набивают ОДНИМ сортом целиком (см. эталон
            # закупщика "Колумбия 13.09 (AST)_AGATA MOS.xls": 1600 стеблей на
            # 4 места - это PINK 400, WHITE 400 и RED 400+400, а не по кусочку
            # каждого сорта в каждой коробке). Раскладываем сорта по коробкам
            # подряд; если ровно не делится - откатываемся на равное деление.
            per_box = sum(it["stems"] for it in items) / pieces
            packed, current, filled = [], [], 0.0
            for it in items:
                left = it["stems"]
                while left > 0:
                    take = min(left, per_box - filled)
                    share = take / it["stems"]
                    current.append({
                        "variety": it["variety"], "length_cm": it["length_cm"],
                        "grade_text": it.get("grade_text"),
                        "stems": int(take) if take == int(take) else take,
                        "price": it["price"],
                        "total": round(it["total"] * share, 2) if it["total"] is not None else None,
                        "farm": it.get("farm"), "product": it.get("product"),
                    })
                    left -= take
                    filled += take
                    if filled >= per_box - 1e-9:
                        packed.append(current)
                        current, filled = [], 0.0
            if current:
                packed.append(current)
            if len(packed) == pieces:
                for group in packed:
                    boxes.append(_box_from(group, box_size))
                return

        for _ in range(pieces):
            box_items = []
            for it in items:
                stems_per_box = it["stems"] / pieces
                if stems_per_box == int(stems_per_box):
                    stems_per_box = int(stems_per_box)
                box_items.append({
                    "variety": it["variety"], "length_cm": it["length_cm"],
                    "grade_text": it.get("grade_text"),
                    "stems": stems_per_box, "price": it["price"],
                    "total": round(it["total"] / pieces, 2) if it["total"] is not None else None,
                    "farm": it.get("farm"), "product": it.get("product"),
                })
            boxes.append(_box_from(box_items, box_size))

    for r in range(data_start, sh.nrows):
        fb_val = _xls_scan(sh, r, fb_c)
        if isinstance(fb_val, float) and fb_val > 0:
            _flush(current_items, pieces, box_size)
            current_items = []
            pieces = _to_int(_xls_scan(sh, r, pieces_c)) or 1
            # Вместимость физической коробки = Full Boxes / Pieces (см. пример
            # от закупщика: Full Boxes=1, Pieces=4 -> 4 коробки по 0.25).
            box_size = round(fb_val / pieces, 4) if pieces else None
            continue

        variety = str(_xls_scan(sh, r, variety_c)).strip()
        type_ = str(_xls_scan(sh, r, type_c)).strip()
        if not variety and not type_:
            continue
        length_val = _xls_scan(sh, r, length_c)
        length_cm = length_val if isinstance(length_val, float) else _to_float(length_val)
        stems_total = _xls_scan(sh, r, stems_c)
        price = _xls_scan(sh, r, price_c) or None
        total = _xls_scan(sh, r, total_c) or None
        if stems_total == "" or current_items is None:
            continue

        farm = str(_xls_scan(sh, r, plantation_c)).strip()
        current_items.append({
            # Колонки Type/Variety/Length у Astoria - это ровно PRODUCT /
            # VARIETY / GRADE итогового файла (подтверждено эталоном
            # закупщика: PRODUCT "SPRAY CARNATION", VARIETY "MIX",
            # GRADE "SELECT"). Раньше Type приклеивался к названию сорта.
            "variety": variety, "product": " ".join(type_.split()) or None,
            "length_cm": length_cm,
            # Длина бывает нечисловой ("FANCY"/"SELECT"/"1000GR") - в этом
            # случае в GRADE идёт текст как есть.
            "grade_text": None if length_cm is not None else (str(length_val).strip() or None),
            "stems": float(stems_total), "price": price, "total": total, "farm": farm,
        })

    _flush(current_items, pieces, box_size)


    # Итоговая строка "TOTAL" - ищем отдельно, т.к. она ниже таблицы позиций.
    total_row_r, _ = _xls_find_label(sh, "TOTAL")
    grand_stems = grand_total = None
    if total_row_r is not None:
        v = _xls_scan(sh, total_row_r, stems_c)
        grand_stems = v if isinstance(v, float) else _to_float(v)
        v = _xls_scan(sh, total_row_r, total_c)
        grand_total = v if isinstance(v, float) else _to_float(v)

    # Транспорт у Astoria указан В САМОМ инвойсе (строки "TRANSPORT" и
    # "WEIGHT: N KG"), отдельная авианакладная для него не нужна - закупщик
    # считает ставку как TRANSPORT / вес (правка 2026-09-10). Итог "TOTAL" в
    # инвойсе включает транспорт и предохлаждение, поэтому FOB по цветам
    # получаем вычитанием - иначе сверка "посчитано против напечатано"
    # ложно ругалась бы на расхождение.
    def _amount_by_label(label):
        row, _c = _xls_find_label(sh, label, max_rows=sh.nrows)
        if row is None:
            return None
        v = _xls_scan(sh, row, total_c)
        return v if isinstance(v, float) else _to_float(v)

    transport_cost = _amount_by_label("TRANSPORT")
    precooling = _amount_by_label("PRECOOLING") or 0

    weight_kg = None
    for r in range(sh.nrows):
        for c in range(sh.ncols):
            m = re.match(r"WEIGHT:\s*([\d.,]+)\s*KG",
                          str(sh.cell_value(r, c)).strip(), re.I)
            if m:
                weight_kg = _to_float(m.group(1))
                break
        if weight_kg is not None:
            break

    if grand_total is not None and transport_cost:
        grand_total = round(grand_total - transport_cost - precooling, 2)

    awb_no = _xls_value_right_of_label(sh, *_xls_find_label(sh, "AWB :", max_rows=sh.nrows)) \
        if _xls_find_label(sh, "AWB :", max_rows=sh.nrows)[0] is not None else None
    forwarder = _xls_value_right_of_label(sh, *_xls_find_label(sh, "FORWARDER :", max_rows=sh.nrows)) \
        if _xls_find_label(sh, "FORWARDER :", max_rows=sh.nrows)[0] is not None else None
    airline = _xls_value_right_of_label(sh, *_xls_find_label(sh, "AIRLAINE :", max_rows=sh.nrows)) \
        if _xls_find_label(sh, "AIRLAINE :", max_rows=sh.nrows)[0] is not None else None

    transport = None
    if transport_cost or weight_kg:
        transport = {"cost_usd": transport_cost, "weight_kg": weight_kg,
                      "awb_no": str(awb_no).strip() if awb_no else None}

    return {
        "source_filename": source_filename,
        "supplier": "Astoria Export",
        "mark": str(_xls_value_right_of_label(sh, client_r, client_c)).strip()
                if client_r is not None else None,
        "invoice_no": str(_xls_value_right_of_label(sh, invoice_no_r, invoice_no_c)).strip()
                      if invoice_no_r is not None else None,
        "invoice_date": str(_xls_value_right_of_label(sh, date_r, date_c)).strip()
                         if date_r is not None else None,
        "awb": str(awb_no).strip() if awb_no else None,
        "hawb": None,
        "forwarder": str(forwarder).strip() if forwarder else "Astoria Export",
        "airline": str(airline).strip() if airline else None,
        "destination": origin_country,
        "transport": transport,
        "boxes": boxes,
        "totals": {"total_stems": grand_stems, "total_fob": grand_total},
    }


# ---------------------------------------------------------------------------
# Шаблон "брокер" (.xls) - брокер в Эквадоре присылает инвойсы мелких ферм
# уже в формате factura: FARM / FARM INVOICE / PRODUCT / BOX / BOX SIZE /
# VARIETY / GRADE / TOTAL STEMS / UNIT PRICE / TOTAL USD, метка - в OBS.
# Первые образцы - INV-15802 (Qualisa Service) и INV-15825 (Ecoroses),
# партия VIKA 21-25.09. Строка с BOX=N - это N одинаковых коробок (стебли и
# сумма на все N), строка без BOX - ещё одна позиция той же коробки.
# ---------------------------------------------------------------------------

def _broker_header(sh):
    for r in range(min(40, sh.nrows)):
        cells = {str(sh.cell_value(r, c)).strip().upper(): c for c in range(sh.ncols)
                 if isinstance(sh.cell_value(r, c), str) and sh.cell_value(r, c).strip()}
        if "FARM INVOICE" in cells and "VARIETY" in cells:
            return r, cells
    return None, None


def detect_broker_xls(wb):
    return _broker_header(wb.sheet_by_index(0))[0] is not None


def parse_broker_xls(wb, source_filename=""):
    sh = wb.sheet_by_index(0)
    header_r, cols = _broker_header(sh)

    def cell(r, name, width=0):
        c = cols.get(name)
        if c is None:
            return ""
        for cc in range(max(0, c - width), min(sh.ncols, c + 1)):
            v = sh.cell_value(r, cc)
            if v != "":
                return v
        return ""

    def head(label):
        r, c = _xls_find_label(sh, label, max_rows=header_r)
        return _xls_value_right_of_label(sh, r, c) if r is not None else ""

    invoice_date = head("DATE")
    if isinstance(invoice_date, float):
        invoice_date = xlrd.xldate_as_datetime(invoice_date, wb.datemode).strftime("%Y-%m-%d")

    boxes, group, marks = [], None, Counter()

    def flush():
        if not group:
            return
        pcs = group["pcs"]
        for _ in range(pcs):
            boxes.append({
                "box_no": str(len(boxes) + 1), "box_type": None,
                "box_size": group["size"], "farm": group["farm"], "product": group["product"],
                "items": [dict(it, stems=_whole(it["stems"] / pcs), total=round(it["total"] / pcs, 2))
                          for it in group["items"]],
            })

    for r in range(header_r + 1, sh.nrows):
        variety = str(cell(r, "VARIETY")).strip()
        if not variety:
            continue  # итоговые строки под таблицей
        grade = cell(r, "GRADE")
        length = _to_float(str(grade)) if str(grade).strip() else None
        item = {
            "variety": variety,
            "length_cm": length,
            "grade_text": None if length is not None else (str(grade).strip() or None),
            "stems": float(cell(r, "TOTAL STEMS") or 0),
            "price": _to_float(str(cell(r, "UNIT PRICE"))),
            "total": float(cell(r, "TOTAL USD") or 0),
        }
        mark = str(cell(r, "OBS", width=1)).strip()
        if mark:
            marks[mark] += 1
        pcs = cell(r, "BOX")
        if pcs != "":
            flush()
            group = {"pcs": int(pcs) or 1, "size": _to_float(str(cell(r, "BOX SIZE"))),
                     "farm": str(cell(r, "FARM")).strip() or None,
                     "product": str(cell(r, "PRODUCT")).strip() or None, "items": [item]}
        elif group is not None:
            group["items"].append(item)
    flush()

    consignee = head("CONSIGNEE")
    farm_invoices = sorted({str(sh.cell_value(r, cols["FARM INVOICE"])).strip()
                            for r in range(header_r + 1, sh.nrows)
                            if str(sh.cell_value(r, cols["FARM INVOICE"])).strip()})
    return {
        "source_filename": source_filename,
        "supplier": "Брокер: " + ", ".join(sorted({b["farm"] for b in boxes if b.get("farm")})),
        "mark": (marks.most_common(1)[0][0] if marks else (str(consignee).strip() or None)),
        "invoice_no": ", ".join(farm_invoices) or None,
        "invoice_date": invoice_date or None,
        # "M.A.W.B" у брокера бывает номером инвойса фермы (INV-15825) -
        # номер рейса берётся из HAWB форвардера.
        "awb": None,
        "hawb": None,
        "forwarder": str(head("CARGO AGENCY")).strip() or None,
        "airline": None,
        "destination": None,
        "boxes": boxes,
        "totals": {"total_stems": sum(it["stems"] for b in boxes for it in b["items"]),
                   "total_fob": round(sum(it["total"] for b in boxes for it in b["items"]), 2)},
    }


XLS_TEMPLATES = [
    ("astoria_export", detect_astoria_xls, parse_astoria_xls),
    ("broker_xls", detect_broker_xls, parse_broker_xls),
]


def parse_invoice_file(path):
    """Единая точка входа: определяет PDF/XLS по расширению и диспетчит
    в соответствующий набор шаблонов."""
    if path.lower().endswith((".xls", ".xlsx")):
        wb = xlrd.open_workbook(path)
        for name, detect, parse in XLS_TEMPLATES:
            if detect(wb):
                return parse(wb, source_filename=path), name
        raise ValueError(
            "Не удалось распознать шаблон поставщика в этом импортном XLS. "
            "Нужно добавить новый шаблон разбора (см. import_parser.py -> XLS_TEMPLATES)."
        )
    return parse_invoice_pdf(path)
