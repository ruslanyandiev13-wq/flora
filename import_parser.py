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

Правило MIX (см. ТЗ, раздел 3): если внутри одной коробки встречаются разные
length_cm - у ВСЕХ позиций этой коробки length_cm принудительно заменяется на
максимальную длину, встретившуюся в коробке. Суммы/кол-во/цена не трогаются.
"""
import re
import pdfplumber
import xlrd

# Доля коробки по первой букве кода типа коробки (F=Full/H=Half/Q=Quarter/E=Eighth).
BOX_SIZE_TABLE = {"F": 1.0, "H": 0.5, "Q": 0.25, "E": 0.125}


def _box_size_from_code(box_type_code):
    if not box_type_code:
        return None
    letter = box_type_code.strip()[:1].upper()
    return BOX_SIZE_TABLE.get(letter)


def apply_mix_rule(boxes):
    """Мутирует boxes на месте: в каждой коробке заменяет length_cm всех
    позиций на максимальную длину, встретившуюся в этой коробке."""
    for box in boxes:
        lengths = [it["length_cm"] for it in box["items"] if it.get("length_cm") is not None]
        if not lengths:
            continue
        max_len = max(lengths)
        for it in box["items"]:
            if it.get("length_cm") is not None:
                it["length_cm"] = max_len
    return boxes


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
    r"(?P<box_type>[HQFE]\S*(?:\s+S-\d+)?)\s+"
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

    apply_mix_rule(boxes)

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
    ("box_no", -1, 50),
    ("tb", 50, 144.5),
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

            box_no_tok = rowdict.get("box_no", "").strip()
            if re.match(r"^\d{2}$", box_no_tok):
                pending_box_no = box_no_tok

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
                current_box = {
                    "box_no": pending_box_no,
                    "box_type": tb,
                    "box_size": BOX_SIZE_TABLE.get(tb),
                    "items": [],
                }
                boxes.append(current_box)
                pending_box_no = None

            if current_box is not None:
                current_box["items"].append({
                    "variety": f"{color} {variety}".strip(), "length_cm": length_cm,
                    "stems": stems, "price": price, "total": total,
                    "bunches": bunches,
                })

    apply_mix_rule(boxes)

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

    apply_mix_rule(boxes)

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

    apply_mix_rule(boxes)

    footer_m = re.search(
        r"^([\d.]+)\s*I\s*I\s*([\d,]+)I\s*TOTALF\.O\.B\.VALUE:I\s*\$([\d,.]+)$",
        full_text, re.M)
    total_stems = float(footer_m.group(2).replace(",", "")) if footer_m else None
    total_fob = float(footer_m.group(3).replace(",", "")) if footer_m else None

    return {
        "source_filename": source_filename,
        "supplier": "Rosas del Corazon (Saftec)",
        "mark": mark_m.group(1) if mark_m else None,
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

    apply_mix_rule(boxes)

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
# Шаблон "TESSA CORP." (Эквадор)
# ---------------------------------------------------------------------------
#
# Самый неровный из всех форматов: описание сорта переносится СВОБОДНО между
# 2-3 текстовыми строками (сколько слов уместилось по ширине), без фиксированной
# привязки к колонке - у пиксельных координат тоже нет надёжной колонки (см.
# анализ в истории разработки). Числа при этом всегда собраны в одну "якорную"
# строку. Полностью зерна к зерну восстановить название сорта в общем случае
# нельзя без разметки PDF; берём лучшее доступное приближение (строки
# непосредственно до/после якоря), гарантируя точность числовых полей
# (stems/price/total), которые для 1С единственно важны.
_TESSA_ANCHOR_RE = re.compile(
    r"^(?:(?P<boxes>\d+)\s+(?P<order>\d+)\s+)?(?P<prefix>.*?)\s*"
    r"(?P<len>\d+)\s+(?P<bun>\d+)\s+(?P<stems>\d+)\s+\$(?P<price>[\d.]+)\s+\$(?P<total>[\d,.]+)$"
)
_TESSA_BOXTYPE_RE = re.compile(r"^([A-Z]{2,4})\s+\S*-\S*$")


def detect_tessa(pdf):
    text = pdf.pages[0].extract_text() or ""
    return "TESSA CORP" in text.upper()


def parse_tessa(pdf, source_filename=""):
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"Invoice Number\s+(\S+)", full_text)
    date_m = re.search(r"Invoice Date\s+([\d/]+)", full_text)
    mark_m = re.search(r"SHIP CUSTOMER\n\S+\s+(\S+)", full_text)
    awb_m = re.search(r"AWB\s+(\S[\d\- ]*\d)", full_text)
    hawb_m = re.search(r"HAWB\s+(\S+)", full_text)
    airline_m = re.search(r"Airline\s+(.+)", full_text)
    forwarder_m = re.search(r"Cargo Agency\s+(\S+)", full_text)
    country_m = re.search(r"Country Final Destination\s+(\S+)", full_text)
    total_m = re.search(r"TOTALS\s+(\d+)\s+(\d+)\s+\$([\d,.]+)", full_text)

    boxes = []
    box_counter = 0
    current_box = None
    current_box_type = None

    for page in pdf.pages:
        lines = (page.extract_text() or "").splitlines()
        for i, line in enumerate(lines):
            m = _TESSA_ANCHOR_RE.match(line.strip())
            if not m or m.group("stems") is None:
                continue
            is_new_box = m.group("order") is not None

            # Строка box-type/label ("HB TESSA-") сдвигает before/after на 1,
            # и её "хвост" (напр. "P" от "TESSA-P") может прилипнуть к строке
            # после якоря, склеившись с началом описания СЛЕДУЮЩЕЙ позиции
            # (напр. "ALSTRO P") - такую строку для описания не используем.
            before_idx = i - 1
            if is_new_box and i >= 2 and _TESSA_BOXTYPE_RE.match(lines[i - 1].strip()):
                current_box_type = _TESSA_BOXTYPE_RE.match(lines[i - 1].strip()).group(1)
                before_idx = i - 2

            # Симметрично: если перед якорем была box-type/label строка, то
            # СРАЗУ после якоря идёт её "хвост" (склеенный с началом описания
            # следующей позиции, напр. "ALSTRO P") - пропускаем на строку дальше.
            after_idx = i + 2 if before_idx == i - 2 else i + 1

            before_txt = lines[before_idx].strip() if 0 <= before_idx < len(lines) else ""
            after_txt = lines[after_idx].strip() if after_idx < len(lines) else ""
            if _TESSA_ANCHOR_RE.match(before_txt) or _TESSA_BOXTYPE_RE.match(before_txt):
                before_txt = ""
            if _TESSA_ANCHOR_RE.match(after_txt) or after_txt.startswith(("TOTALS", "AWB")):
                after_txt = ""

            variety = " ".join(w for w in
                                f"{before_txt} {m.group('prefix')} {after_txt}".split())

            length_cm = _to_float(m.group("len"))
            bun = _to_int(m.group("bun"))
            stems = _to_int(m.group("stems"))
            price = _to_float(m.group("price"))
            total = float(m.group("total").replace(",", ""))

            if is_new_box:
                box_counter += 1
                current_box = {
                    "box_no": m.group("order"),
                    "box_type": current_box_type,
                    "box_size": _box_size_from_code(current_box_type),
                    "items": [],
                }
                boxes.append(current_box)

            if current_box is not None:
                current_box["items"].append({
                    "variety": variety, "length_cm": length_cm, "stems": stems,
                    "price": price, "total": total,
                })

    apply_mix_rule(boxes)

    total_stems = _to_float(total_m.group(2)) if total_m else None
    total_fob = float(total_m.group(3).replace(",", "")) if total_m else None

    return {
        "source_filename": source_filename,
        "supplier": "TESSA CORP.",
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

    apply_mix_rule(boxes)

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
    apply_mix_rule(boxes)

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

    apply_mix_rule(boxes)

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
    ("monterosas", detect_monterosas, parse_monterosas),
    ("tessa", detect_tessa, parse_tessa),
    ("rosaprima", detect_rosaprima, parse_rosaprima),
    ("ceresfarms", detect_ceresfarms, parse_ceresfarms),
    ("utopia", detect_utopia, parse_utopia),
]


def parse_invoice_pdf(path):
    with pdfplumber.open(path) as pdf:
        for name, detect, parse in TEMPLATES:
            if detect(pdf):
                return parse(pdf, source_filename=path), name
        raise ValueError(
            "Не удалось распознать шаблон поставщика в этом импортном PDF. "
            "Нужно добавить новый шаблон разбора (см. import_parser.py -> TEMPLATES)."
        )


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

    boxes = []
    box_counter = 0
    pieces = None
    box_size = None  # Full Boxes / Pieces - вместимость ОДНОЙ физической коробки
    current_items = None  # позиции текущей группы коробок (до раздачи по штукам)

    def _flush(items, pieces, box_size):
        nonlocal box_counter
        if not items or not pieces:
            return
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
                    "total": round(it["total"] / pieces, 2) if it["total"] is not None else None,
                })
            boxes.append({"box_no": str(box_counter), "box_type": None, "box_size": box_size,
                           "farm": items[0].get("farm"), "items": box_items})

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
            "variety": f"{type_} {variety}".strip(), "length_cm": length_cm,
            "stems": float(stems_total), "price": price, "total": total, "farm": farm,
        })

    _flush(current_items, pieces, box_size)

    apply_mix_rule(boxes)

    # Итоговая строка "TOTAL" - ищем отдельно, т.к. она ниже таблицы позиций.
    total_row_r, _ = _xls_find_label(sh, "TOTAL")
    grand_stems = grand_total = None
    if total_row_r is not None:
        v = _xls_scan(sh, total_row_r, stems_c)
        grand_stems = v if isinstance(v, float) else _to_float(v)
        v = _xls_scan(sh, total_row_r, total_c)
        grand_total = v if isinstance(v, float) else _to_float(v)

    return {
        "source_filename": source_filename,
        "supplier": "Astoria Export",
        "mark": str(_xls_value_right_of_label(sh, client_r, client_c)).strip()
                if client_r is not None else None,
        "invoice_no": str(_xls_value_right_of_label(sh, invoice_no_r, invoice_no_c)).strip()
                      if invoice_no_r is not None else None,
        "invoice_date": str(_xls_value_right_of_label(sh, date_r, date_c)).strip()
                         if date_r is not None else None,
        "awb": None,
        "hawb": None,
        "forwarder": "Astoria Export",
        "airline": None,
        "destination": None,
        "boxes": boxes,
        "totals": {"total_stems": grand_stems, "total_fob": grand_total},
    }


XLS_TEMPLATES = [
    ("astoria_export", detect_astoria_xls, parse_astoria_xls),
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
