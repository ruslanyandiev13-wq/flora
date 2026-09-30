"""
Раздел «Поставки» (метки Москвы): что именно приедет в каждой доставке.

Проблема (разбор партии VIKA 21-25.09): номер AWB в инвойсе фермы - это
план. Коробки одного инвойса разъезжаются по разным рейсам (инвойс TESSA
90832511 - тремя рейсами за два дня). Факт виден только в HAWB форвардера:
в Handling Information написано, сколько полных коробок какой фермы летит
("2.00 = POSITANO FARMS S.A.S.").

Как считается (всё заново из документов метки при каждом открытии):
1. Каждая коробка инвойса получает ферму: у TESSA - по коду Loc (TESSA-P ->
   POSITANO, справочник grower_aliases), у остальных - по самому поставщику.
2. HAWB идут по дате вылета. Для каждой строки "доля = ферма" подбираются
   свободные коробки этой фермы, дающие ровно эту долю, - самые ранние по
   дате инвойса. Не хватает коробок - строка попадает в «Нет инвойса».
3. Одинаковые по размеру коробки одной фермы от одной даты, но с разным
   содержимым, неотличимы по документам - такие подборы помечаются
   «Нужно выбрать»; закупщик может вручную привязать коробку к рейсу
   (delivery_pins), ручная привязка применяется раньше автоматики.
4. Поставка - рейсы, прилетевшие в один аэропорт в один день (закупщик,
   2026-09-28). В HAWB напечатана дата вылета, группируем по ней.
5. Перевозка = платный вес HAWB × ставка за кг (Эквадор 8.1, Колумбия 8 -
   settings), делится между коробками по долям.
"""
import datetime
import io
import json
import re
from collections import Counter

import xlwt

import db
import import_parser
from import_combine import combine_by_mark
from import_xls_writer import build_combined_factura_xls

UNITS = 16  # доли коробок считаем в 1/16 полной (самая малая - SB)

AIRPORT_NAMES = {"SVO": "Шереметьево", "VKO": "Внуково", "DME": "Домодедово", "LED": "Пулково"}
COUNTRY_NAMES = {"ecuador": "Эквадор", "colombia": "Колумбия"}

# Ферма коробки, если её не видно в самой коробке: поставщик = одна ферма.
TEMPLATE_GROWER = {
    "matiz_roses": "MATIZ", "monterosas": "MONTEROSAS", "monterosas_v2": "MONTEROSAS",
    "rosas_corazon": "ROSAS DEL CORAZON", "gardaexport": "GARDAEXPORT", "florsani": "FLORSANI",
    "rosaprima": "ROSAPRIMA", "rosaprima_ec": "ROSAPRIMA", "ceresfarms": "CERES",
    "utopia": "UTOPIA",
}
# Форма юрлица в конце названия ("POSITANO FARMS S.A.S.", "... CIA LTDA").
# Только через пробел: код TESSA-S (Solera) или TESSA-A - это не "S.A.".
_LEGAL_TAIL_RE = re.compile(r"(\s+(S\.?\s?A\.?(\s?S\.?)?|CIA\.?|LTDA\.?|SCC|INC\.?|LLC|C\.?\s?L\.?))+\s*$")


# --- фермы -------------------------------------------------------------------

def normalize_name(name):
    name = re.sub(r"\(.*?\)", " ", (name or "").upper()).strip()
    name = _LEGAL_TAIL_RE.sub("", name)
    return " ".join(re.sub(r"[^A-Z0-9]+", " ", name).split())


class Growers:
    """Имя фермы по любому написанию: код TESSA, строка HAWB, поставщик."""

    def __init__(self, aliases):
        self.map = {normalize_name(a["alias"]): a["grower"] for a in aliases}

    def resolve(self, name):
        n = normalize_name(name)
        if not n:
            return None
        if n in self.map:
            return self.map[n]
        # HAWB обрезает длинные названия ("EL CAMPANARIO DE"), а инвойсы
        # дописывают лишнее - совпадение по целым словам с начала.
        matches = [(len(a), g) for a, g in self.map.items()
                   if n.startswith(a + " ") or a.startswith(n + " ")]
        return max(matches)[1] if matches else n


def known_growers():
    """Все фермы, которые можно выбрать: из справочника написаний и все
    поставщики «Импорта»."""
    growers = Growers(db.get_grower_aliases())
    names = {a["grower"] for a in db.get_grower_aliases()}
    names |= {growers.resolve(g) for g in TEMPLATE_GROWER.values()}
    return sorted(n for n in names if n)


def _size_from_code(template, box_type):
    """Размер коробки, если парсер на момент загрузки его не знал (новые коды
    добавляются в import_parser позже - перезагружать инвойс не нужно)."""
    if not box_type:
        return None
    if template == "rosaprima_ec":
        return import_parser._ROSAPRIMA_EC_BOX_SIZE.get(box_type.strip().upper())
    return import_parser._box_size_from_code(box_type)


def box_grower(growers, template, data, box):
    if box.get("farm_code"):
        return growers.resolve(box["farm_code"])
    if box.get("farm"):
        return growers.resolve(box["farm"])
    if template == "tessa":
        return "TESSA (код фермы не найден)"
    return growers.resolve(TEMPLATE_GROWER.get(template) or data.get("supplier") or template)


# --- даты --------------------------------------------------------------------

def parse_date(value, template=None):
    """Дата инвойса в ISO - только для порядка (раньше инвойс - раньше рейс).
    02/03/2026 у TESSA - месяц/день, у остальных - день/месяц."""
    value = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%b/%d/%Y", "%d.%m.%Y", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            pass
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", value)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        month_first = b > 12 or (a <= 12 and template == "tessa")
        month, day = (a, b) if month_first else (b, a)
        try:
            return datetime.date(y, month, day).isoformat()
        except ValueError:
            return None
    return None


_WEEKDAYS = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]


def fmt_day(iso):
    """"26.09, сб" - для заголовков доставок."""
    try:
        d = datetime.date.fromisoformat(iso)
        return f"{d.strftime('%d.%m')}, {_WEEKDAYS[d.weekday()]}"
    except (TypeError, ValueError):
        return iso or "—"


def fmt_date(iso):
    try:
        return datetime.date.fromisoformat(iso).strftime("%d.%m.%Y")
    except (TypeError, ValueError):
        return iso or "—"


# --- подбор коробок ----------------------------------------------------------

def _earliest_subset(units, target):
    """Индексы коробок (в порядке очереди), дающих ровно target, - самые
    ранние из возможных. None, если так не набрать."""
    n = len(units)
    reach = [set() for _ in range(n + 1)]
    reach[n] = {0}
    for i in range(n - 1, -1, -1):
        reach[i] = reach[i + 1] | {s + units[i] for s in reach[i + 1] if s + units[i] <= target}
    if target not in reach[0]:
        return None
    picked, rem = [], target
    for i in range(n):
        if rem == 0:
            break
        if units[i] <= rem and (rem - units[i]) in reach[i + 1]:
            picked.append(i)
            rem -= units[i]
    return picked


def _best_subset(units, target):
    """Точный подбор, а если не выходит - наибольшая доля не больше target."""
    exact = _earliest_subset(units, target)
    if exact is not None:
        return exact
    for t in range(target - 1, 0, -1):
        found = _earliest_subset(units, t)
        if found is not None:
            return found
    return []


def batch_label(batch, batches=None):
    """"BESST · 28.09.2026" - метка + дата первой загрузки. Если в один день
    завели две поставки одной метки, у второй - номер: "BESST · 28.09.2026 (2)"."""
    date = fmt_date((batch.get("created_at") or "")[:10])
    label = f"{batch['mark']} · {date}"
    if batches:
        same = sorted(b["id"] for b in batches
                      if b["mark"] == batch["mark"] and (b.get("created_at") or "")[:10] == (batch.get("created_at") or "")[:10])
        if len(same) > 1 and batch["id"] in same:
            n = same.index(batch["id"]) + 1
            if n > 1:
                label += f" ({n})"
    return label


def build(batch_id):
    """Полная картина по поставке (метка + дата загрузки): доставки, рейсы,
    коробки и что требует внимания. Считается заново из документов."""
    batch = db.get_delivery_batch(batch_id)
    mark = batch["mark"] if batch else None
    growers = Growers(db.get_grower_aliases())
    pins = db.get_delivery_pins()
    docs = db.get_delivery_docs(batch_id)
    rates = {c: float(db.get_setting(f"delivery_rate_kg_{c}", d))
             for c, d in (("ecuador", 8.1), ("colombia", 8))}
    transit_days = int(float(db.get_setting("delivery_transit_days", 1)))

    invoices, boxes = [], []
    for doc in docs:
        if doc["kind"] != "invoice":
            continue
        data = json.loads(doc["data"])
        inv = {"id": doc["id"], "filename": doc["filename"], "template": doc["template"],
               "data": data, "date": parse_date(data.get("invoice_date"), doc["template"]),
               "uploaded_at": doc["uploaded_at"]}
        invoices.append(inv)
        for idx, b in enumerate(data.get("boxes") or []):
            size = b.get("box_size") or _size_from_code(doc["template"], b.get("box_type"))
            boxes.append({
                "key": f"{doc['id']}:{idx}", "doc_id": doc["id"], "idx": idx, "invoice": inv,
                "grower": box_grower(growers, doc["template"], data, b),
                "size": size, "units": round(size * UNITS) if size else None,
                "box": b, "hawb_id": None, "pinned": False,
                "signature": tuple((it.get("variety"), it.get("length_cm"), it.get("stems"))
                                   for it in b.get("items") or []),
            })
    boxes.sort(key=lambda x: (x["invoice"]["date"] or "9999", x["doc_id"], x["idx"]))
    by_key = {b["key"]: b for b in boxes}

    hawbs = []
    for doc in docs:
        if doc["kind"] != "hawb":
            continue
        data = json.loads(doc["data"])
        country = data.get("origin_country") or "ecuador"
        weight = data.get("chargeable_weight") or data.get("gross_weight") or 0
        hawbs.append({
            "id": doc["id"], "filename": doc["filename"], "data": data, "country": country,
            "rate": rates.get(country, rates["ecuador"]), "weight": weight,
            "cost": round(weight * rates.get(country, rates["ecuador"]), 2),
            "lines": [{"name": g["name"], "grower": growers.resolve(g["name"]),
                       "full": g["full_boxes"], "need": round(g["full_boxes"] * UNITS),
                       "boxes": [], "missing": 0, "ambiguous": []}
                      for g in data.get("growers") or []],
            "extra": [],  # вручную привязанные коробки фермы, которой нет в строках HAWB
        })
    hawbs.sort(key=lambda h: (h["data"].get("flight_date") or "", h["data"].get("mawb") or ""))
    by_hawb = {h["id"]: h for h in hawbs}

    # 1. Ручные привязки.
    for key, hawb_id in pins.items():
        box, hawb = by_key.get(key), by_hawb.get(hawb_id)
        if not box or not hawb:
            continue
        box["hawb_id"], box["pinned"] = hawb_id, True
        line = next((l for l in hawb["lines"] if l["grower"] == box["grower"]), None)
        (line["boxes"] if line else hawb["extra"]).append(box)

    # 2. Автоматика: рейсы по порядку, в каждом - ранние коробки фермы.
    for hawb in hawbs:
        for line in hawb["lines"]:
            target = line["need"] - sum(b["units"] or 0 for b in line["boxes"])
            if target <= 0:
                continue
            cands = [b for b in boxes if b["hawb_id"] is None and b["grower"] == line["grower"]
                     and b["units"]]
            picked = [cands[i] for i in _best_subset([c["units"] for c in cands], target)]
            for b in picked:
                b["hawb_id"] = hawb["id"]
                line["boxes"].append(b)
            line["missing"] = target - sum(b["units"] for b in picked)
            # Неотличимые по документам коробки: тот же размер и дата,
            # другое содержимое - какая из них в рейсе, по HAWB не понять.
            for b in picked:
                for c in cands:
                    if (c["hawb_id"] is None and c["units"] == b["units"]
                            and c["invoice"]["date"] == b["invoice"]["date"]
                            and c["signature"] != b["signature"]):
                        line["ambiguous"].append((b, c))

    for hawb in hawbs:
        hawb["boxes"] = [b for l in hawb["lines"] for b in l["boxes"]] + hawb["extra"]
        units_total = sum(l["need"] for l in hawb["lines"]) or sum(b["units"] or 0 for b in hawb["boxes"])
        for b in hawb["boxes"]:
            b["cost"] = round(hawb["cost"] * (b["units"] or 0) / units_total, 2) if units_total else None
        hawb["pieces_mismatch"] = (hawb["data"].get("pieces") is not None
                                   and not any(l["missing"] for l in hawb["lines"])
                                   and hawb["data"]["pieces"] != len(hawb["boxes"]))

    # Коробки, которые могут быть одной и той же, - по одной паре на
    # коробку, чтобы предупреждение не разрасталось.
    for hawb in hawbs:
        for line in hawb["lines"]:
            seen, pairs = set(), []
            for a, b in line["ambiguous"]:
                if a["key"] not in seen:
                    seen.add(a["key"])
                    pairs.append((a, b))
            line["ambiguous"] = pairs

    # Подсказка «это та же ферма»: строка HAWB без инвойса и неотгруженные
    # коробки фермы, которой нет ни в одной HAWB, - скорее всего, форвардер
    # просто пишет её иначе (ROSA PRIMA CIA. LTDA. = ROSAPRIMA).
    hawb_growers = {l["grower"] for h in hawbs for l in h["lines"]}
    orphans = {}
    for b in boxes:
        if b["hawb_id"] is None and b["grower"] not in hawb_growers and b["units"]:
            orphans.setdefault(b["grower"], []).append(b)
    for hawb in hawbs:
        for line in hawb["lines"]:
            line["suggest"] = [g for g, bxs in sorted(orphans.items())
                               if line["missing"] and
                               _earliest_subset([x["units"] for x in bxs], line["missing"]) is not None]

    # 3. Поставки: один аэропорт, один день.
    deliveries = {}
    for hawb in hawbs:
        d = hawb["data"]
        key = f"{d.get('flight_date') or 'без-даты'}_{d.get('airport') or 'XXX'}"
        dl = deliveries.setdefault(key, {"id": key, "date": d.get("flight_date"),
                                         "airport": d.get("airport"), "hawbs": []})
        dl["hawbs"].append(hawb)
    for dl in deliveries.values():
        dl["airport_name"] = AIRPORT_NAMES.get(dl["airport"], dl["airport"])
        dl["pieces"] = sum(h["data"].get("pieces") or 0 for h in dl["hawbs"])
        dl["boxes_found"] = sum(len(h["boxes"]) for h in dl["hawbs"])
        dl["weight"] = round(sum(h["weight"] for h in dl["hawbs"]), 2)
        dl["cost"] = round(sum(h["cost"] for h in dl["hawbs"]), 2)
        dl["stems"] = sum(it.get("stems") or 0 for h in dl["hawbs"] for b in h["boxes"]
                          for it in b["box"].get("items") or [])
        dl["missing"] = [(h, l) for h in dl["hawbs"] for l in h["lines"] if l["missing"]]
        dl["ambiguous"] = [(h, l) for h in dl["hawbs"] for l in h["lines"] if l["ambiguous"]]
        # Дата доставки на склад = вылет + дней в пути (settings).
        try:
            flight = datetime.date.fromisoformat(dl["date"])
            dl["arrival"] = (flight + datetime.timedelta(days=transit_days)).isoformat()
        except (TypeError, ValueError):
            dl["arrival"] = None
        dl["arrived"] = bool(dl["arrival"]) and dl["arrival"] < datetime.date.today().isoformat()
        dl["ready"] = not dl["missing"] and dl["boxes_found"] == dl["pieces"]

    not_shipped = [b for b in boxes if b["hawb_id"] is None]
    not_shipped_by_grower = {}
    for b in not_shipped:
        not_shipped_by_grower.setdefault(b["grower"], []).append(b)
    return {
        "not_shipped_by_grower": sorted(not_shipped_by_grower.items()),
        "hawb_growers": hawb_growers, "orphan_growers": sorted(orphans),
        "transit_days": transit_days, "known_growers": known_growers(),
        "shipped_pieces": sum(len(h["boxes"]) for h in hawbs),
        "mark": mark, "batch": batch,
        "label": batch_label(batch, db.get_delivery_batches()) if batch else "",
        "invoices": invoices, "boxes": boxes, "hawbs": hawbs,
        "deliveries": sorted(deliveries.values(), key=lambda d: d["id"]),
        "not_shipped": not_shipped, "rates": rates, "docs": docs,
        "missing": [(dl, h, l) for dl in deliveries.values() for h, l in dl["missing"]],
        "ambiguous": [(dl, h, l) for dl in deliveries.values() for h, l in dl["ambiguous"]],
    }


# --- файлы -------------------------------------------------------------------

def _delivery_invoices(delivery):
    """Инвойсы, урезанные до коробок этой поставки, - вход для combine_by_mark."""
    per_doc = {}
    for hawb in delivery["hawbs"]:
        for b in hawb["boxes"]:
            inv = b["invoice"]
            entry = per_doc.setdefault(inv["id"], {"inv": inv, "boxes": []})
            entry["boxes"].append((b["idx"], b["box"]))
    result = []
    for entry in sorted(per_doc.values(), key=lambda e: (e["inv"]["date"] or "", e["inv"]["id"])):
        inv = entry["inv"]
        data = dict(inv["data"], boxes=[box for _, box in sorted(entry["boxes"], key=lambda t: t[0])])
        result.append({"filename": inv["filename"], "data": data, "template": inv["template"]})
    return result


def invoice_total_xls(delivery):
    """Invoice total по фактической поставке: коробки - по HAWB, логистика -
    вес × ставка. Формат - тот же, что у импорта (import_xls_writer)."""
    by_mark = combine_by_mark(_delivery_invoices(delivery))
    weight, cost = delivery["weight"], delivery["cost"]
    gross = round(sum(h["data"].get("gross_weight") or 0 for h in delivery["hawbs"]), 2)
    for info in by_mark.values():
        info["awb"] = {
            "pieces": delivery["pieces"], "gross_weight": gross, "chargeable_weight": weight,
            "rate_per_kg": round(cost / weight, 4) if weight else None,
            "rate_all_in": round(cost / weight, 4) if weight else None, "total_awb": cost,
        }
        # Шапку берём из HAWB, а не из инвойсов (там номера AWB - плановые).
        info["awb_number"] = ", ".join(h["data"].get("mawb") or "" for h in delivery["hawbs"])
        info["hawb_number"] = ", ".join(h["data"].get("hawb") or "" for h in delivery["hawbs"])
        info["airline"] = ", ".join(sorted({h["data"].get("airline") or "" for h in delivery["hawbs"]}))
        info["invoice_date"] = fmt_date(delivery["date"])
    buf = io.BytesIO()
    build_combined_factura_xls(buf, by_mark, {"awb_no": info_awb_numbers(delivery),
                                              "total_awb": cost})
    buf.seek(0)
    return buf


def info_awb_numbers(delivery):
    return ", ".join(h["data"].get("mawb") or "" for h in delivery["hawbs"])


_BORDER = "borders: left thin, right thin, top thin, bottom thin"


def acceptance_xls(mark, delivery):
    """Лист для приёмки: по каждому рейсу - коробки с фермой, сортами и
    стеблями и пустая колонка «Принято» для отметки на складе."""
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("приёмка")
    bold = xlwt.easyxf("font: bold on")
    head = xlwt.easyxf("font: bold on; " + _BORDER + "; pattern: pattern solid, fore_colour gray25")
    cell = xlwt.easyxf(_BORDER)
    widths = [22, 16, 12, 8, 34, 8, 10, 12]
    for c, w in enumerate(widths):
        ws.col(c).width = 256 * w

    ws.write(0, 0, f"Поставка {mark}: доставка {fmt_date(delivery.get('arrival'))}, "
                   f"{delivery['airport_name']} (вылет {fmt_date(delivery['date'])})", bold)
    ws.write(1, 0, f"Мест по HAWB: {delivery['pieces']} · коробок найдено: "
                   f"{delivery['boxes_found']} · стеблей: {delivery['stems']} · "
                   f"вес: {delivery['weight']} кг")
    r = 3
    headers = ["Ферма", "Инвойс", "Коробка", "Доля", "Сорт", "Длина", "Стеблей", "Принято"]
    for hawb in delivery["hawbs"]:
        d = hawb["data"]
        ws.write(r, 0, f"Рейс {d.get('mawb')} · HAWB {d.get('hawb')} · {d.get('airline') or ''} · "
                       f"мест {d.get('pieces')} · {hawb['weight']} кг", bold)
        r += 1
        for c, h in enumerate(headers):
            ws.write(r, c, h, head)
        r += 1
        for b in hawb["boxes"]:
            first = True
            for it in b["box"].get("items") or []:
                length = it.get("length_cm") if it.get("length_cm") is not None else it.get("grade_text")
                values = [b["grower"] if first else "", b["invoice"]["data"].get("invoice_no") if first else "",
                          (b["box"].get("box_type") or "") if first else "", b["size"] if first else "",
                          it.get("variety"), length, it.get("stems"), ""]
                for c, v in enumerate(values):
                    ws.write(r, c, "" if v is None else v, cell)
                r += 1
                first = False
        for line in hawb["lines"]:
            if line["missing"]:
                ws.write(r, 0, f"Нет инвойса: {line['name']} — {line['missing'] / UNITS:g} полных коробок",
                         xlwt.easyxf("font: bold on, colour red"))
                r += 1
        r += 1
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
