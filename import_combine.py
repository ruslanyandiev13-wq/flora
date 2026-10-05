"""
Фаза 3 (ТЗ по импортным инвойсам): сборка нескольких уже разобранных
farm-инвойсов (import_parser.py) в один сквозно пронумерованный черновик
по метке (mark) - основа для будущего файла "factura" (Фаза 4).

Структура подтверждена реальным примером `Invoice total 30.08 SIRI MOS.xls`
(лист factura): одна строка на физическую коробку, колонки FARM/№/PRODUCT/
BOX/BOX SIZE/VARIETY/GRADE/TOTAL STEMS/UNIT PRICE/TOTAL USD - у первой
позиции коробки заполнены BOX/BOX SIZE, у последующих позиций той же
коробки (MIX) эти два поля пустые. Ровно так уже устроена внутренняя
структура boxes[].items[] в import_parser.py - переносить один в один.
"""
import re

# Короткие метки поставщика в колонке FARM. "garda"/"tessa"/"matiz" взяты
# буквально из реального примера factura - подтверждено. Остальные шесть -
# лучший читаемый вариант по умолчанию, НЕ подтверждённый реальным файлом
# (эти поставщики пока не встречались в примерах факtура) - если увидите
# другое написание в реальном файле, нужно поправить здесь.
FARM_LABELS = {
    "matiz_roses": "matiz",
    "gardaexport": "garda",
    "tessa": "tessa",
    "florsani": "florsani",
    # 2026-09-11: закупщик поправил - "saftec" это карго-агент (он же стоит в
    # шапке CARGO AGENCY), а плантация называется corazon.
    "rosas_corazon": "corazon",
    "monterosas": "monterosas",
    "monterosas_v2": "monterosas",
    "rosaprima": "rosaprima",
    "rosaprima_ec": "rosaprima",
    "ceresfarms": "ceres",
    "utopia": "utopia",
    "star_roses": "star roses",
}

# Astoria - консолидатор: в колонке FARM должна стоять сама плантация, а не
# "astoria". Бухгалтер пишет их сокращённо - здесь только ПОДТВЕРЖДЁННЫЕ
# написания из его файлов ("Колумбия 13.09 (AST)_AGATA MOS.xls" и
# "Invoice total 30.08 SIRI MOS.xls"). Незнакомые плантации оставляем как
# напечатано в инвойсе - не выдумываем сокращения.
PLANTATION_LABELS = {
    "FLORES DE SERREZUELA S.A.": "serrezuella",
    "FLORES DE SERREZUELA": "serrezuella",
}

# То же для колонки PRODUCT: в инвойсе Astoria тип написан как "SPRAY " -
# закупщик в своём файле расшифровывает его как "SPRAY CARNATION"
# (плантация возит именно гвоздику-спрей). Если встретится "SPRAY" у
# розовой плантации - расшифровку нужно будет уточнить.
PRODUCT_ALIASES = {
    "SPRAY": "SPRAY CARNATION",
    # Astoria: "S  MATHIOLAS" -> MATHIOLA, как в справочнике культур
    # закупщика (правка 2026-10-01, Invoice total BESST 04.10, строки 39-40).
    "S MATHIOLAS": "MATHIOLA",
    "S MATHIOLA": "MATHIOLA",
    "MATHIOLAS": "MATHIOLA",
}

# Колонка PRODUCT - ВНИМАНИЕ: это не постоянное свойство поставщика (см.
# ниже, урок с TESSA): культура определяется по названию сорта, и лишь если
# по нему ничего не понятно - по ферме.
# Ниже - поставщики, у которых культура НЕ роза и в названии сорта её не видно
# (у Gardaexport сорта называются "Pink MIX COLOR" и т.п.). Для всех
# остальных поставщиков продукт по умолчанию - "ROSES <ферма>" (см.
# _product_for_box): так подписано в эталонном файле бухгалтера
# ("ROSES matiz", "ROSES star roses") и так же попросил закупщик для tessa
# и monterosas (правка 2026-09-10).
PRODUCT_LABELS = {
    "gardaexport": "ALSTROEMERIAS",
}

# 2026-09-08: обнаружено на реальных инвойсах, что TESSA отгружает РАЗНЫЕ
# культуры в разных инвойсах одного и того же поставщика (альстромерии,
# гвоздики, розы) - "tessa -> ALSTROEMERIAS" в PRODUCT_LABELS оказалось
# неверной догадкой (было подтверждено только одним примером, где случайно
# оказалась альстромерия). Продукт печатается прямо в названии сорта -
# альстромерии и гвоздики ВСЕГДА печатаются с этим префиксом ("ALSTRO ...",
# "CARNATION ...") - определяем по нему, а не по таблице поставщика. Для роз
# TESSA такого префикса нет нигде в инвойсе (сорт печатается голым именем
# культивара, напр. "PINK FLOYD"/"CANDLELIGHT") - в этом случае product
# остаётся None, а не гадаем "ROSES" без доказательства.
# Список культур взят из самого эталонного файла бухгалтера ("Invoice total
# 30.08 SIRI MOS.xls", колонка-справочник справа от таблицы) - это его
# собственный словарь названий, а не придуманный нами.
_PRODUCT_KEYWORDS = [
    ("ALSTRO", "ALSTROEMERIAS"),
    ("CARNATION", "CARNATION"),
    ("RANUNCULUS", "RANUNCULUS"),
    ("GYPSOPHIL", "GYPSOPHILIA"),
    ("MATHIOLA", "MATHIOLA"),
    ("ANEMONE", "ANEMONE"),
    ("DELPHINIUM", "DELPHINIUM"),
    ("MOLUCELLA", "MOLUCELLA"),
    ("CALLA", "CALLA"),
]


# Длина, начиная с которой розы НЕ объединяются: длинные розы редкие и
# продаются посортово (правило закупщика 2026-09-10: "если в коробке много
# сортов по одной/две пачки 25 или 50 стеблей и длина 40/50/60см - объединяем,
# а всё, что 70 и выше, не объединяем"). На другие культуры порог не
# распространяется - у них объединение всегда (альстромерии 70см закупщик
# принял объединёнными).
ROSES_NO_MERGE_FROM_CM = 70

# Что именно значит "много сортов по одной/две пачки" - уточнено закупщиком
# 2026-09-11 на разборе tests/test1109 (его ручной файл "Invoice total 16.09
# VIKTORIA.xls"): коробку star roses из 16 сортов по 25 стеблей он свёл в MIX,
# а коробки из ДВУХ сортов (monterosas 50+75 стеблей, corazon 100+50) оставил
# построчно. Порог выбран им самим - "от 5 сортов", пачка = 25 стеблей.
ROSES_MIX_MIN_VARIETIES = 5
ROSES_MIX_MAX_STEMS_PER_VARIETY = 50


def normalize_variety(variety):
    """Название сорта прописными буквами, как в файле закупщика.

    - "Pink MIX COLOR"/"White MIX COLOR" (gardaexport) -> "MIX": сорт внутри
      такой коробки не отслеживается, это уже микс от плантации.
    - "ALSTRO PINK PRIMADONNA" (tessa) -> "ALSTRO": сорт альстромерии закупщик
      тоже не ведёт (правка 2026-09-11, строки 57-58).
    """
    v = (variety or "").strip()
    up = v.upper()
    if "MIX COLOR" in up:
        return "MIX"
    if up.startswith("ALSTRO"):
        return "ALSTRO"
    return up


def _normalize_carnation_item(item):
    """Разделяет описание TESSA на сорт и сортность, сохраняя сырой инвойс.

    "Carnation crimea select" + 70 см -> CRIMEA / SELECT. Если сортность
    уже указана отдельным полем (Astoria/брокер или ручная правка), она
    имеет приоритет. Неизвестные окончания названия не угадываем.
    """
    result = dict(item)
    variety = re.sub(r"^CARNATION\s+", "", (item.get("variety") or "").strip().upper())
    match = re.fullmatch(r"(.+?)\s+(SELECT|FANCY)", variety)
    if match:
        variety = match.group(1)
    result["variety"] = variety
    # Ручная длина/сортность, в том числе очистка клетки, важнее суффикса
    # старого описания. Название сорта при этом всё равно очищаем.
    if item.get("_grade_edited"):
        return result
    grade = str(item.get("grade_text") or "").strip().upper()
    if match:
        grade = grade or match.group(2)
    if grade:
        result.update(grade_text=grade, length_cm=None)
    return result


def _dedupe_identical(items):
    """Складывает строки одной культуры с одинаковыми сортом, грейдом и ценой.

    Это не MIX, а склейка дублей: у corazon EXPLORER 80 см по 0.70 напечатан
    тремя строками (200+25+25), а в файле закупщика это одна строка на 250
    стеблей. Сюда же попадают сорта, которые normalize_variety() свела к одному
    имени (несколько ALSTRO ... у tessa) - если исходные названия были разными,
    запоминаем их в merged_from для листа "объяснение".
    """
    out = []
    index = {}
    for it in items:
        variety = normalize_variety(it.get("variety"))
        key = (variety, it.get("length_cm"), it.get("grade_text"), it.get("price"), it.get("product"))
        prev = index.get(key)
        if prev is None:
            copy = dict(it, variety=variety)
            if variety != (it.get("variety") or "").strip().upper():
                copy["renamed_from"] = it.get("variety")
            index[key] = copy
            out.append(copy)
            continue
        prev["stems"] = (prev.get("stems") or 0) + (it.get("stems") or 0)
        prev["total"] = round((prev.get("total") or 0) + (it.get("total") or 0), 2)
        sources = prev.get("merged_from") or [prev.get("renamed_from") or prev["variety"]]
        # Объяснять нечего, если склеились строки одного и того же сорта
        # (EXPLORER 200+25+25) - пишем в лист "объяснение" только реальные миксы.
        if any(name != it.get("variety") for name in sources):
            prev["merged_from"] = sources + [it.get("variety")]
    return out


def _roses_group_merges(group):
    """Можно ли свести эту группу роз в MIX: сортов не меньше пяти и у
    каждого не больше двух пачек."""
    if len(group) < ROSES_MIX_MIN_VARIETIES:
        return False
    return all((it.get("stems") or 0) <= ROSES_MIX_MAX_STEMS_PER_VARIETY for it in group)


def merge_same_grade_items(items, product=None):
    """Сводит позиции одной коробки в строки "MIX" по правилам закупщика.

    - Сначала складываем полные дубли (один сорт, одна длина, одна цена).
    - Гвоздика: сохраняем сорта и сортность, не сводим разные сорта в MIX
      (правка 2026-10-05).
    - Розы: объединяем только позиции короче 70 см И только если сортов в
      группе не меньше ROSES_MIX_MIN_VARIETIES, каждый по 1-2 пачки; иначе
      выводим посортово (правка 2026-09-11: две-три позиции по 50-100 стеблей
      закупщик хочет видеть по сортам).
    - Остальные культуры: объединяем всегда.

    Объединяем по длине, текстовой сортности и цене, а не по одной длине: в подтверждённом
    эталонном файле бухгалтера есть коробка, где две строки MIX имеют
    одинаковую длину 80 см, но разную цену (0.28 и 0.30) - и они оставлены
    отдельными строками.

    В объединённой позиции остаётся "merged_from" - список исходных сортов;
    по нему строится лист "объяснение" в итоговом файле, чтобы закупщик видел,
    что именно было слито, и мог поправить.
    """
    product = (product or "").strip().upper()
    is_roses = product.startswith(("ROSES", "SPRAY ROSES"))
    is_carnation = product == "CARNATION"

    items = _dedupe_identical(items)
    if is_carnation:
        return items

    groups = {}
    order = []
    for it in items:
        key = (it.get("length_cm"), it.get("grade_text"), it.get("price"), it.get("product"))
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(it)

    merged = []
    for key in order:
        group = groups[key]
        length = key[0]
        if is_roses and length is not None and length >= ROSES_NO_MERGE_FROM_CM:
            merged.extend(group)
            continue
        if is_roses and len(group) > 1 and not _roses_group_merges(group):
            merged.extend(group)
            continue
        if len(group) == 1:
            merged.append(group[0])
            continue
        # Сорт объединённой строки: у альстромерии закупщик пишет "ALSTRO",
        # у остальных культур - "MIX".
        name = ("ALSTRO" if all((it.get("variety") or "").upper() == "ALSTRO" for it in group)
                else "MIX")
        sources = []
        for it in group:
            sources.extend(it.get("merged_from") or [it.get("renamed_from") or it.get("variety")])
        merged.append({
            **group[0],
            "variety": name,
            "stems": sum(it.get("stems") or 0 for it in group),
            "total": round(sum(it.get("total") or 0 for it in group), 2),
            "merged_from": sources,
        })
    return merged


def top_length_items(items):
    """Правило "считать по верхней ростовке": все позиции коробки получают
    максимальную встретившуюся в ней длину.

    Применяется ТОЛЬКО к альстромерии (правка закупщика 2026-09-11: "мы
    задавали правило в альстромерию... и это правило встало в розу"). Раньше
    оно жило в import_parser.apply_mix_rule() и применялось ко всем шаблонам -
    из-за этого розы 50 см превращались в 60/70/80 см. Список возвращается
    новый, исходные позиции не мутируются."""
    lengths = [it["length_cm"] for it in items if it.get("length_cm") is not None]
    if not lengths:
        return items
    top = max(lengths)
    return [dict(it, length_cm=top) if it.get("length_cm") is not None else it
            for it in items]


# Плантации, у которых спрей-розы напечатаны слитным префиксом "SP"
# (SPGISELLE, SPWINKCORALKISS). Только подтверждённые - у остальных ферм
# опираемся на явное слово SPRAY, иначе обычные сорта вроде SPARKLE/SPIRIT
# попали бы в спрей по ошибке (решение закупщика 2026-09-11).
SP_PREFIX_TEMPLATES = {"rosas_corazon"}


def is_spray_variety(variety, template=None):
    up = (variety or "").upper().strip()
    if up.startswith("SPRAY"):
        return True
    return template in SP_PREFIX_TEMPLATES and re.match(r"^SP[A-Z]", up) is not None


def _product_for_item(product, variety, template):
    """PRODUCT отдельной строки. Спрей-розы закупщик подписывает "SPRAY ROSES
    <ферма>" - и делает это построчно: в одной коробке corazon у него рядом
    стоят "ROSES corazon" и "SPRAY ROSES corazon" (правка 2026-09-11)."""
    if product and product.upper().startswith("ROSES") and is_spray_variety(variety, template):
        return f"SPRAY {product}"
    return product


def _product_for_box(template, box):
    """PRODUCT коробки: если культуру видно в названии сорта - берём её,
    иначе это розы этой фермы ("ROSES tessa", "ROSES monterosas" и т.д.);
    исключение - фермы из PRODUCT_LABELS, которые возят не розы."""
    for it in box["items"]:
        variety = (it.get("variety") or "").upper()
        for keyword, label in _PRODUCT_KEYWORDS:
            if keyword in variety:
                return label
    if template in PRODUCT_LABELS:
        return PRODUCT_LABELS[template]
    farm = FARM_LABELS.get(template)
    return f"ROSES {farm}" if farm else None


def _sum_transport(mark_invoices):
    """Перевозка из инвойсов Astoria/брокера. Неполную сумму или вес нельзя
    выдавать за итог всей метки; ноль при этом остаётся известным значением."""
    transports = [inv["data"].get("transport") or {} for inv in mark_invoices]
    if not any(transports):
        return None

    def total(field, legacy_weight=False):
        values = [t.get(field, t.get("weight_kg") if legacy_weight else None) for t in transports]
        return round(sum(values), 2) if all(v is not None for v in values) else None

    return {"cost_usd": total("cost_usd"), "weight_kg": total("weight_kg"),
            "gross_weight": total("gross_weight", legacy_weight=True),
            "chargeable_weight": total("chargeable_weight", legacy_weight=True),
            "awb_no": next((t["awb_no"] for t in transports if t.get("awb_no")), None)}


def box_farm_product(inv, box):
    """FARM и PRODUCT коробки так, как они попадут в factura. Для Astoria
    плантация и культура - из самой коробки (у неё Plantation свой на каждую),
    у остальных шаблонов ферма фиксирована по инвойсу, а культура - по
    названию сорта (см. _product_for_box). Ручная правка со страницы проверки
    кладётся в box["farm"]/box["product"] и поэтому имеет приоритет."""
    product = box.get("product") or _product_for_box(inv["template"], box)
    product = PRODUCT_ALIASES.get(" ".join((product or "").upper().split()), product)
    farm = box.get("farm") or FARM_LABELS.get(inv["template"]) or inv["data"]["supplier"]
    farm = PLANTATION_LABELS.get(farm.upper(), farm)
    return farm, product


def split_invoice_by_mark(inv):
    """Инвойс, в котором коробки разных меток (у Astoria метка - в колонке
    Handler каждой коробки: 001C-20260828 - SIRI, BESST и AGATA в одном
    файле), -> по инвойсу на метку. Обычный инвойс возвращается как есть."""
    data = inv["data"]
    marks = []
    for b in data.get("boxes") or []:
        mark = b.get("mark") or data.get("mark")
        if mark not in marks:
            marks.append(mark)
    if len(marks) <= 1:
        return [inv]
    return [dict(inv, data=dict(data, mark=mark,
                                boxes=[b for b in data["boxes"] if (b.get("mark") or data.get("mark")) == mark]))
            for mark in marks]


def combine_by_mark(invoices):
    """invoices: список {"filename":.., "data":.., "template":..} (формат
    pending-сессии import_app.py). Возвращает dict {mark: {...}} - для
    каждой метки сквозная нумерация коробок по всем инвойсам этой метки и
    метка FARM на каждой коробке (для Astoria - из самой коробки, у неё
    Plantation свой на каждую коробку; для остальных шаблонов - фиксированная
    по всему инвойсу)."""
    by_mark = {}
    invoices = [part for inv in invoices for part in split_invoice_by_mark(inv)]
    for inv in invoices:
        mark = inv["data"].get("mark") or "?"
        by_mark.setdefault(mark, []).append(inv)

    result = {}
    for mark, mark_invoices in by_mark.items():
        combined_boxes = []
        box_counter = 0
        for inv in mark_invoices:
            for b in inv["data"]["boxes"]:
                box_counter += 1
                farm, product = box_farm_product(inv, b)
                if (product or "").strip().upper() == "CARNATION":
                    product = "CARNATION"
                # Верхняя ростовка - только у альстромерии, у всех остальных
                # культур длины остаются такими, как напечатаны в инвойсе.
                # Определяем PRODUCT до объединения: после замены названий
                # на MIX префикс SPRAY/SP уже нельзя восстановить.
                items = [dict(it, variety=(it.get("variety") or "").strip().upper(),
                              product=_product_for_item(product, it.get("variety"), inv["template"]))
                         for it in b["items"]]
                if product == "CARNATION":
                    items = [_normalize_carnation_item(it) for it in items]
                if (product or "").upper().startswith("ALSTRO"):
                    items = top_length_items(items)
                # no_merge - признак от парсера (Колумбия: всё построчно,
                # ничего не сводим в MIX).
                if not b.get("no_merge"):
                    items = merge_same_grade_items(items, product)
                # Сортность гортензии: PREMIUM/PREMIUN/PREM -> PR (правка
                # 2026-10-02). Меняем только копии итоговых позиций;
                # числовые длины и грейды остальных культур сохраняем.
                for it in items:
                    if ((it.get("product") or "").strip().upper() == "HYDRANGEAS"
                            and it.get("length_cm") is None
                            and str(it.get("grade_text") or "").strip().upper().startswith("PR")):
                        it["grade_text"] = "PR"
                combined_boxes.append({
                    "box_no": box_counter,
                    "farm": farm,
                    "product": product,
                    "box_size": b.get("box_size"),
                    "items": items,
                    "source": inv["filename"],
                })
        total_stems = sum(it["stems"] for b in combined_boxes for it in b["items"])
        total_fob = sum((it["total"] or 0) for b in combined_boxes for it in b["items"])
        total_full_boxes = sum((b["box_size"] or 0) for b in combined_boxes)

        # Шапка (CONSIGNEE/DESTINATION/CARGO AGENCY/AIRLINE/M.A.W.B/H.A.W.B/
        # DATE) - берём из самих фермерских инвойсов этой метки (поле уже
        # разбирается в import_parser.py per-template), а не только из
        # ручного ввода AWB-веса (см. import_app.py::factura_awb - вес/ставка
        # реально нигде не печатаются в фермерских инвойсах, а вот номер AWB/
        # перевозчик/направление обычно печатаются). Берём первое непустое
        # значение среди инвойсов этой метки.
        def _first(field):
            for inv in mark_invoices:
                v = inv["data"].get(field)
                if v:
                    return v
            return None

        result[mark] = {
            "boxes": combined_boxes,
            "total_stems": total_stems,
            "total_fob": round(total_fob, 2),
            "total_full_boxes": round(total_full_boxes, 3),
            "invoices": mark_invoices,
            "destination": _first("destination"),
            "forwarder": _first("forwarder"),
            "airline": _first("airline"),
            # Названы awb_number/hawb_number (не awb/hawb!), т.к. в
            # import_app.py::factura() ключ info["awb"] уже занят под ручной
            # ввод веса/ставки (см. factura_awb) - совпадение имён затёрло
            # бы номер AWB словарём веса.
            "awb_number": _first("awb"),
            "hawb_number": _first("hawb"),
            "invoice_date": _first("invoice_date"),
            # Транспорт, указанный В САМИХ инвойсах (Astoria). Если он есть -
            # отдельная авианакладная для этой метки не нужна (правка
            # закупщика 2026-09-10). Суммируем по всем инвойсам метки.
            "transport": _sum_transport(mark_invoices),
        }
    return result
