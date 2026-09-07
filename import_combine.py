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
    "rosas_corazon": "saftec",
    "monterosas": "monterosas",
    "rosaprima": "rosaprima",
    "ceresfarms": "ceres",
    "utopia": "utopia",
}

# Колонка PRODUCT в реальном примере factura-файла ("Invoice total 30.08 SIRI
# MOS.xls") - подтверждена только для 3 поставщиков, которые в нём реально
# встретились: garda/tessa -> "ALSTROEMERIAS", matiz -> "ROSES matiz". Для
# остальных - НЕ подтверждено (нет ни одного реального примера с этой
# колонкой для них), поэтому оставляем None, а не гадаем (та же логика, что
# и для неподтверждённых FARM_LABELS выше).
PRODUCT_LABELS = {
    "gardaexport": "ALSTROEMERIAS",
    "tessa": "ALSTROEMERIAS",
    "matiz_roses": "ROSES matiz",
}


def combine_by_mark(invoices):
    """invoices: список {"filename":.., "data":.., "template":..} (формат
    pending-сессии import_app.py). Возвращает dict {mark: {...}} - для
    каждой метки сквозная нумерация коробок по всем инвойсам этой метки и
    метка FARM на каждой коробке (для Astoria - из самой коробки, у неё
    Plantation свой на каждую коробку; для остальных шаблонов - фиксированная
    по всему инвойсу)."""
    by_mark = {}
    for inv in invoices:
        mark = inv["data"].get("mark") or "?"
        by_mark.setdefault(mark, []).append(inv)

    result = {}
    for mark, mark_invoices in by_mark.items():
        combined_boxes = []
        box_counter = 0
        for inv in mark_invoices:
            farm_label = FARM_LABELS.get(inv["template"])
            product_label = PRODUCT_LABELS.get(inv["template"])
            for b in inv["data"]["boxes"]:
                box_counter += 1
                combined_boxes.append({
                    "box_no": box_counter,
                    "farm": b.get("farm") or farm_label or inv["data"]["supplier"],
                    "product": b.get("product") or product_label,
                    "box_size": b.get("box_size"),
                    "items": b["items"],
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
        }
    return result
