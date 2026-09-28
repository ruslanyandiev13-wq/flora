"""
Ручные правки распознанных позиций голландского модуля перед выгрузкой.

Правки живут в pending-сессии (app.py::_save_pending) рядом с исходными
инвойсами - сами разобранные данные не трогаем, поверх них накладываем
pending["edits"]:

    {"<box_no>": {"fust": "AAA",
                  "items": {"<индекс позиции в коробке>": {"aantal": 50, ...}}}}

box_no - сквозной номер коробки из app.py::_combined_boxes; он стабилен, пока
набор загруженных файлов тот же (а он в pending-сессии не меняется).
"""
import copy

# Поля позиции, которые можно править, и чем их разбирать. Описание - текст,
# остальное - числа (запятая как десятичный разделитель тоже принимается).
ITEM_FIELDS = {
    "aantal": "int",
    "omschrijving": "text",
    "prijs": "number",
    "lengte": "int",
    "gew": "number",
    "bedrag": "number",
}
BOX_FIELDS = {"fust": "text"}

FIELD_TITLES = {
    "aantal": "Кол-во", "omschrijving": "Описание", "prijs": "Цена",
    "lengte": "S1", "gew": "S2", "bedrag": "Сумма", "fust": "Тара",
}


def parse_value(field, raw):
    """Значение из поля ввода -> то, что пойдёт в xls. Пустое - None
    (для кол-ва/описания/цены пустое запрещено). ValueError с понятным
    текстом, если число не разобралось."""
    kind = ITEM_FIELDS.get(field) or BOX_FIELDS.get(field)
    if kind is None:
        raise ValueError(f"Поле «{field}» не редактируется")
    raw = ("" if raw is None else str(raw)).strip()
    if not raw:
        if field in ("aantal", "omschrijving", "prijs", "fust"):
            raise ValueError(f"«{FIELD_TITLES[field]}» не может быть пустым")
        return None
    if kind == "text":
        return raw
    try:
        number = float(raw.replace(",", ".").replace(" ", ""))
    except ValueError:
        raise ValueError(f"«{FIELD_TITLES[field]}»: «{raw}» - не число")
    if number < 0:
        raise ValueError(f"«{FIELD_TITLES[field]}» не может быть отрицательным")
    if kind == "int":
        if number != int(number):
            raise ValueError(f"«{FIELD_TITLES[field]}» должно быть целым")
        return int(number)
    return number


def set_edit(edits, box_no, field, value, item_idx=None):
    """Записывает правку в словарь edits (мутирует и возвращает его)."""
    box = edits.setdefault(str(box_no), {})
    if item_idx is None:
        box[field] = value
    else:
        box.setdefault("items", {}).setdefault(str(item_idx), {})[field] = value
    return edits


def apply_edits(boxes, edits):
    """Накладывает правки на объединённые коробки.

    Возвращает (новые коробки, изменение суммы товара). У исправленных
    позиций и коробок появляется "original" - исходные значения исправленных
    полей (для подсветки на странице). Если поменяли кол-во или цену, а
    сумму руками не трогали - сумма пересчитывается как кол-во × цена.
    """
    boxes = copy.deepcopy(boxes)
    money_delta = 0.0
    for b in boxes:
        box_edits = (edits or {}).get(str(b["box_no"]))
        if not box_edits:
            continue
        if "fust" in box_edits:
            b["original"] = {"fust": b["fust"]}
            b["fust"] = box_edits["fust"]
        for idx_str, item_edits in (box_edits.get("items") or {}).items():
            idx = int(idx_str)
            if idx >= len(b["items"]) or not item_edits:
                continue
            it = b["items"][idx]
            it["original"] = {f: it.get(f) for f in item_edits}
            old_bedrag = it.get("bedrag")
            it.update(item_edits)
            if ("bedrag" not in item_edits
                    and ("aantal" in item_edits or "prijs" in item_edits)
                    and it.get("aantal") is not None and it.get("prijs") is not None):
                it["original"].setdefault("bedrag", old_bedrag)
                it["bedrag"] = round(it["aantal"] * it["prijs"], 2)
            money_delta += (it.get("bedrag") or 0) - (old_bedrag or 0)
    return boxes, round(money_delta, 2)


def count_edits(edits):
    n = 0
    for box in (edits or {}).values():
        n += sum(1 for f in box if f != "items")
        n += sum(len(fields) for fields in (box.get("items") or {}).values())
    return n
