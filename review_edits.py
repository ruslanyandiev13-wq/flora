"""
Ручные правки распознанных позиций перед выгрузкой - голландский модуль
(/review) и импортный (/import/review).

Правки живут в pending-сессии рядом с исходными инвойсами - сами разобранные
данные не трогаем, поверх них накладываем pending["edits"]:

    {"<box_no>": {"fust": "AAA",
                  "items": {"<индекс позиции в коробке>": {"aantal": 50, ...}}}}

box_no - сквозной номер коробки по всем загруженным инвойсам (как на
странице проверки); он стабилен, пока набор загруженных файлов тот же (а он
в pending-сессии не меняется).

Какие поля правятся и как разбираются - описано схемой (DUTCH / IMPORT).
"""
import copy
from dataclasses import dataclass, field


@dataclass(frozen=True)
class Schema:
    # поле -> "text" / "number" / "int" / "grade" (число - длина, иначе текст)
    item_fields: dict
    box_fields: dict
    titles: dict
    required: frozenset = field(default_factory=frozenset)
    # Сумма строки пересчитывается как qty × price, если её не правили руками.
    qty: str = None
    price: str = None
    amount: str = None


DUTCH = Schema(
    item_fields={"aantal": "int", "omschrijving": "text", "prijs": "number",
                 "lengte": "int", "gew": "number", "bedrag": "number"},
    box_fields={"fust": "text"},
    titles={"aantal": "Кол-во", "omschrijving": "Описание", "prijs": "Цена",
            "lengte": "S1", "gew": "S2", "bedrag": "Сумма", "fust": "Тара"},
    required=frozenset({"aantal", "omschrijving", "prijs", "fust"}),
    qty="aantal", price="prijs", amount="bedrag",
)

# Импорт: "grade" - одна клетка "Длина / грейд" на странице. Число пишется в
# length_cm, текст (FANCY/PREM у Astoria) - в grade_text; см. grade_edits().
# В edits хранится как одна правка "grade", раскрывается при применении.
IMPORT = Schema(
    item_fields={"variety": "text", "grade": "grade", "stems": "int",
                 "price": "number", "total": "number"},
    box_fields={"farm": "text", "product": "text", "box_size": "number"},
    titles={"variety": "Сорт", "grade": "Длина / грейд", "stems": "Стеблей",
            "price": "Цена", "total": "Сумма", "farm": "FARM",
            "product": "PRODUCT", "box_size": "Доля коробки"},
    required=frozenset({"variety", "stems", "price"}),
    qty="stems", price="price", amount="total",
)


def parse_value(name, raw, schema=DUTCH):
    """Значение из поля ввода -> то, что пойдёт в xls. Пустое - None
    (для обязательных полей запрещено). ValueError с понятным текстом,
    если число не разобралось."""
    kind = schema.item_fields.get(name) or schema.box_fields.get(name)
    if kind is None:
        raise ValueError(f"Поле «{name}» не редактируется")
    title = schema.titles[name]
    raw = ("" if raw is None else str(raw)).strip()
    if not raw:
        if name in schema.required:
            raise ValueError(f"«{title}» не может быть пустым")
        return None
    if kind == "text":
        return raw
    try:
        number = float(raw.replace(",", ".").replace(" ", ""))
    except ValueError:
        if kind == "grade":
            return raw.upper()
        raise ValueError(f"«{title}»: «{raw}» - не число")
    if number < 0:
        raise ValueError(f"«{title}» не может быть отрицательным")
    if kind in ("int", "grade"):
        if number != int(number):
            raise ValueError(f"«{title}» должно быть целым")
        return int(number)
    return number


def grade_edits(value):
    """Клетка "Длина / грейд" импорта -> правки настоящих полей позиции."""
    if isinstance(value, int):
        return {"length_cm": value, "grade_text": None}
    return {"length_cm": None, "grade_text": value}


def display_value(item, name):
    """Значение поля так, как оно показано в клетке страницы проверки."""
    if name == "grade":
        return item.get("length_cm") if item.get("length_cm") is not None else item.get("grade_text")
    return item.get(name)


def set_edit(edits, box_no, name, value, item_idx=None):
    """Записывает правку в словарь edits (мутирует и возвращает его)."""
    box = edits.setdefault(str(box_no), {})
    if item_idx is None:
        box[name] = value
    else:
        box.setdefault("items", {}).setdefault(str(item_idx), {})[name] = value
    return edits


def drop_edit(edits, box_no, name, item_idx=None):
    """Убирает правку (значение вернули к исходному) и опустевшие уровни."""
    box = edits.get(str(box_no))
    if not box:
        return edits
    if item_idx is None:
        box.pop(name, None)
    else:
        items = box.get("items") or {}
        item = items.get(str(item_idx)) or {}
        item.pop(name, None)
        if not item:
            items.pop(str(item_idx), None)
        if not items:
            box.pop("items", None)
    if not box:
        edits.pop(str(box_no))
    return edits


def apply_box_edits(box, box_edits, schema=DUTCH):
    """Накладывает правки одной коробки на неё саму (мутирует box).

    У исправленных коробок и позиций появляется "original" - исходные
    значения исправленных полей (для подсветки на странице). Если поменяли
    кол-во или цену, а сумму руками не трогали - она пересчитывается.
    Возвращает изменение суммы по коробке."""
    if not box_edits:
        return 0.0
    box_level = {k: v for k, v in box_edits.items() if k != "items"}
    if box_level:
        box["original"] = {k: box.get(k) for k in box_level}
        box.update(box_level)
    money_delta = 0.0
    for idx_str, item_edits in (box_edits.get("items") or {}).items():
        idx = int(idx_str)
        if idx >= len(box["items"]) or not item_edits:
            continue
        it = box["items"][idx]
        it["original"] = {k: display_value(it, k) for k in item_edits}
        old_amount = it.get(schema.amount)
        for k, v in item_edits.items():
            it.update(grade_edits(v) if schema.item_fields.get(k) == "grade" else {k: v})
        if (schema.amount not in item_edits
                and (schema.qty in item_edits or schema.price in item_edits)
                and it.get(schema.qty) is not None and it.get(schema.price) is not None):
            it["original"].setdefault(schema.amount, old_amount)
            it[schema.amount] = round(it[schema.qty] * it[schema.price], 2)
        money_delta += (it.get(schema.amount) or 0) - (old_amount or 0)
    return money_delta


def apply_edits(boxes, edits, schema=DUTCH):
    """Правки поверх списка коробок с полем box_no. Возвращает (новые
    коробки, изменение суммы товара); исходный список не меняется."""
    boxes = copy.deepcopy(boxes)
    money_delta = 0.0
    for b in boxes:
        money_delta += apply_box_edits(b, (edits or {}).get(str(b["box_no"])), schema)
    return boxes, round(money_delta, 2)


def strip_original(item):
    """Позиция без служебного "original" - для xls и истории выгрузки."""
    return {k: v for k, v in item.items() if k != "original"}


def count_edits(edits):
    n = 0
    for box in (edits or {}).values():
        n += sum(1 for f in box if f != "items")
        n += sum(len(fields) for fields in (box.get("items") or {}).values())
    return n


def record_edit(edits, payload, original_boxes, schema=DUTCH):
    """Проверяет и записывает правку одной клетки из запроса страницы
    проверки: {"box_no", "item_idx" (нет - правка коробки), "field", "value"}.

    original_boxes - коробки БЕЗ правок (с box_no), по ним проверяется, что
    позиция существует, и узнаётся исходное значение: если его вернули,
    правка убирается, чтобы клетка не подсвечивалась. Возвращает
    (box_no, item_idx, field); ValueError с текстом для пользователя."""
    name = payload.get("field")
    item_idx = payload.get("item_idx")
    try:
        box_no = int(payload.get("box_no"))
        item_idx = None if item_idx is None else int(item_idx)
    except (TypeError, ValueError):
        raise ValueError("Некорректный номер коробки или позиции")
    allowed = schema.box_fields if item_idx is None else schema.item_fields
    if name not in allowed:
        raise ValueError(f"Поле «{name}» не редактируется")
    value = parse_value(name, payload.get("value"), schema)

    box = next((b for b in original_boxes if b["box_no"] == box_no), None)
    if box is None or (item_idx is not None and not 0 <= item_idx < len(box["items"])):
        raise ValueError("Такой позиции нет - обновите страницу")
    source = box if item_idx is None else box["items"][item_idx]

    if value == display_value(source, name):
        drop_edit(edits, box_no, name, item_idx)
    else:
        set_edit(edits, box_no, name, value, item_idx)
    return box_no, item_idx, name


def edit_response(boxes, box_no, item_idx, name, edits, schema=DUTCH):
    """Ответ страницы после сохранения правки: новое значение клетки и
    пересчитанная сумма строки (boxes - коробки уже с правками)."""
    box = next(b for b in boxes if b["box_no"] == box_no)
    target = box if item_idx is None else box["items"][item_idx]
    original = target.get("original") or {}
    return {
        "value": display_value(target, name),
        "original": original.get(name),
        "edited": name in original,
        "amount": None if item_idx is None else target.get(schema.amount),
        "amount_edited": item_idx is not None and schema.amount in original,
        "edits_count": count_edits(edits),
    }
