"""
Пересборка итогового файла после правки (см. ТЗ "Чат-помощник и обратная
связь", раздел 1: "генерация исправленного файла всегда идёт через
существующий код парсинга/сборки... с применённой точечной правкой").

Работает поверх generation_context_json, сохранённого в billing.db при
генерации (см. billing/charge.py). Сейчас поддержан только модуль
'netherlands' - у 'import' ещё нет шага генерации итогового файла (Фаза 4
не согласована), поэтому и корректировать там пока нечего.
"""
import datetime
import os
import uuid

import pdf_parser
from xls_writer import build_xls

GENERATED_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "chat_generated")
os.makedirs(GENERATED_DIR, exist_ok=True)


def _get_by_path(obj, path):
    for part in _split_path(path):
        obj = obj[part] if isinstance(obj, dict) else obj[int(part)]
    return obj


def _set_by_path(obj, path, value):
    parts = _split_path(path)
    for part in parts[:-1]:
        obj = obj[part] if isinstance(obj, dict) else obj[int(part)]
    last = parts[-1]
    if isinstance(obj, list):
        obj[int(last)] = value
    else:
        obj[last] = value


def _split_path(path):
    return [p for p in path.replace("[", ".").replace("]", "").split(".") if p != ""]


def apply_light_correction(generation_context, field_path, proposed_value):
    """Меняет одно поле в generation_context на месте, без повторного
    разбора исходного файла. Возвращает изменённый context (та же ссылка)."""
    _set_by_path(generation_context, field_path, proposed_value)
    return generation_context


def _rebuild_netherlands(generation_context):
    header = dict(generation_context["header"])
    if isinstance(header.get("date"), str):
        header["date"] = datetime.date.fromisoformat(header["date"])
    out_path = os.path.join(GENERATED_DIR, f"{uuid.uuid4().hex}.xls")
    build_xls(out_path, generation_context["boxes"], header, generation_context["total_money"],
              pallet_cost_usd=generation_context.get("pallet_cost_usd"),
              usd_rate=generation_context.get("usd_rate"))
    return out_path


def rebuild(module, generation_context):
    if module == "netherlands":
        return _rebuild_netherlands(generation_context)
    raise ValueError(f"Пересборка для модуля '{module}' пока не поддержана")


def reparse_source_and_replace(module, generation_context, source_file_path, source_filename):
    """'Тяжёлая' правка (ТЗ раздел 2.2): заново гонит исходный файл через
    штатный парсер, заменяет в generation_context только те коробки, что
    пришли из ЭТОГО источника (по source_filename - см. app.py::_combined_boxes,
    у каждой коробки есть поле 'source'), остальные коробки батча не трогает,
    затем пересобирает файл и перенумеровывает коробки сквозным счётчиком."""
    if module != "netherlands":
        raise ValueError(f"Тяжёлая правка для модуля '{module}' пока не поддержана")

    data, _template = pdf_parser.parse_invoice_pdf(source_file_path)
    new_items = [{
        "aantal": it["aantal"], "omschrijving": it["omschrijving"], "prijs": it["prijs"],
        "lengte": it["lengte"], "gew": it["gew"], "bedrag": it["bedrag"],
    } for b in data["boxes"] for it in b["items"]]
    new_boxes_raw = [{"fust": b["fust"], "items": [
        {"aantal": it["aantal"], "omschrijving": it["omschrijving"], "prijs": it["prijs"],
         "lengte": it["lengte"], "gew": it["gew"], "bedrag": it["bedrag"]}
        for it in b["items"]
    ], "source": source_filename} for b in data["boxes"]]

    kept = [b for b in generation_context["boxes"] if b.get("source") != source_filename]
    combined = kept + new_boxes_raw
    for i, b in enumerate(combined, start=1):
        b["box_no"] = i
    generation_context["boxes"] = combined
    generation_context["total_money"] = sum(
        it["bedrag"] for b in combined for it in b["items"]
    )
    out_path = _rebuild_netherlands(generation_context)
    return out_path, generation_context, data
