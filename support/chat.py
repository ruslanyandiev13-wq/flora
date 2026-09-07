"""
Чат-помощник поверх OpenAI function calling (ТЗ "Чат-помощник и обратная
связь", разделы 1-2). Модель НИКОГДА не пишет итоговый файл сама - только
диалог и ограниченный набор function-calling инструментов; реальная пересборка
файла идёт через support/regenerate.py -> существующий код модулей.

Отступление от буквального текста ТЗ: get_recent_documents() и
get_document_details() в спецификации принимают user_id/бессрочный
document_id без проверки владельца. Модель не должна сама решать, чьи
документы ей показывать - это открывало бы одному пользователю доступ к
чужим накладным через чат. Поэтому user_id в схему инструмента НЕ вынесен -
личность всегда берётся из текущей серверной сессии, а get_document_details
проверяет, что документ принадлежит этому пользователю (или роль admin).

correction_type ('light'|'heavy') добавлен в propose_correction сверх
буквальной сигнатуры ТЗ - без этого apply_correction не может знать, нужен
ли повторный разбор PDF (см. ТЗ раздел 2.2, где это два явно разных по
стоимости действия). Модели в system-промпте объяснено, когда какой ставить.
"""
import json
import os
import uuid

from openai import OpenAI

import auth
import document_store
from billing import db as billing_db
from billing.cost_calc import calc_cost, complexity_label
import support.db as support_db
import support.regenerate as regenerate

MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o")

_client = None


def _get_client():
    global _client
    if _client is None:
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            raise RuntimeError("OPENAI_API_KEY не задан в окружении")
        _client = OpenAI(api_key=api_key)
    return _client


TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "get_recent_documents",
            "description": "Получить последние обработанные документы текущего пользователя",
            "parameters": {
                "type": "object",
                "properties": {
                    "limit": {"type": "integer", "description": "Сколько документов вернуть (по умолчанию 5)"}
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_document_details",
            "description": "Получить структурированные данные документа (то же, что видно на странице проверки)",
            "parameters": {
                "type": "object",
                "properties": {"document_id": {"type": "integer"}},
                "required": ["document_id"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "propose_correction",
            "description": "Предложить точечное исправление поля в документе. НЕ применяется "
                            "автоматически - только показывается пользователю как diff.",
            "parameters": {
                "type": "object",
                "properties": {
                    "document_id": {"type": "integer"},
                    "field_path": {"type": "string", "description": "Путь к полю, напр. 'boxes.2.items.0.prijs'"},
                    "current_value": {"description": "Текущее значение поля"},
                    "proposed_value": {"description": "Предлагаемое новое значение"},
                    "reason": {"type": "string"},
                    "correction_type": {
                        "type": "string", "enum": ["light", "heavy"],
                        "description": "'light' - точечная замена значения; 'heavy' - только если "
                                        "значение нельзя надёжно определить без повторного разбора "
                                        "исходного PDF (например, целая пропущенная строка/сорт)",
                    },
                },
                "required": ["document_id", "field_path", "current_value", "proposed_value",
                             "reason", "correction_type"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "apply_correction",
            "description": "Применить ранее предложенное исправление. Вызывать ТОЛЬКО после "
                            "явного подтверждения пользователем ('да, применяй' и т.п.).",
            "parameters": {
                "type": "object",
                "properties": {"correction_id": {"type": "integer"}},
                "required": ["correction_id"],
            },
        },
    },
]

SYSTEM_PROMPT = """Ты - ассистент поддержки в сервисе обработки инвойсов цветочных поставщиков \
(модули "Голландия" и "Импорт"). Помогаешь пользователю с уже обработанными документами.

Правила:
- Ты НЕ можешь и не должен сам переписывать итоговый файл или додумывать цифры. Любое изменение \
данных - только через propose_correction, и оно применяется только через apply_correction ПОСЛЕ \
явного согласия пользователя.
- Сначала пойми, какой документ обсуждается (используй get_recent_documents/get_document_details).
- Если правка - это просто замена одного неверно распознанного значения на верное - \
correction_type='light'. Если для уверенного исправления нужно заново прочитать исходный PDF \
(например, ты не уверен в значении, или пользователь говорит "тут вообще всё не так распознано") - \
correction_type='heavy'.
- Если проблема не сводится к точечной правке (весь формат поставщика распознан неверно и т.п.) - \
не пытайся исправить сам. Сообщи, что это передано разработчику на рассмотрение.
- Перед вызовом apply_correction всегда покажи предложенное изменение простыми словами и дождись \
явного подтверждения.
- Отвечай кратко и по-русски."""


def _current_username():
    user = auth.current_user()
    return user["username"] if user else None


def _is_admin():
    user = auth.current_user()
    return bool(user and user["role"] == "admin")


def _tool_get_recent_documents(args):
    username = _current_username()
    rows = billing_db.get_history(uploaded_by=username, limit=args.get("limit", 5))
    return [{"id": r["id"], "filename": r["source_filename"], "date": r["upload_datetime"],
              "module": r["module"], "status": r["status"]} for r in rows]


def _tool_get_document_details(args):
    doc = billing_db.get_invoice(args["document_id"])
    if not doc:
        return {"error": "Документ не найден"}
    if doc["uploaded_by"] != _current_username() and not _is_admin():
        return {"error": "Нет доступа к этому документу"}
    context = json.loads(doc["generation_context_json"]) if doc.get("generation_context_json") else None
    parsed = json.loads(doc["parsed_data_json"]) if doc.get("parsed_data_json") else None
    return {
        "id": doc["id"], "module": doc["module"], "filename": doc["source_filename"],
        "status": doc["status"], "boxes": (context or {}).get("boxes") or (parsed or {}).get("boxes"),
    }


def _tool_propose_correction(args):
    doc = billing_db.get_invoice(args["document_id"])
    if not doc:
        return {"error": "Документ не найден"}
    if doc["uploaded_by"] != _current_username() and not _is_admin():
        return {"error": "Нет доступа к этому документу"}
    correction_id = support_db.create_correction(
        document_id=args["document_id"], field_path=args["field_path"],
        current_value=args["current_value"], proposed_value=args["proposed_value"],
        reason=args["reason"], correction_type=args.get("correction_type", "light"),
    )
    return {"correction_id": correction_id, "diff": {
        "было": args["current_value"], "станет": args["proposed_value"], "причина": args["reason"],
    }}


def _tool_apply_correction(args):
    correction = support_db.get_correction(args["correction_id"])
    if not correction:
        return {"error": "Исправление не найдено"}
    if correction["applied"]:
        return {"error": "Уже применено"}

    doc = billing_db.get_invoice(correction["document_id"])
    if not doc:
        return {"error": "Документ не найден"}
    if doc["uploaded_by"] != _current_username() and not _is_admin():
        return {"error": "Нет доступа к этому документу"}

    if correction["correction_type"] == "heavy":
        if not doc.get("source_file_path"):
            return {"error": "Исходный файл недоступен - тяжёлая правка невозможна"}
        out_path, new_context, fresh_data = regenerate.reparse_source_and_replace(
            doc["module"], json.loads(doc["generation_context_json"]),
            doc["source_file_path"], doc["source_filename"])
        line_items_count = sum(len(b["items"]) for b in fresh_data["boxes"])
        cost = calc_cost(page_count=doc["page_count"] or 1, line_items_count=line_items_count)
        label = complexity_label(cost)
        billing_db.update_invoice_after_correction(
            correction["document_id"], out_path, new_context, cost, label, line_items_count,
            parsed_data=fresh_data)
        billing_db.charge_tokens(cost, correction["document_id"],
                                  org_id=billing_db.DEFAULT_ORG_ID, reason="chat_correction")
        cost_info = {"cost": cost, "invoice_id": correction["document_id"]}
    else:
        context = json.loads(doc["generation_context_json"]) if doc.get("generation_context_json") else None
        if context is None:
            return {"error": "Для этого документа нет сохранённого контекста генерации - "
                              "лёгкая правка невозможна"}
        regenerate.apply_light_correction(context, correction["field_path"], correction["proposed_value"])
        out_path = regenerate.rebuild(doc["module"], context)
        billing_db.update_invoice_result(correction["document_id"], out_path, generation_context=context)
        billing_db.charge_tokens(3, correction["document_id"],
                                  org_id=billing_db.DEFAULT_ORG_ID, reason="chat_correction")
        cost_info = {"cost": 3, "invoice_id": correction["document_id"]}

    document_store.store_result(correction["document_id"], out_path)
    support_db.mark_correction_applied(args["correction_id"])
    return {"applied": True, "result_file_path": out_path, "tokens_charged": cost_info["cost"]}


_TOOL_IMPL = {
    "get_recent_documents": _tool_get_recent_documents,
    "get_document_details": _tool_get_document_details,
    "propose_correction": _tool_propose_correction,
    "apply_correction": _tool_apply_correction,
}


def start_session(related_document_id=None):
    session_id = uuid.uuid4().hex
    support_db.create_session(session_id, _current_username(), related_document_id=related_document_id)
    return session_id


def send_message(session_id, user_text, max_tool_rounds=6):
    """Один ход диалога: добавляет сообщение пользователя, гоняет tool-calling
    цикл, сохраняет всё в chat_messages, возвращает финальный текст ассистента."""
    client = _get_client()
    support_db.add_message(session_id, "user", user_text)

    history = support_db.get_messages(session_id)
    messages = [{"role": "system", "content": SYSTEM_PROMPT}]
    for m in history:
        messages.append({"role": m["role"], "content": m["content"]})

    for _ in range(max_tool_rounds):
        response = client.chat.completions.create(model=MODEL, messages=messages, tools=TOOLS)
        choice = response.choices[0]
        msg = choice.message

        if not msg.tool_calls:
            support_db.add_message(session_id, "assistant", msg.content or "")
            return msg.content or ""

        messages.append({"role": "assistant", "content": msg.content or "",
                          "tool_calls": [tc.model_dump() for tc in msg.tool_calls]})
        tool_call_log = []
        for tc in msg.tool_calls:
            fn_name = tc.function.name
            try:
                fn_args = json.loads(tc.function.arguments)
            except json.JSONDecodeError:
                fn_args = {}
            impl = _TOOL_IMPL.get(fn_name)
            result = impl(fn_args) if impl else {"error": f"Неизвестный инструмент {fn_name}"}
            tool_call_log.append({"name": fn_name, "args": fn_args, "result": result})
            messages.append({"role": "tool", "tool_call_id": tc.id, "content": json.dumps(result, ensure_ascii=False, default=str)})

        support_db.add_message(session_id, "assistant", msg.content or "(вызов инструментов)",
                                tool_calls=tool_call_log)

    return "Не удалось завершить обработку запроса за разумное число шагов - попробуйте переформулировать."
