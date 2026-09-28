"""
Единая точка интеграции с голландским и импортным модулями (ТЗ "Биллинг /
История", раздел 4): "после успешной генерации итогового файла и в
Голландии, и в Импорте должен вызываться единый billing-хук... Не
дублировать логику списания в каждом модуле отдельно."
"""
import document_store
from billing import db as billing_db
from billing.cost_calc import calc_cost, complexity_label


def charge_for_invoice(org_id, module, metadata):
    """metadata: dict с ключами uploaded_by, supplier_template, page_count,
    line_items_count, awb_free_text, awb_structured, extra_awb_count,
    mix_rule_applied, status ('processed'|'error'|'pending_review'),
    result_file_path (опционально).

    Плюс опционально - для будущих чат-правок (ТЗ "Чат-помощник и обратная
    связь", раздел 1): source_file_path (путь к исходно загруженному файлу),
    parsed_data (структура из pdf_parser/import_parser для ЭТОЙ накладной),
    generation_context (всё, что нужно, чтобы заново вызвать build_xls/etc и
    получить тот же результат - обычно header + полный список коробок пакета).

    Токены списываются, только если status == 'processed' (см. ТЗ, раздел 1:
    при ошибке разбора списания не происходит) - но запись в invoices
    создаётся в любом случае, чтобы ошибки тоже были видны в истории.

    Возвращает {"invoice_id": int, "cost": int, "label": str}.
    """
    # fixed_cost - документ, который считается не по формуле инвойса
    # (HAWB в «Поставках», см. cost_calc.COST_DELIVERY_HAWB).
    cost = metadata["fixed_cost"] if metadata.get("fixed_cost") is not None else calc_cost(
        page_count=metadata.get("page_count", 1),
        line_items_count=metadata.get("line_items_count", 0),
        awb_free_text=metadata.get("awb_free_text", False),
        awb_structured=metadata.get("awb_structured", False),
        extra_awb_count=metadata.get("extra_awb_count", 0),
        mix_rule_applied=metadata.get("mix_rule_applied", False),
    )
    label = complexity_label(cost)
    status = metadata.get("status", "processed")

    awb_count = (1 if (metadata.get("awb_free_text") or metadata.get("awb_structured")) else 0) \
        + metadata.get("extra_awb_count", 0)

    invoice_id = billing_db.record_invoice(
        org_id=org_id, module=module, uploaded_by=metadata.get("uploaded_by", "?"),
        supplier_template=metadata.get("supplier_template"),
        page_count=metadata.get("page_count", 1),
        line_items_count=metadata.get("line_items_count", 0),
        awb_count=awb_count, mix_rule_applied=metadata.get("mix_rule_applied", False),
        complexity_score=cost, complexity_label=label, status=status,
        result_file_path=metadata.get("result_file_path"),
        source_filename=metadata.get("source_filename"),
    )

    if metadata.get("source_file_path") or metadata.get("parsed_data") or metadata.get("generation_context"):
        stored_source = document_store.store_source(invoice_id, metadata.get("source_file_path"))
        billing_db.attach_document(
            invoice_id, source_file_path=stored_source,
            parsed_data=metadata.get("parsed_data"),
            generation_context=metadata.get("generation_context"),
        )

    if status == "processed":
        billing_db.charge_tokens(cost, invoice_id, org_id=org_id,
                                  reason=metadata.get("charge_reason", "invoice_processing"))

    return {"invoice_id": invoice_id, "cost": cost, "label": label}
