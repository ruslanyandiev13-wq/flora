"""
Формула стоимости обработки одной накладной в токенах (ТЗ "Биллинг / История",
раздел 1.1). Веса вынесены в именованные константы, чтобы их можно было менять
централизованно, не трогая роуты и не пересчитывая формулу руками.

Сознательно НЕ включён штраф за новый/неоткалиброванный шаблон поставщика -
освоение нового формата это работа по развитию сервиса, а не то, за что должен
переплачивать заказчик при обычной обработке (см. ТЗ, раздел 1.1).
"""

BASE_COST = 5
COST_PER_EXTRA_PDF_PAGE = 1          # сверх первой страницы
COST_PER_20_LINE_ITEMS = 1
COST_AWB_FREE_TEXT = 3               # AWB со свободным текстом (нужен текстовый парсинг)
COST_AWB_STRUCTURED_FIELD = 1        # AWB со структурированным полем метки
COST_PER_EXTRA_AWB_IN_MARK = 2       # каждый доп. AWB, суммированный в рамках одной метки
COST_MIX_RULE_APPLIED = 1            # правило MIX применено хотя бы к одной коробке
# HAWB форвардера в разделе «Поставки» - не инвойс, а AWB с текстовым
# разбором строк «доля = ферма», поэтому без базовой ставки инвойса.
COST_DELIVERY_HAWB = COST_AWB_FREE_TEXT

COMPLEXITY_LABELS = (
    (10, "🟢 Простая"),
    (25, "🟡 Средняя"),
    (None, "🔴 Сложная"),
)


def complexity_label(score):
    for threshold, label in COMPLEXITY_LABELS:
        if threshold is None or score <= threshold:
            return label


def calc_cost(*, page_count=1, line_items_count=0, awb_free_text=False,
               awb_structured=False, extra_awb_count=0, mix_rule_applied=False):
    """Считает стоимость одной накладной в токенах по формуле из ТЗ.

    page_count - число страниц PDF (для XLS-инвойсов передавать 1).
    awb_free_text / awb_structured - взаимоисключающие флаги первого AWB
        в накладной (если AWB в документе вообще нет - оба False).
    extra_awb_count - количество ДОПОЛНИТЕЛЬНЫХ AWB сверх первого,
        суммированных в рамках одной метки.
    """
    cost = BASE_COST
    cost += max(0, page_count - 1) * COST_PER_EXTRA_PDF_PAGE
    cost += (line_items_count // 20) * COST_PER_20_LINE_ITEMS
    if awb_free_text:
        cost += COST_AWB_FREE_TEXT
    elif awb_structured:
        cost += COST_AWB_STRUCTURED_FIELD
    cost += extra_awb_count * COST_PER_EXTRA_AWB_IN_MARK
    if mix_rule_applied:
        cost += COST_MIX_RULE_APPLIED
    return cost
