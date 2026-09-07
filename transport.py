"""
Расчёт логистики: для каждого кода тары считаем долю стоимости палеты,
пропорциональную количеству коробок этого типа в поставке (коробки не всегда
занимают палету целиком - платим только за фактически используемую долю).

    cost(fust_code) = pallet_cost_usd / capacity_per_pallet(fust_code) * box_count(fust_code)

Итоговая стоимость логистики (USD) = сумма cost() по всем типам тары,
округлённая вверх до целого доллара.
"""
import math


def count_boxes_by_fust(boxes):
    counts = {}
    for b in boxes:
        code = str(b["fust"])
        counts[code] = counts.get(code, 0) + 1
    return counts


def calculate_transport(boxes, tara_mapping, pallet_cost_usd):
    """
    boxes: список коробок [{"fust": "AAA", ...}, ...] (после объединения инвойсов)
    tara_mapping: dict fust_code -> {"capacity_per_pallet": float, "label": str}
    pallet_cost_usd: float

    Возвращает:
      {
        "by_fust": [{"fust_code":.., "count":.., "capacity":.., "cost_usd":..}, ...],
        "missing_fust": ["XXX", ...],  # коды тары без записи в справочнике
        "total_usd": float,  # округлено вверх до целого
      }
    """
    counts = count_boxes_by_fust(boxes)
    by_fust = []
    missing = []
    total = 0.0

    for code, count in sorted(counts.items()):
        mapping = tara_mapping.get(code)
        if not mapping or not mapping.get("capacity_per_pallet"):
            missing.append(code)
            continue
        capacity = mapping["capacity_per_pallet"]
        cost = pallet_cost_usd / capacity * count
        total += cost
        by_fust.append({
            "fust_code": code,
            "label": mapping.get("label") or code,
            "count": count,
            "capacity_per_pallet": capacity,
            "cost_usd": round(cost, 2),
        })

    return {
        "by_fust": by_fust,
        "missing_fust": missing,
        "total_usd": math.ceil(total),
    }
