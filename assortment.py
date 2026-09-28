"""
Приведение названий сортов из паклиста MH Flowers к тому виду, в котором их
ведёт бухгалтер (её файл "Голландия (ассортимент).xlsx", залитый в справочник
`assortment`).

Зачем: MH Flowers печатает в колонке Omschrijving не только сорт, но иногда и
плантацию - "R Tr Summer Dance Water", "Matth Impala Marine Ton Vreugdenhil".
В одних случаях бухгалтер такой хвост убирает, в других он часть названия
позиции ("R Tr Yellow Babe Flora Ola" - это отдельная строка её ассортимента).
Отличить одно от другого по самой накладной нельзя - только по справочнику
(разбор правок бухгалтера от 2026-09-15, файлы в tests/test_holland).

Порядок применения (первое сработавшее правило выигрывает):
 1. Ручная замена из справочника `variety_aliases` - для случаев, которых нет
    ни в паклисте, ни в ассортименте ("Blushing Bride 4-6" -> "Serruria
    Blushing Bride": слова "Serruria" в PDF нет вообще).
 2. Напечатанное название целиком есть в ассортименте - оставляем как есть.
 3. Есть укороченное (отбрасываем слова с конца) - берём написание из
    ассортимента.
 4. Ничего не нашли - берём название, уже очищенное парсером от имени
    плантации по колонке Kweker (см. pdf_parser._strip_kweker_from_omschrijving).
"""
import db


def load_lookup():
    """Читает справочники один раз на разбор накладной, а не на каждую строку."""
    names = {n.strip().lower(): n.strip() for n in db.get_assortment_names()}
    aliases = {a["printed"].strip().lower(): a["canonical"].strip()
               for a in db.get_variety_aliases()}
    return {"names": names, "aliases": aliases}


def canonical_name(printed, fallback=None, lookup=None):
    """printed - как напечатано в PDF, fallback - очищенное парсером название."""
    lookup = lookup or load_lookup()
    names, aliases = lookup["names"], lookup["aliases"]
    printed = (printed or "").strip()
    fallback = (fallback or printed).strip()

    for candidate in (printed, fallback):
        if candidate and candidate.lower() in aliases:
            return aliases[candidate.lower()]
    if printed.lower() in names:
        return names[printed.lower()]
    if fallback.lower() in names:
        return names[fallback.lower()]

    words = printed.split()
    for cut in range(len(words) - 1, 0, -1):
        tail = words[cut:]
        # Отрезаем только хвост, похожий на имя плантации ("... Water",
        # "... Flora Ola"). Всё, где есть цифры, - это ростовка или калибр
        # ("Li Ot Zambesi 5+", "Cymb T Toledo Decorum 80cm"), и бухгалтер такие
        # хвосты оставляет: они часть названия позиции.
        if any(any(ch.isdigit() for ch in word) for word in tail):
            continue
        short = " ".join(words[:cut])
        if short.lower() in names:
            return names[short.lower()]
    return fallback or printed
