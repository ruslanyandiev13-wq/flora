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
 3. Название кончается известной припиской плантации/марки (справочник
    `grower_suffixes`: Decorum, Location Aalsmeer, Water...) - убираем её и
    берём написание из ассортимента, если оно там есть.
 4. Ничего не нашли - берём название, уже очищенное парсером от имени
    плантации по колонке Kweker (см. pdf_parser._strip_kweker_from_omschrijving).

Сравнение с ассортиментом - без учёта регистра, лишних пробелов и точки в
конце ("Chr T Altaj." = "Chr T Altaj").

Раньше в п.3 хвост отрезался по одному слову, пока укороченное название не
находилось в ассортименте - это резало и настоящие сорта: "Chr T Resq Salmon"
превращался в "Chr T Resq", "Chr S Purpetta Red" в "Chr S Purpetta" (разбор
файла бухгалтера за 30.09: такие она оставляет целиком). Отличить приписку
от цвета/сорта по паклисту нельзя, поэтому только явный список.
"""
import db


def _key(name):
    return " ".join((name or "").split()).rstrip(".").strip().lower()


def load_lookup():
    """Читает справочники один раз на разбор накладной, а не на каждую строку."""
    names = {_key(n): n.strip() for n in db.get_assortment_names()}
    aliases = {_key(a["printed"]): a["canonical"].strip()
               for a in db.get_variety_aliases()}
    # Длинные приписки проверяем первыми ("Location Aalsmeer" раньше "Aalsmeer").
    suffixes = sorted((s.lower() for s in db.get_grower_suffixes()), key=len, reverse=True)
    return {"names": names, "aliases": aliases, "suffixes": suffixes}


def strip_suffixes(name, suffixes):
    """Убирает с конца названия известные приписки (можно несколько подряд)."""
    name = " ".join((name or "").split())
    changed = True
    while changed:
        changed = False
        for suffix in suffixes:
            if name.lower().endswith(" " + suffix):
                name = name[:-len(suffix)].strip()
                changed = True
    return name


def canonical_name(printed, fallback=None, lookup=None):
    """printed - как напечатано в PDF, fallback - очищенное парсером название."""
    lookup = lookup or load_lookup()
    names, aliases = lookup["names"], lookup["aliases"]
    printed = (printed or "").strip()
    fallback = (fallback or printed).strip()

    for candidate in (printed, fallback):
        if candidate and _key(candidate) in aliases:
            return aliases[_key(candidate)]
    for candidate in (printed, fallback):
        if _key(candidate) in names:
            return names[_key(candidate)]

    short = strip_suffixes(printed, lookup.get("suffixes") or [])
    if short != " ".join(printed.split()):
        return names.get(_key(short), short)
    return fallback or printed
