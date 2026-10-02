"""
Фаза 4 (импортные инвойсы): сборка итогового .xls-файла "factura" по метке -
формат подтверждён реальным примером бухгалтера
"samples_files_0109/Invoice total 30.08 SIRI MOS.xls" (лист "factura").
Отдельный модуль от голландского xls_writer.py - голландский не трогаем.

Из реального примера воспроизведены:
  - шапка (CONSIGNEE/DESTINATION/CARGO AGENCY/AIRLINE/M.A.W.B/H.A.W.B/DATE/
    TOTAL FULL BOXES/ставка за кг) - те же подписи и позиции ячеек;
  - таблица товаров: FARM(0)/№(1)/PRODUCT(3)/BOX(5)/BOX SIZE(6)/VARIETY(7)/
    GRADE(10)/TOTAL STEMS(12)/UNIT PRICE(14)/TOTAL USD(15,16 - дублируется
    в обе колонки, как в реальном файле, там колонка 16 явно использовалась
    для суммирования итога) - BOX/BOX SIZE заполнены только у первой позиции
    коробки (MIX), как и во внутренней структуре boxes[].items[];
  - итоговый блок: TOTAL STEMS/TOTAL FLOWERS FOB USD/TOTAL CHARGEABLE
    WEIGHT(Kg)/GROSS WEIGHT/TOTAL AWB/TOTAL USD.

НЕ воспроизведено (в реальном файле похоже на разовые пометки бухгалтера,
а не часть шаблона): повтор метки/AWB-номера в строке EMAIL, отдельное поле
"N мест" рядом с FAX, нечто в TRUCK, COMISION/PACKING USD/INLAND USD/
TERMOGRAPHER (данных для них у нас нет), и разрозненные числа в конце листа
(строки 58-77 реального примера).
"""
import xlwt

import import_combine as combine

# Подсветка первой строки коробки в колонке VARIETY - как в файле закупщика
# "Invoice total 09.09 POLINA.xls": жёлтый = начало коробки, оранжевый
# (gold, тот же цвет, что у него) = коробка 0.25, зелёный = меньше 0.25.
# Тонкие границы у каждой ячейки строк товаров, колонки A..P (FARM..TOTAL USD) -
# как в файле закупщика "Колумбия 30.09 (AST)_SIRI.xls" (правка 2026-09-24:
# "границы нужны"). Шапка и итоги - без рамок, как у него.
_BORDERS = "borders: left thin, right thin, top thin, bottom thin"
TABLE_LAST_COL = 15

_CELL = xlwt.easyxf(_BORDERS)
# Только строки товаров factura: №, места, размер коробки, длина, стебли и цены.
_CENTERED_TABLE_COLS = (1, 5, 6, 10, 12, 14, 15)  # B, F, G, K, M, O, P
_CENTERED_CELL = xlwt.easyxf(_BORDERS + "; alignment: horiz center")
_BOX_START_STYLES = {
    "yellow": xlwt.easyxf(_BORDERS + "; pattern: pattern solid, fore_colour yellow"),
    "gold": xlwt.easyxf(_BORDERS + "; pattern: pattern solid, fore_colour gold"),
    "light_green": xlwt.easyxf(_BORDERS + "; pattern: pattern solid, fore_colour light_green"),
}


def _frame_row(ws, r, filled):
    """Дорисовывает рамку пустым ячейкам строки таблицы (xlwt не рисует
    границы у незаписанных ячеек)."""
    for c in range(TABLE_LAST_COL + 1):
        if c not in filled:
            ws.write(r, c, "", _CENTERED_CELL if c in _CENTERED_TABLE_COLS else _CELL)


def _box_start_style(box_size):
    if box_size is not None and box_size < 0.25:
        return _BOX_START_STYLES["light_green"]
    if box_size == 0.25:
        return _BOX_START_STYLES["gold"]
    return _BOX_START_STYLES["yellow"]


_SMALL = "font: height 160"  # на 2 пункта меньше обычных 10
_MARK_CELL = xlwt.easyxf(_BORDERS + "; font: height 160, colour grey50")
_CHECK_CELL = xlwt.easyxf(_SMALL + ", colour grey50")
_DATE = xlwt.easyxf("font: bold on", num_format_str="DD.MM.YYYY")
_DATE_PLAIN = xlwt.easyxf(num_format_str="DD.MM.YYYY")
_TOTAL_YELLOW = xlwt.easyxf("font: bold on; pattern: pattern solid, fore_colour yellow")
_TOTAL_GREEN = xlwt.easyxf("pattern: pattern solid, fore_colour lime")
_TOTAL_LABEL_WRAP = xlwt.easyxf("font: bold on; alignment: wrap on, vert centre")

# Колонки без данных о цветке в таблице позиций - скрываются (правка
# закупщика 2026-10-01/02): E, I, J, L, N. Подписи шапки перенесены в K.
# Подписи итогов, начинающиеся в J,
# и значения в N объединены с соседними колонками и остаются видны.
_HIDDEN_COLS = (4, 8, 9, 11, 13)


def _cell_name(r, c):
    """(9, 1) -> "B10" для формул."""
    letters = ""
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{r + 1}"


def build_combined_factura_xls(output_path, by_mark, awb_doc=None, consignee=None,
                               delivery_date=None):
    """Единый файл на всю партию (одна авианакладная): все коробки всех меток
    идут одной таблицей со сквозной нумерацией, метка вынесена в отдельную
    колонку (колонка 2 - в подтверждённом образце factura она пустая).
    Подтверждено закупщиком 2026-09-09 ("Да, в один. Invoice total").

    by_mark: результат import_combine.combine_by_mark(), уже дополненный
        AWB-блоком (см. import_app.py::_apply_awb).
    awb_doc: разобранная авианакладная (import_parser.parse_awb_pdf), если была
        загружена - из неё берётся номер AWB и общие вес/ставка партии.
    consignee: подпись получателя в B2/B8 ("BESST MOS 2"), по умолчанию - метки.
    delivery_date: дата поставки на склад (datetime.date) - в H8 и на
        листе «Склад приёмка».

    Раскладка листа factura - по правкам закупщика 2026-10-01 (эталон
    "Invoice total 04.10 BESST MOS 2 (AI).xls").
    """
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("factura")

    bold = xlwt.easyxf("font: bold on")
    normal = xlwt.easyxf("")

    filled = set()

    def w(r, c, v, style=normal):
        if v is not None:
            ws.write(r, c, v, style)
            filled.add(c)

    marks = list(by_mark.keys())
    consignee = consignee or ", ".join(marks)

    def first(field):
        for info in by_mark.values():
            if info.get(field):
                return info[field]
        return None

    awb_doc = awb_doc or {}
    total_full_boxes = round(sum(i.get("total_full_boxes") or 0 for i in by_mark.values()), 3)
    total_stems = sum(i.get("total_stems") or 0 for i in by_mark.values())
    total_fob = round(sum(i.get("total_fob") or 0 for i in by_mark.values()), 2)
    # Итог по накладной берём из самой AWB ("Total Prepaid"), если она была
    # загружена: сумма долей по меткам может разойтись с ней на копейку
    # из-за округления при пропорциональном делении.
    total_awb = round(sum((i.get("awb") or {}).get("total_awb") or 0 for i in by_mark.values()), 2)
    has_awb_total = (any((i.get("awb") or {}).get("total_awb") is not None for i in by_mark.values())
                     or awb_doc.get("total_awb") is not None)
    # Напечатанная перевозка инвойса и ручные правки имеют приоритет над
    # общей AWB, как и на экране проверки.
    if awb_doc.get("total_awb") and not any(
            (i.get("awb") or {}).get("source") in ("invoice", "manual") for i in by_mark.values()):
        total_awb = awb_doc["total_awb"]
    total_gross = round(sum((i.get("awb") or {}).get("gross_weight") or 0 for i in by_mark.values()), 2)
    total_chargeable = round(sum((i.get("awb") or {}).get("chargeable_weight") or 0
                                  for i in by_mark.values()), 2)
    # "ставка за кг" - как у закупщика: итог по накладной / платный вес
    # (включает сборы, поэтому не равна тарифу из AWB).
    rate = round(total_awb / total_chargeable, 4) if has_awb_total and total_chargeable else next(
        ((i.get("awb") or {}).get("rate_per_kg") for i in by_mark.values()
         if (i.get("awb") or {}).get("rate_per_kg") is not None), None)
    awb_no = awb_doc.get("awb_no") or first("awb_number")
    box_count = sum(len(i["boxes"]) for i in by_mark.values())

    # Шапка
    w(0, 0, "TO ")
    w(1, 0, "CONSIGNEE ")
    w(1, 1, consignee)
    w(1, 10, "DATE")
    w(1, 16, first("invoice_date"))

    w(2, 0, "DESTINATION")
    w(2, 1, first("destination"))
    w(2, 10, "CARGO AGENCY")
    w(2, 16, first("forwarder"))

    if rate is not None:
        w(3, 16, rate)
        w(3, 17, "ставка за кг")

    w(4, 0, "CITY/COUNTRY")
    w(4, 10, "AIRLINE")
    w(4, 16, first("airline"))

    w(5, 0, "TELEPHONE")
    w(5, 10, "M.A.W.B")
    w(5, 16, awb_no)

    HEADER_ROW = 9
    total_row = HEADER_ROW + 1 + sum(len(b["items"]) for i in by_mark.values() for b in i["boxes"])

    w(6, 0, "FAX")
    # Число мест = число коробок из итоговой строки (F) + слово «мест».
    ws.write(6, 2, xlwt.Formula(_cell_name(total_row, 5)), bold)
    w(6, 3, "мест")
    w(6, 10, "H.A.W.B")
    w(6, 16, first("hawb_number"))

    w(7, 0, "EMAIL")
    ws.write(7, 1, xlwt.Formula("B2"), bold)          # получатель ещё раз, жирным
    if delivery_date:
        ws.write(7, 7, delivery_date, _DATE)           # дата поставки на склад
    w(7, 10, "TOTAL FULL BOXES")
    w(7, 16, total_full_boxes)

    w(8, 0, "awb")
    if awb_no:
        ws.write(8, 1, xlwt.Formula(_cell_name(5, 16)))  # = M.A.W.B из Q6

    w(HEADER_ROW, 0, "FARM", bold)
    w(HEADER_ROW, 2, "метка", xlwt.easyxf(_SMALL + ", colour grey50"))
    w(HEADER_ROW, 3, "PRODUCT", bold)
    w(HEADER_ROW, 5, "BOX", bold)
    w(HEADER_ROW, 6, " BOX SIZE", bold)
    w(HEADER_ROW, 7, "VARIETY", bold)
    w(HEADER_ROW, 10, " GRADE", bold)
    w(HEADER_ROW, 12, "TOTAL STEMS", bold)
    w(HEADER_ROW, 14, "UNIT PRICE", bold)
    w(HEADER_ROW, 15, "TOTAL USD", bold)
    w(HEADER_ROW, 18, "OBS", bold)

    r = HEADER_ROW + 1
    box_no = 0
    for mark, info in by_mark.items():
        for b in info["boxes"]:
            box_no += 1
            first_item = True
            for it in b["items"]:
                filled.clear()
                w(r, 0, b.get("farm"), _CELL)
                w(r, 1, box_no, _CENTERED_CELL)
                # Метка - мелко и строчными: закупщику она не нужна (п.4).
                w(r, 2, (mark or "").lower(), _MARK_CELL)
                # PRODUCT построчно: у спрей-роз он свой ("SPRAY ROSES corazon"),
                # и в одной коробке такие строки соседствуют с обычными.
                w(r, 3, it.get("product") or b.get("product"), _CELL)
                if first_item:
                    w(r, 5, 1, _CENTERED_CELL)
                    w(r, 6, b.get("box_size"), _CENTERED_CELL)
                w(r, 7, it.get("variety"),
                  _box_start_style(b.get("box_size")) if first_item else _CELL)
                # GRADE: длина в см, а если она нечисловая (FANCY/SELECT/
                # 1000GR у Astoria) - текст как в инвойсе.
                w(r, 10, it.get("length_cm") if it.get("length_cm") is not None
                   else it.get("grade_text"), _CENTERED_CELL)
                w(r, 12, it.get("stems"), _CENTERED_CELL)
                w(r, 14, it.get("price"), _CENTERED_CELL)
                w(r, 15, it.get("total"), _CENTERED_CELL)
                # Q - перепроверка закупщика: стебли × цена, формулой (п.3).
                ws.write(r, 16, xlwt.Formula(f"{_cell_name(r, 12)}*{_cell_name(r, 14)}"), _CHECK_CELL)
                _frame_row(ws, r, filled)
                r += 1
                first_item = False

    # Итоговая строка прямо под таблицей - как в файлах закупщика: количество
    # коробок, сумма долей коробок, стебли и деньги в СВОИХ колонках.
    w(r, 5, box_no, bold)
    w(r, 6, total_full_boxes, bold)
    w(r, 12, total_stems)
    w(r, 15, total_fob, bold)
    if r > HEADER_ROW + 1:
        ws.write(r, 16, xlwt.Formula(f"SUM({_cell_name(HEADER_ROW + 1, 16)}:{_cell_name(r - 1, 16)})"),
                 _CHECK_CELL)

    # Блок итогов: подписи объединены J:M (TOTAL AWB / TOTAL USD - J:L),
    # значения стеблей и веса - N:P (п.9-11).
    def label(row, text, last_col=12, wrap=False):
        ws.write_merge(row, row, 9, last_col, text, _TOTAL_LABEL_WRAP if wrap else bold)
        if wrap:
            # При скрытой L длинная подпись веса занимает две строки.
            ws.row(row).height = 560
            ws.row(row).height_mismatch = True

    def value_np(row, v, style=normal):
        ws.write_merge(row, row, 13, 15, v if v is not None else "", style)

    r += 2
    label(r, "TOTAL STEMS")
    value_np(r, total_stems, _TOTAL_YELLOW)
    r += 1
    label(r, "TOTAL FLOWERS FOB USD")
    w(r, 19, total_fob, bold)
    r += 3
    label(r, "TOTAL CHARGEABLE WEIGHT(Kg)", wrap=True)
    value_np(r, total_chargeable or None, _TOTAL_GREEN)
    r += 1
    label(r, "GROSS WEIGHT")
    value_np(r, total_gross or None)
    r += 3
    label(r, "TOTAL AWB", last_col=11)
    w(r, 19, total_awb if has_awb_total else "", _TOTAL_GREEN)
    r += 1
    label(r, "TOTAL USD", last_col=11)
    w(r, 19, round(total_fob + total_awb, 2), _TOTAL_YELLOW)

    for c in _HIDDEN_COLS:
        ws.col(c).hidden = True
    ws.col(7).width = 256 * 26   # VARIETY - соседние I, J скрыты
    ws.col(2).width = 256 * 7

    _add_warehouse_sheet(wb, by_mark, bold, consignee, awb_no, delivery_date)
    _add_roses_sheet(wb, by_mark, bold)
    _add_other_flowers_sheet(wb, by_mark, bold)
    _add_explanation_sheet(wb, by_mark, bold)

    wb.save(output_path)
    return output_path


def _iter_rows(by_mark):
    """(метка, номер коробки, коробка, позиция, первая ли позиция коробки) -
    в том же порядке и с той же сквозной нумерацией, что на листе factura."""
    box_no = 0
    for mark, info in by_mark.items():
        for b in info["boxes"]:
            box_no += 1
            for i, it in enumerate(b["items"]):
                yield mark, box_no, b, it, i == 0


def _grade(it):
    return it.get("length_cm") if it.get("length_cm") is not None else it.get("grade_text")


def _is_roses(b):
    # "SPRAY ROSES corazon" - тоже розы, поэтому ищем ROSES в любом месте.
    return "ROSES" in (b.get("product") or "").upper()


def _add_warehouse_sheet(wb, by_mark, bold, consignee, awb_no, delivery_date):
    """Лист «Склад приёмка» - по образцу закупщика 2026-10-01: сверху места,
    получатель, AWB и дата поставки, ниже - все коробки без цен."""
    ws = wb.add_sheet("Склад приёмка")
    places = sum(len(i["boxes"]) for i in by_mark.values())
    ws.write(3, 2, places, bold)
    ws.write(3, 3, "мест")
    ws.write(4, 1, consignee, bold)
    ws.write(5, 0, "awb")
    ws.write_merge(5, 5, 1, 2, awb_no or "")
    if delivery_date:
        ws.write(5, 5, delivery_date, _DATE)
    head = xlwt.easyxf("font: bold on; " + _BORDERS)
    for c, title in enumerate(["FARM", "№", "PRODUCT", "BOX", "BOX SIZE", "VARIETY", "GRADE", "STEMS"]):
        ws.write(6, c, title, head)
    r = 7
    total = 0
    for _mark, box_no, b, it, first in _iter_rows(by_mark):
        values = [b.get("farm"), box_no, it.get("product") or b.get("product"),
                  1 if first else None, b.get("box_size") if first else None,
                  it.get("variety"), _grade(it), it.get("stems")]
        for c, v in enumerate(values):
            ws.write(r, c, "" if v is None else v, _CELL)
        total += it.get("stems") or 0
        r += 1
    ws.write(r, 6, "ИТОГО", bold)
    ws.write(r, 7, total, bold)
    for c, width in enumerate([16, 5, 20, 6, 9, 26, 8, 9]):
        ws.col(c).width = 256 * width


def _add_other_flowers_sheet(wb, by_mark, bold):
    """Лист «другие цветы» - всё, что не розы (образец закупщика 2026-10-01)."""
    ws = wb.add_sheet("другие цветы")
    ws.write(0, 1, "Позиции не по розе (копия из основного листа, справочно)", bold)
    head = xlwt.easyxf("font: bold on; " + _BORDERS)
    for c, title in enumerate(["FARM", "номер коробки", "PRODUCT", "VARIETY", " GRADE",
                               "TOTAL STEMS", "МЕТКА"], start=1):
        ws.write(3, c, title, head)
    r, total = 4, 0
    for mark, box_no, b, it, _first in _iter_rows(by_mark):
        if _is_roses(b):
            continue
        values = [b.get("farm"), box_no, it.get("product") or b.get("product"),
                  it.get("variety"), _grade(it), it.get("stems"), mark]
        for c, v in enumerate(values, start=1):
            ws.write(r, c, "" if v is None else v, _CELL)
        total += it.get("stems") or 0
        r += 1
    if r == 4:
        ws.write(4, 1, "В этой партии всё - розы.")
    else:
        ws.write(r + 1, 4, "ИТОГО", bold)
        ws.write(r + 1, 6, total, bold)


def _rename_reason(item):
    """Почему сорт в строке называется не так, как в инвойсе."""
    variety = (item.get("variety") or "").upper()
    if variety == combine.MOON_MIX_NAME:
        return "гвоздика серии Moon - всегда MOON MIX"
    if variety == "ALSTRO":
        return "альстромерия - сорт не отслеживается"
    if "MIX COLOR" in (item.get("renamed_from") or "").upper():
        return "плантация уже отгрузила микс"
    return "гвоздика - всегда MIX"


def _add_explanation_sheet(wb, by_mark, bold):
    """Лист "объяснение": что именно и почему было сведено в строку MIX.

    Требование закупщика 2026-09-10 - он должен видеть каждое объединение,
    чтобы либо согласиться, либо поправить, прежде чем пускать Invoice total
    в работу. Поэтому лист добавляется ВСЕГДА, даже если объединений не было.
    """
    ws = wb.add_sheet("объяснение")
    ws.write(0, 0, "Что было объединено в строку MIX", bold)
    ws.write(1, 0, "Правила: гвоздика - всегда MIX (сорт не отслеживается); "
                    "розы объединяются только короче 70 см и только если сортов от "
                    f"{combine.ROSES_MIX_MIN_VARIETIES} и у каждого не больше двух пачек "
                    f"({combine.ROSES_MIX_MAX_STEMS_PER_VARIETY} стеблей), иначе посортово; "
                    "альстромерия - всегда ALSTRO, ростовка по верхней в коробке; "
                    "гвоздики серии Moon (MOONLITE, MOONAQUA...) - всегда MOON MIX; "
                    "остальные культуры объединяются всегда. "
                    "Объединяются только позиции с ОДИНАКОВОЙ длиной и ценой.")

    headers = ["Метка", "Коробка №", "FARM", "PRODUCT", "Длина, см", "Цена",
               "Стеблей итого", "Сумма USD", "Что объединено / заменено"]
    for c, title in enumerate(headers):
        ws.write(3, c, title, bold)

    r = 4
    box_no = 0  # сквозная нумерация, как на основном листе
    for mark, info in by_mark.items():
        for b in info["boxes"]:
            box_no += 1
            for it in b["items"]:
                merged_from = it.get("merged_from")
                renamed_from = it.get("renamed_from")
                if not merged_from and not renamed_from:
                    continue
                if merged_from:
                    note = f"объединено {len(merged_from)} сортов: " + ", ".join(merged_from)
                else:
                    note = f"переименовано из «{renamed_from}» ({_rename_reason(it)})"
                for c, value in enumerate([mark, box_no, b.get("farm"),
                                            it.get("product") or b.get("product"),
                                            it.get("length_cm"), it.get("price"),
                                            it.get("stems"), it.get("total"), note]):
                    if value is not None:
                        ws.write(r, c, value)
                r += 1

    if r == 4:
        ws.write(4, 0, "Объединений не было - все позиции выведены как в инвойсах.")


def _add_roses_sheet(wb, by_mark, bold):
    """Лист "roses": справочная КОПИЯ позиций розовых плантаций.

    Из основного листа ничего не вырезается - это дубль для удобства
    (требование закупщика 2026-09-10). Колонки - по образцу закупщика
    2026-10-01: Farm / Variety / Длина / Стеблей / Метка, у строк позиций
    рамки, как на основном листе.
    """
    ws = wb.add_sheet("roses")
    ws.write(0, 0, "Позиции с розовых плантаций (копия из основного листа, справочно)", bold)

    headers = ["Farm", "Variety", "Длина", "Стеблей", "Метка"]
    for c, title in enumerate(headers):
        ws.write(2, c, title, bold)

    r = 3
    total_stems = 0
    for mark, _box_no, b, it, _first in _iter_rows(by_mark):
        if not _is_roses(b):
            continue
        values = [b.get("farm"), it.get("variety"), _grade(it), it.get("stems"), mark]
        for c, value in enumerate(values):
            ws.write(r, c, "" if value is None else value, _CELL)
        total_stems += it.get("stems") or 0
        r += 1

    if r == 3:
        ws.write(3, 0, "В этой партии позиций с розовых плантаций нет.")
    else:
        r += 1
        ws.write(r, 1, "ИТОГО", bold)
        ws.write(r, 3, total_stems, bold)
