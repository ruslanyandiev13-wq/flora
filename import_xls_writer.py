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
            ws.write(r, c, "", _CELL)


def _box_start_style(box_size):
    if box_size is not None and box_size < 0.25:
        return _BOX_START_STYLES["light_green"]
    if box_size == 0.25:
        return _BOX_START_STYLES["gold"]
    return _BOX_START_STYLES["yellow"]


def build_combined_factura_xls(output_path, by_mark, awb_doc=None):
    """Единый файл на всю партию (одна авианакладная): все коробки всех меток
    идут одной таблицей со сквозной нумерацией, метка вынесена в отдельную
    колонку (колонка 2 - в подтверждённом образце factura она пустая).
    Подтверждено закупщиком 2026-09-09 ("Да, в один. Invoice total").

    by_mark: результат import_combine.combine_by_mark(), уже дополненный
        AWB-блоком (см. import_app.py::_apply_awb).
    awb_doc: разобранная авианакладная (import_parser.parse_awb_pdf), если была
        загружена - из неё берётся номер AWB и общие вес/ставка партии.
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
    if awb_doc.get("total_awb"):
        total_awb = awb_doc["total_awb"]
    total_gross = round(sum((i.get("awb") or {}).get("gross_weight") or 0 for i in by_mark.values()), 2)
    total_chargeable = round(sum((i.get("awb") or {}).get("chargeable_weight") or 0
                                  for i in by_mark.values()), 2)
    # "ставка за кг" - как у закупщика: итог по накладной / платный вес
    # (включает сборы, поэтому не равна тарифу из AWB).
    rate = round(total_awb / total_chargeable, 4) if total_awb and total_chargeable else next(
        ((i.get("awb") or {}).get("rate_per_kg") for i in by_mark.values()
         if (i.get("awb") or {}).get("rate_per_kg")), None)

    # Шапка
    w(0, 0, "TO ")
    w(1, 0, "CONSIGNEE ")
    w(1, 1, ", ".join(marks))
    w(1, 11, "DATE")
    w(1, 16, first("invoice_date"))

    w(2, 0, "DESTINATION")
    w(2, 1, first("destination"))
    w(2, 11, "CARGO AGENCY")
    w(2, 16, first("forwarder"))

    w(4, 0, "CITY/COUNTRY")
    w(4, 11, "AIRLINE")
    w(4, 16, first("airline"))

    w(5, 0, "TELEPHONE")
    w(5, 11, "M.A.W.B")
    w(5, 16, awb_doc.get("awb_no") or first("awb_number"))

    w(6, 0, "FAX")
    w(6, 11, "H.A.W.B")
    w(6, 16, first("hawb_number"))

    w(7, 0, "EMAIL")
    w(7, 11, "TOTAL FULL BOXES")
    w(7, 16, total_full_boxes)

    HEADER_ROW = 9
    w(HEADER_ROW, 0, "FARM", bold)
    w(HEADER_ROW, 2, "МЕТКА", bold)
    w(HEADER_ROW, 3, "PRODUCT", bold)
    w(HEADER_ROW, 5, "BOX", bold)
    w(HEADER_ROW, 6, " BOX SIZE", bold)
    w(HEADER_ROW, 7, "VARIETY", bold)
    w(HEADER_ROW, 10, " GRADE", bold)
    w(HEADER_ROW, 12, "TOTAL STEMS", bold)
    w(HEADER_ROW, 14, "UNIT PRICE", bold)
    w(HEADER_ROW, 15, "TOTAL USD", bold)
    w(HEADER_ROW, 18, "OBS", bold)
    if rate is not None:
        w(HEADER_ROW, 21, rate)
        w(HEADER_ROW, 22, "ставка за кг")

    r = HEADER_ROW + 1
    box_no = 0
    for mark, info in by_mark.items():
        for b in info["boxes"]:
            box_no += 1
            first_item = True
            for it in b["items"]:
                filled.clear()
                w(r, 0, b.get("farm"), _CELL)
                w(r, 1, box_no, _CELL)
                w(r, 2, mark, _CELL)
                # PRODUCT построчно: у спрей-роз он свой ("SPRAY ROSES corazon"),
                # и в одной коробке такие строки соседствуют с обычными.
                w(r, 3, it.get("product") or b.get("product"), _CELL)
                if first_item:
                    w(r, 5, 1, _CELL)
                    w(r, 6, b.get("box_size"), _CELL)
                w(r, 7, it.get("variety"),
                  _box_start_style(b.get("box_size")) if first_item else _CELL)
                # GRADE: длина в см, а если она нечисловая (FANCY/SELECT/
                # 1000GR у Astoria) - текст как в инвойсе.
                w(r, 10, it.get("length_cm") if it.get("length_cm") is not None
                   else it.get("grade_text"), _CELL)
                w(r, 12, it.get("stems"), _CELL)
                w(r, 14, it.get("price"), _CELL)
                w(r, 15, it.get("total"), _CELL)
                w(r, 16, it.get("total"))
                _frame_row(ws, r, filled)
                r += 1
                first_item = False

    # Итоговая строка прямо под таблицей - как в файлах закупщика: количество
    # коробок, сумма долей коробок, стебли и деньги в СВОИХ колонках.
    w(r, 5, box_no, bold)
    w(r, 6, total_full_boxes, bold)
    w(r, 12, total_stems, bold)
    w(r, 15, total_fob, bold)
    w(r, 16, total_fob, bold)

    r += 2
    w(r, 9, "TOTAL STEMS", bold)
    w(r, 13, total_stems)
    r += 1
    w(r, 9, "TOTAL FLOWERS FOB USD", bold)
    w(r, 19, total_fob)
    r += 3
    w(r, 9, "TOTAL CHARGEABLE WEIGHT(Kg)", bold)
    w(r, 13, total_chargeable or None)
    r += 1
    w(r, 9, "GROSS WEIGHT", bold)
    w(r, 13, total_gross or None)
    r += 3
    w(r, 9, "TOTAL AWB", bold)
    w(r, 19, total_awb or None)
    r += 1
    w(r, 9, "TOTAL USD", bold)
    w(r, 19, round(total_fob + total_awb, 2))

    _add_explanation_sheet(wb, by_mark, bold)
    _add_roses_sheet(wb, by_mark, bold)

    wb.save(output_path)
    return output_path


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
    (требование закупщика 2026-09-10).
    """
    ws = wb.add_sheet("roses")
    ws.write(0, 0, "Позиции с розовых плантаций (копия из основного листа, справочно)", bold)

    headers = ["Метка", "Коробка №", "FARM", "PRODUCT", "Доля коробки",
               "VARIETY", "Длина, см", "Стеблей", "Цена", "Сумма USD"]
    for c, title in enumerate(headers):
        ws.write(2, c, title, bold)

    r = 3
    total_stems = total_fob = 0
    box_no = 0  # сквозная нумерация, как на основном листе
    for mark, info in by_mark.items():
        for b in info["boxes"]:
            box_no += 1
            # "SPRAY ROSES corazon" - тоже розы, поэтому ищем ROSES в любом
            # месте названия, а не только в начале.
            if "ROSES" not in (b.get("product") or "").upper():
                continue
            first = True
            for it in b["items"]:
                values = [mark, box_no, b.get("farm"), it.get("product") or b.get("product"),
                          b.get("box_size") if first else None,
                          it.get("variety"),
                          it.get("length_cm") if it.get("length_cm") is not None else it.get("grade_text"),
                          it.get("stems"),
                          it.get("price"), it.get("total")]
                for c, value in enumerate(values):
                    if value is not None:
                        ws.write(r, c, value)
                total_stems += it.get("stems") or 0
                total_fob += it.get("total") or 0
                r += 1
                first = False

    if r == 3:
        ws.write(3, 0, "В этой партии позиций с розовых плантаций нет.")
    else:
        r += 1
        ws.write(r, 5, "ИТОГО", bold)
        ws.write(r, 7, total_stems)
        ws.write(r, 9, round(total_fob, 2))


def build_factura_xls(output_path, mark, info):
    """
    mark: строка-метка (совпадает с CONSIGNEE).
    info: элемент словаря, который возвращает import_combine.combine_by_mark(),
        дополненный AWB-блоком (см. import_app.py::factura() - info["awb"] =
        {"pieces","gross_weight","chargeable_weight","rate_per_kg",
         "weight_per_box","total_awb","awb_per_box","box_count_mismatch"}).
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

    awb = info.get("awb") or {}

    # Шапка
    w(0, 0, "TO ")
    w(1, 0, "CONSIGNEE ")
    w(1, 1, mark)
    w(1, 11, "DATE")
    w(1, 16, info.get("invoice_date"))

    w(2, 0, "DESTINATION")
    w(2, 1, info.get("destination"))
    w(2, 11, "CARGO AGENCY")
    w(2, 16, info.get("forwarder"))

    w(4, 0, "CITY/COUNTRY")
    w(4, 11, "AIRLINE")
    w(4, 16, info.get("airline"))

    w(5, 0, "TELEPHONE")
    w(5, 11, "M.A.W.B")
    w(5, 16, info.get("awb_number"))

    w(6, 0, "FAX")
    w(6, 11, "H.A.W.B")
    w(6, 16, info.get("hawb_number"))

    w(7, 0, "EMAIL")
    w(7, 11, "TOTAL FULL BOXES")
    w(7, 16, info.get("total_full_boxes"))

    HEADER_ROW = 9
    w(HEADER_ROW, 0, "FARM", bold)
    w(HEADER_ROW, 3, "PRODUCT", bold)
    w(HEADER_ROW, 5, "BOX", bold)
    w(HEADER_ROW, 6, " BOX SIZE", bold)
    w(HEADER_ROW, 7, "VARIETY", bold)
    w(HEADER_ROW, 10, " GRADE", bold)
    w(HEADER_ROW, 12, "TOTAL STEMS", bold)
    w(HEADER_ROW, 14, "UNIT PRICE", bold)
    w(HEADER_ROW, 15, "TOTAL USD", bold)
    w(HEADER_ROW, 18, "OBS", bold)
    rate = awb.get("rate_all_in") or awb.get("rate_per_kg")
    if rate is not None:
        w(HEADER_ROW, 21, rate)
        w(HEADER_ROW, 22, "ставка за кг")

    r = HEADER_ROW + 1
    for b in info["boxes"]:
        first = True
        for it in b["items"]:
            filled.clear()
            w(r, 0, b.get("farm"), _CELL)
            w(r, 1, b["box_no"], _CELL)
            w(r, 3, it.get("product") or b.get("product"), _CELL)
            if first:
                w(r, 5, 1, _CELL)
                w(r, 6, b.get("box_size"), _CELL)
            w(r, 7, it.get("variety"),
              _box_start_style(b.get("box_size")) if first else _CELL)
            w(r, 10, it.get("length_cm") if it.get("length_cm") is not None
               else it.get("grade_text"), _CELL)
            w(r, 12, it.get("stems"), _CELL)
            w(r, 14, it.get("price"), _CELL)
            w(r, 15, it.get("total"), _CELL)
            w(r, 16, it.get("total"))
            _frame_row(ws, r, filled)
            r += 1
            first = False

    # Итоговая строка прямо под таблицей (как в файлах закупщика).
    w(r, 5, len(info["boxes"]), bold)
    w(r, 6, info.get("total_full_boxes"), bold)
    w(r, 12, info.get("total_stems"), bold)
    w(r, 15, info.get("total_fob"), bold)
    w(r, 16, info.get("total_fob"), bold)

    r += 2
    w(r, 9, "TOTAL STEMS", bold)
    w(r, 13, info.get("total_stems"))
    r += 1
    w(r, 9, "TOTAL FLOWERS FOB USD", bold)
    w(r, 19, info.get("total_fob"))
    r += 3
    w(r, 9, "TOTAL CHARGEABLE WEIGHT(Kg)", bold)
    w(r, 13, awb.get("chargeable_weight"))
    r += 1
    w(r, 9, "GROSS WEIGHT", bold)
    w(r, 13, awb.get("gross_weight"))
    r += 3
    total_awb = awb.get("total_awb")
    w(r, 9, "TOTAL AWB", bold)
    w(r, 19, total_awb)
    r += 1
    total_fob = info.get("total_fob") or 0
    w(r, 9, "TOTAL USD", bold)
    w(r, 19, round(total_fob + total_awb, 2) if total_awb is not None else total_fob)

    wb.save(output_path)
    return output_path
