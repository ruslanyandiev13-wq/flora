"""
Сборка итогового .xls (формат Excel 97-2003) в точности по структуре, которую
использует бухгалтер для загрузки в 1С (проверено на реальном файле
"Голландия Siri & Chrys MH на 19.07 Трак Вологда для загрузки.xls").

Структура:
  Row0: '', 'PETIA', '', '', '', 'Paklijst Export Unie Flora', '', 'Week', '', 'Vrijdag', '', 'Blad', '', ''
  Row1: пусто
  Row2: '', week_no, '', excel_date_serial, '', blad_no, ...
  Row3: пусто
  Row4: '', 'Debnr.', 'Sub', 'Naam', ...
  Row5: '', debnr, '', naam, ...
  Row6: '', 'Палета,$', pallet_cost_usd, '', 'КурсUSD', usd_rate, ...  (для расчёта логистики в 1С)
  Row7: '', 0, '', 'Kolli','Inhoud','Totaal','Omschrijving','Kweker','Prijs','S1','S2','S3','Fust','Bedrag'  (статичная строка)
  Row8: пусто
  далее для каждой коробки:
    строка-шапка коробки:  '', box_no, fust_code, '', 'Prijs', '', 'S1', 'S2', 's3', '', '', 'Totaal', '', ''
    строка(и)-позиции:     '', aantal, omschrijving, '', prijs, '', lengte, gew, '', '', '', bedrag, '', ''
  далее пусто, пусто
  строка итога: '', 'Totaal Stelen :', total_stems(опционально), '', 'Totaal Stelen :', total_money, ...
"""
import datetime
import xlwt

# Строка-шапка таблицы товаров (где есть текст "Omschrijving"/"Kolli"/"Prijs" и
# т.д.) - на неё завязаны и позиция строки "Палета,$"/"КурсUSD" (прямо над ней),
# и начало строк с коробками (двумя строками ниже, после одной пустой строки).
HEADER_ROW = 7


def _excel_date_serial(py_date):
    epoch = datetime.date(1899, 12, 30)
    return (py_date - epoch).days


def build_xls(output_path, boxes, header, total_money, total_stems=None,
              pallet_cost_usd=None, usd_rate=None):
    """
    boxes: список коробок в объединённом (по всем инвойсам) и перенумерованном виде:
        [{"box_no": 1, "fust": "AAA", "items": [
            {"aantal":80, "omschrijving":"Li Ot Zambesi", "prijs":0.83,
             "lengte":100, "gew": None, "bedrag":66.4}, ...]}]
    header: dict с ключами:
        petia (str, по умолчанию 'PETIA'), paklijst_title (str),
        week_no (int), date (datetime.date), blad_no (int),
        debnr (str), naam (str)
    total_money: итоговая сумма (Totaal Stelen, денежная часть)
    total_stems: итоговое количество стеблей (опционально)
    pallet_cost_usd, usd_rate: используются 1С-обработкой для автоматического
        расчёта логистики по коробкам/товарам - пишутся подписанными ячейками
        "Палета,$"/"КурсUSD" прямо над строкой-шапкой таблицы товаров. Если
        значение ещё не задано пользователем - подпись пишется, число нет.
    """
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Лист1")

    bold = xlwt.easyxf("font: bold on")
    small = xlwt.easyxf("font: height 150")  # 7.5pt
    normal = xlwt.easyxf("")

    written_rows = set()

    def w(r, c, v, style=normal):
        ws.write(r, c, v, style)
        written_rows.add(r)

    # Row 0
    w(0, 1, header.get("petia", "PETIA"))
    w(0, 5, header.get("paklijst_title", "Paklijst Export Unie Flora"))
    w(0, 7, "Week")
    w(0, 9, "Vrijdag")
    w(0, 11, "Blad")

    # Row 2
    w(2, 1, header.get("week_no", 1))
    date_val = header.get("date")
    if isinstance(date_val, datetime.date):
        w(2, 3, _excel_date_serial(date_val))
    elif date_val is not None:
        w(2, 3, date_val)
    w(2, 5, header.get("blad_no", 1))

    # Row 4-5
    w(4, 1, "Debnr.")
    w(4, 2, "Sub")
    w(4, 3, "Naam")
    w(5, 1, header.get("debnr", ""))
    w(5, 3, header.get("naam", ""))

    # Строка над шапкой: "Палета,$" / "КурсUSD" - по ним 1С считает логистику.
    # Прежде чем писать, убеждаемся, что строка ещё ничем не занята (иначе можно
    # молча испортить файл, если структура выше когда-нибудь изменится).
    logistics_row = HEADER_ROW - 1
    if logistics_row in written_rows:
        raise RuntimeError(
            f"Строка {logistics_row} (прямо над шапкой таблицы товаров) уже "
            "содержит данные - запись 'Палета,$'/'КурсUSD' туда испортила бы файл."
        )
    w(logistics_row, 1, "Палета,$")
    if pallet_cost_usd is not None:
        w(logistics_row, 2, pallet_cost_usd)
    w(logistics_row, 4, "КурсUSD")
    if usd_rate is not None:
        w(logistics_row, 5, usd_rate)

    # Row 7 - статичная строка-разделитель (как в реальном файле бухгалтера)
    w(HEADER_ROW, 1, 0)
    for c, val in zip(range(3, 14), ["Kolli", "Inhoud", "Totaal", "Omschrijving",
                                       "Kweker", "Prijs", "S1", "S2", "S3", "Fust", "Bedrag"]):
        w(HEADER_ROW, c, val, small)

    r = HEADER_ROW + 2
    for box in boxes:
        w(r, 1, box["box_no"], bold)
        fust_val = box["fust"]
        if isinstance(fust_val, str) and fust_val.isdigit():
            fust_val = float(fust_val)
        w(r, 2, fust_val, bold)
        w(r, 4, "Prijs", small)
        w(r, 6, "S1", small)
        w(r, 7, "S2", small)
        w(r, 8, "s3", small)
        w(r, 11, "Totaal", small)
        r += 1
        for item in box["items"]:
            w(r, 1, item["aantal"])
            w(r, 2, item["omschrijving"])
            if item.get("prijs") is not None:
                w(r, 4, item["prijs"])
            if item.get("lengte") is not None:
                w(r, 6, item["lengte"])
            if item.get("gew") is not None:
                w(r, 7, item["gew"])
            if item.get("bedrag") is not None:
                w(r, 11, item["bedrag"])
            r += 1

    r += 2
    w(r, 1, "Totaal Stelen :")
    if total_stems is not None:
        w(r, 2, total_stems)
    w(r, 4, "Totaal Stelen :")
    w(r, 5, round(total_money, 2))

    wb.save(output_path)
    return output_path
