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

    def w(r, c, v, style=normal):
        if v is not None:
            ws.write(r, c, v, style)

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
    if awb.get("rate_per_kg") is not None:
        w(HEADER_ROW, 21, awb["rate_per_kg"])
        w(HEADER_ROW, 22, "ставка за кг")

    r = HEADER_ROW + 1
    for b in info["boxes"]:
        first = True
        for it in b["items"]:
            w(r, 0, b.get("farm"))
            w(r, 1, b["box_no"])
            w(r, 3, b.get("product"))
            if first:
                w(r, 5, 1)
                w(r, 6, b.get("box_size"))
            w(r, 7, it.get("variety"))
            w(r, 10, it.get("length_cm"))
            w(r, 12, it.get("stems"))
            w(r, 14, it.get("price"))
            w(r, 15, it.get("total"))
            w(r, 16, it.get("total"))
            r += 1
            first = False

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
