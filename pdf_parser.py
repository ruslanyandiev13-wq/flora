"""
Парсер PDF-инвойсов от голландских поставщиков цветов.

Сейчас поддерживается шаблон "MH Flowers" (таблица Doos/Fust/Aantal/...).
Архитектура рассчитана на добавление новых шаблонов: каждый шаблон - это
отдельная функция detect_*() + parse_*(), которая регистрируется в TEMPLATES.
"""
import re
import pdfplumber


def _cluster_rows(words, y_tol=3.0):
    """Группирует слова в строки по вертикальной координате 'top' с допуском."""
    words = sorted(words, key=lambda w: w['top'])
    rows = []
    current = []
    current_top = None
    for w in words:
        if current_top is None or abs(w['top'] - current_top) <= y_tol:
            current.append(w)
            current_top = w['top'] if current_top is None else current_top
        else:
            rows.append(current)
            current = [w]
            current_top = w['top']
    if current:
        rows.append(current)
    return rows


def _col_for(x0, columns):
    for name, a, b in columns:
        if a <= x0 < b:
            return name
    return "other"


NUM_RE = re.compile(r"^-?\d+([.,]\d+)?$")


def _to_float(s):
    if s is None:
        return None
    s = s.strip().replace(" ", "")
    if not s:
        return None
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _to_int(s):
    f = _to_float(s)
    return int(f) if f is not None else None


# ---------------------------------------------------------------------------
# Шаблон "MH Flowers"
# ---------------------------------------------------------------------------

MHFLOWERS_HEADER_WORDS = {"Doos", "Fust", "Aantal", "Omschrijving", "Kweker",
                           "Lengte", "Gew.", "Prijs", "Bedrag"}


def detect_mhflowers(pdf):
    """Возвращает True, если PDF похож на шаблон MH Flowers."""
    first_page_text = pdf.pages[0].extract_text() or ""
    return "MH Flowers" in first_page_text or "mhflowers" in first_page_text.lower()


def _find_header_columns(page):
    words = page.extract_words()
    header = {}
    for w in words:
        if w['text'] in MHFLOWERS_HEADER_WORDS and 'Doos' == w['text'] or w['text'] in MHFLOWERS_HEADER_WORDS:
            header.setdefault(w['text'], []).append(w)
    # Doos, Fust(1st), Aantal, Omschrijving, Kweker, Lengte, Gew., Fust(2nd), Prijs, Bedrag
    # Нам нужна их x0-координата на странице, где есть весь набор заголовков
    return header


def _build_columns(header_words):
    """header_words: dict text -> list-of-word-dicts (может быть 2 Fust)."""
    # Берём x0 первого вхождения каждого нужного слова
    def x0(text, idx=0):
        return header_words[text][idx]['x0']

    doos = x0("Doos")
    fust1 = x0("Fust", 0)
    aantal = x0("Aantal")
    omschr = x0("Omschrijving")
    lengte = x0("Lengte")
    gew = x0("Gew.")
    fust2 = x0("Fust", 1) if len(header_words["Fust"]) > 1 else gew + 35
    prijs = x0("Prijs")
    bedrag = x0("Bedrag")

    # строим границы колонок как средние точки между соседними x0
    names_order = [("doos", doos), ("fust", fust1), ("aantal", aantal),
                    ("omschrijving", omschr),
                    ("lengte", lengte), ("gew", gew), ("fust2", fust2),
                    ("prijs", prijs), ("bedrag", bedrag)]
    # Колонка Kweker (производитель) есть не во всех инвойсах MH Flowers -
    # в части реальных файлов её нет вообще (см. правки от 2026-09).
    if "Kweker" in header_words:
        names_order.append(("kweker", x0("Kweker")))
    names_order.sort(key=lambda t: t[1])
    columns = []
    for i, (name, start) in enumerate(names_order):
        end = names_order[i + 1][1] if i + 1 < len(names_order) else 9999
        # немного расширяем влево, чтобы не терять слова у самой границы
        columns.append((name, start - 4, end - 4))
    return columns


def parse_mhflowers(pdf, source_filename=""):
    """
    Возвращает dict:
      {
        "invoice_no": str,
        "invoice_date": str,
        "currency": "EUR" (для Голландии всегда EUR),
        "boxes": [
            {"box_no": int, "fust": "AAA", "items": [
                {"aantal": 80, "omschrijving": "Li Ot Zambesi", "kweker": "...",
                 "lengte": 100.0, "gew": None, "fust_mult": "1xAAA",
                 "prijs": 0.83, "bedrag": 66.40}
            ]}
        ],
        "eenmalig_fust": [ {"aantal":6, "omschrijving":"Gerberadoos 12 cm", "fust_code":"612",
                             "fust_prijs":2.00, "totaal":12.00} ],
        "totals": {"subtotaal":..., "eenmalig_fust_totaal":..., "faktuurbedrag":..., "currency": "EUR"}
      }
    """
    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    invoice_no_m = re.search(r"Invoice nr\.\s*([0-9A-Za-z /\-]+)", full_text)
    invoice_date_m = re.search(r"Invoice date\s*([0-9\-]+)", full_text)
    invoice_no = invoice_no_m.group(1).strip() if invoice_no_m else None
    invoice_date = invoice_date_m.group(1).strip() if invoice_date_m else None

    boxes = []
    eenmalig_fust = []
    subtotaal = None
    faktuurbedrag = None
    faktuur_currency = None

    columns = None
    finished_items = False

    for page in pdf.pages:
        words = page.extract_words()
        header_words = _find_header_columns(page)
        required = {"Doos", "Aantal", "Omschrijving", "Lengte", "Gew.", "Prijs", "Bedrag"}
        if required.issubset(header_words.keys()):
            columns = _build_columns(header_words)

        if columns is None:
            continue  # страница ещё не содержит основную таблицу (напр. титул)

        rows = _cluster_rows(words)
        current_box = None  # текущий номер коробки для строк-продолжений

        for row_words in rows:
            row_words = sorted(row_words, key=lambda w: w['x0'])
            top_y = row_words[0]['top']
            # Пропускаем строки самого заголовка таблицы
            texts = [w['text'] for w in row_words]
            if texts[:2] == ["Doos", "Fust"] or "Omschrijving" in texts and "Kweker" in texts:
                continue
            if "Subtotaal" in texts:
                # Subtotaal    3644,6
                nums = [w['text'] for w in row_words if NUM_RE.match(w['text'])]
                if nums:
                    subtotaal = _to_float(nums[-1])
                continue
            if "Faktuurbedrag" in texts:
                # Faktuurbedrag  EUR  4299,30
                cur_m = re.search(r"\b(EUR|USD)\b", " ".join(texts))
                nums = [w['text'] for w in row_words if NUM_RE.match(w['text'])]
                if nums:
                    faktuurbedrag = _to_float(nums[-1])
                if cur_m:
                    faktuur_currency = cur_m.group(1)
                continue
            if "Eenmalig" in texts:
                finished_items = True
                continue
            if finished_items:
                continue
            if texts[:2] == ["Aantal", "Fust"] or (len(texts) >= 2 and texts[0] == "Aantal"):
                continue
            if "Totaal" in texts and "Eenm." in texts:
                continue
            if "Statistic" in texts or "Colli" in texts:
                continue

            rowdict = {}
            for w in row_words:
                c = _col_for(w['x0'], columns)
                rowdict.setdefault(c, []).append(w['text'])
            rowdict = {k: " ".join(v) for k, v in rowdict.items()}

            # Строка "eenmalig fust" внизу таблицы имеет вид:
            #   <aantal>  <omschrijving>  <fust_code>  <fust_huur> <fust_prijs> <totaal>
            # но т.к. использует те же колонки-заготовки, легче распознавать по контексту:
            # если после "bedrag" пусто и в fust2 стоит числовой код, а не "NxCODE" - это скорее
            # eenmalig-fust таблица. Для надёжности разбираем её отдельным блоком по тексту (см. ниже).

            aantal = _to_int(rowdict.get("doos")) if rowdict.get("doos") else None
            fust = rowdict.get("fust")
            row_aantal = _to_int(rowdict.get("aantal"))
            omschr = rowdict.get("omschrijving")
            kweker = rowdict.get("kweker")
            lengte = _to_float(rowdict.get("lengte"))
            gew = _to_float(rowdict.get("gew"))
            fust_mult = rowdict.get("fust2")
            prijs = _to_float(rowdict.get("prijs"))
            bedrag = _to_float(rowdict.get("bedrag"))

            # Иногда PDF склеивает без пробела мульти-код тары и цену за штуку,
            # напр. "1xAAA,2xPPS0,65" (должно быть "1xAAA,2xPPS" + Prijs "0,65") -
            # тогда вся строка целиком попадает в колонку fust2, а колонка prijs
            # остаётся пустой. Иногда при этом последняя буква кода и первая цифра
            # цены ещё и переставляются местами при извлечении текста из PDF
            # (напр. "1xAAA,2xPP0S,65" вместо "1xAAA,2xPPS0,65"). Распознаём оба варианта.
            if prijs is None and fust_mult and "," in fust_mult:
                m = re.match(r"^(.*)([0-9])([A-Za-z]),(\d{2})$", fust_mult)
                if m:
                    fust_mult = m.group(1) + m.group(3)
                    prijs = _to_float(m.group(2) + "," + m.group(4))
                else:
                    m2 = re.match(r"^(.*[A-Za-z])(\d{1,3},\d{2})$", fust_mult)
                    if m2:
                        fust_mult = m2.group(1)
                        prijs = _to_float(m2.group(2))

            if not omschr and row_aantal is None:
                continue  # пустая/служебная строка

            # Строки с "0,00 / 0,00" в начале инвойса (Gerberadoos mini, PP и т.д.) - пропускаем,
            # это преамбула, не относится к товару.
            if row_aantal == 0 and (bedrag == 0 or bedrag is None) and aantal is None:
                continue

            if aantal is not None and fust:
                if current_box is not None and current_box["box_no"] == aantal:
                    # Та же коробка, что и раньше (PDF повторяет номер/тару на каждой
                    # строке товара внутри одной коробки) - просто добавляем позицию.
                    pass
                else:
                    # Начало новой коробки
                    current_box = {"box_no": aantal, "fust": fust, "items": []}
                    boxes.append(current_box)
                if row_aantal is not None and omschr:
                    current_box["items"].append({
                        "aantal": row_aantal, "omschrijving": omschr, "kweker": kweker,
                        "lengte": lengte, "gew": gew, "fust_mult": fust_mult,
                        "prijs": prijs, "bedrag": bedrag,
                    })
            elif row_aantal is not None and omschr and current_box is not None:
                # Продолжение той же коробки (несколько позиций в одной коробке)
                current_box["items"].append({
                    "aantal": row_aantal, "omschrijving": omschr, "kweker": kweker,
                    "lengte": lengte, "gew": gew, "fust_mult": fust_mult,
                    "prijs": prijs, "bedrag": bedrag,
                })

    # --- Разбор блока "Eenmalig Fust" отдельным текстовым проходом (таблица тары) ---
    # Пример строки: "6   Gerberadoos 12 cm   612   0,00   2,00   12,00"
    eenmalig_block = re.search(
        r"Eenmalig Fust\s+Aantal\s+Fust Omschrijving\s+Fust code\s+Fust huur\s+Fust prijs\s+Totaal(.*?)Totaal Eenm\. Fust",
        full_text, re.S)
    if eenmalig_block:
        block_text = eenmalig_block.group(1)
        for line in block_text.strip().splitlines():
            line = line.strip()
            if not line:
                continue
            m = re.match(r"^(\d+)\s+(.*?)\s+([A-Za-z0-9]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)$", line)
            if m:
                eenmalig_fust.append({
                    "aantal": _to_int(m.group(1)),
                    "omschrijving": m.group(2).strip(),
                    "fust_code": m.group(3),
                    "fust_huur": _to_float(m.group(4)),
                    "fust_prijs": _to_float(m.group(5)),
                    "totaal": _to_float(m.group(6)),
                })

    return {
        "source_filename": source_filename,
        "invoice_no": invoice_no,
        "invoice_date": invoice_date,
        "boxes": boxes,
        "eenmalig_fust": eenmalig_fust,
        "totals": {
            "subtotaal": subtotaal,
            "faktuurbedrag": faktuurbedrag,
            "currency": faktuur_currency or "EUR",
        },
    }


TEMPLATES = [
    ("mhflowers", detect_mhflowers, parse_mhflowers),
]


def parse_invoice_pdf(path):
    with pdfplumber.open(path) as pdf:
        for name, detect, parse in TEMPLATES:
            if detect(pdf):
                return parse(pdf, source_filename=path), name
        raise ValueError(
            "Не удалось распознать шаблон поставщика в этом PDF. "
            "Нужно добавить новый шаблон разбора (см. pdf_parser.py -> TEMPLATES)."
        )
