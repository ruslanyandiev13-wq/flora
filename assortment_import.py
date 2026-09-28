"""
Чтение файла ассортимента бухгалтера ("Голландия (ассортимент).xlsx") для
справочника `assortment`.

Формат её файла (подтверждён на реальном, 797 позиций): первая строка -
заголовки "(разбанч) | Название | Рост | Вес | комент", дальше по строке на
позицию. Колонки ищем по заголовку, а не по номеру - чтобы файл можно было
дополнять столбцами, не ломая загрузку.
"""
NAME_HEADERS = ("название", "наименование", "name")
BUNCH_HEADERS = ("разбанч", "(разбанч)", "банч", "bunch")
HEIGHT_HEADERS = ("рост", "высота", "length")
WEIGHT_HEADERS = ("вес", "weight")
COMMENT_HEADERS = ("комент", "комментарий", "коммент")


def _header_index(header_row, variants):
    for idx, value in enumerate(header_row):
        text = str(value or "").strip().lower()
        if text in variants:
            return idx
    return None


def read_xlsx(file_obj):
    """file_obj: путь или файловый объект (request.files[...]).
    Возвращает список dict'ов для db.replace_assortment().

    openpyxl импортируется здесь, а не на уровне модуля: если на сервере его
    вдруг не окажется, должна отвалиться только загрузка ассортимента, а не всё
    приложение (на проде такой импорт в app.py уже уронил сервис в 502)."""
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("на сервере не установлен пакет openpyxl "
                           "(pip install openpyxl) - без него .xlsx не прочитать")
    wb = openpyxl.load_workbook(file_obj, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if header is None:
        raise ValueError("файл пустой")

    i_name = _header_index(header, NAME_HEADERS)
    if i_name is None:
        raise ValueError('не найдена колонка "Название"')
    i_bunch = _header_index(header, BUNCH_HEADERS)
    i_height = _header_index(header, HEIGHT_HEADERS)
    i_weight = _header_index(header, WEIGHT_HEADERS)
    i_comment = _header_index(header, COMMENT_HEADERS)

    def cell(row, idx):
        if idx is None or idx >= len(row):
            return None
        value = row[idx]
        if value is None:
            return None
        text = str(value).strip()
        return text or None

    out = []
    seen = set()
    for row in rows:
        name = cell(row, i_name)
        if not name or name.lower() in seen:
            continue
        seen.add(name.lower())
        bunch = cell(row, i_bunch)
        try:
            bunch = int(float(bunch)) if bunch is not None else None
        except ValueError:
            bunch = None
        out.append({"name": name, "bunch": bunch,
                    "height": cell(row, i_height), "weight": cell(row, i_weight),
                    "comment": cell(row, i_comment)})
    if not out:
        raise ValueError("в файле нет ни одной позиции")
    return out
