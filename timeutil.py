"""
Единое время биллинга - GMT+3 (Москва), независимо от часового пояса сервера.

На проде systemd-сервис работает в UTC, локально - в часовом поясе машины,
из-за чего в истории обработки и в ленте токенов время "уезжало". Все новые
отметки времени в биллинге пишутся через now_str()/today() отсюда, формат
хранения не меняется (наивный ISO "YYYY-MM-DDTHH:MM:SS") - на нём построены
фильтры по датам и группировка по дням в billing/db.py (substr(datetime,1,10),
сравнения строк >= / <=). Смещение зашито константой: переход на летнее время
в GMT+3 отсутствует.
"""
import datetime

TZ = datetime.timezone(datetime.timedelta(hours=3))
TZ_LABEL = "GMT+3"


def now():
    """Текущее время в GMT+3 (naive - как и всё, что уже лежит в базе)."""
    return datetime.datetime.now(TZ).replace(tzinfo=None)


def now_str():
    """Отметка времени для записи в базу: 'YYYY-MM-DDTHH:MM:SS' в GMT+3."""
    return now().isoformat(timespec="seconds")


def today():
    """Сегодняшняя дата по GMT+3 (может отличаться от даты сервера в UTC)."""
    return now().date()


def format_dt(value):
    """Для показа в шаблонах: 'YYYY-MM-DDTHH:MM:SS' -> '11.09.2026 14:35'.

    Значение, которое не удалось разобрать (в т.ч. записи, сделанные до
    перехода на GMT+3, если формат отличался), возвращаем как есть."""
    if not value:
        return ""
    text = str(value).strip().replace("T", " ")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            dt = datetime.datetime.strptime(text, fmt)
        except ValueError:
            continue
        return dt.strftime("%d.%m.%Y") if fmt == "%Y-%m-%d" else dt.strftime("%d.%m.%Y %H:%M")
    return str(value)
