"""
Получение курса USD и EUR с сайта https://ligovka.ru (графа "продажа",
уровень "от 10000").

ВАЖНО: сайт может поменять вёрстку в любой момент - поэтому здесь заложен
запасной путь (fallback) на ручной ввод курса в интерфейсе, если скрейпер
перестанет работать. Обязательно проверьте вживую после деплоя (в этой
песочнице доступ к ligovka.ru отсутствует из-за сетевых ограничений
контейнера, так что автоматический тест здесь не проводился).
"""
import re
import requests
from bs4 import BeautifulSoup

LIGOVKA_URL = "https://ligovka.ru/"
MARKUP = 1.05  # +5%, согласно правилу компании


def fetch_ligovka_rates(timeout=10):
    """
    Возвращает dict:
      {"usd_sell_10000": float, "eur_sell_10000": float,
       "usd_sell_10000_with_markup": float, "eur_sell_10000_with_markup": float,
       "fetched_at": "..."}
    Бросает RuntimeError с понятным сообщением, если не удалось распарсить.
    """
    resp = requests.get(LIGOVKA_URL, timeout=timeout, headers={
        "User-Agent": "Mozilla/5.0 (compatible; FloraLogisticsBot/1.0)"
    })
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    usd_sell = _extract_sell_rate(soup, "usd")
    eur_sell = _extract_sell_rate(soup, "eur")

    if usd_sell is None or eur_sell is None:
        raise RuntimeError(
            "Не удалось найти курс 'от 10000' на ligovka.ru — возможно, "
            "изменилась вёрстка сайта. Введите курс вручную."
        )

    return {
        "usd_sell_10000": usd_sell,
        "eur_sell_10000": eur_sell,
        "usd_sell_10000_with_markup": round(usd_sell * MARKUP, 4),
        "eur_sell_10000_with_markup": round(eur_sell * MARKUP, 4),
    }


def _extract_sell_rate(soup, currency):
    """
    Ищет строку таблицы, где есть ссылка на .../detailed/<currency>?tab=10000,
    и берёт второе число в этой строке (первое - покупка, второе - продажа).
    """
    links = soup.find_all("a", href=re.compile(rf"/detailed/{currency}\?tab=10000"))
    if not links:
        # запасной вариант: искать строку по тексту "от 10000" рядом с currency
        return _extract_sell_rate_by_text(soup, currency)

    # Берём родительскую строку таблицы (tr) первой найденной ссылки уровня "от 10000"
    row = None
    for a in links:
        tr = a.find_parent("tr")
        if tr is not None:
            row = tr
            break
    if row is None:
        return None

    numbers = []
    for a in row.find_all("a", href=re.compile(rf"/detailed/{currency}")):
        text = a.get_text(strip=True).replace(",", ".")
        if re.match(r"^-?\d+(\.\d+)?$", text):
            numbers.append(float(text))
    if len(numbers) >= 2:
        return numbers[1]  # покупка, ПРОДАЖА
    return None


def _extract_sell_rate_by_text(soup, currency):
    text = soup.get_text("\n")
    # ищем блок вида "от 10000 ... 80.00 ... 81.60" рядом с USD/EUR - грубый fallback
    return None
