# =====================================================================
#  Скринер облигаций MOEX + Доход.ру
#  Single-file версия: без эмодзи, с исправленными багами и ускорениями
# =====================================================================
from __future__ import annotations

import hashlib
import html as _html
import io
import json
import re
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from plotly.subplots import make_subplots
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

st.set_page_config(page_title="Скринер облигаций QCKSLVR TRADING", layout="wide")

# ================== НАСТРОЙКИ ==================
BASE_URL = "https://iss.moex.com/iss"
DOHOD_BOND_URL = "https://analytics.dohod.ru/bond/"

TODAY = date.today()
DEFAULT_DATE = TODAY                                   # было date(2026,10,21) — в прошлом
DEFAULT_FORECAST_TARGET_DATE = date(TODAY.year + 1, 1, 1)
DEFAULT_EXTREMUM_DATE = TODAY - timedelta(days=120)
DEFAULT_NDFL = 13.0
DEFAULT_CURRENT_CBR = 14.0
DEFAULT_FORECAST_CBR = 13.75
DEFAULT_MAX_DURATION = 10.0
DEFAULT_MIN_DAYS = 0
DEFAULT_MIN_VOLUME = 0
DEFAULT_MIN_TRADES = 0
DEFAULT_MAX_GROUP = 6
DEFAULT_MIN_POTENTIAL_RETURN = -100.0
DEFAULT_MAX_SPREAD = 5.0
DEFAULT_MIN_CREDIT_QUALITY = 0.0
DEFAULT_MAX_COMPLEXITY = 7.0
DOHOD_FRESH_DAYS = 30

DAILY_BARS = 520
WEEKLY_BARS = 175

PAGE_SIZE = 100          # ISS отдаёт максимум 100 строк на страницу
MAX_PAGES = 120
HTTP_WORKERS = 8         # страниц параллельно внутри одной доски
BOARD_WORKERS = 2        # досок параллельно (итого <= 16 соединений)
HTTP_TIMEOUT = (5, 20)
HISTORY_LOOKBACK_DAYS = 5
TABLE_PAGE_SIZE = 200

BOND_BOARDS = ["TQCB", "TQOB", "TQOY", "TQOD", "TQIF", "TQTF",
               "TQBD", "TQTC", "TQRD", "TQIR", "TQUD"]
BOARD_PRIORITY = {"TQOB": 0, "TQCB": 1, "TQIF": 2, "TQOY": 3, "TQOD": 4,
                  "TQTF": 5, "TQBD": 6, "TQUD": 7, "TQTC": 8, "TQRD": 9, "TQIR": 10}
RUB_CURRENCIES = {"SUR", "RUB", "RUR"}

INCOME_TYPE_MAP = {
    "Эффективная (YTM)": "YTM, %",
    "К оферте (YTC)": "YTC, %",
    "Текущая (купон/цена)": "Текущая доходность, %",
    "Простая (без реинвест.)": "Простая доходность, %",
}

PRESETS = {
    "Замена депозита (консервативный)": {
        "flt_min_days_maturity": 365, "flt_max_days_maturity": 1095,
        "flt_max_duration": 2.0, "flt_max_group": 2, "flt_min_credit_quality": 7.0,
        "flt_max_complexity": 0.0, "flt_only_non_subordinated": True,
        "flt_min_volume": 5_000_000, "flt_max_spread": 0.3,
        "flt_min_potential_return": 3.0, "flt_currency_filter": ["SUR", "RUB"],
        "flt_board_filter": ["TQCB", "TQOB"], "flt_rate_change": 0.0,
        "flt_only_with_price": True, "flt_only_with_maturity": True,
    },
    "Сбалансированный": {
        "flt_min_days_maturity": 365, "flt_max_days_maturity": 1825,
        "flt_max_duration": 3.0, "flt_max_group": 3, "flt_min_credit_quality": 5.0,
        "flt_max_complexity": 1.0, "flt_only_non_subordinated": True,
        "flt_min_volume": 1_000_000, "flt_max_spread": 0.5,
        "flt_min_potential_return": 5.0, "flt_rate_change": -1.0,
        "flt_only_with_price": True, "flt_only_with_maturity": True,
    },
    "Агрессивный / высокая доходность": {
        "flt_min_days_maturity": 180, "flt_max_days_maturity": 3650,
        "flt_max_duration": 5.0, "flt_max_group": 5, "flt_min_credit_quality": 3.0,
        "flt_max_complexity": 3.0, "flt_only_non_subordinated": False,
        "flt_min_volume": 500_000, "flt_max_spread": 1.0,
        "flt_min_potential_return": 8.0, "flt_rate_change": -2.0,
        "flt_only_with_price": True, "flt_only_with_maturity": False,
    },
    "Ставка падает — длинные бумаги": {
        "flt_max_duration": 5.0, "flt_min_days_maturity": 730,
        "flt_max_group": 3, "flt_min_potential_return": 5.0,
        "flt_rate_change": -2.0, "flt_only_with_price": True,
    },
    "Ставка растёт — короткие бумаги": {
        "flt_max_duration": 1.0, "flt_max_days_maturity": 730,
        "flt_max_group": 2, "flt_min_potential_return": 0.0,
        "flt_rate_change": 2.0, "flt_only_with_price": True,
    },
}

EXCEL_COLS = {
    "Эмитент", "Основной заемщик", "Отрасль",
    "Кредитное качество", "Кредитный рэнкинг", "Качество эмитента",
    "Inside Q", "Outside Q", "NetDebt/Equity", "Коэф. ликвидности",
    "Сложность", "Размер выпуска (оценка)", "Медиана оборота, млн",
    "Объём в обращении, млрд", "Лет до даты (Доход.ру)",
    "Тип эмитента (Доход.ру)", "Субординированная", "С гарантией",
    "Название (Доход.ру)", "Тип купона (Доход.ру)",
}


def mark_x(c: str) -> str:
    return f"{c} (X)" if c in EXCEL_COLS else c


# ================== РЕШЕНИЯ ЦБ РФ ==================
# Вынесено во внешний файл: cbr_decisions.json рядом с app.py
# Формат: [{"date": "2026-12-18", "rate": null, "change_bp": null}, ...]
CBR_FILE = Path(__file__).parent / "cbr_decisions.json"

CBR_DECISIONS_FALLBACK = [
    (date(2024, 7, 26), 18.00, +200), (date(2024, 9, 13), 19.00, +100),
    (date(2024, 10, 25), 21.00, +200), (date(2024, 12, 20), 21.00, 0),
    (date(2025, 2, 14), 21.00, 0), (date(2025, 3, 21), 21.00, 0),
    (date(2025, 4, 25), 21.00, 0), (date(2025, 6, 6), 20.00, -100),
    (date(2025, 7, 25), 18.00, -200), (date(2025, 9, 12), 17.00, -100),
    (date(2025, 10, 24), 16.50, -50), (date(2025, 12, 19), 16.00, -50),
    (date(2026, 2, 13), 15.50, -50), (date(2026, 3, 20), 15.00, -50),
    (date(2026, 4, 24), 14.50, -50), (date(2026, 6, 19), 14.25, -25),
    (date(2026, 7, 24), 14.00, -25), (date(2026, 9, 11), 14.00, 0),
    (date(2026, 10, 23), 14.00, 0), (date(2026, 12, 18), None, None),
]


@st.cache_data(ttl=3600, show_spinner=False)
def load_cbr_decisions() -> list:
    try:
        if CBR_FILE.exists():
            raw = json.loads(CBR_FILE.read_text(encoding="utf-8"))
            return [(date.fromisoformat(r["date"]), r.get("rate"), r.get("change_bp"))
                    for r in raw]
    except Exception:
        pass
    return list(CBR_DECISIONS_FALLBACK)


# ================== ЗАГОЛОВКИ (без эмодзи-префикса) ==================
def display_name(col: str) -> str:
    """Перенос длинного заголовка. Разделитель сохраняется, \n ставится ПОСЛЕ него."""
    base = mark_x(col)
    if len(base) > 22:
        for sep in (", ", " "):
            if sep in base:
                parts = base.split(sep)
                mid = max(1, len(parts) // 2)
                return sep.join(parts[:mid]) + sep + "\n" + sep.join(parts[mid:])
    return base


def display_name_short(col: str) -> str:
    return display_name(col).replace("\n", "")


def base_from_display(disp: str) -> str:
    """Обратное преобразование: просто убираем переносы и маркер (X)."""
    base = disp.replace("\n", "")
    if base.endswith(" (X)"):
        base = base[:-4]
    return base.strip()


# ================== СПРАВКИ ==================
HELP = {
    "settlement_date": "Дата расчёта — от неё считаются сроки до погашения и оферты.",
    "forecast_target_date": "Дата прогноза — дата, к которой ожидается изменение ключевой ставки.",
    "rate_change_pp": "Прогноз ставки ЦБ — целевое значение ключевой ставки к выбранной дате. Изменение в п.п. считается автоматически.",
    "ndfl": "Ставка НДФЛ — налог на купонный доход. Уменьшает реальную доходность.",
    "cbr_rate": "Текущая ставка ЦБ — базовый ориентир для оценки доходности.",
    "uploaded_file": "Excel Доход.ру (X) — файл с метриками качества эмитента, которых нет на MOEX.",
    "industries": "Отрасль (X) — сектор экономики. Для диверсификации.",
    "issuer_types": "Тип эмитента (X): Гос / Субфед / Муницип / Корп.",
    "list_levels": "Уровень листинга — 1 (самые ликвидные) … 3 (риск низкой ликвидности).",
    "only_with_price": "Только с ценой — показывать только бумаги, по которым есть сделки.",
    "only_discount": "Только дисконт — цена ниже номинала (100%).",
    "min_spread": "Мин. спред, % — нижняя граница спреда BID/ASK. 0 = без ограничения.",
    "max_spread": "Макс. спред, % — верхняя граница. ОФЗ <0.1%, корпораты 0.1–0.5%, >1% — низкая ликвидность.",
    "min_coupon": "Купон — выплата в % годовых.",
    "max_coupon": "Макс. купон — верхняя граница. Отсекает мусорные бумаги.",
    "coupon_type_filter": "Тип купона: Фикс / Флоатер / Индексируемый.",
    "face_type_filter": "Тип номинала: Постоянный / Амортизируемый / Индексируемый.",
    "min_days_maturity": "Мин. дней до погашения.",
    "max_days_maturity": "Макс. дней до погашения. 0 = без ограничения.",
    "min_eff_days": "Мин. эффективный срок — срок до ближайшей оферты или погашения.",
    "max_eff_days": "Макс. эффективный срок. 0 = без ограничения.",
    "only_with_maturity": "Только с датой погашения — исключить бессрочные (вечные).",
    "only_with_offer": "Только с офертой — бумаги, где можно вернуть номинал.",
    "min_potential_return": "Потенциальная доходность — сколько заработаете за срок от даты расчёта до даты прогноза.",
    "only_with_ytm": "Только с YTM — бумаги с рассчитанной доходностью к погашению.",
    "min_ytm": "Мин. YTM — нижняя граница доходности.",
    "max_ytm": "Макс. YTM — выше 30% часто означает риск дефолта.",
    "min_net_ytm": "Мин. чистый YTM — доходность после НДФЛ.",
    "max_duration": "Дюрация — на сколько % изменится цена при изменении ставки на 1 п.п.",
    "max_group": "Группа надёжности: 1 — ОФЗ/госбанки … 6 — нет данных.",
    "min_credit_quality": "Кредитное качество (X) — оценка надежности от 0 до 10.",
    "min_issuer_quality": "Качество эмитента (X) — устойчивость бизнес-модели от 1 до 10.",
    "min_inside_q": "Inside Q (X) — внутреннее качество долга (баланс и прибыль).",
    "min_outside_q": "Outside Q (X) — внешнее качество (ROE, NetDebt/Equity, изменчивость прибыли).",
    "min_netdebt": "NetDebt/Equity (X) — долговая нагрузка (рейтинг 1-10). Чем выше, тем ниже долг.",
    "min_liquidity_k": "Коэф. ликвидности (X) — способность погасить краткосрочные долги.",
    "max_complexity": "Сложность (X): 0 — простая, 1 — put-оферта, 2 — амортизируемая, 3 — плавающий купон, 4 — индексируемый, 5+ — call, вечная, структурная.",
    "only_non_subordinated": "Только без субординации (X).",
    "only_with_guarantee": "Только с гарантией (X) — бумаги с поручителем или гарантом.",
    "min_volume": "Мин. объём торгов — низкий объём = не купить без влияния на цену.",
    "min_trades": "Мин. сделок за день — мало сделок = цена может быть случайной.",
    "keep_na": "Показывать бумаги без данных — включить строки с пропусками.",
    "max_lot_cost": "Макс. сумма покупки лота. 0 = без ограничения.",
    "currency_filter": "Валюта номинала — рублёвый портфель: только SUR/RUB.",
    "board_filter": "Режим торгов — TQCB (корпоративные), TQOB (ОФЗ), TQOY (юани).",
    "color_filter_on": "Фильтр по цветам — оставить строки с определённой подсветкой.",
    "color_col_filter_on": "Фильтр по цвету столбца.",
    "show_green": "Зелёные — хорошо по показателю.",
    "show_yellow": "Жёлтые — внимание.",
    "show_red": "Красные — риск.",
    "show_uncolored": "Без подсветки — нейтральные бумаги.",
    "reset_btn": "Сбросить — вернуть все фильтры к параметрам по умолчанию.",
    "use_colors": "Красить ячейки — включить/выключить подсветку.",
    "preset_select": "Пресет — готовый набор фильтров под стратегию.",
    "income_type": "Тип доходности — метрика для сортировки и отображения.",
    "extremum_date": "Экстремум — дата, от которой считается изменение RGBI и цены облигации для расчёта Силы актива.",
}

COLUMN_HELP = {
    "SECID": "Биржевой тикер. По клику — карточка на analytics.dohod.ru (по ISIN).",
    "ISIN": "Международный идентификатор. Сопоставление с Excel Доход.ру.",
    "Наименование": "Короткое название выпуска MOEX.",
    "Полное наименование": "Полное официальное название выпуска.",
    "Режим торгов": "Доска MOEX.",
    "Эмитент": "Эмитент (X).",
    "Основной заемщик": "Основной заемщик (X).",
    "Отрасль": "Отрасль (X).",
    "Тип эмитента (Доход.ру)": "Гос/субфед/муницип/корп (X).",
    "Тип ценной бумаги": "Категория MOEX.",
    "Название (Доход.ру)": "Название в базе Доход.ру (X).",
    "Тип бумаги": "Сводная метка: Фикс, Флоатер, Амортизируемая, Суборд.",
    "Флоатер": "Да — купон пересчитывается, защита от роста ставок.",
    "Высокий риск": "Флаг высокого риска MOEX.",
    "Технический дефолт": "Признак техдефолта.",
    "Цена последней сделки, %": "Цена последней сделки в % от номинала.",
    "Средневзвешенная цена, % к номиналу": "Средняя цена сделок за день.",
    "Цена закрытия, %": "Цена на момент закрытия сессии.",
    "Текущая цена, %": "Цена на момент запроса.",
    "Изменение, % к предыдущей": "Изменение цены к предыдущему дню.",
    "Спред, %": "BID/ASK спред. Транзакционные издержки.",
    "Лучшая цена покупки, %": "BID.",
    "Лучшая цена продажи, %": "OFFER.",
    "Ставка купона, %": "Ставка купона в % годовых.",
    "Купон": "Размер купона за период.",
    "Годовой купон": "Сумма купонов за 12 месяцев.",
    "Период купона, дней": "Дней между выплатами.",
    "НКД": "Накопленный купонный доход.",
    "Тип купона": "Фикс/Переменный/Плавающий/Формула/Индексируемый.",
    "Тип номинала": "Постоянный/Переменный/Амортизируемый/Индексируемый.",
    "Номинал": "Основная сумма долга.",
    "Валюта номинала": "SUR/RUB/USD/CNY и т.д.",
    "Размер лота": "Количество облигаций в лоте.",
    "Объем выпуска, шт.": "Общее количество облигаций.",
    "Грязная цена": "Цена покупки с НКД.",
    "Сумма покупки лота": "Сколько денег нужно на один лот.",
    "Уровень листинга": "1/2/3 — требования биржи.",
    "Тип купона (Доход.ру)": "Классификация типа купона (X).",
    "Дата погашения": "Когда эмитент вернёт номинал.",
    "Дата оферты": "Когда инвестор может продать бумагу эмитенту.",
    "Дата оценки доходности": "BUYBACKDATE — к какой дате считается доходность.",
    "Дней до погашения": "Календарных дней до погашения.",
    "Дней до оферты": "Календарных дней до оферты.",
    "Эффективный срок, дней": "Срок до даты оценки доходности.",
    "Лет до даты (Доход.ру)": "Срок в годах (X).",
    "Ближайшая дата, лет": "Срок в годах до ближайшего события.",
    "Событие": "Оферта или погашение — что раньше.",
    "Потенциальная доходность, %": "Сколько заработаете за срок до даты прогноза при заданном сдвиге ставки.",
    "Прогнозная цена, %": "Цена к дате прогноза при заданном изменении ставки.",
    "YTM, %": "Эффективная доходность к погашению.",
    "Чистый YTM, %": "YTM после НДФЛ.",
    "YTC, %": "Доходность к оферте.",
    "YTC call, %": "Доходность к дате call-опциона.",
    "Текущая доходность, %": "Купон / цена × 100%.",
    "Простая доходность, %": "Без реинвестирования, с учётом дисконта к номиналу.",
    "Доходность по последней сделке": "YTM по цене последней сделки.",
    "Доходность по средневзвешенной цене": "YTM по WAP.",
    "Доходность к оферте": "YTM до даты оферты.",
    "Доходность к кол-опциону по срвзв.": "YTM до даты call.",
    "Дюрация, лет": "На сколько % изменится цена при изменении ставки на 1 п.п.",
    "Дюрация, дней": "Дюрация в днях.",
    "Дюрация к дате колл-опциона": "Дюрация до даты call.",
    "Дисконт к номиналу, %": "100 − цена.",
    "Группа надёжности": "Агрегированный показатель кредитного риска (1–6).",
    "Z-спред по цене последней сделки": "Премия к безрисковой кривой, б.п.",
    "Z-спред по средневзвешенной цене": "Z-спред по WAP.",
    "Рыночный риск, %": "Изменение цены при изменении ставок на ваш прогноз.",
    "Сила актива (Δ%)": "Разница между изменением цены облигации и RGBI с даты Экстремум. Больше нуля — обогнала рынок.",
    "Кредитное качество": "Оценка эмитента от 0 до 10 (X).",
    "Кредитный рэнкинг": "Буквенная оценка (X).",
    "Качество эмитента": "Устойчивость бизнес-модели от 1 до 10 (X).",
    "Inside Q": "Внутреннее качество долга (X).",
    "Outside Q": "Внешнее качество (X).",
    "NetDebt/Equity": "Долговая нагрузка (X).",
    "Коэф. ликвидности": "Способность погасить краткосрочные долги (X).",
    "Сложность": "0 — простая … 7 — структурная (X).",
    "Субординированная": "Пониженный приоритет при банкротстве (X).",
    "С гарантией": "Есть поручитель/гарант (X).",
    "Размер выпуска (оценка)": "Рейтинг размера выпуска от 0 до 10 (X).",
    "Объём в обращении, млрд": "Сколько млрд в обращении (X).",
    "Объем в валюте": "Дневной оборот в валюте номинала.",
    "Объем в бумагах": "Дневной оборот в бумагах.",
    "Количество сделок": "Сделок за день.",
    "Объем на покупку, шт.": "BIDDEPTH.",
    "Объем на продажу, шт.": "OFFERDEPTH.",
    "Медиана оборота, млн": "Медианный дневной оборот (X).",
    "Ликвидность (комбинир.)": "Коэффициент ликвидности + количество сделок.",
    "Вмененная плавающая ставка": "Ожидаемая RUONIA + спред.",
    "Вмененная инфляция (BEI)": "Ожидаемая инфляция.",
    "Вмененная ключевая ставка (CBR)": "Ожидаемая средняя ставка ЦБ.",
    "Состояние сессии": "Статус торгов.",
    "Время обновления": "Время последнего обновления.",
    "Время загрузки": "Время ответа сервера ISS.",
}

DATE_COLS = {"Дата погашения", "Дата оферты", "Дата оценки доходности"}
INT_COLS = {"Количество сделок", "Объем в бумагах", "Размер лота", "Объем выпуска, шт.",
            "Уровень листинга", "Дней до погашения", "Дней до оферты",
            "Период купона, дней", "Эффективный срок, дней", "Группа надёжности"}

COLUMN_GROUPS = {
    "Идентификация": [
        "SECID", "ISIN", "Наименование", "Полное наименование", "Режим торгов",
        "Эмитент", "Основной заемщик", "Отрасль", "Тип эмитента (Доход.ру)",
        "Тип ценной бумаги", "Название (Доход.ру)", "Тип бумаги",
    ],
    "Рынок": [
        "Цена последней сделки, %", "Средневзвешенная цена, % к номиналу",
        "Цена закрытия, %", "Текущая цена, %", "Изменение, % к предыдущей",
        "Спред, %", "Лучшая цена покупки, %", "Лучшая цена продажи, %",
    ],
    "Купон": [
        "Ставка купона, %", "Купон", "Годовой купон", "Период купона, дней",
        "НКД", "Тип купона", "Тип номинала", "Номинал", "Валюта номинала",
        "Размер лота", "Объем выпуска, шт.", "Грязная цена", "Сумма покупки лота",
        "Уровень листинга", "Тип купона (Доход.ру)", "Флоатер",
    ],
    "Сроки": [
        "Дата погашения", "Дата оферты", "Дата оценки доходности",
        "Дней до погашения", "Дней до оферты", "Эффективный срок, дней",
        "Лет до даты (Доход.ру)", "Ближайшая дата, лет", "Событие",
    ],
    "Доходность и риск": [
        "Потенциальная доходность, %", "Прогнозная цена, %", "YTM, %", "Чистый YTM, %",
        "YTC, %", "YTC call, %", "Текущая доходность, %", "Простая доходность, %",
        "Доходность по средневзвешенной цене", "Дюрация, лет", "Дюрация, дней",
        "Дюрация к дате колл-опциона", "Дисконт к номиналу, %", "Группа надёжности",
        "Z-спред по цене последней сделки", "Z-спред по средневзвешенной цене",
        "Рыночный риск, %", "Сила актива (Δ%)",
    ],
    "Кредитное качество (X)": [
        "Кредитное качество", "Кредитный рэнкинг", "Качество эмитента",
        "Inside Q", "Outside Q", "NetDebt/Equity", "Коэф. ликвидности",
        "Сложность", "Субординированная", "С гарантией",
        "Размер выпуска (оценка)", "Объём в обращении, млрд",
    ],
    "Ликвидность": [
        "Объем в валюте", "Объем в бумагах", "Количество сделок",
        "Объем на покупку, шт.", "Объем на продажу, шт.",
        "Медиана оборота, млн", "Ликвидность (комбинир.)",
    ],
    "Риск": ["Высокий риск", "Технический дефолт"],
    "Макро": [
        "Вмененная плавающая ставка", "Вмененная инфляция (BEI)",
        "Вмененная ключевая ставка (CBR)", "Состояние сессии",
        "Время обновления", "Время загрузки",
    ],
}

DEFAULT_VISIBLE_COLS = {
    "SECID", "Наименование", "Тип бумаги", "Флоатер",
    "Цена последней сделки, %", "Изменение, % к предыдущей", "Спред, %",
    "Ставка купона, %", "Период купона, дней", "Купон", "Валюта номинала",
    "Дата погашения", "Дней до погашения", "Событие",
    "Дюрация, лет", "Сила актива (Δ%)",
    "Кредитное качество", "Кредитный рэнкинг", "Качество эмитента",
    "Ликвидность (комбинир.)",
}

# ================== ЦВЕТА ==================
_COLOR_BG = {"green": "#c6efce", "yellow": "#ffeb9c", "red": "#ffc7ce"}
_COLOR_FG = {"green": "#006100", "yellow": "#9c6500", "red": "#9c0006"}

# Колонки, для которых существуют правила подсветки.
# Теги считаем только по ним, а не по всем 120 столбцам.
COLORABLE_COLS = [
    "Потенциальная доходность, %", "YTM, %", "Чистый YTM, %", "YTC, %", "YTC call, %",
    "Текущая доходность, %", "Простая доходность, %", "Доходность по средневзвешенной цене",
    "Ставка купона, %", "Дюрация, лет", "Дней до погашения", "Эффективный срок, дней",
    "Цена последней сделки, %", "Дисконт к номиналу, %", "Спред, %", "Группа надёжности",
    "Кредитное качество", "Качество эмитента", "Inside Q", "Outside Q", "NetDebt/Equity",
    "Коэф. ликвидности", "Сложность", "Субординированная", "С гарантией",
    "Объем в валюте", "Количество сделок",
    "Z-спред по цене последней сделки", "Z-спред по средневзвешенной цене",
    "Уровень листинга", "Тип купона", "Рыночный риск, %", "Тип бумаги", "Флоатер",
    "Сила актива (Δ%)",
]


def _is_rub_series(df: pd.DataFrame) -> pd.Series:
    if "Валюта номинала" not in df.columns:
        return pd.Series(False, index=df.index)
    return df["Валюта номинала"].astype(str).str.strip().str.upper().isin(RUB_CURRENCIES)


def compute_color_tags(df, cbr_rate, max_duration, rgbi_change_pct=None, columns=None):
    cols = [c for c in (columns if columns is not None else COLORABLE_COLS) if c in df.columns]
    tags = pd.DataFrame("", index=df.index, columns=cols)

    def mark(col, mask, color):
        if col in tags.columns and mask is not None:
            m = mask.reindex(tags.index).fillna(False).astype(bool)
            tags.loc[m, col] = color

    if "Потенциальная доходность, %" in df.columns:
        s = df["Потенциальная доходность, %"]
        if cbr_rate > 0:
            mark("Потенциальная доходность, %", s.notna() & (s >= cbr_rate), "green")
        mark("Потенциальная доходность, %", s.notna() & (s < 0), "red")

    for col in ("YTM, %", "Чистый YTM, %", "YTC, %", "YTC call, %",
                "Текущая доходность, %", "Простая доходность, %",
                "Доходность по средневзвешенной цене"):
        if col in df.columns:
            s = df[col]
            if cbr_rate > 0:
                mark(col, s.notna() & (s >= cbr_rate), "green")
            mark(col, s.notna() & (s < 0), "red")

    if "Ставка купона, %" in df.columns:
        s = df["Ставка купона, %"]
        if cbr_rate > 0:
            mark("Ставка купона, %", s.notna() & (s >= cbr_rate), "green")
        mark("Ставка купона, %", s.notna() & (s < 5), "red")

    if "Дюрация, лет" in df.columns:
        s = df["Дюрация, лет"]
        mark("Дюрация, лет", s.notna() & (s <= 1) & (s <= max_duration), "green")
        mark("Дюрация, лет", s.notna() & (s > 1) & (s <= max_duration), "yellow")
        mark("Дюрация, лет", s.notna() & (s > max_duration), "red")

    for col in ("Дней до погашения", "Эффективный срок, дней"):
        if col in df.columns:
            s = df[col]
            mark(col, s.notna() & (s >= 365) & (s <= 365 * 3), "green")
            mark(col, s.notna() & ((s < 0) | (s > 365 * 5)), "red")

    if "Цена последней сделки, %" in df.columns:
        s = df["Цена последней сделки, %"]
        mark("Цена последней сделки, %", s.notna() & (s >= 80) & (s <= 100), "green")
        mark("Цена последней сделки, %", s.notna() & (s > 105), "red")

    if "Дисконт к номиналу, %" in df.columns:
        s = df["Дисконт к номиналу, %"]
        mark("Дисконт к номиналу, %", s.notna() & (s >= 5), "green")
        mark("Дисконт к номиналу, %", s.notna() & (s < 0), "red")

    if "Спред, %" in df.columns:
        s = df["Спред, %"]
        mark("Спред, %", s.notna() & (s < 0.2), "green")
        mark("Спред, %", s.notna() & (s >= 0.5) & (s < 1.0), "yellow")
        mark("Спред, %", s.notna() & (s >= 1.0), "red")

    if "Группа надёжности" in df.columns:
        s = df["Группа надёжности"]
        mark("Группа надёжности", s <= 3, "green")
        mark("Группа надёжности", s >= 5, "red")

    for col, hi, lo in [("Кредитное качество", 8, 5), ("Качество эмитента", 8, 5),
                        ("Inside Q", 7, 4), ("Outside Q", 7, 4), ("NetDebt/Equity", 7, 4)]:
        if col in df.columns:
            s = df[col]
            mark(col, s.notna() & (s >= hi), "green")
            mark(col, s.notna() & (s < lo), "red")

    if "Коэф. ликвидности" in df.columns:
        s = df["Коэф. ликвидности"]
        mark("Коэф. ликвидности", s.notna() & (s >= 70), "green")
        mark("Коэф. ликвидности", s.notna() & (s < 50), "red")

    if "Сложность" in df.columns:
        s = df["Сложность"]
        mark("Сложность", s.notna() & (s <= 1), "green")
        mark("Сложность", s.notna() & (s >= 3), "red")

    if "Субординированная" in df.columns:
        s = df["Субординированная"].astype(str).str.strip().str.lower()
        mark("Субординированная", s.isin(["да", "yes"]), "red")

    if "С гарантией" in df.columns:
        s = df["С гарантией"].astype(str).str.strip().str.lower()
        mark("С гарантией", s.isin(["да", "yes"]), "green")

    is_rub = _is_rub_series(df)
    if "Объем в валюте" in df.columns:
        s = df["Объем в валюте"]
        mark("Объем в валюте", s.notna() & is_rub & (s >= 10_000_000), "green")
        mark("Объем в валюте", s.notna() & is_rub & (s < 500_000), "red")

    if "Количество сделок" in df.columns:
        s = df["Количество сделок"]
        mark("Количество сделок", s.notna() & (s >= 100), "green")
        mark("Количество сделок", s.notna() & (s < 5), "red")

    for col in ("Z-спред по цене последней сделки", "Z-спред по средневзвешенной цене"):
        if col in df.columns:
            s = df[col]
            mark(col, s.notna() & (s >= 300), "green")
            mark(col, s.notna() & (s < 50), "red")

    if "Уровень листинга" in df.columns:
        s = df["Уровень листинга"]
        mark("Уровень листинга", s == 1, "green")
        mark("Уровень листинга", s == 3, "red")

    if "Тип купона" in df.columns:
        s = df["Тип купона"].astype(str).str.lower()
        mark("Тип купона", s.str.contains("плавающ|формул|перемен", na=False), "yellow")

    if "Рыночный риск, %" in df.columns:
        s = df["Рыночный риск, %"]
        mark("Рыночный риск, %", s.notna() & (s >= -2.0), "green")
        mark("Рыночный риск, %", s.notna() & (s < -5.0), "red")

    if "Тип бумаги" in df.columns:
        s = df["Тип бумаги"].astype(str)
        mark("Тип бумаги", s.str.contains("Флоатер", na=False), "yellow")
        mark("Тип бумаги", s.str.contains("Суборд", na=False), "red")

    if "Флоатер" in df.columns:
        s = df["Флоатер"].astype(str)
        mark("Флоатер", s.str.contains("Да", na=False), "yellow")

    if "Сила актива (Δ%)" in df.columns and rgbi_change_pct is not None:
        s = df["Сила актива (Δ%)"]
        mark("Сила актива (Δ%)", s.notna() & (s > 0), "green")
        mark("Сила актива (Δ%)", s.notna() & (s <= 0), "red")

    return tags


# ================== КЕШ EXCEL ==================
CACHE_DIR = Path.home() / ".moex_bond_screener"
try:
    CACHE_DIR.mkdir(exist_ok=True)
except Exception:
    CACHE_DIR = Path(".")
DOHOD_FILE = CACHE_DIR / "dohod_latest.xlsx"
DOHOD_META = CACHE_DIR / "dohod_meta.json"


def _file_hash(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def save_dohod_bytes(data: bytes, source: str = "upload"):
    try:
        DOHOD_FILE.write_bytes(data)
        DOHOD_META.write_text(json.dumps({
            "saved_at": datetime.now().isoformat(),
            "source": source, "size_bytes": len(data), "hash": _file_hash(data),
        }, ensure_ascii=False))
    except Exception:
        pass


def load_cached_dohod_bytes():
    try:
        return DOHOD_FILE.read_bytes() if DOHOD_FILE.exists() else None
    except Exception:
        return None


def get_dohod_cache_age_days():
    if not DOHOD_META.exists():
        return None
    try:
        meta = json.loads(DOHOD_META.read_text())
        return (datetime.now() - datetime.fromisoformat(meta["saved_at"])).days
    except Exception:
        return None


def clear_dohod_cache():
    for p in (DOHOD_FILE, DOHOD_META):
        try:
            if p.exists():
                p.unlink()
        except Exception:
            pass


# ================== ОТРАСЛЬ ==================
INDUSTRY_KEYWORDS = {
    "Банк": ["банк", "сбербанк", "втб", "газпромбанк", "альфа-банк", "россельхозбанк",
             "открытие", "псб", "мкб", "совкомбанк", "росбанк", "тинькофф", "тбанк",
             "мтс-банк", "экспобанк", "центр-инвест", "авто финанс", "росэксимбанк",
             "мсп банк", "вэб"],
    "Нефтегаз": ["газпром", "роснефть", "лукойл", "новатэк", "татнефть", "башнефть",
                 "сургутнефтегаз", "славнефть", "нефтегаз", "транснефть",
                 "газпром нефть", "газпром капитал", "нк "],
    "Металлургия": ["металлоинвест", "норильск", "русал", "северсталь", "ммк", "тмк",
                    "омк", "евраз", "селигдар", "полюс", "алроса", "полиметалл",
                    "уральская сталь", "кокс"],
    "Энергетика": ["россети", "фск", "интер рао", "русгидро", "лэнэнерго", "мособлэнерго",
                   "тгк", "огк", "энерг", "т плюс", "мозэс", "мрск", "мосэнергосбыт"],
    "Телеком": ["мтс", "ростелеком", "мегафон", "билайн", "вымпелком", "телеком",
                "метроком", "эр-телеком"],
    "Лизинг": ["лизинг", "втб лиз", "газтехлизинг", "гтлк", "европлан", "совкомбанк лизинг",
               "балтийский лизинг", "интерлизинг", "экономлизинг", "ресо-лизинг",
               "практика лк", "рольф", "контрол лизинг", "мсб-лизинг", "техно лизинг"],
    "Девелопмент": ["пик", "лср", "самолет", "эталон", "а101", "брусника", "ингр",
                    "донстрой", "пионер", "апри", "аквилон", "легенда", "сэтл"],
    "Ритейл": ["магнит", "х5", "пятерочка", "лента", "о'кей", "детский мир",
               "м.видео", "мвидео", "м видео", "окей"],
    "Химия": ["акрон", "фосагро", "еврохим", "нижнекамскнефтехим", "сибур",
              "куйбышевазот", "полипласт", "азот", "агроэко"],
    "Транспорт": ["ржд", "аэрофлот", "трансконтейнер", "нмтп", "фпк", "совкомфлот",
                  "гтлк", "первая грузовая", "пгк", "новотранс", "евротранс",
                  "сегежа", "делпортс", "флот", "флит"],
    "МФО / коллекторы": ["мфк", "мфо", "лайм-займ", "мани мен", "быстроденьги", "займ",
                         "деньги", "пко", "коллект", "джой мани", "фордевинд",
                         "мани капитал", "веббанкир", "смартфакт", "мгкл", "пкб",
                         "айди коллект", "сза"],
    "Госсектор": ["минфин", "правительство", "администрация", "департамент", "город",
                  "область", "край", "республика", "мосфин", "мэф"],
    "Строительство": ["строй", "стройтрест", "мостоотряд"],
    "АПК": ["агро", "зерно", "русагро", "черкизово", "мироторг", "русское море",
            "интарсия", "агроком", "биотех"],
    "Золото": ["gold", "полюс", "селигдар", "selgold"],
    "Ипотечные агенты": ["иа дом", "иадом", "дом.рф", "ипотечный агент", "сфо", "mortgage"],
}
_INDUSTRY_PATTERNS = {
    ind: "|".join(re.escape(k) for k in kws) for ind, kws in INDUSTRY_KEYWORDS.items()
}


def determine_industry_vectorized(df: pd.DataFrame) -> pd.Series:
    text_parts = []
    for col in ("Эмитент", "Наименование", "Основной заемщик",
                "Название (Доход.ру)", "Полное наименование"):
        if col in df.columns:
            text_parts.append(df[col].fillna("").astype(str))
    if not text_parts:
        return pd.Series("Прочее", index=df.index)
    text = text_parts[0]
    for part in text_parts[1:]:
        text = text + " " + part
    text = text.str.lower()
    result = pd.Series("Прочее", index=df.index, dtype=object)
    for industry, pattern in _INDUSTRY_PATTERNS.items():
        mask = text.str.contains(pattern, regex=True, na=False) & (result == "Прочее")
        result[mask] = industry
    return result


# ================== РЕЖИМЫ ТОРГОВ ==================
BOARD_INFO = {
    "TQCB": {"label": "Корпоративные", "settle": "Т+1", "currency": "RUB", "for": "Обычным"},
    "TQOB": {"label": "ОФЗ", "settle": "Т+1", "currency": "RUB", "for": "Обычным"},
    "TQOY": {"label": "Юани", "settle": "Т+1", "currency": "CNY", "for": "Валютным"},
    "TQOD": {"label": "Еврооблигации", "settle": "Т+1", "currency": "валюта", "for": "Квалам"},
    "TQIF": {"label": "Флоатеры", "settle": "Т+1", "currency": "RUB", "for": "Опытным"},
    "TQTF": {"label": "Сектор ПИР", "settle": "Т+1", "currency": "RUB", "for": "Обычным"},
    "TQBD": {"label": "Доп. режим", "settle": "Т+1", "currency": "RUB", "for": "По ситуации"},
    "TQTC": {"label": "Технический", "settle": "—", "currency": "—", "for": "Служебный"},
    "TQRD": {"label": "Для квалов", "settle": "—", "currency": "RUB", "for": "Квалам"},
    "TQIR": {"label": "Иностранные", "settle": "Т+1", "currency": "валюта", "for": "Квалам"},
    "TQUD": {"label": "С ЦК", "settle": "Т+1", "currency": "RUB", "for": "Обычным"},
}


def board_label(code: str) -> str:
    info = BOARD_INFO.get(code)
    if info:
        return f"{code} — {info['label']} · {info['currency']} · {info['settle']}"
    return code


# ================== HTTP-КЛИЕНТ ISS ==================
_SESSION = None
_SESSION_LOCK = threading.Lock()


def http_session() -> requests.Session:
    """Единая сессия: keep-alive + пул соединений + ретраи. Без st.* (потоки!)."""
    global _SESSION
    if _SESSION is None:
        with _SESSION_LOCK:
            if _SESSION is None:
                s = requests.Session()
                retry = Retry(total=3, backoff_factor=0.4,
                              status_forcelist=(429, 500, 502, 503, 504),
                              allowed_methods=frozenset(["GET"]))
                s.mount("https://", HTTPAdapter(
                    pool_connections=HTTP_WORKERS * 2,
                    pool_maxsize=HTTP_WORKERS * 4,
                    max_retries=retry))
                s.headers.update({"User-Agent": "bond-screener/2.0",
                                  "Accept-Encoding": "gzip, deflate"})
                _SESSION = s
    return _SESSION


def _iss_get(path: str, blocks: tuple, params: dict) -> dict:
    p = {"iss.meta": "off", "iss.only": ",".join(blocks)}
    p.update(params)
    r = http_session().get(f"{BASE_URL}/{path}.json", params=p, timeout=HTTP_TIMEOUT)
    r.raise_for_status()
    j = r.json()
    out = {}
    for b in blocks:
        blk = j.get(b) or {}
        out[b] = pd.DataFrame(blk.get("data") or [], columns=blk.get("columns") or [])
    return out


def fetch_paged(path: str, blocks: tuple, params: dict | None = None,
                workers: int = HTTP_WORKERS, max_pages: int = MAX_PAGES):
    """Параллельная пагинация: страницы качаются пачками, стоп на короткой странице."""
    params = dict(params or {})
    acc = {b: [] for b in blocks}
    main = blocks[0]
    err = None
    with ThreadPoolExecutor(max_workers=workers) as ex:
        page = 0
        while page < max_pages:
            idx = range(page, min(page + workers, max_pages))
            futures = [ex.submit(_iss_get, path, blocks,
                                 {**params, "start": i * PAGE_SIZE, "limit": PAGE_SIZE})
                       for i in idx]
            stop = False
            for f in futures:
                try:
                    res = f.result()
                except (requests.RequestException, ValueError) as e:
                    err, stop = str(e), True
                    break
                head = res[main]
                if head.empty:
                    stop = True
                    break
                for b in blocks:
                    if not res[b].empty:
                        acc[b].append(res[b])
                if len(head) < PAGE_SIZE:
                    stop = True
                    break
            if stop:
                break
            page += workers
    frames = {b: (pd.concat(v, ignore_index=True) if v else pd.DataFrame())
              for b, v in acc.items()}
    return frames, err


# ================== ЗАГРУЗЧИКИ ==================
def _load_board_raw(board: str):
    """securities + marketdata одним запросом на доску (вместо двух загрузок)."""
    frames, err = fetch_paged(
        f"engines/stock/markets/bonds/boards/{board}/securities",
        ("securities", "marketdata"))
    sec, mkt = frames["securities"], frames["marketdata"]
    if sec.empty:
        return pd.DataFrame(), err
    if "BOARDID" not in sec.columns:
        sec["BOARDID"] = board
    if not mkt.empty:
        if "BOARDID" not in mkt.columns:
            mkt["BOARDID"] = board
        keys = ["SECID", "BOARDID"]
        mkt = mkt.drop_duplicates(keys)
        # пересекающиеся колонки берём из marketdata (торговые данные авторитетнее)
        dup = [c for c in sec.columns if c in mkt.columns and c not in keys]
        if dup:
            sec = sec.drop(columns=dup)
        sec = sec.merge(mkt, on=keys, how="left")
    return sec, err


@st.cache_data(ttl=600, show_spinner=False)
def load_universe(boards: tuple):
    logs = []
    with ThreadPoolExecutor(max_workers=BOARD_WORKERS) as ex:
        results = list(ex.map(_load_board_raw, boards))
    frames = []
    for board, (dfb, err) in zip(boards, results):
        logs.append(f"{board}: {len(dfb)} строк" + (f" (ошибка: {err})" if err else ""))
        if not dfb.empty:
            frames.append(dfb)
    if not frames:
        return pd.DataFrame(), logs
    df = pd.concat(frames, ignore_index=True)
    df = df.loc[:, ~df.columns.duplicated()]
    df["_prio"] = df["BOARDID"].map(BOARD_PRIORITY).fillna(99)
    before = len(df)
    df = (df.sort_values("_prio")
            .drop_duplicates(subset=["SECID"], keep="first")
            .drop(columns="_prio")
            .reset_index(drop=True))
    logs.append(f"после дедупликации по SECID: {len(df)} (было {before})")
    return df, logs


def _board_closes_raw(board: str, d_from: date, d_to: date):
    frames, err = fetch_paged(
        f"history/engines/stock/markets/bonds/boards/{board}/securities",
        ("history",),
        params={"from": d_from.isoformat(), "till": d_to.isoformat(),
                "history.columns": "SECID,BOARDID,TRADEDATE,CLOSE"})
    return frames["history"], err


@st.cache_data(ttl=86400, show_spinner=False)
def load_closes(boards: tuple, target: date, lookback: int = HISTORY_LOOKBACK_DAYS) -> pd.Series:
    """SECID -> последнее закрытие на дату target (или ближайшее до неё).
    Один набор запросов на доску вместо запроса на каждую бумагу."""
    d_from = target - timedelta(days=lookback)
    with ThreadPoolExecutor(max_workers=BOARD_WORKERS) as ex:
        results = list(ex.map(lambda b: _board_closes_raw(b, d_from, target), boards))
    frames = [f for f, _ in results if not f.empty]
    if not frames:
        return pd.Series(dtype=float)
    df = pd.concat(frames, ignore_index=True)
    if "CLOSE" not in df.columns or "SECID" not in df.columns:
        return pd.Series(dtype=float)
    df["CLOSE"] = pd.to_numeric(df["CLOSE"], errors="coerce")
    df["TRADEDATE"] = pd.to_datetime(df["TRADEDATE"], errors="coerce")
    df = df.dropna(subset=["CLOSE", "TRADEDATE"])
    if df.empty:
        return pd.Series(dtype=float)
    return df.sort_values("TRADEDATE").groupby("SECID")["CLOSE"].last()


@st.cache_data(ttl=600, show_spinner=False)
def fetch_rgbi_daily(from_date: date, till_date: date) -> pd.DataFrame:
    frames, _ = fetch_paged(
        "history/engines/stock/markets/index/securities/RGBI", ("history",),
        params={"from": from_date.isoformat(), "till": till_date.isoformat(),
                "history.columns": "TRADEDATE,OPEN,HIGH,LOW,CLOSE"})
    df = frames["history"]
    if df.empty:
        return df
    df["TRADEDATE"] = pd.to_datetime(df["TRADEDATE"], errors="coerce")
    for c in ("OPEN", "HIGH", "LOW", "CLOSE"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return (df.dropna(subset=["TRADEDATE", "CLOSE"])
              .drop_duplicates("TRADEDATE")
              .sort_values("TRADEDATE")
              .reset_index(drop=True))


@st.cache_data(ttl=60, show_spinner=False)
def fetch_rgbi_current():
    try:
        blocks = _iss_get("engines/stock/markets/index/boards/SNDX/securities/RGBI",
                          ("marketdata",), {})
    except Exception:
        return None
    df = blocks["marketdata"]
    if df.empty:
        return None
    row = df.iloc[0]
    o = pd.to_numeric(row.get("OPEN"), errors="coerce")
    h = pd.to_numeric(row.get("HIGH"), errors="coerce")
    low = pd.to_numeric(row.get("LOW"), errors="coerce")
    c = pd.to_numeric(row.get("LAST"), errors="coerce")
    if pd.isna(c):
        c = pd.to_numeric(row.get("CURRENTVALUE"), errors="coerce")
    if pd.isna(c):
        return None
    o = c if pd.isna(o) else o
    h = c if pd.isna(h) else h
    low = c if pd.isna(low) else low
    return {"TRADEDATE": pd.Timestamp(TODAY), "OPEN": float(o),
            "HIGH": float(h), "LOW": float(low), "CLOSE": float(c)}


# ================== ПАРСЕР EXCEL ==================
DOHOD_RENAME = {
    "ISIN": "ISIN",
    "Название": "Название (Доход.ру)",
    "Эмитент": "Эмитент",
    "Основной заемщик": "Основной заемщик",
    "Объем выпуска в обращении, млрд": "Объём в обращении, млрд",
    "Лет до даты": "Лет до даты (Доход.ру)",
    "Кредитное качество (число, max=10)": "Кредитное качество",
    "Кредитное качество (рэнкинг)": "Кредитный рэнкинг",
    "Качество эмитента (max=10)": "Качество эмитента",
    "Inside Q (max=10)": "Inside Q",
    "Outside Q (max=10)": "Outside Q",
    "NetDebt/Equity (рейтинг) (max=10)": "NetDebt/Equity",
    "Коэф. Ликвидности (max=100)": "Коэф. ликвидности",
    "Медиана дневного оборота (млн в валюте торгов)": "Медиана оборота, млн",
    "Сложность (max=5)": "Сложность",
    "Размер (max=10)": "Размер выпуска (оценка)",
    "Тип купона": "Тип купона (Доход.ру)",
    "Тип эмитента": "Тип эмитента (Доход.ру)",
    "Субординированная (да/нет)": "Субординированная",
    "С гарантией (да/нет)": "С гарантией",
    "Отрасль": "Отрасль",
}


def parse_dohod_file(uploaded):
    try:
        raw = pd.read_excel(io.BytesIO(uploaded) if isinstance(uploaded, (bytes, bytearray))
                            else uploaded, sheet_name=0)
    except Exception as e:
        return pd.DataFrame(), f"Не удалось прочитать Excel: {e}"
    raw.columns = [str(c).strip() for c in raw.columns]
    if "ISIN" not in raw.columns:
        return pd.DataFrame(), "В файле нет столбца ISIN"
    renamed = raw.rename(columns={k: v for k, v in DOHOD_RENAME.items() if k in raw.columns})
    keep = ["ISIN"] + [v for v in DOHOD_RENAME.values()
                       if v in renamed.columns and v != "ISIN"]
    renamed = renamed[keep].copy()
    renamed["ISIN"] = renamed["ISIN"].astype(str).str.strip().str.upper()
    renamed = renamed.drop_duplicates(subset=["ISIN"], keep="first")
    num_cols = ["Кредитное качество", "Качество эмитента", "Inside Q", "Outside Q",
                "NetDebt/Equity", "Коэф. ликвидности", "Сложность",
                "Размер выпуска (оценка)", "Лет до даты (Доход.ру)",
                "Объём в обращении, млрд", "Медиана оборота, млн"]
    for c in num_cols:
        if c in renamed.columns:
            renamed[c] = pd.to_numeric(renamed[c], errors="coerce")
    if "Отрасль" not in renamed.columns or renamed["Отрасль"].isna().all():
        renamed["Отрасль"] = determine_industry_vectorized(renamed)
    return renamed, None


@st.cache_data(ttl=3600, show_spinner=False)
def parse_dohod_cached(file_bytes: bytes):
    return parse_dohod_file(file_bytes)


# ================== МЕТРИКИ (векторные) ==================
def _join_labels(parts: list) -> np.ndarray:
    out = np.full(len(parts[0]), "", dtype=object)
    for p in parts:
        out = np.where(out == "", p, np.where(p == "", out, out + " · " + p))
    return out


def enrich(df, settlement_date, forecast_target_date, ndfl, rate_change_pp):
    df = df.copy()
    ts = pd.Timestamp(settlement_date)

    for src, tgt in [("Дата погашения", "Дней до погашения"),
                     ("Дата оферты", "Дней до оферты")]:
        if src in df.columns:
            df[tgt] = (df[src] - ts).dt.days

    if "Дата оценки доходности" in df.columns:
        eff = (df["Дата оценки доходности"] - ts).dt.days
        df["Эффективный срок, дней"] = eff.where(eff >= 0)
    elif "Дней до погашения" in df.columns:
        df["Эффективный срок, дней"] = df["Дней до погашения"]
    if "Эффективный срок, дней" in df.columns and "Дней до погашения" in df.columns:
        df["Эффективный срок, дней"] = df["Эффективный срок, дней"].fillna(df["Дней до погашения"])

    df["Группа надёжности"] = 5
    if "Кредитное качество" in df.columns:
        kq = df["Кредитное качество"]
        if kq.notna().any():
            df.loc[kq >= 9, "Группа надёжности"] = 1
            df.loc[(kq >= 7) & (kq < 9), "Группа надёжности"] = 2
            df.loc[(kq >= 5) & (kq < 7), "Группа надёжности"] = 3
            df.loc[(kq >= 3) & (kq < 5), "Группа надёжности"] = 4
            df.loc[kq < 3, "Группа надёжности"] = 5
        mask_no_kq = kq.isna()
    else:
        mask_no_kq = pd.Series(True, index=df.index)
    if "Уровень листинга" in df.columns:
        level = pd.to_numeric(df["Уровень листинга"], errors="coerce")
        df.loc[mask_no_kq & (level == 1), "Группа надёжности"] = 2
        df.loc[mask_no_kq & (level == 2), "Группа надёжности"] = 3
        df.loc[mask_no_kq & (level == 3), "Группа надёжности"] = 4
    if "Режим торгов" in df.columns:
        df.loc[df["Режим торгов"] == "TQOB", "Группа надёжности"] = 1
    df["Группа надёжности"] = df["Группа надёжности"].astype(int)

    df["YTM, %"] = np.nan
    if "Доходность по последней сделке" in df.columns:
        df["YTM, %"] = df["Доходность по последней сделке"]
    if "Доходность по средневзвешенной цене" in df.columns:
        df["YTM, %"] = df["YTM, %"].fillna(df["Доходность по средневзвешенной цене"])
    df["Чистый YTM, %"] = df["YTM, %"] * (1 - ndfl / 100)

    if "Купон" in df.columns and "Период купона, дней" in df.columns:
        periods = 365.0 / df["Период купона, дней"].replace(0, np.nan)
        df["Годовой купон"] = df["Купон"] * periods

    price_col = next((c for c in ("Цена последней сделки, %",
                                  "Средневзвешенная цена, % к номиналу")
                      if c in df.columns), None)

    if price_col and "Ставка купона, %" in df.columns:
        df["Текущая доходность, %"] = np.where(
            df[price_col] > 0, df["Ставка купона, %"] / df[price_col] * 100, np.nan)
    if "Дюрация, дней" in df.columns:
        df["Дюрация, лет"] = df["Дюрация, дней"] / 365
    if "Доходность к оферте" in df.columns:
        df["YTC, %"] = df["Доходность к оферте"]
    if "Доходность к кол-опциону по срвзв." in df.columns:
        df["YTC call, %"] = df["Доходность к кол-опциону по срвзв."]
    if price_col and "Номинал" in df.columns and "НКД" in df.columns:
        df["Грязная цена"] = df[price_col] / 100.0 * df["Номинал"] + df["НКД"]
        if "Размер лота" in df.columns:
            df["Сумма покупки лота"] = df["Грязная цена"] * df["Размер лота"]
    if price_col:
        df["Дисконт к номиналу, %"] = 100.0 - df[price_col]

    if all(c in df.columns for c in ("Годовой купон", "Грязная цена", "Номинал", "Дней до погашения")):
        years = (df["Дней до погашения"] / 365.0).where(lambda s: s > 0.05)
        df["Простая доходность, %"] = (
            (df["Годовой купон"] + (df["Номинал"] - df["Грязная цена"]) / years)
            / df["Грязная цена"] * 100.0)
    else:
        df["Простая доходность, %"] = np.nan

    T = (pd.Timestamp(forecast_target_date) - ts).days / 365.0
    if T <= 0:
        T = 1.0 / 365.0

    if price_col and all(c in df.columns for c in ("Дюрация, лет", "Годовой купон", "Номинал", "НКД")):
        delta_y = rate_change_pp / 100.0
        dur = df["Дюрация, лет"]
        price_change_pct = (-dur * delta_y + 0.5 * dur ** 2 * delta_y ** 2) * 100.0
        df["Прогнозная цена, %"] = df[price_col] * (1 + price_change_pct / 100.0)
        future_dirty = df["Прогнозная цена, %"] / 100.0 * df["Номинал"] + df["НКД"]
        df["Потенциальная доходность, %"] = (
            (future_dirty - df["Грязная цена"] + df["Годовой купон"] * T)
            / df["Грязная цена"] * 100.0)
        if "Тип купона" in df.columns:
            fl = df["Тип купона"].astype(str).str.lower().str.contains(
                "плавающ|перемен|формул", na=False)
            df.loc[fl, "Прогнозная цена, %"] = np.nan
            df.loc[fl, "Потенциальная доходность, %"] = np.nan

    # --- Рыночный риск: вектор вместо apply(axis=1) ---
    if "Дюрация, лет" in df.columns:
        fallback = np.where(_is_rub_series(df), 2.0, 1.0)
        delta = rate_change_pp if abs(rate_change_pp) > 0.001 else fallback
        df["Рыночный риск, %"] = -df["Дюрация, лет"] * delta
    else:
        df["Рыночный риск, %"] = np.nan

    # --- Ликвидность: вектор ---
    k = df["Коэф. ликвидности"] if "Коэф. ликвидности" in df.columns else pd.Series(np.nan, index=df.index)
    n = df["Количество сделок"] if "Количество сделок" in df.columns else pd.Series(np.nan, index=df.index)
    kk, nn = k / 100.0, n.clip(upper=100) / 100.0
    df["Ликвидность (комбинир.)"] = np.where(
        k.notna() & n.notna(), kk * 0.6 + nn * 0.4, np.where(k.notna(), kk, nn))

    # --- Ближайшее событие: вектор ---
    nat = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    o = df["Дата оферты"] if "Дата оферты" in df.columns else nat
    m = df["Дата погашения"] if "Дата погашения" in df.columns else nat
    use_offer = o.notna() & (m.isna() | (o <= m))
    nearest = o.where(use_offer, m)
    df["Ближайшая дата, лет"] = (nearest - ts).dt.days / 365.0
    df["Событие"] = np.where(use_offer, "Оферта", np.where(m.notna(), "Погашение", ""))

    # --- Тип бумаги: вектор ---
    empty = pd.Series("", index=df.index, dtype=object)
    cp = (df["Тип купона"].astype(str).str.lower() if "Тип купона" in df.columns else empty)
    fv = (df["Тип номинала"].astype(str).str.lower() if "Тип номинала" in df.columns else empty)
    sub = (df["Субординированная"].astype(str).str.strip().str.lower()
           if "Субординированная" in df.columns else empty)
    is_float = cp.str.contains("плавающ|перемен|формул", na=False).to_numpy()
    parts = [
        np.where(is_float, "Флоатер", np.where(cp.str.contains("фикс", na=False), "Фикс", "")),
        np.where(fv.str.contains("амортиз", na=False), "Амортизируемая", ""),
        np.where(fv.str.contains("индекс", na=False), "Индексируемая", ""),
        np.where(m.isna(), "Вечная", ""),
        np.where(sub.isin(["да", "yes"]), "Суборд", ""),
    ]
    labels = _join_labels(parts)
    df["Тип бумаги"] = np.where(labels == "", "—", labels)
    df["Флоатер"] = np.where(is_float, "Да", "—")
    return df


@st.cache_data(ttl=600, show_spinner=False, max_entries=8)
def enrich_cached(_df, data_key, settlement_date, forecast_target_date, ndfl, rate_change_pp):
    return enrich(_df, settlement_date, forecast_target_date, ndfl, rate_change_pp)


def add_strength(df: pd.DataFrame, closes: pd.Series, rgbi_change_pct) -> pd.DataFrame:
    df = df.copy()
    if rgbi_change_pct is None or "SECID" not in df.columns or closes.empty:
        df["Сила актива (Δ%)"] = np.nan
        return df
    price = df["Цена последней сделки, %"] if "Цена последней сделки, %" in df.columns \
        else pd.Series(np.nan, index=df.index)
    if "Текущая цена, %" in df.columns:
        price = price.fillna(df["Текущая цена, %"])
    hist = df["SECID"].map(closes)
    chg = (price / hist - 1.0) * 100.0
    df["Сила актива (Δ%)"] = (chg - rgbi_change_pct).where(hist > 0)
    return df


# ================== ФИЛЬТРЫ ==================
_SUBORD_YES = {"да", "yes", "true", "1", "+"}


def keep(f, col, mask, keep_na):
    if col not in f.columns:
        return f
    return f[f[col].isna() | mask] if keep_na else f[mask]


def apply_filters(f, o):
    keep_na = o["keep_na"]
    if o["only_with_price"]:
        price_cols = [c for c in ("Цена последней сделки, %",
                                  "Средневзвешенная цена, % к номиналу") if c in f.columns]
        if price_cols:
            mask_price = pd.Series(False, index=f.index)
            for c in price_cols:
                mask_price |= f[c].notna()
            f = f[mask_price]
    if o["only_with_ytm"] and "YTM, %" in f.columns:
        f = f[f["YTM, %"].notna()]
    if "Потенциальная доходность, %" in f.columns:
        f = keep(f, "Потенциальная доходность, %",
                 f["Потенциальная доходность, %"] >= o["min_potential_return"], keep_na)
    if "YTM, %" in f.columns:
        f = keep(f, "YTM, %", (f["YTM, %"] >= o["min_ytm"]) & (f["YTM, %"] <= o["max_ytm"]), keep_na)
    if "Чистый YTM, %" in f.columns:
        f = keep(f, "Чистый YTM, %", f["Чистый YTM, %"] >= o["min_net_ytm"], keep_na)
    if "Дней до погашения" in f.columns:
        f = keep(f, "Дней до погашения",
                 (f["Дней до погашения"] >= o["min_days_maturity"]) &
                 (f["Дней до погашения"] <= o["max_days_maturity"]), keep_na)
    if "Эффективный срок, дней" in f.columns:
        f = keep(f, "Эффективный срок, дней",
                 (f["Эффективный срок, дней"] >= o["min_eff_days"]) &
                 (f["Эффективный срок, дней"] <= o["max_eff_days"]), keep_na)
    if "Дюрация, лет" in f.columns:
        f = keep(f, "Дюрация, лет", f["Дюрация, лет"] <= o["max_duration"], keep_na)
    if "Группа надёжности" in f.columns:
        f = keep(f, "Группа надёжности", f["Группа надёжности"] <= o["max_group"], keep_na)
    for col, key in [("Кредитное качество", "min_credit_quality"),
                     ("Качество эмитента", "min_issuer_quality"),
                     ("Inside Q", "min_inside_q"), ("Outside Q", "min_outside_q"),
                     ("NetDebt/Equity", "min_netdebt"), ("Коэф. ликвидности", "min_liquidity_k")]:
        if col in f.columns and o[key] > 0:
            f = keep(f, col, f[col] >= o[key], keep_na)
    if "Сложность" in f.columns and o["max_complexity"] < 10:
        f = keep(f, "Сложность", f["Сложность"] <= o["max_complexity"], keep_na)
    if o["only_non_subordinated"] and "Субординированная" in f.columns:
        s = f["Субординированная"]
        mask_sub = s.notna() & s.astype(str).str.strip().str.lower().isin(_SUBORD_YES)
        f = f[~mask_sub | s.isna()] if keep_na else f[~mask_sub]
    if o["only_with_guarantee"] and "С гарантией" in f.columns:
        s = f["С гарантией"]
        mask_g = s.notna() & s.astype(str).str.strip().str.lower().isin(_SUBORD_YES)
        f = f[mask_g | s.isna()] if keep_na else f[mask_g]
    if "Тип эмитента (Доход.ру)" in f.columns and o["issuer_types"]:
        f = f[f["Тип эмитента (Доход.ру)"].isna() |
              f["Тип эмитента (Доход.ру)"].isin(o["issuer_types"])]
    if "Отрасль" in f.columns and o["industries"]:
        f = f[f["Отрасль"].isna() | f["Отрасль"].isin(o["industries"])]

    is_rub = _is_rub_series(f)
    if "Объем в валюте" in f.columns:
        mask_rub = (f["Объем в валюте"] >= o["min_volume"]) | ~is_rub
        f = f[f["Объем в валюте"].isna() | mask_rub] if keep_na else f[mask_rub]
    if "Количество сделок" in f.columns:
        f = keep(f, "Количество сделок", f["Количество сделок"] >= o["min_trades"], keep_na)
    if "Сумма покупки лота" in f.columns:
        mask_lot = (f["Сумма покупки лота"] <= o["max_lot_cost"]) | ~is_rub
        f = f[f["Сумма покупки лота"].isna() | mask_lot] if keep_na else f[mask_lot]
    if "Спред, %" in f.columns:
        if o.get("min_spread", 0.0) > 0.0:
            mask = (f["Спред, %"] >= o["min_spread"]) & (f["Спред, %"] <= o["max_spread"])
        else:
            mask = f["Спред, %"] <= o["max_spread"]
        f = keep(f, "Спред, %", mask, keep_na)
    if "Ставка купона, %" in f.columns:
        f = keep(f, "Ставка купона, %",
                 (f["Ставка купона, %"] >= o["min_coupon"]) &
                 (f["Ставка купона, %"] <= o["max_coupon"]), keep_na)
    if o["only_discount"] and "Дисконт к номиналу, %" in f.columns:
        f = keep(f, "Дисконт к номиналу, %", f["Дисконт к номиналу, %"] >= 0, keep_na)
    if "Валюта номинала" in f.columns and o["currency_filter"]:
        f = f[f["Валюта номинала"].isin(o["currency_filter"])]
    if "Режим торгов" in f.columns and o["board_filter"]:
        f = f[f["Режим торгов"].isin(o["board_filter"])]
    if "Уровень листинга" in f.columns and o["list_levels"]:
        f = f[f["Уровень листинга"].isin(o["list_levels"])]
    if o["only_with_offer"] and "Дата оферты" in f.columns:
        f = f[f["Дата оферты"].notna()]
    if o["only_with_maturity"] and "Дата погашения" in f.columns:
        f = f[f["Дата погашения"].notna()]
    if "Тип купона" in f.columns and o["coupon_type_filter"]:
        f = f[f["Тип купона"].astype(str).isin(o["coupon_type_filter"])]
    if "Тип номинала" in f.columns and o["face_type_filter"]:
        f = f[f["Тип номинала"].astype(str).isin(o["face_type_filter"])]
    if "Тип бумаги" in f.columns and o.get("bond_type_filter"):
        pattern = "|".join(re.escape(x) for x in o["bond_type_filter"])
        f = f[f["Тип бумаги"].astype(str).str.contains(pattern, na=False, regex=True)]
    return f


def apply_color_filter(res_df, tags_df, selected_colors):
    if not selected_colors:
        return res_df.iloc[0:0], tags_df.iloc[0:0]
    match = pd.Series(False, index=tags_df.index)
    for color in ("green", "yellow", "red"):
        if color in selected_colors:
            match |= (tags_df == color).any(axis=1)
    if "uncolored" in selected_colors:
        match |= (tags_df == "").all(axis=1)
    return res_df[match], tags_df[match]


# ================== СОСТОЯНИЕ / ВИДЖЕТЫ ==================
FILTER_DEFAULTS = {
    "flt_preset": "—",
    "flt_settlement_date": DEFAULT_DATE,
    "flt_extremum_date": DEFAULT_EXTREMUM_DATE,
    "flt_cbr_rate": DEFAULT_CURRENT_CBR,
    "flt_forecast_cbr": DEFAULT_FORECAST_CBR,
    "flt_forecast_target_date": DEFAULT_FORECAST_TARGET_DATE,
    "flt_ndfl": DEFAULT_NDFL,
    "flt_only_with_price": True, "flt_only_discount": False,
    "flt_min_spread": 0.0, "flt_max_spread": DEFAULT_MAX_SPREAD,
    "flt_min_coupon": 0.0, "flt_max_coupon": 50.0, "flt_max_lot_cost": 0,
    "flt_min_days_maturity": DEFAULT_MIN_DAYS, "flt_max_days_maturity": 0,
    "flt_min_eff_days": 0, "flt_max_eff_days": 0,
    "flt_only_with_maturity": True, "flt_only_with_offer": False,
    "flt_min_potential_return": DEFAULT_MIN_POTENTIAL_RETURN,
    "flt_max_duration": DEFAULT_MAX_DURATION, "flt_max_group": DEFAULT_MAX_GROUP,
    "flt_only_with_ytm": False, "flt_min_ytm": -50.0, "flt_max_ytm": 100.0,
    "flt_min_net_ytm": -50.0,
    "flt_min_credit_quality": DEFAULT_MIN_CREDIT_QUALITY, "flt_min_issuer_quality": 0.0,
    "flt_min_inside_q": 0.0, "flt_min_outside_q": 0.0, "flt_min_netdebt": 0.0,
    "flt_min_liquidity_k": 0.0, "flt_max_complexity": DEFAULT_MAX_COMPLEXITY,
    "flt_only_non_subordinated": True, "flt_only_with_guarantee": False,
    "flt_min_volume": DEFAULT_MIN_VOLUME, "flt_min_trades": DEFAULT_MIN_TRADES,
    "flt_keep_na": False,
    "flt_color_on": False, "flt_color_green": True, "flt_color_yellow": True,
    "flt_color_red": True, "flt_color_uncolored": True,
    "flt_color_col_on": False, "flt_color_col_name": "Спред, %",
    "flt_color_col_colors": ["green"],
    "view_income_type": "Простая (без реинвест.)",
    "view_use_colors": True, "view_sort_dir": "Убыв.",
    "show_cbr_decisions": True, "scen_growing": False, "scen_falling": False,
}


def init_state(defaults: dict):
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def reset_all_filters():
    for k in list(st.session_state.keys()):
        if k == "flt_dohod_file":
            continue
        if k.startswith(("flt_", "colsel_", "view_")):
            del st.session_state[k]


def reset_button_callback():
    reset_all_filters()
    st.toast("Фильтры сброшены к параметрам по умолчанию")


def clear_cache_callback():
    clear_dohod_cache()
    st.toast("Кеш Доход.ру удалён")


def apply_preset_callback():
    name = st.session_state.get("flt_preset", "—")
    preset = PRESETS.get(name)
    if not preset:
        return
    cur = float(st.session_state.get("flt_cbr_rate", DEFAULT_CURRENT_CBR))
    for k, v in preset.items():
        if k == "flt_rate_change":
            st.session_state["flt_forecast_cbr"] = float(min(max(cur + v, 0.0), 100.0))
        else:
            st.session_state[k] = v
    st.toast(f"Пресет «{name}» применён")


def num_input(label, min_v, max_v, step, key, help_text="", is_excel=False, integer=False):
    """Один виджет вместо связки слайдер+поле+2 кнопки (в 4 раза меньше виджетов)."""
    if is_excel:
        st.markdown(f"<span class='x-label'>{_html.escape(label)}</span>",
                    unsafe_allow_html=True)
    if integer:
        return int(st.number_input(label, min_value=int(min_v), max_value=int(max_v),
                                   step=int(step), key=key, help=help_text, format="%d",
                                   label_visibility="collapsed" if is_excel else "visible"))
    return float(st.number_input(label, min_value=float(min_v), max_value=float(max_v),
                                 step=float(step), key=key, help=help_text, format="%.2f",
                                 label_visibility="collapsed" if is_excel else "visible"))


def checkbox_group(options, key_prefix, default_selected, n_cols=2):
    selected = []
    cols = st.columns(n_cols)
    for i, opt in enumerate(options):
        key = f"{key_prefix}_{opt}"
        if key not in st.session_state:
            st.session_state[key] = opt in default_selected
        if cols[i % n_cols].checkbox(str(opt), key=key):
            selected.append(opt)
    return selected


def ensure_multiselect(key, options, default):
    cur = st.session_state.get(key)
    if cur is None:
        st.session_state[key] = [o for o in default if o in options]
    else:
        st.session_state[key] = [o for o in cur if o in options]


def ensure_select(key, options, default):
    if st.session_state.get(key) not in options:
        st.session_state[key] = default if default in options else options[0]


# ================== ТАБЛИЦА ==================
def _format_frame(view: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame(index=view.index, columns=view.columns, dtype=object)
    for c in view.columns:
        s = view[c]
        if c in DATE_COLS or pd.api.types.is_datetime64_any_dtype(s):
            out[c] = pd.to_datetime(s, errors="coerce").dt.strftime("%d.%m.%Y").fillna("")
        elif c in INT_COLS:
            out[c] = (pd.to_numeric(s, errors="coerce").round()
                      .astype("Int64").astype(str).replace("<NA>", ""))
        elif pd.api.types.is_numeric_dtype(s):
            out[c] = s.map(lambda v: "" if pd.isna(v) else f"{v:,.2f}".replace(",", " "))
        else:
            out[c] = s.fillna("").astype(str).map(_html.escape)
    return out


def render_table(res: pd.DataFrame, tags: pd.DataFrame, links: pd.Series, use_colors=True):
    total = len(res)
    if total == 0:
        st.info("Нет строк для отображения.")
        return
    sizes = [100, 200, 500, 1000]
    if st.session_state.get("view_page_size") not in sizes:
        st.session_state["view_page_size"] = TABLE_PAGE_SIZE
    c1, c2, c3 = st.columns([1, 1, 3])
    page_size = c1.selectbox("Строк на странице", sizes, key="view_page_size")
    pages = max(1, (total + page_size - 1) // page_size)
    cur_page = int(st.session_state.get("view_page", 1))
    st.session_state["view_page"] = min(max(1, cur_page), pages)
    page = c2.number_input("Страница", min_value=1, max_value=pages, step=1, key="view_page")
    start, stop = (page - 1) * page_size, min(page * page_size, total)
    c3.caption(f"Всего строк: {total}. Показаны {start + 1}–{stop}.")

    view = res.iloc[start:stop]
    fmt = _format_frame(view)
    if use_colors:
        tg = tags.reindex(index=view.index, columns=view.columns).fillna("")
    else:
        tg = pd.DataFrame("", index=view.index, columns=view.columns)
    tag_arr = tg.to_numpy()
    link_vals = links.reindex(view.index).fillna("").astype(str).to_numpy()
    cols = list(view.columns)
    secid_idx = cols.index("SECID") if "SECID" in cols else -1

    parts = ['<div class="tbl-wrap"><table class="screener"><thead><tr>']
    parts += [f"<th>{_html.escape(display_name(c))}</th>" for c in cols]
    parts.append("</tr></thead><tbody>")
    for i, row in enumerate(fmt.itertuples(index=False, name=None)):
        cells = []
        for j, val in enumerate(row):
            color = tag_arr[i, j]
            cls = f' class="c-{color}"' if color in _COLOR_BG else ""
            if j == secid_idx and link_vals[i]:
                val = (f'<a href="{DOHOD_BOND_URL}{_html.escape(link_vals[i])}" '
                       f'target="_blank" rel="noopener noreferrer">{val}</a>')
            cells.append(f"<td{cls}>{val}</td>")
        parts.append("<tr>" + "".join(cells) + "</tr>")
    parts.append("</tbody></table></div>")
    st.markdown("".join(parts), unsafe_allow_html=True)


# ================== ГРАФИКИ ==================
def resample_weekly(df: pd.DataFrame, weeks: int = WEEKLY_BARS) -> pd.DataFrame:
    if df.empty:
        return df
    res = (df.set_index("TRADEDATE")
             .resample("W-FRI")
             .agg({"OPEN": "first", "HIGH": "max", "LOW": "min", "CLOSE": "last"})
             .dropna().reset_index())
    return res.tail(weeks).reset_index(drop=True)


def make_rgbi_chart(df: pd.DataFrame, title: str, is_daily: bool = False) -> go.Figure:
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Ohlc(
        x=df["TRADEDATE"], open=df["OPEN"], high=df["HIGH"],
        low=df["LOW"], close=df["CLOSE"], name="RGBI",
        increasing_line_color="black", decreasing_line_color="black",
        line=dict(width=2),
        hovertemplate="%{x|%d.%m.%Y}<br>O: %{open:.2f}<br>H: %{high:.2f}"
                      "<br>L: %{low:.2f}<br>C: %{close:.2f}<extra></extra>"),
        secondary_y=False)
    current = float(df["CLOSE"].iloc[-1])
    fig.add_trace(go.Scatter(
        x=[df["TRADEDATE"].iloc[0], df["TRADEDATE"].iloc[-1]],
        y=[current, current], mode="lines",
        line=dict(color="#1a73e8", width=1.5, dash="dot"),
        showlegend=False, hoverinfo="skip"), secondary_y=True)
    fig.add_annotation(x=1.01, xref="paper", y=current, yref="y2",
                       text=f"{current:.2f}", showarrow=False,
                       font=dict(size=12, color="#1a73e8"),
                       bgcolor="rgba(255,255,255,0.85)",
                       bordercolor="#1a73e8", borderwidth=1, xanchor="left")
    y_min, y_max = float(df["LOW"].min()) * 0.999, float(df["HIGH"].max()) * 1.001
    fig.update_layout(
        title=title, height=360, margin=dict(l=10, r=70, t=60, b=10),
        xaxis=dict(showgrid=False, rangeslider_visible=False,
                   rangebreaks=[dict(bounds=["sat", "mon"])] if is_daily else []),
        yaxis=dict(title="RGBI", side="left", range=[y_min, y_max],
                   showgrid=True, gridcolor="#eeeeee"),
        yaxis2=dict(title="Текущая", overlaying="y", side="right",
                    showgrid=False, range=[y_min, y_max]),
        showlegend=False)
    return fig


def add_cbr_decisions_to_chart(fig, decisions, show_past=True, show_future=True):
    for dt, rate, change in decisions:
        is_future = dt > TODAY
        if (is_future and not show_future) or (not is_future and not show_past):
            continue
        if is_future:
            line_color, dash, width = "#888888", "dot", 1
            text_color, bg, border = "#888888", "rgba(255,255,255,0.85)", "#888888"
        elif change is None:
            line_color, dash, width = "#cccccc", "solid", 1
            text_color, bg, border = "#666666", "#f5f5f5", "#9e9e9e"
        elif change < 0:
            line_color, dash, width = "#2e7d32", "solid", 1
            text_color, bg, border = "#1b5e20", "#e8f5e9", "#2e7d32"
        elif change > 0:
            line_color, dash, width = "#c62828", "solid", 1
            text_color, bg, border = "#b71c1c", "#ffebee", "#c62828"
        else:
            line_color, dash, width = "#9e9e9e", "solid", 1
            text_color, bg, border = "#616161", "#f5f5f5", "#9e9e9e"
        ts = pd.Timestamp(dt)
        fig.add_shape(type="line", x0=ts, x1=ts, y0=0, y1=1, yref="paper",
                      line=dict(color=line_color, width=width, dash=dash), opacity=0.7)
        if rate is not None and change is not None:
            rate_str = f"{rate:.2f}".rstrip("0").rstrip(".").replace(".", ",")
            text = f"{rate_str}%" + (f" ({change:+d} б.п.)" if change else "")
        elif is_future:
            text = "план"
        else:
            text = ""
        if text:
            fig.add_annotation(x=ts, xref="x", y=1.0, yref="paper",
                               yanchor="bottom", yshift=2, text=text, showarrow=False,
                               font=dict(size=8, color=text_color), bgcolor=bg,
                               bordercolor=border, borderwidth=1, borderpad=1,
                               opacity=0.6 if is_future else 0.9)
    return fig


# =====================================================================
#  ИНТЕРФЕЙС
# =====================================================================
init_state(FILTER_DEFAULTS)

st.markdown("""<style>
.tbl-wrap{overflow:auto;max-height:620px;border:1px solid #ddd;border-radius:4px}
.screener{border-collapse:collapse;font-size:.85em;width:100%}
.screener th{padding:4px 8px;border:1px solid #ddd;background:#f5f5f5;position:sticky;
 top:0;z-index:1;white-space:pre-line;text-align:left}
.screener td{padding:3px 8px;border:1px solid #eee;white-space:nowrap}
.screener a{color:#1a73e8;text-decoration:none;font-weight:600}
.screener td.c-green{background:#c6efce;color:#006100}
.screener td.c-yellow{background:#ffeb9c;color:#9c6500}
.screener td.c-red{background:#ffc7ce;color:#9c0006}
.x-label{color:#28a745;font-weight:600}
</style>""", unsafe_allow_html=True)

st.title("Скринер облигаций QCKSLVR TRADING")
st.markdown("Сообщество: [https://vk.ru/qckslvrtrading](https://vk.ru/qckslvrtrading)")
st.caption("Столбцы и фильтры из Excel Доход.ру помечены «(X)». "
           "Клик по SECID открывает карточку на analytics.dohod.ru по ISIN.")

# ================== БОКОВОЕ МЕНЮ ==================
with st.sidebar:
    st.header("Параметры стратегии")
    st.selectbox("Пресет", ["—"] + list(PRESETS.keys()),
                 key="flt_preset", help=HELP["preset_select"])
    st.button("Применить пресет", key="apply_preset_btn",
              on_click=apply_preset_callback, use_container_width=True)
    st.markdown("---")
    settlement_date = st.date_input("Дата расчёта", format="DD.MM.YYYY",
                                    key="flt_settlement_date", help=HELP["settlement_date"])
    extremum_date = st.date_input("Экстремум", format="DD.MM.YYYY",
                                  key="flt_extremum_date", help=HELP["extremum_date"])
    st.markdown("---")
    current_cbr = st.number_input("Текущая ставка ЦБ, %", 0.0, 100.0, step=0.25,
                                  key="flt_cbr_rate", help=HELP["cbr_rate"])
    forecast_cbr = st.number_input("Прогноз ставки ЦБ, %", 0.0, 100.0, step=0.25,
                                   key="flt_forecast_cbr", help=HELP["rate_change_pp"])
    forecast_target_date = st.date_input("Дата прогноза", format="DD.MM.YYYY",
                                         key="flt_forecast_target_date",
                                         help=HELP["forecast_target_date"])
    rate_change_pp = forecast_cbr - current_cbr
    ndfl = st.number_input("Ставка НДФЛ, %", 0.0, 50.0, step=0.5,
                           key="flt_ndfl", help=HELP["ndfl"])
    st.markdown("---")
    st.subheader("Данные Доход.ру")
    uploaded_file = st.file_uploader("Excel Доход.ру (X)", type=["xlsx", "xls"],
                                     key="flt_dohod_file", help=HELP["uploaded_file"])
    age = get_dohod_cache_age_days()
    if DOHOD_FILE.exists() and uploaded_file is None:
        if age is None:
            st.caption("Используется локальный кеш Доход.ру.")
        elif age > DOHOD_FRESH_DAYS:
            st.warning(f"Кеш устарел: {age} дней.")
        elif age == 0:
            st.success("Свежий кеш.")
        else:
            st.info(f"Кеш: {age} дн.")
    if DOHOD_FILE.exists() or DOHOD_META.exists():
        st.button("Удалить кеш", key="clear_cache_btn", on_click=clear_cache_callback,
                  use_container_width=True)

# ================== ГРАФИКИ RGBI ==================
with st.spinner("Загрузка истории RGBI..."):
    rgbi_daily = fetch_rgbi_daily(TODAY - timedelta(days=365 * 5 + 30), TODAY)
    curr_rgbi = fetch_rgbi_current()

current_added = False
if curr_rgbi and not rgbi_daily.empty:
    last_date = rgbi_daily["TRADEDATE"].iloc[-1].date()
    curr_date = curr_rgbi["TRADEDATE"].date()
    if last_date < curr_date:
        rgbi_daily = pd.concat([rgbi_daily, pd.DataFrame([curr_rgbi])], ignore_index=True)
        current_added = True
    elif last_date == curr_date:
        rgbi_daily.loc[rgbi_daily.index[-1], ["OPEN", "HIGH", "LOW", "CLOSE"]] = [
            curr_rgbi["OPEN"], curr_rgbi["HIGH"], curr_rgbi["LOW"], curr_rgbi["CLOSE"]]
        current_added = True

rgbi_change_pct = None
if not rgbi_daily.empty:
    ext_df = rgbi_daily[rgbi_daily["TRADEDATE"] <= pd.Timestamp(extremum_date)]
    if not ext_df.empty:
        rgbi_change_pct = (rgbi_daily.iloc[-1]["CLOSE"] / ext_df.iloc[-1]["CLOSE"] - 1) * 100

if not rgbi_daily.empty:
    rgbi_weekly = resample_weekly(rgbi_daily, WEEKLY_BARS)
    rgbi_daily_short = rgbi_daily.tail(DAILY_BARS).reset_index(drop=True)
    show_cbr = st.checkbox("Показывать решения ЦБ по ключевой ставке",
                           key="show_cbr_decisions",
                           help="Вертикальные линии в даты заседаний Совета директоров Банка России.")
    decisions = load_cbr_decisions()
    col_w, col_d = st.columns(2)
    with col_w:
        fig_w = make_rgbi_chart(rgbi_weekly,
                                f"RGBI — недельный таймфрейм ({len(rgbi_weekly)} баров, "
                                f"~{len(rgbi_weekly)/52:.1f} г.)")
        if show_cbr:
            fig_w = add_cbr_decisions_to_chart(fig_w, decisions)
        st.plotly_chart(fig_w, use_container_width=True)
    with col_d:
        fig_d = make_rgbi_chart(rgbi_daily_short,
                                f"RGBI — дневной таймфрейм ({len(rgbi_daily_short)} баров, "
                                f"~{len(rgbi_daily_short)/260:.1f} г.)", is_daily=True)
        if show_cbr:
            fig_d = add_cbr_decisions_to_chart(fig_d, decisions)
        st.plotly_chart(fig_d, use_container_width=True)
    if current_added:
        st.caption(f"Текущий бар от {TODAY:%d.%m.%Y} включён в графики.")
    else:
        st.caption(f"Текущий бар не получен. Последний доступный: "
                   f"{rgbi_daily['TRADEDATE'].iloc[-1]:%d.%m.%Y}.")
    if rgbi_change_pct is not None:
        st.metric("Изменение RGBI с даты Экстремум", f"{rgbi_change_pct:+.2f}%")
else:
    st.warning("Не удалось загрузить историю RGBI с MOEX ISS.")

# ================== ПРОГНОЗ ==================
st.markdown("---")
st.markdown("#### Ваш прогноз по ключевой ставке")
if rate_change_pp < 0:
    st.success(
        f"Снижение ставки: с {current_cbr:.2f}% до {forecast_cbr:.2f}% "
        f"({rate_change_pp:+.2f} п.п.) к {forecast_target_date:%d.%m.%Y}. "
        "Цены на облигации растут — выигрывают длинные бумаги с фиксированным купоном. "
        "Увеличьте максимальную дюрацию (5–7 лет), группа надёжности 1–3.")
elif rate_change_pp > 0:
    st.error(
        f"Рост ставки: с {current_cbr:.2f}% до {forecast_cbr:.2f}% "
        f"({rate_change_pp:+.2f} п.п.) к {forecast_target_date:%d.%m.%Y}. "
        "Длинные бумаги падают — сократите дюрацию (1–2 года) или перейдите во флоатеры.")
else:
    st.info(f"Ставка без изменений: {current_cbr:.2f}% к {forecast_target_date:%d.%m.%Y}. "
            "Нейтральный сценарий: ориентируйтесь на текущую доходность и качество эмитента.")

st.markdown("---")
st.markdown("#### Сценарии рынка (на основе динамики RGBI)")
sc1, sc2 = st.columns(2)
with sc1:
    market_growing = st.checkbox("Индекс растет", key="scen_growing")
with sc2:
    market_falling = st.checkbox("Индекс падает", key="scen_falling")
if market_growing and market_falling:
    st.warning("Выбраны оба сценария. Оставьте только один.")
if market_growing:
    st.success("""
**Рынок растет (RGBI вверх)**

Ставка снижается, доходности ОФЗ падают, спрос на длинные бумаги растёт.

| Параметр | Рекомендация |
|---|---|
| Дюрация, лет | до 5–7 |
| Тип купона | Фиксированный |
| Кредитное качество (X) | от 7 |
| Группа надёжности | 1–3 |
| Прогноз ставки ЦБ | −1…−2 п.п. |
""")
if market_falling:
    st.error("""
**Рынок падает (RGBI вниз)**

Ставка высокая или растёт, длинные бумаги под давлением.

| Параметр | Рекомендация |
|---|---|
| Дюрация, лет | до 2–2.5 |
| Тип купона | Флоатер или короткий Фикс |
| Кредитное качество (X) | от 8 |
| Группа надёжности | 1–2 |
| Прогноз ставки ЦБ | +1…+2 п.п. |
""")

# ================== ЗАГРУЗКА MOEX ==================
st.markdown("---")
with st.spinner("Загрузка инструментов MOEX..."):
    df, load_log = load_universe(tuple(BOND_BOARDS))
if df.empty:
    st.error("MOEX не вернул инструменты (securities пустой).")
    st.stop()

rename_map = {
    "SHORTNAME": "Наименование", "SECNAME": "Полное наименование",
    "MATDATE": "Дата погашения", "OFFERDATE": "Дата оферты",
    "BUYBACKDATE": "Дата оценки доходности",
    "COUPONPERCENT": "Ставка купона, %", "COUPONVALUE": "Купон",
    "FACEVALUE": "Номинал", "FACEUNIT": "Валюта номинала", "ACCRUEDINT": "НКД",
    "LOTSIZE": "Размер лота", "ISSUESIZE": "Объем выпуска, шт.",
    "LISTLEVEL": "Уровень листинга", "SECTYPE": "Тип ценной бумаги",
    "COUPONPERIOD": "Период купона, дней", "COUPON_DETAILS": "Тип купона",
    "FACEVALUE_TYPE": "Тип номинала", "BOARDID": "Режим торгов",
    "YIELD": "Доходность по последней сделке",
    "YIELDATWAPRICE": "Доходность по средневзвешенной цене",
    "YIELDTOOFFER": "Доходность к оферте", "DURATION": "Дюрация, дней",
    "CALLOPTIONYIELD": "Доходность к кол-опциону по срвзв.",
    "CALLOPTIONDURATION": "Дюрация к дате колл-опциона",
    "ZSPREAD": "Z-спред по цене последней сделки",
    "ZSPREADATWAPRICE": "Z-спред по средневзвешенной цене",
    "WAPRICE": "Средневзвешенная цена, % к номиналу",
    "LAST": "Цена последней сделки, %", "CLOSEPRICE": "Цена закрытия, %",
    "LCURRENTPRICE": "Текущая цена, %", "VALTODAY": "Объем в валюте",
    "VOLTODAY": "Объем в бумагах", "NUMTRADES": "Количество сделок",
    "LASTCHANGEPRCNT": "Изменение, % к предыдущей",
    "IRICPICLOSE": "Вмененная плавающая ставка",
    "BEICLOSE": "Вмененная инфляция (BEI)",
    "CBRCLOSE": "Вмененная ключевая ставка (CBR)",
    "TRADINGSTATUS": "Состояние сессии", "UPDATETIME": "Время обновления",
    "SYSTIME": "Время загрузки", "SPREAD": "Спред, %",
    "BID": "Лучшая цена покупки, %", "OFFER": "Лучшая цена продажи, %",
    "BIDDEPTH": "Объем на покупку, шт.", "OFFERDEPTH": "Объем на продажу, шт.",
    "HIGHRISK": "Высокий риск", "HASTECHNICALDEFAULT": "Технический дефолт",
}
df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
df = df.loc[:, ~df.columns.duplicated()]

if "ISIN" in df.columns:
    df["ISIN"] = df["ISIN"].astype(str).str.strip().str.upper()
    df.loc[df["ISIN"].isin(["", "NAN", "NONE", "NAT", "<NA>"]), "ISIN"] = np.nan
else:
    df["ISIN"] = np.nan

for col in ("Дата погашения", "Дата оферты", "Дата оценки доходности"):
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce")

num_cols = [
    "Ставка купона, %", "Купон", "НКД", "Номинал", "Доходность по последней сделке",
    "Доходность по средневзвешенной цене", "Доходность к оферте", "Дюрация, дней",
    "Доходность к кол-опциону по срвзв.", "Дюрация к дате колл-опциона",
    "Z-спред по цене последней сделки", "Z-спред по средневзвешенной цене",
    "Средневзвешенная цена, % к номиналу", "Цена последней сделки, %",
    "Цена закрытия, %", "Текущая цена, %", "Объем в валюте", "Объем в бумагах",
    "Количество сделок", "Изменение, % к предыдущей", "Уровень листинга",
    "Период купона, дней", "Размер лота", "Объем выпуска, шт.",
    "Вмененная плавающая ставка", "Вмененная инфляция (BEI)",
    "Вмененная ключевая ставка (CBR)", "Спред, %",
    "Лучшая цена покупки, %", "Лучшая цена продажи, %",
    "Объем на покупку, шт.", "Объем на продажу, шт.",
]
for col in num_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# ================== EXCEL ==================
dohod_bytes, dohod_origin, dohod_err = None, None, None
dohod_df, new_cols, matched = pd.DataFrame(), [], 0

if uploaded_file is not None:
    try:
        dohod_bytes = uploaded_file.read()
        new_hash = _file_hash(dohod_bytes)
        if st.session_state.get("dohod_hash") != new_hash:
            save_dohod_bytes(dohod_bytes, source="manual_upload")
            st.session_state["dohod_hash"] = new_hash
        dohod_origin = "manual"
    except Exception as e:
        dohod_err = f"Не удалось прочитать файл: {e}"

if dohod_bytes is None:
    cached = load_cached_dohod_bytes()
    if cached:
        dohod_bytes, dohod_origin = cached, "cache"

if dohod_bytes:
    with st.spinner("Парсинг Excel Доход.ру..."):
        dohod_df, dohod_err = parse_dohod_cached(dohod_bytes)
    if dohod_err:
        st.error(dohod_err)
    elif dohod_df.empty:
        st.warning("Файл Доход.ру пустой.")
    else:
        before_cols = set(df.columns)
        df = df.merge(dohod_df, on="ISIN", how="left", suffixes=("", "_др"))
        df = df.loc[:, ~df.columns.duplicated()]
        new_cols = sorted(set(df.columns) - before_cols)
        matched = int(df["Кредитное качество"].notna().sum()) if "Кредитное качество" in df.columns else 0
        if "Отрасль" in df.columns:
            mask = df["Отрасль"].isna() | (df["Отрасль"].astype(str).str.strip() == "")
            if mask.any():
                df.loc[mask, "Отрасль"] = determine_industry_vectorized(df[mask])

# ================== ENRICH ==================
data_key = f"{len(df)}:{len(df.columns)}:{st.session_state.get('dohod_hash', 'none')}"
with st.spinner("Расчёт метрик..."):
    df = enrich_cached(df, data_key, settlement_date, forecast_target_date, ndfl, rate_change_pp)

for c in ("Режим торгов", "Валюта номинала", "Отрасль", "Тип купона", "Тип номинала",
          "Тип бумаги", "Тип эмитента (Доход.ру)"):
    if c in df.columns:
        df[c] = df[c].astype("category")

# ================== ОПЦИИ ФИЛЬТРОВ ==================
def _options(col, cast=None):
    if col not in df.columns:
        return []
    s = df[col].dropna()
    if cast is not None:
        s = s.astype(cast)
    return sorted(s.unique().tolist())


currency_options = _options("Валюта номинала")
board_codes_present = _options("Режим торгов")
coupon_type_options = _options("Тип купона", str)
face_type_options = _options("Тип номинала", str)
list_levels_options = _options("Уровень листинга", int)
issuer_types_options = _options("Тип эмитента (Доход.ру)", str)
industry_options = _options("Отрасль", str)
bond_type_options = _options("Тип бумаги", str)

ensure_multiselect("flt_currency_filter", currency_options, ["SUR", "RUB"])
ensure_multiselect("flt_board_filter", board_codes_present, ["TQCB", "TQOB"])
ensure_multiselect("flt_color_col_colors", ["green", "yellow", "red", "uncolored"], ["green"])
color_col_options = [c for c in COLORABLE_COLS if c in df.columns]
ensure_select("flt_color_col_name", color_col_options or ["—"], "Спред, %")

COLOR_LABELS = {"green": "Зелёные", "yellow": "Жёлтые",
                "red": "Красные", "uncolored": "Без подсветки"}

# ================== ФИЛЬТРЫ (одна форма = один rerun) ==================
with st.expander("Фильтры под стратегию", expanded=True):
    with st.form("filters", border=False):
        r1a, r1b, r1c = st.columns(3)
        with r1a, st.expander("Рынок", expanded=False):
            only_with_price = st.checkbox("Только с ценой", key="flt_only_with_price",
                                          help=HELP["only_with_price"])
            only_discount = st.checkbox("Только дисконт", key="flt_only_discount",
                                        help=HELP["only_discount"])
            min_spread = num_input("Мин. спред, %", 0.0, 5.0, 0.05, "flt_min_spread",
                                   HELP["min_spread"])
            max_spread = num_input("Макс. спред, %", 0.0, 5.0, 0.05, "flt_max_spread",
                                   HELP["max_spread"])
        with r1b, st.expander("Купон", expanded=False):
            min_coupon = num_input("Мин. ставка купона, %", 0.0, 50.0, 0.5,
                                   "flt_min_coupon", HELP["min_coupon"])
            max_coupon = num_input("Макс. ставка купона, %", 0.0, 50.0, 0.5,
                                   "flt_max_coupon", HELP["max_coupon"])
            max_lot_cost = num_input("Макс. сумма покупки лота (0 = без лимита)",
                                     0, 10_000_000, 1000, "flt_max_lot_cost",
                                     HELP["max_lot_cost"], integer=True)
        with r1c, st.expander("Сроки", expanded=False):
            min_days_maturity = num_input("Мин. дней до погашения", 0, 7300, 30,
                                          "flt_min_days_maturity", HELP["min_days_maturity"],
                                          integer=True)
            max_days_maturity = num_input("Макс. дней до погашения (0 = без лимита)",
                                          0, 10950, 30, "flt_max_days_maturity",
                                          HELP["max_days_maturity"], integer=True)
            min_eff_days = num_input("Мин. эффективный срок, дней", 0, 7300, 30,
                                     "flt_min_eff_days", HELP["min_eff_days"], integer=True)
            max_eff_days = num_input("Макс. эффективный срок, дней (0 = без лимита)",
                                     0, 10950, 30, "flt_max_eff_days",
                                     HELP["max_eff_days"], integer=True)
            only_with_maturity = st.checkbox("Только с датой погашения",
                                             key="flt_only_with_maturity",
                                             help=HELP["only_with_maturity"])
            only_with_offer = st.checkbox("Только с офертой", key="flt_only_with_offer",
                                          help=HELP["only_with_offer"])

        r2a, r2b, r2c = st.columns(3)
        with r2a, st.expander("Доходность и риск", expanded=False):
            min_potential_return = num_input("Мин. потенциальная доходность, %", -100.0,
                                             100.0, 0.5, "flt_min_potential_return",
                                             HELP["min_potential_return"])
            max_duration = num_input("Макс. дюрация, лет", 0.0, 10.0, 0.25,
                                     "flt_max_duration", HELP["max_duration"])
            max_group = num_input("Макс. группа надёжности", 1, 6, 1, "flt_max_group",
                                  HELP["max_group"], integer=True)
            only_with_ytm = st.checkbox("Только с YTM", key="flt_only_with_ytm",
                                        help=HELP["only_with_ytm"])
            min_ytm = num_input("Мин. YTM, %", -50.0, 100.0, 0.5, "flt_min_ytm", HELP["min_ytm"])
            max_ytm = num_input("Макс. YTM, %", -50.0, 200.0, 1.0, "flt_max_ytm", HELP["max_ytm"])
            min_net_ytm = num_input("Мин. чистый YTM, %", -50.0, 100.0, 0.5,
                                    "flt_min_net_ytm", HELP["min_net_ytm"])
        with r2b, st.expander("Кредитное качество (X)", expanded=False):
            min_credit_quality = num_input("Мин. кредитное качество (X, из 10)", 0.0, 10.0,
                                           0.5, "flt_min_credit_quality",
                                           HELP["min_credit_quality"], is_excel=True)
            min_issuer_quality = num_input("Мин. качество эмитента (X, из 10)", 0.0, 10.0,
                                           0.5, "flt_min_issuer_quality",
                                           HELP["min_issuer_quality"], is_excel=True)
            min_inside_q = num_input("Мин. Inside Q (X, из 10)", 0.0, 10.0, 0.5,
                                     "flt_min_inside_q", HELP["min_inside_q"], is_excel=True)
            min_outside_q = num_input("Мин. Outside Q (X, из 10)", 0.0, 10.0, 0.5,
                                      "flt_min_outside_q", HELP["min_outside_q"], is_excel=True)
            min_netdebt = num_input("Мин. NetDebt/Equity (X, из 10)", 0.0, 10.0, 0.5,
                                    "flt_min_netdebt", HELP["min_netdebt"], is_excel=True)
            min_liquidity_k = num_input("Мин. коэф. ликвидности (X, из 100)", 0.0, 100.0,
                                        1.0, "flt_min_liquidity_k",
                                        HELP["min_liquidity_k"], is_excel=True)
            max_complexity = num_input("Макс. сложность (X, из 7)", 0.0, 7.0, 1.0,
                                       "flt_max_complexity", HELP["max_complexity"],
                                       is_excel=True)
            only_non_subordinated = st.checkbox("Только без субординации (X)",
                                                key="flt_only_non_subordinated",
                                                help=HELP["only_non_subordinated"])
            only_with_guarantee = st.checkbox("Только с гарантией (X)",
                                              key="flt_only_with_guarantee",
                                              help=HELP["only_with_guarantee"])
        with r2c, st.expander("Ликвидность", expanded=False):
            min_volume = num_input("Мин. объём торгов", 0, 500_000_000, 100_000,
                                   "flt_min_volume", HELP["min_volume"], integer=True)
            min_trades = num_input("Мин. сделок за день", 0, 5000, 10, "flt_min_trades",
                                   HELP["min_trades"], integer=True)

        r3a, r3b, r3c = st.columns(3)
        with r3a, st.expander("Валюта и режимы", expanded=False):
            currency_filter = st.multiselect("Валюта номинала", currency_options,
                                             key="flt_currency_filter", help=HELP["currency_filter"])
            board_filter = st.multiselect("Режим торгов", board_codes_present,
                                          format_func=board_label, key="flt_board_filter",
                                          help=HELP["board_filter"])
        with r3b, st.expander("Отрасль (X)", expanded=False):
            st.caption(HELP["industries"])
            industries = checkbox_group(industry_options, "flt_industry", [], n_cols=2) \
                if industry_options else []
        with r3c, st.expander("Тип эмитента (X)", expanded=False):
            st.caption(HELP["issuer_types"])
            issuer_types = checkbox_group(issuer_types_options, "flt_issuer_type", [], n_cols=1) \
                if issuer_types_options else []

        r4a, r4b, r4c = st.columns(3)
        with r4a, st.expander("Уровень листинга", expanded=False):
            st.caption(HELP["list_levels"])
            list_levels = checkbox_group(list_levels_options, "flt_list_levels",
                                         list_levels_options, n_cols=3) if list_levels_options else []
        with r4b, st.expander("Тип купона", expanded=False):
            coupon_type_filter = checkbox_group(coupon_type_options, "flt_coupon_type", [], 2) \
                if coupon_type_options else []
        with r4c, st.expander("Тип номинала", expanded=False):
            face_type_filter = checkbox_group(face_type_options, "flt_face_type", [], 2) \
                if face_type_options else []

        r5a, r5b, r5c = st.columns(3)
        with r5a, st.expander("Тип бумаги", expanded=False):
            bond_type_filter = checkbox_group(bond_type_options, "flt_bond_type", [], 1) \
                if bond_type_options else []
        with r5b, st.expander("Прочее", expanded=False):
            keep_na = st.checkbox("Показывать бумаги без данных", key="flt_keep_na",
                                  help=HELP["keep_na"])
        with r5c, st.expander("Фильтр по цветам (общий)", expanded=False):
            color_filter_on = st.checkbox("Применить фильтр по цвету", key="flt_color_on",
                                          help=HELP["color_filter_on"])
            show_green = st.checkbox("Зелёные", key="flt_color_green", help=HELP["show_green"])
            show_yellow = st.checkbox("Жёлтые", key="flt_color_yellow", help=HELP["show_yellow"])
            show_red = st.checkbox("Красные", key="flt_color_red", help=HELP["show_red"])
            show_uncolored = st.checkbox("Без подсветки", key="flt_color_uncolored",
                                         help=HELP["show_uncolored"])

        with st.expander("Фильтр по цвету столбца", expanded=False):
            color_col_filter_on = st.checkbox("Применить фильтр", key="flt_color_col_on",
                                              help=HELP["color_col_filter_on"])
            color_col_name = st.selectbox("Столбец", color_col_options or ["—"],
                                          key="flt_color_col_name")
            color_col_colors = st.multiselect("Оставить строки с цветом",
                                              ["green", "yellow", "red", "uncolored"],
                                              format_func=lambda x: COLOR_LABELS.get(x, x),
                                              key="flt_color_col_colors")

        st.form_submit_button("Применить фильтры", use_container_width=True)

    st.button("Сбросить к параметрам по умолчанию", key="reset_btn",
              on_click=reset_button_callback, help=HELP["reset_btn"],
              use_container_width=True)

# ================== ПРИМЕНЕНИЕ ФИЛЬТРОВ ==================
if min_spread > max_spread:
    min_spread, max_spread = max_spread, min_spread

opts = {
    "min_potential_return": min_potential_return,
    "min_credit_quality": min_credit_quality, "min_issuer_quality": min_issuer_quality,
    "min_inside_q": min_inside_q, "min_outside_q": min_outside_q,
    "min_netdebt": min_netdebt, "min_liquidity_k": min_liquidity_k,
    "max_complexity": max_complexity,
    "only_non_subordinated": only_non_subordinated, "only_with_guarantee": only_with_guarantee,
    "issuer_types": issuer_types, "industries": industries,
    "min_ytm": min_ytm, "max_ytm": max_ytm, "min_net_ytm": min_net_ytm,
    "only_with_ytm": only_with_ytm, "only_with_price": only_with_price,
    "only_discount": only_discount,
    "min_days_maturity": min_days_maturity,
    "max_days_maturity": max_days_maturity if max_days_maturity > 0 else 10 ** 9,
    "min_eff_days": min_eff_days,
    "max_eff_days": max_eff_days if max_eff_days > 0 else 10 ** 9,
    "max_duration": max_duration, "max_group": max_group,
    "only_with_maturity": only_with_maturity, "only_with_offer": only_with_offer,
    "min_volume": min_volume, "min_trades": min_trades,
    "min_spread": min_spread, "max_spread": max_spread,
    "max_lot_cost": max_lot_cost if max_lot_cost > 0 else 10 ** 12,
    "min_coupon": min_coupon, "max_coupon": max_coupon,
    "list_levels": list_levels,
    "currency_filter": currency_filter or currency_options,
    "board_filter": board_filter or board_codes_present,
    "coupon_type_filter": coupon_type_filter, "face_type_filter": face_type_filter,
    "bond_type_filter": bond_type_filter, "keep_na": keep_na,
}
filtered = apply_filters(df, opts)

# ================== СИЛА АКТИВА (bulk, без запроса на каждую бумагу) ==================
if not filtered.empty and rgbi_change_pct is not None:
    boards_for_hist = tuple(sorted(filtered["Режим торгов"].dropna().astype(str).unique()))
    with st.spinner("Загрузка исторических цен по доскам..."):
        closes = load_closes(boards_for_hist, extremum_date)
    filtered = add_strength(filtered, closes, rgbi_change_pct)
else:
    filtered = filtered.copy()
    filtered["Сила актива (Δ%)"] = np.nan

# ================== ТИП ДОХОДНОСТИ / СОРТИРОВКА ==================
st.subheader(f"Найдено: {len(filtered)} из {len(df)}")

income_type = st.radio("Какую доходность показывать и по ней сортировать:",
                       list(INCOME_TYPE_MAP.keys()), horizontal=True,
                       key="view_income_type", help=HELP["income_type"])
income_col = INCOME_TYPE_MAP[income_type]

numeric_or_date_cols = [c for c in filtered.columns
                        if pd.api.types.is_numeric_dtype(filtered[c])
                        or pd.api.types.is_datetime64_any_dtype(filtered[c])]
if not numeric_or_date_cols:
    numeric_or_date_cols = list(filtered.columns)

sort_display_map = {display_name_short(c): c for c in numeric_or_date_cols}
sort_display_list = list(sort_display_map.keys())
default_sort = income_col if income_col in numeric_or_date_cols else numeric_or_date_cols[0]
ensure_select("view_sort_by", sort_display_list, display_name_short(default_sort))

s1, s2 = st.columns([3, 1])
with s1:
    sort_by_display = st.selectbox("Сортировка — столбец", sort_display_list, key="view_sort_by")
with s2:
    sort_dir = st.radio("Направление", ["Убыв.", "Возр."], key="view_sort_dir")
sort_by = sort_display_map[sort_by_display]
filtered = filtered.sort_values(sort_by, ascending=(sort_dir == "Возр."), na_position="last")

tags_full = compute_color_tags(filtered, current_cbr, max_duration, rgbi_change_pct)

if color_filter_on:
    selected_colors = {c for c, flag in
                       (("green", show_green), ("yellow", show_yellow),
                        ("red", show_red), ("uncolored", show_uncolored)) if flag}
    filtered, tags_full = apply_color_filter(filtered, tags_full, selected_colors)
    if not selected_colors:
        st.warning("Все цветовые галочки сняты — таблица и экспорт пусты.")

if color_col_filter_on and color_col_name in tags_full.columns and color_col_colors:
    col_tags = tags_full[color_col_name].fillna("").replace({"": "uncolored"})
    mask = col_tags.isin(color_col_colors)
    filtered, tags_full = filtered[mask], tags_full[mask]

# ================== ВЫБОР СТОЛБЦОВ ==================
with st.expander("Столбцы для отображения", expanded=False):
    st.caption("Столбцы с (X) — из файла Доход.ру.")
    selected_cols = []
    for group_name, group_cols in COLUMN_GROUPS.items():
        available = [c for c in group_cols if c in filtered.columns]
        if not available:
            continue
        if "(X)" in group_name:
            st.markdown(f"<span class='x-label'>{group_name}</span>", unsafe_allow_html=True)
        else:
            st.markdown(f"**{group_name}**")
        layout = st.columns(3)
        for i, c in enumerate(available):
            key = f"colsel_{c}"
            if key not in st.session_state:
                st.session_state[key] = c in DEFAULT_VISIBLE_COLS
            if layout[i % 3].checkbox(mark_x(c), key=key, help=COLUMN_HELP.get(c, "")):
                selected_cols.append(c)
    use_colors = st.checkbox("Красить ячейки", key="view_use_colors", help=HELP["use_colors"])

if income_col in filtered.columns and income_col not in selected_cols:
    selected_cols.insert(0, income_col)

# ================== ТАБЛИЦА ==================
if not selected_cols:
    st.info("Выберите хотя бы один столбец. В CSV всё равно попадут все доступные столбцы.")
else:
    res = filtered[selected_cols].copy()
    res = res.loc[:, ~res.columns.duplicated()]
    links = (filtered["ISIN"] if "ISIN" in filtered.columns
             else pd.Series("", index=filtered.index))
    render_table(res, tags_full, links, use_colors=use_colors)

# ================== ЭКСПОРТ ==================
export_df_full = filtered.rename(columns={c: mark_x(c) for c in filtered.columns})
csv = export_df_full.to_csv(index=False).encode("utf-8-sig")
st.download_button("Скачать CSV (полный набор для дашборда)", csv,
                   "moex_bonds_dashboard.csv", "text/csv",
                   help="Все столбцы MOEX, Доход.ру и расчётные метрики. Доход.ру помечены (X).")
with st.expander("Все столбцы, попадающие в CSV", expanded=False):
    st.write(f"Всего столбцов: {len(export_df_full.columns)}. Строк: {len(export_df_full)}.")
    st.code("\n".join(export_df_full.columns.tolist()))

# ================== СВОДКА ==================
if not filtered.empty:
    st.markdown("### Сводка")
    c1, c2, c3, c4 = st.columns(4)
    if income_col in filtered.columns:
        s = filtered[income_col].dropna()
        if not s.empty:
            c1.metric(f"Средняя {income_type}", f"{s.mean():.2f}%")
            c2.metric(f"Медиана {income_type}", f"{s.median():.2f}%")
    if "Дюрация, лет" in filtered.columns:
        d = filtered["Дюрация, лет"].dropna()
        if not d.empty:
            c3.metric("Средняя дюрация", f"{d.mean():.2f} лет")
            c4.metric("Медианная дюрация", f"{d.median():.2f} лет")
    cl, cr = st.columns(2)
    if "Группа надёжности" in filtered.columns:
        with cl:
            st.caption("Группы надёжности")
            st.dataframe(filtered["Группа надёжности"].value_counts().sort_index()
                         .rename("Количество"), use_container_width=True)
    if "Отрасль" in filtered.columns:
        with cr:
            st.caption("Отрасли (X)")
            st.dataframe(filtered["Отрасль"].value_counts().head(15)
                         .rename("Количество"), use_container_width=True)

# ================== СЛУЖЕБНЫЙ БЛОК ==================
with st.expander("Справочная информация и служебные блоки", expanded=False):
    st.markdown("### Методология Доход.ру")
    st.markdown("""
- YTM — всегда эффективная, из денежных потоков.
- Флоатеры — моделирование по форвардной кривой, дюрация до следующего пересчёта.
- Оферты — только настоящие, частичные игнорируются.
- Качество эмитента (1–10): от 5 — высокое.
- Сложность: 0 — фикс; 1 — put; 2 — амортизация; 3 — флоатер; 4 — индексируемый;
  5 — call; 6 — вечная; 7 — структурная.
""")
    st.markdown("### Легенда цветов")
    st.markdown("""
**(X)** — столбец или фильтр из Excel Доход.ру.

| Столбец | Зелёный | Жёлтый | Красный |
|---|---|---|---|
| Потенциальная доходность | от ЦБ | — | < 0 |
| YTM, Чистый YTM, YTC, Текущая, Простая | от ЦБ | — | < 0 |
| Ставка купона | от ЦБ | — | < 5% |
| Дюрация, лет | до 1 | 1…макс | > макс |
| Дисконт к номиналу | от 5% | — | < 0 |
| Спред | < 0.2 | 0.5–1.0 | > 1.0 |
| Кредитное качество (X) | от 8 | — | < 5 |
| Качество эмитента (X) | от 8 | — | < 5 |
| Inside/Outside Q (X) | от 7 | — | < 4 |
| NetDebt/Equity (X) | от 7 | — | < 4 |
| Коэф. ликвидности (X) | от 70 | — | < 50 |
| Сложность (X) | до 1 | — | от 3 |
| Субординированная (X) = Да | — | — | красный |
| С гарантией (X) = Да | зелёный | — | — |
| Группа надёжности | 1–3 | — | 5–6 |
| Z-спред, б.п. | от 300 | — | < 50 |
| Уровень листинга | 1 | — | 3 |
| Рыночный риск | от −2% | — | < −5% |
| Флоатер = Да | — | жёлтый | — |
| Тип бумаги | — | Флоатер | Суборд |
| Сила актива (Δ%) | > 0 | — | <= 0 |
""")
    st.markdown("### Детали загрузки данных")
    for line in load_log:
        st.write("— " + line)
    if dohod_bytes and not dohod_err and not dohod_df.empty:
        origin_label = {"manual": "загруженный файл", "cache": "локальный кеш"}.get(dohod_origin, "")
        st.success(f"Доход.ру ({origin_label}): {len(dohod_df)} строк, "
                   f"сопоставлено: {matched} из {len(df)}")
        if new_cols:
            st.caption("Новые столбцы (X): " + ", ".join(mark_x(c) for c in new_cols))
    elif not dohod_bytes:
        st.info("Excel Доход.ру не загружен.")
    st.markdown("### Режимы торгов")
    rows = [{"Код": c, "Тип": BOARD_INFO.get(c, {}).get("label", "—"),
             "Валюта": BOARD_INFO.get(c, {}).get("currency", "—"),
             "Расчёты": BOARD_INFO.get(c, {}).get("settle", "—"),
             "Кому": BOARD_INFO.get(c, {}).get("for", "—")} for c in board_codes_present]
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
