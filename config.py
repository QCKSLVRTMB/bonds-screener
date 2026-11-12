from datetime import date, timedelta

BASE_URL = "https://iss.moex.com/iss"
DOHOD_BOND_URL = "https://analytics.dohod.ru/bond/"

# только нужные доски: меньше строк = меньше страниц
BOND_BOARDS = ("TQOB", "TQCB", "TQIF", "TQOD", "TQOY", "TQIR", "TQUD", "TQRD")
BOARD_PRIORITY = {b: i for i, b in enumerate(BOND_BOARDS)}
RUB_CURRENCIES = {"SUR", "RUB", "RUR"}

PAGE_SIZE = 100          # ISS отдаёт максимум 100 строк
MAX_PAGES = 120
HTTP_WORKERS = 8         # страниц параллельно внутри одной доски
BOARD_WORKERS = 2        # досок параллельно (итого <= 16 соединений)
TIMEOUT = (5, 20)

TTL_MARKET = 300
TTL_HISTORY = 86_400
HISTORY_LOOKBACK_DAYS = 5   # окно, если в дату экстремума бумага не торговалась

TODAY = date.today()
DEFAULT_SETTLEMENT = TODAY                       # было date(2026,10,21) — в прошлом
DEFAULT_FORECAST_TARGET = date(TODAY.year + 1, 1, 1)
DEFAULT_EXTREMUM = TODAY - timedelta(days=120)

# Обрезка колонок: трафик меньше в 2-3 раза.
# Если доска не знает колонку — клиент автоматически повторит запрос без фильтра.
ISS_COLUMNS = {
    "securities": (
        "SECID,BOARDID,SHORTNAME,SECNAME,ISIN,MATDATE,OFFERDATE,BUYBACKDATE,"
        "COUPONPERCENT,COUPONVALUE,COUPONPERIOD,FACEVALUE,FACEUNIT,ACCRUEDINT,"
        "LOTSIZE,ISSUESIZE,LISTLEVEL,SECTYPE,FACEUNIT,STATUS"
    ),
    "marketdata": (
        "SECID,BOARDID,BID,OFFER,SPREAD,LAST,WAPRICE,CLOSEPRICE,LCURRENTPRICE,"
        "LASTCHANGEPRCNT,YIELD,YIELDATWAPRICE,YIELDTOOFFER,DURATION,ZSPREAD,"
        "ZSPREADATWAPRICE,CALLOPTIONYIELD,CALLOPTIONDURATION,VALTODAY,VOLTODAY,"
        "NUMTRADES,BIDDEPTH,OFFERDEPTH,UPDATETIME,SYSTIME,TRADINGSTATUS,"
        "IRICPICLOSE,BEICLOSE,CBRCLOSE"
    ),
    "history": "SECID,BOARDID,TRADEDATE,CLOSE",
}

TABLE_PAGE_SIZE = 200
