import os

# --- Ozon Seller API credentials ---
CLIENT_ID = os.environ.get("OZON_CLIENT_ID", "")
API_KEY   = os.environ.get("OZON_API_KEY", "")

# --- Report output ---
# Настройки (.env), себестоимости (cost_prices.csv) и кэши — в отдельной папке,
# а не рядом с кодом: пакет ставится через pip/uv и его папка может быть только для чтения.
DATA_DIR = os.environ.get("OZON_AGENT_HOME", os.path.expanduser("~/.ozon-agent"))
os.makedirs(DATA_DIR, exist_ok=True)
REPORT_FOLDER = os.environ.get("OZON_REPORT_FOLDER", os.path.expanduser("~/Downloads"))

# --- Alert thresholds ---
# Остаток считается тревожным, когда хватит меньше чем на N дней
RESTOCK_URGENT_DAYS = 3    # 🔴 срочно
RESTOCK_SOON_DAYS   = 7    # 🟡 скоро
RESTOCK_PLAN_DAYS   = 14   # 🔵 планировать

# --- Артикулы, которые не показываем в отчёте (сняты с продажи) ---
# Точное совпадение offer_id; фильтр только на выходе, вернуть = убрать отсюда.
EXCLUDED_SKUS: set[str] = set()   # например {"SKU-1", "SKU-2"}

# --- API base ---
OZON_API_BASE = "https://api-seller.ozon.ru"

# --- Веб-выгрузки из кабинета (ручной режим) ---
# Папка, куда Chrome качает выгрузки.
DOWNLOADS_DIR = os.environ.get("OZON_DOWNLOADS_DIR", os.path.expanduser("~/Downloads"))

# «Все товары и видимость на Ozon» (Аналитика → Отчёты) → лист «Здоровье карточек».
OZON_HEALTH_CSV = os.environ.get("OZON_HEALTH_CSV", os.path.join(DOWNLOADS_DIR, "products.csv"))

# Самые свежие файлы по маске (имена кабинет даёт с датой) — автопоиск в DOWNLOADS_DIR:
#  «Начисления по товарам» (Финансы → Экономика магазина → Скачать → По товарам)
# «Отчет по начислениям…» — новый формат Ozon (~08.2026), «Отчет по товарам…» — старый.
# Парсер определяет формат сам по заголовку, маски нужны только чтобы найти файл.
OZON_ACCRUALS_PATTERNS = ["Отчет по начислениям*.xlsx", "Отчет по товарам за период*.xlsx",
                          "accruals*by*product*.xlsx"]
#  «Возвраты FBS и FBO» (Аналитика → Отчёты → Возвраты → FBS и FBO)
OZON_RETURNS_PATTERNS = ["returns_report_*.xlsx", "Возвраты*.xlsx"]

# --- Налог (та же ставка, что в Yandex-агенте) ---
TAX_RATE = 0.07
