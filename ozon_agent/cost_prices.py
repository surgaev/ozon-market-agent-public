"""Закупочные цены из cost_prices.csv в папке данных (~/.ozon-agent).

Формат — две колонки, разделитель запятая или точка с запятой (как сохраняет Excel):
    sku,cost
    ABC123,100
    ABC123-red,120.5
Строки с # — комментарии. Цена ищется по точному артикулу, иначе по самому
длинному совпавшему префиксу: ABC123 покрывает ABC123-blue, ABC123-xl и т.д.
Артикул без цены отчёт подсветит отдельно — прибыль по нему будет завышена.
"""
import csv
import io
import os

from .config import DATA_DIR

COST_FILE = os.path.join(DATA_DIR, "cost_prices.csv")
TAX_RATE = float(os.environ.get("OZON_TAX_RATE", "0.07"))   # налог с выручки (УСН 6% + взносы ≈ 7%)


def _load() -> dict[str, float]:
    if not os.path.exists(COST_FILE):
        return {}
    lines = [l for l in open(COST_FILE, encoding="utf-8-sig") if l.strip() and not l.lstrip().startswith("#")]
    if not lines:
        return {}
    delim = ";" if lines[0].count(";") > lines[0].count(",") else ","
    out = {}
    for row in csv.reader(io.StringIO("".join(lines)), delimiter=delim):
        if len(row) < 2 or row[0].strip().lower() in ("sku", "артикул"):
            continue
        try:
            out[row[0].strip()] = float(row[1].strip().replace(" ", "").replace(",", "."))
        except ValueError:
            continue
    return out


_COST_MAP = _load()


def get_cost_price(sku: str) -> float | None:
    if sku in _COST_MAP:
        return _COST_MAP[sku]
    best, best_len = None, 0
    for prefix, cost in _COST_MAP.items():
        if sku.startswith(prefix) and len(prefix) > best_len:
            best, best_len = cost, len(prefix)
    return best
