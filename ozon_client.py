"""
Ozon Seller API client.

Auth: Client-Id + Api-Key headers.

Что доступно (и используется здесь):
  POST /v3/product/list                — список product_id
  POST /v3/product/info/list           — offer_id + все числовые sku (sources[])
  POST /v1/analytics/data              — продажи по дням (ordered_units, revenue)
                                          ⚠ остальные метрики (показы/CTR/буст) Ozon отключил
  POST /v4/product/info/stocks         — остатки FBO/FBS
  POST /v3/finance/transaction/list    — юнит-экономика (комиссия, логистика, услуги)
  POST /v3/posting/fbs|fbo/list        — отправления (кластеры доставки, возвраты)
  POST /v5/product/info/prices         — цены

Числовой sku (в analytics и finance) ≠ product_id. Маппинг sku → offer_id
строится через product/info/list (поля sku + sources[].sku).
"""

import time
import requests
from datetime import date, timedelta
from collections import defaultdict
from config import CLIENT_ID, API_KEY, OZON_API_BASE


def _headers():
    return {"Client-Id": CLIENT_ID, "Api-Key": API_KEY, "Content-Type": "application/json"}


# Кратковременные сетевые сбои (таймаут чтения, обрыв соединения, DNS) не должны
# ронять весь прогон: до этого ReadTimeout на первом же /v3/product/list убивал
# отчёт целиком. Ретраим их так же, как rate limit (ср. _request_raw_with_retry
# в yandex-агенте).
_NET_ERRORS = (requests.exceptions.ConnectionError, requests.exceptions.Timeout)


class ObsoleteEndpoint(Exception):
    """Метод отключён на стороне Озона. Не сетевой сбой и не наша ошибка —
    повторять бесполезно, нужен другой источник данных."""


def _post(path, body, timeout=60, retries=4):
    """POST с ретраем на rate limit (code 8), 429 и СЕТЕВЫЕ сбои."""
    last_exc = None
    for attempt in range(retries):
        try:
            r = requests.post(f"{OZON_API_BASE}{path}", headers=_headers(),
                              json=body, timeout=timeout)
        except _NET_ERRORS as e:
            last_exc = e
            if attempt == retries - 1:
                break
            wait = 10 * (attempt + 1)
            print(f"  сетевой сбой на {path} ({type(e).__name__}), "
                  f"повтор через {wait}s ({attempt+1}/{retries})...")
            time.sleep(wait)
            continue
        # Озон отключает старые версии методов, отвечая 400 «obsolete method
        # cannot be used» (20.09.2026 так закрылись оба /v3/finance/transaction/*).
        # Ретраить бессмысленно — метода больше нет, нужен отдельный сигнал наверх.
        # С 10.2026 /v3/finance/transaction/* отвечают уже 404 — тот же смысл.
        if (r.status_code == 400 and "obsolete method" in r.text) or \
                (r.status_code == 404 and path.startswith("/v3/finance/transaction")):
            raise ObsoleteEndpoint(f"метод {path} отключён Озоном: {r.status_code} {r.text[:120]}")
        if r.status_code == 429 or (r.status_code == 400 and '"code":8' in r.text):
            wait = 5 * (attempt + 1)
            print(f"  rate limit на {path}, ждём {wait}s...")
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r.json()
    if last_exc is not None:
        raise last_exc
    r.raise_for_status()
    return r.json()


# ── Каталог: маппинг sku → offer_id ───────────────────────────────────────────

def get_product_map() -> dict:
    """
    Returns:
    {
      "sku2offer": { numeric_sku(int): offer_id(str) },
      "offers":    { offer_id: {"name": str, "product_id": int, "skus": [int,...]} },
    }
    """
    product_ids = []
    last_id = ""
    while True:
        data = _post("/v3/product/list", {"filter": {}, "limit": 1000, "last_id": last_id})
        res = data.get("result", {})
        items = res.get("items", [])
        product_ids += [i["product_id"] for i in items]
        last_id = res.get("last_id", "")
        if not last_id or not items:
            break
        time.sleep(0.3)

    sku2offer = {}
    offers = {}
    # product/info/list принимает до 1000 product_id за раз
    for i in range(0, len(product_ids), 900):
        chunk = product_ids[i:i + 900]
        data = _post("/v3/product/info/list", {"product_id": chunk})
        items = data.get("items") or data.get("result", {}).get("items", [])
        for it in items:
            offer_id = it.get("offer_id", "")
            name = it.get("name", offer_id)
            pid = it.get("id")
            skus = set()
            if it.get("sku"):
                skus.add(int(it["sku"]))
            for src in (it.get("sources") or []):
                if src.get("sku"):
                    skus.add(int(src["sku"]))
            offers[offer_id] = {"name": name, "product_id": pid, "skus": sorted(skus)}
            for s in skus:
                sku2offer[s] = offer_id
        time.sleep(0.3)
    return {"sku2offer": sku2offer, "offers": offers}


# ── Продажи по дням (единственные живые метрики аналитики) ────────────────────

def get_sales_by_day(date_from: date, date_to: date, sku2offer: dict) -> dict:
    """
    Продажи по дням через /v1/analytics/data (dimension=[sku,day]).
    Живые метрики: ordered_units, revenue.

    Returns: { offer_id: { "YYYY-MM-DD": {"units": float, "revenue": float} } }
    """
    result = defaultdict(lambda: defaultdict(lambda: {"units": 0.0, "revenue": 0.0}))
    offset = 0
    while True:
        data = _post("/v1/analytics/data", {
            "date_from": date_from.isoformat(),
            "date_to": date_to.isoformat(),
            "dimension": ["sku", "day"],
            "metrics": ["ordered_units", "revenue"],
            "limit": 1000,
            "offset": offset,
        })
        rows = data.get("result", {}).get("data", [])
        if not rows:
            break
        for row in rows:
            dims = row["dimensions"]
            sku = int(dims[0]["id"])
            day = dims[1]["id"]
            units, revenue = row["metrics"][0], row["metrics"][1]
            offer_id = sku2offer.get(sku, f"sku:{sku}")
            cell = result[offer_id][day]
            cell["units"] += units
            cell["revenue"] += revenue
        if len(rows) < 1000:
            break
        offset += 1000
        time.sleep(1.0)
    return {k: dict(v) for k, v in result.items()}


# ── Остатки FBO / FBS ──────────────────────────────────────────────────────────

def get_stocks() -> dict:
    """
    Returns: { offer_id: {"fbo": int, "fbs": int, "reserved_fbo": int, "reserved_fbs": int} }
    """
    stocks = defaultdict(lambda: {"fbo": 0, "fbs": 0, "reserved_fbo": 0, "reserved_fbs": 0})
    cursor = ""
    while True:
        data = _post("/v4/product/info/stocks", {
            "cursor": cursor, "filter": {"visibility": "ALL"}, "limit": 1000,
        })
        items = data.get("items", [])
        for it in items:
            offer_id = it.get("offer_id", "")
            for s in it.get("stocks", []):
                typ = (s.get("type") or "").lower()  # "fbo" / "fbs"
                present = s.get("present", 0) or 0
                reserved = s.get("reserved", 0) or 0
                if typ == "fbo":
                    stocks[offer_id]["fbo"] += present
                    stocks[offer_id]["reserved_fbo"] += reserved
                elif typ == "fbs":
                    stocks[offer_id]["fbs"] += present
                    stocks[offer_id]["reserved_fbs"] += reserved
        cursor = data.get("cursor", "")
        if not cursor or not items:
            break
        time.sleep(0.3)
    return dict(stocks)


# ── Юнит-экономика и возвраты (finance/transaction) ───────────────────────────

def _fetch_transactions(date_from: date, date_to: date) -> list[dict]:
    """Все операции finance/transaction за период (с пагинацией)."""
    ops = []
    page = 1
    while True:
        data = _post("/v3/finance/transaction/list", {
            "filter": {
                "date": {
                    "from": date_from.isoformat() + "T00:00:00.000Z",
                    "to": date_to.isoformat() + "T23:59:59.999Z",
                },
                "transaction_type": "all",
            },
            "page": page, "page_size": 1000,
        })
        res = data.get("result", {})
        chunk = res.get("operations", [])
        ops += chunk
        page_count = res.get("page_count", 0)
        if page >= page_count or not chunk:
            break
        page += 1
        time.sleep(0.5)
    return ops


def get_unit_economics(date_from: date, date_to: date, sku2offer: dict) -> dict:
    """
    Юнит-экономика по доставленным заказам + сводка возвратов и накладных расходов.

    Returns:
    {
      "sales": [ {offer_id, name, channel, revenue, commission, services_total,
                  services, net, units, cluster_to} ],   # 1 строка = 1 доставленный товар-позиция
      "returns": { offer_id: {"count": int, "amount": float} },
      "overhead": { "Хранение": float, "Эквайринг": float, "Штрафы": float, ... },
    }
    channel: "FBS" / "FBO" из posting.delivery_schema.
    """
    ops = _fetch_transactions(date_from, date_to)

    sales = []
    returns = defaultdict(lambda: {"count": 0, "amount": 0.0})
    overhead = defaultdict(float)

    for o in ops:
        otype = o.get("operation_type", "")
        posting = o.get("posting") or {}
        channel = (posting.get("delivery_schema") or "").upper()  # FBS/FBO
        items = o.get("items") or []
        services = o.get("services") or []
        services_total = round(sum(s.get("price", 0) for s in services), 2)

        # Доставка покупателю — продажа
        if otype == "OperationAgentDeliveredToCustomer":
            revenue = o.get("accruals_for_sale", 0) or 0
            commission = o.get("sale_commission", 0) or 0
            net = o.get("amount", 0) or 0
            n_items = len(items) or 1
            # делим order-level суммы поровну по позициям
            for it in items:
                sku = it.get("sku")
                offer_id = sku2offer.get(int(sku), f"sku:{sku}") if sku else "?"
                # агрегируем услуги по имени (на позицию — поровну)
                svc = {s["name"]: round(s.get("price", 0) / n_items, 2) for s in services}
                sales.append({
                    "offer_id": offer_id,
                    "name": it.get("name", offer_id),
                    "channel": channel or "FBS",
                    "revenue": round(revenue / n_items, 2),
                    "commission": round(commission / n_items, 2),
                    "services_total": round(services_total / n_items, 2),
                    "services": svc,
                    "net": round(net / n_items, 2),
                    "units": 1,
                })

        # Возвраты / отмены.
        # ВАЖНО различать (баг 2026-07-07, SKU показывал завышенные ₽/шт):
        #  • реальный возврат ВЫКУПЛЕННОГО товара (ClientReturn*/OperationItemReturn) —
        #    уменьшает и деньги, и количество чистых штук;
        #  • сервисные операции с "Return" в имени — OperationReturnGoodsFBSofRMS
        #    («доставка/обработка возврата, отмены, невыкупа» — услуга; невыкуп
        #    вообще не был в units) и SellerReturns* («вывоз товара со склада» —
        #    продавец забирает свой сток) — это ДЕНЬГИ (затраты), но НЕ штуки.
        elif "Return" in otype or "Cancel" in otype:
            amount = o.get("amount", 0) or 0
            is_real_client_return = ("ClientReturn" in otype
                                     or otype == "OperationItemReturn")
            for it in items:
                sku = it.get("sku")
                offer_id = sku2offer.get(int(sku), f"sku:{sku}") if sku else "?"
                if is_real_client_return:
                    returns[offer_id]["count"] += 1
                returns[offer_id]["amount"] += round(amount / (len(items) or 1), 2)

        # Накладные расходы (не привязаны к конкретному SKU)
        else:
            amount = o.get("amount", 0) or 0
            name = o.get("operation_type_name") or otype
            overhead[name] += amount

    return {
        "sales": sales,
        "returns": {k: v for k, v in returns.items()},
        "overhead": {k: round(v, 2) for k, v in overhead.items()},
    }


# ── Отправления: кластеры доставки (логистика по регионам) ────────────────────

def get_postings_clusters(date_from: date, date_to: date) -> dict:
    """
    Кластеры назначения (cluster_to) по доставленным отправлениям FBS+FBO.

    Returns: { cluster_to: {"count": int, "revenue": float, "channel": {FBS:n, FBO:n}} }
    """
    clusters = defaultdict(lambda: {"count": 0, "revenue": 0.0,
                                    "channel": defaultdict(int)})
    for schema, path in (("FBS", "/v3/posting/fbs/list"), ("FBO", "/v2/posting/fbo/list")):
        offset = 0
        while True:
            body = {
                "dir": "DESC",
                "filter": {
                    "since": date_from.isoformat() + "T00:00:00.000Z",
                    "to": date_to.isoformat() + "T23:59:59.999Z",
                },
                "limit": 1000, "offset": offset,
                "with": {"financial_data": True, "analytics_data": True},
            }
            data = _post(path, body)
            res = data.get("result", [])
            # FBS: result={"postings":[...]}; FBO: result=[...] напрямую
            postings = res if isinstance(res, list) else res.get("postings", [])
            if not postings:
                break
            for p in postings:
                fd = p.get("financial_data") or {}
                cluster = fd.get("cluster_to") or (p.get("analytics_data") or {}).get("region") or "—"
                rev = 0.0
                for pr in (fd.get("products") or []):
                    rev += (pr.get("price", 0) or 0) * (pr.get("quantity", 1) or 1)
                c = clusters[cluster]
                c["count"] += 1
                c["revenue"] += rev
                c["channel"][schema] += 1
            if len(postings) < 1000:
                break
            offset += 1000
            time.sleep(0.5)
    return {k: {"count": v["count"], "revenue": round(v["revenue"], 2),
                "channel": dict(v["channel"])} for k, v in clusters.items()}


# ── Цены ────────────────────────────────────────────────────────────────────

def get_prices() -> dict:
    """Returns { offer_id: price(float) } из /v5/product/info/prices."""
    prices = {}
    cursor = ""
    while True:
        data = _post("/v5/product/info/prices", {
            "cursor": cursor, "filter": {"visibility": "ALL"}, "limit": 1000,
        })
        items = data.get("items") or data.get("result", {}).get("items", [])
        for it in items:
            offer_id = it.get("offer_id", "")
            p = it.get("price", {})
            val = p.get("price") or p.get("marketing_price") or 0
            try:
                prices[offer_id] = float(val)
            except (TypeError, ValueError):
                pass
        cursor = data.get("cursor", "")
        if not cursor or not items:
            break
        time.sleep(0.3)
    return prices


def get_realization(year: int, month: int) -> dict:
    """Отчёт о реализации за календарный месяц — ПРОДАЖНАЯ часть экономики по SKU.

    Замена отключённому /v3/finance/transaction/list, но НЕПОЛНАЯ: здесь есть
    выручка, комиссия, баллы Озона и возвраты, но НЕТ услуг (логистика, эквайринг,
    обработка отправлений, промо, платные выплаты) — они только в выгрузке
    «Отчёт по начислениям» из кабинета. Поэтому метод используется как источник
    свежих продаж, а не как полная юнит-экономика.

    Период задаётся ТОЛЬКО месяцем целиком — произвольное окно API не принимает.

    Returns: {offer_id: {name, qty, ret_qty, revenue, commission, bonus, stars, payout}}
    """
    data = _post("/v2/finance/realization", {"year": int(year), "month": int(month)})
    out = {}
    for row in (data.get("result", {}) or {}).get("rows", []) or []:
        item = row.get("item") or {}
        oid = item.get("offer_id") or ""
        if not oid:
            continue
        dc = row.get("delivery_commission") or {}
        rc = row.get("return_commission") or {}
        qty = float(dc.get("quantity") or 0)
        a = out.setdefault(oid, {"name": item.get("name", ""), "qty": 0.0, "ret_qty": 0.0,
                                 "revenue": 0.0, "commission": 0.0, "bonus": 0.0,
                                 "stars": 0.0, "payout": 0.0})
        a["qty"] += qty
        a["ret_qty"] += float(rc.get("quantity") or 0)
        a["revenue"] += float(row.get("seller_price_per_instance") or 0) * qty
        # standard_fee — комиссия за продажу; в отчёте она положительная, приводим
        # к тому же знаку, что и в начислениях (расход = минус)
        a["commission"] -= float(dc.get("standard_fee") or 0)
        a["bonus"] += float(dc.get("bonus") or 0)
        a["stars"] -= float(dc.get("stars") or 0)
        a["payout"] += float(dc.get("total") or 0)
    return out


def get_realization_range(date_from: date, date_to: date) -> dict:
    """Реализация за все месяцы, которые задевает период (API умеет только помесячно).
    Месяцы суммируются; вернётся агрегат по offer_id плюс список охваченных месяцев."""
    months, cur = [], date(date_from.year, date_from.month, 1)
    while cur <= date_to:
        months.append((cur.year, cur.month))
        cur = date(cur.year + 1, 1, 1) if cur.month == 12 else date(cur.year, cur.month + 1, 1)
    agg, done = {}, []
    for y, m in months:
        try:
            part = get_realization(y, m)
        except Exception as e:
            print(f"   ⚠️  реализация за {y}-{m:02d} недоступна: {str(e)[:90]}")
            continue
        done.append(f"{y}-{m:02d}")
        for oid, v in part.items():
            a = agg.setdefault(oid, {k: (v[k] if k == "name" else 0.0) for k in v})
            for k, val in v.items():
                if k != "name":
                    a[k] += val
        time.sleep(0.3)
    return {"by_offer": agg, "months": done}
