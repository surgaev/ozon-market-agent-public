#!/usr/bin/env python3
"""
Ozon Seller daily analytics agent.

Период по умолчанию: последние 14 дней до вчера (сегодня не берём — данные копятся).

Usage:
    OZON_CLIENT_ID=... OZON_API_KEY=... python main.py [--date YYYY-MM-DD] [--days N]
"""

import argparse
import json
import os
import sys
import traceback
from datetime import date, timedelta

from config import (REPORT_FOLDER, OZON_HEALTH_CSV, DOWNLOADS_DIR,
                    OZON_ACCRUALS_PATTERNS, OZON_RETURNS_PATTERNS, EXCLUDED_SKUS)
import ozon_client as oc


def _drop_excluded(obj, skus):
    """Убирает снятые с продажи артикулы из входных данных отчёта (та же логика,
    что в yandex-агенте): ключ словаря = offer_id — долой, строки-словари с полем
    артикула — долой, остальное — рекурсивно вглубь."""
    if not skus:
        return obj
    if hasattr(obj, "columns") and hasattr(obj, "isin"):          # DataFrame
        for col in ("offer_id", "shopSku"):
            if col in obj.columns:
                return obj[~obj[col].isin(skus)]
        return obj
    if isinstance(obj, dict):
        return {k: _drop_excluded(v, skus) for k, v in obj.items()
                if not (k in skus or (isinstance(k, tuple) and k and k[0] in skus))}
    if isinstance(obj, (list, set)):
        keep = [_drop_excluded(x, skus) for x in obj
                if not (isinstance(x, str) and x in skus)
                and not (isinstance(x, dict) and any(x.get(f) in skus
                         for f in ("offer_id", "sku", "shopSku", "offerId")))]
        return set(keep) if isinstance(obj, set) else keep
    return obj
from report_generator import build_report
from web_exports import (parse_products_health, parse_accruals_by_product,
                         parse_returns, find_latest)


def _aggregate_econ(unit_econ: dict) -> dict:
    """Свернуть unit_econ (позиции + возвраты) в компактный dict по offer_id."""
    agg = {}
    for s in unit_econ.get("sales", []):
        oid = s["offer_id"]
        a = agg.setdefault(oid, {"name": s.get("name", oid), "units": 0.0, "revenue": 0.0,
                                 "commission": 0.0, "services": 0.0, "net": 0.0,
                                 "ret_qty": 0.0, "ret_amount": 0.0})
        a["units"] += s.get("units", 0)
        a["revenue"] += s.get("revenue", 0)
        a["commission"] += s.get("commission", 0)
        a["services"] += s.get("services_total", 0)
        a["net"] += s.get("net", 0)
    for oid, r in unit_econ.get("returns", {}).items():
        a = agg.setdefault(oid, {"name": oid, "units": 0.0, "revenue": 0.0, "commission": 0.0,
                                 "services": 0.0, "net": 0.0, "ret_qty": 0.0, "ret_amount": 0.0})
        a["ret_qty"] += r.get("count", 0)
        a["ret_amount"] += r.get("amount", 0)
    return agg


def _get_econ_quarter_cached(run_date, date_from, date_to, product_map, valid_days=7):
    """90-дневная юнит-экономика по offer_id с кэшем на 7 дней (ozon_econ_quarter_cache.json).
    Устойчивая факт-прибыль (без лага выплат). При сбое — пусто (лист покажет «—»)."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ozon_econ_quarter_cache.json")
    if os.path.exists(path):
        try:
            c = json.load(open(path, encoding="utf-8"))
            cached_on = date.fromisoformat(c.get("date", "2000-01-01"))
            if (run_date - cached_on).days < valid_days and c.get("to") == date_to.isoformat():
                return c["data"]
        except Exception:
            pass
    try:
        # Ozon transaction/list ограничен ~месяцем на запрос → тянем 30-дн чанками.
        agg = {}
        cur = date_from
        while cur <= date_to:
            chunk_to = min(cur + timedelta(days=29), date_to)
            ue = oc.get_unit_economics(cur, chunk_to, product_map["sku2offer"])
            for oid, a in _aggregate_econ(ue).items():
                acc = agg.setdefault(oid, {"name": a["name"], "units": 0.0, "revenue": 0.0,
                                           "commission": 0.0, "services": 0.0, "net": 0.0,
                                           "ret_qty": 0.0, "ret_amount": 0.0})
                for k in ("units", "revenue", "commission", "services", "net", "ret_qty", "ret_amount"):
                    acc[k] += a[k]
            cur = chunk_to + timedelta(days=1)
    except Exception as e:
        print(f"   90-дн юнит-экономика недоступна ({e}); ₽/шт 90д будет пустой")
        return {}
    try:
        json.dump({"date": run_date.isoformat(), "from": date_from.isoformat(),
                   "to": date_to.isoformat(), "data": agg},
                  open(path, "w", encoding="utf-8"), ensure_ascii=False)
    except Exception:
        pass
    return agg


def run(run_date: date, days: int = 14):
    print(f"\n=== Ozon Seller Analytics — {run_date} ===\n")
    date_to = run_date - timedelta(days=1)
    date_from = date_to - timedelta(days=days - 1)
    print(f"Период: {date_from} → {date_to}\n")

    print("1. Каталог + маппинг sku→offer_id...")
    product_map = oc.get_product_map()
    print(f"   offer'ов: {len(product_map['offers'])}, sku в маппинге: {len(product_map['sku2offer'])}")

    print("2. Продажи по дням (analytics)...")
    sales_by_day = oc.get_sales_by_day(date_from, date_to, product_map["sku2offer"])
    print(f"   offer'ов с продажами: {len(sales_by_day)}")

    print("3. Остатки FBO/FBS...")
    stocks = oc.get_stocks()
    print(f"   offer'ов с записями остатков: {len(stocks)}")

    print("4. Юнит-экономика (finance/transaction)...")
    # 20.09.2026 Озон отключил /v3/finance/transaction/* («obsolete method»), замены
    # в API нет. Деньги теперь берутся только из ручной выгрузки «Отчёт по
    # начислениям» (см. ниже). Прогон из-за этого не роняем: показы, остатки,
    # цены и логистика по-прежнему свежие.
    try:
        unit_econ = oc.get_unit_economics(date_from, date_to, product_map["sku2offer"])
        print(f"   доставленных позиций: {len(unit_econ['sales'])}, "
              f"возвратов: {len(unit_econ['returns'])}, статей накладных: {len(unit_econ['overhead'])}")
    except oc.ObsoleteEndpoint as e:
        unit_econ = {"sales": [], "returns": {}, "overhead": {}}
        print(f"   ⚠️  {e}")
        print("   → юнит-экономика из API недоступна; листы с деньгами будут "
              "строиться по выгрузке «Отчёт по начислениям» из кабинета")

    # 4а. Отчёт о реализации — ПРОДАЖНАЯ часть экономики через API (замена
    # отключённым транзакциям). Услуг в нём нет, поэтому он не заменяет выгрузку
    # начислений целиком: он нужен, чтобы свежие продажи были в отчёте даже когда
    # выгрузку из кабинета давно не скачивали.
    print("4а. Отчёт о реализации (продажи по SKU через API, помесячно)...")
    try:
        realization = oc.get_realization_range(date_from, date_to)
        print(f"   offer'ов: {len(realization['by_offer'])}; "
              f"месяцы: {', '.join(realization['months']) or '—'}")
    except Exception as e:
        realization = {"by_offer": {}, "months": []}
        print(f"   ⚠️  реализация недоступна ({str(e)[:110]})")

    print("4б. Квартальная юнит-экономика (90д, кэш 7д) — устойчивая факт-прибыль...")
    q_from = run_date - timedelta(days=90)
    econ_quarter = _get_econ_quarter_cached(run_date, q_from, date_to, product_map)
    print(f"   offer'ов за квартал: {len(econ_quarter)}")

    print("5. Кластеры доставки (логистика по регионам)...")
    clusters = oc.get_postings_clusters(date_from, date_to)
    print(f"   кластеров: {len(clusters)}")

    print("6. Цены...")
    prices = oc.get_prices()
    print(f"   цены для {len(prices)} offer'ов")

    print("\n   Веб-выгрузки из кабинета (ручной режим):")
    products_health = parse_products_health(OZON_HEALTH_CSV)
    print(f"   • Здоровье карточек: {len(products_health) or '—'} "
          + ("✓" if products_health else f"(нет {OZON_HEALTH_CSV})"))

    acc_path = find_latest(DOWNLOADS_DIR, OZON_ACCRUALS_PATTERNS)
    accruals = parse_accruals_by_product(acc_path) if acc_path else {}
    print(f"   • Начисления по товарам: {len(accruals) or '—'} "
          + (f"✓ ({os.path.basename(acc_path)})" if accruals else "(не найдено)"))

    ret_path = find_latest(DOWNLOADS_DIR, OZON_RETURNS_PATTERNS)
    returns_detail = parse_returns(ret_path) if ret_path else []
    print(f"   • Возвраты (детально): {len(returns_detail) or '—'} "
          + (f"✓ ({os.path.basename(ret_path)})" if returns_detail else "(не найдено)"))

    print("\n7. Генерируем Excel...")
    if EXCLUDED_SKUS:
        print(f"   Скрываем снятые с продажи: {', '.join(sorted(EXCLUDED_SKUS))}")
    path = (lambda **kw: build_report(**{k: _drop_excluded(v, EXCLUDED_SKUS)
                                         for k, v in kw.items()}))(
        run_date=run_date,
        product_map=product_map,
        sales_by_day=sales_by_day,
        stocks=stocks,
        unit_econ=unit_econ,
        clusters=clusters,
        prices=prices,
        report_folder=REPORT_FOLDER,
        date_from=date_from,
        date_to=date_to,
        products_health=products_health,
        accruals=accruals,
        returns_detail=returns_detail,
        econ_quarter=econ_quarter,
        realization=realization,
    )
    print(f"   Отчёт: {path}")
    print("\n=== Готово ===\n")
    return path


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="YYYY-MM-DD, default=today")
    parser.add_argument("--days", type=int, default=14, help="окно в днях (default 14)")
    args = parser.parse_args()
    run_date = date.fromisoformat(args.date) if args.date else date.today()
    try:
        run(run_date, args.days)
    except Exception:
        print("\n[ERROR]")
        traceback.print_exc()
        sys.exit(1)
