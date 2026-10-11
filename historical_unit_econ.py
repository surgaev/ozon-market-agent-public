#!/usr/bin/env python3
"""
Историческая юнит-экономика Ozon по каждому артикулу (offer_id).

Цель: понять выгодность всех артикулов, которые продавались за период (по умолч. 6 мес).
Источник: /v3/finance/transaction/list — фактические начисления (выручка, комиссия,
логистика+услуги, выплата), возвраты, и накладные расходы (хранение/реклама/эквайринг/
подписки и т.п., не привязанные к SKU).

Метод (как выбрал пользователь):
  • 1 строка = 1 артикул, FBS и FBO объединены (канал — отдельной колонкой-разбивкой).
  • Прибыль с разнесением накладных:
        прямая маржа = выплата − себестоимость − налог 7%
        накладные разносим по артикулам пропорц. их доле в выручке
        чистая прибыль = прямая маржа − разнесённые накладные
  • finance/transaction/list ограничен 1 месяцем на запрос → тянем помесячными окнами.

Usage:
    set -a; source .env; set +a
    python historical_unit_econ.py [--months 6] [--date YYYY-MM-DD]
"""

import argparse
import sys
import traceback
from datetime import date, timedelta
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

import ozon_client as oc
from cost_prices import get_cost_price, TAX_RATE
from config import REPORT_FOLDER
from report_generator import (_fill, _hfont, _tfont, _border, _center, _left,
                              _style_header_row, RED_LIGHT, ORANGE_LT, YELLOW_LT,
                              GREEN_LT, GREY_LT, BLUE_BG, TEAL)


def _month_windows(date_from: date, date_to: date):
    """Разбивает [date_from, date_to] на окна не длиннее ~1 месяца (лимит API)."""
    # 28-дневные окна: всегда < 1 календарного месяца даже когда период
    # захватывает февраль (иначе Ozon отдаёт 400 «период больше 1 месяца»).
    windows = []
    cur = date_from
    while cur <= date_to:
        end = min(cur + timedelta(days=27), date_to)
        windows.append((cur, end))
        cur = end + timedelta(days=1)
    return windows


def collect_history(date_from: date, date_to: date, sku2offer: dict):
    """
    Тянем транзакции помесячно и агрегируем по offer_id.

    Returns:
      per_offer: { offer_id: {name, units, units_fbs, units_fbo, revenue, commission,
                              logistics, net, ret_count, ret_amount} }
      overhead_total: float (сумма накладных, обычно отрицательная)
      overhead_detail: { name: amount }
    """
    per_offer = defaultdict(lambda: {
        "name": "", "units": 0, "units_fbs": 0, "units_fbo": 0,
        "revenue": 0.0, "commission": 0.0, "logistics": 0.0, "net": 0.0,
        "ret_count": 0, "ret_amount": 0.0,
    })
    overhead_total = 0.0
    overhead_detail = defaultdict(float)

    windows = _month_windows(date_from, date_to)
    print(f"  окон по ≤30 дней: {len(windows)}")
    for i, (wf, wt) in enumerate(windows, 1):
        print(f"  [{i}/{len(windows)}] {wf} → {wt} ...", flush=True)
        ops = oc._fetch_transactions(wf, wt)
        print(f"      операций: {len(ops)}")
        for o in ops:
            otype = o.get("operation_type", "")
            posting = o.get("posting") or {}
            channel = (posting.get("delivery_schema") or "").upper()  # FBS/FBO
            items = o.get("items") or []
            services = o.get("services") or []
            services_total = sum(s.get("price", 0) for s in services)

            if otype == "OperationAgentDeliveredToCustomer":
                revenue = o.get("accruals_for_sale", 0) or 0
                commission = o.get("sale_commission", 0) or 0
                net = o.get("amount", 0) or 0
                n_items = len(items) or 1
                for it in items:
                    sku = it.get("sku")
                    oid = sku2offer.get(int(sku), f"sku:{sku}") if sku else "?"
                    a = per_offer[oid]
                    if not a["name"]:
                        a["name"] = it.get("name", oid)
                    a["units"] += 1
                    if channel == "FBO":
                        a["units_fbo"] += 1
                    else:
                        a["units_fbs"] += 1
                    a["revenue"] += revenue / n_items
                    a["commission"] += commission / n_items
                    a["logistics"] += services_total / n_items
                    a["net"] += net / n_items

            elif "Return" in otype or "Cancel" in otype or otype == "OperationItemReturn":
                amount = o.get("amount", 0) or 0
                n_items = len(items) or 1
                for it in items:
                    sku = it.get("sku")
                    oid = sku2offer.get(int(sku), f"sku:{sku}") if sku else "?"
                    a = per_offer[oid]
                    a["ret_count"] += 1
                    a["ret_amount"] += amount / n_items

            else:
                amount = o.get("amount", 0) or 0
                name = o.get("operation_type_name") or otype
                overhead_total += amount
                overhead_detail[name] += amount

    return (dict(per_offer), round(overhead_total, 2),
            {k: round(v, 2) for k, v in overhead_detail.items()})


def build_excel(per_offer, overhead_total, overhead_detail,
                date_from, date_to, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Юнит-экономика 6мес"
    ws.sheet_view.showGridLines = False

    # выручка только по проданным (для разнесения накладных)
    total_revenue = sum(a["revenue"] for a in per_offer.values()) or 1.0
    overhead_abs = -overhead_total  # накладные обычно отрицательны → расход положительный

    heads = [
        "Товар", "offer_id", "Канал\n(FBS/FBO)", "Продано\nшт.", "Возвр.\nшт.",
        # на 1 штуку
        "Цена\n₽/ед", "Комиссия\n₽/ед", "Комиссия\n%", "Логистика\n₽/ед",
        "Себест.\n₽/ед", "Прибыль\n₽/ед", "Маржа\n%",
        # за период
        "Выручка\n₽", "Комиссия\n₽", "Логистика\n₽", "Себест.\n₽", "Налог\n7% ₽",
        "Прямая\nмаржа ₽", "Накладные\n(разнес.) ₽", "Возвраты\n₽",
        "Чистая\nприбыль ₽", "Чистая\nмаржа %",
    ]
    ncol = len(heads)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    ndays = (date_to - date_from).days + 1
    ws["A1"] = (f"Историческая юнит-экономика Ozon по артикулам за {date_from}–{date_to} "
                f"({ndays} дн.). 1 строка = 1 артикул (FBS+FBO вместе). "
                f"Прямая маржа = выплата − себест. − налог 7%. "
                f"Накладные (хранение/реклама/эквайринг/подписки) = {overhead_abs:,.0f}₽ "
                f"разнесены пропорц. выручке. Чистая прибыль = прямая маржа − накладные + возвраты "
                f"(возвраты обычно отрицательны = расход). "
                f"Сортировка: по offer_id. Красная маржа → артикул убыточен.")
    ws["A1"].font = _tfont(sz=10, color="555555"); ws["A1"].alignment = _left()

    # групповые подзаголовки
    ws.merge_cells(start_row=2, start_column=6, end_row=2, end_column=12)
    g1 = ws.cell(row=2, column=6, value="НА 1 ШТУКУ")
    g1.fill = _fill(BLUE_BG); g1.font = _hfont(); g1.alignment = _center()
    ws.merge_cells(start_row=2, start_column=13, end_row=2, end_column=ncol)
    g2 = ws.cell(row=2, column=13, value="ЗА ПЕРИОД (все штуки)")
    g2.fill = _fill(TEAL); g2.font = _hfont(); g2.alignment = _center()
    for c in range(1, 6):
        ws.cell(row=2, column=c).fill = _fill(GREY_LT)

    hr = 3
    for c, h in enumerate(heads, 1):
        ws.cell(row=hr, column=c, value=h)
    _style_header_row(ws, hr, ncol, TEAL)
    for c in range(6, 13):
        ws.cell(row=hr, column=c).fill = _fill(BLUE_BG)
    ws.freeze_panes = "F4"
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 12
    for c in range(4, ncol + 1):
        ws.column_dimensions[get_column_letter(c)].width = 11

    # считаем строки
    rows = []
    missing_cost = []
    for oid, a in per_offer.items():
        if oid == "?" or oid.startswith("sku:"):
            # неизвестный артикул — пропускаем в маппинге, но учитываем ниже отдельно
            pass
        units = a["units"]
        if units == 0:
            continue
        revenue = a["revenue"]
        cost = get_cost_price(oid)
        if cost is None and not oid.startswith("sku:"):
            missing_cost.append(oid)
        cost = cost or 0
        cost_total = cost * units
        tax = revenue * TAX_RATE
        direct = a["net"] - cost_total - tax
        overhead_alloc = overhead_abs * (revenue / total_revenue)
        ret_amount = a["ret_amount"]  # обычно отрицательно = расход на возвраты
        net_profit = direct - overhead_alloc + ret_amount
        net_margin = (net_profit / revenue * 100) if revenue else 0
        rows.append({
            "oid": oid, "a": a, "units": units, "revenue": revenue, "cost": cost,
            "cost_total": cost_total, "tax": tax, "direct": direct,
            "overhead_alloc": overhead_alloc, "ret_amount": ret_amount,
            "net_profit": net_profit, "net_margin": net_margin,
        })

    rows.sort(key=lambda x: x["oid"])

    tot = defaultdict(float)
    r = hr + 1
    for row in rows:
        oid = row["oid"]; a = row["a"]; units = row["units"]; revenue = row["revenue"]
        u = units or 1
        ch = "FBS" if a["units_fbs"] and not a["units_fbo"] else \
             "FBO" if a["units_fbo"] and not a["units_fbs"] else \
             f"FBS {a['units_fbs']}/FBO {a['units_fbo']}"
        rev_u = revenue / u
        comm_u = a["commission"] / u
        comm_pct = (a["commission"] / revenue * 100) if revenue else 0
        log_u = a["logistics"] / u
        profit_u = row["direct"] / u
        margin_u = (profit_u / rev_u * 100) if rev_u else 0

        ws.cell(row=r, column=1, value=(a["name"] or oid)[:60])
        ws.cell(row=r, column=2, value=oid)
        ws.cell(row=r, column=3, value=ch)
        ws.cell(row=r, column=4, value=units)
        rcc = ws.cell(row=r, column=5, value=a["ret_count"] or None)
        # на 1 штуку
        ws.cell(row=r, column=6, value=round(rev_u, 2))
        ws.cell(row=r, column=7, value=round(comm_u, 2))
        ws.cell(row=r, column=8, value=round(comm_pct, 1))
        ws.cell(row=r, column=9, value=round(log_u, 2))
        ws.cell(row=r, column=10, value=round(row["cost"], 2))
        puc = ws.cell(row=r, column=11, value=round(profit_u, 2))
        muc = ws.cell(row=r, column=12, value=round(margin_u, 1))
        # за период
        ws.cell(row=r, column=13, value=round(revenue, 2))
        ws.cell(row=r, column=14, value=round(a["commission"], 2))
        ws.cell(row=r, column=15, value=round(a["logistics"], 2))
        ws.cell(row=r, column=16, value=round(row["cost_total"], 2))
        ws.cell(row=r, column=17, value=round(row["tax"], 2))
        dc = ws.cell(row=r, column=18, value=round(row["direct"], 2))
        ws.cell(row=r, column=19, value=round(row["overhead_alloc"], 2))
        retc = ws.cell(row=r, column=20,
                       value=round(row["ret_amount"], 2) if row["ret_amount"] else None)
        npc = ws.cell(row=r, column=21, value=round(row["net_profit"], 2))
        nmc = ws.cell(row=r, column=22, value=round(row["net_margin"], 1))

        if a["ret_count"]:
            rcc.fill = _fill(ORANGE_LT); rcc.font = _tfont(bold=True, color="E65100")
        if row["ret_amount"]:
            retc.font = _tfont(color="B71C1C" if row["ret_amount"] < 0 else "1B5E20")
        puc.font = _tfont(bold=True, color="1B5E20" if profit_u > 0 else "B71C1C")
        dc.font = _tfont(bold=True, color="1B5E20" if row["direct"] > 0 else "B71C1C")
        npc.font = _tfont(bold=True, color="1B5E20" if row["net_profit"] > 0 else "B71C1C")
        muc.fill = _fill(GREEN_LT if margin_u >= 25 else RED_LIGHT if margin_u < 10 else YELLOW_LT)
        nmc.fill = _fill(GREEN_LT if row["net_margin"] >= 20 else
                         RED_LIGHT if row["net_margin"] < 5 else YELLOW_LT)
        for c in range(1, ncol + 1):
            cell = ws.cell(row=r, column=c); cell.border = _border()
            cell.alignment = _left() if c in (1,) else _center() if c > 2 else _left()

        tot["units"] += units
        tot["ret_count"] += a["ret_count"]
        tot["commission"] += a["commission"]
        tot["logistics"] += a["logistics"]
        for k in ("revenue", "cost_total", "tax", "direct", "overhead_alloc",
                  "ret_amount", "net_profit"):
            tot[k] += row[k]
        r += 1

    # ИТОГО
    ws.cell(row=r, column=1, value="ИТОГО").font = _tfont(bold=True)
    ws.cell(row=r, column=4, value=int(tot["units"])).font = _tfont(bold=True)
    ws.cell(row=r, column=5, value=int(tot["ret_count"]) or None)
    for col, k in ((13, "revenue"), (14, "commission"), (15, "logistics"),
                   (16, "cost_total"), (17, "tax"), (18, "direct"),
                   (19, "overhead_alloc"), (20, "ret_amount"), (21, "net_profit")):
        ws.cell(row=r, column=col, value=round(tot[k], 2)).font = _tfont(bold=True)
    tot_net_margin = (tot["net_profit"] / tot["revenue"] * 100) if tot["revenue"] else 0
    ws.cell(row=r, column=22, value=round(tot_net_margin, 1)).font = _tfont(bold=True)
    for c in range(1, ncol + 1):
        cell = ws.cell(row=r, column=c); cell.border = _border()
        cell.fill = _fill(GREY_LT); cell.alignment = _center() if c > 2 else _left()
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(ncol)}{r-1}"

    # ── Лист 2: расшифровка накладных ──
    ws2 = wb.create_sheet("Накладные расходы")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    ws2["A1"] = (f"Накладные расходы Ozon за {date_from}–{date_to}, не привязанные к SKU "
                 f"(разнесены по артикулам пропорц. выручке). Итого {overhead_abs:,.0f}₽.")
    ws2["A1"].font = _tfont(sz=10, color="555555"); ws2["A1"].alignment = _left()
    ws2.cell(row=2, column=1, value="Статья"); ws2.cell(row=2, column=2, value="Сумма ₽")
    _style_header_row(ws2, 2, 2)
    ws2.column_dimensions["A"].width = 55; ws2.column_dimensions["B"].width = 16
    rr = 3
    for name, amt in sorted(overhead_detail.items(), key=lambda x: x[1]):
        ws2.cell(row=rr, column=1, value=name)
        c = ws2.cell(row=rr, column=2, value=round(amt, 2))
        c.font = _tfont(color="B71C1C" if amt < 0 else "1B5E20")
        for cc in (1, 2):
            ws2.cell(row=rr, column=cc).border = _border()
        rr += 1
    ws2.cell(row=rr, column=1, value="ИТОГО").font = _tfont(bold=True)
    ws2.cell(row=rr, column=2, value=round(overhead_total, 2)).font = _tfont(bold=True)
    for cc in (1, 2):
        ws2.cell(row=rr, column=cc).border = _border(); ws2.cell(row=rr, column=cc).fill = _fill(GREY_LT)

    wb.save(out_path)
    return missing_cost, len(rows), round(tot["net_profit"], 2), round(tot["revenue"], 2)


def run(run_date: date, months: int):
    date_to = run_date - timedelta(days=1)
    date_from = date_to - timedelta(days=months * 30 - 1)
    print(f"\n=== Ozon историческая юнит-экономика: {date_from} → {date_to} ===\n")

    print("1. Каталог + маппинг sku→offer_id...")
    pm = oc.get_product_map()
    print(f"   offer'ов: {len(pm['offers'])}, sku: {len(pm['sku2offer'])}")

    print("2. Транзакции (помесячно)...")
    per_offer, overhead_total, overhead_detail = collect_history(
        date_from, date_to, pm["sku2offer"])
    sold = [o for o, a in per_offer.items() if a["units"] > 0
            and not o.startswith("sku:") and o != "?"]
    print(f"   артикулов с продажами: {len(sold)}; накладных статей: {len(overhead_detail)}")

    out_path = f"{REPORT_FOLDER}/ozon_unit_econ_history_{run_date}.xlsx"
    print("3. Excel...")
    missing, n_rows, net_profit, revenue = build_excel(
        per_offer, overhead_total, overhead_detail, date_from, date_to, out_path)
    print(f"\n   Отчёт: {out_path}")
    print(f"   Артикулов в таблице: {n_rows}")
    print(f"   Выручка: {revenue:,.0f}₽; чистая прибыль (после накладных): {net_profit:,.0f}₽")
    if missing:
        print(f"   ⚠ Без себестоимости ({len(missing)}): {', '.join(sorted(set(missing))[:30])}")
    print("\n=== Готово ===\n")
    return out_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="YYYY-MM-DD, default=today")
    parser.add_argument("--months", type=int, default=6, help="окно в месяцах (default 6)")
    args = parser.parse_args()
    run_date = date.fromisoformat(args.date) if args.date else date.today()
    try:
        run(run_date, args.months)
    except Exception:
        print("\n[ERROR]")
        traceback.print_exc()
        sys.exit(1)
