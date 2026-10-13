"""
Excel-отчёт Ozon — листы:
  1. Дашборд               — сводка продаж, разбивка FBO/FBS, накладные расходы, алерты
  2. Товары                — продажи 14д + остатки + цена по каждому offer_id
  3. Юнит-экономика        — комиссии Ozon, логистика, себестоимость, чистая прибыль/маржа
  4. Остатки FBO/FBS       — остатки + прогноз поставки FBO
  5. Продажи по дням 14д   — матрица offer × день (заказано шт.)
  6. Логистика по регионам — кластеры доставки (count, выручка, канал)
  7. Возвраты              — возвраты/отмены по offer_id

⚠ Воронки трафика (показы/CTR/буст) нет — Ozon отключил эти метрики в Seller API.
"""

import os
from datetime import date, timedelta
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .cost_prices import get_cost_price, TAX_RATE
from .config import RESTOCK_URGENT_DAYS, RESTOCK_SOON_DAYS, RESTOCK_PLAN_DAYS


def _fill_ym_compare(ws, row, oid, ozon_ppu, ym_blend, col_ym, col_diff):
    """Две колонки сравнения с Яндексом. Сопоставляем артикулы ТОЧНЫМ совпадением:
    префиксное «похоже» уже подводило (SKUbl ≠ SKUblk), а тут цена ошибки —
    сравнение экономики разных товаров. Нет пары на Яндексе → прочерк."""
    ym = ym_blend.get(str(oid))
    c1 = ws.cell(row=row, column=col_ym, value=ym if ym is not None else "—")
    if ym is None:
        c1.font = _tfont(color="999999")
        ws.cell(row=row, column=col_diff, value="—").font = _tfont(color="999999")
        return
    c1.font = _tfont(color="1B5E20" if ym > 0 else "B71C1C")
    diff = round((ozon_ppu or 0) - ym, 2)
    c2 = ws.cell(row=row, column=col_diff, value=diff)
    c2.font = _tfont(bold=True, color="1B5E20" if diff > 0 else "B71C1C")
    c2.fill = _fill(GREEN_LT if diff > 0 else RED_LIGHT)


def load_yandex_blend(report_folder: str) -> tuple[dict, str]:
    """⭐ВЗВЕШ ₽/шт по артикулам из САМОГО СВЕЖЕГО отчёта Яндекса в той же папке.

    Читаем готовый xlsx, а не считаем заново: у Яндекс-агента своя механика
    (netting, бакеты 0–30/31–90 дней), дублировать её здесь незачем и опасно —
    разъедется. Берём то, что он уже посчитал.
    Returns: ({sku: ₽/шт}, дата отчёта 'YYYY-MM-DD' или '').
    """
    import glob
    import openpyxl
    files = sorted(glob.glob(os.path.join(report_folder, "ym_report_*.xlsx")))
    if not files:
        return {}, ""
    path = files[-1]
    rep_date = os.path.basename(path).replace("ym_report_", "").replace(".xlsx", "")
    try:
        ws = openpyxl.load_workbook(path, data_only=True, read_only=True)["Сверка выплат"]
    except Exception:
        return {}, ""
    rows = ws.iter_rows(min_row=2, values_only=True)
    head = next(rows, None)
    if not head:
        return {}, ""
    try:
        col = [i for i, h in enumerate(head) if h and "ВЗВЕШ" in str(h)][0]
        sku_col = [i for i, h in enumerate(head) if str(h).strip() == "SKU"][0]
    except IndexError:
        return {}, ""
    out = {}
    for row in rows:
        if len(row) <= max(col, sku_col):
            continue
        sku, val = row[sku_col], row[col]
        if sku and isinstance(val, (int, float)):
            out[str(sku)] = round(float(val), 2)   # одинакова на всех строках SKU
    return out, rep_date

# ── Palette (как в Yandex-отчёте) ─────────────────────────────────────────────
RED_BG = "FF4444"; RED_LIGHT = "FFC7CE"; ORANGE_BG = "FF8C00"; ORANGE_LT = "FFE0B2"
YELLOW_LT = "FFF9C4"; GREEN_LT = "C6EFCE"; GREEN_BG = "27AE60"; BLUE_BG = "2F4F8F"
BLUE_LT = "BDD7EE"; GREY_LT = "F2F2F2"; TEAL = "16A085"

def _fill(c):  return PatternFill("solid", fgColor=c)
def _hfont():  return Font(bold=True, color="FFFFFF", size=11)
def _tfont(sz=11, bold=False, color="000000"): return Font(bold=bold, color=color, size=sz)
def _border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)
def _center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def _left():   return Alignment(horizontal="left", vertical="center", wrap_text=True)

def _style_header_row(ws, row_num, num_cols, bg=BLUE_BG):
    for col in range(1, num_cols + 1):
        c = ws.cell(row=row_num, column=col)
        c.font = _hfont(); c.fill = _fill(bg); c.alignment = _center(); c.border = _border()

def _delta(today, ref):
    try:
        if not ref:
            return "—"
        d = round((today - ref) / ref * 100, 1)
        return f"+{d}%" if d >= 0 else f"{d}%"
    except Exception:
        return "—"


def build_report(run_date, product_map, sales_by_day, stocks, unit_econ,
                 clusters, prices, report_folder, date_from, date_to,
                 products_health=None, accruals=None, returns_detail=None,
                 econ_quarter=None, realization=None, accruals_info=None):
    os.makedirs(report_folder, exist_ok=True)
    offers = product_map["offers"]
    name_of = lambda oid: offers.get(oid, {}).get("name", oid)
    ndays = (date_to - date_from).days + 1
    econ_quarter = econ_quarter or {}

    # Список дней периода
    all_days = [(date_from + timedelta(days=i)).isoformat() for i in range(ndays)]

    # ── Сводные продажи по offer (за весь период) ──
    sales_tot = defaultdict(lambda: {"units": 0.0, "revenue": 0.0})
    for oid, byday in sales_by_day.items():
        for day, cell in byday.items():
            sales_tot[oid]["units"] += cell["units"]
            sales_tot[oid]["revenue"] += cell["revenue"]

    total_units = sum(v["units"] for v in sales_tot.values())
    total_revenue = sum(v["revenue"] for v in sales_tot.values())

    wb = Workbook()

    # ═══════════════════════════════════════════════════════════════════════════
    # SHEET 1: ДАШБОРД
    # ═══════════════════════════════════════════════════════════════════════════
    ws = wb.active
    ws.title = "Дашборд"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 36
    for col in "BCDEF":
        ws.column_dimensions[col].width = 16

    ws.merge_cells("A1:F1")
    ws["A1"] = f"OZON — отчёт · {date_from.isoformat()} → {date_to.isoformat()} ({ndays} дн.)"
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].fill = _fill(BLUE_BG)
    ws["A1"].alignment = _center()
    ws.row_dimensions[1].height = 26

    r = 3
    # Карточки сводки
    def card(row, col, title, value, sub="", color=TEAL):
        ws.cell(row=row, column=col, value=title).font = _tfont(sz=9, color="666666")
        vc = ws.cell(row=row + 1, column=col, value=value)
        vc.font = Font(bold=True, size=14, color=color)
        if sub:
            ws.cell(row=row + 2, column=col, value=sub).font = _tfont(sz=9, color="999999")

    # Канальная разбивка из юнит-экономики (доставленные)
    ch_units = defaultdict(int); ch_revenue = defaultdict(float); ch_net = defaultdict(float)
    for s in unit_econ["sales"]:
        ch_units[s["channel"]] += s["units"]
        ch_revenue[s["channel"]] += s["revenue"]
        ch_net[s["channel"]] += s["net"]
    # /v3/finance/transaction отключён (09.2026) → unit_econ пуст. Тогда доставки
    # и выплату берём из выгрузки начислений, иначе на дашборде одни нули.
    acc_promo = 0.0
    if not unit_econ["sales"] and accruals:
        for a in accruals.values():
            ch = a.get("channel") or "FBS"
            ch_units[ch] += a.get("qty", 0)
            ch_revenue[ch] += a.get("revenue", 0)
            ch_net[ch] += a.get("total", 0)
            acc_promo += sum(v for n, v in a.get("other_detail", {}).items() if "Звёзд" in n or "родвиж" in n)
    delivered_units = sum(ch_units.values())
    delivered_net = sum(ch_net.values())

    card(r, 1, "Заказано (шт.)", int(total_units))
    card(r, 2, "Заказано (₽)", round(total_revenue))
    card(r, 3, "Доставлено (шт.)", int(delivered_units))
    card(r, 4, "Выплата Ozon (₽)", round(delivered_net),
         sub="выручка − комиссия − логистика")
    overhead_sum = sum(unit_econ["overhead"].values())
    card(r, 5, "Накладные (₽)", round(overhead_sum),
         sub="хранение/эквайринг/штрафы", color="B71C1C" if overhead_sum < 0 else TEAL)

    # Продвижение/реклама: промо-статьи из начислений (Звёздные товары, Трафареты,
    # продвижение в поиске и т.п.). Ozon не отдаёт рекламу по SKU — только суммой.
    PROMO_KEYS = ("родвиж", "еклам", "Звёзд", "Трафарет", "Премиум", "Promo", "Буст", "буст")
    promo_sum = sum(v for n, v in unit_econ["overhead"].items()
                    if any(k in n for k in PROMO_KEYS)) + acc_promo
    ws.column_dimensions["F"].width = 16
    card(r, 6, "Продвижение (₽)", round(promo_sum),
         sub="промо/реклама (суммарно)", color="B71C1C" if promo_sum < 0 else TEAL)

    # ── Разбивка по каналам ──
    r = 8
    ws.cell(row=r, column=1, value="ПО КАНАЛАМ (доставленные)").font = Font(bold=True, size=12, color="FFFFFF")
    ws.cell(row=r, column=1).fill = _fill(BLUE_BG)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.cell(row=r, column=1).alignment = _center()
    r += 1
    heads = ["Канал", "Доставлено шт.", "Выручка ₽", "Выплата ₽"]
    for c, h in enumerate(heads, 1):
        ws.cell(row=r, column=c, value=h)
    _style_header_row(ws, r, len(heads), TEAL)
    r += 1
    for ch in ("FBS", "FBO"):
        ws.cell(row=r, column=1, value=ch)
        ws.cell(row=r, column=2, value=ch_units.get(ch, 0))
        ws.cell(row=r, column=3, value=round(ch_revenue.get(ch, 0), 2))
        ws.cell(row=r, column=4, value=round(ch_net.get(ch, 0), 2))
        for c in range(1, 5):
            cell = ws.cell(row=r, column=c)
            cell.border = _border(); cell.alignment = _center() if c > 1 else _left()
        r += 1

    # ── Накладные расходы ──
    r += 1
    ws.cell(row=r, column=1, value="НАКЛАДНЫЕ РАСХОДЫ (не на SKU)").font = Font(bold=True, size=12, color="FFFFFF")
    ws.cell(row=r, column=1).fill = _fill(ORANGE_BG)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.cell(row=r, column=1).alignment = _center()
    r += 1
    ws.cell(row=r, column=1, value="Статья"); ws.cell(row=r, column=2, value="Сумма ₽")
    _style_header_row(ws, r, 2, ORANGE_BG)
    r += 1
    for name, val in sorted(unit_econ["overhead"].items(), key=lambda x: x[1]):
        ws.cell(row=r, column=1, value=name)
        vc = ws.cell(row=r, column=2, value=round(val, 2))
        vc.font = _tfont(color="B71C1C" if val < 0 else "1B5E20")
        for c in (1, 2):
            cell = ws.cell(row=r, column=c); cell.border = _border()
            cell.alignment = _left() if c == 1 else _center()
        r += 1

    # Прогноз поставки нужен листу «Остатки FBO FBS» (блок алертов на дашборде убран).
    restock = _restock_table(stocks, sales_tot, ndays)

    # ═══════════════════════════════════════════════════════════════════════════
    # SHEET 2: ТОВАРЫ
    # ═══════════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Товары")
    ws2.sheet_view.showGridLines = False
    heads = ["Товар", "offer_id", f"Заказано\nшт. ({ndays}д)", f"Выручка\n₽ ({ndays}д)",
             "Ост.\nFBO", "Ост.\nFBS", "Резерв\nFBO+FBS", "Цена\n₽"]
    for c, h in enumerate(heads, 1):
        ws2.cell(row=1, column=c, value=h)
    _style_header_row(ws2, 1, len(heads))
    ws2.freeze_panes = "A2"
    ws2.column_dimensions["A"].width = 42
    ws2.column_dimensions["B"].width = 16
    for c in "CDEFGH":
        ws2.column_dimensions[c].width = 12

    rows2 = []
    all_oids = set(sales_tot) | set(stocks)
    for oid in all_oids:
        st = stocks.get(oid, {})
        s = sales_tot.get(oid, {"units": 0, "revenue": 0})
        rows2.append((oid, s["units"], s["revenue"], st.get("fbo", 0), st.get("fbs", 0),
                      st.get("reserved_fbo", 0) + st.get("reserved_fbs", 0), prices.get(oid, 0)))
    rows2.sort(key=lambda x: -x[2])  # по выручке
    r = 2
    for oid, units, revenue, fbo, fbs, reserved, price in rows2:
        if oid.startswith("sku:"):
            continue
        ws2.cell(row=r, column=1, value=name_of(oid))
        ws2.cell(row=r, column=2, value=oid)
        ws2.cell(row=r, column=3, value=int(units))
        ws2.cell(row=r, column=4, value=round(revenue, 2))
        ws2.cell(row=r, column=5, value=fbo)
        ws2.cell(row=r, column=6, value=fbs)
        ws2.cell(row=r, column=7, value=reserved)
        ws2.cell(row=r, column=8, value=price)
        if fbo == 0 and fbs == 0 and units > 0:
            ws2.cell(row=r, column=5).fill = _fill(RED_LIGHT)
        for c in range(1, len(heads) + 1):
            cell = ws2.cell(row=r, column=c); cell.border = _border()
            cell.alignment = _center() if c > 2 else _left()
        r += 1
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{r-1}"

    # ═══════════════════════════════════════════════════════════════════════════
    # SHEET 3: ЮНИТ-ЭКОНОМИКА (объединённая: факт из начислений + расчёт на единицу)
    # ═══════════════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Юнит-экономика")
    ws3.sheet_view.showGridLines = False
    missing_cost = []
    realization = realization or {}

    if accruals:
        # ── Слева — на 1 штуку, справа — итог за период (все штуки) ──
        heads = ["Товар", "offer_id", "Канал", "Доставл.\nшт.", "Возвр.\nшт.",
                 # НА 1 ШТУКУ (6-12)
                 "Выручка\n₽/ед", "Комиссия\n₽/ед", "Комиссия\n%", "Логистика\n₽/ед",
                 "Себест.\n₽/ед", "Прибыль\n₽/ед", "Маржа\n%",
                 # ЗА ПЕРИОД, все шт. (13-20)
                 "Выручка\n₽", "Комиссия\n₽", "Логистика\n₽",
                 "в т.ч. лог.\nвозвратов ₽", "Прочее\n₽",
                 "Из чего «Прочее»", "Итого\nвыплата ₽", "Прибыль\nза период ₽",
                 # сравнение с Яндексом (21-22)
                 "ЯМ ⭐ВЗВЕШ\n₽/шт", "Δ Озон−ЯМ\n₽/шт"]
        ncol = len(heads)
        ym_blend, ym_date = load_yandex_blend(report_folder)
        COL_YM, COL_YM_DIFF = 21, 22
        # кол-во возвратов по offer_id из детальной веб-выгрузки
        ret_count = defaultdict(int)
        for x in (returns_detail or []):
            ret_count[x["offer_id"]] += x.get("qty", 1)
        ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
        acc_note = ""
        if accruals_info:
            p1, p2 = accruals_info["period"]
            acc_note = (f"ДЕНЬГИ — начисления с {date_from:%d.%m} по {date_to:%d.%m.%Y} (по дате начисления) "
                        f"из выгрузки за {p1:%d.%m}–{p2:%d.%m.%Y}. "
                        + ("" if accruals_info["covered"] else
                           "⚠️ ВЫГРУЗКА ПОКРЫВАЕТ ПЕРИОД НЕ ПОЛНОСТЬЮ — за непокрытые дни денег нет, "
                           "нужна свежая выгрузка. "))
        ws3["A1"] = (acc_note + "Факт по начислениям Ozon. СЛЕВА — на 1 штуку, СПРАВА — итог за период по всем "
                     "доставленным штукам. Прибыль/ед = Выплата/ед − Себест. − налог 7%. "
                     "Себестоимость списывается только на ЧИСТЫЕ штуки (выкуплено − возвращено): "
                     "возврат приезжает годным на склад, товар не потерян, его реальная цена — "
                     "логистика туда-обратно и обработка (они уже сидят в выплате). "
                     "«Возвр. шт.» рядом с «Доставл. шт.» — если минус в итоге, смотри сюда. "
                     "«Из чего Прочее» — расшифровка подписок/штрафов/промо. "
                     "ЖЁЛТЫЕ строки с каналом «API» — продажи из отчёта о реализации, которых "
                     "ещё нет в выгрузке начислений (её качают вручную, она отстаёт). В них есть "
                     "выручка и комиссия, но НЕТ услуг: логистики, эквайринга, промо. Прибыль по "
                     "ним завышена — это оценка сверху. Чтобы они стали полными, нужна свежая "
                     "выгрузка «Отчёт по начислениям» из кабинета. "
                     + (f"Последние две колонки — сравнение с Яндексом: ⭐ВЗВЕШ ₽/шт из отчёта "
                        f"ym_report_{ym_date} (взвешенная прибыль на штуку, окна 0–30 и 31–90 дней "
                        f"с весом 2:1) и разница «Озон − Яндекс». Плюс = на Озоне артикул приносит "
                        f"больше. Артикулы сопоставляются точным совпадением offer_id и SKU; "
                        f"прочерк — пары на Яндексе нет."
                        if ym_blend else "Сравнение с Яндексом недоступно: отчёт ym_report_*.xlsx "
                        "не найден в папке."))
        ws3["A1"].font = _tfont(sz=10, color="555555"); ws3["A1"].alignment = _left()
        # Групповые подзаголовки (строка 2)
        ws3.merge_cells(start_row=2, start_column=6, end_row=2, end_column=12)
        g1 = ws3.cell(row=2, column=6, value="НА 1 ШТУКУ")
        g1.fill = _fill(BLUE_BG); g1.font = _hfont(); g1.alignment = _center()
        ws3.merge_cells(start_row=2, start_column=13, end_row=2, end_column=ncol)
        g2 = ws3.cell(row=2, column=13, value="ЗА ПЕРИОД (все штуки)")
        g2.fill = _fill(TEAL); g2.font = _hfont(); g2.alignment = _center()
        for c in range(1, 6):
            ws3.cell(row=2, column=c).fill = _fill(GREY_LT)
        hr = 3
        for c, h in enumerate(heads, 1):
            ws3.cell(row=hr, column=c, value=h)
        _style_header_row(ws3, hr, ncol, TEAL)
        # перекрасим шапку левой группы в синий для контраста
        for c in range(6, 13):
            ws3.cell(row=hr, column=c).fill = _fill(BLUE_BG)
        ws3.freeze_panes = "F4"
        ws3.column_dimensions["A"].width = 38
        ws3.column_dimensions["B"].width = 14
        for c in range(3, ncol + 1):
            ws3.column_dimensions[get_column_letter(c)].width = 11
        ws3.column_dimensions[get_column_letter(heads.index("Из чего «Прочее»") + 1)].width = 40

        tot = defaultdict(float)
        r = hr + 1
        for oid, a in sorted(accruals.items(), key=lambda x: -x[1]["total"]):
            cost = get_cost_price(oid)
            if cost is None:
                missing_cost.append(oid)
            cost = cost or 0
            qty = a["qty"] or 0
            # Себестоимость — только на ЧИСТЫЕ штуки (выкуплено − возвращено):
            # возврат приезжает годным на склад, товар не потерян. Списывать на
            # него закуп значило бы считать убыток дважды. Реальная цена возврата —
            # логистика туда-обратно и обработка, они и так сидят в a["total"].
            net_qty = max(0, qty - (a.get("ret_qty") or 0))
            cost_total = round(cost * net_qty, 2)
            tax = round(a["revenue"] * TAX_RATE, 2)
            profit = round(a["total"] - cost_total - tax, 2)
            margin = round(profit / a["revenue"] * 100, 1) if a["revenue"] else 0
            u = qty if qty else 1
            rev_u = round(a["revenue"] / u, 2)
            comm_u = round(a["commission"] / u, 2)
            comm_pct = round(a["commission"] / a["revenue"] * 100, 1) if a["revenue"] else 0
            log_u = round(a["logistics"] / u, 2)
            profit_u = round(profit / u, 2)
            other_detail = a.get("other_detail", {})
            why = "; ".join(f"{n}: {v:+.2f}" for n, v in
                            sorted(other_detail.items(), key=lambda x: x[1]))
            rcnt = ret_count.get(oid, 0)
            ret_log = a.get("ret_logistics", 0.0)
            ws3.cell(row=r, column=1, value=a["name"][:55] or name_of(oid))
            ws3.cell(row=r, column=2, value=oid)
            ws3.cell(row=r, column=3, value=a["channel"])
            ws3.cell(row=r, column=4, value=qty)
            rcc = ws3.cell(row=r, column=5, value=rcnt or None)
            # НА 1 ШТУКУ
            ws3.cell(row=r, column=6, value=rev_u)
            ws3.cell(row=r, column=7, value=comm_u)
            ws3.cell(row=r, column=8, value=comm_pct)
            ws3.cell(row=r, column=9, value=log_u)
            ws3.cell(row=r, column=10, value=cost)
            puc = ws3.cell(row=r, column=11, value=profit_u)
            mc = ws3.cell(row=r, column=12, value=margin)
            # ЗА ПЕРИОД
            ws3.cell(row=r, column=13, value=a["revenue"])
            ws3.cell(row=r, column=14, value=a["commission"])
            ws3.cell(row=r, column=15, value=a["logistics"])
            rlc = ws3.cell(row=r, column=16, value=round(ret_log, 2) if ret_log else None)
            ws3.cell(row=r, column=17, value=a["other"])
            ws3.cell(row=r, column=18, value=why)
            ws3.cell(row=r, column=19, value=a["total"])
            pc = ws3.cell(row=r, column=20, value=profit)
            _fill_ym_compare(ws3, r, oid, profit_u, ym_blend, COL_YM, COL_YM_DIFF)
            if rcnt:
                rcc.fill = _fill(ORANGE_LT); rcc.font = _tfont(bold=True, color="E65100")
                rlc.font = _tfont(color="E65100")
            puc.font = _tfont(bold=True, color="1B5E20" if profit_u > 0 else "B71C1C")
            pc.font = _tfont(bold=True, color="1B5E20" if profit > 0 else "B71C1C")
            mc.fill = _fill(GREEN_LT if margin >= 25 else RED_LIGHT if margin < 10 else YELLOW_LT)
            for c in range(1, ncol + 1):
                cell = ws3.cell(row=r, column=c); cell.border = _border()
                cell.alignment = _left() if c in (1, 18) else _center() if c > 2 else _left()
            for k, v in (("qty", qty), ("revenue", a["revenue"]), ("commission", a["commission"]),
                         ("logistics", a["logistics"]), ("ret_logistics", ret_log),
                         ("ret_cnt", rcnt), ("other", a["other"]),
                         ("total", a["total"]), ("profit", profit)):
                tot[k] += v
            r += 1

        # ── Продажи, которых ещё нет в выгрузке начислений (из API реализации) ──
        # Выгрузка из кабинета всегда отстаёт: её качают руками. Реализация же
        # тянется автоматически, поэтому артикулы, проданные ПОСЛЕ даты выгрузки,
        # показываем отдельными строками. Услуг (логистика, эквайринг, промо) в
        # реализации нет — они остаются пустыми, и прибыль по таким строкам
        # ЗАВЫШЕНА: это оценка сверху, а не факт. Мешать два источника в одной
        # строке нельзя — задвоились бы комиссии.
        real_by_offer = (realization or {}).get("by_offer", {})
        api_only = {o: v for o, v in real_by_offer.items()
                    if o not in accruals and (v.get("qty") or 0) > 0}
        for oid, v in sorted(api_only.items(), key=lambda x: -x[1]["revenue"]):
            cost = get_cost_price(oid)
            if cost is None:
                missing_cost.append(oid)
            cost = cost or 0
            qty = v["qty"]
            net_qty = max(0, qty - v.get("ret_qty", 0))
            tax = round(v["revenue"] * TAX_RATE, 2)
            profit = round(v["payout"] - cost * net_qty - tax, 2)
            u = qty or 1
            ws3.cell(row=r, column=1, value=(v["name"] or name_of(oid))[:55])
            ws3.cell(row=r, column=2, value=oid)
            ws3.cell(row=r, column=3, value="API")
            ws3.cell(row=r, column=4, value=int(qty))
            ws3.cell(row=r, column=5, value=int(v["ret_qty"]) or None)
            ws3.cell(row=r, column=6, value=round(v["revenue"] / u, 2))
            ws3.cell(row=r, column=7, value=round(v["commission"] / u, 2))
            ws3.cell(row=r, column=8, value=round(v["commission"] / v["revenue"] * 100, 1)
                     if v["revenue"] else 0)
            ws3.cell(row=r, column=9, value="—")          # логистики в реализации нет
            ws3.cell(row=r, column=10, value=cost)
            puc = ws3.cell(row=r, column=11, value=round(profit / u, 2))
            ws3.cell(row=r, column=12, value="—")
            ws3.cell(row=r, column=13, value=round(v["revenue"], 2))
            ws3.cell(row=r, column=14, value=round(v["commission"], 2))
            ws3.cell(row=r, column=15, value="—")
            ws3.cell(row=r, column=17, value=round(v["stars"], 2) or None)
            ws3.cell(row=r, column=18, value="только продажи из API — услуги не учтены")
            ws3.cell(row=r, column=19, value=round(v["payout"], 2))
            pc = ws3.cell(row=r, column=20, value=profit)
            _fill_ym_compare(ws3, r, oid, round(profit / u, 2), ym_blend, COL_YM, COL_YM_DIFF)
            for cc_ in (puc, pc):
                cc_.font = _tfont(bold=True, color="1B5E20" if profit > 0 else "B71C1C")
            for c in range(1, ncol + 1):
                cell = ws3.cell(row=r, column=c); cell.border = _border()
                cell.fill = _fill(YELLOW_LT)      # жёлтым — данные неполные
                cell.alignment = _left() if c in (1, 18) else _center() if c > 2 else _left()
            r += 1

        # Итоговая строка (только в части «за период»)
        ws3.cell(row=r, column=1, value="ИТОГО").font = _tfont(bold=True)
        ws3.cell(row=r, column=4, value=int(tot["qty"]))
        ws3.cell(row=r, column=5, value=int(tot["ret_cnt"]) or None)
        for col, k in ((13, "revenue"), (14, "commission"), (15, "logistics"),
                       (16, "ret_logistics"), (17, "other"), (19, "total"), (20, "profit")):
            ws3.cell(row=r, column=col, value=round(tot[k], 2)).font = _tfont(bold=True)
        for c in range(1, ncol + 1):
            cell = ws3.cell(row=r, column=c); cell.border = _border()
            cell.fill = _fill(GREY_LT); cell.alignment = _center() if c > 2 else _left()
        last = r
    else:
        # ── Fallback: только расчёт на единицу из finance API (выгрузки нет) ──
        heads = ["Товар", "offer_id", "Канал", "Доставл.\nшт.", "Выручка\n₽/ед.",
                 "Комиссия\nOzon ₽", "Логистика\n+услуги ₽", "Выплата\n₽/ед.",
                 "Себест.\n₽", "Налог\n7%", "Чистая\nприбыль/ед.", "Маржа\n%"]
        ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(heads))
        ws3["A1"] = ("Расчёт на единицу из finance API (веб-выгрузка начислений не найдена — "
                     "скачай её для факта за период с разбивкой логистики).")
        ws3["A1"].font = _tfont(sz=10, color="B45309"); ws3["A1"].alignment = _left()
        hr = 2
        for c, h in enumerate(heads, 1):
            ws3.cell(row=hr, column=c, value=h)
        _style_header_row(ws3, hr, len(heads))
        ws3.freeze_panes = "C3"
        ws3.column_dimensions["A"].width = 38
        ws3.column_dimensions["B"].width = 14
        for c in range(3, len(heads) + 1):
            ws3.column_dimensions[get_column_letter(c)].width = 11

        agg = defaultdict(lambda: {"units": 0, "revenue": 0.0, "commission": 0.0,
                                   "services": 0.0, "net": 0.0})
        for s in unit_econ["sales"]:
            a = agg[(s["offer_id"], s["channel"])]
            a["units"] += s["units"]; a["revenue"] += s["revenue"]
            a["commission"] += s["commission"]; a["services"] += s["services_total"]
            a["net"] += s["net"]
        r = hr + 1
        for (oid, ch), a in sorted(agg.items()):
            u = a["units"] or 1
            rev_u = a["revenue"] / u
            cost = get_cost_price(oid)
            if cost is None and not oid.startswith("sku:"):
                missing_cost.append(oid)
            cost = cost or 0
            tax_u = round(rev_u * TAX_RATE, 2)
            profit_u = round(a["net"] / u - cost - tax_u, 2)
            margin = round(profit_u / rev_u * 100, 1) if rev_u else 0
            ws3.cell(row=r, column=1, value=name_of(oid))
            ws3.cell(row=r, column=2, value=oid)
            ws3.cell(row=r, column=3, value=ch)
            ws3.cell(row=r, column=4, value=a["units"])
            ws3.cell(row=r, column=5, value=round(rev_u, 2))
            ws3.cell(row=r, column=6, value=round(a["commission"] / u, 2))
            ws3.cell(row=r, column=7, value=round(a["services"] / u, 2))
            ws3.cell(row=r, column=8, value=round(a["net"] / u, 2))
            ws3.cell(row=r, column=9, value=cost)
            ws3.cell(row=r, column=10, value=tax_u)
            pc = ws3.cell(row=r, column=11, value=profit_u)
            mc = ws3.cell(row=r, column=12, value=margin)
            pc.font = _tfont(bold=True, color="1B5E20" if profit_u > 0 else "B71C1C")
            mc.fill = _fill(GREEN_LT if margin >= 25 else RED_LIGHT if margin < 10 else YELLOW_LT)
            for c in range(1, len(heads) + 1):
                cell = ws3.cell(row=r, column=c); cell.border = _border()
                cell.alignment = _center() if c > 2 else _left()
            r += 1
        last = r - 1

    if missing_cost:
        ws3.cell(row=last + 2, column=1,
                 value=f"⚠ Себестоимость не найдена: {', '.join(sorted(set(missing_cost)))}")
        ws3.cell(row=last + 2, column=1).font = _tfont(bold=True, color="B71C1C")
        ws3.merge_cells(start_row=last + 2, start_column=1, end_row=last + 2, end_column=len(heads))
    ws3.auto_filter.ref = f"A{hr}:{get_column_letter(len(heads))}{last}"

    # ═══════════════════════════════════════════════════════════════════════════
    # SHEET: ПРОМО «ЗВЁЗДНЫЕ ТОВАРЫ» — маржа с платным промо vs без
    # ═══════════════════════════════════════════════════════════════════════════
    if accruals:
        PROMO_KEYS = ("Звёзд", "Трафарет", "родвиж", "еклам", "буст", "Буст", "Премиум", "Premium")
        promo_rows = []
        for oid, a in accruals.items():
            fee = sum(v for n, v in a.get("other_detail", {}).items()
                      if any(k in n for k in PROMO_KEYS))  # обычно отрицательная
            if abs(fee) < 0.01:
                continue
            cost = get_cost_price(oid) or 0
            qty = a["qty"] or 0
            tax = round(a["revenue"] * TAX_RATE, 2)
            # как и на «Юнит-экономике» — себестоимость только на чистые штуки
            profit_with = round(a["total"] - cost * max(0, qty - (a.get("ret_qty") or 0)) - tax, 2)
            profit_without = round(profit_with - fee, 2)  # без расхода на промо (fee<0 → прибавляем)
            m_with = round(profit_with / a["revenue"] * 100, 1) if a["revenue"] else 0
            m_without = round(profit_without / a["revenue"] * 100, 1) if a["revenue"] else 0
            promo_rows.append((oid, a["name"], qty, a["revenue"], round(fee, 2),
                               profit_with, profit_without, m_with, m_without))

        wsp = wb.create_sheet("Промо Звёздные товары")
        wsp.sheet_view.showGridLines = False
        heads = ["Товар", "offer_id", "Доставл.\nшт.", "Выручка\n₽",
                 "Расход на\nпромо ₽", "Прибыль\nс промо ₽", "Прибыль\nбез промо ₽",
                 "Маржа\nс промо %", "Маржа\nбез промо %", "Разница\nпп"]
        wsp.merge_cells("A1:J1")
        wsp["A1"] = ("«Звёздные товары» — платное промо Ozon. Показано, насколько комиссия за промо "
                     "снижает прибыль/маржу. ВАЖНО: прирост продаж от промо по данным начислений "
                     "не виден — оценивайте «выгодно ли» вместе с динамикой продаж на листе «Продажи по дням».")
        wsp["A1"].font = _tfont(sz=10, color="555555"); wsp["A1"].alignment = _left()
        wsp.row_dimensions[1].height = 42
        hr = 2
        for c, h in enumerate(heads, 1):
            wsp.cell(row=hr, column=c, value=h)
        _style_header_row(wsp, hr, len(heads), ORANGE_BG)
        wsp.freeze_panes = "A3"
        wsp.column_dimensions["A"].width = 40
        wsp.column_dimensions["B"].width = 14
        for c in "CDEFGHIJ":
            wsp.column_dimensions[c].width = 12
        r = hr + 1
        t_fee = t_pw = t_po = t_rev = 0.0
        for oid, name, qty, rev, fee, pw, po, mw, mo in sorted(promo_rows, key=lambda x: x[4]):
            wsp.cell(row=r, column=1, value=(name or name_of(oid))[:55])
            wsp.cell(row=r, column=2, value=oid)
            wsp.cell(row=r, column=3, value=qty)
            wsp.cell(row=r, column=4, value=rev)
            fc = wsp.cell(row=r, column=5, value=fee); fc.font = _tfont(color="E65100")
            pwc = wsp.cell(row=r, column=6, value=pw)
            poc = wsp.cell(row=r, column=7, value=po)
            pwc.font = _tfont(bold=True, color="1B5E20" if pw > 0 else "B71C1C")
            poc.font = _tfont(color="1B5E20" if po > 0 else "B71C1C")
            wsp.cell(row=r, column=8, value=mw)
            wsp.cell(row=r, column=9, value=mo)
            dc = wsp.cell(row=r, column=10, value=round(mw - mo, 1))
            dc.fill = _fill(RED_LIGHT)  # промо всегда снижает маржу
            for c in range(1, len(heads) + 1):
                cell = wsp.cell(row=r, column=c); cell.border = _border()
                cell.alignment = _center() if c > 2 else _left()
            t_fee += fee; t_pw += pw; t_po += po; t_rev += rev
            r += 1
        if promo_rows:
            wsp.cell(row=r, column=1, value="ИТОГО").font = _tfont(bold=True)
            wsp.cell(row=r, column=5, value=round(t_fee, 2)).font = _tfont(bold=True, color="E65100")
            wsp.cell(row=r, column=6, value=round(t_pw, 2)).font = _tfont(bold=True)
            wsp.cell(row=r, column=7, value=round(t_po, 2)).font = _tfont(bold=True)
            wsp.cell(row=r, column=8, value=round(t_pw / t_rev * 100, 1) if t_rev else 0).font = _tfont(bold=True)
            wsp.cell(row=r, column=9, value=round(t_po / t_rev * 100, 1) if t_rev else 0).font = _tfont(bold=True)
            for c in range(1, len(heads) + 1):
                cell = wsp.cell(row=r, column=c); cell.border = _border()
                cell.fill = _fill(GREY_LT); cell.alignment = _center() if c > 2 else _left()
        else:
            wsp.cell(row=hr + 1, column=1, value="За период расходов на «Звёздные товары» нет").font = _tfont(color="999999")

    # ═══════════════════════════════════════════════════════════════════════════
    # SHEET 4: ОСТАТКИ FBO/FBS + ПРОГНОЗ
    # ═══════════════════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet("Остатки FBO FBS")
    ws4.sheet_view.showGridLines = False
    heads = ["Товар", "offer_id", "Ост.\nFBO", "Ост.\nFBS", "Резерв",
             f"Прод./день\n({ndays}д)", "Хватит FBO\nдн.", "Завезти FBO\n~шт.", "Статус"]
    for c, h in enumerate(heads, 1):
        ws4.cell(row=1, column=c, value=h)
    _style_header_row(ws4, 1, len(heads))
    ws4.freeze_panes = "A2"
    ws4.column_dimensions["A"].width = 40
    ws4.column_dimensions["B"].width = 14
    for c in "CDEFGHI":
        ws4.column_dimensions[c].width = 12
    r = 2
    for x in restock:
        oid = x["offer_id"]
        if oid.startswith("sku:"):
            continue
        ws4.cell(row=r, column=1, value=name_of(oid))
        ws4.cell(row=r, column=2, value=oid)
        ws4.cell(row=r, column=3, value=x["fbo"])
        ws4.cell(row=r, column=4, value=x["fbs"])
        ws4.cell(row=r, column=5, value=x["reserved"])
        ws4.cell(row=r, column=6, value=x["avg_daily"])
        ws4.cell(row=r, column=7, value=x["days_left"] if x["days_left"] < 900 else "∞")
        ws4.cell(row=r, column=8, value=x["restock_qty"])
        sc = ws4.cell(row=r, column=9, value=x["urgency"])
        if "СРОЧНО" in x["urgency"]:
            sc.fill = _fill(RED_LIGHT)
        elif "Скоро" in x["urgency"]:
            sc.fill = _fill(YELLOW_LT)
        elif "OK" in x["urgency"]:
            sc.fill = _fill(GREEN_LT)
        for c in range(1, len(heads) + 1):
            cell = ws4.cell(row=r, column=c); cell.border = _border()
            cell.alignment = _center() if c > 2 else _left()
        r += 1
    ws4.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{r-1}"

    # ═══════════════════════════════════════════════════════════════════════════
    # SHEET 5: ПРОДАЖИ ПО ДНЯМ (матрица offer × день, заказано шт.)
    # ═══════════════════════════════════════════════════════════════════════════
    ws5 = wb.create_sheet(f"Продажи по дням {ndays}д")
    ws5.sheet_view.showGridLines = False
    heads = ["Товар", "offer_id"] + [d[5:] for d in all_days] + ["Итого\nшт."]
    for c, h in enumerate(heads, 1):
        ws5.cell(row=1, column=c, value=h)
    _style_header_row(ws5, 1, len(heads))
    ws5.freeze_panes = "C2"
    ws5.column_dimensions["A"].width = 36
    ws5.column_dimensions["B"].width = 14
    for c in range(3, len(heads) + 1):
        ws5.column_dimensions[get_column_letter(c)].width = 7

    # сортировка по сумме за период
    matrix = []
    for oid in sales_by_day:
        if oid.startswith("sku:"):
            continue
        tot = sum(sales_by_day[oid].get(d, {}).get("units", 0) for d in all_days)
        if tot > 0:
            matrix.append((oid, tot))
    matrix.sort(key=lambda x: -x[1])
    r = 2
    for oid, tot in matrix:
        ws5.cell(row=r, column=1, value=name_of(oid))
        ws5.cell(row=r, column=2, value=oid)
        for ci, d in enumerate(all_days):
            u = sales_by_day[oid].get(d, {}).get("units", 0)
            cell = ws5.cell(row=r, column=3 + ci, value=int(u) if u else None)
            if u:
                cell.fill = _fill(GREEN_LT)
            cell.alignment = _center(); cell.border = _border()
        tc = ws5.cell(row=r, column=len(heads), value=int(tot))
        tc.font = _tfont(bold=True); tc.alignment = _center(); tc.border = _border()
        ws5.cell(row=r, column=1).border = _border()
        ws5.cell(row=r, column=2).border = _border()
        r += 1

    # ═══════════════════════════════════════════════════════════════════════════
    # SHEET 6: ЛОГИСТИКА ПО РЕГИОНАМ (кластеры)
    # ═══════════════════════════════════════════════════════════════════════════
    ws6 = wb.create_sheet("Логистика по регионам")
    ws6.sheet_view.showGridLines = False
    ws6.merge_cells("A1:E1")
    ws6["A1"] = "Кластеры доставки Ozon (cluster_to) по отправлениям за период"
    ws6["A1"].font = _tfont(sz=10, color="555555"); ws6["A1"].alignment = _left()
    heads = ["Кластер назначения", "Отправл.\nвсего", "FBS", "FBO", "Выручка\n₽"]
    for c, h in enumerate(heads, 1):
        ws6.cell(row=2, column=c, value=h)
    _style_header_row(ws6, 2, len(heads))
    ws6.column_dimensions["A"].width = 34
    for c in "BCDE":
        ws6.column_dimensions[c].width = 13
    r = 3
    for cluster, v in sorted(clusters.items(), key=lambda x: -x[1]["count"]):
        ws6.cell(row=r, column=1, value=cluster)
        ws6.cell(row=r, column=2, value=v["count"])
        ws6.cell(row=r, column=3, value=v["channel"].get("FBS", 0))
        ws6.cell(row=r, column=4, value=v["channel"].get("FBO", 0))
        ws6.cell(row=r, column=5, value=round(v["revenue"], 2))
        for c in range(1, len(heads) + 1):
            cell = ws6.cell(row=r, column=c); cell.border = _border()
            cell.alignment = _center() if c > 1 else _left()
        r += 1

    # ═══════════════════════════════════════════════════════════════════════════
    # SHEET 7: ВОЗВРАТЫ — детальные из веб-выгрузки, иначе сводка из finance API
    # ═══════════════════════════════════════════════════════════════════════════
    ws7 = wb.create_sheet("Возвраты")
    ws7.sheet_view.showGridLines = False
    if returns_detail:
        ws7.merge_cells("A1:H1")
        ws7["A1"] = "Детальные возвраты из веб-выгрузки «Возвраты FBS и FBO»"
        ws7["A1"].font = _tfont(sz=10, color="555555"); ws7["A1"].alignment = _left()
        heads = ["Дата\nвозврата", "Схема", "offer_id", "Товар", "Причина",
                 "Статус\nвозврата", "Статус\nкомпенс.", "Кол-во"]
        hr = 2
        for c, h in enumerate(heads, 1):
            ws7.cell(row=hr, column=c, value=h)
        _style_header_row(ws7, hr, len(heads))
        ws7.freeze_panes = "A3"
        ws7.column_dimensions["A"].width = 12
        ws7.column_dimensions["B"].width = 7
        ws7.column_dimensions["C"].width = 13
        ws7.column_dimensions["D"].width = 34
        ws7.column_dimensions["E"].width = 26
        ws7.column_dimensions["F"].width = 16
        ws7.column_dimensions["G"].width = 16
        ws7.column_dimensions["H"].width = 8
        r = hr + 1
        for x in sorted(returns_detail, key=lambda z: z["return_date"], reverse=True):
            ws7.cell(row=r, column=1, value=x["return_date"])
            ws7.cell(row=r, column=2, value=x["channel"])
            ws7.cell(row=r, column=3, value=x["offer_id"])
            ws7.cell(row=r, column=4, value=x["name"][:50])
            ws7.cell(row=r, column=5, value=x["reason"])
            ws7.cell(row=r, column=6, value=x["status"])
            ws7.cell(row=r, column=7, value=x["comp_status"])
            ws7.cell(row=r, column=8, value=x["qty"])
            for c in range(1, len(heads) + 1):
                cell = ws7.cell(row=r, column=c); cell.border = _border()
                cell.alignment = _center() if c in (1, 2, 8) else _left()
            r += 1
        ws7.auto_filter.ref = f"A{hr}:{get_column_letter(len(heads))}{r-1}"
    else:
        heads = ["Товар", "offer_id", "Возвратов\nшт.", "Сумма\n₽"]
        for c, h in enumerate(heads, 1):
            ws7.cell(row=1, column=c, value=h)
        _style_header_row(ws7, 1, len(heads))
        ws7.column_dimensions["A"].width = 40
        ws7.column_dimensions["B"].width = 14
        ws7.column_dimensions["C"].width = 12
        ws7.column_dimensions["D"].width = 12
        r = 2
        rets = sorted(unit_econ["returns"].items(), key=lambda x: -x[1]["count"])
        if not rets:
            ws7.cell(row=2, column=1, value="Возвратов за период нет").font = _tfont(color="999999")
        for oid, v in rets:
            ws7.cell(row=r, column=1, value=name_of(oid))
            ws7.cell(row=r, column=2, value=oid)
            ws7.cell(row=r, column=3, value=v["count"])
            ws7.cell(row=r, column=4, value=round(v["amount"], 2))
            for c in range(1, len(heads) + 1):
                cell = ws7.cell(row=r, column=c); cell.border = _border()
                cell.alignment = _center() if c > 2 else _left()
            r += 1

    # ═══════════════════════════════════════════════════════════════════════════
    # SHEET 8: ЗДОРОВЬЕ КАРТОЧЕК (из веб-выгрузки «Все товары и видимость»)
    # ═══════════════════════════════════════════════════════════════════════════
    if products_health:
        ws8 = wb.create_sheet("Здоровье карточек")
        ws8.sheet_view.showGridLines = False
        ws8.merge_cells("A1:I1")
        ws8["A1"] = ("Из веб-выгрузки «Все товары и видимость на Ozon» (Аналитика → Отчёты). "
                     "Воронки показов/CTR в выгрузках Ozon нет — только статус карточки и контент-рейтинг.")
        ws8["A1"].font = _tfont(sz=10, color="555555"); ws8["A1"].alignment = _left()
        heads = ["Товар", "offer_id", "Статус", "Причина\nскрытия", "Контент-\nрейтинг",
                 "Отзывы", "Рейтинг", "Ост.\nFBO+FBS", "Ошибки/предупр."]
        hr = 2
        for c, h in enumerate(heads, 1):
            ws8.cell(row=hr, column=c, value=h)
        _style_header_row(ws8, hr, len(heads))
        ws8.freeze_panes = "A3"
        ws8.column_dimensions["A"].width = 40
        ws8.column_dimensions["B"].width = 14
        ws8.column_dimensions["C"].width = 15
        ws8.column_dimensions["D"].width = 22
        for c in "EFGH":
            ws8.column_dimensions[c].width = 11
        ws8.column_dimensions["I"].width = 28
        # сортировка: сначала «не продаётся», потом по контент-рейтингу (худшие вверх)
        ph = sorted(products_health,
                    key=lambda x: (x["status"] == "Готов к продаже", x["content_rating"]))
        r = hr + 1
        for x in ph:
            ws8.cell(row=r, column=1, value=x["name"][:55])
            ws8.cell(row=r, column=2, value=x["offer_id"])
            sc = ws8.cell(row=r, column=3, value=x["status"])
            ws8.cell(row=r, column=4, value=x["reasons"])
            cr = ws8.cell(row=r, column=5, value=x["content_rating"])
            ws8.cell(row=r, column=6, value=x["reviews"])
            ws8.cell(row=r, column=7, value=x["rating"] or None)
            ws8.cell(row=r, column=8, value=x["fbo"] + x["fbs"] + x["realfbs"])
            ws8.cell(row=r, column=9, value=(x["errors"] or x["warnings"])[:40])
            if x["status"] == "Не продается":
                sc.fill = _fill(RED_LIGHT); sc.font = _tfont(bold=True, color="B71C1C")
            elif x["status"] == "Продается":
                sc.fill = _fill(GREEN_LT)
            if x["content_rating"] and x["content_rating"] < 70:
                cr.fill = _fill(YELLOW_LT)
            elif x["content_rating"] >= 90:
                cr.fill = _fill(GREEN_LT)
            for c in range(1, len(heads) + 1):
                cell = ws8.cell(row=r, column=c); cell.border = _border()
                cell.alignment = _center() if c not in (1, 4, 9) else _left()
            r += 1
        ws8.auto_filter.ref = f"A{hr}:{get_column_letter(len(heads))}{r-1}"

    # ═══════════════════════════════════════════════════════════════════════════
    # ВЫКУПЫ / ФАКТ-ПРИБЫЛЬ (из finance/transaction — реальные деньги по выкупам)
    # ═══════════════════════════════════════════════════════════════════════════
    def _agg_econ(ue):
        agg = {}
        for s in ue.get("sales", []):
            oid = s["offer_id"]
            a = agg.setdefault(oid, {"name": s.get("name", oid), "units": 0.0, "revenue": 0.0,
                                     "commission": 0.0, "services": 0.0, "net": 0.0,
                                     "ret_qty": 0.0, "ret_amount": 0.0})
            a["units"] += s.get("units", 0); a["revenue"] += s.get("revenue", 0)
            a["commission"] += s.get("commission", 0); a["services"] += s.get("services_total", 0)
            a["net"] += s.get("net", 0)
        for oid, rr in ue.get("returns", {}).items():
            a = agg.setdefault(oid, {"name": oid, "units": 0.0, "revenue": 0.0, "commission": 0.0,
                                     "services": 0.0, "net": 0.0, "ret_qty": 0.0, "ret_amount": 0.0})
            a["ret_qty"] += rr.get("count", 0); a["ret_amount"] += rr.get("amount", 0)
        return agg

    def _econ_profit(oid, a):
        """Факт-прибыль по выкупам. Возвраты годные → себест на чистые штуки, налог с нетто."""
        units = a["units"]; retq = a["ret_qty"]; net_u = max(units - retq, 0)
        payout = a["net"] + a["ret_amount"]                # выплата (выкупы − возвраты)
        cost = (get_cost_price(oid) or 0) * net_u
        ret_rev = retq * (a["revenue"] / units) if units > 0 else 0
        tax = round((a["revenue"] - ret_rev) * TAX_RATE, 2)
        profit = round(payout - cost - tax, 2)
        ppu = round(profit / net_u, 2) if net_u > 0 else 0
        margin = round(profit / a["revenue"] * 100, 1) if a["revenue"] else 0
        return net_u, round(payout, 2), round(cost, 2), tax, profit, ppu, margin

    agg14 = _agg_econ(unit_econ)
    if agg14:
        wsb = wb.create_sheet("Выкупы факт-прибыль")
        wsb.sheet_view.showGridLines = False
        heads = ["Товар", "offer_id", f"Выкуплено\nшт ({ndays}д)", "Возвр.\nшт",
                 "Выручка\n₽", "Комиссия\n₽", "Услуги\n₽", "Выплата\n₽",
                 "Себест.\n₽", "Налог\n₽", f"Факт-приб.\n₽ ({ndays}д)",
                 f"₽/шт\n{ndays}д", "₽/шт\n90д", "Маржа\n%"]
        nc = len(heads)
        note = (f"Реальная факт-прибыль по ВЫКУПЛЕННЫМ товарам (finance/transaction, "
                f"OperationAgentDeliveredToCustomer) за {date_from:%d.%m}–{date_to:%d.%m.%Y}. "
                f"Выплата = выкупы − возвраты (реальные деньги). Факт-прибыль = Выплата − Себест "
                f"(на ЧИСТЫЕ штуки) − Налог {TAX_RATE*100:.0f}% (с нетто-выручки). Возвраты годные → "
                f"себест/налог на них не списываем. «₽/шт 90д» — устойчивая квартальная прибыль на "
                f"штуку (без скачков от лага выплат); дневная ₽/шт скачет, ориентируйся на 90д.")
        wsb.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
        wsb.cell(row=1, column=1, value=note).font = _tfont(sz=10, color="555555")
        wsb.cell(row=1, column=1).alignment = _left()
        wsb.row_dimensions[1].height = 60
        for c, h in enumerate(heads, 1):
            wsb.cell(row=2, column=c, value=h)
        _style_header_row(wsb, 2, nc, TEAL)
        wsb.row_dimensions[2].height = 40

        q_by = econ_quarter  # {offer_id: agg}
        rows_data = []
        for oid, a in agg14.items():
            net_u, payout, cost, tax, profit, ppu, margin = _econ_profit(oid, a)
            ppu90 = ""
            qa = q_by.get(oid)
            if qa:
                _, _, _, _, _, ppu90, _ = _econ_profit(oid, qa)
            rows_data.append((oid, a, net_u, payout, cost, tax, profit, ppu, ppu90, margin))
        rows_data.sort(key=lambda x: -x[6])   # по факт-прибыли ↓
        r = 3
        for oid, a, net_u, payout, cost, tax, profit, ppu, ppu90, margin in rows_data:
            vals = [a["name"][:45], oid, int(a["units"]), int(a["ret_qty"]),
                    round(a["revenue"]), round(a["commission"]), round(a["services"]),
                    payout, cost, tax, profit, ppu, ppu90 if ppu90 != "" else "—", margin]
            for c, v in enumerate(vals, 1):
                cell = wsb.cell(row=r, column=c, value=v)
                cell.border = _border()
                cell.alignment = _left() if c == 1 else _center()
                if c in (5, 6, 7, 8, 9, 10, 11) and isinstance(v, (int, float)):
                    cell.number_format = '#,##0'
                if c in (12, 13) and isinstance(v, (int, float)):
                    cell.number_format = '#,##0.0'
            # цвет прибыли и ₽/шт
            for col, val in ((11, profit), (12, ppu), (13, ppu90 if ppu90 != "" else 0)):
                pc = wsb.cell(row=r, column=col)
                pc.font = _tfont(bold=True, color="1B5E20" if (isinstance(val, (int, float)) and val > 0) else "B71C1C")
            mc = wsb.cell(row=r, column=14)
            mc.font = _tfont(bold=True, color="1B5E20" if margin >= 20 else "B71C1C" if margin < 5 else "9A7D0A")
            if (get_cost_price(oid) or 0) == 0:
                wsb.cell(row=r, column=9).font = _tfont(bold=True, color="E65100")
            r += 1
        # ширины
        wsb.column_dimensions["A"].width = 40; wsb.column_dimensions["B"].width = 14
        for i in range(3, nc + 1):
            wsb.column_dimensions[get_column_letter(i)].width = 10
        wsb.freeze_panes = "C3"
        wsb.auto_filter.ref = f"A2:{get_column_letter(nc)}{r-1}"

    # ═══════════════════════════════════════════════════════════════════════════
    # ПОРЯДОК ЛИСТОВ
    # ═══════════════════════════════════════════════════════════════════════════
    desired_order = [
        "Дашборд", "Юнит-экономика", "Выкупы факт-прибыль", "Возвраты",
        f"Продажи по дням {ndays}д", "Остатки FBO FBS",
    ]
    ordered = [wb[n] for n in desired_order if n in wb.sheetnames]
    ordered += [wb[n] for n in wb.sheetnames if n not in desired_order]
    wb._sheets = ordered

    # ── SAVE ──
    path = os.path.join(report_folder, f"ozon_report_{run_date.isoformat()}.xlsx")
    wb.save(path)
    return path


def _restock_table(stocks, sales_tot, ndays):
    """Прогноз поставки FBO по средней скорости продаж за период."""
    rows = []
    all_oids = set(stocks) | set(sales_tot)
    for oid in all_oids:
        st = stocks.get(oid, {})
        fbo = st.get("fbo", 0); fbs = st.get("fbs", 0)
        reserved = st.get("reserved_fbo", 0) + st.get("reserved_fbs", 0)
        sold = sales_tot.get(oid, {}).get("units", 0)
        avg_daily = sold / ndays if ndays else 0
        if avg_daily > 0:
            days_left = round(fbo / avg_daily, 1)
        elif fbo > 0:
            days_left = 999
        else:
            days_left = 0
        restock_qty = max(0, round(avg_daily * 30 - fbo))
        if avg_daily == 0 and fbo == 0:
            urgency = "—"
        elif days_left <= RESTOCK_URGENT_DAYS:
            urgency = "🔴 СРОЧНО"
        elif days_left <= RESTOCK_SOON_DAYS:
            urgency = "🟡 Скоро"
        elif days_left <= RESTOCK_PLAN_DAYS:
            urgency = "🔵 Планировать"
        else:
            urgency = "🟢 OK"
        rows.append({"offer_id": oid, "fbo": fbo, "fbs": fbs, "reserved": reserved,
                     "avg_daily": round(avg_daily, 2), "days_left": days_left,
                     "restock_qty": restock_qty, "urgency": urgency})
    # Сначала живые позиции (по срочности), мёртвые («—») — в самый низ
    rows.sort(key=lambda x: (x["urgency"] == "—", x["days_left"]))
    return rows
