"""
Парсеры Excel/CSV-выгрузок из веб-кабинета Ozon (ручной режим).

Воронка показов/CTR через API недоступна (нет Premium), а в кабинете она
отдаётся canvas-виджетами без выгрузки. Зато кнопки «Скачать отчёт» в
Аналитика → Отчёты дают CSV/Excel, которые здесь и разбираются.

Файлы пользователь скачивает вручную (Chrome залогинен) в папку загрузок.
"""

import csv
import os
import re
import glob
from collections import defaultdict

import openpyxl


def find_latest(folder: str, patterns: list[str]) -> str | None:
    """Самый свежий файл из folder, подходящий под любой glob-паттерн."""
    cands = []
    for p in patterns:
        cands += glob.glob(os.path.join(folder, p))
    return max(cands, key=os.path.getmtime) if cands else None


def _num(s):
    if s is None:
        return 0
    s = str(s).strip().strip("'").replace(",", ".").replace("%", "")
    try:
        return float(s)
    except ValueError:
        return 0


def parse_products_health(path: str) -> list[dict]:
    """
    Выгрузка «Все товары и видимость на Ozon» (products.csv, разделитель ';').

    Returns list of dict per товар:
      offer_id, name, content_rating, status, reasons, reviews, rating,
      fbo, fbs, realfbs, errors, warnings
    """
    if not path or not os.path.exists(path):
        return []
    rows = []
    with open(path, encoding="utf-8-sig") as f:
        rd = csv.DictReader(f, delimiter=";")
        for r in rd:
            rows.append({
                "offer_id":       (r.get("Артикул") or "").strip().strip("'"),
                "name":           (r.get("Название товара") or "").strip(),
                "content_rating": _num(r.get("Контент-рейтинг")),
                "status":         (r.get("Статус товара") or "").strip(),
                "reasons":        (r.get("Причины скрытия") or "").strip(),
                "reviews":        int(_num(r.get("Отзывы"))),
                "rating":         _num(r.get("Рейтинг")),
                "fbo":            int(_num(r.get("Доступно к продаже по схеме FBO, шт."))),
                "fbs":            int(_num(r.get("Доступно к продаже по схеме FBS, шт."))),
                "realfbs":        int(_num(r.get("Доступно к продаже по схеме realFBS, шт."))),
                "errors":         (r.get("Ошибки") or "").strip(),
                "warnings":       (r.get("Предупреждения") or "").strip(),
            })
    return rows


# ── Классификация строк «длинного» отчёта по начислениям ──────────────────────
# Ozon отдаёт группу услуг + тип начисления; раскладываем их по тем же корзинам,
# что и старый «широкий» отчёт, чтобы отчёт-генератор не заметил разницы.
_REV_TYPES = ("Выручка", "Возврат выручки", "Баллы за скидки")
_LOG_GROUPS = ("Услуги доставки", "Услуги FBO")
# «Услуги партнёров» — смешанная группа: часть логистика, часть эквайринг/промо
_LOG_PARTNER_HINTS = ("Обработка отправления", "Доставка до места выдачи",
                      "Обработка возвратов", "Упаковка товара")


def _classify_accrual(group: str, otype: str) -> str:
    """→ 'revenue' | 'commission' | 'logistics' | 'other'."""
    g = (group or "").strip()
    t = (otype or "").strip()
    if t in _REV_TYPES:
        return "revenue"          # «Баллы за скидки» — это НЕ расход, а часть
                                  # цены, которую Ozon закрыл сам (≈46% выручки)
    if g == "Вознаграждение Ozon":
        return "commission"
    if g in _LOG_GROUPS:
        return "logistics"
    if g == "Услуги партнёров":
        return "logistics" if any(h in t for h in _LOG_PARTNER_HINTS) else "other"
    return "other"


def accruals_period(path: str):
    """Период выгрузки «Отчёт по начислениям» из её первой строки
    («Период: 21.09.2026-13.10.2026») → (date_from, date_to) или None
    (старый широкий формат периода не пишет)."""
    if not path or not os.path.exists(path):
        return None
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    first = next(ws.iter_rows(max_row=1, values_only=True), (None,))[0]
    m = re.search(r"(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})", str(first or ""))
    if not m:
        return None
    from datetime import datetime as _dt
    return tuple(_dt.strptime(x, "%d.%m.%Y").date() for x in m.groups())


def _parse_accruals_long(ws, date_from=None, date_to=None) -> dict:
    """Новый формат Ozon (с ~08.2026): 1 строка = 1 операция, 16 колонок,
    заголовок на 2-й строке (в 1-й — «Период: …»).

    Колонки (0-инд.): 0=ID начисления, 2=Группа услуг, 3=Тип начисления,
    4=Артикул, 6=Название, 7=Количество, 11=Схема работы, 15=Сумма итого.
    """
    agg = defaultdict(lambda: {"name": "", "channel": "", "qty": 0, "ret_qty": 0,
                               "revenue": 0.0,
                               "commission": 0.0, "logistics": 0.0, "ret_logistics": 0.0,
                               "other": 0.0, "points": 0.0, "total": 0.0,
                               "other_detail": defaultdict(float)})
    skipped = 0.0
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 16:
            continue
        # Окно отчёта — по дате начисления (кол. 1). Без фильтра в отчёт за 14 дней
        # попадала бы вся выгрузка (у неё свой период, обычно длиннее).
        if date_from and hasattr(r[1], "date") and not (date_from <= r[1].date() <= date_to):
            continue
        oid = str(r[4]).strip() if r[4] else ""
        amount = _num(r[15])
        if not oid:
            skipped += amount     # услуги без привязки к товару (напр. выплаты)
            continue
        group, otype = (r[2] or ""), (r[3] or "")
        a = agg[oid]
        a["name"] = r[6] or a["name"]
        a["channel"] = r[11] or a["channel"]
        if otype == "Выручка":
            a["qty"] += int(_num(r[7]))          # продажа = строка выручки
        if otype == "Возврат выручки":
            # Вернувшийся товар приезжает годным на склад, поэтому себестоимость
            # по нему списывать нельзя — иначе убыток задваивается: и товар на
            # складе, и деньги «потрачены». Считаем штуки возвратов отдельно,
            # чтобы себестоимость шла только на чистые (выкупленные) штуки.
            a["ret_qty"] += int(_num(r[7]))
        if otype == "Баллы за скидки":
            a["points"] += amount                 # для справки, входит в revenue
        bucket = _classify_accrual(group, otype)
        a[bucket] += amount
        if bucket == "other" and abs(amount) >= 0.01:
            a["other_detail"][otype or "—"] += amount
        if ("озврат" in otype) or ("тмен" in otype) or ("евыкуп" in otype):
            if bucket == "logistics":
                a["ret_logistics"] += amount
        a["total"] += amount
    if abs(skipped) >= 0.01:
        print(f"   ⚠️  начисления без артикула (не разнесены по SKU): {skipped:,.2f} ₽")
    return agg


def parse_accruals_by_product(path: str, date_from=None, date_to=None) -> dict:
    """
    Выгрузка начислений в разрезе товаров. Поддерживает ДВА формата:
      • старый «широкий» — «Отчет по товарам за период …»: 1 строка = отправление,
        расходы разнесены по 25 колонкам;
      • новый «длинный» — «Отчет по начислениям …» (Ozon перешёл ~08.2026):
        1 строка = одна операция, группа услуг + тип + сумма, заголовок на 2-й строке.
    Формат определяется автоматически по заголовку.

    Returns: { offer_id: {name, channel, qty, revenue, commission,
                          logistics, ret_logistics, other, points, total, other_detail} }
      revenue — полная цена продавца (деньги покупателя + баллы Ozon за скидку);
                в старом формате это была одна колонка «За продажу или возврат»,
                в новом Ozon разделил её на «Выручка» и «Баллы за скидки» —
                складываем обратно, иначе выручка занижается почти вдвое.
      points  — только новый формат: сколько из revenue закрыл сам Ozon баллами.
    """
    if not path or not os.path.exists(path):
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    hdr2 = [str(c.value or "") for c in ws[2]]
    if "Группа услуг" in hdr2 and "Сумма итого, руб." in hdr2:
        agg = _parse_accruals_long(ws, date_from, date_to)
        return _finalize_accruals(agg)

    rows = list(ws.iter_rows(min_row=2, values_only=True))
    # колонки (0-индекс): 4=склад, 6=offer_id, 7=name, 8=qty, 9=revenue,
    # 11=commission, 12..20 и 23=логистика, 24=итого
    LOG_COLS = [12, 13, 14, 15, 16, 17, 18, 19, 20, 23]
    agg = defaultdict(lambda: {"name": "", "channel": "", "qty": 0, "ret_qty": 0,
                               "revenue": 0.0,
                               "commission": 0.0, "logistics": 0.0, "ret_logistics": 0.0,
                               "other": 0.0, "total": 0.0, "other_detail": defaultdict(float)})
    for r in rows:
        if not r or len(r) < 25 or not r[6]:
            continue
        oid = str(r[6]).strip()
        a = agg[oid]
        a["name"] = r[7] or a["name"]
        a["channel"] = r[4] or a["channel"]
        otype = (r[1] or "")
        if "Доставка покупателю" in otype:
            a["qty"] += int(_num(r[8]))
        rev = _num(r[9]); comm = _num(r[11]); total = _num(r[24])
        log = sum(_num(r[c]) for c in LOG_COLS)
        a["revenue"] += rev
        a["commission"] += comm
        a["logistics"] += log
        # возвратная/отменная логистика — подмножество logistics (для пояснения)
        if ("возврат" in otype) or ("отмен" in otype) or ("невыкуп" in otype) or ("невостреб" in otype):
            a["ret_logistics"] += log
        a["total"] += total
        contrib = total - rev - comm - log  # вклад строки в «Прочее»
        a["other"] += contrib
        if abs(contrib) >= 0.01:
            a["other_detail"][otype or "—"] += contrib

    return _finalize_accruals(agg)


def _finalize_accruals(agg: dict) -> dict:
    """defaultdict → обычный dict с округлением (общий для обоих форматов)."""
    out = {}
    for k, v in agg.items():
        d = {}
        for kk, vv in v.items():
            if kk == "other_detail":
                d[kk] = {n: round(s, 2) for n, s in vv.items() if abs(s) >= 0.01}
            elif isinstance(vv, float):
                d[kk] = round(vv, 2)
            else:
                d[kk] = vv
        d.setdefault("points", 0.0)   # старый формат баллы отдельно не показывает
        out[k] = d
    return out


def parse_returns(path: str) -> list[dict]:
    """
    Выгрузка «Возвраты FBS и FBO». Заголовок в строке 6, данные с 7-й.

    Returns list of dict: {channel, name, offer_id, posting, order_date,
                           return_date, status, comp_status, reason, qty}
    """
    if not path or not os.path.exists(path):
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    out = []
    for r in ws.iter_rows(min_row=7, values_only=True):
        if not r or not r[0] or not r[3]:
            continue
        out.append({
            "channel":     str(r[0]).strip(),
            "name":        str(r[1] or "").strip(),
            "offer_id":    str(r[3]).strip(),
            "order_date":  str(r[5] or "")[:10],
            "return_date": str(r[6] or "")[:10],
            "status":      str(r[7] or "").strip(),
            "comp_status": str(r[9] or "").strip(),
            "reason":      str(r[12] or "").strip(),
            "qty":         int(_num(r[15])) if str(r[15]).strip().isdigit() else 1,
        })
    return out
