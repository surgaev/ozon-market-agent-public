"""Скриншот листа отчёта Ozon для README — на ВЫДУМАННЫХ данных.

Рисует HTML-макет листа «Юнит-экономика» с теми же
цветами и правилами подсветки, что в настоящем отчёте, и снимает их
безголовым Chrome в PNG. Данные магазина сюда не попадают: артикулы, названия
и цифры сгенерированы ниже с фиксированным seed.

Запуск (macOS): python3 docs/make_screenshots.py
"""
import os
import random
import re
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
rnd = random.Random(7)

GREEN, YELLOW, RED = ("#C6EFCE", "#1B5E20"), ("#FFF9C4", "#7D6608"), ("#FFC7CE", "#B71C1C")


def tier(v):
    return GREEN if v > 100 else YELLOW if v >= 50 else RED


def fmt(v, dec=0):
    if v is None:
        return ""
    s = f"{v:,.{dec}f}".replace(",", " ")
    return s.replace("-", "−")


CSS = """
body{margin:0;background:#fff;font-family:Calibri,'Segoe UI',Arial,sans-serif;font-size:12px;color:#000}
.wrap{padding:10px 12px 0 12px;display:inline-block}
.note{color:#555;font-size:11px;max-width:1100px;white-space:normal;margin:0 0 6px 2px;line-height:1.35}
table{border-collapse:collapse}
td,th{border:1px solid #c9c9c9;padding:3px 6px;white-space:nowrap;height:18px}
th{color:#fff;font-weight:700;text-align:center;white-space:pre-line;line-height:1.15;font-size:11px}
td.l{text-align:left} td{text-align:right}
td.c{text-align:center}
b{font-weight:700}
.tabs{margin-top:8px;border-top:1px solid #bbb;padding:4px 0 6px 2px;font-size:11px;color:#333}
.tabs span{display:inline-block;padding:3px 10px;margin-right:2px;border:1px solid #c9c9c9;border-top:none;background:#f3f3f3}
.tabs span.on{background:#fff;font-weight:700;border-bottom:2px solid #217346;color:#217346}
"""

TABS = ["Дашборд", "Юнит-экономика", "Возвраты", "Продажи по дням 30д", "Остатки FBO FBS",
        "Товары", "Промо Звёздные товары", "Логистика по регионам", "Здоровье карточек"]


def tabs(active):
    return '<div class="tabs">' + "".join(
        f'<span class="{"on" if t == active else ""}">{t}</span>' for t in TABS) + "</div>"


# ── выдуманный каталог: (название, offer_id, канал, цена, себестоимость, доставлено, возвратов, ЯМ ⭐ВЗВЕШ) ──
CATALOG = [
    ("Наушники-вкладыши", "EAR-BUD-W", "FBO", 3290, 1150, 14, 1, None),
    ("Ремешок для часов кожаный", "STRAP-LTH-BK", "FBS", 1590, 310, 22, 0, 139),
    ("Держатель для телефона в авто", "HOLD-CAR-02", "FBO", 1390, 260, 31, 2, 74),
    ("Чехол для телефона силиконовый", "CASE-SIL-01", "FBS", 990, 140, 46, 1, 56),
    ("Перчатки сенсорные, L", "GLV-TCH-L", "FBS", 2290, 420, 6, 0, 38),
    ("Кабель USB-C плетёный 2 м", "CBL-USBC-2M", "FBO", 690, 85, 58, 3, 12),
    ("Подставка для ноутбука", "STAND-LP-AL", "FBS", 2590, 980, 3, 1, -99),
]
API_ROW = ("Чехол для планшета 11\"", "CASE-TAB-11", 1890, 316, 2)   # продажа, которой ещё нет в выгрузке
OTHER = ["Звёздные товары", "Эквайринг", "Программы партнёров"]


def unit_html():
    heads1 = ('<tr><th colspan=5 style="background:#F2F2F2"></th>'
              '<th colspan=7 style="background:#2F4F8F">НА 1 ШТУКУ</th>'
              '<th colspan=10 style="background:#16A085">ЗА ПЕРИОД (все штуки)</th></tr>')
    names = [("Товар", "#16A085"), ("offer_id", "#16A085"), ("Канал", "#16A085"), ("Доставл.\nшт.", "#16A085"),
             ("Возвр.\nшт.", "#16A085")] + [(t, "#2F4F8F") for t in (
             "Выручка\n₽/ед", "Комиссия\n₽/ед", "Комиссия\n%", "Логистика\n₽/ед", "Себест.\n₽/ед",
             "Прибыль\n₽/ед", "Маржа\n%")] + [(t, "#16A085") for t in (
             "Выручка\n₽", "Комиссия\n₽", "Логистика\n₽", "в т.ч. лог.\nвозвратов ₽", "Прочее\n₽",
             "Из чего «Прочее»", "Итого\nвыплата ₽", "Прибыль\nза период ₽", "ЯМ ⭐ВЗВЕШ\n₽/шт", "Δ Озон−ЯМ\n₽/шт")]
    h = [heads1, "<tr>" + "".join(f'<th style="background:{c}">{t}</th>' for t, c in names) + "</tr>"]
    rows, tot = [], dict(q=0, r=0, rev=0, com=0, log=0, rl=0, oth=0, out=0, prof=0)
    for name, oid, ch, price, cost_u, q, ret, ym in CATALOG:
        rev = q * price * rnd.uniform(0.97, 1.0)
        com = -rev * rnd.uniform(0.46, 0.52)
        rl = -ret * (15 if ch == "FBO" else 30)
        log = -q * rnd.uniform(80, 140) * (1.3 if ch == "FBO" else 1) + rl
        parts = [-rev * rnd.uniform(0.008, 0.012), -rev * rnd.uniform(0.006, 0.01), rev * rnd.uniform(0.002, 0.005)]
        oth = sum(parts)
        why = "; ".join(f"{n}: {'+' if v > 0 else '−'}{abs(v):.0f}" for n, v in sorted(zip(OTHER, parts), key=lambda x: x[1]))
        out = rev + com + log + oth
        prof = out - cost_u * (q - ret) - 0.07 * rev
        rows.append((name, oid, ch, q, ret, rev, com, log, rl, oth, why, out, prof, cost_u, ym, False))
        for k, v in zip(tot, (q, ret, rev, com, log, rl, oth, out, prof)):
            tot[k] += v
    rows.sort(key=lambda x: -x[12])
    name, oid, price, cost_u, q = API_ROW
    rev = q * price; com = -rev * 0.47
    prof = rev + com - cost_u * q - 0.07 * rev
    rows.append((name, oid, "API", q, 0, rev, com, None, None, None, "только продажи из API — услуги не учтены",
                 rev + com, prof, cost_u, None, True))
    for k, v in zip(("q", "rev", "com", "out", "prof"), (q, rev, com, rev + com, prof)):
        tot[k] += v

    def td(v, dec=0, style="", cls=""):
        return f'<td class="{cls}" style="{style}">{v if isinstance(v, str) else fmt(v, dec)}</td>'
    for name, oid, ch, q, ret, rev, com, log, rl, oth, why, out, prof, cost_u, ym, api in rows:
        bg = "background:#FFF9C4;" if api else ""
        pu = prof / q
        mrg = prof / rev * 100
        pcol = f'color:{"#1B5E20" if prof > 0 else "#B71C1C"};font-weight:700;'
        mfill = "#C6EFCE" if mrg >= 25 else "#FFC7CE" if mrg < 10 else "#FFF9C4"
        rc = f"background:#FFE0B2;color:#E65100;font-weight:700;" if ret else bg
        cells = (td(name, cls="l", style=bg) + td(oid, cls="l", style=bg) + td(ch, cls="c", style=bg) + td(q, style=bg)
                 + td(ret if ret else "", style=rc) + td(rev / q, 2, bg) + td(com / q, 2, bg) + td(com / rev * 100, 1, bg)
                 + td("—" if log is None else log / q, 2, bg) + td(cost_u, 2, bg) + td(pu, 2, bg + pcol)
                 + td(mrg, 1, ("" if api else f"background:{mfill};")) + td(rev, 0, bg) + td(com, 0, bg)
                 + td("—" if log is None else log, 0, bg)
                 + td(rl if rl else "", 0, bg + ("color:#E65100;" if rl else ""))
                 + td("" if oth is None else oth, 0, bg) + td(why, cls="l", style=bg + "font-size:10.5px;color:#333;white-space:normal;min-width:190px;max-width:190px;line-height:1.2;")
                 + td(out, 0, bg) + td(prof, 0, bg + pcol))
        if ym is None:
            cells += td("—", style=bg + "color:#999;") + td("—", style=bg + "color:#999;")
        else:
            d = pu - ym
            cells += (td(ym, 0, bg + f'color:{"#1B5E20" if ym > 0 else "#B71C1C"};')
                      + td(d, 0, f'background:{"#C6EFCE" if d > 0 else "#FFC7CE"};color:{"#1B5E20" if d > 0 else "#B71C1C"};font-weight:700;'))
        h.append(f"<tr>{cells}</tr>")
    g = "background:#F2F2F2;font-weight:700;"
    h.append("<tr>" + td("ИТОГО", cls="l", style=g) + td("", style=g) * 2 + td(tot["q"], style=g) + td(tot["r"], style=g)
             + td("", style=g) * 7 + td(tot["rev"], 0, g) + td(tot["com"], 0, g) + td(tot["log"], 0, g)
             + td(tot["rl"], 0, g) + td(tot["oth"], 0, g) + td("", style=g) + td(tot["out"], 0, g)
             + td(tot["prof"], 0, g) + td("", style=g) * 2 + "</tr>")
    note = ("Факт по начислениям Ozon. СЛЕВА — на 1 штуку, СПРАВА — итог за период. Прибыль = выплата − себестоимость "
            "(только на выкупленные штуки) − налог 7%. Оранжевым — возвраты. ЖЁЛТАЯ строка «API» — продажа из отчёта "
            "о реализации, которой ещё нет в выгрузке начислений (без логистики и услуг — оценка сверху). "
            "Последние колонки — сравнение с прибылью на штуку на Яндекс Маркете. <i>Данные выдуманные.</i>")
    return (f"<html><head><meta charset='utf-8'><style>{CSS}</style></head><body><div class='wrap'>"
            f"<div class='note'>{note}</div><table>{''.join(h)}</table>{tabs('Юнит-экономика')}</div></body></html>")


MEASURE = ("<script>window.onload=()=>{const w=document.querySelector('.wrap');"
           "document.body.setAttribute('data-size',Math.ceil(w.scrollWidth+24)+'x'+Math.ceil(w.scrollHeight+4));}"
           "</script>")


def shoot(html, name):
    """Два прохода: Chrome сначала меряет страницу (--dump-dom), потом снимает
    ровно по её размеру — без обрезки и пустых полей."""
    path = os.path.join(HERE, name + ".html")
    open(path, "w", encoding="utf-8").write(html.replace("</body>", MEASURE + "</body>"))
    common = [CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars"]
    dom = subprocess.run(common + ["--window-size=4000,3000", "--dump-dom", "file://" + path],
                         check=True, capture_output=True, text=True).stdout
    w, h = re.search(r'data-size="(\d+)x(\d+)"', dom).groups()
    png = os.path.join(HERE, name + ".png")
    subprocess.run(common + ["--force-device-scale-factor=2", f"--window-size={w},{h}",
                             f"--screenshot={png}", "file://" + path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    os.remove(path)
    print(f"→ {png} ({w}×{h})")


if __name__ == "__main__":
    shoot(unit_html(), "screenshot-unit-economics")
