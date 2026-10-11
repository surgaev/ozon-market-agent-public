"""Команда `ozon-agent`: init — создать настройки, run — собрать отчёт."""
import argparse
import os
import sys
from datetime import date

DATA_DIR = os.environ.get("OZON_AGENT_HOME", os.path.expanduser("~/.ozon-agent"))
ENV_FILE = os.path.join(DATA_DIR, ".env")
COST_FILE = os.path.join(DATA_DIR, "cost_prices.csv")

ENV_TEMPLATE = """\
# Ключи Seller API: кабинет Ozon → Настройки → Seller API → создать ключ (роль Admin read only
# или с доступом к аналитике, финансам, товарам и остаткам).
OZON_CLIENT_ID=
OZON_API_KEY=
# Куда класть отчёты и где искать выгрузки из кабинета (по умолчанию ~/Downloads):
# OZON_REPORT_FOLDER=/Users/you/Downloads
# OZON_DOWNLOADS_DIR=/Users/you/Downloads
"""
COST_TEMPLATE = """\
# Закупочные цены, ₽. offer_id или его префикс: ABC123 покрывает ABC123-red, ABC123-xl.
# Разделитель — запятая или точка с запятой; можно редактировать в Excel.
sku,cost
ABC123,100
"""


def _load_env():
    if not os.path.exists(ENV_FILE):
        return
    for line in open(ENV_FILE, encoding="utf-8"):
        line = line.split("#", 1)[0].strip()
        if "=" in line:
            k, v = line.split("=", 1)
            if v.strip():
                os.environ.setdefault(k.strip(), v.strip())


def _bind_en0():
    """Исходящие соединения — через Wi-Fi (en0) в обход VPN, который режет API (macOS)."""
    import socket
    import subprocess
    import urllib3.connection as u3c
    import urllib3.util.connection as u3
    ip = subprocess.check_output(["ipconfig", "getifaddr", "en0"]).decode().strip()
    orig = u3.create_connection

    def create_connection(address, timeout=socket._GLOBAL_DEFAULT_TIMEOUT, source_address=None, **kw):
        return orig(address, timeout, (ip, 0), **kw)
    u3.create_connection = create_connection
    u3c.connection.create_connection = create_connection
    print(f"исходящие соединения привязаны к en0 ({ip})")


def cmd_init(_args):
    os.makedirs(DATA_DIR, exist_ok=True)
    for path, body in ((ENV_FILE, ENV_TEMPLATE), (COST_FILE, COST_TEMPLATE)):
        if os.path.exists(path):
            print(f"  уже есть: {path}")
        else:
            open(path, "w", encoding="utf-8").write(body)
            print(f"  создан:   {path}")
    print(f"\nДальше:\n  1. впишите Client-Id и Api-Key в {ENV_FILE}"
          f"\n  2. впишите закупочные цены в {COST_FILE}"
          f"\n  3. скачайте «Отчёт по начислениям» из кабинета в ~/Downloads"
          f"\n  4. запустите: ozon-agent run --days 30")


def cmd_run(args):
    _load_env()
    missing = [k for k in ("OZON_CLIENT_ID", "OZON_API_KEY") if not os.environ.get(k)]
    if missing:
        sys.exit(f"Не заполнено в {ENV_FILE}: {', '.join(missing)}. Сначала: ozon-agent init")
    if not os.path.exists(COST_FILE):
        print(f"⚠️  нет {COST_FILE} — себестоимость будет 0, прибыль завышена")
    if args.en0:
        _bind_en0()
    from . import main   # импорт после .env: config читает переменные при загрузке
    main.run(date.fromisoformat(args.date) if args.date else date.today(), days=args.days)


def main():
    p = argparse.ArgumentParser(prog="ozon-agent", description="Аналитика продавца Ozon → Excel-отчёт")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("init", help=f"создать шаблоны настроек в {DATA_DIR}").set_defaults(func=cmd_init)
    r = sub.add_parser("run", help="собрать отчёт")
    r.add_argument("--days", type=int, default=30, help="окно в днях, заканчивается вчера (по умолчанию 30)")
    r.add_argument("--date", help="дата отчёта YYYY-MM-DD (по умолчанию сегодня)")
    r.add_argument("--en0", action="store_true", help="macOS: в обход VPN через Wi-Fi (en0)")
    r.set_defaults(func=cmd_run)
    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
