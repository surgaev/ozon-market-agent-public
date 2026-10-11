"""Разовый запуск агента Озона с привязкой исходящих соединений к en0.
Нужен только пока активен VPN, который режет api-seller.ozon.ru. Код агента не трогаем."""
import os, sys, socket, subprocess
from datetime import date

AGENT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AGENT); os.chdir(AGENT)

ip = subprocess.check_output(["ipconfig", "getifaddr", "en0"]).decode().strip()
print(f"привязка исходящих соединений к en0 ({ip})")

import urllib3.util.connection as u3
_orig = u3.create_connection
def create_connection(address, timeout=socket._GLOBAL_DEFAULT_TIMEOUT, source_address=None, **kw):
    return _orig(address, timeout, (ip, 0), **kw)
u3.create_connection = create_connection
import urllib3.connection as u3c
u3c.connection.create_connection = create_connection

for k, v in (l.split("=", 1) for l in open(f"{AGENT}/.env") if "=" in l and not l.startswith("#")):
    os.environ.setdefault(k.strip(), v.strip())

import main
main.run(date.today(), days=int(sys.argv[1]))
