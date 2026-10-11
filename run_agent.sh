#!/bin/bash
# Cron/launchd wrapper for Ozon Seller agent

set -e

AGENT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_FILE="$AGENT_DIR/agent.log"

# ── Load .env ──
if [ -f "$AGENT_DIR/.env" ]; then
  export $(grep -v '^#' "$AGENT_DIR/.env" | xargs)
fi

if [ -z "$OZON_CLIENT_ID" ] || [ -z "$OZON_API_KEY" ]; then
  echo "[$(date)] ERROR: OZON_CLIENT_ID / OZON_API_KEY not set" >> "$LOG_FILE"
  exit 1
fi

# ── Ожидание сети (Mac мог только что проснуться) ──
# Проверяем curl'ом по HTTPS, а не пингом: VPN пропускает ICMP, но режет TCP —
# тогда пинг проходит, проверка считает сеть живой, а прогон умирает на таймаутах
# внутри Python (20.09.2026 ровно так и вышло). Любой HTTP-ответ = путь рабочий.
API_URL="https://api-seller.ozon.ru"
DIAG_INTERFACE="en0"   # только для диагностики; трафик через него НЕ гоним
MAX_WAIT=120
WAITED=0
while ! curl -s -m 5 -o /dev/null "$API_URL" 2>/dev/null; do
  if [ $WAITED -ge $MAX_WAIT ]; then
    DEF_IFACE="$(route -n get default 2>/dev/null | awk '/interface:/{print $2}')"
    if curl -s -m 5 --interface "$DIAG_INTERFACE" -o /dev/null "$API_URL" 2>/dev/null; then
      echo "[$(date)] ERROR: API Озона недоступен через маршрут по умолчанию (${DEF_IFACE:-?}), но отвечает через ${DIAG_INTERFACE} — похоже, мешает VPN. Отключи VPN и запусти вручную: bash run_agent.sh --days 30" >> "$LOG_FILE"
      /usr/bin/osascript -e 'display notification "API Озона не отвечает через VPN, но доступен напрямую — отключи VPN и запусти вручную" with title "Ozon Agent: мешает VPN" sound name "Basso"' 2>/dev/null || true
    else
      echo "[$(date)] ERROR: Нет сети после ${MAX_WAIT}s (маршрут по умолчанию: ${DEF_IFACE:-нет}), прогон отменён" >> "$LOG_FILE"
      /usr/bin/osascript -e 'display notification "Нет сети — отчёт по Озону не собран" with title "Ozon Agent: нет сети" sound name "Basso"' 2>/dev/null || true
    fi
    exit 1
  fi
  sleep 5
  WAITED=$((WAITED + 5))
done

echo "[$(date)] Starting Ozon Seller agent..." >> "$LOG_FILE"

cd "$AGENT_DIR"

if [ -d "$AGENT_DIR/venv" ]; then
  source "$AGENT_DIR/venv/bin/activate"
fi

# "$@" — чтобы окно отчёта задавалось снаружи: launchd передаёт «--days 30»,
# вручную можно запустить с любым другим периодом. Без проброса main.py всегда
# брал бы окно по умолчанию (14 дней).
python3 main.py "$@" >> "$LOG_FILE" 2>&1

echo "[$(date)] Agent finished." >> "$LOG_FILE"
