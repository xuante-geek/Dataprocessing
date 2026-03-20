#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PYTHON_BIN="$ROOT_DIR/.venv/bin/python"
LOG_DIR="$ROOT_DIR/logs"
PORT="${DP_PORT:-5001}"
URL="http://127.0.0.1:${PORT}/?autorun=1"
HEALTH_URL="http://127.0.0.1:${PORT}/api/files"

mkdir -p "$LOG_DIR"

is_dp_service_ready() {
  curl -sf "$HEALTH_URL" >/dev/null 2>&1
}

if is_dp_service_ready; then
  open "$URL"
  exit 0
fi

# Port is occupied by another service (not DataProcessing).
if lsof -i :"$PORT" >/dev/null 2>&1; then
  echo "端口 ${PORT} 已被其他进程占用，且不是 DataProcessing 服务。"
  echo "请先释放端口后重试，或改用其他端口：DP_PORT=5001 bash scripts/run_ui.sh"
  lsof -nP -iTCP:"$PORT"
  exit 1
fi

nohup env DP_PORT="$PORT" "$PYTHON_BIN" "$ROOT_DIR/src/app.py" > "$LOG_DIR/autorun.app.log" 2>&1 &
SERVER_PID=$!
disown "$SERVER_PID" 2>/dev/null || true

for _ in {1..50}; do
  if is_dp_service_ready; then
    break
  fi
  sleep 0.2
done

if ! is_dp_service_ready; then
  echo "DataProcessing 服务启动失败，请检查日志：$LOG_DIR/autorun.app.log"
  exit 1
fi

open "$URL"
echo "DataProcessing 已启动 (PID: ${SERVER_PID}), 访问: ${URL}"
