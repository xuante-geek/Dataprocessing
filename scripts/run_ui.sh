#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PYTHON_BIN="$ROOT_DIR/.venv/bin/python"
LOG_DIR="$ROOT_DIR/logs"
URL="http://127.0.0.1:5000/?autorun=1"

mkdir -p "$LOG_DIR"

if lsof -i :5000 >/dev/null 2>&1; then
  open "$URL"
  exit 0
fi

"$PYTHON_BIN" "$ROOT_DIR/src/app.py" > "$LOG_DIR/autorun.app.log" 2>&1 &
SERVER_PID=$!

for _ in {1..50}; do
  if curl -sf http://127.0.0.1:5000/api/files >/dev/null 2>&1; then
    break
  fi
  sleep 0.2
done

open "$URL"

wait "$SERVER_PID"
