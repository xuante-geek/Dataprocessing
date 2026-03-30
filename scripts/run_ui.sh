#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PYTHON_BIN="$ROOT_DIR/.venv/bin/python"
LOG_DIR="$ROOT_DIR/logs"
URL="http://127.0.0.1:5001/?autorun=1"
REQUIRED_PYTHON="3.9.6"

mkdir -p "$LOG_DIR"

if [ ! -x "$PYTHON_BIN" ]; then
  echo "缺少虚拟环境：$PYTHON_BIN"
  echo "请先创建并安装依赖：python3.9 -m venv .venv && .venv/bin/python -m pip install -r requirements.txt"
  exit 1
fi

PY_VERSION="$("$PYTHON_BIN" -c 'import sys; print(".".join(map(str, sys.version_info[:3])))')"
if [ "$PY_VERSION" != "$REQUIRED_PYTHON" ]; then
  echo "Python 版本不符合要求：当前 $PY_VERSION，要求 $REQUIRED_PYTHON"
  echo "请使用 Python 3.9.6 重建 .venv"
  exit 1
fi

if lsof -i :5001 >/dev/null 2>&1; then
  open "$URL"
  exit 0
fi

"$PYTHON_BIN" "$ROOT_DIR/src/app.py" > "$LOG_DIR/autorun.app.log" 2>&1 &
SERVER_PID=$!

for _ in {1..50}; do
  if curl -sf http://127.0.0.1:5001/api/files >/dev/null 2>&1; then
    break
  fi
  sleep 0.2
done

open "$URL"

wait "$SERVER_PID"
