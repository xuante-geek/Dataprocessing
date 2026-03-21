#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PLIST_PATH="$HOME/Library/LaunchAgents/com.xuante.dataprocessing.plist"
PYTHON_BIN="$ROOT_DIR/.venv/bin/python"
SCRIPT_PATH="$ROOT_DIR/scripts/run_frontend_autorun.py"
LOG_DIR="$ROOT_DIR/logs"
STDOUT_LOG="$LOG_DIR/autorun.launchd.out.log"
STDERR_LOG="$LOG_DIR/autorun.launchd.err.log"

mkdir -p "$LOG_DIR"

HOUR="${1:-18}"
MINUTE="${2:-0}"
RUN_AT_LOAD_FLAG="${3:-run}"
RUN_AT_LOAD_TAG="true"
if [ "$RUN_AT_LOAD_FLAG" = "no-run" ]; then
  RUN_AT_LOAD_TAG="false"
fi

if ! [[ "$HOUR" =~ ^[0-9]+$ ]] || [ "$HOUR" -lt 0 ] || [ "$HOUR" -gt 23 ]; then
  echo "Hour 必须是 0-23 的整数"
  exit 1
fi
if ! [[ "$MINUTE" =~ ^[0-9]+$ ]] || [ "$MINUTE" -lt 0 ] || [ "$MINUTE" -gt 59 ]; then
  echo "Minute 必须是 0-59 的整数"
  exit 1
fi
if [ ! -x "$PYTHON_BIN" ]; then
  echo "未找到可执行 Python：$PYTHON_BIN"
  exit 1
fi
if [ ! -f "$SCRIPT_PATH" ]; then
  echo "未找到脚本：$SCRIPT_PATH"
  exit 1
fi

cat > "$PLIST_PATH" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
  <dict>
    <key>Label</key>
    <string>com.xuante.dataprocessing</string>
    <key>ProgramArguments</key>
    <array>
      <string>${PYTHON_BIN}</string>
      <string>${SCRIPT_PATH}</string>
    </array>
    <key>WorkingDirectory</key>
    <string>${ROOT_DIR}</string>
    <key>StandardOutPath</key>
    <string>${STDOUT_LOG}</string>
    <key>StandardErrorPath</key>
    <string>${STDERR_LOG}</string>
    <key>StartCalendarInterval</key>
    <dict>
      <key>Hour</key>
      <integer>${HOUR}</integer>
      <key>Minute</key>
      <integer>${MINUTE}</integer>
    </dict>
    <key>RunAtLoad</key>
    <${RUN_AT_LOAD_TAG}/>
  </dict>
</plist>
EOF

launchctl unload "$PLIST_PATH" >/dev/null 2>&1 || true
launchctl load "$PLIST_PATH"

echo "Launchd 已安装：$PLIST_PATH"
echo "定时执行脚本：$SCRIPT_PATH"
echo "标准输出日志：$STDOUT_LOG"
echo "标准错误日志：$STDERR_LOG"
echo "修改执行时间请编辑 plist 的 StartCalendarInterval。"
