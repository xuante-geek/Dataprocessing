#!/usr/bin/env python3
from __future__ import annotations

import os
from pathlib import Path
import subprocess
import sys
import time
import urllib.error
import urllib.request

try:
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
    from playwright.sync_api import sync_playwright
except ImportError as exc:  # pragma: no cover
    raise SystemExit("缺少依赖：playwright。请先安装 requirements.txt 并安装 chromium。") from exc


ROOT = Path(__file__).resolve().parents[1]
PYTHON_BIN = ROOT / ".venv" / "bin" / "python"
APP_PATH = ROOT / "src" / "app.py"
LOG_DIR = ROOT / "logs"
APP_LOG_PATH = LOG_DIR / "autorun.app.log"

PORT = int(os.environ.get("DP_PORT", "5001"))
BASE_URL = f"http://127.0.0.1:{PORT}"
HEALTH_URL = f"{BASE_URL}/api/files"
AUTORUN_URL = f"{BASE_URL}/?autorun=1"
TIMEOUT_SECONDS = int(os.environ.get("AUTORUN_UI_TIMEOUT_SECONDS", "14400"))


def _is_service_ready(timeout: float = 2.0) -> bool:
    req = urllib.request.Request(HEALTH_URL, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return 200 <= resp.status < 300
    except Exception:
        return False


def _start_service_if_needed() -> subprocess.Popen[bytes] | None:
    if _is_service_ready():
        print(f"[autorun-ui] service already ready: {HEALTH_URL}", flush=True)
        return None

    if not PYTHON_BIN.exists():
        raise RuntimeError(f"未找到 Python：{PYTHON_BIN}")
    if not APP_PATH.exists():
        raise RuntimeError(f"未找到应用入口：{APP_PATH}")

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_handle = APP_LOG_PATH.open("ab")
    env = dict(os.environ)
    env["DP_PORT"] = str(PORT)
    process = subprocess.Popen(
        [str(PYTHON_BIN), str(APP_PATH)],
        cwd=str(ROOT),
        env=env,
        stdout=log_handle,
        stderr=subprocess.STDOUT,
    )

    for _ in range(100):
        if _is_service_ready():
            print(f"[autorun-ui] service started: {HEALTH_URL}", flush=True)
            return process
        exit_code = process.poll()
        if exit_code is not None:
            raise RuntimeError(f"本地服务启动失败，进程已退出（code={exit_code}）。请检查 {APP_LOG_PATH}")
        time.sleep(0.2)
    raise RuntimeError(f"本地服务启动超时，请检查 {APP_LOG_PATH}")


def _wait_for_runall_modal() -> None:
    deadline = time.time() + TIMEOUT_SECONDS
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(headless=False, channel="chrome")
        except Exception:
            browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto(AUTORUN_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(1200)

        modal = page.locator("#modal")
        modal_title = page.locator("#modal-title")
        modal_message = page.locator("#modal-message")

        while time.time() < deadline:
            if page.is_closed():
                raise RuntimeError("浏览器页面已关闭，任务中断。")

            try:
                is_visible = modal.is_visible()
            except PlaywrightTimeoutError:
                is_visible = False

            if is_visible:
                title = modal_title.inner_text().strip()
                message = modal_message.inner_text().strip()
                if title == "全部运行完成" and "所有步骤已按顺序完成" in message:
                    browser.close()
                    return
                if title == "全部运行失败":
                    browser.close()
                    raise RuntimeError(f"全部运行失败：{message}")

            page.wait_for_timeout(1000)

        browser.close()
        raise RuntimeError(f"等待“全部运行完成”超时（>{TIMEOUT_SECONDS} 秒）。")


def _sleep_display_now() -> None:
    subprocess.run(["pmset", "displaysleepnow"], check=False)


def _wake_display_once() -> None:
    subprocess.run(["caffeinate", "-u", "-t", "5"], check=False)


def _start_keep_awake() -> subprocess.Popen[bytes]:
    return subprocess.Popen(["caffeinate", "-dimsu"])


def _stop_process(process: subprocess.Popen[bytes] | None) -> None:
    if process is None:
        return
    if process.poll() is not None:
        return
    process.terminate()
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.kill()


def _stop_keep_awake(process: subprocess.Popen[bytes] | None) -> None:
    if process is None:
        return
    if process.poll() is not None:
        return
    process.terminate()
    try:
        process.wait(timeout=2)
    except subprocess.TimeoutExpired:
        process.kill()


def main() -> int:
    print("[autorun-ui] start", flush=True)
    _wake_display_once()
    keep_awake_proc = _start_keep_awake()
    service_proc: subprocess.Popen[bytes] | None = None

    try:
        service_proc = _start_service_if_needed()
        _wait_for_runall_modal()
        print("[autorun-ui] run-all completed", flush=True)
        _sleep_display_now()
        print("[autorun-ui] display sleep requested", flush=True)
        return 0
    except Exception as exc:
        print(f"[autorun-ui] failed: {exc}", file=sys.stderr, flush=True)
        return 1
    finally:
        _stop_keep_awake(keep_awake_proc)
        _stop_process(service_proc)


if __name__ == "__main__":
    raise SystemExit(main())
