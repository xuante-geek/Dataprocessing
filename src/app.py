from __future__ import annotations

import csv
import datetime as dt
from pathlib import Path

from flask import Flask, jsonify, request

try:
    import openpyxl
except ImportError as exc:  # pragma: no cover - runtime dependency check
    raise SystemExit(
        "缺少依赖：openpyxl。请先安装 requirements.txt 后再运行。"
    ) from exc

BASE_DIR = Path(__file__).resolve().parents[1]
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "docs" / "data"
DOCS_DIR = BASE_DIR / "docs"

app = Flask(__name__, static_folder=str(DOCS_DIR), static_url_path="")


def _cell_to_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, (dt.date, dt.datetime, dt.time)):
        return value.isoformat()
    return str(value)


def convert_xlsx_to_csv(source_path: Path, output_path: Path) -> None:
    workbook = openpyxl.load_workbook(source_path, data_only=True, read_only=True)
    sheet_name = workbook.sheetnames[0] if workbook.sheetnames else None
    if not sheet_name:
        raise ValueError("未找到可用工作表")

    sheet = workbook[sheet_name]
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open("w", encoding="utf-8-sig", newline="") as file_handle:
        writer = csv.writer(file_handle)
        for row in sheet.iter_rows(values_only=True):
            writer.writerow([_cell_to_text(cell) for cell in row])


@app.get("/")
def index() -> object:
    return app.send_static_file("index.html")


@app.get("/api/files")
def list_files() -> object:
    if not INPUT_DIR.exists():
        return jsonify({"files": []})

    files = sorted(p.name for p in INPUT_DIR.glob("*.xlsx"))
    return jsonify({"files": files})


@app.post("/api/convert")
def convert_file() -> object:
    payload = request.get_json(silent=True) or {}
    filename = payload.get("filename")

    if not filename:
        return jsonify({"error": "缺少文件名"}), 400

    safe_name = Path(filename).name
    if safe_name != filename or not safe_name.lower().endswith(".xlsx"):
        return jsonify({"error": "文件名不合法"}), 400

    source_path = INPUT_DIR / safe_name
    if not source_path.exists():
        return jsonify({"error": "文件不存在"}), 404

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / f"{source_path.stem}.csv"

    try:
        convert_xlsx_to_csv(source_path, output_path)
    except Exception as exc:  # pragma: no cover - surfaced to UI
        return jsonify({"error": f"转换失败：{exc}"}), 500

    return jsonify({"output": output_path.name})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
