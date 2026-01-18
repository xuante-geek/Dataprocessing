from __future__ import annotations

import csv
import datetime as dt
import math
import os
from pathlib import Path
from typing import Iterable
from bisect import bisect_left, insort

from flask import Flask, jsonify, request

try:
    import openpyxl
    from openpyxl.utils.datetime import from_excel
    from openpyxl.utils.cell import get_column_letter
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
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return str(value)
        rounded = round(value, 4)
        text = f"{rounded:.4f}"
        return text.rstrip("0").rstrip(".")
    return str(value)

def _round_for_output(value: object) -> object:
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return value
        return round(value, 4)
    return value


def _is_garbled_text(text: str) -> bool:
    if "\ufffd" in text:
        return True
    for char in text:
        code_point = ord(char)
        if code_point < 32 and char not in ("\t", "\n", "\r"):
            return True
    return False


def _parse_date(value: object, *, epoch: dt.datetime) -> dt.date:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date) and not isinstance(value, dt.datetime):
        return value
    if isinstance(value, bool):
        raise ValueError("布尔类型不是有效日期")
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            raise ValueError("数值为 NaN/Inf")
        return from_excel(value, epoch=epoch).date()
    if isinstance(value, str):
        text = value.strip()
        if not text:
            raise ValueError("日期为空白")
        candidates = (
            "%Y-%m-%d",
            "%Y/%m/%d",
            "%Y.%m.%d",
            "%Y%m%d",
            "%Y-%m-%d %H:%M:%S",
            "%Y/%m/%d %H:%M:%S",
        )
        for fmt in candidates:
            try:
                parsed = dt.datetime.strptime(text, fmt)
                return parsed.date()
            except ValueError:
                continue
        try:
            return dt.date.fromisoformat(text)
        except ValueError as exc:
            raise ValueError(f"无法解析日期：{text}") from exc
    raise ValueError(f"不支持的日期类型：{type(value).__name__}")


def _validate_text_or_number(value: object) -> object:
    if value is None:
        raise ValueError("内容空白")
    if isinstance(value, bool):
        raise ValueError("不支持布尔类型")
    if isinstance(value, str):
        text = value.strip()
        if not text:
            raise ValueError("内容空白")
        if _is_garbled_text(text):
            raise ValueError("疑似乱码/控制字符")
        return text
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            raise ValueError("数值为 NaN/Inf")
        return value
    raise ValueError(f"不支持的类型：{type(value).__name__}")


def _validate_header_cell(value: object) -> str:
    if value is None:
        raise ValueError("标题空白")
    if not isinstance(value, str):
        raise ValueError("标题必须为文本")
    text = value.strip()
    if not text:
        raise ValueError("标题空白")
    if _is_garbled_text(text):
        raise ValueError("标题疑似乱码/控制字符")
    return text


def process_xlsx_to_outputs(source_path: Path, output_csv_path: Path, output_xlsx_path: Path) -> None:
    workbook = openpyxl.load_workbook(source_path, data_only=True, read_only=True)
    sheet_name = workbook.sheetnames[0] if workbook.sheetnames else None
    if not sheet_name:
        raise ValueError("未找到可用工作表")

    sheet = workbook[sheet_name]
    rows_iter = sheet.iter_rows(values_only=True)
    header_values = next(rows_iter, None)
    if not header_values:
        raise ValueError("未找到标题行")

    def _is_blank(value: object) -> bool:
        return value is None or (isinstance(value, str) and not value.strip())

    last_col = 0
    for column_index, value in enumerate(header_values, start=1):
        if not _is_blank(value):
            last_col = column_index

    if last_col == 0:
        raise ValueError("标题行为空")

    columns_to_keep: list[int] = [col for col in range(1, last_col + 1) if col not in (2, 3, 4)]
    if 1 not in columns_to_keep:
        columns_to_keep.insert(0, 1)

    output_rows: list[list[object]] = []
    row_dates: list[dt.date] = []

    kept_header: list[str] = []
    for col in columns_to_keep:
        value = header_values[col - 1] if col - 1 < len(header_values) else None
        coordinate = f"{get_column_letter(col)}1"
        try:
            kept_header.append(_validate_header_cell(value))
        except ValueError as exc:
            raise ValueError(f"{coordinate} 标题错误：{exc}") from exc

    output_rows.append(kept_header)

    for row_offset, row_values in enumerate(rows_iter, start=2):
        values = list(row_values[:last_col])
        if len(values) < last_col:
            values.extend([None] * (last_col - len(values)))

        kept_values = [values[col - 1] for col in columns_to_keep]
        if all(_is_blank(value) for value in kept_values):
            continue

        normalized_row: list[object] = []
        for position, col in enumerate(columns_to_keep):
            value = values[col - 1]
            coordinate = f"{get_column_letter(col)}{row_offset}"
            try:
                if position == 0:
                    parsed = _parse_date(value, epoch=workbook.epoch)
                    normalized_row.append(parsed.isoformat())
                else:
                    normalized_row.append(_validate_text_or_number(value))
            except ValueError as exc:
                raise ValueError(f"{coordinate} 内容错误：{exc}") from exc

        output_rows.append(normalized_row)
        row_dates.append(dt.date.fromisoformat(normalized_row[0]))

    if len(output_rows) <= 1:
        raise ValueError("没有可导出的数据行")

    data_rows = output_rows[1:]
    if len(data_rows) != len(row_dates):
        raise ValueError("内部错误：行数不一致")

    sorted_data_rows = [row for _, row in sorted(zip(row_dates, data_rows), key=lambda item: item[0])]
    final_rows = [output_rows[0], *sorted_data_rows]

    output_csv_path.parent.mkdir(parents=True, exist_ok=True)
    with output_csv_path.open("w", encoding="utf-8-sig", newline="") as file_handle:
        writer = csv.writer(file_handle)
        for row in final_rows:
            writer.writerow([_cell_to_text(value) for value in row])

    workbook_out = openpyxl.Workbook()
    sheet_out = workbook_out.active
    sheet_out.title = "processed"
    sheet_out.freeze_panes = "B2"
    for row in final_rows:
        sheet_out.append([_round_for_output(value) for value in row])
    output_xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    workbook_out.save(output_xlsx_path)

def _find_input_xlsx(stem: str) -> Path:
    if not INPUT_DIR.exists():
        raise FileNotFoundError("input/ 目录不存在")

    candidates: list[Path] = []
    for path in INPUT_DIR.iterdir():
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() != ".xlsx":
            continue
        if path.stem.lower() == stem.lower():
            candidates.append(path)

    if not candidates:
        raise FileNotFoundError(f"未找到文件：{stem}.xlsx（请放入 input/）")
    if len(candidates) > 1:
        raise FileNotFoundError(f"找到多个匹配文件：{stem}.xlsx（请保留一个）")
    return candidates[0]


def _coerce_float(value: object) -> float:
    if isinstance(value, bool):
        raise ValueError("不支持布尔类型")
    if value is None:
        raise ValueError("内容空白")
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            raise ValueError("数值为 NaN/Inf")
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            raise ValueError("内容空白")
        if _is_garbled_text(text):
            raise ValueError("疑似乱码/控制字符")
        cleaned = text.replace(",", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1].strip()
        try:
            return float(cleaned)
        except ValueError as exc:
            raise ValueError(f"无法解析为数值：{text}") from exc
    raise ValueError(f"不支持的类型：{type(value).__name__}")


def _normalize_yield(yield_raw: float) -> float:
    if yield_raw > 1.0:
        return yield_raw / 100.0
    return yield_raw


def _iter_rows_values(sheet: object, *, last_col: int) -> Iterable[tuple[object, ...]]:
    for row_values in sheet.iter_rows(values_only=True):
        values = tuple(row_values[:last_col])
        if len(values) < last_col:
            values = values + (None,) * (last_col - len(values))
        yield values


def _validate_expected_header(actual: object, expected: str, coordinate: str) -> None:
    text = _validate_header_cell(actual)
    if text != expected:
        raise ValueError(f"{coordinate} 标题不匹配：期望“{expected}”，实际“{text}”")


def _process_data_pe(source_path: Path) -> list[tuple[dt.date, float, float]]:
    workbook = openpyxl.load_workbook(source_path, data_only=True, read_only=True)
    sheet_name = workbook.sheetnames[0] if workbook.sheetnames else None
    if not sheet_name:
        raise ValueError("data_PE：未找到可用工作表")
    sheet = workbook[sheet_name]

    last_col = 8
    rows_iter = _iter_rows_values(sheet, last_col=last_col)
    header = next(rows_iter, None)
    if not header:
        raise ValueError("data_PE：未找到标题行")

    _validate_expected_header(header[0], "日期", "A1")
    _validate_expected_header(header[3], "PE-TTM-S", "D1")
    _validate_expected_header(header[7], "收盘点位", "H1")

    fill_dates = [
        "2018-08-03",
        "2018-08-06",
        "2018-08-07",
        "2018-08-08",
        "2018-08-09",
        "2018-08-10",
        "2018-08-13",
        "2018-08-14",
        "2018-08-15",
        "2018-08-16",
        "2018-08-17",
        "2018-08-20",
        "2018-08-21",
        "2018-08-22",
        "2018-08-23",
        "2018-08-24",
    ]
    fill_values = [
        3892.88,
        3828.14,
        3933.12,
        3871.35,
        3963.8,
        3979.61,
        3978.56,
        3962.88,
        3876.46,
        3846.75,
        3785.01,
        3814.7,
        3870.75,
        3838.79,
        3856.65,
        3854.99,
    ]
    fill_close_by_date = {
        dt.date.fromisoformat(date): value for date, value in zip(fill_dates, fill_values)
    }

    rows: list[tuple[dt.date, float, float]] = []
    for row_index, values in enumerate(rows_iter, start=2):
        if all(value is None or (isinstance(value, str) and not value.strip()) for value in values):
            continue

        date = _parse_date(values[0], epoch=workbook.epoch)
        try:
            pe = _coerce_float(values[3])
        except ValueError as exc:
            raise ValueError(f"data_PE D{row_index} 内容错误：{exc}") from exc

        close_value = values[7]
        if date in fill_close_by_date and (
            close_value is None or (isinstance(close_value, str) and not close_value.strip())
        ):
            close_value = fill_close_by_date[date]
        try:
            close = _coerce_float(close_value)
        except ValueError as exc:
            raise ValueError(f"data_PE H{row_index} 内容错误：{exc}") from exc

        if pe <= 0:
            raise ValueError(f"data_PE D{row_index} 内容错误：PE 必须为正数")

        rows.append((date, pe, close))

    if not rows:
        raise ValueError("data_PE：没有可用数据行")

    rows.sort(key=lambda item: item[0])
    return rows


def _process_data_bond(source_path: Path) -> list[tuple[dt.date, float, float]]:
    workbook = openpyxl.load_workbook(source_path, data_only=True, read_only=True)
    sheet_name = workbook.sheetnames[0] if workbook.sheetnames else None
    if not sheet_name:
        raise ValueError("data_bond：未找到可用工作表")
    sheet = workbook[sheet_name]

    last_col = 5
    rows_iter = _iter_rows_values(sheet, last_col=last_col)
    header = next(rows_iter, None)
    if not header:
        raise ValueError("data_bond：未找到标题行")

    _validate_expected_header(header[0], "日期", "A1")
    _validate_expected_header(header[4], "十年期收益率", "E1")

    rows: list[tuple[dt.date, float, float]] = []
    for row_index, values in enumerate(rows_iter, start=2):
        if all(value is None or (isinstance(value, str) and not value.strip()) for value in values):
            continue

        date = _parse_date(values[0], epoch=workbook.epoch)
        try:
            yield_raw = _coerce_float(values[4])
        except ValueError as exc:
            raise ValueError(f"data_bond E{row_index} 内容错误：{exc}") from exc

        rows.append((date, yield_raw, _normalize_yield(yield_raw)))

    if not rows:
        raise ValueError("data_bond：没有可用数据行")

    rows.sort(key=lambda item: item[0])
    return rows


def _merge_by_bond_dates(
    bond_rows: list[tuple[dt.date, float, float]],
    pe_rows: list[tuple[dt.date, float, float]],
) -> list[tuple[dt.date, float, float, float]]:
    merged: list[tuple[dt.date, float, float, float]] = []
    pe_index = 0

    for bond_date, bond_yield_raw, bond_yield_decimal in bond_rows:
        while pe_index < len(pe_rows) and pe_rows[pe_index][0] < bond_date:
            pe_index += 1

        if pe_index >= len(pe_rows):
            raise ValueError("合并失败：data_PE 数据不足，无法继续对齐日期")

        pe_date, pe_value, pe_close = pe_rows[pe_index]
        if pe_date >= bond_date:
            merged.append((bond_date, bond_yield_raw, pe_value, pe_close))
            pe_index += 1
            continue

    if not merged:
        raise ValueError("合并失败：未生成任何对齐行")

    return merged


def _compute_erp_rows(
    merged_rows: list[tuple[dt.date, float, float, float]],
    bond_rows: list[tuple[dt.date, float, float]],
) -> list[list[object]]:
    bond_decimal_by_date = {date: decimal for date, _, decimal in bond_rows}
    output: list[list[object]] = [["日期", "十年期收益率", "PE-TTM-S", "收盘点位", "股权风险溢价"]]

    for date, yield_raw, pe_value, close_value in merged_rows:
        bond_yield_decimal = bond_decimal_by_date.get(date)
        if bond_yield_decimal is None:
            raise ValueError("内部错误：未找到收益率小数值")
        erp = (1.0 + 1.0 / pe_value) / (1.0 + bond_yield_decimal) - 1.0
        output.append([date.isoformat(), yield_raw, pe_value, close_value, erp])

    return output


def _write_csv(rows: list[list[object]], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as file_handle:
        writer = csv.writer(file_handle)
        for row in rows:
            writer.writerow([_cell_to_text(value) for value in row])


def _write_xlsx(rows: list[list[object]], path: Path, sheet_title: str) -> None:
    workbook_out = openpyxl.Workbook()
    sheet_out = workbook_out.active
    sheet_out.title = sheet_title
    sheet_out.freeze_panes = "B2"
    for row in rows:
        sheet_out.append([_round_for_output(value) for value in row])
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook_out.save(path)

def _rolling_median(sorted_window: list[float]) -> float:
    size = len(sorted_window)
    if size == 0:
        raise ValueError("窗口为空")
    mid = size // 2
    if size % 2 == 1:
        return float(sorted_window[mid])
    return (float(sorted_window[mid - 1]) + float(sorted_window[mid])) / 2.0


def _rolling_stddevp(sum_values: float, sum_squares: float, size: int) -> float:
    if size <= 0:
        raise ValueError("窗口为空")
    mean = sum_values / size
    variance = (sum_squares / size) - (mean * mean)
    if variance < 0 and variance > -1e-12:
        variance = 0.0
    if variance < 0:
        raise ValueError("方差为负数（数值异常）")
    return math.sqrt(variance)


def _compute_erp_10year_bands(
    erp_rows: list[list[object]],
    *,
    window_size: int = 2000,
) -> list[list[object]]:
    if not erp_rows or len(erp_rows) < 2:
        raise ValueError("ERP 数据为空")

    header = erp_rows[0]
    if len(header) < 5 or header[4] != "股权风险溢价":
        raise ValueError("ERP 表头不符合预期")

    data_rows = erp_rows[1:]
    if len(data_rows) < window_size:
        raise ValueError(f"数据不足：至少需要 {window_size} 行交易日数据")

    output: list[list[object]] = [
        ["日期", "十年期收益率", "PE-TTM-S", "收盘点位", "股权风险溢价", "+2σ", "+1σ", "中位数", "-1σ", "-2σ"]
    ]

    sorted_window: list[float] = []
    queue: list[float] = []
    sum_values = 0.0
    sum_squares = 0.0

    for index, row in enumerate(data_rows):
        erp_value = row[4]
        if not isinstance(erp_value, (int, float)):
            raise ValueError(f"ERP 第 {index + 2} 行数值类型不合法")
        erp_float = float(erp_value)

        insort(sorted_window, erp_float)
        queue.append(erp_float)
        sum_values += erp_float
        sum_squares += erp_float * erp_float

        if len(queue) > window_size:
            leaving = queue.pop(0)
            sum_values -= leaving
            sum_squares -= leaving * leaving
            remove_index = bisect_left(sorted_window, leaving)
            if remove_index >= len(sorted_window) or sorted_window[remove_index] != leaving:
                raise ValueError("内部错误：滚动窗口移除失败")
            sorted_window.pop(remove_index)

        if len(queue) < window_size:
            continue

        median = _rolling_median(sorted_window)
        stddevp = _rolling_stddevp(sum_values, sum_squares, window_size)
        upper2 = median + 2 * stddevp
        upper1 = median + stddevp
        lower1 = median - stddevp
        lower2 = median - 2 * stddevp

        output.append([row[0], row[1], row[2], row[3], erp_float, upper2, upper1, median, lower1, lower2])

    return output


@app.get("/")
def index() -> object:
    return app.send_static_file("index.html")


@app.get("/api/files")
def list_files() -> object:
    if not INPUT_DIR.exists():
        return jsonify({"files": []})

    files = sorted(
        p.name
        for p in INPUT_DIR.iterdir()
        if p.is_file() and p.suffix.lower() == ".xlsx"
        if not p.name.startswith("~$")
    )
    return jsonify({"files": files})


@app.post("/api/convert")
def convert_file() -> object:
    payload = request.get_json(silent=True) or {}
    filename = payload.get("filename")

    if not filename:
        return jsonify({"error": "缺少文件名"}), 400

    safe_name = Path(filename).name
    if safe_name != filename or not safe_name.lower().endswith(".xlsx") or safe_name.startswith("~$"):
        return jsonify({"error": "文件名不合法"}), 400

    source_path = INPUT_DIR / safe_name
    if not source_path.exists():
        return jsonify({"error": "文件不存在"}), 404

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_csv_path = OUTPUT_DIR / f"{source_path.stem}.csv"
    output_xlsx_path = OUTPUT_DIR / f"{source_path.stem}_processed.xlsx"

    try:
        process_xlsx_to_outputs(source_path, output_csv_path, output_xlsx_path)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:  # pragma: no cover - surfaced to UI
        return jsonify({"error": f"转换失败：{exc}"}), 500

    return jsonify({"output_csv": output_csv_path.name, "output_xlsx": output_xlsx_path.name})


@app.post("/api/erp")
def generate_erp() -> object:
    try:
        pe_path = _find_input_xlsx("data_PE")
        bond_path = _find_input_xlsx("data_bond")

        pe_rows = _process_data_pe(pe_path)
        bond_rows = _process_data_bond(bond_path)
        merged_rows = _merge_by_bond_dates(bond_rows, pe_rows)

        pe_clean_rows: list[list[object]] = [["日期", "PE-TTM-S", "收盘点位"]] + [
            [date.isoformat(), pe, close] for date, pe, close in pe_rows
        ]
        bond_clean_rows: list[list[object]] = [["日期", "十年期收益率"]] + [
            [date.isoformat(), yield_raw] for date, yield_raw, _ in bond_rows
        ]
        merged_clean_rows: list[list[object]] = [["日期", "十年期收益率", "PE-TTM-S", "收盘点位"]] + [
            [date.isoformat(), yield_raw, pe, close] for date, yield_raw, pe, close in merged_rows
        ]
        erp_rows = _compute_erp_rows(merged_rows, bond_rows)

        output = {
            "data_PE_clean": ("data_PE_clean.csv", "data_PE_clean.xlsx"),
            "data_bond_clean": ("data_bond_clean.csv", "data_bond_clean.xlsx"),
            "merged": ("merged.csv", "merged.xlsx"),
            "erp": ("ERP.csv", "ERP.xlsx"),
        }

        _write_csv(pe_clean_rows, OUTPUT_DIR / output["data_PE_clean"][0])
        _write_xlsx(pe_clean_rows, OUTPUT_DIR / output["data_PE_clean"][1], "data_PE_clean")
        _write_csv(bond_clean_rows, OUTPUT_DIR / output["data_bond_clean"][0])
        _write_xlsx(bond_clean_rows, OUTPUT_DIR / output["data_bond_clean"][1], "data_bond_clean")
        _write_csv(merged_clean_rows, OUTPUT_DIR / output["merged"][0])
        _write_xlsx(merged_clean_rows, OUTPUT_DIR / output["merged"][1], "merged")
        _write_csv(erp_rows, OUTPUT_DIR / output["erp"][0])
        _write_xlsx(erp_rows, OUTPUT_DIR / output["erp"][1], "ERP")

        return jsonify(
            {
                "outputs": {
                    "data_PE_clean_csv": output["data_PE_clean"][0],
                    "data_PE_clean_xlsx": output["data_PE_clean"][1],
                    "data_bond_clean_csv": output["data_bond_clean"][0],
                    "data_bond_clean_xlsx": output["data_bond_clean"][1],
                    "merged_csv": output["merged"][0],
                    "merged_xlsx": output["merged"][1],
                    "erp_csv": output["erp"][0],
                    "erp_xlsx": output["erp"][1],
                }
            }
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:  # pragma: no cover
        return jsonify({"error": f"生成失败：{exc}"}), 500

@app.post("/api/erp10y")
def generate_erp_10year() -> object:
    try:
        pe_path = _find_input_xlsx("data_PE")
        bond_path = _find_input_xlsx("data_bond")

        pe_rows = _process_data_pe(pe_path)
        bond_rows = _process_data_bond(bond_path)
        merged_rows = _merge_by_bond_dates(bond_rows, pe_rows)
        erp_rows = _compute_erp_rows(merged_rows, bond_rows)

        bands_rows = _compute_erp_10year_bands(erp_rows, window_size=2000)

        csv_name = "ERP_10Year.csv"
        xlsx_name = "ERP_10Year.xlsx"
        _write_csv(bands_rows, OUTPUT_DIR / csv_name)
        _write_xlsx(bands_rows, OUTPUT_DIR / xlsx_name, "ERP_10Year")

        return jsonify({"output_csv": csv_name, "output_xlsx": xlsx_name})
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:  # pragma: no cover
        return jsonify({"error": f"生成失败：{exc}"}), 500


if __name__ == "__main__":
    debug = os.environ.get("DP_DEBUG") == "1"
    app.run(host="127.0.0.1", port=5000, debug=debug, use_reloader=False)
