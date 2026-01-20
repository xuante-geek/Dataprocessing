from __future__ import annotations

from bisect import bisect_left, bisect_right, insort
from collections import deque
import csv
import datetime as dt
import math
import os
from pathlib import Path
from typing import Iterable

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

OUTPUT_DECIMAL_PLACES = 6


def _cell_to_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, (dt.date, dt.datetime, dt.time)):
        return value.isoformat()
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return str(value)
        rounded = round(value, OUTPUT_DECIMAL_PLACES)
        text = f"{rounded:.{OUTPUT_DECIMAL_PLACES}f}"
        return text.rstrip("0").rstrip(".")
    return str(value)

def _round_for_output(value: object) -> object:
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return value
        return round(value, OUTPUT_DECIMAL_PLACES)
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

    def normalize(text: str) -> str:
        return " ".join(text.strip().split()).lower()

    candidates: list[Path] = []
    for path in INPUT_DIR.iterdir():
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() != ".xlsx":
            continue
        if normalize(path.stem) == normalize(stem):
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


def _rolling_percentile(sorted_window: list[float], value: float) -> float:
    window_size = len(sorted_window)
    if window_size <= 0:
        raise ValueError("窗口为空")
    if window_size == 1:
        return 50.0
    left = bisect_left(sorted_window, value)
    right = bisect_right(sorted_window, value)
    rank_low = left + 1
    rank_high = right
    avg_rank = (rank_low + rank_high) / 2.0
    return 100.0 * (avg_rank - 1.0) / (window_size - 1.0)


def _moving_average(values: list[float], window: int) -> list[float | None]:
    if window <= 0:
        raise ValueError("移动平均窗口必须为正整数")
    out: list[float | None] = []
    q: deque[float] = deque()
    sum_values = 0.0
    for value in values:
        q.append(value)
        sum_values += value
        if len(q) > window:
            sum_values -= q.popleft()
        if len(q) == window:
            out.append(sum_values / window)
        else:
            out.append(None)
    return out


def _rolling_percentiles(values: list[float | None], window: int) -> list[float | None]:
    if window <= 0:
        raise ValueError("滚动窗口必须为正整数")

    first_valid = 0
    while first_valid < len(values) and values[first_valid] is None:
        first_valid += 1

    out: list[float | None] = [None] * len(values)
    if first_valid >= len(values):
        return out

    sorted_window: list[float] = []
    q: deque[float] = deque()

    for index in range(first_valid, len(values)):
        current = values[index]
        if current is None:
            out[index] = None
            continue

        insort(sorted_window, float(current))
        q.append(float(current))
        if len(q) > window:
            leaving = q.popleft()
            remove_index = bisect_left(sorted_window, leaving)
            if remove_index >= len(sorted_window) or sorted_window[remove_index] != leaving:
                raise ValueError("内部错误：滚动窗口移除失败")
            sorted_window.pop(remove_index)

        if len(q) < window:
            out[index] = None
            continue

        out[index] = _rolling_percentile(sorted_window, float(current))

    return out


def _load_ratio_series(source_path: Path) -> tuple[list[str], list[float]]:
    workbook = openpyxl.load_workbook(source_path, data_only=True, read_only=True)
    try:
        sheet_name = workbook.sheetnames[0] if workbook.sheetnames else None
        if not sheet_name:
            raise ValueError(f"{source_path.name}：未找到可用工作表")
        sheet = workbook[sheet_name]
        epoch = workbook.epoch

        last_col = 4  # A-D
        rows_iter = _iter_rows_values(sheet, last_col=last_col)
        header = next(rows_iter, None)
        if not header:
            raise ValueError(f"{source_path.name}：未找到标题行")

        header_a = _validate_header_cell(header[0])
        _ = _validate_header_cell(header[3])
        if "日期" not in header_a and header_a.lower() != "date":
            raise ValueError(f"{source_path.name}：A1 标题应为“日期”")

        rows: list[tuple[dt.date, float]] = []
        for _, values in enumerate(rows_iter, start=2):
            try:
                date = _parse_date(values[0], epoch=epoch)
                ratio = _coerce_float(values[3])
                rows.append((date, ratio))
            except Exception:
                continue

        if not rows:
            raise ValueError(f"{source_path.name}：清洗后没有可用数据行")

        rows.sort(key=lambda item: item[0])
        dates = [date.isoformat() for date, _ in rows]
        metrics = [metric for _, metric in rows]
        return dates, metrics
    finally:
        workbook.close()


def _load_erp_series() -> tuple[list[str], list[float], list[float], list[float], list[float]]:
    pe_path = _find_input_xlsx("data_PE")
    bond_path = _find_input_xlsx("data_bond")

    pe_rows = _process_data_pe(pe_path)
    bond_rows = _process_data_bond(bond_path)
    merged_rows = _merge_by_bond_dates(bond_rows, pe_rows)
    erp_rows = _compute_erp_rows(merged_rows, bond_rows)

    dates: list[str] = []
    erp_values: list[float] = []
    bond_yield_values: list[float] = []
    pe_values: list[float] = []
    close_values: list[float] = []
    for row_index, row in enumerate(erp_rows[1:], start=2):
        try:
            date_text = str(row[0])
            _ = dt.date.fromisoformat(date_text)
            value = row[4]
            if not isinstance(value, (int, float)):
                raise ValueError("数值类型不合法")
            dates.append(date_text)
            erp_values.append(float(value))

            yield_value = row[1]
            if not isinstance(yield_value, (int, float)):
                raise ValueError("十年期收益率类型不合法")
            bond_yield_values.append(float(yield_value))

            pe_value = row[2]
            if not isinstance(pe_value, (int, float)):
                raise ValueError("PE 类型不合法")
            pe_values.append(float(pe_value))

            close_value = row[3]
            if not isinstance(close_value, (int, float)):
                raise ValueError("收盘点位类型不合法")
            close_values.append(float(close_value))
        except Exception as exc:
            raise ValueError(f"ERP 第 {row_index} 行数据不合法：{exc}") from exc

    if not dates:
        raise ValueError("ERP 数据为空")
    return dates, erp_values, bond_yield_values, pe_values, close_values


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


def _process_ratio_file(source_path: Path, *, metric_header: str) -> list[list[object]]:
    workbook = openpyxl.load_workbook(source_path, data_only=True, read_only=True)
    sheet_name = workbook.sheetnames[0] if workbook.sheetnames else None
    if not sheet_name:
        raise ValueError(f"{source_path.name}：未找到可用工作表")
    sheet = workbook[sheet_name]

    last_col = 4  # A-D
    rows_iter = _iter_rows_values(sheet, last_col=last_col)
    header = next(rows_iter, None)
    if not header:
        raise ValueError(f"{source_path.name}：未找到标题行")

    header_a = _validate_header_cell(header[0])
    _ = _validate_header_cell(header[3])
    if "日期" not in header_a and header_a.lower() != "date":
        raise ValueError(f"{source_path.name}：A1 标题应为“日期”")

    rows: list[tuple[dt.date, float]] = []
    for row_index, values in enumerate(rows_iter, start=2):
        try:
            date = _parse_date(values[0], epoch=workbook.epoch)
            ratio = _coerce_float(values[3])
            rows.append((date, ratio))
        except Exception:
            continue

    if not rows:
        raise ValueError(f"{source_path.name}：清洗后没有可用数据行")

    rows.sort(key=lambda item: item[0])
    output: list[list[object]] = [[header_a, metric_header]]
    output.extend([[date.isoformat(), ratio] for date, ratio in rows])
    return output

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


def _compute_erp_rolling_bands(
    erp_rows: list[list[object]],
    *,
    window_size: int = 2000,
) -> list[list[object]]:
    if not isinstance(window_size, int) or window_size <= 0:
        raise ValueError("滚动窗口 n 必须为正整数")
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
    queue: deque[float] = deque()
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
            leaving = queue.popleft()
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


def _compute_erp_interval_bands(
    erp_rows: list[list[object]],
    *,
    start_date: dt.date,
    end_date: dt.date,
) -> tuple[dt.date, dt.date, dt.date, dt.date, list[list[object]], float, float]:
    if not erp_rows or len(erp_rows) < 2:
        raise ValueError("ERP 数据为空")

    header = erp_rows[0]
    if len(header) < 5 or header[4] != "股权风险溢价":
        raise ValueError("ERP 表头不符合预期")

    data_rows = erp_rows[1:]
    dates: list[dt.date] = []
    for index, row in enumerate(data_rows, start=2):
        try:
            dates.append(dt.date.fromisoformat(str(row[0])))
        except ValueError as exc:
            raise ValueError(f"ERP 第 {index} 行日期无法解析") from exc

    if not dates:
        raise ValueError("ERP 数据为空")

    earliest = dates[0]
    latest = dates[-1]
    if start_date < earliest:
        raise ValueError(f"起始日期过早：最早日期为 {earliest.isoformat()}")
    if start_date > latest:
        raise ValueError(f"起始日期过晚：最近日期为 {latest.isoformat()}")
    if end_date < earliest:
        raise ValueError(f"终止日期过早：最早日期为 {earliest.isoformat()}")
    if end_date > latest:
        raise ValueError(f"终止日期过晚：最近日期为 {latest.isoformat()}")

    start_index = bisect_left(dates, start_date)
    if start_index >= len(dates):
        raise ValueError(f"起始日期过晚：最近日期为 {latest.isoformat()}")

    end_index = bisect_right(dates, end_date) - 1
    if end_index < 0:
        raise ValueError(f"终止日期过早：最早日期为 {earliest.isoformat()}")

    actual_start = dates[start_index]
    actual_end = dates[end_index]
    if actual_start > actual_end:
        raise ValueError("起始日期不能晚于终止日期（自动调整后）")

    interval_rows = data_rows[start_index : end_index + 1]
    if not interval_rows:
        raise ValueError("区间内没有数据")

    erp_values: list[float] = []
    sum_values = 0.0
    sum_squares = 0.0
    for index, row in enumerate(interval_rows, start=start_index + 2):
        value = row[4]
        if not isinstance(value, (int, float)):
            raise ValueError(f"ERP 第 {index} 行数值类型不合法")
        value_float = float(value)
        erp_values.append(value_float)
        sum_values += value_float
        sum_squares += value_float * value_float

    sorted_values = sorted(erp_values)
    median = _rolling_median(sorted_values)
    stddevp = _rolling_stddevp(sum_values, sum_squares, len(erp_values))
    upper2 = median + 2 * stddevp
    upper1 = median + stddevp
    lower1 = median - stddevp
    lower2 = median - 2 * stddevp

    output: list[list[object]] = [
        ["日期", "十年期收益率", "PE-TTM-S", "收盘点位", "股权风险溢价", "+2σ", "+1σ", "中位数", "-1σ", "-2σ"]
    ]
    for row in interval_rows:
        output.append([row[0], row[1], row[2], row[3], row[4], upper2, upper1, median, lower1, lower2])

    return earliest, latest, actual_start, actual_end, output, median, stddevp


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

        bands_rows = _compute_erp_rolling_bands(erp_rows, window_size=2000)

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


@app.post("/api/erprolling")
def generate_erp_rolling() -> object:
    payload = request.get_json(silent=True) or {}
    n = payload.get("n")

    try:
        if isinstance(n, str):
            try:
                n = int(n.strip())
            except ValueError as exc:
                raise ValueError("n 必须为整数") from exc
        if not isinstance(n, int):
            raise ValueError("n 必须为整数")
        if n < 1 or n > 4000:
            raise ValueError("n 超出范围（1-4000）")

        pe_path = _find_input_xlsx("data_PE")
        bond_path = _find_input_xlsx("data_bond")

        pe_rows = _process_data_pe(pe_path)
        bond_rows = _process_data_bond(bond_path)
        merged_rows = _merge_by_bond_dates(bond_rows, pe_rows)
        erp_rows = _compute_erp_rows(merged_rows, bond_rows)

        bands_rows = _compute_erp_rolling_bands(erp_rows, window_size=n)

        csv_name = "ERP_Rolling Calculation.csv"
        xlsx_name = "ERP_Rolling Calculation.xlsx"
        _write_csv(bands_rows, OUTPUT_DIR / csv_name)
        _write_xlsx(bands_rows, OUTPUT_DIR / xlsx_name, "ERP_Rolling Calculation")

        return jsonify({"output_csv": csv_name, "output_xlsx": xlsx_name, "n": n})
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:  # pragma: no cover
        return jsonify({"error": f"生成失败：{exc}"}), 500


@app.post("/api/erpinterval")
def generate_erp_interval() -> object:
    payload = request.get_json(silent=True) or {}
    start_date_raw = payload.get("start_date")
    end_date_raw = payload.get("end_date")

    try:
        if not isinstance(start_date_raw, str) or not start_date_raw.strip():
            raise ValueError("缺少起始日期 start_date")
        try:
            start_date = dt.date.fromisoformat(start_date_raw.strip())
        except ValueError as exc:
            raise ValueError("起始日期格式必须为 YYYY-MM-DD") from exc

        if end_date_raw is None or (isinstance(end_date_raw, str) and not end_date_raw.strip()):
            end_date = dt.date.today()
        else:
            if not isinstance(end_date_raw, str):
                raise ValueError("终止日期格式必须为 YYYY-MM-DD")
            try:
                end_date = dt.date.fromisoformat(end_date_raw.strip())
            except ValueError as exc:
                raise ValueError("终止日期格式必须为 YYYY-MM-DD") from exc

        pe_path = _find_input_xlsx("data_PE")
        bond_path = _find_input_xlsx("data_bond")

        pe_rows = _process_data_pe(pe_path)
        bond_rows = _process_data_bond(bond_path)
        merged_rows = _merge_by_bond_dates(bond_rows, pe_rows)
        erp_rows = _compute_erp_rows(merged_rows, bond_rows)

        earliest, latest, actual_start, actual_end, output_rows, median, stddevp = _compute_erp_interval_bands(
            erp_rows, start_date=start_date, end_date=end_date
        )

        csv_name = "ERP_Interval.csv"
        xlsx_name = "ERP_Interval.xlsx"
        _write_csv(output_rows, OUTPUT_DIR / csv_name)
        _write_xlsx(output_rows, OUTPUT_DIR / xlsx_name, "ERP_Interval")

        adjusted = actual_start != start_date
        adjusted_end = actual_end != end_date
        return jsonify(
            {
                "output_csv": csv_name,
                "output_xlsx": xlsx_name,
                "input_start_date": start_date.isoformat(),
                "used_start_date": actual_start.isoformat(),
                "input_end_date": end_date.isoformat(),
                "used_end_date": actual_end.isoformat(),
                "earliest_date": earliest.isoformat(),
                "latest_date": latest.isoformat(),
                "adjusted_to_trading_day": adjusted,
                "adjusted_end_to_trading_day": adjusted_end,
                "median": median,
                "stddevp": stddevp,
            }
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:  # pragma: no cover
        return jsonify({"error": f"生成失败：{exc}"}), 500


@app.post("/api/thermometer/clean")
def generate_thermometer_clean() -> object:
    try:
        gdp_path = _find_input_xlsx("data_Ratio GDP")
        volume_path = _find_input_xlsx("data_Ratio Volume")
        lend_path = _find_input_xlsx("data_Ratio Securities Lend")

        gdp_rows = _process_ratio_file(gdp_path, metric_header="总市值/GDP")
        volume_rows = _process_ratio_file(volume_path, metric_header="成交量/总市值")
        lend_rows = _process_ratio_file(lend_path, metric_header="融资融券/总市值")

        outputs = {
            "ratio_gdp": "Ratio_GDP.csv",
            "ratio_volume": "Ratio_Volume.csv",
            "ratio_securities_lend": "Ratio_Securities_Lend.csv",
        }

        _write_csv(gdp_rows, OUTPUT_DIR / outputs["ratio_gdp"])
        _write_csv(volume_rows, OUTPUT_DIR / outputs["ratio_volume"])
        _write_csv(lend_rows, OUTPUT_DIR / outputs["ratio_securities_lend"])

        return jsonify(
            {
                "outputs": {
                    "ratio_gdp_csv": outputs["ratio_gdp"],
                    "ratio_volume_csv": outputs["ratio_volume"],
                    "ratio_securities_lend_csv": outputs["ratio_securities_lend"],
                }
            }
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:  # pragma: no cover
        return jsonify({"error": f"生成失败：{exc}"}), 500


@app.post("/api/thermometer/percentiles")
def generate_thermometer_percentiles() -> object:
    payload = request.get_json(silent=True) or {}

    def get_int(name: str, *, min_value: int, max_value: int) -> int:
        raw = payload.get(name)
        if isinstance(raw, str):
            raw = raw.strip()
            if not raw:
                raise ValueError(f"缺少参数：{name}")
            try:
                raw = int(raw)
            except ValueError as exc:
                raise ValueError(f"{name} 必须为整数") from exc
        if not isinstance(raw, int):
            raise ValueError(f"{name} 必须为整数")
        if raw < min_value or raw > max_value:
            raise ValueError(f"{name} 超出范围（{min_value}-{max_value}）")
        return raw

    try:
        ma_gdp = get_int("moving_average_gdp", min_value=1, max_value=1000)
        rp_gdp = get_int("rolling_period_gdp", min_value=1, max_value=1000)
        ma_volume = get_int("moving_average_volume", min_value=1, max_value=4000)
        rp_volume = get_int("rolling_period_volume", min_value=1, max_value=4000)
        ma_securities = get_int("moving_average_securities", min_value=1, max_value=4000)
        rp_securities = get_int("rolling_period_securities", min_value=1, max_value=4000)
        ma_erp = get_int("moving_erp", min_value=1, max_value=4000)
        rp_erp = get_int("rolling_period_erp", min_value=1, max_value=4000)

        gdp_path = _find_input_xlsx("data_Ratio GDP")
        volume_path = _find_input_xlsx("data_Ratio Volume")
        lend_path = _find_input_xlsx("data_Ratio Securities Lend")

        gdp_dates, gdp_values = _load_ratio_series(gdp_path)
        vol_dates, vol_values = _load_ratio_series(volume_path)
        sec_dates, sec_values = _load_ratio_series(lend_path)
        erp_dates, erp_values, erp_yields, erp_pes, erp_closes = _load_erp_series()

        def build_output(
            dates: list[str],
            values: list[float],
            *,
            metric_header: str,
            ma_window: int,
            rp_window: int,
        ) -> list[list[object]]:
            ma_values = _moving_average(values, ma_window)
            pct_values = _rolling_percentiles(ma_values, rp_window)
            out: list[list[object]] = [["日期", metric_header, "平均移动", "分位"]]
            for index, date_text in enumerate(dates):
                if pct_values[index] is None:
                    continue
                out.append([date_text, values[index], ma_values[index], pct_values[index]])
            return out

        gdp_out = build_output(
            gdp_dates,
            gdp_values,
            metric_header="总市值/GDP",
            ma_window=ma_gdp,
            rp_window=rp_gdp,
        )
        vol_out = build_output(
            vol_dates,
            vol_values,
            metric_header="成交量/总市值",
            ma_window=ma_volume,
            rp_window=rp_volume,
        )
        sec_out = build_output(
            sec_dates,
            sec_values,
            metric_header="融资融券/总市值",
            ma_window=ma_securities,
            rp_window=rp_securities,
        )
        erp_ma_values = _moving_average(erp_values, ma_erp)
        erp_pct_values = _rolling_percentiles(erp_ma_values, rp_erp)
        erp_out: list[list[object]] = [
            ["日期", "股权风险溢价", "平均移动", "分位", "十年期收益率", "PE-TTM-S", "收盘点位"]
        ]
        for index, date_text in enumerate(erp_dates):
            if erp_pct_values[index] is None:
                continue
            erp_out.append(
                [
                    date_text,
                    erp_values[index],
                    erp_ma_values[index],
                    erp_pct_values[index],
                    erp_yields[index],
                    erp_pes[index],
                    erp_closes[index],
                ]
            )

        outputs = {
            "ratio_gdp": "Ratio_GDP_Percentile.csv",
            "ratio_volume": "Ratio_Volume_Percentile.csv",
            "ratio_securities_lend": "Ratio_Securities_Lend_Percentile.csv",
            "erp": "ERP_Percentile.csv",
        }

        _write_csv(gdp_out, OUTPUT_DIR / outputs["ratio_gdp"])
        _write_csv(vol_out, OUTPUT_DIR / outputs["ratio_volume"])
        _write_csv(sec_out, OUTPUT_DIR / outputs["ratio_securities_lend"])
        _write_csv(erp_out, OUTPUT_DIR / outputs["erp"])

        return jsonify(
            {
                "outputs": {
                    "ratio_gdp_csv": outputs["ratio_gdp"],
                    "ratio_volume_csv": outputs["ratio_volume"],
                    "ratio_securities_lend_csv": outputs["ratio_securities_lend"],
                    "erp_csv": outputs["erp"],
                }
            }
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:  # pragma: no cover
        return jsonify({"error": f"生成失败：{exc}"}), 500


if __name__ == "__main__":
    debug = os.environ.get("DP_DEBUG") == "1"
    app.run(host="127.0.0.1", port=5000, debug=debug, use_reloader=False)
