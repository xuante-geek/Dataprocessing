#!/usr/bin/env python3
from __future__ import annotations

import datetime as dt
import logging
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import app  # noqa: E402


LOG_DIR = ROOT / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "autorun.log"
LOGIN_REQUIRED_EXIT = 10


class LoginRequiredError(RuntimeError):
    pass


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def run_download_all() -> None:
    cfg = app._download_load_config()
    download_dir, user_data_dir = app._download_ensure_dirs(cfg)
    tasks = list(cfg.get("tasks", []))

    if app.sync_playwright is None:
        raise RuntimeError("缺少依赖：playwright")
    if not tasks:
        raise RuntimeError("下载任务为空")

    # 后台运行：强制使用 headless
    cfg["headless"] = True
    # 后台模式下登录不可交互，缩短等待时间
    app.DOWNLOAD_LOGIN_WAIT_SECONDS = 5
    # 后台模式下不使用“登录按钮”判定，避免误判强制切前端
    app.DOWNLOAD_LOGIN_STRICT = False

    if not app._download_acquire_lock():
        raise RuntimeError("下载任务锁定中，请稍后再试")

    try:
        with app.sync_playwright() as p:
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(user_data_dir),
                headless=cfg.get("headless", False),
                accept_downloads=True,
                channel="chrome",
            )
            for task in tasks:
                task_url = task.get("url")
                if not task_url:
                    raise RuntimeError(f"任务缺少 URL：{task.get('id')}")
                try:
                    result = app._download_perform_task(task, task_url, context, download_dir)
                except app.DownloadLoginAbort as exc:
                    raise LoginRequiredError(str(exc)) from exc
                except RuntimeError as exc:
                    if "需要登录" in str(exc):
                        raise LoginRequiredError(str(exc)) from exc
                    raise
                validation = result.get("validation", {})
                if not validation.get("ok", True):
                    err_msg = "; ".join(validation.get("errors", [])) or "校验失败"
                    raise RuntimeError(f"{task.get('name', task.get('id'))} 校验失败：{err_msg}")
            context.close()
    finally:
        app._download_release_lock()


def generate_erp_full(include_yield: bool, include_pe: bool, include_erp_percentile: bool) -> None:
    pe_path = app._find_input_xlsx("data_PE")
    bond_path = app._find_input_xlsx("data_bond")

    pe_rows = app._process_data_pe(pe_path)
    bond_rows = app._process_data_bond(bond_path)
    merged_rows = app._merge_by_bond_dates(bond_rows, pe_rows)

    pe_clean_rows: list[list[object]] = [["日期", "PE-TTM-S", "全A点位"]] + [
        [date.isoformat(), pe, close] for date, pe, close in pe_rows
    ]
    bond_clean_rows: list[list[object]] = [["日期", "十年国债收益率"]] + [
        [date.isoformat(), yield_raw] for date, yield_raw, _ in bond_rows
    ]
    merged_clean_rows: list[list[object]] = [["日期", "十年国债收益率", "PE-TTM-S", "全A点位"]] + [
        [date.isoformat(), yield_raw, pe, close] for date, yield_raw, pe, close in merged_rows
    ]
    erp_rows = app._compute_erp_rows(merged_rows, bond_rows)

    drop_names: set[str] = set()
    if not include_yield:
        drop_names.add("十年国债收益率")
    if not include_pe:
        drop_names.add("PE-TTM-S")
    if not include_erp_percentile:
        drop_names.add("股权风险溢价分位")

    pe_clean_rows = app._filter_columns(pe_clean_rows, drop_names)
    bond_clean_rows = app._filter_columns(bond_clean_rows, drop_names)
    merged_clean_rows = app._filter_columns(merged_clean_rows, drop_names)
    erp_rows = app._filter_columns(erp_rows, drop_names)

    app._write_csv(pe_clean_rows, app.OUTPUT_DIR / "data_PE_clean.csv")
    app._write_csv(bond_clean_rows, app.OUTPUT_DIR / "data_bond_clean.csv")
    app._write_csv(merged_clean_rows, app.OUTPUT_DIR / "merged.csv")
    app._write_csv(erp_rows, app.OUTPUT_DIR / "ERP.csv")


def generate_erp_rolling(n: int, include_yield: bool, include_pe: bool, include_erp_percentile: bool) -> None:
    pe_path = app._find_input_xlsx("data_PE")
    bond_path = app._find_input_xlsx("data_bond")

    pe_rows = app._process_data_pe(pe_path)
    bond_rows = app._process_data_bond(bond_path)
    merged_rows = app._merge_by_bond_dates(bond_rows, pe_rows)
    erp_rows = app._compute_erp_rows(merged_rows, bond_rows)

    bands_rows = app._compute_erp_rolling_bands(erp_rows, window_size=n, include_percentile=include_erp_percentile)

    drop_names: set[str] = set()
    if not include_yield:
        drop_names.add("十年国债收益率")
    if not include_pe:
        drop_names.add("PE-TTM-S")
    if not include_erp_percentile:
        drop_names.add("股权风险溢价分位")
    bands_rows = app._filter_columns(bands_rows, drop_names)

    app._write_csv(bands_rows, app.OUTPUT_DIR / "ERP_Rolling Calculation.csv")


def generate_erp_interval(
    start_date: dt.date,
    end_date: dt.date | None,
    include_yield: bool,
    include_pe: bool,
    include_erp_percentile: bool,
) -> None:
    pe_path = app._find_input_xlsx("data_PE")
    bond_path = app._find_input_xlsx("data_bond")

    pe_rows = app._process_data_pe(pe_path)
    bond_rows = app._process_data_bond(bond_path)
    merged_rows = app._merge_by_bond_dates(bond_rows, pe_rows)
    erp_rows = app._compute_erp_rows(merged_rows, bond_rows)

    _, _, _, _, output_rows, _, _ = app._compute_erp_interval_bands(
        erp_rows, start_date=start_date, end_date=end_date
    )

    drop_names: set[str] = set()
    if not include_yield:
        drop_names.add("十年国债收益率")
    if not include_pe:
        drop_names.add("PE-TTM-S")
    if not include_erp_percentile:
        drop_names.add("股权风险溢价分位")
    output_rows = app._filter_columns(output_rows, drop_names)

    csv_path = app.OUTPUT_DIR / "ERP_Interval.csv"
    app._write_csv(output_rows, csv_path)
    app._publish_csv_to_cos(csv_path, "ERP_Interval.csv")


def generate_thermo_percentiles(
    ma_gdp: int,
    rp_gdp: int,
    internal_gdp_mode: str,
    internal_gdp: int | None,
    ma_volume: int,
    rp_volume: int,
    internal_volume_mode: str,
    internal_volume: int | None,
    ma_securities: int,
    rp_securities: int,
    internal_securities_mode: str,
    internal_securities: int | None,
    ma_erp: int,
    rp_erp: int,
    internal_erp_mode: str,
    internal_erp: int | None,
) -> None:
    gdp_path = app._find_input_xlsx("data_Ratio GDP")
    volume_path = app._find_input_xlsx("data_Ratio Volume")
    lend_path = app._find_input_xlsx("data_Ratio Securities Lend")

    gdp_dates, gdp_values = app._load_ratio_series(gdp_path)
    vol_dates, vol_values = app._load_ratio_series(volume_path)
    sec_dates, sec_values = app._load_ratio_series(lend_path)
    erp_dates, erp_values, erp_yields, erp_pes, erp_closes = app._load_erp_series()

    def build_output(
        dates: list[str],
        values: list[float],
        *,
        metric_header: str,
        ma_window: int,
        rp_window: int,
        internal_mode: str,
        min_window: int | None,
    ) -> list[list[object]]:
        ma_values = app._moving_average(values, ma_window)
        if internal_mode == "auto":
            pct_values = app._rolling_percentiles(ma_values, rp_window)
        else:
            if min_window is None:
                raise ValueError("最小递减滚动周期不能为空")
            pct_values = app._rolling_percentiles_with_min_window(ma_values, rp_window, min_window)
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
        internal_mode=internal_gdp_mode,
        min_window=internal_gdp,
    )
    gdp_out = app._keep_weekly_latest_rows(gdp_out)
    vol_out = build_output(
        vol_dates,
        vol_values,
        metric_header="成交量/总市值",
        ma_window=ma_volume,
        rp_window=rp_volume,
        internal_mode=internal_volume_mode,
        min_window=internal_volume,
    )
    sec_out = build_output(
        sec_dates,
        sec_values,
        metric_header="融资融券/总市值",
        ma_window=ma_securities,
        rp_window=rp_securities,
        internal_mode=internal_securities_mode,
        min_window=internal_securities,
    )

    erp_ma_values = app._moving_average(erp_values, ma_erp)
    if internal_erp_mode == "auto":
        erp_pct_values = app._rolling_percentiles(erp_ma_values, rp_erp)
    else:
        if internal_erp is None:
            raise ValueError("股权风险溢价最小递减滚动周期不能为空")
        erp_pct_values = app._rolling_percentiles_with_min_window(erp_ma_values, rp_erp, internal_erp)

    erp_out: list[list[object]] = [
        ["日期", "股权风险溢价", "平均移动", "分位", "十年国债收益率", "PE-TTM-S", "全A点位"]
    ]
    for index, date_text in enumerate(erp_dates):
        if erp_pct_values[index] is None:
            continue
        erp_out.append(
            [
                date_text,
                erp_values[index],
                erp_ma_values[index],
                round(float(erp_pct_values[index]), 1),
                erp_yields[index],
                erp_pes[index],
                erp_closes[index],
            ]
        )

    app._write_csv(gdp_out, app.OUTPUT_DIR / "Ratio_GDP_Percentile.csv")
    app._write_csv(vol_out, app.OUTPUT_DIR / "Ratio_Volume_Percentile.csv")
    app._write_csv(sec_out, app.OUTPUT_DIR / "Ratio_Securities_Lend_Percentile.csv")
    app._write_csv(erp_out, app.OUTPUT_DIR / "ERP_Percentile.csv")


def generate_thermo_merge(
    ma_gdp: int,
    rp_gdp: int,
    internal_gdp_mode: str,
    internal_gdp: int | None,
    ma_volume: int,
    rp_volume: int,
    internal_volume_mode: str,
    internal_volume: int | None,
    ma_securities: int,
    rp_securities: int,
    internal_securities_mode: str,
    internal_securities: int | None,
    ma_erp: int,
    rp_erp: int,
    internal_erp_mode: str,
    internal_erp: int | None,
    weight_gdp: float,
    weight_volume: float,
    weight_securities: float,
    weight_erp: float,
    include_gdp: bool,
    include_volume: bool,
    include_securities: bool,
    include_erp: bool,
    include_yield: bool,
) -> None:
    gdp_path = app._find_input_xlsx("data_Ratio GDP")
    volume_path = app._find_input_xlsx("data_Ratio Volume")
    lend_path = app._find_input_xlsx("data_Ratio Securities Lend")

    gdp_dates, gdp_values = app._load_ratio_series(gdp_path)
    vol_dates, vol_values = app._load_ratio_series(volume_path)
    sec_dates, sec_values = app._load_ratio_series(lend_path)
    erp_dates, erp_values, erp_yields, _, erp_closes = app._load_erp_series()

    gdp_records_full = app._build_percentile_records(
        gdp_dates,
        gdp_values,
        ma_window=ma_gdp,
        rp_window=rp_gdp,
        internal_mode=internal_gdp_mode,
        min_window=internal_gdp,
    )
    vol_records = app._build_percentile_records(
        vol_dates,
        vol_values,
        ma_window=ma_volume,
        rp_window=rp_volume,
        internal_mode=internal_volume_mode,
        min_window=internal_volume,
    )
    sec_records = app._build_percentile_records(
        sec_dates,
        sec_values,
        ma_window=ma_securities,
        rp_window=rp_securities,
        internal_mode=internal_securities_mode,
        min_window=internal_securities,
    )
    erp_records = app._build_erp_percentile_records(
        erp_dates,
        erp_values,
        erp_yields,
        erp_closes,
        ma_window=ma_erp,
        rp_window=rp_erp,
        internal_mode=internal_erp_mode,
        min_window=internal_erp,
    )

    if not (gdp_records_full and vol_records and sec_records and erp_records):
        raise ValueError("数据不足：请检查移动平均与滚动周期参数是否过大")

    vol_start = vol_records[0][0]
    sec_start = sec_records[0][0]
    erp_start = erp_records[0]["date"]

    vol_end = vol_records[-1][0]
    sec_end = sec_records[-1][0]
    erp_end = erp_records[-1]["date"]
    gdp_end_full = gdp_records_full[-1][0]
    date_end = min(gdp_end_full, vol_end, sec_end, erp_end)

    gdp_records_filtered = [record for record in gdp_records_full if record[0] <= date_end]
    gdp_records = app._keep_weekly_latest_records(gdp_records_filtered)

    date_begin = max(vol_start, sec_start, erp_start)
    gdp_dates_only = [d for d, _ in gdp_records]
    gdp_start_index = app._nearest_index(gdp_dates_only, date_begin)
    gdp_end_index = app.bisect_right(gdp_dates_only, date_end) - 1
    if gdp_end_index < gdp_start_index:
        raise ValueError("合并失败：有效时间区间为空")

    vol_dates_only = [d for d, _ in vol_records]
    sec_dates_only = [d for d, _ in sec_records]
    erp_dates_only = [record["date"] for record in erp_records]

    header = ["日期", "股权风险溢价分位", "全A点位", "市场温度"]
    if include_gdp:
        header.insert(1, "市值/GDP分位")
    if include_volume:
        header.insert(2 if include_gdp else 1, "成交量/市值分位")
    if include_securities:
        insert_at = 3 if include_gdp and include_volume else 2 if (include_gdp or include_volume) else 1
        header.insert(insert_at, "融资融券/市值分位")
    if include_erp:
        header.append("股权风险溢价")
    if include_yield:
        header.append("十年国债收益率")

    rows: list[list[object]] = [header]
    one_decimal_columns = {
        "市值/GDP分位",
        "成交量/市值分位",
        "融资融券/市值分位",
        "股权风险溢价分位",
        "市场温度",
        "全A点位",
    }

    def get_pct(records: list[tuple[dt.date, float]], dates_only: list[dt.date], target: dt.date) -> float:
        idx = app._nearest_index(dates_only, target)
        return float(records[idx][1])

    for gdp_idx in range(gdp_start_index, gdp_end_index + 1):
        date_value = gdp_dates_only[gdp_idx]
        gdp_pct = float(gdp_records[gdp_idx][1])
        vol_pct = get_pct(vol_records, vol_dates_only, date_value)
        sec_pct = get_pct(sec_records, sec_dates_only, date_value)

        erp_idx = app._nearest_index(erp_dates_only, date_value)
        erp_record = erp_records[erp_idx]
        erp_pct = float(erp_record["erp_percentile"])
        close_value = float(erp_record["close"])
        erp_value = float(erp_record["erp"])
        yield_value = float(erp_record["yield"])

        temperature = (
            weight_gdp * gdp_pct
            + weight_volume * vol_pct
            + weight_securities * sec_pct
            + weight_erp * (100.0 - erp_pct)
        ) / 100.0

        row_map: dict[str, object] = {
            "日期": date_value.isoformat(),
            "市值/GDP分位": gdp_pct,
            "成交量/市值分位": vol_pct,
            "融资融券/市值分位": sec_pct,
            "股权风险溢价分位": erp_pct,
            "股权风险溢价": erp_value,
            "十年国债收益率": yield_value,
            "全A点位": close_value,
            "市场温度": temperature,
        }
        row_out: list[object] = []
        for col in header:
            value = row_map.get(col, "")
            if col in one_decimal_columns and isinstance(value, (int, float)) and not isinstance(value, bool):
                value = round(float(value), 1)
            row_out.append(value)
        rows.append(row_out)

    output_path = app.OUTPUT_DIR / "Market_Thermometer.csv"
    app._write_csv(rows, output_path)
    app._publish_csv_to_cos(output_path, "Market_Thermometer.csv")


def main() -> int:
    setup_logging()
    logging.info("autorun start")

    # 1) 下载
    try:
        run_download_all()
        logging.info("download complete")
    except LoginRequiredError as exc:
        logging.warning(str(exc))
        return LOGIN_REQUIRED_EXIT
    except Exception as exc:
        logging.exception("download failed: %s", exc)
        return 1

    # 2) ERP（默认可选列关闭）
    include_yield = False
    include_pe = False
    include_erp_percentile = False

    generate_erp_full(include_yield, include_pe, include_erp_percentile)
    logging.info("erp full complete")

    generate_erp_rolling(2400, include_yield, include_pe, include_erp_percentile)
    logging.info("erp rolling complete")

    generate_erp_interval(dt.date(2014, 1, 2), None, include_yield, include_pe, include_erp_percentile)
    logging.info("erp interval complete")

    # 3) 市场温度计（参数与 UI 默认一致）
    ma_gdp = 1
    rp_gdp = 500
    internal_gdp_mode = "auto"
    internal_gdp = None

    ma_volume = 20
    rp_volume = 1000
    internal_volume_mode = "auto"
    internal_volume = None

    ma_securities = 1
    rp_securities = 1000
    internal_securities_mode = "auto"
    internal_securities = None

    ma_erp = 1
    rp_erp = 2400
    internal_erp_mode = "custom"
    internal_erp = 2000

    generate_thermo_percentiles(
        ma_gdp,
        rp_gdp,
        internal_gdp_mode,
        internal_gdp,
        ma_volume,
        rp_volume,
        internal_volume_mode,
        internal_volume,
        ma_securities,
        rp_securities,
        internal_securities_mode,
        internal_securities,
        ma_erp,
        rp_erp,
        internal_erp_mode,
        internal_erp,
    )
    logging.info("thermo percentiles complete")

    weight_gdp = 34
    weight_volume = 33
    weight_securities = 33
    weight_erp = 0

    include_gdp = False
    include_volume = False
    include_securities = False
    include_erp = False
    include_yield = False

    generate_thermo_merge(
        ma_gdp,
        rp_gdp,
        internal_gdp_mode,
        internal_gdp,
        ma_volume,
        rp_volume,
        internal_volume_mode,
        internal_volume,
        ma_securities,
        rp_securities,
        internal_securities_mode,
        internal_securities,
        ma_erp,
        rp_erp,
        internal_erp_mode,
        internal_erp,
        weight_gdp,
        weight_volume,
        weight_securities,
        weight_erp,
        include_gdp,
        include_volume,
        include_securities,
        include_erp,
        include_yield,
    )
    logging.info("thermo merge complete")
    logging.info("autorun success")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
