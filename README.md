# DataProcessing

本地批处理工具（Feature 1）：从 `input/` 读取 Excel（`.xlsx`），导出同名 CSV 到 `docs/data/`。

处理逻辑（当前阶段）：
- 识别第一行为标题行、第一列为日期列
- 校验保留列中是否存在空白/乱码/非法类型（存在则中断并提示单元格坐标）
- 删除 B/C/D 列
- 按第一列日期从远到近排序（旧→新）
- 导出 `CSV`，并额外生成冻结首行/首列的处理后 `Excel`

## 运行

1. 安装依赖：`pip install -r requirements.txt`
2. 启动：`python src/app.py`
3. 打开：`http://127.0.0.1:5000`

## 目录约定

- 输入：`input/`
- 输出：`docs/data/`

## Feature 2：ERP

按钮“生成 ERP（Feature 2）”会自动读取 `input/data_PE.xlsx` 和 `input/data_bond.xlsx`，完成清洗、对齐合并、计算 ERP，并输出到 `docs/data/`（含 `ERP.csv`）。

清洗规则（当前阶段）：
- `data_PE.xlsx`：保留 `日期`、`PE-TTM-S`、`收盘点位`；其中 `2018-08-03` 至 `2018-08-24` 的 `收盘点位` 缺失会按内置清单补齐
- `data_bond.xlsx`：保留 `日期`、`十年期收益率`
- 除上述补齐以外：若仍存在空白/乱码/非法类型单元格，会中断并提示具体单元格坐标

输出格式（当前阶段）：
- 所有数值最多保留小数点后 6 位（导出时统一四舍五入）

## Feature 3：ERP_10Year

按钮“生成 ERP_10Year（Feature 3）”会基于 ERP 数据，使用滚动 2000 个交易日（约 10 年）计算：
- 中位数（MEDIAN）
- 总体标准差（STDEVP）

并生成 5 条布林带列：`+2σ`、`+1σ`、`中位数`、`-1σ`、`-2σ`，输出到 `docs/data/ERP_10Year.csv`。

## Feature 4：ERP_Rolling Calculation

页面提供输入框 `n`（范围 1-4000），用于按滚动 `n` 个交易日计算布林带，导出到：
- `docs/data/ERP_Rolling Calculation.csv`

## Feature 5：ERP_Interval

输入固定区间起始日期与终止日期（`YYYY-MM-DD`）：
- 起始日期若为非交易日会自动顺延到下一个交易日
- 终止日期若为非交易日会自动回退到上一个交易日

程序会计算区间 ERP 中位数和总体标准差，并输出水平布林线到：
- `docs/data/ERP_Interval.csv`

## Feature 7：市场温度计（数据清洗）

读取 `input/` 中以下文件并清洗后导出（会删除含空白/非数值/乱码的行，并保留 `日期` + 指标列）：
- `data_Ratio GDP.xlsx` → `docs/data/Ratio_GDP.csv`
- `data_Ratio Volume.xlsx` → `docs/data/Ratio_Volume.csv`
- `data_Ratio Securities Lend .xlsx` → `docs/data/Ratio_Securities_Lend.csv`

## Feature 8：市场温度计（分位）

基于温度计三份指标与 ERP，支持对“平均移动”后的数值做滚动分位计算，并输出 4 份 CSV：
- `docs/data/Ratio_GDP_Percentile.csv`
- `docs/data/Ratio_Volume_Percentile.csv`
- `docs/data/Ratio_Securities_Lend_Percentile.csv`
- `docs/data/ERP_Percentile.csv`

说明：分位列为空白的行会被删除（仅保留分位已计算完成的行）。

`ERP_Percentile.csv` 额外包含 `十年期收益率`、`PE-TTM-S`、`收盘点位` 三列，便于后续关联使用。

## Feature 9：市场温度计（合并与温度）

以 `市值/GDP` 的周频日期为基准，对齐并合并四个分位因子，并按权重计算市场温度，导出到：
- `docs/data/Market_Thermometer.csv`

温度计算：`市场温度 = (W_GDP*T1 + W_Volume*T2 + W_Securities*T3 + W_ERP*(100-T4)) / 100`（ERP 分位为反向指标）。

`Market_Thermometer.csv` 默认对分位、市场温度、全A点位等列做 1 位小数输出，便于展示。
