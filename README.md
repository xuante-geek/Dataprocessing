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
- 所有数值最多保留小数点后 4 位（导出时统一四舍五入）

## Feature 3：ERP_10Year

按钮“生成 ERP_10Year（Feature 3）”会基于 ERP 数据，使用滚动 2000 个交易日（约 10 年）计算：
- 中位数（MEDIAN）
- 总体标准差（STDEVP）

并生成 5 条布林带列：`+2σ`、`+1σ`、`中位数`、`-1σ`、`-2σ`，输出到 `docs/data/ERP_10Year.csv`。
