# DataProcessing

本地批处理工具（Feature 1）：从 `input/` 读取 Excel（`.xlsx`），导出同名 CSV 到 `docs/data/`。

处理逻辑（当前阶段）：
- 识别第一行为标题行、第一列为日期列
- 校验保留列中是否存在空白/乱码/非法类型（存在则中断并提示单元格坐标）
- 删除 B/C/D 列
- 按第一列日期从远到近排序（旧→新）
- 导出 `CSV`，并额外生成冻结首行/首列的处理后 `Excel`

## 运行（V3_Feature1 已合并 DataDownload）

1. 固定 Python 主版本（两台机器统一）：
   - `python3 --version` 应为 `3.9.6`
   - 项目已提供 `.python-version`，建议配合 `pyenv` 使用
2. 创建并激活虚拟环境（统一使用 `.venv`）：
   - `python3.9 -m venv .venv`
   - `source .venv/bin/activate`
3. 安装依赖（固定版本）：
   - `python -m pip install -r requirements.txt`
4. 安装浏览器内核（仅首次需要）：
   - `python -m playwright install chromium`
5. 启动服务：
   - `python src/app.py`
6. 打开：
   - `http://127.0.0.1:5001`

注意：`scripts/run_ui.sh` 与 `scripts/install_launchd.sh` 会强制检查 `.venv` 和 Python `3.9.6`。

## 目录约定

- 输入：`input/`
- 输出：`docs/data/`
- 下载缓存：`data/user_data/`（用于扫码登录的浏览器持久化）

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

## V3_Feature1：DataDownload 控制台

新增第一个 Tab：**DataDownload 控制台**，用于自动下载 Excel 到 `input/`，减少人工拷贝步骤。

主要特性：
- 任务清单来自 `config/download_config.json`
- 下载结果会覆盖为标准命名（如 `data_PE.xlsx`、`data_bond.xlsx`、`data_Ratio GDP.xlsx`）
- 需要扫码登录时会弹出浏览器窗口，完成扫码后继续执行

常见问题：
- 若提示 `缺少依赖：playwright`，请确保使用同一个 `python3` 执行：
  - `source .venv/bin/activate`
  - `python -m pip install -r requirements.txt`
  - `python -m playwright install chromium`
  - `python src/app.py`

## V3_Feature2：统一 UI 风格（DataDownload 风格）

界面整体视觉统一为 DataDownload 控制台风格：
- 浅蓝背景、白色卡片与更轻的阴影
- 主要按键为蓝色实心，次要按键为白底描边
- 表格、状态胶囊、输入框样式统一为清爽扁平化风格

## V3_Feature3：全部运行

新增右上角“全部运行”按键，顺序依次执行：
1. DataDownload 全部任务
2. 导出完整周期 ERP
3. 导出滚动周期 ERP
4. 导出指定周期 ERP
5. 导出市场温度计分位数据
6. 导出市场温度计总表
7. 导出标普500均线数据
8. 导出纳斯达克均线数据

每一步必须成功后才会触发下一步，任何一步失败将终止并提示错误。

## V3_Feature4：发布 CSV 到远程 COS

在生成本地 CSV 后，程序会自动同步发布到腾讯云 COS（同名覆盖）：
- `ERP_Interval.csv` → `https://anexus-data-1399092305.cos.ap-guangzhou.myqcloud.com/data/ERP_Interval.csv`
- `Market_Thermometer.csv` → `https://anexus-data-1399092305.cos.ap-guangzhou.myqcloud.com/data/Market_Thermometer.csv`

运行前请配置环境变量（必须）：
- `COS_SECRET_ID`
- `COS_SECRET_KEY`

可选环境变量（未配置时使用默认值）：
- `COS_BUCKET`（默认：`anexus-data-1399092305`）
- `COS_REGION`（默认：`ap-guangzhou`）
- `COS_BASE_PATH`（默认：`data`）

发布接口返回中会包含 `remote_url` 字段，用于确认远程地址。

## V3_Feature5：读取 .env

支持从项目根目录的 `.env` 自动读取 COS 配置，避免每次手动 `export`。示例见 `.env.example`。

## V3_Feature6：固定区间终止日期默认最近交易日

ERP 指定区间导出时，终止日期输入框默认空白。若用户未填写，程序会自动使用原始 ERP 数据的最近交易日作为终止日期。若用户手动填写，则按填写日期作为区间终止。

## V3_Feature7：周频对齐优化

`Market_Thermometer` 合并时，GDP 周频去重改为**在共同截止日期内保留当周最新有效值**，避免因周内晚于其它序列的日期而整周被回退（例如本周可用日期是 2/24 时仍可保留 2/24）。

## V3_Feature8：美股指数下载

DataDownload 控制台新增两个任务：
- 标普500：`https://www.touzid.com/indice/fundamental.html#/US.INX`
- 纳斯达克：`https://www.touzid.com/indice/fundamental.html#/US.IXIC`

下载流程均为：选择 `20年` → 粒度切换 `日` → 下载 EXCEL。  
下载后会执行与现有任务一致的校核（时间跨度需 `>=19年`，且日频占比阈值 `>=80%`），通过后输出为：
- `input/data_SP500.xlsx`
- `input/data_NASDAQ.xlsx`

## V3_Feature9：美股指数计算器（清洗 + 均线）

新增第 4 个 Tab：**美股指数计算器**，用于处理 DataDownload 下载后的美股指数 Excel。

包含两个导出按钮：
- 导出标普500均线数据
- 导出纳斯达克均线数据

处理规则（两者一致）：
- 读取 `input/data_SP500.xlsx` 或 `input/data_NASDAQ.xlsx`
- 仅保留 A 列“日期”和 H 列“收盘点位”
- 按日期从远到近排序（旧 → 新）
- 若收盘点位为空白或非数字，则该行删除
- 新增 C 列“参考线”：按均线变量做滚动均值
- C 列为空白的行（前置不足窗口行）会删除后再导出

参数与输出：
- 标普500均线变量：`0-4000`，默认 `850`
- 纳斯达克均线变量：`0-4000`，默认 `850`
- 当均线变量为 `0` 时，按 `1` 处理（等同不平滑）
- 输出：
  - `docs/data/SP500_Average.csv`
  - `docs/data/NASDAQ_Average.csv`

## V3_Feature5：Launchd 定时后台运行

新增 Launchd 定时任务，支持每日自动执行 “下载 → 计算 → 上传 COS”，到点自动打开浏览器并触发“全部运行”。
若电脑处于睡眠状态，Launchd 不会触发；需要开启系统唤醒能力或设置定时唤醒。

关键文件：
- `scripts/run_all.py`：无界面一键执行脚本
- `scripts/install_launchd.sh`：安装 Launchd 任务
- `scripts/uninstall_launchd.sh`：卸载 Launchd 任务

使用步骤：
1. 确保 `.env` 已配置 COS 密钥
2. 确保使用 `.venv`（Python 3.9.6）并安装依赖：
   - `source .venv/bin/activate`
   - `python -m pip install -r requirements.txt`
3. 安装定时任务（默认每天 18:00 执行）：
   - `bash scripts/install_launchd.sh`
   - 自定义时间（例如 07:30）：`bash scripts/install_launchd.sh 7 30`
   - 不要安装后立刻运行（仅到点运行）：`bash scripts/install_launchd.sh 7 30 no-run`
4. 查看日志：
   - `logs/autorun.app.log`

需要调整时间时，可以重新执行安装脚本并传入新时间，或编辑 `~/Library/LaunchAgents/com.xuante.dataprocessing.plist` 内的 `StartCalendarInterval`。

睡眠状态仍需定时运行的设置建议（择一或组合）：
1. 系统设置开启唤醒能力：
   - MacBook：系统设置 → 电池 → 选项 → 打开“为网络访问唤醒”（不同系统版本文字略有差异）
   - Mac mini：系统设置 → 能源节能 → 打开“为网络访问唤醒”
2. 使用 `pmset` 设置每日定时唤醒（示例：每天 11:55 唤醒）：
   - `sudo pmset repeat wakeorpoweron MTWRFSU 11:55:00`
   - 查看当前计划：`pmset -g sched`
   - 清除计划：`sudo pmset repeat cancel`

提示：如需稳定唤醒，建议设备接电并关闭“自动进入睡眠”或开启“防止显示器关闭时自动睡眠”。
