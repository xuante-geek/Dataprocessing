const erpButton = document.getElementById("erp");
const rollingButton = document.getElementById("rolling");
const intervalButton = document.getElementById("interval");
const thermoPercentileButton = document.getElementById("thermo-percentile");
const thermoMergeButton = document.getElementById("thermo-merge");
const rollingNInput = document.getElementById("rolling-n");
const intervalStartInput = document.getElementById("interval-start");
const intervalEndInput = document.getElementById("interval-end");
const downloadRunAllButton = document.getElementById("download-run-all");
const downloadRunSelectedButton = document.getElementById("download-run-selected");
const downloadTaskTable = document.getElementById("download-task-table");
const downloadGlobalStatus = document.getElementById("download-global-status");
const downloadRunSummary = document.getElementById("download-run-summary");
const thermoStatusText = document.getElementById("thermo-status");
const maGdpInput = document.getElementById("ma-gdp");
const rpGdpInput = document.getElementById("rp-gdp");
const internalGdpInput = document.getElementById("internal-gdp");
const internalGdpMode = document.getElementById("internal-gdp-mode");
const maVolumeInput = document.getElementById("ma-volume");
const rpVolumeInput = document.getElementById("rp-volume");
const internalVolumeInput = document.getElementById("internal-volume");
const internalVolumeMode = document.getElementById("internal-volume-mode");
const maSecuritiesInput = document.getElementById("ma-securities");
const rpSecuritiesInput = document.getElementById("rp-securities");
const internalSecuritiesInput = document.getElementById("internal-securities");
const internalSecuritiesMode = document.getElementById("internal-securities-mode");
const maErpInput = document.getElementById("ma-erp");
const rpErpInput = document.getElementById("rp-erp");
const internalErpInput = document.getElementById("internal-erp");
const internalErpMode = document.getElementById("internal-erp-mode");
const wGdpInput = document.getElementById("w-gdp");
const wVolumeInput = document.getElementById("w-volume");
const wSecuritiesInput = document.getElementById("w-securities");
const wErpInput = document.getElementById("w-erp");
const colGdp = document.getElementById("col-gdp");
const colVolume = document.getElementById("col-volume");
const colSecurities = document.getElementById("col-securities");
const colErp = document.getElementById("col-erp");
const colYield = document.getElementById("col-yield");
const erpColYield = document.getElementById("erp-col-yield");
const erpColPe = document.getElementById("erp-col-pe");
const erpColErpPct = document.getElementById("erp-col-erp-pct");
const windowSizeToggle = document.getElementById("window-size-toggle");

const statusText = document.getElementById("status");
const modal = document.getElementById("modal");
const modalTitle = document.getElementById("modal-title");
const modalMessage = document.getElementById("modal-message");
const modalClose = document.getElementById("modal-close");
const pageTitle = document.getElementById("page-title");
const heroStatusText = document.getElementById("hero-status");

const tabDownload = document.getElementById("tab-download");
const tabErp = document.getElementById("tab-erp");
const tabThermo = document.getElementById("tab-thermo");
const panelDownload = document.getElementById("panel-download");
const panelErp = document.getElementById("panel-erp");
const panelThermo = document.getElementById("panel-thermo");

let isBusy = false;
let isServiceAvailable = false;
let downloadTasks = [];
let downloadRunning = false;
let downloadLastRunId = null;
let downloadLastPrompt = null;
let downloadPollTimer = null;
let activePanel = "download";

const syncInternalToggle = (modeEl, inputEl) => {
  const isAuto = modeEl.value === "auto";
  inputEl.disabled = isBusy || isAuto;
  if (isAuto) {
    inputEl.value = "";
  } else if (!inputEl.value) {
    inputEl.value = inputEl.defaultValue || "";
  }
};

const updateDownloadControls = () => {
  const disabled = downloadRunning || !isServiceAvailable;
  if (downloadRunAllButton) {
    downloadRunAllButton.disabled = disabled;
  }
  if (downloadRunSelectedButton) {
    downloadRunSelectedButton.disabled = disabled;
  }
  document.querySelectorAll(".download-row input, .download-row textarea").forEach((input) => {
    input.disabled = disabled;
  });
};

const setStatus = (message) => {
  statusText.textContent = message;
};

const setHeroStatus = (message) => {
  if (!heroStatusText) return;
  heroStatusText.textContent = message;
};

const showModal = (title, message) => {
  modalTitle.textContent = title;
  modalMessage.textContent = message;
  modal.classList.remove("hidden");
};

const hideModal = () => {
  modal.classList.add("hidden");
};

const updateControls = () => {
  erpButton.disabled = isBusy || !isServiceAvailable;
  rollingButton.disabled = isBusy || !isServiceAvailable;
  rollingNInput.disabled = isBusy;
  intervalButton.disabled = isBusy || !isServiceAvailable;
  intervalStartInput.disabled = isBusy;
  intervalEndInput.disabled = isBusy;
  thermoPercentileButton.disabled = isBusy || !isServiceAvailable;
  thermoMergeButton.disabled = isBusy || !isServiceAvailable;
  maGdpInput.disabled = isBusy;
  rpGdpInput.disabled = isBusy;
  internalGdpMode.disabled = isBusy;
  syncInternalToggle(internalGdpMode, internalGdpInput);
  maVolumeInput.disabled = isBusy;
  rpVolumeInput.disabled = isBusy;
  internalVolumeMode.disabled = isBusy;
  syncInternalToggle(internalVolumeMode, internalVolumeInput);
  maSecuritiesInput.disabled = isBusy;
  rpSecuritiesInput.disabled = isBusy;
  internalSecuritiesMode.disabled = isBusy;
  syncInternalToggle(internalSecuritiesMode, internalSecuritiesInput);
  maErpInput.disabled = isBusy;
  rpErpInput.disabled = isBusy;
  internalErpMode.disabled = isBusy;
  syncInternalToggle(internalErpMode, internalErpInput);
  wGdpInput.disabled = isBusy;
  wVolumeInput.disabled = isBusy;
  wSecuritiesInput.disabled = isBusy;
  wErpInput.disabled = isBusy;
  colGdp.disabled = isBusy;
  colVolume.disabled = isBusy;
  colSecurities.disabled = isBusy;
  colErp.disabled = isBusy;
  colYield.disabled = isBusy;
  if (windowSizeToggle) {
    windowSizeToggle.disabled = isBusy;
  }
  if (erpColYield) {
    erpColYield.disabled = isBusy;
  }
  if (erpColPe) {
    erpColPe.disabled = isBusy;
  }
  if (erpColErpPct) {
    erpColErpPct.disabled = isBusy;
  }
  updateDownloadControls();
};

const checkService = async () => {
  if (window.location.protocol === "file:") {
    isServiceAvailable = false;
    updateControls();
    updateHeroStatusForPanel();
    setStatus("请运行：python src/app.py，然后访问 http://127.0.0.1:5000");
    showModal("需要启动本地服务", "请运行：python src/app.py，然后用浏览器打开 http://127.0.0.1:5000");
    return;
  }

  isBusy = true;
  updateControls();
  setHeroStatus("正在连接本地服务...");
  setStatus("");

  try {
    const response = await fetch("/api/files");
    if (!response.ok) {
      throw new Error("本地服务不可用。");
    }
    isServiceAvailable = true;
    updateHeroStatusForPanel();
    thermoStatusText.textContent = "";
    if (!downloadPollTimer) {
      initDownloadPanel();
    }
  } catch (error) {
    isServiceAvailable = false;
    updateHeroStatusForPanel();
    setStatus("本地服务未连接（请确认已运行 python src/app.py）。");
    thermoStatusText.textContent = "";
    showModal("连接失败", "无法连接本地服务，请先运行：python src/app.py");
  } finally {
    isBusy = false;
    updateControls();
  }
};

const postJson = async (url, payload) => {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload || {}),
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || "请求失败。");
  }
  return data;
};

const updateHeroStatusForPanel = () => {
  if (!isServiceAvailable) {
    setHeroStatus("本地服务未连接（请确认已运行 python src/app.py）。");
    return;
  }
  if (activePanel === "download") {
    setHeroStatus("本地服务已连接，可以开始下载。");
  } else {
    setHeroStatus("本地服务已连接，可以开始生成。");
  }
};

const renderDownloadTasks = () => {
  if (!downloadTaskTable) return;
  const rows = downloadTasks.map((task) => {
    const url = task.url ? String(task.url) : "";
    return `
      <div class="download-row" data-id="${task.id}">
        <div class="cell checkbox"><input type="checkbox" class="download-task-check" /></div>
        <div class="cell name">${task.name}</div>
        <div class="cell url"><input type="text" value="${url}" /></div>
        <div class="cell status"><span class="status-pill">待机</span></div>
      </div>
    `;
  });
  downloadTaskTable.innerHTML = `
    <div class="download-row header">
      <div class="cell checkbox"></div>
      <div class="cell name">名称</div>
      <div class="cell url">URL</div>
      <div class="cell status">状态</div>
    </div>
    ${rows.join("")}
  `;
  updateDownloadControls();
};

const loadDownloadConfig = async () => {
  try {
    const res = await fetch("/api/download/config");
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.error || "读取下载配置失败。");
    }
    downloadTasks = data.tasks || [];
    renderDownloadTasks();
  } catch (error) {
    showModal("下载配置失败", error.message);
  }
};

const collectDownloadUrls = () => {
  const urls = {};
  document.querySelectorAll(".download-row[data-id]").forEach((row) => {
    const id = row.dataset.id;
    const input = row.querySelector(".cell.url input");
    if (input && input.value.trim()) {
      urls[id] = input.value.trim();
    }
  });
  return urls;
};

const selectedDownloadTaskIds = () => {
  const ids = [];
  document.querySelectorAll(".download-row[data-id]").forEach((row) => {
    const checked = row.querySelector(".download-task-check")?.checked;
    if (checked) {
      ids.push(row.dataset.id);
    }
  });
  return ids;
};

const runDownloadTasks = async (ids) => {
  if (downloadRunning) {
    showModal("提示", "已有下载任务在运行，请稍后再试。");
    return;
  }
  if (!ids.length) {
    showModal("提示", "请先选择要下载的任务。");
    return;
  }
  try {
    const payload = {
      task_ids: ids,
      urls: collectDownloadUrls(),
      trigger: "manual",
    };
    const res = await fetch("/api/download/run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.message || "启动失败");
    }
    downloadRunning = true;
    updateDownloadControls();
  } catch (error) {
    showModal("启动失败", error.message);
  }
};

const updateDownloadGlobalStatus = (status) => {
  if (!downloadGlobalStatus) return;
  let label = "待机";
  let klass = "";
  if (status.running) {
    label = "运行中";
    klass = "running";
  } else if (status.success === true) {
    label = "完成";
    klass = "success";
  } else if (status.success === false) {
    label = "失败";
    klass = "error";
  }
  downloadGlobalStatus.textContent = label;
  downloadGlobalStatus.className = `status-chip ${klass}`.trim();
};

const updateDownloadTaskRow = (taskId, info) => {
  const row = document.querySelector(`.download-row[data-id='${taskId}']`);
  if (!row) return;
  const pill = row.querySelector(".status-pill");
  if (!pill) return;

  let label = "待机";
  let klass = "";
  if (info.status === "running") {
    label = "进行中";
    klass = "running";
  } else if (info.status === "waiting_login") {
    label = "待登录";
    klass = "warning";
  } else if (info.status === "success") {
    label = "成功";
    klass = "success";
    if (info.validation && info.validation.warnings && info.validation.warnings.length) {
      label = "成功(有警告)";
      klass = "warning";
    }
  } else if (info.status === "error") {
    label = "失败";
    klass = "error";
  }
  pill.textContent = label;
  pill.className = `status-pill ${klass}`.trim();
};

const buildDownloadSummary = (status) => {
  if (!status.started_at) {
    return "暂无运行记录。";
  }
  const lines = [`运行时间: ${status.started_at} → ${status.ended_at || "进行中"}`];
  lines.push(`触发方式: ${status.trigger || "manual"}`);
  lines.push(`结果: ${status.success ? "成功" : "失败/部分失败"}`);
  if (status.message) {
    lines.push(`提示: ${status.message}`);
  }
  return lines.join("<br />");
};

const buildDownloadModalBody = (status) => {
  const rows = [];
  if (status.message) {
    rows.push(`<div>${status.message}</div>`);
  }
  for (const [id, info] of Object.entries(status.tasks || {})) {
    const task = downloadTasks.find((t) => t.id === id);
    const name = task ? task.name : id;
    let line = `${name}: ${info.status || "idle"}`;
    if (info.status === "error" && info.message) {
      line += ` (${info.message})`;
    }
    if (info.validation && info.validation.errors && info.validation.errors.length) {
      line += `，校验错误: ${info.validation.errors.join("；")}`;
    }
    if (info.validation && info.validation.warnings && info.validation.warnings.length) {
      line += `，校验警告: ${info.validation.warnings.join("；")}`;
    }
    rows.push(`<div>${line}</div>`);
  }
  return rows.join("");
};

const pollDownloadStatus = async () => {
  if (!isServiceAvailable) {
    return;
  }
  const res = await fetch("/api/download/status");
  const status = await res.json();

  downloadRunning = !!status.running;
  updateDownloadControls();
  updateDownloadGlobalStatus(status);
  if (downloadRunSummary) {
    downloadRunSummary.innerHTML = buildDownloadSummary(status);
  }

  for (const [id, info] of Object.entries(status.tasks || {})) {
    updateDownloadTaskRow(id, info);
  }

  if (status.running && status.message && status.message.includes("扫码") && status.message !== downloadLastPrompt) {
    downloadLastPrompt = status.message;
    showModal("需要登录", status.message);
  }

  if (status.run_id && status.run_id !== downloadLastRunId && status.ended_at) {
    downloadLastRunId = status.run_id;
    const title = status.success ? "下载完成" : "下载出现失败";
    showModal(title, buildDownloadModalBody(status));
  }
};

const initDownloadPanel = () => {
  loadDownloadConfig().then(() => {
    pollDownloadStatus();
    if (!downloadPollTimer) {
      downloadPollTimer = setInterval(pollDownloadStatus, 5000);
    }
  });
};

const generateErp = async () => {
  isBusy = true;
  updateControls();
  setStatus("正在导出完整周期 ERP...");

  try {
    const payload = {
      include_yield: Boolean(erpColYield && erpColYield.checked),
      include_pe: Boolean(erpColPe && erpColPe.checked),
      include_erp_percentile: Boolean(erpColErpPct && erpColErpPct.checked),
    };
    const data = await postJson("/api/erp", payload);
    const outputs = data.outputs || {};
    const lines = [
      "已生成：",
      outputs.data_PE_clean_csv ? `- docs/data/${outputs.data_PE_clean_csv}` : null,
      outputs.data_bond_clean_csv ? `- docs/data/${outputs.data_bond_clean_csv}` : null,
      outputs.merged_csv ? `- docs/data/${outputs.merged_csv}` : null,
      outputs.erp_csv ? `- docs/data/${outputs.erp_csv}` : null,
    ].filter(Boolean);

    setStatus("导出完成。");
    showModal("完成", lines.join("\n"));
  } catch (error) {
    setStatus("导出失败。");
    showModal("导出失败", error.message);
  } finally {
    isBusy = false;
    updateControls();
  }
};

const parseRollingN = () => {
  const raw = String(rollingNInput.value || "").trim();
  const n = Number(raw);
  if (!Number.isFinite(n) || !Number.isInteger(n)) {
    throw new Error("n 必须为整数（1-4000）。");
  }
  if (n < 1 || n > 4000) {
    throw new Error("n 超出范围（1-4000）。");
  }
  return n;
};

const generateRolling = async () => {
  let n;
  try {
    n = parseRollingN();
  } catch (error) {
    showModal("参数错误", error.message);
    return;
  }

  isBusy = true;
  updateControls();
  setStatus(`正在导出滚动周期 ERP（n=${n}）...`);

  try {
    const payload = {
      n,
      include_yield: Boolean(erpColYield && erpColYield.checked),
      include_pe: Boolean(erpColPe && erpColPe.checked),
      include_erp_percentile: Boolean(erpColErpPct && erpColErpPct.checked),
    };
    const data = await postJson("/api/erprolling", payload);
    const lines = [
      `n = ${data.n}`,
      "已生成：",
      data.output_csv ? `- docs/data/${data.output_csv}` : null,
    ].filter(Boolean);

    setStatus("导出完成。");
    showModal("完成", lines.join("\n"));
  } catch (error) {
    setStatus("导出失败。");
    showModal("导出失败", error.message);
  } finally {
    isBusy = false;
    updateControls();
  }
};

const parseIntervalStart = () => {
  const raw = String(intervalStartInput.value || "").trim();
  if (!raw) {
    throw new Error("请填写起始日期（YYYY-MM-DD）。");
  }
  if (!/^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(raw)) {
    throw new Error("起始日期格式必须为 YYYY-MM-DD。");
  }
  return raw;
};

const parseIntervalEnd = () => {
  const raw = String(intervalEndInput.value || "").trim();
  if (!raw) {
    throw new Error("请填写终止日期（YYYY-MM-DD）。");
  }
  if (!/^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(raw)) {
    throw new Error("终止日期格式必须为 YYYY-MM-DD。");
  }
  return raw;
};

const generateInterval = async () => {
  let startDate;
  let endDate;
  try {
    startDate = parseIntervalStart();
    endDate = parseIntervalEnd();
  } catch (error) {
    showModal("参数错误", error.message);
    return;
  }

  isBusy = true;
  updateControls();
  setStatus(`正在导出指定周期 ERP（${startDate} → ${endDate}）...`);

  try {
    const payload = {
      start_date: startDate,
      end_date: endDate,
      include_yield: Boolean(erpColYield && erpColYield.checked),
      include_pe: Boolean(erpColPe && erpColPe.checked),
      include_erp_percentile: Boolean(erpColErpPct && erpColErpPct.checked),
    };
    const data = await postJson("/api/erpinterval", payload);
    if (data.used_end_date && intervalEndInput.value !== data.used_end_date) {
      intervalEndInput.value = data.used_end_date;
    }
    const adjustedNote = data.adjusted_to_trading_day
      ? `（非交易日已自动调整为 ${data.used_start_date}）`
      : "";
    const adjustedEndNote = data.adjusted_end_to_trading_day
      ? `（非交易日已自动回退为 ${data.used_end_date}）`
      : "";
    const lines = [
      `起始日期：${data.input_start_date} ${adjustedNote}`.trim(),
      `终止日期：${data.input_end_date} ${adjustedEndNote}`.trim(),
      `有效区间：${data.used_start_date} → ${data.used_end_date}`,
      "已生成：",
      data.output_csv ? `- docs/data/${data.output_csv}` : null,
    ].filter(Boolean);

    setStatus("固定区间生成完成。");
    showModal("完成", lines.join("\n"));
  } catch (error) {
    setStatus("导出失败。");
    showModal("导出失败", error.message);
  } finally {
    isBusy = false;
    updateControls();
  }
};

const setActivePanel = (name) => {
  activePanel = name;
  const isDownload = name === "download";
  const isErp = name === "erp";
  const isThermo = name === "thermo";
  tabDownload.classList.toggle("active", isDownload);
  tabErp.classList.toggle("active", isErp);
  tabThermo.classList.toggle("active", isThermo);
  tabDownload.setAttribute("aria-selected", String(isDownload));
  tabErp.setAttribute("aria-selected", String(isErp));
  tabThermo.setAttribute("aria-selected", String(isThermo));
  panelDownload.classList.toggle("hidden", !isDownload);
  panelErp.classList.toggle("hidden", !isErp);
  panelThermo.classList.toggle("hidden", !isThermo);
  pageTitle.textContent = isDownload
    ? "DataDownload 控制台"
    : isThermo
      ? "市场温度计"
      : "股权风险溢价（ERP）处理器";
  updateHeroStatusForPanel();
};

const parseIntInRange = (value, min, max, label) => {
  const raw = String(value || "").trim();
  const n = Number(raw);
  if (!Number.isFinite(n) || !Number.isInteger(n)) {
    throw new Error(`${label} 必须为整数（${min}-${max}）。`);
  }
  if (n < min || n > max) {
    throw new Error(`${label} 超出范围（${min}-${max}）。`);
  }
  return n;
};

const parseInternalValue = (modeEl, inputEl, min, max, label) => {
  if (modeEl.value === "auto") {
    return null;
  }
  return parseIntInRange(inputEl.value, min, max, label);
};

const generateThermoPercentiles = async () => {
  let payload;
  try {
    payload = {
      moving_average_gdp: parseIntInRange(maGdpInput.value, 1, 1000, "总市值/GDP平均移动（周频）"),
      rolling_period_gdp: parseIntInRange(rpGdpInput.value, 1, 1000, "总市值/GDP分位滚动周期（周频）"),
      internal_gdp_mode: internalGdpMode.value,
      internal_gdp: parseInternalValue(
        internalGdpMode,
        internalGdpInput,
        1,
        1000,
        "总市值/GDP最小递减滚动周期（周频）",
      ),
      moving_average_volume: parseIntInRange(maVolumeInput.value, 1, 4000, "成交量平均移动"),
      rolling_period_volume: parseIntInRange(rpVolumeInput.value, 1, 4000, "成交量/总市值分位滚动周期"),
      internal_volume_mode: internalVolumeMode.value,
      internal_volume: parseInternalValue(
        internalVolumeMode,
        internalVolumeInput,
        1,
        4000,
        "成交量/总市值最小递减滚动周期",
      ),
      moving_average_securities: parseIntInRange(maSecuritiesInput.value, 1, 4000, "融资融券平均移动"),
      rolling_period_securities: parseIntInRange(rpSecuritiesInput.value, 1, 4000, "融资融券/总市值分位滚动周期"),
      internal_securities_mode: internalSecuritiesMode.value,
      internal_securities: parseInternalValue(
        internalSecuritiesMode,
        internalSecuritiesInput,
        1,
        4000,
        "融资融券/总市值最小递减滚动周期",
      ),
      moving_erp: parseIntInRange(maErpInput.value, 1, 4000, "股权风险溢价平均移动"),
      rolling_period_erp: parseIntInRange(rpErpInput.value, 1, 4000, "股权风险溢价分位滚动周期"),
      internal_erp_mode: internalErpMode.value,
      internal_erp: parseInternalValue(
        internalErpMode,
        internalErpInput,
        1,
        4000,
        "股权风险溢价最小递减滚动周期",
      ),
      include_window_size: Boolean(windowSizeToggle && windowSizeToggle.checked),
    };
  } catch (error) {
    showModal("参数错误", error.message);
    return;
  }

  isBusy = true;
  updateControls();
  thermoStatusText.textContent = "正在导出市场温度计分位数据（包含清洗）...";

  try {
    const data = await postJson("/api/thermometer/percentiles", payload);
    const outputs = data.outputs || {};
    const lines = [
      "已生成：",
      outputs.ratio_gdp_csv ? `- docs/data/${outputs.ratio_gdp_csv}` : null,
      outputs.ratio_volume_csv ? `- docs/data/${outputs.ratio_volume_csv}` : null,
      outputs.ratio_securities_lend_csv ? `- docs/data/${outputs.ratio_securities_lend_csv}` : null,
      outputs.erp_csv ? `- docs/data/${outputs.erp_csv}` : null,
    ].filter(Boolean);
    thermoStatusText.textContent = "导出完成。";
    showModal("完成", lines.join("\n"));
  } catch (error) {
    thermoStatusText.textContent = "导出失败。";
    showModal("导出失败", error.message);
  } finally {
    isBusy = false;
    updateControls();
  }
};

const parseFloatInRange = (value, min, max, label) => {
  const raw = String(value ?? "").trim();
  const n = Number(raw);
  if (!Number.isFinite(n)) {
    throw new Error(`${label} 必须为数值（${min}-${max}）。`);
  }
  if (n < min || n > max) {
    throw new Error(`${label} 超出范围（${min}-${max}）。`);
  }
  return n;
};

const generateThermoMerge = async () => {
  let payload;
  try {
    payload = {
      moving_average_gdp: parseIntInRange(maGdpInput.value, 1, 1000, "总市值/GDP平均移动（周频）"),
      rolling_period_gdp: parseIntInRange(rpGdpInput.value, 1, 1000, "总市值/GDP分位滚动周期（周频）"),
      internal_gdp_mode: internalGdpMode.value,
      internal_gdp: parseInternalValue(
        internalGdpMode,
        internalGdpInput,
        1,
        1000,
        "总市值/GDP最小递减滚动周期（周频）",
      ),
      moving_average_volume: parseIntInRange(maVolumeInput.value, 1, 4000, "成交量平均移动"),
      rolling_period_volume: parseIntInRange(rpVolumeInput.value, 1, 4000, "成交量/总市值分位滚动周期"),
      internal_volume_mode: internalVolumeMode.value,
      internal_volume: parseInternalValue(
        internalVolumeMode,
        internalVolumeInput,
        1,
        4000,
        "成交量/总市值最小递减滚动周期",
      ),
      moving_average_securities: parseIntInRange(maSecuritiesInput.value, 1, 4000, "融资融券平均移动"),
      rolling_period_securities: parseIntInRange(rpSecuritiesInput.value, 1, 4000, "融资融券/总市值分位滚动周期"),
      internal_securities_mode: internalSecuritiesMode.value,
      internal_securities: parseInternalValue(
        internalSecuritiesMode,
        internalSecuritiesInput,
        1,
        4000,
        "融资融券/总市值最小递减滚动周期",
      ),
      moving_erp: parseIntInRange(maErpInput.value, 1, 4000, "股权风险溢价平均移动"),
      rolling_period_erp: parseIntInRange(rpErpInput.value, 1, 4000, "股权风险溢价分位滚动周期"),
      internal_erp_mode: internalErpMode.value,
      internal_erp: parseInternalValue(
        internalErpMode,
        internalErpInput,
        1,
        4000,
        "股权风险溢价最小递减滚动周期",
      ),

      weight_gdp: parseFloatInRange(wGdpInput.value, 0, 100, "权重：市值/GDP（%）"),
      weight_volume: parseFloatInRange(wVolumeInput.value, 0, 100, "权重：成交量/市值（%）"),
      weight_securities_lend: parseFloatInRange(wSecuritiesInput.value, 0, 100, "权重：融资融券/市值（%）"),
      weight_erp: parseFloatInRange(wErpInput.value, 0, 100, "权重：股权风险溢价分位（%）"),

      include_gdp_percentile: Boolean(colGdp.checked),
      include_volume_percentile: Boolean(colVolume.checked),
      include_securities_percentile: Boolean(colSecurities.checked),
      include_erp: Boolean(colErp.checked),
      include_bond_yield: Boolean(colYield.checked),
    };
    const sum =
      payload.weight_gdp + payload.weight_volume + payload.weight_securities_lend + payload.weight_erp;
    if (sum > 100.000001) {
      throw new Error("权重之和不能超过 100%。");
    }
  } catch (error) {
    showModal("参数错误", error.message);
    return;
  }

  isBusy = true;
  updateControls();
  thermoStatusText.textContent = "正在导出市场温度计总表...";

  try {
    const data = await postJson("/api/thermometer/merge", payload);
    thermoStatusText.textContent = "导出完成。";
    const lines = [
      "已生成：",
      data.output_csv ? `- docs/data/${data.output_csv}` : null,
      data.date_begin_used ? `起始日期：${data.date_begin_used}` : null,
      data.date_end ? `结束日期：${data.date_end}` : null,
    ].filter(Boolean);
    showModal("完成", lines.join("\n"));
  } catch (error) {
    thermoStatusText.textContent = "导出失败。";
    showModal("导出失败", error.message);
  } finally {
    isBusy = false;
    updateControls();
  }
};

erpButton.addEventListener("click", generateErp);
rollingButton.addEventListener("click", generateRolling);
intervalButton.addEventListener("click", generateInterval);
thermoPercentileButton.addEventListener("click", generateThermoPercentiles);
thermoMergeButton.addEventListener("click", generateThermoMerge);
if (downloadRunAllButton) {
  downloadRunAllButton.addEventListener("click", () => {
    const ids = downloadTasks.map((task) => task.id);
    runDownloadTasks(ids);
  });
}
if (downloadRunSelectedButton) {
  downloadRunSelectedButton.addEventListener("click", () => {
    runDownloadTasks(selectedDownloadTaskIds());
  });
}
modalClose.addEventListener("click", hideModal);
modal.addEventListener("click", (event) => {
  if (event.target === modal) {
    hideModal();
  }
});

isServiceAvailable = false;
updateControls();
const bindInternalMode = (modeEl, inputEl) => {
  modeEl.addEventListener("change", () => {
    syncInternalToggle(modeEl, inputEl);
  });
};
bindInternalMode(internalGdpMode, internalGdpInput);
bindInternalMode(internalVolumeMode, internalVolumeInput);
bindInternalMode(internalSecuritiesMode, internalSecuritiesInput);
bindInternalMode(internalErpMode, internalErpInput);
syncInternalToggle(internalGdpMode, internalGdpInput);
syncInternalToggle(internalVolumeMode, internalVolumeInput);
syncInternalToggle(internalSecuritiesMode, internalSecuritiesInput);
syncInternalToggle(internalErpMode, internalErpInput);
const today = new Date();
const localDate = new Date(today.getTime() - today.getTimezoneOffset() * 60000)
  .toISOString()
  .slice(0, 10);
intervalEndInput.value = localDate;
checkService();

if (tabDownload) {
  tabDownload.addEventListener("click", () => setActivePanel("download"));
}
tabErp.addEventListener("click", () => setActivePanel("erp"));
tabThermo.addEventListener("click", () => setActivePanel("thermo"));
setActivePanel("download");
