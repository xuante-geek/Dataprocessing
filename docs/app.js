const erpButton = document.getElementById("erp");
const rollingButton = document.getElementById("rolling");
const intervalButton = document.getElementById("interval");
const thermoPercentileButton = document.getElementById("thermo-percentile");
const thermoMergeButton = document.getElementById("thermo-merge");
const rollingNInput = document.getElementById("rolling-n");
const intervalStartInput = document.getElementById("interval-start");
const intervalEndInput = document.getElementById("interval-end");
const thermoStatusText = document.getElementById("thermo-status");
const thermoMergeStatusText = document.getElementById("thermo-merge-status");
const maGdpInput = document.getElementById("ma-gdp");
const rpGdpInput = document.getElementById("rp-gdp");
const maVolumeInput = document.getElementById("ma-volume");
const rpVolumeInput = document.getElementById("rp-volume");
const maSecuritiesInput = document.getElementById("ma-securities");
const rpSecuritiesInput = document.getElementById("rp-securities");
const maErpInput = document.getElementById("ma-erp");
const rpErpInput = document.getElementById("rp-erp");
const wGdpInput = document.getElementById("w-gdp");
const wVolumeInput = document.getElementById("w-volume");
const wSecuritiesInput = document.getElementById("w-securities");
const wErpInput = document.getElementById("w-erp");
const colGdp = document.getElementById("col-gdp");
const colVolume = document.getElementById("col-volume");
const colSecurities = document.getElementById("col-securities");
const colErp = document.getElementById("col-erp");
const colYield = document.getElementById("col-yield");

const statusText = document.getElementById("status");
const modal = document.getElementById("modal");
const modalTitle = document.getElementById("modal-title");
const modalMessage = document.getElementById("modal-message");
const modalClose = document.getElementById("modal-close");
const pageTitle = document.getElementById("page-title");

const tabErp = document.getElementById("tab-erp");
const tabThermo = document.getElementById("tab-thermo");
const panelErp = document.getElementById("panel-erp");
const panelThermo = document.getElementById("panel-thermo");

let isBusy = false;
let isServiceAvailable = false;

const setStatus = (message) => {
  statusText.textContent = message;
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
  maVolumeInput.disabled = isBusy;
  rpVolumeInput.disabled = isBusy;
  maSecuritiesInput.disabled = isBusy;
  rpSecuritiesInput.disabled = isBusy;
  maErpInput.disabled = isBusy;
  rpErpInput.disabled = isBusy;
  wGdpInput.disabled = isBusy;
  wVolumeInput.disabled = isBusy;
  wSecuritiesInput.disabled = isBusy;
  wErpInput.disabled = isBusy;
  colGdp.disabled = isBusy;
  colVolume.disabled = isBusy;
  colSecurities.disabled = isBusy;
  colErp.disabled = isBusy;
  colYield.disabled = isBusy;
};

const checkService = async () => {
  if (window.location.protocol === "file:") {
    isServiceAvailable = false;
    updateControls();
    setStatus("请运行：python src/app.py，然后访问 http://127.0.0.1:5000");
    showModal("需要启动本地服务", "请运行：python src/app.py，然后用浏览器打开 http://127.0.0.1:5000");
    return;
  }

  isBusy = true;
  updateControls();
  setStatus("正在检查本地服务...");

  try {
    const response = await fetch("/api/files");
    if (!response.ok) {
      throw new Error("本地服务不可用。");
    }
    isServiceAvailable = true;
    setStatus("本地服务已连接，可以开始生成。");
    thermoStatusText.textContent = "本地服务已连接，可以开始导出。";
  } catch (error) {
    isServiceAvailable = false;
    setStatus("本地服务未连接（请确认已运行 python src/app.py）。");
    thermoStatusText.textContent = "本地服务未连接。";
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

const generateErp = async () => {
  isBusy = true;
  updateControls();
  setStatus("正在导出完整周期 ERP...");

  try {
    const data = await postJson("/api/erp");
    const outputs = data.outputs || {};
    const lines = [
      "已生成：",
      outputs.erp_csv ? `- docs/data/${outputs.erp_csv}` : null,
      outputs.erp_xlsx ? `- docs/data/${outputs.erp_xlsx}` : null,
      outputs.merged_csv ? `- docs/data/${outputs.merged_csv}` : null,
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
    const data = await postJson("/api/erprolling", { n });
    const lines = [
      `n = ${data.n}`,
      "已生成：",
      data.output_csv ? `- docs/data/${data.output_csv}` : null,
      data.output_xlsx ? `- docs/data/${data.output_xlsx}` : null,
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
    const data = await postJson("/api/erpinterval", { start_date: startDate, end_date: endDate });
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
      data.output_xlsx ? `- docs/data/${data.output_xlsx}` : null,
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
  const isThermo = name === "thermo";
  tabErp.classList.toggle("active", !isThermo);
  tabThermo.classList.toggle("active", isThermo);
  tabErp.setAttribute("aria-selected", String(!isThermo));
  tabThermo.setAttribute("aria-selected", String(isThermo));
  panelErp.classList.toggle("hidden", isThermo);
  panelThermo.classList.toggle("hidden", !isThermo);
  pageTitle.textContent = isThermo ? "市场温度计" : "股权风险溢价（ERP）处理器";
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

const generateThermoPercentiles = async () => {
  let payload;
  try {
    payload = {
      moving_average_gdp: parseIntInRange(maGdpInput.value, 1, 1000, "总市值/GDP平均移动（周频）"),
      rolling_period_gdp: parseIntInRange(rpGdpInput.value, 1, 1000, "总市值/GDP分位滚动周期（周频）"),
      moving_average_volume: parseIntInRange(maVolumeInput.value, 1, 4000, "成交量平均移动"),
      rolling_period_volume: parseIntInRange(rpVolumeInput.value, 1, 4000, "成交量/总市值分位滚动周期"),
      moving_average_securities: parseIntInRange(maSecuritiesInput.value, 1, 4000, "融资融券平均移动"),
      rolling_period_securities: parseIntInRange(rpSecuritiesInput.value, 1, 4000, "融资融券/总市值分位滚动周期"),
      moving_erp: parseIntInRange(maErpInput.value, 1, 4000, "股权风险溢价平均移动"),
      rolling_period_erp: parseIntInRange(rpErpInput.value, 1, 4000, "股权风险溢价分位滚动周期"),
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
      moving_average_volume: parseIntInRange(maVolumeInput.value, 1, 4000, "成交量平均移动"),
      rolling_period_volume: parseIntInRange(rpVolumeInput.value, 1, 4000, "成交量/总市值分位滚动周期"),
      moving_average_securities: parseIntInRange(maSecuritiesInput.value, 1, 4000, "融资融券平均移动"),
      rolling_period_securities: parseIntInRange(rpSecuritiesInput.value, 1, 4000, "融资融券/总市值分位滚动周期"),
      moving_erp: parseIntInRange(maErpInput.value, 1, 4000, "股权风险溢价平均移动"),
      rolling_period_erp: parseIntInRange(rpErpInput.value, 1, 4000, "股权风险溢价分位滚动周期"),

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
  thermoMergeStatusText.textContent = "正在导出市场温度计总表...";

  try {
    const data = await postJson("/api/thermometer/merge", payload);
    thermoMergeStatusText.textContent = "导出完成。";
    const lines = [
      "已生成：",
      data.output_csv ? `- docs/data/${data.output_csv}` : null,
      data.date_begin_used ? `起始日期：${data.date_begin_used}` : null,
      data.date_end ? `结束日期：${data.date_end}` : null,
    ].filter(Boolean);
    showModal("完成", lines.join("\n"));
  } catch (error) {
    thermoMergeStatusText.textContent = "导出失败。";
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
modalClose.addEventListener("click", hideModal);
modal.addEventListener("click", (event) => {
  if (event.target === modal) {
    hideModal();
  }
});

isServiceAvailable = false;
updateControls();
const today = new Date();
const localDate = new Date(today.getTime() - today.getTimezoneOffset() * 60000)
  .toISOString()
  .slice(0, 10);
intervalEndInput.value = localDate;
checkService();

tabErp.addEventListener("click", () => setActivePanel("erp"));
tabThermo.addEventListener("click", () => setActivePanel("thermo"));
setActivePanel("erp");
