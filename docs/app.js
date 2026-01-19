const erpButton = document.getElementById("erp");
const erp10yButton = document.getElementById("erp10y");
const rollingButton = document.getElementById("rolling");
const intervalButton = document.getElementById("interval");
const rollingNInput = document.getElementById("rolling-n");
const intervalStartInput = document.getElementById("interval-start");
const intervalEndInput = document.getElementById("interval-end");

const statusText = document.getElementById("status");
const modal = document.getElementById("modal");
const modalTitle = document.getElementById("modal-title");
const modalMessage = document.getElementById("modal-message");
const modalClose = document.getElementById("modal-close");

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
  erp10yButton.disabled = isBusy || !isServiceAvailable;
  rollingButton.disabled = isBusy || !isServiceAvailable;
  rollingNInput.disabled = isBusy;
  intervalButton.disabled = isBusy || !isServiceAvailable;
  intervalStartInput.disabled = isBusy;
  intervalEndInput.disabled = isBusy;
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
  } catch (error) {
    isServiceAvailable = false;
    setStatus("本地服务未连接（请确认已运行 python src/app.py）。");
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
  setStatus("正在生成 ERP（Feature 2）...");

  try {
    const data = await postJson("/api/erp");
    const outputs = data.outputs || {};
    const lines = [
      "已生成：",
      outputs.erp_csv ? `- docs/data/${outputs.erp_csv}` : null,
      outputs.erp_xlsx ? `- docs/data/${outputs.erp_xlsx}` : null,
      outputs.merged_csv ? `- docs/data/${outputs.merged_csv}` : null,
    ].filter(Boolean);

    setStatus("ERP 生成完成。");
    showModal("完成", lines.join("\n"));
  } catch (error) {
    setStatus("ERP 生成失败。");
    showModal("生成失败", error.message);
  } finally {
    isBusy = false;
    updateControls();
  }
};

const generateErp10y = async () => {
  isBusy = true;
  updateControls();
  setStatus("正在生成 ERP_10Year（Feature 3）...");

  try {
    const data = await postJson("/api/erp10y");
    const lines = [
      "已生成：",
      data.output_csv ? `- docs/data/${data.output_csv}` : null,
      data.output_xlsx ? `- docs/data/${data.output_xlsx}` : null,
    ].filter(Boolean);

    setStatus("ERP_10Year 生成完成。");
    showModal("完成", lines.join("\n"));
  } catch (error) {
    setStatus("ERP_10Year 生成失败。");
    showModal("生成失败", error.message);
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
  setStatus(`正在生成 ERP_Rolling Calculation（n=${n}）...`);

  try {
    const data = await postJson("/api/erprolling", { n });
    const lines = [
      `n = ${data.n}`,
      "已生成：",
      data.output_csv ? `- docs/data/${data.output_csv}` : null,
      data.output_xlsx ? `- docs/data/${data.output_xlsx}` : null,
    ].filter(Boolean);

    setStatus("滚动计算生成完成。");
    showModal("完成", lines.join("\n"));
  } catch (error) {
    setStatus("滚动计算生成失败。");
    showModal("生成失败", error.message);
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
  setStatus(`正在生成 ERP_Interval（${startDate} → ${endDate}）...`);

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
    setStatus("固定区间生成失败。");
    showModal("生成失败", error.message);
  } finally {
    isBusy = false;
    updateControls();
  }
};

erpButton.addEventListener("click", generateErp);
erp10yButton.addEventListener("click", generateErp10y);
rollingButton.addEventListener("click", generateRolling);
intervalButton.addEventListener("click", generateInterval);
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
