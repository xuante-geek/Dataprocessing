const fileSelect = document.getElementById("file-select");
const convertButton = document.getElementById("convert");
const erpButton = document.getElementById("erp");
const erp10yButton = document.getElementById("erp10y");
const statusText = document.getElementById("status");
const modal = document.getElementById("modal");
const modalTitle = document.getElementById("modal-title");
const modalMessage = document.getElementById("modal-message");
const modalClose = document.getElementById("modal-close");

let isBusy = false;
let hasFiles = false;
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
  fileSelect.disabled = isBusy || !hasFiles;
  convertButton.disabled = isBusy || !hasFiles;
  erpButton.disabled = isBusy || !isServiceAvailable;
  erp10yButton.disabled = isBusy || !isServiceAvailable;
};

const setPlaceholder = (text) => {
  fileSelect.innerHTML = "";
  const option = document.createElement("option");
  option.textContent = text;
  option.value = "";
  fileSelect.appendChild(option);
};

const loadFiles = async () => {
  isBusy = true;
  updateControls();
  setStatus("正在读取 input/ 中的文件...");

  try {
    const response = await fetch("/api/files");
    if (!response.ok) {
      throw new Error("无法读取文件列表。");
    }

    const data = await response.json();
    fileSelect.innerHTML = "";
    isServiceAvailable = true;

    if (!data.files || data.files.length === 0) {
      setPlaceholder("未找到 input/ 中的 .xlsx 文件");
      hasFiles = false;
      setStatus("请将 Excel(.xlsx) 文件放入 input/，然后刷新页面。");
      return;
    }

    data.files.forEach((file) => {
      const option = document.createElement("option");
      option.value = file;
      option.textContent = file;
      fileSelect.appendChild(option);
    });

    hasFiles = true;
    setStatus(`已加载 ${data.files.length} 个文件，可以导出。`);
  } catch (error) {
    setPlaceholder("无法连接本地服务");
    hasFiles = false;
    isServiceAvailable = false;
    setStatus("本地服务未连接（请确认已运行 python src/app.py）。");
    const message =
      error.message.includes("Failed to fetch") || error.message.includes("Load failed")
        ? "请先运行本地服务：python src/app.py，然后刷新页面。"
        : error.message;
    showModal("读取失败", message);
  } finally {
    isBusy = false;
    updateControls();
  }
};

const convertFile = async () => {
  const filename = fileSelect.value;
  if (!filename) {
    showModal("未选择文件", "请选择需要转换的 Excel 文件。");
    return;
  }

  isBusy = true;
  updateControls();
  setStatus("正在导出 CSV...");

  try {
    const response = await fetch("/api/convert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filename }),
    });

    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "导出失败。");
    }

    setStatus("导出完成。");
    const csvName = data.output_csv || "";
    const xlsxName = data.output_xlsx || "";
    const parts = [];
    if (csvName) parts.push(`CSV：docs/data/${csvName}`);
    if (xlsxName) parts.push(`Excel：docs/data/${xlsxName}（已冻结首行/首列）`);
    showModal("导出完成", parts.length ? `已生成\n${parts.join("\n")}` : "已生成输出文件。");
  } catch (error) {
    setStatus("导出失败。");
    showModal("导出失败", error.message);
  } finally {
    isBusy = false;
    updateControls();
  }
};

const generateErp = async () => {
  isBusy = true;
  updateControls();
  setStatus("正在生成 ERP（Feature 2）...");

  try {
    const response = await fetch("/api/erp", { method: "POST" });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "生成失败。");
    }

    const outputs = data.outputs || {};
    const lines = [
      "已生成：",
      outputs.data_PE_clean_csv ? `- docs/data/${outputs.data_PE_clean_csv}` : null,
      outputs.data_bond_clean_csv ? `- docs/data/${outputs.data_bond_clean_csv}` : null,
      outputs.merged_csv ? `- docs/data/${outputs.merged_csv}` : null,
      outputs.erp_csv ? `- docs/data/${outputs.erp_csv}` : null,
      outputs.erp_xlsx ? `- docs/data/${outputs.erp_xlsx}` : null,
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
    const response = await fetch("/api/erp10y", { method: "POST" });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "生成失败。");
    }

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

convertButton.addEventListener("click", convertFile);
erpButton.addEventListener("click", generateErp);
erp10yButton.addEventListener("click", generateErp10y);
modalClose.addEventListener("click", hideModal);
modal.addEventListener("click", (event) => {
  if (event.target === modal) {
    hideModal();
  }
});

if (window.location.protocol === "file:") {
  setPlaceholder("请通过本地服务打开页面");
  hasFiles = false;
  isServiceAvailable = false;
  updateControls();
  setStatus("请运行：python src/app.py，然后访问 http://127.0.0.1:5000");
  showModal("需要启动本地服务", "请运行：python src/app.py，然后用浏览器打开 http://127.0.0.1:5000");
} else {
  setPlaceholder("正在加载文件列表...");
  hasFiles = false;
  isServiceAvailable = true;
  updateControls();
  loadFiles();
}
