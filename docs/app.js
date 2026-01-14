const fileSelect = document.getElementById("file-select");
const convertButton = document.getElementById("convert");
const statusText = document.getElementById("status");
const modal = document.getElementById("modal");
const modalTitle = document.getElementById("modal-title");
const modalMessage = document.getElementById("modal-message");
const modalClose = document.getElementById("modal-close");

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

const setBusy = (isBusy) => {
  convertButton.disabled = isBusy;
  fileSelect.disabled = isBusy;
};

const loadFiles = async () => {
  setBusy(true);
  setStatus("正在读取 input 中的文件...");

  try {
    const response = await fetch("/api/files");
    if (!response.ok) {
      throw new Error("无法读取文件列表。");
    }

    const data = await response.json();
    fileSelect.innerHTML = "";

    if (!data.files || data.files.length === 0) {
      const option = document.createElement("option");
      option.textContent = "未找到 input 中的 .xlsx 文件";
      option.value = "";
      fileSelect.appendChild(option);
      fileSelect.disabled = true;
      convertButton.disabled = true;
      setStatus("等待 Excel 文件放入 input。");
      return;
    }

    fileSelect.disabled = false;
    data.files.forEach((file) => {
      const option = document.createElement("option");
      option.value = file;
      option.textContent = file;
      fileSelect.appendChild(option);
    });

    convertButton.disabled = false;
    setStatus("已就绪，可以导出。");
  } catch (error) {
    convertButton.disabled = true;
    fileSelect.innerHTML = "";
    const option = document.createElement("option");
    option.textContent = "无法连接本地服务";
    option.value = "";
    fileSelect.appendChild(option);
    fileSelect.disabled = true;
    setStatus("本地服务未连接。");
    const message =
      error.message.includes("Failed to fetch") || error.message.includes("Load failed")
        ? "请先运行本地服务：python src/app.py，然后刷新页面。"
        : error.message;
    showModal("读取失败", message);
  } finally {
    setBusy(false);
  }
};

const convertFile = async () => {
  const filename = fileSelect.value;
  if (!filename) {
    showModal("未选择文件", "请选择需要转换的 Excel 文件。");
    return;
  }

  setBusy(true);
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
    showModal("导出完成", `CSV 已保存到 docs/data/${data.output}`);
  } catch (error) {
    setStatus("导出失败。");
    showModal("导出失败", error.message);
  } finally {
    setBusy(false);
  }
};

convertButton.addEventListener("click", convertFile);
modalClose.addEventListener("click", hideModal);
modal.addEventListener("click", (event) => {
  if (event.target === modal) {
    hideModal();
  }
});

if (window.location.protocol === "file:") {
  fileSelect.innerHTML = "";
  const option = document.createElement("option");
  option.textContent = "请通过本地服务打开页面";
  option.value = "";
  fileSelect.appendChild(option);
  fileSelect.disabled = true;
  convertButton.disabled = true;
  setStatus("请运行：python src/app.py，然后访问 http://127.0.0.1:5000");
  showModal("需要启动本地服务", "请运行：python src/app.py，然后用浏览器打开 http://127.0.0.1:5000");
} else {
  loadFiles();
}
