const fileInput = document.querySelector("#file-input");
const dropZone = document.querySelector("#drop-zone");
const selectedFile = document.querySelector("#selected-file");
const fileName = document.querySelector("#file-name");
const fileMeta = document.querySelector("#file-meta");
const changeFile = document.querySelector("#change-file");
const calculateButton = document.querySelector("#calculate-button");
const statusPanel = document.querySelector("#status-panel");
const statusTitle = document.querySelector("#status-title");
const statusDetail = document.querySelector("#status-detail");
const results = document.querySelector("#results");
const resultTime = document.querySelector("#result-time");
const downloadList = document.querySelector("#download-list");
const resultNote = document.querySelector("#result-note");
const errorPanel = document.querySelector("#error-panel");
const errorSummary = document.querySelector("#error-summary");
const errorList = document.querySelector("#error-list");
const errorDownloads = document.querySelector("#error-downloads");

let currentFile = null;
let engineReady = false;
let isWorking = false;
let downloadUrls = [];
let pendingCompletion = null;

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} bayt`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(2)} MB`;
}

function setStatus(state, title, detail) {
  statusPanel.classList.remove("is-ready", "is-success", "is-error", "is-working");
  statusPanel.classList.add(`is-${state}`);
  statusTitle.textContent = title;
  statusDetail.textContent = detail;
}

function clearDownloads() {
  for (const url of downloadUrls) URL.revokeObjectURL(url);
  downloadUrls = [];
  downloadList.replaceChildren();
  errorDownloads.replaceChildren();
}

function clearResults() {
  clearDownloads();
  results.hidden = true;
  errorPanel.hidden = true;
  errorList.replaceChildren();
}

function isExcelFile(file) {
  return file && /\.xlsx$/i.test(file.name) && file.size > 0;
}

function selectFile(file) {
  if (!file || isWorking) return;
  if (!isExcelFile(file)) {
    setStatus("error", "Excel .xlsx faylı tələb olunur", "XLSX formatında boş olmayan kitab seçin.");
    return;
  }
  clearResults();
  currentFile = file;
  fileName.textContent = file.name;
  fileMeta.textContent = `${formatBytes(file.size)} · ilkin fayl dəyişdirilməyəcək`;
  dropZone.hidden = true;
  selectedFile.hidden = false;
  calculateButton.textContent = "Hesabla və faylları hazırla";
  calculateButton.disabled = !engineReady;
  if (engineReady) {
    setStatus("ready", "Hesablama modulu hazırdır", "Fayl seçilib. Hesablamaya başlaya bilərsiniz.");
  }
}

fileInput.addEventListener("change", () => selectFile(fileInput.files[0]));
changeFile.addEventListener("click", () => {
  fileInput.value = "";
  fileInput.click();
});

for (const eventName of ["dragenter", "dragover"]) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("is-dragging");
  });
}

for (const eventName of ["dragleave", "drop"]) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("is-dragging");
  });
}

dropZone.addEventListener("drop", (event) => selectFile(event.dataTransfer.files[0]));

function makeDownload(container, title, description, filename, buffer) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  downloadUrls.push(url);

  const card = document.createElement("div");
  card.className = "download-card";
  const heading = document.createElement("strong");
  heading.textContent = title;
  const copy = document.createElement("span");
  copy.textContent = description;
  const link = document.createElement("a");
  link.className = "download-button";
  link.href = url;
  link.download = filename;
  link.textContent = "Endir";
  card.append(heading, copy, link);
  container.append(card);
}

function updatedFilename(name) {
  const stem = name.replace(/\.xlsx$/i, "").replace(/_AZ$/i, "") || "Inventory_Input";
  return `${stem}_AZ.xlsx`;
}

function finishRequest() {
  isWorking = false;
  changeFile.disabled = false;
  fileInput.disabled = false;
  calculateButton.disabled = !engineReady || !currentFile;
  calculateButton.textContent = "Yenidən hesabla";
}

function settleRequest(method, value) {
  if (!pendingCompletion) return;
  pendingCompletion[method](value);
  pendingCompletion = null;
}

function renderSuccess(meta, files) {
  clearResults();
  makeDownload(
    downloadList,
    "Hesabat",
    "Formulasız, statik hesablama nəticələri.",
    "Inventory_Report.xlsx",
    files.report,
  );
  makeDownload(
    downloadList,
    "Yenilənmiş kitab",
    "Azərbaycan dilində, daimi kodlarla və adi diapazonlarla giriş kitabı.",
    updatedFilename(currentFile.name),
    files.updated,
  );
  const codeText = meta.assigned_codes
    ? `Yeni kodların sayı: ${meta.assigned_codes}.`
    : "Yeni kod tələb olunmadı.";
  const warningText = meta.warning_count
    ? ` Mümkün təkrarlar: ${meta.warning_count}; Yoxlamalar vərəqinə baxın.`
    : "";
  resultNote.textContent = `${codeText}${warningText} İlkin fayl dəyişməyib. İşi endirdiyiniz yenilənmiş kitabda davam etdirin.`;
  resultTime.textContent = `${meta.elapsed.toFixed(1)} san.`;
  results.hidden = false;
  setStatus("success", "Hesablama tamamlandı", "Hər iki fayl endirmək üçün hazırdır.");
  finishRequest();
}

function issueText(issue) {
  const location = issue.Row ? `${issue.Sheet}, sətir ${issue.Row}` : issue.Sheet;
  return `${location}: ${issue.Message}`;
}

function renderFailure(meta, files) {
  clearResults();
  const issues = meta.issues || [];
  errorSummary.textContent = `Xətaların sayı: ${issues.length}. Hesabat yaradılmadı.`;
  for (const issue of issues.slice(0, 12)) {
    const item = document.createElement("li");
    item.textContent = issueText(issue);
    errorList.append(item);
  }
  if (issues.length > 12) {
    const item = document.createElement("li");
    item.textContent = `Digər xətalar: ${issues.length - 12}. Tam siyahı Yoxlamalar vərəqindədir.`;
    errorList.append(item);
  }
  makeDownload(
    errorDownloads,
    "Yoxlamalar",
    "Hər xəta üçün vərəq, sətir və səbəb.",
    "Inventory_Report_Checks.xlsx",
    files.checks,
  );
  if ((meta.migrated || meta.assigned_codes) && files.updated) {
    makeDownload(
      errorDownloads,
      "Azərbaycan dilində giriş kitabı",
      `Yeni kodlar: ${meta.assigned_codes}. Məlumatları bu nüsxədə düzəldin.`,
      updatedFilename(currentFile.name),
      files.updated,
    );
  }
  errorPanel.hidden = false;
  setStatus("error", "Hesablama dayandırıldı", "Göstərilən məlumatları düzəldin və yenidən hesablayın.");
  finishRequest();
}

const workerUrl = new URL("./worker.js", import.meta.url);
workerUrl.search = new URL(import.meta.url).search;
const worker = new Worker(workerUrl, { type: "module" });

worker.addEventListener("message", ({ data }) => {
  if (data.type === "status") {
    if (data.state === "ready") {
      engineReady = true;
      calculateButton.disabled = !currentFile;
    }
    setStatus(data.state, data.title, data.detail);
    return;
  }
  if (data.type === "result") {
    if (data.meta.ok) renderSuccess(data.meta, data.files);
    else renderFailure(data.meta, data.files);
    settleRequest("resolve", data.meta);
    return;
  }
  if (data.type === "process-error" || data.type === "fatal") {
    if (data.type === "fatal") engineReady = false;
    setStatus("error", "Hesablama alınmadı", data.message);
    finishRequest();
    settleRequest("reject", new Error(data.message));
  }
});

worker.addEventListener("error", () => {
  engineReady = false;
  setStatus(
    "error",
    "Hesablama modulu yüklənmədi",
    "İnternet bağlantısını yoxlayın və səhifəni yeniləyin.",
  );
  finishRequest();
  settleRequest("reject", new Error("Hesablama modulu əlçatan deyil."));
});

function startCalculation() {
  if (!currentFile) return Promise.reject(new Error("Əvvəlcə Excel faylını seçin."));
  if (!engineReady) return Promise.reject(new Error("Hesablama modulu hələ hazır deyil."));
  if (isWorking) return Promise.reject(new Error("Hesablama artıq davam edir."));
  clearResults();
  isWorking = true;
  changeFile.disabled = true;
  fileInput.disabled = true;
  calculateButton.disabled = true;
  calculateButton.textContent = "Hesablanır…";
  setStatus("working", "Kitab oxunur", "Məlumatlar yoxlanılır və çatışmayan kodlar verilir.");
  return new Promise((resolve, reject) => {
    pendingCompletion = { resolve, reject };
    currentFile.arrayBuffer().then((buffer) => {
      worker.postMessage(
        { type: "process", name: currentFile.name, buffer },
        [buffer],
      );
    }).catch((error) => {
      setStatus("error", "Fayl oxunmadı", error.message || String(error));
      finishRequest();
      settleRequest("reject", error);
    });
  });
}

calculateButton.addEventListener("click", () => {
  void startCalculation().catch(() => {});
});

function registerWebMcpTool() {
  const context = document.modelContext;
  if (!context?.registerTool) return;
  const lifecycle = new AbortController();
  const registration = context.registerTool(
    {
      name: "calculate_selected_inventory_workbook",
      title: "Seçilmiş anbar kitabını hesabla",
      description: "İstifadəçinin seçdiyi XLSX kitabını aylıq orta maya dəyəri ilə hesablayır və faylları endirmək üçün hazırlayır.",
      inputSchema: {
        type: "object",
        properties: {},
        additionalProperties: false,
      },
      annotations: {
        readOnlyHint: false,
        untrustedContentHint: true,
      },
      async execute(input) {
        if (!input || typeof input !== "object" || Array.isArray(input) || Object.keys(input).length) {
          throw new Error("Bu əməliyyat əlavə parametr qəbul etmir.");
        }
        const meta = await startCalculation();
        return {
          status: meta.ok ? "success" : "input_error",
          assignedCodes: meta.assigned_codes,
          warnings: meta.warning_count,
          errors: meta.issues?.length || 0,
        };
      },
    },
    { signal: lifecycle.signal },
  );
  void Promise.resolve(registration).catch(() => {});
  window.addEventListener("pagehide", () => lifecycle.abort(), { once: true });
}

registerWebMcpTool();
