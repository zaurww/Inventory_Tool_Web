import { loadPyodide } from "https://cdn.jsdelivr.net/pyodide/v314.0.7/full/pyodide.mjs";

const PYODIDE_INDEX = "https://cdn.jsdelivr.net/pyodide/v314.0.7/full/";
const PYTHON_SOURCES = [
  "input_layout.py",
  "inventory.py",
  "prepare_input.py",
  "browser_api.py",
  "migrate_input.py",
  "report_locale.py",
];
const INPUT_PATH = "/tmp/inventory_input.xlsx";
const UPDATED_PATH = "/tmp/inventory_input_updated.xlsx";
const REPORT_PATH = "/tmp/inventory_report.xlsx";
const CHECKS_PATH = "/tmp/inventory_checks.xlsx";

let pyodide = null;

function status(state, title, detail) {
  self.postMessage({ type: "status", state, title, detail });
}

async function loadSource(name) {
  const response = await fetch(new URL(`./py/${name}`, self.location.href), {
    cache: "no-store",
  });
  if (!response.ok) throw new Error(`${name} modulu tapılmadı.`);
  pyodide.FS.writeFile(`/inventory_app/${name}`, await response.text(), {
    encoding: "utf8",
  });
}

async function initialize() {
  status("working", "Hesablama mühərriki yüklənir", "İlk açılış bir qədər vaxt apara bilər.");
  pyodide = await loadPyodide({ indexURL: PYODIDE_INDEX });
  status("working", "Excel modulu yüklənir", "XLSX kitablarının oxunması və yazılması hazırlanır.");
  await pyodide.loadPackage("micropip");
  const micropip = pyodide.pyimport("micropip");
  try {
    await micropip.install(["openpyxl==3.1.5", "et-xmlfile==2.0.0"]);
  } finally {
    micropip.destroy();
  }
  pyodide.FS.mkdirTree("/inventory_app");
  for (const source of PYTHON_SOURCES) await loadSource(source);
  await pyodide.runPythonAsync(`
import sys
if "/inventory_app" not in sys.path:
    sys.path.insert(0, "/inventory_app")
import browser_api
`);
  status("ready", "Hesablama modulu hazırdır", "Excel faylını seçin və hesablamanı başladın.");
}

const readyPromise = initialize().catch((error) => {
  self.postMessage({
    type: "fatal",
    message: error.message || String(error),
  });
  throw error;
});

function removeIfPresent(path) {
  try {
    pyodide.FS.unlink(path);
  } catch {
    // A previous result is optional.
  }
}

function readBuffer(path) {
  const bytes = pyodide.FS.readFile(path);
  return bytes.slice().buffer;
}

async function processWorkbook(name, buffer) {
  await readyPromise;
  for (const path of [INPUT_PATH, UPDATED_PATH, REPORT_PATH, CHECKS_PATH]) {
    removeIfPresent(path);
  }
  pyodide.FS.writeFile(INPUT_PATH, new Uint8Array(buffer));
  pyodide.globals.set("web_input_name", name);
  status("working", "Aylıq orta maya dəyəri hesablanır", "Statik hesabatlar və yoxlamalar hazırlanır.");
  try {
    const metadataJson = await pyodide.runPythonAsync(`
import json
from pathlib import Path
from browser_api import process_workbook

_web_result = process_workbook(Path("/tmp/inventory_input.xlsx").read_bytes(), str(web_input_name))
Path("/tmp/inventory_input_updated.xlsx").write_bytes(_web_result["input_bytes"])
if _web_result["ok"]:
    Path("/tmp/inventory_report.xlsx").write_bytes(_web_result["report_bytes"])
else:
    Path("/tmp/inventory_checks.xlsx").write_bytes(_web_result["checks_bytes"])
_web_metadata = {
    key: value for key, value in _web_result.items()
    if key not in {"input_bytes", "report_bytes", "checks_bytes"}
}
json.dumps(_web_metadata, ensure_ascii=False)
`);
    const meta = JSON.parse(metadataJson);
    const files = { updated: readBuffer(UPDATED_PATH) };
    if (meta.ok) files.report = readBuffer(REPORT_PATH);
    else files.checks = readBuffer(CHECKS_PATH);
    const transfers = Object.values(files);
    self.postMessage({ type: "result", meta, files }, transfers);
  } finally {
    pyodide.globals.delete("web_input_name");
    for (const key of ["_web_result", "_web_metadata"]) {
      if (pyodide.globals.has(key)) pyodide.globals.delete(key);
    }
    for (const path of [INPUT_PATH, UPDATED_PATH, REPORT_PATH, CHECKS_PATH]) {
      removeIfPresent(path);
    }
  }
}

self.addEventListener("message", async ({ data }) => {
  if (data.type !== "process") return;
  try {
    await processWorkbook(data.name, data.buffer);
  } catch (error) {
    const message = error.message || String(error);
    self.postMessage({
      type: "process-error",
      message: message.split("\n").filter(Boolean).at(-1) || message,
    });
  }
});
