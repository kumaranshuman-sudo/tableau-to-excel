/* Loading the .twb/.twbx, remembering it (settings + IndexedDB) and the save location. */
import JSZip from "jszip";
import { parseTableauFormatting } from "../twb/parser.js";
import { showWorkbookLabel } from "./status.js";
import { checkWorkbookMatch } from "../visual/workbook-match.js";

/* =============================================================================
 * parseTwbXmlInBrowser(xmlString) - Parses TITLES from XML
 * ============================================================================= */
export function parseTwbXmlInBrowser(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const titleMap = {};

  const worksheetNodes = doc.getElementsByTagName("worksheet");

  for (let i = 0; i < worksheetNodes.length; i++) {
    const wsNode = worksheetNodes[i];
    const internalName = wsNode.getAttribute("name");
    if (!internalName) continue;

    const layoutNode = getDirectChildByTag(wsNode, "layout-options")
                    || getDirectChildByTag(wsNode, "layout");
    if (!layoutNode) continue;

    const titleNode = getDirectChildByTag(layoutNode, "title");
    if (!titleNode) continue;

    const fmtNode = titleNode.getElementsByTagName("formatted-text")[0];
    if (!fmtNode) continue;

    const runNodes = fmtNode.getElementsByTagName("run");
    const displayTitle = Array.from(runNodes)
      .map(r => r.textContent || "")
      .join("")
      .trim();

    if (displayTitle) {
      titleMap[internalName] = displayTitle;
      console.log(`[TWB parse] Title: "${internalName}" → "${displayTitle}"`);
    }
  }

  return titleMap;
}

export function getDirectChildByTag(parent, tagName) {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child.nodeType === 1 && child.tagName === tagName) return child;
  }
  return null;
}

/* =============================================================================
 * Workbook file handle – the save dialog opens in the loaded workbook's folder.
 * Pages cannot read a file's path, but a FileSystemFileHandle from showOpenFilePicker
 * can be passed to showSaveFilePicker({ startIn }). It is kept in IndexedDB so the
 * next Tableau session still starts there. Runtimes without the File System Access
 * API keep the plain <input type="file"> + download behaviour.
 * ============================================================================= */
export let workbookFileHandle = null;

export const HANDLE_DB = "mark2table", HANDLE_STORE = "handles", HANDLE_KEY = "workbook";

export function openHandleDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(HANDLE_DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(HANDLE_STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function storeWorkbookHandle(handle) {
  try {
    const db = await openHandleDb();
    db.transaction(HANDLE_STORE, "readwrite").objectStore(HANDLE_STORE).put(handle, HANDLE_KEY);
  } catch (e) {
    console.warn("[Workbook] could not remember the workbook folder:", e.message);
  }
}

/* parsed workbook model: Tableau's extension settings cannot hold a large model
 * (saveAsync fails) and it would be lost on the next reload → keep a copy here too */
export async function storeFormatModel(fileName, model, titleMap) {
  try {
    const db = await openHandleDb();
    // one entry per workbook file: any dashboard of a workbook loaded once is matched again later
    db.transaction(HANDLE_STORE, "readwrite").objectStore(HANDLE_STORE)
      .put({ fileName, model, titleMap, savedAt: Date.now() }, "model:" + String(fileName).toLowerCase());
  } catch (e) {
    console.warn("[Workbook] could not keep the workbook formatting in browser storage:", e.message);
  }
}

export async function restoreFormatModels() {
  try {
    const db = await openHandleDb();
    return await new Promise(resolve => {
      const out = [];
      const req = db.transaction(HANDLE_STORE).objectStore(HANDLE_STORE).openCursor();
      req.onsuccess = () => {
        const cur = req.result;
        if (!cur) return resolve(out);
        const key = String(cur.key);
        if ((key.startsWith("model:") || key === "formatModel") && cur.value && cur.value.model) out.push(cur.value);
        cur.continue();
      };
      req.onerror = () => resolve(out);
    });
  } catch (e) {
    return [];
  }
}

export async function restoreWorkbookHandle() {
  try {
    const db = await openHandleDb();
    return await new Promise(resolve => {
      const req = db.transaction(HANDLE_STORE).objectStore(HANDLE_STORE).get(HANDLE_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch (e) {
    return null;
  }
}

/* =============================================================================
 * loadWorkbookFile() - Loads .twb/.twbx and parses BOTH titles AND colors
 * ============================================================================= */
export async function loadWorkbookFile() {
  if (typeof window.showOpenFilePicker === "function") {
    try {
      const [handle] = await window.showOpenFilePicker({
        types: [{ description: "Tableau workbook", accept: { "application/octet-stream": [".twb", ".twbx"] } }]
      });
      const file = await handle.getFile();
      const titleMap = await readWorkbookFile(file);
      if (Object.keys(titleMap).length || FORMAT_MODEL_CACHE) {
        workbookFileHandle = handle;
        await storeWorkbookHandle(handle);
      }
      return titleMap;
    } catch (err) {
      if (err && err.name === "AbortError") return {};            // picker cancelled
      console.warn("[loadWorkbookFile] file picker unavailable, using file input:", err && err.message);
    }
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".twb,.twbx";
    input.style.display = "none";
    document.body.appendChild(input);

    input.onchange = async (event) => {
      document.body.removeChild(input);
      const file = /** @type {HTMLInputElement} */ (event.target).files[0];
      if (!file) {
        console.log("[loadWorkbookFile] No file selected");
        resolve({});
        return;
      }
      resolve(await readWorkbookFile(file));
    };

    input.oncancel = () => {
      document.body.removeChild(input);
      resolve({});
    };

    input.click();
  });
}

/* parses a picked .twb/.twbx: titles + format model, saved to the extension settings */
export async function readWorkbookFile(file) {
  console.log(`[loadWorkbookFile] Reading: ${file.name}`);

  try {
    let xmlString;
    const ext = file.name.split(".").pop().toLowerCase();

    if (ext === "twb") {
      xmlString = await new Promise((res, rej) => {
        const reader = new FileReader();
        reader.onload = (e) => res(e.target.result);
        reader.onerror = () => rej(new Error("FileReader failed reading .twb"));
        reader.readAsText(file, "utf-8");
      });
    } else if (ext === "twbx") {
      const arrayBuffer = await new Promise((res, rej) => {
        const reader = new FileReader();
        reader.onload = (e) => res(e.target.result);
        reader.onerror = () => rej(new Error("FileReader failed reading .twbx"));
        reader.readAsArrayBuffer(file);
      });

      const zip = await JSZip.loadAsync(arrayBuffer);
      const twbEntry = Object.values(zip.files).find(
        f => !f.dir && /\.twb$/i.test(f.name)
      );

      if (!twbEntry) {
        const entries = Object.values(zip.files).filter(f => !f.dir).map(f => f.name).slice(0, 20);
        throw new Error(`No .twb file found inside the .twbx archive. Entries: ${entries.join(", ") || "none"}`);
      }

      xmlString = await twbEntry.async("string");
    } else {
      throw new Error(`Unsupported file type ".${ext}". Please select a .twb or .twbx file.`);
    }

    const titleMap = parseTwbXmlInBrowser(xmlString);
    console.log(`[loadWorkbookFile] Parsed ${Object.keys(titleMap).length} titles`);

    // ── NEW: one DOM-based pass extracts all formatting (fonts, colours, number formats…)
    const formatModel = parseTableauFormatting(xmlString);
    FORMAT_MODEL_CACHE = formatModel;
    await storeFormatModel(file.name, formatModel, titleMap);
    tableau.extensions.settings.set("twbTitleMap", JSON.stringify(titleMap));
    try {
      tableau.extensions.settings.set("twbFormatModel", JSON.stringify(formatModel));
    } catch (e) {
      try { tableau.extensions.settings.erase("twbFormatModel"); } catch (e2) { /* nothing saved */ }
      console.warn("[loadWorkbookFile] Format model too large for settings – kept in memory only:", e.message);
    }
    tableau.extensions.settings.set("twbFileName", file.name);
    try {
      await tableau.extensions.settings.saveAsync();
    } catch (e) {
      console.warn("[loadWorkbookFile] settings.saveAsync failed (model kept in memory):", e.message);
    }

    const sheetCount = Object.keys(formatModel.sheets).length;
    const colorCount = Object.values(formatModel.sheets).filter(s => s.panes.some(p => p.encodings.some(e => e.channel === "color"))).length;
    showWorkbookLabel(file.name, `${Object.keys(titleMap).length} titles, formatting for ${sheetCount} sheets (${colorCount} with colour)`, formatModel);

    return titleMap;

  } catch (err) {
    console.error("[loadWorkbookFile] Error:", err.message);
    alert(`Could not read workbook file:\n${err.message}`);
    return {};
  }
}

/* Asks where to save BEFORE the export runs: the picker needs the click's user
 * activation, which the data fetch would outlast. → handle, null (no picker API /
 * picker failed → plain download), or "cancelled". */
export async function chooseSaveTarget(suggestedName) {
  if (typeof window.showSaveFilePicker !== "function") return null;
  const options = {
    suggestedName,
    types: [{ description: "Excel workbook",
              accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
  };
  if (!workbookFileHandle) workbookFileHandle = await restoreWorkbookHandle();
  try {
    return await window.showSaveFilePicker(workbookFileHandle ? { ...options, startIn: workbookFileHandle } : options);
  } catch (err) {
    if (err && err.name === "AbortError") return "cancelled";
    if (workbookFileHandle) {
      try {                                   // stale handle (file moved / deleted) → default folder
        return await window.showSaveFilePicker(options);
      } catch (err2) {
        if (err2 && err2.name === "AbortError") return "cancelled";
      }
    }
    console.warn("[Export] save dialog unavailable, using browser download:", err && err.message);
    return null;
  }
}

/* =============================================================================
 * getTitleMap() - Gets titles from settings   * ============================================================================= */
export function getTitleMap() {
  try {
    const saved = tableau.extensions.settings.get("twbTitleMap");
    if (!saved) return RESTORED_TITLE_MAP || {};
    return JSON.parse(saved);
  } catch (err) {
    console.warn("[getTitleMap] Could not read settings:", err.message);
    return {};
  }
}

/* =============================================================================
 * getFormatModel() - format model from memory, else from extension settings
 * ============================================================================= */
export let FORMAT_MODEL_CACHE = null;

export function getFormatModel() {
  if (FORMAT_MODEL_CACHE) return FORMAT_MODEL_CACHE;
  try {
    const saved = tableau.extensions.settings.get("twbFormatModel");
    FORMAT_MODEL_CACHE = saved ? JSON.parse(saved) : null;
  } catch (err) {
    console.warn("[getFormatModel] Could not read settings:", err.message);
  }
  return FORMAT_MODEL_CACHE;
}

/* workbook model for THIS dashboard: memory / settings first, else the remembered workbook
 * whose worksheets match the dashboard best (no Load Workbook needed again) */
export let RESTORED_TITLE_MAP = null;

export let RESTORED_FILE_NAME = null;

export async function ensureFormatModel() {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  const rank = m => { const r = checkWorkbookMatch(m, dashboard); return !r ? -1 : r.level === "ok" ? 2 + r.matched : r.level === "partial" ? 1 + r.matched / (r.total || 1) : 0; };
  const current = getFormatModel();
  if (current && rank(current) >= 2) return current;
  const saved = await restoreFormatModels();
  /** @type {{ model: FormatModel, titleMap?: Record<string, string>, fileName: string } | null} */
  let best = null;
  let bestRank = current ? rank(current) : 0;
  saved.forEach(rec => { const r = rank(rec.model); if (r > bestRank) { best = rec; bestRank = r; } });
  if (best) {
    FORMAT_MODEL_CACHE = best.model;
    RESTORED_TITLE_MAP = best.titleMap || null;
    RESTORED_FILE_NAME = best.fileName;
    console.log(`[Workbook] using remembered workbook ${best.fileName} for this dashboard`);
  }
  return FORMAT_MODEL_CACHE;
}

/* main.js restores the cached model from the extension settings at start-up */
export function setFormatModelCache(model) {
  FORMAT_MODEL_CACHE = model;
}
