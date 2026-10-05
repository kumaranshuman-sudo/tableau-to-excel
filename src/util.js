/* Small shared helpers: logging, DOM child lookup, text/number normalisation, export file name. */
import { FORMAT_CONFIG } from "./config.js";

/* ── small helpers ─────────────────────────────────────────────────────── */
export function tfLog(...a) { if (FORMAT_CONFIG.debug) console.log("[Format]", ...a); }

export function tfKids(el, tag) {
  if (!el) return [];
  return Array.from(el.childNodes).filter(n => n.nodeType === 1 && (!tag || n.tagName === tag));
}

export function tfKid(el, tag) { return tfKids(el, tag)[0] || null; }

export function tfNorm(s) { return String(s == null ? "" : s).toLowerCase().replace(/\s+/g, " ").trim(); }

export function tfNum(v) { const n = parseFloat(v); return isFinite(n) ? n : undefined; }

/** default name of the exported file: "<workbook> - <dashboard>.xlsx", without characters Windows forbids
 * @param {string | null} workbookFile the loaded .twb / .twbx @param {string} dashboardName @returns {string} */
export function exportFileName(workbookFile, dashboardName) {
  const clean = s => String(s || "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "").replace(/\s+/g, " ").trim().replace(/[. ]+$/, "");
  const workbook = clean(String(workbookFile || "").replace(/\.twbx?$/i, ""));
  const dashboard = clean(dashboardName) || "Dashboard Export";
  const base = workbook && workbook.toLowerCase() !== dashboard.toLowerCase() ? `${workbook} - ${dashboard}` : dashboard;
  return base.slice(0, 150).replace(/[. ]+$/, "") + ".xlsx";
}

/** copy of o without its undefined properties
 * @template T @param {T} o @returns {T} */
export function tfDefined(o) {
  const r = /** @type {T} */ ({});
  for (const k in o) if (o[k] !== undefined) r[k] = o[k];
  return r;
}
