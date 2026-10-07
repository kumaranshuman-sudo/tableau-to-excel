/* ══════════════════════════════════════════════════════════════════════════════════════════════════════
 * build_table_copy.js — Tableau dashboard extension: Export to Excel (NeenOpal)
 * ──────────────────────────────────────────────────────────────────────────────────────────────────────
 * ONE file, merged from two code bases that grew out of the same v0.6.0 build_table_copy.js:
 *   • the visual export engine (Export_extension_XML / Export2Sheet, src/ modules): tables, KPI cards,
 *     text boxes, images, backgrounds, Tableau-exact layout, NATIVE EXCEL CHARTS (bar, line, area, pie,
 *     donut, scatter, bubble, combo, treemap, waterfall, box plot, Gantt …), Conversion Report sheet;
 *   • the Tableau Cloud features of build_table_copy.js: the workbook loaded automatically through the
 *     backend (Worker), every visible dashboard read through hidden Embedding API views, the one-time
 *     Tableau sign-in, and ONE Excel sheet per dashboard.
 *
 * Each "═══ <path> ═══" banner below starts one of the original modules (same names, same order as the
 * module tree, so a change can be traced back to it). The src/ modules are folded in up to "Donut centre
 * totals and waterfall grand totals" (eab0bfb); this file is now the only source. Sections that belong to the merge:
 *   cloud/backend.js      auto-load from Tableau Cloud (Advanced → Backend URL)
 *   cloud/dashboards.js   the other dashboards (hidden views, sign-in, progress screen, sheet names)
 *   ui/workbook-store.js  applyWorkbookXml(): manual 📁 load and auto-load share one path
 *   export/export.js      exportToExcel() → writeDashboardSheet() per dashboard → charts injected per sheet
 *   export/report.js      writeConversionReports(): one report for every exported dashboard
 *   main.js               panel wiring (start-up)
 * Settings that change the output: FORMAT_CONFIG (config.js section).
 *
 * The page (index.html) must load, before this file:
 *   ./js/tableau.extensions.1.latest.js, ExcelJS 4.4 (window.ExcelJS) and JSZip 3.10 (window.JSZip).
 * Panel element ids used: twb_file_label, load_workbook_btn, export_button, export_status,
 *   export_all (checkbox "All dashboards"), backend_url, backend_key, backend_save (Advanced settings).
 * ══════════════════════════════════════════════════════════════════════════════════════════════════════ */
"use strict";
(function () {

if (typeof ExcelJS === "undefined" || typeof JSZip === "undefined") {
  console.error("[Export] ExcelJS and JSZip must be loaded (script tags in index.html) before build_table_copy.js");
}

/* ═══ format/number-format.js ═════════════════════════════════════════════════════════════════════ */
/* Tableau number formats → Excel number formats. */

/* Tableau text-format → Excel numFmt.
 * - strips Tableau's type prefix (n/c/p/e/*); locale "standard" codes (C1033…) → null (inferred instead),
 *   while short ones are formats ("p0%" = whole percent)
 * - moves thousands-scaling commas behind the decimals: "#,##0,.0K" → "#,##0.0,\"K\""
 * - quotes every literal letter (Excel rejects bare K, M, yrs …)
 * - validates the result; anything doubtful → null so we never write an invalid format */
function tableauToExcelNumFmt(raw) {
  if (!raw) return null;
  let s = String(raw);
  if (/^[A-Za-z]\d{3,}%?$/.test(s)) return null;
  if (/^[ncpes*]/i.test(s) && !/^[#0]/.test(s)) s = s.slice(1);
  const sections = tfSplitSections(s);
  if (!sections || sections.length > 4) return null;
  const out = sections.map(tfConvertSection);
  if (out.some(x => x === null)) return null;
  const res = out.join(";");
  return /[0#]/.test(res) ? res : null;
}

/**
 * A number as Excel shows it in a format, for text Excel does not format itself (label texts taken from
 * cells): literal text, digits with thousands separators, decimals (0 required, # optional), thousands
 * scaling (trailing commas: ,, = millions), percent; sections positive;negative;zero.
 * @param {number} value @param {string} fmt Excel number format @returns {string}
 */
function tfFormatNumber(value, fmt) {
  if (typeof value !== "number" || !isFinite(value)) return value == null ? "" : String(value);
  const secs = tfSplitSections(String(fmt || "General")) || ["General"];
  let sec = secs[0], sign = value < 0 ? "-" : "";
  if (value < 0 && secs.length > 1) { sec = secs[1]; sign = ""; }   // the negative section shows the size
  else if (value === 0 && secs.length > 2) sec = secs[2];
  const x = Math.abs(value);
  let pre = "", post = "", pattern = "", phase = 0;                 // 0 before the number, 1 in it, 2 after
  const text = t => { if (phase === 1) phase = 2; if (phase === 0) pre += t; else post += t; };
  for (let i = 0; i < sec.length; i++) {
    const ch = sec[i];
    if (ch === '"') { const j = sec.indexOf('"', i + 1); text(sec.slice(i + 1, j < 0 ? sec.length : j)); i = j < 0 ? sec.length : j; }
    else if (ch === "\\") { text(sec[i + 1] || ""); i++; }
    else if (ch === "[") { const j = sec.indexOf("]", i); i = j < 0 ? sec.length : j; }      // [Red], [>100]
    else if (ch === "_") { text(" "); i++; }                                                 // _) = a space
    else if (ch === "*") i++;                                                                 // fill character
    else if (/[0#?.,%]/.test(ch) && phase < 2) { phase = 1; pattern += ch; }
    else if (/general/i.test(sec.slice(i, i + 7)) && phase === 0) { pattern = "G"; phase = 1; i += 6; }
    else text(ch);
  }
  if (pattern === "G" || !/[0#?]/.test(pattern)) {
    return sign + pre + (pattern === "G" ? String(+x.toPrecision(10)) : "") + post;
  }
  const pct = (pattern.match(/%/g) || []).length;
  const p = pattern.replace(/%/g, "");
  const dot = p.indexOf(".");
  let ip = dot >= 0 ? p.slice(0, dot) : p, dp = dot >= 0 ? p.slice(dot + 1) : "";
  let scale = 0;                                                   // commas after the last digit: ÷ 1000 each
  const it = ip.match(/,+$/); if (it) { scale += it[0].length; ip = ip.slice(0, -it[0].length); }
  const dt = dp.match(/,+$/); if (dt) { scale += dt[0].length; dp = dp.slice(0, -dt[0].length); }
  const minDec = (dp.match(/0/g) || []).length, maxDec = (dp.match(/[0#?]/g) || []).length;
  let s = (x / Math.pow(1000, scale) * Math.pow(100, pct)).toFixed(maxDec);
  if (maxDec > minDec) s = s.replace(new RegExp(`0{1,${maxDec - minDec}}$`), "").replace(/\.$/, "");
  let [whole, frac] = s.split(".");
  const minInt = (ip.match(/0/g) || []).length;
  if (whole.length < minInt) whole = whole.padStart(minInt, "0");
  if (!minInt && whole === "0" && frac) whole = "";
  if (ip.includes(",")) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return sign + pre + whole + (frac ? "." + frac : "") + "%".repeat(pct) + post;
}

/** an Excel number format with literal text around the number in every section ("4.0" → "4.0 DAYS")
 * @param {string} fmt @param {string} prefix @param {string} suffix @returns {string} */
function tfWrapNumFmt(fmt, prefix, suffix) {
  const q = s => s ? `"${s.replace(/"/g, "")}"` : "";
  return (tfSplitSections(fmt || "General") || ["General"]).slice(0, 3).map(sec => q(prefix) + (sec || "General") + q(suffix)).join(";");
}

function tfSplitSections(s) {
  const secs = []; let cur = "", q = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') q = !q;
    if (ch === "\\" && !q) { cur += ch + (s[i + 1] || ""); i++; continue; }
    if (ch === ";" && !q) { secs.push(cur); cur = ""; continue; }
    cur += ch;
  }
  if (q) return null;                        // unbalanced quotes
  secs.push(cur);
  return secs;
}

function tfConvertSection(sec) {
  // tokens: quoted literal | escaped char | number pattern | other char
  const toks = [];
  const re = /"[^"]*"|\\.|[#0?,.]+(?:[eE][+-][0#]+)?|[\s\S]/g;
  let m, numDone = false;
  while ((m = re.exec(sec))) {
    const t = m[0];
    if (t[0] === '"' || t[0] === "\\") { toks.push(t); continue; }
    if (/^[#0?,.]/.test(t) && /[#0?]/.test(t) && !numDone) {
      numDone = true;
      const [intPart, ...decParts] = t.split(".");
      const dec = decParts.join("");
      const intTrail = (intPart.match(/,+$/) || [""])[0];                 // "#,##0," → scale
      const decScale = (dec.match(/,+$/) || [""])[0];
      const intClean = intPart.slice(0, intPart.length - intTrail.length);
      const decClean = dec.replace(/,+$/, "");
      toks.push(intClean + (decParts.length ? "." + decClean : "") + intTrail + decScale);
      continue;
    }
    if (/[A-Za-z]/.test(t)) {                                              // literal letter → quote
      toks.push('"' + t + '"');
      continue;
    }
    toks.push(t);
  }
  // merge adjacent quoted literals: "K""g" → "Kg"
  return toks.join("").replace(/""/g, "");
}

/* Parse a Tableau formatted value, build an Excel format and only accept it
 * if it reproduces the same number (round-trip check). */
function inferExcelNumFmt(formatted, value) {
  if (formatted == null || typeof value !== "number" || !isFinite(value)) return null;
  const f = String(formatted).trim();
  const m = f.match(/^(\(?)([^\d(]*?)(-?)(\d[\d.,\s\u00A0']*)(.*?)(\)?)$/);
  if (!m) return null;
  const [, lpar, prefixRaw, minus, core, suffixRaw, rpar] = m;
  const neg = !!(lpar && rpar) || !!minus || /^-/.test(prefixRaw);
  const prefix = prefixRaw.replace(/-/g, "").trim() ? prefixRaw.replace(/-/g, "") : "";
  const suffix = suffixRaw;

  // decide decimal separator by trying both interpretations
  const tryParse = dec => {
    const grpSrc = dec === "." ? "[,\\s\\u00A0']" : "[.\\s\\u00A0']";
    const clean = core.trim().replace(new RegExp(grpSrc, "g"), "").replace(dec, ".");
    const n = parseFloat(clean);
    const decimals = clean.includes(".") ? clean.split(".")[1].length : 0;
    return { n, decimals, grouped: new RegExp(grpSrc).test(core.trim()) };
  };
  const unitScale = { "%": 100, K: 1e-3, M: 1e-6, B: 1e-9, Bn: 1e-9 };
  const unit = (suffix.trim().match(/^(%|K|M|Bn|B)\b/) || suffix.trim().match(/^%/) || [])[0] || null;
  const scaled = Math.abs(value) * (unit ? unitScale[unit] : 1);

  for (const dec of [".", ","]) {
    const p = tryParse(dec);
    if (!isFinite(p.n)) continue;
    const tol = 0.5 * Math.pow(10, -p.decimals) + 1e-9 * Math.max(1, scaled);
    let plainPercent = false;
    if (Math.abs(scaled - p.n) > tol) {
      if (unit === "%" && Math.abs(Math.abs(value) - p.n) <= tol) plainPercent = true; // value already ×100
      else continue;
    }
    // Tableau's "Automatic" format pads to ~6 significant digits ("348.000", "4.50000",
    // "24.86108%"); 3+ decimals are that padding, not a chosen format → keep what the value needs
    let decimals = p.decimals;
    if (decimals >= 3) {
      const frac = (String(p.n).split(".")[1] || "").length;
      decimals = Math.min(decimals, frac);
      if (decimals >= 3) decimals = Math.max(0, Math.min(decimals, 2 - Math.floor(Math.log10(Math.abs(p.n) || 1))));
    }
    let body = (p.grouped ? "#,##0" : "0") + (decimals ? "." + "0".repeat(decimals) : "");
    const q = t => t ? '"' + t.replace(/"/g, '""') + '"' : "";
    let rest = suffix;
    if (unit === "%" && !plainPercent) { body += "%"; rest = suffix.replace("%", ""); }
    else if (unit && unit !== "%") { body += ",".repeat({ K: 1, M: 2, B: 3, Bn: 3 }[unit]) + q(unit); rest = suffix.replace(unit, ""); }
    const pos = q(prefix) + body + q(rest);
    if (!neg) return pos;
    if (lpar && rpar) return pos + ";(" + pos + ")";
    return minus ? pos + ";" + q(prefix) + "-" + body + q(rest) : pos + ";-" + pos;
  }
  return null;
}


/* ═══ config.js ═══════════════════════════════════════════════════════════════════════════════════ */
/* Extension-wide settings and enumerations. */

/* ══════════════════════════════════════════════════════════════════════════
 * TABLEAU → EXCEL FORMAT ENGINE
 * ──────────────────────────────────────────────────────────────────────────
 * 1. parseTableauFormatting(xml)  → plain JSON "format model" (DOM based)
 * 2. createSheetFormatter(model, sheetName) → resolves the Tableau cascade
 *      defaults → workbook <style> → worksheet <style> → element → field
 *      → customized-label run
 * 3. tfExcel*() helpers → convert resolved props into ExcelJS styles
 * ══════════════════════════════════════════════════════════════════════════ */

const FORMAT_CONFIG = {
  substituteTableauFonts: true,   // "Tableau Book" etc. are only installed with Tableau Desktop
  tableauFontSubstitute: "Arial",
  writeNativeNumbers: true,       // write real numbers + Excel numFmt instead of text
  applyFallbackHeatmap: false,    // old red/green heatmap when Tableau has no color (not in Tableau → off)
  nativeCharts: true,             // bar/line/area/pie/scatter/combo → editable native Excel charts
  chartPolicy: "data",            // charts that cannot be drawn: "skip" | "data" (export their data as a plain table)
  groupOverflowRows: true,        // collapse rows beyond ROW_GROUP_THRESHOLD into an expandable [+]/[-] group
  autoFilter: "largest",          // Excel allows ONE autofilter per sheet: "largest" table | "none"
  textBoxHeaders: true,           // dashboard text boxes over a table as its column headers (one box per column or per group of panes)
  textBoxTitle: false,            // text boxes are drawn in place now; true = also use the top one as the sheet title
  linkText: "url",                // URL-action cells: "url" shows the link itself, any other string is shown as the text
  sheetGridlines: false,          // Excel's cell grid on the dashboard sheet – Tableau dashboards have none
  printFitToWidth: true,          // printing / PDF: the dashboard one page wide, landscape when it is wider than tall
  tableauTextScale: null,         // Tableau label text vs its pixels (Windows display scaling) for pie label placement; null = the browser's devicePixelRatio
  conversionReport: true,         // a "Conversion Report" sheet: every visual, how it was converted (NATIVE … TABLE_FALLBACK) and why
  fallbackNotes: true,            // a note on each visual exported as its data table, naming the Tableau visual and the reason
  debug: true                     // true: dump the parsed workbook model and [Format] traces to the console
};

const VISUAL_TYPES = Object.freeze({
  TABLE: "TABLE",
  KPI: "KPI",
  HEATMAP: "HEATMAP",
  BAR: "BAR",
  COLUMN: "COLUMN",
  LINE: "LINE",
  AREA: "AREA",
  PIE: "PIE",
  SCATTER: "SCATTER",
  BUBBLE: "BUBBLE",
  COMBO: "COMBO",
  HISTOGRAM: "HISTOGRAM",
  TREEMAP: "TREEMAP",
  WATERFALL: "WATERFALL",
  BOXPLOT: "BOXPLOT",
  MAP: "MAP",
  MAP_FILLED: "MAP_FILLED",
  GANTT: "GANTT",
  CUSTOM: "CUSTOM",
  UNKNOWN: "UNKNOWN"
});

/* Approximate Tableau Desktop defaults. Compare against one of your exports
 * and tweak these once – everything explicitly set in the TWB overrides them. */
const TABLEAU_DEFAULTS = {
  worksheet:  { fontName: "Tableau Book", fontSize: 9, color: "FF333333" },
  fieldLabel: { bold: true },
  title:      { fontName: "Tableau Book", fontSize: 15, color: "FF333333" },
  dashTitle:  { fontName: "Tableau Book", fontSize: 18, color: "FF333333" },
  divider:    { visible: true, style: "thin", color: "FFD4D4D4", level: 1 }
};

/* XML element names used in <style-rule element='…'>. Arrays = aliases seen
 * across Tableau versions; unknown ones are simply ignored. */
const TF_ELEMENTS = {
  datalabel:  ["datalabel"],        // Marks card → Label → font (per pane)
  sheet:      ["all", "worksheet"],
  title:      ["title", "worksheet-title"],
  dashTitle:  ["dash-title", "dashboard-title", "title"],
  pane:       ["pane"],
  cell:       ["cell"],
  header:     ["header", "label"],   // "header" = box (shading/size), "label" = header TEXT (font/colour/underline/alignment)
  fieldLabel: ["field-labels", "field-labels-decoration"],
  divider:    ["table-div"]
};

const TF_DERIV_LABEL = {
  sum: "SUM", avg: "AVG", cnt: "CNT", ctd: "CNTD", min: "MIN", max: "MAX",
  med: "MEDIAN", attr: "ATTR", usr: "AGG", std: "STDEV", stdp: "STDEVP",
  var: "VAR", varp: "VARP", yr: "YEAR", qr: "QUARTER", mn: "MONTH", wk: "WEEK",
  dy: "DAY", hr: "HOUR", tyr: "YEAR", tqr: "QUARTER", tmn: "MONTH", twk: "WEEK", tdy: "DAY"
};


/* ═══ util.js ═════════════════════════════════════════════════════════════════════════════════════ */
/* Small shared helpers: logging, DOM child lookup, text/number normalisation, export file name. */

/* ── small helpers ─────────────────────────────────────────────────────── */
function tfLog(...a) { if (FORMAT_CONFIG.debug) console.log("[Format]", ...a); }

function tfKids(el, tag) {
  if (!el) return [];
  return Array.from(el.childNodes).filter(n => n.nodeType === 1 && (!tag || n.tagName === tag));
}

function tfKid(el, tag) { return tfKids(el, tag)[0] || null; }

function tfNorm(s) { return String(s == null ? "" : s).toLowerCase().replace(/\s+/g, " ").trim(); }

function tfNum(v) { const n = parseFloat(v); return isFinite(n) ? n : undefined; }

/** default name of the exported file: "<workbook> - <dashboard>.xlsx", without characters Windows forbids
 * @param {string | null} workbookFile the loaded .twb / .twbx @param {string} dashboardName @returns {string} */
function exportFileName(workbookFile, dashboardName) {
  const clean = s => String(s || "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "").replace(/\s+/g, " ").trim().replace(/[. ]+$/, "");
  const workbook = clean(String(workbookFile || "").replace(/\.twbx?$/i, ""));
  const dashboard = clean(dashboardName) || "Dashboard Export";
  const base = workbook && workbook.toLowerCase() !== dashboard.toLowerCase() ? `${workbook} - ${dashboard}` : dashboard;
  return base.slice(0, 150).replace(/[. ]+$/, "") + ".xlsx";
}

/** copy of o without its undefined properties
 * @template T @param {T} o @returns {T} */
function tfDefined(o) {
  const r = /** @type {T} */ ({});
  for (const k in o) if (o[k] !== undefined) r[k] = o[k];
  return r;
}


/* ═══ twb/field-ref.js ════════════════════════════════════════════════════════════════════════════ */
/* Tableau field references ("[ds].[sum:Sales:qk]") – parsing and comparison. */

/* ── Field references ──────────────────────────────────────────────────────
 * "[federated.0x].[sum:Sales:qk]" → {ds, deriv:"sum", name:"Sales", type:"qk"}
 * type: qk = continuous measure, nk/ok = discrete                          */
/** @param {string} ref @returns {FieldRef | null} */
function tfParseFieldRef(ref) {
  if (!ref) return null;
  ref = String(ref).trim().replace(/^"|"$/g, "");
  const parts = ref.match(/\[[^\]]*\]/g) || [];
  if (!parts.length) return { raw: ref, ds: null, deriv: null, name: ref, type: null, inner: ref };
  const inner = parts[parts.length - 1].slice(1, -1);
  const ds = parts.length > 1 ? parts[0].slice(1, -1) : null;
  const seg = inner.split(":");
  let deriv = null, name = inner, type = null;
  // "[usr:Calculation_1:ok:9]" → type is the nk/ok/qk segment, not the last one
  let ti = -1;
  for (let i = seg.length - 1; i >= 2; i--) if (/^(nk|ok|qk)$/.test(seg[i])) { ti = i; break; }
  if (ti >= 2) { deriv = seg[0] || null; type = seg[ti]; name = seg.slice(1, ti).join(":"); }
  else if (seg.length === 2 && seg[0] === "") name = seg[1];          // [:Measure Names]
  return { raw: ref, ds, deriv, name, type, inner };
}

/** a quick table calculation's measure: "[cum:sum:Sales:qk:7]" (running total) → "[sum:Sales:qk]"; null for any
 * other field @param {FieldRef} ref @returns {FieldRef | null} */
function tfTableCalcBase(ref) {
  if (!ref || !ref.deriv || !ref.type || !/^(sum|avg|cnt|ctd|min|max|med|attr|usr|std|stdp|var|varp|none):/i.test(ref.name || "")) return null;
  const inner = `${ref.name}:${ref.type}`;
  return tfParseFieldRef(ref.ds ? `[${ref.ds}].[${inner}]` : `[${inner}]`);
}

/** The field references in a shelf or label text. A relationship model's "Count of <table>" is one reference of
 * three parts: "[ds].[__tableau_internal_object_id__].[cnt:Orders.csv_…:qk]".
 * @param {string} text @returns {FieldRef[]} */
function tfExtractRefs(text) {
  if (!text) return [];
  const m = String(text).match(/(?:\[[^\]]+\]\.)?(?:\[__tableau_internal_object_id__\]\.)?\[[^\]]+\]/g) || [];
  return m.map(tfParseFieldRef).filter(Boolean);
}

function tfRefKey(r) { return ((r.ds || "") + "|" + r.inner).toLowerCase(); }

/** @param {Partial<FieldRef> | null} a @param {Partial<FieldRef> | null} b @returns {boolean} */
function tfSameField(a, b) {
  if (!a || !b) return false;
  if (a.inner.toLowerCase() === b.inner.toLowerCase()) return true;
  return tfNorm(a.name) === tfNorm(b.name) && (a.deriv || "none") === (b.deriv || "none");
}


/* ═══ charts/model/axes.js ════════════════════════════════════════════════════════════════════════ */
/* The worksheet's own axis and label settings on a chart spec: hidden axes, axis titles, Edit Axis ranges
 * and tick spacing, tick number formats, grid lines, axis rulers, mark label font, position and text. */

/* Edit Axis → the spec: fixed range ("fixed" / "fixedmin" / "fixedmax"), tick spacing, include zero */
function applySpace(spec, s, axis, noReverse = false) {
  const key = name => axis === "x" ? "x" + name : "value" + name;
  if (s.majorSpacing) spec[key("MajorUnit")] = s.majorSpacing;
  // e.g. a bump chart: rank 1 at the top (a butterfly's reversed wing is negated instead)
  if (s.reverse && !spec.mirrored && !noReverse) spec[axis === "x" ? "xReversed" : "valueReversed"] = true;
  if (/^fixed(min)?$/.test(s.rangeType || "") && s.min !== undefined) spec[key("Min")] = s.min;
  if (/^fixed(max)?$/.test(s.rangeType || "") && s.max !== undefined) spec[key("Max")] = s.max;
  if (s.domainExpand === "false" && axis !== "x") spec.includeZero = false;
}

/* an axis's tick format (Format → Axis → Numbers: the field's "label" text-format) */
const axisNumFmt = (fmt, ref) => (ref && tableauToExcelNumFmt(fmt.labelFormat(ref))) || undefined;

/* mark labels: the worksheet's label font, for bars the label alignment as Excel's label position, and a
 * label text around the value ("<AGG(Days)> DAYS") as literal text in the label's number format */
function applyLabels(spec, fmt, measures) {
  const dl = fmt.dataLabelStyle();
  if (dl.fontName || dl.fontSize || dl.color || dl.bold !== undefined) {
    spec.labelFont = { name: dl.fontName, size: dl.fontSize, color: dl.color ? dl.color.slice(2) : undefined, bold: dl.bold };
  }
  const tpl = fmt.labelTemplate();
  const at = tpl && (tpl.prefix.trim() || tpl.suffix.trim()) ? measures.findIndex(m => m && tfSameField(m, tpl.ref)) : -1;
  if (at === 0 || at === 1) {
    spec.series.forEach(s => {
      if (s.refLine || !!s.secondary !== (at === 1 && measures.length > 1) || s.labelTexts) return;
      s.labelNumFmt = tfWrapNumFmt(s.labelNumFmt || (s.secondary ? spec.secondaryNumFmt : spec.numFmt) || "General", tpl.prefix, tpl.suffix);
    });
  }
  if (spec.kind === "bar" || spec.kind === "combo") {
    const cell = fmt.markCellStyle(null);
    const pos = spec.barDir === "bar" ? { left: "inBase", center: "ctr", right: "inEnd" }[cell.hAlign]
                                     : { bottom: "inBase", middle: "ctr", top: "inEnd" }[cell.vAlign];
    if (pos) spec.labelPos = pos;
  }
}

/**
 * Tableau hides axes ("Show Header" off), renames or removes axis titles (Edit Axis), and turns grid
 * lines and axis rulers off per worksheet; the Excel chart follows the same settings.
 * @param {ChartSpec} spec @param {ChartContext} ctx
 */
function tvApplyWorkbookAxes(spec, ctx) {
  const { vm, roles } = ctx;
  const fmt = vm.fmt;
  if (!fmt.hasModel || roles.source !== "twb" || spec.axesHidden || /^(pie|doughnut|treemap)$/.test(spec.kind)) return;

  if (spec.kind === "scatter" || spec.kind === "bubble") {
    const x = roles.cols.values[0], y = roles.rows.values[0];
    const xi = x ? fmt.axisInfo(x.ref, "cols", "0") : {}, yi = y ? fmt.axisInfo(y.ref, "rows", "0") : {};
    if (xi.hidden) spec.xAxisHidden = true;
    if (xi.title !== undefined) spec.xTitle = xi.title;
    if (yi.hidden) spec.valueAxisHidden = true;
    if (yi.title !== undefined) spec.valueTitle = yi.title;
    if (!fmt.gridlinesShown("rows")) spec.gridlines = false;
    if (!fmt.gridlinesShown("cols")) spec.xGridlines = false;
    if (x) applySpace(spec, fmt.axisSpace(x.ref, "cols", "0"), "x");
    if (y) applySpace(spec, fmt.axisSpace(y.ref, "rows", "0"), "value");
    spec.xAxisNumFmt = axisNumFmt(fmt, x && x.ref);
    spec.valueAxisNumFmt = axisNumFmt(fmt, y && y.ref);
    applyLabels(spec, fmt, [y && y.ref, x && x.ref]);
    return;
  }

  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) return;
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  const refs = roles[valueShelf].axisRefs.length ? roles[valueShelf].axisRefs : roles[valueShelf].values.map(v => v.ref);
  const primary = refs[0] ? fmt.axisInfo(refs[0], valueShelf, "0") : {};
  // Tick Marks: None leaves an axis with no labels – Tableau shows nothing there
  const noTicks = (ref, cls) => fmt.axisSpace(ref, valueShelf, cls).majorShow === "false";
  if (primary.hidden || (refs[0] && noTicks(refs[0], "0"))) spec.valueAxisHidden = true;
  if (primary.title !== undefined) spec.valueTitle = primary.title;
  if (refs[1] && spec.series.some(s => s.secondary)) {
    // class = which axis of that field: a second measure has its own first axis ("0"); the same field on
    // both axes has its second one ("1")
    const cls = tfSameField(refs[0], refs[1]) ? "1" : "0";
    const second = fmt.axisInfo(refs[1], valueShelf, cls);
    if (second.hidden || noTicks(refs[1], cls)) spec.secondaryAxisHidden = true;
    if (second.title !== undefined) spec.secondaryTitle = second.title;
  }
  if (!fmt.gridlinesShown(valueShelf)) spec.gridlines = false;
  // one axis reversed beside another that is not (nor synchronized to it) mirrors two panes – a centred funnel, a
  // butterfly – which one Excel axis cannot: the reversal is left out
  const mirror = refs.length > 1 && refs.slice(1).some(r => {
    const s = fmt.axisSpace(r, valueShelf, tfSameField(r, refs[0]) ? "1" : "0");
    return !s.reverse && !s.synchronized;
  });
  if (refs[0]) applySpace(spec, fmt.axisSpace(refs[0], valueShelf, "0"), "value", mirror);
  spec.valueAxisNumFmt = axisNumFmt(fmt, refs[0]);
  if (refs[1]) spec.secondaryAxisNumFmt = axisNumFmt(fmt, refs[1]);
  applyLabels(spec, fmt, refs);
  // categories: a continuous pill draws an axis, discrete pills draw headers – hidden only if all are
  const dims = [...roles[valueShelf].dims.map(d => ({ d, shelf: valueShelf })), ...roles[catShelf].dims.map(d => ({ d, shelf: catShelf }))];
  const hidden = ({ d, shelf }) => !!d.ref && (d.continuous ? fmt.axisInfo(d.ref, shelf, "0").hidden === true : fmt.isLabelHidden(d.ref));
  if (dims.length && dims.every(hidden)) spec.categoryAxisHidden = true;
  // category labels: horizontal like Tableau's column headers, unless the header is rotated in the workbook
  const inner = roles[catShelf].dims.filter(d => d.ref && !d.continuous).slice(-1)[0];
  const rot = inner ? fmt.headerOrientation(inner.ref) : null;
  spec.categoryRotation = rot || 0;
  if (!fmt.axisLineShown(catShelf)) spec.axisLine = false;
}


/* ═══ data/values.js ══════════════════════════════════════════════════════════════════════════════ */
/* Tableau summary-data values: null checks, text/number access, natural + calendar ordering. */

/* ══════════════════════════════════════════════════════════════════════════
 * VIEW MODEL — rebuild the table Tableau actually draws from summary data
 *   1. pivot Measure Names / Measure Values into one column per measure
 *   2. merge stacked multi-pane rows (one block per pane → one row)
 *   3. sort rows like the view (manual sort, sort-by-measure, else natural)
 *   4. keep only visible columns, in visual order, with Tableau's header labels
 * ══════════════════════════════════════════════════════════════════════════ */
/** @param {DataValue} dv */
function tfIsNull(dv) {
  if (!dv) return true;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  return v === null || v === undefined || v === "%null%";
}

/** @param {DataValue} dv @returns {string} */
function tfDvText(dv) { return tfIsNull(dv) ? "" : String(dv.formattedValue != null ? dv.formattedValue : dv.value); }

/** @param {DataValue} dv @returns {number | null} */
function tfDvNum(dv) {
  if (tfIsNull(dv)) return null;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-eE]/g, ""));
  return isFinite(n) && /^[\s$€£¥(+-]*[\d.,]+/.test(String(v)) ? n : null;
}

/* discrete date parts arrive as names: "January" / "Jan" / "Monday" → calendar position */
const TF_MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

const TF_WEEKDAYS = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];

function tfCalendarRank(text) {
  const t = String(text || "").trim().toLowerCase();
  const m = t.match(/^(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sept?|oct|nov|dec)$/);
  if (m) return { kind: "month", rank: TF_MONTHS.indexOf(m[1].slice(0, 3)) };
  const d = t.match(/^(sunday|monday|tuesday|wednesday|thursday|friday|saturday|sun|mon|tue|wed|thu|fri|sat)$/);
  if (d) return { kind: "weekday", rank: TF_WEEKDAYS.indexOf(d[1].slice(0, 3)) };
  return null;
}

/** View order of two values: numbers, dates, ISO strings, month / weekday names, then text; nulls last.
 * @param {DataValue} a @param {DataValue} b @returns {number} */
function tfNaturalCompare(a, b) {
  const an = tfIsNull(a), bn = tfIsNull(b);
  if (an || bn) return an === bn ? 0 : an ? 1 : -1;           // nulls last
  const av = a.nativeValue !== undefined ? a.nativeValue : a.value;
  const bv = b.nativeValue !== undefined ? b.nativeValue : b.value;
  if (typeof av === "number" && typeof bv === "number") return av - bv;
  if (av instanceof Date && bv instanceof Date) return /** @type {any} */ (av) - /** @type {any} */ (bv);
  const iso = /^\d{4}-\d{2}-\d{2}/;
  if (typeof av === "string" && typeof bv === "string" && iso.test(av) && iso.test(bv)) return av < bv ? -1 : av > bv ? 1 : 0;
  const ra = tfCalendarRank(tfDvText(a)), rb = tfCalendarRank(tfDvText(b));   // MONTH / WEEKDAY names
  if (ra !== null && rb !== null && ra.kind === rb.kind) return ra.rank - rb.rank;
  return tfDvText(a).localeCompare(tfDvText(b), undefined, { numeric: true, sensitivity: "base" });
}


/* ═══ format/colors.js ════════════════════════════════════════════════════════════════════════════ */
/* Colour parsing and interpolation helpers. */

/* Tableau colours: "#rrggbb", "#rgb" or "#rrggbbaa" (aa=00 → transparent) → ARGB */
function tfArgb(v) {
  if (v == null) return undefined;
  let s = String(v).trim().replace(/^#/, "");
  if (/^[0-9a-f]{8}$/i.test(s)) {
    if (s.slice(6) === "00") return null;          // explicit "no fill"
    s = s.slice(0, 6);
  }
  if (/^[0-9a-f]{3}$/i.test(s)) s = s.split("").map(c => c + c).join("");
  if (!/^[0-9a-f]{6}$/i.test(s)) return undefined;
  return "FF" + s.toUpperCase();
}

function tfBrightness(argb) {
  const r = parseInt(argb.substring(2, 4), 16), g = parseInt(argb.substring(4, 6), 16), b = parseInt(argb.substring(6, 8), 16);
  return (r * 299 + g * 587 + b * 114) / 1000;
}

function tfInterpolate(c1, c2, t) {
  const p = (c, i) => parseInt(c.substring(i, i + 2), 16);
  const ch = i => Math.round(p(c1, i) + (p(c2, i) - p(c1, i)) * t).toString(16).padStart(2, "0");
  return ("FF" + ch(2) + ch(4) + ch(6)).toUpperCase();
}

function tfSample(colors, pos) {
  pos = Math.max(0, Math.min(1, pos));
  if (colors.length === 1) return colors[0];
  const seg = pos * (colors.length - 1);
  const i = Math.min(Math.floor(seg), colors.length - 2);
  return tfInterpolate(colors[i], colors[i + 1], seg - i);
}


/* ═══ format/palettes.js ══════════════════════════════════════════════════════════════════════════ */
/* Tableau built-in and automatic colour palettes. */

"use strict";

// ── TABLEAU BUILT-IN PALETTE LOOKUP ──────────────────────────────────────
const TABLEAU_BUILTIN_PALETTES = {
  "blue_10_0": ["#C7DDEA","#AFCFE1","#97C0D7","#7FAFCA","#689BC0","#5487B1","#4475A0","#356790","#2A5783"],
  "orange_10_0": ["#F3C184", "#F0AE62", "#EE9A42", "#EF882D", "#ED7420", "#E25F1D", "#CC531F", "#B54820", "#9E3D22"],
  "green_10_0": ["#B9D9AF", "#A3CF95", "#8BC57D", "#74BA67", "#5DAA56", "#4D984B", "#3D8743", "#31773F", "#24693D"],
  "red_10_0": ["#F3B8AB", "#EEA08E", "#EA8972", "#E9725B", "#EA5C4C", "#E6453C", "#D92C34", "#C71532", "#AE123A"],
  "purple_10_0": ["#ECC6E3", "#E2BCD8", "#D7AFCA", "#CB9FBC", "#BB85A8", "#AC759B", "#9F6B93", "#8E5B86", "#7C4D79"],
  "brown_10_0": ["#E8D5B4","#DEBE8A","#D8A66B","#D08F57","#C97A4B","#BF6740","#B45539","#AA4435","#9F3632"],
  "gray_10_0": ["#E5E5E5","#D4D6D8","#C0C4C8","#AAB0B6","#959DA5","#818A94","#6D7782","#5B6470","#49525E"],
  "gray_warm_10_0": ["#D8D1CE","#CBC2BE","#BBB0AB","#AB9F9A","#9A8E89","#887C77","#776B67","#685D59","#59504E"],
  "blue_teal_10_0": ["#B7D7D1","#9BC9C7","#80BBC0","#67ACC0","#529DBA","#448DAF","#3A7D9F","#336C91","#2C5985"],
  "orange_gold_10_0": ["#E8C85E","#EDB657","#F0A54A","#F08F32","#EF791F","#E96418","#D5531D","#BA4522","#9E3A26"],
  "green_gold_10_0": ["#E5C75A","#CDBE58","#B1B953","#97B64F","#7FB255","#67A957","#529B53","#348347","#146C36"],
  "red_gold_10_0": ["#E8C85A","#EDA951","#F08C4B","#EC7247","#E65E47","#DF4D47","#D33A45","#C32942","#B71D3E"],
  "orange_blue_diverging_10_0": ["#9E3D22","#B54820","#CC531F","#E25F1D","#ED7420","#F3D9BE","#97C0D7","#689BC0","#4475A0","#356790","#2B5C8A"],
  "red_green_diverging_10_0": ["#AE123A","#C61E3F","#D93443","#E94F4A","#F07A66","#F2B3A3","#8BC97D","#6DB65F","#539F50","#3B8447","#24693D"],
  "green_blue_diverging_10_0": ["#24693D","#347D46","#4E9854","#72B464","#9BCF89","#D4DDD9","#A9C9DC","#7DAACE","#5C8FBC","#4373A0","#2A5783"],
  "red_blue_diverging_10_0": ["#A90C38","#C71F3E","#DC393F","#EB5A4F","#F3A091","#E7E2DE","#B8D1E1","#86B0D1","#5F90BC","#44709C","#2E5A87"],
  "red_black_10_0": ["#AE123A","#C71F3E","#DC393F","#EB5A4F","#F3A091","#DDD9D5","#B9BCBC","#969DA1","#78818A","#606A75","#49525E"],
  "gold_purple_diverging_10_0": ["#AD9024","#B89B34","#C6AA50","#D3BA6D","#DDC892","#E3D7D1","#D7C1D2","#C9A5C3","#BB8AB2","#AC7299"],
  "red_green_gold_diverging_10_0": ["#BE2A3E","#D44344","#E75D49","#F07A47","#F2A14A","#E7C65A","#A8BE5E","#77AF5B","#55994E","#3B8746","#22763F"],
  "sunrise_sunset_diverging_10_0": ["#33608C","#556AA0","#7B67A6","#A664A2","#CD6C95","#EC7C79","#F2B15A","#ED8D46","#E56B44","#D24844","#B81840"],
  "orange_blue_white_diverging_10_0": ["#9E3D22","#B94B20","#D45A1D","#EC7420","#F3B562","#F5F1EC","#C6DDEA","#93BED8","#679BC1","#4677A5","#2B5C8A"],
  "red_green_white_diverging_10_0": ["#AE123A","#C8243F","#DC4947","#ED725E","#F3A38F","#F3F1EE","#B9DFAF","#8BC57D","#5DAA56","#3F8847","#24693D"],
  "green_blue_white_diverging_10_0": ["#24693D","#3E864B","#5BA557","#82C06F","#B7DFAE","#F2F3F1","#C7DDEA","#97C0D7","#689BC0","#4475A0","#2A5783"],
  "red_blue_white_diverging_10_0": ["#A90C38","#C71F3E","#DC393F","#EB5A4F","#F3A091","#F4F3F2","#C6DDEA","#93BED8","#679BC1","#4677A5","#2E5A87"],
  "red_black_white_diverging_10_0": ["#AE123A","#C8243F","#DC4947","#ED725E","#F3A38F","#F3F2F1","#D0D2D1","#A7ADB0","#838C93","#626C77","#49525E"],
  "tableau-blue-light": ["#EEF2F7","#E7EDF5","#DFE8F3","#D8E2F0","#D1DDEE","#CBD9ED","#C9D7F1","#C7D9F2","#C4D8F3"],
  "tableau-orange-light": ["#F6F2EE","#F7EBDD","#F8E3CC","#F9DBBE","#FAD5B3","#FBCFA8","#FCCB9F","#FFCC9E","#FFCC9E"],
  "tableau-orange-blue-light": ["#FFCC9E","#FAD1AB","#F4D8BC","#EEE0CC","#EAE4D8","#E8E8E8","#DFE7EF","#D5E0EC","#CBD9E9","#C7D9F1","#C4D8F3"],
  "tableau-map-blue-green": ["#F5F5C8","#EEF2B3","#E0EA9A","#CBE18F","#AFD695","#91CC9D","#72C3A8","#58BCB5","#41B7C4"],
  "tableau-map-temperatur": ["#529985","#669C76","#81A364","#A6B04E","#D2C63F","#F0D347","#F2C04A","#E7A24A","#D4824D","#C26B51"]
};

// ── HELPER: Get built-in palette colors by name ──────────────────────────
function getBuiltInPaletteColors(paletteName) {
  if (!paletteName) return null;
  
  if (TABLEAU_BUILTIN_PALETTES[paletteName]) {
    return TABLEAU_BUILTIN_PALETTES[paletteName];
  }
  
  const lowerName = paletteName.toLowerCase();
  for (const [key, colors] of Object.entries(TABLEAU_BUILTIN_PALETTES)) {
    if (key.toLowerCase() === lowerName) {
      return colors;
    }
  }
  
  const normalized = paletteName.toLowerCase().replace(/[-\s]+/g, '');
  for (const [key, colors] of Object.entries(TABLEAU_BUILTIN_PALETTES)) {
    const keyNormalized = key.toLowerCase().replace(/[-\s]+/g, '');
    if (keyNormalized.includes(normalized) || normalized.includes(keyNormalized)) {
      return colors;
    }
  }
  
  return null;
}

const AUTOMATIC_PALETTE_BY_MARK = {
  automatic:    "blue_teal_10_0",
  bar:          "blue_10_0",
  line:         "blue_10_0",
  area:         "blue_teal_10_0",
  square:       "blue_teal_10_0",
  circle:       "blue_10_0",
  shape:        "blue_10_0",
  text:         "blue_10_0",
  map:          "blue_teal_10_0",
  multipolygon: "blue_teal_10_0",
  pie:          "blue_10_0",
  ganttbar:     "blue_10_0",
  polygon:      "blue_10_0",
  density:      "blue_10_0",
  heatmap:      "blue_10_0"
};

const DEFAULT_AUTOMATIC_PALETTE = "blue_10_0";

function getAutomaticPaletteForMark(markClass) {
  const key = String(markClass || "Automatic").toLowerCase().replace(/[\s_-]+/g, "");
  const name = AUTOMATIC_PALETTE_BY_MARK[key] || DEFAULT_AUTOMATIC_PALETTE;
  if (!AUTOMATIC_PALETTE_BY_MARK[key]) {
    console.log(`[Auto Palette] Mark type "${markClass}" not in table, using ${name}`);
  }
  return {
    name,
    colors: TABLEAU_BUILTIN_PALETTES[name].map(c => "FF" + c.replace("#", "").toUpperCase())
  };
}

const TABLEAU_10 = ["FF4E79A7", "FFF28E2B", "FFE15759", "FF76B7B2", "FF59A14F",
                    "FFEDC948", "FFB07AA1", "FFFF9DA7", "FF9C755F", "FFBAB0AC"];


/* ═══ format/color-scale.js ═══════════════════════════════════════════════════════════════════════ */
/* Colour encodings → value-to-colour scales (categorical and continuous). */

/* ── Colour scales ─────────────────────────────────────────────────────── */
function tfBuildColorScale(fmt, enc, values) {
  const def = enc.def;
  if (enc.continuous) {
    const nums = values.filter(v => typeof v === "number" && isFinite(v));
    if (!nums.length) return null;
    const min = def && def.min !== undefined ? def.min : Math.min(...nums);
    const max = def && def.max !== undefined ? def.max : Math.max(...nums);
    let colors = def && def.customColors.length ? def.customColors : null;
    let diverging = /diverging/i.test((def && (def.paletteName || "")) + " " + (def && (def.paletteType || "")));
    if (!colors && def && def.paletteName && !/^automatic$/i.test(def.paletteName)) {
      const pc = getBuiltInPaletteColors(def.paletteName);
      if (pc) colors = pc.map(tfArgb);
    }
    if (!colors) {                     // Tableau "Automatic"
      if (min < 0 && max > 0) { colors = getBuiltInPaletteColors("orange_blue_diverging_10_0").map(tfArgb); diverging = true; }
      else colors = getAutomaticPaletteForMark(enc.markClass).colors;
    }
    if (def && def.reverse) colors = colors.slice().reverse();
    let center = def && def.center !== undefined ? def.center : null;
    if (diverging && center === null) center = (min < 0 && max > 0) ? 0 : (min + max) / 2;
    tfLog(`Gradient for ${enc.ref.inner}: ${colors.length} stops, range ${min}..${max}${center !== null ? ", center " + center : ""}`);
    return v => {
      if (typeof v !== "number" || !isFinite(v)) return null;
      let pos;
      if (center !== null) {
        const dev = Math.max(Math.abs(max - center), Math.abs(min - center)) || 1;
        pos = 0.5 + (v - center) / (2 * dev);       // symmetric, like "Use full colour range" = off
      } else pos = max === min ? 0.5 : (v - min) / (max - min);
      return tfSample(colors, pos);
    };
  }
  // categorical
  const map = {};
  if (def) for (const [k, c] of Object.entries(def.map)) fmt.bucketAliases(k).forEach(a => { map[tfNorm(a)] = c; });
  const used = new Set(Object.values(map));
  const pool = TABLEAU_10.filter(c => !used.has(c));
  const auto = {};
  [...new Set(values.filter(v => v != null && v !== "").map(tfNorm))].filter(k => !(k in map)).sort().forEach((k, i) => {
    const p = pool.length ? pool : TABLEAU_10;
    auto[k] = p[i % p.length];
  });
  tfLog(`Categorical colours for ${enc.ref.inner}: ${Object.keys(map).length} from TWB, ${Object.keys(auto).length} auto-assigned`);
  return v => (v == null || v === "") ? null : (map[tfNorm(v)] || auto[tfNorm(v)] || null);
}


/* ═══ format/date-format.js ═══════════════════════════════════════════════════════════════════════ */
/* Tableau date label formats (Format → Header → Dates) applied to header values. */

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/**
 * A header value in Tableau's custom date format: "iLLLLL" → "J" (month initial), "MMM" → "Jan",
 * "MMM yy" → "Jan 24", "yyyy" → "2024" … The value can be a full date or a date part Tableau sends as
 * text (month / weekday name, "Q1", year). null when the format needs a part the value does not have.
 * @param {string} raw Tableau text-format ("i…" / "*…" prefix + ICU-style pattern)
 * @param {string} text the formatted value from the summary data
 * @param {DataValue} [dv]
 * @returns {string | null}
 */
function tfFormatDateLabel(raw, text, dv) {
  const pattern = String(raw || "").replace(/^[i*]/, "");
  if (!/[yMLdEQq]/.test(pattern)) return null;
  const t = String(text || "").trim().toLowerCase();
  let date = null, month = null, weekday = null, quarter = null, year = null;
  const native = dv ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : null;
  if (native instanceof Date) date = native;
  else if (typeof native === "string" && /^\d{4}-\d{2}-\d{2}/.test(native)) date = new Date(native.slice(0, 10) + "T00:00:00Z");
  if (date && !isNaN(date.getTime())) {
    year = date.getUTCFullYear(); month = date.getUTCMonth(); weekday = date.getUTCDay(); quarter = Math.floor(month / 3) + 1;
  } else {
    date = null;
    const m = MONTHS.findIndex(n => n.toLowerCase() === t || n.slice(0, 3).toLowerCase() === t);
    if (m >= 0) { month = m; quarter = Math.floor(m / 3) + 1; }
    const d = DAYS.findIndex(n => n.toLowerCase() === t || n.slice(0, 3).toLowerCase() === t);
    if (d >= 0) weekday = d;
    const q = t.match(/^q([1-4])$/);
    if (q) quarter = +q[1];
    if (/^\d{4}$/.test(t)) year = +t;
  }
  if (month === null && weekday === null && quarter === null && year === null) return null;
  let missing = false;
  const out = pattern.replace(/'([^']*)'|y{1,4}|M{1,5}|L{1,5}|d{1,2}|E{1,5}|Q{1,4}|q{1,4}/g, (tok, literal) => {
    if (literal !== undefined) return literal;
    const c = tok[0], n = tok.length;
    const need = v => { if (v === null) missing = true; return v === null; };
    if (c === "y") return need(year) ? "" : n === 2 ? String(year).slice(-2) : String(year);
    if (c === "M" || c === "L") {
      if (need(month)) return "";
      return n >= 5 ? MONTHS[month][0] : n === 4 ? MONTHS[month] : n === 3 ? MONTHS[month].slice(0, 3)
           : n === 2 ? String(month + 1).padStart(2, "0") : String(month + 1);
    }
    if (c === "d") return need(date) ? "" : String(date.getUTCDate()).padStart(n, "0");
    if (c === "E") return need(weekday) ? "" : n >= 5 ? DAYS[weekday][0] : n === 4 ? DAYS[weekday] : DAYS[weekday].slice(0, 3);
    return need(quarter) ? "" : n >= 3 ? "Q" + quarter : String(quarter);
  });
  return missing ? null : out;
}


/* ═══ twb/dashboard-text.js ═══════════════════════════════════════════════════════════════════════ */
/* Dashboard text boxes: header/title runs and zone borders. */

/* Text boxes count as column headers only if ALL hold:
 *   – ≥ 2 text zones, visible, in one row (same top edge ± tol)
 *   – their bottom edge touches the sheet's top edge (± tol)
 *   – they lie within the sheet's horizontal span
 *   – their number equals the number of visible columns
 * Positions are Tableau's own zone coordinates (0–100000), so the check is deterministic. */
const TF_ZONE_TOL = 1500;

function tfZoneText(z) { return (z.runs || []).map(r => r.text).join("").replace(/\u00C6\r?\n?/g, "\n").trim(); }

function tfTextBoxHeaders(dash, sheetName, nCols) {
  const strip = tfTextBoxHeaderStrip(dash, sheetName);
  if (!strip || nCols < 2 || strip.zones.length !== nCols) return null;
  return strip;
}

/* the row of text boxes sitting on a worksheet's top edge, left to right (≥ 2), whatever their number */
function tfTextBoxHeaderStrip(dash, sheetName) {
  const ws = dash.zones.find(z => z.name === sheetName && z.type === "worksheet");
  if (!ws) return null;
  const near = (a, b) => Math.abs(a - b) <= TF_ZONE_TOL;
  const cand = dash.zones.filter(z => z.type === "text" && !z.hidden && z.runs && tfZoneText(z) &&
    near(z.y + z.h, ws.y) && z.x >= ws.x - TF_ZONE_TOL && z.x + z.w <= ws.x + ws.w + TF_ZONE_TOL);
  if (cand.length < 2 || !cand.every(z => near(z.y, cand[0].y))) return null;
  cand.sort((a, b) => a.x - b.x);
  return { ws, zones: cand.map(z => ({ id: z.id, text: tfZoneText(z), x: z.x, w: z.w, h: z.h, runs: z.runs,
    props: (z.runs.find(r => r.text.trim()) || { props: {} }).props })) };
}

/* Dashboard title: Tableau's own title if shown, else the text box above every worksheet
 * with the largest font (ties → topmost). Returns runs for rich text, or null. */
function tfDashboardTitleRuns(model, dashName, usedZoneIds) {
  const dash = model && model.dashboards && model.dashboards[dashName];
  if (!dash) return null;
  const titleZone = dash.zones.find(z => z.type === "title" && !z.hidden);
  if (titleZone && dash.title && dash.title.runs.length) return dash.title.runs;
  if (!FORMAT_CONFIG.textBoxTitle) return null;
  const wsTop = Math.min(...dash.zones.filter(z => z.type === "worksheet").map(z => z.y));
  const size = z => (z.runs.find(r => r.text.trim()) || { props: {} }).props.fontSize || 9;
  const cand = dash.zones.filter(z => z.type === "text" && !z.hidden && z.runs && tfZoneText(z) &&
    !usedZoneIds.has(z.id) && z.y + z.h <= wsTop + TF_ZONE_TOL);
  if (!cand.length) return null;
  cand.sort((a, b) => size(b) - size(a) || a.y - b.y);
  return cand[0].runs;
}

function tfStrokeToBorder(size) {
  if (!size) return null;
  if (size <= 1) return "thin";
  if (size === 2) return "medium";
  return "thick";
}


/* ═══ twb/formatter.js ════════════════════════════════════════════════════════════════════════════ */
/* Resolves Tableau's formatting cascade for one worksheet (createSheetFormatter). */

function tfFormatsToProps(m) {
  const p = {};
  if (m["font-family"]) p.fontName = m["font-family"];
  if (m["font-size"]) p.fontSize = tfNum(m["font-size"]);
  if (m["font-weight"]) p.bold = /bold/i.test(m["font-weight"]);
  if (m["font-style"]) p.italic = /italic/i.test(m["font-style"]);
  if (m["text-decoration"]) { p.underline = /underline/i.test(m["text-decoration"]); p.strike = /line-through/i.test(m["text-decoration"]); }
  if (m["color"]) p.color = tfArgb(m["color"]) || undefined;
  if ("background-color" in m) p.bgColor = tfArgb(m["background-color"]);  // null = transparent
  if (/^(left|center|right)$/.test(m["text-align"] || "")) p.hAlign = m["text-align"];
  const va = { top: "top", center: "middle", middle: "middle", bottom: "bottom" }[m["vertical-align"]];
  if (va) p.vAlign = va;
  if (m["wrap"]) p.wrap = /^(true|on)$/i.test(m["wrap"]);
  if (m["text-format"]) p.numFmtRaw = m["text-format"];
  if (m["band-color"]) p.bandColor = tfArgb(m["band-color"]) || undefined;
  if (m["band-size"]) p.bandSize = tfNum(m["band-size"]);
  if (m["line-visibility"]) p.lineVisible = m["line-visibility"] !== "off";
  if (m["stroke-size"] != null) p.strokeSize = tfNum(m["stroke-size"]);
  if (m["stroke-color"]) p.strokeColor = tfArgb(m["stroke-color"]) || undefined;
  if (m["div-level"]) p.divLevel = tfNum(m["div-level"]);
  if (m["width"]) p.width = tfNum(m["width"]);
  if (m["display"]) p.display = m["display"] !== "false";
  if (m["display-field-labels"]) p.displayFieldLabels = m["display-field-labels"] !== "false";
  if (m["height"]) p.height = tfNum(m["height"]);
  if (m["text-orientation"] != null) p.textOrientation = tfNum(m["text-orientation"]);
  return p;
}

/* precedence inside one element: base < scope < field < field+scope */
function tfCollect(style, elements, opts = {}) {
  const out = {};
  if (!style) return out;
  for (const el of elements) {
    const tiers = [{}, {}, {}, {}];
    for (const f of style.rules[el] || []) {
      if ((f.dataClass || null) !== (opts.dataClass || null)) continue;
      if (f.scope && f.scope !== opts.scope) continue;
      if (f.field && !(opts.field && tfSameField(tfParseFieldRef(f.field), opts.field))) continue;
      tiers[(f.field ? 2 : 0) + (f.scope ? 1 : 0)][f.attr] = f.value;
    }
    Object.assign(out, ...tiers);
  }
  return tfFormatsToProps(out);
}

function tfMerge(...layers) {
  const out = {};
  layers.forEach(l => { if (l) for (const k in l) if (l[k] !== undefined) out[k] = l[k]; });
  return out;
}

function tfFieldInfo(model, ref) {
  if (!model || !ref) return null;
  const k = (tfTableCalcBase(ref) || ref).name.toLowerCase();      // a table calculation: its measure's field
  return model.fields[(ref.ds || "") + "|" + k] || model.fields["|" + k] || null;
}

function tfDisplayNames(model, ref) {
  const info = tfFieldInfo(model, ref);
  const cap = (info && info.caption) || ref.name;
  const agg = TF_DERIV_LABEL[(ref.deriv || "").toLowerCase()];
  const names = [];
  if (agg) names.push(`${agg}(${cap})`, `${agg}(${ref.name})`);
  names.push(cap, ref.name);
  if (ref.name === "Multiple Values") names.push("Measure Values");
  const alias = model && model.measureAliases && (model.measureAliases[tfRefKey(ref)] ||
    Object.entries(model.measureAliases).find(([k]) => k.endsWith("|" + ref.inner.toLowerCase()))?.[1]);
  if (alias) names.unshift(alias);
  return [...new Set(names.map(tfNorm))];
}

/** Format → Dashboard → Dashboard Shading: the dashboard's canvas colour (ARGB), or null
 * @param {FormatModel | null} model @param {string} dashboardName @returns {string | null} */
function tfDashboardShading(model, dashboardName) {
  const dash = model && model.dashboards ? model.dashboards[dashboardName] : null;
  return (dash && tfCollect(dash.style, ["table"]).bgColor) || null;
}

/** @param {FormatModel | null} model @param {string} sheetName */
/**
 * @param {FormatModel | null} model @param {string} sheetName
 * @param {Pane[]} [onlyPanes] scope the Marks-card methods to these panes (one layer of a dual-axis donut)
 */
function createSheetFormatter(model, sheetName, onlyPanes) {
  const sheet = (model && model.sheets && model.sheets[sheetName]) || null;
  const wb = model ? model.workbookStyle : null;
  const st = sheet ? sheet.style : null;
  const panes = onlyPanes || (sheet ? sheet.panes : []);
  /** @param {Pane} p @param {string} attr */
  const markRule = (p, attr) => (((p.style && p.style.rules.mark) || []).find(x => x.attr === attr) || {}).value;
  const matchCache = new Map();

  function base() {
    return tfMerge(TABLEAU_DEFAULTS.worksheet, tfCollect(wb, TF_ELEMENTS.sheet), tfCollect(st, TF_ELEMENTS.sheet));
  }

  /* summary-data column name ("SUM(Sales)") → Tableau field ref. Exact names only,
   * no substring matching (old code matched "Sales" to "Sales Person"). */
  function matchName(colName) {
    const key = tfNorm(colName);
    if (matchCache.has(key)) return matchCache.get(key);
    let best = null, bestScore = 0;
    const stripped = key.replace(/^[a-z]+\((.*)\)$/, "$1");
    for (const ref of sheet ? sheet.fieldRefs : []) {
      const names = tfDisplayNames(model, ref);
      let score = 0;
      if (names[0] === key) score = 3;
      else if (names.includes(key)) score = 2;
      else if (names.includes(stripped)) score = 1;
      if (score > bestScore) { best = ref; bestScore = score; }
    }
    matchCache.set(key, best);
    return best;
  }

  /* summary-data column id ("[federated.x].[pcto:sum:Sales:qk:2]") → the sheet's field ref; null when the id is not
   * a field instance of this sheet */
  function matchFieldId(id) {
    const r = id && /^\[/.test(String(id)) ? tfParseFieldRef(id) : null;
    if (!r || !r.type || !sheet) return null;
    return sheet.fieldRefs.find(x => x.inner.toLowerCase() === r.inner.toLowerCase()) || null;
  }

  function onShelf(ref, shelf) { return !!ref && (sheet ? sheet[shelf] : []).some(r => tfSameField(r, ref)); }
  // headers = DISCRETE pills on Rows/Columns; continuous pills (…:qk) draw axes, not headers
  function isHeaderField(ref) { return !!ref && ref.type !== "qk" && (onShelf(ref, "rows") || onShelf(ref, "cols")); }
  function shelfOf(ref) { return onShelf(ref, "rows") ? "rows" : onShelf(ref, "cols") ? "cols" : null; }

  function labelRunProps(ref) {
    if (!ref) return {};
    for (const p of panes) for (const r of p.labelRuns) if (r.refs.some(x => tfSameField(x, ref))) return r.props;
    return {};
  }
  /* panes that draw this field (multi-pane tables: each column has its own marks card) */
  function owningPanes(ref) {
    const owns = p => ref && (p.labelRuns.some(r => r.refs.some(x => tfSameField(x, ref))) ||
                              p.encodings.some(e => tfSameField(e.field, ref)));
    const hit = panes.filter(owns);
    if (hit.length) return hit;
    if (panes.length === 1) return panes;
    return panes.filter(p => !p.id);                            // the "All" pane
  }
  function paneStyle(ref) {
    const els = TF_ELEMENTS.pane.concat(TF_ELEMENTS.cell, TF_ELEMENTS.datalabel);
    return tfMerge(...owningPanes(ref).map(p => tfCollect(p.style, els, { field: ref })));
  }

  function numFmt(ref, fromHeader) {
    const els = fromHeader ? TF_ELEMENTS.header : TF_ELEMENTS.pane.concat(TF_ELEMENTS.cell);
    const raw = tfCollect(st, els, { field: ref, scope: fromHeader ? "rows" : undefined }).numFmtRaw
             || (ref && (tfFieldInfo(model, ref) || {}).defaultFormat);
    return raw ? tableauToExcelNumFmt(raw) : null;
  }

  return {
    hasModel: !!sheet,
    matchName,
    matchFieldId,
    isHeaderField,

    titleText() { return sheet && sheet.title ? sheet.title.text : null; },
    titleStyle() {
      const runs = sheet && sheet.title ? sheet.title.runs : [];
      const firstRun = (runs.find(r => r.text.trim()) || {}).props;
      return tfMerge(base(), TABLEAU_DEFAULTS.title, tfCollect(wb, TF_ELEMENTS.title), tfCollect(st, TF_ELEMENTS.title), firstRun);
    },

    /* Excel header row (field names) */
    fieldLabelStyle(ref, isMeasure) {
      return tfMerge(base(), TABLEAU_DEFAULTS.fieldLabel,
        tfCollect(wb, TF_ELEMENTS.fieldLabel, { scope: "rows" }),
        tfCollect(st, TF_ELEMENTS.fieldLabel, { field: ref, scope: isMeasure ? "cols" : "rows" }));
    },

    /* dimension values (row headers in Tableau) */
    headerCellStyle(ref) {
      const p = tfMerge(base(), tfCollect(wb, TF_ELEMENTS.header, { scope: "rows" }),
        tfCollect(st, TF_ELEMENTS.header, { field: ref, scope: "rows" }));
      p.numFmt = numFmt(ref, true);
      return p;
    },

    /** a header's text orientation in degrees (-90 = rotated up), null when Tableau's default (horizontal)
     * @param {FieldRef} ref */
    headerOrientation(ref) {
      if (!ref) return null;
      const p = tfMerge(tfCollect(wb, ["label"], { field: ref }), tfCollect(st, ["label"], { field: ref }),
        tfCollect(st, ["label"], { field: ref, scope: "cols" }), tfCollect(st, ["label"], { field: ref, scope: "rows" }));
      return typeof p.textOrientation === "number" ? p.textOrientation : null;
    },

    /* measure values (text marks in Tableau) */
    markCellStyle(ref) {
      const els = TF_ELEMENTS.pane.concat(TF_ELEMENTS.cell, TF_ELEMENTS.datalabel);
      const p = tfMerge(base(), tfCollect(wb, els),
        tfCollect(st, els, { field: ref, scope: "rows" }),
        paneStyle(ref), labelRunProps(ref));
      // cell position comes from Format → Alignment (pane/cell); the label editor's alignment only
      // aligns lines inside the label, so it is used only when no pane/cell alignment is set
      const cellAlign = tfMerge(tfCollect(st, els, { field: ref, scope: "rows" }), paneStyle(ref)).hAlign;
      if (cellAlign) p.hAlign = cellAlign;
      p.numFmt = numFmt(ref, false);
      p.explicitColor = !!(labelRunProps(ref).color || tfCollect(st, TF_ELEMENTS.cell, { field: ref }).color);
      return p;
    },

    divider(scope) {
      const p = tfCollect(st, TF_ELEMENTS.divider, { scope });
      const d = scope === "rows" ? { ...TABLEAU_DEFAULTS.divider } : { ...TABLEAU_DEFAULTS.divider, visible: false };
      if (p.lineVisible !== undefined) d.visible = p.lineVisible;
      if (p.strokeSize !== undefined) { d.style = tfStrokeToBorder(p.strokeSize); if (!d.style) d.visible = false; }
      if (p.strokeColor) d.color = p.strokeColor;
      if (p.divLevel) d.level = p.divLevel;
      return d;
    },

    banding() {
      const pane = tfCollect(st, TF_ELEMENTS.pane, { scope: "rows" });
      const header = tfCollect(st, TF_ELEMENTS.header, { scope: "rows" });
      const table = tfCollect(st, ["table"], { scope: "rows" });
      return { pane: pane.bandColor || null, header: header.bandColor || null,
               size: table.bandSize || pane.bandSize || header.bandSize || 1 };
    },

    widthPx(ref) {
      return tfCollect(st, TF_ELEMENTS.cell.concat(TF_ELEMENTS.header), { field: ref, scope: "cols" }).width
          || tfCollect(st, TF_ELEMENTS.cell.concat(TF_ELEMENTS.header), { field: ref }).width || null;
    },
    /* row height: Format → Cell height, which Tableau stores on the innermost Rows dimension when set by
       dragging the row border */
    rowHeightPx() {
      const own = tfCollect(st, TF_ELEMENTS.cell).height;
      if (own) return own;
      const dims = sheet ? sheet.rows.filter(r => r.type !== "qk") : [];
      for (let i = dims.length - 1; i >= 0; i--) {
        const h = tfCollect(st, TF_ELEMENTS.cell, { field: dims[i] }).height;
        if (h) return h;
      }
      return null;
    },

    /* fields on Text / Label – these are the "marks" that get coloured */
    textRefs() {
      const refs = [];
      panes.forEach(p => {
        p.encodings.filter(e => e.channel === "text" || e.channel === "label").forEach(e => refs.push(e.field));
        p.labelRuns.forEach(r => r.refs.forEach(x => refs.push(x)));
      });
      return refs;
    },

    colorEncoding() {
      for (const p of panes) {
        const out = this.colorEncodingOf(p);
        if (out) return out;
      }
      return null;
    },

    /** a Shape encoding's value → shape map ("Arrows/1-4.png", ":filled/circle"): sheet, then data source
     * @param {FieldRef} ref @returns {Record<string, string> | null} */
    shapeMap(ref) {
      const pool = [st, ...(model ? Object.values(model.datasourceStyles) : []), wb].filter(Boolean);
      for (const s of pool) {
        const def = (s.shapes || []).find(x => tfSameField(x.field, ref)) ||
                    (s.shapes || []).find(x => x.field && tfNorm(x.field.name) === tfNorm(ref.name));
        if (def) return def.map;
      }
      return null;
    },

    /** one pane's Color encoding (a sheet built from marks has one per pane) @param {any} p a pane */
    colorEncodingOf(p) {
      const enc = p.encodings.find(e => e.channel === "color" && e.field);
      if (!enc) return null;
      const ref = enc.field;
      const hasText = p.encodings.some(e => e.channel === "text" || e.channel === "label") || p.labelRuns.length > 0;
      const cls = String(p.markClass || "Automatic");
      const effective = /^automatic$/i.test(cls) && hasText ? "Text" : cls;
      const pool = [st, ...panes.map(x => x.style), model && model.datasourceStyles[ref.ds],
                    ...(model ? Object.values(model.datasourceStyles) : [])].filter(Boolean);
      let def = null;
      for (const s of pool) { def = s.encodings.find(e => tfSameField(e.field, ref)); if (def) break; }
      if (!def) for (const s of pool) { def = s.encodings.find(e => e.field && tfNorm(e.field.name) === tfNorm(ref.name)); if (def) break; }
      const paneRefs = [...p.encodings.filter(e => e.channel === "text" || e.channel === "label").map(e => e.field),
                        ...p.labelRuns.flatMap(r => r.refs)];
      return { ref, def, markClass: effective, applyTo: /^text$/i.test(effective) ? "font" : "fill", paneRefs,
               continuous: ref.type ? /^q/.test(ref.type) : !!(def && def.type === "interpolated") };
    },

    shelfOf,
    sheetModel: sheet,

    /** Format → Shading → Worksheet, workbook < worksheet: ARGB; "none" = switched off (what is behind the
     * sheet shows); undefined = not set (Tableau draws the sheet white). A dashboard's shading colours the
     * dashboard only – sheets on it keep their own (designers set them to match).
     * @returns {string | undefined} */
    sheetShading() {
      let out;
      [wb, st].forEach(s => { const x = s ? tfCollect(s, ["table"]).bgColor : undefined; if (x !== undefined) out = x === null ? "none" : x; });
      return out;
    },
    /** the worksheet's own background colour (ARGB), null when it has none (white or switched off) */
    tableBackground() {
      const s = this.sheetShading();
      return s && s !== "none" ? s : null;
    },
    /* column-header height (px) stored on any Columns-shelf header field, incl. Measure Names */
    headerRowHeightPx() {
      if (!sheet) return null;
      let h = null;
      (st && st.rules.header || []).forEach(f => {
        if (f.attr !== "height" || !f.field) return;
        const r = tfParseFieldRef(f.field);
        if (r.name === "Measure Names" || onShelf(r, "cols")) h = Math.max(h || 0, tfNum(f.value) || 0);
      });
      return h;
    },
    /* URL actions whose source includes this sheet */
    urlActions(dashboardName) {
      return (model && model.actions || []).filter(a =>
        (a.worksheet ? a.worksheet === sheetName : (!a.dashboard || a.dashboard === dashboardName)) &&
        !(a.exclude || []).includes(sheetName));
    },

    /* header label hidden via Format → "Hide" (label display=false) */
    isLabelHidden(ref) { return tfCollect(st, ["label"], { field: ref }).display === false; },

    /* an axis on a shelf: "Show Header" off (display=false) and Edit Axis → title ("" = no title).
     * cls "0" / "1" = primary / secondary axis of a dual axis; rules without a class apply to both */
    axisInfo(ref, shelf, cls) {
      /** @type {{ hidden?: boolean, title?: string }} */
      const out = {};
      if (!st || !ref) return out;
      /** @type {Record<string, string>[]} */
      const tiers = [{}, {}];                                 // without scope < with scope
      for (const f of st.rules.axis || []) {
        if (f.attr !== "display" && f.attr !== "title") continue;
        if (!f.field || !tfSameField(tfParseFieldRef(f.field), ref)) continue;
        if (f.scope && f.scope !== shelf) continue;
        if (f.axisClass !== undefined && cls !== undefined && f.axisClass !== cls) continue;
        tiers[f.scope ? 1 : 0][f.attr] = f.value;
      }
      const v = { ...tiers[0], ...tiers[1] };
      if (v.display !== undefined) out.hidden = v.display === "false";
      if (v.title !== undefined) out.title = v.title;
      return out;
    },
    /* grid lines across a shelf's axis (Format → Lines → Grid Lines): off = stroke 0 or line-visibility off */
    gridlinesShown(shelf) {
      const g = tfMerge(tfCollect(wb, ["gridline"], { scope: shelf }), tfCollect(st, ["gridline"], { scope: shelf }));
      return !(g.strokeSize === 0 || g.lineVisible === false);
    },
    /* Analytics → Reference Line entries of the worksheet */
    referenceLines() { return (sheet && sheet.referenceLines) || []; },
    /* Edit Axis for a field on a shelf: fixed range, tick spacing, include zero */
    axisSpace(ref, shelf, cls) {
      /** @type {AxisSpace} */
      let out = {};
      if (!st || !st.spaces || !ref) return out;
      for (const s of st.spaces) {
        if (!s.field || !tfSameField(tfParseFieldRef(s.field), ref) || (s.scope && s.scope !== shelf)) continue;
        if (s.axisClass !== undefined && cls !== undefined && s.axisClass !== cls) continue;
        out = s;
      }
      return out;
    },
    /* a header's label format (Format → Header → Dates), e.g. "iLLLLL" = month initial */
    labelFormat(ref) { return ref ? tfCollect(st, ["label"], { field: ref }).numFmtRaw || null : null; },
    /** mark label font (datalabel element); its colour only when the user picked one, not "automatic"
     * @returns {Record<string, any>} */
    dataLabelStyle() {
      const raw = {};
      const pool = [wb, st, ...owningPanes(null).map(p => p.style)];      // workbook, worksheet, marks card
      for (const s of pool) for (const f of (s && s.rules.datalabel) || []) if (!f.field) raw[f.attr] = f.value;
      const p = tfMerge(...pool.map(s => tfCollect(s, ["datalabel"])));
      if (raw["color-mode"] && raw["color-mode"] !== "user") delete p.color;
      return p;
    },
    /** the mark label's text around its one field ("<AGG(Days)> DAYS" → prefix "", suffix " DAYS"); null when
     * the label is the default, or holds several fields
     * @returns {{ ref: FieldRef, prefix: string, suffix: string } | null} */
    labelTemplate() {
      const pane = owningPanes(null)[0];
      if (!pane || !pane.labelRuns.length) return null;
      const text = pane.labelRuns.map(r => r.text).join("");
      const tokens = [...text.matchAll(/<([^<>]+)>/g)];
      const ref = tokens.length === 1 ? tfParseFieldRef(tokens[0][1]) : null;
      if (!ref) return null;
      const t = tokens[0], flat = s => s.replace(/\u00C6[ \t]*(?:\r?\n)?|\r?\n/g, " ");
      return { ref, prefix: flat(text.slice(0, t.index)), suffix: flat(text.slice(t.index + t[0].length)) };
    },
    /** reference line formatting: line colour (and its opacity), width, dash, visibility; label font / format
     * @returns {Record<string, any>} */
    reflineStyle() {
      const raw = {};
      const pool = [wb, st, ...owningPanes(null).map(p => p.style)];
      for (const s of pool) for (const f of (s && s.rules.refline) || []) if (!f.field && !f.scope) raw[f.attr] = f.value;
      const alpha = /^#?[0-9a-f]{8}$/i.test(raw["stroke-color"] || "") ? parseInt(raw["stroke-color"].replace(/^#/, "").slice(6), 16) / 255 : 1;
      return { ...tfMerge(...pool.map(s => tfCollect(s, ["refline"]))), dash: raw["line-pattern-only"] || null, strokeAlpha: alpha };
    },

    /* axis rulers (Format → Lines → Axis Rulers) */
    axisLineShown(shelf) {
      const a = tfMerge(tfCollect(wb, ["axis"], { scope: shelf }), tfCollect(st, ["axis"], { scope: shelf }));
      return !(a.strokeSize === 0 || a.lineVisible === false);
    },
    /* "Show field labels for rows/columns" (worksheet display-field-labels) */
    fieldLabelsShown(scope) {
      const v = tfCollect(st, ["worksheet"], { scope }).displayFieldLabels;
      return v === undefined ? true : v;
    },
    /* Rows then Columns, discrete pills only, in shelf order */
    headerRefs() {
      return sheet ? [...sheet.rows, ...sheet.cols].filter(r => r.type !== "qk" && r.name !== "Measure Names") : [];
    },
    /* measure order: Measure Names manual sort, else the Measure Names filter order */
    measureOrder() {
      if (!sheet) return [];
      const ms = sheet.manualSorts.find(m => m.field && m.field.name === "Measure Names");
      const src = ms ? ms.order : sheet.measureFilter;
      return src.map(x => tfParseFieldRef(String(x).replace(/^"|"$/g, ""))).filter(Boolean);
    },
    /* Measure Names text ("Typical home value") → measure field ref */
    measureRefForName(text) {
      const key = tfNorm(text);
      const cands = [...this.measureOrder(), ...(sheet ? sheet.fieldRefs : [])];
      return cands.find(r => tfDisplayNames(model, r).includes(key)) || matchName(text);
    },
    /* panes in visual order with the fields they label */
    panesInOrder() {
      // a pane drawn on its own axis sits where that axis is on Columns / Rows (one pane per measure)
      const shelfPos = p => {
        const x = p.xAxisName ? sheet.cols.findIndex(r => tfSameField(r, tfParseFieldRef(p.xAxisName))) : -1;
        if (x >= 0) return x;
        return p.yAxisName ? sheet.rows.findIndex(r => tfSameField(r, tfParseFieldRef(p.yAxisName))) : -1;
      };
      return panes.map((p, i) => ({
        pane: p, i,
        order: p.xIndex !== undefined && p.xIndex !== null ? p.xIndex
          : shelfPos(p) >= 0 ? shelfPos(p) : (p.xAxisName ? 0 : (panes.length > 1 ? -1 : 0)),
        refs: [...p.labelRuns.flatMap(r => r.refs),
               ...p.encodings.filter(e => e.channel === "text" || e.channel === "label").map(e => e.field)]
      })).sort((a, b) => a.order - b.order || a.i - b.i);
    },
    /* ── pies and donuts ── */
    /** pie layers { outer, hole, inner } on a dual axis (MIN(0) twice); the later axis draws on top.
     *   single pie                                       → { outer }
     *   donut: + a Pie / Circle layer with nothing on its card splitting it into slices → { outer, hole }
     *   nested donut: two coloured pies, the top one smaller                           → { outer, inner }
     * A layer underneath that the top layer covers is dropped (plain pie). The id-less "All" pane draws nothing.
     * null: not a pie (or 3+ layers). */
    pieLayers() {
      const drawn = panes.length > 1 && panes.some(p => p.id) ? panes.filter(p => p.id) : panes;
      const isPieP = p => /^pie$/i.test(p.markClass || "");
      if (drawn.length === 1) return isPieP(drawn[0]) ? { outer: drawn[0], hole: null, inner: null } : null;
      if (drawn.length !== 2) return null;
      const splits = p => p.encodings.some(e => e.channel === "wedge-size" || (/^(color|lod)$/.test(e.channel) && e.field.type !== "qk"));
      const oi = drawn.findIndex(splits), hi = 1 - oi;
      if (oi < 0) return null;
      if (splits(drawn[hi])) {                                   // nested donut
        if (!drawn.every(isPieP)) return null;
        const size = p => parseFloat(markRule(p, "size")) || 1;
        const [bottom, top] = drawn;
        return size(top) < size(bottom) ? { outer: bottom, hole: null, inner: top } : { outer: top, hole: null, inner: null };
      }
      if (!isPieP(drawn[oi]) || !/^(pie|circle)$/i.test(drawn[hi].markClass || "")) return null;
      return { outer: drawn[oi], hole: hi > oi ? drawn[hi] : null, inner: null };
    },
    /** Marks → Color for a pane without a colour field, RRGGBB; Tableau blue if unset */
    markColor() {
      const v = panes.map(p => markRule(p, "mark-color")).find(Boolean);
      return ((v && tfArgb(v)) || TABLEAU_10[0]).slice(2);
    },
    /** the axis measure of each pane, e.g. one layer's own MIN(0) of a dual-axis donut */
    axisRefs() { return panes.map(p => p.yAxisName || p.xAxisName).filter(Boolean).map(tfParseFieldRef); },
    /** fields on the Marks card by channel ("color", "wedge-size", "text" …) @param {RegExp} channelRe */
    encodingRefs(channelRe) { return panes.flatMap(p => p.encodings.filter(e => channelRe.test(e.channel)).map(e => e.field)); },
    baseStyle: base,
    /** Marks → Size slider value; null = never moved */
    markSize() {
      for (const p of panes) {
        const v = parseFloat(markRule(p, "size"));
        if (isFinite(v)) return v;
      }
      return null;
    },
    /** Label → "Allow labels to overlap other marks" off (Tableau's default) → overlapping labels hidden */
    markLabelsCulled() { return panes.some(p => markRule(p, "mark-labels-cull") !== "false"); },
    /** Label → "Show mark labels"; unset = on when something is on Label */
    markLabelsShown() {
      return panes.some(p => {
        const f = markRule(p, "mark-labels-show");
        if (f !== undefined) return f === "true";
        return p.labelRuns.length > 0 || p.encodings.some(e => e.channel === "text" || e.channel === "label");
      });
    },
    /** label text as runs [{text, props}]: the Label editor, else Tableau's default (each field on Label on its
     * own line); text holds <[ds].[field]> placeholders */
    labelRunsTemplate() {
      const p = panes[0];
      if (!p) return [];
      if (p.labelRuns.length) return p.labelRuns.map(r => ({ text: r.text, props: r.props }));
      return p.encodings.filter(e => e.channel === "text" || e.channel === "label")
        .map((e, i) => ({ text: (i ? "\n" : "") + "<" + e.field.raw + ">", props: {} }));
    },
    hasCustomTooltip() { return !!(panes[0] && panes[0].tooltipRuns && panes[0].tooltipRuns.length); },
    /** tooltip text as runs: the Tooltip editor, else Tableau's default "Caption: <field>" per field in the view */
    tooltipTemplate() {
      const p = panes[0];
      if (p && p.tooltipRuns && p.tooltipRuns.length) return p.tooltipRuns.map(r => ({ text: r.text, props: r.props }));
      /** @type {FieldRef[]} */
      const refs = [];
      const add = r => { if (r && !refs.some(x => tfSameField(x, r))) refs.push(r); };
      // dual axis: a pane shows only its own axis measure, not the other layer's
      const axis = p && (p.yAxisName || p.xAxisName) ? tfParseFieldRef(p.yAxisName || p.xAxisName) : null;
      (sheet ? [...sheet.rows, ...sheet.cols] : []).filter(r => !axis || r.type !== "qk" || tfSameField(r, axis)).forEach(add);
      (p ? p.encodings : []).forEach(e => add(e.field));
      // dimensions first, then measures by instance name – as Tableau lists them on the dashboards seen so far
      refs.sort((a, b) => Number(a.type === "qk") - Number(b.type === "qk") || (a.type === "qk" ? a.inner.localeCompare(b.inner) : 0));
      // quick table calc "[pcto:sum:Qty:qk]" → "% of Total Qty along Table (Across)"
      const ALONG = { rows: " along Table (Across)", columns: " along Table (Down)" };
      const cap = r => /^pcto$/i.test(r.deriv || "")
        ? "% of Total " + this.captionFor({ ...r, deriv: r.name.split(":")[0], name: r.name.split(":").pop() }, r.name) +
          (ALONG[String(((sheet && sheet.tableCalcs) || {})[r.inner.toLowerCase()] || "").toLowerCase()] || "")
        : this.captionFor(r, r.name);
      return refs.map((r, i) => ({ text: (i ? "\n" : "") + cap(r) + ": <" + r.raw + ">", props: {} }));
    },
    /** font of a label run: worksheet font < Format → Label < the run's own font @param {Record<string, any>} runProps */
    labelFont(runProps) {
      return tfMerge(base(), tfCollect(wb, TF_ELEMENTS.datalabel), tfCollect(st, TF_ELEMENTS.datalabel),
        ...panes.map(p => tfCollect(p.style, TF_ELEMENTS.datalabel)), runProps);
    },
    /** Excel number format of a field's values @param {FieldRef} ref */
    numFmtFor: ref => numFmt(ref, false),
    /** the sheet's fit on a dashboard: "entire-view" | "fit-width" | "fit-height" | null (Standard) @param {string} dashboardName */
    fitMode(dashboardName) {
      const d = model && model.dashboards && model.dashboards[dashboardName];
      return (d && d.fit && d.fit[sheetName]) || null;
    },

    /* chart vs table: a visible continuous axis, or only chart marks with no text */
    isChart() {
      if (!sheet) return false;
      if (this.pieLayers()) return true;                         // pies usually carry labels
      for (const shelf of ["rows", "cols"]) {
        for (const r of sheet[shelf]) {
          const continuous = r.type === "qk" || r.name === "Multiple Values";
          if (!continuous) continue;
          const ax = tfCollect(st, ["axis"], { field: r, scope: shelf });
          if (ax.display !== false) return true;
        }
      }
      const chartMarks = /^(bar|line|area|pie|map|multipolygon|polygon|ganttbar|density)$/i;
      return panes.length > 0 && panes.every(p => chartMarks.test(p.markClass || "") &&
        !p.encodings.some(e => e.channel === "text" || e.channel === "label") && !p.labelRuns.length);
    },
    manualSortFor(ref) { return sheet ? sheet.manualSorts.find(m => tfSameField(m.field, ref)) : null; },
    measureSortFor(ref) { return sheet ? sheet.measureSorts.find(m => tfSameField(m.field, ref)) : null; },
    /** the sheet's quick table calculations over a measure ("cum:sum:Sales:qk:7" over "sum:Sales:qk") */
    tableCalcRefs(ref) {
      return ref && sheet ? sheet.fieldRefs.filter(r => { const b = tfTableCalcBase(r); return !!b && tfSameField(b, ref); }) : [];
    },
    /** a running total (Quick Table Calculation → Running Total) */
    isRunningTotal(ref) {
      return !!ref && (/^(cum|rsum)$/i.test(ref.deriv || "") || (sheet ? (sheet.runningTotals || []) : []).some(t => tfSameField(t, ref)));
    },
    captionFor(ref, fallbackName) {
      const base = tfTableCalcBase(ref);                       // a running total is captioned as its measure
      if (base) return this.captionFor(base, fallbackName);
      const info = tfFieldInfo(model, ref);
      let cap = info && info.caption ? info.caption
              : ref && !/^Calculation_/.test(ref.name) ? ref.name
              : String(fallbackName || "").replace(/^[A-Z]+\((.*)\)$/, "$1");
      const d = ref && (ref.deriv || "").toLowerCase();
      const pre = { avg: "Avg. ", cnt: "Count of ", ctd: "Count Distinct of ", min: "Min. ", max: "Max. ",
                    med: "Median ", std: "Std. dev. of ", var: "Variance of " }[d];
      return pre ? pre + cap : cap;
    },
    /* title text with <Sheet Name>, <[field]> and <[Parameters].[p]> substituted */
    renderTitle(valuesFor, sheetLabel) {
      const runs = sheet && sheet.title ? sheet.title.runs : null;
      const props = this.titleStyle();
      if (!runs || !runs.length) return { text: sheetLabel, props };
      let text = runs.map(r => r.text).join("").replace(/\u00C6\r?\n?/g, "\n");
      text = text.replace(/<([^<>]+)>/g, (m, inner) => {
        if (/^sheet name$/i.test(inner)) return sheetLabel;
        if (/^(workbook name|page name|page count|page number)$/i.test(inner)) return "";
        if (!/^\[/.test(inner)) return m;
        const ref = tfParseFieldRef(inner);
        if (!ref) return m;
        if (ref.ds === "Parameters") {
          const info = tfFieldInfo(model, ref) || {};
          return info.alias || info.value || "";
        }
        const vals = valuesFor(ref);
        if (!vals.length) return "";
        return vals.length <= 3 ? vals.join(", ") : "*";
      });
      return { text: text.trim(), props };
    },

    /* display names for Measure-Names buckets like "[ds].[sum:Sales:qk]" */
    bucketAliases(bucketKey) {
      if (!/^\[.*\]$/.test(bucketKey)) return [bucketKey];
      const ref = tfParseFieldRef(bucketKey);
      return ref ? [bucketKey, ...tfDisplayNames(model, ref)] : [bucketKey];
    }
  };
}


/* ═══ format/excel-style.js ═══════════════════════════════════════════════════════════════════════ */
/* Resolved Tableau style props → ExcelJS font / fill / alignment, and cell writing. */

function tfExcelFont(p) {
  let name = p.fontName || "Arial";
  let bold = !!p.bold;
  if (FORMAT_CONFIG.substituteTableauFonts && /^tableau\b/i.test(name)) {
    if (/\b(bold|semibold|black|heavy)\b/i.test(name)) bold = true;  // weight lives in the family name
    name = FORMAT_CONFIG.tableauFontSubstitute;
  }
  const f = { name, size: p.fontSize || 9, bold, italic: !!p.italic };
  if (p.underline) f.underline = true;
  if (p.strike) f.strike = true;
  if (p.color) f.color = { argb: p.color };
  return f;
}

function tfExcelAlignment(p, isNumber) {
  // a number shrinks to its column (columns keep the dashboard's widths) instead of showing ####
  return { horizontal: p.hAlign || (isNumber ? "right" : "left"), vertical: p.vAlign || "middle", wrapText: !!p.wrap,
           ...(isNumber && !p.wrap ? { shrinkToFit: true } : {}) };
}

/** @param {string} argb @returns {import("exceljs").Fill | undefined} */
function tfExcelFill(argb) {
  return argb ? { type: "pattern", pattern: "solid", fgColor: { argb } } : undefined;
}

function tfRichRuns(runs, baseProps) {
  const out = [];
  runs.forEach(r => {
    const text = r.text.replace(/\u00C6\r?\n?/g, "\n");
    if (text) out.push({ text, font: tfExcelFont(tfMerge(baseProps, r.props)) });
  });
  return out;
}

/* writes value + style; returns true if the cell holds a real number */
function tfWriteCell(cell, dv, p, extra = {}) {
  const native = dv ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : undefined;
  const formatted = dv ? (dv.formattedValue != null ? dv.formattedValue : (native != null ? String(native) : "")) : "";
  let isNumber = false;
  if (dv && (dv.value === null || dv.nativeValue === null || dv.value === "%null%")) {   // Tableau Null → blank
    cell.value = "";
  } else if (FORMAT_CONFIG.writeNativeNumbers && typeof native === "number" && isFinite(native)) {
    const numFmt = p.numFmt || inferExcelNumFmt(formatted, native);
    if (numFmt) { cell.value = native; cell.numFmt = numFmt; isNumber = true; }
  }
  if (!isNumber && cell.value !== "") cell.value = formatted;
  if (extra.hyperlink) cell.value = { text: extra.hyperlinkText || extra.hyperlink, hyperlink: extra.hyperlink };
  const looksNumeric = isNumber || /^[^\d]{0,3}-?[\d.,]+[^\d]{0,3}$/.test(String(formatted).trim());
  const fill = extra.fill !== undefined ? extra.fill : p.bgColor;
  const font = tfExcelFont(extra.fontColor ? { ...p, color: extra.fontColor } : p);
  if (fill && !extra.fontColor && !p.explicitColor && tfBrightness(fill) < 128) font.color = { argb: "FFFFFFFF" };
  cell.font = font;
  cell.alignment = tfExcelAlignment(p, looksNumeric);
  if (fill) cell.fill = tfExcelFill(fill);
  if (extra.border) {
    const b = {};
    for (const k in extra.border) if (extra.border[k]) b[k] = extra.border[k];
    if (Object.keys(b).length) cell.border = b;
  }
  return isNumber;
}


/* ═══ twb/parser.js ═══════════════════════════════════════════════════════════════════════════════ */
/* TWB/TWBX XML → plain JSON format model (worksheets, panes, styles, dashboards, actions). */

function tfParseRunProps(run) {
  const a = n => run.getAttribute(n);
  const align = { "0": "left", "1": "center", "2": "right" }[a("fontalignment")];
  return tfDefined({
    fontName: a("fontname") || undefined,
    fontSize: tfNum(a("fontsize")),
    bold: a("bold") != null ? a("bold") === "true" : undefined,
    italic: a("italic") != null ? a("italic") === "true" : undefined,
    underline: a("underline") != null ? a("underline") === "true" : undefined,
    color: tfArgb(a("fontcolor")) || undefined,
    hAlign: align
  });
}

/* a dashboard zone's Layout pane formatting: background and border */
function tfZoneStyle(el) {
  if (!el) return undefined;
  const f = {};
  tfKids(el, "format").forEach(x => { f[x.getAttribute("attr")] = x.getAttribute("value"); });
  const style = tfDefined({
    bgColor: tfArgb(f["background-color"]) || undefined,
    borderColor: tfArgb(f["border-color"]) || undefined,
    borderStyle: f["border-style"] || undefined,
    borderWidth: tfNum(f["border-width"])
  });
  return Object.keys(style).length ? style : undefined;
}

function tfParseRuns(ftEl) {
  return tfKids(ftEl, "run").map(r => ({ text: r.textContent || "", props: tfParseRunProps(r) }));
}

function tfBucketKey(t) {
  let s = String(t == null ? "" : t).trim();
  if (/^".*"$/.test(s)) s = s.slice(1, -1).replace(/""/g, '"');
  else if (/^#.*#$/.test(s)) s = s.slice(1, -1);
  if (s === "%null%") s = "null";
  return s.toLowerCase();
}

function tfParseColorEncoding(enc) {
  const a = n => enc.getAttribute(n);
  const cp = enc.getElementsByTagName("color-palette")[0] || null;
  const customColors = cp
    ? Array.from(cp.getElementsByTagName("color")).map(c => tfArgb(c.textContent)).filter(Boolean)
    : [];
  const map = {};
  Array.from(enc.getElementsByTagName("map")).forEach(m => {
    const color = tfArgb(m.getAttribute("to"));
    if (!color) return;
    Array.from(m.getElementsByTagName("bucket")).forEach(b => { map[tfBucketKey(b.textContent)] = color; });
  });
  return {
    field: tfParseFieldRef(a("field")),
    type: a("type"),                                  // "palette" (categorical) | "interpolated" | …
    paletteName: a("palette") || (cp && cp.getAttribute("name")) || null,
    paletteType: cp ? cp.getAttribute("type") : null, // ordered-sequential | ordered-diverging | regular
    customColors,
    reverse: a("reverse") === "true" || (cp && cp.getAttribute("reverse") === "true"),
    // optional range settings (only used if present in your TWB)
    center: tfNum(a("center")), min: tfNum(a("min")), max: tfNum(a("max")),
    map
  };
}

/** @param {Element | null} styleEl @returns {ParsedStyle} */
function tfParseStyle(styleEl) {
  /** @type {ParsedStyle["rules"]} */
  const rules = {};
  const encodings = [];
  if (!styleEl) return { rules, encodings };
  tfKids(styleEl, "style-rule").forEach(rule => {
    const el = rule.getAttribute("element") || "all";
    rules[el] = rules[el] || [];
    tfKids(rule, "format").forEach(f => rules[el].push(tfDefined({
      attr: f.getAttribute("attr"),
      value: f.getAttribute("value"),
      field: f.getAttribute("field") || undefined,
      scope: f.getAttribute("scope") || undefined,
      dataClass: f.getAttribute("data-class") || undefined,
      axisClass: f.getAttribute("class") || undefined
    })));
  });
  Array.from(styleEl.getElementsByTagName("encoding"))
    .filter(e => e.getAttribute("attr") === "color")
    .forEach(e => encodings.push(tfParseColorEncoding(e)));
  // Edit Axis: fixed range, tick spacing, "include zero" (encoding attr="space")
  const spaces = Array.from(styleEl.getElementsByTagName("encoding")).filter(e => e.getAttribute("attr") === "space").map(e => tfDefined({
    field: e.getAttribute("field") || undefined,
    scope: e.getAttribute("scope") || undefined,
    axisClass: e.getAttribute("class") || undefined,
    rangeType: e.getAttribute("range-type") || undefined,
    min: tfNum(e.getAttribute("min")),
    max: tfNum(e.getAttribute("max")),
    majorSpacing: tfNum(e.getAttribute("major-spacing")),
    majorShow: e.getAttribute("major-show") || undefined,      // "false": Edit Axis → Tick Marks → None
    domainExpand: e.getAttribute("domain-expand") || undefined,
    reverse: e.getAttribute("reverse") === "true" || undefined, // Edit Axis → Scale → Reversed (rank 1 at the top)
    fold: e.getAttribute("fold") === "true" || undefined,       // Dual Axis: this axis is folded onto the one before it
    synchronized: e.getAttribute("synchronized") === "true" || undefined   // … with the same scale (Synchronize Axis)
  }));
  // Shape encodings: value → shape ("Zoom Icons/Zoom in.png" = a custom shape, ":filled/circle" = Tableau's)
  const shapes = Array.from(styleEl.getElementsByTagName("encoding")).filter(e => e.getAttribute("attr") === "shape").map(e => {
    /** @type {Record<string, string>} */
    const map = {};
    Array.from(e.getElementsByTagName("map")).forEach(m => Array.from(m.getElementsByTagName("bucket"))
      .forEach(b => { map[tfBucketKey(b.textContent)] = m.getAttribute("to"); }));
    return { field: tfParseFieldRef(e.getAttribute("field")), map };
  }).filter(s => s.field);
  /** @type {ParsedStyle} */
  const out = { rules, encodings };
  if (spaces.length) out.spaces = spaces;
  if (shapes.length) out.shapes = shapes;
  return out;
}

function tfParseTitle(ownerEl) {
  const lo = tfKid(ownerEl, "layout-options") || tfKid(ownerEl, "layout");
  const t = lo && tfKid(lo, "title");
  const ft = t && tfKid(t, "formatted-text");
  if (!ft) return null;
  const runs = tfParseRuns(ft);
  return { text: runs.map(r => r.text).join("").trim(), runs };
}

/**
 * A calculation that is one number ("0.5", "MIN(0)"): an axis that only positions marks, e.g. the columns
 * of a table built from marks.
 * @param {Element | null} calc @returns {number | undefined}
 */
function tfConstantFormula(calc) {
  const m = calc && String(calc.getAttribute("formula") || "").match(/^\s*(?:(?:MIN|MAX|AVG|SUM|ATTR)\s*\(\s*(-?\d+(?:\.\d+)?)\s*\)|(-?\d+(?:\.\d+)?))\s*$/i);
  return m ? Number(m[1] ?? m[2]) : undefined;
}

/**
 * What kind of calculation a calculated field is, from its formula – for the conversion report (the export
 * writes the values Tableau computed; formulas are not translated).
 * @param {string} formula @returns {string}
 */
function tfCalcKind(formula) {
  const f = String(formula || "").replace(/\/\/[^\n]*|\/\*[\s\S]*?\*\//g, " ");
  if (/\{\s*(FIXED|INCLUDE|EXCLUDE)\b|\{\s*(SUM|AVG|MIN|MAX|COUNTD?)\s*\(/i.test(f)) return "LOD expression";
  if (/\b(WINDOW_\w+|RUNNING_\w+|RANK(_\w+)?|INDEX|FIRST|LAST|LOOKUP|PREVIOUS_VALUE|SIZE|TOTAL)\s*\(/i.test(f)) return "table calculation";
  if (/\b(SUM|AVG|MIN|MAX|COUNTD?|MEDIAN|ATTR|STDEVP?|VARP?|PERCENTILE)\s*\(/i.test(f)) return "aggregation";
  if (/\b(DATEPART|DATETRUNC|DATEADD|DATEDIFF|DATENAME|DATEPARSE|MAKEDATE|TODAY|NOW|YEAR|QUARTER|MONTH|WEEK|DAY)\s*\(/i.test(f)) return "date calculation";
  if (/\b(IF|CASE|IIF|ZN|IFNULL|ISNULL)\b/i.test(f)) return "conditional logic";
  if (/\b(MAKEPOINT|MAKELINE|BUFFER|DISTANCE|SCRIPT_\w+|RAWSQL\w*|REGEXP_\w+|SPLIT)\s*\(/i.test(f)) return "Tableau-specific function";
  return /[-+*/]/.test(f) ? "arithmetic" : "expression";
}

/** a datasource column's calculation: a constant, a bin, or a calculated field (its kind and formula)
 * @param {Element | null} calc */
function tfCalcInfo(calc) {
  if (!calc) return {};
  const cls = calc.getAttribute("class") || "";
  const formula = calc.getAttribute("formula") || "";
  if (cls === "bin") return { bin: true };
  if (cls === "categorical-bin") return { group: true };
  if (!formula) return {};
  return { constant: tfConstantFormula(calc), calcKind: tfCalcKind(formula), formula: formula.length > 160 ? formula.slice(0, 157) + "…" : formula };
}

/* Raise when the parser reads something new: a workbook model remembered by an older version is then
   parsed again from its stored XML (ui/workbook-store.js), so an update reaches workbooks loaded before it. */
const FORMAT_MODEL_VERSION = 5;

/** @param {string} xmlString the .twb XML @returns {FormatModel} */
function parseTableauFormatting(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const root = doc.documentElement;
  const model = { version: FORMAT_MODEL_VERSION, workbookStyle: tfParseStyle(tfKid(root, "style")),
                  fields: {}, datasourceStyles: {}, sheets: {}, dashboards: {}, measureAliases: {} };

  // ── datasource columns: captions, roles, default number formats, colour maps
  tfKids(tfKid(root, "datasources"), "datasource").forEach(ds => {
    const dsName = ds.getAttribute("name");
    model.datasourceStyles[dsName] = tfParseStyle(tfKid(ds, "style"));
    tfKids(ds, "column").forEach(col => {
      const n = (col.getAttribute("name") || "").replace(/^\[|\]$/g, "");
      if (!n) return;
      const info = tfDefined({
        name: n, ds: dsName,
        caption: col.getAttribute("caption") || undefined,
        role: col.getAttribute("role") || undefined,
        datatype: col.getAttribute("datatype") || undefined,
        defaultFormat: col.getAttribute("default-format") || undefined,
        alias: col.getAttribute("alias") || undefined,      // parameters: current value's display text
        value: col.getAttribute("value") || undefined,
        param: col.getAttribute("param-domain-type") ? true : undefined,
        geoRole: col.getAttribute("semantic-role") || undefined,   // "[State].[Name]": Tableau geocodes it on a map
        ...tfCalcInfo(tfKid(col, "calculation"))
      });
      model.fields[dsName + "|" + n.toLowerCase()] = info;
      if (!model.fields["|" + n.toLowerCase()]) model.fields["|" + n.toLowerCase()] = info;
      // Measure Names aliases: key "[ds].[sum:HomeValue:qk]" → "Typical home value"
      if (n === ":Measure Names") {
        Array.from(col.getElementsByTagName("alias")).forEach(al => {
          const r = tfParseFieldRef(al.getAttribute("key"));
          if (r) model.measureAliases[tfRefKey(r)] = al.getAttribute("value");
        });
      }
    });
  });

  // ── worksheets
  tfKids(tfKid(root, "worksheets"), "worksheet").forEach(ws => {
    const name = ws.getAttribute("name");
    const table = tfKid(ws, "table");
    const sheet = {
      name,
      title: tfParseTitle(ws),
      style: tfParseStyle(table && tfKid(table, "style")),
      rows: tfExtractRefs(table && tfKid(table, "rows") && tfKid(table, "rows").textContent),
      cols: tfExtractRefs(table && tfKid(table, "cols") && tfKid(table, "cols").textContent),
      // Analysis → Totals → Show Row / Column Grand Totals (<rows total='true'> / <cols total='true'>)
      grandTotals: { rows: !!(table && tfKid(table, "rows") && tfKid(table, "rows").getAttribute("total") === "true"),
                     cols: !!(table && tfKid(table, "cols") && tfKid(table, "cols").getAttribute("total") === "true") },
      panes: [], fieldRefs: []
    };
    const view = table && tfKid(table, "view");
    // columns declared in this sheet's datasource-dependencies (covers blended/secondary sources)
    Array.from(ws.getElementsByTagName("datasource-dependencies")).forEach(dep => {
      const dsName = dep.getAttribute("datasource");
      tfKids(dep, "column").forEach(col => {
        const n = (col.getAttribute("name") || "").replace(/^\[|\]$/g, "");
        if (!n) return;
        const key = dsName + "|" + n.toLowerCase();
        const cur = model.fields[key] || {};
        // a calculation typed into a shelf exists only here (an unnamed calc: "avg(0)" placing a butterfly's labels)
        model.fields[key] = tfMerge(tfDefined({ name: n, ds: dsName, caption: col.getAttribute("caption") || undefined,
          role: col.getAttribute("role") || undefined, datatype: col.getAttribute("datatype") || undefined,
          defaultFormat: col.getAttribute("default-format") || undefined,
          alias: col.getAttribute("alias") || undefined, value: col.getAttribute("value") || undefined,
          ...tfCalcInfo(tfKid(col, "calculation")) }), cur);
        if (!model.fields["|" + n.toLowerCase()]) model.fields["|" + n.toLowerCase()] = model.fields[key];
      });
    });
    sheet.manualSorts = view ? Array.from(view.getElementsByTagName("manual-sort")).map(ms => ({
      field: tfParseFieldRef(ms.getAttribute("column")),
      direction: (ms.getAttribute("direction") || "ASC").toUpperCase(),
      order: Array.from(ms.getElementsByTagName("bucket")).map(b => b.textContent)
    })) : [];
    sheet.measureSorts = view ? [
      ...Array.from(view.getElementsByTagName("shelf-sort-v2")).map(ss => ({
        field: tfParseFieldRef(ss.getAttribute("dimension-to-sort")),
        measure: tfParseFieldRef(ss.getAttribute("measure-to-sort-by")),
        direction: (ss.getAttribute("direction") || "ASC").toUpperCase() })),
      ...Array.from(view.getElementsByTagName("computed-sort")).map(cs => ({
        field: tfParseFieldRef(cs.getAttribute("column")),
        measure: tfParseFieldRef(cs.getAttribute("using")),
        direction: (cs.getAttribute("direction") || "ASC").toUpperCase() }))
    ].filter(x => x.field && x.measure) : [];
    // members kept by the Measure Names filter, in filter order
    sheet.measureFilter = [];
    if (view) tfKids(view, "filter").forEach(f => {
      if (!/\[:Measure Names\]$/.test(f.getAttribute("column") || "")) return;
      Array.from(f.getElementsByTagName("groupfilter")).forEach(g => {
        const m = g.getAttribute("member");
        if (g.getAttribute("function") === "member" && m) sheet.measureFilter.push(m);
      });
    });
    const panesEl = table && tfKid(table, "panes");
    tfKids(panesEl, "pane").forEach(p => {
      const mark = tfKid(p, "mark");
      const encodings = tfKids(tfKid(p, "encodings")).map(e => ({
        channel: e.tagName, field: tfParseFieldRef(e.getAttribute("column"))
      })).filter(e => e.field);
      const cl = tfKid(p, "customized-label");
      const labelRuns = cl ? tfParseRuns(tfKid(cl, "formatted-text")) : [];
      const ct = tfKid(p, "customized-tooltip");                 // Tooltip editor text
      const tooltipRuns = ct ? tfParseRuns(tfKid(ct, "formatted-text")) : [];
      sheet.panes.push({
        id: p.getAttribute("id") || null,
        xIndex: tfNum(p.getAttribute("x-index")),
        xAxisName: p.getAttribute("x-axis-name") || null,
        yIndex: tfNum(p.getAttribute("y-index")),
        yAxisName: p.getAttribute("y-axis-name") || null,     // dual axis / multi-pane: which measure this pane draws
        markClass: mark ? mark.getAttribute("class") : "Automatic",
        encodings,
        labelRuns: labelRuns.map(r => ({ refs: tfExtractRefs(r.text.replace(/^<|>$/g, "")), text: r.text, props: r.props })),
        tooltipRuns,
        style: tfParseStyle(tfKid(p, "style"))
      });
    });
    // Analytics pane box plot: <reference-line formula='iqr' boxplot-whisker-type='…'>
    sheet.boxPlot = Array.from(ws.getElementsByTagName("reference-line")).some(rl =>
      rl.getAttribute("formula") === "iqr" || rl.getAttribute("boxplot-whisker-type") != null);
    // reference lines (Analytics → Reference Line); box-plot whiskers / distributions are not lines
    const refIds = new Set();
    sheet.referenceLines = Array.from(ws.getElementsByTagName("reference-line")).filter(rl => {
      const id = rl.getAttribute("id") || "";
      if (refIds.has(id) || rl.getAttribute("boxplot-whisker-type") != null) return false;
      refIds.add(id);
      return /^(average|mean|median|sum|total|min|max|constant)$/.test(rl.getAttribute("formula") || "");
    }).map(rl => tfDefined({
      formula: rl.getAttribute("formula"),
      labelType: rl.getAttribute("label-type") || "automatic",
      label: rl.getAttribute("label") || undefined,
      scope: rl.getAttribute("scope") || undefined,
      value: tfNum(rl.getAttribute("value")),
      axis: tfParseFieldRef(rl.getAttribute("axis-column")) || undefined,
      field: tfParseFieldRef(rl.getAttribute("value-column")) || undefined
    }));
    // every field the sheet references (for name matching incl. Measure Names)
    const seen = new Set();
    const add = r => { if (r && !seen.has(r.inner.toLowerCase())) { seen.add(r.inner.toLowerCase()); sheet.fieldRefs.push(r); } };
    sheet.runningTotals = [];                                    // waterfall: running-sum table calcs
    sheet.tableCalcs = {};                                       // quick table calc → its "compute using" (ordering-type)
    Array.from(ws.getElementsByTagName("column-instance")).forEach(ci => {
      const depDs = ci.parentNode && ci.parentNode.getAttribute && ci.parentNode.getAttribute("datasource");
      const ref = tfParseFieldRef((depDs ? "[" + depDs + "]." : "") + ci.getAttribute("name"));
      add(ref);
      const calc = tfKid(ci, "table-calc");
      if (ref && calc && /^(cumtotal|runningtotal)$/i.test(calc.getAttribute("type") || "")) sheet.runningTotals.push(ref);
      if (ref && calc) sheet.tableCalcs[ref.inner.toLowerCase()] = calc.getAttribute("ordering-type") || "";
    });
    sheet.rows.forEach(add); sheet.cols.forEach(add);
    sheet.panes.forEach(p => p.encodings.forEach(e => add(e.field)));
    model.sheets[name] = sheet;
  });

  // ── URL actions (not filter/highlight actions) → Excel hyperlinks
  model.actions = [];
  tfKids(tfKid(root, "actions"), "action").forEach(a => {
    const link = tfKid(a, "link"), src = tfKid(a, "source");
    if (!link || tfKid(a, "command")) return;                   // filter/highlight actions have a <command>
    const expr = link.getAttribute("expression") || "";
    if (!expr || /^tsl:/i.test(expr)) return;
    model.actions.push(tfDefined({
      caption: a.getAttribute("caption") || undefined, expression: expr,
      dashboard: src ? src.getAttribute("dashboard") || undefined : undefined,
      worksheet: src ? src.getAttribute("worksheet") || undefined : undefined,
      exclude: src ? Array.from(src.getElementsByTagName("exclude-sheet")).map(x => x.getAttribute("name")) : []
    }));
  });

  // ── dashboards (title + style)
  tfKids(tfKid(root, "dashboards"), "dashboard").forEach(d => {
    const zones = [];
    const walk = (z, parent) => tfKids(z, "zone").forEach(c => {
      zones.push(tfDefined({
        id: c.getAttribute("id"), name: c.getAttribute("name") || undefined,
        type: c.getAttribute("type-v2") || "worksheet",
        showTitle: c.getAttribute("show-title") !== "false",
        hidden: c.getAttribute("hidden-by-user") === "true",
        x: tfNum(c.getAttribute("x")), y: tfNum(c.getAttribute("y")),
        w: tfNum(c.getAttribute("w")), h: tfNum(c.getAttribute("h")),
        runs: tfKid(c, "formatted-text") ? tfParseRuns(tfKid(c, "formatted-text")) : undefined,
        param: c.getAttribute("param") || undefined,             // image zones: the file inside the .twbx; legends: their field
        scaled: c.getAttribute("is-scaled") === "1" || undefined,
        // Center Image is on unless switched off (is-centered='0')
        centered: c.getAttribute("is-centered") ? c.getAttribute("is-centered") === "1" : undefined,
        url: c.getAttribute("url") || undefined,                   // image / button zones: the link it opens
        parent,                                                     // the layout container it sits in
        style: tfZoneStyle(tfKid(c, "zone-style"))
      }));
      walk(c, c.getAttribute("id") || undefined);
    });
    walk(tfKid(d, "zones"), undefined);   // direct child only → phone/tablet layouts are skipped
    const size = tfKid(d, "size");
    model.dashboards[d.getAttribute("name")] = {
      title: tfParseTitle(d), style: tfParseStyle(tfKid(d, "style")), zones,
      width: size ? tfNum(size.getAttribute("maxwidth")) : undefined,
      height: size ? tfNum(size.getAttribute("maxheight")) : undefined
    };
  });

  // ── each sheet's fit on each dashboard (<window class='dashboard'> viewpoints: "entire-view", "fit-width" …)
  tfKids(tfKid(root, "windows"), "window").forEach(w => {
    const dash = w.getAttribute("class") === "dashboard" && model.dashboards[w.getAttribute("name")];
    if (!dash) return;
    dash.fit = {};
    tfKids(tfKid(w, "viewpoints"), "viewpoint").forEach(v => {
      const z = tfKid(v, "zoom");
      if (z) dash.fit[v.getAttribute("name")] = z.getAttribute("type");
    });
  });

  tfDebugDump(model);
  return model;
}

/* Prints every style element/attr found – use this to see what YOUR TWB contains */
function tfDebugDump(model) {
  if (!FORMAT_CONFIG.debug) return;
  const summarize = st => {
    const o = {};
    for (const [el, list] of Object.entries(st.rules)) o[el] = [...new Set(list.map(f => f.attr + (f.scope ? "@" + f.scope : "") + (f.field ? "[field]" : "")))].join(", ");
    return o;
  };
  console.log("[Format] Workbook style:", summarize(model.workbookStyle));
  for (const s of Object.values(model.sheets)) {
    console.log(`[Format] Sheet "${s.name}"`, {
      rules: summarize(s.style),
      marks: s.panes.map(p => p.markClass).join(", "),
      encodings: s.panes.flatMap(p => p.encodings.map(e => e.channel + "=" + e.field.inner)).join(", "),
      colorEncodings: s.style.encodings.length,
      labelRuns: s.panes.reduce((n, p) => n + p.labelRuns.length, 0)
    });
  }
}


/* ═══ charts/model/common.js ══════════════════════════════════════════════════════════════════════ */
/* Shared helpers for building renderer-neutral Excel chart specs. */

"use strict";

const TV_MAX_POINTS = 4000;

const TV_MAX_SERIES = 60;

const TV_DEFAULT_MARK_COLOR = "4E79A7";

function tvHex(argb) {
  if (!argb) return null;
  const s = String(argb).replace(/^#/, "").toUpperCase();
  return s.length === 8 ? s.slice(2) : s;
}

function tvText(dv) { return tfIsNull(dv) ? "Null" : tfDvText(dv); }

function tvMarkToken(cls) {
  const t = String(cls || "").toLowerCase().replace(/[\s_-]+/g, "");
  return t === "automatic" ? "" : t;
}

function tvPaneRule(pane, element, attrName) {
  const list = (pane && pane.style && pane.style.rules && pane.style.rules[element]) || [];
  const f = list.find(x => x.attr === attrName && !x.field && !x.scope);
  return f ? f.value : undefined;
}

/* ── formatting helpers ─────────────────────────────────────────────────── */
function tvChartFont(fmt) {
  const f = tfExcelFont(fmt.headerCellStyle(null));
  return { name: f.name, size: f.size || 9, color: tvHex(f.color && f.color.argb) || "333333" };
}

/** @param {ViewModel} vm @param {number} ci @returns {string} */
function tvNumFmt(vm, ci) {
  const c = vm.cols[ci];
  const st = vm.fmt.sheetModel ? vm.fmt.sheetModel.style : null;
  if (c && c.ref && st) {
    const axisRaw = tfCollect(st, ["axis"], { field: c.ref }).numFmtRaw;
    const axis = axisRaw ? tableauToExcelNumFmt(axisRaw) : null;
    if (axis) return axis;
  }
  const own = c && vm.fmt.markCellStyle(c.ref || null).numFmt;
  if (own) return own;
  for (const row of vm.rows) {
    const dv = row[ci];
    if (tfIsNull(dv)) continue;
    const n = tfDvNum(dv);
    if (n === null) continue;
    return inferExcelNumFmt(dv.formattedValue, n) || "General";
  }
  return "General";
}

/** @param {ViewModel} vm @param {number} ci @returns {string} */
function tvMeasureLabel(vm, ci) {
  const c = vm.cols[ci];
  return (c && (c.label || c.name)) || "Value";
}

/* mark of the pane that draws this measure (dual axis / multi-pane), else the sheet's.
 * Without a workbook, the live spec has one marks card per measure, in shelf order. */
/** @param {Roles} roles @param {FieldRef | null} ref @param {string} [fallback] @param {number} [index] @param {number} [count] @returns {string} */
function tvMeasureMark(roles, ref, fallback, index, count) {
  if (!roles.panes.length && roles.liveMarks && count > 1 && roles.liveMarks.length === count && roles.liveMarks[index]) {
    return roles.liveMarks[index];
  }
  if (ref) {
    const p = roles.panes.find(x => [x.yAxisName, x.xAxisName].some(n => n && tfSameField(tfParseFieldRef(n), ref)));
    const t = p && tvMarkToken(p.markClass);
    if (t) return t;
  }
  const all = roles.panes.find(x => !x.id) || roles.panes[0];
  return (all && tvMarkToken(all.markClass)) || fallback || "bar";
}

/** @param {Roles} roles @param {FieldRef | null} ref @returns {boolean} */
function tvLabelsOn(roles, ref) {
  const shown = roles.panes.some(p => tvPaneRule(p, "mark", "mark-labels-show") === "true");
  if (shown) return true;
  if (!ref) return false;
  return roles.labelRefs.some(r => tfSameField(r, ref) || (r.name === "Multiple Values")) ||
    roles.labelNames.some(n => n === tfNorm(ref.name));
}

/** @param {Roles} roles @returns {string} */
function tvMarkColor(roles) {
  for (const p of roles.panes) {
    const c = tvHex(tfArgb(tvPaneRule(p, "mark", "mark-color")));
    if (c) return c;
  }
  return TV_DEFAULT_MARK_COLOR;
}

/* ── colour scale for the Color encoding (TWB mapping or Tableau automatic) ─ */
/**
 * @param {ViewModel} vm @param {Roles} roles @param {string} [markToken]
 * @returns {((value: any, index?: number) => string | null) | null} value → hex colour
 */
function tvColorScale(vm, roles, markToken) {
  if (!roles.color) return null;
  const fmt = vm.fmt;
  /** @type {Partial<ReturnType<SheetFormatter["colorEncoding"]>>} */
  let enc = fmt.colorEncoding();
  if (!enc) {                                      // live spec only: Tableau automatic palettes
    const ci = roles.color.ci;
    const name = ci >= 0 ? vm.cols[ci].name : "Measure Names";
    enc = { ref: { name, inner: name }, def: null, continuous: !!roles.color.continuous, markClass: markToken || "bar" };
  }
  if (roles.color.measureNames) {
    const scale = tfBuildColorScale(fmt, { ...enc, continuous: false }, []);
    return (value, i) => tvHex(scale && scale(value)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]);
  }
  const ci = roles.color.ci;
  const values = vm.rows.map(r => roles.color.continuous ? tfDvNum(r[ci]) : tvText(r[ci]));
  const scale = tfBuildColorScale(fmt, enc, values);
  return scale ? (value => tvHex(scale(value))) : null;
}

/* legend order of a discrete colour field: manual sort, else natural */
/** @param {ViewModel} vm @param {Roles} roles @returns {string[]} */
function tvColorValues(vm, roles) {
  const ci = roles.color.ci;
  const seen = new Map();
  vm.rows.forEach(r => { const t = tvText(r[ci]); if (!seen.has(t)) seen.set(t, r[ci]); });
  let values = [...seen.keys()];
  const manual = roles.color.ref ? vm.fmt.manualSortFor(roles.color.ref) : null;
  const byMeasure = !manual && roles.color.ref ? vm.fmt.measureSortFor(roles.color.ref) : null;
  if (manual) {
    const rank = new Map(manual.order.map((b, i) => [tfBucketKey(b), i]));
    values.sort((a, b) => (rank.has(tfNorm(a)) ? rank.get(tfNorm(a)) : 1e9) - (rank.has(tfNorm(b)) ? rank.get(tfNorm(b)) : 1e9));
    if (manual.direction === "DESC") values.reverse();
  } else if (byMeasure) {
    // sorted by a measure (Sort → Field: Value Ordered, ascending): its total per colour value
    const mi = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, byMeasure.measure));
    const total = new Map(values.map(v => [v, 0]));
    if (mi >= 0) vm.rows.forEach(r => { const t = tvText(r[ci]); total.set(t, (total.get(t) || 0) + (tfDvNum(r[mi]) || 0)); });
    values.sort((a, b) => (total.get(a) - total.get(b)) * (byMeasure.direction === "DESC" ? -1 : 1) || tfNaturalCompare(seen.get(a), seen.get(b)));
  } else values.sort((a, b) => tfNaturalCompare(seen.get(a), seen.get(b)));
  return values;
}

/** a header's date label format (Format → Header → Dates: "iLLLLL" → J F M …)
 * @param {ViewModel} vm @param {number} ci @returns {(text: string, dv: DataValue) => string} */
function tvLabelFormatter(vm, ci) {
  const col = vm.cols[ci];
  const raw = col && col.ref && vm.fmt.labelFormat ? vm.fmt.labelFormat(col.ref) : null;
  if (!raw) return t => t;
  return (text, dv) => tfFormatDateLabel(raw, text, dv) ?? text;
}

/**
 * A marks card's label as lines of text segments: its custom label (Label → Text), else one line per Text
 * field; each field token replaced by valueOf(ref). A token split over runs still resolves, each segment
 * keeps the font of the run it starts in.
 * @param {any} pane @param {(ref: FieldRef) => string} valueOf
 * @returns {{ text: string, props: Record<string, any> }[][]}
 */
function tvPaneLabel(pane, valueOf) {
  const runs = pane.labelRuns.length ? pane.labelRuns.map(r => ({ text: String(r.text), props: r.props || {} }))
    : pane.encodings.filter(e => e.channel === "text" || e.channel === "label")
        .flatMap((e, i) => [...(i ? [{ text: "\n", props: {} }] : []), { text: `<${e.field.raw}>`, props: {} }]);
  const chars = runs.flatMap(r => [...r.text].map(ch => ({ ch, props: r.props })));
  const text = chars.map(c => c.ch).join("");
  /** @type {{ text: string, props: Record<string, any> }[][]} */
  const lines = [[]];
  const push = (t, props) => {
    if (!t) return;
    const line = lines[lines.length - 1], last = line[line.length - 1];
    if (last && last.props === props) last.text += t; else line.push({ text: t, props });
  };
  const re = /<([^<>]+)>|Æ[ \t]*(?:\r?\n)?|\r?\n/g;
  let at = 0;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    for (let k = at; k < m.index; k++) push(chars[k].ch, chars[k].props);
    if (m[1] !== undefined) { const ref = tfParseFieldRef(m[1]); push(ref ? valueOf(ref) : m[0], chars[m.index].props); }
    else lines.push([]);
    at = m.index + m[0].length;
  }
  for (let k = at; k < text.length; k++) push(chars[k].ch, chars[k].props);
  return lines;
}

/* ── category axis: distinct label tuples in view order ─────────────────── */
/**
 * @param {ViewModel} vm
 * @param {number[]} catCis category columns, outer → inner
 * @param {boolean} sortNative sort by value (continuous / date axes) instead of view order
 * @returns {{ count: number, levels: string[][], indexOf: (row: DataValue[]) => number | undefined }}
 */
function tvCategories(vm, catCis, sortNative) {
  let order = [];
  const seen = new Map();
  vm.rows.forEach(row => {
    const labels = catCis.map(ci => tvText(row[ci]));
    const key = labels.join("\u0001");
    if (!seen.has(key)) { seen.set(key, true); order.push({ key, labels, dvs: catCis.map(ci => row[ci]) }); }
  });
  if (sortNative) {
    order.sort((a, b) => {
      for (let l = 0; l < catCis.length; l++) { const d = tfNaturalCompare(a.dvs[l], b.dvs[l]); if (d) return d; }
      return 0;
    });
  }
  if (!catCis.length) order = [{ key: "", labels: [""] }];
  const index = new Map(order.map((o, i) => [o.key, i]));
  // shown labels follow each header's date format (J F M …); the keys stay the raw values, so January,
  // June and July remain three categories even when all are labelled "J"
  const shown = catCis.map(ci => tvLabelFormatter(vm, ci));
  return {
    count: order.length,
    levels: (catCis.length ? catCis : [null]).map((_, l) => order.map(o => catCis.length ? shown[l](o.labels[l], o.dvs[l]) : o.labels[l])),
    indexOf: row => index.get(catCis.map(ci => tvText(row[ci])).join("\u0001"))
  };
}

/**
 * @param {ViewModel} vm
 * @param {{ count: number, indexOf: (row: DataValue[]) => number | undefined }} cats
 * @param {number} valueCi
 * @param {(row: DataValue[]) => boolean} [filter]
 * @returns {(number | null)[]} one total per category, null = no value
 */
function tvSum(vm, cats, valueCi, filter) {
  const out = new Array(cats.count).fill(null);
  vm.rows.forEach(row => {
    if (filter && !filter(row)) return;
    const n = tfDvNum(row[valueCi]);
    if (n === null) return;
    const i = cats.indexOf(row);
    if (i === undefined) return;
    out[i] = (out[i] || 0) + n;
  });
  return out;
}

/* ══════════════════════════════════════════════════════════════════════════
 * chart builders
 * ══════════════════════════════════════════════════════════════════════════ */
/** @param {ViewModel} vm @returns {Pick<ChartSpec, "font" | "background" | "gridlines">} */
function tvBaseSpec(vm) {
  return {
    font: tvChartFont(vm.fmt),
    // Worksheet shading: its colour, none (see-through: the container shows) or Tableau's white
    background: vm.fmt.sheetShading() === "none" ? null : tvHex(vm.fmt.tableBackground()) || "FFFFFF",
    gridlines: true
  };
}

/** @param {string} mark @returns {Pick<ChartSeries, "type" | "line" | "marker">} */
function tvSeriesType(mark) {
  if (mark === "line") return { type: "line" };
  if (mark === "area") return { type: "area" };
  if (/^(circle|shape|square)$/.test(mark)) return { type: "line", line: false, marker: true };
  return { type: "bar" };
}


/* ═══ charts/model/reflines.js ════════════════════════════════════════════════════════════════════ */
/* Analytics → Reference Line on a chart spec, valued, labelled and styled like Tableau. */

const COMPUTATION = { average: "Average", mean: "Average", median: "Median", sum: "Sum", total: "Total",
                      min: "Minimum", max: "Maximum", constant: "Constant" };

/* Excel number-format literal: "text" (quotes inside dropped) */
const lit = s => s ? `"${String(s).replace(/"/g, "")}"` : "";

/** the line's value over the marks of the chart; null = not computable
 * @param {ReferenceLine} rl @param {number[]} values @returns {number | null} */
function tvReferenceValue(rl, values) {
  if (rl.formula === "constant") return typeof rl.value === "number" ? rl.value : null;
  if (!values.length) return null;
  switch (rl.formula) {
    case "average": case "mean": return values.reduce((a, b) => a + b, 0) / values.length;
    case "median": {
      const s = [...values].sort((a, b) => a - b), m = Math.floor(s.length / 2);
      return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    }
    case "sum": case "total": return values.reduce((a, b) => a + b, 0);
    case "min": return Math.min(...values);
    case "max": return Math.max(...values);
    default: return null;
  }
}

/**
 * The label as an Excel number format on the line's own value, so Excel prints it: "Average"
 * (automatic / computation), the value, or a custom text with <Value> / <Computation>.
 * @param {ReferenceLine} rl @param {string} valueFmt @returns {string | null}
 */
function tvReferenceLabelFormat(rl, valueFmt) {
  const name = COMPUTATION[rl.formula] || "";
  const positive = (String(valueFmt || "General").match(/^((?:"[^"]*"|[^;])*)/) || [])[1] || "General";
  if (rl.labelType === "none") return null;
  if (rl.labelType === "value") return positive;
  if (rl.labelType === "custom" && rl.label) {
    const parts = rl.label.split(/(<Value>|<Computation>)/i);
    if (!parts.some(p => /^<value>$/i.test(p))) return lit(parts.map(p => /^<computation>$/i.test(p) ? name : p.replace(/<[^<>]*>/g, "")).join(""));
    return parts.map(p => /^<value>$/i.test(p) ? positive : /^<computation>$/i.test(p) ? lit(name) : lit(p.replace(/<[^<>]*>/g, ""))).join("");
  }
  // automatic / computation: Tableau labels the line with the computation's name
  return rl.formula === "constant" ? positive : lit(name);
}

/**
 * Adds the worksheet's reference lines (table / pane scope) on the value axis: a flat line series on a
 * vertical value axis, or spec.refLines for horizontal bars (the writer draws those across the bars).
 * Lines per cell are Tableau's per-mark ticks – not drawn. Across horizontal bars on two axes a line is left
 * out and named in the conversion report: the writer would need a third axis group, and Excel cannot open that.
 * @param {ChartSpec} spec @param {ChartContext} ctx
 */
function tvApplyReferenceLines(spec, ctx) {
  const { vm, roles } = ctx;
  const fmt = vm.fmt;
  if (!fmt.hasModel || !fmt.referenceLines || spec.boxPlot || !/^(bar|line|area|combo)$/.test(spec.kind) || !spec.categories) return;
  const lines = fmt.referenceLines().filter(r => !r.scope || r.scope === "per-table" || r.scope === "per-pane");
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!lines.length || !valueShelf) return;
  const measures = roles[valueShelf].values.map(v => v.ref).filter(Boolean);
  const drawn = spec.series.filter(s => !s.refLine && s.color !== null);                   // not invisible helper series
  const numbers = list => list.flatMap(s => s.values || []).filter(v => typeof v === "number" && isFinite(v));
  /* the series a line is computed over: its measure's (value-column, else the axis it sits on) –
     a dual-axis chart has one per measure; without a match, the primary axis's series */
  const seriesOf = rl => {
    const ref = rl.field || rl.axis;
    const m = ref && roles[valueShelf].values.find(v => v.ref && tfSameField(v.ref, ref));
    if (m) {
      const label = tvMeasureLabel(vm, m.ci).trim();
      const own = drawn.filter(s => { const n = s.name.trim(); return n === label || n.startsWith(label + " – "); });
      if (own.length) return own;
    }
    return drawn.filter(s => !s.secondary);
  };
  const style = fmt.reflineStyle();
  const hidden = style.lineVisible === false || style.strokeSize === 0;
  const valueFmt = (style.numFmtRaw && tableauToExcelNumFmt(style.numFmtRaw)) || spec.numFmt || "General";
  const twoAxes = spec.series.some(s => s.secondary);
  let left = 0;
  lines.forEach(rl => {
    if (rl.axis && measures.length && !measures.some(m => tfSameField(m, rl.axis))) return;   // another measure's axis
    const data = seriesOf(rl);
    const value = tvReferenceValue(rl, numbers(data));
    if (value === null) return;
    const labelFmt = tvReferenceLabelFormat(rl, valueFmt);
    if (hidden && !labelFmt) return;
    /** @type {RefLineStyle} */
    const line = { value, labelFmt, name: labelFmt ? tfFormatNumber(value, labelFmt) : "Reference line",
                   color: style.strokeColor ? style.strokeColor.slice(2) : "7F7F7F", alpha: style.strokeAlpha ?? 1,
                   width: style.strokeSize || 1, dash: style.dash === "dashed", hidden,
                   font: { color: style.color ? style.color.slice(2) : undefined, bold: style.bold } };
    if (spec.barDir === "bar" && data.some(s => (s.type || spec.kind) === "bar")) {
      if (twoAxes) left++; else (spec.refLines = spec.refLines || []).push(line);
    }
    // drawn on its measure's axis: on a dual-axis chart that may be the secondary one
    else spec.series.push({ name: line.name, type: "line", color: line.color, line: !hidden,
                            ...(data.length && data.every(s => s.secondary) ? { secondary: true } : {}),
                            marker: false, labels: false, values: spec.categories.levels[0].map(() => value), refLine: line });
  });
  if (left) {
    const c = spec.conversion;
    const note = `reference line${left > 1 ? "s" : ""} left out: Excel cannot draw ${left > 1 ? "them" : "one"} across horizontal bars on two axes`;
    spec.conversion = { ...c, strategy: "APPROXIMATE", note: c && c.note ? `${c.note}; ${note}` : note };
  }
}


/* ═══ visual/classify.js ══════════════════════════════════════════════════════════════════════════ */
/* Visual type classification from marks, shelves and encodings. */

function isKPIViewModel(vm) {
  const vis = vm.order.map(i => vm.cols[i]);
  return vm.rows.length === 1 && vis.length >= 1 && vis.length <= 8 && !vis.some(c => c.isHeader);
}

function normalizeVisualToken(value) {
  return String(value == null ? "" : value).toLowerCase().replace(/[\s_-]+/g, "");
}

/* mark types of every marks card in the live visual specification
 * (marksSpecifications[].type: "bar" | "line" | "gantt-bar" | …) */
function visualSpecMarkTokens(spec) {
  if (!spec || typeof spec !== "object") return [];
  const tokens = [];
  const add = value => {
    if (value == null) return;
    if (Array.isArray(value)) value.forEach(add);
    else if (typeof value === "object") {
      [value.markType, value.markClass, value.mark, value.marksType, value.type, value.class].forEach(add);
    } else tokens.push(normalizeVisualToken(value));
  };
  if (Array.isArray(spec.marksSpecifications) && spec.marksSpecifications.length) {
    spec.marksSpecifications.forEach(add);                    // all cards → dual axis / combos are visible
  } else {
    [spec.markType, spec.markClass, spec.mark, spec.type, spec.visualType, spec.marksType, spec.marks].forEach(add);
    if (spec.activeMarksSpecification && typeof spec.activeMarksSpecification === "object") add(spec.activeMarksSpecification);
    (spec.panes || []).forEach(add);
  }
  return [...new Set(tokens.filter(t => t && t !== "automatic"))];
}

/* continuous MEASURE (draws a value axis) vs continuous date/dimension (draws a category axis) */
const TV_DATE_DERIVS = new Set(["yr", "qr", "mn", "wk", "dy", "hr", "mi", "sc", "my", "md", "mdy", "wd",
                                "tyr", "tqr", "tmn", "twk", "tdy", "thr", "tmi", "tsc"]);

function tvIsMeasureRef(model, ref) {
  if (!ref) return false;
  if (ref.name === "Multiple Values") return true;
  if (ref.type !== "qk") return false;
  if (TV_DATE_DERIVS.has(String(ref.deriv || "").toLowerCase())) return false;
  const info = tfFieldInfo(model, ref) || {};
  if (info.bin || /\(bin\)$/i.test(ref.name)) return false;           // a bin draws the histogram's category axis
  return !(info.role === "dimension" && /^date/i.test(info.datatype || ""));
}

/** a measure that is one number (MIN(0), AVG(1)): it only places marks or labels @param {FormatModel | null} model @param {FieldRef} ref */
function tvIsConstantRef(model, ref) {
  const info = ref && tfFieldInfo(model, ref);
  return !!info && info.constant !== undefined && !info.param;
}

/* what the TWB shelves say about the view: value axes, continuous dimension axes, geography,
 * and the evidence for treemaps, histograms, waterfalls and box plots */
function visualShelfShape(sheet, model) {
  if (!sheet) return null;
  const measures = shelf => sheet[shelf].filter(r => tvIsMeasureRef(model, r)).length;
  const contDim = shelf => sheet[shelf].some(r => r.type === "qk" && !tvIsMeasureRef(model, r));
  const all = [...sheet.rows, ...sheet.cols];
  const encoded = channel => sheet.panes.some(p => p.encodings.some(e => e.channel === channel));
  const datatype = r => String((tfFieldInfo(model, r) || {}).datatype || "");
  return {
    rowMeasures: measures("rows"), colMeasures: measures("cols"),
    continuousDimension: contDim("rows") || contDim("cols"),
    geo: all.some(r => /^(latitude|longitude)( \(generated\))?$/i.test(r.name)),
    filled: encoded("geometry"),
    shelfFields: all.filter(r => r.name !== "Measure Names").length,
    size: encoded("size"),
    binned: all.some(r => /\(bin\)$/i.test(r.name) || !!(tfFieldInfo(model, r) || {}).bin),
    countAxis: all.some(r => tvIsMeasureRef(model, r) && /^(cnt|ctd)$/i.test(r.deriv || "")),
    numericDiscreteDim: all.some(r => r.type === "ok" && !TV_DATE_DERIVS.has(String(r.deriv || "").toLowerCase()) &&
                                      /^(integer|real)$/i.test(datatype(r))),
    runningTotal: all.some(r => /^(cum|rsum)$/i.test(r.deriv || "") ||
                                (sheet.runningTotals || []).some(t => tfSameField(t, r))),
    boxPlot: !!sheet.boxPlot,
    // a measure on Color and nothing on Text / Label: Tableau's Automatic mark is a square (heat map)
    colorMeasure: sheet.panes.some(p => p.encodings.some(e => e.channel === "color" && tvIsMeasureRef(model, e.field))),
    text: sheet.panes.some(p => p.labelRuns.length || p.encodings.some(e => e.channel === "text" || e.channel === "label")),
    // a date on the shelves (discrete MONTH(…) or continuous) → Tableau's Automatic mark is a line
    dateDimension: all.some(r => !tvIsMeasureRef(model, r) && (TV_DATE_DERIVS.has(String(r.deriv || "").toLowerCase()) ||
                                                              (/^date/i.test(datatype(r)) && r.type !== "nk"))),
    shelfDims: all.filter(r => r.name !== "Measure Names" && !tvIsMeasureRef(model, r)).length
  };
}

/* the same evidence from the live visual specification (no workbook loaded) */
function liveSpecShape(spec) {
  if (!spec || typeof spec !== "object" || !(Array.isArray(spec.rowFields) || Array.isArray(spec.columnFields))) return null;
  const name = f => String(f == null ? "" : typeof f === "string" ? f : (f.name || f.fieldName || f.caption || ""));
  const shelf = [...(spec.rowFields || []), ...(spec.columnFields || [])].filter(f => !/^measure names$/i.test(name(f)));
  const encodings = (spec.marksSpecifications || []).flatMap(m => (m && m.encodings) || []);
  return {
    shelfFields: shelf.length,
    size: encodings.some(e => /^size$/i.test(String(e.type || e.encodingType || ""))),
    binned: shelf.some(f => /\(bin\)$/i.test(name(f)))
  };
}

/** Tableau's Automatic mark for the view's shelves: what an Automatic marks card draws
 * @param {ReturnType<typeof visualShelfShape>} shape @returns {string} */
function tvAutomaticMark(shape) {
  return !shape ? "" : shape.geo ? "map" : shape.rowMeasures && shape.colMeasures ? "circle"
    : shape.rowMeasures || shape.colMeasures ? (shape.binned ? "bar" : shape.continuousDimension || shape.dateDimension ? "line" : "bar")
    : shape.shelfFields === 0 && shape.size ? "square"                    // empty shelves + Size → treemap
    : shape.colorMeasure && !shape.text ? "square"                         // dimensions only, a measure on Color → heat map
    : "text";
}

/* mark type: live visual spec (current state) → TWB pane marks → Tableau's "Automatic" rules */
/** @param {any} spec live visual spec @param {ViewModel} vm @param {FormatModel | null} model */
function resolveVisualMarks(spec, vm, model) {
  const sheet = vm.fmt && vm.fmt.sheetModel;
  const shape = visualShelfShape(sheet, model);
  let tokens = visualSpecMarkTokens(spec);
  let source = tokens.length ? "live" : "none";
  const automatic = () => tvAutomaticMark(shape);
  if (!tokens.length && sheet) {
    // the marks drawn: every pane with an id when the sheet has several (the id-less "All" pane draws nothing)
    const drawn = sheet.panes.length > 1 && sheet.panes.some(p => p.id) ? sheet.panes.filter(p => p.id) : sheet.panes;
    tokens = [...new Set(drawn.map(p => normalizeVisualToken(p.markClass)).filter(t => t && t !== "automatic"))];
    // Automatic panes beside text / shape panes draw Tableau's automatic mark (bars of a butterfly chart beside its
    // label column): the chart mark comes first, text last
    const auto = automatic();
    if (tokens.length && tokens.every(t => /^(text|shape|circle|square)$/.test(t)) && /^(bar|line)$/.test(auto) &&
        drawn.some(p => !normalizeVisualToken(p.markClass) || normalizeVisualToken(p.markClass) === "automatic")) tokens.push(auto);
    tokens.sort((a, b) => Number(a === "text") - Number(b === "text"));
    if (tokens.length) source = "twb";
  }
  if (!tokens.length && shape) {
    source = "automatic";
    tokens = [automatic()];
  }
  return { tokens, shape, source };
}

/**
 * Rounded bars: a line from a constant (MIN(0)) to each value, one per member, drawn thick with round caps –
 * Line mark, Measure Names on Path, the constant among Measure Values.
 * @param {ViewModel} vm @param {FormatModel | null} model
 */
function isRoundedBar(vm, model) {
  const sheet = vm.fmt && vm.fmt.sheetModel;
  if (!sheet) return false;
  const path = sheet.panes.some(p => /^line$/i.test(p.markClass || "") && p.encodings.some(e => e.channel === "path" && e.field.name === "Measure Names"));
  const mv = [...sheet.rows, ...sheet.cols].some(r => r.name === "Multiple Values");
  const circles = sheet.panes.some(p => /^(circle|shape)$/i.test(p.markClass || ""));
  return path && mv && !circles && sheet.fieldRefs.some(r => tvIsConstantRef(model, r));
}

/**
 * A table built from marks: every axis on Rows / Columns is a constant ("MIN(0)", "0.5") that only places
 * the marks – one column of circles, squares, shapes or bars per axis, each pane showing its label.
 * @param {ViewModel} vm @param {FormatModel | null} model
 */
function isMarkTable(vm, model) {
  const sheet = vm.fmt && vm.fmt.sheetModel;
  if (!sheet) return false;
  const axes = [...sheet.rows, ...sheet.cols].filter(r => tvIsMeasureRef(model, r));
  if (!axes.length || !sheet.panes.some(p => p.labelRuns.length || p.encodings.some(e => e.channel === "text" || e.channel === "label"))) return false;
  // pies / donuts on MIN(0) axes are charts: only cell-like marks make a table
  if (!sheet.panes.every(p => /^(circle|square|shape|bar|text|automatic)?$/i.test(p.markClass || ""))) return false;
  return axes.every(r => {
    const info = tfFieldInfo(model, r);
    if (info && info.constant !== undefined) return true;
    // no formula to go by: the same value on every row of the data
    const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, r));
    if (ci < 0 || vm.rows.length < 2) return false;
    return new Set(vm.rows.map(row => tfDvNum(row[ci])).filter(v => v !== null)).size <= 1;
  });
}

/**
 * @param {any} spec live visual spec @param {ViewModel} vm @param {FormatModel | null} model
 * @param {string[]} [ev] collects the evidence behind the decision (shown in the conversion report)
 * @returns {VisualType}
 */
function classifyVisualType(spec, vm, model, ev = []) {
  const { tokens, shape, source } = resolveVisualMarks(spec, vm, model);
  const live = liveSpecShape(spec);
  const axes = shape ? shape.rowMeasures + shape.colMeasures : null;     // null → unknown (no workbook)
  const token = tokens[0] || "";
  ev.push(tokens.length ? `${tokens.join(" + ")} mark${tokens.length > 1 ? "s" : ""} (${source === "twb" ? "workbook" : source === "live" ? "live view" : "Tableau's automatic mark"})` : "no mark type known");
  if (shape) ev.push(`${shape.rowMeasures} measure axis on Rows, ${shape.colMeasures} on Columns`);
  const kpiOrTable = why => {
    const kpi = isKPIViewModel(vm);
    ev.push(why, kpi ? "one row of a few values: a KPI" : "drawn as cells");
    return kpi ? VISUAL_TYPES.KPI : VISUAL_TYPES.TABLE;
  };
  // generated lat/long on the shelves: a map whatever the mark (circle → symbol map)
  if (shape && shape.geo) {
    ev.push("generated latitude / longitude on the shelves: a map");
    const filled = shape.filled || tokens.some(t => t === "multipolygon" || t === "polygon");
    if (filled) ev.push(shape.filled ? "geometry encoding: filled areas" : "polygon marks");
    return filled ? VISUAL_TYPES.MAP_FILLED : VISUAL_TYPES.MAP;
  }
  if (axes && isMarkTable(vm, model)) return kpiOrTable("every axis is a constant that only places the marks: a table built from marks");
  if (tokens.includes("text") && tokens.every(t => /^(text|shape|circle|square)$/.test(t))) {
    return kpiOrTable("text marks");                                  // KPI tiles / buttons built from marks
  }
  const cartesian = tokens.filter(t => t === "bar" || t === "line" || t === "area");
  if (new Set(cartesian).size > 1) { ev.push("different marks per axis: a combination chart"); return VISUAL_TYPES.COMBO; }
  if (token === "line" && isRoundedBar(vm, model)) {
    ev.push("a thick line from a constant to each value (Path: Measure Names): rounded bars");
    return VISUAL_TYPES.BAR;
  }
  if (token === "ganttbar" || token === "gantt") {
    if (shape && shape.runningTotal) { ev.push("Gantt bars on a running total: a waterfall"); return VISUAL_TYPES.WATERFALL; }
    ev.push("Gantt bars: start on the axis, length from Size");
    return VISUAL_TYPES.GANTT;
  }
  if (shape && shape.boxPlot && axes > 0) { ev.push("box plot reference distribution (Analytics pane)"); return VISUAL_TYPES.BOXPLOT; }
  const byMark = {
    bar: VISUAL_TYPES.BAR, line: VISUAL_TYPES.LINE, area: VISUAL_TYPES.AREA, pie: VISUAL_TYPES.PIE,
    map: VISUAL_TYPES.MAP, multipolygon: VISUAL_TYPES.MAP_FILLED, polygon: VISUAL_TYPES.MAP_FILLED,
    heatmap: VISUAL_TYPES.MAP, density: VISUAL_TYPES.MAP,               // density marks: Tableau-only visual
    vizextension: VISUAL_TYPES.CUSTOM
  };
  if (token === "bar" && axes === 0) return kpiOrTable("bar marks without a measure axis");
  if (token === "bar" && ((shape && (shape.binned || (shape.countAxis && shape.numericDiscreteDim))) || (live && live.binned))) {
    ev.push(shape && shape.binned || live && live.binned ? "a bin field on the category axis: a histogram" : "counts per numeric value: a histogram");
    return VISUAL_TYPES.HISTOGRAM;
  }
  if (byMark[token]) {
    ev.push(token === "vizextension" ? "a viz extension draws this sheet" : /^(heatmap|density)$/.test(token) ? "density marks" : `${token} marks`);
    return byMark[token];
  }
  if (token === "circle" || token === "shape" || token === "square") {
    // nothing on Rows/Columns + Size → Tableau lays the marks out itself: treemap / packed bubbles
    const free = shape ? shape.shelfFields === 0 && shape.size : live ? live.shelfFields === 0 && live.size : false;
    if (free) {
      ev.push("nothing on Rows / Columns and a Size measure: Tableau packs the marks itself");
      return token === "square" ? VISUAL_TYPES.TREEMAP : VISUAL_TYPES.BUBBLE;
    }
    if (axes === null) { ev.push("no workbook: guessed from the mark"); return token === "circle" ? VISUAL_TYPES.SCATTER : VISUAL_TYPES.TABLE; }
    if (shape.rowMeasures && shape.colMeasures) { ev.push("a measure on both axes: a scatter plot"); return VISUAL_TYPES.SCATTER; }
    if (axes > 0) { ev.push("marks along one measure axis: a dot plot"); return VISUAL_TYPES.LINE; }
    if (token === "square") { ev.push("square marks in a grid of dimensions: a heat map"); return VISUAL_TYPES.HEATMAP; }
    return kpiOrTable(`${token} marks in a grid of dimensions`);
  }
  if (!token && vm.kind === "chart") { ev.push("chart without a known mark"); return VISUAL_TYPES.UNKNOWN; }
  return kpiOrTable(token ? `${token} marks` : "no chart marks");
}


/* ═══ charts/model/cartesian.js ═══════════════════════════════════════════════════════════════════ */
/* Bar / line / area / combo / funnel / histogram chart specs. */

/** a constant measure's value (MIN(0) → 0, AVG(1) → 1), else undefined @param {ChartContext} ctx @param {FieldRef | null} ref */
const constantOf = (ctx, ref) => ref && tvIsConstantRef(ctx.model || null, ref) ? tfFieldInfo(ctx.model, ref).constant : undefined;
/** a "% of Total" quick table calculation @param {FieldRef | null} ref */
const isPercentOfTotal = ref => !!ref && (/^pcto$/i.test(ref.deriv || "") || /^pcto:/i.test(ref.inner || ""));
/** a measure its formula negates ("-COUNT([Customers])"): the left wing of a butterfly chart */
const isNegated = (ctx, ref) => { const info = ref && ctx.model ? tfFieldInfo(ctx.model, ref) : null; return !!info && /^\s*-/.test(info.formula || ""); };
/** a number format that shows negative values without their minus sign (both wings of a butterfly read positive) */
const absFmt = fmt => { const p = (String(fmt || "General").match(/^((?:"[^"]*"|[^;])*)/) || [])[1] || "General"; return `${p};${p}`; };

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvCartesianSpecs(ctx) {
  const { vm, roles } = ctx;
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) throw new Error("no continuous measure axis to chart");
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  const horizontal = valueShelf === "cols";
  // a constant 0 on an axis (MIN(0), AVG(0)) only places labels – a butterfly's names – or the start of rounded bars:
  // not data, so not a series
  const zero = m => constantOf(ctx, m.ref) === 0;
  if (roles[valueShelf].values.some(zero) && roles[valueShelf].values.some(m => !zero(m))) {
    const values = roles[valueShelf].values.filter(m => !zero(m));
    return tvCartesianSpecs({ ...ctx, roles: { ...roles, [valueShelf]: { ...roles[valueShelf], values } } });
  }
  const measures = roles[valueShelf].values;
  const mvMode = measures.some(m => m.mv);
  const catDims = [...roles[valueShelf].dims, ...roles[catShelf].dims];
  const color = roles.color;
  const colorCi = color && !color.measureNames && !color.continuous ? color.ci : -1;
  const colorLevel = colorCi >= 0 ? catDims.findIndex(d => d.ci === colorCi) : -1;
  const scale = tvColorScale(vm, roles, ctx.markToken);
  const fieldLabels = vm.fmt.hasModel ? vm.fmt.fieldLabelsShown(catShelf) : true;
  // rounded bars are lines in Tableau (thick, round caps): Excel's bars
  const asBars = t => ctx.roundedBar && t === "line" ? "bar" : t;
  const sheetMark = asBars(ctx.markToken);
  // an Automatic marks card draws the All card's mark, else Tableau's automatic mark for the shelves (a line over
  // dates beside a card of bars – not the bars of the other card)
  const allPane = roles.panes.find(p => !p.id);
  const inherit = (allPane && asBars(tvMarkToken(allPane.markClass))) || asBars((roles.source === "twb" && ctx.autoMark) || ctx.markToken);
  /** @param {FieldRef | null} ref @param {string} fallback @param {number} [index] @param {number} [count] */
  const measureMark = (ref, fallback, index, count) => {
    const own = ref && roles.panes.find(x => x.id && [x.yAxisName, x.xAxisName].some(n => n && tfSameField(tfParseFieldRef(n), ref)));
    if (own && !tvMarkToken(own.markClass)) return inherit;
    return asBars(tvMeasureMark(roles, ref, fallback, index, count));
  };

  // categories (Measure Names alone on the category shelf → one bar per measure)
  const mnCategory = mvMode && !catDims.length && roles.measureNames === catShelf;
  const cats = mnCategory
    ? { count: measures.length, levels: [measures.map(m => tvMeasureLabel(vm, m.ci))], indexOf: () => 0 }
    : tvCategories(vm, catDims.map(d => d.ci), catDims.some(d => d.continuous));
  if (cats.count > TV_MAX_POINTS) throw new Error(`${cats.count} categories – too many for an Excel chart`);
  const categoryTitle = fieldLabels && catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / ") : "";
  const categories = { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [mnCategory ? "Measure Names" : ""], levels: cats.levels };

  // per-category colours: colour = a category level, or a continuous measure
  let pointColors;
  if (scale && colorLevel >= 0) pointColors = cats.levels[colorLevel].map(v => scale(v));
  else if (scale && color && color.continuous && color.ci >= 0) pointColors = tvSum(vm, cats, color.ci).map(v => scale(v));

  // a dimension on Detail / Label that splits the marks of one category and colour: Tableau draws a line (or a dot)
  // per member – each becomes its own series, never their sum
  const splitDims = (() => {
    if (mnCategory) return [];
    const cand = [...roles.detailDims, ...roles.textDims].filter((ci, i, a) => ci >= 0 && a.indexOf(ci) === i && ci !== colorCi &&
                                                                    !catDims.some(d => d.ci === ci));
    if (!cand.length) return [];
    const seen = new Map();
    for (const r of vm.rows) {
      const k = cats.indexOf(r) + "\u0001" + (colorCi >= 0 ? tvText(r[colorCi]) : "");
      const v = cand.map(ci => tvText(r[ci])).join("\u0001");
      if (!seen.has(k)) seen.set(k, v); else if (seen.get(k) !== v) return cand;
    }
    return [];
  })();

  const seriesFor = (m, opts) => {
    const label = tvMeasureLabel(vm, m.ci);
    const labels = tvLabelsOn(roles, m.ref);
    if (splitDims.length && opts.type && opts.type.type === "line") {
      const keyCis = [...(colorCi >= 0 && colorLevel < 0 && !opts.noColorSplit ? [colorCi] : []), ...splitDims];
      const keyOf = r => keyCis.map(ci => tvText(r[ci])).join(" – ");
      const keys = [...new Set(vm.rows.map(keyOf))];
      if (keys.length > TV_MAX_SERIES) {
        throw new Error(`${keys.length} lines (one per ${splitDims.map(ci => tvMeasureLabel(vm, ci)).join(" / ")}) – too many for an Excel chart`);
      }
      // each line in its member's colour (the colour field's value on its marks), else the marks' colour
      const colourOf = k => { const r = colorCi >= 0 && scale ? vm.rows.find(x => keyOf(x) === k) : null; return r ? scale(tvText(r[colorCi])) : null; };
      return keys.map((k, i) => ({
        name: opts.prefix ? `${label} – ${k}` : k, values: tvSum(vm, cats, m.ci, r => keyOf(r) === k),
        color: colourOf(k) || opts.color || tvHex(TABLEAU_10[i % TABLEAU_10.length]), labels, ...opts.type
      }));
    }
    if (colorCi >= 0 && colorLevel < 0 && !opts.noColorSplit) {
      const values = tvColorValues(vm, roles);
      return values.map((v, i) => ({
        name: opts.prefix ? `${label} – ${v}` : v,
        values: tvSum(vm, cats, m.ci, r => tvText(r[colorCi]) === v),
        color: (scale && scale(v)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        labels, ...opts.type
      }));
    }
    return [{ name: label, values: tvSum(vm, cats, m.ci), color: opts.color, labels,
              pointColors: opts.noPointColors ? undefined : pointColors, ...opts.type,
              ...(opts.alpha !== undefined ? { alpha: opts.alpha } : {}) }];
  };

  // Tableau never draws a line across panes: categories of several levels (a year's quarters, a country's months) start
  // a new line at every change of an outer level – each such stretch its own series, the same colour, blank elsewhere
  const paneBreaks = series => {
    const levels = cats.levels;
    const broken = s => !s.refLine && ((s.type === "line" && s.line !== false) || s.type === "area");
    if (levels.length < 2 || !series.some(broken)) return series;
    const groupOf = i => levels.slice(0, -1).map(lv => lv[i]).join("\u0001");
    const ranges = [];
    for (let i = 0; i < cats.count; i++) {
      const g = groupOf(i), last = ranges[ranges.length - 1];
      if (last && last.g === g) last.to = i; else ranges.push({ g, from: i, to: i });
    }
    if (ranges.length < 2 || ranges.length * series.length > TV_MAX_SERIES) return series;
    return series.flatMap(s => !broken(s) ? [s]
      : ranges.map((r, k) => ({ ...s, name: k ? `${s.name} ` + "\u200b".repeat(k) : s.name,
                                values: s.values.map((v, i) => i >= r.from && i <= r.to ? v : null),
                                pointColors: s.pointColors ? s.pointColors.map((c, i) => i >= r.from && i <= r.to ? c : null) : undefined })));
  };

  // stacked bars / areas of a "% of Total" that fills every category: Excel's 100% stacked chart
  const percentOfTotal = measures.length > 0 && measures.every(m => isPercentOfTotal(m.ref));
  const make = (series, extra) => {
    if (series.length > TV_MAX_SERIES) throw new Error(`${series.length} series – too many for an Excel chart`);
    const kinds = new Set(series.map(s => s.type));
    const kind = kinds.size > 1 || series.some(s => s.secondary) ? "combo" : ([...kinds][0] || "bar");
    const stackable = series.every(s => s.type === "bar" || s.type === "area");
    const stacked = stackable && series.length > 1 &&
      ((colorCi >= 0 && colorLevel < 0) || (mvMode && color && color.measureNames && roles.measureNames !== catShelf));
    const totals = Array.from({ length: cats.count }, (_, i) => series.reduce((a, s) => a + ((s.values && s.values[i]) || 0), 0));
    const percent = stacked && percentOfTotal && totals.some(t => t) && totals.every(t => t === 0 || Math.abs(t - 1) < 0.005);
    // Tableau stacks columns in the colour legend's order from the top; Excel stacks its first series at the
    // bottom → reversed (bars across keep the legend's order from the axis outwards)
    const ordered = paneBreaks(stacked && !horizontal ? [...series].reverse() : series);
    const spec = { ...tvBaseSpec(vm), kind, barDir: horizontal ? "bar" : "col", stacked,
                   categories, categoryTitle, legend: series.length > 1, series: ordered, ...extra };
    if (percent) {
      spec.percent = true;
      if (!/%/.test(spec.numFmt || "")) spec.numFmt = "0%";
    }
    return spec;
  };

  const axisIds = [...new Set(measures.map(m => m.axis || 0))].sort((a, b) => a - b);
  const paneOfAxis = k => {
    const refs = roles[valueShelf].axisRefs || [];
    const ref = refs[k];
    if (ref) {
      const nth = refs.slice(0, k).filter(r => tfSameField(r, ref) || (r.name === ref.name && r.name === "Multiple Values")).length;
      const axisName = p => (horizontal ? p.xAxisName : p.yAxisName);
      const hits = roles.panes.filter(p => { const n = axisName(p) && tfParseFieldRef(axisName(p)); return n && (tfSameField(n, ref) || (n.name === "Multiple Values" && ref.name === "Multiple Values")); });
      if (hits[nth]) return hits[nth];
    }
    return roles.panes.find(x => String(x.id) === String(k + 1)) || null;
  };
  const paneMarkOfAxis = k => { const p = paneOfAxis(k); return p ? asBars(tvMarkToken(p.markClass)) || (p.id ? inherit : "") : ""; };
  const paneColor = (m, i) => {
    const p = paneOfAxis(m.axis || 0);
    return (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) ||
           (p ? tvHex(tfArgb(tvPaneRule(p, "mark", "mark-color"))) : null) || tvHex(TABLEAU_10[i % TABLEAU_10.length]);
  };

  // an axis of custom shapes picked by a field (logos, ▲▼ icons: Shape mark + Shape encoding) is
  // decoration Excel cannot draw → chart the other measures instead of a meaningless extra series
  if (axisIds.length > 1) {
    const decorative = axisIds.filter(k => { const p = paneOfAxis(k); return p && tvMarkToken(p.markClass) === "shape" && p.encodings.some(e => e.channel === "shape"); });
    if (decorative.length && decorative.length < axisIds.length) {
      const keep = measures.filter(m => !decorative.includes(m.axis || 0));
      if (keep.length) { measures.splice(0, measures.length, ...keep); return tvCartesianSpecs({ ...ctx, roles: { ...roles, [valueShelf]: { ...roles[valueShelf], values: keep } } }); }
    }
  }

  // butterfly / population pyramid: two measures back to back – the left one negated by its formula, or drawn on a
  // reversed axis – as bars stacked around zero; the left wing's minus signs are hidden
  const uniqueMeasures = measures.filter((m, i) => measures.findIndex(x => x.ci === m.ci) === i);
  if (horizontal && uniqueMeasures.length === 2 && colorLevel < 0 && colorCi < 0) {
    const reversed = m => !m.mv && !!vm.fmt.axisSpace(m.ref, valueShelf, "0").reverse;
    const left = uniqueMeasures.findIndex(m => isNegated(ctx, m.ref) || reversed(m));
    const right = 1 - left;
    if (left >= 0 && !isNegated(ctx, uniqueMeasures[right].ref) && !reversed(uniqueMeasures[right])) {
      const flip = !isNegated(ctx, uniqueMeasures[left].ref);           // a reversed axis holds positive values
      const series = [left, right].map((k, i) => {
        const m = uniqueMeasures[k];
        const values = tvSum(vm, cats, m.ci).map(v => v === null ? null : k === left && flip ? -v : v);
        return { name: tvMeasureLabel(vm, m.ci).replace(/^AGG\(-(.*)\)$/i, "$1"), values, type: /** @type {"bar"} */ ("bar"),
                 color: paneColor(m, i), labels: tvLabelsOn(roles, m.ref), labelNumFmt: absFmt(tvNumFmt(vm, m.ci)) };
      });
      const spec = make(series, { numFmt: absFmt(tvNumFmt(vm, uniqueMeasures[right].ci)), valueTitle: "", gapWidth: 30 });
      return [{ ...spec, stacked: true, legend: true, mirrored: true,
                conversion: { strategy: "CONSTRUCTED", output: "Back-to-back bar chart (butterfly)",
                              note: "Excel has no butterfly chart: two series stacked around zero, the left one negated, its minus signs hidden" } }];
    }
  }

  // barbell (dumbbell): two values per category joined by a line (Measure Names on the line's Path) → markers for
  // both values and Excel's high-low line between them
  const pathMN = roles.panes.some(p => tvMarkToken(p.markClass) === "line" && p.encodings.some(e => e.channel === "path" && e.field.name === "Measure Names"));
  if (pathMN && mvMode && ctx.type !== VISUAL_TYPES.BAR && uniqueMeasures.length === 2 && !mnCategory && horizontal) {
    const linePane = roles.panes.find(p => tvMarkToken(p.markClass) === "line");
    const [a, b] = uniqueMeasures.map(m => tvSum(vm, cats, m.ci));
    const join = { name: "", color: (linePane && tvHex(tfArgb(tvPaneRule(linePane, "mark", "mark-color")))) || "A0A0A0", line: true, lineWidth: 2,
                   values: a.flatMap((v, i) => [v, b[i], null]), cats: a.flatMap((v, i) => v === null || b[i] === null ? [null, null, null] : [i, i, null]) };
    const dots = uniqueMeasures.map((m, i) => ({ name: tvMeasureLabel(vm, m.ci), color: paneColor(m, i), marker: "circle", markerSize: 9,
                                                 values: [a, b][i], cats: [a, b][i].map((v, c) => v === null ? null : c) }));
    const spec = make([{ name: "", values: a.map(() => null), type: "bar", color: null, labels: false }],
                      { numFmt: tvNumFmt(vm, uniqueMeasures[0].ci), valueTitle: "" });
    return [{ ...spec, legend: false, overlay: [join, ...dots],
              conversion: { strategy: "CONSTRUCTED", output: "Dumbbell: dots and joining lines over the bar chart's category bands",
                            note: "Excel has no dumbbell chart: both values as dots, joined by a line, drawn over horizontal category bands" } }];
  }
  if (pathMN && mvMode && ctx.type !== VISUAL_TYPES.BAR && uniqueMeasures.length === 2 && !mnCategory) {
    const linePane = roles.panes.find(p => tvMarkToken(p.markClass) === "line");
    const series = uniqueMeasures.map((m, i) => ({ name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), type: /** @type {"line"} */ ("line"),
      line: false, marker: true, markerSize: 9, color: paneColor(m, i), labels: tvLabelsOn(roles, m.ref) }));
    const spec = make(series, { numFmt: tvNumFmt(vm, uniqueMeasures[0].ci), valueTitle: "" });
    return [{ ...spec, kind: "line", barDir: "col", stacked: false, legend: true,
              hiLowLines: { color: (linePane && tvHex(tfArgb(tvPaneRule(linePane, "mark", "mark-color")))) || "A0A0A0" },
              conversion: horizontal
                ? { strategy: "APPROXIMATE", output: "Line chart: markers joined by high-low lines (vertical)",
                    note: "drawn vertically: Excel joins markers with high-low lines only on vertical line charts" }
                : { strategy: "CONSTRUCTED", output: "Line chart: markers joined by high-low lines",
                    note: "Excel has no dumbbell chart: both values as markers, joined by high-low lines" } }];
  }

  // the same measure on both axes (styling: line + shape at the points, bar + line, Gantt cap on bars)
  // → drawn once: bar / area / line wins, a line gets markers when the other pane draws shapes
  if (!mvMode && axisIds.length > 1 && measures.every(m => m.ci === measures[0].ci)) {
    const toks = axisIds.map(k => paneMarkOfAxis(k) || sheetMark).filter(Boolean);
    if (toks.includes("bar") && toks.some(x => /^(circle|shape)$/.test(x))) {
      const m = measures[0];
      const base = { name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), color: tvMarkColor(roles), pointColors };
      /** @type {ChartSeries[]} */
      const series = [{ ...base, type: "bar", labels: horizontal && tvLabelsOn(roles, m.ref) }];
      if (!horizontal) series.push({ ...base, name: base.name + " ", type: "line", line: false, marker: true, markerSize: 9, labels: tvLabelsOn(roles, m.ref) });
      const spec = make(series, { numFmt: tvNumFmt(vm, m.ci), gapWidth: 300, valueTitle: tvMeasureLabel(vm, m.ci) });
      const heads = horizontal ? [{ name: base.name, color: base.color, marker: "circle", markerSize: 9, values: base.values, cats: base.values.map((v, i) => v === null ? null : i) }] : undefined;
      return [{ ...spec, stacked: false, legend: false, overlay: heads, conversion: lollipopConversion(horizontal) }];
    }
    const primary = ["bar", "area", "line"].find(x => toks.includes(x)) || toks[0] || "bar";
    const type = { ...tvSeriesType(primary) };
    if (primary === "line" && toks.some(x => /^(circle|shape|square)$/.test(x))) type.marker = true;
    const m = measures[0];
    return [withTargets(make(seriesFor(m, { type, color: tvMarkColor(roles) }), { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) }))];
  }

  // a measure on one axis and Measure Values on the other (Sales bars + MV lines) → combo, axis each
  const groupOf = k => measures.filter(m => (m.axis || 0) === k);
  const sameSets = axisIds.every(k => { const a = groupOf(k).map(m => m.ci).sort().join(), b = groupOf(axisIds[0]).map(m => m.ci).sort().join(); return a === b; });
  if (mvMode && axisIds.length > 1 && !sameSets && !mnCategory) {
    const series = [];
    axisIds.forEach((k, ai) => {
      const tok = paneMarkOfAxis(k) || sheetMark;
      groupOf(k).filter((m, i, a) => a.findIndex(x => x.ci === m.ci) === i).forEach(m => series.push({
        name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), labels: tvLabelsOn(roles, m.ref),
        color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, series.length) : null) || tvHex(TABLEAU_10[series.length % TABLEAU_10.length]),
        ...tvSeriesType(tok), secondary: axisIds.length === 2 && ai === 1
      }));
    });
    const first = groupOf(axisIds[0])[0], second = groupOf(axisIds[1])[0];
    const spec = make(series, { numFmt: tvNumFmt(vm, first.ci), valueTitle: groupOf(axisIds[0]).length === 1 ? tvMeasureLabel(vm, first.ci) : "Value",
                                secondaryNumFmt: tvNumFmt(vm, second.ci), secondaryTitle: groupOf(axisIds[1]).length === 1 ? tvMeasureLabel(vm, second.ci) : "Value" });
    spec.stacked = false;
    spec.legend = series.length > 1;
    return [spec];
  }

  // Measure Values: one chart, one series per measure
  if (mvMode) {
    // dual axis "MV + MV": every measure is listed once per axis, and each axis has its own
    // pane (id 1, 2, …) with its own mark → one entry per measure, marks per axis
    const axisCount = Math.max(...measures.map(m => (m.axis || 0) + 1));
    const axisMarks = Array.from({ length: axisCount }, (_, k) => {
      return (axisCount > 1 && paneMarkOfAxis(k)) || measureMark(null, sheetMark);
    });
    const unique = uniqueMeasures;
    const markSet = [...new Set(axisMarks.map(t => tvSeriesType(t).marker && tvSeriesType(t).line === false ? "dot" : tvSeriesType(t).type))];
    const type = tvSeriesType(axisMarks.find(t => tvSeriesType(t).type === "bar") || axisMarks[0]);

    if (!mnCategory && axisCount > 1 && markSet.length > 1) {
      if (markSet.includes("bar") && markSet.includes("dot")) {
        // lollipop: bar + circle on the same values. Vertical → thin bars + markers;
        // horizontal → Excel cannot mix bar and line orientations, so thin bars only
        const colorOf = (m, i) => (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) ||
          (unique.length === 1 ? tvMarkColor(roles) : tvHex(TABLEAU_10[i % TABLEAU_10.length]));
        const series = [];
        unique.forEach((m, i) => {
          const base = { name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), color: colorOf(m, i), pointColors };
          series.push({ ...base, type: "bar", labels: false });
          if (!horizontal) series.push({ ...base, name: base.name + " ", type: "line", line: false, marker: true, markerSize: 9,
                                         labels: tvLabelsOn(roles, m.ref) });
        });
        if (horizontal) series.forEach(s => { s.labels = tvLabelsOn(roles, unique[0].ref); });
        const spec = make(series, { numFmt: tvNumFmt(vm, unique[0].ci), gapWidth: 300,
                                    valueTitle: unique.length === 1 ? tvMeasureLabel(vm, unique[0].ci) : "Value" });
        spec.stacked = false;                      // the circle sits on the bar end, never on top of it
        spec.legend = unique.length > 1;
        // across: each bar's head as a dot over its band
        if (horizontal) spec.overlay = series.map(s => ({ name: s.name, color: s.color, marker: "circle", markerSize: 9, values: s.values, cats: s.values.map((v, i) => v === null ? null : i) }));
        spec.conversion = lollipopConversion(horizontal);
        return [spec];
      }
      // e.g. area + line of the same measure: overlay both, sharing one axis
      const series = [];
      axisMarks.forEach((token, k) => unique.forEach((m, i) => series.push({
        name: tvMeasureLabel(vm, m.ci) + (k ? " " : ""), values: tvSum(vm, cats, m.ci), labels: k === axisCount - 1 && tvLabelsOn(roles, m.ref),
        color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        ...tvSeriesType(token)
      })));
      const spec = make(series, { numFmt: tvNumFmt(vm, unique[0].ci), valueTitle: unique.length === 1 ? tvMeasureLabel(vm, unique[0].ci) : "Value" });
      spec.stacked = false;
      spec.legend = unique.length > 1;
      return [spec];
    }
    measures.splice(0, measures.length, ...unique);

    // funnel: one category per measure, each row carries exactly one of the measures
    // (per-stage calcs) → one value per stage, drawn centred
    const exclusive = catDims.length === 1 && measures.length > 1 && type.type === "bar" &&
      vm.rows.every(r => measures.filter(m => tfDvNum(r[m.ci]) !== null).length <= 1);
    if (exclusive) return tvFunnelSpecFromStages(ctx, cats, catDims[0].ci, measures, horizontal);

    if (mnCategory) {
      const values = measures.map(m => vm.rows.reduce((s, r) => s + (tfDvNum(r[m.ci]) || 0), 0));
      const pc = color && color.measureNames && scale ? measures.map((m, i) => scale(vm.cols[m.ci].name, i)) : undefined;
      return [make([{ name: "Value", values, color: tvMarkColor(roles), pointColors: pc,
                      labels: tvLabelsOn(roles, measures[0].ref), ...type }],
                   { numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "Value" })];
    }
    const series = measures.map((m, i) => ({
      name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), labels: tvLabelsOn(roles, m.ref),
      color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
      ...type
    }));
    return [make(series, { numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "Value" })];
  }

  // one measure
  if (measures.length === 1) {
    const m = measures[0];
    const type = tvSeriesType(measureMark(m.ref, sheetMark));
    // a horizontal dot plot: each mark (detail included) a dot over its category band, one series per colour
    if (horizontal && type.type === "line" && type.line === false && !mnCategory) {
      const groups = colorCi >= 0 && colorLevel < 0 ? tvColorValues(vm, roles) : [null];
      if (groups.length <= TV_MAX_SERIES) {
        const overlay = groups.map((g, i) => {
          const own = vm.rows.filter(r => (g === null || tvText(r[colorCi]) === g) && tfDvNum(r[m.ci]) !== null && cats.indexOf(r) !== undefined);
          const c = g === null ? tvMarkColor(roles) : (scale && scale(g)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]);
          return { name: g === null ? tvMeasureLabel(vm, m.ci) : g, color: c, marker: "circle", markerSize: 7,
                   values: own.map(r => tfDvNum(r[m.ci])), cats: own.map(r => cats.indexOf(r)) };
        });
        const spec = make([{ name: "", values: new Array(cats.count).fill(null), type: "bar", color: null, labels: false }],
                          { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) });
        return [{ ...spec, legend: groups.length > 1, overlay,
                  conversion: { strategy: "CONSTRUCTED", output: "Dot plot: dots over the bar chart's category bands",
                                note: "Excel has no horizontal dot chart: every mark a dot (XY) over its category's band" } }];
      }
    }
    return [withTargets(make(seriesFor(m, { type, color: tvMarkColor(roles) }),
                             { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) }))];
  }

  // several measures: dual axis / combo (different marks) or separate panes (same mark). Tableau's Dual Axis
  // folds the second axis onto the first; Synchronize Axis gives both one scale
  // each axis's own marks card (found by its axis, else its position), an Automatic one drawing what it inherits
  const marks = measures.map((m, i) => {
    const p = paneOfAxis(m.axis || 0);
    return p && p.id ? asBars(tvMarkToken(p.markClass)) || inherit : measureMark(m.ref, sheetMark, i, measures.length);
  });
  const space = (m, i) => vm.fmt.axisSpace(m.ref, valueShelf, measures.slice(0, i).some(x => tfSameField(x.ref, m.ref)) ? "1" : "0");
  const folded = measures.length === 2 && !!space(measures[1], 1).fold;
  const synced = measures.length === 2 && !!space(measures[1], 1).synchronized;
  if (new Set(marks).size > 1 && measures.length <= 4) {
    const series = [];
    const colourPanes = roles.panes.filter(p => p.encodings.some(e => e.channel === "color"));
    measures.forEach((m, i) => {
      // each axis has its own marks card: the colour field splits only the measure whose card holds it
      // (a KPI trend's min / max dots, not its area); the others keep their card's colour and opacity
      const pane = paneOfAxis(m.axis || 0);
      const split = !pane || !colourPanes.length || colourPanes.includes(pane);
      const own = pane ? tvHex(tfArgb(tvPaneRule(pane, "mark", "mark-color"))) : null;
      const opacity = pane ? Number(tvPaneRule(pane, "mark", "mark-transparency")) : NaN;
      seriesFor(m, {
        type: { ...tvSeriesType(marks[i]), secondary: measures.length === 2 && i === 1 && !synced },
        color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) || (!split && own) ||
               tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        prefix: colorCi >= 0 && split, noPointColors: true, noColorSplit: !split,
        alpha: opacity >= 0 && opacity < 255 ? opacity / 255 : undefined
      }).forEach(s => series.push(s));
    });
    const spec = make(series, {
      numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: tvMeasureLabel(vm, measures[0].ci),
      secondaryNumFmt: tvNumFmt(vm, measures[1].ci),
      secondaryTitle: measures.length === 2 ? tvMeasureLabel(vm, measures[1].ci) : undefined
    });
    // without Dual Axis Tableau draws each measure in a pane of its own: overlaid here, each on its own axis
    if (roles.source === "twb" && !folded) {
      spec.conversion = { strategy: "APPROXIMATE", note: "Tableau draws each measure in its own pane (no dual axis); Excel overlays them, each on its own axis" };
    }
    return [spec];
  }
  // the same mark on a dual axis: one chart. Bars are drawn one in the other, as Tableau draws them (bar in bar,
  // a progress bar over its track): the wider marks (the pane's Size) – else the longer values – on the primary axis
  // behind, the other on the secondary axis in front, narrower when Tableau draws it narrower. Synchronized axes share
  // one scale (the secondary axis is pinned to the primary's range and hidden); lines and areas of a synchronized
  // axis simply share the primary axis
  // a colour field on one axis's card splits only that axis's bars (a trend's highlighted min / max months)
  const colourCards = roles.panes.filter(p => p.encodings.some(e => e.channel === "color"));
  const coloured = m => { const p = paneOfAxis(m.axis || 0); return colorCi >= 0 && colorLevel < 0 && (!p || !colourCards.length || colourCards.includes(p)); };
  if (folded && (colorCi < 0 || measures.filter(coloured).length === 1)) {
    const type = tvSeriesType(marks[0]);
    const bars = type.type === "bar";
    const sizeOf = m => { const p = paneOfAxis(m.axis || 0); return (p && Number(tvPaneRule(p, "mark", "size"))) || 1; };
    const reach = m => Math.max(0, ...tvSum(vm, cats, m.ci).filter(v => v !== null).map(Math.abs));
    let order = [0, 1];
    if (bars) order.sort((a, b) => sizeOf(measures[b]) - sizeOf(measures[a]) || reach(measures[b]) - reach(measures[a]) || a - b);
    const [back, front] = order.map(k => measures[k]);
    const series = [back, front].flatMap((m, i) => seriesFor(m, {
      type: { ...type, secondary: i === 1 && (bars || !synced) }, color: paneColor(m, measures.indexOf(m)),
      noColorSplit: !coloured(m), noPointColors: true
    }).map(s => ({ ...s, labels: s.labels && constantOf(ctx, m.ref) === undefined })));
    const spec = make(series, { numFmt: tvNumFmt(vm, back.ci), valueTitle: tvMeasureLabel(vm, back.ci),
                                secondaryNumFmt: tvNumFmt(vm, front.ci), secondaryTitle: synced ? undefined : tvMeasureLabel(vm, front.ci) });
    // the coloured axis's bars stack by colour (one colour per category as a rule); the other draws whole bars
    spec.stacked = colorCi >= 0 && bars;
    spec.legend = false;
    if (bars && synced) { spec.secondarySync = true; spec.secondaryAxisHidden = true; }
    if (bars && sizeOf(front) < sizeOf(back) * 0.9) { spec.gapWidth = 50; spec.secondaryGapWidth = 250; }
    return [spec];
  }
  if (measures.length <= 6) {
    return measures.map((m, i) => withTargets(make(seriesFor(m, { type: tvSeriesType(marks[i]), color: tvMarkColor(roles) }),
      { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci), paneIndex: i, paneCount: measures.length })));
  }
  const series = measures.map((m, i) => ({ name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci),
    color: tvHex(TABLEAU_10[i % TABLEAU_10.length]), labels: tvLabelsOn(roles, m.ref), ...tvSeriesType(marks[i]) }));
  return [make(series, { numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "Value" })];

  /* bullet graph: a reference line per cell from another field (the target) over bars → a tick at each category's
   * target: dash markers over columns; across horizontal bars, vertical ticks drawn by the writer */
  function withTargets(spec) {
    const fmt = vm.fmt;
    const rl = fmt.hasModel && fmt.referenceLines
      ? fmt.referenceLines().find(r => r.scope === "per-cell" && r.field && r.axis && !tfSameField(r.field, r.axis)) : null;
    if (!rl || !spec.series.some(s => (s.type || spec.kind) === "bar") || spec.percent) return spec;
    const ti = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, rl.field));
    if (ti < 0) return spec;
    // the line's value per cell: the target's average over the cell's marks (one mark per category as a rule)
    const sums = tvSum(vm, cats, ti), counts = new Array(cats.count).fill(0);
    vm.rows.forEach(r => { const i = cats.indexOf(r); if (i !== undefined && tfDvNum(r[ti]) !== null) counts[i]++; });
    const targets = sums.map((v, i) => v === null || !counts[i] ? null : v / counts[i]);
    const name = tvMeasureLabel(vm, ti);
    const conversion = { strategy: /** @type {"CONSTRUCTED"} */ ("CONSTRUCTED"), output: `${horizontal ? "Bar" : "Column"} chart with target ticks (bullet graph)`,
                         note: "Excel has no bullet graph: the actual values as bars, each category's target as a tick" };
    if (horizontal) return { ...spec, targets: { name, values: targets, color: "333333" }, conversion };
    const tick = { name, values: targets, type: /** @type {"line"} */ ("line"), line: false, marker: true, markerSymbol: "dash",
                   markerSize: 20, color: "333333", labels: false };
    return { ...spec, kind: "combo", series: [...spec.series, tick], conversion };
  }
}

/**
 * Names what an Excel chart cannot show of the Tableau view (the conversion becomes APPROXIMATE): bars that sum
 * Tableau's segments per detail member, bar widths from Size, a horizontal line or dot plot drawn vertically.
 * @param {ChartSpec} spec @param {ChartContext} ctx
 */
function tvNoteApproximations(spec, ctx) {
  if (spec.conversion || !/^(bar|line|area|combo)$/.test(spec.kind)) return;
  const { vm, roles } = ctx;
  const notes = [];
  const bars = spec.series.some(s => (s.type || spec.kind) === "bar" && s.color !== null);
  const lines = spec.series.filter(s => !s.refLine).every(s => (s.type || spec.kind) === "line");
  if (lines && spec.barDir === "bar") {
    spec.barDir = "col";                                   // Excel draws line charts vertically only
    notes.push("drawn vertically: Excel has no horizontal line or dot chart");
  }
  // bars sized by a measure (bar widths carry data); a size on another card (shapes beside the bars) is not the bars'
  const sizedBars = roles.panes.some(p => {
    const e = p.encodings.find(x => x.channel === "size");
    const mark = tvMarkToken(p.markClass) || (spec.kind === "bar" ? "bar" : "");
    return e && mark === "bar" && tvIsMeasureRef(ctx.model || null, e.field);
  });
  if (bars && sizedBars) notes.push("Tableau sizes these bars by a measure; Excel bars all have one width");
  if (bars) {
    const shelfDims = [...roles.rows.dims, ...roles.cols.dims].map(d => d.ci);
    const colorCi = roles.color && !roles.color.measureNames && roles.color.ci >= 0 ? roles.color.ci : -1;
    const detail = [...roles.detailDims, ...roles.textDims].filter(ci => ci !== colorCi && !shelfDims.includes(ci));
    if (detail.length) {
      const seen = new Map();
      const split = vm.rows.some(r => {
        const k = [...shelfDims, colorCi].map(ci => ci >= 0 ? String(r[ci] && r[ci].formattedValue) : "").join("\u0001");
        const v = detail.map(ci => String(r[ci] && r[ci].formattedValue)).join("\u0001");
        if (!seen.has(k)) { seen.set(k, v); return false; }
        return seen.get(k) !== v;
      });
      if (split) notes.push(`one bar per category: Tableau's segments per ${detail.map(ci => tvMeasureLabel(vm, ci)).join(" / ")} are summed (totals unchanged)`);
    }
  }
  if (notes.length) spec.conversion = { strategy: "APPROXIMATE", note: notes.join("; ") };
}

/** a lollipop's conversion: thin bars, and a dot at each bar's end */
function lollipopConversion(horizontal) {
  return { strategy: /** @type {"CONSTRUCTED"} */ ("CONSTRUCTED"), output: `${horizontal ? "Bar" : "Column"} chart (thin bars) with a dot on each`,
           note: "Excel has no lollipop chart: thin bars, each with a dot at its end" + (horizontal ? " (dots drawn over the category bands)" : "") };
}

/* ── funnel: stages = categories with one value each → stacked bars of an
 *    invisible half-gap + the value, so every stage is centred ─────────────── */
/**
 * @param {ChartContext} ctx @param {ReturnType<typeof tvCategories>} cats @param {number} catCi
 * @param {RoleValue[]} measures @param {boolean} horizontal
 * @returns {ChartSpec[]}
 */
function tvFunnelSpecFromStages(ctx, cats, catCi, measures, horizontal) {
  const { vm, roles } = ctx;
  const values = new Array(cats.count).fill(null);
  vm.rows.forEach(r => {
    const i = cats.indexOf(r);
    if (i === undefined) return;
    measures.forEach(m => { const n = tfDvNum(r[m.ci]); if (n !== null) values[i] = (values[i] || 0) + n; });
  });
  const max = Math.max(0, ...values.filter(v => v !== null));
  const scale = tvColorScale(vm, roles, "bar");
  const pointColors = scale && roles.color && roles.color.ci === catCi
    ? cats.levels[0].map((v, i) => scale(v) || tvHex(TABLEAU_10[i % TABLEAU_10.length])) : undefined;
  return [{
    ...tvBaseSpec(vm), kind: "bar", barDir: horizontal ? "bar" : "col", stacked: true, gapWidth: 15, legend: false,
    gridlines: false, valueAxisHidden: true,
    categories: { names: [tvMeasureLabel(vm, catCi)], levels: cats.levels }, categoryTitle: "",
    numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "",
    series: [
      { name: "Offset", type: "bar", values: values.map(v => v === null ? null : (max - v) / 2), color: null, labels: false },
      { name: tvMeasureLabel(vm, catCi), type: "bar", values, color: tvMarkColor(roles), pointColors, labels: roles.labelRefs.length > 0 }
    ],
    conversion: { strategy: "CONSTRUCTED", output: "Funnel: centred stacked bars",
                  note: "built from stacked bars with an invisible offset, so it opens in every Excel version" }
  }];
}

/* ── histogram: the bar chart with touching bars ─────────────────────────── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvHistogramSpecs(ctx) {
  return tvCartesianSpecs(ctx).map(s => ({ ...s, gapWidth: 0,
    conversion: { strategy: "CONSTRUCTED", output: `Column chart of Tableau's bins (no gaps)`,
                  note: "Tableau's own bins and counts as touching columns – the same bins Tableau draws, not Excel's automatic ones" } }));
}


/* ═══ charts/model/roles.js ═══════════════════════════════════════════════════════════════════════ */
/* Which summary column sits on which shelf / encoding (TWB or live spec). */

/* ── field roles: which summary column sits on which shelf / encoding ───── */
/**
 * @param {ViewModel} vm
 * @param {FormatModel | null} model
 * @param {any} [liveSpec] the live visual specification, when no workbook was loaded
 * @returns {Roles}
 */
function tvRoles(vm, model, liveSpec) {
  const sheet = vm.fmt.sheetModel;
  /** @type {Roles} */
  const roles = { rows: { values: [], dims: [], axisRefs: [] }, cols: { values: [], dims: [], axisRefs: [] }, measureNames: null,
                  color: null, angle: -1, size: -1, textDims: [], detailDims: [],
                  labelRefs: [], labelNames: [], panes: sheet ? sheet.panes : [], source: "none" };
  const addDim = (list, ci) => {
    if (ci >= 0 && (vm.cols[ci].isHeader || !numeric(vm.cols[ci])) && !list.includes(ci)) list.push(ci);
  };
  const pivoted = vm.cols.map((c, i) => c.pivoted ? i : -1).filter(i => i >= 0);
  const numeric = c => /^(int|float|real|integer|number|double)/i.test(String(c.dataType || "")) || c.pivoted;

  if (sheet) {
    roles.source = "twb";
    const colOf = ref => {
      if (!ref) return -1;
      let i = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
      if (i < 0) {
        const names = tfDisplayNames(model, ref);
        i = vm.cols.findIndex(c => names.includes(tfNorm(c.name)));
      }
      return i;
    };
    for (const shelf of ["rows", "cols"]) {
      let axisNo = 0;                              // each measure / Measure Values entry = one axis ("A + B")
      for (const ref of sheet[shelf]) {
        if (ref.name === "Measure Names") { roles.measureNames = shelf; continue; }
        if (ref.name === "Multiple Values") {
          roles[shelf].axisRefs.push(ref);
          const axis = axisNo++;
          pivoted.forEach(ci => roles[shelf].values.push({ ci, ref: vm.cols[ci].ref, mv: true, axis }));
          continue;
        }
        const ci = colOf(ref);
        if (ci < 0) continue;
        if (tvIsMeasureRef(model, ref)) { roles[shelf].axisRefs.push(ref); roles[shelf].values.push({ ci, ref, axis: axisNo++ }); }
        else roles[shelf].dims.push({ ci, ref, continuous: ref.type === "qk" });
      }
    }
    const enc = channel => {
      for (const p of sheet.panes) {
        const e = p.encodings.find(x => x.channel === channel);
        if (e) return e.field;
      }
      return null;
    };
    const colorRef = enc("color");
    if (colorRef) {
      if (colorRef.name === "Measure Names") roles.color = { measureNames: true, ref: colorRef };
      else {
        const ci = colOf(colorRef);
        if (ci >= 0) roles.color = { ci, ref: colorRef, continuous: tvIsMeasureRef(model, colorRef) || colorRef.type === "qk" };
      }
    }
    const angleRef = enc("wedge-size");
    if (angleRef) roles.angle = angleRef.name === "Multiple Values" ? (pivoted[0] ?? -1) : colOf(angleRef);
    const sizeRef = enc("size");
    if (sizeRef) roles.size = colOf(sizeRef);
    sheet.panes.forEach(p => {
      p.encodings.filter(e => e.channel === "text" || e.channel === "label").forEach(e => {
        roles.labelRefs.push(e.field);
        if (!tvIsMeasureRef(model, e.field)) addDim(roles.textDims, colOf(e.field));
      });
      p.encodings.filter(e => e.channel === "lod").forEach(e => addDim(roles.detailDims, colOf(e.field)));
      p.labelRuns.forEach(r => r.refs.forEach(x => roles.labelRefs.push(x)));
    });
    return roles;
  }

  // ── no workbook loaded: use the live visual specification ──
  const spec = liveSpec && typeof liveSpec === "object" ? liveSpec : null;
  const byName = name => {
    const key = tfNorm(name);
    const strip = s => tfNorm(s).replace(/^[a-z]+\((.*)\)$/, "$1");
    let i = vm.cols.findIndex(c => tfNorm(c.name) === key);
    if (i < 0) i = vm.cols.findIndex(c => strip(c.name) === strip(name));
    return i;
  };
  const fieldName = f => f == null ? "" : typeof f === "string" ? f : (f.name || f.fieldName || f.caption || "");
  if (spec && (Array.isArray(spec.rowFields) || Array.isArray(spec.columnFields))) {
    roles.source = "live";
    for (const [shelf, list] of [["rows", spec.rowFields || []], ["cols", spec.columnFields || []]]) {
      list.forEach(f => {
        const name = fieldName(f);
        if (/^measure names$/i.test(name)) { roles.measureNames = shelf; return; }
        if (/^measure values$/i.test(name)) { pivoted.forEach(ci => roles[shelf].values.push({ ci, ref: vm.cols[ci].ref, mv: true })); return; }
        const ci = byName(name);
        if (ci < 0) return;
        if (numeric(vm.cols[ci]) && !vm.cols[ci].isHeader) roles[shelf].values.push({ ci, ref: vm.cols[ci].ref });
        else roles[shelf].dims.push({ ci, ref: vm.cols[ci].ref, continuous: /^date/i.test(String(vm.cols[ci].dataType || "")) });
      });
    }
    const marks = Array.isArray(spec.marksSpecifications) ? spec.marksSpecifications : [];
    roles.liveMarks = marks.map(m => tvMarkToken(m && (m.type || m.markType)));   // one marks card per measure
    const active = marks[spec.activeMarksSpecificationIndex] || marks[0];
    (active && active.encodings || []).forEach(e => {
      const name = fieldName(e.field);
      const type = String(e.type || e.encodingType || "").toLowerCase();
      if (type === "color") {
        if (/^measure names$/i.test(name)) roles.color = { measureNames: true };
        else { const ci = byName(name); if (ci >= 0) roles.color = { ci, continuous: numeric(vm.cols[ci]) && !vm.cols[ci].isHeader }; }
      } else if (type === "angle") roles.angle = byName(name);
      else if (type === "size") roles.size = byName(name);
      else if (type === "label" || type === "text") {
        roles.labelNames.push(tfNorm(name));
        const ci = byName(name);
        if (ci >= 0 && !numeric(vm.cols[ci])) addDim(roles.textDims, ci);
      } else if (type === "detail") addDim(roles.detailDims, byName(name));
    });
    if (roles.rows.values.length || roles.cols.values.length) return roles;
  }

  // ── nothing structural known: dimensions → categories, measures → rows ──
  roles.source = "summary";
  vm.cols.forEach((c, ci) => {
    if (c.isHeader) roles.cols.dims.push({ ci, ref: c.ref, continuous: /^date/i.test(String(c.dataType || "")) });
    else if (numeric(c)) roles.rows.values.push({ ci, ref: c.ref, mv: !!c.pivoted });
  });
  return roles;
}


/* ═══ charts/model/packing.js ═════════════════════════════════════════════════════════════════════ */
/* Circle packing for packed bubbles: largest bubble in the middle, the others around it. */

/* the two centres where a circle of radius r touches circles p and q (gap included) */
function touchingBoth(p, q, r, gap) {
  const r1 = p.r + r + gap, r2 = q.r + r + gap;
  const dx = q.x - p.x, dy = q.y - p.y, d = Math.hypot(dx, dy);
  if (!d || d > r1 + r2 || d < Math.abs(r1 - r2)) return [];
  const a = (r1 * r1 - r2 * r2 + d * d) / (2 * d);
  const h = Math.sqrt(Math.max(0, r1 * r1 - a * a));
  const mx = p.x + a * dx / d, my = p.y + a * dy / d;
  return [{ x: mx + h * dy / d, y: my - h * dx / d }, { x: mx - h * dy / d, y: my + h * dx / d }];
}

/**
 * Greedy packing, biggest first: every circle goes where it touches two placed circles, as close to
 * the centre as it fits. O(n³) – fine for the few hundred marks a packed-bubble view shows.
 * @param {number[]} radii
 * @param {number} [gap] space between circles, in radius units
 * @returns {{ centres: { x: number, y: number }[], box: { w: number, h: number, cx: number, cy: number } }}
 */
function tvPackCircles(radii, gap = 0) {
  const placed = [];
  const centres = new Array(radii.length);
  const order = radii.map((_, i) => i).sort((a, b) => radii[b] - radii[a]);
  const fits = (x, y, r) => placed.every(p => Math.hypot(p.x - x, p.y - y) >= p.r + r + gap - 1e-9);
  for (const i of order) {
    const r = radii[i];
    let best = null;
    if (!placed.length) best = { x: 0, y: 0 };
    else if (placed.length === 1) best = { x: placed[0].r + r + gap, y: 0 };
    else {
      for (let a = 0; a < placed.length; a++) {
        for (let b = a + 1; b < placed.length; b++) {
          for (const c of touchingBoth(placed[a], placed[b], r, gap)) {
            const d = Math.hypot(c.x, c.y);
            if ((!best || d < best.d) && fits(c.x, c.y, r)) best = { ...c, d };
          }
        }
      }
      if (!best) best = { x: Math.max(...placed.map(p => p.x + p.r)) + r + gap, y: 0 };
    }
    placed.push({ x: best.x, y: best.y, r });
    centres[i] = { x: best.x, y: best.y };
  }
  const x0 = Math.min(...placed.map(p => p.x - p.r)), x1 = Math.max(...placed.map(p => p.x + p.r));
  const y0 = Math.min(...placed.map(p => p.y - p.r)), y1 = Math.max(...placed.map(p => p.y + p.r));
  return { centres, box: { w: x1 - x0, h: y1 - y0, cx: (x0 + x1) / 2, cy: (y0 + y1) / 2 } };
}


/* ═══ charts/writer/constants.js ══════════════════════════════════════════════════════════════════ */
/* OOXML namespaces, relationship and content types, axis ids. */

const NS = {
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
  cx: "http://schemas.microsoft.com/office/drawing/2014/chartex",
  cx1: "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006"
};

const REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

const REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

const REL_CHARTEX = "http://schemas.microsoft.com/office/2014/relationships/chartEx";

const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";

const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const CT_CHARTEX = "application/vnd.ms-office.chartex+xml";

const REL_CHARTSTYLE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";

const REL_CHARTCOLORS = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

const CT_CHARTSTYLE = "application/vnd.ms-office.chartstyle+xml";

const CT_CHARTCOLORS = "application/vnd.ms-office.chartcolorstyle+xml";

const EMU_PER_PX = 9525;

const AX = { cat: 50010, val: 50020, cat2: 50030, val2: 50040, x2: 50050, y2: 50060 };

const C15 = "http://schemas.microsoft.com/office/drawing/2012/chart";

/* XLSX package surgery: relationship and content types of the parts injectCharts adds */
const EMPTY_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.rels}"></Relationships>`;


/* ═══ charts/writer/xml-util.js ═══════════════════════════════════════════════════════════════════ */
/* XML escaping, cell references and other small writer helpers. */

/* ── small helpers ─────────────────────────────────────────────────────── */
function esc(s) {
  return String(s == null ? "" : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function colName(i) {
  let s = "", n = i + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
  return s;
}

/* 0-based (col,row) → 'Sheet'!$A$1[:$B$9] */
function cellRef(sheet, c1, r1, c2, r2) {
  const q = "'" + String(sheet).replace(/'/g, "''") + "'!";
  const a = "$" + colName(c1) + "$" + (r1 + 1);
  return q + a + (c2 === undefined ? "" : ":$" + colName(c2) + "$" + (r2 + 1));
}

function hex(c) {
  const s = String(c || "").replace(/^#/, "").toUpperCase();
  return s.length === 8 ? s.slice(2) : s;
}

function num(v) { return typeof v === "number" && isFinite(v) ? v : null; }

function attr(tag, name) {
  const m = tag.match(new RegExp("\\s" + name.replace(":", "\\:") + "=\"([^\"]*)\""));
  return m ? m[1] : null;
}

/* boundaries of each category level: a new group starts where any level ≤ l changes */
function levelStarts(levels) {
  const n = levels.length ? levels[0].length : 0;
  return levels.map((_, l) => Array.from({ length: n }, (__, i) =>
    i === 0 || levels.slice(0, l + 1).some(lv => lv[i] !== lv[i - 1])));
}


/* ═══ charts/writer/pie.js ════════════════════════════════════════════════════════════════════════ */
/* Pie / doughnut chart parts drawn the way Tableau draws them:
 *   pieChartXml     – the chart: slices with their colours (a selected mark outlined), a pinned plot area so
 *                     the diameter is Tableau's, an optional doughnut hole
 *   userShapesXml   – labels as text boxes at Tableau's label positions (Excel data labels wrap and shrink the
 *                     pie), the donut hole's own fill and the centre text
 *   tipShapeXml     – Tableau's tooltip: an invisible wedge over each slice whose hyperlink ScreenTip is the
 *                     tooltip text (Excel's own chart hover text cannot be changed)
 * Knows nothing about Tableau: colours, fonts and texts arrive resolved. */

const PIE_EMU_PER_PX = 9525;                                        // 96 dpi
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_CDR = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";
const REL_USER_SHAPES = NS.r + "/chartUserShapes";
const CT_USER_SHAPES = "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";
const REL_HYPERLINK = NS.r + "/hyperlink";

/* ── text ─────────────────────────────────────────────────────────────── */
/** @typedef {{ name?: string, size?: number, bold?: boolean, italic?: boolean, underline?: boolean, color?: string }} PieFont */
/** @param {PieFont} [f] */
function fontAttrs(f) {
  f = f || {};
  return ` lang="en-US" sz="${Math.round((f.size || 9) * 100)}" b="${f.bold ? 1 : 0}" i="${f.italic ? 1 : 0}"` +
         (f.underline ? ' u="sng"' : "");
}
/** @param {PieFont} [f] */
function fontKids(f) {
  f = f || {};
  return `<a:solidFill><a:srgbClr val="${f.color || "333333"}"/></a:solidFill>` + (f.name ? `<a:latin typeface="${esc(f.name)}"/>` : "");
}
/** lines [[{text, font}]] → one paragraph per line, one run per piece @param {string} [algn] "l" | "ctr" | "r" */
function parasXml(lines, algn) {
  const pPr = algn ? `<a:pPr algn="${algn}"/>` : "";
  return lines.map(runs => {
    const rs = runs.filter(r => r.text !== "").map(r =>
      `<a:r><a:rPr${fontAttrs(r.font)}>${fontKids(r.font)}</a:rPr><a:t>${esc(r.text)}</a:t></a:r>`).join("");
    return rs ? `<a:p>${pPr}${rs}</a:p>` : `<a:p>${pPr}<a:endParaRPr lang="en-US"/></a:p>`;
  }).join("");
}
const richXml = lines => `<c:rich><a:bodyPr/><a:lstStyle/>${parasXml(lines)}</c:rich>`;
const txPrXml = f => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${fontAttrs(f)}>${fontKids(f)}</a:defRPr></a:pPr>` +
                     `<a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
const NO_LINE = `<a:ln><a:noFill/></a:ln>`;
const fillXml = c => c ? `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>` : `<a:noFill/>`;

/* ── data: worksheet ranges (linked, "Select Data" shows them) or literals ── */
const sheetRef = (sheet, range) => `'${String(sheet).replace(/'/g, "''")}'!${range}`;
const strPts = vals => `<c:ptCount val="${vals.length}"/>` + vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("");
const numPts = vals => `<c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
  vals.map((v, i) => `<c:pt idx="${i}"><c:v>${Number(v)}</c:v></c:pt>`).join("");
const pieCatXml = (vals, ref) => ref
  ? `<c:cat><c:strRef><c:f>${esc(sheetRef(ref.sheet, ref.cat))}</c:f><c:strCache>${strPts(vals)}</c:strCache></c:strRef></c:cat>`
  : `<c:cat><c:strLit>${strPts(vals)}</c:strLit></c:cat>`;
const pieValXml = (vals, ref) => ref
  ? `<c:val><c:numRef><c:f>${esc(sheetRef(ref.sheet, ref.val))}</c:f><c:numCache>${numPts(vals)}</c:numCache></c:numRef></c:val>`
  : `<c:val><c:numLit>${numPts(vals)}</c:numLit></c:val>`;

/* ── chart frame ──────────────────────────────────────────────────────── */
function titleXml(title, font) {
  if (!title) return '<c:autoTitleDeleted val="1"/>';
  return `<c:title><c:tx>${richXml([[{ text: title, font }]])}</c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;
}
/** plot: { x, y, w, h } fractions of the chart (inner plot area) – pins the pie's diameter */
function layoutXml(plot) {
  if (!plot) return "<c:layout/>";
  const v = n => Math.max(0, Math.min(1, n)).toFixed(4);
  return `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
    `<c:x val="${v(plot.x)}"/><c:y val="${v(plot.y)}"/><c:w val="${v(plot.w)}"/><c:h val="${v(plot.h)}"/></c:manualLayout></c:layout>`;
}

/**
 * spec: { seriesName, categories[], values[], ref? {sheet, cat, val}, title?, titleFont?, font?, bg? (RRGGBB, null =
 *   see-through), plot?, holeSize? (% of the diameter → doughnut), hasShapes?, points: [{ color, selected }] }
 * @param {any} spec @returns {string}
 */
function pieChartXml(spec) {
  const dPts = (spec.points || []).map((p, i) =>
    `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${fillXml(p.color)}` +
    (p.selected ? `<a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`
                : `<a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>`) +
    `</c:spPr></c:dPt>`).join("");
  const ser = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>${esc(spec.seriesName || "")}</c:v></c:tx>` +
    `${dPts}${pieCatXml(spec.categories, spec.ref)}${pieValXml(spec.values, spec.ref)}</c:ser>`;
  const plot = spec.holeSize
    ? `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/>` +
      `<c:holeSize val="${Math.max(10, Math.min(90, Math.round(spec.holeSize)))}"/></c:doughnutChart>`    // Excel allows 10–90
    : `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`;
  return XML_DECL +
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:roundedCorners val="0"/><c:chart>${titleXml(spec.title, spec.titleFont)}` +
    `<c:plotArea>${layoutXml(spec.plot)}${plot}<c:spPr><a:noFill/>${NO_LINE}</c:spPr></c:plotArea>` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr>${fillXml(spec.bg)}${NO_LINE}</c:spPr>${txPrXml(spec.font)}` +
    (spec.hasShapes ? `<c:userShapes r:id="rId1"/>` : "") + `</c:chartSpace>`;
}

/**
 * Text boxes / shapes drawn on top of a chart (its chartDrawing part), anchored relative to the chart so they
 * move and scale with it; later items draw on top.
 * texts: [{ x, y, w, h (px inside the chart), lines: [[{text, font}]], align: "l"|"ctr"|"r", anchor: "t"|"ctr"|"b",
 *           shape?: preset geometry ("ellipse" …; default a borderless text box), fill?: RRGGBB }]
 * @param {any[]} texts @param {{ w: number, h: number }} size chart size px @returns {string}
 */
function userShapesXml(texts, size) {
  const fx = (v, d) => Math.max(0, Math.min(1, v / d)).toFixed(5);
  const paras = t => t.lines && t.lines.length ? parasXml(t.lines, t.align || "ctr") : `<a:p><a:endParaRPr lang="en-US"/></a:p>`;
  return XML_DECL + `<c:userShapes xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:cdr="${NS_CDR}">` + texts.map((t, i) =>
    `<cdr:relSizeAnchor><cdr:from><cdr:x>${fx(t.x, size.w)}</cdr:x><cdr:y>${fx(t.y, size.h)}</cdr:y></cdr:from>` +
    `<cdr:to><cdr:x>${fx(t.x + t.w, size.w)}</cdr:x><cdr:y>${fx(t.y + t.h, size.h)}</cdr:y></cdr:to>` +
    `<cdr:sp macro="" textlink=""><cdr:nvSpPr><cdr:cNvPr id="${i + 2}" name="${t.shape ? "Shape" : "Label"} ${i + 1}"/>` +
    `<cdr:cNvSpPr${t.shape ? "" : ' txBox="1"'}/></cdr:nvSpPr>` +
    `<cdr:spPr><a:prstGeom prst="${t.shape || "rect"}"><a:avLst/></a:prstGeom>${fillXml(t.fill)}${NO_LINE}</cdr:spPr>` +
    `<cdr:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${t.anchor || "t"}"/><a:lstStyle/>` +
    `${paras(t)}</cdr:txBody></cdr:sp></cdr:relSizeAnchor>`).join("") + `</c:userShapes>`;
}

const xfrmXml = (x, y, w, h) => `<a:off x="${Math.round(x * PIE_EMU_PER_PX)}" y="${Math.round(y * PIE_EMU_PER_PX)}"/>` +
                                `<a:ext cx="${Math.round(w * PIE_EMU_PER_PX)}" cy="${Math.round(h * PIE_EMU_PER_PX)}"/>`;

/**
 * An invisible wedge over a slice: hovering shows its hyperlink ScreenTip (the Tableau tooltip). A 100 %
 * transparent fill still catches the mouse; noFill would not.
 * tip: { x, y, w, h (px inside the chart), prst ("pie" | "blockArc" | "ellipse"), adj: [avLst values], text }
 * @param {any} t @param {number} id @param {string} rid the hyperlink relationship
 */
function tipShapeXml(t, id, rid) {
  const gd = (t.adj || []).map((v, k) => `<a:gd name="adj${k + 1}" fmla="val ${Math.round(v)}"/>`).join("");
  const text = String(t.text).slice(0, 255);                   // Excel's ScreenTip limit
  return `<xdr:sp macro="" textlink=""><xdr:nvSpPr>` +
    `<xdr:cNvPr id="${id}" name="Tooltip ${id}"><a:hlinkClick xmlns:r="${NS.r}" r:id="${rid}" tooltip="${esc(text).replace(/\n/g, "&#xA;")}"/></xdr:cNvPr>` +
    `<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm>${xfrmXml(t.x, t.y, t.w, t.h)}</a:xfrm><a:prstGeom prst="${t.prst}"><a:avLst>${gd}</a:avLst></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill>${NO_LINE}</xdr:spPr></xdr:sp>`;
}

/**
 * The anchor of a chart with tooltip wedges: chart and wedges grouped, so the wedges keep their exact px
 * positions over the slices whatever the cell sizes.
 * @param {any} chart ChartJob (col / row / offsets, widthPx / heightPx, name) @param {string} chartRid
 * @param {{ next: number }} ids shape ids @param {string[]} tipRids one hyperlink relationship per tip
 */
function groupedAnchorXml(chart, chartRid, ids, tipRids) {
  const w = chart.widthPx, h = chart.heightPx;
  const frame = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${ids.next++}" name="${esc(chart.name || "Chart")}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm>${xfrmXml(0, 0, w, h)}</xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${chartRid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
  const group = `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="${ids.next++}" name="${esc((chart.name || "Chart") + " group")}"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>` +
    `<xdr:grpSpPr><a:xfrm>${xfrmXml(0, 0, w, h)}<a:chOff x="0" y="0"/><a:chExt cx="${Math.round(w * PIE_EMU_PER_PX)}" cy="${Math.round(h * PIE_EMU_PER_PX)}"/></a:xfrm></xdr:grpSpPr>` +
    frame + chart.tips.map((t, k) => tipShapeXml(t, ids.next++, tipRids[k])).join("") + `</xdr:grpSp>`;
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.col}</xdr:col><xdr:colOff>${Math.round((chart.colOffPx || 0) * PIE_EMU_PER_PX)}</xdr:colOff>` +
    `<xdr:row>${chart.row}</xdr:row><xdr:rowOff>${Math.round((chart.rowOffPx || 0) * PIE_EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${Math.round(w * PIE_EMU_PER_PX)}" cy="${Math.round(h * PIE_EMU_PER_PX)}"/>${group}<xdr:clientData/></xdr:oneCellAnchor>`;
}


/* ═══ charts/model/pie.js ═════════════════════════════════════════════════════════════════════════ */
/* ══════════════════════════════════════════════════════════════════════════
 * PIES AND DONUTS – Tableau marks → pie model → chart jobs. One summary row = one Tableau mark = one slice.
 *   Angle (wedge-size)  → slice value; no Angle → equal slices (as Tableau draws them)
 *   Color               → slice fill: palette / custom colours, or gradient for a measure
 *   Detail / other dims → more slices, never merged
 *   Label / Tooltip     → Tableau's template filled with Tableau's formatted values
 *   Rows / Cols dims    → one pie per pane
 * A donut is a dual axis (MIN(0) twice): the slices plus a hole layer on top (its own label – usually the
 * total – drawn in the centre); a nested donut is two coloured pies, the top one smaller.
 * ══════════════════════════════════════════════════════════════════════════ */

/* Measured in Tableau 2026.2 (1000×800 dashboard): a view with nothing on Rows/Columns draws each pane as a
 * 160×160 px cell at the zone's top-left (inside the 4 px zone margin), pie centred in it. Size slider →
 * diameter is linear in the TWB value: untouched = 80 px, 1.4613 → 117 px (80 × 1.46); a pie bigger than its
 * cell is clipped to it. */
const PIE_DIAMETER_PX = 80;
const PIE_CELL_PX = 160;
const ZONE_MARGIN = 4;
/** @param {number | null | undefined} sizeValue Size slider */
const pieDiameterPx = sizeValue => sizeValue === null || sizeValue === undefined ? PIE_DIAMETER_PX : PIE_DIAMETER_PX * Math.max(0.01, sizeValue);

/** Excel chart font from Tableau props; measureName = Tableau's own font, for text measuring */
function pieFont(p) {
  const f = tfExcelFont(p || {});
  return { name: f.name, size: f.size, bold: f.bold, italic: f.italic, underline: !!f.underline,
           color: f.color ? f.color.argb.slice(2) : "333333", measureName: p && p.fontName };
}

/* ── label placement ────────────────────────────────────────────────────────
 * Tableau's pie label placement (measured in Tableau 2026.2): a label sits at its slice's outer mid-point, on
 * the side the slice faces – right half: label starts there, left half: ends there; top half: above, bottom
 * half: below – and is pushed back inside the pane where it would stick out. With "Allow labels to overlap
 * other marks" off, a label hitting an already placed one is hidden, biggest slice first. Tableau draws label
 * text at the screen's DPI scale (9 pt = 15 px at 125 %), the pane itself in plain px. */
const LABEL_CHAR_EM = 0.52;            // text width per character when there is no canvas to measure with
const LABEL_LINE_EM = 1.33;            // Tableau label line height (20 px lines for 15 px text)
const LABEL_PAD = 2;
/** @type {CanvasRenderingContext2D | null} */
let measureCtx = null;
function textWidthPx(text, font, px) {
  try {
    measureCtx = measureCtx || document.createElement("canvas").getContext("2d");
    measureCtx.font = `${font.italic ? "italic " : ""}${font.bold ? "bold " : ""}${px}px "${font.measureName || font.name}", Arial, sans-serif`;
    return measureCtx.measureText(text).width;
  } catch (e) {
    return text.length * px * LABEL_CHAR_EM;
  }
}

/**
 * Per point a label box {x, y, w, h, up} in pane px, or null (no label / hidden).
 * @param {any[]} points @param {{ cx: number, cy: number, R: number, w: number, h: number, cull: boolean, textScale: number }} g
 */
function pieLabelLayout(points, g) {
  const EPS = 1e-6;
  const total = points.reduce((s, p) => s + p.value, 0) || 1;
  let acc = 0;
  const boxes = points.map(p => {
    const mid = (acc + p.value / 2) / total * 2 * Math.PI;      // clockwise from 12 o'clock (firstSliceAng 0)
    acc += p.value;
    if (!p.label || !p.label.length) return null;
    const s = Math.sin(mid), c = Math.cos(mid);
    const right = s > EPS || (Math.abs(s) <= EPS && c > 0);    // exactly 12 / 6 o'clock: the way the circle turns
    const up = c > EPS || (Math.abs(c) <= EPS && s < 0);       // exactly 3 o'clock → below, 9 o'clock → above
    let w = 0, h = 0;
    p.label.forEach(line => {
      const px = f => (f.size || 9) * 4 / 3 * g.textScale;
      w = Math.max(w, line.reduce((t, x) => t + textWidthPx(x.text, x.font, px(x.font)), 0));
      h += (line.length ? Math.max(...line.map(x => px(x.font))) : px({})) * LABEL_LINE_EM;
    });
    w += 2 * LABEL_PAD;
    const ax = g.cx + g.R * s, ay = g.cy - g.R * c;
    return { x: Math.max(0, Math.min(g.w - w, right ? ax : ax - w)),
             y: Math.max(0, Math.min(g.h - h, up ? ay - h : ay)), w, h, up };
  });
  if (g.cull) {
    // collision on the text itself (box minus padding), a few px of slack: Tableau keeps labels that only touch
    const SLACK = 4;
    const hit = (a, b) => Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x) - 2 * LABEL_PAD > SLACK &&
                          Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) > SLACK;
    const kept = [];
    points.map((p, i) => i).sort((a, b) => points[b].value - points[a].value || a - b).forEach(i => {
      if (!boxes[i]) return;
      if (kept.some(k => hit(k, boxes[i]))) boxes[i] = null;
      else kept.push(boxes[i]);
    });
  }
  return boxes;
}

/* ── text ──────────────────────────────────────────────────────────────── */
/** a number in the sheet's own Excel format; null when there is no format or no number */
const formatNum = (v, nf) => v === null || v === undefined || !nf ? null : tfFormatNumber(v, nf);

/** a number shown the way Tableau showed a sibling value: same prefix / suffix, decimals, grouping, K/M/% scale */
function formatLike(dv, v) {
  const s = tfDvText(dv), n = tfDvNum(dv);
  const m = s.match(/^(\D*?)(-?\d[\d,]*(?:\.\d+)?)(.*)$/);
  if (!m || !n) return String(v);
  const shown = parseFloat(m[2].replace(/,/g, ""));
  const scale = shown ? n / shown : 1;
  const dec = (m[2].split(".")[1] || "").length;
  return m[1] + (v / scale).toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec,
    useGrouping: m[2].includes(",") || Math.abs(shown) < 1000 }) + m[3];
}

/**
 * runs [{text, props}] with <[field]> placeholders → lines [[{text, props}]]. Tableau usually splits a
 * placeholder over runs ("<", "[ds].[field]", ">"), so the joined text is rendered and every piece keeps the
 * props of the run it came from (a field value takes its field run's props).
 * @param {{ text: string, props?: any }[]} runs @param {(ref: FieldRef) => string} valueOf
 */
function renderTemplate(runs, valueOf) {
  const texts = runs.map(r => String(r.text || "").replace(/Æ\r?\n?/g, "\n"));
  const text = texts.join("");
  const owner = [];                                            // char index → run index
  texts.forEach((t, i) => { for (let k = 0; k < t.length; k++) owner.push(i); });
  const pieces = [];
  const push = (t, ri) => {
    const last = pieces[pieces.length - 1];
    if (last && last.ri === ri) last.text += t; else pieces.push({ text: t, ri });
  };
  const re = /<([^<>]+)>/g;
  let at = 0, m;
  while ((m = re.exec(text))) {
    for (let k = at; k < m.index; k++) push(text[k], owner[k]);
    const inner = m[1];
    let val = m[0];
    if (/^\[/.test(inner)) { const ref = tfParseFieldRef(inner); val = ref ? valueOf(ref) : ""; }
    else if (/^(sheet name|workbook name|page name|page count|page number)$/i.test(inner)) val = "";
    push(val, owner[m.index + 1]);
    at = m.index + m[0].length;
  }
  for (let k = at; k < text.length; k++) push(text[k], owner[k]);
  const lines = [[]];
  pieces.forEach(p => p.text.split(/\r?\n/).forEach((part, i) => {
    if (i) lines.push([]);
    if (part) lines[lines.length - 1].push({ text: part, props: runs[p.ri].props || {} });
  }));
  while (lines.length && !lines[lines.length - 1].length) lines.pop();
  return lines;
}
const plainText = lines => lines.map(line => line.map(x => x.text).join("")).join("\n");

/** summary rows selected in Tableau (getSelectedMarksAsync), matched on their dimension values */
function selectedRows(cols, data, tables) {
  const out = new Set();
  (tables || []).forEach(t => {
    const pairs = (t.columns || []).map((c, ti) => [cols.findIndex(x => x.name === c.fieldName), ti])
      .filter(([si]) => si >= 0 && !cols[si].numeric);
    if (!pairs.length) return;
    const keys = new Set((t.data || []).map(r => pairs.map(([, ti]) => tfDvText(r[ti])).join("\u0001")));
    data.forEach((r, i) => { if (keys.has(pairs.map(([si]) => tfDvText(r[si])).join("\u0001"))) out.add(i); });
  });
  return out;
}

/* ── model ─────────────────────────────────────────────────────────────── */
/**
 * One pie layer: its slices (value, colour, label, tooltip) per pane, its legend.
 * @param {SheetFormatter} fmt formatter scoped to the layer's pane @param {any} summary
 * @param {{ selectedTables?: any[] }} [opts]
 */
function buildPieSpec(fmt, summary, opts = {}) {
  // fieldId ("[ds].[pcto:sum:Qty:qk]") carries the aggregation – a % of total and its base measure can share
  // the fieldName "SUM(Qty)", so it wins over name matching when present
  const cols = (summary.columns || []).map((c, i) => {
    const name = c.fieldName || c.fieldId || `Col${i + 1}`;
    const byId = /:(qk|nk|ok)(:\d+)?\]$/.test(c.fieldId || "") ? tfParseFieldRef(c.fieldId) : null;
    return { name, ref: byId || fmt.matchName(name), numeric: /^(int|float|real|integer|number)$/i.test(String(c.dataType || "")) };
  });
  const colFor = ref => {
    if (!ref) return -1;
    let i = cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
    if (i < 0 && /^multiple values$/i.test(ref.name)) i = cols.findIndex(c => /^measure values$/i.test(c.name));
    if (i < 0 && /^measure names$/i.test(ref.name)) i = cols.findIndex(c => /^measure names$/i.test(c.name));
    return i;
  };

  const angleRef = fmt.encodingRefs(/^wedge-size$/i)[0] || null;
  const vi = colFor(angleRef);
  let marks = (summary.data || []).map((r, i) => ({ r, i }));
  // dual axis: summary data holds every layer's marks (a donut's hole is one extra row with the total);
  // this layer's rows are the ones where its own axis measure has a value
  const axisCol = fmt.axisRefs().length === 1 ? colFor(fmt.axisRefs()[0]) : -1;
  if (axisCol >= 0) marks = marks.filter(({ r }) => !tfIsNull(r[axisCol]));
  // both layers on one axis measure (MIN(0) twice): the other layer's rows are the ones without this layer's
  // colour member – a donut's hole total comes back as Country = Null
  const enc0 = fmt.colorEncoding(), ci0 = enc0 && !enc0.continuous ? colFor(enc0.ref) : -1;
  if (axisCol >= 0 && ci0 >= 0 && marks.some(({ r }) => !tfIsNull(r[ci0]))) marks = marks.filter(({ r }) => !tfIsNull(r[ci0]));
  if (vi >= 0) marks = marks.filter(({ r }) => { const v = tfDvNum(r[vi]); return v !== null && v > 0; });   // Tableau drops null / ≤ 0 wedges
  if (!marks.length) return null;

  // a field's value as Tableau shows it; % of total computed when the table calc isn't in the summary data
  const totals = {};
  const valueOf = r => ref => {
    const ci = colFor(ref);
    if (ci >= 0) return formatNum(tfDvNum(r[ci]), fmt.numFmtFor(ref)) || tfDvText(r[ci]);
    if (!/^pcto/i.test(ref.deriv || "")) return "";
    const base = tfNorm(ref.name.split(":").pop());
    const bi = cols.findIndex(c => c.numeric && c.ref && tfNorm(c.ref.name) === base);
    if (bi < 0) return "";
    if (totals[bi] === undefined) totals[bi] = marks.reduce((s, m) => s + (tfDvNum(m.r[bi]) || 0), 0);
    // % of the whole table (Tableau's default scope); other "compute using" scopes are not replicated
    const nf = fmt.numFmtFor(ref);
    const dec = nf ? ((nf.match(/\.(0+)/) || [, ""])[1].length) : 2;
    return totals[bi] ? (100 * (tfDvNum(r[bi]) || 0) / totals[bi]).toFixed(dec) + "%" : "";
  };

  // colour + legend
  const enc = fmt.colorEncoding();
  const ci = enc ? colFor(enc.ref) : -1;
  let colorAt = r => fmt.markColor();                          // no colour field: Marks → Color, else Tableau blue
  let legend = null;
  if (ci >= 0) {
    const pick = dv => tfIsNull(dv) ? null
      : enc.continuous ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : tfDvText(dv);
    const vals = marks.map(({ r }) => pick(r[ci]));
    const scale = tfBuildColorScale(fmt, enc, vals);
    if (scale) {
      const rgb = v => { const a = scale(v); return a ? a.slice(2) : null; };
      colorAt = r => rgb(pick(r[ci]));
      const title = fmt.captionFor(enc.ref, cols[ci].name);
      if (enc.continuous) {
        const nums = vals.filter(v => typeof v === "number");
        const lo = Math.min(...nums), hi = Math.max(...nums);
        legend = { title, items: [{ text: String(lo), color: rgb(lo) }, { text: String(hi), color: rgb(hi) }] };   // gradient: its two ends
      } else {
        const uniq = [...new Map(marks.map(({ r }) => [tfDvText(r[ci]), r[ci]])).values()].sort(tfNaturalCompare);
        legend = { title, items: uniq.map(dv => ({ text: tfDvText(dv), color: rgb(tfDvText(dv)) })) };
      }
    }
  }

  // panes: discrete fields on Rows / Columns → one pie each
  const hdr = shelf => cols.map((c, i) => c.ref && fmt.isHeaderField(c.ref) && fmt.shelfOf(c.ref) === shelf ? i : -1).filter(i => i >= 0);
  const rowHdr = hdr("rows"), colHdr = hdr("cols");
  const keyOf = (r, idx) => idx.map(i => tfDvText(r[i])).join(" | ");
  const hdrSet = new Set([...rowHdr, ...colHdr]);
  const dimIdx = cols.map((c, i) => !c.numeric && !hdrSet.has(i) && !/^measure values$/i.test(c.name) ? i : -1).filter(i => i >= 0);

  const valueCaption = vi >= 0 ? fmt.captionFor(angleRef, cols[vi].name) : "Marks";
  const showLabels = fmt.markLabelsShown();
  const labelTpl = fmt.labelRunsTemplate(), tipTpl = fmt.tooltipTemplate();
  const selected = selectedRows(cols, summary.data || [], opts.selectedTables);
  const points = marks.map(({ r, i }) => ({
    row: i,
    rowKey: keyOf(r, rowHdr), colKey: keyOf(r, colHdr),
    category: dimIdx.map(k => tfDvText(r[k])).filter(Boolean).join(", ") || valueCaption,
    value: vi >= 0 ? tfDvNum(r[vi]) : 1,
    color: colorAt(r),
    selected: selected.has(i),
    label: showLabels ? renderTemplate(labelTpl, valueOf(r))
      .map(line => line.map(x => ({ text: x.text, font: pieFont(fmt.labelFont(x.props)) }))) : null,
    tooltip: plainText(renderTemplate(tipTpl, valueOf(r)))
  }));
  const markSize = fmt.markSize();
  tfLog(`Pie: ${points.length} marks, angle=${vi >= 0 ? cols[vi].name : "none"}, colour=${ci >= 0 ? cols[ci].name : "none"}, ` +
        `labels=${showLabels}, selected=${selected.size}, size slider=${markSize === null ? "default" : markSize}`);
  const axis = shelf => ((fmt.sheetModel || {})[shelf] || []).some(r => r.type === "qk");
  return {
    columns: cols, data: summary.data || [], valueCaption, points, legend, markSize,
    cullLabels: fmt.markLabelsCulled(),
    rowKeys: [...new Set(points.map(p => p.rowKey))], colKeys: [...new Set(points.map(p => p.colKey))],
    stretch: { w: axis("cols"), h: axis("rows") },             // a continuous axis (a donut's MIN(0)) → the pane fills the zone
    paneOf: r => keyOf(r, rowHdr) + "\u0001" + keyOf(r, colHdr),   // the pane of any summary row (e.g. a hole row)
    colorRef: enc ? enc.ref : null,
    font: pieFont(fmt.labelFont({})),
    /** @type {any} */ hole: null, /** @type {any} */ inner: null, /** @type {any} */ centerText: null
  };
}

/**
 * Donut hole = the top layer of a dual axis (nothing on its card splitting it), drawn as Excel's doughnut hole.
 * Its label lives at the hole's level of detail (usually the grand total): Tableau returns one summary row per
 * pane for it, which carries the exact value – COUNTD and AVG too; without it the measures are re-aggregated
 * over the slices. @param {SheetFormatter} fmt formatter scoped to the hole's pane @param {any} pie
 */
function buildHoleSpec(fmt, pie) {
  const shown = fmt.markLabelsShown(), tpl = fmt.labelRunsTemplate();
  const ax = fmt.axisRefs()[0];
  const ai = ax ? pie.columns.findIndex(c => c.ref && tfSameField(c.ref, ax)) : -1;
  // the hole's own row: one the slices don't use (on one shared axis measure every row has an axis value)
  const sliceRows = new Set(pie.points.map(p => p.row));
  const axisRows = ai >= 0 ? pie.data.filter(r => !tfIsNull(r[ai])) : [];
  const spare = axisRows.filter(r => !sliceRows.has(pie.data.indexOf(r)));
  const own = new Map((spare.length ? spare : axisRows).map(r => [pie.paneOf(r), r]));
  const valueOf = pts => ref => {
    const k = pie.columns.findIndex(c => c.ref && tfSameField(c.ref, ref));
    if (k < 0) return "";
    const hit = own.get(pts[0].rowKey + "\u0001" + pts[0].colKey);
    if (hit) return formatNum(tfDvNum(hit[k]), fmt.numFmtFor(ref)) || tfDvText(hit[k]);
    const dvs = pts.map(p => pie.data[p.row][k]);
    if (!pie.columns[k].numeric) { const u = [...new Set(dvs.map(tfDvText))]; return u.length === 1 ? u[0] : "*"; }
    const nums = dvs.map(tfDvNum).filter(v => v !== null);
    if (!nums.length) return "";
    const d = (ref.deriv || "").toLowerCase();
    // AVG / CNTD re-aggregated as a plain mean / sum of the slices, not from the raw rows
    const sum = nums.reduce((a, b) => a + b, 0);
    const v = d === "min" ? Math.min(...nums) : d === "max" ? Math.max(...nums) : d === "avg" ? sum / nums.length : sum;
    return formatNum(v, fmt.numFmtFor(ref)) || formatLike(dvs.find(x => tfDvNum(x) !== null), v);
  };
  // hover text: the Tooltip editor if set; else, like Tableau, "Caption: value" for every field the hole's own
  // row has a value for, dimensions first, then by field name
  const tipFor = pts => {
    const hit = pts.length && own.get(pts[0].rowKey + "\u0001" + pts[0].colKey);
    if (!hit || fmt.hasCustomTooltip()) return plainText(renderTemplate(fmt.tooltipTemplate(), valueOf(pts)));
    return pie.columns.map((c, k) => ({ c, k })).filter(({ c, k }) => c.ref && !tfIsNull(hit[k]))
      .sort((a, b) => Number(a.c.numeric) - Number(b.c.numeric) || (a.c.numeric ? a.c.ref.inner.localeCompare(b.c.ref.inner) : 0))
      .map(({ c }) => fmt.captionFor(c.ref, c.name) + ": " + valueOf(pts)(c.ref)).join("\n");
  };
  return {
    markSize: fmt.markSize(),
    tipFor,
    color: fmt.markColor(),                                    // usually white = the background
    labelFor: pts => shown ? renderTemplate(tpl, valueOf(pts))
      .map(line => line.map(x => ({ text: x.text, font: pieFont(fmt.labelFont(x.props)) }))) : null
  };
}

/**
 * Floating centre text: a dashboard text box, or another sheet (e.g. a total), whose centre lies inside a single
 * donut's hole (or a pie's circle) is drawn there with its Tableau text and fonts.
 * g: { stretch, D, holeD, cell, margin } px; dataOf(sheetName) → that sheet's summary data.
 * @returns {{ name: string | null, id: string, lines: any[] } | null} name: the sheet drawn there (null = a text box)
 */
function pieCenterText(model, dashName, sheetName, g, dataOf) {
  const dash = model && model.dashboards && model.dashboards[dashName];
  const z = dash && dash.width && dash.height && dash.zones.find(s => s.type === "worksheet" && s.name === sheetName);
  if (!z) return null;
  const px = (v, w) => v / 100000 * (w ? dash.width : dash.height);
  const W = px(z.w, 1), H = px(z.h, 0);
  const cx = px(z.x, 1) + g.margin + (g.stretch.w ? W - 2 * g.margin : Math.min(g.cell, W - 2 * g.margin)) / 2;
  const cy = px(z.y, 0) + (g.stretch.h ? H - g.margin : Math.min(g.cell, H - g.margin)) / 2;
  const R = (g.holeD || g.D) / 2;
  const hit = dash.zones.find(c => c !== z && !c.hidden && c.w && c.h && px(c.w, 1) < W && px(c.h, 0) < H &&
    (c.type === "text" ? !!(c.runs && tfZoneText(c)) : c.type === "worksheet" && c.name !== sheetName) &&
    Math.hypot(px(c.x + c.w / 2, 1) - cx, px(c.y + c.h / 2, 0) - cy) <= R);
  if (!hit) return null;
  if (hit.type === "text") return { name: null, id: hit.id, lines: renderTemplate(hit.runs, () => "")
    .map(line => line.map(x => ({ text: x.text, font: pieFont(tfMerge(TABLEAU_DEFAULTS.worksheet, x.props)) }))) };
  const sum = dataOf(hit.name);
  if (!sum || !sum.data || !sum.data.length) return null;
  const f = createSheetFormatter(model, hit.name);
  const refs = (sum.columns || []).map(c => /:(qk|nk|ok)(:\d+)?\]$/.test(c.fieldId || "") ? tfParseFieldRef(c.fieldId) : f.matchName(c.fieldName));
  const row = sum.data[0];
  const valueOf = ref => {
    const k = refs.findIndex(r => r && tfSameField(r, ref));
    return k < 0 ? "" : formatNum(tfDvNum(row[k]), f.numFmtFor(ref)) || tfDvText(row[k]);
  };
  const tpl = f.labelRunsTemplate();
  let lines = tpl.length ? renderTemplate(tpl, valueOf) : [];
  if (!lines.some(l => l.length)) lines = [[{ text: row.map(tfDvText).filter(Boolean).join(" "), props: {} }]];   // no label → its values
  return { name: hit.name, id: hit.id, lines: lines.map(line => line.map(x => ({ text: x.text, font: pieFont(f.labelFont(x.props)) }))) };
}

/**
 * The whole pie model of a sheet, or null when it isn't a pie / donut Tableau draws with one or two layers.
 * @param {FormatModel} model @param {string} sheetName @param {any} summary
 * @param {{ dashboardName: string, selectedTables?: any[], dataOf: (name: string) => any }} opts
 */
function buildPieModel(model, sheetName, summary, opts) {
  const fmt = createSheetFormatter(model, sheetName);
  const layers = fmt.pieLayers();
  if (!layers) return null;
  const scoped = pane => createSheetFormatter(model, sheetName, [pane]);
  const pie = buildPieSpec(scoped(layers.outer), summary, opts);
  if (!pie) return null;
  if (layers.hole) pie.hole = buildHoleSpec(scoped(layers.hole), pie);
  if (layers.inner) pie.inner = buildPieSpec(scoped(layers.inner), summary, opts);
  const fit = fmt.fitMode(opts.dashboardName);                 // Entire View / Fit Width / Fit Height → the pane fills the zone
  if (/^(entire-view|fit-width)$/.test(fit || "")) pie.stretch.w = true;
  if (/^(entire-view|fit-height)$/.test(fit || "")) pie.stretch.h = true;
  if (pie.rowKeys.length * pie.colKeys.length === 1) {
    pie.centerText = pieCenterText(model, opts.dashboardName, sheetName,
      { stretch: pie.stretch, D: pieDiameterPx(pie.markSize), holeD: pie.hole ? pieDiameterPx(pie.hole.markSize) : 0,
        cell: PIE_CELL_PX, margin: ZONE_MARGIN }, opts.dataOf);
  }
  return pie;
}

/* ── output ───────────────────────────────────────────────────────────── */
/**
 * The slices on the chart data sheet – every summary column, the category / value Excel plots and Tableau's
 * label and tooltip text – so each chart is linked to real cells. Returns per pane the ranges its series uses.
 * @param {import("exceljs").Worksheet} ds @param {number} startRow 1-based @param {any} pie @param {string} name
 * @returns {{ panes: Map<string, { ref: any, innerRef: any }>, nextRow: number }}
 */
function writePieData(ds, startRow, pie, name) {
  const nCols = pie.columns.length;
  ds.getCell(startRow, 1).value = name;
  ds.getCell(startRow, 1).font = { bold: true };
  // cell by cell: Row.values only takes an Array of ExcelJS's own realm
  const putRow = (r, values) => values.forEach((v, k) => { ds.getCell(r, k + 1).value = v; });
  putRow(startRow + 1, [...pie.columns.map(c => c.name), "Excel category", "Excel value", "Tableau label", "Tableau tooltip"]);
  ds.getRow(startRow + 1).font = { bold: true };
  const catCol = ds.getColumn(nCols + 1).letter, valCol = ds.getColumn(nCols + 2).letter;
  let row = startRow + 2;
  const writeRows = pts => {
    const first = row;
    pts.forEach(p => {
      putRow(row++, [
        ...pie.data[p.row].map((dv, k) => pie.columns[k].numeric ? tfDvNum(dv) : tfDvText(dv)),
        p.category, p.value, p.label ? plainText(p.label) : "", p.tooltip]);
    });
    return { sheet: ds.name, cat: `$${catCol}$${first}:$${catCol}$${row - 1}`, val: `$${valCol}$${first}:$${valCol}$${row - 1}` };
  };
  const panes = new Map();
  pie.rowKeys.forEach(rk => pie.colKeys.forEach(ck => {
    const pts = pie.points.filter(p => p.rowKey === rk && p.colKey === ck);
    if (!pts.length) return;
    const ref = writeRows(pts);
    const innerPts = pie.inner ? pie.inner.points.filter(p => p.rowKey === rk && p.colKey === ck) : [];
    panes.set(rk + "\u0001" + ck, { ref, innerRef: innerPts.length ? writeRows(innerPts) : null });
  }));
  return { panes, nextRow: row + 1 };
}

/**
 * Chart jobs for a pie sheet drawn in a block of w × h px: one exact-size chart per pane (a nested donut's inner
 * pie is a second, see-through chart over the ring's hole), Tableau's labels as text boxes, its tooltip on an
 * invisible wedge over each slice. Positions are px from the block's top-left.
 * @param {any} pie @param {{ w: number, h: number }} block @param {Map<string, any>} paneRefs (writePieData)
 * @param {{ name: string, bg: string, link: string }} o bg RRGGBB; link: where clicking a slice goes
 * @returns {{ x: number, y: number, w: number, h: number, xml: string, shapes: string | null, tips: any[], name: string }[]}
 */
function buildPieCharts(pie, block, paneRefs, o) {
  const M = ZONE_MARGIN;
  // Tableau's cell per pane, from the block's top-left inside the zone margin
  const fitW = (block.w - 2 * M) / pie.colKeys.length, fitH = (block.h - M) / pie.rowKeys.length;
  const pw = pie.stretch.w ? fitW : Math.min(PIE_CELL_PX, fitW);
  const ph = pie.stretch.h ? fitH : Math.min(PIE_CELL_PX, fitH);
  const textScale = FORMAT_CONFIG.tableauTextScale || (typeof window !== "undefined" && window.devicePixelRatio) || 1;
  // Tableau clips an oversized pie to its cell; Excel can't clip, so it fills the cell instead – and the hole /
  // inner ring shrink with it, so the ring keeps Tableau's proportions
  const full = pieDiameterPx(pie.markSize);
  const D = Math.min(full, pw, ph), k = D / full;
  const holeD = pie.hole ? Math.min(pieDiameterPx(pie.hole.markSize) * k, D) : 0;
  const innerD = pie.inner ? Math.min(pieDiameterPx(pie.inner.markSize) * k, D) : 0;
  const single = pie.rowKeys.length * pie.colKeys.length === 1;
  const labelsOf = (pts, R, cull) => {
    const boxes = pieLabelLayout(pts, { cx: pw / 2, cy: ph / 2, R, w: pw, h: ph, cull, textScale });
    return pts.map((p, k) => boxes[k] && { ...boxes[k], lines: p.label, anchor: boxes[k].up ? "b" : "t" }).filter(Boolean);
  };
  // Excel draws slice 1 from 12 o'clock clockwise with |value|; DrawingML angles are 60000ths of a degree
  // clockwise from 3 o'clock
  const tipsOf = (pts, d, hole, size) => {
    const x = (size.w - d) / 2, y = (size.h - d) / 2;
    const total = pts.reduce((s, p) => s + Math.abs(p.value || 0), 0);
    const ang = a => (((a % 360) + 360) % 360) * 60000;
    let a = -90;
    return pts.map(p => {
      const st = a;
      a += total ? 360 * Math.abs(p.value || 0) / total : 0;
      if (a === st || !p.tooltip) return null;
      return { x, y, w: d, h: d, prst: hole ? "blockArc" : "pie",
               adj: hole ? [ang(st), ang(a), 50000 * (1 - hole / d)] : [ang(st), ang(a)], text: p.tooltip, link: o.link };
    }).filter(Boolean);
  };
  const out = [];
  pie.rowKeys.forEach((rk, ri) => pie.colKeys.forEach((ck, cj) => {
    const pts = pie.points.filter(p => p.rowKey === rk && p.colKey === ck);
    const refs = paneRefs.get(rk + "\u0001" + ck);
    if (!pts.length || !refs) return;
    const innerPts = innerD && refs.innerRef ? pie.inner.points.filter(p => p.rowKey === rk && p.colKey === ck) : [];
    // labels of the two rings are culled per ring, not against each other
    const texts = [...labelsOf(pts, D / 2, pie.cullLabels), ...(innerPts.length ? labelsOf(innerPts, innerD / 2, pie.inner.cullLabels) : [])];
    // Excel's hole is see-through (it shows the chart background): a hole in another colour is a filled circle
    // under the labels
    if (holeD && pie.hole.color.toUpperCase() !== o.bg.toUpperCase())
      texts.unshift({ x: (pw - holeD) / 2, y: (ph - holeD) / 2, w: holeD, h: holeD, shape: "ellipse", fill: pie.hole.color, lines: [] });
    // centre: a floating text box / sheet over the hole wins over the hole's own label (Tableau draws it on top)
    const centre = (single && pie.centerText && pie.centerText.lines) || (pie.hole && pie.hole.labelFor(pts));
    if (centre && centre.length && !innerPts.length) texts.push({ x: 0, y: 0, w: pw, h: ph, lines: centre, anchor: "ctr" });
    const x = M + cj * pw, y = ri * ph;
    const spec = {
      seriesName: pie.valueCaption, categories: pts.map(p => p.category), values: pts.map(p => p.value), ref: refs.ref,
      title: [rk, ck].filter(Boolean).join(" | ") || null, titleFont: pie.font, font: pie.font, bg: o.bg,
      plot: { x: (1 - D / pw) / 2, y: (1 - D / ph) / 2, w: D / pw, h: D / ph },
      holeSize: holeD ? 100 * holeD / D : innerPts.length ? 100 * innerD / D : null,
      points: pts.map(p => ({ color: p.color, selected: p.selected })), hasShapes: texts.length > 0
    };
    const holeTip = holeD && pie.hole.tipFor(pts);
    out.push({ x, y, w: pw, h: ph, name: o.name, xml: pieChartXml(spec),
               shapes: texts.length ? userShapesXml(texts, { w: pw, h: ph }) : null,
               tips: [...tipsOf(pts, D, holeD || (innerPts.length ? innerD : 0), { w: pw, h: ph }),
                      ...(holeTip ? [{ x: (pw - holeD) / 2, y: (ph - holeD) / 2, w: holeD, h: holeD, prst: "ellipse", text: holeTip, link: o.link }] : [])] });
    if (!innerPts.length) return;
    // nested donut: the inner pie is its own see-through chart exactly over the ring's hole (drawn after = on top);
    // its labels stay in the ring chart so they can sit over the ring
    const innerTexts = centre && centre.length ? [{ x: 0, y: 0, w: innerD, h: innerD, lines: centre, anchor: "ctr" }] : [];
    const innerSpec = {
      seriesName: pie.inner.valueCaption, categories: innerPts.map(p => p.category), values: innerPts.map(p => p.value),
      ref: refs.innerRef, title: null, font: pie.font, bg: null, plot: { x: 0, y: 0, w: 1, h: 1 }, hasShapes: innerTexts.length > 0,
      points: innerPts.map(p => ({ color: p.color, selected: p.selected }))
    };
    out.push({ x: x + (pw - innerD) / 2, y: y + (ph - innerD) / 2, w: innerD, h: innerD, name: o.name + " (inner)",
               xml: pieChartXml(innerSpec), shapes: innerTexts.length ? userShapesXml(innerTexts, { w: innerD, h: innerD }) : null,
               tips: tipsOf(innerPts, innerD, 0, { w: innerD, h: innerD }) });
  }));
  return out;
}


/* ═══ charts/model/specialized.js ═════════════════════════════════════════════════════════════════ */
/* Pie, scatter, waterfall, box plot, Gantt, treemap, symbol map and packed bubble chart specs. */

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvPieSpec(ctx) {
  const { vm, roles } = ctx;
  let catCi = roles.color && !roles.color.measureNames && !roles.color.continuous ? roles.color.ci : -1;
  if (catCi < 0) {
    const dims = [...roles.rows.dims, ...roles.cols.dims];
    catCi = dims.length ? dims[0].ci : vm.cols.findIndex(c => c.isHeader);
  }
  let angleCi = roles.angle;
  if (angleCi < 0) {
    const v = [...roles.rows.values, ...roles.cols.values][0];
    angleCi = v ? v.ci : vm.cols.findIndex(c => !c.isHeader && c.name !== (vm.cols[catCi] || {}).name);
  }
  const scale = tvColorScale(vm, roles, "pie");
  const spec = /** @type {ChartSpec} */ ({ ...tvBaseSpec(vm), kind: ctx.doughnut ? "doughnut" : "pie", legend: true });
  // Tableau's donut: a pie of slices (colour / angle) and, on the second axis, a smaller plain pie – the hole,
  // in the background colour, carrying the total as its label
  const piePanes = roles.panes.filter(p => tvMarkToken(p.markClass) === "pie");
  const slicing = p => p.encodings.some(e => e.channel === "wedge-size" || e.channel === "color");
  const slicePane = ctx.doughnut ? piePanes.find(slicing) || null : null;
  const holePane = ctx.doughnut ? piePanes.find(p => p.id && !slicing(p)) || null : null;
  let categories, values, cats = null;
  if ((roles.color && roles.color.measureNames) || (catCi < 0 && vm.cols.some(c => c.pivoted))) {
    const measures = vm.cols.map((c, i) => c.pivoted ? i : -1).filter(i => i >= 0);
    categories = { names: ["Measure Names"], levels: [measures.map(ci => tvMeasureLabel(vm, ci))] };
    values = measures.map(ci => vm.rows.reduce((s, r) => s + (tfDvNum(r[ci]) || 0), 0));
    angleCi = measures[0];
  } else {
    if (catCi < 0 || angleCi < 0) throw new Error("pie has no colour dimension or angle measure");
    cats = tvCategories(vm, [catCi], false);
    categories = { names: [tvMeasureLabel(vm, catCi)], levels: cats.levels };
    values = tvSum(vm, cats, angleCi);
  }
  // a dual-pie donut's hole contributes a mark with no category: no value, or the total of the slices → not a slice
  const total = values.reduce((s, v) => s + (v || 0), 0);
  const nullish = l => /^(null|%null%|\(null\))?$/i.test(String(l == null ? "" : l).trim());
  const isHole = i => nullish(categories.levels[0][i]) &&
    (!values[i] || (ctx.doughnut && Math.abs(values[i] - (total - values[i])) <= Math.abs(values[i]) * 0.005));
  const keep = categories.levels[0].map((_, i) => !isHole(i));
  const kept = keep.map((k, i) => k ? i : -1).filter(i => i >= 0);
  if (keep.includes(false)) {
    categories = { ...categories, levels: categories.levels.map(lv => lv.filter((_, i) => keep[i])) };
    values = values.filter((_, i) => keep[i]);
  }
  if (categories.levels[0].length > 200) throw new Error("too many pie slices for an Excel chart");
  const angleRef = vm.cols[angleCi] && vm.cols[angleCi].ref;
  const catRef = catCi >= 0 && vm.cols[catCi].ref;
  // a "% of Total" quick table calc on Label → Excel's percentage label instead of the raw value
  const labelPct = roles.labelRefs.some(r => /^pcto$/i.test(r.deriv || ""));
  const labelValue = !!angleRef && roles.labelRefs.some(r => !/^pcto$/i.test(r.deriv || "") && tfSameField(r, angleRef)) ||
    (!labelPct && tvLabelsOn(roles, angleRef));
  const labelCat = !!catRef && roles.labelRefs.some(r => tfSameField(r, catRef));
  spec.categories = categories;
  spec.numFmt = tvNumFmt(vm, angleCi);
  spec.series = [{
    name: tvMeasureLabel(vm, angleCi),
    color: tvMarkColor(roles),
    values,
    pointColors: categories.levels[0].map((v, i) => (scale && scale(v, i)) || tvHex(TABLEAU_10[i % TABLEAU_10.length])),
    labels: labelValue || labelCat || labelPct,
    labelParts: { value: labelValue, category: labelCat, percent: labelPct },
    labelNumFmt: labelPct && !labelValue ? "0.0%" : undefined
  }];
  if (slicePane && cats) {
    // the field values as Tableau writes them (measures in their number format, empty for null)
    const text = (row, ref) => {
      const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
      if (ci < 0 || !row || tfIsNull(row[ci])) return "";
      const n = tfDvNum(row[ci]);
      return vm.cols[ci].isHeader || n === null ? tvText(row[ci]) : tfFormatNumber(n, tvNumFmt(vm, ci));
    };
    const rowsOf = new Map();
    vm.rows.forEach(r => { const i = cats.indexOf(r); if (i !== undefined && keep[i] && !rowsOf.has(i)) rowsOf.set(i, r); });
    const flat = lines => lines.map(l => l.map(s => s.text).join("")).filter(t => t.trim()).join("\n");
    // slice labels: only the slices' own marks card (blank until its label fields have values)
    const texts = kept.map(i => flat(tvPaneLabel(slicePane, ref => text(rowsOf.get(i), ref))));
    const labelled = texts.some(Boolean) && slicePane.style && tvPaneRule(slicePane, "mark", "mark-labels-show") !== "false";
    Object.assign(spec.series[0], { labels: labelled, labelTexts: labelled ? texts : undefined, labelParts: undefined, labelNumFmt: undefined });
    if (holePane) {
      // the hole's label: its fields over all the slices (the total), in its runs' fonts
      const sumOf = ref => {
        const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
        if (ci < 0) return "";
        if (vm.cols[ci].isHeader) return "";
        const s = vm.rows.reduce((acc, r) => { const i = cats.indexOf(r); return i !== undefined && keep[i] ? acc + (tfDvNum(r[ci]) || 0) : acc; }, 0);
        return tfFormatNumber(s, tvNumFmt(vm, ci));
      };
      spec.centerLabel = tvPaneLabel(holePane, sumOf).filter(l => l.some(s => s.text.trim())).map(l => l.map(s => ({
        text: s.text, bold: !!s.props.bold, size: s.props.fontSize, color: tvHex(s.props.color) || undefined, font: s.props.fontName })));
      // the hole's size against the ring's (the marks' Size sliders)
      const sizeOf = p => Number(tvPaneRule(p, "mark", "size"));
      const ratio = sizeOf(holePane) / sizeOf(slicePane);
      if (ratio > 0 && ratio < 1) spec.holeSize = Math.round(Math.max(10, Math.min(90, ratio * 100)));
    }
  }
  return [spec];
}

/**
 * @param {ChartContext} ctx
 * @param {{ lines?: boolean }} [opts] lines: a Line mark with a measure on both axes – the points joined in order of x
 * @returns {ChartSpec[]}
 */
function tvScatterSpec(ctx, opts = {}) {
  const { vm, roles } = ctx;
  const xm = roles.cols.values[0], ym = roles.rows.values[0];
  const scale = tvColorScale(vm, roles, ctx.markToken);
  const spec = /** @type {ChartSpec} */ ({ ...tvBaseSpec(vm), kind: "scatter", numFmt: tvNumFmt(vm, ym.ci), xNumFmt: tvNumFmt(vm, xm.ci),
                 valueTitle: tvMeasureLabel(vm, ym.ci), xTitle: tvMeasureLabel(vm, xm.ci) });
  const point = r => ({ x: tfDvNum(r[xm.ci]), y: tfDvNum(r[ym.ci]) });
  const rows = vm.rows.filter(r => { const p = point(r); return p.x !== null && p.y !== null; });
  if (opts.lines) rows.sort((a, b) => point(a).x - point(b).x || point(a).y - point(b).y);
  if (rows.length > TV_MAX_POINTS) throw new Error(`${rows.length} marks – too many for an Excel scatter chart`);
  const colorCi = roles.color && !roles.color.measureNames ? roles.color.ci : -1;
  // Label = a dimension (store name …) → Excel "value from cells" labels with that text, when
  // few enough marks to stay readable; numbers only when a measure itself is on Label
  const labelDim = roles.textDims[0] ?? -1;
  const measureLabel = roles.labelRefs.some(r => tfSameField(r, ym.ref) || tfSameField(r, xm.ref));
  const textLabels = labelDim >= 0 && rows.length <= 40;
  const labels = textLabels || measureLabel || (labelDim < 0 && tvLabelsOn(roles, ym.ref));
  const labelOf = r => tvText(r[labelDim]);
  // a Size measure is part of the meaning (Tableau's circles sized by it): Excel's bubble chart, area = size
  const sizeCi = !opts.lines && roles.size >= 0 && !vm.cols[roles.size].isHeader ? roles.size : -1;
  const sizeOf = r => { const n = tfDvNum(r[sizeCi]); return n === null ? null : Math.abs(n); };
  if (colorCi >= 0 && !roles.color.continuous) {
    const values = tvColorValues(vm, roles);
    if (values.length > TV_MAX_SERIES) throw new Error("too many colour values for an Excel chart");
    spec.series = values.map((v, i) => {
      const own = rows.filter(r => tvText(r[colorCi]) === v), pts = own.map(point);
      return { name: v, color: (scale && scale(v)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
               x: pts.map(p => p.x), y: pts.map(p => p.y), labels, ...(opts.lines ? { line: true } : {}),
               ...(sizeCi >= 0 ? { size: own.map(sizeOf) } : {}),
               labelTexts: textLabels ? own.map(labelOf) : undefined };
    });
  } else {
    const pts = rows.map(point);
    spec.series = [{
      name: tvMeasureLabel(vm, ym.ci), color: tvMarkColor(roles), x: pts.map(p => p.x), y: pts.map(p => p.y), labels,
      ...(opts.lines ? { line: true } : {}), ...(sizeCi >= 0 ? { size: rows.map(sizeOf) } : {}),
      labelTexts: textLabels ? rows.map(labelOf) : undefined,
      pointColors: colorCi >= 0 && scale ? rows.map(r => scale(tfDvNum(r[colorCi]))) : undefined
    }];
  }
  if (sizeCi >= 0) Object.assign(spec, { kind: "bubble", sizeTitle: tvMeasureLabel(vm, sizeCi), sizeNumFmt: tvNumFmt(vm, sizeCi) });
  spec.legend = spec.series.length > 1;
  return [spec];
}

/* value axis + category dimensions shared by waterfall / box plot */
/** @param {ChartContext} ctx */
function tvAxisLayout(ctx) {
  const { roles } = ctx;
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) throw new Error("no continuous measure axis to chart");
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  return { valueShelf, horizontal: valueShelf === "cols", measure: roles[valueShelf].values[0],
           catDims: [...roles[valueShelf].dims, ...roles[catShelf].dims] };
}

/* ── waterfall: Gantt bars sized by -measure on a running total →
 *    stacked columns: invisible base + increase / decrease, each step in its category's colour, and the
 *    Text field written over each floating bar ─────────────────────────────── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvWaterfallSpec(ctx) {
  const { vm, roles } = ctx;
  const { horizontal, measure, catDims, valueShelf } = tvAxisLayout(ctx);
  const fieldLabels = vm.fmt.hasModel ? vm.fmt.fieldLabelsShown(valueShelf === "rows" ? "cols" : "rows") : true;
  const cats = tvCategories(vm, catDims.map(d => d.ci), catDims.some(d => d.continuous));
  if (cats.count > TV_MAX_POINTS) throw new Error(`${cats.count} categories – too many for an Excel chart`);
  const running = tvSum(vm, cats, measure.ci);
  if (running.some(v => v !== null && v < 0)) throw new Error("running total goes below zero – not drawable as stacked columns");
  const base = [], up = [], down = [], ends = [];
  let prev = 0;
  running.forEach(v => {
    const cur = v === null ? prev : v;
    const delta = cur - prev;
    base.push(Math.min(prev, cur));
    up.push(delta > 0 ? delta : null);
    down.push(delta < 0 ? -delta : null);
    ends.push(Math.max(prev, cur));                             // the top of the floating bar
    prev = cur;
  });
  const markColor = tvMarkColor(roles);
  const scale = tvColorScale(vm, roles, "ganttbar");
  // colour = the category (each step its own colour), or a measure (usually the step itself)
  const colorLevel = roles.color && !roles.color.continuous ? catDims.findIndex(d => d.ci === roles.color.ci) : -1;
  const stepColors = !scale ? null : colorLevel >= 0 ? cats.levels[colorLevel].map(v => scale(v))
    : roles.color && roles.color.continuous && roles.color.ci >= 0 ? tvSum(vm, cats, roles.color.ci).map(v => scale(v)) : null;
  const label = tvMeasureLabel(vm, measure.ci);
  const numFmt = tvNumFmt(vm, measure.ci);
  // mark labels: the Text field (the step, or the running total) over each bar; Excel writes stacked labels
  // inside the bars, so an invisible line along the bar tops carries them above, as text in its format
  const textCi = roles.labelRefs.map(r => vm.cols.findIndex(c => c.ref && tfSameField(c.ref, r))).find(i => i >= 0);
  const labelCi = textCi === undefined ? measure.ci : textCi;
  const labels = tvLabelsOn(roles, vm.cols[labelCi].ref);
  const labelFmt = tvNumFmt(vm, labelCi);
  const labelValues = labelCi === measure.ci ? running : tvSum(vm, cats, labelCi);
  const above = labels && !horizontal;
  // decreases are drawn as positive heights but labelled as the negative step, like Tableau
  const positive = (numFmt.match(/^((?:"[^"]*"|[^;])*)/) || [])[1] || "General";
  const downFmt = positive === "General" ? "-General" : "-" + positive;
  /** @type {ChartSeries[]} */
  const series = [
    { name: "Base", type: "bar", values: base, color: null, labels: false },
    { name: label + " (increase)", type: "bar", values: up, color: markColor, pointColors: stepColors, labels: labels && !above },
    { name: label + " (decrease)", type: "bar", values: down, color: markColor, pointColors: stepColors, labels: labels && !above, labelNumFmt: downFmt }
  ];
  // Analysis → Totals → Show Grand Totals on the category shelf: Tableau ends with a full bar from zero to the
  // final running total (drawn in grey)
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  const totals = vm.fmt.sheetModel && vm.fmt.sheetModel.grandTotals && vm.fmt.sheetModel.grandTotals[catShelf];
  if (totals && running.length) {
    const n = running.length;
    series.forEach(s => { s.values = [...s.values, null]; if (s.pointColors) s.pointColors = [...s.pointColors, null]; });
    series.push({ name: "Grand Total", type: "bar", values: [...Array(n).fill(null), prev], color: "9E9E9E", labels: labels && !above });
    cats.levels = cats.levels.map((lv, i) => [...lv, i === 0 ? "Grand Total" : ""]);
    ends.push(prev); labelValues.push(labelCi === measure.ci ? prev : null);
  }
  if (above) {
    series.push({ name: label + " (labels)", type: "line", values: ends, color: null, line: false, marker: false, labels: true,
                  labelTexts: labelValues.map(v => v === null ? "" : tfFormatNumber(v, labelFmt)) });
  }
  return [{
    // Tableau's Gantt bars float thin, about as wide as the gaps between them
    ...tvBaseSpec(vm), kind: "bar", barDir: horizontal ? "bar" : "col", stacked: true, gapWidth: 100, legend: false,
    conversion: { strategy: "CONSTRUCTED", output: "Waterfall: stacked columns on an invisible running base",
                  note: "built from stacked columns (invisible base, increases, decreases) so Tableau's step colours and labels stay – Excel's own waterfall cannot colour each step" },
    // "Allow labels to overlap other marks" off (Tableau's default): overlapping labels are left out
    labelCull: !roles.panes.some(p => tvPaneRule(p, "mark", "mark-labels-cull") === "false"),
    categories: { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [""], levels: cats.levels },
    categoryTitle: fieldLabels ? catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / ") : "",
    numFmt, valueTitle: label, series
  }];
}

/* ── box plot: quartiles per category → line chart with up/down bars (box),
 *    high-low lines (whiskers) and a dash at the median ───────────────────── */
function tvQuantile(sorted, p) {                      // linear interpolation, as Tableau / QUARTILE.INC
  const pos = (sorted.length - 1) * p, lo = Math.floor(pos), hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvBoxPlotSpec(ctx) {
  const { vm, roles } = ctx;
  const { measure, catDims } = tvAxisLayout(ctx);
  const cats = tvCategories(vm, catDims.map(d => d.ci), catDims.some(d => d.continuous));
  if (cats.count > 200) throw new Error("too many boxes for an Excel chart");
  const groups = Array.from({ length: cats.count }, () => []);
  vm.rows.forEach(r => {
    const n = tfDvNum(r[measure.ci]);
    const i = cats.indexOf(r);
    if (n !== null && i !== undefined) groups[i].push(n);
  });
  const stats = groups.map(g => {
    if (!g.length) return null;
    const s = g.slice().sort((a, b) => a - b);
    const q1 = tvQuantile(s, 0.25), med = tvQuantile(s, 0.5), q3 = tvQuantile(s, 0.75), iqr = q3 - q1;
    // standard whiskers: the furthest marks within 1.5 × IQR of the box
    const low = s.find(v => v >= q1 - 1.5 * iqr), high = [...s].reverse().find(v => v <= q3 + 1.5 * iqr);
    return { q1, med, q3, low, high };
  });
  const col = k => stats.map(st => st ? st[k] : null);
  const boxColor = tvMarkColor(roles);
  return [{
    ...tvBaseSpec(vm), kind: "line", boxPlot: { color: boxColor }, legend: false, includeZero: false,
    conversion: { strategy: "CONSTRUCTED", output: "Box plot: quartile boxes (up/down bars), whiskers (high-low lines), median marks",
                  note: "Tableau's quartiles and 1.5 × IQR whiskers computed from the marks – Excel's box & whisker chart would compute its own statistics" },
    categories: { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [""], levels: cats.levels },
    categoryTitle: catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / "),
    numFmt: tvNumFmt(vm, measure.ci), valueTitle: tvMeasureLabel(vm, measure.ci),
    // up/down bars span the FIRST and LAST series → Q1 … Q3; high-low lines span all series
    series: [
      { name: "Lower quartile", type: "line", values: col("q1"), color: boxColor, line: false, marker: false },
      { name: "Lower whisker", type: "line", values: col("low"), color: boxColor, line: false, marker: false },
      { name: "Median", type: "line", values: col("med"), color: "333333", line: false, marker: true, markerSymbol: "dash", markerSize: 14 },
      { name: "Upper whisker", type: "line", values: col("high"), color: boxColor, line: false, marker: false },
      { name: "Upper quartile", type: "line", values: col("q3"), color: boxColor, line: false, marker: false }
    ]
  }];
}

/* ── Gantt: per row, marks start at the date axis value and last Size days →
 *    stacked horizontal bars of alternating invisible gaps and visible bars ─ */
const TV_GANTT_MAX_SERIES = 250;

function tvExcelSerial(dv) {
  if (tfIsNull(dv)) return null;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  if (typeof v === "number") return v;                         // already a serial / numeric axis
  const t = v instanceof Date ? v.getTime() : Date.parse(String(v));
  return isNaN(t) ? null : t / 86400000 + 25569;               // days since 1899-12-30 (UTC)
}

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvGanttSpec(ctx) {
  const { vm, roles } = ctx;
  const startDim = [...roles.cols.dims, ...roles.rows.dims].find(d => d.continuous);
  if (!startDim) throw new Error("no continuous date axis for the Gantt bars");
  if (roles.size < 0) throw new Error("no Size measure for the Gantt bar length");
  const startShelf = roles.cols.dims.includes(startDim) ? "cols" : "rows";
  const catDims = roles[startShelf === "cols" ? "rows" : "cols"].dims.filter(d => !d.continuous);
  const cats = tvCategories(vm, catDims.map(d => d.ci), false);
  const colorScale = tvColorScale(vm, roles, "ganttbar");
  const colorCi = roles.color && !roles.color.measureNames ? roles.color.ci : -1;
  const markColor = tvMarkColor(roles);
  const marks = Array.from({ length: cats.count }, () => []);
  vm.rows.forEach(r => {
    const start = tvExcelSerial(r[startDim.ci]), len = tfDvNum(r[roles.size]), i = cats.indexOf(r);
    if (start === null || len === null || i === undefined) return;
    const c = colorCi >= 0 && colorScale ? colorScale(roles.color.continuous ? tfDvNum(r[colorCi]) : tvText(r[colorCi])) : null;
    marks[i].push({ start, end: start + Math.max(0, len), color: c || markColor });
  });
  const all = marks.flat();
  if (!all.length) throw new Error("no Gantt marks with a start date and a length");
  const axisMin = Math.floor(Math.min(...all.map(m => m.start)));
  const axisMax = Math.ceil(Math.max(...all.map(m => m.end)));
  const depth = Math.max(...marks.map(m => m.length));
  if (2 * depth > TV_GANTT_MAX_SERIES) throw new Error(`${depth} bars in one row – too many for an Excel chart`);
  /** @type {ChartSeries[]} */
  const series = [];
  for (let k = 0; k < depth; k++) {
    series.push({ name: `Gap ${k + 1}`, type: "bar", values: [], color: null, labels: false });
    series.push({ name: `Bar ${k + 1}`, type: "bar", values: [], color: markColor, pointColors: [], labels: false });
  }
  marks.forEach((row, i) => {
    row.sort((a, b) => a.start - b.start);
    let edge = 0;
    row.forEach((m, k) => {
      if (m.start < edge - 1e-9) throw new Error("overlapping Gantt bars in one row – not drawable as stacked bars");
      series[2 * k].values[i] = k === 0 ? m.start : m.start - edge;
      series[2 * k + 1].values[i] = m.end - m.start;
      series[2 * k + 1].pointColors[i] = m.color;
      edge = m.end;
    });
    for (let k = row.length; k < depth; k++) { series[2 * k].values[i] = null; series[2 * k + 1].values[i] = null; }
  });
  const dateFmt = axisMax - axisMin > 400 ? "mmm yyyy" : "d mmm yyyy";
  return [{
    ...tvBaseSpec(vm), kind: "bar", barDir: "bar", stacked: true, gapWidth: 40, legend: false,
    conversion: { strategy: "CONSTRUCTED", output: "Gantt: stacked bars with invisible gaps on a date axis",
                  note: "Excel has no Gantt chart: each bar starts at its date and lasts its Size, the gaps before it invisible" },
    valueMin: axisMin, valueMax: axisMax, includeZero: false,
    categories: { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [""], levels: cats.levels },
    categoryTitle: catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / "),
    numFmt: dateFmt, valueTitle: tvMeasureLabel(vm, startDim.ci), series
  }];
}

/* ── treemap: nothing on Rows/Columns, Size = measure, Label/Detail = the
 *    hierarchy → Excel treemap (chartex), each tile in its Tableau colour ── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvTreemapSpec(ctx) {
  const { vm, roles } = ctx;
  if (roles.size < 0) throw new Error("no Size measure for the treemap");
  let levelCis = [...roles.detailDims, ...roles.textDims].filter((ci, i, a) => a.indexOf(ci) === i);
  if (roles.color && !roles.color.continuous && roles.color.ci >= 0 && !levelCis.includes(roles.color.ci)) levelCis.unshift(roles.color.ci);
  if (!levelCis.length) levelCis = vm.cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);
  if (!levelCis.length) throw new Error("no dimension to split the treemap");
  // outermost level = the one with the fewest distinct values
  const distinct = ci => new Set(vm.rows.map(r => tvText(r[ci]))).size;
  levelCis.sort((a, b) => distinct(a) - distinct(b));
  const cats = tvCategories(vm, levelCis, false);
  if (cats.count > 1000) throw new Error(`${cats.count} tiles – too many for an Excel treemap`);
  const sizes = tvSum(vm, cats, roles.size);
  const scale = tvColorScale(vm, roles, "square");
  let pointColors;
  if (scale && roles.color.continuous) pointColors = tvSum(vm, cats, roles.color.ci).map(v => scale(v));
  else if (scale && roles.color.ci >= 0) {
    const level = levelCis.indexOf(roles.color.ci);
    pointColors = cats.levels[level].map((v, i) => scale(v) || tvHex(TABLEAU_10[i % TABLEAU_10.length]));
  }
  // drop empty / negative tiles: a treemap cannot draw them
  const keep = sizes.map((v, i) => v !== null && v > 0 ? i : -1).filter(i => i >= 0);
  // Excel builds the hierarchy from consecutive rows: keep each parent's tiles together,
  // biggest parent first, biggest tile first inside it (Tableau's treemap order)
  const pathKey = (i, l) => cats.levels.slice(0, l + 1).map(lv => lv[i]).join("\u0001");
  const totals = cats.levels.map((_, l) => {
    const m = new Map();
    keep.forEach(i => m.set(pathKey(i, l), (m.get(pathKey(i, l)) || 0) + sizes[i]));
    return m;
  });
  keep.sort((a, b) => {
    for (let l = 0; l < cats.levels.length; l++) {
      const ka = pathKey(a, l), kb = pathKey(b, l);
      if (ka === kb) continue;
      return (totals[l].get(kb) - totals[l].get(ka)) || (ka < kb ? -1 : 1);
    }
    return 0;
  });
  return [{
    ...tvBaseSpec(vm), kind: "treemap", legend: false,
    categories: { names: levelCis.map(ci => tvMeasureLabel(vm, ci)), levels: cats.levels.map(lv => keep.map(i => lv[i])) },
    numFmt: tvNumFmt(vm, roles.size),
    series: [{ name: tvMeasureLabel(vm, roles.size), values: keep.map(i => sizes[i]), color: tvMarkColor(roles),
               pointColors: pointColors ? keep.map(i => pointColors[i]) : undefined, labels: true,
               // Size measure also on Label → "Quality Issue 20K", like Tableau
               labelParts: { category: true, value: roles.labelRefs.some(r => tfSameField(r, vm.cols[roles.size].ref || {})) } }]
  }];
}

/* ── packed bubbles: one bubble per mark, area = Size → bubble chart; the bubbles are packed here
 *    (largest in the middle) and scaled to the chart's size when the chart XML is written ──── */
const TV_MAX_BUBBLES = 200;

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvPackedBubbleSpec(ctx) {
  const { vm, roles } = ctx;
  if (roles.size < 0) throw new Error("no Size measure for the bubbles");
  const colorDim = roles.color && !roles.color.measureNames && !roles.color.continuous && roles.color.ci >= 0 ? roles.color.ci : -1;
  let dims = [colorDim, ...roles.textDims, ...roles.detailDims].filter((ci, i, a) => ci >= 0 && ci !== roles.size && a.indexOf(ci) === i);
  if (!dims.length) dims = vm.cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);
  if (!dims.length) throw new Error("no dimension to split the bubbles");
  const cats = tvCategories(vm, dims, false);
  if (cats.count > TV_MAX_BUBBLES) throw new Error(`${cats.count} bubbles – too many to pack in an Excel chart`);
  const sizes = tvSum(vm, cats, roles.size);
  const keep = sizes.map((v, i) => v !== null && v > 0 ? i : -1).filter(i => i >= 0);   // a bubble needs an area
  if (!keep.length) throw new Error("no bubble with a positive size");
  const max = Math.max(...keep.map(i => sizes[i]));
  const radii = keep.map(i => Math.sqrt(sizes[i] / max));
  const { centres, box } = tvPackCircles(radii, 0.03);
  const scale = tvColorScale(vm, roles, "circle");
  let pointColors;
  if (scale && roles.color && roles.color.continuous && roles.color.ci >= 0) {
    const totals = tvSum(vm, cats, roles.color.ci);
    pointColors = keep.map(i => scale(totals[i]));
  } else if (scale && colorDim >= 0) {
    const level = dims.indexOf(colorDim);
    pointColors = keep.map((i, k) => scale(cats.levels[level][i]) || tvHex(TABLEAU_10[k % TABLEAU_10.length]));
  }
  // Tableau's label and tooltip per bubble: its template filled with the bubble's values – Tableau's own formatted
  // value when the bubble is one row, else the rows' sum in the field's format
  const rowsOf = keep.map(() => /** @type {any[][]} */ ([]));
  const at = new Map(keep.map((i, k) => [i, k]));
  vm.rows.forEach(row => { const k = at.get(cats.indexOf(row)); if (k !== undefined) rowsOf[k].push(row); });
  const valueOf = k => ref => {
    const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
    const rows = rowsOf[k];
    if (ci < 0 || !rows.length) return "";
    if (rows.length === 1 || vm.cols[ci].isHeader) return tfDvText(rows[0][ci]);
    const nums = rows.map(r => tfDvNum(r[ci])).filter(v => v !== null);
    const nf = vm.fmt.numFmtFor ? vm.fmt.numFmtFor(ref) : null;
    return nums.length ? (nf ? tfFormatNumber(nums.reduce((a, b) => a + b, 0), nf) : String(nums.reduce((a, b) => a + b, 0))) : "";
  };
  const text = lines => lines.map(line => line.map(x => x.text).join("")).join("\n");
  const labelTpl = vm.fmt.hasModel ? vm.fmt.labelRunsTemplate() : [];
  const tipTpl = vm.fmt.hasModel ? vm.fmt.tooltipTemplate() : [];
  // without a label template: the label dimension (Label, else Detail, else the innermost level)
  const labelLevel = dims.indexOf(roles.textDims[0] ?? roles.detailDims[0] ?? dims[dims.length - 1]);
  const labels = labelTpl.length ? vm.fmt.markLabelsShown() : roles.textDims.length > 0 || roles.labelRefs.length > 0;
  const labelTexts = labels ? keep.map((i, k) => labelTpl.length ? text(renderTemplate(labelTpl, valueOf(k))) : cats.levels[labelLevel][i]) : undefined;
  const tooltips = tipTpl.length ? keep.map((i, k) => text(renderTemplate(tipTpl, valueOf(k)))) : undefined;
  return [{
    ...tvBaseSpec(vm), kind: "bubble", gridlines: false, axesHidden: true, legend: false,
    xTitle: "Bubble x", numFmt: "General", xNumFmt: "General",
    sizeTitle: tvMeasureLabel(vm, roles.size), sizeNumFmt: tvNumFmt(vm, roles.size), packed: box,
    series: [{
      name: "Bubble y", color: tvMarkColor(roles), x: centres.map(p => p.x), y: centres.map(p => p.y),
      size: keep.map(i => sizes[i]), pointColors, labels, labelTexts, tooltips
    }],
    labelCull: vm.fmt.hasModel ? vm.fmt.markLabelsCulled() : true
  }];
}


/* ═══ charts/model/index.js ═══════════════════════════════════════════════════════════════════════ */
/* Entry point: visual model → Excel chart specs (throws when not representable). */

/* ══════════════════════════════════════════════════════════════════════════
 * TABLEAU VISUAL → EXCEL CHART SPECS
 * ──────────────────────────────────────────────────────────────────────────
 * buildExcelChartSpecs(visualModel, model) reads
 *   – live summary data (already pivoted / merged / sorted by buildViewModel)
 *   – TWB shelves, panes, marks, encodings, colours, number formats, labels
 *   – the live visual specification when no workbook was loaded
 * and returns renderer-neutral chart specs (ChartSpec, src/types.d.ts) for charts/writer.
 * It throws when a visual cannot be represented faithfully, so the caller
 * can fall back to the Tableau image renderer or the data table.
 * ══════════════════════════════════════════════════════════════════════════ */

/**
 * @param {VisualModel} visualModel
 * @param {FormatModel | null} model
 * @returns {ChartSpec[]}
 */
function buildExcelChartSpecs(visualModel, model) {
  const vm = visualModel.viewModel;
  if (!vm || !vm.rows.length) throw new Error("visual has no data rows");
  const roles = tvRoles(vm, model, visualModel.source && visualModel.source.visualSpec);
  const markToken = visualModel.metadata.markToken || "";
  /** @type {ChartContext} */
  const ctx = { vm, roles, markToken, type: visualModel.type, model,
                autoMark: tvAutomaticMark(visualShelfShape(vm.fmt && vm.fmt.sheetModel, model)),
                roundedBar: visualModel.type === VISUAL_TYPES.BAR && isRoundedBar(vm, model) };
  if (roles.source === "summary" && /^(circle|shape|square)$/.test(markToken)) {
    throw new Error("mark layout unknown – load the workbook file for this visual");
  }
  let specs;
  if (visualModel.type === VISUAL_TYPES.HISTOGRAM) specs = tvHistogramSpecs(ctx);
  else if (visualModel.type === VISUAL_TYPES.WATERFALL) specs = tvWaterfallSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.BOXPLOT) specs = tvBoxPlotSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.GANTT) specs = tvGanttSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.TREEMAP) specs = tvTreemapSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.BUBBLE) specs = tvPackedBubbleSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.PIE) {
    ctx.doughnut = roles.panes.filter(p => tvMarkToken(p.markClass) === "pie").length >= 2;
    specs = tvPieSpec(ctx);
  } else if (visualModel.type === VISUAL_TYPES.SCATTER && roles.rows.values.length && roles.cols.values.length) {
    specs = tvScatterSpec(ctx);
  } else if (visualModel.type === VISUAL_TYPES.LINE && roles.rows.values.length && roles.cols.values.length &&
             ![...roles.rows.values, ...roles.cols.values].some(v => v.mv)) {
    specs = tvScatterSpec(ctx, { lines: true });         // a line of one measure against another (a Pareto curve)
  } else {
    specs = tvCartesianSpecs(ctx);
  }
  // reference lines last: a line Excel cannot draw adds its note to the conversion the approximations named
  specs.forEach(s => { s.name = visualModel.metadata.worksheetName; s.rolesSource = roles.source; tvApplyWorkbookAxes(s, ctx); tvNoteApproximations(s, ctx);
                       tvApplyReferenceLines(s, ctx); });
  return specs;
}


/* ═══ export/conditional-format.js ════════════════════════════════════════════════════════════════ */
/* Fallback conditional formatting for tables without Tableau colours. */

function isNumeric(value) {
  if (value === null || value === undefined) return false;
  
  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/%/g, "")
    .trim();
  
  return cleaned !== "" && !isNaN(Number(cleaned));
}

function shouldSkipConditionalFormatting(columnName, value, fieldRole = null) {
  if (!columnName) return false;
  
  const lowerColName = columnName.toLowerCase();
  
  const categoricalKeywords = [
    'year', 'quarter', 'month', 'date', 'time', 
    'category', 'region', 'product', 'name', 'id',
    'country', 'city', 'state', 'department', 'type',
    'status', 'group', 'segment', 'class'
  ];
  
  const geoKeywords = [
    'latitude', 'lat', 'latitud', 'ycoord', 'y_coord',
    'longitude', 'long', 'lon', 'longitud', 'xcoord', 'x_coord',
    'location', 'geo', 'geography', 'coordinates', 'coords',
    'postal', 'zip', 'zipcode', 'postcode', 'address'
  ];
  
  for (const keyword of geoKeywords) {
    if (lowerColName.includes(keyword)) {
      console.log(`🗺️ Skipping conditional formatting for geographic column: "${columnName}"`);
      return true;
    }
  }
  
  for (const keyword of categoricalKeywords) {
    if (lowerColName.includes(keyword)) {
      return true;
    }
  }
  
  if (value !== undefined && value !== null) {
    const strValue = String(value).trim();
    const numValue = parseFloat(strValue);
    
    if (!isNaN(numValue)) {
      if (lowerColName.includes('lat') && numValue >= -90 && numValue <= 90) {
        console.log(`🗺️ Skipping conditional formatting for latitude column: "${columnName}" (value: ${numValue})`);
        return true;
      }
      
      if (lowerColName.includes('lon') && numValue >= -180 && numValue <= 180) {
        console.log(`🗺️ Skipping conditional formatting for longitude column: "${columnName}" (value: ${numValue})`);
        return true;
      }
      
      if (lowerColName.match(/lat/i) && numValue >= -90 && numValue <= 90) {
        console.log(`🗺️ Skipping conditional formatting for geographic coordinate column: "${columnName}"`);
        return true;
      }
      
      if (lowerColName.match(/lon/i) && numValue >= -180 && numValue <= 180) {
        console.log(`🗺️ Skipping conditional formatting for geographic coordinate column: "${columnName}"`);
        return true;
      }
    }
    
    if (/^\d{4}$/.test(strValue)) {
      const yearNum = parseInt(strValue, 10);
      if (yearNum >= 1900 && yearNum <= 2100) {
        return true;
      }
    }
  }
  
  return false;
}

function isGeographicCoordinate(value, columnName) {
  if (!isNumeric(value)) return false;
  
  const lowerColName = columnName.toLowerCase();
  const numValue = parseFloat(String(value).trim());
  
  if (lowerColName.includes('latitude') || lowerColName.includes('lat')) {
    return numValue >= -90 && numValue <= 90;
  }
  
  if (lowerColName.includes('longitude') || lowerColName.includes('lon')) {
    return numValue >= -180 && numValue <= 180;
  }
  
  if (numValue >= -90 && numValue <= 90 && (lowerColName.includes('coord') || lowerColName.includes('geo'))) {
    return true;
  }
  
  if ((numValue >= -180 && numValue <= -90) || (numValue >= 90 && numValue <= 180)) {
    if (lowerColName.includes('coord') || lowerColName.includes('geo')) {
      return true;
    }
  }
  
  return false;
}

function isNumericForFormatting(value, columnName) {
  if (!isNumeric(value)) return false;
  
  if (columnName) {
    const strValue = String(value).trim();
    if (/^\d{4}$/.test(strValue)) {
      const yearNum = parseInt(strValue, 10);
      if (yearNum >= 1900 && yearNum <= 2100) {
        return false;
      }
    }
    
    if (isGeographicCoordinate(value, columnName)) {
      return false;
    }
  }
  
  return true;
}

function getZeroCenteredColor(value, maxAbs) {
  if (maxAbs === 0) {
    return "FFFFFF";
  }
  
  const intensity = Math.min(Math.abs(value) / maxAbs, 1);
  
  let r = 255;
  let g = 255;
  let b = 0;
  
  if (value > 0) {
    r = Math.round(255 * (1 - intensity));
    g = 255;
    b = Math.round(100 * (1 - intensity));
  } else if (value < 0) {
    r = 255;
    g = Math.round(255 * (1 - intensity));
    b = Math.round(100 * (1 - intensity));
  }
  
  return (
    "FF" +
    r.toString(16).padStart(2, "0") +
    g.toString(16).padStart(2, "0") +
    b.toString(16).padStart(2, "0")
  ).toUpperCase();
}

/* ── FALLBACK: Hardcoded conditional formatting (only used if no XML colors) ── */
function applyConditionalFormattingToTable(worksheet, rows, headers, startRow, startCol, colWidths) {
  headers.forEach((header, colIndex) => {
    const columnName = header.fieldName || header.fieldId || `Column_${colIndex + 1}`;
    const firstValue = rows[0]?.[colIndex]?.formattedValue || rows[0]?.[colIndex]?.value;
    
    let shouldSkip = shouldSkipConditionalFormatting(columnName, firstValue);
    
    if (shouldSkip) {
      console.log(`⚠️ Skipping conditional formatting for column: "${columnName}"`);
      return;
    }
    
    const numericValues = [];
    
    rows.forEach(row => {
      const value = row[colIndex]?.formattedValue || row[colIndex]?.value;
      if (isNumericForFormatting(value, columnName)) {
        let numValue = String(value).replace(/,/g, "").replace(/%/g, "");
        numericValues.push(parseFloat(numValue));
      }
    });
    
    if (numericValues.length === 0) return;
    
    const maxAbs = Math.max(...numericValues.map(v => Math.abs(v)));
    
    rows.forEach((row, rowIndex) => {
      const rawValue = row[colIndex]?.formattedValue || row[colIndex]?.value;
      
      if (!isNumericForFormatting(rawValue, columnName)) return;
      
      let numValue = String(rawValue).replace(/,/g, "").replace(/%/g, "");
      const value = parseFloat(numValue);
      
      const excelRow = startRow + rowIndex;
      const excelCol = startCol + colIndex;
      const cell = worksheet.getCell(excelRow + 1, excelCol + 1);
      
      const color = getZeroCenteredColor(value, maxAbs);
      
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color }
      };
    });
  });
}


/* ═══ export/mark-cells.js ════════════════════════════════════════════════════════════════════════ */
/* A table built from marks (one constant axis per column) drawn in cells: each column takes the look of the
 * marks in its pane – circles / squares as the cell's fill, arrow shapes as a glyph before the label, bars as
 * Excel data bars – coloured as Tableau colours them. */

const DEFAULT_MARK = "FF4E79A7";                       // Tableau's default mark colour

/* Tableau's "Arrows" shape palettes: the file number is the direction, the same in every set */
const ARROW_BY_INDEX = ["", "↓", "↘", "→", "↗", "↑", "↖", "←", "↙"];
/* Tableau's own shapes (":filled/circle" …) */
const BUILT_IN = { circle: "●", square: "■", triangle: "▲", diamond: "◆", plus: "+", cross: "✕", asterisk: "✱",
                   "dot-circle": "◉", "dot-square": "▣" };

/** the character drawn for a shape: an arrow's direction, or the shape itself @param {string} name */
function tvShapeGlyph(name) {
  if (!name) return null;
  const arrow = name.match(/^Arrows\/\w+-(\d)\.png$/i);
  if (arrow) return ARROW_BY_INDEX[+arrow[1]] || null;
  const own = name.match(/^:(filled\/)?([a-z-]+)$/i);
  if (own) {
    const g = BUILT_IN[own[2].toLowerCase()];
    if (!g) return null;
    return own[1] ? g : ({ "●": "○", "■": "□", "▲": "△", "◆": "◇" })[g] || g;
  }
  return "●";                                           // a custom picture: a dot in its colour
}

/** a number format with a glyph before the value in every section (two at least: no minus before the glyph)
 * @param {string | undefined} fmt @param {string} glyph */
function tvGlyphNumFmt(fmt, glyph) {
  const secs = (tfSplitSections(fmt || "General") || ["General"]).slice(0, 3).map(s => s || "General");
  if (secs.length === 1) secs.push("-" + secs[0]);
  return secs.map(s => `"${glyph} "${s}`).join(";");
}

/**
 * @typedef {{ kind: "fill" | "glyph" | "bar", colorAt: (row: number) => string | null,
 *             glyphAt: (row: number) => string | null, group: string }} MarkColumn
 * kind fill: circles / squares / shapes without a Shape field; glyph: a Shape field (arrows);
 * bar: bars, sized by their label value. group: columns whose bars share one scale (same colour field).
 */

/**
 * How each value column of a table built from marks looks.
 * @param {SheetFormatter} fmt @param {ViewColumn[]} cols @param {any[][]} rows @param {number[]} order
 * @returns {Map<number, MarkColumn>}
 */
function buildMarkPlan(fmt, cols, rows, order) {
  /** @type {Map<number, MarkColumn>} */
  const plan = new Map();
  const panes = fmt.panesInOrder();
  const colOf = ref => cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
  order.forEach(ci => {
    const c = cols[ci];
    if (c.isHeader || !c.ref) return;
    const pp = panes.find(p => p.refs.some(r => tfSameField(r, c.ref)));
    if (!pp) return;
    const pane = pp.pane, mark = String(pane.markClass || "").toLowerCase();
    if (!/^(circle|square|shape|bar)$/.test(mark)) return;
    // colour: the pane's Color field, else its Marks card colour
    const enc = fmt.colorEncodingOf(pane);
    const fixedRule = ((pane.style && pane.style.rules && pane.style.rules.mark) || []).find(f => f.attr === "mark-color" && !f.field);
    const fixed = (fixedRule && tfArgb(fixedRule.value)) || DEFAULT_MARK;
    let colorAt = () => fixed;
    if (enc) {
      const k = colOf(enc.ref);
      if (k >= 0) {
        const vals = rows.map(r => tfIsNull(r[k]) ? null : enc.continuous ? tfDvNum(r[k]) : tfDvText(r[k]));
        const scale = tfBuildColorScale(fmt, enc, vals);
        if (scale) { const colors = vals.map(v => v === null ? null : scale(v)); colorAt = i => colors[i] || fixed; }
      }
    }
    const shapeEnc = pane.encodings.find(e => e.channel === "shape" && e.field);
    const group = enc ? `color:${enc.ref.inner}` : `col:${ci}`;
    if (mark === "bar") {
      plan.set(ci, { kind: "bar", colorAt, glyphAt: () => null, group });
    } else if (mark === "shape" && shapeEnc) {
      const map = fmt.shapeMap(shapeEnc.field) || {};
      const k = colOf(shapeEnc.field);
      const glyphs = rows.map(r => k >= 0 ? tvShapeGlyph(map[tfBucketKey(tfDvText(r[k]))]) : null);
      plan.set(ci, { kind: "glyph", colorAt, glyphAt: i => glyphs[i] || null, group });
    } else {
      plan.set(ci, { kind: "fill", colorAt, glyphAt: () => null, group });
    }
  });
  return plan;
}

/**
 * Data bars for the bar columns, written after the cells: one rule per cell when the colour varies by row.
 * Columns of one group share a scale; a column of negative values (the left half of a diverging bar) grows
 * leftwards from its right edge.
 * @param {import("exceljs").Worksheet} worksheet @param {Map<number, MarkColumn>} plan
 * @param {ViewColumn[]} cols @param {any[][]} rows @param {number[]} order
 * @param {number} firstRow 0-based sheet row of the first data row
 * @param {(k: number) => number} colAt 0-based sheet column of the table's k-th field
 */
function writeMarkBars(worksheet, plan, cols, rows, order, firstRow, colAt) {
  const bars = order.map((ci, k) => ({ ci, k, m: plan.get(ci) })).filter(x => x.m && x.m.kind === "bar");
  /** @type {Map<string, number>} */
  const maxOf = new Map();
  bars.forEach(({ ci, m }) => rows.forEach(r => {
    const v = tfDvNum(r[ci]);
    if (v !== null) maxOf.set(m.group, Math.max(maxOf.get(m.group) || 0, Math.abs(v)));
  }));
  bars.forEach(({ ci, k, m }) => {
    const M = maxOf.get(m.group);
    if (!M) return;
    const values = rows.map(r => tfDvNum(r[ci]));
    const negative = values.some(v => v !== null && v < 0) && !values.some(v => v !== null && v > 0);
    const cfvo = negative ? [{ type: "num", value: -M }, { type: "num", value: 0 }] : [{ type: "num", value: 0 }, { type: "num", value: M }];
    const rule = argb => ({ type: "dataBar", gradient: false, border: false, minLength: 0, maxLength: 100, cfvo,
      color: { argb }, negativeFillColor: { argb }, negativeBarColorSameAsPositive: true, axisPosition: "auto" });
    const addr = i => worksheet.getCell(firstRow + i + 1, colAt(k) + 1).address;
    const colors = values.map((v, i) => v === null ? null : m.colorAt(i));
    const one = colors.filter(Boolean);
    if (one.length && one.every(x => x === one[0])) {
      worksheet.addConditionalFormatting(/** @type {any} */ ({ ref: `${addr(0)}:${addr(rows.length - 1)}`, rules: [rule(one[0])] }));
    } else {
      colors.forEach((argb, i) => {
        if (argb) worksheet.addConditionalFormatting(/** @type {any} */ ({ ref: addr(i), rules: [rule(argb)] }));
      });
    }
  });
}


/* ═══ export/images.js ════════════════════════════════════════════════════════════════════════════ */
/* Dashboard image objects (logos, icons) → pictures on the sheet, from the files packaged in the .twbx. */

/** @typedef {{ type: "png" | "jpeg" | "gif" | "svg" | null, width: number, height: number }} ImageInfo */

/* below this size (px) an image is a divider line or spacer, not a picture */
const MIN_IMAGE_PX = 8;

/**
 * The file's format and pixel size from its header: PNG IHDR, JPEG SOF, GIF screen descriptor, SVG
 * width / height / viewBox. type null = not a format Excel or the browser can draw.
 * @param {Uint8Array} data @param {string} [path] @returns {ImageInfo}
 */
function imageInfo(data, path = "") {
  const b = data || new Uint8Array(0);
  const u16 = i => (b[i] << 8) | b[i + 1], u32 = i => ((b[i] << 24) >>> 0) + (b[i + 1] << 16) + (b[i + 2] << 8) + b[i + 3];
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return { type: "png", width: u32(16), height: u32(20) };
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return { type: "gif", width: b[6] | (b[7] << 8), height: b[8] | (b[9] << 8) };
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    for (let i = 2; i + 9 < b.length;) {
      if (b[i] !== 0xff) { i++; continue; }
      const m = b[i + 1];
      // SOF0–SOF15 carry the frame size (not DHT C4, JPG C8, DAC CC)
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { type: "jpeg", width: u16(i + 7), height: u16(i + 5) };
      if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { i += 2; continue; }
      i += 2 + u16(i + 2);
    }
    return { type: "jpeg", width: 0, height: 0 };
  }
  const head = new TextDecoder().decode(b.subarray(0, 4096));
  if (/\.svg$/i.test(path) || /<svg[\s>]/i.test(head)) {
    const tag = (head.match(/<svg[^>]*>/i) || [""])[0];
    const attr = n => { const m = tag.match(new RegExp(`\\s${n}\\s*=\\s*["']([\\d.]+)(px)?["']`, "i")); return m ? +m[1] : 0; };
    const vb = (tag.match(/viewBox\s*=\s*["']\s*[-\d.]+[\s,]+[-\d.]+[\s,]+([\d.]+)[\s,]+([\d.]+)/i) || []).slice(1).map(Number);
    return { type: "svg", width: attr("width") || vb[0] || 0, height: attr("height") || vb[1] || 0 };
  }
  return { type: null, width: 0, height: 0 };
}

/**
 * How each image object can be drawn on the grid: "backdrop" (a background behind other objects – not
 * exported, it would cover the cells), "overlay" (an icon on top of a worksheet / text box: floats over
 * that block, host = its id), "tiny" (dividers, spacers) or "block" (a logo with its own zone).
 * Only visible objects count, as Tableau shows them.
 * @param {any[]} objects Tableau DashboardObjects
 * @returns {Map<string, { kind: "backdrop" | "overlay" | "tiny" | "block", host?: string }>}
 */
function classifyImageObjects(objects) {
  const box = o => ({ x: o.position.x, y: o.position.y, w: (o.size || {}).width || 0, h: (o.size || {}).height || 0 });
  const inside = (px, py, r) => px >= r.x && px <= r.x + r.w && py >= r.y && py <= r.y + r.h;
  const shown = (objects || []).filter(o => o && o.position && o.isVisible !== false);
  const others = shown.filter(o => /^(worksheet|text|quick-filter|filter|parameter-control|parameter|legend)$/.test(o.type));
  const out = new Map();
  shown.filter(o => o.type === "image").forEach(o => {
    const r = box(o);
    if (r.w < MIN_IMAGE_PX || r.h < MIN_IMAGE_PX) return out.set(String(o.id), { kind: "tiny" });
    if (others.some(p => { const q = box(p); return inside(q.x + q.w / 2, q.y + q.h / 2, r); })) return out.set(String(o.id), { kind: "backdrop" });
    const host = others.find(p => /^(worksheet|text)$/.test(p.type) && inside(r.x + r.w / 2, r.y + r.h / 2, box(p)));
    out.set(String(o.id), host ? { kind: "overlay", host: String(host.id) } : { kind: "block" });
  });
  return out;
}

/**
 * The picture's rectangle inside its box (px), as Tableau draws an image object: Fit Image scales it to
 * fit keeping its proportions, otherwise it keeps its own size (shrunk only when larger than the box);
 * Center Image (on unless switched off) centres it, otherwise it sits top-left.
 * @param {number} natW @param {number} natH @param {number} boxW @param {number} boxH
 * @param {boolean} [scaled] @param {boolean} [centered] undefined = on
 * @returns {{ x: number, y: number, w: number, h: number }}
 */
function fitImage(natW, natH, boxW, boxH, scaled, centered) {
  const nw = natW || boxW, nh = natH || boxH;
  const k = scaled ? Math.min(boxW / nw, boxH / nh) : Math.min(1, boxW / nw, boxH / nh);
  const w = Math.max(1, Math.round(nw * k)), h = Math.max(1, Math.round(nh * k));
  const center = centered !== false;
  return { x: center ? Math.round((boxW - w) / 2) : 0, y: center ? Math.round((boxH - h) / 2) : 0, w, h };
}

const B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

const B64_INDEX = (() => { const t = new Int16Array(128).fill(-1); for (let i = 0; i < 64; i++) t[B64.charCodeAt(i)] = i; return t; })();

/** base64 text (line breaks, padding allowed) → bytes @param {string} text @returns {Uint8Array} */
function fromBase64(text) {
  const s = String(text || ""), out = new Uint8Array(Math.ceil(s.length * 3 / 4));
  let n = 0, buf = 0, bits = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i), v = c < 128 ? B64_INDEX[c] : -1;
    if (v < 0) continue;                                         // white space, "="
    buf = ((buf << 6) | v) & 0xffffff; bits += 6;
    if (bits >= 8) { bits -= 8; out[n++] = (buf >> bits) & 255; }
  }
  return out.subarray(0, n);
}

/**
 * The custom shapes embedded in the workbook (<external><shapes>: icons used by Shape marks), keyed
 * "shape:<name>" next to the packaged images.
 * @param {string} xmlString the .twb XML @returns {Record<string, WorkbookImage>}
 */
function extractCustomShapes(xmlString) {
  /** @type {Record<string, WorkbookImage>} */
  const out = {};
  const at = String(xmlString || "").lastIndexOf("<external>");
  if (at < 0) return out;
  const unescape = s => s.replace(/&apos;/g, "'").replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  const re = /<shape name=(?:'([^']*)'|"([^"]*)")\s*>([\s\S]*?)<\/shape>/g;
  re.lastIndex = at;
  for (let m = re.exec(xmlString); m; m = re.exec(xmlString)) out["shape:" + unescape(m[1] ?? m[2])] = { data: fromBase64(m[3]) };
  return out;
}

/** @param {Uint8Array} bytes @returns {string} */
function toBase64(bytes) {
  let out = "";
  for (let i = 0; i < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | ((bytes[i + 1] || 0) << 8) | (bytes[i + 2] || 0);
    out += B64[(n >> 18) & 63] + B64[(n >> 12) & 63] + (i + 1 < bytes.length ? B64[(n >> 6) & 63] : "=") + (i + 2 < bytes.length ? B64[n & 63] : "=");
  }
  return out;
}

/* a canvas to draw on, when the page has one (the browser; not the Node test runs) */
function makeCanvas(w, h) {
  try {
    const c = typeof document !== "undefined" && document.createElement ? document.createElement("canvas") : null;
    if (!c || typeof c.getContext !== "function") return null;
    c.width = w; c.height = h;
    return c;
  } catch (e) {
    return null;
  }
}

/* the file drawn at w × h px into a PNG (SVG has to be rasterised for Excel; big photos are shrunk) */
async function rasterize(data, type, w, h) {
  const canvas = makeCanvas(w, h);
  if (!canvas || typeof Image === "undefined" || typeof URL === "undefined" || !URL.createObjectURL) return null;
  const url = URL.createObjectURL(new Blob([data], { type: type === "svg" ? "image/svg+xml" : "image/" + type }));
  try {
    const img = new Image();
    await new Promise((resolve, reject) => { img.onload = resolve; img.onerror = () => reject(new Error("image could not be decoded")); img.src = url; });
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0, w, h);
    const dataUrl = canvas.toDataURL("image/png");
    return dataUrl.slice(dataUrl.indexOf(",") + 1);
  } catch (e) {
    console.warn("[Images] could not draw the image:", e.message);
    return null;
  } finally {
    URL.revokeObjectURL(url);
  }
}

/**
 * The picture for Excel at its drawn size: PNG / JPEG / GIF as packaged; SVG rasterised; a large file
 * shown small (a 1 MB icon in a 40 px zone) re-drawn at twice its size so the workbook stays small.
 * null when it cannot be drawn here (SVG without a canvas, unknown format).
 * @param {{ data: Uint8Array }} file @param {ImageInfo} info @param {number} w @param {number} h drawn size, px
 * @returns {Promise<{ base64: string, extension: "png" | "jpeg" | "gif" } | null>}
 */
async function prepareImage(file, info, w, h) {
  const big = file.data.length > 150000 && info.width > 3 * w && info.height > 3 * h;
  if (info.type === "svg" || (big && info.type !== "gif")) {
    const png = await rasterize(file.data, info.type, Math.max(1, Math.round(w * 2)), Math.max(1, Math.round(h * 2)));
    if (png) return { base64: png, extension: "png" };
    if (info.type === "svg") return null;
  }
  return info.type ? { base64: toBase64(file.data), extension: /** @type {"png" | "jpeg" | "gif"} */ (info.type) } : null;
}

/** a URL Excel opens as a web link ("www.site.com/x" → "https://www.site.com/x"); null for anything else
 * @param {string | undefined} url @returns {string | null} */
function webLink(url) {
  const u = String(url || "").trim();
  if (!u || /^(javascript|data|file|vbscript):/i.test(u)) return null;
  if (/^(https?|mailto):/i.test(u)) return u;
  return /^[\w-]+(\.[\w-]+)+([/?#].*)?$/.test(u) ? "https://" + u : null;
}

/**
 * The sheet position px from the top-left of cell (col, row), as an anchor ExcelJS writes as is:
 * cell + offset in EMU (fractional anchors would be scaled by ExcelJS's own column width guess).
 * @param {number} col @param {number} row 0-based @param {number} x @param {number} y px
 * @param {(c: number) => number} colPx @param {(r: number) => number} rowPx final column widths / row heights
 */
function nativeAnchor(col, row, x, y, colPx, rowPx) {
  let c = col, r = row, dx = Math.max(0, x), dy = Math.max(0, y);
  for (let guard = 0; guard < 500 && colPx(c) > 0 && dx >= colPx(c); guard++) { dx -= colPx(c); c++; }
  for (let guard = 0; guard < 5000 && dy >= rowPx(r); guard++) { dy -= rowPx(r); r++; }
  return { nativeCol: c, nativeColOff: Math.round(dx * 9525), nativeRow: r, nativeRowOff: Math.round(dy * 9525) };
}


/* ═══ export/layout.js ════════════════════════════════════════════════════════════════════════════ */
/* Dashboard layout → Excel grid: positions, sizes and collision handling. */

/* ── Layout constants ─────────────────────────────────────────────────── */
const PX_PER_COL  = 90;

const PX_PER_ROW  = 20;          // one Excel row (15 pt): dashboard y maps 1:1 onto the sheet


const COL_GAP     = 0;

const ROW_GAP     = 0;            // Tableau zones touch; their padding is inside the zone

const ROW_GROUP_THRESHOLD = 8;

const EXCEL_ROW_PX = 20;

// default Excel row (15pt)
const EXCEL_COL_PX = 75;

// a 10-character column – the narrowest width setColumnWidths assigns
const CHART_DATA_SHEET = "Chart Data";

/* pixel box of a chart/image visual: the dashboard zone size minus Tableau's title strip */
function graphicBox(layout, vm) {
  const widthPx = Math.max(40, Math.round((layout && layout.widthPx) || 480));
  const zoneH = Math.round((layout && layout.heightPx) || 320);
  const titleRows = vm.showTitle ? 1 : 0;
  const heightPx = Math.max(60, zoneH - (titleRows ? 28 : 0));
  return { widthPx, heightPx, titleRows,
           // the zone's width on the dashboard grid, so the block lines up with the blocks above / below it
           gridW: layout && layout.widthPx ? layout.gridW : Math.max(2, Math.ceil(widthPx / EXCEL_COL_PX)),
           // the zone's own height on the sheet, so the block below starts where it does on the dashboard
           rows: Math.max(titleRows + 1, Math.round(zoneH / EXCEL_ROW_PX)) };
}

/* zone edges closer than this (px) are one column boundary */
const EDGE_MERGE_PX = 6;

/**
 * Excel columns from the dashboard's own zone edges: every left / right edge of a laid-out block is a column
 * boundary, so each column is as wide (px) as the gap between two edges and every block spans exactly its
 * zone – positions and widths as on the dashboard, whatever the content. A table's fields get their boundaries
 * at the widths the table draws (splitPx, from the zone's left edge; running into free space beside the zone,
 * and scaled down only when that is not enough); KPI tiles split their zone (split). A field or tile crossed by
 * another block's edge spans several columns.
 * @param {{ layout?: any, split?: number[], splitPx?: number[] }[]} items
 * @returns {{ colPx: number[], span: (item: any) => { gridCol: number, gridW: number, cuts: number[] | null,
 *   fields: { offset: number, span: number }[] | null } | null }}
 */
function buildColumnGrid(items) {
  const laid = items.filter(it => it.layout && typeof it.layout.xPx === "number" && it.layout.widthPx > 0);
  if (!laid.length) return { colPx: [], span: () => null };
  const minX = Math.min(...laid.map(it => it.layout.xPx));
  const merge = (list, tol = EDGE_MERGE_PX) => {
    const out = [];
    [...list].sort((a, b) => a - b).forEach(e => { if (!out.length || e - out[out.length - 1] > tol) out.push(e); });
    return out;
  };
  const sideOf = it => [it.layout.xPx - minX, it.layout.xPx - minX + it.layout.widthPx];
  let edges = merge(laid.flatMap(sideOf));
  const nearest = v => edges.reduce((best, e, i) => Math.abs(e - v) < Math.abs(edges[best] - v) ? i : best, 0);
  const extra = [];
  const rowsOverlap = (a, b) => a.layout.yPx < b.layout.yPx + (b.layout.heightPx || 0) - 2 &&
                                b.layout.yPx < a.layout.yPx + (a.layout.heightPx || 0) - 2;
  /* how far a table may run to the right: up to the next block beside it (sharing its rows) – a table's
     columns keep their widths as long as there is room, instead of being squeezed */
  const roomRight = (R, it) => {
    let limit = Infinity;
    laid.forEach(o => {
      if (o === it) return;
      const [l] = sideOf(o);
      if (l >= R - EDGE_MERGE_PX && rowsOverlap(it, o)) limit = Math.min(limit, Math.max(R, l));
    });
    return limit;
  };
  // A table's field boundaries are column edges too. Every block keeps its own edges: where another block's
  // edge falls inside a field, the field spans both columns (merged cells), so tables, KPI tiles and charts all
  // keep their exact widths on one shared set of columns.
  /** @type {Map<any, number[]>} */
  const fieldEdges = new Map();
  laid.forEach(it => {
    if (!it.splitPx || !it.splitPx.length) return;
    const [l, r] = sideOf(it), L = edges[nearest(l)], R = edges[nearest(r)];
    const total = it.splitPx.reduce((a, b) => a + b, 0);
    const room = roomRight(R, it) - L;                                  // squeezed only when that is not enough
    const k = total > room ? room / total : 1;
    let at = L;
    const own = [L];
    it.splitPx.forEach(w => { at += w * k; own.push(Math.round(at)); });
    extra.push(...own);
    fieldEdges.set(it, own);
  });
  // KPI tiles: Tableau gives each the same share of the zone – their boundaries are column edges, whatever other
  // blocks' edges fall in between (a tile then spans several columns)
  /** @type {Map<any, number[]>} */
  const splitCuts = new Map();
  laid.forEach(it => {
    if (it.splitPx && it.splitPx.length) return;
    if (!it.split || it.split.length < 2) return;
    const [l, r] = sideOf(it), L = edges[nearest(l)], R = edges[nearest(r)];
    const total = it.split.reduce((a, b) => a + b, 0);
    let at = L;
    const cuts = it.split.slice(0, -1).map(w => { at += (R - L) * w / total; return Math.round(at); });
    extra.push(...cuts);
    splitCuts.set(it, cuts);
  });
  edges = merge([...edges, ...extra], 2);
  const colPx = edges.slice(1).map((e, i) => e - edges[i]);
  return {
    colPx,
    span: it => {
      if (!laid.includes(it)) return null;
      const [l, r] = sideOf(it), a = nearest(l), b = nearest(r);
      // split blocks: the column (from the block's first) where each next part starts
      const cuts = splitCuts.has(it) ? splitCuts.get(it).map(e => nearest(e) - a) : null;
      // tables: each field's first column (from the block's first) and how many columns it spans
      let fields = null;
      if (fieldEdges.has(it)) {
        const idx = fieldEdges.get(it).map(nearest);
        fields = idx.slice(0, -1).map((c, k) => ({ offset: c - idx[0], span: Math.max(1, idx[k + 1] - c) }));
        fields.forEach((f, k) => { if (k && f.offset < fields[k - 1].offset + fields[k - 1].span) f.offset = fields[k - 1].offset + fields[k - 1].span; });
      }
      return { gridCol: a, gridW: Math.max(1, b - a), cuts, fields };
    }
  };
}

function getExcelColumnName(colIndex) {
  let columnName = "";
  let dividend = colIndex + 1;

  while (dividend > 0) {
    let modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

/* ── Range tracker ────────────────────────────────────────────────────── */
function makeRangeTracker() {
  let minR = Infinity, minC = Infinity, maxR = 0, maxC = 0;
  return {
    update(r, c) {
      if (r < minR) minR = r;
      if (c < minC) minC = c;
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    },
    toRange() {
      return {
        s: { r: minR === Infinity ? 0 : minR, c: minC === Infinity ? 0 : minC },
        e: { r: maxR, c: maxC },
      };
    },
  };
}

/* edges a few px apart (tiled containers with padding) → one grid line, so blocks that line up on
 * the dashboard land in the same row / column */
const SNAP_PX = 8;

/** @param {number[]} values @param {number} [tolerance] @returns {(v: number) => number} */
function snapEdges(values, tolerance = SNAP_PX) {
  const anchor = new Map();
  let current = null;
  [...new Set(values)].sort((a, b) => a - b).forEach(v => {
    if (current === null || v - current > tolerance) current = v;
    anchor.set(v, current);
  });
  return v => anchor.has(v) ? anchor.get(v) : v;
}

/* ── buildLayoutMap ─────────────────────────────────────────────────── */
/* Tableau's DashboardObject: position = { x, y }, size = { width, height } (px) */
/* Tableau's DashboardObjectType → the kinds laid out on the sheet ("filter" / "parameter" as older
 * test fixtures spell them); a DashboardObject's id is its zone id in the workbook */
const OBJECT_KIND = Object.freeze({ worksheet: "worksheet", "quick-filter": "filter", filter: "filter",
                                           "parameter-control": "parameter", parameter: "parameter", text: "text", image: "image" });

function buildLayoutMap(dashboardObjects, titleMap = {}) {
  const map = new Map();
  // images with their own zone (logos) take part in the layout; backgrounds, dividers and icons
  // floating over a sheet do not (export/images.js)
  const images = classifyImageObjects(dashboardObjects || []);

  const positionableObjects = (dashboardObjects || []).filter(
    (obj) => OBJECT_KIND[obj.type] &&
             obj.position &&
             typeof obj.position.x === "number" &&
             typeof obj.position.y === "number" &&
             (obj.type !== "image" || (images.get(String(obj.id)) || {}).kind === "block")
  );

  if (positionableObjects.length === 0) return map;

  let minX = Infinity, minY = Infinity;
  positionableObjects.forEach((obj) => {
    if (obj.position.x < minX) minX = obj.position.x;
    if (obj.position.y < minY) minY = obj.position.y;
  });

  const sizeOf = obj => obj.size || obj.position;
  const snapX = snapEdges(positionableObjects.flatMap(o => [o.position.x, o.position.x + (sizeOf(o).width || 0)]));
  const snapY = snapEdges(positionableObjects.map(o => o.position.y));

  positionableObjects.forEach((obj) => {
    const px = obj.position;
    const { width, height } = sizeOf(obj);
    const gridCol = Math.round((snapX(px.x) - minX) / PX_PER_COL);
    const gridRow = Math.round((snapY(px.y) - minY) / PX_PER_ROW);
    const gridW = Math.max(2, Math.round((width || 180) / PX_PER_COL));
    const gridH = Math.max(3, Math.round((height || 60) / PX_PER_ROW));
    
    const kind = OBJECT_KIND[obj.type];
    let displayName = "";
    if (kind === "text" || kind === "image") displayName = "";
    else if (kind === "worksheet") {
      // a title written with Tableau's <Sheet Name> token names the sheet itself
      displayName = (titleMap[obj.name] && titleMap[obj.name].replace(/<sheet name>/gi, obj.name))
                 || (obj.title && obj.title.trim() ? obj.title.trim() : null)
                 || obj.name;
    } else {
      displayName = (obj.name || "Filter").replace(/[_-]/g, " ");
    }

    // text boxes and images have no unique name – keyed by their zone id
    const key = kind === "text" || kind === "image" ? `${kind}:${obj.id}` : obj.name || `filter_${gridRow}_${gridCol}`;
    // a quick filter can carry its worksheet's name – the worksheet keeps its own position and size
    if (kind !== "worksheet" && map.has(key) && map.get(key).type === "worksheet") return;
    map.set(key, {
      type: kind,
      id: obj.id,
      gridRow: Math.max(0, gridRow),
      gridCol: Math.max(0, gridCol),
      gridW,
      // right edge on the grid: tiles in a strip end where the next one starts
      gridRight: width ? Math.round((snapX(px.x + width) - minX) / PX_PER_COL) : gridCol + gridW,
      gridH,
      widthPx: width || null,
      heightPx: height || null,
      xPx: px.x,
      yPx: px.y,
      displayName,
      originalName: obj.name
    });
  });

  return map;
}

/* ── grouped table rows ─────────────────────────────────────────────────
 * A long table shows its first ROW_GROUP_THRESHOLD rows and collapses the rest into an outline group.
 * Excel hides whole sheet rows, so a chart / image / card beside the table would lose the same rows:
 * the table keeps visible every row such a neighbour uses (it then fills the band like the Tableau
 * zone) and only the rows below are collapsed. Tables beside each other keep their own threshold. */
function isGroupedTable(item) {
  return !!(FORMAT_CONFIG.groupOverflowRows && item.vm && item.type === "worksheet" && !item.isKPI && !item.box &&
            item.vm.rows.length > ROW_GROUP_THRESHOLD);
}

/**
 * Sheet rows a table's header takes: the height Tableau stores for it in 20 px rows (merged down), so a tall
 * header keeps the sheet's rows at 20 px – a single tall row would push the blocks beside the table down.
 * @param {ViewModel} vm
 */
function headerRowSpan(vm) {
  if (!vm.showHeaderRow) return 0;
  // Tableau's stored header height, else the height of the dashboard text boxes drawn as the headers
  const hpx = (vm.fmt && vm.fmt.headerRowHeightPx ? vm.fmt.headerRowHeightPx() : 0) || vm.headerPx || 0;
  return hpx ? Math.max(1, Math.round(hpx / PX_PER_ROW)) : 1;
}

/* sheet row of a table's first data row */
function tableDataStart(item) {
  return item.gridRow + (item.vm.showTitle ? 1 : 0) + headerRowSpan(item.vm);
}

/** sets item.visibleRows on every table that will be grouped @param {ExportItem[]} items */
function setTableVisibleRows(items) {
  items.forEach(t => {
    if (!isGroupedTable(t)) return;
    const start = tableDataStart(t);
    const end = t.gridRow + t.allocatedRows;
    let need = ROW_GROUP_THRESHOLD;
    // the rows Tableau shows in the zone (at the sheet's own row height) stay visible
    if (t.layout && t.layout.heightPx) {
      const headPx = ((t.vm.showTitle ? 1 : 0) + headerRowSpan(t.vm)) * PX_PER_ROW;
      need = Math.max(need, Math.floor((t.layout.heightPx - headPx) / (t.vm.fmt.rowHeightPx() || PX_PER_ROW)));
    }
    items.forEach(o => {
      if (o === t || isGroupedTable(o)) return;
      const oEnd = o.gridRow + (o.allocatedRows || o.rowCount || 0);
      if (o.gridRow < end && oEnd > t.gridRow) need = Math.max(need, oEnd - start);   // shares sheet rows
    });
    t.visibleRows = Math.min(t.vm.rows.length, need);
  });
}

/* rows the writer will produce – used for layout before writing */
function viewModelHeight(vm) {
  const grouped = FORMAT_CONFIG.groupOverflowRows && vm.rows.length > ROW_GROUP_THRESHOLD;
  return (vm.showTitle ? 1 : 0) + headerRowSpan(vm) + vm.rows.length + (grouped ? 1 : 0);
}

function resolveCollisions(zones) {
if (!zones || zones.length === 0) return zones;

// Full row/header/note height a zone actually occupies on the sheet,
// including its own hidden-but-physically-present grouped rows.
function getVisualHeight(zone) {
  return (zone.allocatedRows || zone.rowCount || 5) + ROW_GAP;
}

// ── Group by EXACT gridRow (not a rounded bucket) ──
const rowGroups = new Map();
zones.forEach(zone => {
  const key = zone.gridRow;
  if (!rowGroups.has(key)) rowGroups.set(key, []);
  rowGroups.get(key).push(zone);
});

// ── Horizontal packing within each group ──
const processedGroups = [];
for (const [, group] of rowGroups) {
  group.sort((a, b) => a.gridCol - b.gridCol);
  const minRow = Math.min(...group.map(z => z.gridRow));
  group.forEach(z => { z.gridRow = minRow; });
  let cursor = group[0].gridCol;
  for (const item of group) {
    if (item.gridCol < cursor) item.gridCol = cursor;
    cursor = item.gridCol + item.gridW + COL_GAP;
  }
  const groupBottom = minRow + Math.max(...group.map(z => getVisualHeight(z))) + ROW_GAP;
  processedGroups.push({ minRow, items: group, bottom: groupBottom });
}

// ── Vertical pushes: a row group moves down as a whole, below every earlier group it overlaps
//    (not only the group just above), so blocks that share a top edge on the dashboard stay level ──
processedGroups.sort((a, b) => a.minRow - b.minRow);
const groupsOverlap = (upper, lower) => upper.items.some(u => lower.items.some(l =>
  u.gridCol < l.gridCol + l.gridW && u.gridCol + u.gridW > l.gridCol));
for (let i = 1; i < processedGroups.length; i++) {
  const lower = processedGroups[i];
  let pushBy = 0;
  for (let j = 0; j < i; j++) {
    if (groupsOverlap(processedGroups[j], lower)) pushBy = Math.max(pushBy, processedGroups[j].bottom - lower.minRow);
  }
  if (pushBy > 0) {
    lower.items.forEach(z => { z.gridRow += pushBy; });
    lower.minRow += pushBy;
    lower.bottom += pushBy;
  }
}

// ── Full pairwise safety-net scan across ALL zones, not just adjacent groups ──
let fullPassChanged = true;
const MAX_PASSES = 20;
let pass = 0;
while (fullPassChanged && pass < MAX_PASSES) {
  fullPassChanged = false;
  pass++;
  for (let i = 0; i < zones.length; i++) {
    for (let j = i + 1; j < zones.length; j++) {
      const a = zones[i], b = zones[j];
      const upper = a.gridRow <= b.gridRow ? a : b;
      const lower = a.gridRow <= b.gridRow ? b : a;
      const upperBottom = upper.gridRow + getVisualHeight(upper) + ROW_GAP;
      const colOverlap = upper.gridCol < lower.gridCol + lower.gridW &&
                          upper.gridCol + upper.gridW > lower.gridCol;
      const rowOverlap = lower.gridRow < upperBottom;
      if (colOverlap && rowOverlap) {
        lower.gridRow += (upperBottom - lower.gridRow);
        fullPassChanged = true;
      }
    }
  }
}

return zones;
}


/* ═══ export/visual-writers.js ════════════════════════════════════════════════════════════════════ */
/* Tables and KPI cards written with Tableau formatting; widths and autofilters. */

/* =============================================================================
 * TABLEAU-FORMATTED WRITERS
 * All styling below comes from the TWB format model (createSheetFormatter).
 * If no workbook was loaded, the formatter falls back to TABLEAU_DEFAULTS.
 * ============================================================================= */

/**
 * A colour legend at its dashboard position, as Tableau draws it: the field caption in bold, then a coloured
 * square and the value per item.
 * @param {import("exceljs").Worksheet} worksheet @param {ExportItem} item @param {any} rangeTracker
 * @param {Record<number, number>} colWidths
 */
function writeLegendBlock(worksheet, item, rangeTracker, colWidths) {
  const r = item.gridRow, C = item.gridCol;
  const font = tfExcelFont(item.fmt.baseStyle());
  const title = worksheet.getCell(r + 1, C + 1);
  title.value = item.legend.title;
  title.font = { ...font, bold: true };
  rangeTracker.update(r, C);
  item.legend.items.forEach((it, i) => {
    worksheet.getCell(r + 2 + i, C + 1).value = { richText: [
      { text: "■ ", font: { ...font, size: (font.size || 9) + 3, color: { argb: "FF" + (it.color || "CCCCCC") } } },
      { text: it.text, font }
    ] };
    rangeTracker.update(r + 1 + i, C);
  });
  if (!item.layout || !item.layout.widthPx) {                 // no zone width to keep: fit the longest entry
    const chars = Math.max(item.legend.title.length, ...item.legend.items.map(it => it.text.length + 2));
    colWidths[C] = Math.max(colWidths[C] || 0, chars + 2);
  }
}

function writeTableauTitle(worksheet, r, C, text, p, span) {
  const cell = worksheet.getCell(r + 1, C + 1);
  cell.value = text;
  cell.font = tfExcelFont(p);
  cell.alignment = { horizontal: p.hAlign || "left", vertical: "middle", wrapText: false };
  if (p.bgColor) cell.fill = tfExcelFill(p.bgColor);
  if (span > 1) worksheet.mergeCells(r + 1, C + 1, r + 1, C + span);
  if ((p.fontSize || 0) > 11) worksheet.getRow(r + 1).height = Math.round(p.fontSize * 1.6);
}

/* Which columns get Tableau's mark colour, and what colour each row gets.
 * Tableau colours the MARKS of the pane that holds the Color encoding. */
function buildColorPlan(fmt, colInfo, rows) {
  const enc = fmt.colorEncoding();
  const textRefs = fmt.textRefs();
  let markIdx = colInfo.map((c, i) => (!c.isHeader && (c.pivoted || (c.ref && textRefs.some(t => tfSameField(t, c.ref))))) ? i : -1).filter(i => i >= 0);
  if (!markIdx.length) markIdx = colInfo.map((c, i) => c.isHeader ? -1 : i).filter(i => i >= 0);
  if (!enc) return { enc: null, markIdx: new Set(markIdx), colorAt: () => null };

  if (enc.paneRefs && enc.paneRefs.length) {                    // scope to the colour's own pane
    const scoped = markIdx.filter(i => colInfo[i].pivoted
      ? enc.paneRefs.some(r => r.name === "Multiple Values")
      : enc.paneRefs.some(r => tfSameField(r, colInfo[i].ref)));
    if (scoped.length) markIdx = scoped;
  }
  let colorIdx = colInfo.findIndex(c => c.ref && tfSameField(c.ref, enc.ref));
  if (colorIdx < 0) {
    console.log(`[Format] Colour field ${enc.ref.inner} not present in summary data – no mark colours`);
    return { enc, markIdx: new Set(markIdx), colorAt: () => null };
  }
  const pick = dv => tfIsNull(dv) ? null
    : enc.continuous ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : tfDvText(dv);
  const vals = rows.map(row => pick(row[colorIdx]));
  const scale = tfBuildColorScale(fmt, enc, vals);
  const colors = scale ? vals.map(v => scale(v)) : [];
  return { enc, markIdx: new Set(markIdx), colorAt: i => colors[i] || null };
}

function borderSide(d) { return d && d.visible && d.style ? { style: d.style, color: { argb: d.color } } : undefined; }

function writeKPICardStacked(worksheet, vm, originRow, originCol, rangeTracker) {
  let r = originRow;
  const C = originCol;
  const { fmt, cols, rows, order } = vm;
  const plan = buildColorPlan(fmt, cols, rows);
  if (vm.showTitle) {
    writeTableauTitle(worksheet, r, C, vm.title.text, vm.title.props, 2);
    rangeTracker.update(r, C); rangeTracker.update(r, C + 1);
    r++;
  }
  order.forEach((ci, k) => {
    const info = cols[ci];
    tfWriteCell(worksheet.getCell(r + k + 1, C + 1), { formattedValue: info.label, value: info.label }, fmt.fieldLabelStyle(info.ref, true));
    const vp = fmt.markCellStyle(info.ref);
    const color = plan.markIdx.has(ci) ? plan.colorAt(0) : null;
    const extra = {};
    if (color && plan.enc.applyTo === "fill") extra.fill = color;
    if (color && plan.enc.applyTo === "font" && !vp.explicitColor) extra.fontColor = color;
    tfWriteCell(worksheet.getCell(r + k + 1, C + 2), rows[0][ci], vp, extra);
    rangeTracker.update(r + k, C); rangeTracker.update(r + k, C + 1);
  });
  return (vm.showTitle ? 1 : 0) + order.length;
}

/**
 * Each table column's width in px, as the table is drawn: the width Tableau stores for the field (or for
 * Measure Names, or its text-box header), else an estimate from its label and values in their font.
 * @param {ViewModel} vm @returns {number[]} one per vm.order entry
 */
function tableColumnPx(vm) {
  const { fmt, cols, rows, order } = vm;
  return order.map(ci => {
    const info = cols[ci];
    const measureNamesRef = info.pivoted ? tfParseFieldRef("[:Measure Names]") : null;
    const px = fmt.widthPx(info.ref) || (measureNamesRef && fmt.widthPx(measureNamesRef)) || info.zoneWidthPx;
    if (px) return px;
    const style = info.isHeader ? fmt.headerCellStyle(info.ref) : fmt.markCellStyle(info.ref);
    const scale = (style.fontSize || 9) / 11;
    let len = Math.min((info.label || "").length, 24);
    for (let i = 0; i < Math.min(rows.length, 200); i++) {
      len = Math.max(len, tfDvText(rows[i][ci]).split("\n").reduce((m, l) => Math.max(m, l.length), 0));
    }
    return Math.min(50, Math.ceil(len * scale * 1.15) + 2) * 7;
  });
}

/**
 * @param {number} [visibleRows] rows kept visible before the rest is grouped
 * @param {{ offset: number, span: number }[] | null} [fieldCols] each field's first column (from originCol) and how
 *   many columns it spans on the dashboard grid; null = one column per field
 */
function writeRegularTable(worksheet, vm, originRow, originCol, rangeTracker, allTablesInfo, colWidths, exactWidths,
                                  visibleRows = ROW_GROUP_THRESHOLD, fieldCols = null) {
  let r = originRow;
  const C = originCol;
  const { fmt, cols, rows, order } = vm;
  const numCols = order.length;
  // field k's first sheet column, and its last: a field crossed by another block's column edge spans both columns
  const grid = fieldCols && fieldCols.length === numCols ? fieldCols : null;
  const colAt = k => C + (grid ? grid[k].offset : k);
  const lastAt = k => colAt(k) + (grid ? grid[k].span : 1) - 1;
  const lastCol = lastAt(numCols - 1);
  const wideFields = !!grid && grid.some(f => f.span > 1);
  if (!fmt.hasModel) console.log(`[Format] No TWB format info for "${vm.title.text}" – using Tableau defaults`);

  const plan = buildColorPlan(fmt, cols, rows);
  // a table built from marks: each column looks like its pane's marks (fills, arrows, data bars)
  const marks = vm.markTable ? buildMarkPlan(fmt, cols, rows, order) : new Map();
  const tableBg = fmt.tableBackground();                       // Format → Shading → Worksheet
  const rowDiv = fmt.divider("rows");
  const colDiv = fmt.divider("cols");
  const band = fmt.banding();
  const headerIdx = vm.headerOrder.length ? vm.headerOrder
                  : cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);

  // ── Title (respects the dashboard zone's "Show title") ──
  if (vm.showTitle) {
    const tp = vm.title.props.bgColor ? vm.title.props : { ...vm.title.props, bgColor: tableBg };
    writeTableauTitle(worksheet, r, C, vm.title.text, tp, lastCol - C + 1);
    rangeTracker.update(r, C);
    rangeTracker.update(r, lastCol);
    r++;
  }

  // ── Header row: Measure Names aliases / captions; hidden field labels stay blank ──
  const styles = {};
  order.forEach(ci => { styles[ci] = cols[ci].isHeader ? fmt.headerCellStyle(cols[ci].ref) : fmt.markCellStyle(cols[ci].ref); });
  // ── Column widths: Tableau pixel width → text-box header width → content estimate ──
  const widthOf = {};
  order.forEach((ci, k) => {
    const info = cols[ci];
    const excelCol = colAt(k);
    const measureNamesRef = info.pivoted ? tfParseFieldRef("[:Measure Names]") : null;
    const px = fmt.widthPx(info.ref) || (measureNamesRef && fmt.widthPx(measureNamesRef)) || info.zoneWidthPx;
    if (px) {
      widthOf[ci] = Math.round(px / 7 * 10) / 10;
      exactWidths[excelCol] = Math.max(exactWidths[excelCol] || 0, widthOf[ci]);
      return;
    }
    const scale = (styles[ci].fontSize || 9) / 11;
    let len = Math.min(info.label.length, 24);
    for (let i = 0; i < Math.min(rows.length, 200); i++) {
      const longestLine = tfDvText(rows[i][ci]).split("\n").reduce((m, l) => Math.max(m, l.length), 0);
      len = Math.max(len, longestLine);
    }
    widthOf[ci] = Math.min(50, Math.ceil(len * scale * 1.15) + 2);
    colWidths[excelCol] = Math.max(colWidths[excelCol] || 0, widthOf[ci]);
  });

  const numericCol = ci => { const d = rows.find(rw => !tfIsNull(rw[ci])); return !!d && typeof (d[ci].nativeValue !== undefined ? d[ci].nativeValue : d[ci].value) === "number"; };

  let headerRow = r;
  if (vm.showHeaderRow) {
    const span = headerRowSpan(vm);
    /** @type {[number, number, number, number][]} merged after every header cell is written */
    const merges = [];
    order.forEach((ci, k) => {
      const info = cols[ci];
      let p = fmt.fieldLabelStyle(info.ref, !info.isHeader);
      if (info.labelProps) p = tfMerge(p, { bold: false }, info.labelProps);      // dashboard text box font
      // header label sits over its column: same alignment as the values (numbers right, text left)
      if (!p.hAlign) p.hAlign = info.isHeader ? styles[ci].hAlign : (styles[ci].hAlign || (numericCol(ci) ? "right" : "left"));
      if (/\n/.test(info.label) || info.label.length > 18) p.wrap = true;
      const cell = worksheet.getCell(r + 1, colAt(k) + 1);
      tfWriteCell(cell, { formattedValue: info.label, value: info.label }, p, {
        fill: p.bgColor || tableBg || undefined,
        border: { bottom: borderSide(rowDiv), right: k < numCols - 1 ? borderSide(colDiv) : undefined }
      });
      // a dashboard text box as the header: its runs (fonts, sizes, line breaks) as rich text
      if (info.labelRuns && info.labelRuns.length && !info.headerCovered) {
        cell.value = { richText: tfRichRuns(info.labelRuns, p) };
        cell.alignment = { ...(cell.alignment || {}), wrapText: true };
      }
      // Tableau's header height as 20 px rows merged down, and a header over several panes merged across;
      // the label wraps in it, centred as Tableau draws it (which also keeps it clear of the filter buttons
      // on the last row)
      const across = info.headerSpan || 1;
      const to = lastAt(Math.min(numCols - 1, k + across - 1));     // its last sheet column
      if (!info.headerCovered && (span > 1 || to > colAt(k))) {
        if (span > 1 || across > 1) cell.alignment = { ...(cell.alignment || {}), wrapText: true };
        for (let i = 0; i < span; i++) for (let c = colAt(k); c <= to; c++) {
          if (i || c > colAt(k)) worksheet.getCell(r + 1 + i, c + 1).border = cell.border;
        }
        merges.push([r + 1, colAt(k) + 1, r + span, to + 1]);
      }
      if (span > 1) for (let i = 0; i < span; i++) worksheet.getRow(r + 1 + i).height = 15;
      rangeTracker.update(r + span - 1, lastAt(k));
    });
    merges.forEach(m => worksheet.mergeCells(...m));
    // wrapped labels and no height stored: at most two lines, not a row Excel grows to fit every word
    if (span === 1 && !fmt.headerRowHeightPx() && order.some(ci => /\n/.test(cols[ci].label) || cols[ci].label.length > 18)) {
      worksheet.getRow(r + 1).height = 26;
    }
    headerRow = r + span - 1;                                // the filter buttons go on the header's last row
    r += span;
  }

  const dataStartRow = r;
  const totalRows = rows.length;
  const keepRows = Math.max(ROW_GROUP_THRESHOLD, visibleRows || 0);   // rows a block beside the table uses stay visible
  const needsGrouping = FORMAT_CONFIG.groupOverflowRows && totalRows > keepRows;
  const level = Math.max(1, rowDiv.level || 1);
  const rowHeightPx = fmt.rowHeightPx();

  rows.forEach((row, rowIdx) => {
    const next = rows[rowIdx + 1];
    const dividerHere = !next || !headerIdx.length || headerIdx.slice(0, level).some(hi =>
      tfDvText(row[hi]) !== tfDvText(next[hi]));
    const banded = rowIdx % 2 === 1;                 // banding per VISUAL row (after pivot/merge)
    const rowColor = plan.colorAt(rowIdx);

    order.forEach((ci, k) => {
      let p = styles[ci];
      const extra = {
        border: {
          bottom: dividerHere ? borderSide(rowDiv) : undefined,
          right: k < numCols - 1 ? borderSide(colDiv) : undefined
        }
      };
      // fill precedence: mark colour > own shading > banding > worksheet background
      const bandFill = banded ? (cols[ci].isHeader ? band.header : band.pane) : null;
      extra.fill = p.bgColor || bandFill || tableBg || undefined;
      const mk = marks.get(ci);
      if (rowColor && plan.markIdx.has(ci) && !mk) {
        if (plan.enc.applyTo === "fill") extra.fill = rowColor;
        else if (!p.explicitColor) extra.fontColor = rowColor;
      }
      const cellColor = vm.cellFill ? vm.cellFill(rowIdx, ci) : null;     // a matrix: each cell its own mark's colour
      if (cellColor) extra.fill = cellColor;
      let dv = row[ci];
      const text = tfDvText(dv);
      if (mk && !tfIsNull(dv)) {
        const color = mk.colorAt(rowIdx);
        if (mk.kind === "fill" && color) extra.fill = color;              // circle / square: the cell is the mark
        if (mk.kind === "glyph") {                                        // arrow shape before the label, in its colour
          if (color && !p.explicitColor) extra.fontColor = color;
          const glyph = mk.glyphAt(rowIdx);
          const native = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
          if (glyph && typeof native === "number") p = { ...p, numFmt: tvGlyphNumFmt(p.numFmt || inferExcelNumFmt(text, native) || "General", glyph) };
          else if (glyph) dv = { ...dv, formattedValue: `${glyph} ${text}` };
        }
      }
      if (cols[ci].link) {                                      // URL action → clickable cell
        const url = cols[ci].link.expression.replace(/<([^<>]+)>/g, (m, inner) => {
          const part = cols[ci].link.parts.find(x => ("[" + x.token.inner + "]") === inner || x.token.raw === inner);
          return part && part.ci >= 0 ? tfDvText(row[part.ci]) : "";
        }).trim();
        if (/^(https?:|mailto:|ftp:)/i.test(url)) {
          extra.hyperlink = url;
          extra.hyperlinkText = FORMAT_CONFIG.linkText === "url" ? url : FORMAT_CONFIG.linkText;
          p = { ...p, underline: true, color: p.explicitColor ? p.color : "FF0563C1" };
        }
      }
      const pp = /\n/.test(text) ? { ...p, wrap: true } : p;    // wrap only on real line breaks
      const cell = worksheet.getCell(r + 1, colAt(k) + 1);
      tfWriteCell(cell, dv, pp, extra);
      if (lastAt(k) > colAt(k)) {                                // a field over several columns: one merged cell
        for (let c = colAt(k) + 1; c <= lastAt(k); c++) worksheet.getCell(r + 1, c + 1).style = cell.style;
        worksheet.mergeCells(r + 1, colAt(k) + 1, r + 1, lastAt(k) + 1);
      }
      rangeTracker.update(r, lastAt(k));
    });

    const excelRow = worksheet.getRow(r + 1);
    if (rowHeightPx) excelRow.height = Math.round(rowHeightPx * 0.75);   // only a height Tableau stores
    if (needsGrouping && rowIdx >= keepRows) {
      excelRow.outlineLevel = 1;
      excelRow.hidden = true;
    }
    r++;
  });

  if (marks.size) writeMarkBars(worksheet, marks, cols, rows, order, dataStartRow, colAt);

  if (needsGrouping) {
    const hiddenCount = totalRows - keepRows;
    const noteCell = worksheet.getCell(r + 1, C + 1);
    noteCell.value = `${hiddenCount} rows are hidden — use the row group controls [+] / [-] on the left to expand or collapse`;
    noteCell.font = { italic: true, size: 9, color: { argb: "FF888888" } };
    noteCell.alignment = { horizontal: "left", vertical: "middle" };
    rangeTracker.update(r, C);
    r++;
    worksheet.properties.outlineProperties = { summaryBelow: false, summaryRight: false };
  }

  if (!plan.enc && FORMAT_CONFIG.applyFallbackHeatmap) {
    applyConditionalFormattingToTable(worksheet, rows.map(rw => order.map(ci => rw[ci])),
      order.map(ci => ({ fieldName: cols[ci].name })), dataStartRow, C, {});
  }

  // no filter buttons on a table built from marks, nor on merged fields (a button would sit inside the field)
  if (allTablesInfo && !vm.markTable && !wideFields) {
    allTablesInfo.push({ name: vm.title.text, headerRow: vm.showHeaderRow ? headerRow : dataStartRow,
      leftCol: C, rightCol: lastCol,
      bottomRow: dataStartRow + totalRows - 1,             // last DATA row – the grouping note stays outside the filter
      rowCount: totalRows, hasHeader: vm.showHeaderRow });
  }
  return r - originRow;
}

function applyAutoFilters(worksheet, allTablesInfo) {
  if (FORMAT_CONFIG.autoFilter === "none" || !allTablesInfo || !allTablesInfo.length) return;
  // Excel allows a single autofilter per sheet → put it on the largest table that has a header row
  const t = allTablesInfo.filter(x => x.hasHeader).sort((a, b) => b.rowCount - a.rowCount)[0];
  if (!t) return;
  worksheet.autoFilter = `${getExcelColumnName(t.leftCol)}${t.headerRow + 1}:${getExcelColumnName(t.rightCol)}${t.bottomRow + 1}`;
}

function setColumnWidths(worksheet, colWidths, exactWidths = {}) {
  const keys = [...Object.keys(colWidths), ...Object.keys(exactWidths)].map(Number);
  const maxColIdx = keys.length ? Math.max(...keys) : 0;
  for (let ci = 0; ci <= maxColIdx; ci++) {
    if (exactWidths[ci]) { worksheet.getColumn(ci + 1).width = exactWidths[ci]; continue; }
    worksheet.getColumn(ci + 1).width = Math.min((colWidths[ci] || 10), 50);
  }
}


/* ═══ export/cell-writers.js ══════════════════════════════════════════════════════════════════════ */
/* Low-level cell writers: titles, headers, data cells, filter tables. */

/* ── Quick filter / parameter control ───────────────────────────────────
 * As Tableau draws one: its caption, then the value(s) in a white box with a thin grey border, across the
 * control's zone. */
function writeIndividualFilterTable(worksheet, filterName, filterValues, originRow, originCol, rangeTracker, width = 1) {
  const C = originCol, W = Math.max(1, width);
  const cleanName = filterName.replace(/_(Filter|filter)_\d+$/, "").replace(/_(Values|values)_\d+$/, "");
  const font = tfExcelFont({ ...TABLEAU_DEFAULTS.worksheet, fontSize: 9, color: "FF333333" });     // Tableau fonts → their substitute
  const span = r => { if (W > 1) { try { worksheet.mergeCells(r + 1, C + 1, r + 1, C + W); } catch (e) { /* already merged */ } } };
  const caption = worksheet.getCell(originRow + 1, C + 1);
  caption.value = cleanName;
  caption.font = { ...font, bold: true };
  caption.alignment = { horizontal: "left", vertical: "bottom" };
  span(originRow);
  rangeTracker.update(originRow, C);
  rangeTracker.update(originRow, C + W - 1);
  const values = Array.isArray(filterValues) ? filterValues : [filterValues];
  /** @type {Partial<import("exceljs").Border>} */
  const edge = { style: "thin", color: { argb: "FFBFBFBF" } };
  values.forEach((value, idx) => {
    const r = originRow + 1 + idx;
    for (let c = C; c < C + W; c++) {
      const cell = worksheet.getCell(r + 1, c + 1);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cell.border = { top: idx === 0 ? edge : undefined, bottom: idx === values.length - 1 ? edge : undefined,
                      left: c === C ? edge : undefined, right: c === C + W - 1 ? edge : undefined };
    }
    const cell = worksheet.getCell(r + 1, C + 1);
    cell.value = value;
    cell.font = font;
    cell.alignment = { horizontal: "left", vertical: "middle", shrinkToFit: true };
    span(r);
    rangeTracker.update(r, C + W - 1);
  });
  return 1 + values.length;
}

function writeDashboardTitle(worksheet, dashboardName, originRow, originCol, rangeTracker, titleProps, titleRuns) {
  let r = originRow;
  const C = originCol;
  const p = titleProps || tfMerge(TABLEAU_DEFAULTS.dashTitle);
  writeTableauTitle(worksheet, r, C, dashboardName, p, 5);
  if (titleRuns && titleRuns.length) {                    // text-box title → rich text, one run per Tableau run
    const rich = tfRichRuns(titleRuns, tfMerge(TABLEAU_DEFAULTS.worksheet));
    const cell = worksheet.getCell(r + 1, C + 1);
    cell.value = { richText: rich };
    cell.alignment = { horizontal: (titleRuns[0].props.hAlign || "left"), vertical: "top", wrapText: true };
    const lineSizes = [];
    let cur = 0;
    rich.forEach(x => x.text.split("\n").forEach((seg, i) => {
      if (i > 0) { lineSizes.push(cur); cur = 0; }
      if (seg) cur = Math.max(cur, x.font.size || 9);
    }));
    lineSizes.push(cur);
    worksheet.getRow(r + 1).height = Math.ceil(lineSizes.reduce((a, b) => a + (b || 9) * 1.3, 0) + 4);
  }
  rangeTracker.update(r, C);
  rangeTracker.update(r, C + 4);
  r++;
  r++;
  return r - originRow;
}


/* ═══ export/sheet-data.js ════════════════════════════════════════════════════════════════════════ */
/* Fetching worksheet summary data, filters and filter-value sheets from Tableau. */

async function fetchAllSheetsData(sheets, concurrency = 4) {
const results = [];
const queue = [...sheets];

async function worker() {
  while (queue.length) {
    const sheet = queue.shift();
    try {
      const [dataResult, visualSpecResult] = await Promise.allSettled([
        // every mark, not only the selected ones: with a mark selected Tableau returns just that mark otherwise
        sheet.getSummaryDataAsync({ ignoreSelection: true }),
        typeof sheet.getVisualSpecificationAsync === "function"
          ? sheet.getVisualSpecificationAsync()
          : Promise.reject(new Error("getVisualSpecificationAsync is unavailable"))
      ]);
      if (dataResult.status === "rejected") throw dataResult.reason;
      results.push({
        sheet,
        data: dataResult.value,
        visualSpec: visualSpecResult.status === "fulfilled" ? visualSpecResult.value : null,
        visualSpecError: visualSpecResult.status === "rejected" ? visualSpecResult.reason : null,
        error: null
      });
    } catch (err) {
      results.push({ sheet, data: null, error: err });
    }
  }
}

const workers = Array(concurrency).fill().map(() => worker());
await Promise.all(workers);
return results;
}

/* ── Filter Extraction ─────────────────────────────────────────────────── */
async function extractFilterValuesPerField(sheets) {
  const filterMap = new Map();
  
  for (const worksheet of sheets) {
    let filters = [];
    
    try {
      filters = await worksheet.getFiltersAsync();
    } catch (e) {
      console.warn(`Could not get filters for ${worksheet.name}:`, e);
      continue;
    }
    
    filters.forEach(filter => {
      const ignoredFields = ["Measure Names", "Measure Values"];
      if (ignoredFields.includes(filter.fieldName) || /^Month\(/i.test(filter.fieldName)) {
        return;
      }
      
      if (!filterMap.has(filter.fieldName)) {
        filterMap.set(filter.fieldName, {
          values: new Set(),
          worksheetNames: new Set()
        });
      }
      
      const filterData = filterMap.get(filter.fieldName);
      filterData.worksheetNames.add(worksheet.name);
      
      let value = "";
      
      if (filter.filterType === "categorical") {
        const appliedValues = (filter.appliedValues || []).map(v => v.formattedValue);
        appliedValues.forEach(v => filterData.values.add(v));
        value = appliedValues.join(", ");
      } else if (filter.filterType === "range") {
        // date ranges arrive as "6/1/2025 12:00:00 AM" – a midnight time is not part of the filter
        const end = v => String((v && v.formattedValue) || "").replace(/\s+12:00:00\s*AM$|\s+00:00:00$/i, "");
        value = end(filter.minValue) + " - " + end(filter.maxValue);
        filterData.values.add(value);
      } else {
        value = filter.filterType;
        filterData.values.add(value);
      }
    });
  }
  
  const result = {};
  for (const [fieldName, data] of filterMap.entries()) {
    result[fieldName] = Array.from(data.values);
  }
  
  return result;
}

/** The marks selected on a sheet (a selected pie slice is outlined, as Tableau shows it); [] when unavailable.
 * @param {any} sheet Tableau worksheet @returns {Promise<any[]>} */
async function fetchSelectedMarks(sheet) {
  try {
    const marks = typeof sheet.getSelectedMarksAsync === "function" ? await sheet.getSelectedMarksAsync() : null;
    return (marks && marks.data) || [];
  } catch (e) {
    console.warn(`[Export] selected marks of "${sheet.name}" unavailable: ${e.message}`);
    return [];
  }
}

/* ── Check if worksheet is a filter value table ───────────────────────── */
function isFilterValueWorksheet(sheetName, summaryData) {
  if (/filter[_\- ]?\d+/i.test(sheetName) || 
      /values[_\- ]?\d+/i.test(sheetName) ||
      /_filter_\d+/i.test(sheetName)) {
    return true;
  }
  
  if (summaryData && summaryData.columns && summaryData.columns.length === 1) {
    const firstColumnName = summaryData.columns[0]?.fieldName || "";
    if (firstColumnName === "Values" || firstColumnName === "VALUE") {
      return true;
    }
  }
  
  return false;
}


/* ═══ export/kpi-card.js ══════════════════════════════════════════════════════════════════════════ */
/* KPI cards and dashboard text boxes: formatted text in a block that fills its dashboard zone. */

/* Tableau's formatted text: "Æ" marks a line break – usually followed by a newline, sometimes only a space */
const LINE_BREAK = /Æ[ \t]*(?:\r?\n)?|\r?\n/;

/* white edges between neighbouring cards – Tableau's tile gutters are narrower than a grid column */
/** @type {Partial<import("exceljs").Border>} */
const KPI_GUTTER = { style: "medium", color: { argb: "FFFFFFFF" } };

/**
 * @typedef {{ text?: string, ci?: number, props: Record<string, any> }} KpiSegment  static text or a value column
 * @typedef {{ segments: KpiSegment[], hAlign: "left" | "center" | "right", sizePt: number }} KpiLine
 * @typedef {{ lines: KpiLine[] }} KpiTile
 * @typedef {{ title: { text: string, props: Record<string, any> } | null, tiles: KpiTile[], background: string | null,
 *             padTop: number, padBottom: number, lineHeightsPt: number[], rows: number, widthPx: number[],
 *             zoneWidthPx: number | null }} KpiCard
 */

/**
 * Tableau formatted text (custom label, text box) → lines of segments. Field tokens "<[ds].[field]>"
 * may be split over several runs ("<", "[Parameters].[P 10]", "> Stores by <" …), so they are found
 * in the joined text; each piece keeps the font of the run it comes from.
 * @param {TextRun[]} runs
 * @param {(inner: string, props: Record<string, any>) => KpiSegment | null} token a "<…>" token → segment (null = drop)
 * @param {(props: Record<string, any>) => Record<string, any>} styled static text props from a run's props
 * @returns {KpiSegment[][]}
 */
function formattedLines(runs, token, styled) {
  const full = runs.map(r => r.text).join("");
  const starts = [];
  runs.reduce((pos, r) => { starts.push(pos); return pos + r.text.length; }, 0);
  const runAt = i => { let k = 0; while (k + 1 < starts.length && starts[k + 1] <= i) k++; return k; };
  const lines = [[]];
  const plain = (from, to) => {                              // static text, cut at run and line boundaries
    for (let i = from; i < to;) {
      const k = runAt(i), end = Math.min(to, k + 1 < starts.length ? starts[k + 1] : full.length);
      full.slice(i, end).split(LINE_BREAK).forEach((part, n) => {
        if (n > 0) lines.push([]);
        if (part) lines[lines.length - 1].push({ text: part, props: styled(runs[k].props) });
      });
      i = end;
    }
  };
  let last = 0;
  for (const m of full.matchAll(/<([^<>]+)>/g)) {
    plain(last, m.index);
    const seg = token(m[1], runs[runAt(m.index + 1)].props);  // the run holding the field name
    if (seg) lines[lines.length - 1].push(seg);
    last = m.index + m[0].length;
  }
  plain(last, full.length);
  while (lines.length > 1 && !lines[lines.length - 1].length) lines.pop();          // trailing line breaks
  return lines;
}

/**
 * The tile as Tableau draws it: its custom label (lines, fonts, sizes, colours, alignment), or
 * Tableau's default text layout; sized to fill the dashboard zone. null = Measure Names on Rows –
 * Tableau draws that as a two-column text table, which the stacked label | value writer matches.
 * @param {ViewModel} vm @param {string} sheetName @param {{ heightPx?: number, widthPx?: number } | null} [layout]
 * @returns {KpiCard | null}
 */
function buildKpiCard(vm, sheetName, layout) {
  const fmt = vm.fmt;
  const sheet = fmt.sheetModel;
  const visible = vm.order.filter(ci => !vm.cols[ci].isHeader);
  const colOf = ref => vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
  const labelPane = sheet ? sheet.panes.find(p => p.labelRuns.length) : null;
  /** @type {KpiTile[]} */
  let tiles;
  if (labelPane) {
    const styled = props => tfMerge(fmt.markCellStyle(null), props);
    const lines = formattedLines(labelPane.labelRuns, (inner, props) => {
      if (/^sheet name$/i.test(inner)) return { text: sheetName, props: styled(props) };
      const ref = /^\[/.test(inner) ? tfParseFieldRef(inner) : null;
      const ci = ref ? colOf(ref) : -1;
      return ci >= 0 ? { ci, props: tfMerge(fmt.markCellStyle(vm.cols[ci].ref), props) } : null;
    }, styled);
    tiles = [{ lines: lines.map(segments => kpiLine(segments, fmt)) }];
  } else {
    const onShelf = shelf => !!sheet && sheet[shelf].some(r => r.name === "Measure Names");
    if (onShelf("rows")) return null;
    if (onShelf("cols")) {                                  // Measure Names on Columns: a tile per measure, name above value
      const nameStyle = fmt.headerCellStyle(tfParseFieldRef("[:Measure Names]"));
      tiles = visible.map(ci => ({ lines: [
        kpiLine([{ text: vm.cols[ci].label || vm.cols[ci].name, props: nameStyle }], fmt),
        kpiLine([{ ci, props: fmt.markCellStyle(vm.cols[ci].ref) }], fmt)
      ] }));
    } else {                                                // Text = the measures: one value per line
      tiles = [{ lines: visible.map(ci => kpiLine([{ ci, props: fmt.markCellStyle(vm.cols[ci].ref) }], fmt)) }];
    }
  }
  const depth = Math.max(...tiles.map(t => t.lines.length));
  const lineHeightsPt = Array.from({ length: depth }, (_, i) =>
    Math.ceil(Math.max(...tiles.map(t => (t.lines[i] ? t.lines[i].sizePt : 9))) * 1.35 + 2));
  const title = vm.showTitle ? vm.title : null;
  // padding rows (default 20 px) fill the zone's height around the lines, as Tableau centres the label
  const used = (title ? 20 : 0) + lineHeightsPt.reduce((s, h) => s + h * 4 / 3, 0);
  const pad = layout && layout.heightPx ? Math.max(0, Math.round((layout.heightPx - used) / 20)) : 0;
  const vAlign = fmt.markCellStyle(null).vAlign || "middle";
  const padTop = vAlign === "top" ? 0 : vAlign === "bottom" ? pad : Math.floor(pad / 2);
  const background = fmt.tableBackground();
  const widthPx = tiles.map(t => Math.max(60, ...t.lines.map(l =>
    l.segments.reduce((w, s) => w + textOf(s, vm).length * (s.props.fontSize || 9) * 4 / 3 * 0.58, 0) + 18)));
  return { title, tiles, background, padTop, padBottom: pad - padTop, lineHeightsPt,
           rows: (title ? 1 : 0) + pad + depth, widthPx, zoneWidthPx: (layout && layout.widthPx) || null };
}

/**
 * A dashboard text box (banner, title, note) as Tableau draws it: its runs – fonts, sizes, colours,
 * alignment – on the zone's background colour, filling the zone. Only parameter values can be fields
 * in a dashboard text box; other tokens (page name …) have no value in an export.
 * @param {DashboardZone} zone @param {{ heightPx?: number, widthPx?: number } | null} layout @param {FormatModel | null} model
 * @returns {KpiCard | null}
 */
function buildTextCard(zone, layout, model) {
  const base = tfMerge(TABLEAU_DEFAULTS.worksheet);
  const styled = props => tfMerge(base, props);
  const lines = formattedLines(zone.runs || [], (inner, props) => {
    const ref = /^\[/.test(inner) ? tfParseFieldRef(inner) : null;
    const info = ref && ref.ds === "Parameters" ? tfFieldInfo(model, ref) || {} : null;
    const text = info ? String(info.alias || info.value || "").replace(/^"|"$/g, "") : "";
    return text ? { text, props: styled(props) } : null;
  }, styled);
  if (!lines.some(l => l.some(s => String(s.text).trim()))) return null;
  /** @type {KpiLine[]} */
  const kpiLines = lines.map(segments => ({ segments,
    hAlign: /** @type {KpiLine["hAlign"]} */ ((segments.find(s => s.props.hAlign) || { props: { hAlign: "left" } }).props.hAlign),
    sizePt: Math.max(9, ...segments.map(s => s.props.fontSize || 9)) }));
  const lineHeightsPt = kpiLines.map(l => Math.ceil(l.sizePt * 1.35 + 2));
  const used = lineHeightsPt.reduce((s, h) => s + h * 4 / 3, 0);
  const pad = layout && layout.heightPx ? Math.max(0, Math.round((layout.heightPx - used) / 20)) : 0;
  const widthPx = [Math.max(40, ...kpiLines.map(l => l.segments.reduce((w, s) => w + s.text.length * (s.props.fontSize || 9) * 4 / 3 * 0.58, 0) + 12))];
  return { title: null, tiles: [{ lines: kpiLines }], background: (zone.style && zone.style.bgColor) || null,
           padTop: Math.floor(pad / 2), padBottom: pad - Math.floor(pad / 2), lineHeightsPt, rows: pad + kpiLines.length,
           widthPx, zoneWidthPx: (layout && layout.widthPx) || null };
}

/** a line's alignment: the pane's Format → Alignment, else the label editor's, else centred
 * @returns {KpiLine} */
function kpiLine(segments, fmt) {
  const cellAlign = tfMerge(fmt.markCellStyle(null)).hAlign;
  const runAlign = (segments.find(s => s.props.hAlign) || { props: {} }).props.hAlign;
  const hAlign = /** @type {KpiLine["hAlign"]} */ (/^(left|right)$/.test(cellAlign || runAlign || "") ? (cellAlign || runAlign) : "center");
  return { segments, hAlign, sizePt: Math.max(9, ...segments.map(s => s.props.fontSize || 9)) };
}

function textOf(s, vm) { return s.ci !== undefined ? tfDvText(vm.rows[0][s.ci]) : String(s.text || ""); }

/**
 * Writes the card over gridW columns from (originRow, originCol); returns the rows used.
 * @param {import("exceljs").Worksheet} worksheet @param {KpiCard} card @param {ViewModel | null} vm null for a text box
 * @param {number} originRow @param {number} originCol @param {number} gridW
 * @param {{ update(r: number, c: number): void }} rangeTracker @param {Record<number, number>} colWidths
 * @param {number[] | null} [cuts] the column (from originCol) where each next tile starts – equal shares of the zone
 */
function writeKpiCard(worksheet, card, vm, originRow, originCol, gridW, rangeTracker, colWidths, cuts = null) {
  let r = originRow;
  const C = originCol;
  if (card.title) {
    writeTableauTitle(worksheet, r, C, card.title.text, card.title.props, gridW);
    rangeTracker.update(r, C); rangeTracker.update(r, C + gridW - 1);
    r++;
  }
  const depth = card.lineHeightsPt.length;
  const bodyRows = card.padTop + depth + card.padBottom;
  if (card.background) {
    for (let rr = r; rr < r + bodyRows; rr++) {
      for (let c = C; c < C + gridW; c++) {
        const cell = worksheet.getCell(rr + 1, c + 1);
        cell.fill = tfExcelFill(card.background);
        if (c === C || c === C + gridW - 1) cell.border = { ...(c === C ? { left: KPI_GUTTER } : {}), ...(c === C + gridW - 1 ? { right: KPI_GUTTER } : {}) };
      }
    }
  }
  const plan = vm ? buildColorPlan(vm.fmt, vm.cols, vm.rows) : null;
  const n = card.tiles.length;
  const span = Math.max(1, Math.floor(gridW / n));
  // the tile's columns: between its boundaries on the grid (equal shares of the zone), else an equal count
  const byCuts = cuts && cuts.length === n - 1 && cuts.every((c, i) => c > (i ? cuts[i - 1] : 0) && c < gridW);
  card.tiles.forEach((tile, k) => {
    const c0 = byCuts ? C + (k ? cuts[k - 1] : 0) : C + k * span;
    const c1 = byCuts ? C + (k < n - 1 ? cuts[k] : gridW) - 1 : k === n - 1 ? C + gridW - 1 : c0 + span - 1;
    // as wide as the tile's zone (and never narrower than its text); Excel width units ≈ 7 px
    const tilePx = Math.max(card.widthPx[k], card.zoneWidthPx ? card.zoneWidthPx / n : 0);
    const perCol = Math.ceil(tilePx / (c1 - c0 + 1) / 7);
    for (let c = c0; c <= c1; c++) colWidths[c] = Math.max(colWidths[c] || 0, perCol);
    tile.lines.forEach((ln, i) => {
      const row = r + card.padTop + i;
      if (c1 > c0) worksheet.mergeCells(row + 1, c0 + 1, row + 1, c1 + 1);
      const cell = worksheet.getCell(row + 1, c0 + 1);
      const only = ln.segments.length === 1 ? ln.segments[0] : null;
      if (only && only.ci !== undefined) {                  // a single value: keep the number and its format
        const color = plan && plan.enc && plan.markIdx.has(only.ci) && plan.enc.applyTo === "font" && !only.props.explicitColor ? plan.colorAt(0) : null;
        tfWriteCell(cell, vm.rows[0][only.ci], only.props, { fill: card.background || undefined, fontColor: color || undefined });
      } else if (ln.segments.length) {
        cell.value = { richText: ln.segments.map(s => ({ text: textOf(s, vm), font: tfExcelFont(s.props) })) };
        if (card.background) cell.fill = tfExcelFill(card.background);
      }
      cell.alignment = { horizontal: ln.hAlign, vertical: "middle", indent: ln.hAlign === "center" ? 0 : 1, wrapText: false };
      // a merged range has one style (ExcelJS): both outer edges go on it, Excel draws only the outline
      if (card.background) cell.border = { ...(k === 0 ? { left: KPI_GUTTER } : {}), ...(k === n - 1 ? { right: KPI_GUTTER } : {}) };
      const xr = worksheet.getRow(row + 1);
      xr.height = Math.max(xr.height || 15, card.lineHeightsPt[i]);
      rangeTracker.update(row, c0); rangeTracker.update(row, c1);
    });
  });
  rangeTracker.update(r + bodyRows - 1, C + gridW - 1);
  return (card.title ? 1 : 0) + bodyRows;
}


/* ═══ export/backgrounds.js ═══════════════════════════════════════════════════════════════════════ */
/* Dashboard shading and container / object backgrounds (Layout pane → Background) → the colour of the
 * cells around and under the blocks on the sheet; thin coloured zones (dividers, accent lines) → borders. */

/**
 * A block on the sheet (0-based rows / columns, inclusive) and the zone it comes from.
 * own: the block's own opaque colour (ARGB); "white" = a sheet Tableau draws white (left unfilled);
 * undefined = see-through, the zone behind it shows.
 * @typedef {{ zoneId: string, top: number, left: number, bottom: number, right: number, own?: string }} BgBlock
 */

/**
 * A divider: along the top of row `at` ("h", columns from–to) or the left of column `at` ("v", rows from–to).
 * @typedef {{ dir: "h" | "v", at: number, from: number, to: number, color: string, style: "thin" | "medium" | "thick" }} BgLine
 */

/* thinner than this (px on the dashboard) a coloured zone is a line, not an area */
const LINE_PX = 8;

/* zone coordinate (0–100000) → grid edge: piecewise linear through the edges of the blocks laid out from
 * zones, never decreasing (the layout keeps the dashboard's order) */
function axisMap(anchors, lo, hi) {
  const pts = [[0, lo], ...anchors, [100000, hi]].sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  for (let i = 1; i < pts.length; i++) pts[i][1] = Math.max(pts[i][1], pts[i - 1][1]);
  return v => {
    if (v <= pts[0][0]) return pts[0][1];
    for (let i = 1; i < pts.length; i++) {
      if (v > pts[i][0]) continue;
      const [x0, g0] = pts[i - 1], [x1, g1] = pts[i];
      return x1 === x0 ? g1 : g0 + (g1 - g0) * (v - x0) / (x1 - x0);
    }
    return pts[pts.length - 1][1];
  };
}

/**
 * The colour each cell shows behind the blocks, as on the dashboard: the dashboard shading at the back,
 * then every container / backdrop zone with a background from the largest to the smallest, then each
 * block's own colour. A zone colours the blocks whose zones sit inside it (their centres) and the cells
 * between them – out to the zone's own edges, placed on the grid between the blocks' edges, so a header
 * bar spans the sheet as it spans the dashboard. Zones thinner than a few px are dividers: lines.
 * @param {DashboardZone[]} zones @param {string | null} dashColor ARGB
 * @param {BgBlock[]} blocks
 * @param {{ top: number, left: number, bottom: number, right: number }} canvas the dashboard's cells
 * @param {{ w: number, h: number }} dashPx the dashboard's size in px
 * @returns {{ colorAt: (row: number, col: number) => string | null, lines: BgLine[] }} colorAt: ARGB, null = leave the cell
 */
function backgroundPlan(zones, dashColor, blocks, canvas, dashPx) {
  const zoneById = new Map((zones || []).map(z => [String(z.id), z]));
  const blockZones = new Set(blocks.map(b => b.zoneId));
  const placed = blocks.map(b => ({ b, z: zoneById.get(b.zoneId) })).filter(p => p.z && p.z.w > 0 && p.z.h > 0);
  const mapX = axisMap(placed.flatMap(({ b, z }) => [[z.x, b.left], [z.x + z.w, b.right + 1]]), canvas.left, canvas.right + 1);
  const mapY = axisMap(placed.flatMap(({ b, z }) => [[z.y, b.top], [z.y + z.h, b.bottom + 1]]), canvas.top, canvas.bottom + 1);
  const center = z => ({ x: z.x + z.w / 2, y: z.y + z.h / 2 });
  const inside = (p, z) => p.x >= z.x && p.x <= z.x + z.w && p.y >= z.y && p.y <= z.y + z.h;
  // a zone in a hidden container is not on the dashboard either
  const shown = z => { for (let q = z, n = 0; q && n < 50; q = zoneById.get(String(q.parent)), n++) if (q.hidden) return false; return true; };
  const coloured = (zones || []).map((z, order) => ({ z, order }))
    .filter(({ z }) => z.style && z.style.bgColor && !blockZones.has(String(z.id)) && z.w > 0 && z.h > 0 && shown(z));
  const px = z => ({ w: z.w * dashPx.w / 100000, h: z.h * dashPx.h / 100000 });

  /** @type {BgLine[]} */
  const lines = [];
  coloured.forEach(({ z }) => {
    const { w, h } = px(z);
    if (Math.min(w, h) >= LINE_PX || Math.max(w, h) < LINE_PX) return;
    const thick = Math.min(w, h), style = thick < 1.5 ? "thin" : thick < 2.5 ? "medium" : "thick";
    const c = center(z);
    const line = h < w
      ? { dir: "h", at: Math.round(mapY(c.y)), from: Math.round(mapX(z.x)), to: Math.round(mapX(z.x + z.w)) - 1 }
      : { dir: "v", at: Math.round(mapX(c.x)), from: Math.round(mapY(z.y)), to: Math.round(mapY(z.y + z.h)) - 1 };
    if (line.to >= line.from) lines.push(/** @type {BgLine} */ ({ ...line, color: z.style.bgColor, style }));
  });

  const layers = coloured
    .filter(({ z }) => { const { w, h } = px(z); return Math.min(w, h) >= LINE_PX; })
    .map(({ z, order }) => {
      const members = placed.filter(p => inside(center(p.z), z)).map(p => p.b);
      const rect = { top: Math.round(mapY(z.y)), left: Math.round(mapX(z.x)),
                     bottom: Math.round(mapY(z.y + z.h)) - 1, right: Math.round(mapX(z.x + z.w)) - 1 };
      members.forEach(b => {
        rect.top = Math.min(rect.top, b.top); rect.left = Math.min(rect.left, b.left);
        rect.bottom = Math.max(rect.bottom, b.bottom); rect.right = Math.max(rect.right, b.right);
      });
      if (rect.bottom < rect.top || rect.right < rect.left) return null;
      return { color: z.style.bgColor, area: z.w * z.h, order, members: new Set(members), ...rect };
    })
    .filter(Boolean)
    .sort((a, b) => b.area - a.area || a.order - b.order);             // back to front
  const covers = (b, r, c) => r >= b.top && r <= b.bottom && c >= b.left && c <= b.right;
  const colorAt = (r, c) => {
    const block = blocks.find(b => covers(b, r, c));
    if (block) {
      if (block.own === "white") return null;
      if (block.own) return block.own;
      for (let i = layers.length - 1; i >= 0; i--) if (layers[i].members.has(block)) return layers[i].color;
      return dashColor || null;
    }
    for (let i = layers.length - 1; i >= 0; i--) if (covers(layers[i], r, c)) return layers[i].color;
    return dashColor || null;
  };
  return { colorAt, lines };
}


/* ═══ visual/workbook-match.js ════════════════════════════════════════════════════════════════════ */
/* Does the loaded workbook belong to the running dashboard? */

/* does the loaded workbook belong to the running dashboard? (a different or stale .twb would
 * silently apply the wrong titles, colours and sorts) → { level: "ok" | "partial" | "mismatch", … } */
function checkWorkbookMatch(model, dashboard) {
  if (!model || !dashboard) return null;
  const names = (dashboard.worksheets || []).map(w => w.name);
  const known = new Set(Object.keys(model.sheets || {}).map(tfNorm));
  const missing = names.filter(n => !known.has(tfNorm(n)));
  const matched = names.length - missing.length;
  const dashboardFound = Object.keys(model.dashboards || {}).some(d => tfNorm(d) === tfNorm(dashboard.name));
  const level = !missing.length && dashboardFound ? "ok"
    : names.length && matched / names.length >= 0.5 ? "partial" : "mismatch";
  return { level, dashboard: dashboard.name, dashboardFound, matched, total: names.length, missing };
}

function describeWorkbookMatch(match) {
  if (!match || match.level === "ok") return "";
  const parts = [];
  if (!match.dashboardFound) parts.push(`dashboard "${match.dashboard}" is not in this workbook`);
  if (match.missing.length) {
    parts.push(`${match.matched}/${match.total} worksheets found` +
      ` (missing: ${match.missing.slice(0, 3).join(", ")}${match.missing.length > 3 ? ", …" : ""})`);
  }
  return (match.level === "mismatch" ? "⚠ Wrong workbook? " : "⚠ Workbook partly matches: ") + parts.join("; ");
}


/* ═══ ui/status.js ════════════════════════════════════════════════════════════════════════════════ */
/* Panel status line and loaded-workbook label. */

/* loaded-workbook label, with a warning when the file does not match this dashboard */
function showWorkbookLabel(fileName, details, model) {
  const label = document.getElementById("twb_file_label");
  if (!label) return;
  const match = checkWorkbookMatch(model, tableau.extensions.dashboardContent.dashboard);
  const warning = describeWorkbookMatch(match);
  label.textContent = `${warning ? "" : "✅ "}${fileName} — ${details}` + (warning ? `\n${warning}` : "");
  label.title = match && match.missing.length ? "Not found in the workbook:\n" + match.missing.join("\n") : "";
  label.classList.toggle("status-warning", !!warning);
  if (warning) console.warn("[Workbook]", warning, match);
}

/* progress / outcome line under the Export button */
function setExportStatus(text, warning) {
  const target = document.getElementById("export_status");
  if (!target) return;
  target.textContent = text;
  target.title = "";
  target.classList.toggle("status-warning", !!warning);
}

function appendExportStatus(text) {
  const target = document.getElementById("export_status");
  if (target && target.textContent) target.textContent += " · " + text;
}

function updateVisualStatus(statuses, note) {
  const target = document.getElementById("export_status");
  if (!target) return;
  if (note) {                                      // e.g. the loaded workbook does not match
    updateVisualStatus(statuses);
    target.textContent = note + " · " + target.textContent;
    target.classList.add("status-warning");
    return;
  }
  if (!statuses || !statuses.length) {
    target.textContent = "No worksheet visual status available";
    return;
  }
  // once converted: how (NATIVE … TABLE_FALLBACK, as on the report sheet); before that, which renderer
  if (statuses.every(item => item.strategy)) {
    const order = ["NATIVE", "CONSTRUCTED", "APPROXIMATE", "TABLE_FALLBACK", "IMAGE_FALLBACK", "UNSUPPORTED", "FAILED"];
    const words = { NATIVE: "native", CONSTRUCTED: "constructed", APPROXIMATE: "approximated", TABLE_FALLBACK: "as data tables",
                    IMAGE_FALLBACK: "as pictures", UNSUPPORTED: "unsupported", FAILED: "failed" };
    const parts = order.map(s => [statuses.filter(item => item.strategy === s).length, words[s]]).filter(([n]) => n);
    const flagged = statuses.filter(item => item.strategy !== "NATIVE" && item.strategy !== "CONSTRUCTED");
    target.textContent = `${statuses.length} visuals: ${parts.map(([n, w]) => `${n} ${w}`).join(", ")} – see the Conversion Report sheet`;
    target.title = flagged.map(w => `${w.worksheet} (${w.visual}): ${w.strategy} – ${w.reason}`).join("\n");
    target.classList.toggle("status-warning", statuses.some(item => item.strategy === "FAILED"));
    return;
  }
  const count = renderer => statuses.filter(item => item.renderer === renderer).length;
  const parts = [
    [count("cell"), "tables/KPIs"],
    [count("excel-chart"), "Excel charts"],
    [count("tableau-image"), "Tableau images"],
    [count("data-fallback"), "as data"]
  ].filter(([n]) => n).map(([n, label]) => `${n} ${label}`);
  const warnings = statuses.filter(item => item.status !== "success" && item.status !== "pending");
  target.textContent = `${statuses.length} visuals: ${parts.join(", ")}` +
    (warnings.length ? ` – ${warnings.length} warning${warnings.length === 1 ? "" : "s"} (see console)` : "");
  target.title = warnings.map(w => `${w.worksheet}: ${w.reason}`).join("\n");
  target.classList.toggle("status-warning", warnings.length > 0);
}

/* plain message on the loaded-workbook label (auto-load progress, problems) */
function showWorkbookStatus(text, warning) {
  const label = document.getElementById("twb_file_label");
  if (!label) return;
  label.textContent = text;
  label.title = "";
  label.classList.toggle("status-warning", !!warning);
}

/* what was left out of a multi-dashboard export, and why – one line each under the status */
function appendExportNotes(notes) {
  (notes || []).forEach(n => console.warn("[Export] " + n));
  const target = document.getElementById("export_status");
  if (!target || !notes || !notes.length) return;
  target.textContent = [target.textContent, ...notes.map(n => "⚠️ " + n)].filter(Boolean).join("\n");
  target.classList.add("status-warning");
}


/* ═══ cloud/backend.js ════════════════════════════════════════════════════════════════════════════ */
/* =============================================================================
 * AUTO-FETCH FROM TABLEAU CLOUD (Phase 1)
 * -----------------------------------------------------------------------------
 * Replaces the manual "Load Workbook" step on Tableau Cloud:
 *   1. ask the backend (Worker) which published workbook holds this dashboard
 *   2. download that workbook's XML through the backend
 *   3. parse it exactly like a manually uploaded file (applyWorkbookXml)
 * The manual 📁 button stays as a fallback (Tableau Desktop, backend down).
 * The PAT never reaches this file — only the backend holds it.
 * ============================================================================= */

const DEFAULT_BACKEND_URL = "";      // production: the hosted Worker URL, e.g. "https://export-backend.example.workers.dev"

/** Read a value saved by this extension: workbook settings first, then this browser. */
function readSaved(key) {
  let v = null;
  try { v = tableau.extensions.settings.get(key); } catch (e) { /* ignore */ }
  if (!v) { try { v = localStorage.getItem("tfx_" + key); } catch (e) { /* ignore */ } }
  return v || "";
}

/** Save values to workbook settings (works in edit mode) and this browser (works always). */
async function writeSaved(pairs) {
  Object.keys(pairs).forEach(k => {
    const v = pairs[k];
    try { localStorage.setItem("tfx_" + k, v); } catch (e) { /* ignore */ }
    try {
      if (v === "") tableau.extensions.settings.erase(k);
      else tableau.extensions.settings.set(k, v);
    } catch (e) { /* ignore */ }
  });
  try { await tableau.extensions.settings.saveAsync(); }
  catch (e) { console.warn("[AutoFetch] settings.saveAsync failed (viewing mode?) — kept in this browser only:", e.message); }
}

function getBackendConfig() {
  return {
    url: (readSaved("backendUrl") || DEFAULT_BACKEND_URL).trim().replace(/\/+$/, ""),
    key: readSaved("backendKey").trim(),
  };
}

/** Call the backend. Throws a readable Error on network failure or non-2xx. */
async function backendFetch(path, init = {}) {
  const { url, key } = getBackendConfig();
  if (!url) throw new Error("No backend URL configured");
  const headers = { ...(init.headers || {}) };
  if (key) headers["X-Spike-Key"] = key;
  let res;
  try { res = await fetch(url + path, { ...init, headers }); }
  catch (e) { throw new Error(`Could not reach the backend (${url}) — is it running?`); }
  if (!res.ok) {
    let text = "";
    try { text = await res.text(); } catch (e) { /* ignore */ }
    let msg = text;
    try { msg = JSON.parse(text).error || text; } catch (e) { /* not JSON */ }
    throw new Error(`Backend ${res.status}: ${String(msg).slice(0, 300)}`);
  }
  return res;
}

/** "N titles, formatting for M sheets (K with colour)" – the loaded-workbook label */
function describeModel(titleMap, formatModel, images) {
  const sheets = Object.values((formatModel && formatModel.sheets) || {});
  const colour = sheets.filter(s => (s.panes || []).some(p => (p.encodings || []).some(e => e.channel === "color"))).length;
  const imageCount = images ? Object.keys(images).length : 0;
  return `${Object.keys(titleMap || {}).length} titles, formatting for ${sheets.length} sheets (${colour} with colour)` +
    (imageCount ? `, ${imageCount} image${imageCount > 1 ? "s" : ""}` : "");
}

/**
 * Safety check: does the downloaded XML really contain this dashboard and its worksheets?
 * Catches a stale saved workbook ID (e.g. the workbook was copied with "Save As").
 */
function twbMatchesDashboard(xmlString, dashboard) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const root = doc && doc.documentElement;
  if (!root || root.nodeName !== "workbook") return false;
  const names = tag => {
    const holder = getDirectChildByTag(root, tag + "s");
    const out = new Set();
    if (!holder) return out;
    for (let i = 0; i < holder.childNodes.length; i++) {
      const c = holder.childNodes[i];
      if (c.nodeType === 1 && c.tagName === tag) out.add(c.getAttribute("name"));
    }
    return out;
  };
  const dashboards = names("dashboard"), worksheets = names("worksheet");
  return dashboards.has(dashboard.name) && dashboard.worksheets.every(ws => worksheets.has(ws.name));
}

/** Which published workbook is this? Saved answer first, otherwise ask the backend. */
async function resolveWorkbook(dashboard) {
  const savedId = readSaved("twbWorkbookId");
  if (savedId) return { id: savedId, name: readSaved("twbWorkbookName") || "workbook", fromSaved: true };

  const res = await backendFetch("/resolve", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ dashboardName: dashboard.name, worksheetNames: dashboard.worksheets.map(w => w.name) }),
  });
  const r = await res.json();
  const pick = c => ({ id: c.workbookId, name: c.workbookName, project: c.projectName });
  if (r.decision === "unique") return pick(r.candidates[0]);
  if (r.decision === "ambiguous") {
    const exact = r.candidates.filter(c => c.exact);
    const list = (exact.length ? exact : r.candidates.filter(c => c.score > 0)).map(pick);
    return { ambiguous: true, candidates: list };
  }
  throw new Error(`no published workbook contains dashboard "${dashboard.name}"`);
}

/** Several workbooks match (e.g. copies in two projects): let the user pick once. */
function askUserToChooseWorkbook(candidates) {
  return new Promise(resolve => {
    const label = document.getElementById("twb_file_label");
    const box = document.createElement("div");
    box.id = "twb_picker";
    box.className = "twb-picker";
    const select = document.createElement("select");
    candidates.forEach((c, i) => {
      const o = document.createElement("option");
      o.value = String(i);
      o.textContent = `${c.project ? c.project + " / " : ""}${c.name}`;
      select.appendChild(o);
    });
    const ok = document.createElement("button");
    ok.className = "btn-load";
    ok.textContent = "Use this workbook";
    ok.addEventListener("click", () => { box.remove(); resolve(candidates[Number(select.value)]); });
    box.appendChild(select);
    box.appendChild(ok);
    if (label && label.parentNode) label.parentNode.insertBefore(box, label.nextSibling);
    else document.body.appendChild(box);
  });
}

/** Auto-load on open. Never throws: on any problem it explains and leaves 📁 as the fallback. */
async function autoLoadWorkbook(dashboard) {
  const env = tableau.extensions.environment || {};
  const savedFile = readSaved("twbFileName");
  const fallback = savedFile ? ` — using previously loaded ${savedFile}` : " — use 📁 to load the workbook manually";

  if (env.context === "desktop") {
    if (!savedFile) showWorkbookStatus("Tableau Desktop: auto-load needs Tableau Cloud — use 📁 to load the workbook", true);
    return;
  }
  if (!getBackendConfig().url) {
    if (!savedFile) showWorkbookStatus("Auto-load not set up (Advanced → Backend URL)" + fallback, true);
    return;
  }

  try {
    for (let attempt = 0; attempt < 2; attempt++) {
      showWorkbookStatus("⏳ Finding this workbook on Tableau Cloud…");
      let wb = await resolveWorkbook(dashboard);
      if (wb.ambiguous) {
        showWorkbookStatus(`${wb.candidates.length} published workbooks contain this dashboard — pick the right one:`);
        wb = await askUserToChooseWorkbook(wb.candidates);
      }

      showWorkbookStatus(`⏳ Downloading "${wb.name}" from Tableau Cloud…`);
      let xml;
      try {
        xml = await (await backendFetch("/twb/" + encodeURIComponent(wb.id))).text();
      } catch (e) {
        if (wb.fromSaved && /Backend 404/.test(e.message)) {          // saved workbook was deleted/moved
          await writeSaved({ twbWorkbookId: "", twbWorkbookName: "" });
          continue;
        }
        throw e;
      }

      if (!twbMatchesDashboard(xml, dashboard)) {
        if (wb.fromSaved) {                                            // saved ID points at another workbook
          await writeSaved({ twbWorkbookId: "", twbWorkbookName: "" });
          continue;
        }
        throw new Error("the downloaded workbook doesn't contain this dashboard");
      }

      const { titleMap, formatModel, images } = await applyWorkbookXml(xml, wb.name, false);
      setLoadedWorkbookId(wb.id);            // lets Phase 3 find the other dashboards' views
      if (!wb.fromSaved) await writeSaved({ twbWorkbookId: wb.id, twbWorkbookName: wb.name });
      showWorkbookLabel(`${wb.name} (auto, Tableau Cloud)`, describeModel(titleMap, formatModel, images), formatModel);
      return;
    }
    throw new Error("could not identify this workbook");
  } catch (e) {
    console.warn("[AutoFetch]", e);
    showWorkbookStatus(`⚠️ Auto-load failed: ${e.message}${fallback}`, true);
  }
}

/** Advanced panel: backend URL + key, saved with the workbook. Saving reconnects.
 *  onReload(promise) receives the new auto-load, so Export can wait for it. */
function setupBackendSettings(dashboard, onReload) {
  const urlIn = /** @type {HTMLInputElement} */ (document.getElementById("backend_url"));
  const keyIn = /** @type {HTMLInputElement} */ (document.getElementById("backend_key"));
  const saveBtn = /** @type {HTMLButtonElement} */ (document.getElementById("backend_save"));
  if (!urlIn || !keyIn || !saveBtn) return;
  const cfg = getBackendConfig();
  urlIn.value = readSaved("backendUrl");
  urlIn.placeholder = DEFAULT_BACKEND_URL || "https://…trycloudflare.com";
  keyIn.value = cfg.key;
  saveBtn.addEventListener("click", async () => {
    saveBtn.disabled = true;
    await writeSaved({ backendUrl: urlIn.value.trim(), backendKey: keyIn.value.trim() });
    const reload = autoLoadWorkbook(dashboard);
    if (onReload) onReload(reload);
    await reload;
    saveBtn.disabled = false;
  });
}


/* ═══ ui/workbook-store.js ════════════════════════════════════════════════════════════════════════ */
/* Loading the .twb/.twbx, remembering it (settings + IndexedDB) and the save location. */

/* =============================================================================
 * parseTwbXmlInBrowser(xmlString) - Parses TITLES from XML
 * ============================================================================= */
function parseTwbXmlInBrowser(xmlString) {
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

function getDirectChildByTag(parent, tagName) {
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
let workbookFileHandle = null;

const HANDLE_DB = "mark2table", HANDLE_STORE = "handles", HANDLE_KEY = "workbook";

function openHandleDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(HANDLE_DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(HANDLE_STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function storeWorkbookHandle(handle) {
  try {
    const db = await openHandleDb();
    db.transaction(HANDLE_STORE, "readwrite").objectStore(HANDLE_STORE).put(handle, HANDLE_KEY);
  } catch (e) {
    console.warn("[Workbook] could not remember the workbook folder:", e.message);
  }
}

/* parsed workbook model: Tableau's extension settings cannot hold a large model
 * (saveAsync fails) and it would be lost on the next reload → keep a copy here too, with the workbook XML
 * so a newer version of the extension can parse it again */
async function storeFormatModel(fileName, model, titleMap, xml) {
  try {
    const db = await openHandleDb();
    // one entry per workbook file: any dashboard of a workbook loaded once is matched again later
    db.transaction(HANDLE_STORE, "readwrite").objectStore(HANDLE_STORE)
      .put({ fileName, model, titleMap, xml, savedAt: Date.now() }, "model:" + String(fileName).toLowerCase());
  } catch (e) {
    console.warn("[Workbook] could not keep the workbook formatting in browser storage:", e.message);
  }
}

/** a model parsed by an older version of the extension: false = parsed by this one */
function isFormatModelStale(model) { return !!model && model.version !== FORMAT_MODEL_VERSION; }

/* a remembered workbook parsed by an older version: parsed again from its XML, and stored again */
function freshen(rec) {
  if (!isFormatModelStale(rec.model) || !rec.xml) return rec;
  try {
    const model = parseTableauFormatting(rec.xml);
    console.log(`[Workbook] ${rec.fileName}: formatting read again by this version of the extension`);
    storeFormatModel(rec.fileName, model, rec.titleMap, rec.xml);
    return { ...rec, model };
  } catch (e) {
    console.warn(`[Workbook] ${rec.fileName}: could not read the stored workbook again:`, e.message);
    return rec;
  }
}

async function restoreFormatModels() {
  try {
    const db = await openHandleDb();
    return await new Promise(resolve => {
      const out = [];
      const req = db.transaction(HANDLE_STORE).objectStore(HANDLE_STORE).openCursor();
      req.onsuccess = () => {
        const cur = req.result;
        if (!cur) return resolve(out);
        const key = String(cur.key);
        if ((key.startsWith("model:") || key === "formatModel") && cur.value && cur.value.model) out.push(freshen(cur.value));
        cur.continue();
      };
      req.onerror = () => resolve(out);
    });
  } catch (e) {
    return [];
  }
}

async function restoreWorkbookHandle() {
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
 * Packaged images – the files the dashboards' image objects show (logos, icons). Read from the .twbx
 * on load and kept in IndexedDB under their own key, so a remembered workbook still exports them.
 * ============================================================================= */
/** @type {{ file: string | null, images: Record<string, WorkbookImage> } | null} */
let IMAGE_CACHE = null;

/** "Image/logo.png" → its bytes, for every image zone of the workbook's dashboards found in the archive
 * @param {JSZip} zip @param {FormatModel} model @returns {Promise<Record<string, WorkbookImage>>} */
async function extractWorkbookImages(zip, model) {
  const wanted = new Set();
  Object.values(model.dashboards || {}).forEach(d => (d.zones || []).forEach(z => { if (z.type === "bitmap" && z.param) wanted.add(z.param); }));
  const files = Object.values(zip.files).filter(f => !f.dir);
  const norm = p => String(p).replace(/\\/g, "/").toLowerCase();
  const base = p => norm(p).split("/").pop();
  /** @type {Record<string, WorkbookImage>} */
  const images = {};
  for (const path of wanted) {
    // packaged as written in the zone ("Image/x.png"); a zone that points at the author's disk matches by file name
    const f = files.find(x => norm(x.name) === norm(path)) || files.find(x => base(x.name) === base(path));
    if (f) images[path] = { data: await f.async("uint8array") };
  }
  return images;
}

async function storeWorkbookImages(fileName, images) {
  try {
    const db = await openHandleDb();
    db.transaction(HANDLE_STORE, "readwrite").objectStore(HANDLE_STORE)
      .put({ fileName, images, savedAt: Date.now() }, "images:" + String(fileName).toLowerCase());
  } catch (e) {
    console.warn("[Workbook] could not keep the workbook images in browser storage:", e.message);
  }
}

/** the packaged images of the workbook the formatting comes from ({} for a .twb or none)
 * @returns {Promise<Record<string, WorkbookImage>>} */
async function getWorkbookImages() {
  if (IMAGE_CACHE && IMAGE_CACHE.file === FORMAT_MODEL_FILE) return IMAGE_CACHE.images;
  if (!FORMAT_MODEL_FILE || typeof indexedDB === "undefined") return {};
  try {
    const db = await openHandleDb();
    /** @type {any} */
    const rec = await new Promise(resolve => {
      const req = db.transaction(HANDLE_STORE).objectStore(HANDLE_STORE).get("images:" + FORMAT_MODEL_FILE.toLowerCase());
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
    IMAGE_CACHE = { file: FORMAT_MODEL_FILE, images: (rec && rec.images) || {} };
    return IMAGE_CACHE.images;
  } catch (e) {
    return {};
  }
}

/* =============================================================================
 * loadWorkbookFile() - Loads .twb/.twbx and parses BOTH titles AND colors
 * ============================================================================= */
async function loadWorkbookFile() {
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
async function readWorkbookFile(file) {
  console.log(`[loadWorkbookFile] Reading: ${file.name}`);

  try {
    let xmlString;
    let zip = null;
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

      zip = await JSZip.loadAsync(arrayBuffer);
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

    // same parsing path as the Tableau Cloud auto-fetch; a manual file is also saved into the workbook settings
    const { titleMap, formatModel, images } = await applyWorkbookXml(xmlString, file.name, true, zip);
    showWorkbookLabel(`${file.name} (manual)`, describeModel(titleMap, formatModel, images), formatModel);

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
async function chooseSaveTarget(suggestedName) {
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
 * applyWorkbookXml() – the one place a workbook becomes the active one (manual 📁 load and
 * Tableau Cloud auto-fetch alike): titles + format model parsed, kept in memory and in browser
 * storage (IndexedDB, with the XML so a newer extension version can parse it again), images and
 * custom shapes extracted. persistToSettings: a manual file is also saved into the workbook's
 * extension settings (old behaviour); the auto-fetch is not – it downloads the latest version on
 * every open.
 * ============================================================================= */
/** titles from the most recent load (auto or manual) @type {any} */
let TITLE_MAP_CACHE = null;

/** the workbook the formatting comes from: { id (Tableau Cloud workbook id, auto-load only), name, xml }
 *  – the other dashboards (tab order, views) are found through it */
let LOADED_WORKBOOK = null;

/** @returns {{ id: string | null, name: string, xml: string } | null} */
function loadedWorkbook() { return LOADED_WORKBOOK; }

/** set by the auto-load once it knows which published workbook this is @param {string} id */
function setLoadedWorkbookId(id) { if (LOADED_WORKBOOK) LOADED_WORKBOOK.id = id; }

/**
 * @param {string} xmlString @param {string} sourceName file / workbook name shown and used for the export's file name
 * @param {boolean} persistToSettings @param {JSZip | null} [zip] the .twbx archive (logos and other images)
 */
async function applyWorkbookXml(xmlString, sourceName, persistToSettings, zip = null) {
  const titleMap = parseTwbXmlInBrowser(xmlString);
  console.log(`[Workbook] ${sourceName}: parsed ${Object.keys(titleMap).length} titles`);
  // one DOM-based pass extracts all formatting (fonts, colours, number formats, shelves, marks …)
  const formatModel = parseTableauFormatting(xmlString);
  TITLE_MAP_CACHE = titleMap;
  FORMAT_MODEL_CACHE = formatModel;
  FORMAT_MODEL_FILE = sourceName;
  LOADED_WORKBOOK = { id: null, name: sourceName, xml: xmlString };       // id is set by the auto-load
  await storeFormatModel(sourceName, formatModel, titleMap, xmlString);
  // logos: only a .twbx carries the image files; custom shapes (button icons) are inside the workbook XML
  const images = { ...(zip ? await extractWorkbookImages(zip, formatModel) : {}), ...extractCustomShapes(xmlString) };
  IMAGE_CACHE = { file: sourceName, images };
  if (Object.keys(images).length) await storeWorkbookImages(sourceName, images);
  if (persistToSettings) {
    tableau.extensions.settings.set("twbTitleMap", JSON.stringify(titleMap));
    try {
      tableau.extensions.settings.set("twbFormatModel", JSON.stringify(formatModel));
    } catch (e) {
      try { tableau.extensions.settings.erase("twbFormatModel"); } catch (e2) { /* nothing saved */ }
      console.warn("[Workbook] Format model too large for settings – kept in memory only:", e.message);
    }
    tableau.extensions.settings.set("twbFileName", sourceName);
    try {
      await tableau.extensions.settings.saveAsync();
    } catch (e) {
      console.warn("[Workbook] settings.saveAsync failed (model kept in memory):", e.message);
    }
  }
  return { titleMap, formatModel, images };
}

/* =============================================================================
 * getTitleMap() - Gets titles from settings   * ============================================================================= */
function getTitleMap() {
  if (TITLE_MAP_CACHE) return TITLE_MAP_CACHE;          // latest auto-fetch or manual load
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
let FORMAT_MODEL_CACHE = null;

/* the .twb / .twbx FORMAT_MODEL_CACHE was read from – the export's default file name starts with it */
let FORMAT_MODEL_FILE = null;

/** @returns {string | null} */
function formatModelFileName() { return FORMAT_MODEL_FILE; }

function getFormatModel() {
  if (FORMAT_MODEL_CACHE) return FORMAT_MODEL_CACHE;
  try {
    const saved = tableau.extensions.settings.get("twbFormatModel");
    FORMAT_MODEL_CACHE = saved ? JSON.parse(saved) : null;
    if (FORMAT_MODEL_CACHE) FORMAT_MODEL_FILE = tableau.extensions.settings.get("twbFileName") || null;
  } catch (err) {
    console.warn("[getFormatModel] Could not read settings:", err.message);
  }
  return FORMAT_MODEL_CACHE;
}

/* workbook model for THIS dashboard: memory / settings first, else the remembered workbook
 * whose worksheets match the dashboard best (no Load Workbook needed again) */
let RESTORED_TITLE_MAP = null;

let RESTORED_FILE_NAME = null;

async function ensureFormatModel() {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  const rank = m => { const r = checkWorkbookMatch(m, dashboard); return !r ? -1 : r.level === "ok" ? 2 + r.matched : r.level === "partial" ? 1 + r.matched / (r.total || 1) : 0; };
  const current = getFormatModel();
  if (current && !isFormatModelStale(current) && rank(current) >= 2) return current;
  const saved = await restoreFormatModels();
  /** @type {{ model: FormatModel, titleMap?: Record<string, string>, fileName: string, xml?: string } | null} */
  let best = null;
  // a stale model in memory / settings loses to the same workbook parsed again by this version
  let bestRank = current ? rank(current) - (isFormatModelStale(current) ? 0.5 : 0) : 0;
  saved.forEach(rec => { const r = rank(rec.model); if (r > bestRank) { best = rec; bestRank = r; } });
  if (best) {
    FORMAT_MODEL_CACHE = best.model;
    FORMAT_MODEL_FILE = best.fileName;
    RESTORED_TITLE_MAP = best.titleMap || null;
    RESTORED_FILE_NAME = best.fileName;
    if (best.xml && !LOADED_WORKBOOK) LOADED_WORKBOOK = { id: null, name: best.fileName, xml: best.xml };
    console.log(`[Workbook] using remembered workbook ${best.fileName} for this dashboard`);
  }
  return FORMAT_MODEL_CACHE;
}

/* main.js restores the cached model from the extension settings at start-up */
function setFormatModelCache(model, fileName) {
  FORMAT_MODEL_CACHE = model;
  FORMAT_MODEL_FILE = fileName || null;
}


/* ═══ visual/icon-sheet.js ════════════════════════════════════════════════════════════════════════ */
/* Button / icon worksheets: a custom shape drawn on empty shelves (zoom, info, navigation icons). Tableau
 * shows the icon; its fields are only flags for actions, so the export draws the icon, not a table. */

/**
 * The icon a worksheet shows when it is only Shape marks on empty Rows / Columns, without text: the shape
 * the workbook maps its current value to, else its Marks card shape. null = not an icon sheet.
 * @param {FormatModel | null} model @param {string} sheetName @param {SummaryData} summary
 * @returns {{ shape: string | null } | null} shape: a custom shape ("Zoom Icons/Zoom in.png"); null = one of
 *   Tableau's own shapes (circle, square …), which the export leaves out
 */
function tvIconSheet(model, sheetName, summary) {
  const sheet = model && model.sheets ? model.sheets[sheetName] : null;
  if (!sheet || !sheet.panes.length || sheet.rows.length || sheet.cols.length) return null;
  if (!sheet.panes.every(p => /^shape$/i.test(p.markClass || ""))) return null;
  if (sheet.panes.some(p => p.labelRuns.length || p.encodings.some(e => e.channel === "text" || e.channel === "label"))) return null;
  const pane = sheet.panes[0];
  let shape = null;
  // Shape = a field: the shape of the value it has now (the workbook's map, sheet before data source)
  const enc = pane.encodings.find(e => e.channel === "shape");
  if (enc) {
    const styles = [sheet.style, ...Object.values(model.datasourceStyles || {}), model.workbookStyle];
    const def = styles.flatMap(s => (s && s.shapes) || []).find(s => tfSameField(s.field, enc.field));
    const fmt = createSheetFormatter(model, sheetName);
    const ci = (summary.columns || []).findIndex(c => tfSameField(fmt.matchName(c.fieldName), enc.field));
    // the first mark whose value has a shape (the marks of a button sit on one spot)
    if (def && ci >= 0) shape = (summary.data || []).map(r => def.map[tfBucketKey(tfDvText(r[ci]))]).find(Boolean) || null;
  }
  // else the Marks card shape
  if (!shape) {
    const f = ((pane.style && pane.style.rules && pane.style.rules.mark) || []).find(x => x.attr === "shape" && !x.field);
    shape = f ? f.value : null;
  }
  return { shape: shape && !shape.startsWith(":") ? shape : null };     // ":filled/circle" … = Tableau's own
}


/* ═══ visual/image-renderer.js ════════════════════════════════════════════════════════════════════ */
/* Tableau-drawn images via createVizImageAsync (spec building, SVG → PNG). */

/* createVizImageAsync input spec, built from the same field roles as the native charts
 * (tvRoles, charts/model/roles.js): real shelves, colour, size, labels and view order.
 * One value axis → v1 spec; several measures (combo / dual axis, Measure Values) → v2 spec.
 * The API draws bar / line / area / square / circle / text marks only and has one field per
 * shelf and no Detail channel, so several category levels are joined into one ordered field. */
const VIZ_IMAGE_DISCRETE_PALETTES = new Set(["tableau10_10_0", "tableau20_10_0", "color_blind_10_0", "seattle_grays_10_0",
  "traffic_light_10_0", "superfishel_stone_10_0", "miller_stone_10_0", "nuriel_stone_10_0", "jewel_bright_10_0", "summer_10_0",
  "winter_10_0", "green_orange_cyan_yellow_10_0", "blue_red_brown_10_0", "purple_pink_gray_10_0", "tableau-10",
  "tableau-10-medium", "tableau-20", "cyclic_10_0"]);

const VIZ_IMAGE_ORDER_FIELD = "__order";

/** @param {VisualModel} visualModel @param {number} width @param {number} height @returns {Record<string, any>} */
function buildVizImageSpec(visualModel, width, height) {
  const vm = visualModel.viewModel;
  const T = VISUAL_TYPES;
  const type = visualModel.type;
  if (!vm || !vm.rows.length) throw new Error("visual has no data rows");
  const roles = tvRoles(vm, visualModel.formatModel, visualModel.source && visualModel.source.visualSpec);
  const markToken = visualModel.metadata.markToken || "";

  // readable, unique field names
  const taken = new Set([VIZ_IMAGE_ORDER_FIELD, "__mark"]);
  const keys = vm.cols.map((c, i) => {
    let k = String(c.label || c.name || `Field${i + 1}`);
    while (taken.has(k)) k += " ";
    taken.add(k);
    return k;
  });
  const kinds = new Map();                                   // column → "discrete" | "continuous"
  const field = (ci, continuous, extra) => {
    kinds.set(ci, continuous ? "continuous" : "discrete");
    return { field: keys[ci], type: continuous ? "continuous" : "discrete", ...extra };
  };
  const title = ci => tvMeasureLabel(vm, ci);

  // colour: TWB palette / mapping → named discrete palette or custom gradient end points
  let color = null;
  if (roles.color && !roles.color.measureNames && roles.color.ci >= 0) {
    const ci = roles.color.ci;
    if (roles.color.continuous) {
      const scale = tvColorScale(vm, roles, markToken);
      const nums = vm.rows.map(r => tfDvNum(r[ci])).filter(n => n !== null);
      if (scale && nums.length) {
        const min = Math.min(...nums), max = Math.max(...nums);
        color = min < 0 && max > 0
          ? field(ci, true, { palette: "custom-diverging", start: "#" + scale(min), end: "#" + scale(max) })
          : field(ci, true, { palette: "custom-sequential", end: "#" + scale(max) });
      } else color = field(ci, true);
    } else {
      const enc = visualModel.formatModel && vm.fmt.colorEncoding ? vm.fmt.colorEncoding() : null;
      const named = enc && enc.def && enc.def.paletteName;
      color = field(ci, false, { palette: VIZ_IMAGE_DISCRETE_PALETTES.has(named) ? named : "tableau10_10_0" });
    }
  }
  const sizeField = roles.size >= 0 ? field(roles.size, true) : null;
  const labelDim = roles.textDims[0] ?? roles.detailDims[0] ?? -1;
  /* No Detail channel: without a per-mark field Tableau would aggregate the marks (a scatter of
   * customers collapses to one point per colour). Label them; with many marks the label is an
   * invisible zero-width key so the marks stay separate without covering the chart. */
  const synthetic = {};
  const markLabel = ci => {
    const distinct = new Set(vm.rows.map(r => tvText(r[ci])));
    if (distinct.size <= 40 || roles.textDims.includes(ci)) return field(ci, false);
    const index = new Map([...distinct].map((v, i) => [v, i]));
    synthetic.__mark = row => index.get(tvText(row[ci])).toString(2).replace(/0/g, "​").replace(/1/g, "‌");
    return { field: "__mark", type: "discrete" };
  };

  /** @type {Record<string, any>} createVizImageAsync input */
  let spec = null;
  let categoryKey = null, categoryOf = null;
  const base = {
    description: (visualModel.title && visualModel.title.text) || visualModel.metadata.worksheetName,
    markcolor: "#" + tvMarkColor(roles)
  };

  if (type === T.MAP) {
    const lat = vm.cols.findIndex(c => /latitude/i.test(c.name));
    const lon = vm.cols.findIndex(c => /longitude/i.test(c.name));
    if (lat < 0 || lon < 0) throw new Error("the map has no latitude/longitude in its summary data");
    const encoding = { columns: field(lon, true, { hidden: true }), rows: field(lat, true, { hidden: true }) };
    if (color) encoding.color = color;
    if (sizeField) encoding.size = sizeField;
    const place = roles.detailDims.find(ci => ci !== lat && ci !== lon);
    if (place !== undefined) encoding.text = markLabel(place);
    spec = { ...base, mark: "circle", encoding };
  } else if (type === T.TREEMAP || type === T.BUBBLE) {
    if (roles.size < 0 || labelDim < 0) throw new Error("no Size measure and label dimension to lay out the marks");
    const encoding = { size: sizeField, text: field(labelDim, false) };
    if (color) encoding.color = color;
    spec = { ...base, mark: type === T.TREEMAP ? "square" : "circle", encoding };
  } else if (type === T.SCATTER) {
    const xm = roles.cols.values[0], ym = roles.rows.values[0];
    if (!xm || !ym) throw new Error("scatter needs a measure on Rows and on Columns");
    const encoding = { columns: field(xm.ci, true, { title: title(xm.ci) }), rows: field(ym.ci, true, { title: title(ym.ci) }) };
    if (color) encoding.color = color;
    if (sizeField) encoding.size = sizeField;
    if (labelDim >= 0) encoding.text = markLabel(labelDim);
    spec = { ...base, mark: "circle", encoding };
  } else {
    // bar / line / area / histogram / combo: one ordered category field + one or more value axes
    const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
    if (!valueShelf) throw new Error("no continuous measure axis to draw");
    const catShelf = valueShelf === "rows" ? "cols" : "rows";
    const measures = roles[valueShelf].values;
    const catCis = [...roles[catShelf].dims, ...roles[valueShelf].dims].map(d => d.ci);
    if (catCis.length) {
      const order = new Map();
      categoryOf = row => catCis.map(ci => tvText(row[ci])).join(" · ");
      vm.rows.forEach(r => { const k = categoryOf(r); if (!order.has(k)) order.set(k, order.size); });
      if (catCis.some(ci => /^date/i.test(String(vm.cols[ci].dataType || "")))) {
        const firstRow = new Map();
        vm.rows.forEach(r => { const k = categoryOf(r); if (!firstRow.has(k)) firstRow.set(k, r); });
        const sorted = [...order.keys()].sort((a, b) => {
          for (const ci of catCis) { const d = tfNaturalCompare(firstRow.get(a)[ci], firstRow.get(b)[ci]); if (d) return d; }
          return 0;
        });
        order.clear();
        sorted.forEach((k, i) => order.set(k, i));
      }
      categoryKey = { name: catCis.map(ci => keys[ci]).join(" · "), order };
    }
    const catField = categoryKey ? { field: categoryKey.name, type: "discrete" } : null;
    const markOf = token => token === "line" ? "line" : token === "area" ? "area"
      : /^(circle|shape)$/.test(token) ? "circle" : token === "square" ? "square" : "bar";
    const shelfName = s => s === "cols" ? "columns" : "rows";
    if (measures.length === 1) {
      const m = measures[0];
      const encoding = {
        [shelfName(valueShelf)]: field(m.ci, true, { title: title(m.ci) }),
        sort: catField ? { field: categoryKey.name, sortby: VIZ_IMAGE_ORDER_FIELD, direction: "ascending" } : undefined
      };
      if (catField) encoding[shelfName(catShelf)] = catField;
      if (color) encoding.color = color;
      if (sizeField) encoding.size = sizeField;
      if (labelDim >= 0 && !catCis.includes(labelDim)) encoding.text = markLabel(labelDim);
      if (!encoding.sort) delete encoding.sort;
      spec = { ...base, mark: markOf(tvMeasureMark(roles, m.ref, markToken)), encoding };
    } else {
      const marks = measures.map((m, i) => markOf(tvMeasureMark(roles, m.ref, markToken, i, measures.length)));
      measures.forEach(m => field(m.ci, true));
      spec = {
        version: 2,
        description: base.description,
        vizlayout: { size: { width, height }, showcolorlegend: !!color },
        [shelfName(catShelf)]: catField ? [catField] : [],
        [shelfName(valueShelf)]: measures.map(m => ({ field: keys[m.ci], type: "continuous", title: title(m.ci) })),
        encodingaxis: shelfName(valueShelf),
        defaultencoding: { mark: marks[0] },
        encodings: marks.map(mark => (color ? { mark, color } : { mark }))
      };
    }
  }
  if (!spec.version) spec.size = { width, height };

  // data: every encoded column + the joined category and its order helper
  const perCategory = new Map();
  if (categoryKey) vm.rows.forEach(r => { const k = categoryOf(r); perCategory.set(k, (perCategory.get(k) || 0) + 1); });
  spec.data = {
    values: vm.rows.map(row => {
      const out = {};
      kinds.forEach((kind, ci) => {
        out[keys[ci]] = kind === "continuous" ? tfDvNum(row[ci]) : tvText(row[ci]);
      });
      for (const name in synthetic) out[name] = synthetic[name](row);
      if (categoryKey) {
        const k = categoryOf(row);
        out[categoryKey.name] = k;
        out[VIZ_IMAGE_ORDER_FIELD] = categoryKey.order.get(k) / perCategory.get(k);   // SUM over the category = its rank
      }
      return out;
    })
  };
  return spec;
}

function svgToPngDataUrl(svg, width, height) {
  return new Promise((resolve, reject) => {
    const blob = new Blob([svg], { type: "image/svg+xml" });
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        canvas.getContext("2d").drawImage(image, 0, 0, width, height);
        URL.revokeObjectURL(url);
        resolve(canvas.toDataURL("image/png"));
      } catch (err) {
        URL.revokeObjectURL(url);
        reject(err);
      }
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Tableau returned an SVG that could not be rasterized"));
    };
    image.src = url;
  });
}

/** @param {VisualModel} visualModel @param {number} width @param {number} height @returns {Promise<string>} PNG, base64 */
async function renderTableauImage(visualModel, width, height) {
  if (typeof tableau === "undefined" || !tableau.extensions ||
      typeof tableau.extensions.createVizImageAsync !== "function") {
    throw new Error("Tableau image API is unavailable in this runtime");
  }
  const spec = buildVizImageSpec(visualModel, width, height);
  const svg = await tableau.extensions.createVizImageAsync(spec);
  if (typeof svg !== "string" || !svg.trim()) throw new Error("Tableau image API returned no SVG");
  return svgToPngDataUrl(svg, width, height);
}


/* ═══ visual/matrix.js ════════════════════════════════════════════════════════════════════════════ */
/* Heat maps, highlight tables, calendars and crosstabs as Tableau draws them: a matrix – the Rows shelf's members
 * down, the Columns shelf's members across, one cell per mark holding its label (else its size measure, else its
 * colour measure, else its shape), filled in the mark's colour. And the guard that keeps a table of marks from
 * dropping what the marks show. */

/** @param {SheetModel} sheet the panes Tableau draws: each with an id when there are several */
const drawnPanes = sheet => sheet.panes.length > 1 && sheet.panes.some(p => p.id) ? sheet.panes.filter(p => p.id) : sheet.panes;

/**
 * Pivots the view model into the matrix Tableau shows, in place. Only a view of dimensions on both shelves (no axes, no
 * Measure Names) whose marks each hold one value; anything else keeps its one-row-per-mark table.
 * @param {ViewModel} vm @param {FormatModel | null} model @returns {boolean} pivoted
 */
function pivotMatrix(vm, model) {
  const fmt = vm.fmt, sheet = fmt.sheetModel;
  if (!sheet || vm.markTable) return false;
  const shelf = [...sheet.rows, ...sheet.cols];
  if (shelf.some(r => r.name === "Measure Names" || r.type === "qk" || tvIsMeasureRef(model, r))) return false;
  const colRefs = sheet.cols;
  const ciOf = ref => vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
  const colCis = colRefs.map(ciOf);
  if (!colRefs.length || colCis.some(ci => ci < 0)) return false;
  const rowCis = vm.headerOrder.filter(ci => !colCis.includes(ci) && sheet.rows.some(r => tfSameField(r, vm.cols[ci].ref)));
  if (!rowCis.length) return false;
  const panes = drawnPanes(sheet);
  if (panes.length !== 1) return false;
  const pane = panes[0];
  const enc = ch => pane.encodings.filter(e => e.channel === ch).map(e => e.field);
  const text = [...enc("text"), ...enc("label"), ...pane.labelRuns.flatMap(r => r.refs)].filter((r, i, a) => a.findIndex(x => tfSameField(x, r)) === i);
  if (text.length > 1) return false;
  const size = enc("size").find(r => tvIsMeasureRef(model, r));
  const color = enc("color")[0] || null;
  const valueRef = text[0] || size || (color && tvIsMeasureRef(model, color) ? color : null) || enc("shape")[0];
  const valueCi = valueRef ? ciOf(valueRef) : -1;
  if (valueCi < 0) return false;
  const colorCi = color ? ciOf(color) : -1;

  // one mark per cell: a cell that would need two values is not a matrix
  const colKey = r => colCis.map(ci => tfDvText(r[ci])).join("\u0001");
  const rowKey = r => rowCis.map(ci => tfDvText(r[ci])).join("\u0001");
  /** @type {Map<string, { head: DataValue[], cells: Map<string, DataValue[]> }>} */
  const groups = new Map();
  for (const r of vm.rows) {
    const k = rowKey(r);
    if (!groups.has(k)) groups.set(k, { head: r, cells: new Map() });
    const cells = groups.get(k).cells, ck = colKey(r);
    if (cells.has(ck) && tfDvText(cells.get(ck)[valueCi]) !== tfDvText(r[valueCi])) return false;
    cells.set(ck, r);
  }
  // the Columns members in Tableau's order: manual sort, else natural
  /** @type {Map<string, DataValue[]>} */
  const tuples = new Map();
  vm.rows.forEach(r => { const k = colKey(r); if (!tuples.has(k)) tuples.set(k, colCis.map(ci => r[ci])); });
  const ranks = colRefs.map(ref => {
    const m = fmt.manualSortFor(ref);
    return m ? { map: new Map(m.order.map((b, i) => [tfBucketKey(b), i])), desc: m.direction === "DESC" } : null;
  });
  const keys = [...tuples.keys()].sort((a, b) => {
    const ta = tuples.get(a), tb = tuples.get(b);
    for (let l = 0; l < colCis.length; l++) {
      let d = 0;
      if (ranks[l]) {
        const ra = ranks[l].map.get(tfNorm(tfDvText(ta[l]))), rb = ranks[l].map.get(tfNorm(tfDvText(tb[l])));
        d = (ra === undefined ? 1e9 : ra) - (rb === undefined ? 1e9 : rb);
        if (ranks[l].desc) d = -d;
      }
      if (!d) d = tfNaturalCompare(ta[l], tb[l]);
      if (d) return d;
    }
    return 0;
  });

  // each mark's colour, from the colour field over all marks (Tableau's palette / range)
  /** @type {(row: DataValue[]) => string | null} */
  let colorAt = () => null;
  const colorEnc = color && colorCi >= 0 ? fmt.colorEncodingOf(pane) : null;
  if (colorEnc) {
    const pick = dv => tfIsNull(dv) ? null : colorEnc.continuous ? tfDvNum(dv) : tfDvText(dv);
    const scale = tfBuildColorScale(fmt, colorEnc, vm.rows.map(r => pick(r[colorCi])));
    if (scale) colorAt = r => { const v = pick(r[colorCi]); return v === null ? null : scale(v) || null; };
  }

  const value = vm.cols[valueCi];
  const blank = { value: null, nativeValue: null, formattedValue: "" };
  /** @type {ViewColumn[]} */
  const cols = [...rowCis.map(ci => vm.cols[ci]),
    ...keys.map(k => {
      const label = tuples.get(k).map(tfDvText).join(" / ");
      return { name: label, label, ref: value.ref, dataType: value.dataType, pivoted: true, isHeader: false };
    })];
  /** @type {(string | null)[][]} */
  const fills = [];
  const rows = [...groups.values()].map(g => {
    fills.push(keys.map(k => g.cells.has(k) ? colorAt(g.cells.get(k)) : null));
    return [...rowCis.map(ci => g.head[ci]), ...keys.map(k => g.cells.has(k) ? g.cells.get(k)[valueCi] : blank)];
  });
  const shownRow = rowCis.map((ci, i) => vm.order.includes(ci) ? i : -1).filter(i => i >= 0);
  vm.cols = cols;
  vm.rows = rows;
  vm.order = [...shownRow, ...keys.map((_, k) => rowCis.length + k)];
  vm.headerOrder = rowCis.map((_, i) => i);
  vm.showHeaderRow = true;
  vm.matrix = { colDims: colCis, valueCi, colorCi };
  vm.cellFill = (rowIdx, ci) => ci >= rowCis.length && fills[rowIdx] ? fills[rowIdx][ci - rowCis.length] : null;
  vm.notes.push(`matrix: ${groups.size} rows × ${keys.length} columns of ${value.name}`);
  return true;
}

/**
 * A table of marks must show what the marks show: when its cells would hold no value at all while the marks encode
 * measures (circles sized and coloured on a timeline), the table is lossy – the export writes every column instead
 * (TABLE_FALLBACK). Sets vm.lossless / vm.fallbackReason.
 * @param {ViewModel} vm @param {FormatModel | null} model @returns {boolean}
 */
function checkLossless(vm, model) {
  const sheet = vm.fmt.sheetModel;
  if (!sheet || vm.markTable || vm.matrix) return true;
  if (vm.order.some(ci => !vm.cols[ci].isHeader)) return true;           // the cells carry values (coloured as the marks)
  // measures by colour / size / label, and any field picking the marks' shape (a KPI's ▲▼)
  const encoded = drawnPanes(sheet).flatMap(p => p.encodings.filter(e => (/^(color|size|text|label)$/.test(e.channel) && tvIsMeasureRef(model, e.field)) ||
                                                                         e.channel === "shape")
    .map(e => ({ channel: e.channel === "label" ? "text" : e.channel, ref: e.field })));
  if (!encoded.length) return true;
  const names = [...new Set(encoded.map(e => {
    const c = vm.cols.find(x => x.ref && tfSameField(x.ref, e.ref));
    return `${c ? c.name : e.ref.name} (${e.channel})`;
  }))];
  vm.lossless = false;
  vm.fallbackReason = `its marks show ${names.join(", ")}, which a table of its headers would leave out – every column exported`;
  return false;
}


/* ═══ visual/view-model.js ════════════════════════════════════════════════════════════════════════ */
/* Summary data → view model (pivot, merge, sort, visible columns, titles). */

/**
 * Dashboard text boxes over a table built from marks, where one box can head several panes (e.g. a
 * diverging bar split into a negative and a positive axis): adjacent value columns whose panes are coloured
 * by the same field share a box. Matched only when the groups and the boxes line up one to one; the row
 * headers take the width left of the first box. Sets the columns' labels, widths and header spans.
 * @param {any} dash @param {string} sheetName @param {ViewColumn[]} cols @param {number[]} order
 * @param {SheetFormatter} fmt @returns {{ ids: string[], headerPx: number } | null}
 */
function markTableHeaders(dash, sheetName, cols, order, fmt) {
  const strip = tfTextBoxHeaderStrip(dash, sheetName);
  if (!strip || !dash.width) return null;
  const panes = fmt.panesInOrder();
  const colourOf = ci => {
    const p = panes.find(x => x.refs.some(r => tfSameField(r, cols[ci].ref)));
    const e = p && p.pane.encodings.find(x => x.channel === "color" && x.field);
    return e ? tfRefKey(e.field) : null;
  };
  const values = order.filter(ci => !cols[ci].isHeader), heads = order.filter(ci => cols[ci].isHeader);
  /** @type {{ key: string | null, cols: number[] }[]} */
  const groups = [];
  values.forEach(ci => {
    const key = colourOf(ci), last = groups[groups.length - 1];
    if (last && key && last.key === key) last.cols.push(ci); else groups.push({ key, cols: [ci] });
  });
  if (groups.length !== strip.zones.length) return null;
  const px = units => units / 100000 * dash.width;
  groups.forEach((g, k) => {
    const z = strip.zones[k], first = cols[g.cols[0]];
    first.label = z.text;
    first.labelProps = z.props;
    first.labelRuns = z.runs;
    first.headerSpan = g.cols.length;
    g.cols.forEach((ci, j) => {
      if (j) { cols[ci].label = ""; cols[ci].headerCovered = true; }
      cols[ci].zoneWidthPx = px(z.w) / g.cols.length;
    });
  });
  const left = px(strip.zones[0].x - strip.ws.x);
  if (heads.length && left > 0) heads.forEach(ci => { cols[ci].zoneWidthPx = left / heads.length; });
  return { ids: strip.zones.map(z => z.id), headerPx: Math.max(...strip.zones.map(z => z.h)) / 100000 * (dash.height || 0) };
}

/* Tableau names a quick table calculation after its measure: a running sum and the plain sum both arrive as
 * "SUM(Value Ordered)". Of the columns sharing a name, the one holding the calculation takes the sheet's
 * table-calculation field: a running sum holds the plain column's grand total (its last value), a percent of
 * total adds up to 1; otherwise the calculation is the first of them, as Tableau lists it (cum… before sum…). */
function splitTableCalcColumns(cols, rows, fmt) {
  const groups = new Map();
  cols.forEach((c, i) => { if (c.ref) { const k = tfNorm(c.name); groups.set(k, [...(groups.get(k) || []), i]); } });
  groups.forEach(idx => {
    if (idx.length < 2 || !fmt.hasModel || idx.every(i => cols[i].exact)) return;
    const calcs = fmt.tableCalcRefs(cols[idx[0]].ref);
    if (!calcs.length) return;
    const nums = i => rows.map(r => tfDvNum(r[i])).filter(v => v !== null);
    const sum = i => nums(i).reduce((a, b) => a + b, 0);
    const near = (a, b) => Math.abs(a - b) <= Math.max(1e-9, Math.abs(b) * 1e-6);
    let free = [...idx];
    calcs.forEach(calc => {
      if (free.length < 2) return;
      const test = fmt.isRunningTotal(calc) ? (i, j) => nums(i).some(v => near(v, sum(j)))
                 : /^pcto$/i.test(calc.deriv || "") ? (i) => near(sum(i), 1) || near(sum(i), 100) : null;
      let pick = test ? free.find(i => free.some(j => j !== i && test(i, j))) : undefined;
      if (pick === undefined) pick = free[0];
      cols[pick].ref = calc;
      free = free.filter(i => i !== pick);
    });
  });
}

/**
 * @param {FormatModel | null} model
 * @param {string} sheetName
 * @param {SummaryData} summary
 * @param {{ dashboardName?: string, displayName?: string, [key: string]: any }} [opts]
 * @returns {ViewModel}
 */
function buildViewModel(model, sheetName, summary, opts = {}) {
  const fmt = createSheetFormatter(model, sheetName);
  const notes = [];
  /** @type {ViewColumn[]} */
  // a column's field: its id when Tableau gives one ("[ds].[pcto:sum:Sales:qk]" – exact, even for table
  // calculations captioned like their measure), else its caption ("SUM(Sales)")
  let cols = (summary.columns || []).map((c, i) => {
    const name = c.fieldName || c.fieldId || `Col${i + 1}`;
    const byId = fmt.matchFieldId(c.fieldId);
    return { name, dataType: c.dataType, ref: byId || fmt.matchName(name), exact: !!byId };
  });
  let rows = (summary.data || []).map(r => r.slice());
  splitTableCalcColumns(cols, rows, fmt);

  // ── 1. pivot Measure Names / Measure Values ─────────────────────────────
  const mnI = cols.findIndex(c => /^measure names$/i.test(c.name));
  const mvI = cols.findIndex(c => /^measure values$/i.test(c.name));
  if (mnI >= 0 && mvI >= 0) {
    const keyIdx = cols.map((_, i) => i).filter(i => i !== mnI && i !== mvI);
    const order = fmt.measureOrder();
    const measures = [];                         // [{text, ref, rank}]
    const groups = new Map();
    rows.forEach(r => {
      const text = tfDvText(r[mnI]);
      if (!measures.some(m => m.text === text)) {
        const ref = fmt.measureRefForName(text);
        const pos = ref ? order.findIndex(o => tfSameField(o, ref)) : -1;
        measures.push({ text, ref, rank: pos >= 0 ? pos : 1e6 + measures.length });
      }
      const key = keyIdx.map(i => tfDvText(r[i])).join("\u0001");
      if (!groups.has(key)) groups.set(key, { keyRow: keyIdx.map(i => r[i]), vals: {} });
      const g = groups.get(key);
      if (g.vals[text] === undefined) g.vals[text] = r[mvI];
    });
    measures.sort((a, b) => a.rank - b.rank);
    cols = [...keyIdx.map(i => cols[i]),
            ...measures.map(m => ({ name: m.text, ref: m.ref, pivoted: true, dataType: "float" }))];
    rows = [...groups.values()].map(g => [...g.keyRow, ...measures.map(m => g.vals[m.text] || null)]);
    notes.push(`pivoted ${measures.length} measures → ${rows.length} rows`);
  }

  // column roles
  cols.forEach(c => {
    c.isHeader = !c.pivoted && fmt.hasModel ? fmt.isHeaderField(c.ref)
               : !c.pivoted && !/^(int|float|real|integer|number)$/i.test(String(c.dataType || ""));
  });

  // ── 2. merge stacked multi-pane rows ────────────────────────────────────
  const hdrIdx = cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);
  if (hdrIdx.length && rows.length > 1) {
    const groups = new Map();
    rows.forEach(r => {
      const key = hdrIdx.map(i => tfDvText(r[i])).join("\u0001");
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    });
    if (groups.size < rows.length) {
      let conflict = false;
      const merged = [];
      for (const g of groups.values()) {
        const out = g[0].slice();
        for (let ci = 0; ci < cols.length && !conflict; ci++) {
          const vals = [...new Set(g.filter(r => !tfIsNull(r[ci])).map(r => tfDvText(r[ci])))];
          if (vals.length > 1) conflict = true;
          else { const hit = g.find(r => !tfIsNull(r[ci])); if (hit) out[ci] = hit[ci]; }
        }
        if (conflict) break;
        merged.push(out);
      }
      if (!conflict) { notes.push(`merged ${rows.length} stacked rows → ${merged.length}`); rows = merged; }
      else notes.push("rows share header values but differ – not merged");
    }
  }

  // ── 3. sort like the view ───────────────────────────────────────────────
  const sortLevels = [];
  (fmt.hasModel ? fmt.headerRefs() : []).forEach(ref => {
    const ci = cols.findIndex(c => c.isHeader && c.ref && tfSameField(c.ref, ref));
    if (ci < 0) return;
    const manual = fmt.manualSortFor(ref);
    const byMeasure = fmt.measureSortFor(ref);
    let mi = -1;
    if (byMeasure) mi = cols.findIndex(c => c.ref && tfSameField(c.ref, byMeasure.measure));
    // sort measure not in the summary data (e.g. a funnel sorted by a hidden field) → the row's own total
    sortLevels.push({ ci, manual, byMeasure, mi });
  });
  if (sortLevels.length) {
    // measure sort aggregates the measure over the rows sharing the same outer path
    const aggCache = sortLevels.map((lvl, L) => {
      if (!lvl.byMeasure) return null;
      const m = new Map();
      rows.forEach(r => {
        const k = sortLevels.slice(0, L + 1).map(x => tfDvText(r[x.ci])).join("\u0001");
        const v = lvl.mi >= 0 ? tfDvNum(r[lvl.mi])
          : cols.reduce((s, c, i) => s + (!c.isHeader && tfDvNum(r[i]) !== null ? tfDvNum(r[i]) : 0), 0);
        m.set(k, (m.get(k) || 0) + (v || 0));
      });
      return m;
    });
    const manualRank = lvl => {
      if (!lvl.manual) return null;
      const map = new Map(lvl.manual.order.map((b, i) => [tfBucketKey(b), i]));
      return map;
    };
    const ranks = sortLevels.map(manualRank);
    rows.sort((a, b) => {
      for (let L = 0; L < sortLevels.length; L++) {
        const lvl = sortLevels[L];
        let d = 0;
        if (ranks[L]) {
          const ra = ranks[L].get(tfNorm(tfDvText(a[lvl.ci]))), rb = ranks[L].get(tfNorm(tfDvText(b[lvl.ci])));
          d = (ra === undefined ? 1e9 : ra) - (rb === undefined ? 1e9 : rb);
          if (lvl.manual.direction === "DESC") d = -d;
        } else if (lvl.byMeasure) {
          const path = r => sortLevels.slice(0, L + 1).map(x => tfDvText(r[x.ci])).join("\u0001");
          d = (aggCache[L].get(path(a)) || 0) - (aggCache[L].get(path(b)) || 0);
          if (lvl.byMeasure.direction === "DESC") d = -d;
        }
        if (!d) d = tfNaturalCompare(a[lvl.ci], b[lvl.ci]);
        if (d) return d;
      }
      return 0;
    });
  }

  // ── URL actions → link columns (kept visible even when the field is only on Detail)
  const linkActions = fmt.urlActions(opts.dashboardName);
  linkActions.forEach(a => {
    const refs = tfExtractRefs(a.expression);
    if (!refs.length) return;
    const ci = cols.findIndex(c => !c.isHeader && c.ref && tfNorm(c.ref.name) === tfNorm(refs[0].name));
    if (ci < 0) return;
    cols[ci].link = { caption: a.caption, expression: a.expression,
      parts: refs.map(r => ({ token: r, ci: cols.findIndex(c => c.ref && tfNorm(c.ref.name) === tfNorm(r.name)) })) };
  });

  // ── 4. visible columns in visual order + header labels ──────────────────
  let order;
  if (!fmt.hasModel) {
    order = cols.map((_, i) => i);
  } else {
    const panes = fmt.panesInOrder();
    const headerRefs = fmt.headerRefs();
    /** @type {{ i: number, g: number, a: number, b: number }[]} column, group, then sort keys */
    const place = [];
    cols.forEach((c, i) => {
      if (c.isHeader) {
        if (fmt.isLabelHidden(c.ref)) return;                       // Format → Hide header
        const pos = headerRefs.findIndex(r => tfSameField(r, c.ref));
        place.push({ i, g: 0, a: pos, b: 0 });
      } else if (c.pivoted) {
        place.push({ i, g: 1, a: 0, b: i });
      } else if (!c.ref) {
        place.push({ i, g: 2, a: 1e6, b: i });                     // unknown → keep, at the end
      } else {
        /** @type {{ a: number, b: number } | null} */
        let hit = null;
        panes.forEach((p, pi) => {
          const k = p.refs.findIndex(r => tfSameField(r, c.ref));
          if (k >= 0 && !hit) hit = { a: pi, b: k };
        });
        if (hit) place.push({ i, g: 2, a: hit.a, b: hit.b });      // shown as a mark label
        else if (c.link) {                                         // link column: in the pane that carries the field
          const pi = panes.findIndex(p => p.pane.encodings.some(e => tfSameField(e.field, c.ref)));
          place.push({ i, g: 2, a: pi >= 0 ? pi : 1e5, b: 999 });
        }
        // else: detail / tooltip / colour-only / axis measure → not visible
      }
    });
    place.sort((x, y) => x.g - y.g || x.a - y.a || x.b - y.b);
    order = place.map(p => p.i);
    if (!order.length) { order = cols.map((_, i) => i); notes.push("no visible columns resolved – exporting all"); }
  }

  const labelsRows = fmt.fieldLabelsShown("rows"), labelsCols = fmt.fieldLabelsShown("cols");
  cols.forEach(c => {
    if (c.pivoted) c.label = c.name;                                          // Measure Names alias
    else if (c.isHeader) {
      const shown = fmt.shelfOf(c.ref) === "cols" ? labelsCols : labelsRows;
      c.label = shown ? fmt.captionFor(c.ref, c.name) : "";
    } else c.label = fmt.captionFor(c.ref, c.name);
    if (c.link && c.link.caption) c.label = c.link.caption;
  });

  // ── dashboard text boxes drawn as column headers (strict match, else captions stay) ──
  const dashM = model && opts.dashboardName && model.dashboards && model.dashboards[opts.dashboardName];
  let headerZoneIds = [];
  let headerPx = 0;
  if (FORMAT_CONFIG.textBoxHeaders && dashM && !labelsRows) {
    const tb = tfTextBoxHeaders(dashM, sheetName, order.length);
    if (tb) {
      order.forEach((ci, k) => {
        const z = tb.zones[k];
        cols[ci].label = z.text;
        cols[ci].labelProps = z.props;
        if (dashM.width) cols[ci].zoneWidthPx = z.w / 100000 * dashM.width;
      });
      headerZoneIds = tb.zones.map(z => z.id);
      notes.push(`headers from ${tb.zones.length} dashboard text boxes`);
    } else {
      const grouped = markTableHeaders(dashM, sheetName, cols, order, fmt);
      if (grouped) {
        headerZoneIds = grouped.ids;
        headerPx = grouped.headerPx;
        notes.push(`headers from ${grouped.ids.length} dashboard text boxes over grouped panes`);
      }
    }
  }
  const showHeaderRow = order.some(i => cols[i].label);

  // title
  const valuesFor = ref => {
    const ci = cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
    if (ci < 0) return [];
    return [...new Set(rows.map(r => tfDvText(r[ci])).filter(Boolean))];
  };
  const title = fmt.renderTitle(valuesFor, opts.displayName || sheetName);
  const dash = model && opts.dashboardName && model.dashboards && model.dashboards[opts.dashboardName];
  const zone = dash && dash.zones.find(z => z.name === sheetName && z.type === "worksheet");
  const showTitle = zone ? zone.showTitle : true;

  if (notes.length) tfLog(`View model "${sheetName}": ${notes.join("; ")}`);
  return {
    fmt, cols, rows, order, notes, showHeaderRow, title, showTitle, headerZoneIds, headerPx,
    dashboardName: opts.dashboardName,
    kind: fmt.isChart() ? "chart" : "table",
    headerOrder: sortLevels.map(l => l.ci)          // all header cols (incl. hidden) outer → inner
  };
}


/* ═══ visual/visual-model.js ══════════════════════════════════════════════════════════════════════ */
/* View model + classification → visual model, and renderer selection. */

/**
 * @param {FormatModel | null} model
 * @param {string} sheetName
 * @param {SummaryData} summary
 * @param {{ dashboardName?: string, displayName?: string, visualSpec?: any, visualSpecError?: any, layout?: any }} [opts]
 * @returns {VisualModel}
 */
function buildVisualModel(model, sheetName, summary, opts = {}) {
  const vm = buildViewModel(model, sheetName, summary, opts);
  const spec = opts.visualSpec || null;
  const marks = resolveVisualMarks(spec, vm, model);
  const evidence = [];
  const type = classifyVisualType(spec, vm, model, evidence);
  vm.markTable = (type === VISUAL_TYPES.TABLE || type === VISUAL_TYPES.KPI) && isMarkTable(vm, model);
  // drawn as cells: a heat map / crosstab as Tableau's matrix, and never a table that drops what its marks show
  if (type === VISUAL_TYPES.HEATMAP || type === VISUAL_TYPES.TABLE) {
    if (pivotMatrix(vm, model)) evidence.push("dimensions on Rows and Columns, one mark per cell: a matrix");
    else checkLossless(vm, model);
  }
  const dimensions = vm.order.filter(i => vm.cols[i].isHeader).map(i => vm.cols[i]);
  const measures = vm.order.filter(i => !vm.cols[i].isHeader).map(i => vm.cols[i]);
  return {
    type,
    title: vm.title,
    data: { rows: vm.rows, columns: vm.cols, dimensions, measures },
    axes: spec && (spec.axes || spec.axis) || {},
    encodings: spec && (spec.encodings || spec.encoding) || {},
    panes: spec && spec.panes || [],
    style: spec && (spec.style || spec.styles) || {},
    layout: opts.layout || null,
    source: { visualSpec: spec },
    formatModel: model || null,
    metadata: {
      worksheetName: sheetName,
      markClass: vm.fmt && vm.fmt.colorEncoding ? ((vm.fmt.colorEncoding() || {}).markClass || null) : null,
      markToken: marks.tokens[0] || "",
      markTokens: marks.tokens,
      markSource: marks.source,
      visualSpecAvailable: !!spec,
      visualSpecError: opts.visualSpecError ? String(opts.visualSpecError.message || opts.visualSpecError) : null,
      visualSpecKeys: spec && typeof spec === "object" ? Object.keys(spec) : [],
      imageApiAvailable: typeof tableau !== "undefined" && tableau.extensions &&
        typeof tableau.extensions.createVizImageAsync === "function"
    },
    viewModel: vm,
    diagnostics: [],
    evidence
  };
}

/* cell-native → cells; mappable charts → native Excel chart; the rest → Tableau image → data table */
/** @param {VisualModel} visualModel @returns {RendererDecision} */
function chooseVisualRenderer(visualModel) {
  const type = visualModel.type;
  if (VISUAL_RENDERERS.cellTypes.has(type)) return VISUAL_RENDERERS.cell;
  if (FORMAT_CONFIG.nativeCharts && VISUAL_RENDERERS.chartTypes.has(type)) return VISUAL_RENDERERS.chart;
  return imageOrFallback(visualModel);
}

/** @param {VisualModel} visualModel @param {string} [why] @returns {RendererDecision} */
function imageOrFallback(visualModel, why) {
  const prefix = why ? why + "; " : "";
  if (VISUAL_RENDERERS.imageTypes.has(visualModel.type) && visualModel.metadata.imageApiAvailable) {
    return { ...VISUAL_RENDERERS.image, reason: prefix + VISUAL_RENDERERS.image.reason };
  }
  const reason = !visualModel.metadata.imageApiAvailable && VISUAL_RENDERERS.imageTypes.has(visualModel.type)
    ? "Tableau image API unavailable – exported as data"
    : `${visualModel.type.toLowerCase()} visuals cannot be drawn in Excel – exported as data`;
  return { ...VISUAL_RENDERERS.fallback, reason: prefix + reason };
}

const VISUAL_RENDERERS = Object.freeze({
  cell: Object.freeze({ renderer: "cell", status: "success", reason: "cell-native visual" }),
  chart: Object.freeze({ renderer: "excel-chart", status: "pending", reason: "native Excel chart" }),
  image: Object.freeze({ renderer: "tableau-image", status: "pending", reason: "Tableau image renderer selected" }),
  cellTypes: /** @type {Set<VisualType>} */ (new Set([VISUAL_TYPES.TABLE, VISUAL_TYPES.KPI, VISUAL_TYPES.HEATMAP])),
  chartTypes: /** @type {Set<VisualType>} */ (new Set([VISUAL_TYPES.BAR, VISUAL_TYPES.COLUMN, VISUAL_TYPES.LINE, VISUAL_TYPES.AREA,
                       VISUAL_TYPES.PIE, VISUAL_TYPES.SCATTER, VISUAL_TYPES.COMBO, VISUAL_TYPES.HISTOGRAM,
                       VISUAL_TYPES.WATERFALL, VISUAL_TYPES.BOXPLOT, VISUAL_TYPES.GANTT, VISUAL_TYPES.TREEMAP,
                       VISUAL_TYPES.BUBBLE])),
  // createVizImageAsync draws bar / line / area / square / circle / text marks only (no pie, gantt or polygon
  // marks). Maps – symbol or filled – are neither: without the basemap they don't read as maps, so they are
  // exported as a table of their marks (data-fallback)
  imageTypes: /** @type {Set<VisualType>} */ (new Set([VISUAL_TYPES.BAR, VISUAL_TYPES.COLUMN, VISUAL_TYPES.LINE, VISUAL_TYPES.AREA, VISUAL_TYPES.SCATTER,
                       VISUAL_TYPES.COMBO, VISUAL_TYPES.HISTOGRAM, VISUAL_TYPES.TREEMAP, VISUAL_TYPES.BUBBLE])),
  fallback: Object.freeze({ renderer: "data-fallback", status: "warning", reason: "visual renderer not implemented yet" })
});


/* ═══ visual/semantics.js ═════════════════════════════════════════════════════════════════════════ */
/* What a Tableau visual is, in Tableau's terms, and the evidence for it: the shelves, marks and encodings of the
 * worksheet read as one visual ("Stacked Bar", "Bump Chart", "Filled Map" …). The conversion report shows it; the
 * classifier (classify.js) decides how the visual is drawn, this names what it is. */

/* quick table calculations, as their derivation prefix ("pcto:sum:Sales:qk") */
const TABLE_CALCS = { pcto: "% of Total", cum: "Running Total", rsum: "Running Total", rank: "Rank", pctrank: "Percentile",
                      diff: "Difference", pdiff: "% Difference", mavg: "Moving Average", ytd: "YTD Total",
                      cgr: "Compound Growth Rate", yoy: "Year over Year Growth", ytdgr: "YTD Growth" };

const GEO_GENERATED = /^(latitude|longitude) \(generated\)$/i;

/**
 * A field reference taken apart: its quick table calculations (outer → inner), its aggregation or date part, the
 * underlying field and how Tableau labels it ("% of Total of SUM(Sales)").
 * @param {FormatModel | null} model @param {FieldRef} ref
 */
function describeField(model, ref) {
  if (ref.name === "Measure Names") return { label: "Measure Names", calcs: [], agg: null, info: null, name: ref.name };
  if (ref.name === "Multiple Values") return { label: "Measure Values", calcs: [], agg: null, info: null, name: ref.name };
  const chain = ref.deriv ? [ref.deriv.toLowerCase()] : [];
  let name = ref.name || "";
  for (let m = name.match(/^([a-z_]+):(.+)$/i); m && (TABLE_CALCS[m[1].toLowerCase()] || TF_DERIV_LABEL[m[1].toLowerCase()] ||
       /^(none|usr)$/i.test(m[1])); m = name.match(/^([a-z_]+):(.+)$/i)) {
    chain.push(m[1].toLowerCase());
    name = m[2];
  }
  const calcs = chain.filter(c => TABLE_CALCS[c]);
  const agg = chain.find(c => !TABLE_CALCS[c]) || null;
  const info = model ? tfFieldInfo(model, { ds: ref.ds, deriv: null, name, inner: name, type: ref.type }) : null;
  const caption = (info && info.caption) || (/^Calculation_\d+$/i.test(name) ? "Calculation" : name.replace(/_/g, " "));
  const AGG = !agg || agg === "none" ? null : agg === "usr" ? "AGG" : (TF_DERIV_LABEL[agg] || agg.toUpperCase());
  let label = AGG ? `${AGG}(${caption})` : caption;
  if (calcs.length) label = calcs.map(c => TABLE_CALCS[c]).join(" of ") + " of " + label;
  return { label, calcs, agg, info, name };
}

/** @param {Pane} p @returns {string} the pane's mark ("" = Automatic) */
const markOf = p => String(p.markClass || "").toLowerCase().replace(/[\s_-]+/g, "").replace(/^automatic$/, "");

/**
 * @param {VisualModel} visualModel
 * @param {FormatModel | null} model
 * @returns {VisualSemantics}
 */
function describeVisual(visualModel, model) {
  const vm = visualModel.viewModel;
  const sheet = vm && vm.fmt ? vm.fmt.sheetModel : null;
  const type = visualModel.type;
  const evidence = [...(visualModel.evidence || [])];
  const source = visualModel.metadata.markSource;
  const confidence = sheet && source === "twb" ? "high" : sheet || source === "live" ? "medium" : "low";
  /** @type {VisualSemantics} */
  const out = { visual: "", family: type, confidence, evidence, encodings: {}, calculations: [], flags: {} };
  if (!sheet) {
    out.visual = { TABLE: "Text Table", KPI: "KPI", CUSTOM: "Viz Extension" }[type] || titleCase(type);
    if (!visualModel.formatModel) evidence.push("no workbook loaded: recognised from the live data only");
    return out;
  }
  const label = r => describeField(model, r).label;
  const isMeasure = r => tvIsMeasureRef(model, r);
  const panes = sheet.panes.length > 1 && sheet.panes.some(p => p.id) ? sheet.panes.filter(p => p.id) : sheet.panes;
  // an Automatic card draws the All card's mark, else Tableau's automatic mark for the shelves
  const allPane = sheet.panes.find(p => !p.id);
  const inherit = (allPane && markOf(allPane)) || tvAutomaticMark(visualShelfShape(sheet, model));
  const marks = [...new Set(panes.map(p => markOf(p) || inherit))];
  const enc = ch => {
    const refs = [];
    panes.forEach(p => p.encodings.filter(e => e.channel === ch).forEach(e => { if (!refs.some(r => tfSameField(r, e.field))) refs.push(e.field); }));
    return refs;
  };
  const color = enc("color"), size = enc("size"), shape = enc("shape"), detail = enc("lod"), path = enc("path");
  const text = [...enc("text"), ...enc("label")];
  panes.forEach(p => p.labelRuns.forEach(r => r.refs.forEach(x => { if (!text.some(t => tfSameField(t, x))) text.push(x); })));
  const shelf = name => sheet[name].filter(r => r.name !== "__tableau_internal_object_id__");
  out.encodings = {
    rows: shelf("rows").map(label), columns: shelf("cols").map(label), color: color.map(label), size: size.map(label),
    shape: shape.map(label), label: text.map(label), detail: detail.map(label), path: path.map(label),
    angle: enc("wedge-size").map(label), tooltip: enc("tooltip").map(label)
  };
  Object.keys(out.encodings).forEach(k => { if (!out.encodings[k].length) delete out.encodings[k]; });

  // calculated fields behind the view: Tableau computed them – the export writes their values
  const seen = new Set();
  [...sheet.rows, ...sheet.cols, ...panes.flatMap(p => p.encodings.map(e => e.field))].forEach(r => {
    const d = describeField(model, r);
    if (d.calcs.length && !seen.has(d.label)) { seen.add(d.label); out.calculations.push({ name: d.label, kind: "quick table calculation" }); }
    if (d.info && d.info.calcKind && !d.info.param && !seen.has(d.name)) {
      seen.add(d.name);
      out.calculations.push({ name: d.info.caption || d.name, kind: d.info.calcKind });
    }
  });

  const rowsM = sheet.rows.filter(isMeasure), colsM = sheet.cols.filter(isMeasure);
  const horizontal = colsM.length > 0 && rowsM.length === 0;
  const valueRefs = horizontal ? colsM : rowsM;
  const catShelf = horizontal ? "rows" : "cols";
  const dimsOf = s => sheet[s].filter(r => !isMeasure(r) && r.name !== "Measure Names" && !GEO_GENERATED.test(r.name) &&
                                        r.name !== "__tableau_internal_object_id__");
  const calcsOf = r => describeField(model, r).calcs;
  const colorDim = color.find(r => !isMeasure(r) && r.name !== "Measure Names");
  const colorMeasure = color.find(r => isMeasure(r));
  const constant = r => { const d = describeField(model, r); return !!d.info && d.info.constant !== undefined; };
  const negated = r => { const d = describeField(model, r); return !!d.info && /^\s*-/.test(d.info.formula || ""); };
  const spaces = (sheet.style && sheet.style.spaces) || [];
  const flags = out.flags;
  // Dual Axis: an axis folded onto the one before it (the TWB's fold); measures without it get panes of their own
  flags.dualAxis = spaces.some(s => s.fold);
  flags.separatePanes = !flags.dualAxis && (rowsM.length > 1 || colsM.length > 1);
  flags.percentOfTotal = valueRefs.some(r => calcsOf(r).includes("pcto"));
  flags.runningTotal = valueRefs.some(r => calcsOf(r).includes("cum") || calcsOf(r).includes("rsum"));
  flags.rank = valueRefs.some(r => calcsOf(r).includes("rank"));
  flags.reversedAxis = spaces.some(s => s.reverse);
  flags.constantAxis = valueRefs.some(constant);
  flags.smallMultiples = dimsOf(horizontal ? "cols" : "rows").length > 0 && valueRefs.length > 0;
  flags.detail = detail.some(r => !isMeasure(r));
  if (flags.dualAxis) evidence.push(`dual axis${spaces.some(s => s.fold && s.synchronized) ? " (synchronized)" : ""}`);
  if (flags.separatePanes) evidence.push("several measures, each in its own pane");
  if (flags.percentOfTotal) evidence.push("percent-of-total table calculation on the value axis");
  if (flags.runningTotal) evidence.push("running-total table calculation on the value axis");
  if (flags.rank) evidence.push("rank table calculation on the value axis");
  if (flags.reversedAxis) evidence.push("reversed axis");
  if (flags.smallMultiples) evidence.push("a dimension on the value shelf: one pane per member (small multiples)");
  if (colorDim) evidence.push(`colour by the dimension ${label(colorDim)}`);
  else if (colorMeasure) evidence.push(`colour by the measure ${label(colorMeasure)}`);
  if (size.length) evidence.push(`size by ${size.map(label).join(", ")}`);
  if (path.length) evidence.push(`path by ${path.map(label).join(", ")}`);
  if (flags.detail) evidence.push(`detail: ${detail.filter(r => !isMeasure(r)).map(label).join(", ")}`);

  const refLines = sheet.referenceLines || [];
  const bulletTarget = refLines.some(rl => rl.scope === "per-cell" && rl.field && rl.axis && !tfSameField(rl.field, rl.axis));
  const dateParts = [...sheet.rows, ...sheet.cols].map(r => String(r.deriv || "").toLowerCase());
  const has = m => marks.includes(m);
  const shelfDims = [...dimsOf("rows"), ...dimsOf("cols")];
  const bars = horizontal ? "Bar" : "Column";
  const name = () => {
    switch (type) {
      case VISUAL_TYPES.MAP: case VISUAL_TYPES.MAP_FILLED: {
        const geoDims = [...detail, ...shelfDims, ...color, ...text].filter(r => !isMeasure(r) && !GEO_GENERATED.test(r.name) &&
                                                                             !!(describeField(model, r).info || {}).geoRole);
        flags.geographic = true;
        if (has("density") || has("heatmap")) return "Density Map";
        if (!geoDims.length && has("pie")) return "Donut / Pie built on map layers";
        if (!geoDims.length) { evidence.push("no geographic field: map layers used as a canvas"); return "Map with Layers"; }
        if (has("pie")) return "Pies on a Map";
        if (panes.length > 1 && marks.length > 1) return "Map with Layers";
        if (has("line") && path.length) return "Flow Map";
        if (has("polygon")) return "Polygon Map";
        if (has("multipolygon") || type === VISUAL_TYPES.MAP_FILLED) return "Filled Map";
        return "Symbol Map";
      }
      case VISUAL_TYPES.PIE:
        if (visualModel.pie && (visualModel.pie.inner || (visualModel.pie.layers || 0) > 1)) return "Nested Donut";
        if ((visualModel.pie && visualModel.pie.hole) || (vm.fmt.pieLayers && (vm.fmt.pieLayers() || {}).hole)) return "Donut Chart";
        return shelfDims.length ? "Multiple Pies" : "Pie Chart";
      case VISUAL_TYPES.HEATMAP:
        if (dateParts.includes("wk") && (dateParts.includes("wd") || dateParts.includes("dy"))) return "Calendar Heat Map";
        return text.some(isMeasure) ? "Highlight Table" : "Heat Map";
      case VISUAL_TYPES.TABLE:
        if (vm.markTable) return "Table built from marks";
        if (has("shape") && dimsOf("rows").length && dimsOf("cols").length) return "KPI Indicator Matrix";
        if (has("circle") && size.length && shelfDims.some(r => r.type === "qk")) return "Circle Timeline";
        return dimsOf("rows").length && dimsOf("cols").length ? "Crosstab" : "Text Table";
      case VISUAL_TYPES.KPI: return "KPI (BAN)";
      case VISUAL_TYPES.HISTOGRAM: return "Histogram";
      case VISUAL_TYPES.WATERFALL: return "Waterfall Chart";
      case VISUAL_TYPES.BOXPLOT: return "Box Plot";
      case VISUAL_TYPES.GANTT: return "Gantt Chart";
      case VISUAL_TYPES.TREEMAP: return "Treemap";
      case VISUAL_TYPES.BUBBLE: return "Packed Bubbles";
      case VISUAL_TYPES.CUSTOM: return "Viz Extension";
      case VISUAL_TYPES.SCATTER:
        if (size.length) return "Scatter Plot with Size (bubble)";
        if (refLines.length >= 2 || (colorMeasure || colorDim) && refLines.length) return "Quadrant Chart";
        return "Scatter Plot";
      case VISUAL_TYPES.COMBO:
        if (valueRefs.some(r => calcsOf(r).includes("pcto") && calcsOf(r).includes("cum"))) return "Pareto Chart";
        return `Combination Chart (${marks.filter(Boolean).join(" + ")}, dual axis)`;
      case VISUAL_TYPES.AREA:
        if (flags.percentOfTotal && colorDim) return "100% Stacked Area";
        if (colorDim) return "Stacked Area";
        if (flags.smallMultiples) return "Small Multiple Areas";
        return "Area Chart";
      case VISUAL_TYPES.LINE:
        if (path.some(r => r.name === "Measure Names") && valueRefs.some(r => r.name === "Multiple Values")) {
          return has("circle") ? "Barbell (Dumbbell) Chart" : flags.constantAxis || hasConstantMeasure(model, sheet) ? "Rounded Bar Chart" : "Line Chart";
        }
        if (rowsM.length && colsM.length) return flags.runningTotal && flags.percentOfTotal ? "Pareto Curve (cumulative % against cumulative %)" : "Line (measure against measure)";
        if (flags.rank) return "Bump Chart";
        if (!has("line") && marks.some(m => /^(circle|shape|square)$/.test(m))) return "Dot Plot";
        if (valueRefs.length && valueRefs.every(r => r.name === valueRefs[0].name) && valueRefs.length === 2 && has("circle") &&
            sheet[catShelf].some(r => /^(yr|tyr)$/i.test(r.deriv || "") && r.type !== "qk")) return "Slope Chart";
        if (flags.runningTotal) return "Cumulative Line Chart";
        if (flags.smallMultiples) return "Small Multiple Lines (sparklines)";
        if (flags.dualAxis && valueRefs.length > 1 && !valueRefs.every(r => tfSameField(r, valueRefs[0]))) return "Dual-axis Line Chart";
        if (flags.separatePanes && valueRefs.length > 1) return "Line Charts in separate panes (one per measure)";
        if (colorDim || (flags.detail && has("line"))) return "Multi-line Chart";
        return "Line Chart";
      case VISUAL_TYPES.BAR: {
        if (path.some(r => r.name === "Measure Names")) return "Rounded Bar Chart";
        if (bulletTarget) return "Bullet Graph";
        if (valueRefs.some(negated) || (valueRefs.some(r => r.name === "Multiple Values") && sheet.fieldRefs.some(negated))) return "Butterfly (Diverging) Bar Chart";
        if (flags.percentOfTotal && valueRefs.some(constant)) return "Progress Bar";
        if (has("circle") && has("bar")) return "Lollipop Chart";
        if (flags.percentOfTotal && colorDim) return `100% Stacked ${bars}`;
        if (flags.smallMultiples) return `Small Multiple ${bars}s`;
        if (flags.dualAxis && valueRefs.length > 1 && marks.every(m => m === "bar")) return "Bar in Bar";
        if (flags.dualAxis && valueRefs.length > 1) return `Combination Chart (${marks.filter(Boolean).join(" + ")}, dual axis)`;
        if (flags.separatePanes && valueRefs.length > 1) return `${bars} Charts in separate panes (one per measure)`;
        const catDimRefs = dimsOf(catShelf);
        if (colorDim && !catDimRefs.some(r => tfSameField(r, colorDim)) && !sheet[horizontal ? "cols" : "rows"].some(r => tfSameField(r, colorDim))) return `Stacked ${bars}`;
        // side by side: the colour field is the innermost of several category fields (bars grouped by the outer one)
        if (colorDim && catDimRefs.length > 1 && tfSameField(catDimRefs[catDimRefs.length - 1], colorDim)) return `Side-by-side (grouped) ${bars}`;
        if (colorDim) return `${bars} Chart coloured by ${label(colorDim)}`;
        if (flags.detail) return `Stacked ${bars} (by detail)`;
        if (refLines.length) return `${bars} Chart with Reference Line`;
        if (dateParts.some(d => TV_DATE_DERIVS.has(d))) return `${bars} Chart over Time`;
        return horizontal ? "Horizontal Bar Chart" : "Column Chart";
      }
      default: return titleCase(type);
    }
  };
  out.visual = name();
  return out;
}

/** a measure on the sheet that is a constant (MIN(0), AVG(1)): it only places the marks */
function hasConstantMeasure(model, sheet) {
  return sheet.fieldRefs.some(r => { const d = describeField(model, r); return !!d.info && d.info.constant !== undefined && !d.info.param; });
}

function titleCase(t) { return String(t || "").toLowerCase().replace(/_/g, " ").replace(/\b\w/g, c => c.toUpperCase()); }


/* ═══ visual/strategy.js ══════════════════════════════════════════════════════════════════════════ */
/* How each visual was converted, in the terms of the conversion report:
 *   NATIVE          a genuine Excel equivalent (chart type or formatted cells)
 *   CONSTRUCTED     Excel has no such chart: built from Excel primitives that keep its meaning
 *   APPROXIMATE     an Excel chart that keeps the analytical meaning, with a visible difference (named)
 *   TABLE_FALLBACK  no Excel chart would read correctly: the visual's data as a table
 *   IMAGE_FALLBACK  a picture of the visual (Tableau-rendered), not editable
 *   UNSUPPORTED     not exported
 *   FAILED          exporting it went wrong
 * A correct table always beats a chart that would mislead: anything not provably meaningful falls back. */

const STRATEGY = Object.freeze({
  NATIVE: "NATIVE", CONSTRUCTED: "CONSTRUCTED", APPROXIMATE: "APPROXIMATE", TABLE_FALLBACK: "TABLE_FALLBACK",
  IMAGE_FALLBACK: "IMAGE_FALLBACK", UNSUPPORTED: "UNSUPPORTED", FAILED: "FAILED"
});

/** report order, strongest first */
const STRATEGY_ORDER = ["NATIVE", "CONSTRUCTED", "APPROXIMATE", "TABLE_FALLBACK", "IMAGE_FALLBACK", "UNSUPPORTED", "FAILED"];

const FIDELITY = { NATIVE: "High", CONSTRUCTED: "High", APPROXIMATE: "Medium", TABLE_FALLBACK: "Exact data",
                   IMAGE_FALLBACK: "Picture only", UNSUPPORTED: "None", FAILED: "None" };

/** an Excel chart in words ("Stacked column chart", "Combo chart: column + line, secondary axis") @param {ChartSpec} spec */
function chartOutputName(spec) {
  if (spec.conversion && spec.conversion.output) return spec.conversion.output;
  const dir = spec.barDir === "bar" ? "bar" : "column";
  const grouping = spec.percent ? "100% stacked " : spec.stacked ? "stacked " : "";
  const kinds = [...new Set(spec.series.filter(s => !s.refLine).map(s => s.type || spec.kind))];
  const extras = spec.series.some(s => s.refLine) || (spec.refLines && spec.refLines.length) ? " with reference line" : "";
  switch (spec.kind) {
    case "bar": return cap(`${grouping || "clustered "}${dir} chart`) + extras;
    case "line": return (spec.series.filter(s => !s.refLine).every(s => s.line === false) ? "Line chart (markers only)" : "Line chart") + extras;
    case "area": return cap(`${grouping}area chart`) + extras;
    case "combo": return (kinds.length === 1 && kinds[0] === "bar" ? `Bar-in-bar ${dir} chart (two axes${spec.secondarySync ? ", one scale" : ""})`
                          : `Combo chart: ${kinds.map(k => k === "bar" ? dir : k).join(" + ")}` +
                            (spec.series.some(s => s.secondary) ? `, secondary axis${spec.secondarySync ? " (synchronized)" : ""}` : "")) + extras;
    case "pie": return "Pie chart";
    case "doughnut": return "Doughnut chart";
    case "scatter": return spec.series.some(s => s.line) ? "Scatter chart with lines" : "Scatter chart" + extras;
    case "bubble": return "Bubble chart";
    case "treemap": return "Treemap (Excel 2016+)";
    default: return cap(`${spec.kind} chart`);
  }
}

/**
 * The conversion of one exported visual, once the export has drawn it (renderer, status and chart specs final).
 * @param {VisualModel} visualModel @param {VisualSemantics} [semantics]
 * @returns {Conversion}
 */
function conversionOf(visualModel, semantics) {
  const make = (strategy, excelOutput, reason, fidelity) => ({ strategy, excelOutput, reason, fidelity: fidelity || FIDELITY[strategy] });
  const renderer = visualModel.renderer;
  const why = String(visualModel.statusReason || "");
  if (visualModel.status === "warning" && /injection failed/i.test(why)) return make(STRATEGY.FAILED, "Nothing", why);
  if (visualModel.pie) {
    const pie = visualModel.pie;
    const donut = pie.hole || pie.inner;
    return make(STRATEGY.NATIVE, `${donut ? "Doughnut" : "Pie"} chart${pie.inner ? " (two rings)" : ""}`,
      "Direct Excel equivalent: Tableau's slices, colours and labels" + (pie.hole ? ", the hole with its text" : "") + "; tooltips as hover text");
  }
  if (renderer === "excel-chart") {
    const specs = visualModel.chartSpecs || [];
    const rank = s => STRATEGY_ORDER.indexOf((s.conversion && s.conversion.strategy) || defaultStrategy(s));
    const weakest = specs.slice().sort((a, b) => rank(b) - rank(a))[0];
    if (!weakest) return make(STRATEGY.FAILED, "Nothing", "no chart was produced");
    const strategy = (weakest.conversion && weakest.conversion.strategy) || defaultStrategy(weakest);
    const output = [...new Set(specs.map(chartOutputName))].join("; ") + (specs.length > 1 ? ` (${specs.length} panes)` : "");
    const note = specs.map(s => s.conversion && s.conversion.note).filter(Boolean)[0];
    return make(strategy, output, note || (strategy === STRATEGY.APPROXIMATE ? "Excel's closest chart; the layout differs" : "Direct Excel equivalent"),
                weakest.conversion && weakest.conversion.fidelity);
  }
  if (renderer === "tableau-image") return make(STRATEGY.IMAGE_FALLBACK, "Picture (Tableau-rendered) + its data on the Visual Data sheet",
                                                  why || "no editable Excel equivalent");
  if (renderer === "data-fallback") return make(STRATEGY.TABLE_FALLBACK, "Data table", fallbackReason(visualModel, semantics, why));
  if (renderer === "cell") {
    const vm = visualModel.viewModel;
    if (vm && vm.lossless === false) return make(STRATEGY.TABLE_FALLBACK, "Data table", vm.fallbackReason || "its marks encode values the cells would not show");
    if (vm && vm.matrix) return make(STRATEGY.CONSTRUCTED, "Matrix of cells with Tableau's colour per cell",
                                     "Excel has no heat-map chart: a matrix with the same rows, columns and colours");
    if (vm && vm.markTable) return make(STRATEGY.CONSTRUCTED, "Cells drawn like Tableau's marks (fills, arrows, data bars)",
                                        "a table built from marks: each column styled like its marks");
    if (visualModel.type === VISUAL_TYPES.KPI) return make(STRATEGY.NATIVE, "KPI cells", "Direct Excel equivalent: the values in Tableau's format");
    if (visualModel.type === VISUAL_TYPES.HEATMAP) return make(STRATEGY.CONSTRUCTED, "Table with Tableau's colour per row", "Excel has no heat-map chart");
    return make(STRATEGY.NATIVE, "Formatted cell table", "Direct Excel equivalent: Tableau's table as cells");
  }
  return make(STRATEGY.UNSUPPORTED, "Nothing", why || "no renderer");
}

/** @param {ChartSpec} spec */
function defaultStrategy(spec) {
  if (spec.kind === "bubble" && spec.packed) return STRATEGY.APPROXIMATE;
  return STRATEGY.NATIVE;
}

/** why a visual became a data table, in words a reader of the report understands */
function fallbackReason(visualModel, semantics, why) {
  const t = visualModel.type;
  const visual = semantics ? semantics.visual : "";
  if (t === VISUAL_TYPES.MAP || t === VISUAL_TYPES.MAP_FILLED) {
    const kind = {
      "Density Map": "density marks have no Excel equivalent",
      "Polygon Map": "custom polygons (hex bins, shapes) cannot be drawn by an Excel chart",
      "Flow Map": "paths between locations cannot be drawn on an Excel map",
      "Map with Layers": "several map layers on one map have no Excel equivalent",
      "Pies on a Map": "pies placed on a map have no Excel equivalent",
      "Donut / Pie built on map layers": "a donut built from map layers (no geography) – Excel's charts cannot stack the layers",
      "Filled Map": "Excel's filled map cannot reproduce Tableau's geography and colour scale reliably",
      "Symbol Map": "Excel has no symbol map: marks sized and coloured at locations"
    }[visual] || "Excel cannot reproduce this map";
    return `${kind}; every mark exported with its location fields and values`;
  }
  if (t === VISUAL_TYPES.CUSTOM) return "a viz extension draws this sheet; Excel has no equivalent";
  const m = why.match(/native chart not possible \(([^)]*)\)/i);
  if (m) return `no faithful Excel chart: ${m[1]}`;
  return why.replace(/ – exported as data$/, "") || "no meaningful Excel equivalent";
}

function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }


/* ═══ export/data-sheet.js ════════════════════════════════════════════════════════════════════════ */
/* The data behind visuals exported as pictures (IMAGE_FALLBACK): each one's marks as a plain table on a "Visual Data"
 * sheet, so the numbers travel with the picture (which links to them). */

const DATA_SHEET = "Visual Data";

/**
 * Appends one visual's data (every column but the latitude / longitude Tableau generates) under a title.
 * @param {import("exceljs").Workbook} workbook @param {string} title @param {ViewModel} vm
 * @returns {string} the link to the block ("#'Visual Data'!A5")
 */
function appendVisualData(workbook, title, vm) {
  const ws = workbook.getWorksheet(DATA_SHEET) || workbook.addWorksheet(DATA_SHEET);
  const font = { name: "Arial", size: 9 };
  const start = ws.rowCount ? ws.rowCount + 2 : 1;
  const cols = vm.cols.map((c, i) => i).filter(i => !/^(latitude|longitude) \(generated\)$/i.test(vm.cols[i].name));
  ws.getCell(start, 1).value = title;
  ws.getCell(start, 1).font = { ...font, size: 11, bold: true };
  cols.forEach((ci, k) => {
    const cell = ws.getCell(start + 1, k + 1);
    cell.value = vm.cols[ci].label || vm.cols[ci].name;
    cell.font = { ...font, bold: true };
    cell.border = { bottom: { style: "thin", color: { argb: "FFD4D4D4" } } };
    ws.getColumn(k + 1).width = Math.max(ws.getColumn(k + 1).width || 0, Math.min(40, String(cell.value).length + 4));
  });
  vm.rows.forEach((row, r) => cols.forEach((ci, k) => {
    const dv = row[ci], cell = ws.getCell(start + 2 + r, k + 1);
    const v = tfIsNull(dv) ? null : dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
    cell.value = typeof v === "number" ? v : tfIsNull(dv) ? null : tfDvText(dv);
    if (typeof v === "number") { const nf = inferExcelNumFmt(dv.formattedValue, v); if (nf) cell.numFmt = nf; }   // as Tableau shows it
    cell.font = font;
  }));
  return `#'${DATA_SHEET}'!A${start}`;
}


/* ═══ export/report.js ════════════════════════════════════════════════════════════════════════════ */
/* The conversion report: a sheet after the dashboard that lists every visual – what Tableau shows, how Excel shows
 * it (NATIVE … TABLE_FALLBACK), the fidelity and why – then the dashboard's other objects, the filters and
 * parameters the data reflects, and what was left out. Nothing is converted silently. */

const REPORT_SHEET = "Conversion Report";

const STRATEGY_FILL = { NATIVE: "FFE2F0D9", CONSTRUCTED: "FFDDEBF7", APPROXIMATE: "FFFFF2CC", TABLE_FALLBACK: "FFEDEDED",
                        IMAGE_FALLBACK: "FFE4DFEC", UNSUPPORTED: "FFF8CBAD", FAILED: "FFF4B183" };

/** "15 NATIVE · 3 CONSTRUCTED · 2 TABLE_FALLBACK" @param {{ strategy: string }[]} visuals */
function strategyCounts(visuals) {
  return STRATEGY_ORDER.map(s => [s, visuals.filter(v => v.strategy === s).length]).filter(([, n]) => n);
}

/**
 * @param {import("exceljs").Workbook} workbook
 * @param {{ dashboard: string, workbookFile: string | null, visuals: ReportVisual[],
 *           objects: { name: string, kind: string, output: string }[], skipped: { name: string, reason: string }[],
 *           filters: { name: string, values: string[] }[], parameters: { name: string, value: string }[] }} report
 */
function writeConversionReport(workbook, report) {
  const names = workbook.worksheets.map(w => w.name);
  const ws = workbook.addWorksheet(names.includes(REPORT_SHEET) ? REPORT_SHEET + " (export)" : REPORT_SHEET);
  ws.views = [{ state: "frozen", ySplit: 5, showGridLines: false }];
  const font = { name: "Arial", size: 9 };
  const bold = { ...font, bold: true };
  let r = 1;
  const put = (row, col, value, f = font, extra = {}) => {
    const cell = ws.getCell(row, col);
    cell.value = value;
    cell.font = f;
    cell.alignment = { vertical: "top", wrapText: true, ...(extra.alignment || {}) };
    if (extra.fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: extra.fill } };
    if (extra.border) cell.border = extra.border;
    return cell;
  };
  put(r, 1, "Tableau → Excel conversion report", { ...font, size: 14, bold: true }, { alignment: { wrapText: false } });
  ws.mergeCells(r, 1, r, 6);
  r++;
  put(r, 1, `Dashboard: ${report.dashboard}` + (report.workbookFile ? `   ·   Workbook: ${report.workbookFile}` : "   ·   no workbook loaded"),
      font, { alignment: { wrapText: false } });
  ws.mergeCells(r, 1, r, 6);
  r++;
  const counts = strategyCounts(report.visuals);
  put(r, 1, `${report.visuals.length} visual${report.visuals.length === 1 ? "" : "s"}: ` +
      (counts.map(([s, n]) => `${n} ${s}`).join("   ·   ") || "none"), bold, { alignment: { wrapText: false } });
  ws.mergeCells(r, 1, r, 6);
  r += 2;

  const head = ["#", "Worksheet", "Tableau visual", "Strategy", "Excel output", "Fidelity", "Reason", "Detected from", "Encodings", "Calculations"];
  const widths = [4, 24, 24, 16, 30, 11, 52, 52, 48, 34];
  const line = { style: "thin", color: { argb: "FFD4D4D4" } };
  head.forEach((h, i) => put(r, i + 1, h, bold, { fill: "FFF2F2F2", border: { bottom: line } }));
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const headerRow = r;
  report.visuals.forEach((v, k) => {
    r++;
    const enc = Object.entries(v.semantics.encodings || {}).map(([ch, list]) => `${ch}: ${list.join(", ")}`).join("\n");
    const calcs = (v.semantics.calculations || []).map(c => `${c.name} (${c.kind})`).join("\n");
    const cells = [k + 1, v.worksheet, v.semantics.visual, v.strategy, v.excelOutput, v.fidelity, v.reason,
                   `${v.semantics.evidence.join("; ")} [confidence: ${v.semantics.confidence}]`, enc, calcs];
    cells.forEach((value, i) => put(r, i + 1, value, i === 3 ? bold : font,
      { fill: i === 3 ? STRATEGY_FILL[v.strategy] : undefined, border: { bottom: line } }));
  });
  ws.autoFilter = { from: { row: headerRow, column: 1 }, to: { row: r, column: head.length } };

  const section = title => { r += 2; put(r, 1, title, { ...bold, size: 11 }, { alignment: { wrapText: false } }); ws.mergeCells(r, 1, r, 6); };
  if (report.objects.length) {
    section("Other dashboard objects");
    report.objects.forEach(o => { r++; put(r, 2, o.name); put(r, 3, o.kind); put(r, 5, o.output); });
  }
  if (report.filters.length || report.parameters.length) {
    section("Filters and parameters (the exported data reflects their state at export time)");
    report.filters.forEach(f => { r++; put(r, 2, f.name); put(r, 3, "Filter"); put(r, 5, f.values.join(", ")); });
    report.parameters.forEach(p => { r++; put(r, 2, p.name); put(r, 3, "Parameter"); put(r, 5, p.value); });
  }
  if (report.skipped.length) {
    section("Not exported");
    report.skipped.forEach(s => { r++; put(r, 2, s.name); put(r, 5, s.reason); ws.mergeCells(r, 5, r, 7); });
  }
  section("Notes");
  [
    "Values are Tableau's own results for the view: calculated fields, LOD expressions and table calculations are exported as values, not translated into Excel formulas.",
    "Filters, parameters and sets apply as they were at export time; Excel has no live link back to Tableau.",
    "Tooltips: pies, donuts and packed bubbles show Tableau's tooltip text on hover; elsewhere the tooltip fields are in the exported data only.",
    "Actions, highlighting and drill-down are interactive Tableau features and are not reproduced.",
    "TABLE_FALLBACK means no Excel chart would show the visual correctly: its data is exported as a table instead (the cell carries a note)."
  ].forEach(t => { r++; put(r, 2, t, font, { alignment: { wrapText: true } }); ws.mergeCells(r, 2, r, 7); ws.getRow(r).height = 24; });
  return ws;
}

/**
 * Several dashboards in one file: one report sheet with a section per dashboard (its visuals, other objects,
 * filters / parameters, what was left out), then the notes once. One dashboard: exactly writeConversionReport.
 * @param {import("exceljs").Workbook} workbook
 * @param {Parameters<typeof writeConversionReport>[1][]} reports
 */
function writeConversionReports(workbook, reports) {
  if (reports.length === 1) return writeConversionReport(workbook, reports[0]);
  const names = workbook.worksheets.map(w => w.name);
  const ws = workbook.addWorksheet(names.includes(REPORT_SHEET) ? REPORT_SHEET + " (export)" : REPORT_SHEET);
  ws.views = [{ showGridLines: false }];
  const font = { name: "Arial", size: 9 };
  const bold = { ...font, bold: true };
  const line = { style: "thin", color: { argb: "FFD4D4D4" } };
  let r = 1;
  const put = (row, col, value, f = font, extra = {}) => {
    const cell = ws.getCell(row, col);
    cell.value = value;
    cell.font = f;
    cell.alignment = { vertical: "top", wrapText: true, ...(extra.alignment || {}) };
    if (extra.fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: extra.fill } };
    if (extra.border) cell.border = extra.border;
    return cell;
  };
  const banner = (text, f) => { put(r, 1, text, f, { alignment: { wrapText: false } }); ws.mergeCells(r, 1, r, 6); };
  const all = reports.flatMap(rep => rep.visuals);
  banner("Tableau → Excel conversion report", { ...font, size: 14, bold: true });
  r++;
  banner(`${reports.length} dashboards` + (reports[0].workbookFile ? `   ·   Workbook: ${reports[0].workbookFile}` : "   ·   no workbook loaded"), font);
  r++;
  banner(`${all.length} visual${all.length === 1 ? "" : "s"}: ` + (strategyCounts(all).map(([s, n]) => `${n} ${s}`).join("   ·   ") || "none"), bold);

  const head = ["#", "Worksheet", "Tableau visual", "Strategy", "Excel output", "Fidelity", "Reason", "Detected from", "Encodings", "Calculations"];
  const widths = [4, 24, 24, 16, 30, 11, 52, 52, 48, 34];
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const section = title => { r += 2; banner(title, { ...bold, size: 11 }); };

  reports.forEach(report => {
    r += 2;
    const counts = strategyCounts(report.visuals);
    banner(`Dashboard: ${report.dashboard}`, { ...bold, size: 12 });
    r++;
    banner(`${report.visuals.length} visual${report.visuals.length === 1 ? "" : "s"}: ` +
      (counts.map(([s, n]) => `${n} ${s}`).join("   ·   ") || "none"), font);
    r++;
    head.forEach((h, i) => put(r, i + 1, h, bold, { fill: "FFF2F2F2", border: { bottom: line } }));
    report.visuals.forEach((v, k) => {
      r++;
      const enc = Object.entries(v.semantics.encodings || {}).map(([ch, list]) => `${ch}: ${list.join(", ")}`).join("\n");
      const calcs = (v.semantics.calculations || []).map(c => `${c.name} (${c.kind})`).join("\n");
      const cells = [k + 1, v.worksheet, v.semantics.visual, v.strategy, v.excelOutput, v.fidelity, v.reason,
                     `${v.semantics.evidence.join("; ")} [confidence: ${v.semantics.confidence}]`, enc, calcs];
      cells.forEach((value, i) => put(r, i + 1, value, i === 3 ? bold : font,
        { fill: i === 3 ? STRATEGY_FILL[v.strategy] : undefined, border: { bottom: line } }));
    });
    if (report.objects.length) {
      section("Other dashboard objects");
      report.objects.forEach(o => { r++; put(r, 2, o.name); put(r, 3, o.kind); put(r, 5, o.output); });
    }
    if (report.filters.length || report.parameters.length) {
      section("Filters and parameters (the exported data reflects their state at export time)");
      report.filters.forEach(f => { r++; put(r, 2, f.name); put(r, 3, "Filter"); put(r, 5, f.values.join(", ")); });
      report.parameters.forEach(p => { r++; put(r, 2, p.name); put(r, 3, "Parameter"); put(r, 5, p.value); });
    }
    if (report.skipped.length) {
      section("Not exported");
      report.skipped.forEach(s => { r++; put(r, 2, s.name); put(r, 5, s.reason); ws.mergeCells(r, 5, r, 7); });
    }
  });
  section("Notes");
  [
    "Values are Tableau's own results for the view: calculated fields, LOD expressions and table calculations are exported as values, not translated into Excel formulas.",
    "Filters, parameters and sets apply as they were at export time; Excel has no live link back to Tableau.",
    "Tooltips: pies, donuts and packed bubbles show Tableau's tooltip text on hover; elsewhere the tooltip fields are in the exported data only.",
    "Actions, highlighting and drill-down are interactive Tableau features and are not reproduced.",
    "TABLE_FALLBACK means no Excel chart would show the visual correctly: its data is exported as a table instead (the cell carries a note)."
  ].forEach(t => { r++; put(r, 2, t, font, { alignment: { wrapText: true } }); ws.mergeCells(r, 2, r, 7); ws.getRow(r).height = 24; });
  return ws;
}


/* ═══ charts/writer/chart-data.js ═════════════════════════════════════════════════════════════════ */
/* Chart source data → worksheet cells (charts stay editable and linked). */

/**
 * @param {import("exceljs").Worksheet} ws the (hidden) chart data sheet
 * @param {ChartSpec} spec
 * @param {number} startRow 0-based first row to write
 * @returns {ChartRefs}
 */
function writeChartData(ws, spec, startRow) {
  const sheet = ws.name;
  const put = (r, c, v, fmt) => {
    const cell = ws.getCell(r + 1, c + 1);
    cell.value = v;
    if (fmt) cell.numFmt = fmt;
  };
  ws.getCell(startRow + 1, 1).font = { bold: true };

  if (spec.kind === "scatter" || spec.kind === "bubble") {
    const per = spec.kind === "bubble" ? 3 : 2;                 // x, y (, bubble size) columns per series
    let maxLen = 0;
    /** @type {ChartRefs["series"]} */
    const series = spec.series.map((s, k) => {
      const cx = per * k, cy = cx + 1, n = s.x.length;
      put(startRow, cx, (spec.xTitle || "X") + (spec.series.length > 1 ? " – " + s.name : ""));
      put(startRow, cy, s.name);
      for (let i = 0; i < n; i++) {
        if (num(s.x[i]) !== null) put(startRow + 1 + i, cx, s.x[i], spec.xNumFmt);
        if (num(s.y[i]) !== null) put(startRow + 1 + i, cy, s.y[i], spec.numFmt);
      }
      maxLen = Math.max(maxLen, n);
      /** @type {ChartRefs["series"][number]} */
      const ref = { tx: cellRef(sheet, cy, startRow),
                    x: cellRef(sheet, cx, startRow + 1, cx, startRow + Math.max(1, n)),
                    y: cellRef(sheet, cy, startRow + 1, cy, startRow + Math.max(1, n)) };
      if (per === 3) {
        const cs = cx + 2;
        put(startRow, cs, (spec.sizeTitle || "Size") + (spec.series.length > 1 ? " – " + s.name : ""));
        s.size.forEach((v, i) => { if (num(v) !== null) put(startRow + 1 + i, cs, v, spec.sizeNumFmt); });
        ref.size = cellRef(sheet, cs, startRow + 1, cs, startRow + Math.max(1, n));
      }
      return ref;
    });
    const lblCol = per * spec.series.length;
    spec.series.forEach((s, k) => {
      if (!s.labelTexts) return;
      const c = lblCol + k;
      put(startRow, c, s.name + " – label");
      s.labelTexts.forEach((v, i) => put(startRow + 1 + i, c, v));
      series[k].lbl = cellRef(sheet, c, startRow + 1, c, startRow + Math.max(1, s.labelTexts.length));
    });
    return { series, nextRow: startRow + maxLen + 3 };
  }

  const levels = spec.categories.levels;
  const L = levels.length;
  const N = L ? levels[0].length : 0;
  // treemap: every level written on every row – Excel reads each row as one full hierarchy path
  const starts = spec.kind === "treemap" ? levels.map(lv => lv.map(() => true)) : levelStarts(levels);
  spec.categories.names.forEach((name, l) => put(startRow, l, name || ""));
  spec.series.forEach((s, k) => put(startRow, L + k, s.name));
  for (let i = 0; i < N; i++) {
    for (let l = 0; l < L; l++) if (l === L - 1 || starts[l][i]) put(startRow + 1 + i, l, levels[l][i]);
    spec.series.forEach((s, k) => {
      const v = num(s.values[i]);
      if (v !== null) put(startRow + 1 + i, L + k, v, s.secondary ? spec.secondaryNumFmt : spec.numFmt);
    });
  }
  // label texts taken from cells (Excel's "Value From Cells"), after the series columns
  let lblCol = L + spec.series.length;
  const lbl = spec.series.map(s => {
    if (!s.labelTexts) return null;
    const c = lblCol++;
    put(startRow, c, s.name + " – label");
    s.labelTexts.forEach((t, i) => put(startRow + 1 + i, c, t));
    return cellRef(sheet, c, startRow + 1, c, startRow + Math.max(1, N));
  });
  return {
    cat: cellRef(sheet, 0, startRow + 1, L - 1, startRow + N),
    series: spec.series.map((s, k) => ({
      tx: cellRef(sheet, L + k, startRow),
      val: cellRef(sheet, L + k, startRow + 1, L + k, startRow + N),
      ...(lbl[k] ? { lbl: lbl[k] } : {})
    })),
    nextRow: startRow + N + 3
  };
}


/* ═══ charts/writer/drawingml.js ══════════════════════════════════════════════════════════════════ */
/* DrawingML chart XML (bar / line / area / pie / scatter / bubble / combo). */

function solid(color, alpha) {
  if (!color) return "<a:noFill/>";
  return `<a:solidFill><a:srgbClr val="${hex(color)}">${alpha ? `<a:alpha val="${alpha}"/>` : ""}</a:srgbClr></a:solidFill>`;
}

function line(color, w) { return color ? `<a:ln w="${w || 9525}">${solid(color)}</a:ln>` : `<a:ln><a:noFill/></a:ln>`; }

function runProps(font, o, tag) {
  const sz = Math.round((o.size || font.size || 9) * 100);
  return `<a:${tag} lang="en-US" sz="${sz}" b="${o.bold ? 1 : 0}">${solid(o.color || font.color)}` +
    `<a:latin typeface="${esc(font.name)}"/><a:cs typeface="${esc(font.name)}"/></a:${tag}>`;
}

function txPr(font, o = {}) {
  const def = runProps(font, o, "defRPr").replace(' lang="en-US"', "");
  // rot 0 is written too: it keeps Excel from turning labels it finds crowded
  return `<c:txPr><a:bodyPr${o.rot !== undefined ? ` rot="${o.rot}" vert="horz"` : ""}${o.noWrap ? ' wrap="none"' : ""}/><a:lstStyle/>` +
    `<a:p><a:pPr>${def}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

function title(text, font, vertical) {
  if (!text) return "";
  const def = runProps(font, {}, "defRPr").replace(' lang="en-US"', "");
  return `<c:title><c:tx><c:rich><a:bodyPr${vertical ? ' rot="-5400000" vert="horz"' : ""}/><a:lstStyle/>` +
    `<a:p><a:pPr>${def}</a:pPr><a:r>${runProps(font, {}, "rPr")}<a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx>` +
    `<c:overlay val="0"/></c:title>`;
}

function strCache(values) {
  return `<c:ptCount val="${values.length}"/>` +
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("");
}

function numCache(values) {
  return `<c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>` +
    values.map((v, i) => num(v) === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
}

function serTx(ref, name) {
  return `<c:tx><c:strRef><c:f>${esc(ref)}</c:f><c:strCache>${strCache([name])}</c:strCache></c:strRef></c:tx>`;
}

/** @param {string} ref @param {any[][]} levels @param {string[] | null} [shown] labels shown instead (Tableau's
 * truncated headers): written into the chart, as Excel re-reads linked category cells when it opens the file */
function catXml(ref, levels, shown = null) {
  if (shown && levels.length <= 1) return `<c:cat><c:strLit>${strCache(shown)}</c:strLit></c:cat>`;
  if (levels.length <= 1) {
    return `<c:cat><c:strRef><c:f>${esc(ref)}</c:f><c:strCache>${strCache(levels[0] || [])}</c:strCache></c:strRef></c:cat>`;
  }
  const starts = levelStarts(levels);
  const n = levels[0].length;
  // innermost level first, as Excel writes it
  const lvls = levels.map((lv, l) => `<c:lvl>${lv.map((v, i) =>
      (l === levels.length - 1 || starts[l][i]) ? `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>` : "").join("")}</c:lvl>`).reverse();
  return `<c:cat><c:multiLvlStrRef><c:f>${esc(ref)}</c:f><c:multiLvlStrCache><c:ptCount val="${n}"/>${lvls.join("")}` +
    `</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`;
}

function valXml(tag, ref, values) {
  return `<c:${tag}><c:numRef><c:f>${esc(ref)}</c:f><c:numCache>${numCache(values)}</c:numCache></c:numRef></c:${tag}>`;
}

/** mark labels: the worksheet's label font (else the chart font); labels set from the workbook never wrap
 * @param {ChartSpec} spec */
function labelTxPr(spec) {
  const lf = spec.labelFont;
  if (!lf) return txPr(spec.font, spec.labelPos ? { noWrap: true } : {});
  return txPr({ name: lf.name || spec.font.name, size: lf.size || spec.font.size, color: lf.color || spec.font.color },
              { bold: lf.bold, noWrap: true });
}

/** @param {ChartSpec} spec @param {ChartSeries} s @param {string | null} pos @param {number[]} [hidden] points without a label */
function dLbls(spec, s, pos, hidden) {
  if (!s.labels) return "";
  const drop = (hidden || []).map(i => `<c:dLbl><c:idx val="${i}"/><c:delete val="1"/></c:dLbl>`).join("");
  if (s.labelTexts) {
    // a doughnut takes no label position (Excel refuses the file)
    return `<c:dLbls>${drop}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${labelTxPr(spec)}` +
      `${spec.kind === "doughnut" ? "" : `<c:dLblPos val="${pos || "r"}"/>`}<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>` +
      `<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>` +
      `<c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="${C15}">` +
      `<c15:showDataLabelsRange val="1"/></c:ext></c:extLst></c:dLbls>`;
  }
  const fmt = s.labelNumFmt || (s.secondary ? spec.secondaryNumFmt : spec.numFmt);
  const showVal = s.labelParts ? (s.labelParts.value ? 1 : 0) : 1;
  const showCat = s.labelParts && s.labelParts.category ? 1 : 0;
  return `<c:dLbls>${fmt ? `<c:numFmt formatCode="${esc(fmt)}" sourceLinked="0"/>` : ""}` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${labelTxPr(spec)}` +
    `${pos ? `<c:dLblPos val="${pos}"/>` : ""}<c:showLegendKey val="0"/><c:showVal val="${showVal}"/>` +
    `<c:showCatName val="${showCat}"/><c:showSerName val="0"/><c:showPercent val="${s.labelParts && s.labelParts.percent ? 1 : 0}"/><c:showBubbleSize val="0"/>` +
    `${spec.kind === "pie" || spec.kind === "doughnut" ? '<c:showLeaderLines val="1"/>' : ""}</c:dLbls>`;
}

function markerXml(symbol, color, size) {
  if (!symbol || symbol === "none") return `<c:marker><c:symbol val="none"/></c:marker>`;
  return `<c:marker><c:symbol val="${symbol}"/><c:size val="${size || 6}"/>` +
    `<c:spPr>${solid(color, 85000)}${line(color, 9525)}</c:spPr></c:marker>`;
}

/* a reference line's stroke: colour with its opacity, width (px), dashes; none when switched off */
function refLineLn(rl) {
  if (rl.hidden) return `<a:ln><a:noFill/></a:ln>`;
  return `<a:ln w="${Math.round(rl.width * 9525)}"><a:solidFill><a:srgbClr val="${hex(rl.color)}">` +
    `${rl.alpha < 1 ? `<a:alpha val="${Math.round(rl.alpha * 100000)}"/>` : ""}</a:srgbClr></a:solidFill>` +
    `${rl.dash ? '<a:prstDash val="dash"/>' : ""}</a:ln>`;
}

/* the line's label on one of its points (idx): an Excel number format on the line's value ("Average",
 * "Avg. $1.2M"); valueFlag picks the value Excel formats (showVal = y, showCatName = x of an XY line) */
function refLineLabel(spec, rl, pos, valueFlag = "showVal", idx = 0) {
  if (!rl.labelFmt) return "";
  const font = { name: spec.font.name, size: spec.font.size, color: rl.font.color || "555555" };
  const flags = on => ["showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"]
    .map(f => `<c:${f} val="${on && f === valueFlag ? 1 : 0}"/>`).join("");
  return `<c:dLbls><c:dLbl><c:idx val="${idx}"/><c:numFmt formatCode="${esc(rl.labelFmt)}" sourceLinked="0"/>` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(font, { bold: rl.font.bold, noWrap: true })}` +
    `<c:dLblPos val="${pos}"/>${flags(true)}</c:dLbl>${flags(false)}</c:dLbls>`;
}

/* a vertical reference line across horizontal bars: XY points (value, 0) → (value, 1) on the hidden x2 / y2 axes */
function refLineXY(spec, rl, k) {
  const lit = vals => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
    vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("") + `</c:numLit>`;
  return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/><c:tx><c:v>${esc(rl.name)}</c:v></c:tx>` +
    `<c:spPr>${refLineLn(rl)}</c:spPr><c:marker><c:symbol val="none"/></c:marker>${refLineLabel(spec, rl, "r", "showCatName")}` +
    `<c:xVal>${lit([rl.value, rl.value])}</c:xVal><c:yVal>${lit([0, 1])}</c:yVal><c:smooth val="0"/></c:ser>`;
}

/* bullet targets across horizontal bars: one vertical tick per category at its target, on the hidden x2 / y2 axes
 * (y2 runs 0 – 1 bottom to top; the first category is the top band). Blank points between ticks break the line. */
function targetTicksXY(spec, k) {
  const t = spec.targets, n = t.values.length;
  const xs = [], ys = [];
  t.values.forEach((v, i) => {
    const mid = 1 - (i + 0.5) / n, half = 0.32 / n;
    xs.push(v, v, null); ys.push(mid - half, mid + half, null);
  });
  const lit = vals => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
    vals.map((v, i) => num(v) === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("") + `</c:numLit>`;
  return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/><c:tx><c:v>${esc(t.name || "Target")}</c:v></c:tx>` +
    `<c:spPr><a:ln w="28575" cap="flat">${solid(t.color || "333333")}</a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>` +
    `<c:xVal>${lit(xs)}</c:xVal><c:yVal>${lit(ys)}</c:yVal><c:smooth val="0"/></c:ser>`;
}

/* marks over horizontal bars' category bands (dots of a dot plot or lollipop, a dumbbell's joining lines): XY points
 * (value, band) on the hidden x2 / y2 axes – a band's centre is 1 − (i + ½) / n; a null band breaks the line */
function overlayXY(spec, o, k) {
  const n = spec.categories && spec.categories.levels.length ? spec.categories.levels[0].length : 1;
  const lit = vals => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
    vals.map((v, i) => num(v) === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("") + `</c:numLit>`;
  const ys = o.cats.map(c => c === null || c === undefined ? null : +(1 - (c + 0.5) / n).toFixed(6));
  const xs = o.values.map((v, i) => ys[i] === null ? null : v);
  const ln = o.line ? `<a:ln w="${Math.round((o.lineWidth || 1.5) * 9525)}" cap="rnd">${solid(o.color)}</a:ln>` : `<a:ln w="19050"><a:noFill/></a:ln>`;
  return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/><c:tx><c:v>${esc(o.name || "Marks")}</c:v></c:tx>` +
    `<c:spPr>${ln}</c:spPr>${markerXml(o.marker || "none", o.color, o.markerSize || 7)}` +
    `<c:xVal>${lit(xs)}</c:xVal><c:yVal>${lit(ys)}</c:yVal><c:smooth val="0"/></c:ser>`;
}

/* a hidden value axis pinned to [min, max] (the XY overlay's x2 / y2) */
function hiddenValAx(id, cross, pos, min, max) {
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/><c:max val="${max}"/><c:min val="${min}"/></c:scaling>` +
    `<c:delete val="1"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="none"/><c:crossAx val="${cross}"/><c:crosses val="max"/>` +
    `<c:crossBetween val="midCat"/></c:valAx>`;
}

/** Tableau-like axis ticks: a "nice" step (1 / 2 / 2.5 / 5 × 10ⁿ) giving about `ticks` intervals
 * @param {number} range @param {number} ticks @returns {number | null} */
function niceUnit(range, ticks) {
  if (!(range > 0) || !(ticks > 0)) return null;
  const raw = range / ticks, pow = Math.pow(10, Math.floor(Math.log10(raw))), n = raw / pow;
  return +((n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * pow).toPrecision(12);
}

/**
 * The primary value axis as Tableau draws it: ticks at the workbook's spacing, else a "nice" step about every
 * 55 px (110 px across); the range starts at zero (without "Include zero": at a tick below the data) and ends
 * just past the data, so no empty tick sits above the marks. Stacked bars / areas count as their totals.
 * Synchronized axes always get an explicit range: the secondary axis is pinned to it.
 * @param {ChartSpec} spec @param {{ w: number, h: number } | null} axisPlot approximate plot area in px
 * @returns {{ fixed: { min?: number, max?: number }, unit: number | null }}
 */
function valueRange(spec, axisPlot) {
  if (spec.percent) return { fixed: { min: 0, max: 1 }, unit: spec.valueMajorUnit || 0.2 };    // 100 % stacked: 0 – 100 %
  const fixed = { min: spec.valueMin, max: spec.valueMax };
  const vals = [], stacks = new Map();
  spec.series.filter(s => (!s.secondary || spec.secondarySync) && s.values).forEach(s => {
    const type = s.type || spec.kind;
    if (spec.stacked && !s.refLine && (type === "bar" || type === "area")) {
      const g = stacks.get(type) || [];
      stacks.set(type, g);
      s.values.forEach((v, i) => {
        const x = num(v);
        if (x === null) return;
        const t = g[i] || (g[i] = { pos: 0, neg: 0 });
        if (x >= 0) t.pos += x; else t.neg += x;
      });
    } else s.values.forEach(v => { const x = num(v); if (x !== null) vals.push(x); });
  });
  stacks.forEach(g => g.forEach(t => { if (t) vals.push(t.pos, t.neg); }));
  (spec.refLines || []).forEach(r => vals.push(r.value));
  (spec.targets ? spec.targets.values : []).forEach(v => { if (num(v) !== null) vals.push(v); });
  (spec.overlay || []).forEach(o => o.values.forEach(v => { if (num(v) !== null) vals.push(v); }));
  const ticks = axisPlot ? Math.max(2, Math.round(spec.barDir === "bar" ? axisPlot.w / 110 : axisPlot.h / 55)) : null;
  const aligned = (v, u) => Math.abs(v / u - Math.round(v / u)) < 1e-9;
  const round = v => +v.toPrecision(12);
  const pinned = fixed.min !== undefined && fixed.max !== undefined;
  // Excel labels ticks from the axis minimum: a fixed minimum off the tick step starts at the step below
  // it, so the labels are round numbers as on Tableau's axis (a hidden axis keeps the exact range)
  const dates = /[dmy]/i.test(spec.numFmt || "") && !/[0#]/.test(spec.numFmt || "");   // Gantt: Excel's own date steps
  const settle = (range, unit) => {
    if (!unit || range.min === undefined || aligned(range.min, unit)) return unit;
    if (spec.valueAxisHidden || dates) return null;
    range.min = round(Math.floor(range.min / unit) * unit);
    return unit;
  };
  // the workbook's spacing, unless it would crowd the axis with ticks (data far beyond what it was set for)
  const spacing = span => spec.valueMajorUnit && span / spec.valueMajorUnit <= 40 ? spec.valueMajorUnit : null;
  if (!vals.length || pinned || !(ticks || spec.includeZero === false || spec.secondarySync || (spec.refLines && spec.refLines.length))) {
    const span = pinned ? fixed.max - fixed.min : 0;
    const unit = pinned ? spacing(span) || niceUnit(span, ticks || 5) : spec.valueMajorUnit || null;
    return { fixed, unit: settle(fixed, unit) };
  }
  const zero = spec.includeZero !== false;
  let lo = fixed.min ?? Math.min(...vals), hi = fixed.max ?? Math.max(...vals);
  if (zero) { lo = Math.min(0, lo); hi = Math.max(0, hi); }
  if (!(hi > lo)) hi = lo + (Math.abs(lo) || 1);
  const unit = spacing(hi - lo) || niceUnit(hi - lo, ticks || 4);
  // labels above the marks need room inside the plot
  const outside = spec.series.some(s => s.labels && !s.refLine) && !/^(inBase|ctr|inEnd)$/.test(spec.labelPos || "");
  const pad = (hi - lo) * (outside ? 0.12 : 0.04);
  const range = { min: fixed.min ?? (zero && lo >= 0 ? 0 : round(Math.floor((lo - pad) / unit) * unit)),
                  max: fixed.max ?? (zero && hi <= 0 ? 0 : round(hi + pad)) };
  return { fixed: range, unit: settle(range, unit) };
}

/** data labels taken from cells (scatter / bubble: Label = a dimension; waterfall steps over their bars)
 * @param {ChartSeries} s @param {ChartRefs["series"][number]} r */
function labelRangeXml(s, r) {
  if (!s.labelTexts || !r.lbl) return "";
  return `<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}" xmlns:c15="${C15}"><c15:datalabelsRange>` +
    `<c15:f>${esc(r.lbl)}</c15:f><c15:dlblRangeCache>${strCache(s.labelTexts)}</c15:dlblRangeCache>` +
    `</c15:datalabelsRange></c:ext></c:extLst>`;
}

/** @param {ChartSpec} spec @param {ChartSeries} s @param {number} k @param {ChartRefs} refs @param {string} type
 *  @param {number[]} [hiddenLabels] points whose data label is left out */
function seriesXml(spec, s, k, refs, type, hiddenLabels) {
  const r = refs.series[k];
  const head = `<c:idx val="${k}"/><c:order val="${k}"/>${serTx(r.tx, s.name)}`;
  const pc = s.pointColors || [];
  const levels = spec.categories ? spec.categories.levels : [];
  // the workbook's opacity (Color → Opacity), in DrawingML's 1/1000 %
  const opacity = s.alpha !== undefined && s.alpha < 1 ? Math.round(Math.max(0, s.alpha) * 100000) : null;
  if (type === "bar") {
    const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>` +
      `<c:spPr>${solid(c, opacity)}<a:ln><a:noFill/></a:ln></c:spPr></c:dPt>` : "").join("");
    const pos = spec.labelPos || (spec.stacked ? null : "outEnd");
    return `<c:ser>${head}<c:spPr>${solid(s.color, opacity)}<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>` +
      `${dpts}${dLbls(spec, s, pos)}${catXml(refs.cat, levels, spec.categoryShown)}${valXml("val", r.val, s.values)}${labelRangeXml(s, r)}</c:ser>`;
  }
  if (type === "line" && s.refLine) {         // reference line: flat, no markers; label above the 2nd point, as Tableau's
    const at = Math.min(1, Math.max(0, (s.values || []).length - 1));
    return `<c:ser>${head}<c:spPr>${refLineLn(s.refLine)}</c:spPr><c:marker><c:symbol val="none"/></c:marker>` +
      `${refLineLabel(spec, s.refLine, "t", "showVal", at)}${catXml(refs.cat, levels, spec.categoryShown)}${valXml("val", r.val, s.values)}<c:smooth val="0"/></c:ser>`;
  }
  if (type === "line") {
    const lineSp = s.line === false ? `<a:ln w="28575"><a:noFill/></a:ln>`
      : `<a:ln w="22225" cap="rnd">${solid(s.color, opacity)}<a:round/></a:ln>`;
    const symbol = s.marker ? (s.markerSymbol || "circle") : "none";
    const size = s.markerSize || 7;
    const dpts = s.marker ? pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/>${markerXml(symbol, c, size)}<c:bubble3D val="0"/></c:dPt>` : "").join("") : "";
    return `<c:ser>${head}<c:spPr>${lineSp}</c:spPr>${markerXml(symbol, s.color, size)}${dpts}` +
      `${dLbls(spec, s, "t", hiddenLabels)}${catXml(refs.cat, levels, spec.categoryShown)}${valXml("val", r.val, s.values)}<c:smooth val="0"/>${labelRangeXml(s, r)}</c:ser>`;
  }
  if (type === "area") {
    return `<c:ser>${head}<c:spPr>${solid(s.color, opacity !== null || s.alpha !== undefined ? opacity : spec.stacked ? null : 75000)}<a:ln><a:noFill/></a:ln></c:spPr>` +
      `${dLbls(spec, s, null)}${catXml(refs.cat, levels, spec.categoryShown)}${valXml("val", r.val, s.values)}</c:ser>`;
  }
  if (type === "pie") {
    const dpts = pc.map((c, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/>` +
      `<c:spPr>${solid(c || s.color)}${line("FFFFFF", 12700)}</c:spPr></c:dPt>`).join("");
    return `<c:ser>${head}${dpts}${dLbls(spec, s, spec.kind === "pie" ? "bestFit" : null)}` +
      `${catXml(refs.cat, levels, spec.categoryShown)}${valXml("val", r.val, s.values)}${labelRangeXml(s, r)}</c:ser>`;
  }
  if (type === "bubble") {
    // packed bubbles: opaque with a white outline like Tableau; map marks slightly see-through
    const fill = c => `${solid(c, spec.packed ? null : 80000)}${line("FFFFFF", 9525)}`;
    const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>` +
      `<c:spPr>${fill(c)}</c:spPr></c:dPt>` : "").join("");
    return `<c:ser>${head}<c:spPr>${fill(s.color)}</c:spPr><c:invertIfNegative val="0"/>${dpts}` +
      `${dLbls(spec, s, spec.packed ? "ctr" : "r", hiddenLabels)}${valXml("xVal", r.x, s.x)}${valXml("yVal", r.y, s.y)}` +
      `${valXml("bubbleSize", r.size, s.size)}<c:bubble3D val="0"/>${labelRangeXml(s, r)}</c:ser>`;
  }
  // scatter; a line of measure against measure joins its points in order, without markers (Tableau's line mark)
  const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/>${markerXml("circle", c, 7)}<c:bubble3D val="0"/></c:dPt>` : "").join("");
  if (s.line) {
    return `<c:ser>${head}<c:spPr><a:ln w="22225" cap="rnd">${solid(s.color, opacity)}<a:round/></a:ln></c:spPr>` +
      `${markerXml(s.marker ? "circle" : "none", s.color, 7)}${dLbls(spec, s, "r")}${valXml("xVal", r.x, s.x)}${valXml("yVal", r.y, s.y)}` +
      `<c:smooth val="0"/>${labelRangeXml(s, r)}</c:ser>`;
  }
  return `<c:ser>${head}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>${markerXml("circle", s.color, 7)}${dpts}` +
    `${dLbls(spec, s, "r")}${valXml("xVal", r.x, s.x)}${valXml("yVal", r.y, s.y)}<c:smooth val="0"/>` +
    labelRangeXml(s, r) + `</c:ser>`;
}

/** @param {ChartSpec} spec @param {number} id @param {number} cross @param {{ deleted?: boolean, rot?: number }} [o] rot: label rotation (60000ths of a degree) */
function catAxis(spec, id, cross, o = {}) {
  const horizontal = spec.barDir === "bar";
  const multi = spec.categories && spec.categories.levels.length > 1;
  return `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="${horizontal ? "maxMin" : "minMax"}"/></c:scaling>` +
    `<c:delete val="${o.deleted ? 1 : 0}"/><c:axPos val="${horizontal ? "l" : "b"}"/>` +
    `${o.deleted ? "" : title(spec.categoryTitle, spec.font, horizontal)}` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>` +
    `<c:tickLblPos val="${o.deleted ? "none" : spec.valueReversed ? "nextTo" : "low"}"/><c:spPr>${spec.axisLine === false ? "<a:ln><a:noFill/></a:ln>" : line("D4D4D4", 9525)}</c:spPr>${txPr(spec.font, { rot: o.rot })}` +
    `<c:crossAx val="${cross}"/><c:crosses val="${spec.valueReversed && !o.deleted ? "max" : "autoZero"}"/><c:auto val="1"/><c:lblAlgn val="ctr"/>` +
    `<c:lblOffset val="100"/>${o.rot !== undefined ? '<c:tickLblSkip val="1"/>' : ""}<c:noMultiLvlLbl val="${multi ? 0 : 1}"/></c:catAx>`;
}

/* Tableau axes "include zero" by default; Excel would otherwise auto-scale from a non-zero minimum */
function zeroScaling(values) {
  const nums = values.filter(v => num(v) !== null);
  if (!nums.length) return "";
  if (Math.min(...nums) >= 0) return `<c:min val="0"/>`;
  if (Math.max(...nums) <= 0) return `<c:max val="0"/>`;
  return "";
}

/* axis ticks: Tableau's automatic axis drops decimals and abbreviates thousands (100K, 1.5M);
 * labels and the data cells keep the full number format */
function axisFmt(fmt, values) {
  const nums = (values || []).filter(v => num(v) !== null).map(Math.abs);
  if (!fmt || !nums.length) return fmt || "General";
  const max = Math.max(...nums);
  if (max < 1000) {
    if (/%/.test(fmt)) return fmt;
    if (max >= 10) return fmt.replace(/0.0+/g, "0");
    return fmt.replace(/0.0{2,}/g, "0.0");
  }
  const plain = fmt.match(/^("[^"]*")?(#,##0|0)(.0+)?$/);
  if (plain && max >= 1e6) return (plain[1] || "") + '#,##0.0,,"M"';
  if (plain && max >= 1e4) return (plain[1] || "") + '#,##0,"K"';
  return fmt.replace(/0.0+/g, "0");
}

/**
 * @param {ChartSpec} spec @param {number} id @param {number} cross
 * @param {{ pos?: string, crosses?: string, grid?: boolean, fixed?: { min?: number, max?: number }, title?: string,
 *           numFmt?: string, values?: (number | null)[], deleted?: boolean, hidden?: boolean, lowLabels?: boolean, midCat?: boolean,
 *           majorUnit?: number | null, tickFmt?: string, reversed?: boolean }} [o] tickFmt = the workbook's own tick format, used as is;
 *   hidden = kept for its scale but not drawn (no labels, tick marks or line) – Excel draws the series of a deleted value
 *   axis on the other axis group's value axis
 */
function valAxis(spec, id, cross, o = {}) {
  const horizontal = spec.barDir === "bar" && spec.kind !== "scatter";
  const pos = o.pos || (horizontal ? "b" : "l");
  const crosses = o.crosses || (horizontal ? "max" : "autoZero");
  const grid = o.grid !== false && spec.gridlines !== false
    ? `<c:majorGridlines><c:spPr>${line("EBEBEB", 9525)}</c:spPr></c:majorGridlines>` : "";
  const fixed = o.fixed || {};                       // CT_Scaling order: orientation, max, min
  const scale = fixed.max !== undefined || fixed.min !== undefined
    ? (fixed.max !== undefined ? `<c:max val="${fixed.max}"/>` : "") + (fixed.min !== undefined ? `<c:min val="${fixed.min}"/>` : "")
    : spec.includeZero === false || !o.values ? "" : zeroScaling(o.values);
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="${o.reversed ? "maxMin" : "minMax"}"/>${scale}</c:scaling><c:delete val="${o.deleted ? 1 : 0}"/>` +
    `<c:axPos val="${pos}"/>${grid}${title(o.title, spec.font, pos === "l" || pos === "r")}` +
    `<c:numFmt formatCode="${esc(o.tickFmt || axisFmt(o.numFmt || "General", o.values))}" sourceLinked="0"/><c:majorTickMark val="none"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="${o.hidden ? "none" : o.lowLabels ? "low" : "nextTo"}"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>` +
    `${txPr(spec.font)}<c:crossAx val="${cross}"/><c:crosses val="${crosses}"/>` +
    `<c:crossBetween val="${o.midCat ? "midCat" : "between"}"/>${o.majorUnit ? `<c:majorUnit val="${o.majorUnit}"/>` : ""}</c:valAx>`;
}

/* Excel's largest bubble: diameter / the plot area's shorter side for a bubble scale s (1 = 100 %).
 * Measured in Excel (one bubble, scales 10–300 %, three plot shapes): 0.03 at 10 %, 0.234 at 100 %,
 * 0.485 at 300 % – A·s / (1 + B·s) fits within 1 %. */
const BUBBLE_A = 0.3014, BUBBLE_B = 0.288;
const bubbleRatio = s => BUBBLE_A * s / (1 + BUBBLE_B * s);
const bubbleScaleFor = ratio => ratio >= bubbleRatio(3) ? 3 : Math.max(0.01, ratio / (BUBBLE_A - BUBBLE_B * ratio));

/** a packed-bubble chart's plot area, as fractions of the chart (it fills the chart: no axes) */
const bubbleFrame = spec => ({ x: 0.02, y: 0.03, w: spec.legend ? 0.74 : 0.96, h: 0.94 });

/**
 * Tableau's tooltip on each packed bubble: an invisible circle over it (px inside the chart) whose hyperlink
 * ScreenTip is the tooltip text, placed with the same fit Excel draws the bubbles at.
 * @param {ChartSpec} spec @param {{ widthPx: number, heightPx: number }} size @param {string} link
 */
function bubbleTips(spec, size, link) {
  if (spec.kind !== "bubble" || !spec.packed) return [];
  const f = bubbleFrame(spec);
  const plot = { w: size.widthPx * f.w, h: size.heightPx * f.h };
  const fit = xyFit(spec, plot);
  if (!fit.unit) return [];
  return spec.series.flatMap(s => {
    if (!s.tooltips || !s.size) return [];
    const max = Math.max(...s.size.filter(v => num(v) !== null && v > 0));
    return s.size.map((v, i) => {
      if (!s.tooltips[i] || !(v > 0)) return null;
      const r = Math.sqrt(v / max) * fit.unit;
      const cx = size.widthPx * f.x + (s.x[i] - fit.x.min) * fit.unit, cy = size.heightPx * f.y + (fit.y.max - s.y[i]) * fit.unit;
      return { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, prst: "ellipse", text: s.tooltips[i], link };
    }).filter(Boolean);
  });
}

/**
 * Packed bubbles: fixed axis ranges and the bubble scale for the plot area's size in px, so the bubbles touch
 * without overlapping.
 * @param {ChartSpec} spec @param {{ w: number, h: number } | null} plot
 * @returns {{ bubbleScale: number, unit?: number, x?: { min: number, max: number }, y?: { min: number, max: number } }}
 */
function xyFit(spec, plot) {
  if (!plot || !spec.packed) return { bubbleScale: 100 };
  const W = plot.w, H = plot.h, shorter = Math.min(W, H);
  const b = spec.packed;                                      // radius units: largest bubble radius = 1
  const want = 0.98 * Math.min(W / b.w, H / b.h);             // px per radius unit that fits the packing
  const bubbleScale = Math.max(1, Math.round(100 * bubbleScaleFor(2 * want / shorter)));
  const unit = bubbleRatio(bubbleScale / 100) * shorter / 2;  // px per radius unit as Excel draws it
  return { bubbleScale, unit, x: { min: b.cx - W / 2 / unit, max: b.cx + W / 2 / unit },
           y: { min: b.cy - H / 2 / unit, max: b.cy + H / 2 / unit } };
}

/**
 * Tableau's column headers stay horizontal and a name too long for its slot is cut short with "..": the
 * labels shown on a vertical chart's category axis, or null when every name fits (or the labels are rotated).
 * @param {ChartSpec} spec @param {{ w: number, h: number } | null} axisPlot
 */
function truncatedCategories(spec, axisPlot) {
  const levels = spec.categories ? spec.categories.levels : [];
  if (!axisPlot || spec.barDir === "bar" || spec.categoryRotation || spec.categoryAxisHidden || levels.length !== 1 || !levels[0].length) return null;
  const names = levels[0].map(c => String(c == null ? "" : c));
  const charPx = ((spec.font && spec.font.size) || 9) * 4 / 3 * 0.5;      // average character of the label font
  const fit = Math.max(4, Math.floor((axisPlot.w / names.length - 4) / charPx));
  const shown = names.map(n => n.length > fit ? n.slice(0, Math.max(1, fit - 2)).trimEnd() + ".." : n);
  return shown.some((s, i) => s !== names[i]) ? shown : null;
}

/** @param {ChartSpec} spec @param {ChartRefs} refs @param {{ w: number, h: number } | null} [plot] plot area in px (manual layout)
 *  @param {{ w: number, h: number } | null} [axisPlot] approximate plot area of a chart with axes, for Tableau-like tick spacing */
function plotAreaXml(spec, refs, plot = null, axisPlot = null) {
  spec.categoryShown = truncatedCategories(spec, axisPlot);
  const k = spec.kind;
  if (k === "pie" || k === "doughnut") {
    const ser = seriesXml(spec, spec.series[0], 0, refs, "pie");
    return k === "pie"
      ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
      : `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="${spec.holeSize || 55}"/></c:doughnutChart>`;
  }
  if (k === "scatter" || k === "bubble") {
    const fit = xyFit(spec, plot);
    // packed bubbles: a label is centred on its bubble; with "allow labels to overlap" off (Tableau's default) one
    // that would overlap a bigger bubble's label is hidden, as Tableau hides it
    const fontPx = (spec.font && spec.font.size || 9) * 4 / 3, charPx = fontPx * 0.55;
    const overlapping = s => {
      if (!spec.packed || !s.labelTexts || !fit.unit || spec.labelCull === false) return [];
      const boxes = s.labelTexts.map((t, i) => {
        const lines = String(t || "").split("\n");
        const w = Math.max(...lines.map(l => l.length)) * charPx, h = lines.length * fontPx * 1.2;
        const cx = (s.x[i] - fit.x.min) * fit.unit, cy = (fit.y.max - s.y[i]) * fit.unit;
        return { i, x: cx - w / 2, y: cy - h / 2, w, h, size: s.size[i] || 0 };
      });
      const kept = [], hidden = [];
      boxes.sort((a, b) => b.size - a.size).forEach(b => {
        const hit = kept.some(o => b.x < o.x + o.w && o.x < b.x + b.w && b.y < o.y + o.h && o.y < b.y + b.h);
        (hit ? hidden : kept).push(b);
      });
      return hidden.map(b => b.i);
    };
    const sers = spec.series.map((s, i) => seriesXml(spec, s, i, refs, k, overlapping(s))).join("");
    const group = k === "bubble"
      ? `<c:bubbleChart><c:varyColors val="0"/>${sers}<c:bubbleScale val="${fit.bubbleScale}"/><c:showNegBubbles val="0"/>` +
        `<c:sizeRepresents val="area"/><c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/></c:bubbleChart>`
      : `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}` +
        `<c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/></c:scatterChart>`;
    return group +
      valAxis(spec, AX.cat, AX.val, { pos: "b", crosses: "autoZero", title: spec.axesHidden || spec.xAxisHidden ? null : spec.xTitle,
        numFmt: spec.xNumFmt, tickFmt: spec.xAxisNumFmt, midCat: true, grid: spec.xGridlines !== false, lowLabels: true, values: spec.series.flatMap(s => s.x),
        fixed: fit.x || { min: spec.xMin, max: spec.xMax }, deleted: spec.axesHidden || spec.xAxisHidden, majorUnit: spec.xMajorUnit, reversed: spec.xReversed }) +
      valAxis(spec, AX.val, AX.cat, { pos: "l", crosses: "autoZero", title: spec.axesHidden || spec.valueAxisHidden ? null : spec.valueTitle,
        numFmt: spec.numFmt, tickFmt: spec.valueAxisNumFmt, midCat: true, lowLabels: true, values: spec.series.flatMap(s => s.y),
        fixed: fit.y || { min: spec.valueMin, max: spec.valueMax }, deleted: spec.axesHidden || spec.valueAxisHidden, majorUnit: spec.valueMajorUnit,
        reversed: spec.valueReversed });
  }
  // labels written over the bars (waterfall steps): as Tableau, a label that would overlap one already
  // placed is left out – positions estimated from the plot size and the axis range
  const { fixed, unit } = valueRange(spec, axisPlot);
  /** @type {Map<number, number[]>} */
  const culled = new Map();
  if (axisPlot && spec.labelCull !== false && fixed.min !== undefined && fixed.max !== undefined && fixed.max > fixed.min) {
    const size = (spec.labelFont && spec.labelFont.size) || spec.font.size || 9, h = size * 4 / 3 + 2;
    spec.series.forEach((s, i) => {
      if (!s.labelTexts || (s.type || k) !== "line") return;
      const slot = axisPlot.w / Math.max(1, s.values.length), kept = [], hide = [];
      s.values.forEach((v, j) => {
        const t = s.labelTexts[j], n = num(v);
        if (n === null || !t) return;
        const w = String(t).length * size * 0.62 + 4, x = (j + 0.5) * slot;
        const y = axisPlot.h * (fixed.max - n) / (fixed.max - fixed.min) - h / 2 - 3;   // a label just above the point
        const box = { l: x - w / 2, r: x + w / 2, t: y - h / 2, b: y + h / 2 };
        if (kept.some(q => q.l < box.r && box.l < q.r && q.t < box.b && box.t < q.b)) hide.push(j); else kept.push(box);
      });
      if (hide.length) culled.set(i, hide);
    });
  }
  // bar / line / area / combo: one chart group per (type, axis)
  const groups = [];
  spec.series.forEach((s, i) => {
    const type = s.type || k;
    const key = type + (s.secondary ? "2" : "1");
    let g = groups.find(x => x.key === key);
    if (!g) groups.push(g = { key, type, secondary: !!s.secondary, items: [] });
    g.items.push(seriesXml(spec, s, i, refs, type, culled.get(i)));
  });
  const hasSecondary = groups.some(g => g.secondary);
  const xml = groups.map(g => {
    const ax = g.secondary ? `<c:axId val="${AX.cat2}"/><c:axId val="${AX.val2}"/>` : `<c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/>`;
    const sers = g.items.join("");
    if (g.type === "bar") {
      const grouping = spec.percent ? "percentStacked" : spec.stacked ? "stacked" : "clustered";
      const gap = g.secondary && spec.secondaryGapWidth !== undefined ? spec.secondaryGapWidth : spec.gapWidth ?? 60;
      return `<c:barChart><c:barDir val="${spec.barDir || "col"}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers}` +
        `<c:gapWidth val="${gap}"/>${spec.stacked || spec.overlap ? '<c:overlap val="100"/>' : ""}${ax}</c:barChart>`;
    }
    if (g.type === "area") {
      return `<c:areaChart><c:grouping val="${spec.percent ? "percentStacked" : spec.stacked ? "stacked" : "standard"}"/><c:varyColors val="0"/>${sers}${ax}</c:areaChart>`;
    }
    // box plot: high-low lines = whiskers, up/down bars between first (Q1) and last (Q3) series = box
    const box = spec.boxPlot
      ? `<c:hiLowLines><c:spPr>${line(spec.boxPlot.color, 12700)}</c:spPr></c:hiLowLines>` +
        `<c:upDownBars><c:gapWidth val="${spec.gapWidth ?? 80}"/>` +
        `<c:upBars><c:spPr>${solid(spec.boxPlot.color, 35000)}${line(spec.boxPlot.color, 12700)}</c:spPr></c:upBars>` +
        `<c:downBars><c:spPr>${solid(spec.boxPlot.color, 35000)}${line(spec.boxPlot.color, 12700)}</c:spPr></c:downBars></c:upDownBars>`
      : spec.hiLowLines ? `<c:hiLowLines><c:spPr>${line(spec.hiLowLines.color, 19050)}</c:spPr></c:hiLowLines>` : "";
    return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${box}<c:marker val="1"/>${ax}</c:lineChart>`;
  }).join("");
  const axisValues = secondary => spec.series.filter(s => !!s.secondary === secondary).flatMap(s => s.values);
  // reference lines across horizontal bars: XY lines on hidden x2 / y2 axes, x2 pinned to the bars' value range – never
  // beside a secondary axis: Excel cannot open a workbook with a chart of three axis groups (the model leaves those out)
  let overlay = "", overlayAxes = "";
  const targets = spec.barDir === "bar" && spec.targets ? 1 : 0;
  const dots = spec.barDir === "bar" && spec.overlay ? spec.overlay : [];
  if (spec.barDir === "bar" && !hasSecondary && ((spec.refLines && spec.refLines.length) || targets || dots.length) &&
      fixed.min !== undefined && fixed.max !== undefined) {
    const k0 = spec.series.length + (spec.refLines || []).length + targets;
    overlay = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>` +
      (spec.refLines || []).map((rl, i) => refLineXY(spec, rl, spec.series.length + i)).join("") +
      (targets ? targetTicksXY(spec, spec.series.length + (spec.refLines || []).length) : "") +
      dots.map((o, j) => overlayXY(spec, o, k0 + j)).join("") +
      `<c:axId val="${AX.x2}"/><c:axId val="${AX.y2}"/></c:scatterChart>`;
    overlayAxes = hiddenValAx(AX.x2, AX.y2, "t", fixed.min, fixed.max) + hiddenValAx(AX.y2, AX.x2, "r", 0, 1);
  }
  // a hidden value axis beside a second axis group is kept, not drawn: Excel would draw a deleted axis's series on the
  // other group's axis, losing the scale Tableau gives each axis
  let axes = catAxis(spec, AX.cat, AX.val, { deleted: spec.categoryAxisHidden, rot: (spec.categoryRotation || 0) * 60000 }) +
    valAxis(spec, AX.val, AX.cat, { title: spec.valueAxisHidden ? null : spec.valueTitle, numFmt: spec.numFmt, values: axisValues(false),
                                    fixed, deleted: spec.valueAxisHidden && !hasSecondary, hidden: spec.valueAxisHidden && hasSecondary,
                                    majorUnit: unit, tickFmt: spec.valueAxisNumFmt, reversed: spec.valueReversed });
  if (hasSecondary) {
    axes += catAxis(spec, AX.cat2, AX.val2, { deleted: true }) +
      valAxis(spec, AX.val2, AX.cat2, { pos: spec.barDir === "bar" ? "t" : "r", crosses: "max", grid: false,
        title: spec.secondaryAxisHidden ? null : spec.secondaryTitle, numFmt: spec.secondaryNumFmt || spec.numFmt, values: axisValues(true),
        hidden: spec.secondaryAxisHidden, tickFmt: spec.secondaryAxisNumFmt,
        // synchronized: the same range as the primary axis, so both measures are drawn to one scale
        ...(spec.secondarySync ? { fixed, majorUnit: unit } : {}) });
  }
  return xml + overlay + axes + overlayAxes;
}

/** chart part XML (xl/charts/chartN.xml)
 * @param {ChartSpec} spec @param {ChartRefs} refs
 * @param {{ widthPx?: number, heightPx?: number }} [size] the chart frame, for maps / packed bubbles
 * @returns {string} */
function chartXml(spec, refs, size) {
  // no axes: the plot area fills the frame (room for the legend), so its size in px is known
  const frame = spec.axesHidden && spec.packed ? bubbleFrame(spec) : null;
  const plot = frame && { w: ((size && size.widthPx) || 600) * frame.w, h: ((size && size.heightPx) || 400) * frame.h };
  const axisPlot = size && size.widthPx ? { w: size.widthPx * 0.85, h: (size.heightPx || 300) * 0.72 } : null;
  const layout = frame
    ? `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${frame.x}"/><c:y val="${frame.y}"/><c:w val="${frame.w}"/><c:h val="${frame.h}"/></c:manualLayout></c:layout>`
    : "<c:layout/>";
  const legend = spec.legend
    ? `<c:legend><c:legendPos val="r"/><c:overlay val="0"/>${txPr(spec.font)}</c:legend>` : "";
  // a donut's centre text: the chart title laid over the hole, the ring centred in a fixed plot area
  const center = spec.kind === "doughnut" && spec.centerLabel && spec.centerLabel.length ? spec.centerLabel : null;
  let titleXml = "", plotLayout = layout;
  if (center) {
    const W = (size && size.widthPx) || 400, H = (size && size.heightPx) || 300;
    const box = { x: 0.04, y: 0.04, w: spec.legend ? 0.66 : 0.92, h: 0.92 };
    const pt = s => s.size || spec.font.size || 9;
    const tw = Math.max(...center.map(l => l.reduce((w, s) => w + String(s.text).length * pt(s) * 0.62, 0))) + 10;
    const th = center.reduce((h, l) => h + Math.max(...l.map(pt)) * 4 / 3 * 1.25, 0) + 6;
    const cx = box.x + box.w / 2, cy = box.y + box.h / 2;
    titleXml = `<c:title><c:tx><c:rich><a:bodyPr wrap="none"/><a:lstStyle/>` + center.map(l => `<a:p><a:pPr algn="ctr"><a:defRPr/></a:pPr>` +
        l.map(s => `<a:r>${runProps({ name: s.font || spec.font.name, size: pt(s), color: s.color || spec.font.color }, { bold: s.bold }, "rPr")}` +
          `<a:t>${esc(s.text)}</a:t></a:r>`).join("") + `</a:p>`).join("") + `</c:rich></c:tx>` +
      `<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${Math.max(0, cx - tw / 2 / W).toFixed(4)}"/><c:y val="${Math.max(0, cy - th / 2 / H).toFixed(4)}"/></c:manualLayout></c:layout>` +
      `<c:overlay val="1"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:title>`;
    plotLayout = `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${box.x}"/><c:y val="${box.y}"/><c:w val="${box.w}"/><c:h val="${box.h}"/></c:manualLayout></c:layout>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>` +
    `<c:chart>${titleXml}<c:autoTitleDeleted val="${titleXml ? 0 : 1}"/><c:plotArea>${plotLayout}${plotAreaXml(spec, refs, plot, axisPlot)}` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>${legend}` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr>${spec.background === null ? "<a:noFill/>" : solid(spec.background || "FFFFFF")}<a:ln><a:noFill/></a:ln></c:spPr>${txPr(spec.font)}` +
    `<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>` +
    `<c:pageSetup/></c:printSettings></c:chartSpace>`;
}


/* ═══ charts/writer/chartex.js ════════════════════════════════════════════════════════════════════ */
/* Excel 2016+ chartex parts (treemap) and their drawing anchors. */

/* Excel 2016+ "chartex" part – treemap. Category levels outer → inner in the spec,
 * inner (leaf) level first in the XML. Excel only accepts chartex formulas that are
 * hidden workbook names (_xlchart.v1.N), so refs here are those names, not ranges. */
/** @param {ChartSpec} spec @param {Pick<ChartRefs, "cat" | "series">} refs workbook names, not ranges @param {string} uid series GUID @returns {string} */
function chartExXml(spec, refs, uid) {
  const levels = spec.categories.levels;
  const n = levels.length ? levels[0].length : 0;
  const s = spec.series[0];
  const lvl = lv => `<cx:lvl ptCount="${n}">${lv.map((v, i) => `<cx:pt idx="${i}">${esc(v)}</cx:pt>`).join("")}</cx:lvl>`;
  const vals = s.values.map((v, i) => num(v) === null ? "" : `<cx:pt idx="${i}">${v}</cx:pt>`).join("");
  const tile = c => `<cx:spPr>${solid(c)}${line("FFFFFF", 12700)}</cx:spPr>`;
  const font = spec.font || /** @type {ChartSpec["font"]} */ ({});
  const labelPr = `<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round((font.size || 9) * 100)}">` +
    `<a:latin typeface="${esc(font.name || "Arial")}"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></cx:txPr>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<cx:chartSpace xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:cx="${NS.cx}">` +
    `<cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>${esc(refs.cat)}</cx:f>${levels.slice().reverse().map(lvl).join("")}</cx:strDim>` +
    `<cx:numDim type="size"><cx:f>${esc(refs.series[0].val)}</cx:f><cx:lvl ptCount="${n}" formatCode="${esc(spec.numFmt || "General")}">${vals}</cx:lvl>` +
    `</cx:numDim></cx:data></cx:chartData>` +
    `<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="treemap" uniqueId="${uid}">` +
    `<cx:tx><cx:txData><cx:f>${esc(refs.series[0].tx)}</cx:f><cx:v>${esc(s.name)}</cx:v></cx:txData></cx:tx>` +
    tile(s.color) +
    (s.pointColors || []).map((c, i) => c ? `<cx:dataPt idx="${i}">${tile(c)}</cx:dataPt>` : "").join("") +
    (s.labels ? `<cx:dataLabels pos="inEnd">` +
      (s.labelParts && s.labelParts.value ? `<cx:numFmt formatCode="${esc(spec.numFmt || "General")}" sourceLinked="0"/>` : "") +
      `${labelPr}<cx:visibility seriesName="0" categoryName="1" value="${s.labelParts && s.labelParts.value ? 1 : 0}"/>` +
      (s.labelParts && s.labelParts.value ? `<cx:separator>${"\n"}</cx:separator>` : "") + `</cx:dataLabels>` : "") +
    `<cx:dataId val="0"/><cx:layoutPr><cx:parentLabelLayout val="${levels.length > 1 ? "banner" : "none"}"/></cx:layoutPr>` +
    `</cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>` +
    `<cx:spPr>${spec.background === null ? "<a:noFill/>" : solid(spec.background || "FFFFFF")}<a:ln><a:noFill/></a:ln></cx:spPr></cx:chartSpace>`;
}

/* chartex frames sit in mc:AlternateContent; older Excel shows the fallback rectangle */
/** @param {ChartJob} chart @param {number} id @param {string} rid */
function anchorExXml(chart, id, rid) {
  const cx = Math.max(1, Math.round(chart.widthPx * EMU_PER_PX));
  const cy = Math.max(1, Math.round(chart.heightPx * EMU_PER_PX));
  const name = esc(chart.name || "Chart " + id);
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.col}</xdr:col><xdr:colOff>${Math.round((chart.colOffPx || 0) * EMU_PER_PX)}</xdr:colOff>` +
    `<xdr:row>${chart.row}</xdr:row><xdr:rowOff>${Math.round((chart.rowOffPx || 0) * EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/>` +
    `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:cx1="${NS.cx1}" Requires="cx1">` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>` +
    `<a:graphicData uri="${NS.cx}"><cx:chart xmlns:cx="${NS.cx}" xmlns:r="${NS.r}" r:id="${rid}"/></a:graphicData>` +
    `</a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback>` +
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>` +
    `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/>` +
    `<a:t>This treemap needs Excel 2016 or later.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent>` +
    `<xdr:clientData/></xdr:oneCellAnchor>`;
}


/* ═══ charts/writer/chartex-style.js ══════════════════════════════════════════════════════════════ */
/* Excel 2016 default chart style / colour style parts required by chartex charts. */

/* Excel 2016 default chart style / colour style parts (chartex charts will not load without them) */
const CHARTEX_STYLE = "<cs:chartStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" id=\"410\"><cs:axisTitle><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"bg1\"><a:lumMod val=\"65000\"/></a:schemeClr></a:solidFill><a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill></a:ln></cs:spPr><cs:defRPr sz=\"900\"/></cs:axisTitle><cs:categoryAxis><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz=\"900\"/></cs:categoryAxis><cs:chartArea mods=\"allowNoFillOverride allowNoLineOverride\"><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz=\"1000\"/></cs:chartArea><cs:dataLabel><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></cs:fontRef><cs:defRPr sz=\"900\"/></cs:dataLabel><cs:dataLabelCallout><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"dk1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:ln><a:solidFill><a:schemeClr val=\"dk1\"><a:lumMod val=\"25000\"/><a:lumOff val=\"75000\"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz=\"900\"/><cs:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"clip\" horzOverflow=\"clip\" vert=\"horz\" wrap=\"square\" lIns=\"36576\" tIns=\"18288\" rIns=\"36576\" bIns=\"18288\" anchor=\"ctr\" anchorCtr=\"1\"><a:spAutoFit/></cs:bodyPr></cs:dataLabelCallout><cs:dataPoint><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:fillRef><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill></a:ln></cs:spPr></cs:dataPoint><cs:dataPoint3D><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:fillRef><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></cs:spPr></cs:dataPoint3D><cs:dataPointLine><cs:lnRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:lnRef><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"28575\" cap=\"rnd\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointLine><cs:dataPointMarker><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:fillRef><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill></a:ln></cs:spPr></cs:dataPointMarker><cs:dataPointMarkerLayout symbol=\"circle\" size=\"5\"/><cs:dataPointWireframe><cs:lnRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:lnRef><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"28575\" cap=\"rnd\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointWireframe><cs:dataTable><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz=\"900\"/></cs:dataTable><cs:downBar><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"dk1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"dk1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:downBar><cs:dropLine><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"35000\"/><a:lumOff val=\"65000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:dropLine><cs:errorBar><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:errorBar><cs:floor><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef></cs:floor><cs:gridlineMajor><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:gridlineMajor><cs:gridlineMinor><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:gridlineMinor><cs:hiLoLine><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"75000\"/><a:lumOff val=\"25000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:hiLoLine><cs:leaderLine><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"35000\"/><a:lumOff val=\"65000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:leaderLine><cs:legend><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:defRPr sz=\"900\"/></cs:legend><cs:plotArea mods=\"allowNoFillOverride allowNoLineOverride\"><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef></cs:plotArea><cs:plotArea3D mods=\"allowNoFillOverride allowNoLineOverride\"><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef></cs:plotArea3D><cs:seriesAxis><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz=\"900\"/></cs:seriesAxis><cs:seriesLine><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\"><a:solidFill><a:srgbClr val=\"D9D9D9\"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:seriesLine><cs:title><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:defRPr sz=\"1400\"/></cs:title><cs:trendline><cs:lnRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:lnRef><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"19050\" cap=\"rnd\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"sysDash\"/></a:ln></cs:spPr></cs:trendline><cs:trendlineLabel><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:defRPr sz=\"900\"/></cs:trendlineLabel><cs:upBar><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"dk1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:upBar><cs:valueAxis><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:defRPr sz=\"900\"/></cs:valueAxis><cs:wall><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef></cs:wall></cs:chartStyle>";

const CHARTEX_COLORS = "<cs:colorStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" meth=\"cycle\" id=\"10\"><a:schemeClr val=\"accent1\"/><a:schemeClr val=\"accent2\"/><a:schemeClr val=\"accent3\"/><a:schemeClr val=\"accent4\"/><a:schemeClr val=\"accent5\"/><a:schemeClr val=\"accent6\"/><cs:variation/><cs:variation><a:lumMod val=\"60000\"/></cs:variation><cs:variation><a:lumMod val=\"80000\"/><a:lumOff val=\"20000\"/></cs:variation><cs:variation><a:lumMod val=\"80000\"/></cs:variation><cs:variation><a:lumMod val=\"60000\"/><a:lumOff val=\"40000\"/></cs:variation><cs:variation><a:lumMod val=\"50000\"/></cs:variation><cs:variation><a:lumMod val=\"70000\"/><a:lumOff val=\"30000\"/></cs:variation><cs:variation><a:lumMod val=\"70000\"/></cs:variation><cs:variation><a:lumMod val=\"50000\"/><a:lumOff val=\"50000\"/></cs:variation></cs:colorStyle>";


/* ═══ charts/writer/package.js ════════════════════════════════════════════════════════════════════ */
/* XLSX package surgery: drawing / chart parts, relationships, content types. */

/** <sheetPr> children in schema order (tabColor, outlinePr, pageSetUpPr): ExcelJS writes pageSetUpPr first
 * when a sheet has both fit-to-page and outline groups, and Excel then refuses to open the file.
 * Anything else inside <sheetPr> leaves it as it is. @param {string} xml @returns {string} */
function orderSheetPr(xml) {
  return xml.replace(/<sheetPr\b([^>]*)>([\s\S]*?)<\/sheetPr>/, (m, attrs, inner) => {
    const pick = tag => (inner.match(new RegExp(`<${tag}\\b[^>]*?(?:\\/>|>[\\s\\S]*?<\\/${tag}>)`)) || [""])[0];
    const out = pick("tabColor") + pick("outlinePr") + pick("pageSetUpPr");
    return out.length === inner.length ? `<sheetPr${attrs}>${out}</sheetPr>` : m;
  });
}

/** the package with every worksheet's <sheetPr> in schema order (unchanged buffer when nothing moves)
 * @param {ArrayBuffer | Uint8Array} buffer */
async function fixSheetProperties(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  let changed = false;
  for (const name of Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
    const xml = await zip.file(name).async("string");
    const fixed = orderSheetPr(xml);
    if (fixed !== xml) { zip.file(name, fixed); changed = true; }
  }
  return changed ? zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }) : buffer;
}

/** @param {ChartJob} chart @param {number} id @param {string} rid */
function anchorXml(chart, id, rid) {
  const cx = Math.max(1, Math.round(chart.widthPx * EMU_PER_PX));
  const cy = Math.max(1, Math.round(chart.heightPx * EMU_PER_PX));
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.col}</xdr:col><xdr:colOff>${Math.round((chart.colOffPx || 0) * EMU_PER_PX)}</xdr:colOff>` +
    `<xdr:row>${chart.row}</xdr:row><xdr:rowOff>${Math.round((chart.rowOffPx || 0) * EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>` +
    `<xdr:cNvPr id="${id}" name="${esc(chart.name || "Chart " + id)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>` +
    `<a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${rid}"/></a:graphicData>` +
    `</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
}

function relTarget(relsXml, id) {
  const tag = (relsXml.match(new RegExp(`<Relationship\\b[^>]*\\sId="${id}"[^>]*>`)) || [])[0];
  return tag ? attr(tag, "Target") : null;
}

function nextRelId(relsXml) {
  const ids = [...relsXml.matchAll(/\sId="rId(\d+)"/g)].map(m => +m[1]);
  return "rId" + ((ids.length ? Math.max(...ids) : 0) + 1);
}

function addRel(relsXml, id, type, target) {
  return relsXml.replace("</Relationships>", `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`);
}

function resolve(baseDir, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = (baseDir + target).split("/");
  const out = [];
  parts.forEach(p => { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); });
  return out.join("/");
}

function dirOf(path) { return path.replace(/[^/]+$/, ""); }

function relsPathOf(path) { return dirOf(path) + "_rels/" + path.replace(/^.*\//, "") + ".rels"; }

function nextFreeIndex(zip, prefix, ext) {
  let n = 1;
  while (zip.file(`${prefix}${n}${ext}`)) n++;
  return n;
}

function addOverride(ct, part, type) {
  if (ct.includes(`PartName="/${part}"`)) return ct;
  return ct.replace("</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
}

/* <drawing> must precede these elements in CT_Worksheet */
function insertDrawingTag(sheetXml, rid) {
  const tag = `<drawing r:id="${rid}"/>`;
  const after = ["<legacyDrawing", "<legacyDrawingHF", "<drawingHF", "<picture", "<oleObjects", "<controls",
                 "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"];
  for (const t of after) {
    const i = sheetXml.indexOf(t);
    if (i >= 0) return sheetXml.slice(0, i) + tag + sheetXml.slice(i);
  }
  return sheetXml;
}

/* opts: { sheetIndex, charts: [{ spec, refs, col, row, widthPx, heightPx, name }] } */
/**
 * Adds the charts to the dashboard sheet of a written XLSX.
 * @param {ArrayBuffer | Uint8Array} buffer workbook.xlsx.writeBuffer() output
 * @param {{ sheetIndex?: number, sheetName?: string, charts: ChartJob[] }} opts
 * @returns {Promise<ArrayBuffer | Uint8Array>}
 */
async function injectCharts(buffer, opts) {
  const charts = (opts.charts || []).filter(Boolean);
  if (!charts.length) return buffer;
  const zip = await JSZip.loadAsync(buffer);
  const read = p => zip.file(p) ? zip.file(p).async("string") : Promise.resolve(null);

  const wbXml = await read("xl/workbook.xml");
  const wbRels = await read("xl/_rels/workbook.xml.rels");
  const sheetTags = wbXml.match(/<sheet\b[^>]*>/g) || [];
  // the dashboard's sheet: by name when given (several dashboards, hidden data sheets in between), else by position
  const unescapeXml = v => String(v || "").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  const sheetTag = opts.sheetName != null ? sheetTags.find(t => unescapeXml(attr(t, "name")) === opts.sheetName)
    : sheetTags[opts.sheetIndex || 0];
  if (!sheetTag) throw new Error("Dashboard sheet not found in workbook package");
  const sheetPath = resolve("xl/", relTarget(wbRels, attr(sheetTag, "r:id")));
  const sheetRelsPath = relsPathOf(sheetPath);
  let sheetXml = await read(sheetPath);
  let sheetRels = (await read(sheetRelsPath)) || EMPTY_RELS;
  let ct = await read("[Content_Types].xml");

  if (!/<worksheet\b[^>]*\sxmlns:r=/.test(sheetXml)) {
    sheetXml = sheetXml.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS.r}"`);
  }

  // re-use the drawing ExcelJS created for images, else create one
  let drawingPath, drawingXml, drawingRels;
  const existingRid = (sheetXml.match(/<drawing\b[^>]*r:id="([^"]+)"/) || [])[1];
  if (existingRid) {
    drawingPath = resolve(dirOf(sheetPath), relTarget(sheetRels, existingRid));
    drawingXml = await read(drawingPath);
    drawingRels = (await read(relsPathOf(drawingPath))) || EMPTY_RELS;
  } else {
    const n = nextFreeIndex(zip, "xl/drawings/drawing", ".xml");
    drawingPath = `xl/drawings/drawing${n}.xml`;
    drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"></xdr:wsDr>`;
    drawingRels = EMPTY_RELS;
    const rid = nextRelId(sheetRels);
    sheetRels = addRel(sheetRels, rid, REL_DRAWING, `../drawings/drawing${n}.xml`);
    sheetXml = insertDrawingTag(sheetXml, rid);
    ct = addOverride(ct, drawingPath, CT_DRAWING);
    zip.file(drawingPath, drawingXml);                    // claim the name now: charts' label parts are drawings too
  }
  if (!/<xdr:wsDr\b[^>]*\sxmlns:a=/.test(drawingXml)) {
    drawingXml = drawingXml.replace(/<xdr:wsDr\b/, `<xdr:wsDr xmlns:a="${NS.a}"`);
  }

  const ids = [...drawingXml.matchAll(/<xdr:cNvPr\b[^>]*\sid="(\d+)"/g)].map(m => +m[1]);
  let shapeId = (ids.length ? Math.max(...ids) : 1) + 1;
  let anchors = "";
  // chartex data ranges → hidden defined names, as Excel writes them
  const usedNames = [...wbXml.matchAll(/name="_xlchart\.v1\.(\d+)"/g)].map(m => +m[1]);
  let nameIndex = usedNames.length ? Math.max(...usedNames) + 1 : 0;
  const newNames = [];
  const defineName = ref => {
    const name = `_xlchart.v1.${nameIndex++}`;
    newNames.push(`<definedName name="${name}" hidden="1">${esc(ref)}</definedName>`);
    return name;
  };
  for (const chart of charts) {
    const rid = nextRelId(drawingRels);
    if (chart.spec.kind === "treemap") {
      const n = nextFreeIndex(zip, "xl/charts/chartEx", ".xml");
      const chartPath = `xl/charts/chartEx${n}.xml`;
      const uid = `{6F1D3C1E-0000-4000-8000-${String(n).padStart(12, "0")}}`;
      const named = { cat: defineName(chart.refs.cat),
                      series: chart.refs.series.map(s => ({ tx: defineName(s.tx), val: defineName(s.val) })) };
      zip.file(chartPath, chartExXml(chart.spec, named, uid));
      ct = addOverride(ct, chartPath, CT_CHARTEX);
      const xmlHead = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
      const styleN = nextFreeIndex(zip, "xl/charts/style", ".xml");
      zip.file(`xl/charts/style${styleN}.xml`, xmlHead + CHARTEX_STYLE);
      zip.file(`xl/charts/colors${styleN}.xml`, xmlHead + CHARTEX_COLORS);
      ct = addOverride(ct, `xl/charts/style${styleN}.xml`, CT_CHARTSTYLE);
      ct = addOverride(ct, `xl/charts/colors${styleN}.xml`, CT_CHARTCOLORS);
      zip.file(relsPathOf(chartPath), addRel(addRel(EMPTY_RELS, "rId1", REL_CHARTSTYLE, `style${styleN}.xml`),
                                             "rId2", REL_CHARTCOLORS, `colors${styleN}.xml`));
      drawingRels = addRel(drawingRels, rid, REL_CHARTEX, `../charts/chartEx${n}.xml`);
      anchors += anchorExXml(chart, shapeId++, rid);
      continue;
    }
    const n = nextFreeIndex(zip, "xl/charts/chart", ".xml");
    const chartPath = `xl/charts/chart${n}.xml`;
    // a pie / donut arrives as finished chart XML, with its label shapes and tooltip wedges
    zip.file(chartPath, chart.xml || chartXml(chart.spec, chart.refs, chart));
    ct = addOverride(ct, chartPath, CT_CHART);
    drawingRels = addRel(drawingRels, rid, REL_CHART, `../charts/chart${n}.xml`);
    if (chart.shapes) {                                    // labels / centre text: the chart's own drawing part
      const d = nextFreeIndex(zip, "xl/drawings/drawing", ".xml");
      const shapesPath = `xl/drawings/drawing${d}.xml`;
      zip.file(shapesPath, chart.shapes);
      ct = addOverride(ct, shapesPath, CT_USER_SHAPES);
      zip.file(relsPathOf(chartPath), addRel(EMPTY_RELS, "rId1", REL_USER_SHAPES, `../drawings/drawing${d}.xml`));
    }
    if (chart.tips && chart.tips.length) {                // tooltips: hyperlink ScreenTips on wedges over the slices
      const tipRids = chart.tips.map(t => {
        const tid = nextRelId(drawingRels);
        drawingRels = addRel(drawingRels, tid, REL_HYPERLINK, esc(t.link));
        return tid;
      });
      const ids = { next: shapeId };
      anchors += groupedAnchorXml(chart, rid, ids, tipRids);
      shapeId = ids.next;
      continue;
    }
    anchors += anchorXml(chart, shapeId++, rid);
  }
  // charts first: pictures ExcelJS placed (icons over a chart) stay in front of them
  drawingXml = drawingXml.replace(/(<xdr:wsDr\b[^>]*>)/, (m) => m + anchors);
  if (newNames.length) {                         // CT_Workbook: definedNames follow sheets
    const defs = newNames.join("");
    const wb = wbXml.includes("</definedNames>") ? wbXml.replace("</definedNames>", defs + "</definedNames>")
      : wbXml.includes("<definedNames/>") ? wbXml.replace("<definedNames/>", `<definedNames>${defs}</definedNames>`)
      : wbXml.replace("</sheets>", `</sheets><definedNames>${defs}</definedNames>`);
    zip.file("xl/workbook.xml", wb);
  }

  zip.file(sheetPath, sheetXml);
  zip.file(sheetRelsPath, sheetRels);
  zip.file(drawingPath, drawingXml);
  zip.file(relsPathOf(drawingPath), drawingRels);
  zip.file("[Content_Types].xml", ct);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}


/* ═══ charts/writer/index.js ══════════════════════════════════════════════════════════════════════ */
/* ══════════════════════════════════════════════════════════════════════════
 * EXCEL NATIVE CHART WRITER
 * ──────────────────────────────────────────────────────────────────────────
 * ExcelJS cannot author charts, so native charts are produced in two steps:
 *   1. writeChartData(ws, spec, row) – the chart's source data is written to a
 *      (hidden) data sheet with ExcelJS, so every chart stays editable and is
 *      linked to real cells.
 *   2. injectCharts(buffer, opts) – after workbook.xlsx.writeBuffer() the XLSX
 *      package is opened with JSZip and DrawingML chart parts are added to the
 *      dashboard sheet (re-using the drawing ExcelJS made for images, if any).
 *
 * Chart spec (produced by charts/model):
 *   { kind: "bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"bubble"|"combo"|"treemap",
 *     barDir: "col"|"bar", stacked, gapWidth, categories: { names:[], levels:[[…]] },
 *     series: [{ name, type, color (null = invisible), values|x+y(+size), pointColors, secondary,
 *                line, marker, markerSymbol, markerSize, labels, labelNumFmt }],
 *     numFmt, secondaryNumFmt, xNumFmt, valueTitle, secondaryTitle, valueMin, valueMax,
 *     boxPlot: { color } (line chart → up/down bars + high-low lines),
 *     categoryTitle, legend, gridlines, font: { name, size, color },
 *     axesHidden + packed (packed bubbles): axis ranges fitted to the chart size }
 * "treemap" is written as an Excel 2016+ chartex part with an older-Excel fallback shape.
 * Colours are 6-digit hex ("4E79A7"); ARGB ("FF4E79A7") is accepted too.
 * ══════════════════════════════════════════════════════════════════════════ */

const ExcelChartWriter = { writeChartData, injectCharts, chartXml, chartExXml, bubbleTips };


/* ═══ cloud/dashboards.js ═════════════════════════════════════════════════════════════════════════ */
/* =============================================================================
 * OTHER DASHBOARDS (Phase 3) — read dashboards the extension is NOT placed on.
 * -----------------------------------------------------------------------------
 * The Extensions API only sees its own dashboard. For the others we open each one
 * in an invisible Embedding API viz, in the viewer's OWN Tableau session (so row-level
 * security and permissions apply), read every worksheet's data, then close it.
 * The result has the shape the export pipeline uses for the live dashboard:
 *   { dashboard: { name, objects, worksheets, getParametersAsync },
 *     sheets: [{ name, getSummaryDataAsync, getVisualSpecificationAsync, getFiltersAsync, getSelectedMarksAsync }] }
 * Everything is read up front into plain snapshots, so nothing keeps talking to a closed viz.
 * ============================================================================= */

const TFX_WINDOW = typeof window !== "undefined" ? /** @type {any} */ (window) : {};
const VIZ_LOAD_TIMEOUT_MS = TFX_WINDOW.TFX_VIZ_TIMEOUT_MS || 90000;   // window override = tests only
// No "size known" from Tableau within this time = the viz is showing Tableau's sign-in page.
const SIGNIN_DETECT_MS = TFX_WINDOW.TFX_SIGNIN_DETECT_MS || 15000;
// How long the user gets to complete the one-time sign-in.
const SIGNIN_WAIT_MS = TFX_WINDOW.TFX_SIGNIN_WAIT_MS || 300000;
// Upper bound for one live Extensions API read of the current dashboard.
const LIVE_READ_TIMEOUT_MS = TFX_WINDOW.TFX_LIVE_READ_TIMEOUT_MS || 120000;
let EMBED_LIB_PROMISE = null;

/** Visible dashboards in Tableau tab order, from the workbook XML (<windows>). */
function dashboardTabOrder(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const root = doc && doc.documentElement;
  if (!root) return [];
  const windows = getDirectChildByTag(root, "windows");
  const out = [];
  if (windows) {
    for (let i = 0; i < windows.childNodes.length; i++) {
      const w = windows.childNodes[i];
      if (w.nodeType === 1 && w.tagName === "window" && w.getAttribute("class") === "dashboard" &&
          w.getAttribute("hidden") !== "true") out.push(w.getAttribute("name"));
    }
  }
  if (out.length) return out;
  const dashboards = getDirectChildByTag(root, "dashboards");       // fallback: definition order
  if (!dashboards) return [];
  for (let i = 0; i < dashboards.childNodes.length; i++) {
    const d = dashboards.childNodes[i];
    if (d.nodeType === 1 && d.tagName === "dashboard") out.push(d.getAttribute("name"));
  }
  return out;
}

/** Dashboard size from the XML (fixed-size dashboards), so the hidden viz renders at its real size. */
function dashboardPixelSize(xmlString, name) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const d = [...doc.getElementsByTagName("dashboard")].find(x => x.getAttribute("name") === name);
  const size = d && getDirectChildByTag(d, "size");
  const num = k => (size && +size.getAttribute(k)) || 0;
  return { width: num("maxwidth") || num("minwidth") || 1200, height: num("maxheight") || num("minheight") || 900 };
}

function loadEmbeddingApi(libUrl) {
  if (window.customElements && customElements.get("tableau-viz")) return Promise.resolve();
  if (!EMBED_LIB_PROMISE) {
    EMBED_LIB_PROMISE = import(libUrl).catch(e => {
      EMBED_LIB_PROMISE = null;
      throw new Error(`could not load Tableau's Embedding API (${e.message})`);
    });
  }
  return EMBED_LIB_PROMISE;
}

/**
 * One-shot debug switch: localStorage.setItem("tfx_debugVizOnce", "1") shows the hidden dashboards
 * on screen for the NEXT export only, then switches itself off (see clearDebugVizSwitch).
 * The old persistent key "tfx_debugViz" is ignored and removed, so it can't stay on by accident.
 */
function debugVizOn() {
  try {
    localStorage.removeItem("tfx_debugViz");
    return localStorage.getItem("tfx_debugVizOnce") === "1";
  } catch (e) { return false; }
}
function clearDebugVizSwitch() {
  try { localStorage.removeItem("tfx_debugVizOnce"); localStorage.removeItem("tfx_debugViz"); } catch (e) { /* ignore */ }
}

/**
 * Host for the hidden vizzes. It sits INSIDE the visible area but fully transparent:
 * Chrome pauses rendering of cross-origin frames that are off-screen, and a paused
 * Tableau viz never becomes interactive — so "off-screen" is not an option.
 */
function hiddenVizHost() {
  let host = document.getElementById("tfx_viz_host");
  if (!host) {
    host = document.createElement("div");
    host.id = "tfx_viz_host";
    host.setAttribute("aria-hidden", "true");
    document.body.appendChild(host);
  }
  host.style.cssText = debugVizOn()
    ? "position:fixed;left:0;top:0;width:100%;height:100%;overflow:auto;z-index:9999;background:#fff;outline:3px dashed #b3261e;"
    : "position:fixed;left:0;top:0;overflow:hidden;opacity:0;pointer-events:none;";
  return host;
}

/**
 * Open one view invisibly; resolves with the <tableau-viz> once it's interactive.
 * Tracks Tableau's earlier "size known" event so a timeout can say WHICH stage failed:
 * never loaded (sign-in / cookies / embedding blocked) vs loaded but never finished drawing.
 */
function openHiddenViz(embedUrl, size) {
  return new Promise((resolve, reject) => {
    const viz = /** @type {any} */ (document.createElement("tableau-viz"));
    viz.setAttribute("src", embedUrl);
    viz.setAttribute("toolbar", "hidden");
    viz.setAttribute("hide-tabs", "");
    viz.setAttribute("width", String(size.width));
    viz.setAttribute("height", String(size.height));
    let sizeKnown = false;
    viz.addEventListener("firstvizsizeknown", () => { sizeKnown = true; });
    // Signed-out embeds never report a size (Tableau shows its sign-in page instead).
    const signInTimer = setTimeout(() => {
      if (sizeKnown) return;
      clearTimeout(timer);
      viz.remove();
      const err = /** @type {any} */ (new Error("needs Tableau sign-in"));
      err.code = "SIGN_IN_NEEDED";
      reject(err);
    }, SIGNIN_DETECT_MS);
    const timer = setTimeout(() => {
      clearTimeout(signInTimer);
      viz.remove();
      const secs = VIZ_LOAD_TIMEOUT_MS / 1000;
      reject(new Error(sizeKnown
        ? `loaded but never finished drawing within ${secs}s (browser paused it?)`
        : `never loaded within ${secs}s — usually a Tableau sign-in/cookie problem, or embedding blocked for this site`));
    }, VIZ_LOAD_TIMEOUT_MS);
    viz.addEventListener("firstinteractive", () => { clearTimeout(timer); clearTimeout(signInTimer); resolve(viz); });
    viz.addEventListener("vizloaderror", e => {
      clearTimeout(timer);
      clearTimeout(signInTimer);
      viz.remove();
      let detail = e && e.detail;
      try { detail = JSON.stringify(detail); } catch (x) { /* keep as is */ }
      reject(new Error("Tableau refused to open it: " + detail));
    });
    hiddenVizHost().appendChild(viz);
  });
}

/* ── plain copies: only what the export uses, so nothing keeps a reference to a closed viz ──────── */
const plainValue = v => (v ? { value: v.value, nativeValue: v.nativeValue, formattedValue: v.formattedValue } : v);

/** a DataTable (summary data, selected marks): columns incl. fieldId (pies match measures by it) */
function plainSummary(t) {
  return {
    columns: (t.columns || []).map(c => ({ fieldName: c.fieldName, fieldId: c.fieldId, dataType: c.dataType, index: c.index,
                                           isReferenced: c.isReferenced })),
    data: (t.data || []).map(r => r.map(plainValue)),
  };
}
function plainFilter(f) {
  const fv = v => (v ? { formattedValue: v.formattedValue, value: v.value } : v);
  return { fieldName: f.fieldName, filterType: f.filterType,
           appliedValues: (f.appliedValues || []).map(fv), minValue: fv(f.minValue), maxValue: fv(f.maxValue) };
}
function plainParameter(p) {
  return { name: p.name, currentValue: p.currentValue ? { value: p.currentValue.value, formattedValue: p.currentValue.formattedValue } : null };
}
/** a DashboardObject: what the layout uses (the id is the zone id in the workbook – text boxes / images need it) */
function plainObject(o) {
  return { id: o.id, type: o.type, name: o.name, isVisible: o.isVisible !== false, isFloating: !!o.isFloating,
           position: o.position ? { x: o.position.x, y: o.position.y } : undefined,
           size: o.size ? { width: o.size.width, height: o.size.height } : undefined };
}

async function readWorksheetSummary(ws) {
  if (typeof ws.getSummaryDataReaderAsync === "function") {
    const reader = await ws.getSummaryDataReaderAsync(undefined, { ignoreSelection: true });
    try { return plainSummary(await reader.getAllPagesAsync()); }
    finally { try { await reader.releaseAsync(); } catch (e) { /* ignore */ } }
  }
  return plainSummary(await ws.getSummaryDataAsync({ ignoreSelection: true }));
}

/** a worksheet stand-in answering from what was read (same methods the export calls on a live worksheet) */
function snapshotSheet(name, s) {
  const answer = (value, error) => async () => { if (error) throw error; return value; };
  return {
    name,
    getSummaryDataAsync: answer(s.summary, s.summaryError),
    getVisualSpecificationAsync: answer(s.visualSpec, s.visualSpec ? null : (s.visualSpecError || new Error("getVisualSpecificationAsync is unavailable"))),
    getFiltersAsync: async () => s.filters || [],
    getSelectedMarksAsync: async () => ({ data: s.selected || [] })
  };
}

/** Reads one worksheet completely: data, visual specification, filters, selected marks.
 *  wrap(promise, what) bounds each call (live dashboard) or passes it through. */
async function readSheet(ws, wrap, plain) {
  const s = { summary: null, summaryError: null, visualSpec: null, visualSpecError: null, filters: [], selected: [] };
  try { s.summary = plain ? await wrap(readWorksheetSummary(ws), `Reading "${ws.name}"`)
                          : await wrap(ws.getSummaryDataAsync({ ignoreSelection: true }), `Reading "${ws.name}"`); }
  catch (e) { s.summaryError = e; }
  if (typeof ws.getVisualSpecificationAsync === "function") {
    try { s.visualSpec = await wrap(ws.getVisualSpecificationAsync(), `Visual specification of "${ws.name}"`); }
    catch (e) { s.visualSpecError = e; }
  }
  try { s.filters = ((await wrap(ws.getFiltersAsync(), `Filters of "${ws.name}"`)) || []).map(f => plain ? plainFilter(f) : f); }
  catch (e) { /* export works without filter values */ }
  if (typeof ws.getSelectedMarksAsync === "function") {
    try {
      const marks = await wrap(ws.getSelectedMarksAsync(), `Selected marks of "${ws.name}"`);
      s.selected = ((marks && marks.data) || []).map(t => plain ? plainSummary(t) : t);
    } catch (e) { /* no selection outline */ }
  }
  return snapshotSheet(ws.name, s);
}

/** Read every worksheet of the dashboard shown in `viz` into the export's input shape. */
async function readEmbeddedDashboard(viz) {
  const sheet = viz.workbook.activeSheet;
  if (!sheet || sheet.sheetType !== "dashboard") throw new Error("opened view is not a dashboard");
  const pass = p => p;
  const sheets = [];
  for (const ws of sheet.worksheets || []) sheets.push(await readSheet(ws, pass, true));
  let params = [];
  try {
    const wb = viz.workbook;
    params = typeof wb.getParametersAsync === "function" ? ((await wb.getParametersAsync()) || []).map(plainParameter) : [];
  } catch (e) { /* no parameter controls exported */ }
  const objects = (sheet.objects || []).map(plainObject);
  return { dashboard: { name: sheet.name, objects, worksheets: sheets, getParametersAsync: async () => params }, sheets };
}

/**
 * Loading screen over the extension while other dashboards are read. The hidden vizzes sit
 * underneath it (they must stay on screen for Chrome to render them), so the user sees progress
 * instead of dashboards flashing by. The debug view and the sign-in panel sit above it.
 */
function showExportOverlay(title, detail) {
  if (!document.getElementById("tfx_loading_css")) {
    const css = document.createElement("style");
    css.id = "tfx_loading_css";
    css.textContent =
      "#tfx_loading{position:fixed;inset:0;z-index:9000;background:#fff;display:flex;flex-direction:column;" +
      "align-items:center;justify-content:center;gap:10px;padding:16px;text-align:center;" +
      "font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#1c2b26;}" +
      "#tfx_loading .tfx-spin{width:28px;height:28px;border:3px solid #dfe6e2;border-top-color:#1c7a4d;" +
      "border-radius:50%;animation:tfxspin .8s linear infinite;}" +
      "#tfx_loading .tfx-title{font-size:13px;font-weight:600;}" +
      "#tfx_loading .tfx-detail{font-size:11px;color:#5b6b65;max-width:260px;line-height:1.4;}" +
      "@keyframes tfxspin{to{transform:rotate(360deg)}}";
    document.head.appendChild(css);
  }
  let overlay = document.getElementById("tfx_loading");
  if (!overlay) {
    overlay = document.createElement("div");
    overlay.id = "tfx_loading";
    overlay.setAttribute("role", "status");
    overlay.setAttribute("aria-live", "polite");
    overlay.innerHTML = '<div class="tfx-spin"></div><div class="tfx-title"></div><div class="tfx-detail"></div>';
    document.body.appendChild(overlay);
  }
  overlay.querySelector(".tfx-title").textContent = title;
  overlay.querySelector(".tfx-detail").textContent = detail || "";
}

function hideExportOverlay() {
  const overlay = document.getElementById("tfx_loading");
  if (overlay) overlay.remove();
}

/** Rejects with a readable error if `promise` doesn't settle in time (nothing may hang the export). */
function withTimeout(promise, ms, what) {
  let t;
  return Promise.race([
    promise,
    new Promise((_, reject) => { t = setTimeout(() => reject(new Error(`${what} took longer than ${ms / 1000}s`)), ms); }),
  ]).finally(() => clearTimeout(t));
}

/**
 * Read the CURRENT dashboard's data, visual specifications, filters, selected marks and parameters up front,
 * through the Extensions API, BEFORE the Embedding API is loaded. Once embedded vizzes are connected inside
 * this frame, Extensions API calls can stop getting answers (both libraries talk to Tableau through
 * postMessage), which left the export stuck on "Building Excel…". Returns the same { dashboard, sheets }
 * shape, with each sheet answering from the snapshot.
 */
async function snapshotLiveDashboard(current) {
  const bounded = (p, what) => withTimeout(p, LIVE_READ_TIMEOUT_MS, what);
  const sheets = [];
  for (const ws of current.sheets) sheets.push(await readSheet(ws, bounded, false));
  let params = [];
  try {
    const d = current.dashboard;
    params = typeof d.getParametersAsync === "function" ? (await bounded(d.getParametersAsync(), "Parameters")) || [] : [];
  } catch (e) { /* no parameter controls exported */ }
  const d = current.dashboard;
  return { dashboard: { name: d.name, objects: d.objects, worksheets: sheets, getParametersAsync: async () => params }, sheets };
}

/**
 * One-time Tableau sign-in for embedded dashboards.
 * The extension's frame can't reuse the user's normal Tableau login, so Tableau shows its own
 * "Sign in to Tableau Cloud" button inside the embedded view. We show that view (with a short
 * explanation) and wait until it becomes interactive = signed in. Resolves true / false (skipped).
 */
function promptTableauSignIn(embedUrl) {
  return new Promise(resolve => {
    const overlay = document.createElement("div");
    overlay.id = "tfx_signin";
    overlay.style.cssText = "position:fixed;inset:0;z-index:10000;background:#fff;display:flex;flex-direction:column;";
    const bar = document.createElement("div");
    bar.style.cssText = "display:flex;gap:8px;align-items:center;padding:8px 10px;border-bottom:1px solid #dfe6e2;font-size:12px;line-height:1.4;";
    const msg = document.createElement("div");
    msg.style.flex = "1";
    msg.textContent = "To export the other dashboards, Tableau needs you to sign in once. Click \u201cSign in to Tableau Cloud\u201d below.";
    const skip = document.createElement("button");
    skip.className = "btn-load";
    skip.style.width = "auto";
    skip.textContent = "Skip";
    bar.appendChild(msg);
    bar.appendChild(skip);
    const box = document.createElement("div");
    box.style.cssText = "flex:1;min-height:0;overflow:auto;";
    overlay.appendChild(bar);
    overlay.appendChild(box);
    document.body.appendChild(overlay);

    const viz = document.createElement("tableau-viz");
    viz.setAttribute("src", embedUrl);
    viz.setAttribute("toolbar", "hidden");
    viz.setAttribute("hide-tabs", "");
    viz.setAttribute("width", String(Math.max(300, box.clientWidth || 0)));
    viz.setAttribute("height", String(Math.max(300, box.clientHeight || 0)));

    let finished = false;
    const finish = ok => {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      overlay.remove();
      resolve(ok);
    };
    const timer = setTimeout(() => finish(false), SIGNIN_WAIT_MS);
    viz.addEventListener("firstinteractive", () => finish(true));
    skip.addEventListener("click", () => finish(false));
    box.appendChild(viz);
  });
}

/**
 * All visible dashboards in tab order: the current one from the Extensions API (live
 * filters), the others via hidden vizzes. Never throws for a single dashboard — failures
 * become `notes`, and the export continues with whatever could be read.
 */
async function collectAllDashboards(current, onProgress) {
  const notes = [];
  const wb = loadedWorkbook();
  if (!wb || !wb.xml) return { targets: [current], notes: ["Other dashboards: load the workbook first."] };
  const order = dashboardTabOrder(wb.xml);
  if (!order.includes(current.dashboard.name)) order.unshift(current.dashboard.name);
  if (order.length === 1) return { targets: [current], notes };
  if (!wb.id) {
    return { targets: [current], notes: ["Other dashboards are only available when the workbook is auto-loaded from Tableau Cloud (not from a manual file)."] };
  }

  // Current dashboard first — before any embedded viz exists in this frame.
  current = await snapshotLiveDashboard(current);

  let views, libUrl;
  try {
    const r = await (await backendFetch(`/workbooks/${encodeURIComponent(wb.id)}/views`)).json();
    views = r.views || [];
    libUrl = r.embedLibUrl;
    if (!libUrl || !views.every(v => v.embedUrl)) throw new Error("backend is outdated — update worker.js");
    await loadEmbeddingApi(libUrl);
  } catch (e) {
    return { targets: [current], notes: [`Other dashboards skipped: ${e.message}`] };
  }

  const targets = [];
  const others = order.filter(n => n !== current.dashboard.name);
  let done = 0, abort = null;
  for (const name of order) {
    if (name === current.dashboard.name) { targets.push(current); continue; }
    done++;
    if (abort) { notes.push(`"${name}": skipped (${abort})`); continue; }
    const view = views.find(v => v.name === name);
    if (!view) { notes.push(`"${name}": not published as a tab — skipped`); continue; }
    if (onProgress) onProgress(name, done, others.length);
    let viz = null;
    try {
      const size = dashboardPixelSize(wb.xml, name);
      try {
        viz = await openHiddenViz(view.embedUrl, size);
      } catch (e) {
        if (e.code !== "SIGN_IN_NEEDED") throw e;
        if (onProgress) onProgress(name, done, others.length, "signin");
        if (!(await promptTableauSignIn(view.embedUrl))) throw new Error("Tableau sign-in was skipped");
        if (onProgress) onProgress(name, done, others.length);
        viz = await openHiddenViz(view.embedUrl, size);       // signed in now → open hidden again
      }
      targets.push(await readEmbeddedDashboard(viz));
    } catch (e) {
      notes.push(`"${name}": ${e.message}`);
      if (/sign-in|refused|Embedding API|never finished drawing/i.test(e.message)) abort = "same problem as the previous dashboard";
    } finally {
      if (viz) viz.remove();
    }
  }
  removeHiddenVizHost();
  clearDebugVizSwitch();
  return { targets, notes };
}

/** Remove the hidden-viz host once reading is done, so nothing (debug view included) covers the extension. */
function removeHiddenVizHost() {
  const host = document.getElementById("tfx_viz_host");
  if (host) host.remove();
}

/* =============================================================================
 * uniqueExcelSheetName() — a valid, unique Excel sheet name for a dashboard.
 * Excel rules: max 31 chars; none of \ / ? * [ ] : ; not empty; must not start or
 * end with an apostrophe; "History" is reserved; names are unique ignoring case.
 * usedLower: Set of lower-cased names already taken (updated here).
 * ============================================================================= */
function uniqueExcelSheetName(dashboardName, usedLower) {
  let base = String(dashboardName || "Dashboard Export")
    .replace(/[\\\/\*\?\[\]:]/g, "")
    .replace(/^'+|'+$/g, "")
    .slice(0, 31)
    .replace(/^'+|'+$/g, "");
  if (!base.trim()) base = "Dashboard Export";
  if (base.toLowerCase() === "history") base = "History (1)";
  let name = base, n = 2;
  while (usedLower.has(name.toLowerCase())) {
    const suffix = ` (${n++})`;
    name = base.slice(0, 31 - suffix.length).replace(/[\s']+$/, "") + suffix;
  }
  usedLower.add(name.toLowerCase());
  return name;
}


/* ═══ export/export.js ════════════════════════════════════════════════════════════════════════════ */
/* The export: Tableau dashboard → formatted XLSX with native charts. */
const makeExportFileName = exportFileName;

/* the longest a Tableau-rendered picture may take (it can stall once embedded dashboards were opened) */
const IMAGE_RENDER_TIMEOUT_MS = 30000;

/* =========================================================================
 * exportToExcel() — one click → ONE .xlsx, ONE sheet per dashboard.
 *   Tableau Cloud with "All dashboards" ticked: every visible dashboard in tab order (the others are read
 *   through hidden embedded views – cloud/dashboards.js); otherwise only the dashboard the extension sits on.
 *   Each dashboard is written by writeDashboardSheet() (tables, KPI cards, text, images, native charts);
 *   the charts are injected sheet by sheet after ExcelJS has written the file; one Conversion Report covers all.
 * ========================================================================= */
async function exportToExcel() {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  const sheets = dashboard.worksheets;
  const btn = /** @type {HTMLButtonElement} */ (document.getElementById("export_button"));
  const btnText = btn.textContent;
  btn.disabled = true;
  /** @type {string[]} dashboards left out, and why */
  let notes = [];

  try {
    if (!(await ensureFormatModel()) && !window.confirm(
        "No Tableau workbook is loaded for this dashboard.\n\n" +
        "Without it the extension cannot tell which sheets are charts, so they will be exported as tables.\n\n" +
        "Cancel, then click 📁 Load Workbook to export real charts – or OK to export tables anyway.")) {
      setExportStatus("Export stopped – click 📁 Load Workbook first so charts are exported as charts", true);
      return;
    }
    const wantAll = wantsAllDashboards();
    // "<workbook> - <dashboard>.xlsx" for one dashboard, "<workbook>.xlsx" when every dashboard can be read
    // (the workbook auto-loaded from Tableau Cloud – a manual file exports only this dashboard)
    const wbFile = formatModelFileName();
    const allName = wantAll && wbFile && loadedWorkbook().id ? String(wbFile).replace(/\.twbx?$/i, "") : null;
    const exportFileName = makeExportFileName(wbFile, allName || dashboard.name);
    setExportStatus("Choose where to save the Excel file…");
    const saveTarget = await chooseSaveTarget(exportFileName);       // opens in the loaded workbook's folder
    if (saveTarget === "cancelled") { setExportStatus("Export cancelled – no file was saved"); return; }

    const titleMap = getTitleMap();
    console.log(`[Export] Using ${Object.keys(titleMap).length} titles`);
    const fmtModel = await ensureFormatModel();
    console.log(`[Export] Format model: ${fmtModel ? Object.keys(fmtModel.sheets).length + " sheets" : "none – load the workbook for exact formatting"}`);

    // ── which dashboards: the current one, or every visible one (Tableau Cloud) ──
    const current = { dashboard, sheets };
    let targets = [current];
    if (wantAll) {
      const collected = await collectAllDashboards(current, (name, i, n, stage) => {
        if (stage === "signin") {
          btn.textContent = "🔐 Waiting for Tableau sign-in…";
          hideExportOverlay();                       // the sign-in panel needs to be seen
          return;
        }
        btn.textContent = `⏳ Reading "${name}" (${i}/${n})…`;
        showExportOverlay(`Reading dashboard ${i} of ${n}`, `"${name}" — opening it in the background to read its tables and charts.`);
      });
      targets = collected.targets;
      notes = collected.notes || [];
    }
    btn.textContent = "⏳ Building Excel…";
    if (targets.length > 1) showExportOverlay("Building Excel…", `${targets.length} dashboards, one sheet each.`);

    // ── one sheet per dashboard ──
    const workbook = new ExcelJS.Workbook();
    // the helper sheets' names are taken: a dashboard called "Chart Data" becomes "Chart Data (2)"
    const usedSheetNames = new Set([CHART_DATA_SHEET, DATA_SHEET, REPORT_SHEET].map(s => s.toLowerCase()));
    const chartData = { sheet: null, row: 0 };                  // one hidden "Chart Data" sheet for every dashboard
    const written = [];
    for (const target of targets) {
      const sheetName = uniqueExcelSheetName(target.dashboard.name, usedSheetNames);
      const result = await writeDashboardSheet(workbook, sheetName, target, { titleMap, fmtModel, chartData });
      if (result) written.push(result);
      else {
        usedSheetNames.delete(sheetName.toLowerCase());
        console.log(`[Export] "${target.dashboard.name}" has no exportable data – no sheet added`);
        if (targets.length > 1) notes.push(`"${target.dashboard.name}": no data in any visible worksheet — no sheet added`);
      }
    }
    if (!written.length) throw new Error("No data found in any visible worksheet.");

    if (FORMAT_CONFIG.conversionReport) writeConversionReports(workbook, written.map(w => w.report));

    setExportStatus("Building the Excel file…");
    /** @type {any} the XLSX bytes: ExcelJS's buffer, then injectCharts' Uint8Array */
    let buffer = await workbook.xlsx.writeBuffer();
    // fit-to-page next to collapsed row groups: ExcelJS writes <sheetPr> out of order – put it right
    if (written.some(w => w.worksheet.pageSetup.fitToPage && w.worksheet.properties.outlineProperties)) buffer = await fixSheetProperties(buffer);
    // native charts, sheet by sheet: sized to their blocks now that the final column widths / row heights are known
    for (const w of written) {
      const jobs = w.finalizeCharts();
      if (!jobs.length) continue;
      try {
        buffer = await ExcelChartWriter.injectCharts(buffer, { sheetName: w.worksheet.name, charts: jobs });
        console.log(`[Charts] ${w.worksheet.name}: ${jobs.length} native Excel chart(s) added`);
      } catch (err) {
        console.error(`[Charts] ${w.worksheet.name}: could not add native charts`, err);
        jobs.forEach(job => {
          const status = w.visualStatuses.find(entry => entry.worksheet === job.item.name);
          if (status) { status.status = "warning"; status.reason = `chart injection failed: ${err.message}`; }
        });
      }
    }

    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    if (saveTarget) {
      const writable = await saveTarget.createWritable();
      await writable.write(blob);
      await writable.close();
      console.log(`[Export] saved to ${saveTarget.name}`);
    } else {
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = exportFileName;
      link.click();
      URL.revokeObjectURL(link.href);
    }
    // the panel's summary: every visual of every dashboard ("Dashboard / Sheet" when there are several), plus
    // where the file went and what was left out
    const multi = written.length > 1;
    const visualStatuses = written.flatMap(w => multi
      ? w.visualStatuses.map(s => ({ ...s, worksheet: `${w.dashboardName} / ${s.worksheet}` })) : w.visualStatuses);
    const workbookWarning = [...new Set(written.map(w => w.workbookWarning).filter(Boolean))].join(" · ");
    updateVisualStatus(visualStatuses, workbookWarning);
    appendExportStatus((multi ? `${written.length} dashboards, one sheet each – ` : "") +
      (saveTarget ? `saved as ${saveTarget.name}` : `downloaded as ${exportFileName}`));
    appendExportNotes(notes);

    console.log(`✅ Export completed with Tableau formatting and native charts — ${written.length} sheet(s): ${written.map(w => w.worksheet.name).join(", ")}`);

  } catch (err) {
    console.error("[Export]", err);
    setExportStatus("Export failed: " + err.message, true);
    appendExportNotes(notes);
    alert("Export failed. Check console (F12) for details.\n\n" + err.message);
  } finally {
    removeHiddenVizHost();
    hideExportOverlay();
    btn.textContent = btnText;
    btn.disabled = false;
  }
}

/* "All dashboards" ticked (default) on Tableau Cloud / Server, with the workbook loaded and more than one
 * dashboard in it → every visible dashboard. Unticked or Tableau Desktop: only this one. (A manual file has no
 * Tableau Cloud id: collectAllDashboards then exports this one and says why.) */
function wantsAllDashboards() {
  const allBox = /** @type {HTMLInputElement} */ (document.getElementById("export_all"));
  const env = tableau.extensions.environment || {};
  if ((allBox && !allBox.checked) || env.context === "desktop") return false;
  const wb = loadedWorkbook();
  return !!(wb && wb.xml && dashboardTabOrder(wb.xml).length > 1);
}

/* =========================================================================
 * writeDashboardSheet() — ONE dashboard onto ONE new Excel sheet: tables, KPI cards, text boxes, images,
 * filters / parameters and native charts at their dashboard positions.
 * target: { dashboard: { name, objects, worksheets, getParametersAsync }, sheets: Tableau worksheets (live, or
 * snapshots of another dashboard) }; ctx: { titleMap, fmtModel, chartData: the shared hidden data sheet }.
 * Returns null (and adds no sheet) when there is nothing to export, else
 * { worksheet, dashboardName, visualStatuses, workbookWarning, report, finalizeCharts() → ChartJob[] }
 * (finalizeCharts runs after workbook.xlsx.writeBuffer(), with the sheet's final column widths / row heights).
 * ========================================================================= */
async function writeDashboardSheet(workbook, sheetName, target, ctx) {
    const dashboard = target.dashboard;
    const sheets = target.sheets;
    const { titleMap, fmtModel, chartData } = ctx;
    const filterValuesMap = await extractFilterValuesPerField(sheets);
    console.log(`📊 [${dashboard.name}] Filter values per field:`, filterValuesMap);
    const workbookWarning = fmtModel && isFormatModelStale(fmtModel)
      ? "⚠ The workbook was loaded with an older version of the extension – click 📁 Load Workbook once more for the latest formatting"
      : fmtModel ? describeWorkbookMatch(checkWorkbookMatch(fmtModel, dashboard))
      : "⚠ Workbook not loaded – charts can only be recognised from the workbook, so they may be exported as tables. Click 📁 Load Workbook, then export again";
    if (workbookWarning) console.warn("[Export]", workbookWarning);

    const layoutMap = buildLayoutMap(dashboard.objects || [], titleMap);

    // ── 1. Build DZV map ── (a sheet squeezed into a 1 px zone is a common way to hide it: not shown either)
    const dzvMap = {};
    (dashboard.objects || [])
      .filter(obj => obj.type === "worksheet")
      .forEach(obj => {
        const tiny = obj.size && (obj.size.width < 8 || obj.size.height < 8);
        dzvMap[obj.name] = obj.isVisible !== false && !tiny;
      });
    console.log("[DZV] Visibility map:", dzvMap);

    // ── 2. Fetch all sheets in parallel ──
    setExportStatus(`Reading ${sheets.length} worksheets of "${dashboard.name}" from Tableau…`);
    const allSheetsData = await fetchAllSheetsData(sheets);

    /** @type {{ name: string, shape: string, layout: any }[]} button / icon sheets, drawn as their icon */
    const iconSheets = [];
    /** @type {ExportItem[]} */
    const filterValueItems = [];
    /** @type {ExportItem[]} */
    const dataWorksheetItems = [];
    const visualStatuses = [];
    /** sheets drawn inside a donut's hole (its centre text): not exported as blocks of their own */
    const pieCentreSheets = new Set();
    /** @type {ExportItem[]} colour legends of pies, drawn at their dashboard position */
    const pieLegendItems = [];
    /** @type {{ name: string, reason: string }[]} worksheets not exported, and why (conversion report) */
    const skipped = [];
    /* A pie's colour legends shown on the dashboard (the chart has none of its own): the TWB legend zones of the
       sheet, matched to each ring by the field they show, laid out at their dashboard position. TWB zones use
       0–100000 units: they are mapped to API pixels through the sheet's own zone. */
    /** @returns {ExportItem[]} */
    const pieLegends = (pie, sheetName, vm) => {
      const dashM = fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name];
      const sheetZone = dashM && dashM.zones.find(z => z.type === "worksheet" && z.name === sheetName && z.w && z.h);
      const api = (dashboard.objects || []).find(o => o.type === "worksheet" && o.name === sheetName && o.size);
      if (!sheetZone || !api) return [];
      const zones = dashM.zones.filter(z => z.type === "color" && z.name === sheetName && !z.hidden);
      const sx = api.size.width / sheetZone.w, sy = api.size.height / sheetZone.h;
      return [pie, pie.inner].filter(s => s && s.legend).map((s, i) => {
        const z = zones.find(z => s.colorRef && z.param && tfSameField(tfParseFieldRef(z.param), s.colorRef)) || (i === 0 && zones.length === 1 ? zones[0] : null);
        if (!z) return null;
        const key = `legend:${z.id}`;
        const obj = { type: "worksheet", name: key, id: z.id, isVisible: true,
                      position: { x: api.position.x + (z.x - sheetZone.x) * sx, y: api.position.y + (z.y - sheetZone.y) * sy },
                      size: { width: z.w * sx, height: z.h * sy } };
        const layout = buildLayoutMap([...(dashboard.objects || []), obj], titleMap).get(key);
        return layout ? { type: /** @type {"legend"} */ ("legend"), name: key, visualName: s.legend.title, layout, legend: s.legend, fmt: vm.fmt, columns: [],
                          rowCount: 1 + s.legend.items.length } : null;
      }).filter(Boolean);
    };

    for (const { sheet, data, visualSpec, visualSpecError, error } of allSheetsData) {
      if (error) {
        console.warn(`Skipping "${sheet.name}": ${error.message}`);
        skipped.push({ name: sheet.name, reason: `FAILED – Tableau did not return its data: ${error.message}` });
        continue;
      }

      // ---- DZV check ----
      if (dzvMap[sheet.name] === false) {
        console.log(`[DZV] Skipping hidden sheet: "${sheet.name}"`);
        skipped.push({ name: sheet.name, reason: "hidden on the dashboard (not visible, or squeezed into a tiny zone)" });
        continue;
      }

      let summaryData = data;  // already fetched
      if (!summaryData.columns || summaryData.columns.length === 0) {
        console.warn(`[Export] Skipping "${sheet.name}": summary data has no columns`);
        skipped.push({ name: sheet.name, reason: "no data: Tableau returned no columns for it" });
        continue;
      }
      if (!summaryData.data || summaryData.data.length === 0) {
        console.warn(`[Export] Skipping "${sheet.name}": summary data has no rows`);
        skipped.push({ name: sheet.name, reason: "no data: no marks with the current filters" });
        continue;
      }

      const layout = layoutMap.get(sheet.name);
      const visualName = (layout && layout.displayName) ? layout.displayName : sheet.name;

      // a button / icon sheet (a custom shape on empty shelves): drawn as its icon, not as a table of its flags
      const icon = fmtModel ? tvIconSheet(fmtModel, sheet.name, summaryData) : null;
      if (icon) {
        if (icon.shape && layout) iconSheets.push({ name: sheet.name, shape: icon.shape, layout });
        else {
          console.log(`[Icons] ${sheet.name}: a Tableau shape – not exported`);
          skipped.push({ name: sheet.name, reason: "a button / icon drawn with one of Tableau's built-in shapes" });
        }
        continue;
      }

      if (isFilterValueWorksheet(sheet.name, summaryData)) {
        let matchedFilterName = null;
        let matchedValues = null;

        for (const [filterField, values] of Object.entries(filterValuesMap)) {
          const cleanFilterField = filterField.toLowerCase().replace(/[^a-z]/g, '');
          const cleanSheetName = sheet.name.toLowerCase().replace(/[^a-z]/g, '');
          if (cleanSheetName.includes(cleanFilterField) || cleanFilterField.includes(cleanSheetName)) {
            matchedFilterName = filterField;
            matchedValues = values;
            break;
          }
        }

        if (!matchedValues) {
          matchedFilterName = visualName;
          matchedValues = [`Values: ${summaryData.data.length} items`];
        }

        filterValueItems.push({
          type: "filterValue",
          name: sheet.name,
          visualName: visualName,
          filterName: matchedFilterName,
          values: matchedValues,
          layout: layout,
          rowCount: 1 + (Array.isArray(matchedValues) ? matchedValues.length : 1),
          originalData: summaryData
        });
      } else {
        // ── NEW: rebuild the visual table (pivot, merge, sort, visible columns, title) ──
        const visualModel = buildVisualModel(fmtModel, sheet.name, summaryData,
          { dashboardName: dashboard.name, displayName: visualName, visualSpec, visualSpecError, layout });
        console.log(`[VisualSpec] ${sheet.name}:`, {
          available: visualModel.metadata.visualSpecAvailable,
          keys: visualModel.metadata.visualSpecKeys,
          type: visualModel.type,
          marks: visualModel.metadata.markTokens,
          markSource: visualModel.metadata.markSource,
          error: visualModel.metadata.visualSpecError,
          spec: visualModel.source.visualSpec
        });
        let renderDecision = chooseVisualRenderer(visualModel);
        // a pie or donut (one or two layers): drawn the way Tableau draws it – slices, labels, hole, tooltips
        if (fmtModel && FORMAT_CONFIG.nativeCharts && visualModel.viewModel.fmt.pieLayers()) {
          const pie = buildPieModel(fmtModel, sheet.name, summaryData, {
            dashboardName: dashboard.name, selectedTables: await fetchSelectedMarks(sheet),
            dataOf: name => (allSheetsData.find(x => x.sheet.name === name) || { data: null }).data
          });
          if (pie) {
            visualModel.pie = pie;
            visualModel.type = VISUAL_TYPES.PIE;
            renderDecision = VISUAL_RENDERERS.chart;
            if (pie.centerText) {                              // the text drawn in the hole is not exported again
              if (pie.centerText.name) pieCentreSheets.add(pie.centerText.name);
              else visualModel.viewModel.headerZoneIds.push(pie.centerText.id);
            }
            pieLegendItems.push(...pieLegends(pie, sheet.name, visualModel.viewModel));
          }
        }
        if (renderDecision.renderer === "excel-chart" && !visualModel.pie) {
          try {
            visualModel.chartSpecs = buildExcelChartSpecs(visualModel, fmtModel);
            // Tableau draws legends as separate dashboard cards, never inside the view: the chart keeps a
            // legend only when the dashboard shows a colour / size / shape legend for this sheet
            const dashZones = fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name] ? fmtModel.dashboards[dashboard.name].zones : null;
            if (dashZones && !dashZones.some(z => /^(color|size|shape)$/.test(z.type) && z.name === sheet.name && !z.hidden)) {
              visualModel.chartSpecs.forEach(s => { s.legend = false; });
            }
          } catch (err) {
            console.warn(`[Visual] ${sheet.name}: native chart not possible – ${err.message}`);
            renderDecision = imageOrFallback(visualModel, `native chart not possible (${err.message})`);
          }
        }
        visualModel.renderer = renderDecision.renderer;
        visualModel.status = renderDecision.status;
        visualModel.statusReason = renderDecision.reason;
        visualStatuses.push({
          worksheet: sheet.name,
          type: visualModel.type,
          renderer: renderDecision.renderer,
          status: renderDecision.status,
          reason: renderDecision.reason,
          visualSpecAvailable: visualModel.metadata.visualSpecAvailable,
          imageApiAvailable: visualModel.metadata.imageApiAvailable
        });
        console.log(`[Visual] ${sheet.name}: ${visualModel.type} -> ${renderDecision.renderer} (${renderDecision.reason})`);

        if (renderDecision.renderer === "data-fallback" && visualModel.viewModel.kind === "chart" &&
            FORMAT_CONFIG.chartPolicy === "skip") {
          console.warn(`[Export] "${sheet.name}" is a chart - skipped by FORMAT_CONFIG.chartPolicy`);
          skipped.push({ name: sheet.name, reason: `UNSUPPORTED – ${renderDecision.reason} (chartPolicy "skip")` });
          continue;
        }
        let vm = visualModel.viewModel;
        if (!vm.rows.length || !vm.order.length) {
          console.warn(`[Export] "${sheet.name}" produced an empty formatted view; rebuilding from live summary data`);
          const rawVisualModel = buildVisualModel(null, sheet.name, summaryData,
            { dashboardName: dashboard.name, displayName: visualName, visualSpec, visualSpecError, layout });
          if (rawVisualModel.viewModel.rows.length && rawVisualModel.viewModel.order.length) {
            vm = rawVisualModel.viewModel;
            visualModel.viewModel = vm;
            visualModel.data.rows = vm.rows;
            visualModel.data.columns = vm.cols;
            visualModel.data.dimensions = vm.order.filter(i => vm.cols[i].isHeader).map(i => vm.cols[i]);
            visualModel.data.measures = vm.order.filter(i => !vm.cols[i].isHeader).map(i => vm.cols[i]);
            visualModel.statusReason = "TWB view was empty; exported live summary data";
          }
        }
        if (!vm.rows.length || !vm.order.length) {
          console.warn(`[Export] Skipping "${sheet.name}": no renderable rows or columns after fallback`);
          skipped.push({ name: sheet.name, reason: "FAILED – no rows or columns left to write" });
          continue;
        }
        // a chart exported as data – or a table whose cells would drop what its marks show: axis/colour/detail fields
        // are the data – show every column, except the latitude / longitude Tableau generates only to place a map's marks
        if ((visualModel.renderer === "data-fallback" && !VISUAL_RENDERERS.cellTypes.has(visualModel.type)) || vm.lossless === false) {
          vm.order = vm.cols.map((_, i) => i).filter(i => !/^(latitude|longitude) \(generated\)$/i.test(vm.cols[i].name));
          vm.cols.forEach(c => { if (!c.label) c.label = c.name; });
          vm.showHeaderRow = true;
        }
        const isGraphic = visualModel.renderer === "excel-chart" || visualModel.renderer === "tableau-image";
        const box = isGraphic ? graphicBox(layout, vm) : null;
        const isKPI = !isGraphic && isKPIViewModel(vm);
        // KPI tile: rebuilt from its Tableau label and filling its zone (null = drawn as a label | value table)
        const kpiCard = isKPI ? buildKpiCard(vm, sheet.name, layout) : null;
        const kpiW = kpiCard && layout && layout.widthPx ? Math.max(kpiCard.tiles.length, layout.gridRight - layout.gridCol) : null;
        dataWorksheetItems.push({
          name: sheet.name,
          visualName: vm.title.text,
          layout,
          isKPI,
          kpiCard,
          type: "worksheet",
          vm,
          visualModel,
          box,                                                 // chart/image: pixel size from the dashboard zone
          fixedGridW: box ? box.gridW : kpiW,
          columns: vm.order,                                   // only its length is used for layout
          rowCount: box ? box.rows : kpiCard ? kpiCard.rows : isKPI ? (vm.showTitle ? 1 : 0) + vm.order.length : viewModelHeight(vm)
        });
      }
    }

    // ── parameter controls (e.g. Start Date / End Date): not filters, so read them separately ──
    /** @type {{ name: string, value: string }[]} every parameter's value at export time (conversion report) */
    const reportParams = [];
    try {
      const params = typeof dashboard.getParametersAsync === "function" ? await dashboard.getParametersAsync() : [];
      const controls = (dashboard.objects || []).filter(o => OBJECT_KIND[o.type] === "parameter" && o.isVisible !== false);
      const used = new Set();
      const norm = s => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
      params.forEach(param => {
        const raw = param.currentValue ? (param.currentValue.formattedValue ?? String(param.currentValue.value)) : "";
        const value = String(raw).replace(/\s+12:00:00\s*AM$|\s+00:00:00$/i, "");   // dates: no midnight time
        reportParams.push({ name: param.name, value });
        let control = controls.find(o => !used.has(o) && norm(o.name) === norm(param.name)) ||
                      controls.find(o => !used.has(o) && norm(o.name).includes(norm(param.name)));
        if (!control) return;                                      // parameter not shown on this dashboard
        used.add(control);
        filterValueItems.push({
          type: "filterValue", isParameter: true, name: param.name, visualName: param.name,
          filterName: param.name, values: [value], layout: layoutMap.get(control.name), rowCount: 2
        });
      });
      if (params.length) console.log(`[Parameters] ${used.size} of ${params.length} parameters have a control on the dashboard`);
    } catch (err) {
      console.warn("[Parameters] could not read parameters:", err.message);
    }

    // Tableau's number + trend tile: a chart in the zone right under / over a KPI card, same position and
    // width, takes the card's columns so the two line up as one card
    const near = (a, b) => Math.abs(a - b) <= 8;
    dataWorksheetItems.filter(it => it.kpiCard && it.fixedGridW && it.layout).forEach(card => {
      const c = card.layout;
      dataWorksheetItems.forEach(other => {
        const o = other.layout;
        if (other === card || !other.box || !o || !o.widthPx || !near(o.xPx, c.xPx) || !near(o.widthPx, c.widthPx)) return;
        if (near(o.yPx, c.yPx + c.heightPx) || near(c.yPx, o.yPx + o.heightPx)) { other.fixedGridW = card.fixedGridW; other.pairedCard = true; }
      });
    });

    // ── dashboard text boxes (banners, titles, notes): drawn where Tableau puts them, as Tableau formats them
    const dashZones = (fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name] || { zones: [] }).zones;
    const headerZoneIds = new Set(dataWorksheetItems.flatMap(it => it.vm ? it.vm.headerZoneIds : []));   // already table headers
    /** @type {ExportItem[]} */
    const textItems = [];
    (dashboard.objects || []).filter(o => OBJECT_KIND[o.type] === "text" && o.isVisible !== false).forEach(o => {
      const zone = dashZones.find(z => z.type === "text" && String(z.id) === String(o.id));
      if (!zone || zone.hidden || headerZoneIds.has(zone.id)) return;
      const layout = layoutMap.get(`text:${o.id}`);
      const card = buildTextCard(zone, layout, fmtModel);
      if (!card) return;
      textItems.push({ type: "text", name: `text:${o.id}`, visualName: tfZoneText(zone), layout, textCard: card,
        rowCount: card.rows, columns: [], fixedGridW: layout && layout.widthPx ? Math.max(1, layout.gridRight - layout.gridCol) : null });
    });
    if (textItems.length) console.log(`[Text] ${textItems.length} dashboard text box(es)`);

    // ── dashboard images (logos, icons) from the files packaged in the .twbx: a logo in its own zone is
    //    a block of the layout, an icon over a sheet floats on that sheet's block, backgrounds are left out
    const imageKinds = classifyImageObjects(dashboard.objects || []);
    const imageObjects = (dashboard.objects || []).filter(o => o.type === "image" && o.isVisible !== false && imageKinds.has(String(o.id)));
    const workbookImages = imageObjects.length || iconSheets.length ? await getWorkbookImages() : {};
    /** @type {ExportItem[]} */
    const imageItems = [];
    /** @type {(ImageBlock & { host: string, xPx: number, yPx: number })[]} */
    const overlayImages = [];
    imageObjects.forEach(o => {
      const cls = imageKinds.get(String(o.id));
      const zone = dashZones.find(z => z.type === "bitmap" && String(z.id) === String(o.id));
      if (!zone || !zone.param || zone.hidden) return;
      if (cls.kind === "backdrop" || cls.kind === "tiny") { console.log(`[Images] ${zone.param}: ${cls.kind === "tiny" ? "divider" : "background"} – not exported`); return; }
      const file = workbookImages[zone.param];
      if (!file) { console.warn(`[Images] ${zone.param} is not in the loaded workbook – load the .twbx to export it`); return; }
      const info = imageInfo(file.data, zone.param);
      if (!info.type) { console.warn(`[Images] ${zone.param}: not a PNG, JPEG, GIF or SVG file`); return; }
      /** @type {ImageBlock} */
      const image = { file, info, zone, url: webLink(zone.url), widthPx: o.size.width, heightPx: o.size.height };
      if (cls.kind === "overlay") { overlayImages.push({ ...image, host: cls.host, xPx: o.position.x, yPx: o.position.y }); return; }
      const layout = layoutMap.get(`image:${o.id}`);
      if (!layout) return;
      imageItems.push({ type: "image", name: `image:${o.id}`, visualName: zone.param, layout, image, columns: [],
        rowCount: Math.max(1, Math.round(o.size.height / EXCEL_ROW_PX)),
        fixedGridW: layout.widthPx ? Math.max(1, layout.gridRight - layout.gridCol) : null });
    });
    // icon sheets: the shape centred in the sheet's zone at about Tableau's mark size
    iconSheets.forEach(({ name, shape, layout }) => {
      const file = workbookImages["shape:" + shape];
      if (!file) { console.warn(`[Icons] ${name}: shape "${shape}" not in the loaded workbook`); return; }
      const info = imageInfo(file.data, shape);
      const w = layout.widthPx || 40, h = layout.heightPx || 40;
      const box = Math.min(28, Math.max(12, Math.round(Math.min(w, h) * 0.5)));
      if (!info.type) return;
      imageItems.push({ type: "image", name, visualName: name, layout, columns: [],
        image: { file, info, zone: { param: shape, scaled: true, centered: true }, url: null, widthPx: box, heightPx: box },
        rowCount: Math.max(1, Math.round(h / EXCEL_ROW_PX)),
        fixedGridW: layout.widthPx ? Math.max(1, layout.gridRight - layout.gridCol) : null });
    });
    if (imageItems.length || overlayImages.length) console.log(`[Images] ${imageItems.length + overlayImages.length} image(s)`);

    const allItems = [...filterValueItems, ...dataWorksheetItems.filter(it => !pieCentreSheets.has(it.name)),
                      ...textItems, ...imageItems, ...pieLegendItems];

    if (allItems.length === 0) {
      return null;   // nothing exportable on this dashboard → no sheet for it
    }

    console.table(visualStatuses);

    // columns from the dashboard's zone edges: every block keeps its Tableau position and width; a table gets
    // a column per field inside its zone, split by how much its fields hold
    allItems.forEach(item => {
      if (item.vm && item.type === "worksheet" && !item.box && !item.kpiCard && !item.isKPI) {
        item.splitPx = tableColumnPx(item.vm);                 // the widths Tableau draws the table at
      } else if (item.isKPI && !item.kpiCard) item.split = [1, 1];
      else if (item.kpiCard && item.kpiCard.tiles.length > 1) item.split = item.kpiCard.tiles.map(() => 1);
    });
    const grid = buildColumnGrid(allItems);
    const placedItems = allItems.map((item, idx) => {
      const span = grid.span(item);
      if (item.layout && span) {
        // a table covers every column it writes (it may run past its zone into free space)
        const fieldsEnd = span.fields && span.fields.length ? span.fields[span.fields.length - 1].offset + span.fields[span.fields.length - 1].span : 0;
        const gridW = Math.max(span.gridW, fieldsEnd);
        return { ...item, gridRow: item.layout.gridRow, gridCol: span.gridCol, gridW, allocatedRows: item.rowCount,
                 splitCuts: span.cuts, fieldCols: span.fields };
      }
      if (item.layout) {
        const l = item.layout;
        let calculatedWidth = l.gridW;

        if (item.fixedGridW) {
          calculatedWidth = item.fixedGridW;
        } else if (item.type === "filterValue") {
          calculatedWidth = Math.max(l.gridW, 3);
        } else if (item.isKPI) {
          calculatedWidth = Math.max(l.gridW, Math.min(6, item.columns.length + 1));
        } else {
          calculatedWidth = Math.max(l.gridW, Math.min(25, item.columns.length + 1));
        }

        return {
          ...item,
          gridRow: l.gridRow,
          gridCol: l.gridCol,
          gridW: calculatedWidth,
          allocatedRows: item.rowCount,
        };
      } else {
        const itemsPerRow = 3;
        const rowIdx = Math.floor(idx / itemsPerRow);
        const colIdx = idx % itemsPerRow;
        const gridRow = rowIdx * (Math.max(8, item.rowCount) + ROW_GAP);
        const gridCol = colIdx * 12;
        const gridW = item.fixedGridW || ((item.type === "filterValue") ? 4 : (item.isKPI ? 5 : 12));

        return {
          ...item,
          gridRow: gridRow,
          gridCol: gridCol,
          gridW: gridW,
          allocatedRows: item.rowCount,
        };
      }
    });

    resolveCollisions(placedItems);

// SNAP-OUT-OF-GROUPED-ROWS PASS
// Excel hides entire physical rows, not per-column cells, so a long table's collapsed rows would
// also hide whatever sits beside it. Charts, images, cards and filter lists that share rows with
// the table keep those rows visible (setTableVisibleRows); a table that would start inside another
// table's collapsed rows is pushed below that table's full physical row range.
let snapChanged = true;
const MAX_SNAP_PASSES = 10;
let snapPass = 0;

while (FORMAT_CONFIG.groupOverflowRows && snapChanged && snapPass < MAX_SNAP_PASSES) {
  snapChanged = false;
  snapPass++;
  setTableVisibleRows(placedItems);               // charts / cards beside a long table keep its rows visible

  for (const item of placedItems) {
    for (const other of placedItems) {
      if (item === other) continue;
      if (!other.visibleRows || other.visibleRows >= other.vm.rows.length) continue;  // other hides no rows

      const otherTotal = other.allocatedRows || other.rowCount || 0;
      const otherVisibleEnd = tableDataStart(other) + other.visibleRows;               // first hidden row
      const otherPhysicalEnd = other.gridRow + otherTotal + ROW_GAP;

      if (item.gridRow >= otherVisibleEnd && item.gridRow < otherPhysicalEnd) {
        item.gridRow = otherPhysicalEnd;
        snapChanged = true;
      }
    }
  }
}
setTableVisibleRows(placedItems);

    const worksheet = workbook.addWorksheet(sheetName);
    // the sheet as the dashboard's canvas: no cell grid, as in Tableau
    worksheet.views = [{ showGridLines: !!FORMAT_CONFIG.sheetGridlines }];

    /** @type {Record<number, number>} */
    const colWidths = {};
    const exactWidths = {};
    const allTablesInfo = [];
    const tracker = makeRangeTracker();

    let currentRow = 0;
    const dashboardName = dashboard.name || "Dashboard Export";
    let dashTitleProps = null;
    const dashFmt = fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name];
    if (dashFmt) {
      const run = dashFmt.title && (dashFmt.title.runs.find(x => x.text.trim()) || {}).props;
      dashTitleProps = tfMerge(TABLEAU_DEFAULTS.dashTitle,
        tfCollect(fmtModel.workbookStyle, TF_ELEMENTS.dashTitle), tfCollect(dashFmt.style, TF_ELEMENTS.dashTitle), run);
    }
    const usedZones = new Set(dataWorksheetItems.flatMap(it => it.vm ? it.vm.headerZoneIds : []));
    const titleRuns = tfDashboardTitleRuns(fmtModel, dashboard.name, usedZones);
    // Tableau shows the dashboard title only when its title zone is on (text boxes are drawn in place)
    const titleShown = !dashFmt || dashFmt.zones.some(z => z.type === "title" && !z.hidden);
    const titleHeight = writeDashboardTitle(worksheet, titleShown ? dashboardName : "", currentRow, 0, tracker, dashTitleProps,
      titleShown ? titleRuns : null);
    currentRow += titleHeight;


    const adjustedItems = placedItems.map(item => ({
      ...item,
      gridRow: item.gridRow + currentRow
    }));

    // native charts: data goes to a hidden sheet now, chart parts are injected after writeBuffer()
    const chartJobs = [];
    const reserveGraphicBlock = item => {
      for (let c = item.gridCol; c < item.gridCol + item.gridW; c++) colWidths[c] = Math.max(colWidths[c] || 0, 10);
      tracker.update(item.gridRow, item.gridCol);
      tracker.update(item.gridRow + item.allocatedRows - 1, item.gridCol + item.gridW - 1);
    };
    const writeGraphicTitle = item => {
      if (!item.vm.showTitle) return 0;
      writeTableauTitle(worksheet, item.gridRow, item.gridCol, item.vm.title.text, item.vm.title.props, item.gridW);
      return 1;
    };

    /** @type {{ item: ExportItem, top: number, panes: Map<string, any> }[]} pies, laid out with the final column widths */
    const pieJobs = [];
    // one hidden data sheet for the charts of every dashboard in the file (ctx.chartData: { sheet, row })
    const chartDataSheetOf = () => {
      if (!chartData.sheet) chartData.sheet = workbook.addWorksheet(CHART_DATA_SHEET, { state: "hidden" });
      return chartData.sheet;
    };

    for (let i = 0; i < adjustedItems.length; i++) {
      const item = adjustedItems[i];

      if (item.visualModel && item.visualModel.pie) {
        const status = visualStatuses.find(entry => entry.worksheet === item.name);
        const titleRows = writeGraphicTitle(item);
        const data = writePieData(chartDataSheetOf(), chartData.row + 1, item.visualModel.pie, item.name);
        chartData.row = data.nextRow - 1;
        pieJobs.push({ item, top: item.gridRow + titleRows, panes: data.panes });
        reserveGraphicBlock(item);
        const pie = item.visualModel.pie;
        item.visualModel.status = "success";
        item.visualModel.statusReason = `native Excel ${pie.hole || pie.inner ? "donut" : "pie"} chart with Tableau labels and tooltips`;
        if (status) { status.status = "success"; status.reason = item.visualModel.statusReason; }
        continue;
      }

      if (item.type === "legend") {
        writeLegendBlock(worksheet, item, tracker, colWidths);
        continue;
      }

      if (item.visualModel && item.visualModel.renderer === "excel-chart") {
        const status = visualStatuses.find(entry => entry.worksheet === item.name);
        const titleRows = writeGraphicTitle(item);
        const specs = item.visualModel.chartSpecs;
        const paneH = Math.floor(item.box.heightPx / specs.length);     // separate panes stack vertically
        specs.forEach((spec, k) => {
          const refs = ExcelChartWriter.writeChartData(chartDataSheetOf(), spec, chartData.row);
          chartData.row = refs.nextRow;
          const offPx = k * paneH;
          chartJobs.push({
            spec, refs, item, name: item.visualName || item.name,
            col: item.gridCol, row: item.gridRow + titleRows + Math.floor(offPx / EXCEL_ROW_PX),
            rowOffPx: offPx % EXCEL_ROW_PX, widthPx: item.box.widthPx, heightPx: paneH,
            top: item.gridRow + titleRows, pane: k, panes: specs.length
          });
        });
        reserveGraphicBlock(item);
        item.visualModel.status = "success";
        item.visualModel.statusReason = `native Excel ${[...new Set(specs.map(s => s.kind))].join("/")} chart` +
          (specs.length > 1 ? ` (${specs.length} panes)` : "");
        if (status) { status.status = "success"; status.reason = item.visualModel.statusReason; }
        continue;
      }

      if (item.visualModel && item.visualModel.renderer === "tableau-image") {
        const imageWidth = item.box ? item.box.widthPx : Math.max(320, item.gridW * PX_PER_COL);
        const imageHeight = item.box ? item.box.heightPx : Math.max(160, (item.allocatedRows || item.rowCount || 8) * PX_PER_ROW);
        const status = visualStatuses.find(entry => entry.worksheet === item.name);
        try {
          // bounded: after other dashboards were read through embedded views, Tableau may never answer
          const imageBase64 = await withTimeout(renderTableauImage(item.visualModel, imageWidth, imageHeight),
            IMAGE_RENDER_TIMEOUT_MS, "Tableau's picture of the visual");
          const titleRows = writeGraphicTitle(item);
          const imageId = workbook.addImage({ base64: imageBase64, extension: "png" });
          // a picture is not data: the visual's marks go to the Visual Data sheet too, and the picture links to them
          const dataLink = appendVisualData(workbook, item.visualName || item.name, item.vm);
          worksheet.addImage(imageId, /** @type {any} */ ({
            tl: { col: item.gridCol, row: item.gridRow + titleRows },
            ext: { width: imageWidth, height: imageHeight },
            hyperlinks: { hyperlink: dataLink, tooltip: "The data of this picture (Visual Data sheet)" }
          }));
          reserveGraphicBlock(item);
          item.visualModel.status = "success";
          item.visualModel.statusReason = "Tableau-rendered picture; its data on the Visual Data sheet";
          if (status) {
            status.status = "success";
            status.reason = item.visualModel.statusReason;
          }
          continue;
        } catch (err) {
          item.visualModel.renderer = "data-fallback";
          item.visualModel.status = "warning";
          item.visualModel.statusReason = `image renderer failed: ${err.message}`;
          // exported as its data: every column, as for any chart that falls back to a table
          item.vm.order = item.vm.cols.map((_, i) => i).filter(i => !/^(latitude|longitude) \(generated\)$/i.test(item.vm.cols[i].name));
          item.vm.cols.forEach(c => { if (!c.label) c.label = c.name; });
          item.vm.showHeaderRow = true;
          if (status) {
            status.renderer = "data-fallback";
            status.status = "warning";
            status.reason = item.visualModel.statusReason;
          }
          console.warn(`[Visual] ${item.name}: image renderer failed; using data fallback`, err);
        }
      }

      if (item.type === "filterValue") {
        writeIndividualFilterTable(
          worksheet,
          item.filterName,
          item.values,
          item.gridRow,
          item.gridCol,
          tracker,
          item.gridW
        );
        if (!item.layout) colWidths[item.gridCol] = Math.max(colWidths[item.gridCol] || 0, 35);
      } else if (item.image) {
        tracker.update(item.gridRow, item.gridCol);                 // the picture goes in once the columns are sized
        tracker.update(item.gridRow + item.allocatedRows - 1, item.gridCol + item.gridW - 1);
      } else if (item.textCard) {
        writeKpiCard(worksheet, item.textCard, null, item.gridRow, item.gridCol, item.gridW, tracker, colWidths);
      } else if (item.kpiCard) {
        writeKpiCard(worksheet, item.kpiCard, item.vm, item.gridRow, item.gridCol, item.gridW, tracker, colWidths, item.splitCuts);
      } else if (item.isKPI) {
        writeKPICardStacked(worksheet, item.vm, item.gridRow, item.gridCol, tracker);
        colWidths[item.gridCol] = Math.max(colWidths[item.gridCol] || 0, 22);
        colWidths[item.gridCol + 1] = Math.max(colWidths[item.gridCol + 1] || 0, 18);
      } else {
        writeRegularTable(worksheet, item.vm, item.gridRow, item.gridCol, tracker, allTablesInfo, colWidths, exactWidths, item.visibleRows,
                          item.fieldCols);
      }
    }

    // ── how each visual was converted (NATIVE … TABLE_FALLBACK), and why: the report sheet, the panel status, and a
    //    note on every visual exported as its data so the reader of the sheet knows what it stands for
    /** @type {ReportVisual[]} */
    const reportVisuals = [];
    adjustedItems.filter(it => it.type === "worksheet" && it.visualModel).forEach(item => {
      const semantics = describeVisual(item.visualModel, fmtModel);
      const conversion = conversionOf(item.visualModel, semantics);
      reportVisuals.push({ worksheet: item.name, semantics, ...conversion });
      const status = visualStatuses.find(entry => entry.worksheet === item.name);
      if (status) Object.assign(status, { strategy: conversion.strategy, visual: semantics.visual });
      if (FORMAT_CONFIG.fallbackNotes && conversion.strategy === STRATEGY.TABLE_FALLBACK) {
        worksheet.getCell(item.gridRow + 1, item.gridCol + 1).note =
          `Tableau visual: ${semantics.visual}\nExported as: its data (TABLE_FALLBACK)\nWhy: ${conversion.reason}`;
      }
    });


    // a KPI card directly above / below a chart of the same width and colour (Tableau's number + trend
    // tile): the layout gap between them gets the card colour too, so they read as one card
    const sameColor = (a, b) => !!a && !!b && String(a).replace(/^#/, "").slice(-6).toUpperCase() === String(b).replace(/^#/, "").slice(-6).toUpperCase();
    adjustedItems.filter(it => it.kpiCard && it.kpiCard.background).forEach(card => {
      const color = card.kpiCard.background;
      const cardEnd = card.gridRow + card.allocatedRows;
      adjustedItems.forEach(other => {
        if (other === card || other.gridCol !== card.gridCol || Math.abs(other.gridW - card.gridW) > 1) return;
        const specs = other.visualModel && other.visualModel.renderer === "excel-chart" ? other.visualModel.chartSpecs : null;
        if (!sameColor(specs && specs.length ? specs[0].background : null, color)) return;
        const otherEnd = other.gridRow + other.allocatedRows;
        const [from, to] = other.gridRow >= cardEnd ? [cardEnd, other.gridRow] : otherEnd <= card.gridRow ? [otherEnd, card.gridRow] : [0, 0];
        if (to - from < 1 || to - from > 2) return;
        for (let rr = from; rr < to; rr++) {
          for (let c = card.gridCol; c < card.gridCol + card.gridW; c++) {
            const cell = worksheet.getCell(rr + 1, c + 1);
            cell.fill = tfExcelFill(color);
            if (c === card.gridCol) cell.border = { left: KPI_GUTTER };
            if (c === card.gridCol + card.gridW - 1) cell.border = { ...(cell.border || {}), right: KPI_GUTTER };
          }
        }
      });
    });

    // bordered layout containers (Layout pane → Border, e.g. a card around a column of sheets): their
    // outline around the blocks they hold
    const zoneById = new Map(dashZones.map(z => [String(z.id), z]));
    const within = (zoneId, containerId) => {
      for (let z = zoneById.get(String(zoneId)); z && z.parent; z = zoneById.get(String(z.parent))) if (String(z.parent) === containerId) return true;
      return false;
    };
    dashZones.filter(z => /^layout-/.test(z.type) && !z.hidden && z.style && z.style.borderStyle && z.style.borderStyle !== "none" &&
                          (z.style.borderWidth || 0) > 0).forEach(container => {
      const members = adjustedItems.filter(it => it.layout && it.layout.id !== undefined && within(it.layout.id, String(container.id)));
      if (!members.length) return;
      const top = Math.min(...members.map(m => m.gridRow)), left = Math.min(...members.map(m => m.gridCol));
      const bottom = Math.max(...members.map(m => m.gridRow + m.allocatedRows)) - 1;
      const right = Math.max(...members.map(m => m.gridCol + m.gridW)) - 1;
      /** @type {Partial<import("exceljs").Border>} */
      const side = { style: /** @type {import("exceljs").BorderStyle} */ (tfStrokeToBorder(container.style.borderWidth) || "thin"),
                     color: { argb: container.style.borderColor || "FFD4D4D4" } };
      const edge = (r, c, sides) => { const cell = worksheet.getCell(r + 1, c + 1); cell.border = { ...(cell.border || {}), ...sides }; };
      for (let c = left; c <= right; c++) { edge(top, c, { top: side }); edge(bottom, c, { bottom: side }); }
      for (let r = top; r <= bottom; r++) { edge(r, left, { left: side }); edge(r, right, { right: side }); }
    });

    // dashboard shading, container / backdrop backgrounds and each object's own: the cells around and under
    // the blocks take the colour Tableau shows there; cells a block coloured itself keep their colour
    const dashColor = tfDashboardShading(fmtModel, dashboard.name);
    const bgBlocks = adjustedItems.filter(it => it.layout && it.layout.id !== undefined).map(it => {
      const zone = zoneById.get(String(it.layout.id));
      const zoneBg = (zone && zone.style && zone.style.bgColor) || undefined;
      // a worksheet draws its shading over its zone: its colour, Tableau's white, or see-through
      const shading = it.type === "worksheet" && it.vm ? it.vm.fmt.sheetShading() : "none";
      const own = shading === "none" ? zoneBg : shading || "white";
      return { zoneId: String(it.layout.id), top: it.gridRow, left: it.gridCol,
               bottom: it.gridRow + it.allocatedRows - 1, right: it.gridCol + it.gridW - 1, own };
    });
    // the dashboard's extent in px (its objects), for the size of zones given in 0–100000 units
    const dashExt = (dashboard.objects || []).reduce((a, o) => o.position && o.size
      ? { w: Math.max(a.w, o.position.x + o.size.width), h: Math.max(a.h, o.position.y + o.size.height) } : a, { w: 0, h: 0 });
    if (bgBlocks.length && (dashColor || dashZones.some(z => z.style && z.style.bgColor))) {
      const used = tracker.toRange();
      const plan = backgroundPlan(dashZones, dashColor, bgBlocks,
        { top: Math.min(...bgBlocks.map(b => b.top)), left: 0, bottom: used.e.r, right: used.e.c },
        { w: (dashFmt && dashFmt.width) || dashExt.w || 1200, h: (dashFmt && dashFmt.height) || dashExt.h || 800 });
      for (let r = 0; r <= used.e.r; r++) {
        for (let c = 0; c <= used.e.c; c++) {
          const color = plan.colorAt(r, c);
          if (!color) continue;
          const cell = worksheet.getCell(r + 1, c + 1);
          const target = cell.isMerged ? cell.master : cell;
          if (target.fill && target.fill.type === "pattern" && target.fill.pattern !== "none") continue;
          target.style = { ...target.style, fill: tfExcelFill(color) };      // a style of its own, not shared
        }
      }
      // dividers and accent lines: borders along the block edges and through the gaps, never across a block
      plan.lines.forEach(ln => {
        /** @type {Partial<import("exceljs").Border>} */
        const side = { style: ln.style, color: { argb: ln.color } };
        for (let k = ln.from; k <= ln.to; k++) {
          const across = bgBlocks.some(b => ln.dir === "h" ? b.top < ln.at && ln.at <= b.bottom && k >= b.left && k <= b.right
                                                         : b.left < ln.at && ln.at <= b.right && k >= b.top && k <= b.bottom);
          if (across) continue;
          const cell = ln.dir === "h" ? worksheet.getCell(ln.at + 1, k + 1) : worksheet.getCell(k + 1, ln.at + 1);
          const key = ln.dir === "h" ? "top" : "left";
          if (cell.border && cell.border[key]) continue;                   // a table's own rule stays
          cell.style = { ...cell.style, border: { ...(cell.border || {}), [key]: side } };
        }
      });
    }

    // the grid's columns at their dashboard widths (Excel: px = 7 × stored width); content no longer widens them
    grid.colPx.forEach((px, c) => { exactWidths[c] = Math.max(0.1, Math.round(px / 7 * 100) / 100); });
    setColumnWidths(worksheet, colWidths, exactWidths);
    applyAutoFilters(worksheet, allTablesInfo);
    if (FORMAT_CONFIG.printFitToWidth) {
      // printing / Save as PDF: the whole dashboard, one page wide, oriented like the dashboard
      const used = tracker.toRange();
      worksheet.pageSetup = { ...worksheet.pageSetup, orientation: dashExt.w >= dashExt.h ? "landscape" : "portrait",
        fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
        printArea: `A$1:${getExcelColumnName(used.e.c)}$${used.e.r + 1}` };          // ExcelJS adds the column's $
    }

    // every row a definite 15 pt (20 px): without one, Excel derives the row height from the display's scaling
    // (e.g. 19 px at 125 %), while charts and pictures keep their size – they would run into the block below.
    // Rows with wrapped text keep Excel's fit.
    {
      const lastRow = tracker.toRange().e.r + 1;
      for (let r = 1; r <= lastRow; r++) {
        const row = worksheet.getRow(r);
        if (row.height || row.hidden) continue;
        let wraps = false;
        row.eachCell(cell => { if (cell.alignment && cell.alignment.wrapText) wraps = true; });
        if (!wraps) row.height = 15;
      }
      // a row taller than 20 px (a big-font text box line) pushes everything below it down: give the excess back
      // in the following rows that hold nothing (no value, border or merge, no chart or picture over them), so
      // blocks further down start where they do on the dashboard
      const covered = new Set();
      [...chartJobs.map(j => j.item), ...adjustedItems.filter(it => it.image)].forEach(it => {
        for (let r = it.gridRow; r < it.gridRow + (it.allocatedRows || 0); r++) covered.add(r);
      });
      overlayImages.forEach(img => {
        const host = adjustedItems.find(it => it.layout && String(it.layout.id) === img.host);
        if (host) for (let r = host.gridRow; r < host.gridRow + (host.allocatedRows || 0); r++) covered.add(r);
      });
      const holdsNothing = ri => {
        if (covered.has(ri)) return false;
        let used = false;
        worksheet.getRow(ri + 1).eachCell(cell => {
          const b = cell.border || {};
          if (cell.isMerged || (cell.value !== null && cell.value !== undefined && cell.value !== "") || b.top || b.bottom || b.left || b.right) used = true;
        });
        return !used;
      };
      const dashTop = Math.min(...adjustedItems.filter(it => it.layout).map(it => it.gridRow));
      let drift = 0;
      for (let ri = isFinite(dashTop) ? dashTop : lastRow; ri < lastRow; ri++) {
        const row = worksheet.getRow(ri + 1);
        if (row.hidden) continue;
        const px = row.height ? Math.round(row.height * 4 / 3) : EXCEL_ROW_PX;
        if (drift > 0 && px === EXCEL_ROW_PX && holdsNothing(ri)) {
          const give = Math.min(drift, EXCEL_ROW_PX - 1);
          row.height = Math.round((EXCEL_ROW_PX - give) * 0.75 * 4) / 4;
          drift -= give;
        } else drift += px - EXCEL_ROW_PX;
        if (drift < 0) drift = 0;
      }
    }
    // final column widths / row heights in px: charts and pictures are fitted to their blocks with them
    // a column's stored width already holds Excel's padding: 7 px per unit (Calibri 11), 64 px when unset
    const colPx = ci => { const w = worksheet.getColumn(ci + 1).width; return w ? Math.round(w * 7) : 64; };
    const rowPx = ri => {                            // collapsed rows take no space on screen
      const row = worksheet.getRow(ri + 1);
      if (row.hidden) return 0;
      return row.height ? Math.round(row.height * 4 / 3) : EXCEL_ROW_PX;
    };
    const blockPx = it => {
      let w = 0, h = 0;
      for (let c = it.gridCol; c < it.gridCol + it.gridW; c++) w += colPx(c);
      for (let r = it.gridRow; r < it.gridRow + it.allocatedRows; r++) h += rowPx(r);
      return { w, h };
    };
    /* a picture at (x, y) px from the block's top-left cell, w × h px, linked like its zone */
    const addPicture = async (image, col, row, x, y, w, h) => {
      const pic = await prepareImage(image.file, image.info, w, h);
      if (!pic) { console.warn(`[Images] ${image.zone.param}: cannot be drawn in this browser`); return false; }
      const id = workbook.addImage({ base64: pic.base64, extension: pic.extension });
      worksheet.addImage(id, /** @type {any} */ ({
        tl: nativeAnchor(col, row, x, y, colPx, rowPx), ext: { width: w, height: h }, editAs: "oneCell",
        ...(image.url ? { hyperlinks: { hyperlink: image.url, tooltip: image.url } } : {})
      }));
      return true;
    };
    let picturesAdded = 0;
    for (const item of adjustedItems.filter(it => it.image)) {
      const img = item.image, block = blockPx(item);
      // Fit Image / Center Image inside the block, never larger than Tableau draws it in its zone
      const fit = fitImage(img.info.width, img.info.height, block.w, block.h, img.zone.scaled, img.zone.centered);
      const cap = fitImage(img.info.width, img.info.height, img.widthPx, img.heightPx, img.zone.scaled, img.zone.centered);
      const w = Math.min(fit.w, cap.w), h = Math.min(fit.h, cap.h);
      const centered = img.zone.centered !== false;
      const x = centered ? Math.round((block.w - w) / 2) : 0, y = centered ? Math.round((block.h - h) / 2) : 0;
      if (await addPicture(img, item.gridCol, item.gridRow, x, y, w, h)) picturesAdded++;
    }
    for (const img of overlayImages) {
      // an icon over a sheet: the same spot of that sheet's block, scaled with the block
      const host = adjustedItems.find(it => it.layout && String(it.layout.id) === img.host);
      if (!host || !host.layout.widthPx || !host.layout.heightPx) continue;
      const block = blockPx(host);
      const sx = block.w / host.layout.widthPx, sy = block.h / host.layout.heightPx, k = Math.max(0.5, Math.min(1.5, sx, sy));
      const fit = fitImage(img.info.width, img.info.height, img.widthPx * k, img.heightPx * k, img.zone.scaled, img.zone.centered);
      if (await addPicture(img, host.gridCol, host.gridRow, Math.round((img.xPx - host.layout.xPx) * sx) + fit.x,
                           Math.round((img.yPx - host.layout.yPx) * sy) + fit.y, fit.w, fit.h)) picturesAdded++;
    }
    if (picturesAdded) console.log(`[Images] ${picturesAdded} picture(s) added`);

    // the Conversion Report's part for this dashboard (written once for all dashboards by exportToExcel)
    const objects = [
      ...textItems.map(it => ({ name: String(it.visualName || "").slice(0, 60) || it.name, kind: "Text box", output: "Cells, with Tableau's fonts" })),
      ...imageItems.map(it => ({ name: it.visualName, kind: iconSheets.some(s => s.name === it.name) ? "Button / icon sheet" : "Image",
                                 output: "Picture" })),
      ...pieLegendItems.map(it => ({ name: it.visualName, kind: "Colour legend", output: "Cells: a colour key per item" })),
      ...filterValueItems.map(it => ({ name: it.filterName, kind: it.isParameter ? "Parameter control" : "Filter card",
                                       output: it.isParameter ? "Cells: its value" : "Cells: the selected values" })),
      ...[...pieCentreSheets].map(name => ({ name, kind: "Worksheet in a donut's hole", output: "The donut's centre text" }))
    ];
    const report = {
      dashboard: dashboardName, workbookFile: formatModelFileName(), visuals: reportVisuals, objects, skipped,
      filters: Object.entries(filterValuesMap).map(([name, values]) => ({ name, values: values.map(String) })),
      parameters: reportParams
    };

    // native charts, once ExcelJS has written the file: pies and donuts get Tableau's pane cells inside the block
    // (slices link back to the block's first cell); every chart is fitted to its block's final px size
    const finalizeCharts = () => {
      /** @type {ChartJob[]} */
      const pieChartJobs = [];
      pieJobs.forEach(({ item, top, panes }) => {
        let w = 0, h = 0;
        for (let c = item.gridCol; c < item.gridCol + item.gridW; c++) w += colPx(c);
        for (let r = top; r < item.gridRow + item.allocatedRows; r++) h += rowPx(r);
        const bg = (item.vm.fmt.tableBackground() || "FFFFFFFF").slice(2);
        const link = `#'${worksheet.name.replace(/'/g, "''")}'!${worksheet.getCell(top + 1, item.gridCol + 1).address}`;
        buildPieCharts(item.visualModel.pie, { w, h }, panes, { name: item.visualName || item.name, bg, link }).forEach(c => {
          const at = nativeAnchor(item.gridCol, top, c.x, c.y, colPx, rowPx);
          pieChartJobs.push({ spec: /** @type {any} */ ({ kind: "pie" }), refs: /** @type {any} */ (null), item, name: c.name,
            col: at.nativeCol, row: at.nativeRow, colOffPx: at.nativeColOff / 9525, rowOffPx: at.nativeRowOff / 9525,
            widthPx: c.w, heightPx: c.h, xml: c.xml, shapes: c.shapes, tips: c.tips });
        });
      });
      if (chartJobs.length || pieChartJobs.length) {
        // fit every chart to its own block now that the final column widths/row heights are known:
        // tables widen the columns, so a chart kept at its Tableau pixel size would end short of the
        // tables / charts aligned with it on the dashboard → stretch it to the block's edges
        chartJobs.forEach(job => {
          let w = 0, h = 0;
          for (let c = job.item.gridCol; c < job.item.gridCol + job.item.gridW; c++) w += colPx(c);
          for (let r = job.top; r < job.item.gridRow + job.item.allocatedRows; r++) h += rowPx(r);
          job.widthPx = Math.max(40, w);
          if (job.item.pairedCard) { job.colOffPx = 2; job.widthPx = Math.max(40, w - 4); }     // the card's white edges
          // stacked panes split the block's real height (rows may be taller, shorter or collapsed), each from its own spot
          const paneH = Math.floor(h / job.panes);
          let r = job.top, off = job.pane * paneH;
          for (let guard = 0; guard < 5000 && off >= rowPx(r) && (rowPx(r) > 0 || off > 0); guard++) { off -= rowPx(r); r++; }
          job.row = r;
          job.rowOffPx = off;
          job.heightPx = Math.max(20, job.pane === job.panes - 1 ? h - job.pane * paneH : paneH);
          // packed bubbles: Tableau's tooltip on each bubble (a click follows the link to the chart's own cell)
          const link = `#'${worksheet.name.replace(/'/g, "''")}'!${worksheet.getCell(job.row + 1, job.col + 1).address}`;
          const tips = ExcelChartWriter.bubbleTips(job.spec, job, link);
          if (tips.length) job.tips = tips;
        });
        return [...chartJobs, ...pieChartJobs];
      }
      return [];
    };

    return { worksheet, dashboardName, visualStatuses, workbookWarning, report, finalizeCharts };
}


/* ═══ main.js ═════════════════════════════════════════════════════════════════════════════════════ */
/* Entry point: wires the panel once Tableau has initialised the extension –
 *   📁 Load workbook (manual .twb / .twbx), the loaded-workbook label (remembered workbook), the Tableau Cloud
 *   auto-load (backend, Advanced settings) and Export (waits for a running auto-load first). */

/* the loaded-workbook label at start-up: the workbook saved with the dashboard, else one remembered in this browser */
function restoreWorkbookLabel() {
  const fileLabel = document.getElementById("twb_file_label");
  if (!fileLabel) return Promise.resolve();
  const savedFileName = tableau.extensions.settings.get("twbFileName");
  const savedTitleMap = tableau.extensions.settings.get("twbTitleMap");
  const savedFormat = tableau.extensions.settings.get("twbFormatModel");
  if (savedFileName && savedTitleMap) {
    const titleCount = Object.keys(JSON.parse(savedTitleMap)).length;
    if (savedFormat) setFormatModelCache(JSON.parse(savedFormat), savedFileName);
    return ensureFormatModel().then(model => {
      if (model) {
        showWorkbookLabel(savedFileName, `${titleCount} titles, formatting for ${Object.keys(model.sheets || {}).length} sheets`, model);
      } else {
        fileLabel.textContent = `${savedFileName} — formatting not available after reload.
⚠ Click 📁 Load Workbook again so charts are exported as charts`;
        fileLabel.classList.add("status-warning");
      }
    });
  }
  fileLabel.textContent = "No workbook loaded — click 📁 to load titles and colors";
  return ensureFormatModel().then(model => {
    if (model && RESTORED_FILE_NAME) {
      showWorkbookLabel(RESTORED_FILE_NAME, `remembered from an earlier load, formatting for ${Object.keys(model.sheets || {}).length} sheets`, model);
    } else {
      fileLabel.textContent = "⚠ No workbook loaded — click 📁 Load Workbook, otherwise charts are exported as tables";
      fileLabel.classList.add("status-warning");
    }
  });
}

document.addEventListener("DOMContentLoaded", () => {
  tableau.extensions.initializeAsync().then(() => {
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    /** the export waits on this while the workbook is being loaded automatically */
    let WORKBOOK_READY = Promise.resolve();

    const loadBtn = /** @type {HTMLButtonElement} */ (document.getElementById("load_workbook_btn"));
    if (loadBtn) {
      loadBtn.addEventListener("click", async () => {
        loadBtn.disabled = true;
        loadBtn.textContent = "⏳ Loading...";
        await loadWorkbookFile();
        loadBtn.disabled = false;
        loadBtn.textContent = "📁 Load workbook file manually (.twb / .twbx)";
      });
    }

    // the remembered workbook first, then Phase 1: fetch the workbook from Tableau Cloud automatically
    // (falls back to 📁 on any problem; Tableau Desktop keeps the manual load)
    WORKBOOK_READY = restoreWorkbookLabel()
      .catch(err => console.warn("[Workbook] could not restore the remembered workbook:", err))
      .then(() => autoLoadWorkbook(dashboard));
    setupBackendSettings(dashboard, reload => { WORKBOOK_READY = reload; });

    const exportBtn = /** @type {HTMLButtonElement} */ (document.getElementById("export_button"));
    if (exportBtn) {
      exportBtn.addEventListener("click", async () => {
        // If the auto-load is still running, wait for it so the export uses the right formatting.
        const original = exportBtn.textContent;
        exportBtn.disabled = true;
        exportBtn.textContent = "⏳ Loading workbook…";
        try { await WORKBOOK_READY; } finally {
          exportBtn.textContent = original;
          exportBtn.disabled = false;
        }
        exportToExcel();
      });
    }

  }).catch((err) => {
    console.error("[Tableau init]", err);
    alert("Failed to connect to Tableau: " + err.message);
    if (document.getElementById("export_button")) {
      /** @type {HTMLButtonElement} */ (document.getElementById("export_button")).disabled = true;
    }
  });
});

})();
