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

export const FORMAT_CONFIG = {
  substituteTableauFonts: true,   // "Tableau Book" etc. are only installed with Tableau Desktop
  tableauFontSubstitute: "Arial",
  writeNativeNumbers: true,       // write real numbers + Excel numFmt instead of text
  applyFallbackHeatmap: false,    // old red/green heatmap when Tableau has no color (not in Tableau → off)
  nativeCharts: true,             // bar/line/area/pie/scatter/combo → editable native Excel charts
  chartPolicy: "data",            // charts that cannot be drawn: "skip" | "data" (export their data as a plain table)
  groupOverflowRows: true,        // collapse rows beyond ROW_GROUP_THRESHOLD into an expandable [+]/[-] group
  autoFilter: "largest",          // Excel allows ONE autofilter per sheet: "largest" table | "none"
  textBoxHeaders: true,           // use dashboard text boxes as column headers – only when they line up exactly with the table
  textBoxTitle: false,            // text boxes are drawn in place now; true = also use the top one as the sheet title
  linkText: "url",                // URL-action cells: "url" shows the link itself, any other string is shown as the text
  sheetGridlines: false,          // Excel's cell grid on the dashboard sheet – Tableau dashboards have none
  printFitToWidth: true,          // printing / PDF: the dashboard one page wide, landscape when it is wider than tall
  debug: true
};

export const VISUAL_TYPES = Object.freeze({
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
export const TABLEAU_DEFAULTS = {
  worksheet:  { fontName: "Tableau Book", fontSize: 9, color: "FF333333" },
  fieldLabel: { bold: true },
  title:      { fontName: "Tableau Book", fontSize: 15, color: "FF333333" },
  dashTitle:  { fontName: "Tableau Book", fontSize: 18, color: "FF333333" },
  divider:    { visible: true, style: "thin", color: "FFD4D4D4", level: 1 }
};

/* XML element names used in <style-rule element='…'>. Arrays = aliases seen
 * across Tableau versions; unknown ones are simply ignored. */
export const TF_ELEMENTS = {
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

export const TF_DERIV_LABEL = {
  sum: "SUM", avg: "AVG", cnt: "CNT", ctd: "CNTD", min: "MIN", max: "MAX",
  med: "MEDIAN", attr: "ATTR", usr: "AGG", std: "STDEV", stdp: "STDEVP",
  var: "VAR", varp: "VARP", yr: "YEAR", qr: "QUARTER", mn: "MONTH", wk: "WEEK",
  dy: "DAY", hr: "HOUR", tyr: "YEAR", tqr: "QUARTER", tmn: "MONTH", twk: "WEEK", tdy: "DAY"
};
