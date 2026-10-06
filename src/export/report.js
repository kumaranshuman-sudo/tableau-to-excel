/* The conversion report: a sheet after the dashboard that lists every visual – what Tableau shows, how Excel shows
 * it (NATIVE … TABLE_FALLBACK), the fidelity and why – then the dashboard's other objects, the filters and
 * parameters the data reflects, and what was left out. Nothing is converted silently. */
import { STRATEGY_ORDER } from "../visual/strategy.js";

export const REPORT_SHEET = "Conversion Report";

const STRATEGY_FILL = { NATIVE: "FFE2F0D9", CONSTRUCTED: "FFDDEBF7", APPROXIMATE: "FFFFF2CC", TABLE_FALLBACK: "FFEDEDED",
                        IMAGE_FALLBACK: "FFE4DFEC", UNSUPPORTED: "FFF8CBAD", FAILED: "FFF4B183" };

/** "15 NATIVE · 3 CONSTRUCTED · 2 TABLE_FALLBACK" @param {{ strategy: string }[]} visuals */
export function strategyCounts(visuals) {
  return STRATEGY_ORDER.map(s => [s, visuals.filter(v => v.strategy === s).length]).filter(([, n]) => n);
}

/**
 * @param {import("exceljs").Workbook} workbook
 * @param {{ dashboard: string, workbookFile: string | null, visuals: ReportVisual[],
 *           objects: { name: string, kind: string, output: string }[], skipped: { name: string, reason: string }[],
 *           filters: { name: string, values: string[] }[], parameters: { name: string, value: string }[] }} report
 */
export function writeConversionReport(workbook, report) {
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
