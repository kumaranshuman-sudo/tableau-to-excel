/* The data behind visuals exported as pictures (IMAGE_FALLBACK): each one's marks as a plain table on a "Visual Data"
 * sheet, so the numbers travel with the picture (which links to them). */
import { tfDvText, tfIsNull } from "../data/values.js";
import { inferExcelNumFmt } from "../format/number-format.js";

export const DATA_SHEET = "Visual Data";

/**
 * Appends one visual's data (every column but the latitude / longitude Tableau generates) under a title.
 * @param {import("exceljs").Workbook} workbook @param {string} title @param {ViewModel} vm
 * @returns {string} the link to the block ("#'Visual Data'!A5")
 */
export function appendVisualData(workbook, title, vm) {
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
