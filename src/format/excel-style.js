/* Resolved Tableau style props → ExcelJS font / fill / alignment, and cell writing. */
import { FORMAT_CONFIG } from "../config.js";
import { tfBrightness } from "./colors.js";
import { inferExcelNumFmt } from "./number-format.js";
import { tfMerge } from "../twb/formatter.js";

export function tfExcelFont(p) {
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

export function tfExcelAlignment(p, isNumber) {
  return { horizontal: p.hAlign || (isNumber ? "right" : "left"), vertical: p.vAlign || "middle", wrapText: !!p.wrap };
}

/** @param {string} argb @returns {import("exceljs").Fill | undefined} */
export function tfExcelFill(argb) {
  return argb ? { type: "pattern", pattern: "solid", fgColor: { argb } } : undefined;
}

export function tfRichRuns(runs, baseProps) {
  const out = [];
  runs.forEach(r => {
    const text = r.text.replace(/\u00C6\r?\n?/g, "\n");
    if (text) out.push({ text, font: tfExcelFont(tfMerge(baseProps, r.props)) });
  });
  return out;
}

/* writes value + style; returns true if the cell holds a real number */
export function tfWriteCell(cell, dv, p, extra = {}) {
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
