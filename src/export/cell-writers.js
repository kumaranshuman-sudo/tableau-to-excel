/* Low-level cell writers: titles, headers, data cells, filter tables. */
import { TABLEAU_DEFAULTS } from "../config.js";
import { writeTableauTitle } from "./visual-writers.js";
import { tfExcelFont, tfRichRuns } from "../format/excel-style.js";
import { tfMerge } from "../twb/formatter.js";

/* ── Quick filter / parameter control ───────────────────────────────────
 * As Tableau draws one: its caption, then the value(s) in a white box with a thin grey border, across the
 * control's zone. */
export function writeIndividualFilterTable(worksheet, filterName, filterValues, originRow, originCol, rangeTracker, width = 1) {
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

/* ── ExcelJS Cell Styling Functions ───────────────────────────────────── */
export function setCellValue(worksheet, row, col, value, options = {}) {
  const cell = worksheet.getCell(row + 1, col + 1);
  
  if (typeof value === 'number') {
    cell.value = value;
  } else {
    cell.value = value;
  }
  
  cell.font = {
    bold: options.bold || false,
    size: options.size || 11,
    name: 'Calibri',
    italic: options.italic || false
  };
  
  if (options.color) {
    cell.font.color = { argb: options.color };
  }
  
  cell.alignment = {
    vertical: 'middle',
    horizontal: options.align || 'left',
    wrapText: options.wrapText || false
  };
  
  if (options.bgColor) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: options.bgColor }
    };
  }
  
  if (options.border) {
    cell.border = options.border;
  }
}

export const thinBorder = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};

export const thickBorder = {
  top: { style: 'medium' },
  left: { style: 'medium' },
  bottom: { style: 'medium' },
  right: { style: 'medium' }
};

export function subtitleCell(worksheet, row, col, value) {
  setCellValue(worksheet, row, col, value, {
    bold: false,
    size: 11,
    color: "FF666666",
    align: "center"
  });
}

export function titleCell(worksheet, row, col, value) {
  setCellValue(worksheet, row, col, value, {
    bold: true,
    size: 12,
    bgColor: "FFE8F0FE",
    align: "center",
    border: thickBorder
  });
}

export function tableHeaderCell(worksheet, row, col, value) {
  setCellValue(worksheet, row, col, value, {
    bold: true,
    size: 11,
    bgColor: "FF1A73E8",
    color: "FFFFFFFF",
    align: "center",
    border: thinBorder,
    wrapText: true
  });
}

export function tableDataCell(worksheet, row, col, value, rowIndex = 0) {
  const bgColor = (rowIndex % 2 === 0) ? null : "FFF9F9F9";
  setCellValue(worksheet, row, col, value, {
    align: "left",
    size: 11,
    bgColor: bgColor,
    border: thinBorder,
    wrapText: true
  });
}

export function writeDashboardTitle(worksheet, dashboardName, exportDate, originRow, originCol, rangeTracker, titleProps, titleRuns) {
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
