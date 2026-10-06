/* Tables and KPI cards written with Tableau formatting; widths and autofilters. */
import { FORMAT_CONFIG } from "../config.js";
import { tfDvText, tfIsNull } from "../data/values.js";
import { applyConditionalFormattingToTable } from "./conditional-format.js";
import { buildMarkPlan, tvGlyphNumFmt, writeMarkBars } from "./mark-cells.js";
import { ROW_GROUP_THRESHOLD, getExcelColumnName, headerRowSpan } from "./layout.js";
import { tfBuildColorScale } from "../format/color-scale.js";
import { inferExcelNumFmt } from "../format/number-format.js";
import { tfExcelFill, tfExcelFont, tfRichRuns, tfWriteCell } from "../format/excel-style.js";
import { tfParseFieldRef, tfSameField } from "../twb/field-ref.js";
import { tfMerge } from "../twb/formatter.js";

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
export function writeLegendBlock(worksheet, item, rangeTracker, colWidths) {
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

export function writeTableauTitle(worksheet, r, C, text, p, span) {
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
export function buildColorPlan(fmt, colInfo, rows) {
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

export function borderSide(d) { return d && d.visible && d.style ? { style: d.style, color: { argb: d.color } } : undefined; }

export function writeKPICardStacked(worksheet, vm, originRow, originCol, rangeTracker) {
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
export function tableColumnPx(vm) {
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
export function writeRegularTable(worksheet, vm, originRow, originCol, rangeTracker, allTablesInfo, colWidths, exactWidths,
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

export function applyAutoFilters(worksheet, allTablesInfo) {
  if (FORMAT_CONFIG.autoFilter === "none" || !allTablesInfo || !allTablesInfo.length) return;
  // Excel allows a single autofilter per sheet → put it on the largest table that has a header row
  const t = allTablesInfo.filter(x => x.hasHeader).sort((a, b) => b.rowCount - a.rowCount)[0];
  if (!t) return;
  worksheet.autoFilter = `${getExcelColumnName(t.leftCol)}${t.headerRow + 1}:${getExcelColumnName(t.rightCol)}${t.bottomRow + 1}`;
}

export function setColumnWidths(worksheet, colWidths, exactWidths = {}) {
  const keys = [...Object.keys(colWidths), ...Object.keys(exactWidths)].map(Number);
  const maxColIdx = keys.length ? Math.max(...keys) : 0;
  for (let ci = 0; ci <= maxColIdx; ci++) {
    if (exactWidths[ci]) { worksheet.getColumn(ci + 1).width = exactWidths[ci]; continue; }
    worksheet.getColumn(ci + 1).width = Math.min((colWidths[ci] || 10), 50);
  }
}
