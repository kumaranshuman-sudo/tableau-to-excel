/* KPI cards and dashboard text boxes: formatted text in a block that fills its dashboard zone. */
import { tfDvText } from "../data/values.js";
import { tfExcelFill, tfExcelFont, tfWriteCell } from "../format/excel-style.js";
import { tfParseFieldRef, tfSameField } from "../twb/field-ref.js";
import { TABLEAU_DEFAULTS } from "../config.js";
import { tfFieldInfo, tfMerge } from "../twb/formatter.js";
import { buildColorPlan, writeTableauTitle } from "./visual-writers.js";

/* Tableau's formatted text: "Æ" marks a line break – usually followed by a newline, sometimes only a space */
const LINE_BREAK = /Æ[ \t]*(?:\r?\n)?|\r?\n/;

/* white edges between neighbouring cards – Tableau's tile gutters are narrower than a grid column */
/** @type {Partial<import("exceljs").Border>} */
export const KPI_GUTTER = { style: "medium", color: { argb: "FFFFFFFF" } };

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
export function formattedLines(runs, token, styled) {
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
export function buildKpiCard(vm, sheetName, layout) {
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
    tiles = [{ lines: lines.map(segments => line(segments, fmt)) }];
  } else {
    const onShelf = shelf => !!sheet && sheet[shelf].some(r => r.name === "Measure Names");
    if (onShelf("rows")) return null;
    if (onShelf("cols")) {                                  // Measure Names on Columns: a tile per measure, name above value
      const nameStyle = fmt.headerCellStyle(tfParseFieldRef("[:Measure Names]"));
      tiles = visible.map(ci => ({ lines: [
        line([{ text: vm.cols[ci].label || vm.cols[ci].name, props: nameStyle }], fmt),
        line([{ ci, props: fmt.markCellStyle(vm.cols[ci].ref) }], fmt)
      ] }));
    } else {                                                // Text = the measures: one value per line
      tiles = [{ lines: visible.map(ci => line([{ ci, props: fmt.markCellStyle(vm.cols[ci].ref) }], fmt)) }];
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
export function buildTextCard(zone, layout, model) {
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
function line(segments, fmt) {
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
 */
export function writeKpiCard(worksheet, card, vm, originRow, originCol, gridW, rangeTracker, colWidths) {
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
  card.tiles.forEach((tile, k) => {
    const c0 = C + k * span, c1 = k === n - 1 ? C + gridW - 1 : c0 + span - 1;
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
