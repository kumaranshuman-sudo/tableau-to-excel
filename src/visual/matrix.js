/* Heat maps, highlight tables, calendars and crosstabs as Tableau draws them: a matrix – the Rows shelf's members
 * down, the Columns shelf's members across, one cell per mark holding its label (else its size measure, else its
 * colour measure, else its shape), filled in the mark's colour. And the guard that keeps a table of marks from
 * dropping what the marks show. */
import { tfDvNum, tfDvText, tfIsNull, tfNaturalCompare } from "../data/values.js";
import { tfBuildColorScale } from "../format/color-scale.js";
import { tfSameField } from "../twb/field-ref.js";
import { tfBucketKey } from "../twb/parser.js";
import { tfNorm } from "../util.js";
import { tvIsMeasureRef } from "./classify.js";

/** @param {SheetModel} sheet the panes Tableau draws: each with an id when there are several */
const drawnPanes = sheet => sheet.panes.length > 1 && sheet.panes.some(p => p.id) ? sheet.panes.filter(p => p.id) : sheet.panes;

/**
 * Pivots the view model into the matrix Tableau shows, in place. Only a view of dimensions on both shelves (no axes, no
 * Measure Names) whose marks each hold one value; anything else keeps its one-row-per-mark table.
 * @param {ViewModel} vm @param {FormatModel | null} model @returns {boolean} pivoted
 */
export function pivotMatrix(vm, model) {
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
export function checkLossless(vm, model) {
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
