/* Visual type classification from marks, shelves and encodings. */
import { VISUAL_TYPES } from "../config.js";
import { tfDvNum } from "../data/values.js";
import { tfSameField } from "../twb/field-ref.js";
import { tfFieldInfo } from "../twb/formatter.js";

export function isKPIViewModel(vm) {
  const vis = vm.order.map(i => vm.cols[i]);
  return vm.rows.length === 1 && vis.length >= 1 && vis.length <= 8 && !vis.some(c => c.isHeader);
}

export function normalizeVisualToken(value) {
  return String(value == null ? "" : value).toLowerCase().replace(/[\s_-]+/g, "");
}

/* mark types of every marks card in the live visual specification
 * (marksSpecifications[].type: "bar" | "line" | "gantt-bar" | …) */
export function visualSpecMarkTokens(spec) {
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
export const TV_DATE_DERIVS = new Set(["yr", "qr", "mn", "wk", "dy", "hr", "mi", "sc", "my", "md", "mdy", "wd",
                                "tyr", "tqr", "tmn", "twk", "tdy", "thr", "tmi", "tsc"]);

export function tvIsMeasureRef(model, ref) {
  if (!ref) return false;
  if (ref.name === "Multiple Values") return true;
  if (ref.type !== "qk") return false;
  if (TV_DATE_DERIVS.has(String(ref.deriv || "").toLowerCase())) return false;
  const info = tfFieldInfo(model, ref) || {};
  return !(info.role === "dimension" && /^date/i.test(info.datatype || ""));
}

/* what the TWB shelves say about the view: value axes, continuous dimension axes, geography,
 * and the evidence for treemaps, histograms, waterfalls and box plots */
export function visualShelfShape(sheet, model) {
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
    binned: all.some(r => /\(bin\)$/i.test(r.name)),
    countAxis: all.some(r => tvIsMeasureRef(model, r) && /^(cnt|ctd)$/i.test(r.deriv || "")),
    numericDiscreteDim: all.some(r => r.type === "ok" && !TV_DATE_DERIVS.has(String(r.deriv || "").toLowerCase()) &&
                                      /^(integer|real)$/i.test(datatype(r))),
    runningTotal: all.some(r => /^(cum|rsum)$/i.test(r.deriv || "") ||
                                (sheet.runningTotals || []).some(t => tfSameField(t, r))),
    boxPlot: !!sheet.boxPlot,
    // a date on the shelves (discrete MONTH(…) or continuous) → Tableau's Automatic mark is a line
    dateDimension: all.some(r => !tvIsMeasureRef(model, r) && (TV_DATE_DERIVS.has(String(r.deriv || "").toLowerCase()) ||
                                                              (/^date/i.test(datatype(r)) && r.type !== "nk"))),
    shelfDims: all.filter(r => r.name !== "Measure Names" && !tvIsMeasureRef(model, r)).length
  };
}

/* the same evidence from the live visual specification (no workbook loaded) */
export function liveSpecShape(spec) {
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

/* mark type: live visual spec (current state) → TWB pane marks → Tableau's "Automatic" rules */
/** @param {any} spec live visual spec @param {ViewModel} vm @param {FormatModel | null} model */
export function resolveVisualMarks(spec, vm, model) {
  const sheet = vm.fmt && vm.fmt.sheetModel;
  const shape = visualShelfShape(sheet, model);
  let tokens = visualSpecMarkTokens(spec);
  let source = tokens.length ? "live" : "none";
  if (!tokens.length && sheet) {
    tokens = [...new Set(sheet.panes.map(p => normalizeVisualToken(p.markClass)).filter(t => t && t !== "automatic"))];
    if (tokens.length) source = "twb";
  }
  if (!tokens.length && shape) {
    source = "automatic";
    if (shape.geo) tokens = ["map"];
    else if (shape.rowMeasures && shape.colMeasures) tokens = ["circle"];
    else if (shape.rowMeasures || shape.colMeasures) tokens = [shape.continuousDimension || shape.dateDimension ? "line" : "bar"];
    else if (shape.shelfFields === 0 && shape.size) tokens = ["square"];   // empty shelves + Size → treemap
    else tokens = ["text"];
  }
  return { tokens, shape, source };
}

/**
 * A table built from marks: every axis on Rows / Columns is a constant ("MIN(0)", "0.5") that only places
 * the marks – one column of circles, squares, shapes or bars per axis, each pane showing its label.
 * @param {ViewModel} vm @param {FormatModel | null} model
 */
export function isMarkTable(vm, model) {
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

/** @param {any} spec live visual spec @param {ViewModel} vm @param {FormatModel | null} model @returns {VisualType} */
export function classifyVisualType(spec, vm, model) {
  const { tokens, shape } = resolveVisualMarks(spec, vm, model);
  const live = liveSpecShape(spec);
  const axes = shape ? shape.rowMeasures + shape.colMeasures : null;     // null → unknown (no workbook)
  const token = tokens[0] || "";
  // generated lat/long on the shelves: a map whatever the mark (circle → symbol map)
  if (shape && shape.geo) {
    return shape.filled || tokens.some(t => t === "multipolygon" || t === "polygon") ? VISUAL_TYPES.MAP_FILLED : VISUAL_TYPES.MAP;
  }
  if (axes && isMarkTable(vm, model)) return isKPIViewModel(vm) ? VISUAL_TYPES.KPI : VISUAL_TYPES.TABLE;
  if (tokens.includes("text") && tokens.every(t => /^(text|shape|circle|square)$/.test(t))) {
    return isKPIViewModel(vm) ? VISUAL_TYPES.KPI : VISUAL_TYPES.TABLE;      // KPI tiles / buttons built from marks
  }
  const cartesian = tokens.filter(t => t === "bar" || t === "line" || t === "area");
  if (new Set(cartesian).size > 1) return VISUAL_TYPES.COMBO;
  if (token === "ganttbar" || token === "gantt") {
    return shape && shape.runningTotal ? VISUAL_TYPES.WATERFALL : VISUAL_TYPES.GANTT;
  }
  if (shape && shape.boxPlot && axes > 0) return VISUAL_TYPES.BOXPLOT;
  const byMark = {
    bar: VISUAL_TYPES.BAR, line: VISUAL_TYPES.LINE, area: VISUAL_TYPES.AREA, pie: VISUAL_TYPES.PIE,
    map: VISUAL_TYPES.MAP, multipolygon: VISUAL_TYPES.MAP_FILLED, polygon: VISUAL_TYPES.MAP_FILLED,
    heatmap: VISUAL_TYPES.MAP, density: VISUAL_TYPES.MAP,               // density marks: Tableau-only visual
    vizextension: VISUAL_TYPES.CUSTOM
  };
  if (token === "bar" && axes === 0) return isKPIViewModel(vm) ? VISUAL_TYPES.KPI : VISUAL_TYPES.TABLE;
  if (token === "bar" && ((shape && (shape.binned || (shape.countAxis && shape.numericDiscreteDim))) || (live && live.binned))) {
    return VISUAL_TYPES.HISTOGRAM;
  }
  if (byMark[token]) return byMark[token];
  if (token === "circle" || token === "shape" || token === "square") {
    // nothing on Rows/Columns + Size → Tableau lays the marks out itself: treemap / packed bubbles
    const free = shape ? shape.shelfFields === 0 && shape.size : live ? live.shelfFields === 0 && live.size : false;
    if (free) return token === "square" ? VISUAL_TYPES.TREEMAP : VISUAL_TYPES.BUBBLE;
    if (axes === null) return token === "circle" ? VISUAL_TYPES.SCATTER : VISUAL_TYPES.TABLE;
    if (shape.rowMeasures && shape.colMeasures) return VISUAL_TYPES.SCATTER;
    if (axes > 0) return VISUAL_TYPES.LINE;                              // dot plot: markers on one axis
    return token === "square" ? VISUAL_TYPES.HEATMAP : VISUAL_TYPES.TABLE;
  }
  if (!token && vm.kind === "chart") return VISUAL_TYPES.UNKNOWN;
  if (isKPIViewModel(vm)) return VISUAL_TYPES.KPI;
  return VISUAL_TYPES.TABLE;
}
