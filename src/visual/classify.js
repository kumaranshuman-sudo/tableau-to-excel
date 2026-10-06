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
  if (info.bin || /\(bin\)$/i.test(ref.name)) return false;           // a bin draws the histogram's category axis
  return !(info.role === "dimension" && /^date/i.test(info.datatype || ""));
}

/** a measure that is one number (MIN(0), AVG(1)): it only places marks or labels @param {FormatModel | null} model @param {FieldRef} ref */
export function tvIsConstantRef(model, ref) {
  const info = ref && tfFieldInfo(model, ref);
  return !!info && info.constant !== undefined && !info.param;
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

/** Tableau's Automatic mark for the view's shelves: what an Automatic marks card draws
 * @param {ReturnType<typeof visualShelfShape>} shape @returns {string} */
export function tvAutomaticMark(shape) {
  return !shape ? "" : shape.geo ? "map" : shape.rowMeasures && shape.colMeasures ? "circle"
    : shape.rowMeasures || shape.colMeasures ? (shape.binned ? "bar" : shape.continuousDimension || shape.dateDimension ? "line" : "bar")
    : shape.shelfFields === 0 && shape.size ? "square"                    // empty shelves + Size → treemap
    : shape.colorMeasure && !shape.text ? "square"                         // dimensions only, a measure on Color → heat map
    : "text";
}

/* mark type: live visual spec (current state) → TWB pane marks → Tableau's "Automatic" rules */
/** @param {any} spec live visual spec @param {ViewModel} vm @param {FormatModel | null} model */
export function resolveVisualMarks(spec, vm, model) {
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
export function isRoundedBar(vm, model) {
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

/**
 * @param {any} spec live visual spec @param {ViewModel} vm @param {FormatModel | null} model
 * @param {string[]} [ev] collects the evidence behind the decision (shown in the conversion report)
 * @returns {VisualType}
 */
export function classifyVisualType(spec, vm, model, ev = []) {
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
