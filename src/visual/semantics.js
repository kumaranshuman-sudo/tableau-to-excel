/* What a Tableau visual is, in Tableau's terms, and the evidence for it: the shelves, marks and encodings of the
 * worksheet read as one visual ("Stacked Bar", "Bump Chart", "Filled Map" …). The conversion report shows it; the
 * classifier (classify.js) decides how the visual is drawn, this names what it is. */
import { TF_DERIV_LABEL, VISUAL_TYPES } from "../config.js";
import { tfFieldInfo } from "../twb/formatter.js";
import { tfSameField } from "../twb/field-ref.js";
import { TV_DATE_DERIVS, tvAutomaticMark, tvIsMeasureRef, visualShelfShape } from "./classify.js";

/* quick table calculations, as their derivation prefix ("pcto:sum:Sales:qk") */
const TABLE_CALCS = { pcto: "% of Total", cum: "Running Total", rsum: "Running Total", rank: "Rank", pctrank: "Percentile",
                      diff: "Difference", pdiff: "% Difference", mavg: "Moving Average", ytd: "YTD Total",
                      cgr: "Compound Growth Rate", yoy: "Year over Year Growth", ytdgr: "YTD Growth" };

const GEO_GENERATED = /^(latitude|longitude) \(generated\)$/i;

/**
 * A field reference taken apart: its quick table calculations (outer → inner), its aggregation or date part, the
 * underlying field and how Tableau labels it ("% of Total of SUM(Sales)").
 * @param {FormatModel | null} model @param {FieldRef} ref
 */
export function describeField(model, ref) {
  if (ref.name === "Measure Names") return { label: "Measure Names", calcs: [], agg: null, info: null, name: ref.name };
  if (ref.name === "Multiple Values") return { label: "Measure Values", calcs: [], agg: null, info: null, name: ref.name };
  const chain = ref.deriv ? [ref.deriv.toLowerCase()] : [];
  let name = ref.name || "";
  for (let m = name.match(/^([a-z_]+):(.+)$/i); m && (TABLE_CALCS[m[1].toLowerCase()] || TF_DERIV_LABEL[m[1].toLowerCase()] ||
       /^(none|usr)$/i.test(m[1])); m = name.match(/^([a-z_]+):(.+)$/i)) {
    chain.push(m[1].toLowerCase());
    name = m[2];
  }
  const calcs = chain.filter(c => TABLE_CALCS[c]);
  const agg = chain.find(c => !TABLE_CALCS[c]) || null;
  const info = model ? tfFieldInfo(model, { ds: ref.ds, deriv: null, name, inner: name, type: ref.type }) : null;
  const caption = (info && info.caption) || (/^Calculation_\d+$/i.test(name) ? "Calculation" : name.replace(/_/g, " "));
  const AGG = !agg || agg === "none" ? null : agg === "usr" ? "AGG" : (TF_DERIV_LABEL[agg] || agg.toUpperCase());
  let label = AGG ? `${AGG}(${caption})` : caption;
  if (calcs.length) label = calcs.map(c => TABLE_CALCS[c]).join(" of ") + " of " + label;
  return { label, calcs, agg, info, name };
}

/** @param {Pane} p @returns {string} the pane's mark ("" = Automatic) */
const markOf = p => String(p.markClass || "").toLowerCase().replace(/[\s_-]+/g, "").replace(/^automatic$/, "");

/**
 * @param {VisualModel} visualModel
 * @param {FormatModel | null} model
 * @returns {VisualSemantics}
 */
export function describeVisual(visualModel, model) {
  const vm = visualModel.viewModel;
  const sheet = vm && vm.fmt ? vm.fmt.sheetModel : null;
  const type = visualModel.type;
  const evidence = [...(visualModel.evidence || [])];
  const source = visualModel.metadata.markSource;
  const confidence = sheet && source === "twb" ? "high" : sheet || source === "live" ? "medium" : "low";
  /** @type {VisualSemantics} */
  const out = { visual: "", family: type, confidence, evidence, encodings: {}, calculations: [], flags: {} };
  if (!sheet) {
    out.visual = { TABLE: "Text Table", KPI: "KPI", CUSTOM: "Viz Extension" }[type] || titleCase(type);
    if (!visualModel.formatModel) evidence.push("no workbook loaded: recognised from the live data only");
    return out;
  }
  const label = r => describeField(model, r).label;
  const isMeasure = r => tvIsMeasureRef(model, r);
  const panes = sheet.panes.length > 1 && sheet.panes.some(p => p.id) ? sheet.panes.filter(p => p.id) : sheet.panes;
  // an Automatic card draws the All card's mark, else Tableau's automatic mark for the shelves
  const allPane = sheet.panes.find(p => !p.id);
  const inherit = (allPane && markOf(allPane)) || tvAutomaticMark(visualShelfShape(sheet, model));
  const marks = [...new Set(panes.map(p => markOf(p) || inherit))];
  const enc = ch => {
    const refs = [];
    panes.forEach(p => p.encodings.filter(e => e.channel === ch).forEach(e => { if (!refs.some(r => tfSameField(r, e.field))) refs.push(e.field); }));
    return refs;
  };
  const color = enc("color"), size = enc("size"), shape = enc("shape"), detail = enc("lod"), path = enc("path");
  const text = [...enc("text"), ...enc("label")];
  panes.forEach(p => p.labelRuns.forEach(r => r.refs.forEach(x => { if (!text.some(t => tfSameField(t, x))) text.push(x); })));
  const shelf = name => sheet[name].filter(r => r.name !== "__tableau_internal_object_id__");
  out.encodings = {
    rows: shelf("rows").map(label), columns: shelf("cols").map(label), color: color.map(label), size: size.map(label),
    shape: shape.map(label), label: text.map(label), detail: detail.map(label), path: path.map(label),
    angle: enc("wedge-size").map(label), tooltip: enc("tooltip").map(label)
  };
  Object.keys(out.encodings).forEach(k => { if (!out.encodings[k].length) delete out.encodings[k]; });

  // calculated fields behind the view: Tableau computed them – the export writes their values
  const seen = new Set();
  [...sheet.rows, ...sheet.cols, ...panes.flatMap(p => p.encodings.map(e => e.field))].forEach(r => {
    const d = describeField(model, r);
    if (d.calcs.length && !seen.has(d.label)) { seen.add(d.label); out.calculations.push({ name: d.label, kind: "quick table calculation" }); }
    if (d.info && d.info.calcKind && !d.info.param && !seen.has(d.name)) {
      seen.add(d.name);
      out.calculations.push({ name: d.info.caption || d.name, kind: d.info.calcKind });
    }
  });

  const rowsM = sheet.rows.filter(isMeasure), colsM = sheet.cols.filter(isMeasure);
  const horizontal = colsM.length > 0 && rowsM.length === 0;
  const valueRefs = horizontal ? colsM : rowsM;
  const catShelf = horizontal ? "rows" : "cols";
  const dimsOf = s => sheet[s].filter(r => !isMeasure(r) && r.name !== "Measure Names" && !GEO_GENERATED.test(r.name) &&
                                        r.name !== "__tableau_internal_object_id__");
  const calcsOf = r => describeField(model, r).calcs;
  const colorDim = color.find(r => !isMeasure(r) && r.name !== "Measure Names");
  const colorMeasure = color.find(r => isMeasure(r));
  const constant = r => { const d = describeField(model, r); return !!d.info && d.info.constant !== undefined; };
  const negated = r => { const d = describeField(model, r); return !!d.info && /^\s*-/.test(d.info.formula || ""); };
  const spaces = (sheet.style && sheet.style.spaces) || [];
  const flags = out.flags;
  // Dual Axis: an axis folded onto the one before it (the TWB's fold); measures without it get panes of their own
  flags.dualAxis = spaces.some(s => s.fold);
  flags.separatePanes = !flags.dualAxis && (rowsM.length > 1 || colsM.length > 1);
  flags.percentOfTotal = valueRefs.some(r => calcsOf(r).includes("pcto"));
  flags.runningTotal = valueRefs.some(r => calcsOf(r).includes("cum") || calcsOf(r).includes("rsum"));
  flags.rank = valueRefs.some(r => calcsOf(r).includes("rank"));
  flags.reversedAxis = spaces.some(s => s.reverse);
  flags.constantAxis = valueRefs.some(constant);
  flags.smallMultiples = dimsOf(horizontal ? "cols" : "rows").length > 0 && valueRefs.length > 0;
  flags.detail = detail.some(r => !isMeasure(r));
  if (flags.dualAxis) evidence.push(`dual axis${spaces.some(s => s.fold && s.synchronized) ? " (synchronized)" : ""}`);
  if (flags.separatePanes) evidence.push("several measures, each in its own pane");
  if (flags.percentOfTotal) evidence.push("percent-of-total table calculation on the value axis");
  if (flags.runningTotal) evidence.push("running-total table calculation on the value axis");
  if (flags.rank) evidence.push("rank table calculation on the value axis");
  if (flags.reversedAxis) evidence.push("reversed axis");
  if (flags.smallMultiples) evidence.push("a dimension on the value shelf: one pane per member (small multiples)");
  if (colorDim) evidence.push(`colour by the dimension ${label(colorDim)}`);
  else if (colorMeasure) evidence.push(`colour by the measure ${label(colorMeasure)}`);
  if (size.length) evidence.push(`size by ${size.map(label).join(", ")}`);
  if (path.length) evidence.push(`path by ${path.map(label).join(", ")}`);
  if (flags.detail) evidence.push(`detail: ${detail.filter(r => !isMeasure(r)).map(label).join(", ")}`);

  const refLines = sheet.referenceLines || [];
  const bulletTarget = refLines.some(rl => rl.scope === "per-cell" && rl.field && rl.axis && !tfSameField(rl.field, rl.axis));
  const dateParts = [...sheet.rows, ...sheet.cols].map(r => String(r.deriv || "").toLowerCase());
  const has = m => marks.includes(m);
  const shelfDims = [...dimsOf("rows"), ...dimsOf("cols")];
  const bars = horizontal ? "Bar" : "Column";
  const name = () => {
    switch (type) {
      case VISUAL_TYPES.MAP: case VISUAL_TYPES.MAP_FILLED: {
        const geoDims = [...detail, ...shelfDims, ...color, ...text].filter(r => !isMeasure(r) && !GEO_GENERATED.test(r.name) &&
                                                                             !!(describeField(model, r).info || {}).geoRole);
        flags.geographic = true;
        if (has("density") || has("heatmap")) return "Density Map";
        if (!geoDims.length && has("pie")) return "Donut / Pie built on map layers";
        if (!geoDims.length) { evidence.push("no geographic field: map layers used as a canvas"); return "Map with Layers"; }
        if (has("pie")) return "Pies on a Map";
        if (panes.length > 1 && marks.length > 1) return "Map with Layers";
        if (has("line") && path.length) return "Flow Map";
        if (has("polygon")) return "Polygon Map";
        if (has("multipolygon") || type === VISUAL_TYPES.MAP_FILLED) return "Filled Map";
        return "Symbol Map";
      }
      case VISUAL_TYPES.PIE:
        if (visualModel.pie && (visualModel.pie.inner || (visualModel.pie.layers || 0) > 1)) return "Nested Donut";
        if ((visualModel.pie && visualModel.pie.hole) || (vm.fmt.pieLayers && (vm.fmt.pieLayers() || {}).hole)) return "Donut Chart";
        return shelfDims.length ? "Multiple Pies" : "Pie Chart";
      case VISUAL_TYPES.HEATMAP:
        if (dateParts.includes("wk") && (dateParts.includes("wd") || dateParts.includes("dy"))) return "Calendar Heat Map";
        return text.some(isMeasure) ? "Highlight Table" : "Heat Map";
      case VISUAL_TYPES.TABLE:
        if (vm.markTable) return "Table built from marks";
        if (has("shape") && dimsOf("rows").length && dimsOf("cols").length) return "KPI Indicator Matrix";
        if (has("circle") && size.length && shelfDims.some(r => r.type === "qk")) return "Circle Timeline";
        return dimsOf("rows").length && dimsOf("cols").length ? "Crosstab" : "Text Table";
      case VISUAL_TYPES.KPI: return "KPI (BAN)";
      case VISUAL_TYPES.HISTOGRAM: return "Histogram";
      case VISUAL_TYPES.WATERFALL: return "Waterfall Chart";
      case VISUAL_TYPES.BOXPLOT: return "Box Plot";
      case VISUAL_TYPES.GANTT: return "Gantt Chart";
      case VISUAL_TYPES.TREEMAP: return "Treemap";
      case VISUAL_TYPES.BUBBLE: return "Packed Bubbles";
      case VISUAL_TYPES.CUSTOM: return "Viz Extension";
      case VISUAL_TYPES.SCATTER:
        if (size.length) return "Scatter Plot with Size (bubble)";
        if (refLines.length >= 2 || (colorMeasure || colorDim) && refLines.length) return "Quadrant Chart";
        return "Scatter Plot";
      case VISUAL_TYPES.COMBO:
        if (valueRefs.some(r => calcsOf(r).includes("pcto") && calcsOf(r).includes("cum"))) return "Pareto Chart";
        return `Combination Chart (${marks.filter(Boolean).join(" + ")}, dual axis)`;
      case VISUAL_TYPES.AREA:
        if (flags.percentOfTotal && colorDim) return "100% Stacked Area";
        if (colorDim) return "Stacked Area";
        if (flags.smallMultiples) return "Small Multiple Areas";
        return "Area Chart";
      case VISUAL_TYPES.LINE:
        if (path.some(r => r.name === "Measure Names") && valueRefs.some(r => r.name === "Multiple Values")) {
          return has("circle") ? "Barbell (Dumbbell) Chart" : flags.constantAxis || hasConstantMeasure(model, sheet) ? "Rounded Bar Chart" : "Line Chart";
        }
        if (rowsM.length && colsM.length) return flags.runningTotal && flags.percentOfTotal ? "Pareto Curve (cumulative % against cumulative %)" : "Line (measure against measure)";
        if (flags.rank) return "Bump Chart";
        if (!has("line") && marks.some(m => /^(circle|shape|square)$/.test(m))) return "Dot Plot";
        if (valueRefs.length && valueRefs.every(r => r.name === valueRefs[0].name) && valueRefs.length === 2 && has("circle") &&
            sheet[catShelf].some(r => /^(yr|tyr)$/i.test(r.deriv || "") && r.type !== "qk")) return "Slope Chart";
        if (flags.runningTotal) return "Cumulative Line Chart";
        if (flags.smallMultiples) return "Small Multiple Lines (sparklines)";
        if (flags.dualAxis && valueRefs.length > 1 && !valueRefs.every(r => tfSameField(r, valueRefs[0]))) return "Dual-axis Line Chart";
        if (flags.separatePanes && valueRefs.length > 1) return "Line Charts in separate panes (one per measure)";
        if (colorDim || (flags.detail && has("line"))) return "Multi-line Chart";
        return "Line Chart";
      case VISUAL_TYPES.BAR: {
        if (path.some(r => r.name === "Measure Names")) return "Rounded Bar Chart";
        if (bulletTarget) return "Bullet Graph";
        if (valueRefs.some(negated) || (valueRefs.some(r => r.name === "Multiple Values") && sheet.fieldRefs.some(negated))) return "Butterfly (Diverging) Bar Chart";
        if (flags.percentOfTotal && valueRefs.some(constant)) return "Progress Bar";
        if (has("circle") && has("bar")) return "Lollipop Chart";
        if (flags.percentOfTotal && colorDim) return `100% Stacked ${bars}`;
        if (flags.smallMultiples) return `Small Multiple ${bars}s`;
        if (flags.dualAxis && valueRefs.length > 1 && marks.every(m => m === "bar")) return "Bar in Bar";
        if (flags.dualAxis && valueRefs.length > 1) return `Combination Chart (${marks.filter(Boolean).join(" + ")}, dual axis)`;
        if (flags.separatePanes && valueRefs.length > 1) return `${bars} Charts in separate panes (one per measure)`;
        const catDimRefs = dimsOf(catShelf);
        if (colorDim && !catDimRefs.some(r => tfSameField(r, colorDim)) && !sheet[horizontal ? "cols" : "rows"].some(r => tfSameField(r, colorDim))) return `Stacked ${bars}`;
        // side by side: the colour field is the innermost of several category fields (bars grouped by the outer one)
        if (colorDim && catDimRefs.length > 1 && tfSameField(catDimRefs[catDimRefs.length - 1], colorDim)) return `Side-by-side (grouped) ${bars}`;
        if (colorDim) return `${bars} Chart coloured by ${label(colorDim)}`;
        if (flags.detail) return `Stacked ${bars} (by detail)`;
        if (refLines.length) return `${bars} Chart with Reference Line`;
        if (dateParts.some(d => TV_DATE_DERIVS.has(d))) return `${bars} Chart over Time`;
        return horizontal ? "Horizontal Bar Chart" : "Column Chart";
      }
      default: return titleCase(type);
    }
  };
  out.visual = name();
  return out;
}

/** a measure on the sheet that is a constant (MIN(0), AVG(1)): it only places the marks */
function hasConstantMeasure(model, sheet) {
  return sheet.fieldRefs.some(r => { const d = describeField(model, r); return !!d.info && d.info.constant !== undefined && !d.info.param; });
}

function titleCase(t) { return String(t || "").toLowerCase().replace(/_/g, " ").replace(/\b\w/g, c => c.toUpperCase()); }
