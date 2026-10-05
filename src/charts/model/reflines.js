/* Analytics → Reference Line on a chart spec, valued, labelled and styled like Tableau. */
import { tableauToExcelNumFmt } from "../../format/number-format.js";
import { tfSameField } from "../../twb/field-ref.js";
import { tvMeasureLabel } from "./common.js";

const COMPUTATION = { average: "Average", mean: "Average", median: "Median", sum: "Sum", total: "Total",
                      min: "Minimum", max: "Maximum", constant: "Constant" };

/* Excel number-format literal: "text" (quotes inside dropped) */
const lit = s => s ? `"${String(s).replace(/"/g, "")}"` : "";

/** the line's value over the marks of the chart; null = not computable
 * @param {ReferenceLine} rl @param {number[]} values @returns {number | null} */
export function tvReferenceValue(rl, values) {
  if (rl.formula === "constant") return typeof rl.value === "number" ? rl.value : null;
  if (!values.length) return null;
  switch (rl.formula) {
    case "average": case "mean": return values.reduce((a, b) => a + b, 0) / values.length;
    case "median": {
      const s = [...values].sort((a, b) => a - b), m = Math.floor(s.length / 2);
      return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    }
    case "sum": case "total": return values.reduce((a, b) => a + b, 0);
    case "min": return Math.min(...values);
    case "max": return Math.max(...values);
    default: return null;
  }
}

/**
 * The label as an Excel number format on the line's own value, so Excel prints it: "Average"
 * (automatic / computation), the value, or a custom text with <Value> / <Computation>.
 * @param {ReferenceLine} rl @param {string} valueFmt @returns {string | null}
 */
export function tvReferenceLabelFormat(rl, valueFmt) {
  const name = COMPUTATION[rl.formula] || "";
  const positive = (String(valueFmt || "General").match(/^((?:"[^"]*"|[^;])*)/) || [])[1] || "General";
  if (rl.labelType === "none") return null;
  if (rl.labelType === "value") return positive;
  if (rl.labelType === "custom" && rl.label) {
    const parts = rl.label.split(/(<Value>|<Computation>)/i);
    if (!parts.some(p => /^<value>$/i.test(p))) return lit(parts.map(p => /^<computation>$/i.test(p) ? name : p.replace(/<[^<>]*>/g, "")).join(""));
    return parts.map(p => /^<value>$/i.test(p) ? positive : /^<computation>$/i.test(p) ? lit(name) : lit(p.replace(/<[^<>]*>/g, ""))).join("");
  }
  // automatic / computation: Tableau labels the line with the computation's name
  return rl.formula === "constant" ? positive : lit(name);
}

/**
 * Adds the worksheet's reference lines (table / pane scope) on the value axis: a flat line series on a
 * vertical value axis, or spec.refLines for horizontal bars (the writer draws those across the bars).
 * Lines per cell are Tableau's per-mark ticks – not drawn.
 * @param {ChartSpec} spec @param {ChartContext} ctx
 */
export function tvApplyReferenceLines(spec, ctx) {
  const { vm, roles } = ctx;
  const fmt = vm.fmt;
  if (!fmt.hasModel || !fmt.referenceLines || spec.boxPlot || !/^(bar|line|area|combo)$/.test(spec.kind) || !spec.categories) return;
  const lines = fmt.referenceLines().filter(r => !r.scope || r.scope === "per-table" || r.scope === "per-pane");
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!lines.length || !valueShelf) return;
  const measures = roles[valueShelf].values.map(v => v.ref).filter(Boolean);
  const drawn = spec.series.filter(s => !s.refLine && s.color !== null);                   // not invisible helper series
  const numbers = list => list.flatMap(s => s.values || []).filter(v => typeof v === "number" && isFinite(v));
  /* the series a line is computed over: its measure's (value-column, else the axis it sits on) –
     a dual-axis chart has one per measure; without a match, the primary axis's series */
  const seriesOf = rl => {
    const ref = rl.field || rl.axis;
    const m = ref && roles[valueShelf].values.find(v => v.ref && tfSameField(v.ref, ref));
    if (m) {
      const label = tvMeasureLabel(vm, m.ci).trim();
      const own = drawn.filter(s => { const n = s.name.trim(); return n === label || n.startsWith(label + " – "); });
      if (own.length) return own;
    }
    return drawn.filter(s => !s.secondary);
  };
  const style = fmt.reflineStyle();
  const hidden = style.lineVisible === false || style.strokeSize === 0;
  const valueFmt = (style.numFmtRaw && tableauToExcelNumFmt(style.numFmtRaw)) || spec.numFmt || "General";
  lines.forEach(rl => {
    if (rl.axis && measures.length && !measures.some(m => tfSameField(m, rl.axis))) return;   // another measure's axis
    const data = seriesOf(rl);
    const value = tvReferenceValue(rl, numbers(data));
    if (value === null) return;
    const labelFmt = tvReferenceLabelFormat(rl, valueFmt);
    if (hidden && !labelFmt) return;
    /** @type {RefLineStyle} */
    const line = { value, labelFmt, color: style.strokeColor ? style.strokeColor.slice(2) : "7F7F7F", alpha: style.strokeAlpha ?? 1,
                   width: style.strokeSize || 1, dash: style.dash === "dashed", hidden,
                   font: { color: style.color ? style.color.slice(2) : undefined, bold: style.bold } };
    if (spec.barDir === "bar" && data.some(s => (s.type || spec.kind) === "bar")) (spec.refLines = spec.refLines || []).push(line);
    // drawn on its measure's axis: on a dual-axis chart that may be the secondary one
    else spec.series.push({ name: (labelFmt || "Reference line").replace(/"/g, ""), type: "line", color: line.color, line: !hidden,
                            ...(data.length && data.every(s => s.secondary) ? { secondary: true } : {}),
                            marker: false, labels: false, values: spec.categories.levels[0].map(() => value), refLine: line });
  });
}
