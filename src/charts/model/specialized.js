/* Pie, scatter, waterfall, box plot, Gantt and treemap chart specs. */
import { TV_MAX_POINTS, TV_MAX_SERIES, tvBaseSpec, tvCategories, tvColorScale, tvColorValues, tvHex, tvLabelsOn, tvMarkColor, tvMeasureLabel, tvNumFmt, tvSum, tvText } from "./common.js";
import { tfDvNum, tfIsNull } from "../../data/values.js";
import { TABLEAU_10 } from "../../format/palettes.js";
import { tfSameField } from "../../twb/field-ref.js";

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvPieSpec(ctx) {
  const { vm, roles } = ctx;
  let catCi = roles.color && !roles.color.measureNames && !roles.color.continuous ? roles.color.ci : -1;
  if (catCi < 0) {
    const dims = [...roles.rows.dims, ...roles.cols.dims];
    catCi = dims.length ? dims[0].ci : vm.cols.findIndex(c => c.isHeader);
  }
  let angleCi = roles.angle;
  if (angleCi < 0) {
    const v = [...roles.rows.values, ...roles.cols.values][0];
    angleCi = v ? v.ci : vm.cols.findIndex(c => !c.isHeader && c.name !== (vm.cols[catCi] || {}).name);
  }
  const scale = tvColorScale(vm, roles, "pie");
  const spec = /** @type {ChartSpec} */ ({ ...tvBaseSpec(vm), kind: ctx.doughnut ? "doughnut" : "pie", legend: true });
  let categories, values;
  if ((roles.color && roles.color.measureNames) || (catCi < 0 && vm.cols.some(c => c.pivoted))) {
    const measures = vm.cols.map((c, i) => c.pivoted ? i : -1).filter(i => i >= 0);
    categories = { names: ["Measure Names"], levels: [measures.map(ci => tvMeasureLabel(vm, ci))] };
    values = measures.map(ci => vm.rows.reduce((s, r) => s + (tfDvNum(r[ci]) || 0), 0));
    angleCi = measures[0];
  } else {
    if (catCi < 0 || angleCi < 0) throw new Error("pie has no colour dimension or angle measure");
    const cats = tvCategories(vm, [catCi], false);
    categories = { names: [tvMeasureLabel(vm, catCi)], levels: cats.levels };
    values = tvSum(vm, cats, angleCi);
  }
  // a dual-pie donut's hole pane contributes a row with no category and no value → not a slice
  const keep = categories.levels[0].map((label, i) => !(label === "Null" && !values[i]));
  if (keep.includes(false)) {
    categories = { ...categories, levels: categories.levels.map(lv => lv.filter((_, i) => keep[i])) };
    values = values.filter((_, i) => keep[i]);
  }
  if (categories.levels[0].length > 200) throw new Error("too many pie slices for an Excel chart");
  const angleRef = vm.cols[angleCi] && vm.cols[angleCi].ref;
  const catRef = catCi >= 0 && vm.cols[catCi].ref;
  // a "% of Total" quick table calc on Label → Excel's percentage label instead of the raw value
  const labelPct = roles.labelRefs.some(r => /^pcto$/i.test(r.deriv || ""));
  const labelValue = !!angleRef && roles.labelRefs.some(r => !/^pcto$/i.test(r.deriv || "") && tfSameField(r, angleRef)) ||
    (!labelPct && tvLabelsOn(roles, angleRef));
  const labelCat = !!catRef && roles.labelRefs.some(r => tfSameField(r, catRef));
  spec.categories = categories;
  spec.numFmt = tvNumFmt(vm, angleCi);
  spec.series = [{
    name: tvMeasureLabel(vm, angleCi),
    color: tvMarkColor(roles),
    values,
    pointColors: categories.levels[0].map((v, i) => (scale && scale(v, i)) || tvHex(TABLEAU_10[i % TABLEAU_10.length])),
    labels: labelValue || labelCat || labelPct,
    labelParts: { value: labelValue, category: labelCat, percent: labelPct },
    labelNumFmt: labelPct && !labelValue ? "0.0%" : undefined
  }];
  return [spec];
}

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvScatterSpec(ctx) {
  const { vm, roles } = ctx;
  const xm = roles.cols.values[0], ym = roles.rows.values[0];
  const scale = tvColorScale(vm, roles, ctx.markToken);
  const spec = /** @type {ChartSpec} */ ({ ...tvBaseSpec(vm), kind: "scatter", numFmt: tvNumFmt(vm, ym.ci), xNumFmt: tvNumFmt(vm, xm.ci),
                 valueTitle: tvMeasureLabel(vm, ym.ci), xTitle: tvMeasureLabel(vm, xm.ci) });
  const point = r => ({ x: tfDvNum(r[xm.ci]), y: tfDvNum(r[ym.ci]) });
  const rows = vm.rows.filter(r => { const p = point(r); return p.x !== null && p.y !== null; });
  if (rows.length > TV_MAX_POINTS) throw new Error(`${rows.length} marks – too many for an Excel scatter chart`);
  const colorCi = roles.color && !roles.color.measureNames ? roles.color.ci : -1;
  // Label = a dimension (store name …) → Excel "value from cells" labels with that text, when
  // few enough marks to stay readable; numbers only when a measure itself is on Label
  const labelDim = roles.textDims[0] ?? -1;
  const measureLabel = roles.labelRefs.some(r => tfSameField(r, ym.ref) || tfSameField(r, xm.ref));
  const textLabels = labelDim >= 0 && rows.length <= 40;
  const labels = textLabels || measureLabel || (labelDim < 0 && tvLabelsOn(roles, ym.ref));
  const labelOf = r => tvText(r[labelDim]);
  if (colorCi >= 0 && !roles.color.continuous) {
    const values = tvColorValues(vm, roles);
    if (values.length > TV_MAX_SERIES) throw new Error("too many colour values for an Excel chart");
    spec.series = values.map((v, i) => {
      const pts = rows.filter(r => tvText(r[colorCi]) === v).map(point);
      return { name: v, color: (scale && scale(v)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
               x: pts.map(p => p.x), y: pts.map(p => p.y), labels,
               labelTexts: textLabels ? rows.filter(r => tvText(r[colorCi]) === v).map(labelOf) : undefined };
    });
  } else {
    const pts = rows.map(point);
    spec.series = [{
      name: tvMeasureLabel(vm, ym.ci), color: tvMarkColor(roles), x: pts.map(p => p.x), y: pts.map(p => p.y), labels,
      labelTexts: textLabels ? rows.map(labelOf) : undefined,
      pointColors: colorCi >= 0 && scale ? rows.map(r => scale(tfDvNum(r[colorCi]))) : undefined
    }];
  }
  spec.legend = spec.series.length > 1;
  return [spec];
}

/* value axis + category dimensions shared by waterfall / box plot */
/** @param {ChartContext} ctx */
export function tvAxisLayout(ctx) {
  const { roles } = ctx;
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) throw new Error("no continuous measure axis to chart");
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  return { valueShelf, horizontal: valueShelf === "cols", measure: roles[valueShelf].values[0],
           catDims: [...roles[valueShelf].dims, ...roles[catShelf].dims] };
}

/* ── waterfall: Gantt bars sized by -measure on a running total →
 *    stacked columns: invisible base + increase / decrease ─────────────────── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvWaterfallSpec(ctx) {
  const { vm, roles } = ctx;
  const { horizontal, measure, catDims } = tvAxisLayout(ctx);
  const cats = tvCategories(vm, catDims.map(d => d.ci), catDims.some(d => d.continuous));
  if (cats.count > TV_MAX_POINTS) throw new Error(`${cats.count} categories – too many for an Excel chart`);
  const running = tvSum(vm, cats, measure.ci);
  if (running.some(v => v !== null && v < 0)) throw new Error("running total goes below zero – not drawable as stacked columns");
  const base = [], up = [], down = [];
  let prev = 0;
  running.forEach(v => {
    const cur = v === null ? prev : v;
    const delta = cur - prev;
    base.push(Math.min(prev, cur));
    up.push(delta > 0 ? delta : null);
    down.push(delta < 0 ? -delta : null);
    prev = cur;
  });
  const markColor = tvMarkColor(roles);
  const scale = tvColorScale(vm, roles, "ganttbar");
  // colour = a measure (usually the step itself) → colour each step; else Tableau's single mark colour
  const stepColors = scale && roles.color && roles.color.continuous && roles.color.ci >= 0
    ? tvSum(vm, cats, roles.color.ci).map(v => scale(v)) : null;
  const labels = tvLabelsOn(roles, measure.ref);
  const label = tvMeasureLabel(vm, measure.ci);
  const numFmt = tvNumFmt(vm, measure.ci);
  // decreases are drawn as positive heights but labelled as the negative step, like Tableau
  const positive = (numFmt.match(/^((?:"[^"]*"|[^;])*)/) || [])[1] || "General";
  const downFmt = positive === "General" ? "-General" : "-" + positive;
  return [{
    ...tvBaseSpec(vm), kind: "bar", barDir: horizontal ? "bar" : "col", stacked: true, gapWidth: 30, legend: false,
    categories: { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [""], levels: cats.levels },
    categoryTitle: catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / "),
    numFmt, valueTitle: label,
    series: [
      { name: "Base", type: "bar", values: base, color: null, labels: false },
      { name: label + " (increase)", type: "bar", values: up, color: markColor, pointColors: stepColors, labels },
      { name: label + " (decrease)", type: "bar", values: down, color: markColor, pointColors: stepColors, labels, labelNumFmt: downFmt }
    ]
  }];
}

/* ── box plot: quartiles per category → line chart with up/down bars (box),
 *    high-low lines (whiskers) and a dash at the median ───────────────────── */
export function tvQuantile(sorted, p) {                      // linear interpolation, as Tableau / QUARTILE.INC
  const pos = (sorted.length - 1) * p, lo = Math.floor(pos), hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvBoxPlotSpec(ctx) {
  const { vm, roles } = ctx;
  const { measure, catDims } = tvAxisLayout(ctx);
  const cats = tvCategories(vm, catDims.map(d => d.ci), catDims.some(d => d.continuous));
  if (cats.count > 200) throw new Error("too many boxes for an Excel chart");
  const groups = Array.from({ length: cats.count }, () => []);
  vm.rows.forEach(r => {
    const n = tfDvNum(r[measure.ci]);
    const i = cats.indexOf(r);
    if (n !== null && i !== undefined) groups[i].push(n);
  });
  const stats = groups.map(g => {
    if (!g.length) return null;
    const s = g.slice().sort((a, b) => a - b);
    const q1 = tvQuantile(s, 0.25), med = tvQuantile(s, 0.5), q3 = tvQuantile(s, 0.75), iqr = q3 - q1;
    // standard whiskers: the furthest marks within 1.5 × IQR of the box
    const low = s.find(v => v >= q1 - 1.5 * iqr), high = [...s].reverse().find(v => v <= q3 + 1.5 * iqr);
    return { q1, med, q3, low, high };
  });
  const col = k => stats.map(st => st ? st[k] : null);
  const boxColor = tvMarkColor(roles);
  return [{
    ...tvBaseSpec(vm), kind: "line", boxPlot: { color: boxColor }, legend: false, includeZero: false,
    categories: { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [""], levels: cats.levels },
    categoryTitle: catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / "),
    numFmt: tvNumFmt(vm, measure.ci), valueTitle: tvMeasureLabel(vm, measure.ci),
    // up/down bars span the FIRST and LAST series → Q1 … Q3; high-low lines span all series
    series: [
      { name: "Lower quartile", type: "line", values: col("q1"), color: boxColor, line: false, marker: false },
      { name: "Lower whisker", type: "line", values: col("low"), color: boxColor, line: false, marker: false },
      { name: "Median", type: "line", values: col("med"), color: "333333", line: false, marker: true, markerSymbol: "dash", markerSize: 14 },
      { name: "Upper whisker", type: "line", values: col("high"), color: boxColor, line: false, marker: false },
      { name: "Upper quartile", type: "line", values: col("q3"), color: boxColor, line: false, marker: false }
    ]
  }];
}

/* ── Gantt: per row, marks start at the date axis value and last Size days →
 *    stacked horizontal bars of alternating invisible gaps and visible bars ─ */
export const TV_GANTT_MAX_SERIES = 250;

export function tvExcelSerial(dv) {
  if (tfIsNull(dv)) return null;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  if (typeof v === "number") return v;                         // already a serial / numeric axis
  const t = v instanceof Date ? v.getTime() : Date.parse(String(v));
  return isNaN(t) ? null : t / 86400000 + 25569;               // days since 1899-12-30 (UTC)
}

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvGanttSpec(ctx) {
  const { vm, roles } = ctx;
  const startDim = [...roles.cols.dims, ...roles.rows.dims].find(d => d.continuous);
  if (!startDim) throw new Error("no continuous date axis for the Gantt bars");
  if (roles.size < 0) throw new Error("no Size measure for the Gantt bar length");
  const startShelf = roles.cols.dims.includes(startDim) ? "cols" : "rows";
  const catDims = roles[startShelf === "cols" ? "rows" : "cols"].dims.filter(d => !d.continuous);
  const cats = tvCategories(vm, catDims.map(d => d.ci), false);
  const colorScale = tvColorScale(vm, roles, "ganttbar");
  const colorCi = roles.color && !roles.color.measureNames ? roles.color.ci : -1;
  const markColor = tvMarkColor(roles);
  const marks = Array.from({ length: cats.count }, () => []);
  vm.rows.forEach(r => {
    const start = tvExcelSerial(r[startDim.ci]), len = tfDvNum(r[roles.size]), i = cats.indexOf(r);
    if (start === null || len === null || i === undefined) return;
    const c = colorCi >= 0 && colorScale ? colorScale(roles.color.continuous ? tfDvNum(r[colorCi]) : tvText(r[colorCi])) : null;
    marks[i].push({ start, end: start + Math.max(0, len), color: c || markColor });
  });
  const all = marks.flat();
  if (!all.length) throw new Error("no Gantt marks with a start date and a length");
  const axisMin = Math.floor(Math.min(...all.map(m => m.start)));
  const axisMax = Math.ceil(Math.max(...all.map(m => m.end)));
  const depth = Math.max(...marks.map(m => m.length));
  if (2 * depth > TV_GANTT_MAX_SERIES) throw new Error(`${depth} bars in one row – too many for an Excel chart`);
  /** @type {ChartSeries[]} */
  const series = [];
  for (let k = 0; k < depth; k++) {
    series.push({ name: `Gap ${k + 1}`, type: "bar", values: [], color: null, labels: false });
    series.push({ name: `Bar ${k + 1}`, type: "bar", values: [], color: markColor, pointColors: [], labels: false });
  }
  marks.forEach((row, i) => {
    row.sort((a, b) => a.start - b.start);
    let edge = 0;
    row.forEach((m, k) => {
      if (m.start < edge - 1e-9) throw new Error("overlapping Gantt bars in one row – not drawable as stacked bars");
      series[2 * k].values[i] = k === 0 ? m.start : m.start - edge;
      series[2 * k + 1].values[i] = m.end - m.start;
      series[2 * k + 1].pointColors[i] = m.color;
      edge = m.end;
    });
    for (let k = row.length; k < depth; k++) { series[2 * k].values[i] = null; series[2 * k + 1].values[i] = null; }
  });
  const dateFmt = axisMax - axisMin > 400 ? "mmm yyyy" : "d mmm yyyy";
  return [{
    ...tvBaseSpec(vm), kind: "bar", barDir: "bar", stacked: true, gapWidth: 40, legend: false,
    valueMin: axisMin, valueMax: axisMax, includeZero: false,
    categories: { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [""], levels: cats.levels },
    categoryTitle: catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / "),
    numFmt: dateFmt, valueTitle: tvMeasureLabel(vm, startDim.ci), series
  }];
}

/* ── treemap: nothing on Rows/Columns, Size = measure, Label/Detail = the
 *    hierarchy → Excel treemap (chartex), each tile in its Tableau colour ── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvTreemapSpec(ctx) {
  const { vm, roles } = ctx;
  if (roles.size < 0) throw new Error("no Size measure for the treemap");
  let levelCis = [...roles.detailDims, ...roles.textDims].filter((ci, i, a) => a.indexOf(ci) === i);
  if (roles.color && !roles.color.continuous && roles.color.ci >= 0 && !levelCis.includes(roles.color.ci)) levelCis.unshift(roles.color.ci);
  if (!levelCis.length) levelCis = vm.cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);
  if (!levelCis.length) throw new Error("no dimension to split the treemap");
  // outermost level = the one with the fewest distinct values
  const distinct = ci => new Set(vm.rows.map(r => tvText(r[ci]))).size;
  levelCis.sort((a, b) => distinct(a) - distinct(b));
  const cats = tvCategories(vm, levelCis, false);
  if (cats.count > 1000) throw new Error(`${cats.count} tiles – too many for an Excel treemap`);
  const sizes = tvSum(vm, cats, roles.size);
  const scale = tvColorScale(vm, roles, "square");
  let pointColors;
  if (scale && roles.color.continuous) pointColors = tvSum(vm, cats, roles.color.ci).map(v => scale(v));
  else if (scale && roles.color.ci >= 0) {
    const level = levelCis.indexOf(roles.color.ci);
    pointColors = cats.levels[level].map((v, i) => scale(v) || tvHex(TABLEAU_10[i % TABLEAU_10.length]));
  }
  // drop empty / negative tiles: a treemap cannot draw them
  const keep = sizes.map((v, i) => v !== null && v > 0 ? i : -1).filter(i => i >= 0);
  // Excel builds the hierarchy from consecutive rows: keep each parent's tiles together,
  // biggest parent first, biggest tile first inside it (Tableau's treemap order)
  const pathKey = (i, l) => cats.levels.slice(0, l + 1).map(lv => lv[i]).join("\u0001");
  const totals = cats.levels.map((_, l) => {
    const m = new Map();
    keep.forEach(i => m.set(pathKey(i, l), (m.get(pathKey(i, l)) || 0) + sizes[i]));
    return m;
  });
  keep.sort((a, b) => {
    for (let l = 0; l < cats.levels.length; l++) {
      const ka = pathKey(a, l), kb = pathKey(b, l);
      if (ka === kb) continue;
      return (totals[l].get(kb) - totals[l].get(ka)) || (ka < kb ? -1 : 1);
    }
    return 0;
  });
  return [{
    ...tvBaseSpec(vm), kind: "treemap", legend: false,
    categories: { names: levelCis.map(ci => tvMeasureLabel(vm, ci)), levels: cats.levels.map(lv => keep.map(i => lv[i])) },
    numFmt: tvNumFmt(vm, roles.size),
    series: [{ name: tvMeasureLabel(vm, roles.size), values: keep.map(i => sizes[i]), color: tvMarkColor(roles),
               pointColors: pointColors ? keep.map(i => pointColors[i]) : undefined, labels: true,
               // Size measure also on Label → "Quality Issue 20K", like Tableau
               labelParts: { category: true, value: roles.labelRefs.some(r => tfSameField(r, vm.cols[roles.size].ref || {})) } }]
  }];
}
