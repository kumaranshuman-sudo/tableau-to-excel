/* Pie, scatter, waterfall, box plot, Gantt, treemap, symbol map and packed bubble chart specs. */
import { TV_MAX_POINTS, TV_MAX_SERIES, tvBaseSpec, tvCategories, tvColorScale, tvColorValues, tvHex, tvLabelsOn, tvMarkColor, tvMarkToken, tvMeasureLabel, tvNumFmt, tvPaneLabel, tvPaneRule, tvSum, tvText } from "./common.js";
import { tfDvNum, tfDvText, tfIsNull } from "../../data/values.js";
import { tfFormatNumber } from "../../format/number-format.js";
import { TABLEAU_10 } from "../../format/palettes.js";
import { tfSameField } from "../../twb/field-ref.js";
import { tvPackCircles } from "./packing.js";
import { renderTemplate } from "./pie.js";

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
  // Tableau's donut: a pie of slices (colour / angle) and, on the second axis, a smaller plain pie – the hole,
  // in the background colour, carrying the total as its label
  const piePanes = roles.panes.filter(p => tvMarkToken(p.markClass) === "pie");
  const slicing = p => p.encodings.some(e => e.channel === "wedge-size" || e.channel === "color");
  const slicePane = ctx.doughnut ? piePanes.find(slicing) || null : null;
  const holePane = ctx.doughnut ? piePanes.find(p => p.id && !slicing(p)) || null : null;
  let categories, values, cats = null;
  if ((roles.color && roles.color.measureNames) || (catCi < 0 && vm.cols.some(c => c.pivoted))) {
    const measures = vm.cols.map((c, i) => c.pivoted ? i : -1).filter(i => i >= 0);
    categories = { names: ["Measure Names"], levels: [measures.map(ci => tvMeasureLabel(vm, ci))] };
    values = measures.map(ci => vm.rows.reduce((s, r) => s + (tfDvNum(r[ci]) || 0), 0));
    angleCi = measures[0];
  } else {
    if (catCi < 0 || angleCi < 0) throw new Error("pie has no colour dimension or angle measure");
    cats = tvCategories(vm, [catCi], false);
    categories = { names: [tvMeasureLabel(vm, catCi)], levels: cats.levels };
    values = tvSum(vm, cats, angleCi);
  }
  // a dual-pie donut's hole contributes a mark with no category: no value, or the total of the slices → not a slice
  const total = values.reduce((s, v) => s + (v || 0), 0);
  const nullish = l => /^(null|%null%|\(null\))?$/i.test(String(l == null ? "" : l).trim());
  const isHole = i => nullish(categories.levels[0][i]) &&
    (!values[i] || (ctx.doughnut && Math.abs(values[i] - (total - values[i])) <= Math.abs(values[i]) * 0.005));
  const keep = categories.levels[0].map((_, i) => !isHole(i));
  const kept = keep.map((k, i) => k ? i : -1).filter(i => i >= 0);
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
  if (slicePane && cats) {
    // the field values as Tableau writes them (measures in their number format, empty for null)
    const text = (row, ref) => {
      const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
      if (ci < 0 || !row || tfIsNull(row[ci])) return "";
      const n = tfDvNum(row[ci]);
      return vm.cols[ci].isHeader || n === null ? tvText(row[ci]) : tfFormatNumber(n, tvNumFmt(vm, ci));
    };
    const rowsOf = new Map();
    vm.rows.forEach(r => { const i = cats.indexOf(r); if (i !== undefined && keep[i] && !rowsOf.has(i)) rowsOf.set(i, r); });
    const flat = lines => lines.map(l => l.map(s => s.text).join("")).filter(t => t.trim()).join("\n");
    // slice labels: only the slices' own marks card (blank until its label fields have values)
    const texts = kept.map(i => flat(tvPaneLabel(slicePane, ref => text(rowsOf.get(i), ref))));
    const labelled = texts.some(Boolean) && slicePane.style && tvPaneRule(slicePane, "mark", "mark-labels-show") !== "false";
    Object.assign(spec.series[0], { labels: labelled, labelTexts: labelled ? texts : undefined, labelParts: undefined, labelNumFmt: undefined });
    if (holePane) {
      // the hole's label: its fields over all the slices (the total), in its runs' fonts
      const sumOf = ref => {
        const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
        if (ci < 0) return "";
        if (vm.cols[ci].isHeader) return "";
        const s = vm.rows.reduce((acc, r) => { const i = cats.indexOf(r); return i !== undefined && keep[i] ? acc + (tfDvNum(r[ci]) || 0) : acc; }, 0);
        return tfFormatNumber(s, tvNumFmt(vm, ci));
      };
      spec.centerLabel = tvPaneLabel(holePane, sumOf).filter(l => l.some(s => s.text.trim())).map(l => l.map(s => ({
        text: s.text, bold: !!s.props.bold, size: s.props.fontSize, color: tvHex(s.props.color) || undefined, font: s.props.fontName })));
      // the hole's size against the ring's (the marks' Size sliders)
      const sizeOf = p => Number(tvPaneRule(p, "mark", "size"));
      const ratio = sizeOf(holePane) / sizeOf(slicePane);
      if (ratio > 0 && ratio < 1) spec.holeSize = Math.round(Math.max(10, Math.min(90, ratio * 100)));
    }
  }
  return [spec];
}

/**
 * @param {ChartContext} ctx
 * @param {{ lines?: boolean }} [opts] lines: a Line mark with a measure on both axes – the points joined in order of x
 * @returns {ChartSpec[]}
 */
export function tvScatterSpec(ctx, opts = {}) {
  const { vm, roles } = ctx;
  const xm = roles.cols.values[0], ym = roles.rows.values[0];
  const scale = tvColorScale(vm, roles, ctx.markToken);
  const spec = /** @type {ChartSpec} */ ({ ...tvBaseSpec(vm), kind: "scatter", numFmt: tvNumFmt(vm, ym.ci), xNumFmt: tvNumFmt(vm, xm.ci),
                 valueTitle: tvMeasureLabel(vm, ym.ci), xTitle: tvMeasureLabel(vm, xm.ci) });
  const point = r => ({ x: tfDvNum(r[xm.ci]), y: tfDvNum(r[ym.ci]) });
  const rows = vm.rows.filter(r => { const p = point(r); return p.x !== null && p.y !== null; });
  if (opts.lines) rows.sort((a, b) => point(a).x - point(b).x || point(a).y - point(b).y);
  if (rows.length > TV_MAX_POINTS) throw new Error(`${rows.length} marks – too many for an Excel scatter chart`);
  const colorCi = roles.color && !roles.color.measureNames ? roles.color.ci : -1;
  // Label = a dimension (store name …) → Excel "value from cells" labels with that text, when
  // few enough marks to stay readable; numbers only when a measure itself is on Label
  const labelDim = roles.textDims[0] ?? -1;
  const measureLabel = roles.labelRefs.some(r => tfSameField(r, ym.ref) || tfSameField(r, xm.ref));
  const textLabels = labelDim >= 0 && rows.length <= 40;
  const labels = textLabels || measureLabel || (labelDim < 0 && tvLabelsOn(roles, ym.ref));
  const labelOf = r => tvText(r[labelDim]);
  // a Size measure is part of the meaning (Tableau's circles sized by it): Excel's bubble chart, area = size
  const sizeCi = !opts.lines && roles.size >= 0 && !vm.cols[roles.size].isHeader ? roles.size : -1;
  const sizeOf = r => { const n = tfDvNum(r[sizeCi]); return n === null ? null : Math.abs(n); };
  if (colorCi >= 0 && !roles.color.continuous) {
    const values = tvColorValues(vm, roles);
    if (values.length > TV_MAX_SERIES) throw new Error("too many colour values for an Excel chart");
    spec.series = values.map((v, i) => {
      const own = rows.filter(r => tvText(r[colorCi]) === v), pts = own.map(point);
      return { name: v, color: (scale && scale(v)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
               x: pts.map(p => p.x), y: pts.map(p => p.y), labels, ...(opts.lines ? { line: true } : {}),
               ...(sizeCi >= 0 ? { size: own.map(sizeOf) } : {}),
               labelTexts: textLabels ? own.map(labelOf) : undefined };
    });
  } else {
    const pts = rows.map(point);
    spec.series = [{
      name: tvMeasureLabel(vm, ym.ci), color: tvMarkColor(roles), x: pts.map(p => p.x), y: pts.map(p => p.y), labels,
      ...(opts.lines ? { line: true } : {}), ...(sizeCi >= 0 ? { size: rows.map(sizeOf) } : {}),
      labelTexts: textLabels ? rows.map(labelOf) : undefined,
      pointColors: colorCi >= 0 && scale ? rows.map(r => scale(tfDvNum(r[colorCi]))) : undefined
    }];
  }
  if (sizeCi >= 0) Object.assign(spec, { kind: "bubble", sizeTitle: tvMeasureLabel(vm, sizeCi), sizeNumFmt: tvNumFmt(vm, sizeCi) });
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
 *    stacked columns: invisible base + increase / decrease, each step in its category's colour, and the
 *    Text field written over each floating bar ─────────────────────────────── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvWaterfallSpec(ctx) {
  const { vm, roles } = ctx;
  const { horizontal, measure, catDims, valueShelf } = tvAxisLayout(ctx);
  const fieldLabels = vm.fmt.hasModel ? vm.fmt.fieldLabelsShown(valueShelf === "rows" ? "cols" : "rows") : true;
  const cats = tvCategories(vm, catDims.map(d => d.ci), catDims.some(d => d.continuous));
  if (cats.count > TV_MAX_POINTS) throw new Error(`${cats.count} categories – too many for an Excel chart`);
  const running = tvSum(vm, cats, measure.ci);
  if (running.some(v => v !== null && v < 0)) throw new Error("running total goes below zero – not drawable as stacked columns");
  const base = [], up = [], down = [], ends = [];
  let prev = 0;
  running.forEach(v => {
    const cur = v === null ? prev : v;
    const delta = cur - prev;
    base.push(Math.min(prev, cur));
    up.push(delta > 0 ? delta : null);
    down.push(delta < 0 ? -delta : null);
    ends.push(Math.max(prev, cur));                             // the top of the floating bar
    prev = cur;
  });
  const markColor = tvMarkColor(roles);
  const scale = tvColorScale(vm, roles, "ganttbar");
  // colour = the category (each step its own colour), or a measure (usually the step itself)
  const colorLevel = roles.color && !roles.color.continuous ? catDims.findIndex(d => d.ci === roles.color.ci) : -1;
  const stepColors = !scale ? null : colorLevel >= 0 ? cats.levels[colorLevel].map(v => scale(v))
    : roles.color && roles.color.continuous && roles.color.ci >= 0 ? tvSum(vm, cats, roles.color.ci).map(v => scale(v)) : null;
  const label = tvMeasureLabel(vm, measure.ci);
  const numFmt = tvNumFmt(vm, measure.ci);
  // mark labels: the Text field (the step, or the running total) over each bar; Excel writes stacked labels
  // inside the bars, so an invisible line along the bar tops carries them above, as text in its format
  const textCi = roles.labelRefs.map(r => vm.cols.findIndex(c => c.ref && tfSameField(c.ref, r))).find(i => i >= 0);
  const labelCi = textCi === undefined ? measure.ci : textCi;
  const labels = tvLabelsOn(roles, vm.cols[labelCi].ref);
  const labelFmt = tvNumFmt(vm, labelCi);
  const labelValues = labelCi === measure.ci ? running : tvSum(vm, cats, labelCi);
  const above = labels && !horizontal;
  // decreases are drawn as positive heights but labelled as the negative step, like Tableau
  const positive = (numFmt.match(/^((?:"[^"]*"|[^;])*)/) || [])[1] || "General";
  const downFmt = positive === "General" ? "-General" : "-" + positive;
  /** @type {ChartSeries[]} */
  const series = [
    { name: "Base", type: "bar", values: base, color: null, labels: false },
    { name: label + " (increase)", type: "bar", values: up, color: markColor, pointColors: stepColors, labels: labels && !above },
    { name: label + " (decrease)", type: "bar", values: down, color: markColor, pointColors: stepColors, labels: labels && !above, labelNumFmt: downFmt }
  ];
  // Analysis → Totals → Show Grand Totals on the category shelf: Tableau ends with a full bar from zero to the
  // final running total (drawn in grey)
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  const totals = vm.fmt.sheetModel && vm.fmt.sheetModel.grandTotals && vm.fmt.sheetModel.grandTotals[catShelf];
  if (totals && running.length) {
    const n = running.length;
    series.forEach(s => { s.values = [...s.values, null]; if (s.pointColors) s.pointColors = [...s.pointColors, null]; });
    series.push({ name: "Grand Total", type: "bar", values: [...Array(n).fill(null), prev], color: "9E9E9E", labels: labels && !above });
    cats.levels = cats.levels.map((lv, i) => [...lv, i === 0 ? "Grand Total" : ""]);
    ends.push(prev); labelValues.push(labelCi === measure.ci ? prev : null);
  }
  if (above) {
    series.push({ name: label + " (labels)", type: "line", values: ends, color: null, line: false, marker: false, labels: true,
                  labelTexts: labelValues.map(v => v === null ? "" : tfFormatNumber(v, labelFmt)) });
  }
  return [{
    // Tableau's Gantt bars float thin, about as wide as the gaps between them
    ...tvBaseSpec(vm), kind: "bar", barDir: horizontal ? "bar" : "col", stacked: true, gapWidth: 100, legend: false,
    conversion: { strategy: "CONSTRUCTED", output: "Waterfall: stacked columns on an invisible running base",
                  note: "built from stacked columns (invisible base, increases, decreases) so Tableau's step colours and labels stay – Excel's own waterfall cannot colour each step" },
    // "Allow labels to overlap other marks" off (Tableau's default): overlapping labels are left out
    labelCull: !roles.panes.some(p => tvPaneRule(p, "mark", "mark-labels-cull") === "false"),
    categories: { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [""], levels: cats.levels },
    categoryTitle: fieldLabels ? catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / ") : "",
    numFmt, valueTitle: label, series
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
    conversion: { strategy: "CONSTRUCTED", output: "Box plot: quartile boxes (up/down bars), whiskers (high-low lines), median marks",
                  note: "Tableau's quartiles and 1.5 × IQR whiskers computed from the marks – Excel's box & whisker chart would compute its own statistics" },
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
    conversion: { strategy: "CONSTRUCTED", output: "Gantt: stacked bars with invisible gaps on a date axis",
                  note: "Excel has no Gantt chart: each bar starts at its date and lasts its Size, the gaps before it invisible" },
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

/* ── packed bubbles: one bubble per mark, area = Size → bubble chart; the bubbles are packed here
 *    (largest in the middle) and scaled to the chart's size when the chart XML is written ──── */
export const TV_MAX_BUBBLES = 200;

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvPackedBubbleSpec(ctx) {
  const { vm, roles } = ctx;
  if (roles.size < 0) throw new Error("no Size measure for the bubbles");
  const colorDim = roles.color && !roles.color.measureNames && !roles.color.continuous && roles.color.ci >= 0 ? roles.color.ci : -1;
  let dims = [colorDim, ...roles.textDims, ...roles.detailDims].filter((ci, i, a) => ci >= 0 && ci !== roles.size && a.indexOf(ci) === i);
  if (!dims.length) dims = vm.cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);
  if (!dims.length) throw new Error("no dimension to split the bubbles");
  const cats = tvCategories(vm, dims, false);
  if (cats.count > TV_MAX_BUBBLES) throw new Error(`${cats.count} bubbles – too many to pack in an Excel chart`);
  const sizes = tvSum(vm, cats, roles.size);
  const keep = sizes.map((v, i) => v !== null && v > 0 ? i : -1).filter(i => i >= 0);   // a bubble needs an area
  if (!keep.length) throw new Error("no bubble with a positive size");
  const max = Math.max(...keep.map(i => sizes[i]));
  const radii = keep.map(i => Math.sqrt(sizes[i] / max));
  const { centres, box } = tvPackCircles(radii, 0.03);
  const scale = tvColorScale(vm, roles, "circle");
  let pointColors;
  if (scale && roles.color && roles.color.continuous && roles.color.ci >= 0) {
    const totals = tvSum(vm, cats, roles.color.ci);
    pointColors = keep.map(i => scale(totals[i]));
  } else if (scale && colorDim >= 0) {
    const level = dims.indexOf(colorDim);
    pointColors = keep.map((i, k) => scale(cats.levels[level][i]) || tvHex(TABLEAU_10[k % TABLEAU_10.length]));
  }
  // Tableau's label and tooltip per bubble: its template filled with the bubble's values – Tableau's own formatted
  // value when the bubble is one row, else the rows' sum in the field's format
  const rowsOf = keep.map(() => /** @type {any[][]} */ ([]));
  const at = new Map(keep.map((i, k) => [i, k]));
  vm.rows.forEach(row => { const k = at.get(cats.indexOf(row)); if (k !== undefined) rowsOf[k].push(row); });
  const valueOf = k => ref => {
    const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
    const rows = rowsOf[k];
    if (ci < 0 || !rows.length) return "";
    if (rows.length === 1 || vm.cols[ci].isHeader) return tfDvText(rows[0][ci]);
    const nums = rows.map(r => tfDvNum(r[ci])).filter(v => v !== null);
    const nf = vm.fmt.numFmtFor ? vm.fmt.numFmtFor(ref) : null;
    return nums.length ? (nf ? tfFormatNumber(nums.reduce((a, b) => a + b, 0), nf) : String(nums.reduce((a, b) => a + b, 0))) : "";
  };
  const text = lines => lines.map(line => line.map(x => x.text).join("")).join("\n");
  const labelTpl = vm.fmt.hasModel ? vm.fmt.labelRunsTemplate() : [];
  const tipTpl = vm.fmt.hasModel ? vm.fmt.tooltipTemplate() : [];
  // without a label template: the label dimension (Label, else Detail, else the innermost level)
  const labelLevel = dims.indexOf(roles.textDims[0] ?? roles.detailDims[0] ?? dims[dims.length - 1]);
  const labels = labelTpl.length ? vm.fmt.markLabelsShown() : roles.textDims.length > 0 || roles.labelRefs.length > 0;
  const labelTexts = labels ? keep.map((i, k) => labelTpl.length ? text(renderTemplate(labelTpl, valueOf(k))) : cats.levels[labelLevel][i]) : undefined;
  const tooltips = tipTpl.length ? keep.map((i, k) => text(renderTemplate(tipTpl, valueOf(k)))) : undefined;
  return [{
    ...tvBaseSpec(vm), kind: "bubble", gridlines: false, axesHidden: true, legend: false,
    xTitle: "Bubble x", numFmt: "General", xNumFmt: "General",
    sizeTitle: tvMeasureLabel(vm, roles.size), sizeNumFmt: tvNumFmt(vm, roles.size), packed: box,
    series: [{
      name: "Bubble y", color: tvMarkColor(roles), x: centres.map(p => p.x), y: centres.map(p => p.y),
      size: keep.map(i => sizes[i]), pointColors, labels, labelTexts, tooltips
    }],
    labelCull: vm.fmt.hasModel ? vm.fmt.markLabelsCulled() : true
  }];
}
