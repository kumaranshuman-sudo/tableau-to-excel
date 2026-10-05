/* Bar / line / area / combo / funnel / histogram chart specs. */
import { TV_MAX_POINTS, TV_MAX_SERIES, tvBaseSpec, tvCategories, tvColorScale, tvColorValues, tvHex, tvLabelsOn, tvMarkColor, tvMarkToken, tvMeasureLabel, tvMeasureMark, tvNumFmt, tvPaneRule, tvSeriesType, tvSum, tvText } from "./common.js";
import { tfDvNum } from "../../data/values.js";
import { tfArgb } from "../../format/colors.js";
import { TABLEAU_10 } from "../../format/palettes.js";
import { tfParseFieldRef, tfSameField } from "../../twb/field-ref.js";

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvCartesianSpecs(ctx) {
  const { vm, roles } = ctx;
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) throw new Error("no continuous measure axis to chart");
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  const horizontal = valueShelf === "cols";
  const measures = roles[valueShelf].values;
  const mvMode = measures.some(m => m.mv);
  const catDims = [...roles[valueShelf].dims, ...roles[catShelf].dims];
  const color = roles.color;
  const colorCi = color && !color.measureNames && !color.continuous ? color.ci : -1;
  const colorLevel = colorCi >= 0 ? catDims.findIndex(d => d.ci === colorCi) : -1;
  const scale = tvColorScale(vm, roles, ctx.markToken);
  const fieldLabels = vm.fmt.hasModel ? vm.fmt.fieldLabelsShown(catShelf) : true;
  const sheetMark = ctx.markToken;

  // categories (Measure Names alone on the category shelf → one bar per measure)
  const mnCategory = mvMode && !catDims.length && roles.measureNames === catShelf;
  const cats = mnCategory
    ? { count: measures.length, levels: [measures.map(m => tvMeasureLabel(vm, m.ci))], indexOf: () => 0 }
    : tvCategories(vm, catDims.map(d => d.ci), catDims.some(d => d.continuous));
  if (cats.count > TV_MAX_POINTS) throw new Error(`${cats.count} categories – too many for an Excel chart`);
  const categoryTitle = fieldLabels && catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / ") : "";
  const categories = { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [mnCategory ? "Measure Names" : ""], levels: cats.levels };

  // per-category colours: colour = a category level, or a continuous measure
  let pointColors;
  if (scale && colorLevel >= 0) pointColors = cats.levels[colorLevel].map(v => scale(v));
  else if (scale && color && color.continuous && color.ci >= 0) pointColors = tvSum(vm, cats, color.ci).map(v => scale(v));

  const seriesFor = (m, opts) => {
    const label = tvMeasureLabel(vm, m.ci);
    const labels = tvLabelsOn(roles, m.ref);
    if (colorCi >= 0 && colorLevel < 0 && !opts.noColorSplit) {
      const values = tvColorValues(vm, roles);
      return values.map((v, i) => ({
        name: opts.prefix ? `${label} – ${v}` : v,
        values: tvSum(vm, cats, m.ci, r => tvText(r[colorCi]) === v),
        color: (scale && scale(v)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        labels, ...opts.type
      }));
    }
    return [{ name: label, values: tvSum(vm, cats, m.ci), color: opts.color, labels,
              pointColors: opts.noPointColors ? undefined : pointColors, ...opts.type,
              ...(opts.alpha !== undefined ? { alpha: opts.alpha } : {}) }];
  };

  const make = (series, extra) => {
    if (series.length > TV_MAX_SERIES) throw new Error(`${series.length} series – too many for an Excel chart`);
    const kinds = new Set(series.map(s => s.type));
    const kind = kinds.size > 1 || series.some(s => s.secondary) ? "combo" : ([...kinds][0] || "bar");
    const stackable = series.every(s => s.type === "bar" || s.type === "area");
    const stacked = stackable && series.length > 1 &&
      ((colorCi >= 0 && colorLevel < 0) || (mvMode && color && color.measureNames && roles.measureNames !== catShelf));
    // Tableau stacks columns in the colour legend's order from the top; Excel stacks its first series at the
    // bottom → reversed (bars across keep the legend's order from the axis outwards)
    const ordered = stacked && !horizontal ? [...series].reverse() : series;
    return { ...tvBaseSpec(vm), kind, barDir: horizontal ? "bar" : "col", stacked,
             categories, categoryTitle, legend: series.length > 1, series: ordered, ...extra };
  };

  const axisIds = [...new Set(measures.map(m => m.axis || 0))].sort((a, b) => a - b);
  const paneOfAxis = k => {
    const refs = roles[valueShelf].axisRefs || [];
    const ref = refs[k];
    if (ref) {
      const nth = refs.slice(0, k).filter(r => tfSameField(r, ref) || (r.name === ref.name && r.name === "Multiple Values")).length;
      const axisName = p => (horizontal ? p.xAxisName : p.yAxisName);
      const hits = roles.panes.filter(p => { const n = axisName(p) && tfParseFieldRef(axisName(p)); return n && (tfSameField(n, ref) || (n.name === "Multiple Values" && ref.name === "Multiple Values")); });
      if (hits[nth]) return hits[nth];
    }
    return roles.panes.find(x => String(x.id) === String(k + 1)) || null;
  };
  const paneMarkOfAxis = k => { const p = paneOfAxis(k); return p ? tvMarkToken(p.markClass) : ""; };

  // an axis of custom shapes picked by a field (logos, ▲▼ icons: Shape mark + Shape encoding) is
  // decoration Excel cannot draw → chart the other measures instead of a meaningless extra series
  if (axisIds.length > 1) {
    const decorative = axisIds.filter(k => { const p = paneOfAxis(k); return p && tvMarkToken(p.markClass) === "shape" && p.encodings.some(e => e.channel === "shape"); });
    if (decorative.length && decorative.length < axisIds.length) {
      const keep = measures.filter(m => !decorative.includes(m.axis || 0));
      if (keep.length) { measures.splice(0, measures.length, ...keep); return tvCartesianSpecs({ ...ctx, roles: { ...roles, [valueShelf]: { ...roles[valueShelf], values: keep } } }); }
    }
  }

  // the same measure on both axes (styling: line + shape at the points, bar + line, Gantt cap on bars)
  // → drawn once: bar / area / line wins, a line gets markers when the other pane draws shapes
  if (!mvMode && axisIds.length > 1 && measures.every(m => m.ci === measures[0].ci)) {
    const toks = axisIds.map(k => paneMarkOfAxis(k) || sheetMark).filter(Boolean);
    const primary = ["bar", "area", "line"].find(x => toks.includes(x)) || toks[0] || "bar";
    const type = { ...tvSeriesType(primary) };
    if (primary === "line" && toks.some(x => /^(circle|shape|square)$/.test(x))) type.marker = true;
    const m = measures[0];
    return [make(seriesFor(m, { type, color: tvMarkColor(roles) }), { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) })];
  }

  // a measure on one axis and Measure Values on the other (Sales bars + MV lines) → combo, axis each
  const groupOf = k => measures.filter(m => (m.axis || 0) === k);
  const sameSets = axisIds.every(k => { const a = groupOf(k).map(m => m.ci).sort().join(), b = groupOf(axisIds[0]).map(m => m.ci).sort().join(); return a === b; });
  if (mvMode && axisIds.length > 1 && !sameSets && !mnCategory) {
    const series = [];
    axisIds.forEach((k, ai) => {
      const tok = paneMarkOfAxis(k) || sheetMark;
      groupOf(k).filter((m, i, a) => a.findIndex(x => x.ci === m.ci) === i).forEach(m => series.push({
        name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), labels: tvLabelsOn(roles, m.ref),
        color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, series.length) : null) || tvHex(TABLEAU_10[series.length % TABLEAU_10.length]),
        ...tvSeriesType(tok), secondary: axisIds.length === 2 && ai === 1
      }));
    });
    const first = groupOf(axisIds[0])[0], second = groupOf(axisIds[1])[0];
    const spec = make(series, { numFmt: tvNumFmt(vm, first.ci), valueTitle: groupOf(axisIds[0]).length === 1 ? tvMeasureLabel(vm, first.ci) : "Value",
                                secondaryNumFmt: tvNumFmt(vm, second.ci), secondaryTitle: groupOf(axisIds[1]).length === 1 ? tvMeasureLabel(vm, second.ci) : "Value" });
    spec.stacked = false;
    spec.legend = series.length > 1;
    return [spec];
  }

  // Measure Values: one chart, one series per measure
  if (mvMode) {
    // dual axis "MV + MV": every measure is listed once per axis, and each axis has its own
    // pane (id 1, 2, …) with its own mark → one entry per measure, marks per axis
    const axisCount = Math.max(...measures.map(m => (m.axis || 0) + 1));
    const axisMarks = Array.from({ length: axisCount }, (_, k) => {
      return (axisCount > 1 && paneMarkOfAxis(k)) || tvMeasureMark(roles, null, sheetMark);
    });
    const unique = measures.filter((m, i) => measures.findIndex(x => x.ci === m.ci) === i);
    const markSet = [...new Set(axisMarks.map(t => tvSeriesType(t).marker && tvSeriesType(t).line === false ? "dot" : tvSeriesType(t).type))];
    const type = tvSeriesType(axisMarks.find(t => tvSeriesType(t).type === "bar") || axisMarks[0]);

    if (!mnCategory && axisCount > 1 && markSet.length > 1) {
      if (markSet.includes("bar") && markSet.includes("dot")) {
        // lollipop: bar + circle on the same values. Vertical → thin bars + markers;
        // horizontal → Excel cannot mix bar and line orientations, so thin bars only
        const colorOf = (m, i) => (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) ||
          (unique.length === 1 ? tvMarkColor(roles) : tvHex(TABLEAU_10[i % TABLEAU_10.length]));
        const series = [];
        unique.forEach((m, i) => {
          const base = { name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), color: colorOf(m, i), pointColors };
          series.push({ ...base, type: "bar", labels: false });
          if (!horizontal) series.push({ ...base, name: base.name + " ", type: "line", line: false, marker: true, markerSize: 9,
                                         labels: tvLabelsOn(roles, m.ref) });
        });
        if (horizontal) series.forEach(s => { s.labels = tvLabelsOn(roles, unique[0].ref); });
        const spec = make(series, { numFmt: tvNumFmt(vm, unique[0].ci), gapWidth: 300,
                                    valueTitle: unique.length === 1 ? tvMeasureLabel(vm, unique[0].ci) : "Value" });
        spec.stacked = false;                      // the circle sits on the bar end, never on top of it
        spec.legend = unique.length > 1;
        return [spec];
      }
      // e.g. area + line of the same measure: overlay both, sharing one axis
      const series = [];
      axisMarks.forEach((token, k) => unique.forEach((m, i) => series.push({
        name: tvMeasureLabel(vm, m.ci) + (k ? " " : ""), values: tvSum(vm, cats, m.ci), labels: k === axisCount - 1 && tvLabelsOn(roles, m.ref),
        color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        ...tvSeriesType(token)
      })));
      const spec = make(series, { numFmt: tvNumFmt(vm, unique[0].ci), valueTitle: unique.length === 1 ? tvMeasureLabel(vm, unique[0].ci) : "Value" });
      spec.stacked = false;
      spec.legend = unique.length > 1;
      return [spec];
    }
    measures.splice(0, measures.length, ...unique);

    // funnel: one category per measure, each row carries exactly one of the measures
    // (per-stage calcs) → one value per stage, drawn centred
    const exclusive = catDims.length === 1 && measures.length > 1 && type.type === "bar" &&
      vm.rows.every(r => measures.filter(m => tfDvNum(r[m.ci]) !== null).length <= 1);
    if (exclusive) return tvFunnelSpecFromStages(ctx, cats, catDims[0].ci, measures, horizontal);

    if (mnCategory) {
      const values = measures.map(m => vm.rows.reduce((s, r) => s + (tfDvNum(r[m.ci]) || 0), 0));
      const pc = color && color.measureNames && scale ? measures.map((m, i) => scale(vm.cols[m.ci].name, i)) : undefined;
      return [make([{ name: "Value", values, color: tvMarkColor(roles), pointColors: pc,
                      labels: tvLabelsOn(roles, measures[0].ref), ...type }],
                   { numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "Value" })];
    }
    const series = measures.map((m, i) => ({
      name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), labels: tvLabelsOn(roles, m.ref),
      color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
      ...type
    }));
    return [make(series, { numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "Value" })];
  }

  // one measure
  if (measures.length === 1) {
    const m = measures[0];
    const type = tvSeriesType(tvMeasureMark(roles, m.ref, sheetMark));
    return [make(seriesFor(m, { type, color: tvMarkColor(roles) }),
                 { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) })];
  }

  // several measures: dual axis / combo (different marks) or separate panes (same mark)
  const marks = measures.map((m, i) => tvMeasureMark(roles, m.ref, sheetMark, i, measures.length));
  if (new Set(marks).size > 1 && measures.length <= 4) {
    const series = [];
    const colourPanes = roles.panes.filter(p => p.encodings.some(e => e.channel === "color"));
    measures.forEach((m, i) => {
      // each axis has its own marks card: the colour field splits only the measure whose card holds it
      // (a KPI trend's min / max dots, not its area); the others keep their card's colour and opacity
      const pane = paneOfAxis(m.axis || 0);
      const split = !pane || !colourPanes.length || colourPanes.includes(pane);
      const own = pane ? tvHex(tfArgb(tvPaneRule(pane, "mark", "mark-color"))) : null;
      const opacity = pane ? Number(tvPaneRule(pane, "mark", "mark-transparency")) : NaN;
      seriesFor(m, {
        type: { ...tvSeriesType(marks[i]), secondary: measures.length === 2 && i === 1 },
        color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) || (!split && own) ||
               tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        prefix: colorCi >= 0 && split, noPointColors: true, noColorSplit: !split,
        alpha: opacity >= 0 && opacity < 255 ? opacity / 255 : undefined
      }).forEach(s => series.push(s));
    });
    return [make(series, {
      numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: tvMeasureLabel(vm, measures[0].ci),
      secondaryNumFmt: tvNumFmt(vm, measures[1].ci),
      secondaryTitle: measures.length === 2 ? tvMeasureLabel(vm, measures[1].ci) : undefined
    })];
  }
  if (measures.length <= 4) {
    return measures.map((m, i) => make(seriesFor(m, { type: tvSeriesType(marks[i]), color: tvMarkColor(roles) }),
      { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci), paneIndex: i, paneCount: measures.length }));
  }
  const series = measures.map((m, i) => ({ name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci),
    color: tvHex(TABLEAU_10[i % TABLEAU_10.length]), labels: tvLabelsOn(roles, m.ref), ...tvSeriesType(marks[i]) }));
  return [make(series, { numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "Value" })];
}

/* ── funnel: stages = categories with one value each → stacked bars of an
 *    invisible half-gap + the value, so every stage is centred ─────────────── */
/**
 * @param {ChartContext} ctx @param {ReturnType<typeof tvCategories>} cats @param {number} catCi
 * @param {RoleValue[]} measures @param {boolean} horizontal
 * @returns {ChartSpec[]}
 */
export function tvFunnelSpecFromStages(ctx, cats, catCi, measures, horizontal) {
  const { vm, roles } = ctx;
  const values = new Array(cats.count).fill(null);
  vm.rows.forEach(r => {
    const i = cats.indexOf(r);
    if (i === undefined) return;
    measures.forEach(m => { const n = tfDvNum(r[m.ci]); if (n !== null) values[i] = (values[i] || 0) + n; });
  });
  const max = Math.max(0, ...values.filter(v => v !== null));
  const scale = tvColorScale(vm, roles, "bar");
  const pointColors = scale && roles.color && roles.color.ci === catCi
    ? cats.levels[0].map((v, i) => scale(v) || tvHex(TABLEAU_10[i % TABLEAU_10.length])) : undefined;
  return [{
    ...tvBaseSpec(vm), kind: "bar", barDir: horizontal ? "bar" : "col", stacked: true, gapWidth: 15, legend: false,
    gridlines: false, valueAxisHidden: true,
    categories: { names: [tvMeasureLabel(vm, catCi)], levels: cats.levels }, categoryTitle: "",
    numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "",
    series: [
      { name: "Offset", type: "bar", values: values.map(v => v === null ? null : (max - v) / 2), color: null, labels: false },
      { name: tvMeasureLabel(vm, catCi), type: "bar", values, color: tvMarkColor(roles), pointColors, labels: roles.labelRefs.length > 0 }
    ]
  }];
}

/* ── histogram: the bar chart with touching bars ─────────────────────────── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvHistogramSpecs(ctx) {
  return tvCartesianSpecs(ctx).map(s => ({ ...s, gapWidth: 0 }));
}
