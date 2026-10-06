/* Bar / line / area / combo / funnel / histogram chart specs. */
import { TV_MAX_POINTS, TV_MAX_SERIES, tvBaseSpec, tvCategories, tvColorScale, tvColorValues, tvHex, tvLabelsOn, tvMarkColor, tvMarkToken, tvMeasureLabel, tvMeasureMark, tvNumFmt, tvPaneRule, tvSeriesType, tvSum, tvText } from "./common.js";
import { VISUAL_TYPES } from "../../config.js";
import { tfDvNum } from "../../data/values.js";
import { tfArgb } from "../../format/colors.js";
import { TABLEAU_10 } from "../../format/palettes.js";
import { tfParseFieldRef, tfSameField } from "../../twb/field-ref.js";
import { tfFieldInfo } from "../../twb/formatter.js";
import { tvIsConstantRef, tvIsMeasureRef } from "../../visual/classify.js";

/** a constant measure's value (MIN(0) → 0, AVG(1) → 1), else undefined @param {ChartContext} ctx @param {FieldRef | null} ref */
const constantOf = (ctx, ref) => ref && tvIsConstantRef(ctx.model || null, ref) ? tfFieldInfo(ctx.model, ref).constant : undefined;
/** a "% of Total" quick table calculation @param {FieldRef | null} ref */
const isPercentOfTotal = ref => !!ref && (/^pcto$/i.test(ref.deriv || "") || /^pcto:/i.test(ref.inner || ""));
/** a measure its formula negates ("-COUNT([Customers])"): the left wing of a butterfly chart */
const isNegated = (ctx, ref) => { const info = ref && ctx.model ? tfFieldInfo(ctx.model, ref) : null; return !!info && /^\s*-/.test(info.formula || ""); };
/** a number format that shows negative values without their minus sign (both wings of a butterfly read positive) */
const absFmt = fmt => { const p = (String(fmt || "General").match(/^((?:"[^"]*"|[^;])*)/) || [])[1] || "General"; return `${p};${p}`; };

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvCartesianSpecs(ctx) {
  const { vm, roles } = ctx;
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) throw new Error("no continuous measure axis to chart");
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  const horizontal = valueShelf === "cols";
  // a constant 0 on an axis (MIN(0), AVG(0)) only places labels – a butterfly's names – or the start of rounded bars:
  // not data, so not a series
  const zero = m => constantOf(ctx, m.ref) === 0;
  if (roles[valueShelf].values.some(zero) && roles[valueShelf].values.some(m => !zero(m))) {
    const values = roles[valueShelf].values.filter(m => !zero(m));
    return tvCartesianSpecs({ ...ctx, roles: { ...roles, [valueShelf]: { ...roles[valueShelf], values } } });
  }
  const measures = roles[valueShelf].values;
  const mvMode = measures.some(m => m.mv);
  const catDims = [...roles[valueShelf].dims, ...roles[catShelf].dims];
  const color = roles.color;
  const colorCi = color && !color.measureNames && !color.continuous ? color.ci : -1;
  const colorLevel = colorCi >= 0 ? catDims.findIndex(d => d.ci === colorCi) : -1;
  const scale = tvColorScale(vm, roles, ctx.markToken);
  const fieldLabels = vm.fmt.hasModel ? vm.fmt.fieldLabelsShown(catShelf) : true;
  // rounded bars are lines in Tableau (thick, round caps): Excel's bars
  const asBars = t => ctx.roundedBar && t === "line" ? "bar" : t;
  const sheetMark = asBars(ctx.markToken);
  // an Automatic marks card draws the All card's mark, else Tableau's automatic mark for the shelves (a line over
  // dates beside a card of bars – not the bars of the other card)
  const allPane = roles.panes.find(p => !p.id);
  const inherit = (allPane && asBars(tvMarkToken(allPane.markClass))) || asBars((roles.source === "twb" && ctx.autoMark) || ctx.markToken);
  /** @param {FieldRef | null} ref @param {string} fallback @param {number} [index] @param {number} [count] */
  const measureMark = (ref, fallback, index, count) => {
    const own = ref && roles.panes.find(x => x.id && [x.yAxisName, x.xAxisName].some(n => n && tfSameField(tfParseFieldRef(n), ref)));
    if (own && !tvMarkToken(own.markClass)) return inherit;
    return asBars(tvMeasureMark(roles, ref, fallback, index, count));
  };

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

  // a dimension on Detail / Label that splits the marks of one category and colour: Tableau draws a line (or a dot)
  // per member – each becomes its own series, never their sum
  const splitDims = (() => {
    if (mnCategory) return [];
    const cand = [...roles.detailDims, ...roles.textDims].filter((ci, i, a) => ci >= 0 && a.indexOf(ci) === i && ci !== colorCi &&
                                                                    !catDims.some(d => d.ci === ci));
    if (!cand.length) return [];
    const seen = new Map();
    for (const r of vm.rows) {
      const k = cats.indexOf(r) + "\u0001" + (colorCi >= 0 ? tvText(r[colorCi]) : "");
      const v = cand.map(ci => tvText(r[ci])).join("\u0001");
      if (!seen.has(k)) seen.set(k, v); else if (seen.get(k) !== v) return cand;
    }
    return [];
  })();

  const seriesFor = (m, opts) => {
    const label = tvMeasureLabel(vm, m.ci);
    const labels = tvLabelsOn(roles, m.ref);
    if (splitDims.length && opts.type && opts.type.type === "line") {
      const keyCis = [...(colorCi >= 0 && colorLevel < 0 && !opts.noColorSplit ? [colorCi] : []), ...splitDims];
      const keyOf = r => keyCis.map(ci => tvText(r[ci])).join(" – ");
      const keys = [...new Set(vm.rows.map(keyOf))];
      if (keys.length > TV_MAX_SERIES) {
        throw new Error(`${keys.length} lines (one per ${splitDims.map(ci => tvMeasureLabel(vm, ci)).join(" / ")}) – too many for an Excel chart`);
      }
      // each line in its member's colour (the colour field's value on its marks), else the marks' colour
      const colourOf = k => { const r = colorCi >= 0 && scale ? vm.rows.find(x => keyOf(x) === k) : null; return r ? scale(tvText(r[colorCi])) : null; };
      return keys.map((k, i) => ({
        name: opts.prefix ? `${label} – ${k}` : k, values: tvSum(vm, cats, m.ci, r => keyOf(r) === k),
        color: colourOf(k) || opts.color || tvHex(TABLEAU_10[i % TABLEAU_10.length]), labels, ...opts.type
      }));
    }
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

  // Tableau never draws a line across panes: categories of several levels (a year's quarters, a country's months) start
  // a new line at every change of an outer level – each such stretch its own series, the same colour, blank elsewhere
  const paneBreaks = series => {
    const levels = cats.levels;
    const broken = s => !s.refLine && ((s.type === "line" && s.line !== false) || s.type === "area");
    if (levels.length < 2 || !series.some(broken)) return series;
    const groupOf = i => levels.slice(0, -1).map(lv => lv[i]).join("\u0001");
    const ranges = [];
    for (let i = 0; i < cats.count; i++) {
      const g = groupOf(i), last = ranges[ranges.length - 1];
      if (last && last.g === g) last.to = i; else ranges.push({ g, from: i, to: i });
    }
    if (ranges.length < 2 || ranges.length * series.length > TV_MAX_SERIES) return series;
    return series.flatMap(s => !broken(s) ? [s]
      : ranges.map((r, k) => ({ ...s, name: k ? `${s.name} ` + "\u200b".repeat(k) : s.name,
                                values: s.values.map((v, i) => i >= r.from && i <= r.to ? v : null),
                                pointColors: s.pointColors ? s.pointColors.map((c, i) => i >= r.from && i <= r.to ? c : null) : undefined })));
  };

  // stacked bars / areas of a "% of Total" that fills every category: Excel's 100% stacked chart
  const percentOfTotal = measures.length > 0 && measures.every(m => isPercentOfTotal(m.ref));
  const make = (series, extra) => {
    if (series.length > TV_MAX_SERIES) throw new Error(`${series.length} series – too many for an Excel chart`);
    const kinds = new Set(series.map(s => s.type));
    const kind = kinds.size > 1 || series.some(s => s.secondary) ? "combo" : ([...kinds][0] || "bar");
    const stackable = series.every(s => s.type === "bar" || s.type === "area");
    const stacked = stackable && series.length > 1 &&
      ((colorCi >= 0 && colorLevel < 0) || (mvMode && color && color.measureNames && roles.measureNames !== catShelf));
    const totals = Array.from({ length: cats.count }, (_, i) => series.reduce((a, s) => a + ((s.values && s.values[i]) || 0), 0));
    const percent = stacked && percentOfTotal && totals.some(t => t) && totals.every(t => t === 0 || Math.abs(t - 1) < 0.005);
    // Tableau stacks columns in the colour legend's order from the top; Excel stacks its first series at the
    // bottom → reversed (bars across keep the legend's order from the axis outwards)
    const ordered = paneBreaks(stacked && !horizontal ? [...series].reverse() : series);
    const spec = { ...tvBaseSpec(vm), kind, barDir: horizontal ? "bar" : "col", stacked,
                   categories, categoryTitle, legend: series.length > 1, series: ordered, ...extra };
    if (percent) {
      spec.percent = true;
      if (!/%/.test(spec.numFmt || "")) spec.numFmt = "0%";
    }
    return spec;
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
  const paneMarkOfAxis = k => { const p = paneOfAxis(k); return p ? asBars(tvMarkToken(p.markClass)) || (p.id ? inherit : "") : ""; };
  const paneColor = (m, i) => {
    const p = paneOfAxis(m.axis || 0);
    return (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) ||
           (p ? tvHex(tfArgb(tvPaneRule(p, "mark", "mark-color"))) : null) || tvHex(TABLEAU_10[i % TABLEAU_10.length]);
  };

  // an axis of custom shapes picked by a field (logos, ▲▼ icons: Shape mark + Shape encoding) is
  // decoration Excel cannot draw → chart the other measures instead of a meaningless extra series
  if (axisIds.length > 1) {
    const decorative = axisIds.filter(k => { const p = paneOfAxis(k); return p && tvMarkToken(p.markClass) === "shape" && p.encodings.some(e => e.channel === "shape"); });
    if (decorative.length && decorative.length < axisIds.length) {
      const keep = measures.filter(m => !decorative.includes(m.axis || 0));
      if (keep.length) { measures.splice(0, measures.length, ...keep); return tvCartesianSpecs({ ...ctx, roles: { ...roles, [valueShelf]: { ...roles[valueShelf], values: keep } } }); }
    }
  }

  // butterfly / population pyramid: two measures back to back – the left one negated by its formula, or drawn on a
  // reversed axis – as bars stacked around zero; the left wing's minus signs are hidden
  const uniqueMeasures = measures.filter((m, i) => measures.findIndex(x => x.ci === m.ci) === i);
  if (horizontal && uniqueMeasures.length === 2 && colorLevel < 0 && colorCi < 0) {
    const reversed = m => !m.mv && !!vm.fmt.axisSpace(m.ref, valueShelf, "0").reverse;
    const left = uniqueMeasures.findIndex(m => isNegated(ctx, m.ref) || reversed(m));
    const right = 1 - left;
    if (left >= 0 && !isNegated(ctx, uniqueMeasures[right].ref) && !reversed(uniqueMeasures[right])) {
      const flip = !isNegated(ctx, uniqueMeasures[left].ref);           // a reversed axis holds positive values
      const series = [left, right].map((k, i) => {
        const m = uniqueMeasures[k];
        const values = tvSum(vm, cats, m.ci).map(v => v === null ? null : k === left && flip ? -v : v);
        return { name: tvMeasureLabel(vm, m.ci).replace(/^AGG\(-(.*)\)$/i, "$1"), values, type: /** @type {"bar"} */ ("bar"),
                 color: paneColor(m, i), labels: tvLabelsOn(roles, m.ref), labelNumFmt: absFmt(tvNumFmt(vm, m.ci)) };
      });
      const spec = make(series, { numFmt: absFmt(tvNumFmt(vm, uniqueMeasures[right].ci)), valueTitle: "", gapWidth: 30 });
      return [{ ...spec, stacked: true, legend: true, mirrored: true,
                conversion: { strategy: "CONSTRUCTED", output: "Back-to-back bar chart (butterfly)",
                              note: "Excel has no butterfly chart: two series stacked around zero, the left one negated, its minus signs hidden" } }];
    }
  }

  // barbell (dumbbell): two values per category joined by a line (Measure Names on the line's Path) → markers for
  // both values and Excel's high-low line between them
  const pathMN = roles.panes.some(p => tvMarkToken(p.markClass) === "line" && p.encodings.some(e => e.channel === "path" && e.field.name === "Measure Names"));
  if (pathMN && mvMode && ctx.type !== VISUAL_TYPES.BAR && uniqueMeasures.length === 2 && !mnCategory && horizontal) {
    const linePane = roles.panes.find(p => tvMarkToken(p.markClass) === "line");
    const [a, b] = uniqueMeasures.map(m => tvSum(vm, cats, m.ci));
    const join = { name: "", color: (linePane && tvHex(tfArgb(tvPaneRule(linePane, "mark", "mark-color")))) || "A0A0A0", line: true, lineWidth: 2,
                   values: a.flatMap((v, i) => [v, b[i], null]), cats: a.flatMap((v, i) => v === null || b[i] === null ? [null, null, null] : [i, i, null]) };
    const dots = uniqueMeasures.map((m, i) => ({ name: tvMeasureLabel(vm, m.ci), color: paneColor(m, i), marker: "circle", markerSize: 9,
                                                 values: [a, b][i], cats: [a, b][i].map((v, c) => v === null ? null : c) }));
    const spec = make([{ name: "", values: a.map(() => null), type: "bar", color: null, labels: false }],
                      { numFmt: tvNumFmt(vm, uniqueMeasures[0].ci), valueTitle: "" });
    return [{ ...spec, legend: false, overlay: [join, ...dots],
              conversion: { strategy: "CONSTRUCTED", output: "Dumbbell: dots and joining lines over the bar chart's category bands",
                            note: "Excel has no dumbbell chart: both values as dots, joined by a line, drawn over horizontal category bands" } }];
  }
  if (pathMN && mvMode && ctx.type !== VISUAL_TYPES.BAR && uniqueMeasures.length === 2 && !mnCategory) {
    const linePane = roles.panes.find(p => tvMarkToken(p.markClass) === "line");
    const series = uniqueMeasures.map((m, i) => ({ name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), type: /** @type {"line"} */ ("line"),
      line: false, marker: true, markerSize: 9, color: paneColor(m, i), labels: tvLabelsOn(roles, m.ref) }));
    const spec = make(series, { numFmt: tvNumFmt(vm, uniqueMeasures[0].ci), valueTitle: "" });
    return [{ ...spec, kind: "line", barDir: "col", stacked: false, legend: true,
              hiLowLines: { color: (linePane && tvHex(tfArgb(tvPaneRule(linePane, "mark", "mark-color")))) || "A0A0A0" },
              conversion: horizontal
                ? { strategy: "APPROXIMATE", output: "Line chart: markers joined by high-low lines (vertical)",
                    note: "drawn vertically: Excel joins markers with high-low lines only on vertical line charts" }
                : { strategy: "CONSTRUCTED", output: "Line chart: markers joined by high-low lines",
                    note: "Excel has no dumbbell chart: both values as markers, joined by high-low lines" } }];
  }

  // the same measure on both axes (styling: line + shape at the points, bar + line, Gantt cap on bars)
  // → drawn once: bar / area / line wins, a line gets markers when the other pane draws shapes
  if (!mvMode && axisIds.length > 1 && measures.every(m => m.ci === measures[0].ci)) {
    const toks = axisIds.map(k => paneMarkOfAxis(k) || sheetMark).filter(Boolean);
    if (toks.includes("bar") && toks.some(x => /^(circle|shape)$/.test(x))) {
      const m = measures[0];
      const base = { name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), color: tvMarkColor(roles), pointColors };
      /** @type {ChartSeries[]} */
      const series = [{ ...base, type: "bar", labels: horizontal && tvLabelsOn(roles, m.ref) }];
      if (!horizontal) series.push({ ...base, name: base.name + " ", type: "line", line: false, marker: true, markerSize: 9, labels: tvLabelsOn(roles, m.ref) });
      const spec = make(series, { numFmt: tvNumFmt(vm, m.ci), gapWidth: 300, valueTitle: tvMeasureLabel(vm, m.ci) });
      const heads = horizontal ? [{ name: base.name, color: base.color, marker: "circle", markerSize: 9, values: base.values, cats: base.values.map((v, i) => v === null ? null : i) }] : undefined;
      return [{ ...spec, stacked: false, legend: false, overlay: heads, conversion: lollipopConversion(horizontal) }];
    }
    const primary = ["bar", "area", "line"].find(x => toks.includes(x)) || toks[0] || "bar";
    const type = { ...tvSeriesType(primary) };
    if (primary === "line" && toks.some(x => /^(circle|shape|square)$/.test(x))) type.marker = true;
    const m = measures[0];
    return [withTargets(make(seriesFor(m, { type, color: tvMarkColor(roles) }), { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) }))];
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
      return (axisCount > 1 && paneMarkOfAxis(k)) || measureMark(null, sheetMark);
    });
    const unique = uniqueMeasures;
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
        // across: each bar's head as a dot over its band
        if (horizontal) spec.overlay = series.map(s => ({ name: s.name, color: s.color, marker: "circle", markerSize: 9, values: s.values, cats: s.values.map((v, i) => v === null ? null : i) }));
        spec.conversion = lollipopConversion(horizontal);
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
    const type = tvSeriesType(measureMark(m.ref, sheetMark));
    // a horizontal dot plot: each mark (detail included) a dot over its category band, one series per colour
    if (horizontal && type.type === "line" && type.line === false && !mnCategory) {
      const groups = colorCi >= 0 && colorLevel < 0 ? tvColorValues(vm, roles) : [null];
      if (groups.length <= TV_MAX_SERIES) {
        const overlay = groups.map((g, i) => {
          const own = vm.rows.filter(r => (g === null || tvText(r[colorCi]) === g) && tfDvNum(r[m.ci]) !== null && cats.indexOf(r) !== undefined);
          const c = g === null ? tvMarkColor(roles) : (scale && scale(g)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]);
          return { name: g === null ? tvMeasureLabel(vm, m.ci) : g, color: c, marker: "circle", markerSize: 7,
                   values: own.map(r => tfDvNum(r[m.ci])), cats: own.map(r => cats.indexOf(r)) };
        });
        const spec = make([{ name: "", values: new Array(cats.count).fill(null), type: "bar", color: null, labels: false }],
                          { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) });
        return [{ ...spec, legend: groups.length > 1, overlay,
                  conversion: { strategy: "CONSTRUCTED", output: "Dot plot: dots over the bar chart's category bands",
                                note: "Excel has no horizontal dot chart: every mark a dot (XY) over its category's band" } }];
      }
    }
    return [withTargets(make(seriesFor(m, { type, color: tvMarkColor(roles) }),
                             { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) }))];
  }

  // several measures: dual axis / combo (different marks) or separate panes (same mark). Tableau's Dual Axis
  // folds the second axis onto the first; Synchronize Axis gives both one scale
  // each axis's own marks card (found by its axis, else its position), an Automatic one drawing what it inherits
  const marks = measures.map((m, i) => {
    const p = paneOfAxis(m.axis || 0);
    return p && p.id ? asBars(tvMarkToken(p.markClass)) || inherit : measureMark(m.ref, sheetMark, i, measures.length);
  });
  const space = (m, i) => vm.fmt.axisSpace(m.ref, valueShelf, measures.slice(0, i).some(x => tfSameField(x.ref, m.ref)) ? "1" : "0");
  const folded = measures.length === 2 && !!space(measures[1], 1).fold;
  const synced = measures.length === 2 && !!space(measures[1], 1).synchronized;
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
        type: { ...tvSeriesType(marks[i]), secondary: measures.length === 2 && i === 1 && !synced },
        color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) || (!split && own) ||
               tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        prefix: colorCi >= 0 && split, noPointColors: true, noColorSplit: !split,
        alpha: opacity >= 0 && opacity < 255 ? opacity / 255 : undefined
      }).forEach(s => series.push(s));
    });
    const spec = make(series, {
      numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: tvMeasureLabel(vm, measures[0].ci),
      secondaryNumFmt: tvNumFmt(vm, measures[1].ci),
      secondaryTitle: measures.length === 2 ? tvMeasureLabel(vm, measures[1].ci) : undefined
    });
    // without Dual Axis Tableau draws each measure in a pane of its own: overlaid here, each on its own axis
    if (roles.source === "twb" && !folded) {
      spec.conversion = { strategy: "APPROXIMATE", note: "Tableau draws each measure in its own pane (no dual axis); Excel overlays them, each on its own axis" };
    }
    return [spec];
  }
  // the same mark on a dual axis: one chart. Bars are drawn one in the other, as Tableau draws them (bar in bar,
  // a progress bar over its track): the wider marks (the pane's Size) – else the longer values – on the primary axis
  // behind, the other on the secondary axis in front, narrower when Tableau draws it narrower. Synchronized axes share
  // one scale (the secondary axis is pinned to the primary's range and hidden); lines and areas of a synchronized
  // axis simply share the primary axis
  // a colour field on one axis's card splits only that axis's bars (a trend's highlighted min / max months)
  const colourCards = roles.panes.filter(p => p.encodings.some(e => e.channel === "color"));
  const coloured = m => { const p = paneOfAxis(m.axis || 0); return colorCi >= 0 && colorLevel < 0 && (!p || !colourCards.length || colourCards.includes(p)); };
  if (folded && (colorCi < 0 || measures.filter(coloured).length === 1)) {
    const type = tvSeriesType(marks[0]);
    const bars = type.type === "bar";
    const sizeOf = m => { const p = paneOfAxis(m.axis || 0); return (p && Number(tvPaneRule(p, "mark", "size"))) || 1; };
    const reach = m => Math.max(0, ...tvSum(vm, cats, m.ci).filter(v => v !== null).map(Math.abs));
    let order = [0, 1];
    if (bars) order.sort((a, b) => sizeOf(measures[b]) - sizeOf(measures[a]) || reach(measures[b]) - reach(measures[a]) || a - b);
    const [back, front] = order.map(k => measures[k]);
    const series = [back, front].flatMap((m, i) => seriesFor(m, {
      type: { ...type, secondary: i === 1 && (bars || !synced) }, color: paneColor(m, measures.indexOf(m)),
      noColorSplit: !coloured(m), noPointColors: true
    }).map(s => ({ ...s, labels: s.labels && constantOf(ctx, m.ref) === undefined })));
    const spec = make(series, { numFmt: tvNumFmt(vm, back.ci), valueTitle: tvMeasureLabel(vm, back.ci),
                                secondaryNumFmt: tvNumFmt(vm, front.ci), secondaryTitle: synced ? undefined : tvMeasureLabel(vm, front.ci) });
    // the coloured axis's bars stack by colour (one colour per category as a rule); the other draws whole bars
    spec.stacked = colorCi >= 0 && bars;
    spec.legend = false;
    if (bars && synced) { spec.secondarySync = true; spec.secondaryAxisHidden = true; }
    if (bars && sizeOf(front) < sizeOf(back) * 0.9) { spec.gapWidth = 50; spec.secondaryGapWidth = 250; }
    return [spec];
  }
  if (measures.length <= 6) {
    return measures.map((m, i) => withTargets(make(seriesFor(m, { type: tvSeriesType(marks[i]), color: tvMarkColor(roles) }),
      { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci), paneIndex: i, paneCount: measures.length })));
  }
  const series = measures.map((m, i) => ({ name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci),
    color: tvHex(TABLEAU_10[i % TABLEAU_10.length]), labels: tvLabelsOn(roles, m.ref), ...tvSeriesType(marks[i]) }));
  return [make(series, { numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "Value" })];

  /* bullet graph: a reference line per cell from another field (the target) over bars → a tick at each category's
   * target: dash markers over columns; across horizontal bars, vertical ticks drawn by the writer */
  function withTargets(spec) {
    const fmt = vm.fmt;
    const rl = fmt.hasModel && fmt.referenceLines
      ? fmt.referenceLines().find(r => r.scope === "per-cell" && r.field && r.axis && !tfSameField(r.field, r.axis)) : null;
    if (!rl || !spec.series.some(s => (s.type || spec.kind) === "bar") || spec.percent) return spec;
    const ti = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, rl.field));
    if (ti < 0) return spec;
    // the line's value per cell: the target's average over the cell's marks (one mark per category as a rule)
    const sums = tvSum(vm, cats, ti), counts = new Array(cats.count).fill(0);
    vm.rows.forEach(r => { const i = cats.indexOf(r); if (i !== undefined && tfDvNum(r[ti]) !== null) counts[i]++; });
    const targets = sums.map((v, i) => v === null || !counts[i] ? null : v / counts[i]);
    const name = tvMeasureLabel(vm, ti);
    const conversion = { strategy: /** @type {"CONSTRUCTED"} */ ("CONSTRUCTED"), output: `${horizontal ? "Bar" : "Column"} chart with target ticks (bullet graph)`,
                         note: "Excel has no bullet graph: the actual values as bars, each category's target as a tick" };
    if (horizontal) return { ...spec, targets: { name, values: targets, color: "333333" }, conversion };
    const tick = { name, values: targets, type: /** @type {"line"} */ ("line"), line: false, marker: true, markerSymbol: "dash",
                   markerSize: 20, color: "333333", labels: false };
    return { ...spec, kind: "combo", series: [...spec.series, tick], conversion };
  }
}

/**
 * Names what an Excel chart cannot show of the Tableau view (the conversion becomes APPROXIMATE): bars that sum
 * Tableau's segments per detail member, bar widths from Size, a horizontal line or dot plot drawn vertically.
 * @param {ChartSpec} spec @param {ChartContext} ctx
 */
export function tvNoteApproximations(spec, ctx) {
  if (spec.conversion || !/^(bar|line|area|combo)$/.test(spec.kind)) return;
  const { vm, roles } = ctx;
  const notes = [];
  const bars = spec.series.some(s => (s.type || spec.kind) === "bar" && s.color !== null);
  const lines = spec.series.filter(s => !s.refLine).every(s => (s.type || spec.kind) === "line");
  if (lines && spec.barDir === "bar") {
    spec.barDir = "col";                                   // Excel draws line charts vertically only
    notes.push("drawn vertically: Excel has no horizontal line or dot chart");
  }
  // bars sized by a measure (bar widths carry data); a size on another card (shapes beside the bars) is not the bars'
  const sizedBars = roles.panes.some(p => {
    const e = p.encodings.find(x => x.channel === "size");
    const mark = tvMarkToken(p.markClass) || (spec.kind === "bar" ? "bar" : "");
    return e && mark === "bar" && tvIsMeasureRef(ctx.model || null, e.field);
  });
  if (bars && sizedBars) notes.push("Tableau sizes these bars by a measure; Excel bars all have one width");
  if (bars) {
    const shelfDims = [...roles.rows.dims, ...roles.cols.dims].map(d => d.ci);
    const colorCi = roles.color && !roles.color.measureNames && roles.color.ci >= 0 ? roles.color.ci : -1;
    const detail = [...roles.detailDims, ...roles.textDims].filter(ci => ci !== colorCi && !shelfDims.includes(ci));
    if (detail.length) {
      const seen = new Map();
      const split = vm.rows.some(r => {
        const k = [...shelfDims, colorCi].map(ci => ci >= 0 ? String(r[ci] && r[ci].formattedValue) : "").join("\u0001");
        const v = detail.map(ci => String(r[ci] && r[ci].formattedValue)).join("\u0001");
        if (!seen.has(k)) { seen.set(k, v); return false; }
        return seen.get(k) !== v;
      });
      if (split) notes.push(`one bar per category: Tableau's segments per ${detail.map(ci => tvMeasureLabel(vm, ci)).join(" / ")} are summed (totals unchanged)`);
    }
  }
  if (notes.length) spec.conversion = { strategy: "APPROXIMATE", note: notes.join("; ") };
}

/** a lollipop's conversion: thin bars, and a dot at each bar's end */
function lollipopConversion(horizontal) {
  return { strategy: /** @type {"CONSTRUCTED"} */ ("CONSTRUCTED"), output: `${horizontal ? "Bar" : "Column"} chart (thin bars) with a dot on each`,
           note: "Excel has no lollipop chart: thin bars, each with a dot at its end" + (horizontal ? " (dots drawn over the category bands)" : "") };
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
    ],
    conversion: { strategy: "CONSTRUCTED", output: "Funnel: centred stacked bars",
                  note: "built from stacked bars with an invisible offset, so it opens in every Excel version" }
  }];
}

/* ── histogram: the bar chart with touching bars ─────────────────────────── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
export function tvHistogramSpecs(ctx) {
  return tvCartesianSpecs(ctx).map(s => ({ ...s, gapWidth: 0,
    conversion: { strategy: "CONSTRUCTED", output: `Column chart of Tableau's bins (no gaps)`,
                  note: "Tableau's own bins and counts as touching columns – the same bins Tableau draws, not Excel's automatic ones" } }));
}
