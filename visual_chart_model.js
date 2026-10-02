"use strict";

/* ══════════════════════════════════════════════════════════════════════════
 * TABLEAU VISUAL → EXCEL CHART SPECS
 * ──────────────────────────────────────────────────────────────────────────
 * buildExcelChartSpecs(visualModel, model) reads
 *   – live summary data (already pivoted / merged / sorted by buildViewModel)
 *   – TWB shelves, panes, marks, encodings, colours, number formats, labels
 *   – the live visual specification when no workbook was loaded
 * and returns renderer-neutral chart specs for excel_chart_writer.js.
 * It throws when a visual cannot be represented faithfully, so the caller
 * can fall back to the Tableau image renderer or the data table.
 *
 * Uses the tf* helpers, tvIsMeasureRef and VISUAL_TYPES from
 * build_table_copy.js (globals, resolved at call time).
 * ══════════════════════════════════════════════════════════════════════════ */

const TV_MAX_POINTS = 4000;
const TV_MAX_SERIES = 60;
const TV_DEFAULT_MARK_COLOR = "4E79A7";

function tvHex(argb) {
  if (!argb) return null;
  const s = String(argb).replace(/^#/, "").toUpperCase();
  return s.length === 8 ? s.slice(2) : s;
}
function tvText(dv) { return tfIsNull(dv) ? "Null" : tfDvText(dv); }
function tvMarkToken(cls) {
  const t = String(cls || "").toLowerCase().replace(/[\s_-]+/g, "");
  return t === "automatic" ? "" : t;
}
function tvPaneRule(pane, element, attrName) {
  const list = (pane && pane.style && pane.style.rules && pane.style.rules[element]) || [];
  const f = list.find(x => x.attr === attrName && !x.field && !x.scope);
  return f ? f.value : undefined;
}

/* ── field roles: which summary column sits on which shelf / encoding ───── */
function tvRoles(vm, model, liveSpec) {
  const sheet = vm.fmt.sheetModel;
  const roles = { rows: { values: [], dims: [] }, cols: { values: [], dims: [] }, measureNames: null,
                  color: null, angle: -1, size: -1, textDims: [], detailDims: [],
                  labelRefs: [], labelNames: [], panes: sheet ? sheet.panes : [], source: "none" };
  const addDim = (list, ci) => {
    if (ci >= 0 && (vm.cols[ci].isHeader || !numeric(vm.cols[ci])) && !list.includes(ci)) list.push(ci);
  };
  const pivoted = vm.cols.map((c, i) => c.pivoted ? i : -1).filter(i => i >= 0);
  const numeric = c => /^(int|float|real|integer|number|double)/i.test(String(c.dataType || "")) || c.pivoted;

  if (sheet) {
    roles.source = "twb";
    const colOf = ref => {
      if (!ref) return -1;
      let i = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
      if (i < 0) {
        const names = tfDisplayNames(model, ref);
        i = vm.cols.findIndex(c => names.includes(tfNorm(c.name)));
      }
      return i;
    };
    for (const shelf of ["rows", "cols"]) {
      for (const ref of sheet[shelf]) {
        if (ref.name === "Measure Names") { roles.measureNames = shelf; continue; }
        if (ref.name === "Multiple Values") {
          pivoted.forEach(ci => roles[shelf].values.push({ ci, ref: vm.cols[ci].ref, mv: true }));
          continue;
        }
        const ci = colOf(ref);
        if (ci < 0) continue;
        if (tvIsMeasureRef(model, ref)) roles[shelf].values.push({ ci, ref });
        else roles[shelf].dims.push({ ci, ref, continuous: ref.type === "qk" });
      }
    }
    const enc = channel => {
      for (const p of sheet.panes) {
        const e = p.encodings.find(x => x.channel === channel);
        if (e) return e.field;
      }
      return null;
    };
    const colorRef = enc("color");
    if (colorRef) {
      if (colorRef.name === "Measure Names") roles.color = { measureNames: true, ref: colorRef };
      else {
        const ci = colOf(colorRef);
        if (ci >= 0) roles.color = { ci, ref: colorRef, continuous: tvIsMeasureRef(model, colorRef) || colorRef.type === "qk" };
      }
    }
    const angleRef = enc("wedge-size");
    if (angleRef) roles.angle = angleRef.name === "Multiple Values" ? (pivoted[0] ?? -1) : colOf(angleRef);
    const sizeRef = enc("size");
    if (sizeRef) roles.size = colOf(sizeRef);
    sheet.panes.forEach(p => {
      p.encodings.filter(e => e.channel === "text" || e.channel === "label").forEach(e => {
        roles.labelRefs.push(e.field);
        if (!tvIsMeasureRef(model, e.field)) addDim(roles.textDims, colOf(e.field));
      });
      p.encodings.filter(e => e.channel === "lod").forEach(e => addDim(roles.detailDims, colOf(e.field)));
      p.labelRuns.forEach(r => r.refs.forEach(x => roles.labelRefs.push(x)));
    });
    return roles;
  }

  // ── no workbook loaded: use the live visual specification ──
  const spec = liveSpec && typeof liveSpec === "object" ? liveSpec : null;
  const byName = name => {
    const key = tfNorm(name);
    const strip = s => tfNorm(s).replace(/^[a-z]+\((.*)\)$/, "$1");
    let i = vm.cols.findIndex(c => tfNorm(c.name) === key);
    if (i < 0) i = vm.cols.findIndex(c => strip(c.name) === strip(name));
    return i;
  };
  const fieldName = f => f == null ? "" : typeof f === "string" ? f : (f.name || f.fieldName || f.caption || "");
  if (spec && (Array.isArray(spec.rowFields) || Array.isArray(spec.columnFields))) {
    roles.source = "live";
    for (const [shelf, list] of [["rows", spec.rowFields || []], ["cols", spec.columnFields || []]]) {
      list.forEach(f => {
        const name = fieldName(f);
        if (/^measure names$/i.test(name)) { roles.measureNames = shelf; return; }
        if (/^measure values$/i.test(name)) { pivoted.forEach(ci => roles[shelf].values.push({ ci, ref: vm.cols[ci].ref, mv: true })); return; }
        const ci = byName(name);
        if (ci < 0) return;
        if (numeric(vm.cols[ci]) && !vm.cols[ci].isHeader) roles[shelf].values.push({ ci, ref: vm.cols[ci].ref });
        else roles[shelf].dims.push({ ci, ref: vm.cols[ci].ref, continuous: /^date/i.test(String(vm.cols[ci].dataType || "")) });
      });
    }
    const marks = Array.isArray(spec.marksSpecifications) ? spec.marksSpecifications : [];
    roles.liveMarks = marks.map(m => tvMarkToken(m && (m.type || m.markType)));   // one marks card per measure
    const active = marks[spec.activeMarksSpecificationIndex] || marks[0];
    (active && active.encodings || []).forEach(e => {
      const name = fieldName(e.field);
      const type = String(e.type || e.encodingType || "").toLowerCase();
      if (type === "color") {
        if (/^measure names$/i.test(name)) roles.color = { measureNames: true };
        else { const ci = byName(name); if (ci >= 0) roles.color = { ci, continuous: numeric(vm.cols[ci]) && !vm.cols[ci].isHeader }; }
      } else if (type === "angle") roles.angle = byName(name);
      else if (type === "size") roles.size = byName(name);
      else if (type === "label" || type === "text") {
        roles.labelNames.push(tfNorm(name));
        const ci = byName(name);
        if (ci >= 0 && !numeric(vm.cols[ci])) addDim(roles.textDims, ci);
      } else if (type === "detail") addDim(roles.detailDims, byName(name));
    });
    if (roles.rows.values.length || roles.cols.values.length) return roles;
  }

  // ── nothing structural known: dimensions → categories, measures → rows ──
  roles.source = "summary";
  vm.cols.forEach((c, ci) => {
    if (c.isHeader) roles.cols.dims.push({ ci, ref: c.ref, continuous: /^date/i.test(String(c.dataType || "")) });
    else if (numeric(c)) roles.rows.values.push({ ci, ref: c.ref, mv: !!c.pivoted });
  });
  return roles;
}

/* ── formatting helpers ─────────────────────────────────────────────────── */
function tvChartFont(fmt) {
  const f = tfExcelFont(fmt.headerCellStyle(null));
  return { name: f.name, size: f.size || 9, color: tvHex(f.color && f.color.argb) || "333333" };
}
function tvNumFmt(vm, ci) {
  const c = vm.cols[ci];
  const st = vm.fmt.sheetModel ? vm.fmt.sheetModel.style : null;
  if (c && c.ref && st) {
    const axisRaw = tfCollect(st, ["axis"], { field: c.ref }).numFmtRaw;
    const axis = axisRaw ? tableauToExcelNumFmt(axisRaw) : null;
    if (axis) return axis;
  }
  const own = c && vm.fmt.markCellStyle(c.ref || null).numFmt;
  if (own) return own;
  for (const row of vm.rows) {
    const dv = row[ci];
    if (tfIsNull(dv)) continue;
    const n = tfDvNum(dv);
    if (n === null) continue;
    return inferExcelNumFmt(dv.formattedValue, n) || "General";
  }
  return "General";
}
function tvMeasureLabel(vm, ci) {
  const c = vm.cols[ci];
  return (c && (c.label || c.name)) || "Value";
}
/* mark of the pane that draws this measure (dual axis / multi-pane), else the sheet's.
 * Without a workbook, the live spec has one marks card per measure, in shelf order. */
function tvMeasureMark(roles, ref, fallback, index, count) {
  if (!roles.panes.length && roles.liveMarks && count > 1 && roles.liveMarks.length === count && roles.liveMarks[index]) {
    return roles.liveMarks[index];
  }
  if (ref) {
    const p = roles.panes.find(x => [x.yAxisName, x.xAxisName].some(n => n && tfSameField(tfParseFieldRef(n), ref)));
    const t = p && tvMarkToken(p.markClass);
    if (t) return t;
  }
  const all = roles.panes.find(x => !x.id) || roles.panes[0];
  return (all && tvMarkToken(all.markClass)) || fallback || "bar";
}
function tvLabelsOn(roles, ref) {
  const shown = roles.panes.some(p => tvPaneRule(p, "mark", "mark-labels-show") === "true");
  if (shown) return true;
  if (!ref) return false;
  return roles.labelRefs.some(r => tfSameField(r, ref) || (r.name === "Multiple Values")) ||
    roles.labelNames.some(n => n === tfNorm(ref.name));
}
function tvMarkColor(roles) {
  for (const p of roles.panes) {
    const c = tvHex(tfArgb(tvPaneRule(p, "mark", "mark-color")));
    if (c) return c;
  }
  return TV_DEFAULT_MARK_COLOR;
}

/* ── colour scale for the Color encoding (TWB mapping or Tableau automatic) ─ */
function tvColorScale(vm, roles, markToken) {
  if (!roles.color) return null;
  const fmt = vm.fmt;
  let enc = fmt.colorEncoding();
  if (!enc) {                                      // live spec only: Tableau automatic palettes
    const ci = roles.color.ci;
    const name = ci >= 0 ? vm.cols[ci].name : "Measure Names";
    enc = { ref: { name, inner: name }, def: null, continuous: !!roles.color.continuous, markClass: markToken || "bar" };
  }
  if (roles.color.measureNames) {
    const scale = tfBuildColorScale(fmt, { ...enc, continuous: false }, []);
    return (value, i) => tvHex(scale && scale(value)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]);
  }
  const ci = roles.color.ci;
  const values = vm.rows.map(r => roles.color.continuous ? tfDvNum(r[ci]) : tvText(r[ci]));
  const scale = tfBuildColorScale(fmt, enc, values);
  return scale ? (value => tvHex(scale(value))) : null;
}
/* legend order of a discrete colour field: manual sort, else natural */
function tvColorValues(vm, roles) {
  const ci = roles.color.ci;
  const seen = new Map();
  vm.rows.forEach(r => { const t = tvText(r[ci]); if (!seen.has(t)) seen.set(t, r[ci]); });
  let values = [...seen.keys()];
  const manual = roles.color.ref ? vm.fmt.manualSortFor(roles.color.ref) : null;
  if (manual) {
    const rank = new Map(manual.order.map((b, i) => [tfBucketKey(b), i]));
    values.sort((a, b) => (rank.has(tfNorm(a)) ? rank.get(tfNorm(a)) : 1e9) - (rank.has(tfNorm(b)) ? rank.get(tfNorm(b)) : 1e9));
    if (manual.direction === "DESC") values.reverse();
  } else values.sort((a, b) => tfNaturalCompare(seen.get(a), seen.get(b)));
  return values;
}

/* ── category axis: distinct label tuples in view order ─────────────────── */
function tvCategories(vm, catCis, sortNative) {
  let order = [];
  const seen = new Map();
  vm.rows.forEach(row => {
    const labels = catCis.map(ci => tvText(row[ci]));
    const key = labels.join("\u0001");
    if (!seen.has(key)) { seen.set(key, true); order.push({ key, labels, dvs: catCis.map(ci => row[ci]) }); }
  });
  if (sortNative) {
    order.sort((a, b) => {
      for (let l = 0; l < catCis.length; l++) { const d = tfNaturalCompare(a.dvs[l], b.dvs[l]); if (d) return d; }
      return 0;
    });
  }
  if (!catCis.length) order = [{ key: "", labels: [""] }];
  const index = new Map(order.map((o, i) => [o.key, i]));
  return {
    count: order.length,
    levels: (catCis.length ? catCis : [null]).map((_, l) => order.map(o => o.labels[l])),
    indexOf: row => index.get(catCis.map(ci => tvText(row[ci])).join("\u0001"))
  };
}
function tvSum(vm, cats, valueCi, filter) {
  const out = new Array(cats.count).fill(null);
  vm.rows.forEach(row => {
    if (filter && !filter(row)) return;
    const n = tfDvNum(row[valueCi]);
    if (n === null) return;
    const i = cats.indexOf(row);
    if (i === undefined) return;
    out[i] = (out[i] || 0) + n;
  });
  return out;
}

/* ══════════════════════════════════════════════════════════════════════════
 * chart builders
 * ══════════════════════════════════════════════════════════════════════════ */
function tvBaseSpec(vm) {
  return {
    font: tvChartFont(vm.fmt),
    background: tvHex(vm.fmt.tableBackground(vm.dashboardName)) || "FFFFFF",
    gridlines: true
  };
}
function tvSeriesType(mark) {
  if (mark === "line") return { type: "line" };
  if (mark === "area") return { type: "area" };
  if (/^(circle|shape|square)$/.test(mark)) return { type: "line", line: false, marker: true };
  return { type: "bar" };
}

function tvPieSpec(ctx) {
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
  const spec = { ...tvBaseSpec(vm), kind: ctx.doughnut ? "doughnut" : "pie", legend: true };
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
  if (categories.levels[0].length > 200) throw new Error("too many pie slices for an Excel chart");
  const angleRef = vm.cols[angleCi] && vm.cols[angleCi].ref;
  const catRef = catCi >= 0 && vm.cols[catCi].ref;
  const labelValue = tvLabelsOn(roles, angleRef);
  const labelCat = !!catRef && roles.labelRefs.some(r => tfSameField(r, catRef));
  spec.categories = categories;
  spec.numFmt = tvNumFmt(vm, angleCi);
  spec.series = [{
    name: tvMeasureLabel(vm, angleCi),
    color: tvMarkColor(roles),
    values,
    pointColors: categories.levels[0].map((v, i) => (scale && scale(v, i)) || tvHex(TABLEAU_10[i % TABLEAU_10.length])),
    labels: labelValue || labelCat,
    labelParts: { value: labelValue, category: labelCat }
  }];
  return [spec];
}

function tvScatterSpec(ctx) {
  const { vm, roles } = ctx;
  const xm = roles.cols.values[0], ym = roles.rows.values[0];
  const scale = tvColorScale(vm, roles, ctx.markToken);
  const spec = { ...tvBaseSpec(vm), kind: "scatter", numFmt: tvNumFmt(vm, ym.ci), xNumFmt: tvNumFmt(vm, xm.ci),
                 valueTitle: tvMeasureLabel(vm, ym.ci), xTitle: tvMeasureLabel(vm, xm.ci) };
  const point = r => ({ x: tfDvNum(r[xm.ci]), y: tfDvNum(r[ym.ci]) });
  const rows = vm.rows.filter(r => { const p = point(r); return p.x !== null && p.y !== null; });
  if (rows.length > TV_MAX_POINTS) throw new Error(`${rows.length} marks – too many for an Excel scatter chart`);
  const colorCi = roles.color && !roles.color.measureNames ? roles.color.ci : -1;
  const labels = tvLabelsOn(roles, ym.ref);
  if (colorCi >= 0 && !roles.color.continuous) {
    const values = tvColorValues(vm, roles);
    if (values.length > TV_MAX_SERIES) throw new Error("too many colour values for an Excel chart");
    spec.series = values.map((v, i) => {
      const pts = rows.filter(r => tvText(r[colorCi]) === v).map(point);
      return { name: v, color: (scale && scale(v)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
               x: pts.map(p => p.x), y: pts.map(p => p.y), labels };
    });
  } else {
    const pts = rows.map(point);
    spec.series = [{
      name: tvMeasureLabel(vm, ym.ci), color: tvMarkColor(roles), x: pts.map(p => p.x), y: pts.map(p => p.y), labels,
      pointColors: colorCi >= 0 && scale ? rows.map(r => scale(tfDvNum(r[colorCi]))) : undefined
    }];
  }
  spec.legend = spec.series.length > 1;
  return [spec];
}

function tvCartesianSpecs(ctx) {
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
    if (colorCi >= 0 && colorLevel < 0) {
      const values = tvColorValues(vm, roles);
      return values.map((v, i) => ({
        name: opts.prefix ? `${label} – ${v}` : v,
        values: tvSum(vm, cats, m.ci, r => tvText(r[colorCi]) === v),
        color: (scale && scale(v)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        labels, ...opts.type
      }));
    }
    return [{ name: label, values: tvSum(vm, cats, m.ci), color: opts.color, labels,
              pointColors: opts.noPointColors ? undefined : pointColors, ...opts.type }];
  };

  const make = (series, extra) => {
    if (series.length > TV_MAX_SERIES) throw new Error(`${series.length} series – too many for an Excel chart`);
    const kinds = new Set(series.map(s => s.type));
    const kind = kinds.size > 1 || series.some(s => s.secondary) ? "combo" : ([...kinds][0] || "bar");
    const stackable = series.every(s => s.type === "bar" || s.type === "area");
    const stacked = stackable && series.length > 1 &&
      ((colorCi >= 0 && colorLevel < 0) || (mvMode && color && color.measureNames && roles.measureNames !== catShelf));
    return { ...tvBaseSpec(vm), kind, barDir: horizontal ? "bar" : "col", stacked,
             categories, categoryTitle, legend: series.length > 1, series, ...extra };
  };

  // Measure Values: one chart, one series per measure
  if (mvMode) {
    const type = tvSeriesType(tvMeasureMark(roles, null, sheetMark));
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
    measures.forEach((m, i) => seriesFor(m, {
      type: { ...tvSeriesType(marks[i]), secondary: measures.length === 2 && i === 1 },
      color: tvHex(TABLEAU_10[i % TABLEAU_10.length]), prefix: colorCi >= 0, noPointColors: true
    }).forEach(s => series.push(s)));
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

/* ── histogram: the bar chart with touching bars ─────────────────────────── */
function tvHistogramSpecs(ctx) {
  return tvCartesianSpecs(ctx).map(s => ({ ...s, gapWidth: 0 }));
}

/* value axis + category dimensions shared by waterfall / box plot */
function tvAxisLayout(ctx) {
  const { roles } = ctx;
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) throw new Error("no continuous measure axis to chart");
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  return { valueShelf, horizontal: valueShelf === "cols", measure: roles[valueShelf].values[0],
           catDims: [...roles[valueShelf].dims, ...roles[catShelf].dims] };
}

/* ── waterfall: Gantt bars sized by -measure on a running total →
 *    stacked columns: invisible base + increase / decrease ─────────────────── */
function tvWaterfallSpec(ctx) {
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
function tvQuantile(sorted, p) {                      // linear interpolation, as Tableau / QUARTILE.INC
  const pos = (sorted.length - 1) * p, lo = Math.floor(pos), hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}
function tvBoxPlotSpec(ctx) {
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
const TV_GANTT_MAX_SERIES = 250;
function tvExcelSerial(dv) {
  if (tfIsNull(dv)) return null;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  if (typeof v === "number") return v;                         // already a serial / numeric axis
  const t = v instanceof Date ? v.getTime() : Date.parse(String(v));
  return isNaN(t) ? null : t / 86400000 + 25569;               // days since 1899-12-30 (UTC)
}
function tvGanttSpec(ctx) {
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
function tvTreemapSpec(ctx) {
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
  return [{
    ...tvBaseSpec(vm), kind: "treemap", legend: false,
    categories: { names: levelCis.map(ci => tvMeasureLabel(vm, ci)), levels: cats.levels.map(lv => keep.map(i => lv[i])) },
    numFmt: tvNumFmt(vm, roles.size),
    series: [{ name: tvMeasureLabel(vm, roles.size), values: keep.map(i => sizes[i]), color: tvMarkColor(roles),
               pointColors: pointColors ? keep.map(i => pointColors[i]) : undefined, labels: true }]
  }];
}

function buildExcelChartSpecs(visualModel, model) {
  const vm = visualModel.viewModel;
  if (!vm || !vm.rows.length) throw new Error("visual has no data rows");
  const roles = tvRoles(vm, model, visualModel.source && visualModel.source.visualSpec);
  const markToken = visualModel.metadata.markToken || "";
  const ctx = { vm, roles, markToken, type: visualModel.type };
  if (roles.source === "summary" && /^(circle|shape|square)$/.test(markToken)) {
    throw new Error("mark layout unknown – load the workbook file for this visual");
  }
  let specs;
  if (visualModel.type === VISUAL_TYPES.HISTOGRAM) specs = tvHistogramSpecs(ctx);
  else if (visualModel.type === VISUAL_TYPES.WATERFALL) specs = tvWaterfallSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.BOXPLOT) specs = tvBoxPlotSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.GANTT) specs = tvGanttSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.TREEMAP) specs = tvTreemapSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.PIE) {
    ctx.doughnut = roles.panes.filter(p => tvMarkToken(p.markClass) === "pie").length >= 2;
    specs = tvPieSpec(ctx);
  } else if (visualModel.type === VISUAL_TYPES.SCATTER && roles.rows.values.length && roles.cols.values.length) {
    specs = tvScatterSpec(ctx);
  } else {
    specs = tvCartesianSpecs(ctx);
  }
  specs.forEach(s => { s.name = visualModel.metadata.worksheetName; s.rolesSource = roles.source; });
  return specs;
}

if (typeof module !== "undefined" && module.exports) module.exports = { buildExcelChartSpecs, tvRoles, tvCategories };
