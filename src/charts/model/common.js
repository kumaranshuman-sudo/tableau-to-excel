/* Shared helpers for building renderer-neutral Excel chart specs. */
import { tfDvNum, tfDvText, tfIsNull, tfNaturalCompare } from "../../data/values.js";
import { tfBuildColorScale } from "../../format/color-scale.js";
import { tfArgb } from "../../format/colors.js";
import { tfFormatDateLabel } from "../../format/date-format.js";
import { tfExcelFont } from "../../format/excel-style.js";
import { inferExcelNumFmt, tableauToExcelNumFmt } from "../../format/number-format.js";
import { TABLEAU_10 } from "../../format/palettes.js";
import { tfParseFieldRef, tfSameField } from "../../twb/field-ref.js";
import { tfCollect } from "../../twb/formatter.js";
import { tfBucketKey } from "../../twb/parser.js";
import { tfNorm } from "../../util.js";

"use strict";

export const TV_MAX_POINTS = 4000;

export const TV_MAX_SERIES = 60;

export const TV_DEFAULT_MARK_COLOR = "4E79A7";

export function tvHex(argb) {
  if (!argb) return null;
  const s = String(argb).replace(/^#/, "").toUpperCase();
  return s.length === 8 ? s.slice(2) : s;
}

export function tvText(dv) { return tfIsNull(dv) ? "Null" : tfDvText(dv); }

export function tvMarkToken(cls) {
  const t = String(cls || "").toLowerCase().replace(/[\s_-]+/g, "");
  return t === "automatic" ? "" : t;
}

export function tvPaneRule(pane, element, attrName) {
  const list = (pane && pane.style && pane.style.rules && pane.style.rules[element]) || [];
  const f = list.find(x => x.attr === attrName && !x.field && !x.scope);
  return f ? f.value : undefined;
}

/* ── formatting helpers ─────────────────────────────────────────────────── */
export function tvChartFont(fmt) {
  const f = tfExcelFont(fmt.headerCellStyle(null));
  return { name: f.name, size: f.size || 9, color: tvHex(f.color && f.color.argb) || "333333" };
}

/** @param {ViewModel} vm @param {number} ci @returns {string} */
export function tvNumFmt(vm, ci) {
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

/** @param {ViewModel} vm @param {number} ci @returns {string} */
export function tvMeasureLabel(vm, ci) {
  const c = vm.cols[ci];
  return (c && (c.label || c.name)) || "Value";
}

/* mark of the pane that draws this measure (dual axis / multi-pane), else the sheet's.
 * Without a workbook, the live spec has one marks card per measure, in shelf order. */
/** @param {Roles} roles @param {FieldRef | null} ref @param {string} [fallback] @param {number} [index] @param {number} [count] @returns {string} */
export function tvMeasureMark(roles, ref, fallback, index, count) {
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

/** @param {Roles} roles @param {FieldRef | null} ref @returns {boolean} */
export function tvLabelsOn(roles, ref) {
  const shown = roles.panes.some(p => tvPaneRule(p, "mark", "mark-labels-show") === "true");
  if (shown) return true;
  if (!ref) return false;
  return roles.labelRefs.some(r => tfSameField(r, ref) || (r.name === "Multiple Values")) ||
    roles.labelNames.some(n => n === tfNorm(ref.name));
}

/** @param {Roles} roles @returns {string} */
export function tvMarkColor(roles) {
  for (const p of roles.panes) {
    const c = tvHex(tfArgb(tvPaneRule(p, "mark", "mark-color")));
    if (c) return c;
  }
  return TV_DEFAULT_MARK_COLOR;
}

/* ── colour scale for the Color encoding (TWB mapping or Tableau automatic) ─ */
/**
 * @param {ViewModel} vm @param {Roles} roles @param {string} [markToken]
 * @returns {((value: any, index?: number) => string | null) | null} value → hex colour
 */
export function tvColorScale(vm, roles, markToken) {
  if (!roles.color) return null;
  const fmt = vm.fmt;
  /** @type {Partial<ReturnType<SheetFormatter["colorEncoding"]>>} */
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
/** @param {ViewModel} vm @param {Roles} roles @returns {string[]} */
export function tvColorValues(vm, roles) {
  const ci = roles.color.ci;
  const seen = new Map();
  vm.rows.forEach(r => { const t = tvText(r[ci]); if (!seen.has(t)) seen.set(t, r[ci]); });
  let values = [...seen.keys()];
  const manual = roles.color.ref ? vm.fmt.manualSortFor(roles.color.ref) : null;
  const byMeasure = !manual && roles.color.ref ? vm.fmt.measureSortFor(roles.color.ref) : null;
  if (manual) {
    const rank = new Map(manual.order.map((b, i) => [tfBucketKey(b), i]));
    values.sort((a, b) => (rank.has(tfNorm(a)) ? rank.get(tfNorm(a)) : 1e9) - (rank.has(tfNorm(b)) ? rank.get(tfNorm(b)) : 1e9));
    if (manual.direction === "DESC") values.reverse();
  } else if (byMeasure) {
    // sorted by a measure (Sort → Field: Value Ordered, ascending): its total per colour value
    const mi = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, byMeasure.measure));
    const total = new Map(values.map(v => [v, 0]));
    if (mi >= 0) vm.rows.forEach(r => { const t = tvText(r[ci]); total.set(t, (total.get(t) || 0) + (tfDvNum(r[mi]) || 0)); });
    values.sort((a, b) => (total.get(a) - total.get(b)) * (byMeasure.direction === "DESC" ? -1 : 1) || tfNaturalCompare(seen.get(a), seen.get(b)));
  } else values.sort((a, b) => tfNaturalCompare(seen.get(a), seen.get(b)));
  return values;
}

/** a header's date label format (Format → Header → Dates: "iLLLLL" → J F M …)
 * @param {ViewModel} vm @param {number} ci @returns {(text: string, dv: DataValue) => string} */
export function tvLabelFormatter(vm, ci) {
  const col = vm.cols[ci];
  const raw = col && col.ref && vm.fmt.labelFormat ? vm.fmt.labelFormat(col.ref) : null;
  if (!raw) return t => t;
  return (text, dv) => tfFormatDateLabel(raw, text, dv) ?? text;
}

/**
 * A marks card's label as lines of text segments: its custom label (Label → Text), else one line per Text
 * field; each field token replaced by valueOf(ref). A token split over runs still resolves, each segment
 * keeps the font of the run it starts in.
 * @param {any} pane @param {(ref: FieldRef) => string} valueOf
 * @returns {{ text: string, props: Record<string, any> }[][]}
 */
export function tvPaneLabel(pane, valueOf) {
  const runs = pane.labelRuns.length ? pane.labelRuns.map(r => ({ text: String(r.text), props: r.props || {} }))
    : pane.encodings.filter(e => e.channel === "text" || e.channel === "label")
        .flatMap((e, i) => [...(i ? [{ text: "\n", props: {} }] : []), { text: `<${e.field.raw}>`, props: {} }]);
  const chars = runs.flatMap(r => [...r.text].map(ch => ({ ch, props: r.props })));
  const text = chars.map(c => c.ch).join("");
  /** @type {{ text: string, props: Record<string, any> }[][]} */
  const lines = [[]];
  const push = (t, props) => {
    if (!t) return;
    const line = lines[lines.length - 1], last = line[line.length - 1];
    if (last && last.props === props) last.text += t; else line.push({ text: t, props });
  };
  const re = /<([^<>]+)>|Æ[ \t]*(?:\r?\n)?|\r?\n/g;
  let at = 0;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    for (let k = at; k < m.index; k++) push(chars[k].ch, chars[k].props);
    if (m[1] !== undefined) { const ref = tfParseFieldRef(m[1]); push(ref ? valueOf(ref) : m[0], chars[m.index].props); }
    else lines.push([]);
    at = m.index + m[0].length;
  }
  for (let k = at; k < text.length; k++) push(chars[k].ch, chars[k].props);
  return lines;
}

/* ── category axis: distinct label tuples in view order ─────────────────── */
/**
 * @param {ViewModel} vm
 * @param {number[]} catCis category columns, outer → inner
 * @param {boolean} sortNative sort by value (continuous / date axes) instead of view order
 * @returns {{ count: number, levels: string[][], indexOf: (row: DataValue[]) => number | undefined }}
 */
export function tvCategories(vm, catCis, sortNative) {
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
  // shown labels follow each header's date format (J F M …); the keys stay the raw values, so January,
  // June and July remain three categories even when all are labelled "J"
  const shown = catCis.map(ci => tvLabelFormatter(vm, ci));
  return {
    count: order.length,
    levels: (catCis.length ? catCis : [null]).map((_, l) => order.map(o => catCis.length ? shown[l](o.labels[l], o.dvs[l]) : o.labels[l])),
    indexOf: row => index.get(catCis.map(ci => tvText(row[ci])).join("\u0001"))
  };
}

/**
 * @param {ViewModel} vm
 * @param {{ count: number, indexOf: (row: DataValue[]) => number | undefined }} cats
 * @param {number} valueCi
 * @param {(row: DataValue[]) => boolean} [filter]
 * @returns {(number | null)[]} one total per category, null = no value
 */
export function tvSum(vm, cats, valueCi, filter) {
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
/** @param {ViewModel} vm @returns {Pick<ChartSpec, "font" | "background" | "gridlines">} */
export function tvBaseSpec(vm) {
  return {
    font: tvChartFont(vm.fmt),
    // Worksheet shading: its colour, none (see-through: the container shows) or Tableau's white
    background: vm.fmt.sheetShading() === "none" ? null : tvHex(vm.fmt.tableBackground()) || "FFFFFF",
    gridlines: true
  };
}

/** @param {string} mark @returns {Pick<ChartSeries, "type" | "line" | "marker">} */
export function tvSeriesType(mark) {
  if (mark === "line") return { type: "line" };
  if (mark === "area") return { type: "area" };
  if (/^(circle|shape|square)$/.test(mark)) return { type: "line", line: false, marker: true };
  return { type: "bar" };
}
