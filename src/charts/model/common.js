/* Shared helpers for building renderer-neutral Excel chart specs. */
import { tfDvNum, tfDvText, tfIsNull, tfNaturalCompare } from "../../data/values.js";
import { tfBuildColorScale } from "../../format/color-scale.js";
import { tfArgb } from "../../format/colors.js";
import { tfExcelFont } from "../../format/excel-style.js";
import { inferExcelNumFmt, tableauToExcelNumFmt } from "../../format/number-format.js";
import { TABLEAU_10 } from "../../format/palettes.js";
import { tfParseFieldRef, tfSameField } from "../../twb/field-ref.js";
import { tfCollect } from "../../twb/formatter.js";
import { tfBucketKey } from "../../twb/parser.js";
import { tfNorm } from "../../util.js";

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
  if (manual) {
    const rank = new Map(manual.order.map((b, i) => [tfBucketKey(b), i]));
    values.sort((a, b) => (rank.has(tfNorm(a)) ? rank.get(tfNorm(a)) : 1e9) - (rank.has(tfNorm(b)) ? rank.get(tfNorm(b)) : 1e9));
    if (manual.direction === "DESC") values.reverse();
  } else values.sort((a, b) => tfNaturalCompare(seen.get(a), seen.get(b)));
  return values;
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
  return {
    count: order.length,
    levels: (catCis.length ? catCis : [null]).map((_, l) => order.map(o => o.labels[l])),
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
    background: tvHex(vm.fmt.tableBackground(vm.dashboardName)) || "FFFFFF",
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
