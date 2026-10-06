/* A table built from marks (one constant axis per column) drawn in cells: each column takes the look of the
 * marks in its pane – circles / squares as the cell's fill, arrow shapes as a glyph before the label, bars as
 * Excel data bars – coloured as Tableau colours them. */
import { tfDvNum, tfDvText, tfIsNull } from "../data/values.js";
import { tfBuildColorScale } from "../format/color-scale.js";
import { tfArgb } from "../format/colors.js";
import { tfSplitSections } from "../format/number-format.js";
import { tfSameField } from "../twb/field-ref.js";
import { tfBucketKey } from "../twb/parser.js";

const DEFAULT_MARK = "FF4E79A7";                       // Tableau's default mark colour

/* Tableau's "Arrows" shape palettes: the file number is the direction, the same in every set */
const ARROW_BY_INDEX = ["", "↓", "↘", "→", "↗", "↑", "↖", "←", "↙"];
/* Tableau's own shapes (":filled/circle" …) */
const BUILT_IN = { circle: "●", square: "■", triangle: "▲", diamond: "◆", plus: "+", cross: "✕", asterisk: "✱",
                   "dot-circle": "◉", "dot-square": "▣" };

/** the character drawn for a shape: an arrow's direction, or the shape itself @param {string} name */
export function tvShapeGlyph(name) {
  if (!name) return null;
  const arrow = name.match(/^Arrows\/\w+-(\d)\.png$/i);
  if (arrow) return ARROW_BY_INDEX[+arrow[1]] || null;
  const own = name.match(/^:(filled\/)?([a-z-]+)$/i);
  if (own) {
    const g = BUILT_IN[own[2].toLowerCase()];
    if (!g) return null;
    return own[1] ? g : ({ "●": "○", "■": "□", "▲": "△", "◆": "◇" })[g] || g;
  }
  return "●";                                           // a custom picture: a dot in its colour
}

/** a number format with a glyph before the value in every section (two at least: no minus before the glyph)
 * @param {string | undefined} fmt @param {string} glyph */
export function tvGlyphNumFmt(fmt, glyph) {
  const secs = (tfSplitSections(fmt || "General") || ["General"]).slice(0, 3).map(s => s || "General");
  if (secs.length === 1) secs.push("-" + secs[0]);
  return secs.map(s => `"${glyph} "${s}`).join(";");
}

/**
 * @typedef {{ kind: "fill" | "glyph" | "bar", colorAt: (row: number) => string | null,
 *             glyphAt: (row: number) => string | null, group: string }} MarkColumn
 * kind fill: circles / squares / shapes without a Shape field; glyph: a Shape field (arrows);
 * bar: bars, sized by their label value. group: columns whose bars share one scale (same colour field).
 */

/**
 * How each value column of a table built from marks looks.
 * @param {SheetFormatter} fmt @param {ViewColumn[]} cols @param {any[][]} rows @param {number[]} order
 * @returns {Map<number, MarkColumn>}
 */
export function buildMarkPlan(fmt, cols, rows, order) {
  /** @type {Map<number, MarkColumn>} */
  const plan = new Map();
  const panes = fmt.panesInOrder();
  const colOf = ref => cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
  order.forEach(ci => {
    const c = cols[ci];
    if (c.isHeader || !c.ref) return;
    const pp = panes.find(p => p.refs.some(r => tfSameField(r, c.ref)));
    if (!pp) return;
    const pane = pp.pane, mark = String(pane.markClass || "").toLowerCase();
    if (!/^(circle|square|shape|bar)$/.test(mark)) return;
    // colour: the pane's Color field, else its Marks card colour
    const enc = fmt.colorEncodingOf(pane);
    const fixedRule = ((pane.style && pane.style.rules && pane.style.rules.mark) || []).find(f => f.attr === "mark-color" && !f.field);
    const fixed = (fixedRule && tfArgb(fixedRule.value)) || DEFAULT_MARK;
    let colorAt = () => fixed;
    if (enc) {
      const k = colOf(enc.ref);
      if (k >= 0) {
        const vals = rows.map(r => tfIsNull(r[k]) ? null : enc.continuous ? tfDvNum(r[k]) : tfDvText(r[k]));
        const scale = tfBuildColorScale(fmt, enc, vals);
        if (scale) { const colors = vals.map(v => v === null ? null : scale(v)); colorAt = i => colors[i] || fixed; }
      }
    }
    const shapeEnc = pane.encodings.find(e => e.channel === "shape" && e.field);
    const group = enc ? `color:${enc.ref.inner}` : `col:${ci}`;
    if (mark === "bar") {
      plan.set(ci, { kind: "bar", colorAt, glyphAt: () => null, group });
    } else if (mark === "shape" && shapeEnc) {
      const map = fmt.shapeMap(shapeEnc.field) || {};
      const k = colOf(shapeEnc.field);
      const glyphs = rows.map(r => k >= 0 ? tvShapeGlyph(map[tfBucketKey(tfDvText(r[k]))]) : null);
      plan.set(ci, { kind: "glyph", colorAt, glyphAt: i => glyphs[i] || null, group });
    } else {
      plan.set(ci, { kind: "fill", colorAt, glyphAt: () => null, group });
    }
  });
  return plan;
}

/**
 * Data bars for the bar columns, written after the cells: one rule per cell when the colour varies by row.
 * Columns of one group share a scale; a column of negative values (the left half of a diverging bar) grows
 * leftwards from its right edge.
 * @param {import("exceljs").Worksheet} worksheet @param {Map<number, MarkColumn>} plan
 * @param {ViewColumn[]} cols @param {any[][]} rows @param {number[]} order
 * @param {number} firstRow 0-based sheet row of the first data row @param {number} C 0-based first column
 */
export function writeMarkBars(worksheet, plan, cols, rows, order, firstRow, C) {
  const bars = order.map((ci, k) => ({ ci, k, m: plan.get(ci) })).filter(x => x.m && x.m.kind === "bar");
  /** @type {Map<string, number>} */
  const maxOf = new Map();
  bars.forEach(({ ci, m }) => rows.forEach(r => {
    const v = tfDvNum(r[ci]);
    if (v !== null) maxOf.set(m.group, Math.max(maxOf.get(m.group) || 0, Math.abs(v)));
  }));
  bars.forEach(({ ci, k, m }) => {
    const M = maxOf.get(m.group);
    if (!M) return;
    const values = rows.map(r => tfDvNum(r[ci]));
    const negative = values.some(v => v !== null && v < 0) && !values.some(v => v !== null && v > 0);
    const cfvo = negative ? [{ type: "num", value: -M }, { type: "num", value: 0 }] : [{ type: "num", value: 0 }, { type: "num", value: M }];
    const rule = argb => ({ type: "dataBar", gradient: false, border: false, minLength: 0, maxLength: 100, cfvo,
      color: { argb }, negativeFillColor: { argb }, negativeBarColorSameAsPositive: true, axisPosition: "auto" });
    const addr = i => worksheet.getCell(firstRow + i + 1, C + k + 1).address;
    const colors = values.map((v, i) => v === null ? null : m.colorAt(i));
    const one = colors.filter(Boolean);
    if (one.length && one.every(x => x === one[0])) {
      worksheet.addConditionalFormatting(/** @type {any} */ ({ ref: `${addr(0)}:${addr(rows.length - 1)}`, rules: [rule(one[0])] }));
    } else {
      colors.forEach((argb, i) => {
        if (argb) worksheet.addConditionalFormatting(/** @type {any} */ ({ ref: addr(i), rules: [rule(argb)] }));
      });
    }
  });
}
