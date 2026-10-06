/* DrawingML chart XML (bar / line / area / pie / scatter / bubble / combo). */
import { AX, C15, NS } from "./constants.js";
import { esc, hex, levelStarts, num } from "./xml-util.js";

export function solid(color, alpha) {
  if (!color) return "<a:noFill/>";
  return `<a:solidFill><a:srgbClr val="${hex(color)}">${alpha ? `<a:alpha val="${alpha}"/>` : ""}</a:srgbClr></a:solidFill>`;
}

export function line(color, w) { return color ? `<a:ln w="${w || 9525}">${solid(color)}</a:ln>` : `<a:ln><a:noFill/></a:ln>`; }

export function runProps(font, o, tag) {
  const sz = Math.round((o.size || font.size || 9) * 100);
  return `<a:${tag} lang="en-US" sz="${sz}" b="${o.bold ? 1 : 0}">${solid(o.color || font.color)}` +
    `<a:latin typeface="${esc(font.name)}"/><a:cs typeface="${esc(font.name)}"/></a:${tag}>`;
}

export function txPr(font, o = {}) {
  const def = runProps(font, o, "defRPr").replace(' lang="en-US"', "");
  // rot 0 is written too: it keeps Excel from turning labels it finds crowded
  return `<c:txPr><a:bodyPr${o.rot !== undefined ? ` rot="${o.rot}" vert="horz"` : ""}${o.noWrap ? ' wrap="none"' : ""}/><a:lstStyle/>` +
    `<a:p><a:pPr>${def}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

export function title(text, font, vertical) {
  if (!text) return "";
  const def = runProps(font, {}, "defRPr").replace(' lang="en-US"', "");
  return `<c:title><c:tx><c:rich><a:bodyPr${vertical ? ' rot="-5400000" vert="horz"' : ""}/><a:lstStyle/>` +
    `<a:p><a:pPr>${def}</a:pPr><a:r>${runProps(font, {}, "rPr")}<a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx>` +
    `<c:overlay val="0"/></c:title>`;
}

export function strCache(values) {
  return `<c:ptCount val="${values.length}"/>` +
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("");
}

export function numCache(values) {
  return `<c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>` +
    values.map((v, i) => num(v) === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
}

export function serTx(ref, name) {
  return `<c:tx><c:strRef><c:f>${esc(ref)}</c:f><c:strCache>${strCache([name])}</c:strCache></c:strRef></c:tx>`;
}

/** @param {string} ref @param {any[][]} levels @param {string[] | null} [shown] labels shown instead (Tableau's
 * truncated headers): written into the chart, as Excel re-reads linked category cells when it opens the file */
export function catXml(ref, levels, shown = null) {
  if (shown && levels.length <= 1) return `<c:cat><c:strLit>${strCache(shown)}</c:strLit></c:cat>`;
  if (levels.length <= 1) {
    return `<c:cat><c:strRef><c:f>${esc(ref)}</c:f><c:strCache>${strCache(levels[0] || [])}</c:strCache></c:strRef></c:cat>`;
  }
  const starts = levelStarts(levels);
  const n = levels[0].length;
  // innermost level first, as Excel writes it
  const lvls = levels.map((lv, l) => `<c:lvl>${lv.map((v, i) =>
      (l === levels.length - 1 || starts[l][i]) ? `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>` : "").join("")}</c:lvl>`).reverse();
  return `<c:cat><c:multiLvlStrRef><c:f>${esc(ref)}</c:f><c:multiLvlStrCache><c:ptCount val="${n}"/>${lvls.join("")}` +
    `</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`;
}

export function valXml(tag, ref, values) {
  return `<c:${tag}><c:numRef><c:f>${esc(ref)}</c:f><c:numCache>${numCache(values)}</c:numCache></c:numRef></c:${tag}>`;
}

/** mark labels: the worksheet's label font (else the chart font); labels set from the workbook never wrap
 * @param {ChartSpec} spec */
export function labelTxPr(spec) {
  const lf = spec.labelFont;
  if (!lf) return txPr(spec.font, spec.labelPos ? { noWrap: true } : {});
  return txPr({ name: lf.name || spec.font.name, size: lf.size || spec.font.size, color: lf.color || spec.font.color },
              { bold: lf.bold, noWrap: true });
}

/** @param {ChartSpec} spec @param {ChartSeries} s @param {string | null} pos @param {number[]} [hidden] points without a label */
export function dLbls(spec, s, pos, hidden) {
  if (!s.labels) return "";
  const drop = (hidden || []).map(i => `<c:dLbl><c:idx val="${i}"/><c:delete val="1"/></c:dLbl>`).join("");
  if (s.labelTexts) {
    // a doughnut takes no label position (Excel refuses the file)
    return `<c:dLbls>${drop}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${labelTxPr(spec)}` +
      `${spec.kind === "doughnut" ? "" : `<c:dLblPos val="${pos || "r"}"/>`}<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>` +
      `<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>` +
      `<c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="${C15}">` +
      `<c15:showDataLabelsRange val="1"/></c:ext></c:extLst></c:dLbls>`;
  }
  const fmt = s.labelNumFmt || (s.secondary ? spec.secondaryNumFmt : spec.numFmt);
  const showVal = s.labelParts ? (s.labelParts.value ? 1 : 0) : 1;
  const showCat = s.labelParts && s.labelParts.category ? 1 : 0;
  return `<c:dLbls>${fmt ? `<c:numFmt formatCode="${esc(fmt)}" sourceLinked="0"/>` : ""}` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${labelTxPr(spec)}` +
    `${pos ? `<c:dLblPos val="${pos}"/>` : ""}<c:showLegendKey val="0"/><c:showVal val="${showVal}"/>` +
    `<c:showCatName val="${showCat}"/><c:showSerName val="0"/><c:showPercent val="${s.labelParts && s.labelParts.percent ? 1 : 0}"/><c:showBubbleSize val="0"/>` +
    `${spec.kind === "pie" || spec.kind === "doughnut" ? '<c:showLeaderLines val="1"/>' : ""}</c:dLbls>`;
}

export function markerXml(symbol, color, size) {
  if (!symbol || symbol === "none") return `<c:marker><c:symbol val="none"/></c:marker>`;
  return `<c:marker><c:symbol val="${symbol}"/><c:size val="${size || 6}"/>` +
    `<c:spPr>${solid(color, 85000)}${line(color, 9525)}</c:spPr></c:marker>`;
}

/* a reference line's stroke: colour with its opacity, width (px), dashes; none when switched off */
export function refLineLn(rl) {
  if (rl.hidden) return `<a:ln><a:noFill/></a:ln>`;
  return `<a:ln w="${Math.round(rl.width * 9525)}"><a:solidFill><a:srgbClr val="${hex(rl.color)}">` +
    `${rl.alpha < 1 ? `<a:alpha val="${Math.round(rl.alpha * 100000)}"/>` : ""}</a:srgbClr></a:solidFill>` +
    `${rl.dash ? '<a:prstDash val="dash"/>' : ""}</a:ln>`;
}

/* the line's label on one of its points (idx): an Excel number format on the line's value ("Average",
 * "Avg. $1.2M"); valueFlag picks the value Excel formats (showVal = y, showCatName = x of an XY line) */
export function refLineLabel(spec, rl, pos, valueFlag = "showVal", idx = 0) {
  if (!rl.labelFmt) return "";
  const font = { name: spec.font.name, size: spec.font.size, color: rl.font.color || "555555" };
  const flags = on => ["showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"]
    .map(f => `<c:${f} val="${on && f === valueFlag ? 1 : 0}"/>`).join("");
  return `<c:dLbls><c:dLbl><c:idx val="${idx}"/><c:numFmt formatCode="${esc(rl.labelFmt)}" sourceLinked="0"/>` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(font, { bold: rl.font.bold, noWrap: true })}` +
    `<c:dLblPos val="${pos}"/>${flags(true)}</c:dLbl>${flags(false)}</c:dLbls>`;
}

/* a vertical reference line across horizontal bars: XY points (value, 0) → (value, 1) on the hidden x2 / y2 axes */
export function refLineXY(spec, rl, k) {
  const lit = vals => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
    vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("") + `</c:numLit>`;
  return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/><c:tx><c:v>${esc((rl.labelFmt || "Reference line").replace(/"/g, ""))}</c:v></c:tx>` +
    `<c:spPr>${refLineLn(rl)}</c:spPr><c:marker><c:symbol val="none"/></c:marker>${refLineLabel(spec, rl, "r", "showCatName")}` +
    `<c:xVal>${lit([rl.value, rl.value])}</c:xVal><c:yVal>${lit([0, 1])}</c:yVal><c:smooth val="0"/></c:ser>`;
}

/* a hidden value axis pinned to [min, max] (the XY overlay's x2 / y2) */
export function hiddenValAx(id, cross, pos, min, max) {
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/><c:max val="${max}"/><c:min val="${min}"/></c:scaling>` +
    `<c:delete val="1"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="none"/><c:crossAx val="${cross}"/><c:crosses val="max"/>` +
    `<c:crossBetween val="midCat"/></c:valAx>`;
}

/** Tableau-like axis ticks: a "nice" step (1 / 2 / 2.5 / 5 × 10ⁿ) giving about `ticks` intervals
 * @param {number} range @param {number} ticks @returns {number | null} */
export function niceUnit(range, ticks) {
  if (!(range > 0) || !(ticks > 0)) return null;
  const raw = range / ticks, pow = Math.pow(10, Math.floor(Math.log10(raw))), n = raw / pow;
  return +((n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * pow).toPrecision(12);
}

/**
 * The primary value axis as Tableau draws it: ticks at the workbook's spacing, else a "nice" step about every
 * 55 px (110 px across); the range starts at zero (without "Include zero": at a tick below the data) and ends
 * just past the data, so no empty tick sits above the marks. Stacked bars / areas count as their totals.
 * @param {ChartSpec} spec @param {{ w: number, h: number } | null} axisPlot approximate plot area in px
 * @returns {{ fixed: { min?: number, max?: number }, unit: number | null }}
 */
export function valueRange(spec, axisPlot) {
  const fixed = { min: spec.valueMin, max: spec.valueMax };
  const vals = [], stacks = new Map();
  spec.series.filter(s => !s.secondary && s.values).forEach(s => {
    const type = s.type || spec.kind;
    if (spec.stacked && !s.refLine && (type === "bar" || type === "area")) {
      const g = stacks.get(type) || [];
      stacks.set(type, g);
      s.values.forEach((v, i) => {
        const x = num(v);
        if (x === null) return;
        const t = g[i] || (g[i] = { pos: 0, neg: 0 });
        if (x >= 0) t.pos += x; else t.neg += x;
      });
    } else s.values.forEach(v => { const x = num(v); if (x !== null) vals.push(x); });
  });
  stacks.forEach(g => g.forEach(t => { if (t) vals.push(t.pos, t.neg); }));
  (spec.refLines || []).forEach(r => vals.push(r.value));
  const ticks = axisPlot ? Math.max(2, Math.round(spec.barDir === "bar" ? axisPlot.w / 110 : axisPlot.h / 55)) : null;
  const aligned = (v, u) => Math.abs(v / u - Math.round(v / u)) < 1e-9;
  const round = v => +v.toPrecision(12);
  const pinned = fixed.min !== undefined && fixed.max !== undefined;
  // Excel labels ticks from the axis minimum: a fixed minimum off the tick step starts at the step below
  // it, so the labels are round numbers as on Tableau's axis (a hidden axis keeps the exact range)
  const dates = /[dmy]/i.test(spec.numFmt || "") && !/[0#]/.test(spec.numFmt || "");   // Gantt: Excel's own date steps
  const settle = (range, unit) => {
    if (!unit || range.min === undefined || aligned(range.min, unit)) return unit;
    if (spec.valueAxisHidden || dates) return null;
    range.min = round(Math.floor(range.min / unit) * unit);
    return unit;
  };
  // the workbook's spacing, unless it would crowd the axis with ticks (data far beyond what it was set for)
  const spacing = span => spec.valueMajorUnit && span / spec.valueMajorUnit <= 40 ? spec.valueMajorUnit : null;
  if (!vals.length || pinned || !(ticks || spec.includeZero === false || (spec.refLines && spec.refLines.length))) {
    const span = pinned ? fixed.max - fixed.min : 0;
    const unit = pinned ? spacing(span) || niceUnit(span, ticks || 5) : spec.valueMajorUnit || null;
    return { fixed, unit: settle(fixed, unit) };
  }
  const zero = spec.includeZero !== false;
  let lo = fixed.min ?? Math.min(...vals), hi = fixed.max ?? Math.max(...vals);
  if (zero) { lo = Math.min(0, lo); hi = Math.max(0, hi); }
  if (!(hi > lo)) hi = lo + (Math.abs(lo) || 1);
  const unit = spacing(hi - lo) || niceUnit(hi - lo, ticks || 4);
  // labels above the marks need room inside the plot
  const outside = spec.series.some(s => s.labels && !s.refLine) && !/^(inBase|ctr|inEnd)$/.test(spec.labelPos || "");
  const pad = (hi - lo) * (outside ? 0.12 : 0.04);
  const range = { min: fixed.min ?? (zero && lo >= 0 ? 0 : round(Math.floor((lo - pad) / unit) * unit)),
                  max: fixed.max ?? (zero && hi <= 0 ? 0 : round(hi + pad)) };
  return { fixed: range, unit: settle(range, unit) };
}

/** data labels taken from cells (scatter / bubble: Label = a dimension; waterfall steps over their bars)
 * @param {ChartSeries} s @param {ChartRefs["series"][number]} r */
export function labelRangeXml(s, r) {
  if (!s.labelTexts || !r.lbl) return "";
  return `<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}" xmlns:c15="${C15}"><c15:datalabelsRange>` +
    `<c15:f>${esc(r.lbl)}</c15:f><c15:dlblRangeCache>${strCache(s.labelTexts)}</c15:dlblRangeCache>` +
    `</c15:datalabelsRange></c:ext></c:extLst>`;
}

/** @param {ChartSpec} spec @param {ChartSeries} s @param {number} k @param {ChartRefs} refs @param {string} type
 *  @param {number[]} [hiddenLabels] points whose data label is left out */
export function seriesXml(spec, s, k, refs, type, hiddenLabels) {
  const r = refs.series[k];
  const head = `<c:idx val="${k}"/><c:order val="${k}"/>${serTx(r.tx, s.name)}`;
  const pc = s.pointColors || [];
  const levels = spec.categories ? spec.categories.levels : [];
  // the workbook's opacity (Color → Opacity), in DrawingML's 1/1000 %
  const opacity = s.alpha !== undefined && s.alpha < 1 ? Math.round(Math.max(0, s.alpha) * 100000) : null;
  if (type === "bar") {
    const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>` +
      `<c:spPr>${solid(c, opacity)}<a:ln><a:noFill/></a:ln></c:spPr></c:dPt>` : "").join("");
    const pos = spec.labelPos || (spec.stacked ? null : "outEnd");
    return `<c:ser>${head}<c:spPr>${solid(s.color, opacity)}<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>` +
      `${dpts}${dLbls(spec, s, pos)}${catXml(refs.cat, levels, spec.categoryShown)}${valXml("val", r.val, s.values)}${labelRangeXml(s, r)}</c:ser>`;
  }
  if (type === "line" && s.refLine) {         // reference line: flat, no markers; label above the 2nd point, as Tableau's
    const at = Math.min(1, Math.max(0, (s.values || []).length - 1));
    return `<c:ser>${head}<c:spPr>${refLineLn(s.refLine)}</c:spPr><c:marker><c:symbol val="none"/></c:marker>` +
      `${refLineLabel(spec, s.refLine, "t", "showVal", at)}${catXml(refs.cat, levels, spec.categoryShown)}${valXml("val", r.val, s.values)}<c:smooth val="0"/></c:ser>`;
  }
  if (type === "line") {
    const lineSp = s.line === false ? `<a:ln w="28575"><a:noFill/></a:ln>`
      : `<a:ln w="22225" cap="rnd">${solid(s.color, opacity)}<a:round/></a:ln>`;
    const symbol = s.marker ? (s.markerSymbol || "circle") : "none";
    const size = s.markerSize || 7;
    const dpts = s.marker ? pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/>${markerXml(symbol, c, size)}<c:bubble3D val="0"/></c:dPt>` : "").join("") : "";
    return `<c:ser>${head}<c:spPr>${lineSp}</c:spPr>${markerXml(symbol, s.color, size)}${dpts}` +
      `${dLbls(spec, s, "t", hiddenLabels)}${catXml(refs.cat, levels, spec.categoryShown)}${valXml("val", r.val, s.values)}<c:smooth val="0"/>${labelRangeXml(s, r)}</c:ser>`;
  }
  if (type === "area") {
    return `<c:ser>${head}<c:spPr>${solid(s.color, opacity !== null || s.alpha !== undefined ? opacity : spec.stacked ? null : 75000)}<a:ln><a:noFill/></a:ln></c:spPr>` +
      `${dLbls(spec, s, null)}${catXml(refs.cat, levels, spec.categoryShown)}${valXml("val", r.val, s.values)}</c:ser>`;
  }
  if (type === "pie") {
    const dpts = pc.map((c, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/>` +
      `<c:spPr>${solid(c || s.color)}${line("FFFFFF", 12700)}</c:spPr></c:dPt>`).join("");
    return `<c:ser>${head}${dpts}${dLbls(spec, s, spec.kind === "pie" ? "bestFit" : null)}` +
      `${catXml(refs.cat, levels, spec.categoryShown)}${valXml("val", r.val, s.values)}${labelRangeXml(s, r)}</c:ser>`;
  }
  if (type === "bubble") {
    // packed bubbles: opaque with a white outline like Tableau; map marks slightly see-through
    const fill = c => `${solid(c, spec.packed ? null : 80000)}${line("FFFFFF", 9525)}`;
    const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>` +
      `<c:spPr>${fill(c)}</c:spPr></c:dPt>` : "").join("");
    return `<c:ser>${head}<c:spPr>${fill(s.color)}</c:spPr><c:invertIfNegative val="0"/>${dpts}` +
      `${dLbls(spec, s, spec.packed ? "ctr" : "r", hiddenLabels)}${valXml("xVal", r.x, s.x)}${valXml("yVal", r.y, s.y)}` +
      `${valXml("bubbleSize", r.size, s.size)}<c:bubble3D val="0"/>${labelRangeXml(s, r)}</c:ser>`;
  }
  // scatter
  const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/>${markerXml("circle", c, 7)}<c:bubble3D val="0"/></c:dPt>` : "").join("");
  return `<c:ser>${head}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>${markerXml("circle", s.color, 7)}${dpts}` +
    `${dLbls(spec, s, "r")}${valXml("xVal", r.x, s.x)}${valXml("yVal", r.y, s.y)}<c:smooth val="0"/>` +
    labelRangeXml(s, r) + `</c:ser>`;
}

/** @param {ChartSpec} spec @param {number} id @param {number} cross @param {{ deleted?: boolean, rot?: number }} [o] rot: label rotation (60000ths of a degree) */
export function catAxis(spec, id, cross, o = {}) {
  const horizontal = spec.barDir === "bar";
  const multi = spec.categories && spec.categories.levels.length > 1;
  return `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="${horizontal ? "maxMin" : "minMax"}"/></c:scaling>` +
    `<c:delete val="${o.deleted ? 1 : 0}"/><c:axPos val="${horizontal ? "l" : "b"}"/>` +
    `${o.deleted ? "" : title(spec.categoryTitle, spec.font, horizontal)}` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>` +
    `<c:tickLblPos val="${o.deleted ? "none" : "low"}"/><c:spPr>${spec.axisLine === false ? "<a:ln><a:noFill/></a:ln>" : line("D4D4D4", 9525)}</c:spPr>${txPr(spec.font, { rot: o.rot })}` +
    `<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>` +
    `<c:lblOffset val="100"/>${o.rot !== undefined ? '<c:tickLblSkip val="1"/>' : ""}<c:noMultiLvlLbl val="${multi ? 0 : 1}"/></c:catAx>`;
}

/* Tableau axes "include zero" by default; Excel would otherwise auto-scale from a non-zero minimum */
export function zeroScaling(values) {
  const nums = values.filter(v => num(v) !== null);
  if (!nums.length) return "";
  if (Math.min(...nums) >= 0) return `<c:min val="0"/>`;
  if (Math.max(...nums) <= 0) return `<c:max val="0"/>`;
  return "";
}

/* axis ticks: Tableau's automatic axis drops decimals and abbreviates thousands (100K, 1.5M);
 * labels and the data cells keep the full number format */
export function axisFmt(fmt, values) {
  const nums = (values || []).filter(v => num(v) !== null).map(Math.abs);
  if (!fmt || !nums.length) return fmt || "General";
  const max = Math.max(...nums);
  if (max < 1000) {
    if (/%/.test(fmt)) return fmt;
    if (max >= 10) return fmt.replace(/0.0+/g, "0");
    return fmt.replace(/0.0{2,}/g, "0.0");
  }
  const plain = fmt.match(/^("[^"]*")?(#,##0|0)(.0+)?$/);
  if (plain && max >= 1e6) return (plain[1] || "") + '#,##0.0,,"M"';
  if (plain && max >= 1e4) return (plain[1] || "") + '#,##0,"K"';
  return fmt.replace(/0.0+/g, "0");
}

/**
 * @param {ChartSpec} spec @param {number} id @param {number} cross
 * @param {{ pos?: string, crosses?: string, grid?: boolean, fixed?: { min?: number, max?: number }, title?: string,
 *           numFmt?: string, values?: (number | null)[], deleted?: boolean, lowLabels?: boolean, midCat?: boolean,
 *           majorUnit?: number | null, tickFmt?: string }} [o] tickFmt = the workbook's own tick format, used as is
 */
export function valAxis(spec, id, cross, o = {}) {
  const horizontal = spec.barDir === "bar" && spec.kind !== "scatter";
  const pos = o.pos || (horizontal ? "b" : "l");
  const crosses = o.crosses || (horizontal ? "max" : "autoZero");
  const grid = o.grid !== false && spec.gridlines !== false
    ? `<c:majorGridlines><c:spPr>${line("EBEBEB", 9525)}</c:spPr></c:majorGridlines>` : "";
  const fixed = o.fixed || {};                       // CT_Scaling order: orientation, max, min
  const scale = fixed.max !== undefined || fixed.min !== undefined
    ? (fixed.max !== undefined ? `<c:max val="${fixed.max}"/>` : "") + (fixed.min !== undefined ? `<c:min val="${fixed.min}"/>` : "")
    : spec.includeZero === false || !o.values ? "" : zeroScaling(o.values);
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${scale}</c:scaling><c:delete val="${o.deleted ? 1 : 0}"/>` +
    `<c:axPos val="${pos}"/>${grid}${title(o.title, spec.font, pos === "l" || pos === "r")}` +
    `<c:numFmt formatCode="${esc(o.tickFmt || axisFmt(o.numFmt || "General", o.values))}" sourceLinked="0"/><c:majorTickMark val="none"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="${o.lowLabels ? "low" : "nextTo"}"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>` +
    `${txPr(spec.font)}<c:crossAx val="${cross}"/><c:crosses val="${crosses}"/>` +
    `<c:crossBetween val="${o.midCat ? "midCat" : "between"}"/>${o.majorUnit ? `<c:majorUnit val="${o.majorUnit}"/>` : ""}</c:valAx>`;
}

/* Excel's largest bubble: diameter / the plot area's shorter side for a bubble scale s (1 = 100 %).
 * Measured in Excel (one bubble, scales 10–300 %, three plot shapes): 0.03 at 10 %, 0.234 at 100 %,
 * 0.485 at 300 % – A·s / (1 + B·s) fits within 1 %. */
const BUBBLE_A = 0.3014, BUBBLE_B = 0.288;
export const bubbleRatio = s => BUBBLE_A * s / (1 + BUBBLE_B * s);
export const bubbleScaleFor = ratio => ratio >= bubbleRatio(3) ? 3 : Math.max(0.01, ratio / (BUBBLE_A - BUBBLE_B * ratio));

/**
 * Maps and packed bubbles: fixed axis ranges (and the bubble scale) for the plot area's size in px, so
 * map marks keep the geography's shape and packed bubbles touch without overlapping.
 * @param {ChartSpec} spec @param {{ w: number, h: number } | null} plot
 * @returns {{ bubbleScale: number, unit?: number, x?: { min: number, max: number }, y?: { min: number, max: number } }}
 */
export function xyFit(spec, plot) {
  if (!plot || !(spec.packed || spec.aspect)) return { bubbleScale: 100 };
  const W = plot.w, H = plot.h, shorter = Math.min(W, H);
  if (spec.packed) {
    const b = spec.packed;                                    // radius units: largest bubble radius = 1
    const want = 0.98 * Math.min(W / b.w, H / b.h);           // px per radius unit that fits the packing
    const bubbleScale = Math.max(1, Math.round(100 * bubbleScaleFor(2 * want / shorter)));
    const unit = bubbleRatio(bubbleScale / 100) * shorter / 2;   // px per radius unit as Excel draws it
    return { bubbleScale, unit, x: { min: b.cx - W / 2 / unit, max: b.cx + W / 2 / unit },
             y: { min: b.cy - H / 2 / unit, max: b.cy + H / 2 / unit } };
  }
  const bubbleScale = Math.max(1, Math.round(100 * bubbleScaleFor(spec.markRatio || 0.1)));
  const xs = spec.series.flatMap(s => s.x).filter(v => num(v) !== null);
  const ys = spec.series.flatMap(s => s.y).filter(v => num(v) !== null);
  const pad = spec.kind === "bubble" ? bubbleRatio(bubbleScale / 100) * shorter / 2 + 4 : 10;   // px around the outer marks
  const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
  // degrees of latitude per px; a degree of longitude is cos(latitude) as long
  const k = Math.max((y1 - y0) / Math.max(1, H - 2 * pad), (x1 - x0) * spec.aspect.xScale / Math.max(1, W - 2 * pad), 1e-6);
  const kx = k / spec.aspect.xScale, cx = (x0 + x1) / 2, cy = (y0 + y1) / 2;
  return { bubbleScale, x: { min: cx - kx * W / 2, max: cx + kx * W / 2 }, y: { min: cy - k * H / 2, max: cy + k * H / 2 } };
}

/**
 * Tableau's column headers stay horizontal and a name too long for its slot is cut short with "..": the
 * labels shown on a vertical chart's category axis, or null when every name fits (or the labels are rotated).
 * @param {ChartSpec} spec @param {{ w: number, h: number } | null} axisPlot
 */
export function truncatedCategories(spec, axisPlot) {
  const levels = spec.categories ? spec.categories.levels : [];
  if (!axisPlot || spec.barDir === "bar" || spec.categoryRotation || spec.categoryAxisHidden || levels.length !== 1 || !levels[0].length) return null;
  const names = levels[0].map(c => String(c == null ? "" : c));
  const charPx = ((spec.font && spec.font.size) || 9) * 4 / 3 * 0.5;      // average character of the label font
  const fit = Math.max(4, Math.floor((axisPlot.w / names.length - 4) / charPx));
  const shown = names.map(n => n.length > fit ? n.slice(0, Math.max(1, fit - 2)).trimEnd() + ".." : n);
  return shown.some((s, i) => s !== names[i]) ? shown : null;
}

/** @param {ChartSpec} spec @param {ChartRefs} refs @param {{ w: number, h: number } | null} [plot] plot area in px (manual layout)
 *  @param {{ w: number, h: number } | null} [axisPlot] approximate plot area of a chart with axes, for Tableau-like tick spacing */
export function plotAreaXml(spec, refs, plot = null, axisPlot = null) {
  spec.categoryShown = truncatedCategories(spec, axisPlot);
  const k = spec.kind;
  if (k === "pie" || k === "doughnut") {
    const ser = seriesXml(spec, spec.series[0], 0, refs, "pie");
    return k === "pie"
      ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
      : `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="${spec.holeSize || 55}"/></c:doughnutChart>`;
  }
  if (k === "scatter" || k === "bubble") {
    const fit = xyFit(spec, plot);
    const sizes = spec.series.flatMap(s => s.size || []).filter(v => num(v) !== null && v > 0);
    const maxSize = sizes.length ? Math.max(...sizes) : 1;
    const charPx = (spec.font && spec.font.size || 9) * 4 / 3 * 0.55;          // average glyph width
    const tooSmall = s => !spec.packed || !s.labelTexts || !fit.unit ? [] :
      s.labelTexts.map((t, i) => String(t || "").length * charPx > 1.9 * fit.unit * Math.sqrt(Math.max(0, s.size[i]) / maxSize) ? i : -1)
        .filter(i => i >= 0);
    const sers = spec.series.map((s, i) => seriesXml(spec, s, i, refs, k, tooSmall(s))).join("");
    const group = k === "bubble"
      ? `<c:bubbleChart><c:varyColors val="0"/>${sers}<c:bubbleScale val="${fit.bubbleScale}"/><c:showNegBubbles val="0"/>` +
        `<c:sizeRepresents val="area"/><c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/></c:bubbleChart>`
      : `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}` +
        `<c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/></c:scatterChart>`;
    return group +
      valAxis(spec, AX.cat, AX.val, { pos: "b", crosses: "autoZero", title: spec.axesHidden || spec.xAxisHidden ? null : spec.xTitle,
        numFmt: spec.xNumFmt, tickFmt: spec.xAxisNumFmt, midCat: true, grid: spec.xGridlines !== false, lowLabels: true, values: spec.series.flatMap(s => s.x),
        fixed: fit.x || { min: spec.xMin, max: spec.xMax }, deleted: spec.axesHidden || spec.xAxisHidden, majorUnit: spec.xMajorUnit }) +
      valAxis(spec, AX.val, AX.cat, { pos: "l", crosses: "autoZero", title: spec.axesHidden || spec.valueAxisHidden ? null : spec.valueTitle,
        numFmt: spec.numFmt, tickFmt: spec.valueAxisNumFmt, midCat: true, lowLabels: true, values: spec.series.flatMap(s => s.y),
        fixed: fit.y || { min: spec.valueMin, max: spec.valueMax }, deleted: spec.axesHidden || spec.valueAxisHidden, majorUnit: spec.valueMajorUnit });
  }
  // labels written over the bars (waterfall steps): as Tableau, a label that would overlap one already
  // placed is left out – positions estimated from the plot size and the axis range
  const { fixed, unit } = valueRange(spec, axisPlot);
  /** @type {Map<number, number[]>} */
  const culled = new Map();
  if (axisPlot && spec.labelCull !== false && fixed.min !== undefined && fixed.max !== undefined && fixed.max > fixed.min) {
    const size = (spec.labelFont && spec.labelFont.size) || spec.font.size || 9, h = size * 4 / 3 + 2;
    spec.series.forEach((s, i) => {
      if (!s.labelTexts || (s.type || k) !== "line") return;
      const slot = axisPlot.w / Math.max(1, s.values.length), kept = [], hide = [];
      s.values.forEach((v, j) => {
        const t = s.labelTexts[j], n = num(v);
        if (n === null || !t) return;
        const w = String(t).length * size * 0.62 + 4, x = (j + 0.5) * slot;
        const y = axisPlot.h * (fixed.max - n) / (fixed.max - fixed.min) - h / 2 - 3;   // a label just above the point
        const box = { l: x - w / 2, r: x + w / 2, t: y - h / 2, b: y + h / 2 };
        if (kept.some(q => q.l < box.r && box.l < q.r && q.t < box.b && box.t < q.b)) hide.push(j); else kept.push(box);
      });
      if (hide.length) culled.set(i, hide);
    });
  }
  // bar / line / area / combo: one chart group per (type, axis)
  const groups = [];
  spec.series.forEach((s, i) => {
    const type = s.type || k;
    const key = type + (s.secondary ? "2" : "1");
    let g = groups.find(x => x.key === key);
    if (!g) groups.push(g = { key, type, secondary: !!s.secondary, items: [] });
    g.items.push(seriesXml(spec, s, i, refs, type, culled.get(i)));
  });
  const hasSecondary = groups.some(g => g.secondary);
  const xml = groups.map(g => {
    const ax = g.secondary ? `<c:axId val="${AX.cat2}"/><c:axId val="${AX.val2}"/>` : `<c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/>`;
    const sers = g.items.join("");
    if (g.type === "bar") {
      const grouping = spec.stacked ? "stacked" : "clustered";
      return `<c:barChart><c:barDir val="${spec.barDir || "col"}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers}` +
        `<c:gapWidth val="${spec.gapWidth ?? 60}"/>${spec.stacked ? '<c:overlap val="100"/>' : ""}${ax}</c:barChart>`;
    }
    if (g.type === "area") {
      return `<c:areaChart><c:grouping val="${spec.stacked ? "stacked" : "standard"}"/><c:varyColors val="0"/>${sers}${ax}</c:areaChart>`;
    }
    // box plot: high-low lines = whiskers, up/down bars between first (Q1) and last (Q3) series = box
    const box = spec.boxPlot
      ? `<c:hiLowLines><c:spPr>${line(spec.boxPlot.color, 12700)}</c:spPr></c:hiLowLines>` +
        `<c:upDownBars><c:gapWidth val="${spec.gapWidth ?? 80}"/>` +
        `<c:upBars><c:spPr>${solid(spec.boxPlot.color, 35000)}${line(spec.boxPlot.color, 12700)}</c:spPr></c:upBars>` +
        `<c:downBars><c:spPr>${solid(spec.boxPlot.color, 35000)}${line(spec.boxPlot.color, 12700)}</c:spPr></c:downBars></c:upDownBars>`
      : "";
    return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${box}<c:marker val="1"/>${ax}</c:lineChart>`;
  }).join("");
  const axisValues = secondary => spec.series.filter(s => !!s.secondary === secondary).flatMap(s => s.values);
  // reference lines across horizontal bars: XY lines on hidden x2 / y2 axes, x2 pinned to the bars' value range
  let overlay = "", overlayAxes = "";
  if (spec.barDir === "bar" && spec.refLines && spec.refLines.length && fixed.min !== undefined && fixed.max !== undefined) {
    overlay = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>` +
      spec.refLines.map((rl, i) => refLineXY(spec, rl, spec.series.length + i)).join("") +
      `<c:axId val="${AX.x2}"/><c:axId val="${AX.y2}"/></c:scatterChart>`;
    overlayAxes = hiddenValAx(AX.x2, AX.y2, "t", fixed.min, fixed.max) + hiddenValAx(AX.y2, AX.x2, "r", 0, 1);
  }
  let axes = catAxis(spec, AX.cat, AX.val, { deleted: spec.categoryAxisHidden, rot: (spec.categoryRotation || 0) * 60000 }) +
    valAxis(spec, AX.val, AX.cat, { title: spec.valueAxisHidden ? null : spec.valueTitle, numFmt: spec.numFmt, values: axisValues(false),
                                    fixed, deleted: spec.valueAxisHidden, majorUnit: unit, tickFmt: spec.valueAxisNumFmt });
  if (hasSecondary) {
    axes += catAxis(spec, AX.cat2, AX.val2, { deleted: true }) +
      valAxis(spec, AX.val2, AX.cat2, { pos: spec.barDir === "bar" ? "t" : "r", crosses: "max", grid: false,
        title: spec.secondaryAxisHidden ? null : spec.secondaryTitle, numFmt: spec.secondaryNumFmt || spec.numFmt, values: axisValues(true),
        deleted: spec.secondaryAxisHidden, tickFmt: spec.secondaryAxisNumFmt });
  }
  return xml + overlay + axes + overlayAxes;
}

/** chart part XML (xl/charts/chartN.xml)
 * @param {ChartSpec} spec @param {ChartRefs} refs
 * @param {{ widthPx?: number, heightPx?: number }} [size] the chart frame, for maps / packed bubbles
 * @returns {string} */
export function chartXml(spec, refs, size) {
  // no axes: the plot area fills the frame (room for the legend), so its size in px is known
  const frame = spec.axesHidden && (spec.packed || spec.aspect) ? { x: 0.02, y: 0.03, w: spec.legend ? 0.74 : 0.96, h: 0.94 } : null;
  const plot = frame && { w: ((size && size.widthPx) || 600) * frame.w, h: ((size && size.heightPx) || 400) * frame.h };
  const axisPlot = size && size.widthPx ? { w: size.widthPx * 0.85, h: (size.heightPx || 300) * 0.72 } : null;
  const layout = frame
    ? `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${frame.x}"/><c:y val="${frame.y}"/><c:w val="${frame.w}"/><c:h val="${frame.h}"/></c:manualLayout></c:layout>`
    : "<c:layout/>";
  const legend = spec.legend
    ? `<c:legend><c:legendPos val="r"/><c:overlay val="0"/>${txPr(spec.font)}</c:legend>` : "";
  // a donut's centre text: the chart title laid over the hole, the ring centred in a fixed plot area
  const center = spec.kind === "doughnut" && spec.centerLabel && spec.centerLabel.length ? spec.centerLabel : null;
  let titleXml = "", plotLayout = layout;
  if (center) {
    const W = (size && size.widthPx) || 400, H = (size && size.heightPx) || 300;
    const box = { x: 0.04, y: 0.04, w: spec.legend ? 0.66 : 0.92, h: 0.92 };
    const pt = s => s.size || spec.font.size || 9;
    const tw = Math.max(...center.map(l => l.reduce((w, s) => w + String(s.text).length * pt(s) * 0.62, 0))) + 10;
    const th = center.reduce((h, l) => h + Math.max(...l.map(pt)) * 4 / 3 * 1.25, 0) + 6;
    const cx = box.x + box.w / 2, cy = box.y + box.h / 2;
    titleXml = `<c:title><c:tx><c:rich><a:bodyPr wrap="none"/><a:lstStyle/>` + center.map(l => `<a:p><a:pPr algn="ctr"><a:defRPr/></a:pPr>` +
        l.map(s => `<a:r>${runProps({ name: s.font || spec.font.name, size: pt(s), color: s.color || spec.font.color }, { bold: s.bold }, "rPr")}` +
          `<a:t>${esc(s.text)}</a:t></a:r>`).join("") + `</a:p>`).join("") + `</c:rich></c:tx>` +
      `<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${Math.max(0, cx - tw / 2 / W).toFixed(4)}"/><c:y val="${Math.max(0, cy - th / 2 / H).toFixed(4)}"/></c:manualLayout></c:layout>` +
      `<c:overlay val="1"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:title>`;
    plotLayout = `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${box.x}"/><c:y val="${box.y}"/><c:w val="${box.w}"/><c:h val="${box.h}"/></c:manualLayout></c:layout>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>` +
    `<c:chart>${titleXml}<c:autoTitleDeleted val="${titleXml ? 0 : 1}"/><c:plotArea>${plotLayout}${plotAreaXml(spec, refs, plot, axisPlot)}` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>${legend}` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr>${spec.background === null ? "<a:noFill/>" : solid(spec.background || "FFFFFF")}<a:ln><a:noFill/></a:ln></c:spPr>${txPr(spec.font)}` +
    `<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>` +
    `<c:pageSetup/></c:printSettings></c:chartSpace>`;
}
