/* DrawingML chart XML (bar / line / area / pie / scatter / combo). */
import { AX, C15, NS } from "./constants.js";
import { esc, hex, levelStarts, num } from "./xml-util.js";

/* ══════════════════════════════════════════════════════════════════════
 * 2. DrawingML chart XML
 * ══════════════════════════════════════════════════════════════════════ */
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
  return `<c:txPr><a:bodyPr${o.rot ? ` rot="${o.rot}" vert="horz"` : ""}/><a:lstStyle/>` +
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

export function catXml(ref, levels) {
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

/** @param {ChartSpec} spec @param {ChartSeries} s @param {string | null} pos */
export function dLbls(spec, s, pos) {
  if (!s.labels) return "";
  if (s.labelTexts) {
    return `<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(spec.font)}` +
      `<c:dLblPos val="${pos || "r"}"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>` +
      `<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>` +
      `<c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="${C15}">` +
      `<c15:showDataLabelsRange val="1"/></c:ext></c:extLst></c:dLbls>`;
  }
  const fmt = s.labelNumFmt || (s.secondary ? spec.secondaryNumFmt : spec.numFmt);
  const showVal = s.labelParts ? (s.labelParts.value ? 1 : 0) : 1;
  const showCat = s.labelParts && s.labelParts.category ? 1 : 0;
  return `<c:dLbls>${fmt ? `<c:numFmt formatCode="${esc(fmt)}" sourceLinked="0"/>` : ""}` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(spec.font)}` +
    `${pos ? `<c:dLblPos val="${pos}"/>` : ""}<c:showLegendKey val="0"/><c:showVal val="${showVal}"/>` +
    `<c:showCatName val="${showCat}"/><c:showSerName val="0"/><c:showPercent val="${s.labelParts && s.labelParts.percent ? 1 : 0}"/><c:showBubbleSize val="0"/>` +
    `${spec.kind === "pie" || spec.kind === "doughnut" ? '<c:showLeaderLines val="1"/>' : ""}</c:dLbls>`;
}

export function markerXml(symbol, color, size) {
  if (!symbol || symbol === "none") return `<c:marker><c:symbol val="none"/></c:marker>`;
  return `<c:marker><c:symbol val="${symbol}"/><c:size val="${size || 6}"/>` +
    `<c:spPr>${solid(color, 85000)}${line(color, 9525)}</c:spPr></c:marker>`;
}

/** @param {ChartSpec} spec @param {ChartSeries} s @param {number} k @param {ChartRefs} refs @param {string} type */
export function seriesXml(spec, s, k, refs, type) {
  const r = refs.series[k];
  const head = `<c:idx val="${k}"/><c:order val="${k}"/>${serTx(r.tx, s.name)}`;
  const pc = s.pointColors || [];
  const levels = spec.categories ? spec.categories.levels : [];
  if (type === "bar") {
    const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>` +
      `<c:spPr>${solid(c)}<a:ln><a:noFill/></a:ln></c:spPr></c:dPt>` : "").join("");
    const pos = spec.stacked ? null : "outEnd";
    return `<c:ser>${head}<c:spPr>${solid(s.color)}<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>` +
      `${dpts}${dLbls(spec, s, pos)}${catXml(refs.cat, levels)}${valXml("val", r.val, s.values)}</c:ser>`;
  }
  if (type === "line") {
    const lineSp = s.line === false ? `<a:ln w="28575"><a:noFill/></a:ln>`
      : `<a:ln w="22225" cap="rnd">${solid(s.color)}<a:round/></a:ln>`;
    const symbol = s.marker ? (s.markerSymbol || "circle") : "none";
    const size = s.markerSize || 7;
    const dpts = s.marker ? pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/>${markerXml(symbol, c, size)}<c:bubble3D val="0"/></c:dPt>` : "").join("") : "";
    return `<c:ser>${head}<c:spPr>${lineSp}</c:spPr>${markerXml(symbol, s.color, size)}${dpts}` +
      `${dLbls(spec, s, "t")}${catXml(refs.cat, levels)}${valXml("val", r.val, s.values)}<c:smooth val="0"/></c:ser>`;
  }
  if (type === "area") {
    return `<c:ser>${head}<c:spPr>${solid(s.color, spec.stacked ? null : 75000)}<a:ln><a:noFill/></a:ln></c:spPr>` +
      `${dLbls(spec, s, null)}${catXml(refs.cat, levels)}${valXml("val", r.val, s.values)}</c:ser>`;
  }
  if (type === "pie") {
    const dpts = pc.map((c, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/>` +
      `<c:spPr>${solid(c || s.color)}${line("FFFFFF", 12700)}</c:spPr></c:dPt>`).join("");
    return `<c:ser>${head}${dpts}${dLbls(spec, s, spec.kind === "pie" ? "bestFit" : null)}` +
      `${catXml(refs.cat, levels)}${valXml("val", r.val, s.values)}</c:ser>`;
  }
  // scatter
  const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/>${markerXml("circle", c, 7)}<c:bubble3D val="0"/></c:dPt>` : "").join("");
  return `<c:ser>${head}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>${markerXml("circle", s.color, 7)}${dpts}` +
    `${dLbls(spec, s, "r")}${valXml("xVal", r.x, s.x)}${valXml("yVal", r.y, s.y)}<c:smooth val="0"/>` +
    (s.labelTexts && r.lbl
      ? `<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}" xmlns:c15="${C15}"><c15:datalabelsRange>` +
        `<c15:f>${esc(r.lbl)}</c15:f><c15:dlblRangeCache>${strCache(s.labelTexts)}</c15:dlblRangeCache>` +
        `</c15:datalabelsRange></c:ext></c:extLst>` : "") + `</c:ser>`;
}

/** @param {ChartSpec} spec @param {number} id @param {number} cross @param {{ deleted?: boolean }} [o] */
export function catAxis(spec, id, cross, o = {}) {
  const horizontal = spec.barDir === "bar";
  const multi = spec.categories && spec.categories.levels.length > 1;
  return `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="${horizontal ? "maxMin" : "minMax"}"/></c:scaling>` +
    `<c:delete val="${o.deleted ? 1 : 0}"/><c:axPos val="${horizontal ? "l" : "b"}"/>` +
    `${o.deleted ? "" : title(spec.categoryTitle, spec.font, horizontal)}` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>` +
    `<c:tickLblPos val="${o.deleted ? "none" : "low"}"/><c:spPr>${line("D4D4D4", 9525)}</c:spPr>${txPr(spec.font)}` +
    `<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>` +
    `<c:lblOffset val="100"/><c:noMultiLvlLbl val="${multi ? 0 : 1}"/></c:catAx>`;
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
 *           numFmt?: string, values?: (number | null)[], deleted?: boolean, lowLabels?: boolean, midCat?: boolean }} [o]
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
    `<c:numFmt formatCode="${esc(axisFmt(o.numFmt || "General", o.values))}" sourceLinked="0"/><c:majorTickMark val="none"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="${o.lowLabels ? "low" : "nextTo"}"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>` +
    `${txPr(spec.font)}<c:crossAx val="${cross}"/><c:crosses val="${crosses}"/>` +
    `<c:crossBetween val="${o.midCat ? "midCat" : "between"}"/></c:valAx>`;
}

/** @param {ChartSpec} spec @param {ChartRefs} refs */
export function plotAreaXml(spec, refs) {
  const k = spec.kind;
  if (k === "pie" || k === "doughnut") {
    const ser = seriesXml(spec, spec.series[0], 0, refs, "pie");
    return k === "pie"
      ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
      : `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
  }
  if (k === "scatter") {
    const sers = spec.series.map((s, i) => seriesXml(spec, s, i, refs, "scatter")).join("");
    return `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}` +
      `<c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/></c:scatterChart>` +
      valAxis(spec, AX.cat, AX.val, { pos: "b", crosses: "autoZero", title: spec.xTitle, numFmt: spec.xNumFmt, midCat: true,
        grid: true, lowLabels: true, values: spec.series.flatMap(s => s.x) }) +
      valAxis(spec, AX.val, AX.cat, { pos: "l", crosses: "autoZero", title: spec.valueTitle, numFmt: spec.numFmt, midCat: true,
        lowLabels: true, values: spec.series.flatMap(s => s.y) });
  }
  // bar / line / area / combo: one chart group per (type, axis)
  const groups = [];
  spec.series.forEach((s, i) => {
    const type = s.type || k;
    const key = type + (s.secondary ? "2" : "1");
    let g = groups.find(x => x.key === key);
    if (!g) groups.push(g = { key, type, secondary: !!s.secondary, items: [] });
    g.items.push(seriesXml(spec, s, i, refs, type));
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
  let axes = catAxis(spec, AX.cat, AX.val) +
    valAxis(spec, AX.val, AX.cat, { title: spec.valueTitle, numFmt: spec.numFmt, values: axisValues(false),
                                    fixed: { min: spec.valueMin, max: spec.valueMax }, deleted: spec.valueAxisHidden });
  if (hasSecondary) {
    axes += catAxis(spec, AX.cat2, AX.val2, { deleted: true }) +
      valAxis(spec, AX.val2, AX.cat2, { pos: spec.barDir === "bar" ? "t" : "r", crosses: "max", grid: false,
        title: spec.secondaryTitle, numFmt: spec.secondaryNumFmt || spec.numFmt, values: axisValues(true) });
  }
  return xml + axes;
}

/** chart part XML (xl/charts/chartN.xml)
 * @param {ChartSpec} spec @param {ChartRefs} refs @returns {string} */
export function chartXml(spec, refs) {
  const legend = spec.legend
    ? `<c:legend><c:legendPos val="r"/><c:overlay val="0"/>${txPr(spec.font)}</c:legend>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>` +
    `<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plotAreaXml(spec, refs)}` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>${legend}` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr>${solid(spec.background || "FFFFFF")}<a:ln><a:noFill/></a:ln></c:spPr>${txPr(spec.font)}` +
    `<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>` +
    `<c:pageSetup/></c:printSettings></c:chartSpace>`;
}
