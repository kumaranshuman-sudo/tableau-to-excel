"use strict";

/* ══════════════════════════════════════════════════════════════════════════
 * EXCEL NATIVE CHART WRITER
 * ──────────────────────────────────────────────────────────────────────────
 * ExcelJS cannot author charts, so native charts are produced in two steps:
 *   1. writeChartData(ws, spec, row) – the chart's source data is written to a
 *      (hidden) data sheet with ExcelJS, so every chart stays editable and is
 *      linked to real cells.
 *   2. injectCharts(buffer, opts) – after workbook.xlsx.writeBuffer() the XLSX
 *      package is opened with JSZip and DrawingML chart parts are added to the
 *      dashboard sheet (re-using the drawing ExcelJS made for images, if any).
 *
 * Chart spec (produced by visual_chart_model.js):
 *   { kind: "bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"combo",
 *     barDir: "col"|"bar", stacked, categories: { names:[], levels:[[…]] },
 *     series: [{ name, type, color, values|x+y, pointColors, secondary,
 *                line, marker, labels }],
 *     numFmt, secondaryNumFmt, xNumFmt, valueTitle, secondaryTitle,
 *     categoryTitle, legend, gridlines, font: { name, size, color } }
 * Colours are 6-digit hex ("4E79A7"); ARGB ("FF4E79A7") is accepted too.
 * ══════════════════════════════════════════════════════════════════════════ */
const ExcelChartWriter = (function () {

  const NS = {
    c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    rels: "http://schemas.openxmlformats.org/package/2006/relationships"
  };
  const REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
  const REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
  const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";
  const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
  const EMU_PER_PX = 9525;
  const AX = { cat: 50010, val: 50020, cat2: 50030, val2: 50040 };

  /* ── small helpers ─────────────────────────────────────────────────────── */
  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function colName(i) {
    let s = "", n = i + 1;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
    return s;
  }
  /* 0-based (col,row) → 'Sheet'!$A$1[:$B$9] */
  function cellRef(sheet, c1, r1, c2, r2) {
    const q = "'" + String(sheet).replace(/'/g, "''") + "'!";
    const a = "$" + colName(c1) + "$" + (r1 + 1);
    return q + a + (c2 === undefined ? "" : ":$" + colName(c2) + "$" + (r2 + 1));
  }
  function hex(c) {
    const s = String(c || "").replace(/^#/, "").toUpperCase();
    return s.length === 8 ? s.slice(2) : s;
  }
  function num(v) { return typeof v === "number" && isFinite(v) ? v : null; }
  function attr(tag, name) {
    const m = tag.match(new RegExp("\\s" + name.replace(":", "\\:") + "=\"([^\"]*)\""));
    return m ? m[1] : null;
  }

  /* boundaries of each category level: a new group starts where any level ≤ l changes */
  function levelStarts(levels) {
    const n = levels.length ? levels[0].length : 0;
    return levels.map((_, l) => Array.from({ length: n }, (__, i) =>
      i === 0 || levels.slice(0, l + 1).some(lv => lv[i] !== lv[i - 1])));
  }

  /* ══════════════════════════════════════════════════════════════════════
   * 1. CHART DATA → worksheet cells
   * ══════════════════════════════════════════════════════════════════════ */
  function writeChartData(ws, spec, startRow) {
    const sheet = ws.name;
    const put = (r, c, v, fmt) => {
      const cell = ws.getCell(r + 1, c + 1);
      cell.value = v;
      if (fmt) cell.numFmt = fmt;
    };
    ws.getCell(startRow + 1, 1).font = { bold: true };

    if (spec.kind === "scatter") {
      let maxLen = 0;
      const series = spec.series.map((s, k) => {
        const cx = 2 * k, cy = 2 * k + 1, n = s.x.length;
        put(startRow, cx, (spec.xTitle || "X") + (spec.series.length > 1 ? " – " + s.name : ""));
        put(startRow, cy, s.name);
        for (let i = 0; i < n; i++) {
          if (num(s.x[i]) !== null) put(startRow + 1 + i, cx, s.x[i], spec.xNumFmt);
          if (num(s.y[i]) !== null) put(startRow + 1 + i, cy, s.y[i], spec.numFmt);
        }
        maxLen = Math.max(maxLen, n);
        return { tx: cellRef(sheet, cy, startRow),
                 x: cellRef(sheet, cx, startRow + 1, cx, startRow + Math.max(1, n)),
                 y: cellRef(sheet, cy, startRow + 1, cy, startRow + Math.max(1, n)) };
      });
      return { series, nextRow: startRow + maxLen + 3 };
    }

    const levels = spec.categories.levels;
    const L = levels.length;
    const N = L ? levels[0].length : 0;
    const starts = levelStarts(levels);
    spec.categories.names.forEach((name, l) => put(startRow, l, name || ""));
    spec.series.forEach((s, k) => put(startRow, L + k, s.name));
    for (let i = 0; i < N; i++) {
      for (let l = 0; l < L; l++) if (l === L - 1 || starts[l][i]) put(startRow + 1 + i, l, levels[l][i]);
      spec.series.forEach((s, k) => {
        const v = num(s.values[i]);
        if (v !== null) put(startRow + 1 + i, L + k, v, s.secondary ? spec.secondaryNumFmt : spec.numFmt);
      });
    }
    return {
      cat: cellRef(sheet, 0, startRow + 1, L - 1, startRow + N),
      series: spec.series.map((s, k) => ({
        tx: cellRef(sheet, L + k, startRow),
        val: cellRef(sheet, L + k, startRow + 1, L + k, startRow + N)
      })),
      nextRow: startRow + N + 3
    };
  }

  /* ══════════════════════════════════════════════════════════════════════
   * 2. DrawingML chart XML
   * ══════════════════════════════════════════════════════════════════════ */
  function solid(color, alpha) {
    if (!color) return "<a:noFill/>";
    return `<a:solidFill><a:srgbClr val="${hex(color)}">${alpha ? `<a:alpha val="${alpha}"/>` : ""}</a:srgbClr></a:solidFill>`;
  }
  function line(color, w) { return color ? `<a:ln w="${w || 9525}">${solid(color)}</a:ln>` : `<a:ln><a:noFill/></a:ln>`; }
  function runProps(font, o, tag) {
    const sz = Math.round((o.size || font.size || 9) * 100);
    return `<a:${tag} lang="en-US" sz="${sz}" b="${o.bold ? 1 : 0}">${solid(o.color || font.color)}` +
      `<a:latin typeface="${esc(font.name)}"/><a:cs typeface="${esc(font.name)}"/></a:${tag}>`;
  }
  function txPr(font, o = {}) {
    const def = runProps(font, o, "defRPr").replace(' lang="en-US"', "");
    return `<c:txPr><a:bodyPr${o.rot ? ` rot="${o.rot}" vert="horz"` : ""}/><a:lstStyle/>` +
      `<a:p><a:pPr>${def}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
  }
  function title(text, font, vertical) {
    if (!text) return "";
    const def = runProps(font, {}, "defRPr").replace(' lang="en-US"', "");
    return `<c:title><c:tx><c:rich><a:bodyPr${vertical ? ' rot="-5400000" vert="horz"' : ""}/><a:lstStyle/>` +
      `<a:p><a:pPr>${def}</a:pPr><a:r>${runProps(font, {}, "rPr")}<a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx>` +
      `<c:overlay val="0"/></c:title>`;
  }

  function strCache(values) {
    return `<c:ptCount val="${values.length}"/>` +
      values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("");
  }
  function numCache(values) {
    return `<c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>` +
      values.map((v, i) => num(v) === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  }
  function serTx(ref, name) {
    return `<c:tx><c:strRef><c:f>${esc(ref)}</c:f><c:strCache>${strCache([name])}</c:strCache></c:strRef></c:tx>`;
  }
  function catXml(ref, levels) {
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
  function valXml(tag, ref, values) {
    return `<c:${tag}><c:numRef><c:f>${esc(ref)}</c:f><c:numCache>${numCache(values)}</c:numCache></c:numRef></c:${tag}>`;
  }

  function dLbls(spec, s, pos) {
    if (!s.labels) return "";
    const fmt = s.secondary ? spec.secondaryNumFmt : spec.numFmt;
    const showVal = s.labelParts ? (s.labelParts.value ? 1 : 0) : 1;
    const showCat = s.labelParts && s.labelParts.category ? 1 : 0;
    return `<c:dLbls>${fmt ? `<c:numFmt formatCode="${esc(fmt)}" sourceLinked="0"/>` : ""}` +
      `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(spec.font)}` +
      `${pos ? `<c:dLblPos val="${pos}"/>` : ""}<c:showLegendKey val="0"/><c:showVal val="${showVal}"/>` +
      `<c:showCatName val="${showCat}"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>` +
      `${spec.kind === "pie" || spec.kind === "doughnut" ? '<c:showLeaderLines val="1"/>' : ""}</c:dLbls>`;
  }
  function markerXml(symbol, color, size) {
    if (!symbol || symbol === "none") return `<c:marker><c:symbol val="none"/></c:marker>`;
    return `<c:marker><c:symbol val="${symbol}"/><c:size val="${size || 6}"/>` +
      `<c:spPr>${solid(color, 85000)}${line(color, 9525)}</c:spPr></c:marker>`;
  }

  function seriesXml(spec, s, k, refs, type) {
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
      const symbol = s.marker ? "circle" : "none";
      const dpts = s.marker ? pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/>${markerXml("circle", c, 7)}<c:bubble3D val="0"/></c:dPt>` : "").join("") : "";
      return `<c:ser>${head}<c:spPr>${lineSp}</c:spPr>${markerXml(symbol, s.color, 7)}${dpts}` +
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
      `${dLbls(spec, s, "t")}${valXml("xVal", r.x, s.x)}${valXml("yVal", r.y, s.y)}<c:smooth val="0"/></c:ser>`;
  }

  function catAxis(spec, id, cross, o = {}) {
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
  function zeroScaling(values) {
    const nums = values.filter(v => num(v) !== null);
    if (!nums.length) return "";
    if (Math.min(...nums) >= 0) return `<c:min val="0"/>`;
    if (Math.max(...nums) <= 0) return `<c:max val="0"/>`;
    return "";
  }
  function valAxis(spec, id, cross, o = {}) {
    const horizontal = spec.barDir === "bar" && spec.kind !== "scatter";
    const pos = o.pos || (horizontal ? "b" : "l");
    const crosses = o.crosses || (horizontal ? "max" : "autoZero");
    const grid = o.grid !== false && spec.gridlines !== false
      ? `<c:majorGridlines><c:spPr>${line("EBEBEB", 9525)}</c:spPr></c:majorGridlines>` : "";
    const scale = spec.includeZero === false || !o.values ? "" : zeroScaling(o.values);
    return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${scale}</c:scaling><c:delete val="0"/>` +
      `<c:axPos val="${pos}"/>${grid}${title(o.title, spec.font, pos === "l" || pos === "r")}` +
      `<c:numFmt formatCode="${esc(o.numFmt || "General")}" sourceLinked="0"/><c:majorTickMark val="none"/>` +
      `<c:minorTickMark val="none"/><c:tickLblPos val="${o.lowLabels ? "low" : "nextTo"}"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>` +
      `${txPr(spec.font)}<c:crossAx val="${cross}"/><c:crosses val="${crosses}"/>` +
      `<c:crossBetween val="${o.midCat ? "midCat" : "between"}"/></c:valAx>`;
  }

  function plotAreaXml(spec, refs) {
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
          `<c:gapWidth val="${spec.gapWidth || 60}"/>${spec.stacked ? '<c:overlap val="100"/>' : ""}${ax}</c:barChart>`;
      }
      if (g.type === "area") {
        return `<c:areaChart><c:grouping val="${spec.stacked ? "stacked" : "standard"}"/><c:varyColors val="0"/>${sers}${ax}</c:areaChart>`;
      }
      return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${ax}</c:lineChart>`;
    }).join("");
    const axisValues = secondary => spec.series.filter(s => !!s.secondary === secondary).flatMap(s => s.values);
    let axes = catAxis(spec, AX.cat, AX.val) +
      valAxis(spec, AX.val, AX.cat, { title: spec.valueTitle, numFmt: spec.numFmt, values: axisValues(false) });
    if (hasSecondary) {
      axes += catAxis(spec, AX.cat2, AX.val2, { deleted: true }) +
        valAxis(spec, AX.val2, AX.cat2, { pos: spec.barDir === "bar" ? "t" : "r", crosses: "max", grid: false,
          title: spec.secondaryTitle, numFmt: spec.secondaryNumFmt || spec.numFmt, values: axisValues(true) });
    }
    return xml + axes;
  }

  function chartXml(spec, refs) {
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

  function anchorXml(chart, id, rid) {
    const cx = Math.max(1, Math.round(chart.widthPx * EMU_PER_PX));
    const cy = Math.max(1, Math.round(chart.heightPx * EMU_PER_PX));
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.col}</xdr:col><xdr:colOff>${Math.round((chart.colOffPx || 0) * EMU_PER_PX)}</xdr:colOff>` +
      `<xdr:row>${chart.row}</xdr:row><xdr:rowOff>${Math.round((chart.rowOffPx || 0) * EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
      `<xdr:ext cx="${cx}" cy="${cy}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>` +
      `<xdr:cNvPr id="${id}" name="${esc(chart.name || "Chart " + id)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>` +
      `<a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${rid}"/></a:graphicData>` +
      `</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
  }

  /* ══════════════════════════════════════════════════════════════════════
   * 3. XLSX package surgery
   * ══════════════════════════════════════════════════════════════════════ */
  const EMPTY_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.rels}"></Relationships>`;

  function relTarget(relsXml, id) {
    const tag = (relsXml.match(new RegExp(`<Relationship\\b[^>]*\\sId="${id}"[^>]*>`)) || [])[0];
    return tag ? attr(tag, "Target") : null;
  }
  function nextRelId(relsXml) {
    const ids = [...relsXml.matchAll(/\sId="rId(\d+)"/g)].map(m => +m[1]);
    return "rId" + ((ids.length ? Math.max(...ids) : 0) + 1);
  }
  function addRel(relsXml, id, type, target) {
    return relsXml.replace("</Relationships>", `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`);
  }
  function resolve(baseDir, target) {
    if (target.startsWith("/")) return target.slice(1);
    const parts = (baseDir + target).split("/");
    const out = [];
    parts.forEach(p => { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); });
    return out.join("/");
  }
  function dirOf(path) { return path.replace(/[^/]+$/, ""); }
  function relsPathOf(path) { return dirOf(path) + "_rels/" + path.replace(/^.*\//, "") + ".rels"; }
  function nextFreeIndex(zip, prefix, ext) {
    let n = 1;
    while (zip.file(`${prefix}${n}${ext}`)) n++;
    return n;
  }
  function addOverride(ct, part, type) {
    if (ct.includes(`PartName="/${part}"`)) return ct;
    return ct.replace("</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
  }
  /* <drawing> must precede these elements in CT_Worksheet */
  function insertDrawingTag(sheetXml, rid) {
    const tag = `<drawing r:id="${rid}"/>`;
    const after = ["<legacyDrawing", "<legacyDrawingHF", "<drawingHF", "<picture", "<oleObjects", "<controls",
                   "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"];
    for (const t of after) {
      const i = sheetXml.indexOf(t);
      if (i >= 0) return sheetXml.slice(0, i) + tag + sheetXml.slice(i);
    }
    return sheetXml;
  }

  /* opts: { sheetIndex, charts: [{ spec, refs, col, row, widthPx, heightPx, name }] } */
  async function injectCharts(buffer, opts) {
    if (typeof JSZip === "undefined") throw new Error("JSZip is required to add native Excel charts");
    const charts = (opts.charts || []).filter(Boolean);
    if (!charts.length) return buffer;
    const zip = await JSZip.loadAsync(buffer);
    const read = p => zip.file(p) ? zip.file(p).async("string") : Promise.resolve(null);

    const wbXml = await read("xl/workbook.xml");
    const wbRels = await read("xl/_rels/workbook.xml.rels");
    const sheetTags = wbXml.match(/<sheet\b[^>]*>/g) || [];
    const sheetTag = sheetTags[opts.sheetIndex || 0];
    if (!sheetTag) throw new Error("Dashboard sheet not found in workbook package");
    const sheetPath = resolve("xl/", relTarget(wbRels, attr(sheetTag, "r:id")));
    const sheetRelsPath = relsPathOf(sheetPath);
    let sheetXml = await read(sheetPath);
    let sheetRels = (await read(sheetRelsPath)) || EMPTY_RELS;
    let ct = await read("[Content_Types].xml");

    if (!/<worksheet\b[^>]*\sxmlns:r=/.test(sheetXml)) {
      sheetXml = sheetXml.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS.r}"`);
    }

    // re-use the drawing ExcelJS created for images, else create one
    let drawingPath, drawingXml, drawingRels;
    const existingRid = (sheetXml.match(/<drawing\b[^>]*r:id="([^"]+)"/) || [])[1];
    if (existingRid) {
      drawingPath = resolve(dirOf(sheetPath), relTarget(sheetRels, existingRid));
      drawingXml = await read(drawingPath);
      drawingRels = (await read(relsPathOf(drawingPath))) || EMPTY_RELS;
    } else {
      const n = nextFreeIndex(zip, "xl/drawings/drawing", ".xml");
      drawingPath = `xl/drawings/drawing${n}.xml`;
      drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"></xdr:wsDr>`;
      drawingRels = EMPTY_RELS;
      const rid = nextRelId(sheetRels);
      sheetRels = addRel(sheetRels, rid, REL_DRAWING, `../drawings/drawing${n}.xml`);
      sheetXml = insertDrawingTag(sheetXml, rid);
      ct = addOverride(ct, drawingPath, CT_DRAWING);
    }
    if (!/<xdr:wsDr\b[^>]*\sxmlns:a=/.test(drawingXml)) {
      drawingXml = drawingXml.replace(/<xdr:wsDr\b/, `<xdr:wsDr xmlns:a="${NS.a}"`);
    }

    const ids = [...drawingXml.matchAll(/<xdr:cNvPr\b[^>]*\sid="(\d+)"/g)].map(m => +m[1]);
    let shapeId = (ids.length ? Math.max(...ids) : 1) + 1;
    let anchors = "";
    for (const chart of charts) {
      const n = nextFreeIndex(zip, "xl/charts/chart", ".xml");
      const chartPath = `xl/charts/chart${n}.xml`;
      zip.file(chartPath, chartXml(chart.spec, chart.refs));
      ct = addOverride(ct, chartPath, CT_CHART);
      const rid = nextRelId(drawingRels);
      drawingRels = addRel(drawingRels, rid, REL_CHART, `../charts/chart${n}.xml`);
      anchors += anchorXml(chart, shapeId++, rid);
    }
    drawingXml = drawingXml.replace("</xdr:wsDr>", anchors + "</xdr:wsDr>");

    zip.file(sheetPath, sheetXml);
    zip.file(sheetRelsPath, sheetRels);
    zip.file(drawingPath, drawingXml);
    zip.file(relsPathOf(drawingPath), drawingRels);
    zip.file("[Content_Types].xml", ct);
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  return { writeChartData, injectCharts, chartXml };
})();

if (typeof module !== "undefined" && module.exports) module.exports = { ExcelChartWriter };
