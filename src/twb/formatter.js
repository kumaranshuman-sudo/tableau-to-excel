/* Resolves Tableau's formatting cascade for one worksheet (createSheetFormatter). */
import { TABLEAU_DEFAULTS, TF_DERIV_LABEL, TF_ELEMENTS } from "../config.js";
import { tfArgb } from "../format/colors.js";
import { tableauToExcelNumFmt } from "../format/number-format.js";
import { tfStrokeToBorder } from "./dashboard-text.js";
import { tfParseFieldRef, tfRefKey, tfSameField } from "./field-ref.js";
import { tfNorm, tfNum } from "../util.js";

export function tfFormatsToProps(m) {
  const p = {};
  if (m["font-family"]) p.fontName = m["font-family"];
  if (m["font-size"]) p.fontSize = tfNum(m["font-size"]);
  if (m["font-weight"]) p.bold = /bold/i.test(m["font-weight"]);
  if (m["font-style"]) p.italic = /italic/i.test(m["font-style"]);
  if (m["text-decoration"]) { p.underline = /underline/i.test(m["text-decoration"]); p.strike = /line-through/i.test(m["text-decoration"]); }
  if (m["color"]) p.color = tfArgb(m["color"]) || undefined;
  if ("background-color" in m) p.bgColor = tfArgb(m["background-color"]);  // null = transparent
  if (/^(left|center|right)$/.test(m["text-align"] || "")) p.hAlign = m["text-align"];
  const va = { top: "top", center: "middle", middle: "middle", bottom: "bottom" }[m["vertical-align"]];
  if (va) p.vAlign = va;
  if (m["wrap"]) p.wrap = /^(true|on)$/i.test(m["wrap"]);
  if (m["text-format"]) p.numFmtRaw = m["text-format"];
  if (m["band-color"]) p.bandColor = tfArgb(m["band-color"]) || undefined;
  if (m["band-size"]) p.bandSize = tfNum(m["band-size"]);
  if (m["line-visibility"]) p.lineVisible = m["line-visibility"] !== "off";
  if (m["stroke-size"] != null) p.strokeSize = tfNum(m["stroke-size"]);
  if (m["stroke-color"]) p.strokeColor = tfArgb(m["stroke-color"]) || undefined;
  if (m["div-level"]) p.divLevel = tfNum(m["div-level"]);
  if (m["width"]) p.width = tfNum(m["width"]);
  if (m["display"]) p.display = m["display"] !== "false";
  if (m["display-field-labels"]) p.displayFieldLabels = m["display-field-labels"] !== "false";
  if (m["height"]) p.height = tfNum(m["height"]);
  return p;
}

/* precedence inside one element: base < scope < field < field+scope */
export function tfCollect(style, elements, opts = {}) {
  const out = {};
  if (!style) return out;
  for (const el of elements) {
    const tiers = [{}, {}, {}, {}];
    for (const f of style.rules[el] || []) {
      if ((f.dataClass || null) !== (opts.dataClass || null)) continue;
      if (f.scope && f.scope !== opts.scope) continue;
      if (f.field && !(opts.field && tfSameField(tfParseFieldRef(f.field), opts.field))) continue;
      tiers[(f.field ? 2 : 0) + (f.scope ? 1 : 0)][f.attr] = f.value;
    }
    Object.assign(out, ...tiers);
  }
  return tfFormatsToProps(out);
}

export function tfMerge(...layers) {
  const out = {};
  layers.forEach(l => { if (l) for (const k in l) if (l[k] !== undefined) out[k] = l[k]; });
  return out;
}

export function tfFieldInfo(model, ref) {
  if (!model || !ref) return null;
  const k = ref.name.toLowerCase();
  return model.fields[(ref.ds || "") + "|" + k] || model.fields["|" + k] || null;
}

export function tfDisplayNames(model, ref) {
  const info = tfFieldInfo(model, ref);
  const cap = (info && info.caption) || ref.name;
  const agg = TF_DERIV_LABEL[(ref.deriv || "").toLowerCase()];
  const names = [];
  if (agg) names.push(`${agg}(${cap})`, `${agg}(${ref.name})`);
  names.push(cap, ref.name);
  if (ref.name === "Multiple Values") names.push("Measure Values");
  const alias = model && model.measureAliases && (model.measureAliases[tfRefKey(ref)] ||
    Object.entries(model.measureAliases).find(([k]) => k.endsWith("|" + ref.inner.toLowerCase()))?.[1]);
  if (alias) names.unshift(alias);
  return [...new Set(names.map(tfNorm))];
}

/** @param {FormatModel | null} model @param {string} sheetName */
export function createSheetFormatter(model, sheetName) {
  const sheet = (model && model.sheets && model.sheets[sheetName]) || null;
  const wb = model ? model.workbookStyle : null;
  const st = sheet ? sheet.style : null;
  const panes = sheet ? sheet.panes : [];
  const matchCache = new Map();

  function base() {
    return tfMerge(TABLEAU_DEFAULTS.worksheet, tfCollect(wb, TF_ELEMENTS.sheet), tfCollect(st, TF_ELEMENTS.sheet));
  }

  /* summary-data column name ("SUM(Sales)") → Tableau field ref. Exact names only,
   * no substring matching (old code matched "Sales" to "Sales Person"). */
  function matchName(colName) {
    const key = tfNorm(colName);
    if (matchCache.has(key)) return matchCache.get(key);
    let best = null, bestScore = 0;
    const stripped = key.replace(/^[a-z]+\((.*)\)$/, "$1");
    for (const ref of sheet ? sheet.fieldRefs : []) {
      const names = tfDisplayNames(model, ref);
      let score = 0;
      if (names[0] === key) score = 3;
      else if (names.includes(key)) score = 2;
      else if (names.includes(stripped)) score = 1;
      if (score > bestScore) { best = ref; bestScore = score; }
    }
    matchCache.set(key, best);
    return best;
  }

  function onShelf(ref, shelf) { return !!ref && (sheet ? sheet[shelf] : []).some(r => tfSameField(r, ref)); }
  // headers = DISCRETE pills on Rows/Columns; continuous pills (…:qk) draw axes, not headers
  function isHeaderField(ref) { return !!ref && ref.type !== "qk" && (onShelf(ref, "rows") || onShelf(ref, "cols")); }
  function shelfOf(ref) { return onShelf(ref, "rows") ? "rows" : onShelf(ref, "cols") ? "cols" : null; }

  function labelRunProps(ref) {
    if (!ref) return {};
    for (const p of panes) for (const r of p.labelRuns) if (r.refs.some(x => tfSameField(x, ref))) return r.props;
    return {};
  }
  /* panes that draw this field (multi-pane tables: each column has its own marks card) */
  function owningPanes(ref) {
    const owns = p => ref && (p.labelRuns.some(r => r.refs.some(x => tfSameField(x, ref))) ||
                              p.encodings.some(e => tfSameField(e.field, ref)));
    const hit = panes.filter(owns);
    if (hit.length) return hit;
    if (panes.length === 1) return panes;
    return panes.filter(p => !p.id);                            // the "All" pane
  }
  function paneStyle(ref) {
    const els = TF_ELEMENTS.pane.concat(TF_ELEMENTS.cell, TF_ELEMENTS.datalabel);
    return tfMerge(...owningPanes(ref).map(p => tfCollect(p.style, els, { field: ref })));
  }

  function numFmt(ref, fromHeader) {
    const els = fromHeader ? TF_ELEMENTS.header : TF_ELEMENTS.pane.concat(TF_ELEMENTS.cell);
    const raw = tfCollect(st, els, { field: ref, scope: fromHeader ? "rows" : undefined }).numFmtRaw
             || (ref && (tfFieldInfo(model, ref) || {}).defaultFormat);
    return raw ? tableauToExcelNumFmt(raw) : null;
  }

  return {
    hasModel: !!sheet,
    matchName,
    isHeaderField,

    titleText() { return sheet && sheet.title ? sheet.title.text : null; },
    titleStyle() {
      const runs = sheet && sheet.title ? sheet.title.runs : [];
      const firstRun = (runs.find(r => r.text.trim()) || {}).props;
      return tfMerge(base(), TABLEAU_DEFAULTS.title, tfCollect(wb, TF_ELEMENTS.title), tfCollect(st, TF_ELEMENTS.title), firstRun);
    },

    /* Excel header row (field names) */
    fieldLabelStyle(ref, isMeasure) {
      return tfMerge(base(), TABLEAU_DEFAULTS.fieldLabel,
        tfCollect(wb, TF_ELEMENTS.fieldLabel, { scope: "rows" }),
        tfCollect(st, TF_ELEMENTS.fieldLabel, { field: ref, scope: isMeasure ? "cols" : "rows" }));
    },

    /* dimension values (row headers in Tableau) */
    headerCellStyle(ref) {
      const p = tfMerge(base(), tfCollect(wb, TF_ELEMENTS.header, { scope: "rows" }),
        tfCollect(st, TF_ELEMENTS.header, { field: ref, scope: "rows" }));
      p.numFmt = numFmt(ref, true);
      return p;
    },

    /* measure values (text marks in Tableau) */
    markCellStyle(ref) {
      const els = TF_ELEMENTS.pane.concat(TF_ELEMENTS.cell, TF_ELEMENTS.datalabel);
      const p = tfMerge(base(), tfCollect(wb, els),
        tfCollect(st, els, { field: ref, scope: "rows" }),
        paneStyle(ref), labelRunProps(ref));
      // cell position comes from Format → Alignment (pane/cell); the label editor's alignment only
      // aligns lines inside the label, so it is used only when no pane/cell alignment is set
      const cellAlign = tfMerge(tfCollect(st, els, { field: ref, scope: "rows" }), paneStyle(ref)).hAlign;
      if (cellAlign) p.hAlign = cellAlign;
      p.numFmt = numFmt(ref, false);
      p.explicitColor = !!(labelRunProps(ref).color || tfCollect(st, TF_ELEMENTS.cell, { field: ref }).color);
      return p;
    },

    divider(scope) {
      const p = tfCollect(st, TF_ELEMENTS.divider, { scope });
      const d = scope === "rows" ? { ...TABLEAU_DEFAULTS.divider } : { ...TABLEAU_DEFAULTS.divider, visible: false };
      if (p.lineVisible !== undefined) d.visible = p.lineVisible;
      if (p.strokeSize !== undefined) { d.style = tfStrokeToBorder(p.strokeSize); if (!d.style) d.visible = false; }
      if (p.strokeColor) d.color = p.strokeColor;
      if (p.divLevel) d.level = p.divLevel;
      return d;
    },

    banding() {
      const pane = tfCollect(st, TF_ELEMENTS.pane, { scope: "rows" });
      const header = tfCollect(st, TF_ELEMENTS.header, { scope: "rows" });
      const table = tfCollect(st, ["table"], { scope: "rows" });
      return { pane: pane.bandColor || null, header: header.bandColor || null,
               size: table.bandSize || pane.bandSize || header.bandSize || 1 };
    },

    widthPx(ref) {
      return tfCollect(st, TF_ELEMENTS.cell.concat(TF_ELEMENTS.header), { field: ref, scope: "cols" }).width
          || tfCollect(st, TF_ELEMENTS.cell.concat(TF_ELEMENTS.header), { field: ref }).width || null;
    },
    rowHeightPx() { return tfCollect(st, TF_ELEMENTS.cell).height || null; },

    /* fields on Text / Label – these are the "marks" that get coloured */
    textRefs() {
      const refs = [];
      panes.forEach(p => {
        p.encodings.filter(e => e.channel === "text" || e.channel === "label").forEach(e => refs.push(e.field));
        p.labelRuns.forEach(r => r.refs.forEach(x => refs.push(x)));
      });
      return refs;
    },

    colorEncoding() {
      for (const p of panes) {
        const enc = p.encodings.find(e => e.channel === "color");
        if (!enc) continue;
        const ref = enc.field;
        const hasText = p.encodings.some(e => e.channel === "text" || e.channel === "label") || p.labelRuns.length > 0;
        const cls = String(p.markClass || "Automatic");
        const effective = /^automatic$/i.test(cls) && hasText ? "Text" : cls;
        const pool = [st, ...panes.map(x => x.style), model && model.datasourceStyles[ref.ds],
                      ...(model ? Object.values(model.datasourceStyles) : [])].filter(Boolean);
        let def = null;
        for (const s of pool) { def = s.encodings.find(e => tfSameField(e.field, ref)); if (def) break; }
        if (!def) for (const s of pool) { def = s.encodings.find(e => e.field && tfNorm(e.field.name) === tfNorm(ref.name)); if (def) break; }
        const paneRefs = [...p.encodings.filter(e => e.channel === "text" || e.channel === "label").map(e => e.field),
                          ...p.labelRuns.flatMap(r => r.refs)];
        return { ref, def, markClass: effective, applyTo: /^text$/i.test(effective) ? "font" : "fill", paneRefs,
                 continuous: ref.type ? /^q/.test(ref.type) : !!(def && def.type === "interpolated") };
      }
      return null;
    },

    shelfOf,
    sheetModel: sheet,

    /* Format → Shading → Worksheet: workbook < dashboard < worksheet */
    tableBackground(dashboardName) {
      const dash = dashboardName && model && model.dashboards ? model.dashboards[dashboardName] : null;
      const v = [tfCollect(wb, ["table"]).bgColor, dash ? tfCollect(dash.style, ["table"]).bgColor : undefined,
                 tfCollect(st, ["table"]).bgColor];
      let out = null;
      v.forEach(x => { if (x !== undefined) out = x; });       // later level wins; null = explicitly none
      return out;
    },
    /* column-header height (px) stored on any Columns-shelf header field, incl. Measure Names */
    headerRowHeightPx() {
      if (!sheet) return null;
      let h = null;
      (st && st.rules.header || []).forEach(f => {
        if (f.attr !== "height" || !f.field) return;
        const r = tfParseFieldRef(f.field);
        if (r.name === "Measure Names" || onShelf(r, "cols")) h = Math.max(h || 0, tfNum(f.value) || 0);
      });
      return h;
    },
    /* URL actions whose source includes this sheet */
    urlActions(dashboardName) {
      return (model && model.actions || []).filter(a =>
        (a.worksheet ? a.worksheet === sheetName : (!a.dashboard || a.dashboard === dashboardName)) &&
        !(a.exclude || []).includes(sheetName));
    },

    /* header label hidden via Format → "Hide" (label display=false) */
    isLabelHidden(ref) { return tfCollect(st, ["label"], { field: ref }).display === false; },

    /* an axis on a shelf: "Show Header" off (display=false) and Edit Axis → title ("" = no title).
     * cls "0" / "1" = primary / secondary axis of a dual axis; rules without a class apply to both */
    axisInfo(ref, shelf, cls) {
      /** @type {{ hidden?: boolean, title?: string }} */
      const out = {};
      if (!st || !ref) return out;
      /** @type {Record<string, string>[]} */
      const tiers = [{}, {}];                                 // without scope < with scope
      for (const f of st.rules.axis || []) {
        if (f.attr !== "display" && f.attr !== "title") continue;
        if (!f.field || !tfSameField(tfParseFieldRef(f.field), ref)) continue;
        if (f.scope && f.scope !== shelf) continue;
        if (f.axisClass !== undefined && cls !== undefined && f.axisClass !== cls) continue;
        tiers[f.scope ? 1 : 0][f.attr] = f.value;
      }
      const v = { ...tiers[0], ...tiers[1] };
      if (v.display !== undefined) out.hidden = v.display === "false";
      if (v.title !== undefined) out.title = v.title;
      return out;
    },
    /* grid lines across a shelf's axis (Format → Lines → Grid Lines): off = stroke 0 or line-visibility off */
    gridlinesShown(shelf) {
      const g = tfMerge(tfCollect(wb, ["gridline"], { scope: shelf }), tfCollect(st, ["gridline"], { scope: shelf }));
      return !(g.strokeSize === 0 || g.lineVisible === false);
    },
    /* Analytics → Reference Line entries of the worksheet */
    referenceLines() { return (sheet && sheet.referenceLines) || []; },
    /* Edit Axis for a field on a shelf: fixed range, tick spacing, include zero */
    axisSpace(ref, shelf, cls) {
      /** @type {AxisSpace} */
      let out = {};
      if (!st || !st.spaces || !ref) return out;
      for (const s of st.spaces) {
        if (!s.field || !tfSameField(tfParseFieldRef(s.field), ref) || (s.scope && s.scope !== shelf)) continue;
        if (s.axisClass !== undefined && cls !== undefined && s.axisClass !== cls) continue;
        out = s;
      }
      return out;
    },
    /* a header's label format (Format → Header → Dates), e.g. "iLLLLL" = month initial */
    labelFormat(ref) { return ref ? tfCollect(st, ["label"], { field: ref }).numFmtRaw || null : null; },
    /** mark label font (datalabel element); its colour only when the user picked one, not "automatic"
     * @returns {Record<string, any>} */
    dataLabelStyle() {
      const raw = {};
      const pool = [wb, st, ...owningPanes(null).map(p => p.style)];      // workbook, worksheet, marks card
      for (const s of pool) for (const f of (s && s.rules.datalabel) || []) if (!f.field) raw[f.attr] = f.value;
      const p = tfMerge(...pool.map(s => tfCollect(s, ["datalabel"])));
      if (raw["color-mode"] && raw["color-mode"] !== "user") delete p.color;
      return p;
    },
    /** the mark label's text around its one field ("<AGG(Days)> DAYS" → prefix "", suffix " DAYS"); null when
     * the label is the default, or holds several fields
     * @returns {{ ref: FieldRef, prefix: string, suffix: string } | null} */
    labelTemplate() {
      const pane = owningPanes(null)[0];
      if (!pane || !pane.labelRuns.length) return null;
      const text = pane.labelRuns.map(r => r.text).join("");
      const tokens = [...text.matchAll(/<([^<>]+)>/g)];
      const ref = tokens.length === 1 ? tfParseFieldRef(tokens[0][1]) : null;
      if (!ref) return null;
      const t = tokens[0], flat = s => s.replace(/\u00C6[ \t]*(?:\r?\n)?|\r?\n/g, " ");
      return { ref, prefix: flat(text.slice(0, t.index)), suffix: flat(text.slice(t.index + t[0].length)) };
    },
    /** reference line formatting: line colour (and its opacity), width, dash, visibility; label font / format
     * @returns {Record<string, any>} */
    reflineStyle() {
      const raw = {};
      const pool = [wb, st, ...owningPanes(null).map(p => p.style)];
      for (const s of pool) for (const f of (s && s.rules.refline) || []) if (!f.field && !f.scope) raw[f.attr] = f.value;
      const alpha = /^#?[0-9a-f]{8}$/i.test(raw["stroke-color"] || "") ? parseInt(raw["stroke-color"].replace(/^#/, "").slice(6), 16) / 255 : 1;
      return { ...tfMerge(...pool.map(s => tfCollect(s, ["refline"]))), dash: raw["line-pattern-only"] || null, strokeAlpha: alpha };
    },

    /* axis rulers (Format → Lines → Axis Rulers) */
    axisLineShown(shelf) {
      const a = tfMerge(tfCollect(wb, ["axis"], { scope: shelf }), tfCollect(st, ["axis"], { scope: shelf }));
      return !(a.strokeSize === 0 || a.lineVisible === false);
    },
    /* "Show field labels for rows/columns" (worksheet display-field-labels) */
    fieldLabelsShown(scope) {
      const v = tfCollect(st, ["worksheet"], { scope }).displayFieldLabels;
      return v === undefined ? true : v;
    },
    /* Rows then Columns, discrete pills only, in shelf order */
    headerRefs() {
      return sheet ? [...sheet.rows, ...sheet.cols].filter(r => r.type !== "qk" && r.name !== "Measure Names") : [];
    },
    /* measure order: Measure Names manual sort, else the Measure Names filter order */
    measureOrder() {
      if (!sheet) return [];
      const ms = sheet.manualSorts.find(m => m.field && m.field.name === "Measure Names");
      const src = ms ? ms.order : sheet.measureFilter;
      return src.map(x => tfParseFieldRef(String(x).replace(/^"|"$/g, ""))).filter(Boolean);
    },
    /* Measure Names text ("Typical home value") → measure field ref */
    measureRefForName(text) {
      const key = tfNorm(text);
      const cands = [...this.measureOrder(), ...(sheet ? sheet.fieldRefs : [])];
      return cands.find(r => tfDisplayNames(model, r).includes(key)) || matchName(text);
    },
    /* panes in visual order with the fields they label */
    panesInOrder() {
      return panes.map((p, i) => ({
        pane: p, i,
        order: p.xIndex !== undefined ? p.xIndex : (p.xAxisName ? 0 : (panes.length > 1 ? -1 : 0)),
        refs: [...p.labelRuns.flatMap(r => r.refs),
               ...p.encodings.filter(e => e.channel === "text" || e.channel === "label").map(e => e.field)]
      })).sort((a, b) => a.order - b.order || a.i - b.i);
    },
    /* chart vs table: a visible continuous axis, or only chart marks with no text */
    isChart() {
      if (!sheet) return false;
      for (const shelf of ["rows", "cols"]) {
        for (const r of sheet[shelf]) {
          const continuous = r.type === "qk" || r.name === "Multiple Values";
          if (!continuous) continue;
          const ax = tfCollect(st, ["axis"], { field: r, scope: shelf });
          if (ax.display !== false) return true;
        }
      }
      const chartMarks = /^(bar|line|area|pie|map|multipolygon|polygon|ganttbar|density)$/i;
      return panes.length > 0 && panes.every(p => chartMarks.test(p.markClass || "") &&
        !p.encodings.some(e => e.channel === "text" || e.channel === "label") && !p.labelRuns.length);
    },
    manualSortFor(ref) { return sheet ? sheet.manualSorts.find(m => tfSameField(m.field, ref)) : null; },
    measureSortFor(ref) { return sheet ? sheet.measureSorts.find(m => tfSameField(m.field, ref)) : null; },
    captionFor(ref, fallbackName) {
      const info = tfFieldInfo(model, ref);
      let cap = info && info.caption ? info.caption
              : ref && !/^Calculation_/.test(ref.name) ? ref.name
              : String(fallbackName || "").replace(/^[A-Z]+\((.*)\)$/, "$1");
      const d = ref && (ref.deriv || "").toLowerCase();
      const pre = { avg: "Avg. ", cnt: "Count of ", ctd: "Count Distinct of ", min: "Min. ", max: "Max. ",
                    med: "Median ", std: "Std. dev. of ", var: "Variance of " }[d];
      return pre ? pre + cap : cap;
    },
    /* title text with <Sheet Name>, <[field]> and <[Parameters].[p]> substituted */
    renderTitle(valuesFor, sheetLabel) {
      const runs = sheet && sheet.title ? sheet.title.runs : null;
      const props = this.titleStyle();
      if (!runs || !runs.length) return { text: sheetLabel, props };
      let text = runs.map(r => r.text).join("").replace(/\u00C6\r?\n?/g, "\n");
      text = text.replace(/<([^<>]+)>/g, (m, inner) => {
        if (/^sheet name$/i.test(inner)) return sheetLabel;
        if (/^(workbook name|page name|page count|page number)$/i.test(inner)) return "";
        if (!/^\[/.test(inner)) return m;
        const ref = tfParseFieldRef(inner);
        if (!ref) return m;
        if (ref.ds === "Parameters") {
          const info = tfFieldInfo(model, ref) || {};
          return info.alias || info.value || "";
        }
        const vals = valuesFor(ref);
        if (!vals.length) return "";
        return vals.length <= 3 ? vals.join(", ") : "*";
      });
      return { text: text.trim(), props };
    },

    /* display names for Measure-Names buckets like "[ds].[sum:Sales:qk]" */
    bucketAliases(bucketKey) {
      if (!/^\[.*\]$/.test(bucketKey)) return [bucketKey];
      const ref = tfParseFieldRef(bucketKey);
      return ref ? [bucketKey, ...tfDisplayNames(model, ref)] : [bucketKey];
    }
  };
}
