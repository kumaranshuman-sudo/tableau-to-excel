/* TWB/TWBX XML → plain JSON format model (worksheets, panes, styles, dashboards, actions). */
import { FORMAT_CONFIG } from "../config.js";
import { tfArgb } from "../format/colors.js";
import { tfExtractRefs, tfParseFieldRef, tfRefKey } from "./field-ref.js";
import { tfMerge } from "./formatter.js";
import { tfDefined, tfKid, tfKids, tfNum } from "../util.js";

export function tfParseRunProps(run) {
  const a = n => run.getAttribute(n);
  const align = { "0": "left", "1": "center", "2": "right" }[a("fontalignment")];
  return tfDefined({
    fontName: a("fontname") || undefined,
    fontSize: tfNum(a("fontsize")),
    bold: a("bold") != null ? a("bold") === "true" : undefined,
    italic: a("italic") != null ? a("italic") === "true" : undefined,
    underline: a("underline") != null ? a("underline") === "true" : undefined,
    color: tfArgb(a("fontcolor")) || undefined,
    hAlign: align
  });
}

/* a dashboard zone's Layout pane formatting: background and border */
export function tfZoneStyle(el) {
  if (!el) return undefined;
  const f = {};
  tfKids(el, "format").forEach(x => { f[x.getAttribute("attr")] = x.getAttribute("value"); });
  const style = tfDefined({
    bgColor: tfArgb(f["background-color"]) || undefined,
    borderColor: tfArgb(f["border-color"]) || undefined,
    borderStyle: f["border-style"] || undefined,
    borderWidth: tfNum(f["border-width"])
  });
  return Object.keys(style).length ? style : undefined;
}

export function tfParseRuns(ftEl) {
  return tfKids(ftEl, "run").map(r => ({ text: r.textContent || "", props: tfParseRunProps(r) }));
}

export function tfBucketKey(t) {
  let s = String(t == null ? "" : t).trim();
  if (/^".*"$/.test(s)) s = s.slice(1, -1).replace(/""/g, '"');
  else if (/^#.*#$/.test(s)) s = s.slice(1, -1);
  if (s === "%null%") s = "null";
  return s.toLowerCase();
}

export function tfParseColorEncoding(enc) {
  const a = n => enc.getAttribute(n);
  const cp = enc.getElementsByTagName("color-palette")[0] || null;
  const customColors = cp
    ? Array.from(cp.getElementsByTagName("color")).map(c => tfArgb(c.textContent)).filter(Boolean)
    : [];
  const map = {};
  Array.from(enc.getElementsByTagName("map")).forEach(m => {
    const color = tfArgb(m.getAttribute("to"));
    if (!color) return;
    Array.from(m.getElementsByTagName("bucket")).forEach(b => { map[tfBucketKey(b.textContent)] = color; });
  });
  return {
    field: tfParseFieldRef(a("field")),
    type: a("type"),                                  // "palette" (categorical) | "interpolated" | …
    paletteName: a("palette") || (cp && cp.getAttribute("name")) || null,
    paletteType: cp ? cp.getAttribute("type") : null, // ordered-sequential | ordered-diverging | regular
    customColors,
    reverse: a("reverse") === "true" || (cp && cp.getAttribute("reverse") === "true"),
    // optional range settings (only used if present in your TWB)
    center: tfNum(a("center")), min: tfNum(a("min")), max: tfNum(a("max")),
    map
  };
}

/** @param {Element | null} styleEl @returns {ParsedStyle} */
export function tfParseStyle(styleEl) {
  /** @type {ParsedStyle["rules"]} */
  const rules = {};
  const encodings = [];
  if (!styleEl) return { rules, encodings };
  tfKids(styleEl, "style-rule").forEach(rule => {
    const el = rule.getAttribute("element") || "all";
    rules[el] = rules[el] || [];
    tfKids(rule, "format").forEach(f => rules[el].push(tfDefined({
      attr: f.getAttribute("attr"),
      value: f.getAttribute("value"),
      field: f.getAttribute("field") || undefined,
      scope: f.getAttribute("scope") || undefined,
      dataClass: f.getAttribute("data-class") || undefined,
      axisClass: f.getAttribute("class") || undefined
    })));
  });
  Array.from(styleEl.getElementsByTagName("encoding"))
    .filter(e => e.getAttribute("attr") === "color")
    .forEach(e => encodings.push(tfParseColorEncoding(e)));
  // Edit Axis: fixed range, tick spacing, "include zero" (encoding attr="space")
  const spaces = Array.from(styleEl.getElementsByTagName("encoding")).filter(e => e.getAttribute("attr") === "space").map(e => tfDefined({
    field: e.getAttribute("field") || undefined,
    scope: e.getAttribute("scope") || undefined,
    axisClass: e.getAttribute("class") || undefined,
    rangeType: e.getAttribute("range-type") || undefined,
    min: tfNum(e.getAttribute("min")),
    max: tfNum(e.getAttribute("max")),
    majorSpacing: tfNum(e.getAttribute("major-spacing")),
    majorShow: e.getAttribute("major-show") || undefined,      // "false": Edit Axis → Tick Marks → None
    domainExpand: e.getAttribute("domain-expand") || undefined
  }));
  // Shape encodings: value → shape ("Zoom Icons/Zoom in.png" = a custom shape, ":filled/circle" = Tableau's)
  const shapes = Array.from(styleEl.getElementsByTagName("encoding")).filter(e => e.getAttribute("attr") === "shape").map(e => {
    /** @type {Record<string, string>} */
    const map = {};
    Array.from(e.getElementsByTagName("map")).forEach(m => Array.from(m.getElementsByTagName("bucket"))
      .forEach(b => { map[tfBucketKey(b.textContent)] = m.getAttribute("to"); }));
    return { field: tfParseFieldRef(e.getAttribute("field")), map };
  }).filter(s => s.field);
  /** @type {ParsedStyle} */
  const out = { rules, encodings };
  if (spaces.length) out.spaces = spaces;
  if (shapes.length) out.shapes = shapes;
  return out;
}

export function tfParseTitle(ownerEl) {
  const lo = tfKid(ownerEl, "layout-options") || tfKid(ownerEl, "layout");
  const t = lo && tfKid(lo, "title");
  const ft = t && tfKid(t, "formatted-text");
  if (!ft) return null;
  const runs = tfParseRuns(ft);
  return { text: runs.map(r => r.text).join("").trim(), runs };
}

/**
 * A calculation that is one number ("0.5", "MIN(0)"): an axis that only positions marks, e.g. the columns
 * of a table built from marks.
 * @param {Element | null} calc @returns {number | undefined}
 */
function tfConstantFormula(calc) {
  const m = calc && String(calc.getAttribute("formula") || "").match(/^\s*(?:(?:MIN|MAX|AVG|SUM|ATTR)\s*\(\s*(-?\d+(?:\.\d+)?)\s*\)|(-?\d+(?:\.\d+)?))\s*$/i);
  return m ? Number(m[1] ?? m[2]) : undefined;
}

/* Raise when the parser reads something new: a workbook model remembered by an older version is then
   parsed again from its stored XML (ui/workbook-store.js), so an update reaches workbooks loaded before it. */
export const FORMAT_MODEL_VERSION = 3;

/** @param {string} xmlString the .twb XML @returns {FormatModel} */
export function parseTableauFormatting(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const root = doc.documentElement;
  const model = { version: FORMAT_MODEL_VERSION, workbookStyle: tfParseStyle(tfKid(root, "style")),
                  fields: {}, datasourceStyles: {}, sheets: {}, dashboards: {}, measureAliases: {} };

  // ── datasource columns: captions, roles, default number formats, colour maps
  tfKids(tfKid(root, "datasources"), "datasource").forEach(ds => {
    const dsName = ds.getAttribute("name");
    model.datasourceStyles[dsName] = tfParseStyle(tfKid(ds, "style"));
    tfKids(ds, "column").forEach(col => {
      const n = (col.getAttribute("name") || "").replace(/^\[|\]$/g, "");
      if (!n) return;
      const info = tfDefined({
        name: n, ds: dsName,
        caption: col.getAttribute("caption") || undefined,
        role: col.getAttribute("role") || undefined,
        datatype: col.getAttribute("datatype") || undefined,
        defaultFormat: col.getAttribute("default-format") || undefined,
        alias: col.getAttribute("alias") || undefined,      // parameters: current value's display text
        value: col.getAttribute("value") || undefined,
        constant: tfConstantFormula(tfKid(col, "calculation"))
      });
      model.fields[dsName + "|" + n.toLowerCase()] = info;
      if (!model.fields["|" + n.toLowerCase()]) model.fields["|" + n.toLowerCase()] = info;
      // Measure Names aliases: key "[ds].[sum:HomeValue:qk]" → "Typical home value"
      if (n === ":Measure Names") {
        Array.from(col.getElementsByTagName("alias")).forEach(al => {
          const r = tfParseFieldRef(al.getAttribute("key"));
          if (r) model.measureAliases[tfRefKey(r)] = al.getAttribute("value");
        });
      }
    });
  });

  // ── worksheets
  tfKids(tfKid(root, "worksheets"), "worksheet").forEach(ws => {
    const name = ws.getAttribute("name");
    const table = tfKid(ws, "table");
    const sheet = {
      name,
      title: tfParseTitle(ws),
      style: tfParseStyle(table && tfKid(table, "style")),
      rows: tfExtractRefs(table && tfKid(table, "rows") && tfKid(table, "rows").textContent),
      cols: tfExtractRefs(table && tfKid(table, "cols") && tfKid(table, "cols").textContent),
      panes: [], fieldRefs: []
    };
    const view = table && tfKid(table, "view");
    // columns declared in this sheet's datasource-dependencies (covers blended/secondary sources)
    Array.from(ws.getElementsByTagName("datasource-dependencies")).forEach(dep => {
      const dsName = dep.getAttribute("datasource");
      tfKids(dep, "column").forEach(col => {
        const n = (col.getAttribute("name") || "").replace(/^\[|\]$/g, "");
        if (!n) return;
        const key = dsName + "|" + n.toLowerCase();
        const cur = model.fields[key] || {};
        model.fields[key] = tfMerge(tfDefined({ name: n, ds: dsName, caption: col.getAttribute("caption") || undefined,
          role: col.getAttribute("role") || undefined, datatype: col.getAttribute("datatype") || undefined,
          defaultFormat: col.getAttribute("default-format") || undefined,
          alias: col.getAttribute("alias") || undefined, value: col.getAttribute("value") || undefined }), cur);
        if (!model.fields["|" + n.toLowerCase()]) model.fields["|" + n.toLowerCase()] = model.fields[key];
      });
    });
    sheet.manualSorts = view ? Array.from(view.getElementsByTagName("manual-sort")).map(ms => ({
      field: tfParseFieldRef(ms.getAttribute("column")),
      direction: (ms.getAttribute("direction") || "ASC").toUpperCase(),
      order: Array.from(ms.getElementsByTagName("bucket")).map(b => b.textContent)
    })) : [];
    sheet.measureSorts = view ? [
      ...Array.from(view.getElementsByTagName("shelf-sort-v2")).map(ss => ({
        field: tfParseFieldRef(ss.getAttribute("dimension-to-sort")),
        measure: tfParseFieldRef(ss.getAttribute("measure-to-sort-by")),
        direction: (ss.getAttribute("direction") || "ASC").toUpperCase() })),
      ...Array.from(view.getElementsByTagName("computed-sort")).map(cs => ({
        field: tfParseFieldRef(cs.getAttribute("column")),
        measure: tfParseFieldRef(cs.getAttribute("using")),
        direction: (cs.getAttribute("direction") || "ASC").toUpperCase() }))
    ].filter(x => x.field && x.measure) : [];
    // members kept by the Measure Names filter, in filter order
    sheet.measureFilter = [];
    if (view) tfKids(view, "filter").forEach(f => {
      if (!/\[:Measure Names\]$/.test(f.getAttribute("column") || "")) return;
      Array.from(f.getElementsByTagName("groupfilter")).forEach(g => {
        const m = g.getAttribute("member");
        if (g.getAttribute("function") === "member" && m) sheet.measureFilter.push(m);
      });
    });
    const panesEl = table && tfKid(table, "panes");
    tfKids(panesEl, "pane").forEach(p => {
      const mark = tfKid(p, "mark");
      const encodings = tfKids(tfKid(p, "encodings")).map(e => ({
        channel: e.tagName, field: tfParseFieldRef(e.getAttribute("column"))
      })).filter(e => e.field);
      const cl = tfKid(p, "customized-label");
      const labelRuns = cl ? tfParseRuns(tfKid(cl, "formatted-text")) : [];
      const ct = tfKid(p, "customized-tooltip");                 // Tooltip editor text
      const tooltipRuns = ct ? tfParseRuns(tfKid(ct, "formatted-text")) : [];
      sheet.panes.push({
        id: p.getAttribute("id") || null,
        xIndex: tfNum(p.getAttribute("x-index")),
        xAxisName: p.getAttribute("x-axis-name") || null,
        yIndex: tfNum(p.getAttribute("y-index")),
        yAxisName: p.getAttribute("y-axis-name") || null,     // dual axis / multi-pane: which measure this pane draws
        markClass: mark ? mark.getAttribute("class") : "Automatic",
        encodings,
        labelRuns: labelRuns.map(r => ({ refs: tfExtractRefs(r.text.replace(/^<|>$/g, "")), text: r.text, props: r.props })),
        tooltipRuns,
        style: tfParseStyle(tfKid(p, "style"))
      });
    });
    // Analytics pane box plot: <reference-line formula='iqr' boxplot-whisker-type='…'>
    sheet.boxPlot = Array.from(ws.getElementsByTagName("reference-line")).some(rl =>
      rl.getAttribute("formula") === "iqr" || rl.getAttribute("boxplot-whisker-type") != null);
    // reference lines (Analytics → Reference Line); box-plot whiskers / distributions are not lines
    const refIds = new Set();
    sheet.referenceLines = Array.from(ws.getElementsByTagName("reference-line")).filter(rl => {
      const id = rl.getAttribute("id") || "";
      if (refIds.has(id) || rl.getAttribute("boxplot-whisker-type") != null) return false;
      refIds.add(id);
      return /^(average|mean|median|sum|total|min|max|constant)$/.test(rl.getAttribute("formula") || "");
    }).map(rl => tfDefined({
      formula: rl.getAttribute("formula"),
      labelType: rl.getAttribute("label-type") || "automatic",
      label: rl.getAttribute("label") || undefined,
      scope: rl.getAttribute("scope") || undefined,
      value: tfNum(rl.getAttribute("value")),
      axis: tfParseFieldRef(rl.getAttribute("axis-column")) || undefined,
      field: tfParseFieldRef(rl.getAttribute("value-column")) || undefined
    }));
    // every field the sheet references (for name matching incl. Measure Names)
    const seen = new Set();
    const add = r => { if (r && !seen.has(r.inner.toLowerCase())) { seen.add(r.inner.toLowerCase()); sheet.fieldRefs.push(r); } };
    sheet.runningTotals = [];                                    // waterfall: running-sum table calcs
    sheet.tableCalcs = {};                                       // quick table calc → its "compute using" (ordering-type)
    Array.from(ws.getElementsByTagName("column-instance")).forEach(ci => {
      const depDs = ci.parentNode && ci.parentNode.getAttribute && ci.parentNode.getAttribute("datasource");
      const ref = tfParseFieldRef((depDs ? "[" + depDs + "]." : "") + ci.getAttribute("name"));
      add(ref);
      const calc = tfKid(ci, "table-calc");
      if (ref && calc && /^(cumtotal|runningtotal)$/i.test(calc.getAttribute("type") || "")) sheet.runningTotals.push(ref);
      if (ref && calc) sheet.tableCalcs[ref.inner.toLowerCase()] = calc.getAttribute("ordering-type") || "";
    });
    sheet.rows.forEach(add); sheet.cols.forEach(add);
    sheet.panes.forEach(p => p.encodings.forEach(e => add(e.field)));
    model.sheets[name] = sheet;
  });

  // ── URL actions (not filter/highlight actions) → Excel hyperlinks
  model.actions = [];
  tfKids(tfKid(root, "actions"), "action").forEach(a => {
    const link = tfKid(a, "link"), src = tfKid(a, "source");
    if (!link || tfKid(a, "command")) return;                   // filter/highlight actions have a <command>
    const expr = link.getAttribute("expression") || "";
    if (!expr || /^tsl:/i.test(expr)) return;
    model.actions.push(tfDefined({
      caption: a.getAttribute("caption") || undefined, expression: expr,
      dashboard: src ? src.getAttribute("dashboard") || undefined : undefined,
      worksheet: src ? src.getAttribute("worksheet") || undefined : undefined,
      exclude: src ? Array.from(src.getElementsByTagName("exclude-sheet")).map(x => x.getAttribute("name")) : []
    }));
  });

  // ── dashboards (title + style)
  tfKids(tfKid(root, "dashboards"), "dashboard").forEach(d => {
    const zones = [];
    const walk = (z, parent) => tfKids(z, "zone").forEach(c => {
      zones.push(tfDefined({
        id: c.getAttribute("id"), name: c.getAttribute("name") || undefined,
        type: c.getAttribute("type-v2") || "worksheet",
        showTitle: c.getAttribute("show-title") !== "false",
        hidden: c.getAttribute("hidden-by-user") === "true",
        x: tfNum(c.getAttribute("x")), y: tfNum(c.getAttribute("y")),
        w: tfNum(c.getAttribute("w")), h: tfNum(c.getAttribute("h")),
        runs: tfKid(c, "formatted-text") ? tfParseRuns(tfKid(c, "formatted-text")) : undefined,
        param: c.getAttribute("param") || undefined,             // image zones: the file inside the .twbx; legends: their field
        scaled: c.getAttribute("is-scaled") === "1" || undefined,
        // Center Image is on unless switched off (is-centered='0')
        centered: c.getAttribute("is-centered") ? c.getAttribute("is-centered") === "1" : undefined,
        url: c.getAttribute("url") || undefined,                   // image / button zones: the link it opens
        parent,                                                     // the layout container it sits in
        style: tfZoneStyle(tfKid(c, "zone-style"))
      }));
      walk(c, c.getAttribute("id") || undefined);
    });
    walk(tfKid(d, "zones"), undefined);   // direct child only → phone/tablet layouts are skipped
    const size = tfKid(d, "size");
    model.dashboards[d.getAttribute("name")] = {
      title: tfParseTitle(d), style: tfParseStyle(tfKid(d, "style")), zones,
      width: size ? tfNum(size.getAttribute("maxwidth")) : undefined,
      height: size ? tfNum(size.getAttribute("maxheight")) : undefined
    };
  });

  // ── each sheet's fit on each dashboard (<window class='dashboard'> viewpoints: "entire-view", "fit-width" …)
  tfKids(tfKid(root, "windows"), "window").forEach(w => {
    const dash = w.getAttribute("class") === "dashboard" && model.dashboards[w.getAttribute("name")];
    if (!dash) return;
    dash.fit = {};
    tfKids(tfKid(w, "viewpoints"), "viewpoint").forEach(v => {
      const z = tfKid(v, "zoom");
      if (z) dash.fit[v.getAttribute("name")] = z.getAttribute("type");
    });
  });

  tfDebugDump(model);
  return model;
}

/* Prints every style element/attr found – use this to see what YOUR TWB contains */
export function tfDebugDump(model) {
  if (!FORMAT_CONFIG.debug) return;
  const summarize = st => {
    const o = {};
    for (const [el, list] of Object.entries(st.rules)) o[el] = [...new Set(list.map(f => f.attr + (f.scope ? "@" + f.scope : "") + (f.field ? "[field]" : "")))].join(", ");
    return o;
  };
  console.log("[Format] Workbook style:", summarize(model.workbookStyle));
  for (const s of Object.values(model.sheets)) {
    console.log(`[Format] Sheet "${s.name}"`, {
      rules: summarize(s.style),
      marks: s.panes.map(p => p.markClass).join(", "),
      encodings: s.panes.flatMap(p => p.encodings.map(e => e.channel + "=" + e.field.inner)).join(", "),
      colorEncodings: s.style.encodings.length,
      labelRuns: s.panes.reduce((n, p) => n + p.labelRuns.length, 0)
    });
  }
}
