"use strict";

// ── TABLEAU BUILT-IN PALETTE LOOKUP ──────────────────────────────────────
const TABLEAU_BUILTIN_PALETTES = {
  "blue_10_0": ["#C7DDEA","#AFCFE1","#97C0D7","#7FAFCA","#689BC0","#5487B1","#4475A0","#356790","#2A5783"],
  "orange_10_0": ["#F3C184", "#F0AE62", "#EE9A42", "#EF882D", "#ED7420", "#E25F1D", "#CC531F", "#B54820", "#9E3D22"],
  "green_10_0": ["#B9D9AF", "#A3CF95", "#8BC57D", "#74BA67", "#5DAA56", "#4D984B", "#3D8743", "#31773F", "#24693D"],
  "red_10_0": ["#F3B8AB", "#EEA08E", "#EA8972", "#E9725B", "#EA5C4C", "#E6453C", "#D92C34", "#C71532", "#AE123A"],
  "purple_10_0": ["#ECC6E3", "#E2BCD8", "#D7AFCA", "#CB9FBC", "#BB85A8", "#AC759B", "#9F6B93", "#8E5B86", "#7C4D79"],
  "brown_10_0": ["#E8D5B4","#DEBE8A","#D8A66B","#D08F57","#C97A4B","#BF6740","#B45539","#AA4435","#9F3632"],
  "gray_10_0": ["#E5E5E5","#D4D6D8","#C0C4C8","#AAB0B6","#959DA5","#818A94","#6D7782","#5B6470","#49525E"],
  "gray_warm_10_0": ["#D8D1CE","#CBC2BE","#BBB0AB","#AB9F9A","#9A8E89","#887C77","#776B67","#685D59","#59504E"],
  "blue_teal_10_0": ["#B7D7D1","#9BC9C7","#80BBC0","#67ACC0","#529DBA","#448DAF","#3A7D9F","#336C91","#2C5985"],
  "orange_gold_10_0": ["#E8C85E","#EDB657","#F0A54A","#F08F32","#EF791F","#E96418","#D5531D","#BA4522","#9E3A26"],
  "green_gold_10_0": ["#E5C75A","#CDBE58","#B1B953","#97B64F","#7FB255","#67A957","#529B53","#348347","#146C36"],
  "red_gold_10_0": ["#E8C85A","#EDA951","#F08C4B","#EC7247","#E65E47","#DF4D47","#D33A45","#C32942","#B71D3E"],
  "orange_blue_diverging_10_0": ["#9E3D22","#B54820","#CC531F","#E25F1D","#ED7420","#F3D9BE","#97C0D7","#689BC0","#4475A0","#356790","#2B5C8A"],
  "red_green_diverging_10_0": ["#AE123A","#C61E3F","#D93443","#E94F4A","#F07A66","#F2B3A3","#8BC97D","#6DB65F","#539F50","#3B8447","#24693D"],
  "green_blue_diverging_10_0": ["#24693D","#347D46","#4E9854","#72B464","#9BCF89","#D4DDD9","#A9C9DC","#7DAACE","#5C8FBC","#4373A0","#2A5783"],
  "red_blue_diverging_10_0": ["#A90C38","#C71F3E","#DC393F","#EB5A4F","#F3A091","#E7E2DE","#B8D1E1","#86B0D1","#5F90BC","#44709C","#2E5A87"],
  "red_black_10_0": ["#AE123A","#C71F3E","#DC393F","#EB5A4F","#F3A091","#DDD9D5","#B9BCBC","#969DA1","#78818A","#606A75","#49525E"],
  "gold_purple_diverging_10_0": ["#AD9024","#B89B34","#C6AA50","#D3BA6D","#DDC892","#E3D7D1","#D7C1D2","#C9A5C3","#BB8AB2","#AC7299"],
  "red_green_gold_diverging_10_0": ["#BE2A3E","#D44344","#E75D49","#F07A47","#F2A14A","#E7C65A","#A8BE5E","#77AF5B","#55994E","#3B8746","#22763F"],
  "sunrise_sunset_diverging_10_0": ["#33608C","#556AA0","#7B67A6","#A664A2","#CD6C95","#EC7C79","#F2B15A","#ED8D46","#E56B44","#D24844","#B81840"],
  "orange_blue_white_diverging_10_0": ["#9E3D22","#B94B20","#D45A1D","#EC7420","#F3B562","#F5F1EC","#C6DDEA","#93BED8","#679BC1","#4677A5","#2B5C8A"],
  "red_green_white_diverging_10_0": ["#AE123A","#C8243F","#DC4947","#ED725E","#F3A38F","#F3F1EE","#B9DFAF","#8BC57D","#5DAA56","#3F8847","#24693D"],
  "green_blue_white_diverging_10_0": ["#24693D","#3E864B","#5BA557","#82C06F","#B7DFAE","#F2F3F1","#C7DDEA","#97C0D7","#689BC0","#4475A0","#2A5783"],
  "red_blue_white_diverging_10_0": ["#A90C38","#C71F3E","#DC393F","#EB5A4F","#F3A091","#F4F3F2","#C6DDEA","#93BED8","#679BC1","#4677A5","#2E5A87"],
  "red_black_white_diverging_10_0": ["#AE123A","#C8243F","#DC4947","#ED725E","#F3A38F","#F3F2F1","#D0D2D1","#A7ADB0","#838C93","#626C77","#49525E"],
  "tableau-blue-light": ["#EEF2F7","#E7EDF5","#DFE8F3","#D8E2F0","#D1DDEE","#CBD9ED","#C9D7F1","#C7D9F2","#C4D8F3"],
  "tableau-orange-light": ["#F6F2EE","#F7EBDD","#F8E3CC","#F9DBBE","#FAD5B3","#FBCFA8","#FCCB9F","#FFCC9E","#FFCC9E"],
  "tableau-orange-blue-light": ["#FFCC9E","#FAD1AB","#F4D8BC","#EEE0CC","#EAE4D8","#E8E8E8","#DFE7EF","#D5E0EC","#CBD9E9","#C7D9F1","#C4D8F3"],
  "tableau-map-blue-green": ["#F5F5C8","#EEF2B3","#E0EA9A","#CBE18F","#AFD695","#91CC9D","#72C3A8","#58BCB5","#41B7C4"],
  "tableau-map-temperatur": ["#529985","#669C76","#81A364","#A6B04E","#D2C63F","#F0D347","#F2C04A","#E7A24A","#D4824D","#C26B51"]
};

// ── HELPER: Get built-in palette colors by name ──────────────────────────
function getBuiltInPaletteColors(paletteName) {
  if (!paletteName) return null;
  
  if (TABLEAU_BUILTIN_PALETTES[paletteName]) {
    return TABLEAU_BUILTIN_PALETTES[paletteName];
  }
  
  const lowerName = paletteName.toLowerCase();
  for (const [key, colors] of Object.entries(TABLEAU_BUILTIN_PALETTES)) {
    if (key.toLowerCase() === lowerName) {
      return colors;
    }
  }
  
  const normalized = paletteName.toLowerCase().replace(/[-\s]+/g, '');
  for (const [key, colors] of Object.entries(TABLEAU_BUILTIN_PALETTES)) {
    const keyNormalized = key.toLowerCase().replace(/[-\s]+/g, '');
    if (keyNormalized.includes(normalized) || normalized.includes(keyNormalized)) {
      return colors;
    }
  }
  
  return null;
}

const AUTOMATIC_PALETTE_BY_MARK = {
  automatic:    "blue_teal_10_0",
  bar:          "blue_10_0",
  line:         "blue_10_0",
  area:         "blue_teal_10_0",
  square:       "blue_teal_10_0",
  circle:       "blue_10_0",
  shape:        "blue_10_0",
  text:         "blue_10_0",
  map:          "blue_teal_10_0",
  multipolygon: "blue_teal_10_0",
  pie:          "blue_10_0",
  ganttbar:     "blue_10_0",
  polygon:      "blue_10_0",
  density:      "blue_10_0",
  heatmap:      "blue_10_0"
};
const DEFAULT_AUTOMATIC_PALETTE = "blue_10_0";

function getAutomaticPaletteForMark(markClass) {
  const key = String(markClass || "Automatic").toLowerCase().replace(/[\s_-]+/g, "");
  const name = AUTOMATIC_PALETTE_BY_MARK[key] || DEFAULT_AUTOMATIC_PALETTE;
  if (!AUTOMATIC_PALETTE_BY_MARK[key]) {
    console.log(`[Auto Palette] Mark type "${markClass}" not in table, using ${name}`);
  }
  return {
    name,
    colors: TABLEAU_BUILTIN_PALETTES[name].map(c => "FF" + c.replace("#", "").toUpperCase())
  };
}


/* ══════════════════════════════════════════════════════════════════════════
 * TABLEAU → EXCEL FORMAT ENGINE
 * ──────────────────────────────────────────────────────────────────────────
 * 1. parseTableauFormatting(xml)  → plain JSON "format model" (DOM based)
 * 2. createSheetFormatter(model, sheetName) → resolves the Tableau cascade
 *      defaults → workbook <style> → worksheet <style> → element → field
 *      → customized-label run
 * 3. tfExcel*() helpers → convert resolved props into ExcelJS styles
 * ══════════════════════════════════════════════════════════════════════════ */

const FORMAT_CONFIG = {
  substituteTableauFonts: true,   // "Tableau Book" etc. are only installed with Tableau Desktop
  tableauFontSubstitute: "Arial",
  writeNativeNumbers: true,       // write real numbers + Excel numFmt instead of text
  applyFallbackHeatmap: false,    // old red/green heatmap when Tableau has no color (not in Tableau → off)
  chartPolicy: "data",            // charts on the dashboard: "skip" | "data" (export their data as a plain table)
  groupOverflowRows: true,        // collapse rows beyond ROW_GROUP_THRESHOLD into an expandable [+]/[-] group
  autoFilter: "largest",          // Excel allows ONE autofilter per sheet: "largest" table | "none"
  textBoxHeaders: true,           // use dashboard text boxes as column headers – only when they line up exactly with the table
  textBoxTitle: true,             // use the top dashboard text box as the dashboard title (rich text)
  linkText: "url",                // URL-action cells: "url" shows the link itself, any other string is shown as the text
  debug: true
};

const VISUAL_TYPES = Object.freeze({
  TABLE: "TABLE",
  KPI: "KPI",
  HEATMAP: "HEATMAP",
  BAR: "BAR",
  COLUMN: "COLUMN",
  LINE: "LINE",
  AREA: "AREA",
  PIE: "PIE",
  SCATTER: "SCATTER",
  MAP: "MAP",
  GANTT: "GANTT",
  CUSTOM: "CUSTOM",
  UNKNOWN: "UNKNOWN"
});

/* Approximate Tableau Desktop defaults. Compare against one of your exports
 * and tweak these once – everything explicitly set in the TWB overrides them. */
const TABLEAU_DEFAULTS = {
  worksheet:  { fontName: "Tableau Book", fontSize: 9, color: "FF333333" },
  fieldLabel: { bold: true },
  title:      { fontName: "Tableau Book", fontSize: 15, color: "FF333333" },
  dashTitle:  { fontName: "Tableau Book", fontSize: 18, color: "FF333333" },
  divider:    { visible: true, style: "thin", color: "FFD4D4D4", level: 1 }
};

/* XML element names used in <style-rule element='…'>. Arrays = aliases seen
 * across Tableau versions; unknown ones are simply ignored. */
const TF_ELEMENTS = {
  datalabel:  ["datalabel"],        // Marks card → Label → font (per pane)
  sheet:      ["all", "worksheet"],
  title:      ["title", "worksheet-title"],
  dashTitle:  ["dash-title", "dashboard-title", "title"],
  pane:       ["pane"],
  cell:       ["cell"],
  header:     ["header", "label"],   // "header" = box (shading/size), "label" = header TEXT (font/colour/underline/alignment)
  fieldLabel: ["field-labels", "field-labels-decoration"],
  divider:    ["table-div"]
};

const TABLEAU_10 = ["FF4E79A7", "FFF28E2B", "FFE15759", "FF76B7B2", "FF59A14F",
                    "FFEDC948", "FFB07AA1", "FFFF9DA7", "FF9C755F", "FFBAB0AC"];

const TF_DERIV_LABEL = {
  sum: "SUM", avg: "AVG", cnt: "CNT", ctd: "CNTD", min: "MIN", max: "MAX",
  med: "MEDIAN", attr: "ATTR", usr: "AGG", std: "STDEV", stdp: "STDEVP",
  var: "VAR", varp: "VARP", yr: "YEAR", qr: "QUARTER", mn: "MONTH", wk: "WEEK",
  dy: "DAY", hr: "HOUR", tyr: "YEAR", tqr: "QUARTER", tmn: "MONTH", twk: "WEEK", tdy: "DAY"
};

/* ── small helpers ─────────────────────────────────────────────────────── */
function tfLog(...a) { if (FORMAT_CONFIG.debug) console.log("[Format]", ...a); }
function tfKids(el, tag) {
  if (!el) return [];
  return Array.from(el.childNodes).filter(n => n.nodeType === 1 && (!tag || n.tagName === tag));
}
function tfKid(el, tag) { return tfKids(el, tag)[0] || null; }
function tfNorm(s) { return String(s == null ? "" : s).toLowerCase().replace(/\s+/g, " ").trim(); }
function tfNum(v) { const n = parseFloat(v); return isFinite(n) ? n : undefined; }
function tfDefined(o) {
  const r = {};
  for (const k in o) if (o[k] !== undefined) r[k] = o[k];
  return r;
}

/* Tableau colours: "#rrggbb", "#rgb" or "#rrggbbaa" (aa=00 → transparent) → ARGB */
function tfArgb(v) {
  if (v == null) return undefined;
  let s = String(v).trim().replace(/^#/, "");
  if (/^[0-9a-f]{8}$/i.test(s)) {
    if (s.slice(6) === "00") return null;          // explicit "no fill"
    s = s.slice(0, 6);
  }
  if (/^[0-9a-f]{3}$/i.test(s)) s = s.split("").map(c => c + c).join("");
  if (!/^[0-9a-f]{6}$/i.test(s)) return undefined;
  return "FF" + s.toUpperCase();
}

function tfBrightness(argb) {
  const r = parseInt(argb.substring(2, 4), 16), g = parseInt(argb.substring(4, 6), 16), b = parseInt(argb.substring(6, 8), 16);
  return (r * 299 + g * 587 + b * 114) / 1000;
}

function tfInterpolate(c1, c2, t) {
  const p = (c, i) => parseInt(c.substring(i, i + 2), 16);
  const ch = i => Math.round(p(c1, i) + (p(c2, i) - p(c1, i)) * t).toString(16).padStart(2, "0");
  return ("FF" + ch(2) + ch(4) + ch(6)).toUpperCase();
}
function tfSample(colors, pos) {
  pos = Math.max(0, Math.min(1, pos));
  if (colors.length === 1) return colors[0];
  const seg = pos * (colors.length - 1);
  const i = Math.min(Math.floor(seg), colors.length - 2);
  return tfInterpolate(colors[i], colors[i + 1], seg - i);
}

/* ── Field references ──────────────────────────────────────────────────────
 * "[federated.0x].[sum:Sales:qk]" → {ds, deriv:"sum", name:"Sales", type:"qk"}
 * type: qk = continuous measure, nk/ok = discrete                          */
function tfParseFieldRef(ref) {
  if (!ref) return null;
  ref = String(ref).trim().replace(/^"|"$/g, "");
  const parts = ref.match(/\[[^\]]*\]/g) || [];
  if (!parts.length) return { raw: ref, ds: null, deriv: null, name: ref, type: null, inner: ref };
  const inner = parts[parts.length - 1].slice(1, -1);
  const ds = parts.length > 1 ? parts[0].slice(1, -1) : null;
  const seg = inner.split(":");
  let deriv = null, name = inner, type = null;
  // "[usr:Calculation_1:ok:9]" → type is the nk/ok/qk segment, not the last one
  let ti = -1;
  for (let i = seg.length - 1; i >= 2; i--) if (/^(nk|ok|qk)$/.test(seg[i])) { ti = i; break; }
  if (ti >= 2) { deriv = seg[0] || null; type = seg[ti]; name = seg.slice(1, ti).join(":"); }
  else if (seg.length === 2 && seg[0] === "") name = seg[1];          // [:Measure Names]
  return { raw: ref, ds, deriv, name, type, inner };
}
function tfExtractRefs(text) {
  if (!text) return [];
  const m = String(text).match(/(?:\[[^\]]+\]\.)?\[[^\]]+\]/g) || [];
  return m.map(tfParseFieldRef).filter(Boolean);
}
function tfRefKey(r) { return ((r.ds || "") + "|" + r.inner).toLowerCase(); }
function tfSameField(a, b) {
  if (!a || !b) return false;
  if (a.inner.toLowerCase() === b.inner.toLowerCase()) return true;
  return tfNorm(a.name) === tfNorm(b.name) && (a.deriv || "none") === (b.deriv || "none");
}

/* ══════════════════════════════════════════════════════════════════════════
 * 1. PARSER
 * ══════════════════════════════════════════════════════════════════════════ */
function tfParseRunProps(run) {
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
function tfParseRuns(ftEl) {
  return tfKids(ftEl, "run").map(r => ({ text: r.textContent || "", props: tfParseRunProps(r) }));
}

function tfBucketKey(t) {
  let s = String(t == null ? "" : t).trim();
  if (/^".*"$/.test(s)) s = s.slice(1, -1).replace(/""/g, '"');
  else if (/^#.*#$/.test(s)) s = s.slice(1, -1);
  if (s === "%null%") s = "null";
  return s.toLowerCase();
}

function tfParseColorEncoding(enc) {
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

function tfParseStyle(styleEl) {
  const rules = {}, encodings = [];
  if (!styleEl) return { rules, encodings };
  tfKids(styleEl, "style-rule").forEach(rule => {
    const el = rule.getAttribute("element") || "all";
    rules[el] = rules[el] || [];
    tfKids(rule, "format").forEach(f => rules[el].push(tfDefined({
      attr: f.getAttribute("attr"),
      value: f.getAttribute("value"),
      field: f.getAttribute("field") || undefined,
      scope: f.getAttribute("scope") || undefined,
      dataClass: f.getAttribute("data-class") || undefined
    })));
  });
  Array.from(styleEl.getElementsByTagName("encoding"))
    .filter(e => e.getAttribute("attr") === "color")
    .forEach(e => encodings.push(tfParseColorEncoding(e)));
  return { rules, encodings };
}

function tfParseTitle(ownerEl) {
  const lo = tfKid(ownerEl, "layout-options") || tfKid(ownerEl, "layout");
  const t = lo && tfKid(lo, "title");
  const ft = t && tfKid(t, "formatted-text");
  if (!ft) return null;
  const runs = tfParseRuns(ft);
  return { text: runs.map(r => r.text).join("").trim(), runs };
}

function parseTableauFormatting(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const root = doc.documentElement;
  const model = { version: 2, workbookStyle: tfParseStyle(tfKid(root, "style")),
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
        value: col.getAttribute("value") || undefined
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
      sheet.panes.push({
        id: p.getAttribute("id") || null,
        xIndex: tfNum(p.getAttribute("x-index")),
        xAxisName: p.getAttribute("x-axis-name") || null,
        markClass: mark ? mark.getAttribute("class") : "Automatic",
        encodings,
        labelRuns: labelRuns.map(r => ({ refs: tfExtractRefs(r.text.replace(/^<|>$/g, "")), text: r.text, props: r.props })),
        style: tfParseStyle(tfKid(p, "style"))
      });
    });
    // every field the sheet references (for name matching incl. Measure Names)
    const seen = new Set();
    const add = r => { if (r && !seen.has(r.inner.toLowerCase())) { seen.add(r.inner.toLowerCase()); sheet.fieldRefs.push(r); } };
    Array.from(ws.getElementsByTagName("column-instance")).forEach(ci => {
      const depDs = ci.parentNode && ci.parentNode.getAttribute && ci.parentNode.getAttribute("datasource");
      add(tfParseFieldRef((depDs ? "[" + depDs + "]." : "") + ci.getAttribute("name")));
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
    const walk = z => tfKids(z, "zone").forEach(c => {
      zones.push(tfDefined({
        id: c.getAttribute("id"), name: c.getAttribute("name") || undefined,
        type: c.getAttribute("type-v2") || "worksheet",
        showTitle: c.getAttribute("show-title") !== "false",
        hidden: c.getAttribute("hidden-by-user") === "true",
        x: tfNum(c.getAttribute("x")), y: tfNum(c.getAttribute("y")),
        w: tfNum(c.getAttribute("w")), h: tfNum(c.getAttribute("h")),
        runs: tfKid(c, "formatted-text") ? tfParseRuns(tfKid(c, "formatted-text")) : undefined
      }));
      walk(c);
    });
    walk(tfKid(d, "zones"));            // direct child only → phone/tablet layouts are skipped
    const size = tfKid(d, "size");
    model.dashboards[d.getAttribute("name")] = {
      title: tfParseTitle(d), style: tfParseStyle(tfKid(d, "style")), zones,
      width: size ? tfNum(size.getAttribute("maxwidth")) : undefined,
      height: size ? tfNum(size.getAttribute("maxheight")) : undefined
    };
  });

  tfDebugDump(model);
  return model;
}

/* Prints every style element/attr found – use this to see what YOUR TWB contains */
function tfDebugDump(model) {
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

/* ══════════════════════════════════════════════════════════════════════════
 * 2. RESOLVER
 * ══════════════════════════════════════════════════════════════════════════ */
function tfFormatsToProps(m) {
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
function tfCollect(style, elements, opts = {}) {
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

function tfMerge(...layers) {
  const out = {};
  layers.forEach(l => { if (l) for (const k in l) if (l[k] !== undefined) out[k] = l[k]; });
  return out;
}

function tfFieldInfo(model, ref) {
  if (!model || !ref) return null;
  const k = ref.name.toLowerCase();
  return model.fields[(ref.ds || "") + "|" + k] || model.fields["|" + k] || null;
}

function tfDisplayNames(model, ref) {
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

function createSheetFormatter(model, sheetName) {
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

/* ══════════════════════════════════════════════════════════════════════════
 * VIEW MODEL — rebuild the table Tableau actually draws from summary data
 *   1. pivot Measure Names / Measure Values into one column per measure
 *   2. merge stacked multi-pane rows (one block per pane → one row)
 *   3. sort rows like the view (manual sort, sort-by-measure, else natural)
 *   4. keep only visible columns, in visual order, with Tableau's header labels
 * ══════════════════════════════════════════════════════════════════════════ */
function tfIsNull(dv) {
  if (!dv) return true;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  return v === null || v === undefined || v === "%null%";
}
function tfDvText(dv) { return tfIsNull(dv) ? "" : String(dv.formattedValue != null ? dv.formattedValue : dv.value); }
function tfDvNum(dv) {
  if (tfIsNull(dv)) return null;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-eE]/g, ""));
  return isFinite(n) && /^[\s$€£¥(+-]*[\d.,]+/.test(String(v)) ? n : null;
}
function tfNaturalCompare(a, b) {
  const an = tfIsNull(a), bn = tfIsNull(b);
  if (an || bn) return an === bn ? 0 : an ? 1 : -1;           // nulls last
  const av = a.nativeValue !== undefined ? a.nativeValue : a.value;
  const bv = b.nativeValue !== undefined ? b.nativeValue : b.value;
  if (typeof av === "number" && typeof bv === "number") return av - bv;
  if (av instanceof Date && bv instanceof Date) return av - bv;
  return tfDvText(a).localeCompare(tfDvText(b), undefined, { numeric: true, sensitivity: "base" });
}

function isKPIViewModel(vm) {
  const vis = vm.order.map(i => vm.cols[i]);
  return vm.rows.length === 1 && vis.length >= 1 && vis.length <= 8 && !vis.some(c => c.isHeader);
}

function normalizeVisualToken(value) {
  return String(value == null ? "" : value).toLowerCase().replace(/[\s_-]+/g, "");
}

function visualSpecMarkTokens(spec) {
  if (!spec || typeof spec !== "object") return [];
  const tokens = [];
  const add = value => {
    if (value == null) return;
    if (Array.isArray(value)) value.forEach(add);
    else if (typeof value === "object") {
      [value.markType, value.markClass, value.mark, value.marksType, value.type, value.class, value.name].forEach(add);
    } else tokens.push(normalizeVisualToken(value));
  };
  [spec.markType, spec.markClass, spec.mark, spec.type, spec.visualType, spec.marksType, spec.marks].forEach(add);
  if (Array.isArray(spec.marksSpecifications)) {
    const activeIndex = Number.isInteger(spec.activeMarksSpecificationIndex)
      ? spec.activeMarksSpecificationIndex
      : -1;
    if (spec.marksSpecifications[activeIndex]) add(spec.marksSpecifications[activeIndex]);
    else spec.marksSpecifications.forEach(add);
  }
  if (spec.activeMarksSpecification && typeof spec.activeMarksSpecification === "object") {
    add(spec.activeMarksSpecification);
  }
  (spec.panes || []).forEach(add);
  return [...new Set(tokens.filter(Boolean))];
}

function classifyVisualType(spec, vm) {
  const tokens = visualSpecMarkTokens(spec);
  if (vm && vm.fmt && typeof vm.fmt.colorEncoding === "function") {
    const encoding = vm.fmt.colorEncoding();
    if (encoding && encoding.markClass) tokens.push(normalizeVisualToken(encoding.markClass));
  }
  const token = tokens.find(Boolean) || "";
  const byMark = {
    bar: VISUAL_TYPES.BAR,
    line: VISUAL_TYPES.LINE,
    area: VISUAL_TYPES.AREA,
    pie: VISUAL_TYPES.PIE,
    circle: VISUAL_TYPES.SCATTER,
    scatter: VISUAL_TYPES.SCATTER,
    map: VISUAL_TYPES.MAP,
    multipolygon: VISUAL_TYPES.MAP,
    polygon: VISUAL_TYPES.MAP,
    ganttbar: VISUAL_TYPES.GANTT,
    gantt: VISUAL_TYPES.GANTT,
    heatmap: VISUAL_TYPES.HEATMAP,
    vizextension: VISUAL_TYPES.CUSTOM
  };
  if (byMark[token]) return byMark[token];
  if (vm.kind === "chart") return VISUAL_TYPES.UNKNOWN;
  if (isKPIViewModel(vm)) return VISUAL_TYPES.KPI;
  return VISUAL_TYPES.TABLE;
}

function buildVisualModel(model, sheetName, summary, opts = {}) {
  const vm = buildViewModel(model, sheetName, summary, opts);
  const spec = opts.visualSpec || null;
  const type = classifyVisualType(spec, vm);
  const dimensions = vm.order.filter(i => vm.cols[i].isHeader).map(i => vm.cols[i]);
  const measures = vm.order.filter(i => !vm.cols[i].isHeader).map(i => vm.cols[i]);
  return {
    type,
    title: vm.title,
    data: { rows: vm.rows, columns: vm.cols, dimensions, measures },
    axes: spec && (spec.axes || spec.axis) || {},
    encodings: spec && (spec.encodings || spec.encoding) || {},
    panes: spec && spec.panes || [],
    style: spec && (spec.style || spec.styles) || {},
    layout: opts.layout || null,
    source: { visualSpec: spec },
    metadata: {
      worksheetName: sheetName,
      markClass: vm.fmt && vm.fmt.colorEncoding ? ((vm.fmt.colorEncoding() || {}).markClass || null) : null,
      visualSpecAvailable: !!spec,
      visualSpecError: opts.visualSpecError ? String(opts.visualSpecError.message || opts.visualSpecError) : null,
      visualSpecKeys: spec && typeof spec === "object" ? Object.keys(spec) : [],
      imageApiAvailable: typeof tableau !== "undefined" && tableau.extensions &&
        typeof tableau.extensions.createVizImageAsync === "function"
    },
    viewModel: vm,
    diagnostics: []
  };
}

function chooseVisualRenderer(visualModel) {
  const cellTypes = new Set([VISUAL_TYPES.TABLE, VISUAL_TYPES.KPI, VISUAL_TYPES.HEATMAP]);
  if (cellTypes.has(visualModel.type)) return VISUAL_RENDERERS.cell;
  if (VISUAL_RENDERERS.imageTypes.has(visualModel.type) && visualModel.metadata.imageApiAvailable) {
    return VISUAL_RENDERERS.image;
  }
  if (!visualModel.metadata.visualSpecAvailable) {
    return { ...VISUAL_RENDERERS.fallback, reason: "visual specification unavailable" };
  }
  return VISUAL_RENDERERS.fallback;
}

const VISUAL_RENDERERS = Object.freeze({
  cell: Object.freeze({ renderer: "cell", status: "success", reason: "cell-native visual" }),
  image: Object.freeze({ renderer: "tableau-image", status: "pending", reason: "Tableau image renderer selected" }),
  imageTypes: new Set([VISUAL_TYPES.BAR, VISUAL_TYPES.LINE, VISUAL_TYPES.AREA, VISUAL_TYPES.PIE, VISUAL_TYPES.SCATTER, VISUAL_TYPES.MAP]),
  fallback: Object.freeze({ renderer: "data-fallback", status: "warning", reason: "visual renderer not implemented yet" })
});

function visualDataValue(dv) {
  if (tfIsNull(dv)) return null;
  return dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
}

function buildVizImageSpec(visualModel, width, height) {
  const vm = visualModel.viewModel;
  const visibleColumns = vm.order.map(index => vm.cols[index]);
  const values = vm.rows.map(row => {
    const value = {};
    visibleColumns.forEach((column, position) => {
      value[column.name || `Field${position + 1}`] = visualDataValue(row[vm.order[position]]);
    });
    return value;
  });
  const dimensions = visualModel.data.dimensions;
  const measures = visualModel.data.measures;
  const dimension = dimensions[0] || visibleColumns[0];
  const numericMeasures = measures.length ? measures : visibleColumns.filter(column => !column.isHeader);
  const firstMeasure = numericMeasures[0] || visibleColumns[1];
  const secondMeasure = numericMeasures[1] || null;
  const latitude = visibleColumns.find(column => /^(latitude|generated latitude)$/i.test(column.name) || /latitude/i.test(column.name));
  const longitude = visibleColumns.find(column => /^(longitude|generated longitude)$/i.test(column.name) || /longitude/i.test(column.name));
  if (visualModel.type === VISUAL_TYPES.MAP && (!latitude || !longitude)) {
    throw new Error("The map has no usable latitude and longitude mapping");
  }
  if (!dimension || !firstMeasure) throw new Error("The visual has no usable dimension and measure mapping");

  const mark = {
    [VISUAL_TYPES.BAR]: "bar",
    [VISUAL_TYPES.LINE]: "line",
    [VISUAL_TYPES.AREA]: "area",
    [VISUAL_TYPES.PIE]: "pie",
    [VISUAL_TYPES.SCATTER]: "circle",
    [VISUAL_TYPES.MAP]: "map"
  }[visualModel.type];
  let encoding;
  if (visualModel.type === VISUAL_TYPES.MAP) {
    encoding = { columns: { field: longitude.name, type: "continuous" }, rows: { field: latitude.name, type: "continuous" } };
  } else if (visualModel.type === VISUAL_TYPES.SCATTER && secondMeasure) {
    encoding = { columns: { field: firstMeasure.name, type: "continuous" }, rows: { field: secondMeasure.name, type: "continuous" } };
  } else if (visualModel.type === VISUAL_TYPES.PIE) {
    encoding = { color: { field: dimension.name, type: "discrete" }, angle: { field: firstMeasure.name, type: "continuous" } };
  } else {
    encoding = { columns: { field: dimension.name, type: "discrete" }, rows: { field: firstMeasure.name, type: "continuous" } };
  }
  const color = dimensions[1];
  if (visualModel.type === VISUAL_TYPES.MAP) {
    const mapColor = dimensions.find(column => column.name !== latitude.name && column.name !== longitude.name);
    if (mapColor) encoding.color = { field: mapColor.name, type: "discrete", palette: "tableau20_10_0" };
  } else if (color && color.name !== dimension.name && visualModel.type !== VISUAL_TYPES.PIE) {
    encoding.color = { field: color.name, type: "discrete", palette: "tableau20_10_0" };
  }
  return {
    version: 2,
    description: visualModel.title && visualModel.title.text || visualModel.metadata.worksheetName,
    size: { width, height },
    data: { values },
    mark,
    markcolor: "#4E79A7",
    encoding
  };
}

function svgToPngDataUrl(svg, width, height) {
  return new Promise((resolve, reject) => {
    const blob = new Blob([svg], { type: "image/svg+xml" });
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        canvas.getContext("2d").drawImage(image, 0, 0, width, height);
        URL.revokeObjectURL(url);
        resolve(canvas.toDataURL("image/png"));
      } catch (err) {
        URL.revokeObjectURL(url);
        reject(err);
      }
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Tableau returned an SVG that could not be rasterized"));
    };
    image.src = url;
  });
}

async function renderTableauImage(visualModel, width, height) {
  if (typeof tableau === "undefined" || !tableau.extensions ||
      typeof tableau.extensions.createVizImageAsync !== "function") {
    throw new Error("Tableau image API is unavailable in this runtime");
  }
  const spec = buildVizImageSpec(visualModel, width, height);
  const svg = await tableau.extensions.createVizImageAsync(spec);
  if (typeof svg !== "string" || !svg.trim()) throw new Error("Tableau image API returned no SVG");
  return svgToPngDataUrl(svg, width, height);
}

function buildViewModel(model, sheetName, summary, opts = {}) {
  const fmt = createSheetFormatter(model, sheetName);
  const notes = [];
  let cols = (summary.columns || []).map((c, i) => {
    const name = c.fieldName || c.fieldId || `Col${i + 1}`;
    return { name, dataType: c.dataType, ref: fmt.matchName(name) };
  });
  let rows = (summary.data || []).map(r => r.slice());

  // ── 1. pivot Measure Names / Measure Values ─────────────────────────────
  const mnI = cols.findIndex(c => /^measure names$/i.test(c.name));
  const mvI = cols.findIndex(c => /^measure values$/i.test(c.name));
  if (mnI >= 0 && mvI >= 0) {
    const keyIdx = cols.map((_, i) => i).filter(i => i !== mnI && i !== mvI);
    const order = fmt.measureOrder();
    const measures = [];                         // [{text, ref, rank}]
    const groups = new Map();
    rows.forEach(r => {
      const text = tfDvText(r[mnI]);
      if (!measures.some(m => m.text === text)) {
        const ref = fmt.measureRefForName(text);
        const pos = ref ? order.findIndex(o => tfSameField(o, ref)) : -1;
        measures.push({ text, ref, rank: pos >= 0 ? pos : 1e6 + measures.length });
      }
      const key = keyIdx.map(i => tfDvText(r[i])).join("\u0001");
      if (!groups.has(key)) groups.set(key, { keyRow: keyIdx.map(i => r[i]), vals: {} });
      const g = groups.get(key);
      if (g.vals[text] === undefined) g.vals[text] = r[mvI];
    });
    measures.sort((a, b) => a.rank - b.rank);
    cols = [...keyIdx.map(i => cols[i]),
            ...measures.map(m => ({ name: m.text, ref: m.ref, pivoted: true, dataType: "float" }))];
    rows = [...groups.values()].map(g => [...g.keyRow, ...measures.map(m => g.vals[m.text] || null)]);
    notes.push(`pivoted ${measures.length} measures → ${rows.length} rows`);
  }

  // column roles
  cols.forEach(c => {
    c.isHeader = !c.pivoted && fmt.hasModel ? fmt.isHeaderField(c.ref)
               : !c.pivoted && !/^(int|float|real|integer|number)$/i.test(String(c.dataType || ""));
  });

  // ── 2. merge stacked multi-pane rows ────────────────────────────────────
  const hdrIdx = cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);
  if (hdrIdx.length && rows.length > 1) {
    const groups = new Map();
    rows.forEach(r => {
      const key = hdrIdx.map(i => tfDvText(r[i])).join("\u0001");
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    });
    if (groups.size < rows.length) {
      let conflict = false;
      const merged = [];
      for (const g of groups.values()) {
        const out = g[0].slice();
        for (let ci = 0; ci < cols.length && !conflict; ci++) {
          const vals = [...new Set(g.filter(r => !tfIsNull(r[ci])).map(r => tfDvText(r[ci])))];
          if (vals.length > 1) conflict = true;
          else { const hit = g.find(r => !tfIsNull(r[ci])); if (hit) out[ci] = hit[ci]; }
        }
        if (conflict) break;
        merged.push(out);
      }
      if (!conflict) { notes.push(`merged ${rows.length} stacked rows → ${merged.length}`); rows = merged; }
      else notes.push("rows share header values but differ – not merged");
    }
  }

  // ── 3. sort like the view ───────────────────────────────────────────────
  const sortLevels = [];
  (fmt.hasModel ? fmt.headerRefs() : []).forEach(ref => {
    const ci = cols.findIndex(c => c.isHeader && c.ref && tfSameField(c.ref, ref));
    if (ci < 0) return;
    const manual = fmt.manualSortFor(ref);
    const byMeasure = fmt.measureSortFor(ref);
    let mi = -1;
    if (byMeasure) mi = cols.findIndex(c => c.ref && tfSameField(c.ref, byMeasure.measure));
    sortLevels.push({ ci, manual, byMeasure: mi >= 0 ? byMeasure : null, mi });
  });
  if (sortLevels.length) {
    // measure sort aggregates the measure over the rows sharing the same outer path
    const aggCache = sortLevels.map((lvl, L) => {
      if (!lvl.byMeasure) return null;
      const m = new Map();
      rows.forEach(r => {
        const k = sortLevels.slice(0, L + 1).map(x => tfDvText(r[x.ci])).join("\u0001");
        m.set(k, (m.get(k) || 0) + (tfDvNum(r[lvl.mi]) || 0));
      });
      return m;
    });
    const manualRank = lvl => {
      if (!lvl.manual) return null;
      const map = new Map(lvl.manual.order.map((b, i) => [tfBucketKey(b), i]));
      return map;
    };
    const ranks = sortLevels.map(manualRank);
    rows.sort((a, b) => {
      for (let L = 0; L < sortLevels.length; L++) {
        const lvl = sortLevels[L];
        let d = 0;
        if (ranks[L]) {
          const ra = ranks[L].get(tfNorm(tfDvText(a[lvl.ci]))), rb = ranks[L].get(tfNorm(tfDvText(b[lvl.ci])));
          d = (ra === undefined ? 1e9 : ra) - (rb === undefined ? 1e9 : rb);
          if (lvl.manual.direction === "DESC") d = -d;
        } else if (lvl.byMeasure) {
          const path = r => sortLevels.slice(0, L + 1).map(x => tfDvText(r[x.ci])).join("\u0001");
          d = (aggCache[L].get(path(a)) || 0) - (aggCache[L].get(path(b)) || 0);
          if (lvl.byMeasure.direction === "DESC") d = -d;
        }
        if (!d) d = tfNaturalCompare(a[lvl.ci], b[lvl.ci]);
        if (d) return d;
      }
      return 0;
    });
  }

  // ── URL actions → link columns (kept visible even when the field is only on Detail)
  const linkActions = fmt.urlActions(opts.dashboardName);
  linkActions.forEach(a => {
    const refs = tfExtractRefs(a.expression);
    if (!refs.length) return;
    const ci = cols.findIndex(c => !c.isHeader && c.ref && tfNorm(c.ref.name) === tfNorm(refs[0].name));
    if (ci < 0) return;
    cols[ci].link = { caption: a.caption, expression: a.expression,
      parts: refs.map(r => ({ token: r, ci: cols.findIndex(c => c.ref && tfNorm(c.ref.name) === tfNorm(r.name)) })) };
  });

  // ── 4. visible columns in visual order + header labels ──────────────────
  let order;
  if (!fmt.hasModel) {
    order = cols.map((_, i) => i);
  } else {
    const panes = fmt.panesInOrder();
    const headerRefs = fmt.headerRefs();
    const place = [];
    cols.forEach((c, i) => {
      if (c.isHeader) {
        if (fmt.isLabelHidden(c.ref)) return;                       // Format → Hide header
        const pos = headerRefs.findIndex(r => tfSameField(r, c.ref));
        place.push({ i, g: 0, a: pos, b: 0 });
      } else if (c.pivoted) {
        place.push({ i, g: 1, a: 0, b: i });
      } else if (!c.ref) {
        place.push({ i, g: 2, a: 1e6, b: i });                     // unknown → keep, at the end
      } else {
        let hit = null;
        panes.forEach((p, pi) => {
          const k = p.refs.findIndex(r => tfSameField(r, c.ref));
          if (k >= 0 && !hit) hit = { a: pi, b: k };
        });
        if (hit) place.push({ i, g: 2, a: hit.a, b: hit.b });      // shown as a mark label
        else if (c.link) {                                         // link column: in the pane that carries the field
          const pi = panes.findIndex(p => p.pane.encodings.some(e => tfSameField(e.field, c.ref)));
          place.push({ i, g: 2, a: pi >= 0 ? pi : 1e5, b: 999 });
        }
        // else: detail / tooltip / colour-only / axis measure → not visible
      }
    });
    place.sort((x, y) => x.g - y.g || x.a - y.a || x.b - y.b);
    order = place.map(p => p.i);
    if (!order.length) { order = cols.map((_, i) => i); notes.push("no visible columns resolved – exporting all"); }
  }

  const labelsRows = fmt.fieldLabelsShown("rows"), labelsCols = fmt.fieldLabelsShown("cols");
  cols.forEach(c => {
    if (c.pivoted) c.label = c.name;                                          // Measure Names alias
    else if (c.isHeader) {
      const shown = fmt.shelfOf(c.ref) === "cols" ? labelsCols : labelsRows;
      c.label = shown ? fmt.captionFor(c.ref, c.name) : "";
    } else c.label = fmt.captionFor(c.ref, c.name);
    if (c.link && c.link.caption) c.label = c.link.caption;
  });

  // ── dashboard text boxes drawn as column headers (strict match, else captions stay) ──
  const dashM = model && opts.dashboardName && model.dashboards && model.dashboards[opts.dashboardName];
  let headerZoneIds = [];
  if (FORMAT_CONFIG.textBoxHeaders && dashM && !labelsRows) {
    const tb = tfTextBoxHeaders(dashM, sheetName, order.length);
    if (tb) {
      order.forEach((ci, k) => {
        const z = tb.zones[k];
        cols[ci].label = z.text;
        cols[ci].labelProps = z.props;
        if (dashM.width) cols[ci].zoneWidthPx = z.w / 100000 * dashM.width;
      });
      headerZoneIds = tb.zones.map(z => z.id);
      notes.push(`headers from ${tb.zones.length} dashboard text boxes`);
    }
  }
  const showHeaderRow = order.some(i => cols[i].label);

  // title
  const valuesFor = ref => {
    const ci = cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
    if (ci < 0) return [];
    return [...new Set(rows.map(r => tfDvText(r[ci])).filter(Boolean))];
  };
  const title = fmt.renderTitle(valuesFor, opts.displayName || sheetName);
  const dash = model && opts.dashboardName && model.dashboards && model.dashboards[opts.dashboardName];
  const zone = dash && dash.zones.find(z => z.name === sheetName && z.type === "worksheet");
  const showTitle = zone ? zone.showTitle : true;

  if (notes.length) tfLog(`View model "${sheetName}": ${notes.join("; ")}`);
  return {
    fmt, cols, rows, order, notes, showHeaderRow, title, showTitle, headerZoneIds,
    dashboardName: opts.dashboardName,
    kind: fmt.isChart() ? "chart" : "table",
    headerOrder: sortLevels.map(l => l.ci)          // all header cols (incl. hidden) outer → inner
  };
}

/* Text boxes count as column headers only if ALL hold:
 *   – ≥ 2 text zones, visible, in one row (same top edge ± tol)
 *   – their bottom edge touches the sheet's top edge (± tol)
 *   – they lie within the sheet's horizontal span
 *   – their number equals the number of visible columns
 * Positions are Tableau's own zone coordinates (0–100000), so the check is deterministic. */
const TF_ZONE_TOL = 1500;
function tfZoneText(z) { return (z.runs || []).map(r => r.text).join("").replace(/\u00C6\r?\n?/g, "\n").trim(); }
function tfTextBoxHeaders(dash, sheetName, nCols) {
  const ws = dash.zones.find(z => z.name === sheetName && z.type === "worksheet");
  if (!ws || nCols < 2) return null;
  const near = (a, b) => Math.abs(a - b) <= TF_ZONE_TOL;
  const cand = dash.zones.filter(z => z.type === "text" && !z.hidden && z.runs && tfZoneText(z) &&
    near(z.y + z.h, ws.y) && z.x >= ws.x - TF_ZONE_TOL && z.x + z.w <= ws.x + ws.w + TF_ZONE_TOL);
  if (cand.length !== nCols || !cand.every(z => near(z.y, cand[0].y))) return null;
  cand.sort((a, b) => a.x - b.x);
  return { zones: cand.map(z => ({ id: z.id, text: tfZoneText(z), w: z.w,
    props: (z.runs.find(r => r.text.trim()) || { props: {} }).props })) };
}
/* Dashboard title: Tableau's own title if shown, else the text box above every worksheet
 * with the largest font (ties → topmost). Returns runs for rich text, or null. */
function tfDashboardTitleRuns(model, dashName, usedZoneIds) {
  const dash = model && model.dashboards && model.dashboards[dashName];
  if (!dash) return null;
  const titleZone = dash.zones.find(z => z.type === "title" && !z.hidden);
  if (titleZone && dash.title && dash.title.runs.length) return dash.title.runs;
  if (!FORMAT_CONFIG.textBoxTitle) return null;
  const wsTop = Math.min(...dash.zones.filter(z => z.type === "worksheet").map(z => z.y));
  const size = z => (z.runs.find(r => r.text.trim()) || { props: {} }).props.fontSize || 9;
  const cand = dash.zones.filter(z => z.type === "text" && !z.hidden && z.runs && tfZoneText(z) &&
    !usedZoneIds.has(z.id) && z.y + z.h <= wsTop + TF_ZONE_TOL);
  if (!cand.length) return null;
  cand.sort((a, b) => size(b) - size(a) || a.y - b.y);
  return cand[0].runs;
}

function tfStrokeToBorder(size) {
  if (!size) return null;
  if (size <= 1) return "thin";
  if (size === 2) return "medium";
  return "thick";
}

/* ── Colour scales ─────────────────────────────────────────────────────── */
function tfBuildColorScale(fmt, enc, values) {
  const def = enc.def;
  if (enc.continuous) {
    const nums = values.filter(v => typeof v === "number" && isFinite(v));
    if (!nums.length) return null;
    const min = def && def.min !== undefined ? def.min : Math.min(...nums);
    const max = def && def.max !== undefined ? def.max : Math.max(...nums);
    let colors = def && def.customColors.length ? def.customColors : null;
    let diverging = /diverging/i.test((def && (def.paletteName || "")) + " " + (def && (def.paletteType || "")));
    if (!colors && def && def.paletteName && !/^automatic$/i.test(def.paletteName)) {
      const pc = getBuiltInPaletteColors(def.paletteName);
      if (pc) colors = pc.map(tfArgb);
    }
    if (!colors) {                     // Tableau "Automatic"
      if (min < 0 && max > 0) { colors = getBuiltInPaletteColors("orange_blue_diverging_10_0").map(tfArgb); diverging = true; }
      else colors = getAutomaticPaletteForMark(enc.markClass).colors;
    }
    if (def && def.reverse) colors = colors.slice().reverse();
    let center = def && def.center !== undefined ? def.center : null;
    if (diverging && center === null) center = (min < 0 && max > 0) ? 0 : (min + max) / 2;
    tfLog(`Gradient for ${enc.ref.inner}: ${colors.length} stops, range ${min}..${max}${center !== null ? ", center " + center : ""}`);
    return v => {
      if (typeof v !== "number" || !isFinite(v)) return null;
      let pos;
      if (center !== null) {
        const dev = Math.max(Math.abs(max - center), Math.abs(min - center)) || 1;
        pos = 0.5 + (v - center) / (2 * dev);       // symmetric, like "Use full colour range" = off
      } else pos = max === min ? 0.5 : (v - min) / (max - min);
      return tfSample(colors, pos);
    };
  }
  // categorical
  const map = {};
  if (def) for (const [k, c] of Object.entries(def.map)) fmt.bucketAliases(k).forEach(a => { map[tfNorm(a)] = c; });
  const used = new Set(Object.values(map));
  const pool = TABLEAU_10.filter(c => !used.has(c));
  const auto = {};
  [...new Set(values.filter(v => v != null && v !== "").map(tfNorm))].filter(k => !(k in map)).sort().forEach((k, i) => {
    const p = pool.length ? pool : TABLEAU_10;
    auto[k] = p[i % p.length];
  });
  tfLog(`Categorical colours for ${enc.ref.inner}: ${Object.keys(map).length} from TWB, ${Object.keys(auto).length} auto-assigned`);
  return v => (v == null || v === "") ? null : (map[tfNorm(v)] || auto[tfNorm(v)] || null);
}

/* ══════════════════════════════════════════════════════════════════════════
 * 3. NUMBER FORMATS
 * ══════════════════════════════════════════════════════════════════════════ */

/* Tableau text-format → Excel numFmt.
 * - strips Tableau's type prefix (n/c/p/e/*); locale "standard" codes (C1033…) → null (inferred instead)
 * - moves thousands-scaling commas behind the decimals: "#,##0,.0K" → "#,##0.0,\"K\""
 * - quotes every literal letter (Excel rejects bare K, M, yrs …)
 * - validates the result; anything doubtful → null so we never write an invalid format */
function tableauToExcelNumFmt(raw) {
  if (!raw) return null;
  let s = String(raw);
  if (/^[A-Za-z]\d*%?$/.test(s)) return null;
  if (/^[ncpes*]/i.test(s) && !/^[#0]/.test(s)) s = s.slice(1);
  const sections = tfSplitSections(s);
  if (!sections || sections.length > 4) return null;
  const out = sections.map(tfConvertSection);
  if (out.some(x => x === null)) return null;
  const res = out.join(";");
  return /[0#]/.test(res) ? res : null;
}
function tfSplitSections(s) {
  const secs = []; let cur = "", q = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') q = !q;
    if (ch === "\\" && !q) { cur += ch + (s[i + 1] || ""); i++; continue; }
    if (ch === ";" && !q) { secs.push(cur); cur = ""; continue; }
    cur += ch;
  }
  if (q) return null;                        // unbalanced quotes
  secs.push(cur);
  return secs;
}
function tfConvertSection(sec) {
  // tokens: quoted literal | escaped char | number pattern | other char
  const toks = [];
  const re = /"[^"]*"|\\.|[#0?,.]+(?:[eE][+-][0#]+)?|[\s\S]/g;
  let m, numDone = false;
  while ((m = re.exec(sec))) {
    const t = m[0];
    if (t[0] === '"' || t[0] === "\\") { toks.push(t); continue; }
    if (/^[#0?,.]/.test(t) && /[#0?]/.test(t) && !numDone) {
      numDone = true;
      const [intPart, ...decParts] = t.split(".");
      const dec = decParts.join("");
      const intTrail = (intPart.match(/,+$/) || [""])[0];                 // "#,##0," → scale
      const decScale = (dec.match(/,+$/) || [""])[0];
      const intClean = intPart.slice(0, intPart.length - intTrail.length);
      const decClean = dec.replace(/,+$/, "");
      toks.push(intClean + (decParts.length ? "." + decClean : "") + intTrail + decScale);
      continue;
    }
    if (/[A-Za-z]/.test(t)) {                                              // literal letter → quote
      toks.push('"' + t + '"');
      continue;
    }
    toks.push(t);
  }
  // merge adjacent quoted literals: "K""g" → "Kg"
  return toks.join("").replace(/""/g, "");
}

/* Parse a Tableau formatted value, build an Excel format and only accept it
 * if it reproduces the same number (round-trip check). */
function inferExcelNumFmt(formatted, value) {
  if (formatted == null || typeof value !== "number" || !isFinite(value)) return null;
  const f = String(formatted).trim();
  const m = f.match(/^(\(?)([^\d(]*?)(-?)(\d[\d.,\s\u00A0']*)(.*?)(\)?)$/);
  if (!m) return null;
  const [, lpar, prefixRaw, minus, core, suffixRaw, rpar] = m;
  const neg = !!(lpar && rpar) || !!minus || /^-/.test(prefixRaw);
  const prefix = prefixRaw.replace(/-/g, "").trim() ? prefixRaw.replace(/-/g, "") : "";
  const suffix = suffixRaw;

  // decide decimal separator by trying both interpretations
  const tryParse = dec => {
    const grpSrc = dec === "." ? "[,\\s\\u00A0']" : "[.\\s\\u00A0']";
    const clean = core.trim().replace(new RegExp(grpSrc, "g"), "").replace(dec, ".");
    const n = parseFloat(clean);
    const decimals = clean.includes(".") ? clean.split(".")[1].length : 0;
    return { n, decimals, grouped: new RegExp(grpSrc).test(core.trim()) };
  };
  const unitScale = { "%": 100, K: 1e-3, M: 1e-6, B: 1e-9, Bn: 1e-9 };
  const unit = (suffix.trim().match(/^(%|K|M|Bn|B)\b/) || suffix.trim().match(/^%/) || [])[0] || null;
  const scaled = Math.abs(value) * (unit ? unitScale[unit] : 1);

  for (const dec of [".", ","]) {
    const p = tryParse(dec);
    if (!isFinite(p.n)) continue;
    const tol = 0.5 * Math.pow(10, -p.decimals) + 1e-9 * Math.max(1, scaled);
    let plainPercent = false;
    if (Math.abs(scaled - p.n) > tol) {
      if (unit === "%" && Math.abs(Math.abs(value) - p.n) <= tol) plainPercent = true; // value already ×100
      else continue;
    }
    let body = (p.grouped ? "#,##0" : "0") + (p.decimals ? "." + "0".repeat(p.decimals) : "");
    const q = t => t ? '"' + t.replace(/"/g, '""') + '"' : "";
    let rest = suffix;
    if (unit === "%" && !plainPercent) { body += "%"; rest = suffix.replace("%", ""); }
    else if (unit && unit !== "%") { body += ",".repeat({ K: 1, M: 2, B: 3, Bn: 3 }[unit]) + q(unit); rest = suffix.replace(unit, ""); }
    const pos = q(prefix) + body + q(rest);
    if (!neg) return pos;
    if (lpar && rpar) return pos + ";(" + pos + ")";
    return minus ? pos + ";" + q(prefix) + "-" + body + q(rest) : pos + ";-" + pos;
  }
  return null;
}

/* ══════════════════════════════════════════════════════════════════════════
 * 4. EXCEL MAPPING
 * ══════════════════════════════════════════════════════════════════════════ */
function tfExcelFont(p) {
  let name = p.fontName || "Arial";
  let bold = !!p.bold;
  if (FORMAT_CONFIG.substituteTableauFonts && /^tableau\b/i.test(name)) {
    if (/\b(bold|semibold|black|heavy)\b/i.test(name)) bold = true;  // weight lives in the family name
    name = FORMAT_CONFIG.tableauFontSubstitute;
  }
  const f = { name, size: p.fontSize || 9, bold, italic: !!p.italic };
  if (p.underline) f.underline = true;
  if (p.strike) f.strike = true;
  if (p.color) f.color = { argb: p.color };
  return f;
}
function tfExcelAlignment(p, isNumber) {
  return { horizontal: p.hAlign || (isNumber ? "right" : "left"), vertical: p.vAlign || "middle", wrapText: !!p.wrap };
}
function tfExcelFill(argb) {
  return argb ? { type: "pattern", pattern: "solid", fgColor: { argb } } : undefined;
}

function tfRichRuns(runs, baseProps) {
  const out = [];
  runs.forEach(r => {
    const text = r.text.replace(/\u00C6\r?\n?/g, "\n");
    if (text) out.push({ text, font: tfExcelFont(tfMerge(baseProps, r.props)) });
  });
  return out;
}

/* writes value + style; returns true if the cell holds a real number */
function tfWriteCell(cell, dv, p, extra = {}) {
  const native = dv ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : undefined;
  const formatted = dv ? (dv.formattedValue != null ? dv.formattedValue : (native != null ? String(native) : "")) : "";
  let isNumber = false;
  if (dv && (dv.value === null || dv.nativeValue === null || dv.value === "%null%")) {   // Tableau Null → blank
    cell.value = "";
  } else if (FORMAT_CONFIG.writeNativeNumbers && typeof native === "number" && isFinite(native)) {
    const numFmt = p.numFmt || inferExcelNumFmt(formatted, native);
    if (numFmt) { cell.value = native; cell.numFmt = numFmt; isNumber = true; }
  }
  if (!isNumber && cell.value !== "") cell.value = formatted;
  if (extra.hyperlink) cell.value = { text: extra.hyperlinkText || extra.hyperlink, hyperlink: extra.hyperlink };
  const looksNumeric = isNumber || /^[^\d]{0,3}-?[\d.,]+[^\d]{0,3}$/.test(String(formatted).trim());
  const fill = extra.fill !== undefined ? extra.fill : p.bgColor;
  const font = tfExcelFont(extra.fontColor ? { ...p, color: extra.fontColor } : p);
  if (fill && !extra.fontColor && !p.explicitColor && tfBrightness(fill) < 128) font.color = { argb: "FFFFFFFF" };
  cell.font = font;
  cell.alignment = tfExcelAlignment(p, looksNumeric);
  if (fill) cell.fill = tfExcelFill(fill);
  if (extra.border) {
    const b = {};
    for (const k in extra.border) if (extra.border[k]) b[k] = extra.border[k];
    if (Object.keys(b).length) cell.border = b;
  }
  return isNumber;
}

if (typeof module !== "undefined" && module.exports) {  // lets you unit-test the engine in Node
  module.exports = { parseTableauFormatting, createSheetFormatter, tfBuildColorScale, tableauToExcelNumFmt,
    inferExcelNumFmt, tfWriteCell, tfExcelFont, tfArgb, tfParseFieldRef, buildVisualModel,
    classifyVisualType, chooseVisualRenderer, buildVizImageSpec, VISUAL_TYPES, VISUAL_RENDERERS, FORMAT_CONFIG, TABLEAU_DEFAULTS };
}

(function () {

  /* ── Layout constants ─────────────────────────────────────────────────── */
  const PX_PER_COL  = 90;
  const PX_PER_ROW  = 22;
  const TITLE_ROWS  = 1;
  const COL_GAP     = 0;  
  const ROW_GAP     = 1;  
  const ROW_GROUP_THRESHOLD = 8;

  function getExcelColumnName(colIndex) {
    let columnName = "";
    let dividend = colIndex + 1;

    while (dividend > 0) {
      let modulo = (dividend - 1) % 26;
      columnName = String.fromCharCode(65 + modulo) + columnName;
      dividend = Math.floor((dividend - modulo) / 26);
    }

    return columnName;
  }

  async function fetchAllSheetsData(sheets, concurrency = 4) {
  const results = [];
  const queue = [...sheets];
  
  async function worker() {
    while (queue.length) {
      const sheet = queue.shift();
      try {
        const [dataResult, visualSpecResult] = await Promise.allSettled([
          sheet.getSummaryDataAsync(),
          typeof sheet.getVisualSpecificationAsync === "function"
            ? sheet.getVisualSpecificationAsync()
            : Promise.reject(new Error("getVisualSpecificationAsync is unavailable"))
        ]);
        if (dataResult.status === "rejected") throw dataResult.reason;
        results.push({
          sheet,
          data: dataResult.value,
          visualSpec: visualSpecResult.status === "fulfilled" ? visualSpecResult.value : null,
          visualSpecError: visualSpecResult.status === "rejected" ? visualSpecResult.reason : null,
          error: null
        });
      } catch (err) {
        results.push({ sheet, data: null, error: err });
      }
    }
  }
  
  const workers = Array(concurrency).fill().map(() => worker());
  await Promise.all(workers);
  return results;
}

  /* ── Range tracker ────────────────────────────────────────────────────── */
  function makeRangeTracker() {
    let minR = Infinity, minC = Infinity, maxR = 0, maxC = 0;
    return {
      update(r, c) {
        if (r < minR) minR = r;
        if (c < minC) minC = c;
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
      },
      toRange() {
        return {
          s: { r: minR === Infinity ? 0 : minR, c: minC === Infinity ? 0 : minC },
          e: { r: maxR, c: maxC },
        };
      },
    };
  }

  function updateVisualStatus(statuses) {
    const target = document.getElementById("export_status");
    if (!target) return;
    if (!statuses || !statuses.length) {
      target.textContent = "No worksheet visual status available";
      return;
    }
    const counts = statuses.reduce((out, item) => {
      const key = item.status === "success" ? "native" : "fallback";
      out[key] = (out[key] || 0) + 1;
      return out;
    }, {});
    const warnings = statuses.filter(item => item.status !== "success").length;
    target.textContent = `${statuses.length} visuals discovered - ${counts.native || 0} cell-native, ${counts.fallback || 0} data fallback${warnings ? ` (${warnings} warning${warnings === 1 ? "" : "s"})` : ""}`;
    target.classList.toggle("status-warning", warnings > 0);
  }

  function isNumeric(value) {
    if (value === null || value === undefined) return false;
    
    const cleaned = String(value)
      .replace(/,/g, "")
      .replace(/%/g, "")
      .trim();
    
    return cleaned !== "" && !isNaN(Number(cleaned));
  }

  function shouldSkipConditionalFormatting(columnName, value, fieldRole = null) {
    if (!columnName) return false;
    
    const lowerColName = columnName.toLowerCase();
    
    const categoricalKeywords = [
      'year', 'quarter', 'month', 'date', 'time', 
      'category', 'region', 'product', 'name', 'id',
      'country', 'city', 'state', 'department', 'type',
      'status', 'group', 'segment', 'class'
    ];
    
    const geoKeywords = [
      'latitude', 'lat', 'latitud', 'ycoord', 'y_coord',
      'longitude', 'long', 'lon', 'longitud', 'xcoord', 'x_coord',
      'location', 'geo', 'geography', 'coordinates', 'coords',
      'postal', 'zip', 'zipcode', 'postcode', 'address'
    ];
    
    for (const keyword of geoKeywords) {
      if (lowerColName.includes(keyword)) {
        console.log(`🗺️ Skipping conditional formatting for geographic column: "${columnName}"`);
        return true;
      }
    }
    
    for (const keyword of categoricalKeywords) {
      if (lowerColName.includes(keyword)) {
        return true;
      }
    }
    
    if (value !== undefined && value !== null) {
      const strValue = String(value).trim();
      const numValue = parseFloat(strValue);
      
      if (!isNaN(numValue)) {
        if (lowerColName.includes('lat') && numValue >= -90 && numValue <= 90) {
          console.log(`🗺️ Skipping conditional formatting for latitude column: "${columnName}" (value: ${numValue})`);
          return true;
        }
        
        if (lowerColName.includes('lon') && numValue >= -180 && numValue <= 180) {
          console.log(`🗺️ Skipping conditional formatting for longitude column: "${columnName}" (value: ${numValue})`);
          return true;
        }
        
        if (lowerColName.match(/lat/i) && numValue >= -90 && numValue <= 90) {
          console.log(`🗺️ Skipping conditional formatting for geographic coordinate column: "${columnName}"`);
          return true;
        }
        
        if (lowerColName.match(/lon/i) && numValue >= -180 && numValue <= 180) {
          console.log(`🗺️ Skipping conditional formatting for geographic coordinate column: "${columnName}"`);
          return true;
        }
      }
      
      if (/^\d{4}$/.test(strValue)) {
        const yearNum = parseInt(strValue, 10);
        if (yearNum >= 1900 && yearNum <= 2100) {
          return true;
        }
      }
    }
    
    return false;
  }

  function isGeographicCoordinate(value, columnName) {
    if (!isNumeric(value)) return false;
    
    const lowerColName = columnName.toLowerCase();
    const numValue = parseFloat(String(value).trim());
    
    if (lowerColName.includes('latitude') || lowerColName.includes('lat')) {
      return numValue >= -90 && numValue <= 90;
    }
    
    if (lowerColName.includes('longitude') || lowerColName.includes('lon')) {
      return numValue >= -180 && numValue <= 180;
    }
    
    if (numValue >= -90 && numValue <= 90 && (lowerColName.includes('coord') || lowerColName.includes('geo'))) {
      return true;
    }
    
    if ((numValue >= -180 && numValue <= -90) || (numValue >= 90 && numValue <= 180)) {
      if (lowerColName.includes('coord') || lowerColName.includes('geo')) {
        return true;
      }
    }
    
    return false;
  }

  function isNumericForFormatting(value, columnName) {
    if (!isNumeric(value)) return false;
    
    if (columnName) {
      const strValue = String(value).trim();
      if (/^\d{4}$/.test(strValue)) {
        const yearNum = parseInt(strValue, 10);
        if (yearNum >= 1900 && yearNum <= 2100) {
          return false;
        }
      }
      
      if (isGeographicCoordinate(value, columnName)) {
        return false;
      }
    }
    
    return true;
  }

  function getZeroCenteredColor(value, maxAbs) {
    if (maxAbs === 0) {
      return "FFFFFF";
    }
    
    const intensity = Math.min(Math.abs(value) / maxAbs, 1);
    
    let r = 255;
    let g = 255;
    let b = 0;
    
    if (value > 0) {
      r = Math.round(255 * (1 - intensity));
      g = 255;
      b = Math.round(100 * (1 - intensity));
    } else if (value < 0) {
      r = 255;
      g = Math.round(255 * (1 - intensity));
      b = Math.round(100 * (1 - intensity));
    }
    
    return (
      "FF" +
      r.toString(16).padStart(2, "0") +
      g.toString(16).padStart(2, "0") +
      b.toString(16).padStart(2, "0")
    ).toUpperCase();
  }

  /* ── FALLBACK: Hardcoded conditional formatting (only used if no XML colors) ── */
  function applyConditionalFormattingToTable(worksheet, rows, headers, startRow, startCol, colWidths) {
    headers.forEach((header, colIndex) => {
      const columnName = header.fieldName || header.fieldId || `Column_${colIndex + 1}`;
      const firstValue = rows[0]?.[colIndex]?.formattedValue || rows[0]?.[colIndex]?.value;
      
      let shouldSkip = shouldSkipConditionalFormatting(columnName, firstValue);
      
      if (shouldSkip) {
        console.log(`⚠️ Skipping conditional formatting for column: "${columnName}"`);
        return;
      }
      
      const numericValues = [];
      
      rows.forEach(row => {
        const value = row[colIndex]?.formattedValue || row[colIndex]?.value;
        if (isNumericForFormatting(value, columnName)) {
          let numValue = String(value).replace(/,/g, "").replace(/%/g, "");
          numericValues.push(parseFloat(numValue));
        }
      });
      
      if (numericValues.length === 0) return;
      
      const maxAbs = Math.max(...numericValues.map(v => Math.abs(v)));
      
      rows.forEach((row, rowIndex) => {
        const rawValue = row[colIndex]?.formattedValue || row[colIndex]?.value;
        
        if (!isNumericForFormatting(rawValue, columnName)) return;
        
        let numValue = String(rawValue).replace(/,/g, "").replace(/%/g, "");
        const value = parseFloat(numValue);
        
        const excelRow = startRow + rowIndex;
        const excelCol = startCol + colIndex;
        const cell = worksheet.getCell(excelRow + 1, excelCol + 1);
        
        const color = getZeroCenteredColor(value, maxAbs);
        
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: color }
        };
      });
    });
  }

  /* ── Filter Extraction ─────────────────────────────────────────────────── */
  async function extractFilterValuesPerField(sheets) {
    const filterMap = new Map();
    
    for (const worksheet of sheets) {
      let filters = [];
      
      try {
        filters = await worksheet.getFiltersAsync();
      } catch (e) {
        console.warn(`Could not get filters for ${worksheet.name}:`, e);
        continue;
      }
      
      filters.forEach(filter => {
        const ignoredFields = ["Measure Names", "Measure Values"];
        if (ignoredFields.includes(filter.fieldName) || /^Month\(/i.test(filter.fieldName)) {
          return;
        }
        
        if (!filterMap.has(filter.fieldName)) {
          filterMap.set(filter.fieldName, {
            values: new Set(),
            worksheetNames: new Set()
          });
        }
        
        const filterData = filterMap.get(filter.fieldName);
        filterData.worksheetNames.add(worksheet.name);
        
        let value = "";
        
        if (filter.filterType === "categorical") {
          const appliedValues = (filter.appliedValues || []).map(v => v.formattedValue);
          appliedValues.forEach(v => filterData.values.add(v));
          value = appliedValues.join(", ");
        } else if (filter.filterType === "range") {
          value = (filter.minValue?.formattedValue || "") + " - " + (filter.maxValue?.formattedValue || "");
          filterData.values.add(value);
        } else {
          value = filter.filterType;
          filterData.values.add(value);
        }
      });
    }
    
    const result = {};
    for (const [fieldName, data] of filterMap.entries()) {
      result[fieldName] = Array.from(data.values);
    }
    
    return result;
  }

  /* ── Check if worksheet is a filter value table ───────────────────────── */
  function isFilterValueWorksheet(sheetName, summaryData) {
    if (/filter[_\- ]?\d+/i.test(sheetName) || 
        /values[_\- ]?\d+/i.test(sheetName) ||
        /_filter_\d+/i.test(sheetName)) {
      return true;
    }
    
    if (summaryData && summaryData.columns && summaryData.columns.length === 1) {
      const firstColumnName = summaryData.columns[0]?.fieldName || "";
      if (firstColumnName === "Values" || firstColumnName === "VALUE") {
        return true;
      }
    }
    
    return false;
  }

  /* ── Write Individual Filter Value Table ──────────────────────────────── */
  function writeIndividualFilterTable(worksheet, filterName, filterValues, originRow, originCol, rangeTracker) {
    let r = originRow;
    const C = originCol;

    let cleanName = filterName.replace(/_(Filter|filter)_\d+$/, '');
    cleanName = cleanName.replace(/_(Values|values)_\d+$/, '');

    titleCell(worksheet, r, C, `📋 ${cleanName}`);
    rangeTracker.update(r, C);
    r++;

    tableHeaderCell(worksheet, r, C, "SELECTED VALUE(S)");
    rangeTracker.update(r, C);
    r++;

    const values = Array.isArray(filterValues) ? filterValues : [filterValues];
    values.forEach((value, idx) => {
      tableDataCell(worksheet, r + idx, C, value, idx);
      rangeTracker.update(r + idx, C);
    });

    return 2 + values.length;
  }

  /* ── ExcelJS Cell Styling Functions ───────────────────────────────────── */
  function setCellValue(worksheet, row, col, value, options = {}) {
    const cell = worksheet.getCell(row + 1, col + 1);
    
    if (typeof value === 'number') {
      cell.value = value;
    } else {
      cell.value = value;
    }
    
    cell.font = {
      bold: options.bold || false,
      size: options.size || 11,
      name: 'Calibri',
      italic: options.italic || false
    };
    
    if (options.color) {
      cell.font.color = { argb: options.color };
    }
    
    cell.alignment = {
      vertical: 'middle',
      horizontal: options.align || 'left',
      wrapText: options.wrapText || false
    };
    
    if (options.bgColor) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: options.bgColor }
      };
    }
    
    if (options.border) {
      cell.border = options.border;
    }
  }

  const thinBorder = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };
  
  const thickBorder = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' }
  };

  function subtitleCell(worksheet, row, col, value) {
    setCellValue(worksheet, row, col, value, {
      bold: false,
      size: 11,
      color: "FF666666",
      align: "center"
    });
  }

  function titleCell(worksheet, row, col, value) {
    setCellValue(worksheet, row, col, value, {
      bold: true,
      size: 12,
      bgColor: "FFE8F0FE",
      align: "center",
      border: thickBorder
    });
  }

  function tableHeaderCell(worksheet, row, col, value) {
    setCellValue(worksheet, row, col, value, {
      bold: true,
      size: 11,
      bgColor: "FF1A73E8",
      color: "FFFFFFFF",
      align: "center",
      border: thinBorder,
      wrapText: true
    });
  }

  function tableDataCell(worksheet, row, col, value, rowIndex = 0) {
    const bgColor = (rowIndex % 2 === 0) ? null : "FFF9F9F9";
    setCellValue(worksheet, row, col, value, {
      align: "left",
      size: 11,
      bgColor: bgColor,
      border: thinBorder,
      wrapText: true
    });
  }

  function writeDashboardTitle(worksheet, dashboardName, exportDate, originRow, originCol, rangeTracker, titleProps, titleRuns) {
    let r = originRow;
    const C = originCol;
    const p = titleProps || tfMerge(TABLEAU_DEFAULTS.dashTitle);
    writeTableauTitle(worksheet, r, C, dashboardName, p, 5);
    if (titleRuns && titleRuns.length) {                    // text-box title → rich text, one run per Tableau run
      const rich = tfRichRuns(titleRuns, tfMerge(TABLEAU_DEFAULTS.worksheet));
      const cell = worksheet.getCell(r + 1, C + 1);
      cell.value = { richText: rich };
      cell.alignment = { horizontal: (titleRuns[0].props.hAlign || "left"), vertical: "top", wrapText: true };
      const lineSizes = [];
      let cur = 0;
      rich.forEach(x => x.text.split("\n").forEach((seg, i) => {
        if (i > 0) { lineSizes.push(cur); cur = 0; }
        if (seg) cur = Math.max(cur, x.font.size || 9);
      }));
      lineSizes.push(cur);
      worksheet.getRow(r + 1).height = Math.ceil(lineSizes.reduce((a, b) => a + (b || 9) * 1.3, 0) + 4);
    }
    rangeTracker.update(r, C);
    rangeTracker.update(r, C + 4);
    r++;

    subtitleCell(worksheet, r, C, `Exported on: ${exportDate}`);
    worksheet.getCell(r + 1, C + 1).font = { name: tfExcelFont(p).name, size: 9, color: { argb: "FF888888" } };
    worksheet.mergeCells(r + 1, C + 1, r + 1, C + 5);
    rangeTracker.update(r, C);
    rangeTracker.update(r, C + 4);
    r++;
    r++;
    return r - originRow;
  }

  /* =============================================================================
   * parseTwbXmlInBrowser(xmlString) - Parses TITLES from XML
   * ============================================================================= */
  function parseTwbXmlInBrowser(xmlString) {
    const doc = new DOMParser().parseFromString(xmlString, "text/xml");
    const titleMap = {};

    const worksheetNodes = doc.getElementsByTagName("worksheet");

    for (let i = 0; i < worksheetNodes.length; i++) {
      const wsNode = worksheetNodes[i];
      const internalName = wsNode.getAttribute("name");
      if (!internalName) continue;

      const layoutNode = getDirectChildByTag(wsNode, "layout-options")
                      || getDirectChildByTag(wsNode, "layout");
      if (!layoutNode) continue;

      const titleNode = getDirectChildByTag(layoutNode, "title");
      if (!titleNode) continue;

      const fmtNode = titleNode.getElementsByTagName("formatted-text")[0];
      if (!fmtNode) continue;

      const runNodes = fmtNode.getElementsByTagName("run");
      const displayTitle = Array.from(runNodes)
        .map(r => r.textContent || "")
        .join("")
        .trim();

      if (displayTitle) {
        titleMap[internalName] = displayTitle;
        console.log(`[TWB parse] Title: "${internalName}" → "${displayTitle}"`);
      }
    }

    return titleMap;
  }

  function getDirectChildByTag(parent, tagName) {
    for (let i = 0; i < parent.childNodes.length; i++) {
      const child = parent.childNodes[i];
      if (child.nodeType === 1 && child.tagName === tagName) return child;
    }
    return null;
  }

  /* =============================================================================
   * loadWorkbookFile() - Loads .twb/.twbx and parses BOTH titles AND colors
   * ============================================================================= */
  async function loadWorkbookFile() {
    return new Promise((resolve) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".twb,.twbx";
      input.style.display = "none";
      document.body.appendChild(input);

      input.onchange = async (event) => {
        document.body.removeChild(input);

        const file = event.target.files[0];
        if (!file) {
          console.log("[loadWorkbookFile] No file selected");
          resolve({});
          return;
        }

        console.log(`[loadWorkbookFile] Reading: ${file.name}`);

        try {
          let xmlString;
          const ext = file.name.split(".").pop().toLowerCase();

          if (ext === "twb") {
            xmlString = await new Promise((res, rej) => {
              const reader = new FileReader();
              reader.onload = (e) => res(e.target.result);
              reader.onerror = () => rej(new Error("FileReader failed reading .twb"));
              reader.readAsText(file, "utf-8");
            });
          } else if (ext === "twbx") {
            const arrayBuffer = await new Promise((res, rej) => {
              const reader = new FileReader();
              reader.onload = (e) => res(e.target.result);
              reader.onerror = () => rej(new Error("FileReader failed reading .twbx"));
              reader.readAsArrayBuffer(file);
            });

            if (typeof JSZip === "undefined") {
              throw new Error("JSZip not loaded. Add the JSZip script tag to index.html.");
            }

            const zip = await JSZip.loadAsync(arrayBuffer);
            const twbEntry = Object.values(zip.files).find(
              f => !f.dir && /\.twb$/i.test(f.name)
            );

            if (!twbEntry) {
              const entries = Object.values(zip.files).filter(f => !f.dir).map(f => f.name).slice(0, 20);
              throw new Error(`No .twb file found inside the .twbx archive. Entries: ${entries.join(", ") || "none"}`);
            }

            xmlString = await twbEntry.async("string");
          } else {
            throw new Error(`Unsupported file type ".${ext}". Please select a .twb or .twbx file.`);
          }

          const titleMap = parseTwbXmlInBrowser(xmlString);
          console.log(`[loadWorkbookFile] Parsed ${Object.keys(titleMap).length} titles`);

          // ── NEW: one DOM-based pass extracts all formatting (fonts, colours, number formats…)
          const formatModel = parseTableauFormatting(xmlString);
          FORMAT_MODEL_CACHE = formatModel;
          tableau.extensions.settings.set("twbTitleMap", JSON.stringify(titleMap));
          try {
            tableau.extensions.settings.set("twbFormatModel", JSON.stringify(formatModel));
          } catch (e) {
            console.warn("[loadWorkbookFile] Format model too large for settings – kept in memory only:", e.message);
          }
          tableau.extensions.settings.set("twbFileName", file.name);
          try {
            await tableau.extensions.settings.saveAsync();
          } catch (e) {
            console.warn("[loadWorkbookFile] settings.saveAsync failed (model kept in memory):", e.message);
          }

          const fileLabel = document.getElementById("twb_file_label");
          if (fileLabel) {
            const sheetCount = Object.keys(formatModel.sheets).length;
            const colorCount = Object.values(formatModel.sheets).filter(s => s.panes.some(p => p.encodings.some(e => e.channel === "color"))).length;
            fileLabel.textContent = `✅ ${file.name} — ${Object.keys(titleMap).length} titles, formatting for ${sheetCount} sheets (${colorCount} with colour)`;
          }


          resolve(titleMap);

        } catch (err) {
          console.error("[loadWorkbookFile] Error:", err.message);
          alert(`Could not read workbook file:\n${err.message}`);
          resolve({});
        }
      };

      input.oncancel = () => {
        document.body.removeChild(input);
        resolve({});
      };

      input.click();
    });
  }


  /* =============================================================================
   * getTitleMap() - Gets titles from settings   * ============================================================================= */
  function getTitleMap() {
    try {
      const saved = tableau.extensions.settings.get("twbTitleMap");
      if (!saved) return {};
      return JSON.parse(saved);
    } catch (err) {
      console.warn("[getTitleMap] Could not read settings:", err.message);
      return {};
    }
  }

  /* =============================================================================
   * getFormatModel() - format model from memory, else from extension settings
   * ============================================================================= */
  let FORMAT_MODEL_CACHE = null;
  function getFormatModel() {
    if (FORMAT_MODEL_CACHE) return FORMAT_MODEL_CACHE;
    try {
      const saved = tableau.extensions.settings.get("twbFormatModel");
      FORMAT_MODEL_CACHE = saved ? JSON.parse(saved) : null;
    } catch (err) {
      console.warn("[getFormatModel] Could not read settings:", err.message);
    }
    return FORMAT_MODEL_CACHE;
  }


  /* ── buildLayoutMap ─────────────────────────────────────────────────── */
  function buildLayoutMap(dashboardObjects, titleMap = {}) {
    const map = new Map();
    
    const positionableObjects = (dashboardObjects || []).filter(
      (obj) => (obj.type === "worksheet" || obj.type === "filter" || obj.type === "parameter") &&
               obj.position &&
               typeof obj.position.x === "number" &&
               typeof obj.position.y === "number"
    );

    if (positionableObjects.length === 0) return map;

    let minX = Infinity, minY = Infinity;
    positionableObjects.forEach((obj) => {
      if (obj.position.x < minX) minX = obj.position.x;
      if (obj.position.y < minY) minY = obj.position.y;
    });

    positionableObjects.forEach((obj) => {
      const px = obj.position;
      const gridCol = Math.round((px.x - minX) / PX_PER_COL);
      const gridRow = Math.round((px.y - minY) / PX_PER_ROW);
      const gridW = Math.max(2, Math.round((px.width || 180) / PX_PER_COL));
      const gridH = Math.max(3, Math.round((px.height || 60) / PX_PER_ROW));
      
      let displayName = "";
      if (obj.type === "worksheet") {
        displayName = titleMap[obj.name]
                   || (obj.title && obj.title.trim() ? obj.title.trim() : null)
                   || obj.name;
      } else {
        displayName = (obj.name || "Filter").replace(/[_-]/g, " ");
      }

      map.set(obj.name || `filter_${gridRow}_${gridCol}`, {
        type: obj.type,
        gridRow: Math.max(0, gridRow),
        gridCol: Math.max(0, gridCol),
        gridW,
        gridH,
        displayName,
        originalName: obj.name
      });
    });

    return map;
  }

  /* =============================================================================
   * TABLEAU-FORMATTED WRITERS
   * All styling below comes from the TWB format model (createSheetFormatter).
   * If no workbook was loaded, the formatter falls back to TABLEAU_DEFAULTS.
   * ============================================================================= */

  function writeTableauTitle(worksheet, r, C, text, p, span) {
    const cell = worksheet.getCell(r + 1, C + 1);
    cell.value = text;
    cell.font = tfExcelFont(p);
    cell.alignment = { horizontal: p.hAlign || "left", vertical: "middle", wrapText: false };
    if (p.bgColor) cell.fill = tfExcelFill(p.bgColor);
    if (span > 1) worksheet.mergeCells(r + 1, C + 1, r + 1, C + span);
    if ((p.fontSize || 0) > 11) worksheet.getRow(r + 1).height = Math.round(p.fontSize * 1.6);
  }

  /* Which columns get Tableau's mark colour, and what colour each row gets.
   * Tableau colours the MARKS of the pane that holds the Color encoding. */
  function buildColorPlan(fmt, colInfo, rows) {
    const enc = fmt.colorEncoding();
    const textRefs = fmt.textRefs();
    let markIdx = colInfo.map((c, i) => (!c.isHeader && (c.pivoted || (c.ref && textRefs.some(t => tfSameField(t, c.ref))))) ? i : -1).filter(i => i >= 0);
    if (!markIdx.length) markIdx = colInfo.map((c, i) => c.isHeader ? -1 : i).filter(i => i >= 0);
    if (!enc) return { enc: null, markIdx: new Set(markIdx), colorAt: () => null };

    if (enc.paneRefs && enc.paneRefs.length) {                    // scope to the colour's own pane
      const scoped = markIdx.filter(i => colInfo[i].pivoted
        ? enc.paneRefs.some(r => r.name === "Multiple Values")
        : enc.paneRefs.some(r => tfSameField(r, colInfo[i].ref)));
      if (scoped.length) markIdx = scoped;
    }
    let colorIdx = colInfo.findIndex(c => c.ref && tfSameField(c.ref, enc.ref));
    if (colorIdx < 0) {
      console.log(`[Format] Colour field ${enc.ref.inner} not present in summary data – no mark colours`);
      return { enc, markIdx: new Set(markIdx), colorAt: () => null };
    }
    const pick = dv => tfIsNull(dv) ? null
      : enc.continuous ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : tfDvText(dv);
    const vals = rows.map(row => pick(row[colorIdx]));
    const scale = tfBuildColorScale(fmt, enc, vals);
    const colors = scale ? vals.map(v => scale(v)) : [];
    return { enc, markIdx: new Set(markIdx), colorAt: i => colors[i] || null };
  }

  function borderSide(d) { return d && d.visible && d.style ? { style: d.style, color: { argb: d.color } } : undefined; }

  /* rows the writer will produce – used for layout before writing */
  function viewModelHeight(vm) {
    const grouped = FORMAT_CONFIG.groupOverflowRows && vm.rows.length > ROW_GROUP_THRESHOLD;
    return (vm.showTitle ? 1 : 0) + (vm.showHeaderRow ? 1 : 0) + vm.rows.length + (grouped ? 1 : 0);
  }

  function writeKPICardStacked(worksheet, vm, originRow, originCol, rangeTracker) {
    let r = originRow;
    const C = originCol;
    const { fmt, cols, rows, order } = vm;
    const plan = buildColorPlan(fmt, cols, rows);
    if (vm.showTitle) {
      writeTableauTitle(worksheet, r, C, vm.title.text, vm.title.props, 2);
      rangeTracker.update(r, C); rangeTracker.update(r, C + 1);
      r++;
    }
    order.forEach((ci, k) => {
      const info = cols[ci];
      tfWriteCell(worksheet.getCell(r + k + 1, C + 1), { formattedValue: info.label, value: info.label }, fmt.fieldLabelStyle(info.ref, true));
      const vp = fmt.markCellStyle(info.ref);
      const color = plan.markIdx.has(ci) ? plan.colorAt(0) : null;
      const extra = {};
      if (color && plan.enc.applyTo === "fill") extra.fill = color;
      if (color && plan.enc.applyTo === "font" && !vp.explicitColor) extra.fontColor = color;
      tfWriteCell(worksheet.getCell(r + k + 1, C + 2), rows[0][ci], vp, extra);
      rangeTracker.update(r + k, C); rangeTracker.update(r + k, C + 1);
    });
    return (vm.showTitle ? 1 : 0) + order.length;
  }

  function writeRegularTable(worksheet, vm, originRow, originCol, rangeTracker, allTablesInfo, colWidths, exactWidths) {
    let r = originRow;
    const C = originCol;
    const { fmt, cols, rows, order } = vm;
    const numCols = order.length;
    if (!fmt.hasModel) console.log(`[Format] No TWB format info for "${vm.title.text}" – using Tableau defaults`);

    const plan = buildColorPlan(fmt, cols, rows);
    const tableBg = fmt.tableBackground(vm.dashboardName) || null;   // Format → Shading → Worksheet
    const rowDiv = fmt.divider("rows");
    const colDiv = fmt.divider("cols");
    const band = fmt.banding();
    const headerIdx = vm.headerOrder.length ? vm.headerOrder
                    : cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);

    // ── Title (respects the dashboard zone's "Show title") ──
    if (vm.showTitle) {
      const tp = vm.title.props.bgColor ? vm.title.props : { ...vm.title.props, bgColor: tableBg };
      writeTableauTitle(worksheet, r, C, vm.title.text, tp, numCols);
      rangeTracker.update(r, C);
      rangeTracker.update(r, C + numCols - 1);
      r++;
    }

    // ── Header row: Measure Names aliases / captions; hidden field labels stay blank ──
    const styles = {};
    order.forEach(ci => { styles[ci] = cols[ci].isHeader ? fmt.headerCellStyle(cols[ci].ref) : fmt.markCellStyle(cols[ci].ref); });
    // ── Column widths: Tableau pixel width → text-box header width → content estimate ──
    const widthOf = {};
    order.forEach((ci, k) => {
      const info = cols[ci];
      const excelCol = C + k;
      const measureNamesRef = info.pivoted ? tfParseFieldRef("[:Measure Names]") : null;
      const px = fmt.widthPx(info.ref) || (measureNamesRef && fmt.widthPx(measureNamesRef)) || info.zoneWidthPx;
      if (px) {
        widthOf[ci] = Math.round(px / 7 * 10) / 10;
        exactWidths[excelCol] = Math.max(exactWidths[excelCol] || 0, widthOf[ci]);
        return;
      }
      const scale = (styles[ci].fontSize || 9) / 11;
      let len = Math.min(info.label.length, 24);
      for (let i = 0; i < Math.min(rows.length, 200); i++) {
        const longestLine = tfDvText(rows[i][ci]).split("\n").reduce((m, l) => Math.max(m, l.length), 0);
        len = Math.max(len, longestLine);
      }
      widthOf[ci] = Math.min(50, Math.ceil(len * scale * 1.15) + 2);
      colWidths[excelCol] = Math.max(colWidths[excelCol] || 0, widthOf[ci]);
    });

    const numericCol = ci => { const d = rows.find(rw => !tfIsNull(rw[ci])); return !!d && typeof (d[ci].nativeValue !== undefined ? d[ci].nativeValue : d[ci].value) === "number"; };

    const headerRow = r;
    if (vm.showHeaderRow) {
      order.forEach((ci, k) => {
        const info = cols[ci];
        let p = fmt.fieldLabelStyle(info.ref, !info.isHeader);
        if (info.labelProps) p = tfMerge(p, { bold: false }, info.labelProps);      // dashboard text box font
        // header label sits over its column: same alignment as the values (numbers right, text left)
        if (!p.hAlign) p.hAlign = info.isHeader ? styles[ci].hAlign : (styles[ci].hAlign || (numericCol(ci) ? "right" : "left"));
        if (/\n/.test(info.label) || info.label.length > 18) p.wrap = true;
        tfWriteCell(worksheet.getCell(r + 1, C + k + 1), { formattedValue: info.label, value: info.label }, p, {
          fill: p.bgColor || tableBg || undefined,
          border: { bottom: borderSide(rowDiv), right: k < numCols - 1 ? borderSide(colDiv) : undefined }
        });
        rangeTracker.update(r, C + k);
      });
      const hpx = fmt.headerRowHeightPx();                    // only a height Tableau stores explicitly
      if (hpx) worksheet.getRow(r + 1).height = Math.round(hpx * 0.75);
      r++;
    }

    const dataStartRow = r;
    const totalRows = rows.length;
    const needsGrouping = FORMAT_CONFIG.groupOverflowRows && totalRows > ROW_GROUP_THRESHOLD;
    const level = Math.max(1, rowDiv.level || 1);
    const rowHeightPx = fmt.rowHeightPx();

    rows.forEach((row, rowIdx) => {
      const next = rows[rowIdx + 1];
      const dividerHere = !next || !headerIdx.length || headerIdx.slice(0, level).some(hi =>
        tfDvText(row[hi]) !== tfDvText(next[hi]));
      const banded = rowIdx % 2 === 1;                 // banding per VISUAL row (after pivot/merge)
      const rowColor = plan.colorAt(rowIdx);

      order.forEach((ci, k) => {
        let p = styles[ci];
        const extra = {
          border: {
            bottom: dividerHere ? borderSide(rowDiv) : undefined,
            right: k < numCols - 1 ? borderSide(colDiv) : undefined
          }
        };
        // fill precedence: mark colour > own shading > banding > worksheet background
        const bandFill = banded ? (cols[ci].isHeader ? band.header : band.pane) : null;
        extra.fill = p.bgColor || bandFill || tableBg || undefined;
        if (rowColor && plan.markIdx.has(ci)) {
          if (plan.enc.applyTo === "fill") extra.fill = rowColor;
          else if (!p.explicitColor) extra.fontColor = rowColor;
        }
        const dv = row[ci];
        const text = tfDvText(dv);
        if (cols[ci].link) {                                      // URL action → clickable cell
          const url = cols[ci].link.expression.replace(/<([^<>]+)>/g, (m, inner) => {
            const part = cols[ci].link.parts.find(x => ("[" + x.token.inner + "]") === inner || x.token.raw === inner);
            return part && part.ci >= 0 ? tfDvText(row[part.ci]) : "";
          }).trim();
          if (/^(https?:|mailto:|ftp:)/i.test(url)) {
            extra.hyperlink = url;
            extra.hyperlinkText = FORMAT_CONFIG.linkText === "url" ? url : FORMAT_CONFIG.linkText;
            p = { ...p, underline: true, color: p.explicitColor ? p.color : "FF0563C1" };
          }
        }
        const pp = /\n/.test(text) ? { ...p, wrap: true } : p;    // wrap only on real line breaks
        tfWriteCell(worksheet.getCell(r + 1, C + k + 1), dv, pp, extra);
        rangeTracker.update(r, C + k);
      });

      const excelRow = worksheet.getRow(r + 1);
      if (rowHeightPx) excelRow.height = Math.round(rowHeightPx * 0.75);   // only a height Tableau stores
      if (needsGrouping && rowIdx >= ROW_GROUP_THRESHOLD) {
        excelRow.outlineLevel = 1;
        excelRow.hidden = true;
      }
      r++;
    });

    if (needsGrouping) {
      const hiddenCount = totalRows - ROW_GROUP_THRESHOLD;
      const noteCell = worksheet.getCell(r + 1, C + 1);
      noteCell.value = `${hiddenCount} rows are hidden — use the row group controls [+] / [-] on the left to expand or collapse`;
      noteCell.font = { italic: true, size: 9, color: { argb: "FF888888" } };
      noteCell.alignment = { horizontal: "left", vertical: "middle" };
      rangeTracker.update(r, C);
      r++;
      worksheet.properties.outlineProperties = { summaryBelow: false, summaryRight: false };
    }

    if (!plan.enc && FORMAT_CONFIG.applyFallbackHeatmap) {
      applyConditionalFormattingToTable(worksheet, rows.map(rw => order.map(ci => rw[ci])),
        order.map(ci => ({ fieldName: cols[ci].name })), dataStartRow, C, {});
    }

    if (allTablesInfo) {
      allTablesInfo.push({ name: vm.title.text, headerRow: vm.showHeaderRow ? headerRow : dataStartRow,
        leftCol: C, rightCol: C + numCols - 1,
        bottomRow: dataStartRow + totalRows - 1,             // last DATA row – the grouping note stays outside the filter
        rowCount: totalRows, hasHeader: vm.showHeaderRow });
    }
    return r - originRow;
  }

  function resolveCollisions(zones) {
  if (!zones || zones.length === 0) return zones;

  // Full row/header/note height a zone actually occupies on the sheet,
  // including its own hidden-but-physically-present grouped rows.
  function getVisualHeight(zone) {
    return (zone.allocatedRows || zone.rowCount || 5) + ROW_GAP;
  }

  // ── Group by EXACT gridRow (not a rounded bucket) ──
  const rowGroups = new Map();
  zones.forEach(zone => {
    const key = zone.gridRow;
    if (!rowGroups.has(key)) rowGroups.set(key, []);
    rowGroups.get(key).push(zone);
  });

  // ── Horizontal packing within each group ──
  const processedGroups = [];
  for (const [, group] of rowGroups) {
    group.sort((a, b) => a.gridCol - b.gridCol);
    const minRow = Math.min(...group.map(z => z.gridRow));
    group.forEach(z => { z.gridRow = minRow; });
    let cursor = group[0].gridCol;
    for (const item of group) {
      if (item.gridCol < cursor) item.gridCol = cursor;
      cursor = item.gridCol + item.gridW + COL_GAP;
    }
    const groupBottom = minRow + Math.max(...group.map(z => getVisualHeight(z))) + ROW_GAP;
    processedGroups.push({ minRow, items: group, bottom: groupBottom });
  }

  // ── Cascade vertical pushes until stable (not just one pass) ──
  processedGroups.sort((a, b) => a.minRow - b.minRow);
  let changed = true;
  while (changed) {
    changed = false;
    for (let i = 0; i < processedGroups.length - 1; i++) {
      const upper = processedGroups[i];
      const lower = processedGroups[i + 1];
      let horizontalOverlap = false;
      for (const u of upper.items) {
        for (const l of lower.items) {
          const uLeft = u.gridCol, uRight = u.gridCol + u.gridW;
          const lLeft = l.gridCol, lRight = l.gridCol + l.gridW;
          if (uLeft < lRight && uRight > lLeft) { horizontalOverlap = true; break; }
        }
        if (horizontalOverlap) break;
      }
      if (horizontalOverlap && lower.minRow < upper.bottom) {
        const pushBy = upper.bottom - lower.minRow;
        lower.items.forEach(z => { z.gridRow += pushBy; });
        lower.minRow += pushBy;
        lower.bottom += pushBy;
        changed = true;
      }
    }
  }

  // ── Full pairwise safety-net scan across ALL zones, not just adjacent groups ──
  let fullPassChanged = true;
  const MAX_PASSES = 20;
  let pass = 0;
  while (fullPassChanged && pass < MAX_PASSES) {
    fullPassChanged = false;
    pass++;
    for (let i = 0; i < zones.length; i++) {
      for (let j = i + 1; j < zones.length; j++) {
        const a = zones[i], b = zones[j];
        const upper = a.gridRow <= b.gridRow ? a : b;
        const lower = a.gridRow <= b.gridRow ? b : a;
        const upperBottom = upper.gridRow + getVisualHeight(upper) + ROW_GAP;
        const colOverlap = upper.gridCol < lower.gridCol + lower.gridW &&
                            upper.gridCol + upper.gridW > lower.gridCol;
        const rowOverlap = lower.gridRow < upperBottom;
        if (colOverlap && rowOverlap) {
          lower.gridRow += (upperBottom - lower.gridRow);
          fullPassChanged = true;
        }
      }
    }
  }

  return zones;
}

  function applyAutoFilters(worksheet, allTablesInfo) {
    if (FORMAT_CONFIG.autoFilter === "none" || !allTablesInfo || !allTablesInfo.length) return;
    // Excel allows a single autofilter per sheet → put it on the largest table that has a header row
    const t = allTablesInfo.filter(x => x.hasHeader).sort((a, b) => b.rowCount - a.rowCount)[0];
    if (!t) return;
    worksheet.autoFilter = `${getExcelColumnName(t.leftCol)}${t.headerRow + 1}:${getExcelColumnName(t.rightCol)}${t.bottomRow + 1}`;
  }

  function setColumnWidths(worksheet, colWidths, exactWidths = {}) {
    const keys = [...Object.keys(colWidths), ...Object.keys(exactWidths)].map(Number);
    const maxColIdx = keys.length ? Math.max(...keys) : 0;
    for (let ci = 0; ci <= maxColIdx; ci++) {
      if (exactWidths[ci]) { worksheet.getColumn(ci + 1).width = exactWidths[ci]; continue; }
      worksheet.getColumn(ci + 1).width = Math.min((colWidths[ci] || 10), 50);
    }
  }


  /* ── Main Export Logic ─────────────────────────────────────────────────── */
  document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const sheets = dashboard.worksheets;

      const loadBtn = document.getElementById("load_workbook_btn");
      if (loadBtn) {
        loadBtn.addEventListener("click", async () => {
          loadBtn.disabled = true;
          loadBtn.textContent = "⏳ Loading...";
          await loadWorkbookFile();
          loadBtn.disabled = false;
          loadBtn.textContent = "📁 Load Workbook";
        });
      }

      const fileLabel = document.getElementById("twb_file_label");
      if (fileLabel) {
        const savedFileName = tableau.extensions.settings.get("twbFileName");
        const savedTitleMap = tableau.extensions.settings.get("twbTitleMap");
        const savedFormat = tableau.extensions.settings.get("twbFormatModel");
        if (savedFileName && savedTitleMap) {
          const titleCount = Object.keys(JSON.parse(savedTitleMap)).length;
          const sheetCount = savedFormat ? Object.keys(JSON.parse(savedFormat).sheets || {}).length : 0;
          fileLabel.textContent = `✅ ${savedFileName} — ${titleCount} titles, formatting for ${sheetCount} sheets`;
        } else {
          fileLabel.textContent = "No workbook loaded — click 📁 to load titles and colors";
        }
      }

      const exportBtn = document.getElementById("export_button");
      if (exportBtn) {
        exportBtn.addEventListener("click", exportToExcel);
      }

      async function exportToExcel() {
  const btn = document.getElementById("export_button");
  btn.disabled = true;

  try {
    const filterValuesMap = await extractFilterValuesPerField(sheets);
    console.log("📊 Filter values per field:", filterValuesMap);

    const titleMap = getTitleMap();
    console.log(`[Export] Using ${Object.keys(titleMap).length} titles`);

    const fmtModel = getFormatModel();
    console.log(`[Export] Format model: ${fmtModel ? Object.keys(fmtModel.sheets).length + " sheets" : "none – load the workbook for exact formatting"}`);

    const layoutMap = buildLayoutMap(dashboard.objects || [], titleMap);

    // ── 1. Build DZV map ──
    const dzvMap = {};
    (dashboard.objects || [])
      .filter(obj => obj.type === "worksheet")
      .forEach(obj => {
        dzvMap[obj.name] = obj.isVisible;
      });
    console.log("[DZV] Visibility map:", dzvMap);

    // ── 2. Fetch all sheets in parallel ──
    const allSheetsData = await fetchAllSheetsData(sheets);

    const filterValueItems = [];
    const dataWorksheetItems = [];
    const visualStatuses = [];

    for (const { sheet, data, visualSpec, visualSpecError, error } of allSheetsData) {
      if (error) {
        console.warn(`Skipping "${sheet.name}": ${error.message}`);
        continue;
      }

      // ---- DZV check ----
      if (dzvMap[sheet.name] === false) {
        console.log(`[DZV] Skipping hidden sheet: "${sheet.name}"`);
        continue;
      }

      let summaryData = data;  // already fetched
      if (!summaryData.columns || summaryData.columns.length === 0) {
        console.warn(`[Export] Skipping "${sheet.name}": summary data has no columns`);
        continue;
      }
      if (!summaryData.data || summaryData.data.length === 0) {
        console.warn(`[Export] Skipping "${sheet.name}": summary data has no rows`);
        continue;
      }

      const layout = layoutMap.get(sheet.name);
      const visualName = (layout && layout.displayName) ? layout.displayName : sheet.name;

      if (isFilterValueWorksheet(sheet.name, summaryData)) {
        let matchedFilterName = null;
        let matchedValues = null;

        for (const [filterField, values] of Object.entries(filterValuesMap)) {
          const cleanFilterField = filterField.toLowerCase().replace(/[^a-z]/g, '');
          const cleanSheetName = sheet.name.toLowerCase().replace(/[^a-z]/g, '');
          if (cleanSheetName.includes(cleanFilterField) || cleanFilterField.includes(cleanSheetName)) {
            matchedFilterName = filterField;
            matchedValues = values;
            break;
          }
        }

        if (!matchedValues) {
          matchedFilterName = visualName;
          matchedValues = [`Values: ${summaryData.data.length} items`];
        }

        filterValueItems.push({
          type: "filterValue",
          name: sheet.name,
          visualName: visualName,
          filterName: matchedFilterName,
          values: matchedValues,
          layout: layout,
          rowCount: 2 + (Array.isArray(matchedValues) ? matchedValues.length : 1),
          originalData: summaryData
        });
      } else {
        // ── NEW: rebuild the visual table (pivot, merge, sort, visible columns, title) ──
        const visualModel = buildVisualModel(fmtModel, sheet.name, summaryData,
          { dashboardName: dashboard.name, displayName: visualName, visualSpec, visualSpecError, layout });
        console.log(`[VisualSpec] ${sheet.name}:`, {
          available: visualModel.metadata.visualSpecAvailable,
          keys: visualModel.metadata.visualSpecKeys,
          type: visualModel.type,
          markClass: visualModel.metadata.markClass,
          error: visualModel.metadata.visualSpecError
        });
        const renderDecision = chooseVisualRenderer(visualModel);
        visualModel.renderer = renderDecision.renderer;
        visualModel.status = renderDecision.status;
        visualModel.statusReason = renderDecision.reason;
        visualStatuses.push({
          worksheet: sheet.name,
          type: visualModel.type,
          renderer: renderDecision.renderer,
          status: renderDecision.status,
          reason: renderDecision.reason,
          visualSpecAvailable: visualModel.metadata.visualSpecAvailable,
          imageApiAvailable: visualModel.metadata.imageApiAvailable
        });
        console.log(`[Visual] ${sheet.name}: ${visualModel.type} -> ${renderDecision.renderer} (${renderDecision.reason})`);

        if (visualModel.viewModel.kind === "chart" && FORMAT_CONFIG.chartPolicy === "skip") {
          console.warn(`[Export] "${sheet.name}" is a chart - skipped by FORMAT_CONFIG.chartPolicy`);
          continue;
        }
        let vm = visualModel.viewModel;
        if (!vm.rows.length || !vm.order.length) {
          console.warn(`[Export] "${sheet.name}" produced an empty formatted view; rebuilding from live summary data`);
          const rawVisualModel = buildVisualModel(null, sheet.name, summaryData,
            { dashboardName: dashboard.name, displayName: visualName, visualSpec, visualSpecError, layout });
          if (rawVisualModel.viewModel.rows.length && rawVisualModel.viewModel.order.length) {
            vm = rawVisualModel.viewModel;
            visualModel.viewModel = vm;
            visualModel.data.rows = vm.rows;
            visualModel.data.columns = vm.cols;
            visualModel.data.dimensions = vm.order.filter(i => vm.cols[i].isHeader).map(i => vm.cols[i]);
            visualModel.data.measures = vm.order.filter(i => !vm.cols[i].isHeader).map(i => vm.cols[i]);
            visualModel.statusReason = "TWB view was empty; exported live summary data";
          }
        }
        if (!vm.rows.length || !vm.order.length) {
          console.warn(`[Export] Skipping "${sheet.name}": no renderable rows or columns after fallback`);
          continue;
        }
        const isKPI = isKPIViewModel(vm);
        dataWorksheetItems.push({
          name: sheet.name,
          visualName: vm.title.text,
          layout,
          isKPI,
          type: "worksheet",
          vm,
          visualModel,
          columns: vm.order,                                   // only its length is used for layout
          rowCount: isKPI ? (vm.showTitle ? 1 : 0) + vm.order.length : viewModelHeight(vm)
        });
      }
    }

    const allItems = [...filterValueItems, ...dataWorksheetItems];

    if (allItems.length === 0) {
      throw new Error("No data found in any visible worksheet.");
    }

    console.table(visualStatuses);
    updateVisualStatus(visualStatuses);

    const placedItems = allItems.map((item, idx) => {
      if (item.layout) {
        const l = item.layout;
        let calculatedWidth = l.gridW;

        if (item.type === "filterValue") {
          calculatedWidth = Math.max(l.gridW, 3);
        } else if (item.isKPI) {
          calculatedWidth = Math.max(l.gridW, Math.min(6, item.columns.length + 1));
        } else {
          calculatedWidth = Math.max(l.gridW, Math.min(25, item.columns.length + 1));
        }

        return {
          ...item,
          gridRow: l.gridRow,
          gridCol: l.gridCol,
          gridW: calculatedWidth,
          allocatedRows: item.rowCount,
        };
      } else {
        const itemsPerRow = 3;
        const rowIdx = Math.floor(idx / itemsPerRow);
        const colIdx = idx % itemsPerRow;
        const gridRow = rowIdx * (Math.max(8, item.rowCount) + ROW_GAP);
        const gridCol = colIdx * 12;
        const gridW = (item.type === "filterValue") ? 4 : (item.isKPI ? 5 : 12);

        return {
          ...item,
          gridRow: gridRow,
          gridCol: gridCol,
          gridW: gridW,
          allocatedRows: item.rowCount,
        };
      }
    });

    resolveCollisions(placedItems);

// SNAP-OUT-OF-GROUPED-ROWS PASS
// Excel hides entire physical rows, not per-column cells. resolveCollisions
// only pushes items that horizontally overlap another item's column range —
// so a table in column K can still be placed on a row that a totally
// different table (in column A) has marked hidden=true for its own
// row-grouping. When the workbook opens with that group collapsed, the
// column-K table vanishes too, even though nothing "collided" by column.
// This pass detects that and pushes the item below the other table's full
// physical row range (not just its visible rows).
let snapChanged = true;
const MAX_SNAP_PASSES = 10;
let snapPass = 0;

while (FORMAT_CONFIG.groupOverflowRows && snapChanged && snapPass < MAX_SNAP_PASSES) {
  snapChanged = false;
  snapPass++;

  for (const item of placedItems) {
    for (const other of placedItems) {
      if (item === other) continue;

      const otherTotal = other.allocatedRows || other.rowCount || 0;
      if (otherTotal <= ROW_GROUP_THRESHOLD + 2) continue; // other has no hidden rows

      const otherVisibleEnd = other.gridRow + ROW_GROUP_THRESHOLD + 3;
      const otherPhysicalEnd = other.gridRow + otherTotal + ROW_GAP;

      if (item.gridRow >= otherVisibleEnd && item.gridRow < otherPhysicalEnd) {
        item.gridRow = otherPhysicalEnd;
        snapChanged = true;
      }
    }
  }
}

    const workbook = new ExcelJS.Workbook();
    const sheetName = (dashboard.name || "Dashboard Export")
      .replace(/[\\\/\*\?\[\]:]/g, "")
      .slice(0, 31);
    const worksheet = workbook.addWorksheet(sheetName);

    const colWidths = {};
    const exactWidths = {};
    const allTablesInfo = [];
    const tracker = makeRangeTracker();

    let currentRow = 0;
    const dashboardName = dashboard.name || "Dashboard Export";
    const exportDate = new Date().toLocaleString();
    let dashTitleProps = null;
    const dashFmt = fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name];
    if (dashFmt) {
      const run = dashFmt.title && (dashFmt.title.runs.find(x => x.text.trim()) || {}).props;
      dashTitleProps = tfMerge(TABLEAU_DEFAULTS.dashTitle,
        tfCollect(fmtModel.workbookStyle, TF_ELEMENTS.dashTitle), tfCollect(dashFmt.style, TF_ELEMENTS.dashTitle), run);
    }
    const usedZones = new Set(dataWorksheetItems.flatMap(it => it.vm ? it.vm.headerZoneIds : []));
    const titleRuns = tfDashboardTitleRuns(fmtModel, dashboard.name, usedZones);
    const titleHeight = writeDashboardTitle(worksheet, dashboardName, exportDate, currentRow, 0, tracker, dashTitleProps, titleRuns);
    currentRow += titleHeight;


    const adjustedItems = placedItems.map(item => ({
      ...item,
      gridRow: item.gridRow + currentRow
    }));

    for (let i = 0; i < adjustedItems.length; i++) {
      const item = adjustedItems[i];

      if (item.visualModel && item.visualModel.renderer === "tableau-image") {
        const imageWidth = Math.max(320, item.gridW * PX_PER_COL);
        const imageHeight = Math.max(160, (item.allocatedRows || item.rowCount || 8) * PX_PER_ROW);
        const status = visualStatuses.find(entry => entry.worksheet === item.name);
        try {
          const imageBase64 = await renderTableauImage(item.visualModel, imageWidth, imageHeight);
          const imageId = workbook.addImage({ base64: imageBase64, extension: "png" });
          worksheet.addImage(imageId, {
            tl: { col: item.gridCol, row: item.gridRow },
            ext: { width: imageWidth, height: imageHeight }
          });
          tracker.update(item.gridRow, item.gridCol);
          tracker.update(item.gridRow + Math.ceil(imageHeight / PX_PER_ROW), item.gridCol + Math.ceil(imageWidth / PX_PER_COL));
          item.visualModel.status = "success";
          item.visualModel.statusReason = "Tableau SVG rendered and embedded as PNG";
          if (status) {
            status.status = "success";
            status.reason = item.visualModel.statusReason;
          }
          continue;
        } catch (err) {
          item.visualModel.renderer = "data-fallback";
          item.visualModel.status = "warning";
          item.visualModel.statusReason = `image renderer failed: ${err.message}`;
          if (status) {
            status.renderer = "data-fallback";
            status.status = "warning";
            status.reason = item.visualModel.statusReason;
          }
          console.warn(`[Visual] ${item.name}: image renderer failed; using data fallback`, err);
        }
      }

      if (item.type === "filterValue") {
        writeIndividualFilterTable(
          worksheet,
          item.filterName,
          item.values,
          item.gridRow,
          item.gridCol,
          tracker
        );
        colWidths[item.gridCol] = Math.max(colWidths[item.gridCol] || 0, 35);
      } else if (item.isKPI) {
        writeKPICardStacked(worksheet, item.vm, item.gridRow, item.gridCol, tracker);
        colWidths[item.gridCol] = Math.max(colWidths[item.gridCol] || 0, 22);
        colWidths[item.gridCol + 1] = Math.max(colWidths[item.gridCol + 1] || 0, 18);
      } else {
        writeRegularTable(worksheet, item.vm, item.gridRow, item.gridCol, tracker, allTablesInfo, colWidths, exactWidths);
      }
    }

    updateVisualStatus(visualStatuses);

    setColumnWidths(worksheet, colWidths, exactWidths);
    applyAutoFilters(worksheet, allTablesInfo);

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
    const fname = `${sheetName}_${stamp}.xlsx`;

    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fname;
    link.click();
    URL.revokeObjectURL(link.href);

    console.log("✅ Export completed with Tableau formatting (fonts, colours, number formats, borders)");

  } catch (err) {
    console.error("[Export]", err);
    alert("Export failed. Check console (F12) for details.\n\n" + err.message);
  } finally {
    btn.disabled = false;
  }
}

    }).catch((err) => {
      console.error("[Tableau init]", err);
      alert("Failed to connect to Tableau: " + err.message);
      if (document.getElementById("export_button")) {
        document.getElementById("export_button").disabled = true;
      }
    });
  });

})();