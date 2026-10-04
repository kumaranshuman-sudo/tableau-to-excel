/* Imports the extension modules (src/testing.js) and runs every worksheet of the
 * sample workbook (plus synthetic waterfall / box plot sheets) through
 * classify → renderer → chart spec / image spec, on synthetic Superstore data.
 *   node tests/harness.js                  → table + check against EXPECTED (exit 1 on mismatch)
 *   node tests/harness.js xlsx <file>      → also writes an XLSX with every native chart
 *   node tests/harness.js json "<sheet>" [image] → dump one sheet's chart / image spec
 * tests/run.ps1 runs the check, builds the XLSX and opens it in Excel. */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import shims from "./shims.js";

const { DOMParser, writeZip } = shims;
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");

/* sheet → [visual type, renderer, native chart kind or image mark] */
const EXPECTED = {
  "KPI - Total Sales": ["KPI", "cell", ""],
  "Text Table": ["TABLE", "cell", ""],
  "Highlight Table": ["HEATMAP", "cell", ""],
  "Heat Map": ["HEATMAP", "cell", ""],
  "Vertical Bar": ["BAR", "excel-chart", "bar"],
  "Horizontal Bar (sorted)": ["BAR", "excel-chart", "bar"],
  "Stacked Bar": ["BAR", "excel-chart", "bar"],
  "Side-by-Side Bar": ["BAR", "excel-chart", "bar"],
  "Line (Continuous)": ["LINE", "excel-chart", "line"],
  "Line (Discrete, by Category)": ["LINE", "excel-chart", "line"],
  "Area Chart": ["AREA", "excel-chart", "area"],
  "Pie Chart": ["PIE", "excel-chart", "pie"],
  "Scatter Plot": ["SCATTER", "excel-chart", "scatter"],
  "Circle View": ["LINE", "excel-chart", "line"],
  "Dual Combination": ["COMBO", "excel-chart", "combo"],
  "Measure Values Bar": ["BAR", "excel-chart", "bar"],
  "Histogram": ["HISTOGRAM", "excel-chart", "bar"],
  "Gantt Chart": ["GANTT", "excel-chart", "bar"],
  "Treemap": ["TREEMAP", "excel-chart", "treemap"],
  "Packed Bubbles": ["BUBBLE", "tableau-image", "img:circle"],
  "Symbol Map": ["MAP", "tableau-image", "img:circle"],
  "Filled Map": ["MAP_FILLED", "data-fallback", ""],
  "Waterfall": ["WATERFALL", "excel-chart", "bar"],
  "Box Plot": ["BOXPLOT", "excel-chart", "line"],
  "Tree Map (Automatic)": ["TREEMAP", "excel-chart", "treemap"]
};
// the modules read browser globals at call time; the parser needs a DOMParser
globalThis.DOMParser = DOMParser;
globalThis.tableau = { extensions: { createVizImageAsync() {} } };
// keep the extension's own logging out of the test output
const print = console.log.bind(console);
console.log = console.warn = console.info = () => {};
const X = await import(pathToFileURL(path.join(ROOT, "src", "testing.js")).href);
const model = X.parseTableauFormatting(fs.readFileSync(path.join(ROOT, "test_workbooks", "Visual Gallery.twb.xml"), "utf8"));

// sheets the sample lacks: waterfall (Gantt + running total) and box plot (Circle View + Analytics box plot)
{
  const clone = o => JSON.parse(JSON.stringify(o));
  const wf = clone(model.sheets["Vertical Bar"]);
  wf.name = "Waterfall";
  wf.cols = wf.cols.map(r => ({ ...r, name: "Sub-Category", inner: "none:Sub-Category:nk", raw: r.raw.replace("Category", "Sub-Category") }));
  wf.rows = wf.rows.map(r => ({ ...r, name: "Profit", inner: "sum:Profit:qk", raw: r.raw.replace("Sales", "Profit") }));
  wf.fieldRefs = [...wf.rows, ...wf.cols];
  wf.runningTotals = clone(wf.rows);
  wf.panes.forEach(p => { p.markClass = "GanttBar"; });
  model.sheets["Waterfall"] = wf;
  // treemap built with Tableau's Automatic mark (empty shelves + Size): Tableau draws squares
  const tma = clone(model.sheets["Treemap"]);
  tma.name = "Tree Map (Automatic)";
  tma.panes.forEach(p => { p.markClass = "Automatic"; });
  model.sheets["Tree Map (Automatic)"] = tma;
  const bp = clone(model.sheets["Circle View"]);
  bp.name = "Box Plot";
  bp.boxPlot = true;
  model.sheets["Box Plot"] = bp;
}

/* ── synthetic Superstore ────────────────────────────────────────────────── */
let seed = 7;
const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
const pick = a => a[Math.floor(rnd() * a.length)];
const SUB = { Furniture: ["Bookcases", "Chairs", "Furnishings", "Tables"],
  "Office Supplies": ["Appliances", "Art", "Binders", "Envelopes", "Fasteners", "Labels", "Paper", "Storage", "Supplies"],
  Technology: ["Accessories", "Copiers", "Machines", "Phones"] };
const STATES = [["California", 36.1, -119.6], ["New York", 42.1, -74.9], ["Texas", 31.0, -97.5], ["Washington", 47.4, -121.4],
  ["Florida", 27.7, -81.6], ["Illinois", 40.3, -89.0], ["Ohio", 40.3, -82.7], ["Georgia", 33.0, -83.6]];
const recs = [];
for (let i = 0; i < 2500; i++) {
  const cat = pick(Object.keys(SUB));
  const st = pick(STATES);
  const d = new Date(Date.UTC(2023 + Math.floor(rnd() * 4), Math.floor(rnd() * 12), 1 + Math.floor(rnd() * 28)));
  const sales = Math.round((20 + rnd() * 900) * (cat === "Technology" ? 1.6 : 1) * 100) / 100;
  recs.push({ Region: pick(["Central", "East", "South", "West"]), Category: cat, "Sub-Category": pick(SUB[cat]),
    Segment: pick(["Consumer", "Corporate", "Home Office"]), "Ship Mode": pick(["First Class", "Same Day", "Second Class", "Standard Class"]),
    date: d, Sales: sales, Profit: Math.round(sales * (rnd() * 0.6 - 0.15) * 100) / 100, Quantity: 1 + Math.floor(rnd() * 14),
    "Order ID": "CA-" + i, "Customer Name": "Customer " + Math.floor(rnd() * 120), State: st[0], lat: st[1], lon: st[2],
    days: Math.floor(rnd() * 7) });
}
const fmtNum = n => n.toLocaleString("en-US", { maximumFractionDigits: 0 });
const dims = {
  Region: { t: "string", f: r => r.Region }, Category: { t: "string", f: r => r.Category },
  "Sub-Category": { t: "string", f: r => r["Sub-Category"] }, Segment: { t: "string", f: r => r.Segment },
  "Ship Mode": { t: "string", f: r => r["Ship Mode"] }, "Customer Name": { t: "string", f: r => r["Customer Name"] },
  "State/Province": { t: "string", f: r => r.State }, "Country/Region": { t: "string", f: () => "United States" },
  "MONTH(Order Date)": { t: "date", f: r => new Date(Date.UTC(r.date.getUTCFullYear(), r.date.getUTCMonth(), 1)),
                         fmt: v => v.toLocaleString("en-US", { month: "long", year: "numeric", timeZone: "UTC" }) },
  "YEAR(Order Date)": { t: "int", f: r => r.date.getUTCFullYear(), fmt: String },
  "QUARTER(Order Date)": { t: "string", f: r => "Q" + (Math.floor(r.date.getUTCMonth() / 3) + 1) },
  Quantity: { t: "int", f: r => r.Quantity, fmt: String },
  "Latitude (generated)": { t: "float", f: r => r.lat, fmt: String }, "Longitude (generated)": { t: "float", f: r => r.lon, fmt: String }
};
const meas = {
  "SUM(Sales)": { agg: g => g.reduce((s, r) => s + r.Sales, 0), fmt: v => "$" + fmtNum(v) },
  "SUM(Profit)": { agg: g => g.reduce((s, r) => s + r.Profit, 0), fmt: v => "$" + fmtNum(v) },
  "CNT(Order ID)": { agg: g => g.length, fmt: fmtNum, t: "int" },
  "AVG(Days to Ship Actual)": { agg: g => g.reduce((s, r) => s + r.days, 0) / g.length, fmt: v => v.toFixed(1) },
  "AGG(Profit Ratio)": { agg: g => g.reduce((s, r) => s + r.Profit, 0) / g.reduce((s, r) => s + r.Sales, 0), fmt: v => (v * 100).toFixed(1) + "%" }
};
const dv = (v, f) => ({ value: v, nativeValue: v, formattedValue: f ? f(v) : String(v) });
function summary(dimNames, measNames, measureValues) {
  const groups = new Map();
  recs.forEach(r => {
    const key = dimNames.map(d => String(+dims[d].f(r) || dims[d].f(r))).join("|");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  });
  const columns = dimNames.map((d, i) => ({ fieldName: d, dataType: dims[d].t, index: i }));
  const data = [];
  for (const g of groups.values()) {
    const keyRow = dimNames.map(d => dv(dims[d].f(g[0]), dims[d].fmt));
    if (measureValues) measNames.forEach(m => data.push([...keyRow, dv(m), dv(meas[m].agg(g), meas[m].fmt)]));
    else data.push([...keyRow, ...measNames.map(m => dv(meas[m].agg(g), meas[m].fmt))]);
  }
  if (measureValues) columns.push({ fieldName: "Measure Names", dataType: "string" }, { fieldName: "Measure Values", dataType: "float" });
  else measNames.forEach(m => columns.push({ fieldName: m, dataType: meas[m].t || "float" }));
  return { columns, data };
}
const SHEETS = {
  "KPI - Total Sales": summary([], ["SUM(Sales)"]),
  "Text Table": summary(["Region", "Category"], ["SUM(Sales)", "SUM(Profit)"], true),
  "Highlight Table": summary(["Sub-Category", "Region"], ["SUM(Profit)"]),
  "Heat Map": summary(["Sub-Category", "Segment"], ["SUM(Profit)", "SUM(Sales)"]),
  "Vertical Bar": summary(["Category"], ["SUM(Sales)"]),
  "Horizontal Bar (sorted)": summary(["Sub-Category"], ["SUM(Sales)"]),
  "Stacked Bar": summary(["Region", "Segment"], ["SUM(Sales)"]),
  "Side-by-Side Bar": summary(["Region", "Segment"], ["SUM(Sales)"]),
  "Line (Continuous)": summary(["MONTH(Order Date)"], ["SUM(Sales)"]),
  "Line (Discrete, by Category)": summary(["YEAR(Order Date)", "QUARTER(Order Date)", "Category"], ["SUM(Profit)"]),
  "Area Chart": summary(["MONTH(Order Date)", "Category"], ["SUM(Sales)"]),
  "Pie Chart": summary(["Segment"], ["SUM(Sales)"]),
  "Scatter Plot": summary(["Region", "Customer Name"], ["SUM(Sales)", "SUM(Profit)"]),
  "Circle View": summary(["Category", "Segment", "Sub-Category"], ["SUM(Sales)"]),
  "Dual Combination": summary(["QUARTER(Order Date)"], ["SUM(Sales)", "SUM(Profit)"]),
  "Measure Values Bar": summary(["Region"], ["SUM(Sales)", "SUM(Profit)"], true),
  "Histogram": summary(["Quantity"], ["CNT(Order ID)"]),
  "Gantt Chart": summary(["Ship Mode", "MONTH(Order Date)"], ["AVG(Days to Ship Actual)"]),
  "Treemap": summary(["Sub-Category"], ["SUM(Profit)", "SUM(Sales)"]),
  "Packed Bubbles": summary(["Category", "Sub-Category"], ["SUM(Sales)"]),
  "Symbol Map": summary(["State/Province", "Country/Region", "Latitude (generated)", "Longitude (generated)"], ["SUM(Profit)", "SUM(Sales)"]),
  "Filled Map": summary(["State/Province", "Country/Region", "Latitude (generated)", "Longitude (generated)"], ["AGG(Profit Ratio)"]),
  "Waterfall": (() => {                                     // running sum of profit per sub-category
    const s = summary(["Sub-Category"], ["SUM(Profit)"]);
    let run = 0;
    s.data.sort((a, b) => a[0].value.localeCompare(b[0].value)).forEach(r => { run += r[1].value * (r[0].value < "F" ? 1 : -0.3); r[1] = dv(run, meas["SUM(Profit)"].fmt); });
    return s;
  })(),
  "Box Plot": summary(["Category", "Segment", "Sub-Category"], ["SUM(Sales)"]),
  "Tree Map (Automatic)": summary(["Sub-Category"], ["SUM(Profit)", "SUM(Sales)"])
};

/* ── run ─────────────────────────────────────────────────────────────────── */
const results = [];
for (const [name, sum] of Object.entries(SHEETS)) {
  const visual = X.buildVisualModel(model, name, sum, { dashboardName: "Visual Gallery", displayName: name });
  let decision = X.chooseVisualRenderer(visual);
  let specs = null, image = null, err = null;
  if (decision.renderer === "excel-chart") {
    try { specs = X.buildExcelChartSpecs(visual, model); }
    catch (e) { err = e.message; decision = X.imageOrFallback(visual, `native chart not possible (${e.message})`); }
  }
  if (decision.renderer === "tableau-image") {
    try { image = X.buildVizImageSpec(visual, 480, 300); } catch (e) { err = "image spec: " + e.message; }
  }
  results.push({ name, visual, decision, specs, image, err });
}

const mode = process.argv[2];
if (mode === "json") {
  const pick = process.argv[3];
  const r = results.find(x => x.name === pick);
  if (!r.image) { try { r.image = X.buildVizImageSpec(r.visual, 480, 300); } catch (e) { r.err = (r.err || "") + " | image spec: " + e.message; } }
  if (process.argv[4] === "image") print(r.err || "", JSON.stringify(r.image && { ...r.image, data: { n: r.image.data.values.length, first: r.image.data.values.slice(0, 2) } }));
  else print(JSON.stringify({ type: r.visual.type, decision: r.decision, err: r.err, image: r.image && { ...r.image, data: { values: r.image.data.values.slice(0, 3) } },
    specs: r.specs && r.specs.map(s => ({ ...s, series: s.series.map(x => ({ ...x, values: x.values && x.values.slice(0, 6), x: x.x && x.x.slice(0, 3), y: x.y && x.y.slice(0, 3) })) })) }, null, 1));
} else {
  const failures = [];
  for (const r of results) {
    const kinds = r.specs ? [...new Set(r.specs.map(s => s.kind))].join("/") + (r.specs.length > 1 ? ` x${r.specs.length}` : "") : "";
    const img = r.image ? `img:${r.image.mark}` : "";
    const got = [r.visual.type, r.decision.renderer, kinds || img];
    const want = EXPECTED[r.name];
    const ok = want && want.every((w, i) => w === got[i]);
    if (!ok) failures.push(`${r.name}: expected ${want ? want.join(" / ") : "(no expectation)"}, got ${got.join(" / ")}`);
    print(`${ok ? "ok  " : "FAIL"} ${r.name.padEnd(30)} ${got[0].padEnd(10)} ${got[1].padEnd(14)} ${got[2].padEnd(14)} ${r.err || ""}`);
  }
  const good = X.checkWorkbookMatch(model, { name: "Visual Gallery", worksheets: Object.keys(SHEETS).map(name => ({ name })) });
  const wrong = X.checkWorkbookMatch(model, { name: "Sales Overview", worksheets: [{ name: "Sales by Region" }, { name: "Vertical Bar" }, { name: "Profit Trend" }] });
  if (good.level !== "ok") failures.push(`workbook match: expected ok for the right workbook, got ${good.level}`);
  if (wrong.level !== "mismatch") failures.push(`workbook match: expected mismatch for a wrong workbook, got ${wrong.level}`);
  print(`\nworkbook match: right file → ${good.level}, wrong file → ${wrong.level}`);
  if (failures.length) {
    console.error(`\n${failures.length} check(s) failed:\n  ` + failures.join("\n  "));
    process.exitCode = 1;
  } else print(`\nall ${results.length} sheets classified and rendered as expected`);
}

/* ── XLSX with every native chart, for checking in real Excel ─────────────── */
if (mode === "xlsx") {
  const out = process.argv[3];
  const cells = new Map();
  const ws = { name: "Chart Data", getCell(r, c) {
    const k = r + "," + c;
    if (!cells.has(k)) cells.set(k, { r, c });
    return cells.get(k);
  } };
  let row = 0;
  const jobs = [];
  const only = process.env.ONLY ? process.env.ONLY.split(",") : null;
  results.filter(r => r.specs && (!only || only.includes(r.name))).forEach((r, n) => {
    r.specs.forEach((spec, k) => {
      const refs = X.ExcelChartWriter.writeChartData(ws, spec, row);
      row = refs.nextRow;
      jobs.push({ spec, refs, name: r.name + (r.specs.length > 1 ? " " + (k + 1) : ""), col: (jobs.length % 3) * 9, row: Math.floor(jobs.length / 3) * 22,
                  widthPx: 600, heightPx: 400 });
    });
  });
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const colName = i => { let s = "", n = i; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); } return s; };
  const byRow = new Map();
  for (const cell of cells.values()) { if (cell.value === undefined) continue; if (!byRow.has(cell.r)) byRow.set(cell.r, []); byRow.get(cell.r).push(cell); }
  const rowsXml = [...byRow.keys()].sort((a, b) => a - b).map(r => `<row r="${r}">` + byRow.get(r).sort((a, b) => a.c - b.c).map(c => {
    const ref = colName(c.c) + r;
    return typeof c.value === "number" ? `<c r="${ref}"><v>${c.value}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t>${esc(c.value)}</t></is></c>`;
  }).join("") + "</row>").join("");
  const sheet = d => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${d}</sheetData></worksheet>`;
  const pkg = {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Dashboard" sheetId="1" r:id="rId1"/><sheet name="Chart Data" sheetId="2" r:id="rId2"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
    "xl/worksheets/sheet1.xml": sheet(""),
    "xl/worksheets/sheet2.xml": sheet(rowsXml)
  };
  X.ExcelChartWriter.injectCharts(writeZip(pkg), { sheetIndex: 0, charts: jobs }).then(buf => {
    fs.writeFileSync(out, Buffer.from(buf));
    fs.writeFileSync(out + ".names.txt", jobs.map(j => j.name).join("\n"));
    print(`\nwrote ${out} with ${jobs.length} charts`);
  });
}
