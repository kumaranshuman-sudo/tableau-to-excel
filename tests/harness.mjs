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
  "Packed Bubbles": ["BUBBLE", "excel-chart", "bubble"],
  "Symbol Map": ["MAP", "excel-chart", "bubble"],
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

  // dashboard layout → Excel grid
  const check = (ok, what) => { if (!ok) failures.push("layout: " + what); return ok; };
  // without snapping "Right" (5 px right of Wide's edge, 8 px lower) would round to column 11 / row 1
  const objects = [
    { type: "worksheet", name: "Wide", position: { x: 0, y: 10 }, size: { width: 940, height: 400 } },
    { type: "worksheet", name: "Right", position: { x: 945, y: 18 }, size: { width: 450, height: 400 } },
    { type: "filter", name: "Wide", position: { x: 1400, y: 0 }, size: { width: 160, height: 80 } }   // filter named after its sheet
  ];
  const map = X.buildLayoutMap(objects);
  const wide = map.get("Wide"), right = map.get("Right");
  const layoutOk = [
    check(wide.widthPx === 940 && wide.gridW === 10, `zone size comes from DashboardObject.size (got ${wide.widthPx}px, ${wide.gridW} cols)`),
    check(wide.type === "worksheet", "a filter with the worksheet's name must not replace the worksheet's entry"),
    check(right.gridCol === 10 && right.gridRow === wide.gridRow, `edges a few px apart snap to one grid line (got row ${right.gridRow}, col ${right.gridCol})`),
    check(X.graphicBox(wide, { showTitle: false }).gridW === 10, "a chart block is as wide as its zone on the grid")
  ];
  // a row of blocks moves down as a whole when one of them sits under a tall block
  const zones = [
    { name: "tall", gridRow: 0, gridCol: 0, gridW: 4, allocatedRows: 20 },
    { name: "small", gridRow: 3, gridCol: 8, gridW: 4, allocatedRows: 3 },
    { name: "under tall", gridRow: 10, gridCol: 0, gridW: 4, allocatedRows: 5 },
    { name: "under small", gridRow: 10, gridCol: 8, gridW: 4, allocatedRows: 5 }
  ];
  X.resolveCollisions(zones);
  layoutOk.push(check(zones[2].gridRow === zones[3].gridRow && zones[2].gridRow >= 21,
    `blocks that share a top edge stay level (got rows ${zones[2].gridRow} / ${zones[3].gridRow})`));
  // a long table keeps the rows a chart beside it uses visible
  const table = { type: "worksheet", name: "table", gridRow: 0, gridCol: 0, gridW: 4, allocatedRows: 43,
                  vm: { rows: new Array(40).fill([]), showTitle: true, showHeaderRow: true } };
  const chart = { type: "worksheet", name: "chart", gridRow: 0, gridCol: 4, gridW: 6, allocatedRows: 20, box: {}, vm: table.vm };
  X.setTableVisibleRows([table, chart]);
  layoutOk.push(check(table.visibleRows === 18, `a table beside a chart keeps its rows visible (got ${table.visibleRows})`));
  print(`layout: ${layoutOk.filter(Boolean).length}/${layoutOk.length} checks passed`);

  // default export file name: "<workbook> - <dashboard>.xlsx"
  const names = [
    [["Superstore.twbx", "Executive Overview"], "Superstore - Executive Overview.xlsx"],
    [["Retail Dashboard (1).twb", "Store Performance"], "Retail Dashboard (1) - Store Performance.xlsx"],
    [[null, "Executive Overview"], "Executive Overview.xlsx"],                       // no workbook loaded
    [["Sales.twbx", "Sales"], "Sales.xlsx"],                                          // same name once
    [["Q1/Q2.twbx", "Profit: West?"], "Q1Q2 - Profit West.xlsx"]                      // characters Windows forbids
  ];
  const nameFails = names.filter(([args, want]) => X.exportFileName(...args) !== want);
  nameFails.forEach(([args, want]) => failures.push(`file name: ${JSON.stringify(args)} → ${X.exportFileName(...args)}, expected ${want}`));
  print(`file name: ${names.length - nameFails.length}/${names.length} checks passed`);

  // a small workbook: a KPI tile with a custom label (value over caption, green background) and a
  // trend sheet whose axes / grid lines are switched off – the export must follow both
  const twb = `<?xml version='1.0' encoding='utf-8' ?><workbook><datasources/><worksheets>
    <worksheet name='KPI'><table><view/><style><style-rule element='table'><format attr='background-color' value='#e6f1e2'/></style-rule></style>
      <panes><pane><mark class='Automatic'/><encodings><text column='[ds].[sum:Sales:qk]'/></encodings>
        <customized-label><formatted-text><run bold='true' fontsize='15'>&lt;[ds].[sum:Sales:qk]&gt;</run><run>\u00C6&#10;</run><run fontsize='11'>Total Sales</run></formatted-text></customized-label>
      </pane></panes><rows/><cols/></table></worksheet>
    <worksheet name='Trend'><table><view/><style>
        <style-rule element='axis'><format attr='display' class='0' field='[ds].[sum:Sales:qk]' scope='rows' value='false'/><format attr='title' class='1' field='[ds].[sum:Sales:qk]' scope='rows' value=''/></style-rule>
        <style-rule element='gridline'><format attr='stroke-size' value='0'/></style-rule></style>
      <panes><pane><mark class='Line'/></pane></panes><rows>[ds].[sum:Sales:qk]</rows><cols>[ds].[mn:Order Date:ok]</cols></table></worksheet>
  </worksheets><dashboards/></workbook>`;
  const kpiModel = X.parseTableauFormatting(twb);
  const sales = X.tfParseFieldRef("[ds].[sum:Sales:qk]");
  const trendFmt = X.createSheetFormatter(kpiModel, "Trend");
  const fidelity = [
    check(trendFmt.axisInfo(sales, "rows", "0").hidden === true, "axis hidden on the primary axis (Show Header off)"),
    check(trendFmt.axisInfo(sales, "rows", "1").hidden === undefined && trendFmt.axisInfo(sales, "rows", "1").title === "",
      "secondary-axis rules stay on the secondary axis (title removed there only)"),
    check(trendFmt.gridlinesShown("rows") === false, "grid lines switched off in the worksheet")
  ];
  const kpiVm = X.buildViewModel(kpiModel, "KPI", { columns: [{ fieldName: "SUM(Sales)", dataType: "float" }],
    data: [[{ value: 2412000, nativeValue: 2412000, formattedValue: "$2.4M" }]] }, {});
  const card = X.buildKpiCard(kpiVm, "KPI", { widthPx: 238, heightPx: 70 });
  const lines = card ? card.tiles[0].lines : [];
  fidelity.push(
    check(!!card && card.tiles.length === 1 && lines.length === 2, `KPI card: one tile, two lines from the custom label (got ${card ? card.tiles.length + " / " + lines.length : "none"})`),
    check(lines[0] && lines[0].segments[0].ci === 0 && lines[0].sizePt === 15 && lines[0].segments[0].props.bold === true, "KPI card: the value line keeps the label's 15pt bold"),
    check(lines[1] && lines[1].segments[0].text === "Total Sales" && lines[1].sizePt === 11, "KPI card: the caption line keeps its text and 11pt"),
    check(!!card && card.background === "FFE6F1E2", `KPI card: the worksheet background becomes the card colour (got ${card && card.background})`)
  );
  // dashboard text boxes: banner colour and fonts kept, field tokens split over runs still resolve
  const banner = X.buildTextCard({ runs: [{ text: "SALES", props: { fontName: "Poppins SemiBold", fontSize: 16, color: "FFFFFFFF", hAlign: "center" } }],
    style: { bgColor: "FF295D79" } }, { widthPx: 338, heightPx: 67 }, null);
  const split = X.buildTextCard({ runs: ["<", "[Parameters].[Parameter 10]", "> Stores by <", "[Parameters].[Parameter 8]", ">"]
    .map(text => ({ text, props: { bold: true, fontSize: 12 } })) }, { heightPx: 30 },
    { fields: { "Parameters|parameter 10": { value: '"Top"' }, "Parameters|parameter 8": { alias: "Sales" } } });
  fidelity.push(
    check(!!banner && banner.background === "FF295D79" && banner.tiles[0].lines[0].hAlign === "center" && banner.tiles[0].lines[0].sizePt === 16,
      "text box: background, alignment and font size from the workbook"),
    check(!!split && split.tiles[0].lines[0].segments.map(s => s.text).join("") === "Top Stores by Sales",
      `text box: parameter tokens split over runs resolve (got ${split && JSON.stringify(split.tiles[0].lines[0].segments.map(s => s.text).join(""))})`)
  );
  const nav = X.buildTextCard({ runs: [{ text: "Navigation", props: {} }, { text: "Æ ", props: {} }, { text: "Menu", props: {} }] }, null, null);
  fidelity.push(check(!!nav && nav.tiles[0].lines.length === 2, "text box: the line-break mark followed by a space still breaks the line"));
  // Tableau's DashboardObject types: quick filters, parameter controls and text boxes are laid out
  const apiMap = X.buildLayoutMap([
    { id: 1, type: "worksheet", name: "Sheet", position: { x: 0, y: 100 }, size: { width: 400, height: 300 } },
    { id: 2, type: "quick-filter", name: "Region", position: { x: 420, y: 100 }, size: { width: 160, height: 80 } },
    { id: 3, type: "parameter-control", name: "Start Date", position: { x: 420, y: 200 }, size: { width: 160, height: 60 } },
    { id: 24, type: "text", name: "Text", position: { x: 0, y: 0 }, size: { width: 580, height: 60 } }
  ]);
  fidelity.push(check(apiMap.has("Region") && apiMap.get("Region").type === "filter" && apiMap.get("Start Date").type === "parameter" &&
    apiMap.has("text:24") && apiMap.get("text:24").id === 24, "layout: quick-filter / parameter-control / text objects from the Extensions API"));
  print(`fidelity: ${fidelity.filter(Boolean).length}/${fidelity.length} checks passed`);

  // chart settings from the workbook, as on a KPI dashboard: a trend with month initials, tick spacing and
  // an average line; bars labelled inside in the marks card's white font with text around the value
  const chartTwb = `<?xml version='1.0' encoding='utf-8' ?><workbook><datasources/><worksheets>
    <worksheet name='Trend'><table><view/><style>
        <style-rule element='axis'><encoding attr='space' class='0' field='[ds].[sum:Sales:qk]' field-type='quantitative' major-spacing='50000.0' scope='rows' type='space'/></style-rule>
        <style-rule element='label'><format attr='text-format' field='[ds].[mn:Order Date:ok]' value='iLLLLL'/><format attr='text-format' field='[ds].[sum:Sales:qk]' value='p0%'/></style-rule></style>
      <panes><pane><mark class='Line'/><reference-line axis-column='[ds].[sum:Sales:qk]' formula='average' id='refline0' label-type='automatic' scope='per-table' value-column='[ds].[sum:Sales:qk]'/></pane></panes>
      <rows>[ds].[sum:Sales:qk]</rows><cols>[ds].[mn:Order Date:ok]</cols></table></worksheet>
    <worksheet name='By Region'><table><view/><style/>
      <panes><pane><mark class='Bar'/><reference-line axis-column='[ds].[sum:Days:qk]' formula='average' id='refline0' label-type='automatic' scope='per-table' value-column='[ds].[sum:Days:qk]'/>
        <customized-label><formatted-text><run>&lt;</run><run>[ds].[sum:Days:qk]</run><run>&gt; DAYS</run></formatted-text></customized-label>
        <style><style-rule element='cell'><format attr='text-align' value='left'/></style-rule>
          <style-rule element='datalabel'><format attr='color-mode' value='user'/><format attr='color' value='#ffffff'/><format attr='font-family' value='Poppins SemiBold'/></style-rule>
          <style-rule element='mark'><format attr='mark-labels-show' value='true'/></style-rule></style>
      </pane></panes><rows>[ds].[none:Region:nk]</rows><cols>[ds].[sum:Days:qk]</cols></table></worksheet>
  </worksheets><dashboards/></workbook>`;
  const chartModel = X.parseTableauFormatting(chartTwb);
  const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const monthly = [43971, 20301, 58872, 36522, 44261, 52982, 45264, 63121, 87867, 77777, 118448, 96182];
  const specsOf = (sheet, cols, rows) => {
    const visual = X.buildVisualModel(chartModel, sheet, { columns: cols.map((c, i) => ({ fieldName: c[0], dataType: c[1], index: i })), data: rows },
      { dashboardName: "", displayName: sheet });
    return X.chooseVisualRenderer(visual).renderer === "excel-chart" ? X.buildExcelChartSpecs(visual, chartModel) : [];
  };
  const [trend] = specsOf("Trend", [["MONTH(Order Date)", "string"], ["SUM(Sales)", "float"]], MONTHS.map((m, i) => [dv(m), dv(monthly[i])]));
  const [bars] = specsOf("By Region", [["Region", "string"], ["SUM(Days)", "float"]],
    [["East", 4.0], ["Central", 3.9], ["South", 3.9], ["West", 3.8]].map(r => [dv(r[0]), dv(r[1])]));
  // chart part XML without the data sheet: references to cells that are not written
  const partXml = (spec, size) => X.chartXml(spec, { cat: "'Chart Data'!$A$2:$A$13", series: spec.series.map(() => ({ tx: "'Chart Data'!$B$1", val: "'Chart Data'!$B$2:$B$13" })), nextRow: 0 }, size);
  const avgLine = trend && trend.series.find(s => s.refLine);
  const trendXml = trend ? partXml(trend, { widthPx: 338, heightPx: 250 }) : "";
  const range = trend ? X.valueRange(trend, { w: 287, h: 180 }) : null;
  const charts = [
    check(!!trend && trend.categories.levels[0].join("") === "JFMAMJJASOND", `month initials from the header's date format (got ${trend && trend.categories.levels[0].join(" ")})`),
    check(!!trend && trend.valueMajorUnit === 50000 && /<c:majorUnit val="50000"\/>/.test(trendXml), "the workbook's tick spacing on the value axis"),
    check(!!range && range.fixed.min === 0 && range.fixed.max > 118448 && range.fixed.max < 130000 && range.unit === 50000,
      `the axis ends just past the data, not at an empty tick (got ${range && JSON.stringify(range)})`),
    check(!!trend && trend.valueAxisNumFmt === "0%", `the axis's own tick format (got ${trend && trend.valueAxisNumFmt})`),
    check(!!avgLine && Math.abs(avgLine.values[0] - monthly.reduce((a, b) => a + b) / 12) < 1e-6 && avgLine.refLine.labelFmt === '"Average"' &&
      /<c:dLbl><c:idx val="1"\/><c:numFmt formatCode="&quot;Average&quot;"/.test(trendXml), "the average line, labelled \"Average\" over the second point"),
    check(!!bars && !!bars.labelFont && bars.labelFont.color === "FFFFFF" && bars.labelFont.name === "Poppins SemiBold" && bars.labelPos === "inBase",
      `bar labels in the marks card's font, inside the bars at the left (got ${bars && JSON.stringify(bars.labelFont)} ${bars && bars.labelPos})`),
    check(!!bars && bars.series.filter(s => !s.refLine).every(s => /" DAYS"$/.test(s.labelNumFmt || "")),
      `label text around the value as literal text (got ${bars && bars.series[0].labelNumFmt})`),
    check(!!bars && (bars.refLines || []).length === 1 && /<c:scatterChart>/.test(partXml(bars, { widthPx: 338, heightPx: 241 })),
      "an average line across horizontal bars (an XY line over the bars)"),
    check((r => r.unit > 0 && r.fixed.min % r.unit === 0 && r.fixed.min <= -1396.17 && r.fixed.min > -1396.17 - r.unit && r.fixed.max === 761.1)(
      X.valueRange({ series: [{ values: [100, 400] }], kind: "line", valueMin: -1396.17, valueMax: 761.1 }, { w: 300, h: 200 })),
      "a fixed axis range starts at the round tick below its minimum, so the labels are round numbers"),
    check(X.tableauToExcelNumFmt("p0%") === "0%" && X.tfWrapNumFmt("#,##0.0;-#,##0.0", "", " DAYS") === '#,##0.0" DAYS";-#,##0.0" DAYS"',
      "short percent formats and label text in number formats"),
    check(X.tfFormatDateLabel("iLLLLL", "March") === "M" && X.tfFormatDateLabel("MMM yy", "", { value: "2024-03-05" }) === "Mar 24" &&
      X.tfFormatDateLabel("yyyy", "Q1") === null, "Tableau date label formats")
  ];
  print(`chart settings: ${charts.filter(Boolean).length}/${charts.length} checks passed`);

  // dashboard images: format and size from the file header, backgrounds / dividers / overlays told apart,
  // Fit Image keeping proportions, links that Excel opens
  const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 1, 44, 0, 0, 0, 100, 8, 6, 0, 0, 0]);
  const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 64, 0, 32, 0, 0, 0]);
  const jpg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 17, 8, 0, 48, 0, 96, 3, 0, 0]);
  const svg = new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 40"><rect/></svg>`);
  const info = [X.imageInfo(png), X.imageInfo(gif), X.imageInfo(jpg), X.imageInfo(svg, "logo.svg")];
  const objs = [
    { id: 1, type: "image", position: { x: 0, y: 0 }, size: { width: 1200, height: 800 } },          // background
    { id: 2, type: "image", position: { x: 20, y: 10 }, size: { width: 130, height: 40 } },          // logo
    { id: 3, type: "image", position: { x: 160, y: 10 }, size: { width: 3, height: 40 } },           // divider
    { id: 4, type: "image", position: { x: 420, y: 140 }, size: { width: 40, height: 40 } },         // icon on a sheet
    { id: 10, type: "worksheet", name: "KPI", position: { x: 400, y: 120 }, size: { width: 300, height: 200 } },
    { id: 11, type: "text", name: "Text", position: { x: 170, y: 10 }, size: { width: 600, height: 40 } }
  ];
  const kinds = X.classifyImageObjects(objs);
  const imgMap = X.buildLayoutMap(objs);
  const fit = X.fitImage(200, 100, 100, 100, true);
  const pic = await X.prepareImage({ data: png }, info[0], 30, 10);
  const imageChecks = [
    check(info[0].type === "png" && info[0].width === 300 && info[0].height === 100 && info[1].type === "gif" && info[1].width === 64 &&
      info[2].type === "jpeg" && info[2].width === 96 && info[2].height === 48 && info[3].type === "svg" && info[3].width === 120,
      `image format and size from the file header (got ${JSON.stringify(info)})`),
    check(kinds.get("1").kind === "backdrop" && kinds.get("2").kind === "block" && kinds.get("3").kind === "tiny" &&
      kinds.get("4").kind === "overlay" && kinds.get("4").host === "10", `backgrounds, logos, dividers and icons on sheets (got ${JSON.stringify([...kinds])})`),
    check(imgMap.has("image:2") && !imgMap.has("image:1") && !imgMap.has("image:4") && imgMap.get("KPI").gridCol === Math.round((400 - 20) / X.PX_PER_COL),
      "only logos take part in the layout; a background does not shift the grid"),
    check(fit.w === 100 && fit.h === 50 && fit.x === 0 && fit.y === 25, `Fit Image keeps the proportions, centred (got ${JSON.stringify(fit)})`),
    check(X.webLink("www.linkedin.com/in/x") === "https://www.linkedin.com/in/x" && X.webLink("javascript:alert(1)") === null, "image links Excel can open"),
    check(JSON.stringify(X.nativeAnchor(0, 0, 100, 30, () => 64, () => 20)) === JSON.stringify({ nativeCol: 1, nativeColOff: 36 * 9525, nativeRow: 1, nativeRowOff: 10 * 9525 }),
      "picture anchors in cells + EMU offsets"),
    check(!!pic && pic.extension === "png" && pic.base64 === Buffer.from(png).toString("base64"), "a PNG goes in as packaged when there is no canvas"),
    // print setup next to collapsed row groups: Excel refuses <sheetPr> children out of schema order
    check(X.orderSheetPr('<sheetPr><pageSetUpPr fitToPage="1"/><outlinePr summaryBelow="0"/></sheetPr>') ===
      '<sheetPr><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr>', "sheet properties in the order Excel requires")
  ];
  print(`images: ${imageChecks.filter(Boolean).length}/${imageChecks.length} checks passed`);
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
