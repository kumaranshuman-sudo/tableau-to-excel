# Mark2Table / Export2Sheet

> Tableau dashboard extension that converts dashboards into editable, formatted Excel workbooks.

[![Repository](https://img.shields.io/badge/GitHub-kumaranshuman--sudo-181717?style=flat-square&logo=github)](https://github.com/kumaranshuman-sudo/tableau-to-excel)
[![Runtime](https://img.shields.io/badge/Runtime-Tableau%20Desktop-1C1C1C?style=flat-square)](https://www.tableau.com/products/desktop)
[![Build](https://img.shields.io/badge/Build-esbuild-FFCF00?style=flat-square)](https://esbuild.github.io/)

## What it solves

Tableau dashboards are optimized for interactive analysis, while Excel is often still required for downstream editing, reporting, and distribution. Mark2Table bridges that gap by exporting dashboard content into a workbook while preserving as much semantic and visual structure as possible: each object at its dashboard position, with Tableau's fonts, colours and number formats.

The exporter recognises what each visual is from its shelves, marks and encodings, and picks the closest representation that keeps its meaning, in this order:

| Strategy | Excel output | Examples |
|---|---|---|
| **NATIVE** | a genuine Excel equivalent | text tables, KPIs, bar / column / line / area (stacked, 100 % stacked), pie, donut, scatter, bubble, combo and dual axis, treemap |
| **CONSTRUCTED** | Excel primitives that keep the meaning | heat map / highlight table / calendar (a matrix with Tableau's colour per cell), waterfall, box plot, Gantt, funnel, bullet graph, butterfly, dumbbell, lollipop, horizontal dot plot, histogram of Tableau's bins, tables built from marks |
| **APPROXIMATE** | Excel's closest chart, the difference named | packed bubbles (Excel's bubble chart), bars Tableau sizes by a measure, bars summing Tableau's detail segments |
| **TABLE_FALLBACK** | the visual's data as a table | symbol / filled / density / polygon / flow maps, map layers, viz extensions, anything no Excel chart would show correctly |
| **IMAGE_FALLBACK** | a Tableau-rendered picture, its data on a *Visual Data* sheet | charts Tableau can draw when the native chart is impossible |

A correct table always beats a misleading chart: a visual is never drawn as an Excel chart that merely looks similar. Every export adds a **Conversion Report** sheet listing each visual, the Tableau visual it is, the strategy, the Excel output, the fidelity, the reason and the evidence it was recognised from, plus the filters and parameters the data reflects. A visual exported as data carries a cell note with the same reason.

Dashboard formatting – palettes, number formats, layout and workbook styles – is applied throughout.

## Architecture

```text
Tableau Dashboard
       │
       ▼
Tableau Extensions API
       │
       ├── Workbook / TWB / TWBX parser
       │        └── formatting cascade + assets
       │
       ├── Dashboard data extraction
       │
       ▼
View model (pivot, sort, matrix) ─── visual classification (type + evidence)
       │
       ▼
Semantics (Tableau visual, encodings) ─── conversion strategy (NATIVE … TABLE_FALLBACK)
       │
       ▼
Renderer-neutral chart model
       │
       ├── Cell writer ───────────────► XLSX cells
       ├── Chart writer ──────────────► Native Excel charts
       ├── Image/data fallback ───────► XLSX pictures / data tables
       └── Report writer ─────────────► Conversion Report sheet
```

## Engineering highlights

- **TWB/TWBX parsing:** reads workbook metadata and formatting rules rather than treating the dashboard as a screenshot.
- **Renderer-neutral chart model:** separates Tableau visual interpretation from Excel DrawingML generation.
- **Native chart generation:** supported visuals become editable Excel charts.
- **Formatting pipeline:** palettes, number formats, dimensions and workbook styles are resolved before writing.
- **Local dependency bundling:** ExcelJS and JSZip are bundled into the extension; runtime CDN dependencies are avoided.
- **Type checking:** JSDoc-based checks validate shared shapes in `src/types.d.ts`.

## Setup

Requires Node.js 18 or later.

```bash
npm install
```

## Run it in Tableau Desktop

```bash
npm run serve
```

Builds the extension into `dist/`, rebuilds on every change under `src/`, and serves `dist/` on `http://localhost:5500/index.html` – the URL in `Export.trex`. In Tableau, add an Extension object to a dashboard and choose **Access Local Extensions → Export.trex** (or **Reload** an existing one).

`npm run build` makes a one-off build.

In the panel, **Load Workbook** reads the dashboard's `.twbx` (or `.twb`) for its formatting. Load the packaged `.twbx` to also export logos, custom shapes and other images: a `.twb` does not contain them.

## Deploy

```bash
npm run build
```

1. Host the contents of `dist/` on any HTTPS web server (Tableau requires HTTPS for non-localhost extensions). `dist/` is self-contained: ExcelJS and JSZip are bundled, nothing loads from a CDN.
2. In `Export.trex`, set `<source-location><url>` to the hosted `index.html`, and raise `extension-version` for each release.
3. Tableau Server / Cloud: add the URL to the extension allow list (Settings → Extensions).
4. Hand out the updated `Export.trex`; users add it through **Access Local Extensions**.

The bundle name carries a content hash (`extension.<hash>.js`), so browsers never run a stale copy.

## How an export runs

`export/export.js` → `exportToExcel()` drives one export:

1. **Read** – every worksheet's summary data from Tableau (`export/sheet-data.js`) and the dashboard objects' positions; the loaded workbook's format model (`twb/parser.js`, cached by `ui/workbook-store.js`).
2. **Model** – per worksheet a view model (`visual/view-model.js`: pivot, sort, visible columns, headers), a visual type with the evidence for it (`visual/classify.js`), heat maps and crosstabs pivoted into Tableau's matrix (`visual/matrix.js`), then a renderer (`visual/visual-model.js`): cells, a native chart, an image or a data table. Summary columns are matched to the workbook's fields by Tableau's field id (`fieldId`), else by caption.
3. **Lay out** – every object at its dashboard position: columns from the zone edges, rows at 20 px (`export/layout.js`).
4. **Write** – tables, KPI cards, text, filters and pictures as cells (`export/*-writers.js`, `export/kpi-card.js`, `export/mark-cells.js`), backgrounds (`export/backgrounds.js`), then native charts injected into the XLSX package (`charts/writer/`).

5. **Report** – each visual's semantics (`visual/semantics.js`: "Bump Chart", "Filled Map" …, its shelves, encodings and calculated fields) and conversion (`visual/strategy.js`) go to the Conversion Report sheet (`export/report.js`) and the panel's status line.

Pies and donuts have their own path (`charts/model/pie.js` → `charts/writer/pie.js`): Tableau's pane cells and mark size, labels as text boxes at Tableau's label positions, the donut hole with its total, nested donuts, one pie per pane, a colour legend at its dashboard position, and Tableau's tooltip on each slice – an invisible wedge whose hyperlink ScreenTip is the tooltip text.

What the chart model reads from the workbook, beyond shelves and marks: Dual Axis and Synchronize Axis (`fold` / `synchronized` on the axis), reversed axes (bump charts), constants that only place labels or bar starts (`MIN(0)`), negated measures (butterfly charts), per-cell reference lines (bullet targets), "% of Total" calculations (100 % stacked), Detail / Label fields that split marks (one line per member, never their sum), and pane boundaries (Tableau never draws a line across panes).

Adding support for a new kind of visual usually means a rule in `visual/classify.js` and either a chart spec builder in `charts/model/` or a cell writer in `export/`; a construction or approximation names itself in `spec.conversion` so the report says what it is.

Settings that change the output are in `src/config.js` (`FORMAT_CONFIG`): `conversionReport` (the report sheet), `fallbackNotes` (notes on visuals exported as data), `chartPolicy` and more; `debug: true` logs the parsed workbook model and formatting traces to the browser console.

## Repository layout

```text
Export.trex               extension manifest (points Tableau at the hosted index.html)
src/
  main.js                 startup and Tableau extension wiring
  config.js               formatting configuration and visual types
  util.js                 small shared helpers (XML, numbers, logging, file names)
  twb/                    TWB/TWBX parsing and formatting cascade
  data/                   summary-data values and ordering
  format/                 palettes, colour scales, number formats and ExcelJS styles
  visual/                 view model, matrix pivot, classification, semantics, conversion strategy, renderer selection
  charts/model/           visual model → renderer-neutral Excel chart specs (pie.js: pies and donuts)
  charts/writer/          chart specs → DrawingML / chartex parts (pie.js: pie parts, label shapes, tooltip wedges)
  export/                 dashboard layout, data fetching, workbook writing, Conversion Report and Visual Data sheets
  ui/                     panel status and workbook storage
  types.d.ts              shared type definitions (not bundled)
  index.html              extension panel template (the build inserts the hashed bundle name)
js/                       Tableau Extensions API library (copied into dist/)
scripts/build.mjs         esbuild bundle + development server
dist/                     build output – what gets deployed (not in git)
.github/workflows/ci.yml  CI: install, type check and build on every push / pull request
```

## Current scope and limitations

This is a **Tableau dashboard extension**: it runs inside a dashboard in Tableau Desktop (or Tableau Server / Cloud once allowed there), not as a standalone exporter. End-to-end validation therefore requires Tableau and representative workbooks.

Visual fidelity follows the strategies above; the Conversion Report names every construction, approximation and fallback. Known limits:

- **Maps** of every kind are exported as tables of their marks (location fields and values): Excel's map charts cannot reproduce Tableau's geography, layers or colour scales reliably.
- **Calculations** – calculated fields, LOD expressions, table calculations – arrive as Tableau's computed values; formulas are not translated into Excel formulas. The report lists each calculation and its kind.
- **Filters and parameters** apply as they were at export time; the report lists them. There are no slicers: Excel has no live link back to Tableau.
- **Tooltips** appear on hover for pies, donuts and packed bubbles; elsewhere their fields are in the data only. Actions, highlighting and drill-down are not reproduced.
- **Shapes and sizes** Excel cannot draw (custom shape images in charts, bar widths by a measure, rounded bar ends) are left out and named in the report.

The project should be evaluated against representative dashboards containing the visual types and formatting features it claims to support.

## Development

```bash
npm run typecheck
npm run build
```

For regression testing, use representative Tableau workbooks and verify:

1. dashboard dimensions and object ordering;
2. extracted values and filters;
3. number formats and palettes;
4. chart type and series mapping;
5. native Excel chart editability;
6. fallback behaviour for unsupported visuals;
7. the Conversion Report: every visual's Tableau visual, strategy and reason as expected.

## License

See the repository license for details.
