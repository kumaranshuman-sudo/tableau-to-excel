# Mark2Table / Export2Sheet

Tableau dashboard extension that exports a dashboard to a formatted Excel workbook: tables and KPIs
as styled cells, charts as native (editable) Excel charts, the rest as Tableau-drawn images or data.

## Setup

```bash
npm install
```

## Run it in Tableau Desktop

```bash
npm run serve
```

Builds the extension into `dist/`, rebuilds on every change under `src/`, and serves `dist/` on
`http://localhost:5500/index.html` – the URL in `Export.trex`. In Tableau, add an Extension object to a
dashboard and choose **Access Local Extensions → Export.trex** (or **Reload** an existing one).

`npm run build` makes a one-off build.

In the panel, **Load Workbook** reads the dashboard's `.twbx` (or `.twb`) for its formatting. Load the
packaged `.twbx` to also export logos and other image objects: a `.twb` does not contain the image files.

## Type check

```bash
npm run typecheck    # JSDoc type check of src/ (shared shapes in src/types.d.ts)
```

## Layout

```
src/
  main.js                 start-up: wires the panel once Tableau has initialised the extension
  config.js               settings (FORMAT_CONFIG), visual types, Tableau defaults
  twb/                    TWB/TWBX parsing and Tableau's formatting cascade
  data/                   summary-data values and ordering
  format/                 palettes, colour scales, number formats, ExcelJS styles
  visual/                 view model, visual classification, renderer choice, Tableau images
  charts/model/           visual model → renderer-neutral Excel chart specs
  charts/writer/          chart specs → DrawingML / chartex parts injected into the XLSX
  export/                 dashboard layout, data fetching, cell writers, the export itself
  ui/                     panel status and workbook storage (settings + IndexedDB)
  types.d.ts              shared shapes for the type check (chart spec, visual model …), not bundled
  index.html              panel page template (the build inserts the hashed bundle name)
scripts/build.mjs         esbuild bundle + dist/index.html + dev server
```

ExcelJS and JSZip are bundled from `node_modules` – the extension loads nothing from a CDN.
