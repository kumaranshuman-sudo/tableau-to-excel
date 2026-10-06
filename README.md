# Mark2Table / Export2Sheet

> Tableau Desktop extension that converts dashboards into editable, formatted Excel workbooks.

[![Repository](https://img.shields.io/badge/GitHub-anxhumandev-181717?style=flat-square&logo=github)](https://github.com/anxhumandev/tableau-to-excel)
[![Runtime](https://img.shields.io/badge/Runtime-Tableau%20Desktop-1C1C1C?style=flat-square)](https://www.tableau.com/products/desktop)
[![Build](https://img.shields.io/badge/Build-esbuild-FFCF00?style=flat-square)](https://esbuild.github.io/)

## What it solves

Tableau dashboards are optimized for interactive analysis, while Excel is often still required for downstream editing, reporting, and distribution. Mark2Table bridges that gap by exporting dashboard content into a workbook while preserving as much semantic and visual structure as possible.

The exporter chooses the appropriate representation per visual:

- **Tables / KPIs** → formatted Excel cells
- **Supported charts** → native, editable Excel charts
- **Unsupported or Tableau-rendered visuals** → images or extracted data
- **Dashboard formatting** → palettes, number formats, layout and workbook-derived styling

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
Visual classification
       │
       ▼
Renderer-neutral chart model
       │
       ├── Cell writer ───────────────► XLSX cells
       ├── Chart writer ──────────────► Native Excel charts
       └── Image/data fallback ───────► XLSX assets / data
```

## Engineering highlights

- **TWB/TWBX parsing:** reads workbook metadata and formatting rules rather than treating the dashboard as a screenshot.
- **Renderer-neutral chart model:** separates Tableau visual interpretation from Excel DrawingML generation.
- **Native chart generation:** supported visuals become editable Excel charts.
- **Formatting pipeline:** palettes, number formats, dimensions and workbook styles are resolved before writing.
- **Local dependency bundling:** ExcelJS and JSZip are bundled into the extension; runtime CDN dependencies are avoided.
- **Type checking:** JSDoc-based checks validate shared shapes in `src/types.d.ts`.

## Setup

```bash
npm install
```

## Run it in Tableau Desktop

```bash
npm run serve
```

Builds the extension into `dist/`, rebuilds on every change under `src/`, and serves `dist/` on `http://localhost:5500/index.html` – the URL in `Export.trex`. In Tableau, add an Extension object to a dashboard and choose **Access Local Extensions → Export.trex** (or **Reload** an existing one).

`npm run build` makes a one-off build.

In the panel, **Load Workbook** reads the dashboard's `.twbx` (or `.twb`) for its formatting. Load the packaged `.twbx` to also export logos and other image objects: a `.twb` does not contain the image files.

## Type check

```bash
npm run typecheck
```

## Repository layout

```text
src/
  main.js                 startup and Tableau extension wiring
  config.js               formatting configuration and visual types
  twb/                    TWB/TWBX parsing and formatting cascade
  data/                   summary-data values and ordering
  format/                 palettes, colour scales, number formats and ExcelJS styles
  visual/                 view model, visual classification and renderer selection
  charts/model/            visual model → renderer-neutral Excel chart specs
  charts/writer/           chart specs → DrawingML / chartex parts
  export/                  dashboard layout, data fetching and workbook writing
  ui/                     panel status and workbook storage
  types.d.ts              shared type definitions
scripts/build.mjs          esbuild bundle + development server
index.html                extension panel
```

## Current scope and limitations

This is an **extension for Tableau Desktop**, not a standalone Tableau Server/Cloud exporter. End-to-end validation therefore requires Tableau Desktop and representative workbooks.

Visual fidelity is intentionally implementation-dependent: supported visuals are exported as native Excel charts, while unsupported or Tableau-specific visuals may fall back to images/data.

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
6. fallback behaviour for unsupported visuals.

## License

See the repository license for details.
