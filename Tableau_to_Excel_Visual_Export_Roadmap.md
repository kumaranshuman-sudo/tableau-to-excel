# Tableau-to-Excel Visual Export — Current State, Target State, Feasibility & Implementation Plan

**Project:** Mark2Table / Export2Sheet
**Current build:** `Export_extension_XML(1).zip`
**Extension version:** `0.6.0`
**Primary implementation:** `build_table_copy.js`
**Date:** 30 September 2026

---

## 1. Executive Summary

The current extension is already a second-generation Tableau-to-Excel exporter. It does much more than export raw worksheet data: it parses Tableau workbook XML, reconstructs Tableau formatting and view semantics, resolves field-level formatting, handles titles and aliases, rebuilds Measure Names/Measure Values layouts, reproduces sorts, handles dynamic-zone visibility, and writes formatted Excel cells.

The current limitation is that chart worksheets are classified as `chart` and are skipped because:

```javascript
FORMAT_CONFIG.chartPolicy = "skip"
```

The next product step is to make the exporter support Tableau visual types rather than treating non-table worksheets as unsupported. The recommended architecture is a **hybrid visual exporter**:

1. Preserve the current cell renderer for tables, KPIs, heatmaps, text, and other cell-native visuals.
2. Add a normalized `VisualModel` that combines live Tableau state with the TWB-derived formatting model.
3. Use native Excel charts where the Tableau visual maps cleanly to Excel.
4. Use Tableau-rendered SVG/PNG images for visuals that cannot be faithfully represented by Excel's chart model.
5. Always preserve the current-dashboard state: filters, parameters where available, visibility, sorting, field values, titles, and formatting.
6. Provide a deterministic fallback for unsupported visuals instead of silently skipping them.

The goal should **not** be “every Tableau visual becomes an editable Excel chart.” That is unrealistic for maps, custom marks, density plots, polygon maps, and several Tableau-specific visual behaviors. The correct goal is:

> **Every supported Tableau visual is exported either as an editable Excel-native visual or as a visually faithful Tableau-rendered visual, while unsupported visuals degrade gracefully with clear status information.**

---

# 2. Current State

## 2.1 Current architecture

The current exporter has four major layers:

```text
                    Tableau Dashboard
                           │
             ┌─────────────┴─────────────┐
             │                           │
             ▼                           ▼
       Tableau API                    TWB/TWBX
             │                           │
             │                           ▼
             │                  Tableau XML parser
             │                           │
             ▼                           ▼
        Live view state          TWB format model
             │                           │
             └────────────┬──────────────┘
                          ▼
                   View-model builder
                          │
                          ▼
                   Excel cell renderer
                          │
                          ▼
                         XLSX
```

### Live Tableau API responsibilities

The current implementation obtains:

- Dashboard and worksheet objects.
- Worksheet summary data.
- Current filters.
- Dynamic Zone Visibility state.
- Dashboard object positions.
- Worksheet names and display information.
- Current worksheet data used to construct the export view model.

### TWB/TWBX responsibilities

The current implementation reads the selected `.twb` or `.twbx` locally in the browser and extracts a normalized formatting model containing information such as:

- Worksheet titles.
- Dashboard titles/text-box information.
- Datasource fields and aliases.
- Worksheet styles.
- Pane information.
- Mark classes.
- Color encodings.
- Categorical color mappings.
- Continuous color palettes.
- Number formats.
- Font rules.
- Background/banding/dividers.
- Manual and measure-based sorts.
- Measure Names aliases.
- Other worksheet/dashboard formatting metadata.

---

## 2.2 Current export pipeline

The current flow is approximately:

```text
1. Initialize Tableau extension
2. Load saved TWB metadata if available
3. Read current dashboard filters
4. Read dashboard objects / visibility
5. Fetch worksheet summary data
6. Build a view model per worksheet
7. Detect KPI vs table vs chart
8. Resolve collisions and dashboard layout
9. Create ExcelJS workbook
10. Write titles / KPI / tables into Excel cells
11. Apply Tableau-derived formatting
12. Apply Tableau-derived colors
13. Apply number formats and native numeric values
14. Group long table rows
15. Set widths
16. Apply worksheet AutoFilter where possible
17. Generate XLSX in memory
18. Trigger browser download
```

---

## 2.3 Current visual support

### Strongly supported today

- Crosstabs / standard tables.
- KPI-style single-row views.
- Tables with field-level Tableau formatting.
- Measure Names / Measure Values reconstruction.
- Multi-pane row merging in supported cases.
- Tableau-derived fonts.
- Tableau-derived number formats.
- Categorical colors.
- Continuous/diverging colors where recoverable from TWB.
- Backgrounds, banding, and borders/dividers.
- Hidden labels/fields in supported cases.
- Text-box-based headers in supported cases.
- Dashboard and worksheet title reconstruction.
- Dashboard layout reconstruction into an Excel grid.
- Large-table row grouping.
- Current filter state.
- Dynamic Zone Visibility.
- Excel-native numbers rather than always writing formatted strings.

### Currently not supported as actual Excel visuals

Chart worksheets are currently classified as `chart` and skipped under:

```javascript
FORMAT_CONFIG.chartPolicy = "skip";
```

The configuration has a `"data"` alternative intended to export chart data as a plain table, but there is currently no visual renderer that reconstructs the chart itself.

---

# 3. Current Strengths That Should Be Preserved

The following components should **not** be rewritten during visual-export work unless a concrete defect requires it.

## 3.1 TWB format model

`parseTableauFormatting()` already provides a normalized representation of workbook formatting. This should remain the single source of truth for static workbook formatting.

## 3.2 Sheet formatter / cascade resolver

`createSheetFormatter()` and its supporting resolution logic provide a useful abstraction for inheriting and overriding Tableau formatting at workbook, worksheet, field, pane, and scoped levels.

This should become the common style engine for all future visual renderers.

## 3.3 View-model reconstruction

`buildViewModel()` is already doing work that is valuable to chart export:

- Measure Names / Measure Values pivoting.
- Multi-pane merging.
- Visible field reconstruction.
- Tableau sort reconstruction.
- Title resolution.
- Header reconstruction.
- Label visibility.

This should evolve into a generic `VisualModel` pipeline rather than being replaced.

## 3.4 Layout and collision engine

The current dashboard-to-Excel spatial conversion and collision handling should remain the common layout layer.

Every future renderer should return content to the same dashboard layout system rather than implementing its own positioning logic.

---

# 4. Target State

## 4.1 Target architecture

The recommended target architecture is:

```text
                         Tableau Dashboard
                                │
              ┌─────────────────┼─────────────────┐
              │                 │                 │
              ▼                 ▼                 ▼
       Summary Data      Visual Specification   Filters /
                                               Parameters /
                                               Visibility
              │                 │                 │
              └─────────────────┼─────────────────┘
                                │
                                ▼
                       Live Tableau Model
                                │
                                │
                                ▼
                           TWB Format Model
                                │
                                ▼
                       Normalized Visual Model
                                │
                                ▼
                        Visual Type Classifier
                                │
          ┌─────────────────────┼──────────────────────┐
          │                     │                      │
          ▼                     ▼                      ▼
   Cell Renderer          Excel Chart Renderer    Image Renderer
          │                     │                      │
          │                     │               Tableau-rendered
          │                     │                  SVG/PNG
          │                     │                      │
          └─────────────────────┼──────────────────────┘
                                ▼
                      Shared Dashboard Layout
                                │
                                ▼
                           XLSX Generator
```

---

## 4.2 The normalized `VisualModel`

The central design change should be to introduce a normalized visual representation between Tableau extraction and Excel rendering.

Conceptually:

```javascript
{
  type: "bar",
  title: "Sales by Region",

  data: {
    rows: [...],
    dimensions: [...],
    measures: [...]
  },

  axes: {
    x: {...},
    y: {...},
    secondaryY: {...}
  },

  encodings: {
    color: {...},
    size: {...},
    text: {...},
    detail: {...},
    tooltip: {...}
  },

  panes: [...],

  style: {...},

  layout: {...},

  metadata: {
    worksheetName: "...",
    markClass: "bar"
  }
}
```

The point is not the exact schema. The important architectural principle is:

> **Renderers must consume a stable internal visual model rather than directly parsing TWB XML or raw Tableau API responses.**

---

# 5. Data Sources: Live API vs TWB XML

A hard architectural rule should be adopted for the new system.

## Live Tableau state = current state

Use the live Extensions API for:

- Current data values.
- Current filters.
- Current worksheet state.
- Dynamic Zone Visibility.
- Current dashboard object positions.
- Current visual specification where available.
- Current data ordering.
- Current interaction state that affects the displayed result.

## TWB/TWBX = workbook metadata

Use the workbook for:

- Static formatting rules.
- Font details.
- Number formats.
- Color palettes.
- Explicit category-color mappings.
- Field aliases.
- Mark/pane metadata.
- Workbook authoring rules.
- Dashboard text/title definitions.
- Additional semantics not exposed cleanly through the live API.

The exporter should never assume that the TWB alone represents the current state of the dashboard.

---

# 6. Feasibility Assessment

## 6.1 Feasibility scale

- **High:** Can be implemented reliably with the current architecture with relatively low risk.
- **Medium:** Feasible, but requires additional modeling or chart-specific handling.
- **Low:** Possible only as an approximation or image export.
- **Very Low:** Exact editable Excel reproduction is not a realistic goal; image/fallback is the appropriate solution.

---

## 6.2 Visual-by-visual feasibility

| Tableau visual | Native Excel | Tableau-rendered image | Feasibility | Recommended approach |
|---|---|---|---|---|
| Text | Yes | Not needed | High | Excel cells |
| KPI | Yes | Not needed | High | Excel cells |
| Crosstab | Yes | Not needed | High | Excel cells |
| Highlight table | Yes | Optional | High | Excel cells + fills |
| Heatmap | Yes via cells | Optional | High | Excel cells + color |
| Bar | Yes | Yes | High | Native Excel first |
| Column | Yes | Yes | High | Native Excel first |
| Line | Yes | Yes | High | Native Excel first |
| Area | Yes | Yes | High | Native Excel first |
| Pie | Yes | Yes | High | Native Excel first |
| Doughnut | Yes | Yes | High | Native Excel first |
| Scatter | Yes | Yes | High | Native Excel first |
| Bubble | Yes | Yes | Medium/High | Native Excel first |
| Treemap | Yes in modern Excel | Yes | Medium/High | Native or image fallback |
| Histogram | Yes after correct binning | Yes | Medium/High | Native Excel |
| Waterfall | Yes | Yes | Medium | Native or image fallback |
| Combo chart | Yes | Yes | Medium/High | Native where mapping is clean |
| Dual axis | Possible | Yes | Medium/High | Native or image fallback |
| Gantt | Approximation | Yes | Medium | Image first; native approximation later |
| Funnel | Approximation | Yes | Medium | Image first |
| Box plot | Possible | Yes | Medium/Low | Image first |
| Symbol map | Limited | Yes | Medium | Tableau image |
| Filled map | Poor mapping | Yes | Low | Tableau image |
| Density map | No meaningful native equivalent | Yes | Low | Tableau image |
| Polygon map | No | Yes | Very Low | Tableau image |
| Shape map | Limited | Yes | Low | Tableau image |
| Word cloud | No | Yes | Low | Tableau image |
| Sankey | No | Yes/custom | Very Low | Tableau image or future custom renderer |
| Custom Viz / Extension | No generic mapping | Sometimes | Very Low | Image/fallback |
| Unknown/future visual | No | Potentially | Unknown | Graceful fallback |

---

# 7. Recommended Rendering Strategy

The exporter should have three rendering modes.

## 7.1 Cell renderer

Continue using the current cell-based renderer for:

- Tables.
- KPI cards.
- Heatmaps.
- Highlight tables.
- Text-based visuals.
- Filter-value displays.

These are not really “charts” and should not be forced into Excel chart objects.

## 7.2 Native Excel chart renderer

Use native Excel charts when the Tableau encoding maps cleanly to Excel.

Initial candidates:

- Bar.
- Column.
- Line.
- Area.
- Pie.
- Doughnut.
- Scatter.
- Bubble.
- Simple combination charts.
- Histogram.
- Treemap where practical.

The native renderer should prioritize semantic fidelity and editability rather than pixel-perfect Tableau styling.

## 7.3 Tableau-rendered image renderer

Use Tableau's own visualization rendering path for difficult visuals.

Conceptual pipeline:

```text
Visual Model
     ↓
Tableau visual specification
     ↓
Tableau visualization rendering
     ↓
SVG
     ↓
PNG when required by workbook embedding
     ↓
Excel image
```

This is the preferred approach for:

- Maps.
- Density maps.
- Polygon maps.
- Complex dual-axis views.
- Gantt views with complex formatting.
- Box plots.
- Custom marks.
- Views whose Tableau styling cannot be represented faithfully by Excel.

---

# 8. Visual Classification

A dedicated classifier should replace the current broad `chart` classification.

Suggested types:

```text
TABLE
KPI
HEATMAP
BAR
COLUMN
LINE
AREA
PIE
DONUT
SCATTER
BUBBLE
TREEMAP
HISTOGRAM
WATERFALL
COMBO
DUAL_AXIS
GANTT
BOXPLOT
FUNNEL
MAP_SYMBOL
MAP_FILLED
MAP_DENSITY
MAP_POLYGON
CUSTOM
UNKNOWN
```

Recommended precedence:

```text
1. Live visual specification
2. TWB pane / mark metadata
3. TWB structural evidence
4. Existing view-model heuristics
5. UNKNOWN
```

This avoids relying on summary-data shape as the primary classification mechanism.

---

# 9. Major Changes Required

## Change 1 — Introduce `VisualModel`

Create a normalized representation of each worksheet before rendering.

### Why

Currently the view model is primarily optimized for cell tables. Charts need explicit concepts for axes, series, encodings, mark types, panes, and visual options.

### Result

All renderers consume the same data model.

---

## Change 2 — Obtain live visual specification

Investigate and integrate `worksheet.getVisualSpecificationAsync()` where supported by the bundled Tableau Extensions API version.

Use it as the preferred source for current visual semantics.

### Do not

Treat the TWB as the authoritative current visual state.

---

## Change 3 — Build a renderer registry

Use a registry instead of a large chain of conditionals.

Conceptually:

```javascript
const VISUAL_RENDERERS = {
  table: renderTable,
  kpi: renderKPI,
  heatmap: renderHeatmap,
  bar: renderBar,
  column: renderColumn,
  line: renderLine,
  area: renderArea,
  pie: renderPie,
  donut: renderDonut,
  scatter: renderScatter,
  bubble: renderBubble,
  treemap: renderTreemap,
  combo: renderCombo,
  map: renderMapImage,
  fallback: renderFallback
};
```

---

## Change 4 — Add chart placement as a first-class content type

The existing dashboard layout engine should support:

```text
CELL_BLOCK
CHART_BLOCK
IMAGE_BLOCK
```

All three must share:

- Grid coordinates.
- Width/height.
- Collision handling.
- Snap-out-of-grouped-row handling.
- Dashboard title offsets.

---

## Change 5 — Add Excel chart generation

The current ExcelJS stack does not provide a convenient high-level chart authoring layer for the required chart coverage.

Before implementation, evaluate:

1. A chart-capable XLSX library.
2. Direct OOXML chart generation.
3. A hybrid approach where ExcelJS creates the workbook and chart OOXML is injected into the XLSX package.

The decision should consider:

- Browser compatibility.
- Bundle size.
- License.
- Client-network constraints.
- Chart coverage.
- Maintenance.
- Dual-axis support.
- Treemap support.
- Image embedding.

---

## Change 6 — Add Tableau-image renderer

Where native Excel mapping is weak, use Tableau's rendering capabilities to create a visual asset and embed it into the workbook.

The renderer should preserve the visual's dashboard position and dimensions.

---

## Change 7 — Add visual status reporting

Do not silently skip visuals.

The export UI should be able to report something like:

```text
Export complete

✓ 8 tables
✓ 5 KPIs
✓ 6 native Excel charts
✓ 3 Tableau-rendered visuals
⚠ 1 unsupported custom visual exported as image
```

This is essential for trust.

---

# 10. Implementation Plan

## Phase 0 — Establish a stable baseline

### Objective

Protect the currently working exporter before introducing visual rendering.

### Tasks

- Create a copy/tag of the current working build.
- Freeze existing table/KPI behavior.
- Capture representative test dashboards.
- Record expected XLSX output for those dashboards.
- Verify current TWB loading and persistence.
- Verify current layout/collision behavior.

### Exit criteria

Existing table/KPI exports are unchanged after refactoring.

---

## Phase 1 — Visual discovery layer

### Objective

Understand every worksheet before deciding how to export it.

### Tasks

Implement a diagnostic extraction pass producing:

```text
worksheet
  ├─ mark type
  ├─ rows
  ├─ columns
  ├─ panes
  ├─ color
  ├─ size
  ├─ text
  ├─ detail
  ├─ axes
  ├─ field references
  └─ current summary data
```

Integrate `getVisualSpecificationAsync()` where available.

### Output

A debug representation of every dashboard visual.

### Exit criteria

For a selected test dashboard, the extension can correctly explain why a worksheet is classified as table/KPI/bar/line/etc.

---

## Phase 2 — Normalized `VisualModel`

### Objective

Separate extraction from rendering.

### Tasks

Create:

```text
buildVisualModel()
```

and move chart-relevant logic out of the export loop.

The model should include:

- Visual type.
- Dimensions.
- Measures.
- Axes.
- Series.
- Encodings.
- Color.
- Size.
- Labels.
- Pane structure.
- Sort.
- Title.
- Number formats.
- Style.
- Layout dimensions.

### Exit criteria

The current table/KPI renderer can consume the normalized model without behavioral regressions.

---

## Phase 3 — Renderer framework

### Objective

Introduce pluggable visual renderers.

### Tasks

Add:

```text
CellRenderer
ExcelChartRenderer
ImageRenderer
```

plus a visual renderer registry.

Add renderer selection logic:

```text
visual type
    ↓
canNativeRender?
    ↓
YES → Excel renderer
NO  → canTableauRenderImage?
          ↓
       YES → Image renderer
          ↓
       NO → fallback renderer
```

### Exit criteria

Adding a new visual type does not require rewriting the main export loop.

---

## Phase 4 — Tableau-rendered visual proof of concept

### Objective

Get visual export working end-to-end with minimal implementation risk.

### First candidates

- Bar.
- Column.
- Line.
- Area.
- Scatter.
- Pie.

### Tasks

- Construct the required Tableau visualization input specification.
- Generate SVG.
- Convert SVG to embeddable image format where necessary.
- Add image to Excel at the dashboard object's coordinates.
- Preserve width/height.
- Preserve title positioning.

### Exit criteria

A dashboard containing tables plus charts can be exported to a single XLSX where tables are editable cells and charts are visible Tableau-rendered visuals.

---

## Phase 5 — Native Excel charts

### Objective

Replace image output with editable Excel charts where the mapping is robust.

### Recommended order

1. Bar.
2. Column.
3. Line.
4. Area.
5. Pie.
6. Doughnut.
7. Scatter.
8. Bubble.
9. Histogram.
10. Treemap.
11. Combo / dual-axis.

### For each chart

Implement the following separately:

```text
Data mapping
Category axis
Value axis
Series mapping
Color mapping
Legend
Axis title
Number format
Data labels
Sort
Secondary axis where applicable
Dashboard positioning
```

### Exit criteria

Each chart has an automated comparison against the source Tableau visual for data correctness and basic visual semantics.

---

## Phase 6 — Complex visuals

Implement or classify:

- Gantt.
- Waterfall.
- Funnel.
- Box plot.
- Treemap.
- Dual-axis.
- Multi-pane charts.

Use native Excel only where the mapping is stable.

Use Tableau-rendered images when exact semantics or styling would otherwise be lost.

---

## Phase 7 — Maps and special visuals

Use image export for:

- Symbol maps.
- Filled maps.
- Density maps.
- Polygon maps.
- Shape maps.
- Custom marks.
- Word clouds.
- Sankey-style custom visuals.
- Viz extensions where supported.

Do not force these into generic Excel chart objects.

---

## Phase 8 — Fallback and diagnostics

### Objective

Make the exporter trustworthy when something is unsupported.

For every visual, record:

```text
worksheet
visual type
renderer
success/failure
fallback used
reason
```

Example:

```text
Sales Map
Type: MAP_DENSITY
Renderer: TABLEAU_IMAGE
Status: success
```

or:

```text
Custom Flow
Type: CUSTOM
Renderer: FALLBACK
Status: warning
Reason: unsupported visual type
```

---

## Phase 9 — Performance and reliability

### Data acquisition

Evaluate replacing the deprecated-style summary-data path with Tableau's newer reader/pagination APIs for large datasets where appropriate.

### Concurrency

Keep bounded concurrency. Do not render dozens of large SVGs/images simultaneously.

Recommended architecture:

```text
Data extraction queue
       ↓
Visual render queue
       ↓
Workbook assembly
```

### Memory

Release large SVG/PNG buffers after the image is inserted into the workbook when possible.

### Network dependencies

The current build loads ExcelJS, JSZip, jQuery, and Bootstrap from external CDNs. For enterprise deployment, evaluate bundling dependencies locally to avoid client-tenant/network failures.

---

# 11. Testing Strategy

Visual export cannot be validated only by “the XLSX opened.” It needs multiple dimensions of correctness.

## 11.1 Data correctness

Compare:

- Categories.
- Measures.
- Aggregations.
- Sort order.
- Filtered values.
- Null handling.
- Number formats.

## 11.2 Semantic correctness

Check:

- Correct chart type.
- Correct x/y mapping.
- Correct series.
- Correct color encoding.
- Correct size encoding.
- Correct secondary axis.
- Correct legend.

## 11.3 Visual correctness

Check:

- Position.
- Width/height.
- Title.
- Fonts.
- Colors.
- Axis formatting.
- Background.
- Labels.
- Gridlines where applicable.

## 11.4 Regression testing

Every change must be checked against the existing table/KPI dashboards.

---

# 12. Test Dashboard Matrix

A strong test suite should deliberately include:

```text
Dashboard A — Tables + KPI only
Dashboard B — Bar + table
Dashboard C — Line + multiple series
Dashboard D — Area + color encoding
Dashboard E — Scatter + size + color
Dashboard F — Pie + labels
Dashboard G — Dual axis / combo
Dashboard H — Heatmap
Dashboard I — Treemap
Dashboard J — Gantt
Dashboard K — Filled map
Dashboard L — Density map
Dashboard M — Multi-pane visual
Dashboard N — Measure Names / Measure Values chart
Dashboard O — Dynamic Zone Visibility
Dashboard P — Complex filters / hierarchy
Dashboard Q — Large-data dashboard
Dashboard R — Custom/unsupported visual
```

The last dashboard in each class should explicitly test fallback behavior.

---

# 13. Important Technical Risks

## Risk 1 — Tableau visual specification mismatch

The live visual API and TWB XML may expose overlapping but not identical information.

### Mitigation

Define ownership clearly:

```text
Live API → current visual state
TWB → workbook metadata / formatting
```

---

## Risk 2 — Exact Tableau styling cannot always map to Excel

Tableau and Excel use different rendering engines and visual models.

### Mitigation

Use image rendering for views where native Excel loses important semantics.

---

## Risk 3 — ExcelJS chart limitations

The current ExcelJS setup does not provide the simple chart-authoring API required for broad native chart coverage.

### Mitigation

Make the chart-generation layer independent from the cell renderer and evaluate a chart-capable library versus OOXML generation.

---

## Risk 4 — Large data / large charts

Scatter plots, detailed marks, and large dashboards can create large in-memory representations.

### Mitigation

Use pagination/readers, bounded concurrency, and rendering queues.

---

## Risk 5 — Wrong or stale TWB file

Because the XML metadata is loaded from a user-selected workbook, a different workbook can theoretically be paired with the current Tableau dashboard.

### Mitigation options

- Display loaded workbook name.
- Compare workbook/dashboard/worksheet names.
- Validate worksheet names against current dashboard.
- Warn when coverage is low.
- Potentially store a workbook fingerprint/hash where appropriate.

---

## Risk 6 — Tableau XML schema evolution

The parser depends on TWB structures that may change across Tableau versions.

### Mitigation

Keep parsing modular, validate assumptions, and maintain fixture workbooks from multiple Tableau versions.

---

## Risk 7 — External CDN dependencies

Current production-like functionality depends on network-accessible CDN libraries.

### Mitigation

Bundle dependencies locally for enterprise deployment where allowed.

---

# 14. Recommended Product Behavior

The export process should not expose the internal complexity to the user.

The user should simply click:

```text
Export to Excel
```

The exporter internally decides:

```text
Table       → editable cells
KPI         → editable cells
Heatmap     → editable cells + formatting
Bar         → native Excel chart if supported
Line        → native Excel chart if supported
Scatter     → native Excel chart if supported
Map         → Tableau-rendered image
Complex     → best available renderer
Unsupported → explicit fallback/warning
```

---

# 15. Recommended Configuration Model

The configuration should eventually look conceptually like:

```javascript
const VISUAL_EXPORT_CONFIG = {
  nativeCharts: true,
  tableauImageFallback: true,

  unsupportedVisualBehavior: "image-or-warning",

  charts: {
    bar: "native",
    column: "native",
    line: "native",
    area: "native",
    pie: "native",
    donut: "native",
    scatter: "native",
    bubble: "native",
    treemap: "native-or-image",
    combo: "native-or-image",
    gantt: "image",
    boxplot: "image",
    map: "image"
  },

  preserveTitles: true,
  preserveColors: true,
  preserveFonts: true,
  preserveNumberFormats: true,
  preserveSort: true,
  preserveFilters: true,
  preserveDynamicZoneVisibility: true
};
```

The exact configuration names can change; the principle is to make rendering policy explicit and testable.

---

# 16. Definition of Done

The visual-export project should be considered successful when all of the following are true:

### Coverage

A broad range of common Tableau visuals is exported instead of silently skipped.

### Fidelity

Data, sort order, filters, titles, major encodings, and formatting are preserved.

### Editability

Common simple charts are editable Excel charts where practical.

### Visual fidelity

Complex visuals fall back to Tableau-rendered images rather than badly approximated charts.

### Layout fidelity

Charts, tables, KPIs, and images remain in the intended dashboard arrangement without collisions.

### Reliability

One unsupported visual does not break the entire workbook export.

### Transparency

The user can tell which visuals were exported natively, which were rendered as images, and which required a fallback.

### Regression safety

Existing table/KPI exports continue to work identically after visual support is introduced.

---

# 17. Final Recommended Roadmap

```text
CURRENT
  │
  │ tables/KPI + TWB formatting
  │ charts skipped
  ▼
PHASE 1
  Visual discovery
  + live visual specification
  ▼
PHASE 2
  Normalized VisualModel
  ▼
PHASE 3
  Renderer registry
  + common layout interface
  ▼
PHASE 4
  Tableau-rendered chart images
  ├── bar
  ├── column
  ├── line
  ├── area
  ├── scatter
  └── pie
  ▼
PHASE 5
  Native Excel charts
  ├── bar
  ├── column
  ├── line
  ├── area
  ├── pie/donut
  ├── scatter/bubble
  └── simple combo
  ▼
PHASE 6
  Complex visual types
  ├── treemap
  ├── histogram
  ├── waterfall
  ├── dual-axis
  ├── gantt
  └── box plot
  ▼
PHASE 7
  Maps + custom visuals
  → Tableau image renderer
  ▼
PHASE 8
  Fallbacks + diagnostics
  ▼
PHASE 9
  Performance + enterprise hardening
  ▼
TARGET
  Tableau dashboard → XLSX
  with editable native visuals where practical
  and Tableau-rendered images where native mapping is poor.
```

---

# 18. Bottom-Line Recommendation

The project is **feasible** and the current codebase is a strong starting point. The wrong strategy would be to keep extending `build_table_copy.js` with visual-specific conditionals and try to recreate every Tableau visual manually.

The strongest implementation is:

```text
Current TWB parser
        +
Current format resolver
        +
Current view-model logic
        +
Live Tableau visual specification
        ↓
Normalized VisualModel
        ↓
Renderer abstraction
        ├── Excel-native renderer
        ├── Tableau-image renderer
        └── Cell renderer
```

The first implementation milestone should be **visual discovery + `VisualModel` + Tableau-image rendering for a small group of standard charts**. This gives immediate chart coverage while keeping the current table/KPI export stable. Native Excel charts can then be introduced one type at a time where the mapping is genuinely robust.

The guiding principle should be:

> **Prefer semantic/editable Excel output when Excel can represent the Tableau visual correctly; otherwise preserve Tableau's actual visual appearance through Tableau-rendered imagery. Never silently replace a complex Tableau visualization with a misleading Excel approximation.**
