/* Shared shapes for the type check (npm run typecheck). Nothing here is emitted or bundled.
 * The JS modules refer to these names in JSDoc (@param {ChartSpec} spec …); the interfaces are
 * global, so no import is needed. Strict mode is off: every property may be null/undefined.  */

/** The Tableau Extensions API, loaded by js/tableau.extensions.1.latest.js before the bundle. */
declare const tableau: any;

/* ── Tableau summary data ────────────────────────────────────────────────── */

/** One cell of Worksheet.getSummaryDataAsync() data. */
interface DataValue {
  value: any;
  nativeValue?: any;
  formattedValue?: string;
}

interface SummaryColumn {
  fieldName: string;
  fieldId?: string;
  dataType?: string;
  index?: number;
}

/** Worksheet.getSummaryDataAsync() result (the parts the export reads). */
interface SummaryData {
  columns: SummaryColumn[];
  data: DataValue[][];
}

/* ── TWB format model (twb/parser.js; also persisted as JSON in the extension settings) ── */

/** "[federated.0x].[sum:Sales:qk]" → { ds, deriv: "sum", name: "Sales", type: "qk", inner: "sum:Sales:qk" } */
interface FieldRef {
  raw?: string;
  ds?: string | null;
  /** aggregation / date part / table calc: sum, avg, yr, tmn, pcto, usr, none … */
  deriv?: string | null;
  name: string;
  /** qk = continuous measure, nk / ok = discrete */
  type?: string | null;
  inner: string;
}

/** One style rule from a <style-rule> / <format> element. */
interface StyleRule {
  attr: string;
  value: string;
  /** raw field reference ("[ds].[sum:Sales:qk]"); absent = applies to every field */
  field?: string;
  scope?: string;
  dataClass?: string;
  /** axis rules: "0" = primary, "1" = secondary axis of a dual axis */
  axisClass?: string;
}

interface ParsedStyle {
  /** element name ("cell", "header", "mark", "axis" …) → its rules */
  rules: Record<string, StyleRule[]>;
  encodings: ColorEncoding[];
  /** Edit Axis settings (encoding attr="space") */
  spaces?: AxisSpace[];
  /** Shape encodings: value (bucket key) → shape name */
  shapes?: { field: FieldRef; map: Record<string, string> }[];
}

/** Edit Axis: fixed range ("fixed" / "fixedmin" / "fixedmax"), tick spacing, include zero */
interface AxisSpace {
  field?: string;
  scope?: string;
  axisClass?: string;
  rangeType?: string;
  min?: number;
  max?: number;
  majorSpacing?: number;
  /** "false" = Tick Marks: None – no ticks and no tick labels: the axis shows nothing */
  majorShow?: string;
  /** "false" = the axis does not extend to zero */
  domainExpand?: string;
  /** Edit Axis → Scale → Reversed */
  reverse?: boolean;
  /** Dual Axis: folded onto the axis before it (overlaid), and with the same scale */
  fold?: boolean;
  synchronized?: boolean;
}

/** Analytics → Reference Line */
interface ReferenceLine {
  formula: string;
  /** automatic / value / computation / custom / none */
  labelType: string;
  /** custom label, with <Value> / <Computation> tokens */
  label?: string;
  scope?: string;
  /** constant lines */
  value?: number;
  axis?: FieldRef;
  field?: FieldRef;
}

/** A <encoding attr="color"> palette / interpolation definition. */
interface ColorEncoding {
  field: FieldRef | null;
  type: string;
  paletteName: string | null;
  paletteType: string | null;
  customColors: string[];
  reverse: boolean;
  center: number | null;
  min: number | null;
  max: number | null;
  /** bucket key → ARGB colour */
  map: Record<string, string>;
}

interface TextRun {
  text: string;
  props: Record<string, any>;
}

interface Pane {
  id: string | null;
  xIndex: number | null;
  xAxisName: string | null;
  yIndex: number | null;
  /** dual axis / multi-pane: the measure this pane draws */
  yAxisName: string | null;
  /** Tableau mark class: Automatic, Bar, Line, Area, Circle, Shape, Square, Pie, GanttBar, Text … */
  markClass: string;
  encodings: { channel: string; field: FieldRef }[];
  labelRuns: { refs: FieldRef[]; text: string; props: Record<string, any> }[];
  /** Tooltip editor text, [] = Tableau's default tooltip */
  tooltipRuns?: TextRun[];
  style: ParsedStyle;
}

interface SheetModel {
  name: string;
  title: { text: string; runs: TextRun[] } | null;
  style: ParsedStyle;
  rows: FieldRef[];
  cols: FieldRef[];
  panes: Pane[];
  fieldRefs: FieldRef[];
  manualSorts?: { field: FieldRef; direction: string; order: string[] }[];
  measureSorts?: { field: FieldRef; measure: FieldRef; direction: string }[];
  measureFilter?: string[];
  boxPlot?: boolean;
  referenceLines?: ReferenceLine[];
  runningTotals?: FieldRef[];
  /** quick table calcs: field inner name → "compute using" (ordering-type: rows, columns …) */
  tableCalcs?: Record<string, string>;
}

/** A dashboard zone; x / y / w / h are in Tableau's 0–100000 units. */
interface DashboardZone {
  id: string;
  name?: string;
  /** worksheet, filter, paramctrl, text, bitmap, layout-basic … */
  type: string;
  showTitle: boolean;
  hidden: boolean;
  x: number;
  y: number;
  w: number;
  h: number;
  runs?: TextRun[];
  /** image zones: the file inside the .twbx; Fit Image; Center Image (undefined = on) */
  param?: string;
  scaled?: boolean;
  centered?: boolean;
  /** image / button zones: the URL it opens */
  url?: string;
  /** id of the layout container the zone sits in */
  parent?: string;
  /** Layout pane formatting: background and border (ARGB colours) */
  style?: { bgColor?: string; borderColor?: string; borderStyle?: string; borderWidth?: number };
}

interface DashboardModel {
  title: { text: string; runs: TextRun[] } | null;
  style: ParsedStyle;
  zones: DashboardZone[];
  width?: number;
  height?: number;
  /** each sheet's fit on the dashboard: "entire-view" | "fit-width" | "fit-height" (none = Standard) */
  fit?: Record<string, string>;
}

interface FieldInfo {
  name: string;
  ds: string;
  caption?: string;
  role?: string;
  datatype?: string;
  defaultFormat?: string;
  alias?: string;
  value?: string;
  /** a calculation that is a single number ("MIN(0)") */
  constant?: number;
  /** a calculated field: what kind (parser.js tfCalcKind) and its formula, shortened */
  calcKind?: string;
  formula?: string;
  /** a bin of a measure / a group of a dimension */
  bin?: boolean;
  group?: boolean;
  /** a parameter */
  param?: boolean;
  /** a geographic role ("[State].[Name]") */
  geoRole?: string;
}

/** parseTableauFormatting() result. */
interface FormatModel {
  version: number;
  workbookStyle: ParsedStyle;
  /** "ds|name" and "|name" (lower-case) → column info */
  fields: Record<string, FieldInfo>;
  datasourceStyles: Record<string, ParsedStyle>;
  sheets: Record<string, SheetModel>;
  dashboards: Record<string, DashboardModel>;
  /** tfRefKey(ref) → Measure Names alias */
  measureAliases: Record<string, string>;
  actions?: { caption?: string; expression: string; dashboard?: string; worksheet?: string; exclude: string[] }[];
}

/** createSheetFormatter(model, sheetName): Tableau's formatting cascade for one worksheet. */
type SheetFormatter = ReturnType<typeof import("./twb/formatter.js").createSheetFormatter>;

/* ── view model / visual model (visual/) ─────────────────────────────────── */

interface ViewColumn {
  name: string;
  /** ref resolved from the column's field id (exact), not its caption */
  exact?: boolean;
  dataType?: string;
  ref: FieldRef | null;
  /** a Measure Values measure turned into its own column */
  pivoted?: boolean;
  /** dimension (row/column header) rather than a measure */
  isHeader?: boolean;
  label?: string;
  labelProps?: Record<string, any>;
  /** a dashboard text box as the header: its runs */
  labelRuns?: any[];
  /** the header spans this many columns (one text box over several panes) */
  headerSpan?: number;
  /** under another column's spanning header */
  headerCovered?: boolean;
  zoneWidthPx?: number;
  link?: { caption?: string; [key: string]: any };
}

/** buildViewModel(): summary data pivoted, merged and sorted like the Tableau view. */
interface ViewModel {
  fmt: SheetFormatter;
  cols: ViewColumn[];
  rows: DataValue[][];
  /** indexes into cols of the visible columns, in display order */
  order: number[];
  notes: string[];
  showHeaderRow: boolean;
  title: { text: string; props: Record<string, any> };
  showTitle: boolean;
  headerZoneIds: string[];
  /** a table built from marks (constant axes): columns drawn like their panes' marks */
  markTable?: boolean;
  /** height (px) of the dashboard text boxes drawn as the headers, 0 = none */
  headerPx?: number;
  dashboardName?: string;
  kind: "chart" | "table";
  /** all header columns (hidden ones too), outer → inner */
  headerOrder: number[];
  /** a heat map / highlight table pivoted into a matrix: the Columns shelf's members as columns (view-model.js) */
  matrix?: { colDims: number[]; valueCi: number; colorCi: number };
  /** per-cell fill of a matrix (ARGB), null = the column's own */
  cellFill?: (rowIdx: number, ci: number) => string | null;
  /** false: the cells would not show what the marks encode, so the export writes every column (TABLE_FALLBACK) */
  lossless?: boolean;
  fallbackReason?: string;
}

/** Value of VISUAL_TYPES (config.js): "TABLE", "BAR", "LINE", "PIE", "TREEMAP" … */
type VisualType = string;

type RendererName = "cell" | "excel-chart" | "tableau-image" | "data-fallback";

/** chooseVisualRenderer() / imageOrFallback(): how a visual is exported. */
interface RendererDecision {
  renderer: RendererName;
  status: "success" | "pending" | "warning";
  reason: string;
}

/** buildVisualModel(): one worksheet's view model, classification and the renderer chosen for it. */
interface VisualModel {
  type: VisualType;
  title: ViewModel["title"];
  data: { rows: DataValue[][]; columns: ViewColumn[]; dimensions: ViewColumn[]; measures: ViewColumn[] };
  axes: any;
  encodings: any;
  panes: any[];
  style: any;
  layout: any;
  source: { visualSpec: any };
  formatModel: FormatModel | null;
  metadata: {
    worksheetName: string;
    markClass: string | null;
    markToken: string;
    markTokens: string[];
    markSource: string;
    visualSpecAvailable: boolean;
    visualSpecError: string | null;
    visualSpecKeys: string[];
    imageApiAvailable: boolean;
  };
  viewModel: ViewModel;
  diagnostics: any[];
  /** why the classifier chose the type (conversion report) */
  evidence?: string[];
  /* set by the export once a renderer is chosen */
  renderer?: RendererName;
  status?: string;
  statusReason?: string;
  chartSpecs?: ChartSpec[];
  /** a pie / donut drawn by charts/model/pie.js */
  pie?: any;
}

/* ── chart model (charts/model) ──────────────────────────────────────────── */

/** A measure on a shelf: summary column index + its field. */
interface RoleValue {
  ci: number;
  ref: FieldRef | null;
  /** part of Measure Values */
  mv?: boolean;
  /** axis number on the shelf ("A + B" = two axes) */
  axis?: number;
}

interface RoleDim {
  ci: number;
  ref: FieldRef | null;
  continuous: boolean;
}

interface ShelfRoles {
  values: RoleValue[];
  dims: RoleDim[];
  axisRefs: FieldRef[];
}

/** Where tvRoles() found the field roles: TWB shelves, the live visual spec, or summary columns only. */
type RolesSource = "twb" | "live" | "summary" | "none";

/** tvRoles(): which summary column sits on which shelf / encoding. */
interface Roles {
  rows: ShelfRoles;
  cols: ShelfRoles;
  /** shelf holding Measure Names: "rows" | "cols" */
  measureNames: string | null;
  color: { ci?: number; ref?: FieldRef; continuous?: boolean; measureNames?: boolean } | null;
  /** column index of the pie angle measure, -1 = none */
  angle: number;
  /** column index of the Size measure, -1 = none */
  size: number;
  textDims: number[];
  detailDims: number[];
  labelRefs: FieldRef[];
  labelNames: string[];
  panes: Pane[];
  source: RolesSource;
  /** live spec only: one mark token per marks card */
  liveMarks?: string[];
}

/** What every chart spec builder receives. */
interface ChartContext {
  vm: ViewModel;
  /** the workbook's format model (field captions, calculations, constants), null without one */
  model?: FormatModel | null;
  /** the mark Tableau's Automatic draws for these shelves (an Automatic marks card beside explicit ones) */
  autoMark?: string;
  /** rounded bars: thick lines in Tableau, bars in Excel */
  roundedBar?: boolean;
  roles: Roles;
  markToken: string;
  type: VisualType;
  doughnut?: boolean;
}

type ChartKind = "bar" | "line" | "area" | "combo" | "pie" | "doughnut" | "scatter" | "bubble" | "treemap";

interface ChartCategories {
  /** one name per category level, outer → inner */
  names: string[];
  /** one label array per level, all the same length */
  levels: string[][];
}

/** One chart series. Colours are 6-digit hex ("4E79A7"); null = invisible (Gantt gaps, waterfall bases). */
interface ChartSeries {
  name: string;
  /** category charts: one value per category */
  values?: (number | null)[];
  /** scatter / bubble only */
  x?: (number | null)[];
  y?: (number | null)[];
  /** bubble only: bubble area */
  size?: (number | null)[];
  color: string | null;
  /** per-point colours (colour = a category level or a measure) */
  pointColors?: (string | null)[];
  /** combo charts: this series' chart type */
  type?: "bar" | "line" | "area";
  /** drawn on the secondary value axis */
  secondary?: boolean;
  /** line series: false = markers only */
  line?: boolean;
  /** fill / line opacity 0–1 (Tableau's Color → Opacity) */
  alpha?: number;
  marker?: boolean;
  markerSymbol?: string;
  markerSize?: number;
  labels?: boolean;
  labelParts?: { value?: boolean; category?: boolean; percent?: boolean };
  labelNumFmt?: string;
  /** scatter: "value from cells" label text per point */
  labelTexts?: string[];
  /** packed bubbles: Tableau's tooltip text per point */
  tooltips?: string[];
  /** a reference line drawn as a flat line series */
  refLine?: RefLineStyle;
}

/** a reference line: its value, label (an Excel number format, so Excel prints "Average" / "Avg. $1.2M"), line style */
interface RefLineStyle {
  value: number;
  /** null = no label */
  labelFmt: string | null;
  color: string;
  alpha: number;
  width: number;
  dash: boolean;
  /** line switched off in the workbook (the label may still show) */
  hidden: boolean;
  font: { color?: string; bold?: boolean };
}

/** A renderer-neutral Excel chart, produced by buildExcelChartSpecs() and drawn by ExcelChartWriter. */
interface ChartSpec {
  kind: ChartKind;
  font: { name: string; size: number; color: string };
  /** chart area colour; null = see-through (Worksheet shading "None") */
  background: string | null;
  gridlines: boolean;
  legend?: boolean;
  /** bar charts: "col" = vertical, "bar" = horizontal */
  barDir?: "col" | "bar";
  stacked?: boolean;
  gapWidth?: number;
  categories?: ChartCategories;
  categoryTitle?: string;
  series: ChartSeries[];
  numFmt?: string;
  valueTitle?: string;
  secondaryNumFmt?: string;
  secondaryTitle?: string;
  /** bubble size column (data sheet header / number format) */
  sizeTitle?: string;
  sizeNumFmt?: string;
  /** packed bubbles: no axes, plot area fills the chart */
  axesHidden?: boolean;
  /** packed bubbles: extent of the packing in radius units (largest bubble radius = 1), centre cx / cy */
  packed?: { w: number; h: number; cx: number; cy: number };
  /** scatter x axis */
  xNumFmt?: string;
  xTitle?: string;
  /** fixed value axis bounds (Gantt dates) */
  valueMin?: number;
  valueMax?: number;
  /** false = let Excel auto-scale instead of starting the axis at zero */
  includeZero?: boolean;
  valueAxisHidden?: boolean;
  /** the worksheet hides these axes / lines (Show Header off, Format → Lines) */
  categoryAxisHidden?: boolean;
  /** category labels' rotation in degrees (Tableau's header text orientation), 0 = horizontal */
  categoryRotation?: number;
  /** category labels shown on the axis (names cut short like Tableau's headers), set by the writer */
  categoryShown?: string[] | null;
  secondaryAxisHidden?: boolean;
  xAxisHidden?: boolean;
  /** scatter: grid lines along the x axis */
  xGridlines?: boolean;
  /** false = no axis ruler on the category axis */
  axisLine?: boolean;
  /** Format → Axis → Numbers: tick formats as set in the workbook (else derived from the value format) */
  valueAxisNumFmt?: string;
  secondaryAxisNumFmt?: string;
  xAxisNumFmt?: string;
  /** Edit Axis: tick spacing; scatter x axis range */
  valueMajorUnit?: number;
  xMajorUnit?: number;
  xMin?: number;
  xMax?: number;
  /** the worksheet's mark label font (hex colour) and, for bars, Excel's label position */
  labelFont?: { name?: string; size?: number; color?: string; bold?: boolean };
  labelPos?: "inBase" | "ctr" | "inEnd" | "outEnd";
  /** false = labels may overlap (Tableau: "Allow labels to overlap other marks"); otherwise overlapping ones are left out */
  labelCull?: boolean;
  /** reference lines across horizontal bars (drawn as vertical lines over the bars) */
  refLines?: RefLineStyle[];
  /** doughnut: the hole as % of the ring; the text in the hole (lines of runs) */
  holeSize?: number;
  centerLabel?: { text: string; bold?: boolean; size?: number; color?: string; font?: string }[][];
  /** line chart drawn as a box plot: up/down bars + high-low lines */
  boxPlot?: { color: string };
  /** separate panes of one worksheet, stacked vertically */
  paneIndex?: number;
  paneCount?: number;
  /** set by buildExcelChartSpecs */
  name?: string;
  rolesSource?: RolesSource;
  /** 100% stacked bars / areas */
  percent?: boolean;
  /** bars of a synchronized dual axis drawn over each other on one axis (progress bar over its track) */
  overlap?: boolean;
  /** dual-axis bars: the secondary axis's bars narrower, inside the primary ones (bar in bar) */
  secondaryGapWidth?: number;
  /** Synchronize Axis: the secondary value axis takes the primary's range (one scale for both) */
  secondarySync?: boolean;
  /** line chart: high-low lines joining each category's markers (dumbbell) */
  hiLowLines?: { color: string };
  /** horizontal bullet graph: one target tick per category, drawn over the bars */
  targets?: { name: string; values: (number | null)[]; color: string };
  /** horizontal bars: marks drawn over the category bands – values on the value axis, cats = band index (null = a gap) */
  overlay?: { name: string; color: string; values: (number | null)[]; cats: (number | null)[]; marker?: string; markerSize?: number;
              line?: boolean; lineWidth?: number }[];
  /** scatter: Edit Axis → Reversed on the x axis */
  xReversed?: boolean;
  /** back-to-back bars (butterfly): one wing negated, so an axis Tableau reversed for it stays as it is */
  mirrored?: boolean;
  /** Edit Axis → Reversed on the value axis (bump charts: rank 1 at the top) */
  valueReversed?: boolean;
  /** how this chart stands for the Tableau visual when it is not a direct Excel equivalent (conversion report) */
  conversion?: { strategy: "NATIVE" | "CONSTRUCTED" | "APPROXIMATE"; output?: string; note?: string; fidelity?: string };
}

/* ── chart writer (charts/writer) ────────────────────────────────────────── */

/** writeChartData(): the cell ranges a chart's XML points at. */
interface ChartRefs {
  /** category range (category charts) */
  cat?: string;
  series: { tx: string; val?: string; x?: string; y?: string; size?: string; lbl?: string }[];
  /** first free row on the data sheet after this chart */
  nextRow: number;
}

/** A chart placed on the dashboard sheet, for injectCharts(). */
interface ChartJob {
  spec: ChartSpec;
  refs: ChartRefs;
  name: string;
  /** top-left anchor cell (0-based) and pixel offset inside it */
  col: number;
  row: number;
  colOffPx?: number;
  rowOffPx?: number;
  widthPx: number;
  heightPx: number;
  /** the chart block's first row under its title; which of how many stacked panes */
  top?: number;
  pane?: number;
  panes?: number;
  item?: any;
  /** a pie / donut: finished chart XML, its label shapes part, its tooltip wedges (px inside the chart) */
  xml?: string;
  shapes?: string | null;
  tips?: { x: number; y: number; w: number; h: number; prst: string; adj?: number[]; text: string; link: string }[];
}

/* ── export (export/export.js) ───────────────────────────────────────────── */

/** One block laid out on the dashboard sheet: a worksheet (table, KPI card, chart or image) or a
 *  filter / parameter value list. */
/** a picture packaged in the .twbx (Image/logo.png) */
interface WorkbookImage {
  data: Uint8Array;
}

/** a dashboard image object to draw: the file, its format / size, its zone and link */
interface ImageBlock {
  file: WorkbookImage;
  info: import("./export/images.js").ImageInfo;
  /** the image zone (Fit Image, Center Image, link); an icon sheet gives only what applies */
  zone: Partial<DashboardZone>;
  url: string | null;
  widthPx: number;
  heightPx: number;
}

interface ExportItem {
  type: "worksheet" | "filterValue" | "text" | "image" | "legend";
  name: string;
  visualName: string;
  /** dashboard zone → grid position (export/layout.js) */
  layout: any;
  rowCount: number;
  /* filter / parameter value lists */
  filterName?: string;
  values?: any[];
  isParameter?: boolean;
  originalData?: SummaryData;
  /* worksheets */
  isKPI?: boolean;
  vm?: ViewModel;
  visualModel?: VisualModel;
  /** chart / image: pixel size from the dashboard zone */
  box?: { widthPx: number; heightPx: number; gridW: number; rows: number; [key: string]: any } | null;
  fixedGridW?: number | null;
  columns?: number[];
  /* placement, added by the layout pass */
  gridRow?: number;
  gridCol?: number;
  gridW?: number;
  allocatedRows?: number;
  /** grouped tables: data rows kept visible – blocks beside the table use them (setTableVisibleRows) */
  visibleRows?: number;
  /** dashboard text box (banner, title, note) drawn in its zone (export/kpi-card.js) */
  textCard?: import("./export/kpi-card.js").KpiCard;
  /** chart sharing a KPI card's columns (Tableau's number + trend tile) */
  pairedCard?: boolean;
  /** KPI tile rebuilt from its Tableau label (export/kpi-card.js); null = label | value table */
  kpiCard?: import("./export/kpi-card.js").KpiCard | null;
  /** relative widths of the columns the block needs in its zone (table fields, KPI tiles) */
  split?: number[];
  /** a split block's part boundaries: the column (from its first) where each next part starts */
  splitCuts?: number[] | null;
  /** a table's fields on the grid: each one's first column (from the table's first) and how many it spans */
  fieldCols?: { offset: number; span: number }[] | null;
  /** a table's column widths in px (Tableau's own, else from its content) */
  splitPx?: number[];
  /** dashboard image object (logo, icon) with its own zone */
  image?: ImageBlock;
  /** a pie's colour legend: its title and swatches, the formatter of its sheet */
  legend?: { title: string; items: { text: string; color: string | null }[] };
  fmt?: SheetFormatter;
}

/* ── browser APIs the panel feature-detects (File System Access; Chromium only) ── */
interface Window {
  showOpenFilePicker?: (options?: any) => Promise<any[]>;
  showSaveFilePicker?: (options?: any) => Promise<any>;
}

/* ── conversion report (visual/semantics.js, visual/strategy.js, export/report.js) ─────────────────────── */

/** describeVisual(): the Tableau visual in Tableau's terms, and the evidence. */
interface VisualSemantics {
  /** "Stacked Bar", "Bump Chart", "Filled Map" … */
  visual: string;
  family: VisualType;
  confidence: "high" | "medium" | "low";
  evidence: string[];
  /** shelf / encoding → field labels ("rows": ["Category"], "color": ["SUM(Profit)"]) */
  encodings: Record<string, string[]>;
  calculations: { name: string; kind: string }[];
  flags: Record<string, boolean>;
}

/** conversionOf(): how one visual was exported. */
interface Conversion {
  strategy: "NATIVE" | "CONSTRUCTED" | "APPROXIMATE" | "TABLE_FALLBACK" | "IMAGE_FALLBACK" | "UNSUPPORTED" | "FAILED";
  excelOutput: string;
  fidelity: string;
  reason: string;
}

interface ReportVisual extends Conversion {
  worksheet: string;
  semantics: VisualSemantics;
}
