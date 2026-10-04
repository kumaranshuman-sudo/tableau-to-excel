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
}

interface ParsedStyle {
  /** element name ("cell", "header", "mark", "axis" …) → its rules */
  rules: Record<string, StyleRule[]>;
  encodings: ColorEncoding[];
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
  runningTotals?: FieldRef[];
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
}

interface DashboardModel {
  title: { text: string; runs: TextRun[] } | null;
  style: ParsedStyle;
  zones: DashboardZone[];
  width?: number;
  height?: number;
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
  dataType?: string;
  ref: FieldRef | null;
  /** a Measure Values measure turned into its own column */
  pivoted?: boolean;
  /** dimension (row/column header) rather than a measure */
  isHeader?: boolean;
  label?: string;
  labelProps?: Record<string, any>;
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
  dashboardName?: string;
  kind: "chart" | "table";
  /** all header columns (hidden ones too), outer → inner */
  headerOrder: number[];
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
  /* set by the export once a renderer is chosen */
  renderer?: RendererName;
  status?: string;
  statusReason?: string;
  chartSpecs?: ChartSpec[];
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
  roles: Roles;
  markToken: string;
  type: VisualType;
  doughnut?: boolean;
}

type ChartKind = "bar" | "line" | "area" | "combo" | "pie" | "doughnut" | "scatter" | "treemap";

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
  /** scatter only */
  x?: (number | null)[];
  y?: (number | null)[];
  color: string | null;
  /** per-point colours (colour = a category level or a measure) */
  pointColors?: (string | null)[];
  /** combo charts: this series' chart type */
  type?: "bar" | "line" | "area";
  /** drawn on the secondary value axis */
  secondary?: boolean;
  /** line series: false = markers only */
  line?: boolean;
  marker?: boolean;
  markerSymbol?: string;
  markerSize?: number;
  labels?: boolean;
  labelParts?: { value?: boolean; category?: boolean; percent?: boolean };
  labelNumFmt?: string;
  /** scatter: "value from cells" label text per point */
  labelTexts?: string[];
}

/** A renderer-neutral Excel chart, produced by buildExcelChartSpecs() and drawn by ExcelChartWriter. */
interface ChartSpec {
  kind: ChartKind;
  font: { name: string; size: number; color: string };
  background: string;
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
  /** scatter x axis */
  xNumFmt?: string;
  xTitle?: string;
  /** fixed value axis bounds (Gantt dates) */
  valueMin?: number;
  valueMax?: number;
  /** false = let Excel auto-scale instead of starting the axis at zero */
  includeZero?: boolean;
  valueAxisHidden?: boolean;
  /** line chart drawn as a box plot: up/down bars + high-low lines */
  boxPlot?: { color: string };
  /** separate panes of one worksheet, stacked vertically */
  paneIndex?: number;
  paneCount?: number;
  /** set by buildExcelChartSpecs */
  name?: string;
  rolesSource?: RolesSource;
}

/* ── chart writer (charts/writer) ────────────────────────────────────────── */

/** writeChartData(): the cell ranges a chart's XML points at. */
interface ChartRefs {
  /** category range (category charts) */
  cat?: string;
  series: { tx: string; val?: string; x?: string; y?: string; lbl?: string }[];
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
  item?: any;
}

/* ── export (export/export.js) ───────────────────────────────────────────── */

/** One block laid out on the dashboard sheet: a worksheet (table, KPI card, chart or image) or a
 *  filter / parameter value list. */
interface ExportItem {
  type: "worksheet" | "filterValue";
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
}

/* ── browser APIs the panel feature-detects (File System Access; Chromium only) ── */
interface Window {
  showOpenFilePicker?: (options?: any) => Promise<any[]>;
  showSaveFilePicker?: (options?: any) => Promise<any>;
}
