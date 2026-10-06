/* View model + classification → visual model, and renderer selection. */
import { FORMAT_CONFIG, VISUAL_TYPES } from "../config.js";
import { classifyVisualType, isMarkTable, resolveVisualMarks } from "./classify.js";
import { checkLossless, pivotMatrix } from "./matrix.js";
import { buildViewModel } from "./view-model.js";

/**
 * @param {FormatModel | null} model
 * @param {string} sheetName
 * @param {SummaryData} summary
 * @param {{ dashboardName?: string, displayName?: string, visualSpec?: any, visualSpecError?: any, layout?: any }} [opts]
 * @returns {VisualModel}
 */
export function buildVisualModel(model, sheetName, summary, opts = {}) {
  const vm = buildViewModel(model, sheetName, summary, opts);
  const spec = opts.visualSpec || null;
  const marks = resolveVisualMarks(spec, vm, model);
  const evidence = [];
  const type = classifyVisualType(spec, vm, model, evidence);
  vm.markTable = (type === VISUAL_TYPES.TABLE || type === VISUAL_TYPES.KPI) && isMarkTable(vm, model);
  // drawn as cells: a heat map / crosstab as Tableau's matrix, and never a table that drops what its marks show
  if (type === VISUAL_TYPES.HEATMAP || type === VISUAL_TYPES.TABLE) {
    if (pivotMatrix(vm, model)) evidence.push("dimensions on Rows and Columns, one mark per cell: a matrix");
    else checkLossless(vm, model);
  }
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
    formatModel: model || null,
    metadata: {
      worksheetName: sheetName,
      markClass: vm.fmt && vm.fmt.colorEncoding ? ((vm.fmt.colorEncoding() || {}).markClass || null) : null,
      markToken: marks.tokens[0] || "",
      markTokens: marks.tokens,
      markSource: marks.source,
      visualSpecAvailable: !!spec,
      visualSpecError: opts.visualSpecError ? String(opts.visualSpecError.message || opts.visualSpecError) : null,
      visualSpecKeys: spec && typeof spec === "object" ? Object.keys(spec) : [],
      imageApiAvailable: typeof tableau !== "undefined" && tableau.extensions &&
        typeof tableau.extensions.createVizImageAsync === "function"
    },
    viewModel: vm,
    diagnostics: [],
    evidence
  };
}

/* cell-native → cells; mappable charts → native Excel chart; the rest → Tableau image → data table */
/** @param {VisualModel} visualModel @returns {RendererDecision} */
export function chooseVisualRenderer(visualModel) {
  const type = visualModel.type;
  if (VISUAL_RENDERERS.cellTypes.has(type)) return VISUAL_RENDERERS.cell;
  if (FORMAT_CONFIG.nativeCharts && VISUAL_RENDERERS.chartTypes.has(type)) return VISUAL_RENDERERS.chart;
  return imageOrFallback(visualModel);
}

/** @param {VisualModel} visualModel @param {string} [why] @returns {RendererDecision} */
export function imageOrFallback(visualModel, why) {
  const prefix = why ? why + "; " : "";
  if (VISUAL_RENDERERS.imageTypes.has(visualModel.type) && visualModel.metadata.imageApiAvailable) {
    return { ...VISUAL_RENDERERS.image, reason: prefix + VISUAL_RENDERERS.image.reason };
  }
  const reason = !visualModel.metadata.imageApiAvailable && VISUAL_RENDERERS.imageTypes.has(visualModel.type)
    ? "Tableau image API unavailable – exported as data"
    : `${visualModel.type.toLowerCase()} visuals cannot be drawn in Excel – exported as data`;
  return { ...VISUAL_RENDERERS.fallback, reason: prefix + reason };
}

export const VISUAL_RENDERERS = Object.freeze({
  cell: Object.freeze({ renderer: "cell", status: "success", reason: "cell-native visual" }),
  chart: Object.freeze({ renderer: "excel-chart", status: "pending", reason: "native Excel chart" }),
  image: Object.freeze({ renderer: "tableau-image", status: "pending", reason: "Tableau image renderer selected" }),
  cellTypes: /** @type {Set<VisualType>} */ (new Set([VISUAL_TYPES.TABLE, VISUAL_TYPES.KPI, VISUAL_TYPES.HEATMAP])),
  chartTypes: /** @type {Set<VisualType>} */ (new Set([VISUAL_TYPES.BAR, VISUAL_TYPES.COLUMN, VISUAL_TYPES.LINE, VISUAL_TYPES.AREA,
                       VISUAL_TYPES.PIE, VISUAL_TYPES.SCATTER, VISUAL_TYPES.COMBO, VISUAL_TYPES.HISTOGRAM,
                       VISUAL_TYPES.WATERFALL, VISUAL_TYPES.BOXPLOT, VISUAL_TYPES.GANTT, VISUAL_TYPES.TREEMAP,
                       VISUAL_TYPES.BUBBLE])),
  // createVizImageAsync draws bar / line / area / square / circle / text marks only (no pie, gantt or polygon
  // marks). Maps – symbol or filled – are neither: without the basemap they don't read as maps, so they are
  // exported as a table of their marks (data-fallback)
  imageTypes: /** @type {Set<VisualType>} */ (new Set([VISUAL_TYPES.BAR, VISUAL_TYPES.COLUMN, VISUAL_TYPES.LINE, VISUAL_TYPES.AREA, VISUAL_TYPES.SCATTER,
                       VISUAL_TYPES.COMBO, VISUAL_TYPES.HISTOGRAM, VISUAL_TYPES.TREEMAP, VISUAL_TYPES.BUBBLE])),
  fallback: Object.freeze({ renderer: "data-fallback", status: "warning", reason: "visual renderer not implemented yet" })
});
