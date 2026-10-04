/* Entry point: visual model → Excel chart specs (throws when not representable). */
import { tvCartesianSpecs, tvHistogramSpecs } from "./cartesian.js";
import { tvMarkToken } from "./common.js";
import { tvRoles } from "./roles.js";
import { tvBoxPlotSpec, tvGanttSpec, tvPieSpec, tvScatterSpec, tvTreemapSpec, tvWaterfallSpec } from "./specialized.js";
import { VISUAL_TYPES } from "../../config.js";

/**
 * @param {VisualModel} visualModel
 * @param {FormatModel | null} model
 * @returns {ChartSpec[]}
 */
export function buildExcelChartSpecs(visualModel, model) {
  const vm = visualModel.viewModel;
  if (!vm || !vm.rows.length) throw new Error("visual has no data rows");
  const roles = tvRoles(vm, model, visualModel.source && visualModel.source.visualSpec);
  const markToken = visualModel.metadata.markToken || "";
  /** @type {ChartContext} */
  const ctx = { vm, roles, markToken, type: visualModel.type };
  if (roles.source === "summary" && /^(circle|shape|square)$/.test(markToken)) {
    throw new Error("mark layout unknown – load the workbook file for this visual");
  }
  let specs;
  if (visualModel.type === VISUAL_TYPES.HISTOGRAM) specs = tvHistogramSpecs(ctx);
  else if (visualModel.type === VISUAL_TYPES.WATERFALL) specs = tvWaterfallSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.BOXPLOT) specs = tvBoxPlotSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.GANTT) specs = tvGanttSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.TREEMAP) specs = tvTreemapSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.PIE) {
    ctx.doughnut = roles.panes.filter(p => tvMarkToken(p.markClass) === "pie").length >= 2;
    specs = tvPieSpec(ctx);
  } else if (visualModel.type === VISUAL_TYPES.SCATTER && roles.rows.values.length && roles.cols.values.length) {
    specs = tvScatterSpec(ctx);
  } else {
    specs = tvCartesianSpecs(ctx);
  }
  specs.forEach(s => { s.name = visualModel.metadata.worksheetName; s.rolesSource = roles.source; });
  return specs;
}
