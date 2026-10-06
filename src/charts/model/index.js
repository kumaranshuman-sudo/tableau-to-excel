/* Entry point: visual model → Excel chart specs (throws when not representable). */
import { tvApplyWorkbookAxes } from "./axes.js";
import { tvApplyReferenceLines } from "./reflines.js";
import { tvCartesianSpecs, tvHistogramSpecs } from "./cartesian.js";
import { tvMarkToken } from "./common.js";
import { tvRoles } from "./roles.js";
import { tvBoxPlotSpec, tvGanttSpec, tvPackedBubbleSpec, tvPieSpec, tvScatterSpec, tvTreemapSpec, tvWaterfallSpec } from "./specialized.js";
import { VISUAL_TYPES } from "../../config.js";

/* ══════════════════════════════════════════════════════════════════════════
 * TABLEAU VISUAL → EXCEL CHART SPECS
 * ──────────────────────────────────────────────────────────────────────────
 * buildExcelChartSpecs(visualModel, model) reads
 *   – live summary data (already pivoted / merged / sorted by buildViewModel)
 *   – TWB shelves, panes, marks, encodings, colours, number formats, labels
 *   – the live visual specification when no workbook was loaded
 * and returns renderer-neutral chart specs (ChartSpec, src/types.d.ts) for charts/writer.
 * It throws when a visual cannot be represented faithfully, so the caller
 * can fall back to the Tableau image renderer or the data table.
 * ══════════════════════════════════════════════════════════════════════════ */

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
  else if (visualModel.type === VISUAL_TYPES.BUBBLE) specs = tvPackedBubbleSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.PIE) {
    ctx.doughnut = roles.panes.filter(p => tvMarkToken(p.markClass) === "pie").length >= 2;
    specs = tvPieSpec(ctx);
  } else if (visualModel.type === VISUAL_TYPES.SCATTER && roles.rows.values.length && roles.cols.values.length) {
    specs = tvScatterSpec(ctx);
  } else {
    specs = tvCartesianSpecs(ctx);
  }
  specs.forEach(s => { s.name = visualModel.metadata.worksheetName; s.rolesSource = roles.source; tvApplyWorkbookAxes(s, ctx); tvApplyReferenceLines(s, ctx); });
  return specs;
}
