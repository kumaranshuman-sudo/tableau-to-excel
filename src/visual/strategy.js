/* How each visual was converted, in the terms of the conversion report:
 *   NATIVE          a genuine Excel equivalent (chart type or formatted cells)
 *   CONSTRUCTED     Excel has no such chart: built from Excel primitives that keep its meaning
 *   APPROXIMATE     an Excel chart that keeps the analytical meaning, with a visible difference (named)
 *   TABLE_FALLBACK  no Excel chart would read correctly: the visual's data as a table
 *   IMAGE_FALLBACK  a picture of the visual (Tableau-rendered), not editable
 *   UNSUPPORTED     not exported
 *   FAILED          exporting it went wrong
 * A correct table always beats a chart that would mislead: anything not provably meaningful falls back. */
import { VISUAL_TYPES } from "../config.js";

export const STRATEGY = Object.freeze({
  NATIVE: "NATIVE", CONSTRUCTED: "CONSTRUCTED", APPROXIMATE: "APPROXIMATE", TABLE_FALLBACK: "TABLE_FALLBACK",
  IMAGE_FALLBACK: "IMAGE_FALLBACK", UNSUPPORTED: "UNSUPPORTED", FAILED: "FAILED"
});

/** report order, strongest first */
export const STRATEGY_ORDER = ["NATIVE", "CONSTRUCTED", "APPROXIMATE", "TABLE_FALLBACK", "IMAGE_FALLBACK", "UNSUPPORTED", "FAILED"];

const FIDELITY = { NATIVE: "High", CONSTRUCTED: "High", APPROXIMATE: "Medium", TABLE_FALLBACK: "Exact data",
                   IMAGE_FALLBACK: "Picture only", UNSUPPORTED: "None", FAILED: "None" };

/** an Excel chart in words ("Stacked column chart", "Combo chart: column + line, secondary axis") @param {ChartSpec} spec */
export function chartOutputName(spec) {
  if (spec.conversion && spec.conversion.output) return spec.conversion.output;
  const dir = spec.barDir === "bar" ? "bar" : "column";
  const grouping = spec.percent ? "100% stacked " : spec.stacked ? "stacked " : "";
  const kinds = [...new Set(spec.series.filter(s => !s.refLine).map(s => s.type || spec.kind))];
  const extras = spec.series.some(s => s.refLine) || (spec.refLines && spec.refLines.length) ? " with reference line" : "";
  switch (spec.kind) {
    case "bar": return cap(`${grouping || "clustered "}${dir} chart`) + extras;
    case "line": return (spec.series.filter(s => !s.refLine).every(s => s.line === false) ? "Line chart (markers only)" : "Line chart") + extras;
    case "area": return cap(`${grouping}area chart`) + extras;
    case "combo": return (kinds.length === 1 && kinds[0] === "bar" ? `Bar-in-bar ${dir} chart (two axes${spec.secondarySync ? ", one scale" : ""})`
                          : `Combo chart: ${kinds.map(k => k === "bar" ? dir : k).join(" + ")}` +
                            (spec.series.some(s => s.secondary) ? `, secondary axis${spec.secondarySync ? " (synchronized)" : ""}` : "")) + extras;
    case "pie": return "Pie chart";
    case "doughnut": return "Doughnut chart";
    case "scatter": return spec.series.some(s => s.line) ? "Scatter chart with lines" : "Scatter chart" + extras;
    case "bubble": return "Bubble chart";
    case "treemap": return "Treemap (Excel 2016+)";
    default: return cap(`${spec.kind} chart`);
  }
}

/**
 * The conversion of one exported visual, once the export has drawn it (renderer, status and chart specs final).
 * @param {VisualModel} visualModel @param {VisualSemantics} [semantics]
 * @returns {Conversion}
 */
export function conversionOf(visualModel, semantics) {
  const make = (strategy, excelOutput, reason, fidelity) => ({ strategy, excelOutput, reason, fidelity: fidelity || FIDELITY[strategy] });
  const renderer = visualModel.renderer;
  const why = String(visualModel.statusReason || "");
  if (visualModel.status === "warning" && /injection failed/i.test(why)) return make(STRATEGY.FAILED, "Nothing", why);
  if (visualModel.pie) {
    const pie = visualModel.pie;
    const donut = pie.hole || pie.inner;
    return make(STRATEGY.NATIVE, `${donut ? "Doughnut" : "Pie"} chart${pie.inner ? " (two rings)" : ""}`,
      "Direct Excel equivalent: Tableau's slices, colours and labels" + (pie.hole ? ", the hole with its text" : "") + "; tooltips as hover text");
  }
  if (renderer === "excel-chart") {
    const specs = visualModel.chartSpecs || [];
    const rank = s => STRATEGY_ORDER.indexOf((s.conversion && s.conversion.strategy) || defaultStrategy(s));
    const weakest = specs.slice().sort((a, b) => rank(b) - rank(a))[0];
    if (!weakest) return make(STRATEGY.FAILED, "Nothing", "no chart was produced");
    const strategy = (weakest.conversion && weakest.conversion.strategy) || defaultStrategy(weakest);
    const output = [...new Set(specs.map(chartOutputName))].join("; ") + (specs.length > 1 ? ` (${specs.length} panes)` : "");
    const note = specs.map(s => s.conversion && s.conversion.note).filter(Boolean)[0];
    return make(strategy, output, note || (strategy === STRATEGY.APPROXIMATE ? "Excel's closest chart; the layout differs" : "Direct Excel equivalent"),
                weakest.conversion && weakest.conversion.fidelity);
  }
  if (renderer === "tableau-image") return make(STRATEGY.IMAGE_FALLBACK, "Picture (Tableau-rendered) + its data on the Visual Data sheet",
                                                  why || "no editable Excel equivalent");
  if (renderer === "data-fallback") return make(STRATEGY.TABLE_FALLBACK, "Data table", fallbackReason(visualModel, semantics, why));
  if (renderer === "cell") {
    const vm = visualModel.viewModel;
    if (vm && vm.lossless === false) return make(STRATEGY.TABLE_FALLBACK, "Data table", vm.fallbackReason || "its marks encode values the cells would not show");
    if (vm && vm.matrix) return make(STRATEGY.CONSTRUCTED, "Matrix of cells with Tableau's colour per cell",
                                     "Excel has no heat-map chart: a matrix with the same rows, columns and colours");
    if (vm && vm.markTable) return make(STRATEGY.CONSTRUCTED, "Cells drawn like Tableau's marks (fills, arrows, data bars)",
                                        "a table built from marks: each column styled like its marks");
    if (visualModel.type === VISUAL_TYPES.KPI) return make(STRATEGY.NATIVE, "KPI cells", "Direct Excel equivalent: the values in Tableau's format");
    if (visualModel.type === VISUAL_TYPES.HEATMAP) return make(STRATEGY.CONSTRUCTED, "Table with Tableau's colour per row", "Excel has no heat-map chart");
    return make(STRATEGY.NATIVE, "Formatted cell table", "Direct Excel equivalent: Tableau's table as cells");
  }
  return make(STRATEGY.UNSUPPORTED, "Nothing", why || "no renderer");
}

/** @param {ChartSpec} spec */
function defaultStrategy(spec) {
  if (spec.kind === "bubble" && spec.packed) return STRATEGY.APPROXIMATE;
  return STRATEGY.NATIVE;
}

/** why a visual became a data table, in words a reader of the report understands */
function fallbackReason(visualModel, semantics, why) {
  const t = visualModel.type;
  const visual = semantics ? semantics.visual : "";
  if (t === VISUAL_TYPES.MAP || t === VISUAL_TYPES.MAP_FILLED) {
    const kind = {
      "Density Map": "density marks have no Excel equivalent",
      "Polygon Map": "custom polygons (hex bins, shapes) cannot be drawn by an Excel chart",
      "Flow Map": "paths between locations cannot be drawn on an Excel map",
      "Map with Layers": "several map layers on one map have no Excel equivalent",
      "Pies on a Map": "pies placed on a map have no Excel equivalent",
      "Donut / Pie built on map layers": "a donut built from map layers (no geography) – Excel's charts cannot stack the layers",
      "Filled Map": "Excel's filled map cannot reproduce Tableau's geography and colour scale reliably",
      "Symbol Map": "Excel has no symbol map: marks sized and coloured at locations"
    }[visual] || "Excel cannot reproduce this map";
    return `${kind}; every mark exported with its location fields and values`;
  }
  if (t === VISUAL_TYPES.CUSTOM) return "a viz extension draws this sheet; Excel has no equivalent";
  const m = why.match(/native chart not possible \(([^)]*)\)/i);
  if (m) return `no faithful Excel chart: ${m[1]}`;
  return why.replace(/ – exported as data$/, "") || "no meaningful Excel equivalent";
}

function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
