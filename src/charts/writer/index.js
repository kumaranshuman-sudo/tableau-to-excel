/* ══════════════════════════════════════════════════════════════════════════
 * EXCEL NATIVE CHART WRITER
 * ──────────────────────────────────────────────────────────────────────────
 * ExcelJS cannot author charts, so native charts are produced in two steps:
 *   1. writeChartData(ws, spec, row) – the chart's source data is written to a
 *      (hidden) data sheet with ExcelJS, so every chart stays editable and is
 *      linked to real cells.
 *   2. injectCharts(buffer, opts) – after workbook.xlsx.writeBuffer() the XLSX
 *      package is opened with JSZip and DrawingML chart parts are added to the
 *      dashboard sheet (re-using the drawing ExcelJS made for images, if any).
 *
 * Chart spec (produced by charts/model):
 *   { kind: "bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"bubble"|"combo"|"treemap",
 *     barDir: "col"|"bar", stacked, gapWidth, categories: { names:[], levels:[[…]] },
 *     series: [{ name, type, color (null = invisible), values|x+y(+size), pointColors, secondary,
 *                line, marker, markerSymbol, markerSize, labels, labelNumFmt }],
 *     numFmt, secondaryNumFmt, xNumFmt, valueTitle, secondaryTitle, valueMin, valueMax,
 *     boxPlot: { color } (line chart → up/down bars + high-low lines),
 *     categoryTitle, legend, gridlines, font: { name, size, color },
 *     axesHidden + aspect (symbol maps) / packed (packed bubbles): axis ranges fitted to the chart size }
 * "treemap" is written as an Excel 2016+ chartex part with an older-Excel fallback shape.
 * Colours are 6-digit hex ("4E79A7"); ARGB ("FF4E79A7") is accepted too.
 * ══════════════════════════════════════════════════════════════════════════ */
import { writeChartData } from "./chart-data.js";
import { chartXml } from "./drawingml.js";
import { chartExXml } from "./chartex.js";
import { injectCharts } from "./package.js";

export const ExcelChartWriter = { writeChartData, injectCharts, chartXml, chartExXml };
