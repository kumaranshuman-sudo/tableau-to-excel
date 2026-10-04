/* The export: Tableau dashboard → formatted XLSX with native charts. */
import ExcelJS from "exceljs";
import { buildExcelChartSpecs } from "../charts/model/index.js";
import { FORMAT_CONFIG, TABLEAU_DEFAULTS, TF_ELEMENTS, VISUAL_TYPES } from "../config.js";
import { writeDashboardTitle, writeIndividualFilterTable } from "./cell-writers.js";
import { CHART_DATA_SHEET, EXCEL_ROW_PX, PX_PER_COL, PX_PER_ROW, ROW_GAP, ROW_GROUP_THRESHOLD, buildLayoutMap, graphicBox, makeRangeTracker, resolveCollisions, viewModelHeight } from "./layout.js";
import { extractFilterValuesPerField, fetchAllSheetsData, isFilterValueWorksheet } from "./sheet-data.js";
import { applyAutoFilters, setColumnWidths, writeKPICardStacked, writeRegularTable, writeTableauTitle } from "./visual-writers.js";
import { tfDashboardTitleRuns } from "../twb/dashboard-text.js";
import { tfCollect, tfMerge } from "../twb/formatter.js";
import { appendExportStatus, setExportStatus, updateVisualStatus } from "../ui/status.js";
import { chooseSaveTarget, ensureFormatModel, getTitleMap } from "../ui/workbook-store.js";
import { isKPIViewModel } from "../visual/classify.js";
import { renderTableauImage } from "../visual/image-renderer.js";
import { VISUAL_RENDERERS, buildVisualModel, chooseVisualRenderer, imageOrFallback } from "../visual/visual-model.js";
import { checkWorkbookMatch, describeWorkbookMatch } from "../visual/workbook-match.js";
import { ExcelChartWriter } from "../charts/writer/index.js";

export async function exportToExcel() {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  const sheets = dashboard.worksheets;
  const btn = /** @type {HTMLButtonElement} */ (document.getElementById("export_button"));
  btn.disabled = true;

  try {
    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
    const exportFileName = `${(dashboard.name || "Dashboard Export").replace(/[\\\/\*\?\[\]:]/g, "").slice(0, 31)}_${stamp}.xlsx`;
    if (!(await ensureFormatModel()) && !window.confirm(
        "No Tableau workbook is loaded for this dashboard.\n\n" +
        "Without it the extension cannot tell which sheets are charts, so they will be exported as tables.\n\n" +
        "Cancel, then click 📁 Load Workbook to export real charts – or OK to export tables anyway.")) {
      setExportStatus("Export stopped – click 📁 Load Workbook first so charts are exported as charts", true);
      return;
    }
    setExportStatus("Choose where to save the Excel file…");
    const saveTarget = await chooseSaveTarget(exportFileName);       // opens in the loaded workbook's folder
    if (saveTarget === "cancelled") { setExportStatus("Export cancelled – no file was saved"); return; }

    const filterValuesMap = await extractFilterValuesPerField(sheets);
    console.log("📊 Filter values per field:", filterValuesMap);

    const titleMap = getTitleMap();
    console.log(`[Export] Using ${Object.keys(titleMap).length} titles`);

    const fmtModel = await ensureFormatModel();
    console.log(`[Export] Format model: ${fmtModel ? Object.keys(fmtModel.sheets).length + " sheets" : "none – load the workbook for exact formatting"}`);
    const workbookWarning = fmtModel ? describeWorkbookMatch(checkWorkbookMatch(fmtModel, dashboard))
      : "⚠ Workbook not loaded – charts can only be recognised from the workbook, so they may be exported as tables. Click 📁 Load Workbook, then export again";
    if (workbookWarning) console.warn("[Export]", workbookWarning);

    const layoutMap = buildLayoutMap(dashboard.objects || [], titleMap);

    // ── 1. Build DZV map ──
    const dzvMap = {};
    (dashboard.objects || [])
      .filter(obj => obj.type === "worksheet")
      .forEach(obj => {
        dzvMap[obj.name] = obj.isVisible;
      });
    console.log("[DZV] Visibility map:", dzvMap);

    // ── 2. Fetch all sheets in parallel ──
    setExportStatus(`Reading ${sheets.length} worksheets from Tableau…`);
    const allSheetsData = await fetchAllSheetsData(sheets);

    /** @type {ExportItem[]} */
    const filterValueItems = [];
    /** @type {ExportItem[]} */
    const dataWorksheetItems = [];
    const visualStatuses = [];

    for (const { sheet, data, visualSpec, visualSpecError, error } of allSheetsData) {
      if (error) {
        console.warn(`Skipping "${sheet.name}": ${error.message}`);
        continue;
      }

      // ---- DZV check ----
      if (dzvMap[sheet.name] === false) {
        console.log(`[DZV] Skipping hidden sheet: "${sheet.name}"`);
        continue;
      }

      let summaryData = data;  // already fetched
      if (!summaryData.columns || summaryData.columns.length === 0) {
        console.warn(`[Export] Skipping "${sheet.name}": summary data has no columns`);
        continue;
      }
      if (!summaryData.data || summaryData.data.length === 0) {
        console.warn(`[Export] Skipping "${sheet.name}": summary data has no rows`);
        continue;
      }

      const layout = layoutMap.get(sheet.name);
      const visualName = (layout && layout.displayName) ? layout.displayName : sheet.name;

      if (isFilterValueWorksheet(sheet.name, summaryData)) {
        let matchedFilterName = null;
        let matchedValues = null;

        for (const [filterField, values] of Object.entries(filterValuesMap)) {
          const cleanFilterField = filterField.toLowerCase().replace(/[^a-z]/g, '');
          const cleanSheetName = sheet.name.toLowerCase().replace(/[^a-z]/g, '');
          if (cleanSheetName.includes(cleanFilterField) || cleanFilterField.includes(cleanSheetName)) {
            matchedFilterName = filterField;
            matchedValues = values;
            break;
          }
        }

        if (!matchedValues) {
          matchedFilterName = visualName;
          matchedValues = [`Values: ${summaryData.data.length} items`];
        }

        filterValueItems.push({
          type: "filterValue",
          name: sheet.name,
          visualName: visualName,
          filterName: matchedFilterName,
          values: matchedValues,
          layout: layout,
          rowCount: 2 + (Array.isArray(matchedValues) ? matchedValues.length : 1),
          originalData: summaryData
        });
      } else {
        // ── NEW: rebuild the visual table (pivot, merge, sort, visible columns, title) ──
        const visualModel = buildVisualModel(fmtModel, sheet.name, summaryData,
          { dashboardName: dashboard.name, displayName: visualName, visualSpec, visualSpecError, layout });
        console.log(`[VisualSpec] ${sheet.name}:`, {
          available: visualModel.metadata.visualSpecAvailable,
          keys: visualModel.metadata.visualSpecKeys,
          type: visualModel.type,
          marks: visualModel.metadata.markTokens,
          markSource: visualModel.metadata.markSource,
          error: visualModel.metadata.visualSpecError,
          spec: visualModel.source.visualSpec
        });
        let renderDecision = chooseVisualRenderer(visualModel);
        if (renderDecision.renderer === "excel-chart") {
          try {
            visualModel.chartSpecs = buildExcelChartSpecs(visualModel, fmtModel);
          } catch (err) {
            console.warn(`[Visual] ${sheet.name}: native chart not possible – ${err.message}`);
            renderDecision = imageOrFallback(visualModel, `native chart not possible (${err.message})`);
          }
        }
        visualModel.renderer = renderDecision.renderer;
        visualModel.status = renderDecision.status;
        visualModel.statusReason = renderDecision.reason;
        visualStatuses.push({
          worksheet: sheet.name,
          type: visualModel.type,
          renderer: renderDecision.renderer,
          status: renderDecision.status,
          reason: renderDecision.reason,
          visualSpecAvailable: visualModel.metadata.visualSpecAvailable,
          imageApiAvailable: visualModel.metadata.imageApiAvailable
        });
        console.log(`[Visual] ${sheet.name}: ${visualModel.type} -> ${renderDecision.renderer} (${renderDecision.reason})`);

        if (renderDecision.renderer === "data-fallback" && visualModel.viewModel.kind === "chart" &&
            FORMAT_CONFIG.chartPolicy === "skip") {
          console.warn(`[Export] "${sheet.name}" is a chart - skipped by FORMAT_CONFIG.chartPolicy`);
          continue;
        }
        let vm = visualModel.viewModel;
        if (!vm.rows.length || !vm.order.length) {
          console.warn(`[Export] "${sheet.name}" produced an empty formatted view; rebuilding from live summary data`);
          const rawVisualModel = buildVisualModel(null, sheet.name, summaryData,
            { dashboardName: dashboard.name, displayName: visualName, visualSpec, visualSpecError, layout });
          if (rawVisualModel.viewModel.rows.length && rawVisualModel.viewModel.order.length) {
            vm = rawVisualModel.viewModel;
            visualModel.viewModel = vm;
            visualModel.data.rows = vm.rows;
            visualModel.data.columns = vm.cols;
            visualModel.data.dimensions = vm.order.filter(i => vm.cols[i].isHeader).map(i => vm.cols[i]);
            visualModel.data.measures = vm.order.filter(i => !vm.cols[i].isHeader).map(i => vm.cols[i]);
            visualModel.statusReason = "TWB view was empty; exported live summary data";
          }
        }
        if (!vm.rows.length || !vm.order.length) {
          console.warn(`[Export] Skipping "${sheet.name}": no renderable rows or columns after fallback`);
          continue;
        }
        // a chart exported as data: axis/colour/detail fields are the data – show every column
        if (visualModel.renderer === "data-fallback" && !VISUAL_RENDERERS.cellTypes.has(visualModel.type)) {
          vm.order = vm.cols.map((_, i) => i);
          vm.cols.forEach(c => { if (!c.label) c.label = c.name; });
          vm.showHeaderRow = true;
        }
        const isGraphic = visualModel.renderer === "excel-chart" || visualModel.renderer === "tableau-image";
        const box = isGraphic ? graphicBox(layout, vm) : null;
        const isKPI = !isGraphic && isKPIViewModel(vm);
        dataWorksheetItems.push({
          name: sheet.name,
          visualName: vm.title.text,
          layout,
          isKPI,
          type: "worksheet",
          vm,
          visualModel,
          box,                                                 // chart/image: pixel size from the dashboard zone
          fixedGridW: box ? box.gridW : null,
          columns: vm.order,                                   // only its length is used for layout
          rowCount: box ? box.rows : isKPI ? (vm.showTitle ? 1 : 0) + vm.order.length : viewModelHeight(vm)
        });
      }
    }

    // ── parameter controls (e.g. Start Date / End Date): not filters, so read them separately ──
    try {
      const params = typeof dashboard.getParametersAsync === "function" ? await dashboard.getParametersAsync() : [];
      const controls = (dashboard.objects || []).filter(o => o.type === "parameter" && o.isVisible !== false);
      const used = new Set();
      const norm = s => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
      params.forEach(param => {
        let control = controls.find(o => !used.has(o) && norm(o.name) === norm(param.name)) ||
                      controls.find(o => !used.has(o) && norm(o.name).includes(norm(param.name)));
        if (!control) return;                                      // parameter not shown on this dashboard
        used.add(control);
        const raw = param.currentValue ? (param.currentValue.formattedValue ?? String(param.currentValue.value)) : "";
        const value = String(raw).replace(/\s+12:00:00\s*AM$|\s+00:00:00$/i, "");   // dates: no midnight time
        filterValueItems.push({
          type: "filterValue", isParameter: true, name: param.name, visualName: param.name,
          filterName: param.name, values: [value], layout: layoutMap.get(control.name), rowCount: 3
        });
      });
      if (params.length) console.log(`[Parameters] ${used.size} of ${params.length} parameters have a control on the dashboard`);
    } catch (err) {
      console.warn("[Parameters] could not read parameters:", err.message);
    }

    const allItems = [...filterValueItems, ...dataWorksheetItems];

    if (allItems.length === 0) {
      throw new Error("No data found in any visible worksheet.");
    }

    console.table(visualStatuses);
    updateVisualStatus(visualStatuses, workbookWarning);

    const placedItems = allItems.map((item, idx) => {
      if (item.layout) {
        const l = item.layout;
        let calculatedWidth = l.gridW;

        if (item.fixedGridW) {
          calculatedWidth = item.fixedGridW;
        } else if (item.type === "filterValue") {
          calculatedWidth = Math.max(l.gridW, 3);
        } else if (item.isKPI) {
          calculatedWidth = Math.max(l.gridW, Math.min(6, item.columns.length + 1));
        } else {
          calculatedWidth = Math.max(l.gridW, Math.min(25, item.columns.length + 1));
        }

        return {
          ...item,
          gridRow: l.gridRow,
          gridCol: l.gridCol,
          gridW: calculatedWidth,
          allocatedRows: item.rowCount,
        };
      } else {
        const itemsPerRow = 3;
        const rowIdx = Math.floor(idx / itemsPerRow);
        const colIdx = idx % itemsPerRow;
        const gridRow = rowIdx * (Math.max(8, item.rowCount) + ROW_GAP);
        const gridCol = colIdx * 12;
        const gridW = item.fixedGridW || ((item.type === "filterValue") ? 4 : (item.isKPI ? 5 : 12));

        return {
          ...item,
          gridRow: gridRow,
          gridCol: gridCol,
          gridW: gridW,
          allocatedRows: item.rowCount,
        };
      }
    });

    resolveCollisions(placedItems);

// SNAP-OUT-OF-GROUPED-ROWS PASS
// Excel hides entire physical rows, not per-column cells. resolveCollisions
// only pushes items that horizontally overlap another item's column range —
// so a table in column K can still be placed on a row that a totally
// different table (in column A) has marked hidden=true for its own
// row-grouping. When the workbook opens with that group collapsed, the
// column-K table vanishes too, even though nothing "collided" by column.
// This pass detects that and pushes the item below the other table's full
// physical row range (not just its visible rows).
let snapChanged = true;
const MAX_SNAP_PASSES = 10;
let snapPass = 0;

while (FORMAT_CONFIG.groupOverflowRows && snapChanged && snapPass < MAX_SNAP_PASSES) {
  snapChanged = false;
  snapPass++;

  for (const item of placedItems) {
    for (const other of placedItems) {
      if (item === other) continue;

      const otherTotal = other.allocatedRows || other.rowCount || 0;
      if (otherTotal <= ROW_GROUP_THRESHOLD + 2) continue; // other has no hidden rows

      const otherVisibleEnd = other.gridRow + ROW_GROUP_THRESHOLD + 3;
      const otherPhysicalEnd = other.gridRow + otherTotal + ROW_GAP;

      if (item.gridRow >= otherVisibleEnd && item.gridRow < otherPhysicalEnd) {
        item.gridRow = otherPhysicalEnd;
        snapChanged = true;
      }
    }
  }
}

    const workbook = new ExcelJS.Workbook();
    const sheetName = (dashboard.name || "Dashboard Export")
      .replace(/[\\\/\*\?\[\]:]/g, "")
      .slice(0, 31);
    const worksheet = workbook.addWorksheet(sheetName);

    const colWidths = {};
    const exactWidths = {};
    const allTablesInfo = [];
    const tracker = makeRangeTracker();

    let currentRow = 0;
    const dashboardName = dashboard.name || "Dashboard Export";
    const exportDate = new Date().toLocaleString();
    let dashTitleProps = null;
    const dashFmt = fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name];
    if (dashFmt) {
      const run = dashFmt.title && (dashFmt.title.runs.find(x => x.text.trim()) || {}).props;
      dashTitleProps = tfMerge(TABLEAU_DEFAULTS.dashTitle,
        tfCollect(fmtModel.workbookStyle, TF_ELEMENTS.dashTitle), tfCollect(dashFmt.style, TF_ELEMENTS.dashTitle), run);
    }
    const usedZones = new Set(dataWorksheetItems.flatMap(it => it.vm ? it.vm.headerZoneIds : []));
    const titleRuns = tfDashboardTitleRuns(fmtModel, dashboard.name, usedZones);
    const titleHeight = writeDashboardTitle(worksheet, dashboardName, exportDate, currentRow, 0, tracker, dashTitleProps, titleRuns);
    currentRow += titleHeight;


    const adjustedItems = placedItems.map(item => ({
      ...item,
      gridRow: item.gridRow + currentRow
    }));

    // native charts: data goes to a hidden sheet now, chart parts are injected after writeBuffer()
    const chartJobs = [];
    let chartDataSheet = null;
    let chartDataRow = 0;
    const reserveGraphicBlock = item => {
      for (let c = item.gridCol; c < item.gridCol + item.gridW; c++) colWidths[c] = Math.max(colWidths[c] || 0, 10);
      tracker.update(item.gridRow, item.gridCol);
      tracker.update(item.gridRow + item.allocatedRows - 1, item.gridCol + item.gridW - 1);
    };
    const writeGraphicTitle = item => {
      if (!item.vm.showTitle) return 0;
      writeTableauTitle(worksheet, item.gridRow, item.gridCol, item.vm.title.text, item.vm.title.props, item.gridW);
      return 1;
    };

    for (let i = 0; i < adjustedItems.length; i++) {
      const item = adjustedItems[i];

      if (item.visualModel && item.visualModel.renderer === "excel-chart") {
        const status = visualStatuses.find(entry => entry.worksheet === item.name);
        const titleRows = writeGraphicTitle(item);
        const specs = item.visualModel.chartSpecs;
        const paneH = Math.floor(item.box.heightPx / specs.length);     // separate panes stack vertically
        if (!chartDataSheet) {
          const name = CHART_DATA_SHEET === sheetName ? CHART_DATA_SHEET + " (export)" : CHART_DATA_SHEET;
          chartDataSheet = workbook.addWorksheet(name, { state: "hidden" });
        }
        specs.forEach((spec, k) => {
          const refs = ExcelChartWriter.writeChartData(chartDataSheet, spec, chartDataRow);
          chartDataRow = refs.nextRow;
          const offPx = k * paneH;
          chartJobs.push({
            spec, refs, item, name: item.visualName || item.name,
            col: item.gridCol, row: item.gridRow + titleRows + Math.floor(offPx / EXCEL_ROW_PX),
            rowOffPx: offPx % EXCEL_ROW_PX, widthPx: item.box.widthPx, heightPx: paneH
          });
        });
        reserveGraphicBlock(item);
        item.visualModel.status = "success";
        item.visualModel.statusReason = `native Excel ${[...new Set(specs.map(s => s.kind))].join("/")} chart` +
          (specs.length > 1 ? ` (${specs.length} panes)` : "");
        if (status) { status.status = "success"; status.reason = item.visualModel.statusReason; }
        continue;
      }

      if (item.visualModel && item.visualModel.renderer === "tableau-image") {
        const imageWidth = item.box ? item.box.widthPx : Math.max(320, item.gridW * PX_PER_COL);
        const imageHeight = item.box ? item.box.heightPx : Math.max(160, (item.allocatedRows || item.rowCount || 8) * PX_PER_ROW);
        const status = visualStatuses.find(entry => entry.worksheet === item.name);
        try {
          const imageBase64 = await renderTableauImage(item.visualModel, imageWidth, imageHeight);
          const titleRows = writeGraphicTitle(item);
          const imageId = workbook.addImage({ base64: imageBase64, extension: "png" });
          worksheet.addImage(imageId, {
            tl: { col: item.gridCol, row: item.gridRow + titleRows },
            ext: { width: imageWidth, height: imageHeight }
          });
          reserveGraphicBlock(item);
          item.visualModel.status = "success";
          item.visualModel.statusReason = "Tableau SVG rendered and embedded as PNG" +
            (item.visualModel.type === VISUAL_TYPES.MAP ? " (marks on latitude/longitude, no basemap)" : "");
          if (status) {
            status.status = "success";
            status.reason = item.visualModel.statusReason;
          }
          continue;
        } catch (err) {
          item.visualModel.renderer = "data-fallback";
          item.visualModel.status = "warning";
          item.visualModel.statusReason = `image renderer failed: ${err.message}`;
          if (status) {
            status.renderer = "data-fallback";
            status.status = "warning";
            status.reason = item.visualModel.statusReason;
          }
          console.warn(`[Visual] ${item.name}: image renderer failed; using data fallback`, err);
        }
      }

      if (item.type === "filterValue") {
        writeIndividualFilterTable(
          worksheet,
          item.filterName,
          item.values,
          item.gridRow,
          item.gridCol,
          tracker,
          item.isParameter ? "PARAMETER VALUE" : undefined
        );
        colWidths[item.gridCol] = Math.max(colWidths[item.gridCol] || 0, 35);
      } else if (item.isKPI) {
        writeKPICardStacked(worksheet, item.vm, item.gridRow, item.gridCol, tracker);
        colWidths[item.gridCol] = Math.max(colWidths[item.gridCol] || 0, 22);
        colWidths[item.gridCol + 1] = Math.max(colWidths[item.gridCol + 1] || 0, 18);
      } else {
        writeRegularTable(worksheet, item.vm, item.gridRow, item.gridCol, tracker, allTablesInfo, colWidths, exactWidths);
      }
    }

    updateVisualStatus(visualStatuses, workbookWarning);

    setColumnWidths(worksheet, colWidths, exactWidths);
    applyAutoFilters(worksheet, allTablesInfo);

    setExportStatus("Building the Excel file…");
    /** @type {any} the XLSX bytes: ExcelJS's buffer, then injectCharts' Uint8Array */
    let buffer = await workbook.xlsx.writeBuffer();
    if (chartJobs.length) {
      // fit every chart to its own block now that the final column widths/row heights are known:
      // tables widen the columns, so a chart kept at its Tableau pixel size would end short of the
      // tables / charts aligned with it on the dashboard → stretch it to the block's edges
      const colPx = ci => Math.floor((worksheet.getColumn(ci + 1).width || 8.43) * 7 + 5);
      const rowPx = ri => { const h = worksheet.getRow(ri + 1).height; return h ? Math.round(h * 4 / 3) : EXCEL_ROW_PX; };
      chartJobs.forEach(job => {
        let w = 0, h = -job.rowOffPx;
        for (let c = job.item.gridCol; c < job.item.gridCol + job.item.gridW; c++) w += colPx(c);
        for (let r = job.row; r < job.item.gridRow + job.item.allocatedRows; r++) h += rowPx(r);
        const panes = job.item.visualModel.chartSpecs.length;     // stacked panes share the block height
        job.widthPx = Math.max(160, w);
        job.heightPx = Math.max(60, panes > 1 ? Math.min(job.heightPx, h) : h);
      });
      try {
        buffer = await ExcelChartWriter.injectCharts(buffer, { sheetIndex: 0, charts: chartJobs });
        console.log(`[Charts] ${chartJobs.length} native Excel chart(s) added`);
      } catch (err) {
        console.error("[Charts] could not add native charts", err);
        chartJobs.forEach(job => {
          const status = visualStatuses.find(entry => entry.worksheet === job.item.name);
          if (status) { status.status = "warning"; status.reason = `chart injection failed: ${err.message}`; }
        });
        updateVisualStatus(visualStatuses, workbookWarning);
      }
    }
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    if (saveTarget) {
      const writable = await saveTarget.createWritable();
      await writable.write(blob);
      await writable.close();
      console.log(`[Export] saved to ${saveTarget.name}`);
      appendExportStatus(`saved as ${saveTarget.name}`);
    } else {
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = exportFileName;
      link.click();
      URL.revokeObjectURL(link.href);
    }

    console.log("✅ Export completed with Tableau formatting (fonts, colours, number formats, borders)");

  } catch (err) {
    console.error("[Export]", err);
    setExportStatus("Export failed: " + err.message, true);
    alert("Export failed. Check console (F12) for details.\n\n" + err.message);
  } finally {
    btn.disabled = false;
  }
}
