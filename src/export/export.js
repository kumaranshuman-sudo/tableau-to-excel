/* The export: Tableau dashboard → formatted XLSX with native charts. */
import ExcelJS from "exceljs";
import { buildExcelChartSpecs } from "../charts/model/index.js";
import { FORMAT_CONFIG, TABLEAU_DEFAULTS, TF_ELEMENTS, VISUAL_TYPES } from "../config.js";
import { writeDashboardTitle, writeIndividualFilterTable } from "./cell-writers.js";
import { CHART_DATA_SHEET, EXCEL_ROW_PX, OBJECT_KIND, PX_PER_COL, PX_PER_ROW, ROW_GAP, buildLayoutMap, getExcelColumnName, graphicBox, makeRangeTracker, resolveCollisions, setTableVisibleRows, tableDataStart, viewModelHeight } from "./layout.js";
import { extractFilterValuesPerField, fetchAllSheetsData, isFilterValueWorksheet } from "./sheet-data.js";
import { applyAutoFilters, setColumnWidths, writeKPICardStacked, writeRegularTable, writeTableauTitle } from "./visual-writers.js";
import { KPI_GUTTER, buildKpiCard, buildTextCard, writeKpiCard } from "./kpi-card.js";
import { tfStrokeToBorder, tfZoneText } from "../twb/dashboard-text.js";
import { tfExcelFill } from "../format/excel-style.js";
import { tfDashboardTitleRuns } from "../twb/dashboard-text.js";
import { tfCollect, tfDashboardShading, tfMerge } from "../twb/formatter.js";
import { backgroundPlan } from "./backgrounds.js";
import { appendExportStatus, setExportStatus, updateVisualStatus } from "../ui/status.js";
import { chooseSaveTarget, ensureFormatModel, formatModelFileName, getTitleMap, getWorkbookImages } from "../ui/workbook-store.js";
import { classifyImageObjects, fitImage, imageInfo, nativeAnchor, prepareImage, webLink } from "./images.js";
import { tvIconSheet } from "../visual/icon-sheet.js";
import { isKPIViewModel } from "../visual/classify.js";
import { renderTableauImage } from "../visual/image-renderer.js";
import { VISUAL_RENDERERS, buildVisualModel, chooseVisualRenderer, imageOrFallback } from "../visual/visual-model.js";
import { checkWorkbookMatch, describeWorkbookMatch } from "../visual/workbook-match.js";
import { ExcelChartWriter } from "../charts/writer/index.js";
import { fixSheetProperties } from "../charts/writer/package.js";
import { exportFileName as makeExportFileName } from "../util.js";

export async function exportToExcel() {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  const sheets = dashboard.worksheets;
  const btn = /** @type {HTMLButtonElement} */ (document.getElementById("export_button"));
  btn.disabled = true;

  try {
    if (!(await ensureFormatModel()) && !window.confirm(
        "No Tableau workbook is loaded for this dashboard.\n\n" +
        "Without it the extension cannot tell which sheets are charts, so they will be exported as tables.\n\n" +
        "Cancel, then click 📁 Load Workbook to export real charts – or OK to export tables anyway.")) {
      setExportStatus("Export stopped – click 📁 Load Workbook first so charts are exported as charts", true);
      return;
    }
    // "<workbook> - <dashboard>.xlsx": the workbook the formatting comes from (remembered or loaded)
    const exportFileName = makeExportFileName(formatModelFileName(), dashboard.name);
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

    // ── 1. Build DZV map ── (a sheet squeezed into a 1 px zone is a common way to hide it: not shown either)
    const dzvMap = {};
    (dashboard.objects || [])
      .filter(obj => obj.type === "worksheet")
      .forEach(obj => {
        const tiny = obj.size && (obj.size.width < 8 || obj.size.height < 8);
        dzvMap[obj.name] = obj.isVisible !== false && !tiny;
      });
    console.log("[DZV] Visibility map:", dzvMap);

    // ── 2. Fetch all sheets in parallel ──
    setExportStatus(`Reading ${sheets.length} worksheets from Tableau…`);
    const allSheetsData = await fetchAllSheetsData(sheets);

    /** @type {{ name: string, shape: string, layout: any }[]} button / icon sheets, drawn as their icon */
    const iconSheets = [];
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

      // a button / icon sheet (a custom shape on empty shelves): drawn as its icon, not as a table of its flags
      const icon = fmtModel ? tvIconSheet(fmtModel, sheet.name, summaryData) : null;
      if (icon) {
        if (icon.shape && layout) iconSheets.push({ name: sheet.name, shape: icon.shape, layout });
        else console.log(`[Icons] ${sheet.name}: a Tableau shape – not exported`);
        continue;
      }

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
            // Tableau draws legends as separate dashboard cards, never inside the view: the chart keeps a
            // legend only when the dashboard shows a colour / size / shape legend for this sheet
            const dashZones = fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name] ? fmtModel.dashboards[dashboard.name].zones : null;
            if (dashZones && !dashZones.some(z => /^(color|size|shape)$/.test(z.type) && z.name === sheet.name && !z.hidden)) {
              visualModel.chartSpecs.forEach(s => { s.legend = false; });
            }
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
        // KPI tile: rebuilt from its Tableau label and filling its zone (null = drawn as a label | value table)
        const kpiCard = isKPI ? buildKpiCard(vm, sheet.name, layout) : null;
        const kpiW = kpiCard && layout && layout.widthPx ? Math.max(kpiCard.tiles.length, layout.gridRight - layout.gridCol) : null;
        dataWorksheetItems.push({
          name: sheet.name,
          visualName: vm.title.text,
          layout,
          isKPI,
          kpiCard,
          type: "worksheet",
          vm,
          visualModel,
          box,                                                 // chart/image: pixel size from the dashboard zone
          fixedGridW: box ? box.gridW : kpiW,
          columns: vm.order,                                   // only its length is used for layout
          rowCount: box ? box.rows : kpiCard ? kpiCard.rows : isKPI ? (vm.showTitle ? 1 : 0) + vm.order.length : viewModelHeight(vm)
        });
      }
    }

    // ── parameter controls (e.g. Start Date / End Date): not filters, so read them separately ──
    try {
      const params = typeof dashboard.getParametersAsync === "function" ? await dashboard.getParametersAsync() : [];
      const controls = (dashboard.objects || []).filter(o => OBJECT_KIND[o.type] === "parameter" && o.isVisible !== false);
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

    // Tableau's number + trend tile: a chart in the zone right under / over a KPI card, same position and
    // width, takes the card's columns so the two line up as one card
    const near = (a, b) => Math.abs(a - b) <= 8;
    dataWorksheetItems.filter(it => it.kpiCard && it.fixedGridW && it.layout).forEach(card => {
      const c = card.layout;
      dataWorksheetItems.forEach(other => {
        const o = other.layout;
        if (other === card || !other.box || !o || !o.widthPx || !near(o.xPx, c.xPx) || !near(o.widthPx, c.widthPx)) return;
        if (near(o.yPx, c.yPx + c.heightPx) || near(c.yPx, o.yPx + o.heightPx)) { other.fixedGridW = card.fixedGridW; other.pairedCard = true; }
      });
    });

    // ── dashboard text boxes (banners, titles, notes): drawn where Tableau puts them, as Tableau formats them
    const dashZones = (fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name] || { zones: [] }).zones;
    const headerZoneIds = new Set(dataWorksheetItems.flatMap(it => it.vm ? it.vm.headerZoneIds : []));   // already table headers
    /** @type {ExportItem[]} */
    const textItems = [];
    (dashboard.objects || []).filter(o => OBJECT_KIND[o.type] === "text" && o.isVisible !== false).forEach(o => {
      const zone = dashZones.find(z => z.type === "text" && String(z.id) === String(o.id));
      if (!zone || zone.hidden || headerZoneIds.has(zone.id)) return;
      const layout = layoutMap.get(`text:${o.id}`);
      const card = buildTextCard(zone, layout, fmtModel);
      if (!card) return;
      textItems.push({ type: "text", name: `text:${o.id}`, visualName: tfZoneText(zone), layout, textCard: card,
        rowCount: card.rows, columns: [], fixedGridW: layout && layout.widthPx ? Math.max(1, layout.gridRight - layout.gridCol) : null });
    });
    if (textItems.length) console.log(`[Text] ${textItems.length} dashboard text box(es)`);

    // ── dashboard images (logos, icons) from the files packaged in the .twbx: a logo in its own zone is
    //    a block of the layout, an icon over a sheet floats on that sheet's block, backgrounds are left out
    const imageKinds = classifyImageObjects(dashboard.objects || []);
    const imageObjects = (dashboard.objects || []).filter(o => o.type === "image" && o.isVisible !== false && imageKinds.has(String(o.id)));
    const workbookImages = imageObjects.length || iconSheets.length ? await getWorkbookImages() : {};
    /** @type {ExportItem[]} */
    const imageItems = [];
    /** @type {(ImageBlock & { host: string, xPx: number, yPx: number })[]} */
    const overlayImages = [];
    imageObjects.forEach(o => {
      const cls = imageKinds.get(String(o.id));
      const zone = dashZones.find(z => z.type === "bitmap" && String(z.id) === String(o.id));
      if (!zone || !zone.param || zone.hidden) return;
      if (cls.kind === "backdrop" || cls.kind === "tiny") { console.log(`[Images] ${zone.param}: ${cls.kind === "tiny" ? "divider" : "background"} – not exported`); return; }
      const file = workbookImages[zone.param];
      if (!file) { console.warn(`[Images] ${zone.param} is not in the loaded workbook – load the .twbx to export it`); return; }
      const info = imageInfo(file.data, zone.param);
      if (!info.type) { console.warn(`[Images] ${zone.param}: not a PNG, JPEG, GIF or SVG file`); return; }
      /** @type {ImageBlock} */
      const image = { file, info, zone, url: webLink(zone.url), widthPx: o.size.width, heightPx: o.size.height };
      if (cls.kind === "overlay") { overlayImages.push({ ...image, host: cls.host, xPx: o.position.x, yPx: o.position.y }); return; }
      const layout = layoutMap.get(`image:${o.id}`);
      if (!layout) return;
      imageItems.push({ type: "image", name: `image:${o.id}`, visualName: zone.param, layout, image, columns: [],
        rowCount: Math.max(1, Math.round(o.size.height / EXCEL_ROW_PX)),
        fixedGridW: layout.widthPx ? Math.max(1, layout.gridRight - layout.gridCol) : null });
    });
    // icon sheets: the shape centred in the sheet's zone at about Tableau's mark size
    iconSheets.forEach(({ name, shape, layout }) => {
      const file = workbookImages["shape:" + shape];
      if (!file) { console.warn(`[Icons] ${name}: shape "${shape}" not in the loaded workbook`); return; }
      const info = imageInfo(file.data, shape);
      const w = layout.widthPx || 40, h = layout.heightPx || 40;
      const box = Math.min(28, Math.max(12, Math.round(Math.min(w, h) * 0.5)));
      if (!info.type) return;
      imageItems.push({ type: "image", name, visualName: name, layout, columns: [],
        image: { file, info, zone: { param: shape, scaled: true, centered: true }, url: null, widthPx: box, heightPx: box },
        rowCount: Math.max(1, Math.round(h / EXCEL_ROW_PX)),
        fixedGridW: layout.widthPx ? Math.max(1, layout.gridRight - layout.gridCol) : null });
    });
    if (imageItems.length || overlayImages.length) console.log(`[Images] ${imageItems.length + overlayImages.length} image(s)`);

    const allItems = [...filterValueItems, ...dataWorksheetItems, ...textItems, ...imageItems];

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
// Excel hides entire physical rows, not per-column cells, so a long table's collapsed rows would
// also hide whatever sits beside it. Charts, images, cards and filter lists that share rows with
// the table keep those rows visible (setTableVisibleRows); a table that would start inside another
// table's collapsed rows is pushed below that table's full physical row range.
let snapChanged = true;
const MAX_SNAP_PASSES = 10;
let snapPass = 0;

while (FORMAT_CONFIG.groupOverflowRows && snapChanged && snapPass < MAX_SNAP_PASSES) {
  snapChanged = false;
  snapPass++;
  setTableVisibleRows(placedItems);               // charts / cards beside a long table keep its rows visible

  for (const item of placedItems) {
    for (const other of placedItems) {
      if (item === other) continue;
      if (!other.visibleRows || other.visibleRows >= other.vm.rows.length) continue;  // other hides no rows

      const otherTotal = other.allocatedRows || other.rowCount || 0;
      const otherVisibleEnd = tableDataStart(other) + other.visibleRows;               // first hidden row
      const otherPhysicalEnd = other.gridRow + otherTotal + ROW_GAP;

      if (item.gridRow >= otherVisibleEnd && item.gridRow < otherPhysicalEnd) {
        item.gridRow = otherPhysicalEnd;
        snapChanged = true;
      }
    }
  }
}
setTableVisibleRows(placedItems);

    const workbook = new ExcelJS.Workbook();
    const sheetName = (dashboard.name || "Dashboard Export")
      .replace(/[\\\/\*\?\[\]:]/g, "")
      .slice(0, 31);
    const worksheet = workbook.addWorksheet(sheetName);
    // the sheet as the dashboard's canvas: no cell grid, as in Tableau
    worksheet.views = [{ showGridLines: !!FORMAT_CONFIG.sheetGridlines }];

    /** @type {Record<number, number>} */
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
    // Tableau shows the dashboard title only when its title zone is on (text boxes are drawn in place)
    const titleShown = !dashFmt || dashFmt.zones.some(z => z.type === "title" && !z.hidden);
    const titleHeight = writeDashboardTitle(worksheet, titleShown ? dashboardName : "", exportDate, currentRow, 0, tracker, dashTitleProps,
      titleShown ? titleRuns : null);
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
      } else if (item.image) {
        tracker.update(item.gridRow, item.gridCol);                 // the picture goes in once the columns are sized
        tracker.update(item.gridRow + item.allocatedRows - 1, item.gridCol + item.gridW - 1);
      } else if (item.textCard) {
        writeKpiCard(worksheet, item.textCard, null, item.gridRow, item.gridCol, item.gridW, tracker, colWidths);
      } else if (item.kpiCard) {
        writeKpiCard(worksheet, item.kpiCard, item.vm, item.gridRow, item.gridCol, item.gridW, tracker, colWidths);
      } else if (item.isKPI) {
        writeKPICardStacked(worksheet, item.vm, item.gridRow, item.gridCol, tracker);
        colWidths[item.gridCol] = Math.max(colWidths[item.gridCol] || 0, 22);
        colWidths[item.gridCol + 1] = Math.max(colWidths[item.gridCol + 1] || 0, 18);
      } else {
        writeRegularTable(worksheet, item.vm, item.gridRow, item.gridCol, tracker, allTablesInfo, colWidths, exactWidths, item.visibleRows);
      }
    }

    updateVisualStatus(visualStatuses, workbookWarning);

    // a KPI card directly above / below a chart of the same width and colour (Tableau's number + trend
    // tile): the layout gap between them gets the card colour too, so they read as one card
    const sameColor = (a, b) => !!a && !!b && String(a).replace(/^#/, "").slice(-6).toUpperCase() === String(b).replace(/^#/, "").slice(-6).toUpperCase();
    adjustedItems.filter(it => it.kpiCard && it.kpiCard.background).forEach(card => {
      const color = card.kpiCard.background;
      const cardEnd = card.gridRow + card.allocatedRows;
      adjustedItems.forEach(other => {
        if (other === card || other.gridCol !== card.gridCol || Math.abs(other.gridW - card.gridW) > 1) return;
        const specs = other.visualModel && other.visualModel.renderer === "excel-chart" ? other.visualModel.chartSpecs : null;
        if (!sameColor(specs && specs.length ? specs[0].background : null, color)) return;
        const otherEnd = other.gridRow + other.allocatedRows;
        const [from, to] = other.gridRow >= cardEnd ? [cardEnd, other.gridRow] : otherEnd <= card.gridRow ? [otherEnd, card.gridRow] : [0, 0];
        if (to - from < 1 || to - from > 2) return;
        for (let rr = from; rr < to; rr++) {
          for (let c = card.gridCol; c < card.gridCol + card.gridW; c++) {
            const cell = worksheet.getCell(rr + 1, c + 1);
            cell.fill = tfExcelFill(color);
            if (c === card.gridCol) cell.border = { left: KPI_GUTTER };
            if (c === card.gridCol + card.gridW - 1) cell.border = { ...(cell.border || {}), right: KPI_GUTTER };
          }
        }
      });
    });

    // bordered layout containers (Layout pane → Border, e.g. a card around a column of sheets): their
    // outline around the blocks they hold
    const zoneById = new Map(dashZones.map(z => [String(z.id), z]));
    const within = (zoneId, containerId) => {
      for (let z = zoneById.get(String(zoneId)); z && z.parent; z = zoneById.get(String(z.parent))) if (String(z.parent) === containerId) return true;
      return false;
    };
    dashZones.filter(z => /^layout-/.test(z.type) && !z.hidden && z.style && z.style.borderStyle && z.style.borderStyle !== "none" &&
                          (z.style.borderWidth || 0) > 0).forEach(container => {
      const members = adjustedItems.filter(it => it.layout && it.layout.id !== undefined && within(it.layout.id, String(container.id)));
      if (!members.length) return;
      const top = Math.min(...members.map(m => m.gridRow)), left = Math.min(...members.map(m => m.gridCol));
      const bottom = Math.max(...members.map(m => m.gridRow + m.allocatedRows)) - 1;
      const right = Math.max(...members.map(m => m.gridCol + m.gridW)) - 1;
      /** @type {Partial<import("exceljs").Border>} */
      const side = { style: /** @type {import("exceljs").BorderStyle} */ (tfStrokeToBorder(container.style.borderWidth) || "thin"),
                     color: { argb: container.style.borderColor || "FFD4D4D4" } };
      const edge = (r, c, sides) => { const cell = worksheet.getCell(r + 1, c + 1); cell.border = { ...(cell.border || {}), ...sides }; };
      for (let c = left; c <= right; c++) { edge(top, c, { top: side }); edge(bottom, c, { bottom: side }); }
      for (let r = top; r <= bottom; r++) { edge(r, left, { left: side }); edge(r, right, { right: side }); }
    });

    // dashboard shading, container / backdrop backgrounds and each object's own: the cells around and under
    // the blocks take the colour Tableau shows there; cells a block coloured itself keep their colour
    const dashColor = tfDashboardShading(fmtModel, dashboard.name);
    const bgBlocks = adjustedItems.filter(it => it.layout && it.layout.id !== undefined).map(it => {
      const zone = zoneById.get(String(it.layout.id));
      const zoneBg = (zone && zone.style && zone.style.bgColor) || undefined;
      // a worksheet draws its shading over its zone: its colour, Tableau's white, or see-through
      const shading = it.type === "worksheet" && it.vm ? it.vm.fmt.sheetShading() : "none";
      const own = shading === "none" ? zoneBg : shading || "white";
      return { zoneId: String(it.layout.id), top: it.gridRow, left: it.gridCol,
               bottom: it.gridRow + it.allocatedRows - 1, right: it.gridCol + it.gridW - 1, own };
    });
    // the dashboard's extent in px (its objects), for the size of zones given in 0–100000 units
    const dashExt = (dashboard.objects || []).reduce((a, o) => o.position && o.size
      ? { w: Math.max(a.w, o.position.x + o.size.width), h: Math.max(a.h, o.position.y + o.size.height) } : a, { w: 0, h: 0 });
    if (bgBlocks.length && (dashColor || dashZones.some(z => z.style && z.style.bgColor))) {
      const used = tracker.toRange();
      const plan = backgroundPlan(dashZones, dashColor, bgBlocks,
        { top: Math.min(...bgBlocks.map(b => b.top)), left: 0, bottom: used.e.r, right: used.e.c },
        { w: (dashFmt && dashFmt.width) || dashExt.w || 1200, h: (dashFmt && dashFmt.height) || dashExt.h || 800 });
      for (let r = 0; r <= used.e.r; r++) {
        for (let c = 0; c <= used.e.c; c++) {
          const color = plan.colorAt(r, c);
          if (!color) continue;
          const cell = worksheet.getCell(r + 1, c + 1);
          const target = cell.isMerged ? cell.master : cell;
          if (target.fill && target.fill.type === "pattern" && target.fill.pattern !== "none") continue;
          target.style = { ...target.style, fill: tfExcelFill(color) };      // a style of its own, not shared
        }
      }
      // dividers and accent lines: borders along the block edges and through the gaps, never across a block
      plan.lines.forEach(ln => {
        /** @type {Partial<import("exceljs").Border>} */
        const side = { style: ln.style, color: { argb: ln.color } };
        for (let k = ln.from; k <= ln.to; k++) {
          const across = bgBlocks.some(b => ln.dir === "h" ? b.top < ln.at && ln.at <= b.bottom && k >= b.left && k <= b.right
                                                         : b.left < ln.at && ln.at <= b.right && k >= b.top && k <= b.bottom);
          if (across) continue;
          const cell = ln.dir === "h" ? worksheet.getCell(ln.at + 1, k + 1) : worksheet.getCell(k + 1, ln.at + 1);
          const key = ln.dir === "h" ? "top" : "left";
          if (cell.border && cell.border[key]) continue;                   // a table's own rule stays
          cell.style = { ...cell.style, border: { ...(cell.border || {}), [key]: side } };
        }
      });
    }

    setColumnWidths(worksheet, colWidths, exactWidths);
    applyAutoFilters(worksheet, allTablesInfo);
    if (FORMAT_CONFIG.printFitToWidth) {
      // printing / Save as PDF: the whole dashboard, one page wide, oriented like the dashboard
      const used = tracker.toRange();
      worksheet.pageSetup = { ...worksheet.pageSetup, orientation: dashExt.w >= dashExt.h ? "landscape" : "portrait",
        fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
        printArea: `A$1:${getExcelColumnName(used.e.c)}$${used.e.r + 1}` };          // ExcelJS adds the column's $
    }

    // final column widths / row heights in px: charts and pictures are fitted to their blocks with them
    const colPx = ci => Math.floor((worksheet.getColumn(ci + 1).width || 8.43) * 7 + 5);
    const rowPx = ri => {                            // collapsed rows take no space on screen
      const row = worksheet.getRow(ri + 1);
      if (row.hidden) return 0;
      return row.height ? Math.round(row.height * 4 / 3) : EXCEL_ROW_PX;
    };
    const blockPx = it => {
      let w = 0, h = 0;
      for (let c = it.gridCol; c < it.gridCol + it.gridW; c++) w += colPx(c);
      for (let r = it.gridRow; r < it.gridRow + it.allocatedRows; r++) h += rowPx(r);
      return { w, h };
    };
    /* a picture at (x, y) px from the block's top-left cell, w × h px, linked like its zone */
    const addPicture = async (image, col, row, x, y, w, h) => {
      const pic = await prepareImage(image.file, image.info, w, h);
      if (!pic) { console.warn(`[Images] ${image.zone.param}: cannot be drawn in this browser`); return false; }
      const id = workbook.addImage({ base64: pic.base64, extension: pic.extension });
      worksheet.addImage(id, /** @type {any} */ ({
        tl: nativeAnchor(col, row, x, y, colPx, rowPx), ext: { width: w, height: h }, editAs: "oneCell",
        ...(image.url ? { hyperlinks: { hyperlink: image.url, tooltip: image.url } } : {})
      }));
      return true;
    };
    let picturesAdded = 0;
    for (const item of adjustedItems.filter(it => it.image)) {
      const img = item.image, block = blockPx(item);
      // Fit Image / Center Image inside the block, never larger than Tableau draws it in its zone
      const fit = fitImage(img.info.width, img.info.height, block.w, block.h, img.zone.scaled, img.zone.centered);
      const cap = fitImage(img.info.width, img.info.height, img.widthPx, img.heightPx, img.zone.scaled, img.zone.centered);
      const w = Math.min(fit.w, cap.w), h = Math.min(fit.h, cap.h);
      const centered = img.zone.centered !== false;
      const x = centered ? Math.round((block.w - w) / 2) : 0, y = centered ? Math.round((block.h - h) / 2) : 0;
      if (await addPicture(img, item.gridCol, item.gridRow, x, y, w, h)) picturesAdded++;
    }
    for (const img of overlayImages) {
      // an icon over a sheet: the same spot of that sheet's block, scaled with the block
      const host = adjustedItems.find(it => it.layout && String(it.layout.id) === img.host);
      if (!host || !host.layout.widthPx || !host.layout.heightPx) continue;
      const block = blockPx(host);
      const sx = block.w / host.layout.widthPx, sy = block.h / host.layout.heightPx, k = Math.max(0.5, Math.min(1.5, sx, sy));
      const fit = fitImage(img.info.width, img.info.height, img.widthPx * k, img.heightPx * k, img.zone.scaled, img.zone.centered);
      if (await addPicture(img, host.gridCol, host.gridRow, Math.round((img.xPx - host.layout.xPx) * sx) + fit.x,
                           Math.round((img.yPx - host.layout.yPx) * sy) + fit.y, fit.w, fit.h)) picturesAdded++;
    }
    if (picturesAdded) console.log(`[Images] ${picturesAdded} picture(s) added`);

    setExportStatus("Building the Excel file…");
    /** @type {any} the XLSX bytes: ExcelJS's buffer, then injectCharts' Uint8Array */
    let buffer = await workbook.xlsx.writeBuffer();
    // fit-to-page next to collapsed row groups: ExcelJS writes <sheetPr> out of order – put it right
    if (worksheet.pageSetup.fitToPage && worksheet.properties.outlineProperties) buffer = await fixSheetProperties(buffer);
    if (chartJobs.length) {
      // fit every chart to its own block now that the final column widths/row heights are known:
      // tables widen the columns, so a chart kept at its Tableau pixel size would end short of the
      // tables / charts aligned with it on the dashboard → stretch it to the block's edges
      chartJobs.forEach(job => {
        let w = 0, h = -job.rowOffPx;
        for (let c = job.item.gridCol; c < job.item.gridCol + job.item.gridW; c++) w += colPx(c);
        for (let r = job.row; r < job.item.gridRow + job.item.allocatedRows; r++) h += rowPx(r);
        const panes = job.item.visualModel.chartSpecs.length;     // stacked panes share the block height
        job.widthPx = Math.max(160, w);
        if (job.item.pairedCard) { job.colOffPx = 2; job.widthPx = Math.max(160, w - 4); }     // the card's white edges
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
    } else {
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = exportFileName;
      link.click();
      URL.revokeObjectURL(link.href);
    }
    // the visual summary again ("Building the Excel file…" replaced it), plus where the file went
    updateVisualStatus(visualStatuses, workbookWarning);
    appendExportStatus(saveTarget ? `saved as ${saveTarget.name}` : `downloaded as ${exportFileName}`);

    console.log("✅ Export completed with Tableau formatting (fonts, colours, number formats, borders)");

  } catch (err) {
    console.error("[Export]", err);
    setExportStatus("Export failed: " + err.message, true);
    alert("Export failed. Check console (F12) for details.\n\n" + err.message);
  } finally {
    btn.disabled = false;
  }
}
