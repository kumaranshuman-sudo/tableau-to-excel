/* Dashboard layout → Excel grid: positions, sizes and collision handling. */
import { FORMAT_CONFIG } from "../config.js";
import { classifyImageObjects } from "./images.js";

/* ── Layout constants ─────────────────────────────────────────────────── */
export const PX_PER_COL  = 90;

export const PX_PER_ROW  = 20;          // one Excel row (15 pt): dashboard y maps 1:1 onto the sheet

export const TITLE_ROWS  = 1;

export const COL_GAP     = 0;

export const ROW_GAP     = 0;            // Tableau zones touch; their padding is inside the zone

export const ROW_GROUP_THRESHOLD = 8;

export const EXCEL_ROW_PX = 20;

// default Excel row (15pt)
export const EXCEL_COL_PX = 75;

// a 10-character column – the narrowest width setColumnWidths assigns
export const CHART_DATA_SHEET = "Chart Data";

/* pixel box of a chart/image visual: the dashboard zone size minus Tableau's title strip */
export function graphicBox(layout, vm) {
  const widthPx = Math.max(40, Math.round((layout && layout.widthPx) || 480));
  const zoneH = Math.round((layout && layout.heightPx) || 320);
  const titleRows = vm.showTitle ? 1 : 0;
  const heightPx = Math.max(60, zoneH - (titleRows ? 28 : 0));
  return { widthPx, heightPx, titleRows,
           // the zone's width on the dashboard grid, so the block lines up with the blocks above / below it
           gridW: layout && layout.widthPx ? layout.gridW : Math.max(2, Math.ceil(widthPx / EXCEL_COL_PX)),
           // the zone's own height on the sheet, so the block below starts where it does on the dashboard
           rows: Math.max(titleRows + 1, Math.round(zoneH / EXCEL_ROW_PX)) };
}

/* zone edges closer than this (px) are one column boundary */
export const EDGE_MERGE_PX = 6;

/**
 * Excel columns from the dashboard's own zone edges: every left / right edge of a laid-out block is a column
 * boundary, so each column is as wide (px) as the gap between two edges and every block spans exactly its
 * zone – positions and widths as on the dashboard, whatever the content. A table zone gets one column per
 * field: at the widths the table draws (splitPx, from the zone's left edge; scaled down only when they would
 * run past the zone), else split by relative widths.
 * @param {{ layout?: any, split?: number[], splitPx?: number[] }[]} items
 * @returns {{ colPx: number[], span: (item: any) => { gridCol: number, gridW: number } | null }}
 */
export function buildColumnGrid(items) {
  const laid = items.filter(it => it.layout && typeof it.layout.xPx === "number" && it.layout.widthPx > 0);
  if (!laid.length) return { colPx: [], span: () => null };
  const minX = Math.min(...laid.map(it => it.layout.xPx));
  const merge = (list, tol = EDGE_MERGE_PX) => {
    const out = [];
    [...list].sort((a, b) => a - b).forEach(e => { if (!out.length || e - out[out.length - 1] > tol) out.push(e); });
    return out;
  };
  const sideOf = it => [it.layout.xPx - minX, it.layout.xPx - minX + it.layout.widthPx];
  let edges = merge(laid.flatMap(sideOf));
  const nearest = v => edges.reduce((best, e, i) => Math.abs(e - v) < Math.abs(edges[best] - v) ? i : best, 0);
  const extra = [];
  /** @type {[number, number[]][]} a table's span and its own column edges */
  const tables = [];
  laid.forEach(it => {
    if (it.splitPx && it.splitPx.length) {
      const [l, r] = sideOf(it), L = edges[nearest(l)], R = edges[nearest(r)];
      const total = it.splitPx.reduce((a, b) => a + b, 0);
      const k = total > R - L ? (R - L) / total : 1;
      let at = L;
      const own = [L];
      it.splitPx.forEach(w => { at += w * k; own.push(Math.round(at)); });
      extra.push(...own);
      tables.push([L, own]);
      return;
    }
    if (!it.split || it.split.length < 2) return;
    const [l, r] = sideOf(it), L = edges[nearest(l)], R = edges[nearest(r)];
    if (edges.filter(e => e > L && e < R).length + 1 >= it.split.length) return;
    const total = it.split.reduce((a, b) => a + b, 0);
    let at = L;
    it.split.slice(0, -1).forEach(w => { at += (R - L) * w / total; extra.push(Math.round(at)); });
  });
  edges = merge([...edges, ...extra], 2);
  // inside a table only its own column edges: another block's edge there would cut a field's column in two
  tables.forEach(([L, own]) => {
    const end = own[own.length - 1];
    edges = edges.filter(e => e <= L || e >= end || own.some(o => Math.abs(o - e) <= 2));
  });
  const colPx = edges.slice(1).map((e, i) => e - edges[i]);
  return {
    colPx,
    span: it => {
      if (!laid.includes(it)) return null;
      const [l, r] = sideOf(it), a = nearest(l), b = nearest(r);
      return { gridCol: a, gridW: Math.max(1, b - a) };
    }
  };
}

export function getExcelColumnName(colIndex) {
  let columnName = "";
  let dividend = colIndex + 1;

  while (dividend > 0) {
    let modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

/* ── Range tracker ────────────────────────────────────────────────────── */
export function makeRangeTracker() {
  let minR = Infinity, minC = Infinity, maxR = 0, maxC = 0;
  return {
    update(r, c) {
      if (r < minR) minR = r;
      if (c < minC) minC = c;
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    },
    toRange() {
      return {
        s: { r: minR === Infinity ? 0 : minR, c: minC === Infinity ? 0 : minC },
        e: { r: maxR, c: maxC },
      };
    },
  };
}

/* edges a few px apart (tiled containers with padding) → one grid line, so blocks that line up on
 * the dashboard land in the same row / column */
export const SNAP_PX = 8;

/** @param {number[]} values @param {number} [tolerance] @returns {(v: number) => number} */
export function snapEdges(values, tolerance = SNAP_PX) {
  const anchor = new Map();
  let current = null;
  [...new Set(values)].sort((a, b) => a - b).forEach(v => {
    if (current === null || v - current > tolerance) current = v;
    anchor.set(v, current);
  });
  return v => anchor.has(v) ? anchor.get(v) : v;
}

/* ── buildLayoutMap ─────────────────────────────────────────────────── */
/* Tableau's DashboardObject: position = { x, y }, size = { width, height } (px) */
/* Tableau's DashboardObjectType → the kinds laid out on the sheet ("filter" / "parameter" as older
 * test fixtures spell them); a DashboardObject's id is its zone id in the workbook */
export const OBJECT_KIND = Object.freeze({ worksheet: "worksheet", "quick-filter": "filter", filter: "filter",
                                           "parameter-control": "parameter", parameter: "parameter", text: "text", image: "image" });

export function buildLayoutMap(dashboardObjects, titleMap = {}) {
  const map = new Map();
  // images with their own zone (logos) take part in the layout; backgrounds, dividers and icons
  // floating over a sheet do not (export/images.js)
  const images = classifyImageObjects(dashboardObjects || []);

  const positionableObjects = (dashboardObjects || []).filter(
    (obj) => OBJECT_KIND[obj.type] &&
             obj.position &&
             typeof obj.position.x === "number" &&
             typeof obj.position.y === "number" &&
             (obj.type !== "image" || (images.get(String(obj.id)) || {}).kind === "block")
  );

  if (positionableObjects.length === 0) return map;

  let minX = Infinity, minY = Infinity;
  positionableObjects.forEach((obj) => {
    if (obj.position.x < minX) minX = obj.position.x;
    if (obj.position.y < minY) minY = obj.position.y;
  });

  const sizeOf = obj => obj.size || obj.position;
  const snapX = snapEdges(positionableObjects.flatMap(o => [o.position.x, o.position.x + (sizeOf(o).width || 0)]));
  const snapY = snapEdges(positionableObjects.map(o => o.position.y));

  positionableObjects.forEach((obj) => {
    const px = obj.position;
    const { width, height } = sizeOf(obj);
    const gridCol = Math.round((snapX(px.x) - minX) / PX_PER_COL);
    const gridRow = Math.round((snapY(px.y) - minY) / PX_PER_ROW);
    const gridW = Math.max(2, Math.round((width || 180) / PX_PER_COL));
    const gridH = Math.max(3, Math.round((height || 60) / PX_PER_ROW));
    
    const kind = OBJECT_KIND[obj.type];
    let displayName = "";
    if (kind === "text" || kind === "image") displayName = "";
    else if (kind === "worksheet") {
      displayName = titleMap[obj.name]
                 || (obj.title && obj.title.trim() ? obj.title.trim() : null)
                 || obj.name;
    } else {
      displayName = (obj.name || "Filter").replace(/[_-]/g, " ");
    }

    // text boxes and images have no unique name – keyed by their zone id
    const key = kind === "text" || kind === "image" ? `${kind}:${obj.id}` : obj.name || `filter_${gridRow}_${gridCol}`;
    // a quick filter can carry its worksheet's name – the worksheet keeps its own position and size
    if (kind !== "worksheet" && map.has(key) && map.get(key).type === "worksheet") return;
    map.set(key, {
      type: kind,
      id: obj.id,
      gridRow: Math.max(0, gridRow),
      gridCol: Math.max(0, gridCol),
      gridW,
      // right edge on the grid: tiles in a strip end where the next one starts
      gridRight: width ? Math.round((snapX(px.x + width) - minX) / PX_PER_COL) : gridCol + gridW,
      gridH,
      widthPx: width || null,
      heightPx: height || null,
      xPx: px.x,
      yPx: px.y,
      displayName,
      originalName: obj.name
    });
  });

  return map;
}

/* ── grouped table rows ─────────────────────────────────────────────────
 * A long table shows its first ROW_GROUP_THRESHOLD rows and collapses the rest into an outline group.
 * Excel hides whole sheet rows, so a chart / image / card beside the table would lose the same rows:
 * the table keeps visible every row such a neighbour uses (it then fills the band like the Tableau
 * zone) and only the rows below are collapsed. Tables beside each other keep their own threshold. */
export function isGroupedTable(item) {
  return !!(FORMAT_CONFIG.groupOverflowRows && item.vm && item.type === "worksheet" && !item.isKPI && !item.box &&
            item.vm.rows.length > ROW_GROUP_THRESHOLD);
}

/**
 * Sheet rows a table's header takes: the height Tableau stores for it in 20 px rows (merged down), so a tall
 * header keeps the sheet's rows at 20 px – a single tall row would push the blocks beside the table down.
 * @param {ViewModel} vm
 */
export function headerRowSpan(vm) {
  if (!vm.showHeaderRow) return 0;
  const hpx = vm.fmt && vm.fmt.headerRowHeightPx ? vm.fmt.headerRowHeightPx() : 0;
  return hpx ? Math.max(1, Math.round(hpx / PX_PER_ROW)) : 1;
}

/* sheet row of a table's first data row */
export function tableDataStart(item) {
  return item.gridRow + (item.vm.showTitle ? 1 : 0) + headerRowSpan(item.vm);
}

/** sets item.visibleRows on every table that will be grouped @param {ExportItem[]} items */
export function setTableVisibleRows(items) {
  items.forEach(t => {
    if (!isGroupedTable(t)) return;
    const start = tableDataStart(t);
    const end = t.gridRow + t.allocatedRows;
    let need = ROW_GROUP_THRESHOLD;
    items.forEach(o => {
      if (o === t || isGroupedTable(o)) return;
      const oEnd = o.gridRow + (o.allocatedRows || o.rowCount || 0);
      if (o.gridRow < end && oEnd > t.gridRow) need = Math.max(need, oEnd - start);   // shares sheet rows
    });
    t.visibleRows = Math.min(t.vm.rows.length, need);
  });
}

/* rows the writer will produce – used for layout before writing */
export function viewModelHeight(vm) {
  const grouped = FORMAT_CONFIG.groupOverflowRows && vm.rows.length > ROW_GROUP_THRESHOLD;
  return (vm.showTitle ? 1 : 0) + headerRowSpan(vm) + vm.rows.length + (grouped ? 1 : 0);
}

export function resolveCollisions(zones) {
if (!zones || zones.length === 0) return zones;

// Full row/header/note height a zone actually occupies on the sheet,
// including its own hidden-but-physically-present grouped rows.
function getVisualHeight(zone) {
  return (zone.allocatedRows || zone.rowCount || 5) + ROW_GAP;
}

// ── Group by EXACT gridRow (not a rounded bucket) ──
const rowGroups = new Map();
zones.forEach(zone => {
  const key = zone.gridRow;
  if (!rowGroups.has(key)) rowGroups.set(key, []);
  rowGroups.get(key).push(zone);
});

// ── Horizontal packing within each group ──
const processedGroups = [];
for (const [, group] of rowGroups) {
  group.sort((a, b) => a.gridCol - b.gridCol);
  const minRow = Math.min(...group.map(z => z.gridRow));
  group.forEach(z => { z.gridRow = minRow; });
  let cursor = group[0].gridCol;
  for (const item of group) {
    if (item.gridCol < cursor) item.gridCol = cursor;
    cursor = item.gridCol + item.gridW + COL_GAP;
  }
  const groupBottom = minRow + Math.max(...group.map(z => getVisualHeight(z))) + ROW_GAP;
  processedGroups.push({ minRow, items: group, bottom: groupBottom });
}

// ── Vertical pushes: a row group moves down as a whole, below every earlier group it overlaps
//    (not only the group just above), so blocks that share a top edge on the dashboard stay level ──
processedGroups.sort((a, b) => a.minRow - b.minRow);
const groupsOverlap = (upper, lower) => upper.items.some(u => lower.items.some(l =>
  u.gridCol < l.gridCol + l.gridW && u.gridCol + u.gridW > l.gridCol));
for (let i = 1; i < processedGroups.length; i++) {
  const lower = processedGroups[i];
  let pushBy = 0;
  for (let j = 0; j < i; j++) {
    if (groupsOverlap(processedGroups[j], lower)) pushBy = Math.max(pushBy, processedGroups[j].bottom - lower.minRow);
  }
  if (pushBy > 0) {
    lower.items.forEach(z => { z.gridRow += pushBy; });
    lower.minRow += pushBy;
    lower.bottom += pushBy;
  }
}

// ── Full pairwise safety-net scan across ALL zones, not just adjacent groups ──
let fullPassChanged = true;
const MAX_PASSES = 20;
let pass = 0;
while (fullPassChanged && pass < MAX_PASSES) {
  fullPassChanged = false;
  pass++;
  for (let i = 0; i < zones.length; i++) {
    for (let j = i + 1; j < zones.length; j++) {
      const a = zones[i], b = zones[j];
      const upper = a.gridRow <= b.gridRow ? a : b;
      const lower = a.gridRow <= b.gridRow ? b : a;
      const upperBottom = upper.gridRow + getVisualHeight(upper) + ROW_GAP;
      const colOverlap = upper.gridCol < lower.gridCol + lower.gridW &&
                          upper.gridCol + upper.gridW > lower.gridCol;
      const rowOverlap = lower.gridRow < upperBottom;
      if (colOverlap && rowOverlap) {
        lower.gridRow += (upperBottom - lower.gridRow);
        fullPassChanged = true;
      }
    }
  }
}

return zones;
}
