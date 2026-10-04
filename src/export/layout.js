/* Dashboard layout → Excel grid: positions, sizes and collision handling. */
import { FORMAT_CONFIG } from "../config.js";

/* ── Layout constants ─────────────────────────────────────────────────── */
export const PX_PER_COL  = 90;

export const PX_PER_ROW  = 22;

export const TITLE_ROWS  = 1;

export const COL_GAP     = 0;

export const ROW_GAP     = 1;

export const ROW_GROUP_THRESHOLD = 8;

export const EXCEL_ROW_PX = 20;

// default Excel row (15pt)
export const EXCEL_COL_PX = 75;

// a 10-character column – the narrowest width setColumnWidths assigns
export const CHART_DATA_SHEET = "Chart Data";

/* pixel box of a chart/image visual: the dashboard zone size minus Tableau's title strip */
export function graphicBox(layout, vm) {
  const widthPx = Math.max(160, Math.round((layout && layout.widthPx) || 480));
  const zoneH = Math.round((layout && layout.heightPx) || 320);
  const titleRows = vm.showTitle ? 1 : 0;
  const heightPx = Math.max(120, zoneH - (titleRows ? 28 : 0));
  return { widthPx, heightPx, titleRows,
           gridW: Math.max(2, Math.ceil(widthPx / EXCEL_COL_PX)),
           rows: titleRows + Math.ceil(heightPx / EXCEL_ROW_PX) };
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

/* ── buildLayoutMap ─────────────────────────────────────────────────── */
export function buildLayoutMap(dashboardObjects, titleMap = {}) {
  const map = new Map();
  
  const positionableObjects = (dashboardObjects || []).filter(
    (obj) => (obj.type === "worksheet" || obj.type === "filter" || obj.type === "parameter") &&
             obj.position &&
             typeof obj.position.x === "number" &&
             typeof obj.position.y === "number"
  );

  if (positionableObjects.length === 0) return map;

  let minX = Infinity, minY = Infinity;
  positionableObjects.forEach((obj) => {
    if (obj.position.x < minX) minX = obj.position.x;
    if (obj.position.y < minY) minY = obj.position.y;
  });

  positionableObjects.forEach((obj) => {
    const px = obj.position;
    const gridCol = Math.round((px.x - minX) / PX_PER_COL);
    const gridRow = Math.round((px.y - minY) / PX_PER_ROW);
    const gridW = Math.max(2, Math.round((px.width || 180) / PX_PER_COL));
    const gridH = Math.max(3, Math.round((px.height || 60) / PX_PER_ROW));
    
    let displayName = "";
    if (obj.type === "worksheet") {
      displayName = titleMap[obj.name]
                 || (obj.title && obj.title.trim() ? obj.title.trim() : null)
                 || obj.name;
    } else {
      displayName = (obj.name || "Filter").replace(/[_-]/g, " ");
    }

    map.set(obj.name || `filter_${gridRow}_${gridCol}`, {
      type: obj.type,
      gridRow: Math.max(0, gridRow),
      gridCol: Math.max(0, gridCol),
      gridW,
      gridH,
      widthPx: px.width || null,
      heightPx: px.height || null,
      displayName,
      originalName: obj.name
    });
  });

  return map;
}

/* rows the writer will produce – used for layout before writing */
export function viewModelHeight(vm) {
  const grouped = FORMAT_CONFIG.groupOverflowRows && vm.rows.length > ROW_GROUP_THRESHOLD;
  return (vm.showTitle ? 1 : 0) + (vm.showHeaderRow ? 1 : 0) + vm.rows.length + (grouped ? 1 : 0);
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

// ── Cascade vertical pushes until stable (not just one pass) ──
processedGroups.sort((a, b) => a.minRow - b.minRow);
let changed = true;
while (changed) {
  changed = false;
  for (let i = 0; i < processedGroups.length - 1; i++) {
    const upper = processedGroups[i];
    const lower = processedGroups[i + 1];
    let horizontalOverlap = false;
    for (const u of upper.items) {
      for (const l of lower.items) {
        const uLeft = u.gridCol, uRight = u.gridCol + u.gridW;
        const lLeft = l.gridCol, lRight = l.gridCol + l.gridW;
        if (uLeft < lRight && uRight > lLeft) { horizontalOverlap = true; break; }
      }
      if (horizontalOverlap) break;
    }
    if (horizontalOverlap && lower.minRow < upper.bottom) {
      const pushBy = upper.bottom - lower.minRow;
      lower.items.forEach(z => { z.gridRow += pushBy; });
      lower.minRow += pushBy;
      lower.bottom += pushBy;
      changed = true;
    }
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
