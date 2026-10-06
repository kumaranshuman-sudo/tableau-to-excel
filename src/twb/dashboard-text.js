/* Dashboard text boxes: header/title runs and zone borders. */
import { FORMAT_CONFIG } from "../config.js";

/* Text boxes count as column headers only if ALL hold:
 *   – ≥ 2 text zones, visible, in one row (same top edge ± tol)
 *   – their bottom edge touches the sheet's top edge (± tol)
 *   – they lie within the sheet's horizontal span
 *   – their number equals the number of visible columns
 * Positions are Tableau's own zone coordinates (0–100000), so the check is deterministic. */
export const TF_ZONE_TOL = 1500;

export function tfZoneText(z) { return (z.runs || []).map(r => r.text).join("").replace(/\u00C6\r?\n?/g, "\n").trim(); }

export function tfTextBoxHeaders(dash, sheetName, nCols) {
  const strip = tfTextBoxHeaderStrip(dash, sheetName);
  if (!strip || nCols < 2 || strip.zones.length !== nCols) return null;
  return strip;
}

/* the row of text boxes sitting on a worksheet's top edge, left to right (≥ 2), whatever their number */
export function tfTextBoxHeaderStrip(dash, sheetName) {
  const ws = dash.zones.find(z => z.name === sheetName && z.type === "worksheet");
  if (!ws) return null;
  const near = (a, b) => Math.abs(a - b) <= TF_ZONE_TOL;
  const cand = dash.zones.filter(z => z.type === "text" && !z.hidden && z.runs && tfZoneText(z) &&
    near(z.y + z.h, ws.y) && z.x >= ws.x - TF_ZONE_TOL && z.x + z.w <= ws.x + ws.w + TF_ZONE_TOL);
  if (cand.length < 2 || !cand.every(z => near(z.y, cand[0].y))) return null;
  cand.sort((a, b) => a.x - b.x);
  return { ws, zones: cand.map(z => ({ id: z.id, text: tfZoneText(z), x: z.x, w: z.w, h: z.h, runs: z.runs,
    props: (z.runs.find(r => r.text.trim()) || { props: {} }).props })) };
}

/* Dashboard title: Tableau's own title if shown, else the text box above every worksheet
 * with the largest font (ties → topmost). Returns runs for rich text, or null. */
export function tfDashboardTitleRuns(model, dashName, usedZoneIds) {
  const dash = model && model.dashboards && model.dashboards[dashName];
  if (!dash) return null;
  const titleZone = dash.zones.find(z => z.type === "title" && !z.hidden);
  if (titleZone && dash.title && dash.title.runs.length) return dash.title.runs;
  if (!FORMAT_CONFIG.textBoxTitle) return null;
  const wsTop = Math.min(...dash.zones.filter(z => z.type === "worksheet").map(z => z.y));
  const size = z => (z.runs.find(r => r.text.trim()) || { props: {} }).props.fontSize || 9;
  const cand = dash.zones.filter(z => z.type === "text" && !z.hidden && z.runs && tfZoneText(z) &&
    !usedZoneIds.has(z.id) && z.y + z.h <= wsTop + TF_ZONE_TOL);
  if (!cand.length) return null;
  cand.sort((a, b) => size(b) - size(a) || a.y - b.y);
  return cand[0].runs;
}

export function tfStrokeToBorder(size) {
  if (!size) return null;
  if (size <= 1) return "thin";
  if (size === 2) return "medium";
  return "thick";
}
