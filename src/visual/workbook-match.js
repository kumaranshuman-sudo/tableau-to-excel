/* Does the loaded workbook belong to the running dashboard? */
import { tfNorm } from "../util.js";

/* does the loaded workbook belong to the running dashboard? (a different or stale .twb would
 * silently apply the wrong titles, colours and sorts) → { level: "ok" | "partial" | "mismatch", … } */
export function checkWorkbookMatch(model, dashboard) {
  if (!model || !dashboard) return null;
  const names = (dashboard.worksheets || []).map(w => w.name);
  const known = new Set(Object.keys(model.sheets || {}).map(tfNorm));
  const missing = names.filter(n => !known.has(tfNorm(n)));
  const matched = names.length - missing.length;
  const dashboardFound = Object.keys(model.dashboards || {}).some(d => tfNorm(d) === tfNorm(dashboard.name));
  const level = !missing.length && dashboardFound ? "ok"
    : names.length && matched / names.length >= 0.5 ? "partial" : "mismatch";
  return { level, dashboard: dashboard.name, dashboardFound, matched, total: names.length, missing };
}

export function describeWorkbookMatch(match) {
  if (!match || match.level === "ok") return "";
  const parts = [];
  if (!match.dashboardFound) parts.push(`dashboard "${match.dashboard}" is not in this workbook`);
  if (match.missing.length) {
    parts.push(`${match.matched}/${match.total} worksheets found` +
      ` (missing: ${match.missing.slice(0, 3).join(", ")}${match.missing.length > 3 ? ", …" : ""})`);
  }
  return (match.level === "mismatch" ? "⚠ Wrong workbook? " : "⚠ Workbook partly matches: ") + parts.join("; ");
}
