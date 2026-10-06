/* Panel status line and loaded-workbook label. */
import { checkWorkbookMatch, describeWorkbookMatch } from "../visual/workbook-match.js";

/* loaded-workbook label, with a warning when the file does not match this dashboard */
export function showWorkbookLabel(fileName, details, model) {
  const label = document.getElementById("twb_file_label");
  if (!label) return;
  const match = checkWorkbookMatch(model, tableau.extensions.dashboardContent.dashboard);
  const warning = describeWorkbookMatch(match);
  label.textContent = `${warning ? "" : "✅ "}${fileName} — ${details}` + (warning ? `\n${warning}` : "");
  label.title = match && match.missing.length ? "Not found in the workbook:\n" + match.missing.join("\n") : "";
  label.classList.toggle("status-warning", !!warning);
  if (warning) console.warn("[Workbook]", warning, match);
}

/* progress / outcome line under the Export button */
export function setExportStatus(text, warning) {
  const target = document.getElementById("export_status");
  if (!target) return;
  target.textContent = text;
  target.title = "";
  target.classList.toggle("status-warning", !!warning);
}

export function appendExportStatus(text) {
  const target = document.getElementById("export_status");
  if (target && target.textContent) target.textContent += " · " + text;
}

export function updateVisualStatus(statuses, note) {
  const target = document.getElementById("export_status");
  if (!target) return;
  if (note) {                                      // e.g. the loaded workbook does not match
    updateVisualStatus(statuses);
    target.textContent = note + " · " + target.textContent;
    target.classList.add("status-warning");
    return;
  }
  if (!statuses || !statuses.length) {
    target.textContent = "No worksheet visual status available";
    return;
  }
  // once converted: how (NATIVE … TABLE_FALLBACK, as on the report sheet); before that, which renderer
  if (statuses.every(item => item.strategy)) {
    const order = ["NATIVE", "CONSTRUCTED", "APPROXIMATE", "TABLE_FALLBACK", "IMAGE_FALLBACK", "UNSUPPORTED", "FAILED"];
    const words = { NATIVE: "native", CONSTRUCTED: "constructed", APPROXIMATE: "approximated", TABLE_FALLBACK: "as data tables",
                    IMAGE_FALLBACK: "as pictures", UNSUPPORTED: "unsupported", FAILED: "failed" };
    const parts = order.map(s => [statuses.filter(item => item.strategy === s).length, words[s]]).filter(([n]) => n);
    const flagged = statuses.filter(item => item.strategy !== "NATIVE" && item.strategy !== "CONSTRUCTED");
    target.textContent = `${statuses.length} visuals: ${parts.map(([n, w]) => `${n} ${w}`).join(", ")} – see the Conversion Report sheet`;
    target.title = flagged.map(w => `${w.worksheet} (${w.visual}): ${w.strategy} – ${w.reason}`).join("\n");
    target.classList.toggle("status-warning", statuses.some(item => item.strategy === "FAILED"));
    return;
  }
  const count = renderer => statuses.filter(item => item.renderer === renderer).length;
  const parts = [
    [count("cell"), "tables/KPIs"],
    [count("excel-chart"), "Excel charts"],
    [count("tableau-image"), "Tableau images"],
    [count("data-fallback"), "as data"]
  ].filter(([n]) => n).map(([n, label]) => `${n} ${label}`);
  const warnings = statuses.filter(item => item.status !== "success" && item.status !== "pending");
  target.textContent = `${statuses.length} visuals: ${parts.join(", ")}` +
    (warnings.length ? ` – ${warnings.length} warning${warnings.length === 1 ? "" : "s"} (see console)` : "");
  target.title = warnings.map(w => `${w.worksheet}: ${w.reason}`).join("\n");
  target.classList.toggle("status-warning", warnings.length > 0);
}
