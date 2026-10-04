/* Entry point: wires the panel (Load Workbook, Export) once Tableau has initialised the extension. */
import { exportToExcel } from "./export/export.js";
import { showWorkbookLabel } from "./ui/status.js";
import { RESTORED_FILE_NAME, ensureFormatModel, loadWorkbookFile, setFormatModelCache } from "./ui/workbook-store.js";

document.addEventListener("DOMContentLoaded", () => {
  tableau.extensions.initializeAsync().then(() => {

    const loadBtn = /** @type {HTMLButtonElement} */ (document.getElementById("load_workbook_btn"));
    if (loadBtn) {
      loadBtn.addEventListener("click", async () => {
        loadBtn.disabled = true;
        loadBtn.textContent = "⏳ Loading...";
        await loadWorkbookFile();
        loadBtn.disabled = false;
        loadBtn.textContent = "📁 Load Workbook";
      });
    }

    const fileLabel = document.getElementById("twb_file_label");
    if (fileLabel) {
      const savedFileName = tableau.extensions.settings.get("twbFileName");
      const savedTitleMap = tableau.extensions.settings.get("twbTitleMap");
      const savedFormat = tableau.extensions.settings.get("twbFormatModel");
      if (savedFileName && savedTitleMap) {
        const titleCount = Object.keys(JSON.parse(savedTitleMap)).length;
        if (savedFormat) setFormatModelCache(JSON.parse(savedFormat));
        ensureFormatModel().then(model => {
          if (model) {
            showWorkbookLabel(savedFileName, `${titleCount} titles, formatting for ${Object.keys(model.sheets || {}).length} sheets`, model);
          } else {
            fileLabel.textContent = `${savedFileName} — formatting not available after reload.
⚠ Click 📁 Load Workbook again so charts are exported as charts`;
            fileLabel.classList.add("status-warning");
          }
        });
      } else {
        fileLabel.textContent = "No workbook loaded — click 📁 to load titles and colors";
        ensureFormatModel().then(model => {
          if (model && RESTORED_FILE_NAME) {
            showWorkbookLabel(RESTORED_FILE_NAME, `remembered from an earlier load, formatting for ${Object.keys(model.sheets || {}).length} sheets`, model);
          } else {
            fileLabel.textContent = "⚠ No workbook loaded — click 📁 Load Workbook, otherwise charts are exported as tables";
            fileLabel.classList.add("status-warning");
          }
        });
      }
    }

    const exportBtn = /** @type {HTMLButtonElement} */ (document.getElementById("export_button"));
    if (exportBtn) {
      exportBtn.addEventListener("click", exportToExcel);
    }


  }).catch((err) => {
    console.error("[Tableau init]", err);
    alert("Failed to connect to Tableau: " + err.message);
    if (document.getElementById("export_button")) {
      /** @type {HTMLButtonElement} */ (document.getElementById("export_button")).disabled = true;
    }
  });
});
