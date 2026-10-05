/* Button / icon worksheets: a custom shape drawn on empty shelves (zoom, info, navigation icons). Tableau
 * shows the icon; its fields are only flags for actions, so the export draws the icon, not a table. */
import { tfDvText } from "../data/values.js";
import { tfSameField } from "../twb/field-ref.js";
import { createSheetFormatter } from "../twb/formatter.js";
import { tfBucketKey } from "../twb/parser.js";

/**
 * The icon a worksheet shows when it is only Shape marks on empty Rows / Columns, without text: the shape
 * the workbook maps its current value to, else its Marks card shape. null = not an icon sheet.
 * @param {FormatModel | null} model @param {string} sheetName @param {SummaryData} summary
 * @returns {{ shape: string | null } | null} shape: a custom shape ("Zoom Icons/Zoom in.png"); null = one of
 *   Tableau's own shapes (circle, square …), which the export leaves out
 */
export function tvIconSheet(model, sheetName, summary) {
  const sheet = model && model.sheets ? model.sheets[sheetName] : null;
  if (!sheet || !sheet.panes.length || sheet.rows.length || sheet.cols.length) return null;
  if (!sheet.panes.every(p => /^shape$/i.test(p.markClass || ""))) return null;
  if (sheet.panes.some(p => p.labelRuns.length || p.encodings.some(e => e.channel === "text" || e.channel === "label"))) return null;
  const pane = sheet.panes[0];
  let shape = null;
  // Shape = a field: the shape of the value it has now (the workbook's map, sheet before data source)
  const enc = pane.encodings.find(e => e.channel === "shape");
  if (enc) {
    const styles = [sheet.style, ...Object.values(model.datasourceStyles || {}), model.workbookStyle];
    const def = styles.flatMap(s => (s && s.shapes) || []).find(s => tfSameField(s.field, enc.field));
    const fmt = createSheetFormatter(model, sheetName);
    const ci = (summary.columns || []).findIndex(c => tfSameField(fmt.matchName(c.fieldName), enc.field));
    // the first mark whose value has a shape (the marks of a button sit on one spot)
    if (def && ci >= 0) shape = (summary.data || []).map(r => def.map[tfBucketKey(tfDvText(r[ci]))]).find(Boolean) || null;
  }
  // else the Marks card shape
  if (!shape) {
    const f = ((pane.style && pane.style.rules && pane.style.rules.mark) || []).find(x => x.attr === "shape" && !x.field);
    shape = f ? f.value : null;
  }
  return { shape: shape && !shape.startsWith(":") ? shape : null };     // ":filled/circle" … = Tableau's own
}
