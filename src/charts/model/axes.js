/* The worksheet's own axis and label settings on a chart spec: hidden axes, axis titles, Edit Axis ranges
 * and tick spacing, tick number formats, grid lines, axis rulers, mark label font, position and text. */
import { tableauToExcelNumFmt, tfWrapNumFmt } from "../../format/number-format.js";
import { tfSameField } from "../../twb/field-ref.js";

/* Edit Axis → the spec: fixed range ("fixed" / "fixedmin" / "fixedmax"), tick spacing, include zero */
function applySpace(spec, s, axis) {
  const key = name => axis === "x" ? "x" + name : "value" + name;
  if (s.majorSpacing) spec[key("MajorUnit")] = s.majorSpacing;
  if (/^fixed(min)?$/.test(s.rangeType || "") && s.min !== undefined) spec[key("Min")] = s.min;
  if (/^fixed(max)?$/.test(s.rangeType || "") && s.max !== undefined) spec[key("Max")] = s.max;
  if (s.domainExpand === "false" && axis !== "x") spec.includeZero = false;
}

/* an axis's tick format (Format → Axis → Numbers: the field's "label" text-format) */
const axisNumFmt = (fmt, ref) => (ref && tableauToExcelNumFmt(fmt.labelFormat(ref))) || undefined;

/* mark labels: the worksheet's label font, for bars the label alignment as Excel's label position, and a
 * label text around the value ("<AGG(Days)> DAYS") as literal text in the label's number format */
function applyLabels(spec, fmt, measures) {
  const dl = fmt.dataLabelStyle();
  if (dl.fontName || dl.fontSize || dl.color || dl.bold !== undefined) {
    spec.labelFont = { name: dl.fontName, size: dl.fontSize, color: dl.color ? dl.color.slice(2) : undefined, bold: dl.bold };
  }
  const tpl = fmt.labelTemplate();
  const at = tpl && (tpl.prefix.trim() || tpl.suffix.trim()) ? measures.findIndex(m => m && tfSameField(m, tpl.ref)) : -1;
  if (at === 0 || at === 1) {
    spec.series.forEach(s => {
      if (s.refLine || !!s.secondary !== (at === 1 && measures.length > 1) || s.labelTexts) return;
      s.labelNumFmt = tfWrapNumFmt(s.labelNumFmt || (s.secondary ? spec.secondaryNumFmt : spec.numFmt) || "General", tpl.prefix, tpl.suffix);
    });
  }
  if (spec.kind === "bar" || spec.kind === "combo") {
    const cell = fmt.markCellStyle(null);
    const pos = spec.barDir === "bar" ? { left: "inBase", center: "ctr", right: "inEnd" }[cell.hAlign]
                                     : { bottom: "inBase", middle: "ctr", top: "inEnd" }[cell.vAlign];
    if (pos) spec.labelPos = pos;
  }
}

/**
 * Tableau hides axes ("Show Header" off), renames or removes axis titles (Edit Axis), and turns grid
 * lines and axis rulers off per worksheet; the Excel chart follows the same settings.
 * @param {ChartSpec} spec @param {ChartContext} ctx
 */
export function tvApplyWorkbookAxes(spec, ctx) {
  const { vm, roles } = ctx;
  const fmt = vm.fmt;
  if (!fmt.hasModel || roles.source !== "twb" || spec.axesHidden || /^(pie|doughnut|treemap)$/.test(spec.kind)) return;

  if (spec.kind === "scatter" || spec.kind === "bubble") {
    const x = roles.cols.values[0], y = roles.rows.values[0];
    const xi = x ? fmt.axisInfo(x.ref, "cols", "0") : {}, yi = y ? fmt.axisInfo(y.ref, "rows", "0") : {};
    if (xi.hidden) spec.xAxisHidden = true;
    if (xi.title !== undefined) spec.xTitle = xi.title;
    if (yi.hidden) spec.valueAxisHidden = true;
    if (yi.title !== undefined) spec.valueTitle = yi.title;
    if (!fmt.gridlinesShown("rows")) spec.gridlines = false;
    if (!fmt.gridlinesShown("cols")) spec.xGridlines = false;
    if (x) applySpace(spec, fmt.axisSpace(x.ref, "cols", "0"), "x");
    if (y) applySpace(spec, fmt.axisSpace(y.ref, "rows", "0"), "value");
    spec.xAxisNumFmt = axisNumFmt(fmt, x && x.ref);
    spec.valueAxisNumFmt = axisNumFmt(fmt, y && y.ref);
    applyLabels(spec, fmt, [y && y.ref, x && x.ref]);
    return;
  }

  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) return;
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  const refs = roles[valueShelf].axisRefs.length ? roles[valueShelf].axisRefs : roles[valueShelf].values.map(v => v.ref);
  const primary = refs[0] ? fmt.axisInfo(refs[0], valueShelf, "0") : {};
  if (primary.hidden) spec.valueAxisHidden = true;
  if (primary.title !== undefined) spec.valueTitle = primary.title;
  if (refs[1] && spec.series.some(s => s.secondary)) {
    // class = which axis of that field: a second measure has its own first axis ("0"); the same field on
    // both axes has its second one ("1")
    const second = fmt.axisInfo(refs[1], valueShelf, tfSameField(refs[0], refs[1]) ? "1" : "0");
    if (second.hidden) spec.secondaryAxisHidden = true;
    if (second.title !== undefined) spec.secondaryTitle = second.title;
  }
  if (!fmt.gridlinesShown(valueShelf)) spec.gridlines = false;
  if (refs[0]) applySpace(spec, fmt.axisSpace(refs[0], valueShelf, "0"), "value");
  spec.valueAxisNumFmt = axisNumFmt(fmt, refs[0]);
  if (refs[1]) spec.secondaryAxisNumFmt = axisNumFmt(fmt, refs[1]);
  applyLabels(spec, fmt, refs);
  // categories: a continuous pill draws an axis, discrete pills draw headers – hidden only if all are
  const dims = [...roles[valueShelf].dims.map(d => ({ d, shelf: valueShelf })), ...roles[catShelf].dims.map(d => ({ d, shelf: catShelf }))];
  const hidden = ({ d, shelf }) => !!d.ref && (d.continuous ? fmt.axisInfo(d.ref, shelf, "0").hidden === true : fmt.isLabelHidden(d.ref));
  if (dims.length && dims.every(hidden)) spec.categoryAxisHidden = true;
  if (!fmt.axisLineShown(catShelf)) spec.axisLine = false;
}
