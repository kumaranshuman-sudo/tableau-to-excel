/* ══════════════════════════════════════════════════════════════════════════
 * PIES AND DONUTS – Tableau marks → pie model → chart jobs. One summary row = one Tableau mark = one slice.
 *   Angle (wedge-size)  → slice value; no Angle → equal slices (as Tableau draws them)
 *   Color               → slice fill: palette / custom colours, or gradient for a measure
 *   Detail / other dims → more slices, never merged
 *   Label / Tooltip     → Tableau's template filled with Tableau's formatted values
 *   Rows / Cols dims    → one pie per pane
 * A donut is a dual axis (MIN(0) twice): the slices plus a hole layer on top (its own label – usually the
 * total – drawn in the centre); a nested donut is two coloured pies, the top one smaller.
 * ══════════════════════════════════════════════════════════════════════════ */
import { FORMAT_CONFIG, TABLEAU_DEFAULTS } from "../../config.js";
import { tfDvNum, tfDvText, tfIsNull, tfNaturalCompare } from "../../data/values.js";
import { tfBuildColorScale } from "../../format/color-scale.js";
import { tfExcelFont } from "../../format/excel-style.js";
import { tfFormatNumber } from "../../format/number-format.js";
import { tfZoneText } from "../../twb/dashboard-text.js";
import { tfParseFieldRef, tfSameField } from "../../twb/field-ref.js";
import { createSheetFormatter, tfMerge } from "../../twb/formatter.js";
import { tfLog, tfNorm } from "../../util.js";
import { pieChartXml, userShapesXml } from "../writer/pie.js";

/* Measured in Tableau 2026.2 (1000×800 dashboard): a view with nothing on Rows/Columns draws each pane as a
 * 160×160 px cell at the zone's top-left (inside the 4 px zone margin), pie centred in it. Size slider →
 * diameter is linear in the TWB value: untouched = 80 px, 1.4613 → 117 px (80 × 1.46); a pie bigger than its
 * cell is clipped to it. */
const PIE_DIAMETER_PX = 80;
const PIE_CELL_PX = 160;
const ZONE_MARGIN = 4;
/** @param {number | null | undefined} sizeValue Size slider */
export const pieDiameterPx = sizeValue => sizeValue === null || sizeValue === undefined ? PIE_DIAMETER_PX : PIE_DIAMETER_PX * Math.max(0.01, sizeValue);

/** Excel chart font from Tableau props; measureName = Tableau's own font, for text measuring */
function pieFont(p) {
  const f = tfExcelFont(p || {});
  return { name: f.name, size: f.size, bold: f.bold, italic: f.italic, underline: !!f.underline,
           color: f.color ? f.color.argb.slice(2) : "333333", measureName: p && p.fontName };
}

/* ── label placement ────────────────────────────────────────────────────────
 * Tableau's pie label placement (measured in Tableau 2026.2): a label sits at its slice's outer mid-point, on
 * the side the slice faces – right half: label starts there, left half: ends there; top half: above, bottom
 * half: below – and is pushed back inside the pane where it would stick out. With "Allow labels to overlap
 * other marks" off, a label hitting an already placed one is hidden, biggest slice first. Tableau draws label
 * text at the screen's DPI scale (9 pt = 15 px at 125 %), the pane itself in plain px. */
const LABEL_CHAR_EM = 0.52;            // text width per character when there is no canvas to measure with
const LABEL_LINE_EM = 1.33;            // Tableau label line height (20 px lines for 15 px text)
const LABEL_PAD = 2;
/** @type {CanvasRenderingContext2D | null} */
let measureCtx = null;
function textWidthPx(text, font, px) {
  try {
    measureCtx = measureCtx || document.createElement("canvas").getContext("2d");
    measureCtx.font = `${font.italic ? "italic " : ""}${font.bold ? "bold " : ""}${px}px "${font.measureName || font.name}", Arial, sans-serif`;
    return measureCtx.measureText(text).width;
  } catch (e) {
    return text.length * px * LABEL_CHAR_EM;
  }
}

/**
 * Per point a label box {x, y, w, h, up} in pane px, or null (no label / hidden).
 * @param {any[]} points @param {{ cx: number, cy: number, R: number, w: number, h: number, cull: boolean, textScale: number }} g
 */
export function pieLabelLayout(points, g) {
  const EPS = 1e-6;
  const total = points.reduce((s, p) => s + p.value, 0) || 1;
  let acc = 0;
  const boxes = points.map(p => {
    const mid = (acc + p.value / 2) / total * 2 * Math.PI;      // clockwise from 12 o'clock (firstSliceAng 0)
    acc += p.value;
    if (!p.label || !p.label.length) return null;
    const s = Math.sin(mid), c = Math.cos(mid);
    const right = s > EPS || (Math.abs(s) <= EPS && c > 0);    // exactly 12 / 6 o'clock: the way the circle turns
    const up = c > EPS || (Math.abs(c) <= EPS && s < 0);       // exactly 3 o'clock → below, 9 o'clock → above
    let w = 0, h = 0;
    p.label.forEach(line => {
      const px = f => (f.size || 9) * 4 / 3 * g.textScale;
      w = Math.max(w, line.reduce((t, x) => t + textWidthPx(x.text, x.font, px(x.font)), 0));
      h += (line.length ? Math.max(...line.map(x => px(x.font))) : px({})) * LABEL_LINE_EM;
    });
    w += 2 * LABEL_PAD;
    const ax = g.cx + g.R * s, ay = g.cy - g.R * c;
    return { x: Math.max(0, Math.min(g.w - w, right ? ax : ax - w)),
             y: Math.max(0, Math.min(g.h - h, up ? ay - h : ay)), w, h, up };
  });
  if (g.cull) {
    // collision on the text itself (box minus padding), a few px of slack: Tableau keeps labels that only touch
    const SLACK = 4;
    const hit = (a, b) => Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x) - 2 * LABEL_PAD > SLACK &&
                          Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) > SLACK;
    const kept = [];
    points.map((p, i) => i).sort((a, b) => points[b].value - points[a].value || a - b).forEach(i => {
      if (!boxes[i]) return;
      if (kept.some(k => hit(k, boxes[i]))) boxes[i] = null;
      else kept.push(boxes[i]);
    });
  }
  return boxes;
}

/* ── text ──────────────────────────────────────────────────────────────── */
/** a number in the sheet's own Excel format; null when there is no format or no number */
const formatNum = (v, nf) => v === null || v === undefined || !nf ? null : tfFormatNumber(v, nf);

/** a number shown the way Tableau showed a sibling value: same prefix / suffix, decimals, grouping, K/M/% scale */
function formatLike(dv, v) {
  const s = tfDvText(dv), n = tfDvNum(dv);
  const m = s.match(/^(\D*?)(-?\d[\d,]*(?:\.\d+)?)(.*)$/);
  if (!m || !n) return String(v);
  const shown = parseFloat(m[2].replace(/,/g, ""));
  const scale = shown ? n / shown : 1;
  const dec = (m[2].split(".")[1] || "").length;
  return m[1] + (v / scale).toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec,
    useGrouping: m[2].includes(",") || Math.abs(shown) < 1000 }) + m[3];
}

/**
 * runs [{text, props}] with <[field]> placeholders → lines [[{text, props}]]. Tableau usually splits a
 * placeholder over runs ("<", "[ds].[field]", ">"), so the joined text is rendered and every piece keeps the
 * props of the run it came from (a field value takes its field run's props).
 * @param {{ text: string, props?: any }[]} runs @param {(ref: FieldRef) => string} valueOf
 */
export function renderTemplate(runs, valueOf) {
  const texts = runs.map(r => String(r.text || "").replace(/Æ\r?\n?/g, "\n"));
  const text = texts.join("");
  const owner = [];                                            // char index → run index
  texts.forEach((t, i) => { for (let k = 0; k < t.length; k++) owner.push(i); });
  const pieces = [];
  const push = (t, ri) => {
    const last = pieces[pieces.length - 1];
    if (last && last.ri === ri) last.text += t; else pieces.push({ text: t, ri });
  };
  const re = /<([^<>]+)>/g;
  let at = 0, m;
  while ((m = re.exec(text))) {
    for (let k = at; k < m.index; k++) push(text[k], owner[k]);
    const inner = m[1];
    let val = m[0];
    if (/^\[/.test(inner)) { const ref = tfParseFieldRef(inner); val = ref ? valueOf(ref) : ""; }
    else if (/^(sheet name|workbook name|page name|page count|page number)$/i.test(inner)) val = "";
    push(val, owner[m.index + 1]);
    at = m.index + m[0].length;
  }
  for (let k = at; k < text.length; k++) push(text[k], owner[k]);
  const lines = [[]];
  pieces.forEach(p => p.text.split(/\r?\n/).forEach((part, i) => {
    if (i) lines.push([]);
    if (part) lines[lines.length - 1].push({ text: part, props: runs[p.ri].props || {} });
  }));
  while (lines.length && !lines[lines.length - 1].length) lines.pop();
  return lines;
}
const plainText = lines => lines.map(line => line.map(x => x.text).join("")).join("\n");

/** summary rows selected in Tableau (getSelectedMarksAsync), matched on their dimension values */
function selectedRows(cols, data, tables) {
  const out = new Set();
  (tables || []).forEach(t => {
    const pairs = (t.columns || []).map((c, ti) => [cols.findIndex(x => x.name === c.fieldName), ti])
      .filter(([si]) => si >= 0 && !cols[si].numeric);
    if (!pairs.length) return;
    const keys = new Set((t.data || []).map(r => pairs.map(([, ti]) => tfDvText(r[ti])).join("\u0001")));
    data.forEach((r, i) => { if (keys.has(pairs.map(([si]) => tfDvText(r[si])).join("\u0001"))) out.add(i); });
  });
  return out;
}

/* ── model ─────────────────────────────────────────────────────────────── */
/**
 * One pie layer: its slices (value, colour, label, tooltip) per pane, its legend.
 * @param {SheetFormatter} fmt formatter scoped to the layer's pane @param {any} summary
 * @param {{ selectedTables?: any[] }} [opts]
 */
export function buildPieSpec(fmt, summary, opts = {}) {
  // fieldId ("[ds].[pcto:sum:Qty:qk]") carries the aggregation – a % of total and its base measure can share
  // the fieldName "SUM(Qty)", so it wins over name matching when present
  const cols = (summary.columns || []).map((c, i) => {
    const name = c.fieldName || c.fieldId || `Col${i + 1}`;
    const byId = /:(qk|nk|ok)(:\d+)?\]$/.test(c.fieldId || "") ? tfParseFieldRef(c.fieldId) : null;
    return { name, ref: byId || fmt.matchName(name), numeric: /^(int|float|real|integer|number)$/i.test(String(c.dataType || "")) };
  });
  const colFor = ref => {
    if (!ref) return -1;
    let i = cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
    if (i < 0 && /^multiple values$/i.test(ref.name)) i = cols.findIndex(c => /^measure values$/i.test(c.name));
    if (i < 0 && /^measure names$/i.test(ref.name)) i = cols.findIndex(c => /^measure names$/i.test(c.name));
    return i;
  };

  const angleRef = fmt.encodingRefs(/^wedge-size$/i)[0] || null;
  const vi = colFor(angleRef);
  let marks = (summary.data || []).map((r, i) => ({ r, i }));
  // dual axis: summary data holds every layer's marks (a donut's hole is one extra row with the total);
  // this layer's rows are the ones where its own axis measure has a value
  const axisCol = fmt.axisRefs().length === 1 ? colFor(fmt.axisRefs()[0]) : -1;
  if (axisCol >= 0) marks = marks.filter(({ r }) => !tfIsNull(r[axisCol]));
  if (vi >= 0) marks = marks.filter(({ r }) => { const v = tfDvNum(r[vi]); return v !== null && v > 0; });   // Tableau drops null / ≤ 0 wedges
  if (!marks.length) return null;

  // a field's value as Tableau shows it; % of total computed when the table calc isn't in the summary data
  const totals = {};
  const valueOf = r => ref => {
    const ci = colFor(ref);
    if (ci >= 0) return formatNum(tfDvNum(r[ci]), fmt.numFmtFor(ref)) || tfDvText(r[ci]);
    if (!/^pcto/i.test(ref.deriv || "")) return "";
    const base = tfNorm(ref.name.split(":").pop());
    const bi = cols.findIndex(c => c.numeric && c.ref && tfNorm(c.ref.name) === base);
    if (bi < 0) return "";
    if (totals[bi] === undefined) totals[bi] = marks.reduce((s, m) => s + (tfDvNum(m.r[bi]) || 0), 0);
    // % of the whole table (Tableau's default scope); other "compute using" scopes are not replicated
    const nf = fmt.numFmtFor(ref);
    const dec = nf ? ((nf.match(/\.(0+)/) || [, ""])[1].length) : 2;
    return totals[bi] ? (100 * (tfDvNum(r[bi]) || 0) / totals[bi]).toFixed(dec) + "%" : "";
  };

  // colour + legend
  const enc = fmt.colorEncoding();
  const ci = enc ? colFor(enc.ref) : -1;
  let colorAt = r => fmt.markColor();                          // no colour field: Marks → Color, else Tableau blue
  let legend = null;
  if (ci >= 0) {
    const pick = dv => tfIsNull(dv) ? null
      : enc.continuous ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : tfDvText(dv);
    const vals = marks.map(({ r }) => pick(r[ci]));
    const scale = tfBuildColorScale(fmt, enc, vals);
    if (scale) {
      const rgb = v => { const a = scale(v); return a ? a.slice(2) : null; };
      colorAt = r => rgb(pick(r[ci]));
      const title = fmt.captionFor(enc.ref, cols[ci].name);
      if (enc.continuous) {
        const nums = vals.filter(v => typeof v === "number");
        const lo = Math.min(...nums), hi = Math.max(...nums);
        legend = { title, items: [{ text: String(lo), color: rgb(lo) }, { text: String(hi), color: rgb(hi) }] };   // gradient: its two ends
      } else {
        const uniq = [...new Map(marks.map(({ r }) => [tfDvText(r[ci]), r[ci]])).values()].sort(tfNaturalCompare);
        legend = { title, items: uniq.map(dv => ({ text: tfDvText(dv), color: rgb(tfDvText(dv)) })) };
      }
    }
  }

  // panes: discrete fields on Rows / Columns → one pie each
  const hdr = shelf => cols.map((c, i) => c.ref && fmt.isHeaderField(c.ref) && fmt.shelfOf(c.ref) === shelf ? i : -1).filter(i => i >= 0);
  const rowHdr = hdr("rows"), colHdr = hdr("cols");
  const keyOf = (r, idx) => idx.map(i => tfDvText(r[i])).join(" | ");
  const hdrSet = new Set([...rowHdr, ...colHdr]);
  const dimIdx = cols.map((c, i) => !c.numeric && !hdrSet.has(i) && !/^measure values$/i.test(c.name) ? i : -1).filter(i => i >= 0);

  const valueCaption = vi >= 0 ? fmt.captionFor(angleRef, cols[vi].name) : "Marks";
  const showLabels = fmt.markLabelsShown();
  const labelTpl = fmt.labelRunsTemplate(), tipTpl = fmt.tooltipTemplate();
  const selected = selectedRows(cols, summary.data || [], opts.selectedTables);
  const points = marks.map(({ r, i }) => ({
    row: i,
    rowKey: keyOf(r, rowHdr), colKey: keyOf(r, colHdr),
    category: dimIdx.map(k => tfDvText(r[k])).filter(Boolean).join(", ") || valueCaption,
    value: vi >= 0 ? tfDvNum(r[vi]) : 1,
    color: colorAt(r),
    selected: selected.has(i),
    label: showLabels ? renderTemplate(labelTpl, valueOf(r))
      .map(line => line.map(x => ({ text: x.text, font: pieFont(fmt.labelFont(x.props)) }))) : null,
    tooltip: plainText(renderTemplate(tipTpl, valueOf(r)))
  }));
  const markSize = fmt.markSize();
  tfLog(`Pie: ${points.length} marks, angle=${vi >= 0 ? cols[vi].name : "none"}, colour=${ci >= 0 ? cols[ci].name : "none"}, ` +
        `labels=${showLabels}, selected=${selected.size}, size slider=${markSize === null ? "default" : markSize}`);
  const axis = shelf => ((fmt.sheetModel || {})[shelf] || []).some(r => r.type === "qk");
  return {
    columns: cols, data: summary.data || [], valueCaption, points, legend, markSize,
    cullLabels: fmt.markLabelsCulled(),
    rowKeys: [...new Set(points.map(p => p.rowKey))], colKeys: [...new Set(points.map(p => p.colKey))],
    stretch: { w: axis("cols"), h: axis("rows") },             // a continuous axis (a donut's MIN(0)) → the pane fills the zone
    paneOf: r => keyOf(r, rowHdr) + "\u0001" + keyOf(r, colHdr),   // the pane of any summary row (e.g. a hole row)
    colorRef: enc ? enc.ref : null,
    font: pieFont(fmt.labelFont({})),
    /** @type {any} */ hole: null, /** @type {any} */ inner: null, /** @type {any} */ centerText: null
  };
}

/**
 * Donut hole = the top layer of a dual axis (nothing on its card splitting it), drawn as Excel's doughnut hole.
 * Its label lives at the hole's level of detail (usually the grand total): Tableau returns one summary row per
 * pane for it, which carries the exact value – COUNTD and AVG too; without it the measures are re-aggregated
 * over the slices. @param {SheetFormatter} fmt formatter scoped to the hole's pane @param {any} pie
 */
export function buildHoleSpec(fmt, pie) {
  const shown = fmt.markLabelsShown(), tpl = fmt.labelRunsTemplate();
  const ax = fmt.axisRefs()[0];
  const ai = ax ? pie.columns.findIndex(c => c.ref && tfSameField(c.ref, ax)) : -1;
  const own = new Map(ai >= 0 ? pie.data.filter(r => !tfIsNull(r[ai])).map(r => [pie.paneOf(r), r]) : []);
  const valueOf = pts => ref => {
    const k = pie.columns.findIndex(c => c.ref && tfSameField(c.ref, ref));
    if (k < 0) return "";
    const hit = own.get(pts[0].rowKey + "\u0001" + pts[0].colKey);
    if (hit) return formatNum(tfDvNum(hit[k]), fmt.numFmtFor(ref)) || tfDvText(hit[k]);
    const dvs = pts.map(p => pie.data[p.row][k]);
    if (!pie.columns[k].numeric) { const u = [...new Set(dvs.map(tfDvText))]; return u.length === 1 ? u[0] : "*"; }
    const nums = dvs.map(tfDvNum).filter(v => v !== null);
    if (!nums.length) return "";
    const d = (ref.deriv || "").toLowerCase();
    // AVG / CNTD re-aggregated as a plain mean / sum of the slices, not from the raw rows
    const sum = nums.reduce((a, b) => a + b, 0);
    const v = d === "min" ? Math.min(...nums) : d === "max" ? Math.max(...nums) : d === "avg" ? sum / nums.length : sum;
    return formatNum(v, fmt.numFmtFor(ref)) || formatLike(dvs.find(x => tfDvNum(x) !== null), v);
  };
  // hover text: the Tooltip editor if set; else, like Tableau, "Caption: value" for every field the hole's own
  // row has a value for, dimensions first, then by field name
  const tipFor = pts => {
    const hit = pts.length && own.get(pts[0].rowKey + "\u0001" + pts[0].colKey);
    if (!hit || fmt.hasCustomTooltip()) return plainText(renderTemplate(fmt.tooltipTemplate(), valueOf(pts)));
    return pie.columns.map((c, k) => ({ c, k })).filter(({ c, k }) => c.ref && !tfIsNull(hit[k]))
      .sort((a, b) => Number(a.c.numeric) - Number(b.c.numeric) || (a.c.numeric ? a.c.ref.inner.localeCompare(b.c.ref.inner) : 0))
      .map(({ c }) => fmt.captionFor(c.ref, c.name) + ": " + valueOf(pts)(c.ref)).join("\n");
  };
  return {
    markSize: fmt.markSize(),
    tipFor,
    color: fmt.markColor(),                                    // usually white = the background
    labelFor: pts => shown ? renderTemplate(tpl, valueOf(pts))
      .map(line => line.map(x => ({ text: x.text, font: pieFont(fmt.labelFont(x.props)) }))) : null
  };
}

/**
 * Floating centre text: a dashboard text box, or another sheet (e.g. a total), whose centre lies inside a single
 * donut's hole (or a pie's circle) is drawn there with its Tableau text and fonts.
 * g: { stretch, D, holeD, cell, margin } px; dataOf(sheetName) → that sheet's summary data.
 * @returns {{ name: string | null, id: string, lines: any[] } | null} name: the sheet drawn there (null = a text box)
 */
export function pieCenterText(model, dashName, sheetName, g, dataOf) {
  const dash = model && model.dashboards && model.dashboards[dashName];
  const z = dash && dash.width && dash.height && dash.zones.find(s => s.type === "worksheet" && s.name === sheetName);
  if (!z) return null;
  const px = (v, w) => v / 100000 * (w ? dash.width : dash.height);
  const W = px(z.w, 1), H = px(z.h, 0);
  const cx = px(z.x, 1) + g.margin + (g.stretch.w ? W - 2 * g.margin : Math.min(g.cell, W - 2 * g.margin)) / 2;
  const cy = px(z.y, 0) + (g.stretch.h ? H - g.margin : Math.min(g.cell, H - g.margin)) / 2;
  const R = (g.holeD || g.D) / 2;
  const hit = dash.zones.find(c => c !== z && !c.hidden && c.w && c.h && px(c.w, 1) < W && px(c.h, 0) < H &&
    (c.type === "text" ? !!(c.runs && tfZoneText(c)) : c.type === "worksheet" && c.name !== sheetName) &&
    Math.hypot(px(c.x + c.w / 2, 1) - cx, px(c.y + c.h / 2, 0) - cy) <= R);
  if (!hit) return null;
  if (hit.type === "text") return { name: null, id: hit.id, lines: renderTemplate(hit.runs, () => "")
    .map(line => line.map(x => ({ text: x.text, font: pieFont(tfMerge(TABLEAU_DEFAULTS.worksheet, x.props)) }))) };
  const sum = dataOf(hit.name);
  if (!sum || !sum.data || !sum.data.length) return null;
  const f = createSheetFormatter(model, hit.name);
  const refs = (sum.columns || []).map(c => /:(qk|nk|ok)(:\d+)?\]$/.test(c.fieldId || "") ? tfParseFieldRef(c.fieldId) : f.matchName(c.fieldName));
  const row = sum.data[0];
  const valueOf = ref => {
    const k = refs.findIndex(r => r && tfSameField(r, ref));
    return k < 0 ? "" : formatNum(tfDvNum(row[k]), f.numFmtFor(ref)) || tfDvText(row[k]);
  };
  const tpl = f.labelRunsTemplate();
  let lines = tpl.length ? renderTemplate(tpl, valueOf) : [];
  if (!lines.some(l => l.length)) lines = [[{ text: row.map(tfDvText).filter(Boolean).join(" "), props: {} }]];   // no label → its values
  return { name: hit.name, id: hit.id, lines: lines.map(line => line.map(x => ({ text: x.text, font: pieFont(f.labelFont(x.props)) }))) };
}

/**
 * The whole pie model of a sheet, or null when it isn't a pie / donut Tableau draws with one or two layers.
 * @param {FormatModel} model @param {string} sheetName @param {any} summary
 * @param {{ dashboardName: string, selectedTables?: any[], dataOf: (name: string) => any }} opts
 */
export function buildPieModel(model, sheetName, summary, opts) {
  const fmt = createSheetFormatter(model, sheetName);
  const layers = fmt.pieLayers();
  if (!layers) return null;
  const scoped = pane => createSheetFormatter(model, sheetName, [pane]);
  const pie = buildPieSpec(scoped(layers.outer), summary, opts);
  if (!pie) return null;
  if (layers.hole) pie.hole = buildHoleSpec(scoped(layers.hole), pie);
  if (layers.inner) pie.inner = buildPieSpec(scoped(layers.inner), summary, opts);
  const fit = fmt.fitMode(opts.dashboardName);                 // Entire View / Fit Width / Fit Height → the pane fills the zone
  if (/^(entire-view|fit-width)$/.test(fit || "")) pie.stretch.w = true;
  if (/^(entire-view|fit-height)$/.test(fit || "")) pie.stretch.h = true;
  if (pie.rowKeys.length * pie.colKeys.length === 1) {
    pie.centerText = pieCenterText(model, opts.dashboardName, sheetName,
      { stretch: pie.stretch, D: pieDiameterPx(pie.markSize), holeD: pie.hole ? pieDiameterPx(pie.hole.markSize) : 0,
        cell: PIE_CELL_PX, margin: ZONE_MARGIN }, opts.dataOf);
  }
  return pie;
}

/* ── output ───────────────────────────────────────────────────────────── */
/**
 * The slices on the chart data sheet – every summary column, the category / value Excel plots and Tableau's
 * label and tooltip text – so each chart is linked to real cells. Returns per pane the ranges its series uses.
 * @param {import("exceljs").Worksheet} ds @param {number} startRow 1-based @param {any} pie @param {string} name
 * @returns {{ panes: Map<string, { ref: any, innerRef: any }>, nextRow: number }}
 */
export function writePieData(ds, startRow, pie, name) {
  const nCols = pie.columns.length;
  ds.getCell(startRow, 1).value = name;
  ds.getCell(startRow, 1).font = { bold: true };
  // cell by cell: Row.values only takes an Array of ExcelJS's own realm
  const putRow = (r, values) => values.forEach((v, k) => { ds.getCell(r, k + 1).value = v; });
  putRow(startRow + 1, [...pie.columns.map(c => c.name), "Excel category", "Excel value", "Tableau label", "Tableau tooltip"]);
  ds.getRow(startRow + 1).font = { bold: true };
  const catCol = ds.getColumn(nCols + 1).letter, valCol = ds.getColumn(nCols + 2).letter;
  let row = startRow + 2;
  const writeRows = pts => {
    const first = row;
    pts.forEach(p => {
      putRow(row++, [
        ...pie.data[p.row].map((dv, k) => pie.columns[k].numeric ? tfDvNum(dv) : tfDvText(dv)),
        p.category, p.value, p.label ? plainText(p.label) : "", p.tooltip]);
    });
    return { sheet: ds.name, cat: `$${catCol}$${first}:$${catCol}$${row - 1}`, val: `$${valCol}$${first}:$${valCol}$${row - 1}` };
  };
  const panes = new Map();
  pie.rowKeys.forEach(rk => pie.colKeys.forEach(ck => {
    const pts = pie.points.filter(p => p.rowKey === rk && p.colKey === ck);
    if (!pts.length) return;
    const ref = writeRows(pts);
    const innerPts = pie.inner ? pie.inner.points.filter(p => p.rowKey === rk && p.colKey === ck) : [];
    panes.set(rk + "\u0001" + ck, { ref, innerRef: innerPts.length ? writeRows(innerPts) : null });
  }));
  return { panes, nextRow: row + 1 };
}

/**
 * Chart jobs for a pie sheet drawn in a block of w × h px: one exact-size chart per pane (a nested donut's inner
 * pie is a second, see-through chart over the ring's hole), Tableau's labels as text boxes, its tooltip on an
 * invisible wedge over each slice. Positions are px from the block's top-left.
 * @param {any} pie @param {{ w: number, h: number }} block @param {Map<string, any>} paneRefs (writePieData)
 * @param {{ name: string, bg: string, link: string }} o bg RRGGBB; link: where clicking a slice goes
 * @returns {{ x: number, y: number, w: number, h: number, xml: string, shapes: string | null, tips: any[], name: string }[]}
 */
export function buildPieCharts(pie, block, paneRefs, o) {
  const M = ZONE_MARGIN;
  // Tableau's cell per pane, from the block's top-left inside the zone margin
  const fitW = (block.w - 2 * M) / pie.colKeys.length, fitH = (block.h - M) / pie.rowKeys.length;
  const pw = pie.stretch.w ? fitW : Math.min(PIE_CELL_PX, fitW);
  const ph = pie.stretch.h ? fitH : Math.min(PIE_CELL_PX, fitH);
  const textScale = FORMAT_CONFIG.tableauTextScale || (typeof window !== "undefined" && window.devicePixelRatio) || 1;
  // Tableau clips an oversized pie to its cell; Excel can't clip, so it fills the cell instead – and the hole /
  // inner ring shrink with it, so the ring keeps Tableau's proportions
  const full = pieDiameterPx(pie.markSize);
  const D = Math.min(full, pw, ph), k = D / full;
  const holeD = pie.hole ? Math.min(pieDiameterPx(pie.hole.markSize) * k, D) : 0;
  const innerD = pie.inner ? Math.min(pieDiameterPx(pie.inner.markSize) * k, D) : 0;
  const single = pie.rowKeys.length * pie.colKeys.length === 1;
  const labelsOf = (pts, R, cull) => {
    const boxes = pieLabelLayout(pts, { cx: pw / 2, cy: ph / 2, R, w: pw, h: ph, cull, textScale });
    return pts.map((p, k) => boxes[k] && { ...boxes[k], lines: p.label, anchor: boxes[k].up ? "b" : "t" }).filter(Boolean);
  };
  // Excel draws slice 1 from 12 o'clock clockwise with |value|; DrawingML angles are 60000ths of a degree
  // clockwise from 3 o'clock
  const tipsOf = (pts, d, hole, size) => {
    const x = (size.w - d) / 2, y = (size.h - d) / 2;
    const total = pts.reduce((s, p) => s + Math.abs(p.value || 0), 0);
    const ang = a => (((a % 360) + 360) % 360) * 60000;
    let a = -90;
    return pts.map(p => {
      const st = a;
      a += total ? 360 * Math.abs(p.value || 0) / total : 0;
      if (a === st || !p.tooltip) return null;
      return { x, y, w: d, h: d, prst: hole ? "blockArc" : "pie",
               adj: hole ? [ang(st), ang(a), 50000 * (1 - hole / d)] : [ang(st), ang(a)], text: p.tooltip, link: o.link };
    }).filter(Boolean);
  };
  const out = [];
  pie.rowKeys.forEach((rk, ri) => pie.colKeys.forEach((ck, cj) => {
    const pts = pie.points.filter(p => p.rowKey === rk && p.colKey === ck);
    const refs = paneRefs.get(rk + "\u0001" + ck);
    if (!pts.length || !refs) return;
    const innerPts = innerD && refs.innerRef ? pie.inner.points.filter(p => p.rowKey === rk && p.colKey === ck) : [];
    // labels of the two rings are culled per ring, not against each other
    const texts = [...labelsOf(pts, D / 2, pie.cullLabels), ...(innerPts.length ? labelsOf(innerPts, innerD / 2, pie.inner.cullLabels) : [])];
    // Excel's hole is see-through (it shows the chart background): a hole in another colour is a filled circle
    // under the labels
    if (holeD && pie.hole.color.toUpperCase() !== o.bg.toUpperCase())
      texts.unshift({ x: (pw - holeD) / 2, y: (ph - holeD) / 2, w: holeD, h: holeD, shape: "ellipse", fill: pie.hole.color, lines: [] });
    // centre: a floating text box / sheet over the hole wins over the hole's own label (Tableau draws it on top)
    const centre = (single && pie.centerText && pie.centerText.lines) || (pie.hole && pie.hole.labelFor(pts));
    if (centre && centre.length && !innerPts.length) texts.push({ x: 0, y: 0, w: pw, h: ph, lines: centre, anchor: "ctr" });
    const x = M + cj * pw, y = ri * ph;
    const spec = {
      seriesName: pie.valueCaption, categories: pts.map(p => p.category), values: pts.map(p => p.value), ref: refs.ref,
      title: [rk, ck].filter(Boolean).join(" | ") || null, titleFont: pie.font, font: pie.font, bg: o.bg,
      plot: { x: (1 - D / pw) / 2, y: (1 - D / ph) / 2, w: D / pw, h: D / ph },
      holeSize: holeD ? 100 * holeD / D : innerPts.length ? 100 * innerD / D : null,
      points: pts.map(p => ({ color: p.color, selected: p.selected })), hasShapes: texts.length > 0
    };
    const holeTip = holeD && pie.hole.tipFor(pts);
    out.push({ x, y, w: pw, h: ph, name: o.name, xml: pieChartXml(spec),
               shapes: texts.length ? userShapesXml(texts, { w: pw, h: ph }) : null,
               tips: [...tipsOf(pts, D, holeD || (innerPts.length ? innerD : 0), { w: pw, h: ph }),
                      ...(holeTip ? [{ x: (pw - holeD) / 2, y: (ph - holeD) / 2, w: holeD, h: holeD, prst: "ellipse", text: holeTip, link: o.link }] : [])] });
    if (!innerPts.length) return;
    // nested donut: the inner pie is its own see-through chart exactly over the ring's hole (drawn after = on top);
    // its labels stay in the ring chart so they can sit over the ring
    const innerTexts = centre && centre.length ? [{ x: 0, y: 0, w: innerD, h: innerD, lines: centre, anchor: "ctr" }] : [];
    const innerSpec = {
      seriesName: pie.inner.valueCaption, categories: innerPts.map(p => p.category), values: innerPts.map(p => p.value),
      ref: refs.innerRef, title: null, font: pie.font, bg: null, plot: { x: 0, y: 0, w: 1, h: 1 }, hasShapes: innerTexts.length > 0,
      points: innerPts.map(p => ({ color: p.color, selected: p.selected }))
    };
    out.push({ x: x + (pw - innerD) / 2, y: y + (ph - innerD) / 2, w: innerD, h: innerD, name: o.name + " (inner)",
               xml: pieChartXml(innerSpec), shapes: innerTexts.length ? userShapesXml(innerTexts, { w: innerD, h: innerD }) : null,
               tips: tipsOf(innerPts, innerD, 0, { w: innerD, h: innerD }) });
  }));
  return out;
}
