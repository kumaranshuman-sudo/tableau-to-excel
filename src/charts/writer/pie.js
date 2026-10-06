/* Pie / doughnut chart parts drawn the way Tableau draws them:
 *   pieChartXml     – the chart: slices with their colours (a selected mark outlined), a pinned plot area so
 *                     the diameter is Tableau's, an optional doughnut hole
 *   userShapesXml   – labels as text boxes at Tableau's label positions (Excel data labels wrap and shrink the
 *                     pie), the donut hole's own fill and the centre text
 *   tipShapeXml     – Tableau's tooltip: an invisible wedge over each slice whose hyperlink ScreenTip is the
 *                     tooltip text (Excel's own chart hover text cannot be changed)
 * Knows nothing about Tableau: colours, fonts and texts arrive resolved. */
import { NS } from "./constants.js";
import { esc } from "./xml-util.js";

const EMU_PER_PX = 9525;                                        // 96 dpi
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_CDR = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";
export const REL_USER_SHAPES = NS.r + "/chartUserShapes";
export const CT_USER_SHAPES = "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";
export const REL_HYPERLINK = NS.r + "/hyperlink";

/* ── text ─────────────────────────────────────────────────────────────── */
/** @typedef {{ name?: string, size?: number, bold?: boolean, italic?: boolean, underline?: boolean, color?: string }} PieFont */
/** @param {PieFont} [f] */
function fontAttrs(f) {
  f = f || {};
  return ` lang="en-US" sz="${Math.round((f.size || 9) * 100)}" b="${f.bold ? 1 : 0}" i="${f.italic ? 1 : 0}"` +
         (f.underline ? ' u="sng"' : "");
}
/** @param {PieFont} [f] */
function fontKids(f) {
  f = f || {};
  return `<a:solidFill><a:srgbClr val="${f.color || "333333"}"/></a:solidFill>` + (f.name ? `<a:latin typeface="${esc(f.name)}"/>` : "");
}
/** lines [[{text, font}]] → one paragraph per line, one run per piece @param {string} [algn] "l" | "ctr" | "r" */
function parasXml(lines, algn) {
  const pPr = algn ? `<a:pPr algn="${algn}"/>` : "";
  return lines.map(runs => {
    const rs = runs.filter(r => r.text !== "").map(r =>
      `<a:r><a:rPr${fontAttrs(r.font)}>${fontKids(r.font)}</a:rPr><a:t>${esc(r.text)}</a:t></a:r>`).join("");
    return rs ? `<a:p>${pPr}${rs}</a:p>` : `<a:p>${pPr}<a:endParaRPr lang="en-US"/></a:p>`;
  }).join("");
}
const richXml = lines => `<c:rich><a:bodyPr/><a:lstStyle/>${parasXml(lines)}</c:rich>`;
const txPrXml = f => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${fontAttrs(f)}>${fontKids(f)}</a:defRPr></a:pPr>` +
                     `<a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
const NO_LINE = `<a:ln><a:noFill/></a:ln>`;
const fillXml = c => c ? `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>` : `<a:noFill/>`;

/* ── data: worksheet ranges (linked, "Select Data" shows them) or literals ── */
const sheetRef = (sheet, range) => `'${String(sheet).replace(/'/g, "''")}'!${range}`;
const strPts = vals => `<c:ptCount val="${vals.length}"/>` + vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("");
const numPts = vals => `<c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
  vals.map((v, i) => `<c:pt idx="${i}"><c:v>${Number(v)}</c:v></c:pt>`).join("");
const catXml = (vals, ref) => ref
  ? `<c:cat><c:strRef><c:f>${esc(sheetRef(ref.sheet, ref.cat))}</c:f><c:strCache>${strPts(vals)}</c:strCache></c:strRef></c:cat>`
  : `<c:cat><c:strLit>${strPts(vals)}</c:strLit></c:cat>`;
const valXml = (vals, ref) => ref
  ? `<c:val><c:numRef><c:f>${esc(sheetRef(ref.sheet, ref.val))}</c:f><c:numCache>${numPts(vals)}</c:numCache></c:numRef></c:val>`
  : `<c:val><c:numLit>${numPts(vals)}</c:numLit></c:val>`;

/* ── chart frame ──────────────────────────────────────────────────────── */
function titleXml(title, font) {
  if (!title) return '<c:autoTitleDeleted val="1"/>';
  return `<c:title><c:tx>${richXml([[{ text: title, font }]])}</c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;
}
/** plot: { x, y, w, h } fractions of the chart (inner plot area) – pins the pie's diameter */
function layoutXml(plot) {
  if (!plot) return "<c:layout/>";
  const v = n => Math.max(0, Math.min(1, n)).toFixed(4);
  return `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
    `<c:x val="${v(plot.x)}"/><c:y val="${v(plot.y)}"/><c:w val="${v(plot.w)}"/><c:h val="${v(plot.h)}"/></c:manualLayout></c:layout>`;
}

/**
 * spec: { seriesName, categories[], values[], ref? {sheet, cat, val}, title?, titleFont?, font?, bg? (RRGGBB, null =
 *   see-through), plot?, holeSize? (% of the diameter → doughnut), hasShapes?, points: [{ color, selected }] }
 * @param {any} spec @returns {string}
 */
export function pieChartXml(spec) {
  const dPts = (spec.points || []).map((p, i) =>
    `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${fillXml(p.color)}` +
    (p.selected ? `<a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`
                : `<a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>`) +
    `</c:spPr></c:dPt>`).join("");
  const ser = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>${esc(spec.seriesName || "")}</c:v></c:tx>` +
    `${dPts}${catXml(spec.categories, spec.ref)}${valXml(spec.values, spec.ref)}</c:ser>`;
  const plot = spec.holeSize
    ? `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/>` +
      `<c:holeSize val="${Math.max(10, Math.min(90, Math.round(spec.holeSize)))}"/></c:doughnutChart>`    // Excel allows 10–90
    : `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`;
  return XML_DECL +
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:roundedCorners val="0"/><c:chart>${titleXml(spec.title, spec.titleFont)}` +
    `<c:plotArea>${layoutXml(spec.plot)}${plot}<c:spPr><a:noFill/>${NO_LINE}</c:spPr></c:plotArea>` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr>${fillXml(spec.bg)}${NO_LINE}</c:spPr>${txPrXml(spec.font)}` +
    (spec.hasShapes ? `<c:userShapes r:id="rId1"/>` : "") + `</c:chartSpace>`;
}

/**
 * Text boxes / shapes drawn on top of a chart (its chartDrawing part), anchored relative to the chart so they
 * move and scale with it; later items draw on top.
 * texts: [{ x, y, w, h (px inside the chart), lines: [[{text, font}]], align: "l"|"ctr"|"r", anchor: "t"|"ctr"|"b",
 *           shape?: preset geometry ("ellipse" …; default a borderless text box), fill?: RRGGBB }]
 * @param {any[]} texts @param {{ w: number, h: number }} size chart size px @returns {string}
 */
export function userShapesXml(texts, size) {
  const fx = (v, d) => Math.max(0, Math.min(1, v / d)).toFixed(5);
  const paras = t => t.lines && t.lines.length ? parasXml(t.lines, t.align || "ctr") : `<a:p><a:endParaRPr lang="en-US"/></a:p>`;
  return XML_DECL + `<c:userShapes xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:cdr="${NS_CDR}">` + texts.map((t, i) =>
    `<cdr:relSizeAnchor><cdr:from><cdr:x>${fx(t.x, size.w)}</cdr:x><cdr:y>${fx(t.y, size.h)}</cdr:y></cdr:from>` +
    `<cdr:to><cdr:x>${fx(t.x + t.w, size.w)}</cdr:x><cdr:y>${fx(t.y + t.h, size.h)}</cdr:y></cdr:to>` +
    `<cdr:sp macro="" textlink=""><cdr:nvSpPr><cdr:cNvPr id="${i + 2}" name="${t.shape ? "Shape" : "Label"} ${i + 1}"/>` +
    `<cdr:cNvSpPr${t.shape ? "" : ' txBox="1"'}/></cdr:nvSpPr>` +
    `<cdr:spPr><a:prstGeom prst="${t.shape || "rect"}"><a:avLst/></a:prstGeom>${fillXml(t.fill)}${NO_LINE}</cdr:spPr>` +
    `<cdr:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${t.anchor || "t"}"/><a:lstStyle/>` +
    `${paras(t)}</cdr:txBody></cdr:sp></cdr:relSizeAnchor>`).join("") + `</c:userShapes>`;
}

const xfrmXml = (x, y, w, h) => `<a:off x="${Math.round(x * EMU_PER_PX)}" y="${Math.round(y * EMU_PER_PX)}"/>` +
                                `<a:ext cx="${Math.round(w * EMU_PER_PX)}" cy="${Math.round(h * EMU_PER_PX)}"/>`;

/**
 * An invisible wedge over a slice: hovering shows its hyperlink ScreenTip (the Tableau tooltip). A 100 %
 * transparent fill still catches the mouse; noFill would not.
 * tip: { x, y, w, h (px inside the chart), prst ("pie" | "blockArc" | "ellipse"), adj: [avLst values], text }
 * @param {any} t @param {number} id @param {string} rid the hyperlink relationship
 */
export function tipShapeXml(t, id, rid) {
  const gd = (t.adj || []).map((v, k) => `<a:gd name="adj${k + 1}" fmla="val ${Math.round(v)}"/>`).join("");
  const text = String(t.text).slice(0, 255);                   // Excel's ScreenTip limit
  return `<xdr:sp macro="" textlink=""><xdr:nvSpPr>` +
    `<xdr:cNvPr id="${id}" name="Tooltip ${id}"><a:hlinkClick xmlns:r="${NS.r}" r:id="${rid}" tooltip="${esc(text).replace(/\n/g, "&#xA;")}"/></xdr:cNvPr>` +
    `<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm>${xfrmXml(t.x, t.y, t.w, t.h)}</a:xfrm><a:prstGeom prst="${t.prst}"><a:avLst>${gd}</a:avLst></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill>${NO_LINE}</xdr:spPr></xdr:sp>`;
}

/**
 * The anchor of a chart with tooltip wedges: chart and wedges grouped, so the wedges keep their exact px
 * positions over the slices whatever the cell sizes.
 * @param {any} chart ChartJob (col / row / offsets, widthPx / heightPx, name) @param {string} chartRid
 * @param {{ next: number }} ids shape ids @param {string[]} tipRids one hyperlink relationship per tip
 */
export function groupedAnchorXml(chart, chartRid, ids, tipRids) {
  const w = chart.widthPx, h = chart.heightPx;
  const frame = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${ids.next++}" name="${esc(chart.name || "Chart")}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm>${xfrmXml(0, 0, w, h)}</xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${chartRid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
  const group = `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="${ids.next++}" name="${esc((chart.name || "Chart") + " group")}"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>` +
    `<xdr:grpSpPr><a:xfrm>${xfrmXml(0, 0, w, h)}<a:chOff x="0" y="0"/><a:chExt cx="${Math.round(w * EMU_PER_PX)}" cy="${Math.round(h * EMU_PER_PX)}"/></a:xfrm></xdr:grpSpPr>` +
    frame + chart.tips.map((t, k) => tipShapeXml(t, ids.next++, tipRids[k])).join("") + `</xdr:grpSp>`;
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.col}</xdr:col><xdr:colOff>${Math.round((chart.colOffPx || 0) * EMU_PER_PX)}</xdr:colOff>` +
    `<xdr:row>${chart.row}</xdr:row><xdr:rowOff>${Math.round((chart.rowOffPx || 0) * EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${Math.round(w * EMU_PER_PX)}" cy="${Math.round(h * EMU_PER_PX)}"/>${group}<xdr:clientData/></xdr:oneCellAnchor>`;
}
