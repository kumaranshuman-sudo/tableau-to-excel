/* Excel 2016+ chartex parts (treemap) and their drawing anchors. */
import { EMU_PER_PX, NS } from "./constants.js";
import { line, solid } from "./drawingml.js";
import { esc, num } from "./xml-util.js";

/* Excel 2016+ "chartex" part – treemap. Category levels outer → inner in the spec,
 * inner (leaf) level first in the XML. Excel only accepts chartex formulas that are
 * hidden workbook names (_xlchart.v1.N), so refs here are those names, not ranges. */
/** @param {ChartSpec} spec @param {Pick<ChartRefs, "cat" | "series">} refs workbook names, not ranges @param {string} uid series GUID @returns {string} */
export function chartExXml(spec, refs, uid) {
  const levels = spec.categories.levels;
  const n = levels.length ? levels[0].length : 0;
  const s = spec.series[0];
  const lvl = lv => `<cx:lvl ptCount="${n}">${lv.map((v, i) => `<cx:pt idx="${i}">${esc(v)}</cx:pt>`).join("")}</cx:lvl>`;
  const vals = s.values.map((v, i) => num(v) === null ? "" : `<cx:pt idx="${i}">${v}</cx:pt>`).join("");
  const tile = c => `<cx:spPr>${solid(c)}${line("FFFFFF", 12700)}</cx:spPr>`;
  const font = spec.font || /** @type {ChartSpec["font"]} */ ({});
  const labelPr = `<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round((font.size || 9) * 100)}">` +
    `<a:latin typeface="${esc(font.name || "Arial")}"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></cx:txPr>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<cx:chartSpace xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:cx="${NS.cx}">` +
    `<cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>${esc(refs.cat)}</cx:f>${levels.slice().reverse().map(lvl).join("")}</cx:strDim>` +
    `<cx:numDim type="size"><cx:f>${esc(refs.series[0].val)}</cx:f><cx:lvl ptCount="${n}" formatCode="${esc(spec.numFmt || "General")}">${vals}</cx:lvl>` +
    `</cx:numDim></cx:data></cx:chartData>` +
    `<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="treemap" uniqueId="${uid}">` +
    `<cx:tx><cx:txData><cx:f>${esc(refs.series[0].tx)}</cx:f><cx:v>${esc(s.name)}</cx:v></cx:txData></cx:tx>` +
    tile(s.color) +
    (s.pointColors || []).map((c, i) => c ? `<cx:dataPt idx="${i}">${tile(c)}</cx:dataPt>` : "").join("") +
    (s.labels ? `<cx:dataLabels pos="inEnd">` +
      (s.labelParts && s.labelParts.value ? `<cx:numFmt formatCode="${esc(spec.numFmt || "General")}" sourceLinked="0"/>` : "") +
      `${labelPr}<cx:visibility seriesName="0" categoryName="1" value="${s.labelParts && s.labelParts.value ? 1 : 0}"/>` +
      (s.labelParts && s.labelParts.value ? `<cx:separator>${"\n"}</cx:separator>` : "") + `</cx:dataLabels>` : "") +
    `<cx:dataId val="0"/><cx:layoutPr><cx:parentLabelLayout val="${levels.length > 1 ? "banner" : "none"}"/></cx:layoutPr>` +
    `</cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>` +
    `<cx:spPr>${solid(spec.background || "FFFFFF")}<a:ln><a:noFill/></a:ln></cx:spPr></cx:chartSpace>`;
}

/* chartex frames sit in mc:AlternateContent; older Excel shows the fallback rectangle */
/** @param {ChartJob} chart @param {number} id @param {string} rid */
export function anchorExXml(chart, id, rid) {
  const cx = Math.max(1, Math.round(chart.widthPx * EMU_PER_PX));
  const cy = Math.max(1, Math.round(chart.heightPx * EMU_PER_PX));
  const name = esc(chart.name || "Chart " + id);
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.col}</xdr:col><xdr:colOff>${Math.round((chart.colOffPx || 0) * EMU_PER_PX)}</xdr:colOff>` +
    `<xdr:row>${chart.row}</xdr:row><xdr:rowOff>${Math.round((chart.rowOffPx || 0) * EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/>` +
    `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:cx1="${NS.cx1}" Requires="cx1">` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>` +
    `<a:graphicData uri="${NS.cx}"><cx:chart xmlns:cx="${NS.cx}" xmlns:r="${NS.r}" r:id="${rid}"/></a:graphicData>` +
    `</a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback>` +
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>` +
    `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/>` +
    `<a:t>This treemap needs Excel 2016 or later.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent>` +
    `<xdr:clientData/></xdr:oneCellAnchor>`;
}
