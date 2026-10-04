/* XLSX package surgery: drawing / chart parts, relationships, content types. */
import JSZip from "jszip";
import { CHARTEX_COLORS, CHARTEX_STYLE } from "./chartex-style.js";
import { anchorExXml, chartExXml } from "./chartex.js";
import { CT_CHART, CT_CHARTCOLORS, CT_CHARTEX, CT_CHARTSTYLE, CT_DRAWING, EMPTY_RELS, EMU_PER_PX, NS, REL_CHART, REL_CHARTCOLORS, REL_CHARTEX, REL_CHARTSTYLE, REL_DRAWING } from "./constants.js";
import { chartXml } from "./drawingml.js";
import { attr, esc } from "./xml-util.js";

/** @param {ChartJob} chart @param {number} id @param {string} rid */
export function anchorXml(chart, id, rid) {
  const cx = Math.max(1, Math.round(chart.widthPx * EMU_PER_PX));
  const cy = Math.max(1, Math.round(chart.heightPx * EMU_PER_PX));
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.col}</xdr:col><xdr:colOff>${Math.round((chart.colOffPx || 0) * EMU_PER_PX)}</xdr:colOff>` +
    `<xdr:row>${chart.row}</xdr:row><xdr:rowOff>${Math.round((chart.rowOffPx || 0) * EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>` +
    `<xdr:cNvPr id="${id}" name="${esc(chart.name || "Chart " + id)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>` +
    `<a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${rid}"/></a:graphicData>` +
    `</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
}

export function relTarget(relsXml, id) {
  const tag = (relsXml.match(new RegExp(`<Relationship\\b[^>]*\\sId="${id}"[^>]*>`)) || [])[0];
  return tag ? attr(tag, "Target") : null;
}

export function nextRelId(relsXml) {
  const ids = [...relsXml.matchAll(/\sId="rId(\d+)"/g)].map(m => +m[1]);
  return "rId" + ((ids.length ? Math.max(...ids) : 0) + 1);
}

export function addRel(relsXml, id, type, target) {
  return relsXml.replace("</Relationships>", `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`);
}

export function resolve(baseDir, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = (baseDir + target).split("/");
  const out = [];
  parts.forEach(p => { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); });
  return out.join("/");
}

export function dirOf(path) { return path.replace(/[^/]+$/, ""); }

export function relsPathOf(path) { return dirOf(path) + "_rels/" + path.replace(/^.*\//, "") + ".rels"; }

export function nextFreeIndex(zip, prefix, ext) {
  let n = 1;
  while (zip.file(`${prefix}${n}${ext}`)) n++;
  return n;
}

export function addOverride(ct, part, type) {
  if (ct.includes(`PartName="/${part}"`)) return ct;
  return ct.replace("</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
}

/* <drawing> must precede these elements in CT_Worksheet */
export function insertDrawingTag(sheetXml, rid) {
  const tag = `<drawing r:id="${rid}"/>`;
  const after = ["<legacyDrawing", "<legacyDrawingHF", "<drawingHF", "<picture", "<oleObjects", "<controls",
                 "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"];
  for (const t of after) {
    const i = sheetXml.indexOf(t);
    if (i >= 0) return sheetXml.slice(0, i) + tag + sheetXml.slice(i);
  }
  return sheetXml;
}

/* opts: { sheetIndex, charts: [{ spec, refs, col, row, widthPx, heightPx, name }] } */
/**
 * Adds the charts to the dashboard sheet of a written XLSX.
 * @param {ArrayBuffer | Uint8Array} buffer workbook.xlsx.writeBuffer() output
 * @param {{ sheetIndex: number, charts: ChartJob[] }} opts
 * @returns {Promise<ArrayBuffer | Uint8Array>}
 */
export async function injectCharts(buffer, opts) {
  if (typeof JSZip === "undefined") throw new Error("JSZip is required to add native Excel charts");
  const charts = (opts.charts || []).filter(Boolean);
  if (!charts.length) return buffer;
  const zip = await JSZip.loadAsync(buffer);
  const read = p => zip.file(p) ? zip.file(p).async("string") : Promise.resolve(null);

  const wbXml = await read("xl/workbook.xml");
  const wbRels = await read("xl/_rels/workbook.xml.rels");
  const sheetTags = wbXml.match(/<sheet\b[^>]*>/g) || [];
  const sheetTag = sheetTags[opts.sheetIndex || 0];
  if (!sheetTag) throw new Error("Dashboard sheet not found in workbook package");
  const sheetPath = resolve("xl/", relTarget(wbRels, attr(sheetTag, "r:id")));
  const sheetRelsPath = relsPathOf(sheetPath);
  let sheetXml = await read(sheetPath);
  let sheetRels = (await read(sheetRelsPath)) || EMPTY_RELS;
  let ct = await read("[Content_Types].xml");

  if (!/<worksheet\b[^>]*\sxmlns:r=/.test(sheetXml)) {
    sheetXml = sheetXml.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS.r}"`);
  }

  // re-use the drawing ExcelJS created for images, else create one
  let drawingPath, drawingXml, drawingRels;
  const existingRid = (sheetXml.match(/<drawing\b[^>]*r:id="([^"]+)"/) || [])[1];
  if (existingRid) {
    drawingPath = resolve(dirOf(sheetPath), relTarget(sheetRels, existingRid));
    drawingXml = await read(drawingPath);
    drawingRels = (await read(relsPathOf(drawingPath))) || EMPTY_RELS;
  } else {
    const n = nextFreeIndex(zip, "xl/drawings/drawing", ".xml");
    drawingPath = `xl/drawings/drawing${n}.xml`;
    drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"></xdr:wsDr>`;
    drawingRels = EMPTY_RELS;
    const rid = nextRelId(sheetRels);
    sheetRels = addRel(sheetRels, rid, REL_DRAWING, `../drawings/drawing${n}.xml`);
    sheetXml = insertDrawingTag(sheetXml, rid);
    ct = addOverride(ct, drawingPath, CT_DRAWING);
  }
  if (!/<xdr:wsDr\b[^>]*\sxmlns:a=/.test(drawingXml)) {
    drawingXml = drawingXml.replace(/<xdr:wsDr\b/, `<xdr:wsDr xmlns:a="${NS.a}"`);
  }

  const ids = [...drawingXml.matchAll(/<xdr:cNvPr\b[^>]*\sid="(\d+)"/g)].map(m => +m[1]);
  let shapeId = (ids.length ? Math.max(...ids) : 1) + 1;
  let anchors = "";
  // chartex data ranges → hidden defined names, as Excel writes them
  const usedNames = [...wbXml.matchAll(/name="_xlchart\.v1\.(\d+)"/g)].map(m => +m[1]);
  let nameIndex = usedNames.length ? Math.max(...usedNames) + 1 : 0;
  const newNames = [];
  const defineName = ref => {
    const name = `_xlchart.v1.${nameIndex++}`;
    newNames.push(`<definedName name="${name}" hidden="1">${esc(ref)}</definedName>`);
    return name;
  };
  for (const chart of charts) {
    const rid = nextRelId(drawingRels);
    if (chart.spec.kind === "treemap") {
      const n = nextFreeIndex(zip, "xl/charts/chartEx", ".xml");
      const chartPath = `xl/charts/chartEx${n}.xml`;
      const uid = `{6F1D3C1E-0000-4000-8000-${String(n).padStart(12, "0")}}`;
      const named = { cat: defineName(chart.refs.cat),
                      series: chart.refs.series.map(s => ({ tx: defineName(s.tx), val: defineName(s.val) })) };
      zip.file(chartPath, chartExXml(chart.spec, named, uid));
      ct = addOverride(ct, chartPath, CT_CHARTEX);
      const xmlHead = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
      const styleN = nextFreeIndex(zip, "xl/charts/style", ".xml");
      zip.file(`xl/charts/style${styleN}.xml`, xmlHead + CHARTEX_STYLE);
      zip.file(`xl/charts/colors${styleN}.xml`, xmlHead + CHARTEX_COLORS);
      ct = addOverride(ct, `xl/charts/style${styleN}.xml`, CT_CHARTSTYLE);
      ct = addOverride(ct, `xl/charts/colors${styleN}.xml`, CT_CHARTCOLORS);
      zip.file(relsPathOf(chartPath), addRel(addRel(EMPTY_RELS, "rId1", REL_CHARTSTYLE, `style${styleN}.xml`),
                                             "rId2", REL_CHARTCOLORS, `colors${styleN}.xml`));
      drawingRels = addRel(drawingRels, rid, REL_CHARTEX, `../charts/chartEx${n}.xml`);
      anchors += anchorExXml(chart, shapeId++, rid);
      continue;
    }
    const n = nextFreeIndex(zip, "xl/charts/chart", ".xml");
    const chartPath = `xl/charts/chart${n}.xml`;
    zip.file(chartPath, chartXml(chart.spec, chart.refs));
    ct = addOverride(ct, chartPath, CT_CHART);
    drawingRels = addRel(drawingRels, rid, REL_CHART, `../charts/chart${n}.xml`);
    anchors += anchorXml(chart, shapeId++, rid);
  }
  drawingXml = drawingXml.replace("</xdr:wsDr>", anchors + "</xdr:wsDr>");
  if (newNames.length) {                         // CT_Workbook: definedNames follow sheets
    const defs = newNames.join("");
    const wb = wbXml.includes("</definedNames>") ? wbXml.replace("</definedNames>", defs + "</definedNames>")
      : wbXml.includes("<definedNames/>") ? wbXml.replace("<definedNames/>", `<definedNames>${defs}</definedNames>`)
      : wbXml.replace("</sheets>", `</sheets><definedNames>${defs}</definedNames>`);
    zip.file("xl/workbook.xml", wb);
  }

  zip.file(sheetPath, sheetXml);
  zip.file(sheetRelsPath, sheetRels);
  zip.file(drawingPath, drawingXml);
  zip.file(relsPathOf(drawingPath), drawingRels);
  zip.file("[Content_Types].xml", ct);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
