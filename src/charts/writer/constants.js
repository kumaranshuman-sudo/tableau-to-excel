/* OOXML namespaces, relationship and content types, axis ids. */

export const NS = {
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
  cx: "http://schemas.microsoft.com/office/drawing/2014/chartex",
  cx1: "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006"
};

export const REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

export const REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

export const REL_CHARTEX = "http://schemas.microsoft.com/office/2014/relationships/chartEx";

export const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";

export const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

export const CT_CHARTEX = "application/vnd.ms-office.chartex+xml";

export const REL_CHARTSTYLE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";

export const REL_CHARTCOLORS = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

export const CT_CHARTSTYLE = "application/vnd.ms-office.chartstyle+xml";

export const CT_CHARTCOLORS = "application/vnd.ms-office.chartcolorstyle+xml";

export const EMU_PER_PX = 9525;

export const AX = { cat: 50010, val: 50020, cat2: 50030, val2: 50040, x2: 50050, y2: 50060 };

export const C15 = "http://schemas.microsoft.com/office/drawing/2012/chart";

/* XLSX package surgery: relationship and content types of the parts injectCharts adds */
export const EMPTY_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.rels}"></Relationships>`;
