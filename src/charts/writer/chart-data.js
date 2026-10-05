/* Chart source data → worksheet cells (charts stay editable and linked). */
import { cellRef, levelStarts, num } from "./xml-util.js";

/**
 * @param {import("exceljs").Worksheet} ws the (hidden) chart data sheet
 * @param {ChartSpec} spec
 * @param {number} startRow 0-based first row to write
 * @returns {ChartRefs}
 */
export function writeChartData(ws, spec, startRow) {
  const sheet = ws.name;
  const put = (r, c, v, fmt) => {
    const cell = ws.getCell(r + 1, c + 1);
    cell.value = v;
    if (fmt) cell.numFmt = fmt;
  };
  ws.getCell(startRow + 1, 1).font = { bold: true };

  if (spec.kind === "scatter" || spec.kind === "bubble") {
    const per = spec.kind === "bubble" ? 3 : 2;                 // x, y (, bubble size) columns per series
    let maxLen = 0;
    /** @type {ChartRefs["series"]} */
    const series = spec.series.map((s, k) => {
      const cx = per * k, cy = cx + 1, n = s.x.length;
      put(startRow, cx, (spec.xTitle || "X") + (spec.series.length > 1 ? " – " + s.name : ""));
      put(startRow, cy, s.name);
      for (let i = 0; i < n; i++) {
        if (num(s.x[i]) !== null) put(startRow + 1 + i, cx, s.x[i], spec.xNumFmt);
        if (num(s.y[i]) !== null) put(startRow + 1 + i, cy, s.y[i], spec.numFmt);
      }
      maxLen = Math.max(maxLen, n);
      /** @type {ChartRefs["series"][number]} */
      const ref = { tx: cellRef(sheet, cy, startRow),
                    x: cellRef(sheet, cx, startRow + 1, cx, startRow + Math.max(1, n)),
                    y: cellRef(sheet, cy, startRow + 1, cy, startRow + Math.max(1, n)) };
      if (per === 3) {
        const cs = cx + 2;
        put(startRow, cs, (spec.sizeTitle || "Size") + (spec.series.length > 1 ? " – " + s.name : ""));
        s.size.forEach((v, i) => { if (num(v) !== null) put(startRow + 1 + i, cs, v, spec.sizeNumFmt); });
        ref.size = cellRef(sheet, cs, startRow + 1, cs, startRow + Math.max(1, n));
      }
      return ref;
    });
    const lblCol = per * spec.series.length;
    spec.series.forEach((s, k) => {
      if (!s.labelTexts) return;
      const c = lblCol + k;
      put(startRow, c, s.name + " – label");
      s.labelTexts.forEach((v, i) => put(startRow + 1 + i, c, v));
      series[k].lbl = cellRef(sheet, c, startRow + 1, c, startRow + Math.max(1, s.labelTexts.length));
    });
    return { series, nextRow: startRow + maxLen + 3 };
  }

  const levels = spec.categories.levels;
  const L = levels.length;
  const N = L ? levels[0].length : 0;
  // treemap: every level written on every row – Excel reads each row as one full hierarchy path
  const starts = spec.kind === "treemap" ? levels.map(lv => lv.map(() => true)) : levelStarts(levels);
  spec.categories.names.forEach((name, l) => put(startRow, l, name || ""));
  spec.series.forEach((s, k) => put(startRow, L + k, s.name));
  for (let i = 0; i < N; i++) {
    for (let l = 0; l < L; l++) if (l === L - 1 || starts[l][i]) put(startRow + 1 + i, l, levels[l][i]);
    spec.series.forEach((s, k) => {
      const v = num(s.values[i]);
      if (v !== null) put(startRow + 1 + i, L + k, v, s.secondary ? spec.secondaryNumFmt : spec.numFmt);
    });
  }
  // label texts taken from cells (Excel's "Value From Cells"), after the series columns
  let lblCol = L + spec.series.length;
  const lbl = spec.series.map(s => {
    if (!s.labelTexts) return null;
    const c = lblCol++;
    put(startRow, c, s.name + " – label");
    s.labelTexts.forEach((t, i) => put(startRow + 1 + i, c, t));
    return cellRef(sheet, c, startRow + 1, c, startRow + Math.max(1, N));
  });
  return {
    cat: cellRef(sheet, 0, startRow + 1, L - 1, startRow + N),
    series: spec.series.map((s, k) => ({
      tx: cellRef(sheet, L + k, startRow),
      val: cellRef(sheet, L + k, startRow + 1, L + k, startRow + N),
      ...(lbl[k] ? { lbl: lbl[k] } : {})
    })),
    nextRow: startRow + N + 3
  };
}
