/* Summary data → view model (pivot, merge, sort, visible columns, titles). */
import { FORMAT_CONFIG } from "../config.js";
import { tfDvNum, tfDvText, tfIsNull, tfNaturalCompare } from "../data/values.js";
import { tfTextBoxHeaders } from "../twb/dashboard-text.js";
import { tfExtractRefs, tfSameField } from "../twb/field-ref.js";
import { createSheetFormatter } from "../twb/formatter.js";
import { tfBucketKey } from "../twb/parser.js";
import { tfLog, tfNorm } from "../util.js";

/* Tableau names a quick table calculation after its measure: a running sum and the plain sum both arrive as
 * "SUM(Value Ordered)". Of the columns sharing a name, the one holding the calculation takes the sheet's
 * table-calculation field: a running sum holds the plain column's grand total (its last value), a percent of
 * total adds up to 1; otherwise the calculation is the first of them, as Tableau lists it (cum… before sum…). */
function splitTableCalcColumns(cols, rows, fmt) {
  const groups = new Map();
  cols.forEach((c, i) => { if (c.ref) { const k = tfNorm(c.name); groups.set(k, [...(groups.get(k) || []), i]); } });
  groups.forEach(idx => {
    if (idx.length < 2 || !fmt.hasModel) return;
    const calcs = fmt.tableCalcRefs(cols[idx[0]].ref);
    if (!calcs.length) return;
    const nums = i => rows.map(r => tfDvNum(r[i])).filter(v => v !== null);
    const sum = i => nums(i).reduce((a, b) => a + b, 0);
    const near = (a, b) => Math.abs(a - b) <= Math.max(1e-9, Math.abs(b) * 1e-6);
    let free = [...idx];
    calcs.forEach(calc => {
      if (free.length < 2) return;
      const test = fmt.isRunningTotal(calc) ? (i, j) => nums(i).some(v => near(v, sum(j)))
                 : /^pcto$/i.test(calc.deriv || "") ? (i) => near(sum(i), 1) || near(sum(i), 100) : null;
      let pick = test ? free.find(i => free.some(j => j !== i && test(i, j))) : undefined;
      if (pick === undefined) pick = free[0];
      cols[pick].ref = calc;
      free = free.filter(i => i !== pick);
    });
  });
}

/**
 * @param {FormatModel | null} model
 * @param {string} sheetName
 * @param {SummaryData} summary
 * @param {{ dashboardName?: string, displayName?: string, [key: string]: any }} [opts]
 * @returns {ViewModel}
 */
export function buildViewModel(model, sheetName, summary, opts = {}) {
  const fmt = createSheetFormatter(model, sheetName);
  const notes = [];
  /** @type {ViewColumn[]} */
  let cols = (summary.columns || []).map((c, i) => {
    const name = c.fieldName || c.fieldId || `Col${i + 1}`;
    return { name, dataType: c.dataType, ref: fmt.matchName(name) };
  });
  let rows = (summary.data || []).map(r => r.slice());
  splitTableCalcColumns(cols, rows, fmt);

  // ── 1. pivot Measure Names / Measure Values ─────────────────────────────
  const mnI = cols.findIndex(c => /^measure names$/i.test(c.name));
  const mvI = cols.findIndex(c => /^measure values$/i.test(c.name));
  if (mnI >= 0 && mvI >= 0) {
    const keyIdx = cols.map((_, i) => i).filter(i => i !== mnI && i !== mvI);
    const order = fmt.measureOrder();
    const measures = [];                         // [{text, ref, rank}]
    const groups = new Map();
    rows.forEach(r => {
      const text = tfDvText(r[mnI]);
      if (!measures.some(m => m.text === text)) {
        const ref = fmt.measureRefForName(text);
        const pos = ref ? order.findIndex(o => tfSameField(o, ref)) : -1;
        measures.push({ text, ref, rank: pos >= 0 ? pos : 1e6 + measures.length });
      }
      const key = keyIdx.map(i => tfDvText(r[i])).join("\u0001");
      if (!groups.has(key)) groups.set(key, { keyRow: keyIdx.map(i => r[i]), vals: {} });
      const g = groups.get(key);
      if (g.vals[text] === undefined) g.vals[text] = r[mvI];
    });
    measures.sort((a, b) => a.rank - b.rank);
    cols = [...keyIdx.map(i => cols[i]),
            ...measures.map(m => ({ name: m.text, ref: m.ref, pivoted: true, dataType: "float" }))];
    rows = [...groups.values()].map(g => [...g.keyRow, ...measures.map(m => g.vals[m.text] || null)]);
    notes.push(`pivoted ${measures.length} measures → ${rows.length} rows`);
  }

  // column roles
  cols.forEach(c => {
    c.isHeader = !c.pivoted && fmt.hasModel ? fmt.isHeaderField(c.ref)
               : !c.pivoted && !/^(int|float|real|integer|number)$/i.test(String(c.dataType || ""));
  });

  // ── 2. merge stacked multi-pane rows ────────────────────────────────────
  const hdrIdx = cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);
  if (hdrIdx.length && rows.length > 1) {
    const groups = new Map();
    rows.forEach(r => {
      const key = hdrIdx.map(i => tfDvText(r[i])).join("\u0001");
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    });
    if (groups.size < rows.length) {
      let conflict = false;
      const merged = [];
      for (const g of groups.values()) {
        const out = g[0].slice();
        for (let ci = 0; ci < cols.length && !conflict; ci++) {
          const vals = [...new Set(g.filter(r => !tfIsNull(r[ci])).map(r => tfDvText(r[ci])))];
          if (vals.length > 1) conflict = true;
          else { const hit = g.find(r => !tfIsNull(r[ci])); if (hit) out[ci] = hit[ci]; }
        }
        if (conflict) break;
        merged.push(out);
      }
      if (!conflict) { notes.push(`merged ${rows.length} stacked rows → ${merged.length}`); rows = merged; }
      else notes.push("rows share header values but differ – not merged");
    }
  }

  // ── 3. sort like the view ───────────────────────────────────────────────
  const sortLevels = [];
  (fmt.hasModel ? fmt.headerRefs() : []).forEach(ref => {
    const ci = cols.findIndex(c => c.isHeader && c.ref && tfSameField(c.ref, ref));
    if (ci < 0) return;
    const manual = fmt.manualSortFor(ref);
    const byMeasure = fmt.measureSortFor(ref);
    let mi = -1;
    if (byMeasure) mi = cols.findIndex(c => c.ref && tfSameField(c.ref, byMeasure.measure));
    // sort measure not in the summary data (e.g. a funnel sorted by a hidden field) → the row's own total
    sortLevels.push({ ci, manual, byMeasure, mi });
  });
  if (sortLevels.length) {
    // measure sort aggregates the measure over the rows sharing the same outer path
    const aggCache = sortLevels.map((lvl, L) => {
      if (!lvl.byMeasure) return null;
      const m = new Map();
      rows.forEach(r => {
        const k = sortLevels.slice(0, L + 1).map(x => tfDvText(r[x.ci])).join("\u0001");
        const v = lvl.mi >= 0 ? tfDvNum(r[lvl.mi])
          : cols.reduce((s, c, i) => s + (!c.isHeader && tfDvNum(r[i]) !== null ? tfDvNum(r[i]) : 0), 0);
        m.set(k, (m.get(k) || 0) + (v || 0));
      });
      return m;
    });
    const manualRank = lvl => {
      if (!lvl.manual) return null;
      const map = new Map(lvl.manual.order.map((b, i) => [tfBucketKey(b), i]));
      return map;
    };
    const ranks = sortLevels.map(manualRank);
    rows.sort((a, b) => {
      for (let L = 0; L < sortLevels.length; L++) {
        const lvl = sortLevels[L];
        let d = 0;
        if (ranks[L]) {
          const ra = ranks[L].get(tfNorm(tfDvText(a[lvl.ci]))), rb = ranks[L].get(tfNorm(tfDvText(b[lvl.ci])));
          d = (ra === undefined ? 1e9 : ra) - (rb === undefined ? 1e9 : rb);
          if (lvl.manual.direction === "DESC") d = -d;
        } else if (lvl.byMeasure) {
          const path = r => sortLevels.slice(0, L + 1).map(x => tfDvText(r[x.ci])).join("\u0001");
          d = (aggCache[L].get(path(a)) || 0) - (aggCache[L].get(path(b)) || 0);
          if (lvl.byMeasure.direction === "DESC") d = -d;
        }
        if (!d) d = tfNaturalCompare(a[lvl.ci], b[lvl.ci]);
        if (d) return d;
      }
      return 0;
    });
  }

  // ── URL actions → link columns (kept visible even when the field is only on Detail)
  const linkActions = fmt.urlActions(opts.dashboardName);
  linkActions.forEach(a => {
    const refs = tfExtractRefs(a.expression);
    if (!refs.length) return;
    const ci = cols.findIndex(c => !c.isHeader && c.ref && tfNorm(c.ref.name) === tfNorm(refs[0].name));
    if (ci < 0) return;
    cols[ci].link = { caption: a.caption, expression: a.expression,
      parts: refs.map(r => ({ token: r, ci: cols.findIndex(c => c.ref && tfNorm(c.ref.name) === tfNorm(r.name)) })) };
  });

  // ── 4. visible columns in visual order + header labels ──────────────────
  let order;
  if (!fmt.hasModel) {
    order = cols.map((_, i) => i);
  } else {
    const panes = fmt.panesInOrder();
    const headerRefs = fmt.headerRefs();
    /** @type {{ i: number, g: number, a: number, b: number }[]} column, group, then sort keys */
    const place = [];
    cols.forEach((c, i) => {
      if (c.isHeader) {
        if (fmt.isLabelHidden(c.ref)) return;                       // Format → Hide header
        const pos = headerRefs.findIndex(r => tfSameField(r, c.ref));
        place.push({ i, g: 0, a: pos, b: 0 });
      } else if (c.pivoted) {
        place.push({ i, g: 1, a: 0, b: i });
      } else if (!c.ref) {
        place.push({ i, g: 2, a: 1e6, b: i });                     // unknown → keep, at the end
      } else {
        /** @type {{ a: number, b: number } | null} */
        let hit = null;
        panes.forEach((p, pi) => {
          const k = p.refs.findIndex(r => tfSameField(r, c.ref));
          if (k >= 0 && !hit) hit = { a: pi, b: k };
        });
        if (hit) place.push({ i, g: 2, a: hit.a, b: hit.b });      // shown as a mark label
        else if (c.link) {                                         // link column: in the pane that carries the field
          const pi = panes.findIndex(p => p.pane.encodings.some(e => tfSameField(e.field, c.ref)));
          place.push({ i, g: 2, a: pi >= 0 ? pi : 1e5, b: 999 });
        }
        // else: detail / tooltip / colour-only / axis measure → not visible
      }
    });
    place.sort((x, y) => x.g - y.g || x.a - y.a || x.b - y.b);
    order = place.map(p => p.i);
    if (!order.length) { order = cols.map((_, i) => i); notes.push("no visible columns resolved – exporting all"); }
  }

  const labelsRows = fmt.fieldLabelsShown("rows"), labelsCols = fmt.fieldLabelsShown("cols");
  cols.forEach(c => {
    if (c.pivoted) c.label = c.name;                                          // Measure Names alias
    else if (c.isHeader) {
      const shown = fmt.shelfOf(c.ref) === "cols" ? labelsCols : labelsRows;
      c.label = shown ? fmt.captionFor(c.ref, c.name) : "";
    } else c.label = fmt.captionFor(c.ref, c.name);
    if (c.link && c.link.caption) c.label = c.link.caption;
  });

  // ── dashboard text boxes drawn as column headers (strict match, else captions stay) ──
  const dashM = model && opts.dashboardName && model.dashboards && model.dashboards[opts.dashboardName];
  let headerZoneIds = [];
  if (FORMAT_CONFIG.textBoxHeaders && dashM && !labelsRows) {
    const tb = tfTextBoxHeaders(dashM, sheetName, order.length);
    if (tb) {
      order.forEach((ci, k) => {
        const z = tb.zones[k];
        cols[ci].label = z.text;
        cols[ci].labelProps = z.props;
        if (dashM.width) cols[ci].zoneWidthPx = z.w / 100000 * dashM.width;
      });
      headerZoneIds = tb.zones.map(z => z.id);
      notes.push(`headers from ${tb.zones.length} dashboard text boxes`);
    }
  }
  const showHeaderRow = order.some(i => cols[i].label);

  // title
  const valuesFor = ref => {
    const ci = cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
    if (ci < 0) return [];
    return [...new Set(rows.map(r => tfDvText(r[ci])).filter(Boolean))];
  };
  const title = fmt.renderTitle(valuesFor, opts.displayName || sheetName);
  const dash = model && opts.dashboardName && model.dashboards && model.dashboards[opts.dashboardName];
  const zone = dash && dash.zones.find(z => z.name === sheetName && z.type === "worksheet");
  const showTitle = zone ? zone.showTitle : true;

  if (notes.length) tfLog(`View model "${sheetName}": ${notes.join("; ")}`);
  return {
    fmt, cols, rows, order, notes, showHeaderRow, title, showTitle, headerZoneIds,
    dashboardName: opts.dashboardName,
    kind: fmt.isChart() ? "chart" : "table",
    headerOrder: sortLevels.map(l => l.ci)          // all header cols (incl. hidden) outer → inner
  };
}
