/* Tableau summary-data values: null checks, text/number access, natural + calendar ordering. */

/* ══════════════════════════════════════════════════════════════════════════
 * VIEW MODEL — rebuild the table Tableau actually draws from summary data
 *   1. pivot Measure Names / Measure Values into one column per measure
 *   2. merge stacked multi-pane rows (one block per pane → one row)
 *   3. sort rows like the view (manual sort, sort-by-measure, else natural)
 *   4. keep only visible columns, in visual order, with Tableau's header labels
 * ══════════════════════════════════════════════════════════════════════════ */
/** @param {DataValue} dv */
export function tfIsNull(dv) {
  if (!dv) return true;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  return v === null || v === undefined || v === "%null%";
}

/** @param {DataValue} dv @returns {string} */
export function tfDvText(dv) { return tfIsNull(dv) ? "" : String(dv.formattedValue != null ? dv.formattedValue : dv.value); }

/** @param {DataValue} dv @returns {number | null} */
export function tfDvNum(dv) {
  if (tfIsNull(dv)) return null;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-eE]/g, ""));
  return isFinite(n) && /^[\s$€£¥(+-]*[\d.,]+/.test(String(v)) ? n : null;
}

/* discrete date parts arrive as names: "January" / "Jan" / "Monday" → calendar position */
export const TF_MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

export const TF_WEEKDAYS = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];

export function tfCalendarRank(text) {
  const t = String(text || "").trim().toLowerCase();
  const m = t.match(/^(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sept?|oct|nov|dec)$/);
  if (m) return { kind: "month", rank: TF_MONTHS.indexOf(m[1].slice(0, 3)) };
  const d = t.match(/^(sunday|monday|tuesday|wednesday|thursday|friday|saturday|sun|mon|tue|wed|thu|fri|sat)$/);
  if (d) return { kind: "weekday", rank: TF_WEEKDAYS.indexOf(d[1].slice(0, 3)) };
  return null;
}

/** View order of two values: numbers, dates, ISO strings, month / weekday names, then text; nulls last.
 * @param {DataValue} a @param {DataValue} b @returns {number} */
export function tfNaturalCompare(a, b) {
  const an = tfIsNull(a), bn = tfIsNull(b);
  if (an || bn) return an === bn ? 0 : an ? 1 : -1;           // nulls last
  const av = a.nativeValue !== undefined ? a.nativeValue : a.value;
  const bv = b.nativeValue !== undefined ? b.nativeValue : b.value;
  if (typeof av === "number" && typeof bv === "number") return av - bv;
  if (av instanceof Date && bv instanceof Date) return /** @type {any} */ (av) - /** @type {any} */ (bv);
  const iso = /^\d{4}-\d{2}-\d{2}/;
  if (typeof av === "string" && typeof bv === "string" && iso.test(av) && iso.test(bv)) return av < bv ? -1 : av > bv ? 1 : 0;
  const ra = tfCalendarRank(tfDvText(a)), rb = tfCalendarRank(tfDvText(b));   // MONTH / WEEKDAY names
  if (ra !== null && rb !== null && ra.kind === rb.kind) return ra.rank - rb.rank;
  return tfDvText(a).localeCompare(tfDvText(b), undefined, { numeric: true, sensitivity: "base" });
}
