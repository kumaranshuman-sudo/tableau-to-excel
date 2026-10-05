/* Tableau date label formats (Format → Header → Dates) applied to header values. */

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/**
 * A header value in Tableau's custom date format: "iLLLLL" → "J" (month initial), "MMM" → "Jan",
 * "MMM yy" → "Jan 24", "yyyy" → "2024" … The value can be a full date or a date part Tableau sends as
 * text (month / weekday name, "Q1", year). null when the format needs a part the value does not have.
 * @param {string} raw Tableau text-format ("i…" / "*…" prefix + ICU-style pattern)
 * @param {string} text the formatted value from the summary data
 * @param {DataValue} [dv]
 * @returns {string | null}
 */
export function tfFormatDateLabel(raw, text, dv) {
  const pattern = String(raw || "").replace(/^[i*]/, "");
  if (!/[yMLdEQq]/.test(pattern)) return null;
  const t = String(text || "").trim().toLowerCase();
  let date = null, month = null, weekday = null, quarter = null, year = null;
  const native = dv ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : null;
  if (native instanceof Date) date = native;
  else if (typeof native === "string" && /^\d{4}-\d{2}-\d{2}/.test(native)) date = new Date(native.slice(0, 10) + "T00:00:00Z");
  if (date && !isNaN(date.getTime())) {
    year = date.getUTCFullYear(); month = date.getUTCMonth(); weekday = date.getUTCDay(); quarter = Math.floor(month / 3) + 1;
  } else {
    date = null;
    const m = MONTHS.findIndex(n => n.toLowerCase() === t || n.slice(0, 3).toLowerCase() === t);
    if (m >= 0) { month = m; quarter = Math.floor(m / 3) + 1; }
    const d = DAYS.findIndex(n => n.toLowerCase() === t || n.slice(0, 3).toLowerCase() === t);
    if (d >= 0) weekday = d;
    const q = t.match(/^q([1-4])$/);
    if (q) quarter = +q[1];
    if (/^\d{4}$/.test(t)) year = +t;
  }
  if (month === null && weekday === null && quarter === null && year === null) return null;
  let missing = false;
  const out = pattern.replace(/'([^']*)'|y{1,4}|M{1,5}|L{1,5}|d{1,2}|E{1,5}|Q{1,4}|q{1,4}/g, (tok, literal) => {
    if (literal !== undefined) return literal;
    const c = tok[0], n = tok.length;
    const need = v => { if (v === null) missing = true; return v === null; };
    if (c === "y") return need(year) ? "" : n === 2 ? String(year).slice(-2) : String(year);
    if (c === "M" || c === "L") {
      if (need(month)) return "";
      return n >= 5 ? MONTHS[month][0] : n === 4 ? MONTHS[month] : n === 3 ? MONTHS[month].slice(0, 3)
           : n === 2 ? String(month + 1).padStart(2, "0") : String(month + 1);
    }
    if (c === "d") return need(date) ? "" : String(date.getUTCDate()).padStart(n, "0");
    if (c === "E") return need(weekday) ? "" : n >= 5 ? DAYS[weekday][0] : n === 4 ? DAYS[weekday] : DAYS[weekday].slice(0, 3);
    return need(quarter) ? "" : n >= 3 ? "Q" + quarter : String(quarter);
  });
  return missing ? null : out;
}
