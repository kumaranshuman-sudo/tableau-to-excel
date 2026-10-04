/* XML escaping, cell references and other small writer helpers. */

/* ── small helpers ─────────────────────────────────────────────────────── */
export function esc(s) {
  return String(s == null ? "" : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export function colName(i) {
  let s = "", n = i + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
  return s;
}

/* 0-based (col,row) → 'Sheet'!$A$1[:$B$9] */
export function cellRef(sheet, c1, r1, c2, r2) {
  const q = "'" + String(sheet).replace(/'/g, "''") + "'!";
  const a = "$" + colName(c1) + "$" + (r1 + 1);
  return q + a + (c2 === undefined ? "" : ":$" + colName(c2) + "$" + (r2 + 1));
}

export function hex(c) {
  const s = String(c || "").replace(/^#/, "").toUpperCase();
  return s.length === 8 ? s.slice(2) : s;
}

export function num(v) { return typeof v === "number" && isFinite(v) ? v : null; }

export function attr(tag, name) {
  const m = tag.match(new RegExp("\\s" + name.replace(":", "\\:") + "=\"([^\"]*)\""));
  return m ? m[1] : null;
}

/* boundaries of each category level: a new group starts where any level ≤ l changes */
export function levelStarts(levels) {
  const n = levels.length ? levels[0].length : 0;
  return levels.map((_, l) => Array.from({ length: n }, (__, i) =>
    i === 0 || levels.slice(0, l + 1).some(lv => lv[i] !== lv[i - 1])));
}
