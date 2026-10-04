/* Tableau number formats → Excel number formats. */

/* Tableau text-format → Excel numFmt.
 * - strips Tableau's type prefix (n/c/p/e/*); locale "standard" codes (C1033…) → null (inferred instead)
 * - moves thousands-scaling commas behind the decimals: "#,##0,.0K" → "#,##0.0,\"K\""
 * - quotes every literal letter (Excel rejects bare K, M, yrs …)
 * - validates the result; anything doubtful → null so we never write an invalid format */
export function tableauToExcelNumFmt(raw) {
  if (!raw) return null;
  let s = String(raw);
  if (/^[A-Za-z]\d*%?$/.test(s)) return null;
  if (/^[ncpes*]/i.test(s) && !/^[#0]/.test(s)) s = s.slice(1);
  const sections = tfSplitSections(s);
  if (!sections || sections.length > 4) return null;
  const out = sections.map(tfConvertSection);
  if (out.some(x => x === null)) return null;
  const res = out.join(";");
  return /[0#]/.test(res) ? res : null;
}

export function tfSplitSections(s) {
  const secs = []; let cur = "", q = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') q = !q;
    if (ch === "\\" && !q) { cur += ch + (s[i + 1] || ""); i++; continue; }
    if (ch === ";" && !q) { secs.push(cur); cur = ""; continue; }
    cur += ch;
  }
  if (q) return null;                        // unbalanced quotes
  secs.push(cur);
  return secs;
}

export function tfConvertSection(sec) {
  // tokens: quoted literal | escaped char | number pattern | other char
  const toks = [];
  const re = /"[^"]*"|\\.|[#0?,.]+(?:[eE][+-][0#]+)?|[\s\S]/g;
  let m, numDone = false;
  while ((m = re.exec(sec))) {
    const t = m[0];
    if (t[0] === '"' || t[0] === "\\") { toks.push(t); continue; }
    if (/^[#0?,.]/.test(t) && /[#0?]/.test(t) && !numDone) {
      numDone = true;
      const [intPart, ...decParts] = t.split(".");
      const dec = decParts.join("");
      const intTrail = (intPart.match(/,+$/) || [""])[0];                 // "#,##0," → scale
      const decScale = (dec.match(/,+$/) || [""])[0];
      const intClean = intPart.slice(0, intPart.length - intTrail.length);
      const decClean = dec.replace(/,+$/, "");
      toks.push(intClean + (decParts.length ? "." + decClean : "") + intTrail + decScale);
      continue;
    }
    if (/[A-Za-z]/.test(t)) {                                              // literal letter → quote
      toks.push('"' + t + '"');
      continue;
    }
    toks.push(t);
  }
  // merge adjacent quoted literals: "K""g" → "Kg"
  return toks.join("").replace(/""/g, "");
}

/* Parse a Tableau formatted value, build an Excel format and only accept it
 * if it reproduces the same number (round-trip check). */
export function inferExcelNumFmt(formatted, value) {
  if (formatted == null || typeof value !== "number" || !isFinite(value)) return null;
  const f = String(formatted).trim();
  const m = f.match(/^(\(?)([^\d(]*?)(-?)(\d[\d.,\s\u00A0']*)(.*?)(\)?)$/);
  if (!m) return null;
  const [, lpar, prefixRaw, minus, core, suffixRaw, rpar] = m;
  const neg = !!(lpar && rpar) || !!minus || /^-/.test(prefixRaw);
  const prefix = prefixRaw.replace(/-/g, "").trim() ? prefixRaw.replace(/-/g, "") : "";
  const suffix = suffixRaw;

  // decide decimal separator by trying both interpretations
  const tryParse = dec => {
    const grpSrc = dec === "." ? "[,\\s\\u00A0']" : "[.\\s\\u00A0']";
    const clean = core.trim().replace(new RegExp(grpSrc, "g"), "").replace(dec, ".");
    const n = parseFloat(clean);
    const decimals = clean.includes(".") ? clean.split(".")[1].length : 0;
    return { n, decimals, grouped: new RegExp(grpSrc).test(core.trim()) };
  };
  const unitScale = { "%": 100, K: 1e-3, M: 1e-6, B: 1e-9, Bn: 1e-9 };
  const unit = (suffix.trim().match(/^(%|K|M|Bn|B)\b/) || suffix.trim().match(/^%/) || [])[0] || null;
  const scaled = Math.abs(value) * (unit ? unitScale[unit] : 1);

  for (const dec of [".", ","]) {
    const p = tryParse(dec);
    if (!isFinite(p.n)) continue;
    const tol = 0.5 * Math.pow(10, -p.decimals) + 1e-9 * Math.max(1, scaled);
    let plainPercent = false;
    if (Math.abs(scaled - p.n) > tol) {
      if (unit === "%" && Math.abs(Math.abs(value) - p.n) <= tol) plainPercent = true; // value already ×100
      else continue;
    }
    // Tableau's "Automatic" format pads to ~6 significant digits ("348.000", "4.50000",
    // "24.86108%"); 3+ decimals are that padding, not a chosen format → keep what the value needs
    let decimals = p.decimals;
    if (decimals >= 3) {
      const frac = (String(p.n).split(".")[1] || "").length;
      decimals = Math.min(decimals, frac);
      if (decimals >= 3) decimals = Math.max(0, Math.min(decimals, 2 - Math.floor(Math.log10(Math.abs(p.n) || 1))));
    }
    let body = (p.grouped ? "#,##0" : "0") + (decimals ? "." + "0".repeat(decimals) : "");
    const q = t => t ? '"' + t.replace(/"/g, '""') + '"' : "";
    let rest = suffix;
    if (unit === "%" && !plainPercent) { body += "%"; rest = suffix.replace("%", ""); }
    else if (unit && unit !== "%") { body += ",".repeat({ K: 1, M: 2, B: 3, Bn: 3 }[unit]) + q(unit); rest = suffix.replace(unit, ""); }
    const pos = q(prefix) + body + q(rest);
    if (!neg) return pos;
    if (lpar && rpar) return pos + ";(" + pos + ")";
    return minus ? pos + ";" + q(prefix) + "-" + body + q(rest) : pos + ";-" + pos;
  }
  return null;
}
