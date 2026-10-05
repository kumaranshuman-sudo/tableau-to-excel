/* Tableau number formats → Excel number formats. */

/* Tableau text-format → Excel numFmt.
 * - strips Tableau's type prefix (n/c/p/e/*); locale "standard" codes (C1033…) → null (inferred instead),
 *   while short ones are formats ("p0%" = whole percent)
 * - moves thousands-scaling commas behind the decimals: "#,##0,.0K" → "#,##0.0,\"K\""
 * - quotes every literal letter (Excel rejects bare K, M, yrs …)
 * - validates the result; anything doubtful → null so we never write an invalid format */
export function tableauToExcelNumFmt(raw) {
  if (!raw) return null;
  let s = String(raw);
  if (/^[A-Za-z]\d{3,}%?$/.test(s)) return null;
  if (/^[ncpes*]/i.test(s) && !/^[#0]/.test(s)) s = s.slice(1);
  const sections = tfSplitSections(s);
  if (!sections || sections.length > 4) return null;
  const out = sections.map(tfConvertSection);
  if (out.some(x => x === null)) return null;
  const res = out.join(";");
  return /[0#]/.test(res) ? res : null;
}

/**
 * A number as Excel shows it in a format, for text Excel does not format itself (label texts taken from
 * cells): literal text, digits with thousands separators, decimals (0 required, # optional), thousands
 * scaling (trailing commas: ,, = millions), percent; sections positive;negative;zero.
 * @param {number} value @param {string} fmt Excel number format @returns {string}
 */
export function tfFormatNumber(value, fmt) {
  if (typeof value !== "number" || !isFinite(value)) return value == null ? "" : String(value);
  const secs = tfSplitSections(String(fmt || "General")) || ["General"];
  let sec = secs[0], sign = value < 0 ? "-" : "";
  if (value < 0 && secs.length > 1) { sec = secs[1]; sign = ""; }   // the negative section shows the size
  else if (value === 0 && secs.length > 2) sec = secs[2];
  const x = Math.abs(value);
  let pre = "", post = "", pattern = "", phase = 0;                 // 0 before the number, 1 in it, 2 after
  const text = t => { if (phase === 1) phase = 2; if (phase === 0) pre += t; else post += t; };
  for (let i = 0; i < sec.length; i++) {
    const ch = sec[i];
    if (ch === '"') { const j = sec.indexOf('"', i + 1); text(sec.slice(i + 1, j < 0 ? sec.length : j)); i = j < 0 ? sec.length : j; }
    else if (ch === "\\") { text(sec[i + 1] || ""); i++; }
    else if (ch === "[") { const j = sec.indexOf("]", i); i = j < 0 ? sec.length : j; }      // [Red], [>100]
    else if (ch === "_") { text(" "); i++; }                                                 // _) = a space
    else if (ch === "*") i++;                                                                 // fill character
    else if (/[0#?.,%]/.test(ch) && phase < 2) { phase = 1; pattern += ch; }
    else if (/general/i.test(sec.slice(i, i + 7)) && phase === 0) { pattern = "G"; phase = 1; i += 6; }
    else text(ch);
  }
  if (pattern === "G" || !/[0#?]/.test(pattern)) {
    return sign + pre + (pattern === "G" ? String(+x.toPrecision(10)) : "") + post;
  }
  const pct = (pattern.match(/%/g) || []).length;
  const p = pattern.replace(/%/g, "");
  const dot = p.indexOf(".");
  let ip = dot >= 0 ? p.slice(0, dot) : p, dp = dot >= 0 ? p.slice(dot + 1) : "";
  let scale = 0;                                                   // commas after the last digit: ÷ 1000 each
  const it = ip.match(/,+$/); if (it) { scale += it[0].length; ip = ip.slice(0, -it[0].length); }
  const dt = dp.match(/,+$/); if (dt) { scale += dt[0].length; dp = dp.slice(0, -dt[0].length); }
  const minDec = (dp.match(/0/g) || []).length, maxDec = (dp.match(/[0#?]/g) || []).length;
  let s = (x / Math.pow(1000, scale) * Math.pow(100, pct)).toFixed(maxDec);
  if (maxDec > minDec) s = s.replace(new RegExp(`0{1,${maxDec - minDec}}$`), "").replace(/\.$/, "");
  let [whole, frac] = s.split(".");
  const minInt = (ip.match(/0/g) || []).length;
  if (whole.length < minInt) whole = whole.padStart(minInt, "0");
  if (!minInt && whole === "0" && frac) whole = "";
  if (ip.includes(",")) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return sign + pre + whole + (frac ? "." + frac : "") + "%".repeat(pct) + post;
}

/** an Excel number format with literal text around the number in every section ("4.0" → "4.0 DAYS")
 * @param {string} fmt @param {string} prefix @param {string} suffix @returns {string} */
export function tfWrapNumFmt(fmt, prefix, suffix) {
  const q = s => s ? `"${s.replace(/"/g, "")}"` : "";
  return (tfSplitSections(fmt || "General") || ["General"]).slice(0, 3).map(sec => q(prefix) + (sec || "General") + q(suffix)).join(";");
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
