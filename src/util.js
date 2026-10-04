/* Small shared helpers: logging, DOM child lookup, text/number normalisation. */
import { FORMAT_CONFIG } from "./config.js";

/* ── small helpers ─────────────────────────────────────────────────────── */
export function tfLog(...a) { if (FORMAT_CONFIG.debug) console.log("[Format]", ...a); }

export function tfKids(el, tag) {
  if (!el) return [];
  return Array.from(el.childNodes).filter(n => n.nodeType === 1 && (!tag || n.tagName === tag));
}

export function tfKid(el, tag) { return tfKids(el, tag)[0] || null; }

export function tfNorm(s) { return String(s == null ? "" : s).toLowerCase().replace(/\s+/g, " ").trim(); }

export function tfNum(v) { const n = parseFloat(v); return isFinite(n) ? n : undefined; }

/** copy of o without its undefined properties
 * @template T @param {T} o @returns {T} */
export function tfDefined(o) {
  const r = /** @type {T} */ ({});
  for (const k in o) if (o[k] !== undefined) r[k] = o[k];
  return r;
}
