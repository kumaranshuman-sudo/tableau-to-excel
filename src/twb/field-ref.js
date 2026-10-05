/* Tableau field references ("[ds].[sum:Sales:qk]") – parsing and comparison. */
import { tfNorm } from "../util.js";

/* ── Field references ──────────────────────────────────────────────────────
 * "[federated.0x].[sum:Sales:qk]" → {ds, deriv:"sum", name:"Sales", type:"qk"}
 * type: qk = continuous measure, nk/ok = discrete                          */
/** @param {string} ref @returns {FieldRef | null} */
export function tfParseFieldRef(ref) {
  if (!ref) return null;
  ref = String(ref).trim().replace(/^"|"$/g, "");
  const parts = ref.match(/\[[^\]]*\]/g) || [];
  if (!parts.length) return { raw: ref, ds: null, deriv: null, name: ref, type: null, inner: ref };
  const inner = parts[parts.length - 1].slice(1, -1);
  const ds = parts.length > 1 ? parts[0].slice(1, -1) : null;
  const seg = inner.split(":");
  let deriv = null, name = inner, type = null;
  // "[usr:Calculation_1:ok:9]" → type is the nk/ok/qk segment, not the last one
  let ti = -1;
  for (let i = seg.length - 1; i >= 2; i--) if (/^(nk|ok|qk)$/.test(seg[i])) { ti = i; break; }
  if (ti >= 2) { deriv = seg[0] || null; type = seg[ti]; name = seg.slice(1, ti).join(":"); }
  else if (seg.length === 2 && seg[0] === "") name = seg[1];          // [:Measure Names]
  return { raw: ref, ds, deriv, name, type, inner };
}

/** a quick table calculation's measure: "[cum:sum:Sales:qk:7]" (running total) → "[sum:Sales:qk]"; null for any
 * other field @param {FieldRef} ref @returns {FieldRef | null} */
export function tfTableCalcBase(ref) {
  if (!ref || !ref.deriv || !ref.type || !/^(sum|avg|cnt|ctd|min|max|med|attr|usr|std|stdp|var|varp|none):/i.test(ref.name || "")) return null;
  const inner = `${ref.name}:${ref.type}`;
  return tfParseFieldRef(ref.ds ? `[${ref.ds}].[${inner}]` : `[${inner}]`);
}

/** @param {string} text @returns {FieldRef[]} */
export function tfExtractRefs(text) {
  if (!text) return [];
  const m = String(text).match(/(?:\[[^\]]+\]\.)?\[[^\]]+\]/g) || [];
  return m.map(tfParseFieldRef).filter(Boolean);
}

export function tfRefKey(r) { return ((r.ds || "") + "|" + r.inner).toLowerCase(); }

/** @param {Partial<FieldRef> | null} a @param {Partial<FieldRef> | null} b @returns {boolean} */
export function tfSameField(a, b) {
  if (!a || !b) return false;
  if (a.inner.toLowerCase() === b.inner.toLowerCase()) return true;
  return tfNorm(a.name) === tfNorm(b.name) && (a.deriv || "none") === (b.deriv || "none");
}
