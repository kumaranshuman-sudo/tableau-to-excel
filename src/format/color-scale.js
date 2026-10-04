/* Colour encodings → value-to-colour scales (categorical and continuous). */
import { tfArgb, tfSample } from "./colors.js";
import { TABLEAU_10, getAutomaticPaletteForMark, getBuiltInPaletteColors } from "./palettes.js";
import { tfLog, tfNorm } from "../util.js";

/* ── Colour scales ─────────────────────────────────────────────────────── */
export function tfBuildColorScale(fmt, enc, values) {
  const def = enc.def;
  if (enc.continuous) {
    const nums = values.filter(v => typeof v === "number" && isFinite(v));
    if (!nums.length) return null;
    const min = def && def.min !== undefined ? def.min : Math.min(...nums);
    const max = def && def.max !== undefined ? def.max : Math.max(...nums);
    let colors = def && def.customColors.length ? def.customColors : null;
    let diverging = /diverging/i.test((def && (def.paletteName || "")) + " " + (def && (def.paletteType || "")));
    if (!colors && def && def.paletteName && !/^automatic$/i.test(def.paletteName)) {
      const pc = getBuiltInPaletteColors(def.paletteName);
      if (pc) colors = pc.map(tfArgb);
    }
    if (!colors) {                     // Tableau "Automatic"
      if (min < 0 && max > 0) { colors = getBuiltInPaletteColors("orange_blue_diverging_10_0").map(tfArgb); diverging = true; }
      else colors = getAutomaticPaletteForMark(enc.markClass).colors;
    }
    if (def && def.reverse) colors = colors.slice().reverse();
    let center = def && def.center !== undefined ? def.center : null;
    if (diverging && center === null) center = (min < 0 && max > 0) ? 0 : (min + max) / 2;
    tfLog(`Gradient for ${enc.ref.inner}: ${colors.length} stops, range ${min}..${max}${center !== null ? ", center " + center : ""}`);
    return v => {
      if (typeof v !== "number" || !isFinite(v)) return null;
      let pos;
      if (center !== null) {
        const dev = Math.max(Math.abs(max - center), Math.abs(min - center)) || 1;
        pos = 0.5 + (v - center) / (2 * dev);       // symmetric, like "Use full colour range" = off
      } else pos = max === min ? 0.5 : (v - min) / (max - min);
      return tfSample(colors, pos);
    };
  }
  // categorical
  const map = {};
  if (def) for (const [k, c] of Object.entries(def.map)) fmt.bucketAliases(k).forEach(a => { map[tfNorm(a)] = c; });
  const used = new Set(Object.values(map));
  const pool = TABLEAU_10.filter(c => !used.has(c));
  const auto = {};
  [...new Set(values.filter(v => v != null && v !== "").map(tfNorm))].filter(k => !(k in map)).sort().forEach((k, i) => {
    const p = pool.length ? pool : TABLEAU_10;
    auto[k] = p[i % p.length];
  });
  tfLog(`Categorical colours for ${enc.ref.inner}: ${Object.keys(map).length} from TWB, ${Object.keys(auto).length} auto-assigned`);
  return v => (v == null || v === "") ? null : (map[tfNorm(v)] || auto[tfNorm(v)] || null);
}
