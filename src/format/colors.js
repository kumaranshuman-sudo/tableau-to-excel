/* Colour parsing and interpolation helpers. */

/* Tableau colours: "#rrggbb", "#rgb" or "#rrggbbaa" (aa=00 → transparent) → ARGB */
export function tfArgb(v) {
  if (v == null) return undefined;
  let s = String(v).trim().replace(/^#/, "");
  if (/^[0-9a-f]{8}$/i.test(s)) {
    if (s.slice(6) === "00") return null;          // explicit "no fill"
    s = s.slice(0, 6);
  }
  if (/^[0-9a-f]{3}$/i.test(s)) s = s.split("").map(c => c + c).join("");
  if (!/^[0-9a-f]{6}$/i.test(s)) return undefined;
  return "FF" + s.toUpperCase();
}

export function tfBrightness(argb) {
  const r = parseInt(argb.substring(2, 4), 16), g = parseInt(argb.substring(4, 6), 16), b = parseInt(argb.substring(6, 8), 16);
  return (r * 299 + g * 587 + b * 114) / 1000;
}

export function tfInterpolate(c1, c2, t) {
  const p = (c, i) => parseInt(c.substring(i, i + 2), 16);
  const ch = i => Math.round(p(c1, i) + (p(c2, i) - p(c1, i)) * t).toString(16).padStart(2, "0");
  return ("FF" + ch(2) + ch(4) + ch(6)).toUpperCase();
}

export function tfSample(colors, pos) {
  pos = Math.max(0, Math.min(1, pos));
  if (colors.length === 1) return colors[0];
  const seg = pos * (colors.length - 1);
  const i = Math.min(Math.floor(seg), colors.length - 2);
  return tfInterpolate(colors[i], colors[i + 1], seg - i);
}
