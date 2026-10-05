/* Dashboard image objects (logos, icons) → pictures on the sheet, from the files packaged in the .twbx. */

/** @typedef {{ type: "png" | "jpeg" | "gif" | "svg" | null, width: number, height: number }} ImageInfo */

/* below this size (px) an image is a divider line or spacer, not a picture */
const MIN_IMAGE_PX = 8;

/**
 * The file's format and pixel size from its header: PNG IHDR, JPEG SOF, GIF screen descriptor, SVG
 * width / height / viewBox. type null = not a format Excel or the browser can draw.
 * @param {Uint8Array} data @param {string} [path] @returns {ImageInfo}
 */
export function imageInfo(data, path = "") {
  const b = data || new Uint8Array(0);
  const u16 = i => (b[i] << 8) | b[i + 1], u32 = i => ((b[i] << 24) >>> 0) + (b[i + 1] << 16) + (b[i + 2] << 8) + b[i + 3];
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return { type: "png", width: u32(16), height: u32(20) };
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return { type: "gif", width: b[6] | (b[7] << 8), height: b[8] | (b[9] << 8) };
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    for (let i = 2; i + 9 < b.length;) {
      if (b[i] !== 0xff) { i++; continue; }
      const m = b[i + 1];
      // SOF0–SOF15 carry the frame size (not DHT C4, JPG C8, DAC CC)
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { type: "jpeg", width: u16(i + 7), height: u16(i + 5) };
      if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { i += 2; continue; }
      i += 2 + u16(i + 2);
    }
    return { type: "jpeg", width: 0, height: 0 };
  }
  const head = new TextDecoder().decode(b.subarray(0, 4096));
  if (/\.svg$/i.test(path) || /<svg[\s>]/i.test(head)) {
    const tag = (head.match(/<svg[^>]*>/i) || [""])[0];
    const attr = n => { const m = tag.match(new RegExp(`\\s${n}\\s*=\\s*["']([\\d.]+)(px)?["']`, "i")); return m ? +m[1] : 0; };
    const vb = (tag.match(/viewBox\s*=\s*["']\s*[-\d.]+[\s,]+[-\d.]+[\s,]+([\d.]+)[\s,]+([\d.]+)/i) || []).slice(1).map(Number);
    return { type: "svg", width: attr("width") || vb[0] || 0, height: attr("height") || vb[1] || 0 };
  }
  return { type: null, width: 0, height: 0 };
}

/**
 * How each image object can be drawn on the grid: "backdrop" (a background behind other objects – not
 * exported, it would cover the cells), "overlay" (an icon on top of a worksheet / text box: floats over
 * that block, host = its id), "tiny" (dividers, spacers) or "block" (a logo with its own zone).
 * Only visible objects count, as Tableau shows them.
 * @param {any[]} objects Tableau DashboardObjects
 * @returns {Map<string, { kind: "backdrop" | "overlay" | "tiny" | "block", host?: string }>}
 */
export function classifyImageObjects(objects) {
  const box = o => ({ x: o.position.x, y: o.position.y, w: (o.size || {}).width || 0, h: (o.size || {}).height || 0 });
  const inside = (px, py, r) => px >= r.x && px <= r.x + r.w && py >= r.y && py <= r.y + r.h;
  const shown = (objects || []).filter(o => o && o.position && o.isVisible !== false);
  const others = shown.filter(o => /^(worksheet|text|quick-filter|filter|parameter-control|parameter|legend)$/.test(o.type));
  const out = new Map();
  shown.filter(o => o.type === "image").forEach(o => {
    const r = box(o);
    if (r.w < MIN_IMAGE_PX || r.h < MIN_IMAGE_PX) return out.set(String(o.id), { kind: "tiny" });
    if (others.some(p => { const q = box(p); return inside(q.x + q.w / 2, q.y + q.h / 2, r); })) return out.set(String(o.id), { kind: "backdrop" });
    const host = others.find(p => /^(worksheet|text)$/.test(p.type) && inside(r.x + r.w / 2, r.y + r.h / 2, box(p)));
    out.set(String(o.id), host ? { kind: "overlay", host: String(host.id) } : { kind: "block" });
  });
  return out;
}

/**
 * The picture's rectangle inside its box (px), as Tableau draws an image object: Fit Image scales it to
 * fit keeping its proportions, otherwise it keeps its own size (shrunk only when larger than the box);
 * Center Image (on unless switched off) centres it, otherwise it sits top-left.
 * @param {number} natW @param {number} natH @param {number} boxW @param {number} boxH
 * @param {boolean} [scaled] @param {boolean} [centered] undefined = on
 * @returns {{ x: number, y: number, w: number, h: number }}
 */
export function fitImage(natW, natH, boxW, boxH, scaled, centered) {
  const nw = natW || boxW, nh = natH || boxH;
  const k = scaled ? Math.min(boxW / nw, boxH / nh) : Math.min(1, boxW / nw, boxH / nh);
  const w = Math.max(1, Math.round(nw * k)), h = Math.max(1, Math.round(nh * k));
  const center = centered !== false;
  return { x: center ? Math.round((boxW - w) / 2) : 0, y: center ? Math.round((boxH - h) / 2) : 0, w, h };
}

const B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/** @param {Uint8Array} bytes @returns {string} */
export function toBase64(bytes) {
  let out = "";
  for (let i = 0; i < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | ((bytes[i + 1] || 0) << 8) | (bytes[i + 2] || 0);
    out += B64[(n >> 18) & 63] + B64[(n >> 12) & 63] + (i + 1 < bytes.length ? B64[(n >> 6) & 63] : "=") + (i + 2 < bytes.length ? B64[n & 63] : "=");
  }
  return out;
}

/* a canvas to draw on, when the page has one (the browser; not the Node test runs) */
function makeCanvas(w, h) {
  try {
    const c = typeof document !== "undefined" && document.createElement ? document.createElement("canvas") : null;
    if (!c || typeof c.getContext !== "function") return null;
    c.width = w; c.height = h;
    return c;
  } catch (e) {
    return null;
  }
}

/* the file drawn at w × h px into a PNG (SVG has to be rasterised for Excel; big photos are shrunk) */
async function rasterize(data, type, w, h) {
  const canvas = makeCanvas(w, h);
  if (!canvas || typeof Image === "undefined" || typeof URL === "undefined" || !URL.createObjectURL) return null;
  const url = URL.createObjectURL(new Blob([data], { type: type === "svg" ? "image/svg+xml" : "image/" + type }));
  try {
    const img = new Image();
    await new Promise((resolve, reject) => { img.onload = resolve; img.onerror = () => reject(new Error("image could not be decoded")); img.src = url; });
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0, w, h);
    const dataUrl = canvas.toDataURL("image/png");
    return dataUrl.slice(dataUrl.indexOf(",") + 1);
  } catch (e) {
    console.warn("[Images] could not draw the image:", e.message);
    return null;
  } finally {
    URL.revokeObjectURL(url);
  }
}

/**
 * The picture for Excel at its drawn size: PNG / JPEG / GIF as packaged; SVG rasterised; a large file
 * shown small (a 1 MB icon in a 40 px zone) re-drawn at twice its size so the workbook stays small.
 * null when it cannot be drawn here (SVG without a canvas, unknown format).
 * @param {{ data: Uint8Array }} file @param {ImageInfo} info @param {number} w @param {number} h drawn size, px
 * @returns {Promise<{ base64: string, extension: "png" | "jpeg" | "gif" } | null>}
 */
export async function prepareImage(file, info, w, h) {
  const big = file.data.length > 150000 && info.width > 3 * w && info.height > 3 * h;
  if (info.type === "svg" || (big && info.type !== "gif")) {
    const png = await rasterize(file.data, info.type, Math.max(1, Math.round(w * 2)), Math.max(1, Math.round(h * 2)));
    if (png) return { base64: png, extension: "png" };
    if (info.type === "svg") return null;
  }
  return info.type ? { base64: toBase64(file.data), extension: /** @type {"png" | "jpeg" | "gif"} */ (info.type) } : null;
}

/** a URL Excel opens as a web link ("www.site.com/x" → "https://www.site.com/x"); null for anything else
 * @param {string | undefined} url @returns {string | null} */
export function webLink(url) {
  const u = String(url || "").trim();
  if (!u || /^(javascript|data|file|vbscript):/i.test(u)) return null;
  if (/^(https?|mailto):/i.test(u)) return u;
  return /^[\w-]+(\.[\w-]+)+([/?#].*)?$/.test(u) ? "https://" + u : null;
}

/**
 * The sheet position px from the top-left of cell (col, row), as an anchor ExcelJS writes as is:
 * cell + offset in EMU (fractional anchors would be scaled by ExcelJS's own column width guess).
 * @param {number} col @param {number} row 0-based @param {number} x @param {number} y px
 * @param {(c: number) => number} colPx @param {(r: number) => number} rowPx final column widths / row heights
 */
export function nativeAnchor(col, row, x, y, colPx, rowPx) {
  let c = col, r = row, dx = Math.max(0, x), dy = Math.max(0, y);
  for (let guard = 0; guard < 500 && colPx(c) > 0 && dx >= colPx(c); guard++) { dx -= colPx(c); c++; }
  for (let guard = 0; guard < 5000 && dy >= rowPx(r); guard++) { dy -= rowPx(r); r++; }
  return { nativeCol: c, nativeColOff: Math.round(dx * 9525), nativeRow: r, nativeRowOff: Math.round(dy * 9525) };
}
