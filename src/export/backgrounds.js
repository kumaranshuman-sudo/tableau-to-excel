/* Dashboard shading and container / object backgrounds (Layout pane → Background) → the colour of the
 * cells around and under the blocks on the sheet; thin coloured zones (dividers, accent lines) → borders. */

/**
 * A block on the sheet (0-based rows / columns, inclusive) and the zone it comes from.
 * own: the block's own opaque colour (ARGB); "white" = a sheet Tableau draws white (left unfilled);
 * undefined = see-through, the zone behind it shows.
 * @typedef {{ zoneId: string, top: number, left: number, bottom: number, right: number, own?: string }} BgBlock
 */

/**
 * A divider: along the top of row `at` ("h", columns from–to) or the left of column `at` ("v", rows from–to).
 * @typedef {{ dir: "h" | "v", at: number, from: number, to: number, color: string, style: "thin" | "medium" | "thick" }} BgLine
 */

/* thinner than this (px on the dashboard) a coloured zone is a line, not an area */
const LINE_PX = 8;

/* zone coordinate (0–100000) → grid edge: piecewise linear through the edges of the blocks laid out from
 * zones, never decreasing (the layout keeps the dashboard's order) */
function axisMap(anchors, lo, hi) {
  const pts = [[0, lo], ...anchors, [100000, hi]].sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  for (let i = 1; i < pts.length; i++) pts[i][1] = Math.max(pts[i][1], pts[i - 1][1]);
  return v => {
    if (v <= pts[0][0]) return pts[0][1];
    for (let i = 1; i < pts.length; i++) {
      if (v > pts[i][0]) continue;
      const [x0, g0] = pts[i - 1], [x1, g1] = pts[i];
      return x1 === x0 ? g1 : g0 + (g1 - g0) * (v - x0) / (x1 - x0);
    }
    return pts[pts.length - 1][1];
  };
}

/**
 * The colour each cell shows behind the blocks, as on the dashboard: the dashboard shading at the back,
 * then every container / backdrop zone with a background from the largest to the smallest, then each
 * block's own colour. A zone colours the blocks whose zones sit inside it (their centres) and the cells
 * between them – out to the zone's own edges, placed on the grid between the blocks' edges, so a header
 * bar spans the sheet as it spans the dashboard. Zones thinner than a few px are dividers: lines.
 * @param {DashboardZone[]} zones @param {string | null} dashColor ARGB
 * @param {BgBlock[]} blocks
 * @param {{ top: number, left: number, bottom: number, right: number }} canvas the dashboard's cells
 * @param {{ w: number, h: number }} dashPx the dashboard's size in px
 * @returns {{ colorAt: (row: number, col: number) => string | null, lines: BgLine[] }} colorAt: ARGB, null = leave the cell
 */
export function backgroundPlan(zones, dashColor, blocks, canvas, dashPx) {
  const zoneById = new Map((zones || []).map(z => [String(z.id), z]));
  const blockZones = new Set(blocks.map(b => b.zoneId));
  const placed = blocks.map(b => ({ b, z: zoneById.get(b.zoneId) })).filter(p => p.z && p.z.w > 0 && p.z.h > 0);
  const mapX = axisMap(placed.flatMap(({ b, z }) => [[z.x, b.left], [z.x + z.w, b.right + 1]]), canvas.left, canvas.right + 1);
  const mapY = axisMap(placed.flatMap(({ b, z }) => [[z.y, b.top], [z.y + z.h, b.bottom + 1]]), canvas.top, canvas.bottom + 1);
  const center = z => ({ x: z.x + z.w / 2, y: z.y + z.h / 2 });
  const inside = (p, z) => p.x >= z.x && p.x <= z.x + z.w && p.y >= z.y && p.y <= z.y + z.h;
  // a zone in a hidden container is not on the dashboard either
  const shown = z => { for (let q = z, n = 0; q && n < 50; q = zoneById.get(String(q.parent)), n++) if (q.hidden) return false; return true; };
  const coloured = (zones || []).map((z, order) => ({ z, order }))
    .filter(({ z }) => z.style && z.style.bgColor && !blockZones.has(String(z.id)) && z.w > 0 && z.h > 0 && shown(z));
  const px = z => ({ w: z.w * dashPx.w / 100000, h: z.h * dashPx.h / 100000 });

  /** @type {BgLine[]} */
  const lines = [];
  coloured.forEach(({ z }) => {
    const { w, h } = px(z);
    if (Math.min(w, h) >= LINE_PX || Math.max(w, h) < LINE_PX) return;
    const thick = Math.min(w, h), style = thick < 1.5 ? "thin" : thick < 2.5 ? "medium" : "thick";
    const c = center(z);
    const line = h < w
      ? { dir: "h", at: Math.round(mapY(c.y)), from: Math.round(mapX(z.x)), to: Math.round(mapX(z.x + z.w)) - 1 }
      : { dir: "v", at: Math.round(mapX(c.x)), from: Math.round(mapY(z.y)), to: Math.round(mapY(z.y + z.h)) - 1 };
    if (line.to >= line.from) lines.push(/** @type {BgLine} */ ({ ...line, color: z.style.bgColor, style }));
  });

  const layers = coloured
    .filter(({ z }) => { const { w, h } = px(z); return Math.min(w, h) >= LINE_PX; })
    .map(({ z, order }) => {
      const members = placed.filter(p => inside(center(p.z), z)).map(p => p.b);
      const rect = { top: Math.round(mapY(z.y)), left: Math.round(mapX(z.x)),
                     bottom: Math.round(mapY(z.y + z.h)) - 1, right: Math.round(mapX(z.x + z.w)) - 1 };
      members.forEach(b => {
        rect.top = Math.min(rect.top, b.top); rect.left = Math.min(rect.left, b.left);
        rect.bottom = Math.max(rect.bottom, b.bottom); rect.right = Math.max(rect.right, b.right);
      });
      if (rect.bottom < rect.top || rect.right < rect.left) return null;
      return { color: z.style.bgColor, area: z.w * z.h, order, members: new Set(members), ...rect };
    })
    .filter(Boolean)
    .sort((a, b) => b.area - a.area || a.order - b.order);             // back to front
  const covers = (b, r, c) => r >= b.top && r <= b.bottom && c >= b.left && c <= b.right;
  const colorAt = (r, c) => {
    const block = blocks.find(b => covers(b, r, c));
    if (block) {
      if (block.own === "white") return null;
      if (block.own) return block.own;
      for (let i = layers.length - 1; i >= 0; i--) if (layers[i].members.has(block)) return layers[i].color;
      return dashColor || null;
    }
    for (let i = layers.length - 1; i >= 0; i--) if (covers(layers[i], r, c)) return layers[i].color;
    return dashColor || null;
  };
  return { colorAt, lines };
}
