/* Circle packing for packed bubbles: largest bubble in the middle, the others around it. */

/* the two centres where a circle of radius r touches circles p and q (gap included) */
function touchingBoth(p, q, r, gap) {
  const r1 = p.r + r + gap, r2 = q.r + r + gap;
  const dx = q.x - p.x, dy = q.y - p.y, d = Math.hypot(dx, dy);
  if (!d || d > r1 + r2 || d < Math.abs(r1 - r2)) return [];
  const a = (r1 * r1 - r2 * r2 + d * d) / (2 * d);
  const h = Math.sqrt(Math.max(0, r1 * r1 - a * a));
  const mx = p.x + a * dx / d, my = p.y + a * dy / d;
  return [{ x: mx + h * dy / d, y: my - h * dx / d }, { x: mx - h * dy / d, y: my + h * dx / d }];
}

/**
 * Greedy packing, biggest first: every circle goes where it touches two placed circles, as close to
 * the centre as it fits. O(n³) – fine for the few hundred marks a packed-bubble view shows.
 * @param {number[]} radii
 * @param {number} [gap] space between circles, in radius units
 * @returns {{ centres: { x: number, y: number }[], box: { w: number, h: number, cx: number, cy: number } }}
 */
export function tvPackCircles(radii, gap = 0) {
  const placed = [];
  const centres = new Array(radii.length);
  const order = radii.map((_, i) => i).sort((a, b) => radii[b] - radii[a]);
  const fits = (x, y, r) => placed.every(p => Math.hypot(p.x - x, p.y - y) >= p.r + r + gap - 1e-9);
  for (const i of order) {
    const r = radii[i];
    let best = null;
    if (!placed.length) best = { x: 0, y: 0 };
    else if (placed.length === 1) best = { x: placed[0].r + r + gap, y: 0 };
    else {
      for (let a = 0; a < placed.length; a++) {
        for (let b = a + 1; b < placed.length; b++) {
          for (const c of touchingBoth(placed[a], placed[b], r, gap)) {
            const d = Math.hypot(c.x, c.y);
            if ((!best || d < best.d) && fits(c.x, c.y, r)) best = { ...c, d };
          }
        }
      }
      if (!best) best = { x: Math.max(...placed.map(p => p.x + p.r)) + r + gap, y: 0 };
    }
    placed.push({ x: best.x, y: best.y, r });
    centres[i] = { x: best.x, y: best.y };
  }
  const x0 = Math.min(...placed.map(p => p.x - p.r)), x1 = Math.max(...placed.map(p => p.x + p.r));
  const y0 = Math.min(...placed.map(p => p.y - p.r)), y1 = Math.max(...placed.map(p => p.y + p.r));
  return { centres, box: { w: x1 - x0, h: y1 - y0, cx: (x0 + x1) / 2, cy: (y0 + y1) / 2 } };
}
