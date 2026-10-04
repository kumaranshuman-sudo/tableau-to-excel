/* Tableau-drawn images via createVizImageAsync (spec building, SVG → PNG). */
import { tvColorScale, tvMarkColor, tvMeasureLabel, tvMeasureMark, tvText } from "../charts/model/common.js";
import { tvRoles } from "../charts/model/roles.js";
import { VISUAL_TYPES } from "../config.js";
import { tfDvNum, tfNaturalCompare } from "../data/values.js";

/* createVizImageAsync input spec, built from the same field roles as the native charts
 * (tvRoles in visual_chart_model.js): real shelves, colour, size, labels and view order.
 * One value axis → v1 spec; several measures (combo / dual axis, Measure Values) → v2 spec.
 * The API draws bar / line / area / square / circle / text marks only and has one field per
 * shelf and no Detail channel, so several category levels are joined into one ordered field. */
export const VIZ_IMAGE_DISCRETE_PALETTES = new Set(["tableau10_10_0", "tableau20_10_0", "color_blind_10_0", "seattle_grays_10_0",
  "traffic_light_10_0", "superfishel_stone_10_0", "miller_stone_10_0", "nuriel_stone_10_0", "jewel_bright_10_0", "summer_10_0",
  "winter_10_0", "green_orange_cyan_yellow_10_0", "blue_red_brown_10_0", "purple_pink_gray_10_0", "tableau-10",
  "tableau-10-medium", "tableau-20", "cyclic_10_0"]);

export const VIZ_IMAGE_ORDER_FIELD = "__order";

/** @param {VisualModel} visualModel @param {number} width @param {number} height @returns {Record<string, any>} */
export function buildVizImageSpec(visualModel, width, height) {
  const vm = visualModel.viewModel;
  const T = VISUAL_TYPES;
  const type = visualModel.type;
  if (!vm || !vm.rows.length) throw new Error("visual has no data rows");
  const roles = tvRoles(vm, visualModel.formatModel, visualModel.source && visualModel.source.visualSpec);
  const markToken = visualModel.metadata.markToken || "";

  // readable, unique field names
  const taken = new Set([VIZ_IMAGE_ORDER_FIELD, "__mark"]);
  const keys = vm.cols.map((c, i) => {
    let k = String(c.label || c.name || `Field${i + 1}`);
    while (taken.has(k)) k += " ";
    taken.add(k);
    return k;
  });
  const kinds = new Map();                                   // column → "discrete" | "continuous"
  const field = (ci, continuous, extra) => {
    kinds.set(ci, continuous ? "continuous" : "discrete");
    return { field: keys[ci], type: continuous ? "continuous" : "discrete", ...extra };
  };
  const title = ci => tvMeasureLabel(vm, ci);

  // colour: TWB palette / mapping → named discrete palette or custom gradient end points
  let color = null;
  if (roles.color && !roles.color.measureNames && roles.color.ci >= 0) {
    const ci = roles.color.ci;
    if (roles.color.continuous) {
      const scale = tvColorScale(vm, roles, markToken);
      const nums = vm.rows.map(r => tfDvNum(r[ci])).filter(n => n !== null);
      if (scale && nums.length) {
        const min = Math.min(...nums), max = Math.max(...nums);
        color = min < 0 && max > 0
          ? field(ci, true, { palette: "custom-diverging", start: "#" + scale(min), end: "#" + scale(max) })
          : field(ci, true, { palette: "custom-sequential", end: "#" + scale(max) });
      } else color = field(ci, true);
    } else {
      const enc = visualModel.formatModel && vm.fmt.colorEncoding ? vm.fmt.colorEncoding() : null;
      const named = enc && enc.def && enc.def.paletteName;
      color = field(ci, false, { palette: VIZ_IMAGE_DISCRETE_PALETTES.has(named) ? named : "tableau10_10_0" });
    }
  }
  const sizeField = roles.size >= 0 ? field(roles.size, true) : null;
  const labelDim = roles.textDims[0] ?? roles.detailDims[0] ?? -1;
  /* No Detail channel: without a per-mark field Tableau would aggregate the marks (a scatter of
   * customers collapses to one point per colour). Label them; with many marks the label is an
   * invisible zero-width key so the marks stay separate without covering the chart. */
  const synthetic = {};
  const markLabel = ci => {
    const distinct = new Set(vm.rows.map(r => tvText(r[ci])));
    if (distinct.size <= 40 || roles.textDims.includes(ci)) return field(ci, false);
    const index = new Map([...distinct].map((v, i) => [v, i]));
    synthetic.__mark = row => index.get(tvText(row[ci])).toString(2).replace(/0/g, "​").replace(/1/g, "‌");
    return { field: "__mark", type: "discrete" };
  };

  /** @type {Record<string, any>} createVizImageAsync input */
  let spec = null;
  let categoryKey = null, categoryOf = null;
  const base = {
    description: (visualModel.title && visualModel.title.text) || visualModel.metadata.worksheetName,
    markcolor: "#" + tvMarkColor(roles)
  };

  if (type === T.MAP) {
    const lat = vm.cols.findIndex(c => /latitude/i.test(c.name));
    const lon = vm.cols.findIndex(c => /longitude/i.test(c.name));
    if (lat < 0 || lon < 0) throw new Error("the map has no latitude/longitude in its summary data");
    const encoding = { columns: field(lon, true, { hidden: true }), rows: field(lat, true, { hidden: true }) };
    if (color) encoding.color = color;
    if (sizeField) encoding.size = sizeField;
    const place = roles.detailDims.find(ci => ci !== lat && ci !== lon);
    if (place !== undefined) encoding.text = markLabel(place);
    spec = { ...base, mark: "circle", encoding };
  } else if (type === T.TREEMAP || type === T.BUBBLE) {
    if (roles.size < 0 || labelDim < 0) throw new Error("no Size measure and label dimension to lay out the marks");
    const encoding = { size: sizeField, text: field(labelDim, false) };
    if (color) encoding.color = color;
    spec = { ...base, mark: type === T.TREEMAP ? "square" : "circle", encoding };
  } else if (type === T.SCATTER) {
    const xm = roles.cols.values[0], ym = roles.rows.values[0];
    if (!xm || !ym) throw new Error("scatter needs a measure on Rows and on Columns");
    const encoding = { columns: field(xm.ci, true, { title: title(xm.ci) }), rows: field(ym.ci, true, { title: title(ym.ci) }) };
    if (color) encoding.color = color;
    if (sizeField) encoding.size = sizeField;
    if (labelDim >= 0) encoding.text = markLabel(labelDim);
    spec = { ...base, mark: "circle", encoding };
  } else {
    // bar / line / area / histogram / combo: one ordered category field + one or more value axes
    const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
    if (!valueShelf) throw new Error("no continuous measure axis to draw");
    const catShelf = valueShelf === "rows" ? "cols" : "rows";
    const measures = roles[valueShelf].values;
    const catCis = [...roles[catShelf].dims, ...roles[valueShelf].dims].map(d => d.ci);
    if (catCis.length) {
      const order = new Map();
      categoryOf = row => catCis.map(ci => tvText(row[ci])).join(" · ");
      vm.rows.forEach(r => { const k = categoryOf(r); if (!order.has(k)) order.set(k, order.size); });
      if (catCis.some(ci => /^date/i.test(String(vm.cols[ci].dataType || "")))) {
        const firstRow = new Map();
        vm.rows.forEach(r => { const k = categoryOf(r); if (!firstRow.has(k)) firstRow.set(k, r); });
        const sorted = [...order.keys()].sort((a, b) => {
          for (const ci of catCis) { const d = tfNaturalCompare(firstRow.get(a)[ci], firstRow.get(b)[ci]); if (d) return d; }
          return 0;
        });
        order.clear();
        sorted.forEach((k, i) => order.set(k, i));
      }
      categoryKey = { name: catCis.map(ci => keys[ci]).join(" · "), order };
    }
    const catField = categoryKey ? { field: categoryKey.name, type: "discrete" } : null;
    const markOf = token => token === "line" ? "line" : token === "area" ? "area"
      : /^(circle|shape)$/.test(token) ? "circle" : token === "square" ? "square" : "bar";
    const shelfName = s => s === "cols" ? "columns" : "rows";
    if (measures.length === 1) {
      const m = measures[0];
      const encoding = {
        [shelfName(valueShelf)]: field(m.ci, true, { title: title(m.ci) }),
        sort: catField ? { field: categoryKey.name, sortby: VIZ_IMAGE_ORDER_FIELD, direction: "ascending" } : undefined
      };
      if (catField) encoding[shelfName(catShelf)] = catField;
      if (color) encoding.color = color;
      if (sizeField) encoding.size = sizeField;
      if (labelDim >= 0 && !catCis.includes(labelDim)) encoding.text = markLabel(labelDim);
      if (!encoding.sort) delete encoding.sort;
      spec = { ...base, mark: markOf(tvMeasureMark(roles, m.ref, markToken)), encoding };
    } else {
      const marks = measures.map((m, i) => markOf(tvMeasureMark(roles, m.ref, markToken, i, measures.length)));
      measures.forEach(m => field(m.ci, true));
      spec = {
        version: 2,
        description: base.description,
        vizlayout: { size: { width, height }, showcolorlegend: !!color },
        [shelfName(catShelf)]: catField ? [catField] : [],
        [shelfName(valueShelf)]: measures.map(m => ({ field: keys[m.ci], type: "continuous", title: title(m.ci) })),
        encodingaxis: shelfName(valueShelf),
        defaultencoding: { mark: marks[0] },
        encodings: marks.map(mark => (color ? { mark, color } : { mark }))
      };
    }
  }
  if (!spec.version) spec.size = { width, height };

  // data: every encoded column + the joined category and its order helper
  const perCategory = new Map();
  if (categoryKey) vm.rows.forEach(r => { const k = categoryOf(r); perCategory.set(k, (perCategory.get(k) || 0) + 1); });
  spec.data = {
    values: vm.rows.map(row => {
      const out = {};
      kinds.forEach((kind, ci) => {
        out[keys[ci]] = kind === "continuous" ? tfDvNum(row[ci]) : tvText(row[ci]);
      });
      for (const name in synthetic) out[name] = synthetic[name](row);
      if (categoryKey) {
        const k = categoryOf(row);
        out[categoryKey.name] = k;
        out[VIZ_IMAGE_ORDER_FIELD] = categoryKey.order.get(k) / perCategory.get(k);   // SUM over the category = its rank
      }
      return out;
    })
  };
  return spec;
}

export function svgToPngDataUrl(svg, width, height) {
  return new Promise((resolve, reject) => {
    const blob = new Blob([svg], { type: "image/svg+xml" });
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        canvas.getContext("2d").drawImage(image, 0, 0, width, height);
        URL.revokeObjectURL(url);
        resolve(canvas.toDataURL("image/png"));
      } catch (err) {
        URL.revokeObjectURL(url);
        reject(err);
      }
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Tableau returned an SVG that could not be rasterized"));
    };
    image.src = url;
  });
}

/** @param {VisualModel} visualModel @param {number} width @param {number} height @returns {Promise<string>} PNG, base64 */
export async function renderTableauImage(visualModel, width, height) {
  if (typeof tableau === "undefined" || !tableau.extensions ||
      typeof tableau.extensions.createVizImageAsync !== "function") {
    throw new Error("Tableau image API is unavailable in this runtime");
  }
  const spec = buildVizImageSpec(visualModel, width, height);
  const svg = await tableau.extensions.createVizImageAsync(spec);
  if (typeof svg !== "string" || !svg.trim()) throw new Error("Tableau image API returned no SVG");
  return svgToPngDataUrl(svg, width, height);
}
