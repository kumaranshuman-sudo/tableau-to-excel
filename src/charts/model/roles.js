/* Which summary column sits on which shelf / encoding (TWB or live spec). */
import { tvMarkToken } from "./common.js";
import { tfSameField } from "../../twb/field-ref.js";
import { tfDisplayNames } from "../../twb/formatter.js";
import { tfNorm } from "../../util.js";
import { tvIsMeasureRef } from "../../visual/classify.js";

/* ── field roles: which summary column sits on which shelf / encoding ───── */
/**
 * @param {ViewModel} vm
 * @param {FormatModel | null} model
 * @param {any} [liveSpec] the live visual specification, when no workbook was loaded
 * @returns {Roles}
 */
export function tvRoles(vm, model, liveSpec) {
  const sheet = vm.fmt.sheetModel;
  /** @type {Roles} */
  const roles = { rows: { values: [], dims: [], axisRefs: [] }, cols: { values: [], dims: [], axisRefs: [] }, measureNames: null,
                  color: null, angle: -1, size: -1, textDims: [], detailDims: [],
                  labelRefs: [], labelNames: [], panes: sheet ? sheet.panes : [], source: "none" };
  const addDim = (list, ci) => {
    if (ci >= 0 && (vm.cols[ci].isHeader || !numeric(vm.cols[ci])) && !list.includes(ci)) list.push(ci);
  };
  const pivoted = vm.cols.map((c, i) => c.pivoted ? i : -1).filter(i => i >= 0);
  const numeric = c => /^(int|float|real|integer|number|double)/i.test(String(c.dataType || "")) || c.pivoted;

  if (sheet) {
    roles.source = "twb";
    const colOf = ref => {
      if (!ref) return -1;
      let i = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
      if (i < 0) {
        const names = tfDisplayNames(model, ref);
        i = vm.cols.findIndex(c => names.includes(tfNorm(c.name)));
      }
      return i;
    };
    for (const shelf of ["rows", "cols"]) {
      let axisNo = 0;                              // each measure / Measure Values entry = one axis ("A + B")
      for (const ref of sheet[shelf]) {
        if (ref.name === "Measure Names") { roles.measureNames = shelf; continue; }
        if (ref.name === "Multiple Values") {
          roles[shelf].axisRefs.push(ref);
          const axis = axisNo++;
          pivoted.forEach(ci => roles[shelf].values.push({ ci, ref: vm.cols[ci].ref, mv: true, axis }));
          continue;
        }
        const ci = colOf(ref);
        if (ci < 0) continue;
        if (tvIsMeasureRef(model, ref)) { roles[shelf].axisRefs.push(ref); roles[shelf].values.push({ ci, ref, axis: axisNo++ }); }
        else roles[shelf].dims.push({ ci, ref, continuous: ref.type === "qk" });
      }
    }
    const enc = channel => {
      for (const p of sheet.panes) {
        const e = p.encodings.find(x => x.channel === channel);
        if (e) return e.field;
      }
      return null;
    };
    const colorRef = enc("color");
    if (colorRef) {
      if (colorRef.name === "Measure Names") roles.color = { measureNames: true, ref: colorRef };
      else {
        const ci = colOf(colorRef);
        if (ci >= 0) roles.color = { ci, ref: colorRef, continuous: tvIsMeasureRef(model, colorRef) || colorRef.type === "qk" };
      }
    }
    const angleRef = enc("wedge-size");
    if (angleRef) roles.angle = angleRef.name === "Multiple Values" ? (pivoted[0] ?? -1) : colOf(angleRef);
    const sizeRef = enc("size");
    if (sizeRef) roles.size = colOf(sizeRef);
    sheet.panes.forEach(p => {
      p.encodings.filter(e => e.channel === "text" || e.channel === "label").forEach(e => {
        roles.labelRefs.push(e.field);
        if (!tvIsMeasureRef(model, e.field)) addDim(roles.textDims, colOf(e.field));
      });
      p.encodings.filter(e => e.channel === "lod").forEach(e => addDim(roles.detailDims, colOf(e.field)));
      p.labelRuns.forEach(r => r.refs.forEach(x => roles.labelRefs.push(x)));
    });
    return roles;
  }

  // ── no workbook loaded: use the live visual specification ──
  const spec = liveSpec && typeof liveSpec === "object" ? liveSpec : null;
  const byName = name => {
    const key = tfNorm(name);
    const strip = s => tfNorm(s).replace(/^[a-z]+\((.*)\)$/, "$1");
    let i = vm.cols.findIndex(c => tfNorm(c.name) === key);
    if (i < 0) i = vm.cols.findIndex(c => strip(c.name) === strip(name));
    return i;
  };
  const fieldName = f => f == null ? "" : typeof f === "string" ? f : (f.name || f.fieldName || f.caption || "");
  if (spec && (Array.isArray(spec.rowFields) || Array.isArray(spec.columnFields))) {
    roles.source = "live";
    for (const [shelf, list] of [["rows", spec.rowFields || []], ["cols", spec.columnFields || []]]) {
      list.forEach(f => {
        const name = fieldName(f);
        if (/^measure names$/i.test(name)) { roles.measureNames = shelf; return; }
        if (/^measure values$/i.test(name)) { pivoted.forEach(ci => roles[shelf].values.push({ ci, ref: vm.cols[ci].ref, mv: true })); return; }
        const ci = byName(name);
        if (ci < 0) return;
        if (numeric(vm.cols[ci]) && !vm.cols[ci].isHeader) roles[shelf].values.push({ ci, ref: vm.cols[ci].ref });
        else roles[shelf].dims.push({ ci, ref: vm.cols[ci].ref, continuous: /^date/i.test(String(vm.cols[ci].dataType || "")) });
      });
    }
    const marks = Array.isArray(spec.marksSpecifications) ? spec.marksSpecifications : [];
    roles.liveMarks = marks.map(m => tvMarkToken(m && (m.type || m.markType)));   // one marks card per measure
    const active = marks[spec.activeMarksSpecificationIndex] || marks[0];
    (active && active.encodings || []).forEach(e => {
      const name = fieldName(e.field);
      const type = String(e.type || e.encodingType || "").toLowerCase();
      if (type === "color") {
        if (/^measure names$/i.test(name)) roles.color = { measureNames: true };
        else { const ci = byName(name); if (ci >= 0) roles.color = { ci, continuous: numeric(vm.cols[ci]) && !vm.cols[ci].isHeader }; }
      } else if (type === "angle") roles.angle = byName(name);
      else if (type === "size") roles.size = byName(name);
      else if (type === "label" || type === "text") {
        roles.labelNames.push(tfNorm(name));
        const ci = byName(name);
        if (ci >= 0 && !numeric(vm.cols[ci])) addDim(roles.textDims, ci);
      } else if (type === "detail") addDim(roles.detailDims, byName(name));
    });
    if (roles.rows.values.length || roles.cols.values.length) return roles;
  }

  // ── nothing structural known: dimensions → categories, measures → rows ──
  roles.source = "summary";
  vm.cols.forEach((c, ci) => {
    if (c.isHeader) roles.cols.dims.push({ ci, ref: c.ref, continuous: /^date/i.test(String(c.dataType || "")) });
    else if (numeric(c)) roles.rows.values.push({ ci, ref: c.ref, mv: !!c.pivoted });
  });
  return roles;
}
