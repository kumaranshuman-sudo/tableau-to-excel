/* Fetching worksheet summary data, filters and filter-value sheets from Tableau. */

export async function fetchAllSheetsData(sheets, concurrency = 4) {
const results = [];
const queue = [...sheets];

async function worker() {
  while (queue.length) {
    const sheet = queue.shift();
    try {
      const [dataResult, visualSpecResult] = await Promise.allSettled([
        // every mark, not only the selected ones: with a mark selected Tableau returns just that mark otherwise
        sheet.getSummaryDataAsync({ ignoreSelection: true }),
        typeof sheet.getVisualSpecificationAsync === "function"
          ? sheet.getVisualSpecificationAsync()
          : Promise.reject(new Error("getVisualSpecificationAsync is unavailable"))
      ]);
      if (dataResult.status === "rejected") throw dataResult.reason;
      results.push({
        sheet,
        data: dataResult.value,
        visualSpec: visualSpecResult.status === "fulfilled" ? visualSpecResult.value : null,
        visualSpecError: visualSpecResult.status === "rejected" ? visualSpecResult.reason : null,
        error: null
      });
    } catch (err) {
      results.push({ sheet, data: null, error: err });
    }
  }
}

const workers = Array(concurrency).fill().map(() => worker());
await Promise.all(workers);
return results;
}

/* ── Filter Extraction ─────────────────────────────────────────────────── */
export async function extractFilterValuesPerField(sheets) {
  const filterMap = new Map();
  
  for (const worksheet of sheets) {
    let filters = [];
    
    try {
      filters = await worksheet.getFiltersAsync();
    } catch (e) {
      console.warn(`Could not get filters for ${worksheet.name}:`, e);
      continue;
    }
    
    filters.forEach(filter => {
      const ignoredFields = ["Measure Names", "Measure Values"];
      if (ignoredFields.includes(filter.fieldName) || /^Month\(/i.test(filter.fieldName)) {
        return;
      }
      
      if (!filterMap.has(filter.fieldName)) {
        filterMap.set(filter.fieldName, {
          values: new Set(),
          worksheetNames: new Set()
        });
      }
      
      const filterData = filterMap.get(filter.fieldName);
      filterData.worksheetNames.add(worksheet.name);
      
      let value = "";
      
      if (filter.filterType === "categorical") {
        const appliedValues = (filter.appliedValues || []).map(v => v.formattedValue);
        appliedValues.forEach(v => filterData.values.add(v));
        value = appliedValues.join(", ");
      } else if (filter.filterType === "range") {
        // date ranges arrive as "6/1/2025 12:00:00 AM" – a midnight time is not part of the filter
        const end = v => String((v && v.formattedValue) || "").replace(/\s+12:00:00\s*AM$|\s+00:00:00$/i, "");
        value = end(filter.minValue) + " - " + end(filter.maxValue);
        filterData.values.add(value);
      } else {
        value = filter.filterType;
        filterData.values.add(value);
      }
    });
  }
  
  const result = {};
  for (const [fieldName, data] of filterMap.entries()) {
    result[fieldName] = Array.from(data.values);
  }
  
  return result;
}

/** The marks selected on a sheet (a selected pie slice is outlined, as Tableau shows it); [] when unavailable.
 * @param {any} sheet Tableau worksheet @returns {Promise<any[]>} */
export async function fetchSelectedMarks(sheet) {
  try {
    const marks = typeof sheet.getSelectedMarksAsync === "function" ? await sheet.getSelectedMarksAsync() : null;
    return (marks && marks.data) || [];
  } catch (e) {
    console.warn(`[Export] selected marks of "${sheet.name}" unavailable: ${e.message}`);
    return [];
  }
}

/* ── Check if worksheet is a filter value table ───────────────────────── */
export function isFilterValueWorksheet(sheetName, summaryData) {
  if (/filter[_\- ]?\d+/i.test(sheetName) || 
      /values[_\- ]?\d+/i.test(sheetName) ||
      /_filter_\d+/i.test(sheetName)) {
    return true;
  }
  
  if (summaryData && summaryData.columns && summaryData.columns.length === 1) {
    const firstColumnName = summaryData.columns[0]?.fieldName || "";
    if (firstColumnName === "Values" || firstColumnName === "VALUE") {
      return true;
    }
  }
  
  return false;
}
