/* Fallback conditional formatting for tables without Tableau colours. */

export function isNumeric(value) {
  if (value === null || value === undefined) return false;
  
  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/%/g, "")
    .trim();
  
  return cleaned !== "" && !isNaN(Number(cleaned));
}

export function shouldSkipConditionalFormatting(columnName, value, fieldRole = null) {
  if (!columnName) return false;
  
  const lowerColName = columnName.toLowerCase();
  
  const categoricalKeywords = [
    'year', 'quarter', 'month', 'date', 'time', 
    'category', 'region', 'product', 'name', 'id',
    'country', 'city', 'state', 'department', 'type',
    'status', 'group', 'segment', 'class'
  ];
  
  const geoKeywords = [
    'latitude', 'lat', 'latitud', 'ycoord', 'y_coord',
    'longitude', 'long', 'lon', 'longitud', 'xcoord', 'x_coord',
    'location', 'geo', 'geography', 'coordinates', 'coords',
    'postal', 'zip', 'zipcode', 'postcode', 'address'
  ];
  
  for (const keyword of geoKeywords) {
    if (lowerColName.includes(keyword)) {
      console.log(`🗺️ Skipping conditional formatting for geographic column: "${columnName}"`);
      return true;
    }
  }
  
  for (const keyword of categoricalKeywords) {
    if (lowerColName.includes(keyword)) {
      return true;
    }
  }
  
  if (value !== undefined && value !== null) {
    const strValue = String(value).trim();
    const numValue = parseFloat(strValue);
    
    if (!isNaN(numValue)) {
      if (lowerColName.includes('lat') && numValue >= -90 && numValue <= 90) {
        console.log(`🗺️ Skipping conditional formatting for latitude column: "${columnName}" (value: ${numValue})`);
        return true;
      }
      
      if (lowerColName.includes('lon') && numValue >= -180 && numValue <= 180) {
        console.log(`🗺️ Skipping conditional formatting for longitude column: "${columnName}" (value: ${numValue})`);
        return true;
      }
      
      if (lowerColName.match(/lat/i) && numValue >= -90 && numValue <= 90) {
        console.log(`🗺️ Skipping conditional formatting for geographic coordinate column: "${columnName}"`);
        return true;
      }
      
      if (lowerColName.match(/lon/i) && numValue >= -180 && numValue <= 180) {
        console.log(`🗺️ Skipping conditional formatting for geographic coordinate column: "${columnName}"`);
        return true;
      }
    }
    
    if (/^\d{4}$/.test(strValue)) {
      const yearNum = parseInt(strValue, 10);
      if (yearNum >= 1900 && yearNum <= 2100) {
        return true;
      }
    }
  }
  
  return false;
}

export function isGeographicCoordinate(value, columnName) {
  if (!isNumeric(value)) return false;
  
  const lowerColName = columnName.toLowerCase();
  const numValue = parseFloat(String(value).trim());
  
  if (lowerColName.includes('latitude') || lowerColName.includes('lat')) {
    return numValue >= -90 && numValue <= 90;
  }
  
  if (lowerColName.includes('longitude') || lowerColName.includes('lon')) {
    return numValue >= -180 && numValue <= 180;
  }
  
  if (numValue >= -90 && numValue <= 90 && (lowerColName.includes('coord') || lowerColName.includes('geo'))) {
    return true;
  }
  
  if ((numValue >= -180 && numValue <= -90) || (numValue >= 90 && numValue <= 180)) {
    if (lowerColName.includes('coord') || lowerColName.includes('geo')) {
      return true;
    }
  }
  
  return false;
}

export function isNumericForFormatting(value, columnName) {
  if (!isNumeric(value)) return false;
  
  if (columnName) {
    const strValue = String(value).trim();
    if (/^\d{4}$/.test(strValue)) {
      const yearNum = parseInt(strValue, 10);
      if (yearNum >= 1900 && yearNum <= 2100) {
        return false;
      }
    }
    
    if (isGeographicCoordinate(value, columnName)) {
      return false;
    }
  }
  
  return true;
}

export function getZeroCenteredColor(value, maxAbs) {
  if (maxAbs === 0) {
    return "FFFFFF";
  }
  
  const intensity = Math.min(Math.abs(value) / maxAbs, 1);
  
  let r = 255;
  let g = 255;
  let b = 0;
  
  if (value > 0) {
    r = Math.round(255 * (1 - intensity));
    g = 255;
    b = Math.round(100 * (1 - intensity));
  } else if (value < 0) {
    r = 255;
    g = Math.round(255 * (1 - intensity));
    b = Math.round(100 * (1 - intensity));
  }
  
  return (
    "FF" +
    r.toString(16).padStart(2, "0") +
    g.toString(16).padStart(2, "0") +
    b.toString(16).padStart(2, "0")
  ).toUpperCase();
}

/* ── FALLBACK: Hardcoded conditional formatting (only used if no XML colors) ── */
export function applyConditionalFormattingToTable(worksheet, rows, headers, startRow, startCol, colWidths) {
  headers.forEach((header, colIndex) => {
    const columnName = header.fieldName || header.fieldId || `Column_${colIndex + 1}`;
    const firstValue = rows[0]?.[colIndex]?.formattedValue || rows[0]?.[colIndex]?.value;
    
    let shouldSkip = shouldSkipConditionalFormatting(columnName, firstValue);
    
    if (shouldSkip) {
      console.log(`⚠️ Skipping conditional formatting for column: "${columnName}"`);
      return;
    }
    
    const numericValues = [];
    
    rows.forEach(row => {
      const value = row[colIndex]?.formattedValue || row[colIndex]?.value;
      if (isNumericForFormatting(value, columnName)) {
        let numValue = String(value).replace(/,/g, "").replace(/%/g, "");
        numericValues.push(parseFloat(numValue));
      }
    });
    
    if (numericValues.length === 0) return;
    
    const maxAbs = Math.max(...numericValues.map(v => Math.abs(v)));
    
    rows.forEach((row, rowIndex) => {
      const rawValue = row[colIndex]?.formattedValue || row[colIndex]?.value;
      
      if (!isNumericForFormatting(rawValue, columnName)) return;
      
      let numValue = String(rawValue).replace(/,/g, "").replace(/%/g, "");
      const value = parseFloat(numValue);
      
      const excelRow = startRow + rowIndex;
      const excelCol = startCol + colIndex;
      const cell = worksheet.getCell(excelRow + 1, excelCol + 1);
      
      const color = getZeroCenteredColor(value, maxAbs);
      
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color }
      };
    });
  });
}
