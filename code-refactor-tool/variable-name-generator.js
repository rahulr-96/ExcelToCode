const XLSX = require("xlsx");
const FileSystem = require("fs");

/**
 * Converts a string to a sanitized camelCase variable name
 */
function sanitizeVariableName(str) {
  if (!str) return "";
  // Remove all non-alphanumeric characters except space and underscore
  str = str.replace(/[^a-zA-Z0-9 _ |]/g, "");
  // Split by space or underscore, lowercase first word, capitalize the rest
  const parts = str.split(/[\s_]+/);
  return parts
    .map((word, i) => (i === 0 ? word.toLowerCase() : word[0].toUpperCase() + word.slice(1).toLowerCase()))
    .join("");
}

/**
 * Finds the nearest header above (row) and to the left (column) for a given cell.
 * @param {Object} sheet - XLSX sheet object
 * @param {Number} row - 1-based row index
 * @param {Number} col - 1-based column index
 * @param {Array} merges - list of merged ranges in the sheet
 */
function findNearestHeader(sheet, row, col, merges) {
  // Helper to check if cell is inside a merged range
  function getMergedValue(r, c) {
    for (const merge of merges) {
      if (
        r >= merge.s.r + 1 &&
        r <= merge.e.r + 1 &&
        c >= merge.s.c + 1 &&
        c <= merge.e.c + 1
      ) {
        const topLeft = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
        return sheet[topLeft] ? sheet[topLeft].v : null;
      }
    }
    return null;
  }

  // Look upwards for row header
  let headerRow = null;
  for (let r = row - 1; r >= 1; r--) {
    const cellAddress = XLSX.utils.encode_cell({ r: r - 1, c: col - 1 });
    const value = (sheet[cellAddress] ? sheet[cellAddress].v : null) || getMergedValue(r, col);
    if (typeof value === "string" && value.trim()) {
      headerRow = value.trim();
      break;
    }
  }

  // Look leftwards for column header
  let headerCol = null;
  for (let c = col - 1; c >= 1; c--) {
    const cellAddress = XLSX.utils.encode_cell({ r: row - 1, c: c - 1 });
    const value = (sheet[cellAddress] ? sheet[cellAddress].v : null) || getMergedValue(row, c);
    if (typeof value === "string" && value.trim()) {
      headerCol = value.trim();
      break;
    }
  }


  let rawName = "";
  if (headerRow && headerCol) rawName = `${headerRow}_${headerCol}`;
  else if (headerRow) rawName = headerRow;
  else if (headerCol) rawName = headerCol;
  else rawName = `R${row}C${col}`;

  return sanitizeVariableName(rawName);
}

/**
 * Reads an Excel file and returns a mapping of Sheet_Cell -> variableName
 */
function excelToVariableMap(filePath) {
  const workbook = XLSX.readFile(filePath, { cellFormula: true });
  const mapping = {};

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const merges = sheet["!merges"] || [];
    const range = XLSX.utils.decode_range(sheet["!ref"]);

    for (let r = range.s.r + 1; r <= range.e.r + 1; r++) {
      for (let c = range.s.c + 1; c <= range.e.c + 1; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
        const cell = sheet[cellAddress];
        if (!cell) continue; // skip empty cells
        if (cell.t === "s") continue; //skip string cells
        const key = `${sheetName}_${cellAddress}`;
        mapping[key] = findNearestHeader(sheet, r, c, merges);
      }
    }
  });

  return mapping;
}

function main() {
  const args = process.argv.slice(2);

  if (args.length < 2) {
    console.error("Usage: node variable-name-generator.js <input_cs> <output_json>");
    process.exit(1);
  }

  const inputFile = args[0];
  const outputFile = args[1];

  // Generate the variable map
  const variableMap = excelToVariableMap(inputFile);

  // Save as JSON
  FileSystem.writeFile(outputFile, JSON.stringify(variableMap, null, 2), (err) => {
    if (err) throw err;
    console.log(`Variable map saved to ${outputFile}`);
  });
}

main();