// Office.onReady((info) => {
//   if (info.host === Office.HostType.Excel) {
//     document.getElementById("run").onclick = fastCWAHU_FromMappingSheet;
//     const s = document.getElementById("status");
//     if (s) s.textContent = "";
//   }
//   document.getElementById("selectHeader").onclick = async () => {
//     await Excel.run(async (ctx) => {
//       const range = ctx.workbook.getSelectedRange();
//       range.name = "HeaderRange";
//       await ctx.sync();
//       updateStatus("✅ HeaderRange defined from selected cells.");
//     });
//   };
// });

// async function fastCWAHU_FromMappingSheet() {
//   try {
//     const files = await pickRtfFiles();
//     if (!files.length) return;

//     await Excel.run(async (ctx) => {
//       const ws = ctx.workbook.worksheets.getActiveWorksheet();
//       const { ws: mapWS, wasCreated } = await ensureMappingSheetExists(ctx);
//       if (wasCreated) updateStatus("✅ Mapping sheet was created.");

//       const headers = await buildHeaderMap(ws, ctx);
//       const mDict = await loadMappingDict(mapWS, ctx);

//       let nextRow = 6;
//       let tempRow = {};
//       let curSection = "";

//       for (const file of files) {
//         const text = await rtfToPlain(file);
//         const lines = text.replace(/\r\n?/g, "\n").split("\n");

//         for (const ln of lines) {
//           const txt = ln.trim();
//           if (!txt) continue;

//           if (/Sizing Data|Cooling Coil|Outdoor Ventilation/i.test(txt)) {
//             curSection = txt.toUpperCase();
//           }

//           const unitHeader = mDict["AIR SYSTEM NAME"]?.[0]?.[0] ?? "";
//           if (unitHeader && /Air System Name/i.test(txt)) {
//             if (Object.keys(tempRow).length > 1) {
//               writeRow(ws, headers, ++nextRow, tempRow);
//             }
//             tempRow = {};
//             tempRow[unitHeader] = extractAfter(txt, "Air System Name");
//             continue;
//           }

//           const hits = matchMappedTermsBySection(txt, curSection, mDict, headers);
//           Object.assign(tempRow, hits);
//         }
//       }

//       if (Object.keys(tempRow).length > 1) {
//         writeRow(ws, headers, ++nextRow, tempRow);
//       }

//       await ctx.sync();
//       updateStatus("✅ Import complete.");
//     });
//   } catch (err) {
//     console.error(err);
//     updateStatus("❌ " + err.message);
//   }
// }

// function pickRtfFiles() {
//   return new Promise((resolve) => {
//     const input = document.createElement("input");
//     input.type = "file";
//     input.accept = ".rtf";
//     input.multiple = true;
//     input.style.display = "none";
//     document.body.appendChild(input);
//     input.onchange = () => {
//       resolve([...input.files]);
//       document.body.removeChild(input);
//     };
//     input.click();
//   });
// }

// function updateStatus(msg) {
//   const el = document.getElementById("status");
//   if (el) el.textContent = msg;
// }

// function rtfToPlain(file) {
//   return new Promise((res, rej) => {
//     const reader = new FileReader();
//     reader.onload = () => {
//       const rtf = reader.result;
//       const txt = rtf
//         .replace(/\\par[d]?/g, "\n")
//         .replace(/\\'[0-9a-f]{2}/gi, (m) => String.fromCharCode(parseInt(m.substr(2), 16)))
//         .replace(/\\[a-z]+\d* ?/gi, "")
//         .replace(/[{}]/g, "");
//       res(txt);
//     };
//     reader.onerror = () => rej(reader.error);
//     reader.readAsText(file);
//   });
// }

// async function ensureMappingSheetExists(ctx) {
//   const sheets = ctx.workbook.worksheets;
//   let ws = sheets.getItemOrNullObject("Mapping");
//   await ctx.sync();

//   let wasCreated = false;
//   if (ws.isNullObject) {
//     ws = sheets.add("Mapping");
//     wasCreated = true;
//     await ctx.sync();

//     ws.getRange("A1:B1").merge(true);
//     ws.getRange("A1").values = [["Schedule header"]];
//     ws.getRange("C1:D1").merge(true);
//     ws.getRange("C1").values = [["Import Data"]];
//     ws.getRange("A2:D2").values = [["Section", "Column", "Section", "HAP term"]];

//     const s = [
//       ["", "", "Supply Fan Sizing Data", "Actual max L/s"],
//       ["", "", "Supply Fan Sizing Data", "Fan static"],
//       ["", "", "Supply Fan Sizing Data", "Fan motor kW"],
//       ["", "", "", ""],
//       ["", "", "Return Fan Sizing Data", "Actual max L/s"],
//       ["", "", "Return Fan Sizing Data", "Fan motor kW"],
//       ["", "", "Return Fan Sizing Data", "Fan static"],
//       ["", "", "", ""],
//       ["", "", "Outdoor Ventilation Air Data", "Design airflow L/s"],
//       ["", "", "", ""],
//       ["", "", "Central Cooling Coil Sizing Data", "Total coil load"],
//       ["", "", "Central Cooling Coil Sizing Data", "Sensible coil load"],
//       ["", "", "", ""],
//       ["", "", "Air System Information", "Air System Name"]
//     ];
//     ws.getRange("A3:D17").values = s;

//     const cols = ws.getRange("A:D");
//     cols.columnWidth = 140;
//     cols.format.horizontalAlignment = "Center";
//     cols.format.verticalAlignment = "Center";
//     const block = ws.getRange("A1:D17");
//     block.format.font.name = "Calibri";
//     block.format.font.size = 12;
//     block.getCell(0, 0).getResizedRange(1, 3).format.font.bold = true;
//     ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"].forEach(
//       (b) => (block.format.borders.getItem(b).style = "Continuous")
//     );

//     await ctx.sync();
//   }
//   return { ws, wasCreated };
// }

// async function buildHeaderMap(ws, ctx) {
//   let rangeObj = ws.names.getItemOrNullObject("HeaderRange");
//   await ctx.sync();

//   if (rangeObj.isNullObject) {
//     const selectedRange = ctx.workbook.getSelectedRange();
//     selectedRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "values"]);
//     await ctx.sync();

//     ws.names.add("HeaderRange", selectedRange);
//     rangeObj = selectedRange;
//   } else {
//     rangeObj = rangeObj.getRange();
//     rangeObj.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "values"]);
//     await ctx.sync();
//   }

//   const { rowCount, columnCount, columnIndex, values } = rangeObj;
//   const map = {};

//   for (let c = 0; c < columnCount; c++) {
//     let top = "", bottom = "";
//     for (let r = 0; r < rowCount && !top; r++) top = (values[r][c] || "").toString().trim();
//     for (let r = rowCount - 1; r >= 0 && !bottom; r--) bottom = (values[r][c] || "").toString().trim();

//     const key = top.toUpperCase() === bottom.toUpperCase()
//       ? top
//       : top && bottom
//         ? `${top}|${bottom}`
//         : top || `|${bottom}`;

//     if (key && !map[key.toUpperCase()]) {
//       map[key.toUpperCase()] = columnIndex + c + 1;
//     }
//   }

//   return map;
// }


// async function loadMappingDict(sheet, ctx) {
//   const rng = sheet.getUsedRangeOrNullObject();
//   await ctx.sync();
//   if (rng.isNullObject) return {};

//   rng.load("values");
//   await ctx.sync();

//   const dict = {};
//   rng.values.slice(2).forEach((row) => {
//     const [schedA, schedB, section, hapTerm] = row.map((v) => (v || "").toString());
//     if (!schedB || !hapTerm) return;
//     const schedHeader = `${schedA ? schedA + "|" : ""}${schedB}`.toUpperCase();
//     (dict[hapTerm.toUpperCase()] = dict[hapTerm.toUpperCase()] || []).push([schedHeader, section.toUpperCase()]);
//   });
//   return dict;
// }

// function matchMappedTermsBySection(txt, currentSection, mappingDict, headerMap) {
//   const res = {};
//   const clean = cleanControlChars(txt).trim();

//   for (const hapKey of Object.keys(mappingDict)) {
//     if (clean.toLowerCase().startsWith(hapKey.toLowerCase())) {
//       const nextChar = clean.charAt(hapKey.length) || " ";
//       if (" .:-/".includes(nextChar)) {
//         const extracted = extractNumber(clean);
//         for (const [schedHeader, reqSection] of mappingDict[hapKey]) {
//           if (!reqSection || currentSection.includes(reqSection)) {
//             if (headerMap[schedHeader.toUpperCase()]) {
//               res[schedHeader] = schedHeader.includes("POWER")
//                 ? stdPower(parseFloat(extracted))
//                 : extracted;
//             }
//           }
//         }
//       }
//     }
//   }
//   return res;
// }

// function cleanControlChars(str = "") {
//   return [...str].filter((ch) => ch.charCodeAt(0) >= 32 && ch.charCodeAt(0) <= 126).join("");
// }

// function extractNumber(input = "") {
//   let out = "";
//   for (let i = input.length - 1; i >= 0; i--) {
//     const ch = input[i];
//     if (/[0-9\-\./\s]/.test(ch)) out = ch + out;
//     else break;
//   }
//   return out.trim().replace(/^\/|\/$/g, "");
// }

// function stdPower(val) {
//   const table = [0, 0.09, 0.19, 0.38, 0.56, 0.75, 1.13, 1.5, 2.25, 3.75, 5.6, 7.5, 11.3, 15, 18.8, 22.5, 30, 37.5, 45, 56.3, 75, 93, 113.5, 150];
//   for (const p of table) if (p >= val) return p;
//   return val;
// }

// function extractAfter(txt, needle) {
//   return txt.split(new RegExp(needle, "i"))[1]?.trim() || "";
// }

// function writeRow(ws, headerMap, row, data) {
//   Object.entries(data).forEach(([hdr, val]) => {
//     const col = headerMap[hdr.toUpperCase()];
//     if (col) ws.getCell(row - 1, col - 1).values = [[val]];
//   });
// }






import {
  logoutRequestLocal
} from "../firebase-auth.js";

// Global variables to store header range and RTF content
let headerRange = null; // Stores the selected Excel header range
let rtfTextContent = null; // Stores the parsed content from uploaded RTF file

// Initializes the add-in when Office is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Set up event handlers for UI buttons
    document.getElementById("selectHeader").onclick = selectHeader;
    document.getElementById("uploadRtf").onchange = handleRtfUpload;
    document.getElementById("importRtf").onclick = importRtfToExcel;
    document.getElementById("requestLogout").onclick = requestLogout;
  }
});

/**
 * Displays a message in the UI message box and logs to console
 * @param {string} msg - The message to display
 */
function showMessage(msg) {
  const el = document.getElementById("messageBox");
  if (el) el.textContent = msg;
  console.log("MESSAGE:", msg);
}

/**
 * Deeply trims a string by removing all whitespace and special Unicode characters
 * @param {string} str - The string to clean
 * @returns {string} The cleaned string in uppercase
 */
function deepTrim(str = "") {
  return str.replace(/\s+/g, "").replace(/[\u200B-\u200D\uFEFF]/g, "").toUpperCase();
}

/**
 * Handles selecting a header range in Excel
 */
async function selectHeader() {
  try {
    await Excel.run(async (ctx) => {
      // Get the currently selected range in Excel
      const range = ctx.workbook.getSelectedRange();
      // Load range properties we need
      range.load(["address", "values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
      ctx.trackedObjects.add(range);
      await ctx.sync();

      // Store the selected range globally
      headerRange = range;
      range.name = "HeaderRange";
      await ctx.sync();

      showMessage(`✅ Header selected: ${range.address}`);
    });
  } catch (error) {
    console.error("Error selecting header:", error);
    showMessage("❌ Failed to define HeaderRange.");
  }
}

/**
 * Handles uploading and parsing an RTF file
 * @param {Event} event - The file upload event
 */
function handleRtfUpload(event) {
  const file = event.target.files[0];
  const reader = new FileReader();
  
  reader.onload = (e) => {
    const content = e.target.result;
    // Parse RTF content by:
    // 1. Replacing RTF commands with newlines
    // 2. Converting hex codes to characters
    // 3. Removing remaining RTF syntax
    // 4. Splitting into lines and cleaning each line
    rtfTextContent = content
      .replace(/\\pard?/g, "\n")
      .replace(/\\'[0-9a-fA-F]{2}/g, (m) => String.fromCharCode(parseInt(m.slice(2), 16)))
      .replace(/\\[^ ]+ ?|[{}]/g, "")
      .split("\n")
      .map((line) => line.trim())
      .filter((line) => line.length > 0);
      
    showMessage("✅ RTF file uploaded.");
    console.log("RTF Lines:", rtfTextContent);
  };
  reader.readAsText(file);
}

/**
 * Imports the parsed RTF content into Excel based on the selected header
 */
async function importRtfToExcel() {
  if (!headerRange || !rtfTextContent) {
    showMessage("⚠️ Please select header and upload RTF file first.");
    return;
  }

  await Excel.run(async (context) => {
    context.trackedObjects.add(headerRange);
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const mapWS = context.workbook.worksheets.getItem("Mapping");
    const usedRange = mapWS.getUsedRange();
    usedRange.load("values");
    headerRange.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();

    // Clear existing data below the header
    const startClearRow = headerRange.rowIndex + headerRange.rowCount;
    const clearRange = sheet.getRangeByIndexes(startClearRow, headerRange.columnIndex, 1000, headerRange.columnCount);
    clearRange.clear(Excel.ClearApplyTo.contents);

    // Build data structures for mapping
    const headerMap = buildHeaderMap(headerRange);
    const mappingDict = buildMappingDict(usedRange.values);

    // Process RTF content
    let currentSection = "";
    let iRow = headerRange.rowIndex + headerRange.rowCount;
    let tempRow = {};
    const writtenUnits = {};
    const unitHeader = getUnitHeader(mappingDict);

    for (let i = 0; i < rtfTextContent.length; i++) {
      const txt = rtfTextContent[i].trim();
      if (!txt) continue;

      // Detect section headers
      if (/Sizing Data|Cooling Coil|Outdoor Ventilation|Air System Information/i.test(txt)) {
        currentSection = txt.toUpperCase();
      }

      // Handle system name lines
      if (/Air System Name/i.test(txt)) {
        const systemName = extractValueFromLine(txt);
        if (Object.keys(tempRow).length > 1 && tempRow[unitHeader] && !writtenUnits[tempRow[unitHeader]]) {
          if (writeRow(sheet, headerMap, iRow, tempRow)) iRow++;
          writtenUnits[tempRow[unitHeader]] = true;
        }
        tempRow = {};
        if (unitHeader) tempRow[unitHeader] = systemName;
        continue;
      }

      // Match terms from the current section
      const matches = matchMappedTermsBySection(txt, currentSection, mappingDict, headerMap);
      if (Object.keys(matches).length > 0) {
        Object.assign(tempRow, matches);
      }
    }

    // Write any remaining data
    if (Object.keys(tempRow).length > 1 && tempRow[unitHeader] && !writtenUnits[tempRow[unitHeader]]) {
      if (writeRow(sheet, headerMap, iRow, tempRow)) iRow++;
    }

    await context.sync();
    showMessage("✅ RTF import complete.");
  });
}

/**
 * Gets the header key for unit names from the mapping dictionary
 * @param {Object} mappingDict - The mapping dictionary
 * @returns {string} The unit header key
 */
function getUnitHeader(mappingDict) {
  if (mappingDict["AIR SYSTEM NAME"]) {
    const pair = mappingDict["AIR SYSTEM NAME"][0];
    return pair[0];
  }
  return "";
}

/**
 * Extracts values from a line of text
 * @param {string} text - The text to parse
 * @returns {string} The extracted value
 */
function extractValueFromLine(text) {
  // First try to match number pairs (e.g., "1.5 / 2.0")
  const pairMatch = text.match(/\d+\.\d+\s*\/\s*\d+\.\d+/);
  if (pairMatch) return pairMatch[0];

  // Otherwise split on "Air System Name" and return the second part
  const parts = text.split(/Air System Name/i);
  return parts[1] ? parts[1].trim() : "";
}

/**
 * Builds a mapping between header names and column indices
 * @param {Object} headerRange - The Excel range containing headers
 * @returns {Object} A dictionary mapping header keys to column indices
 */
function buildHeaderMap(headerRange) {
  const headerMap = {};
  const values = headerRange.values;
  const rowCount = headerRange.rowCount;
  const colCount = headerRange.columnCount;
  
  // For each column, get the top and bottom header values
  for (let c = 0; c < colCount; c++) {
    let top = "", bottom = "";
    for (let r = 0; r < rowCount; r++) if (values[r][c]) { top = values[r][c]; break; }
    for (let r = rowCount - 1; r >= 0; r--) if (values[r][c]) { bottom = values[r][c]; break; }
    
    // Create a compound key if top and bottom are different
    let key = (top && bottom && top !== bottom) ? `${top}|${bottom}` : (top || bottom || "");
    key = key.toUpperCase().replace(/\s+/g, "");
    if (key) headerMap[key] = headerRange.columnIndex + c;
  }
  return headerMap;
}

/**
 * Builds a mapping dictionary from the Mapping worksheet
 * @param {Array} data - The values from the Mapping worksheet
 * @returns {Object} A dictionary mapping terms to header/section pairs
 */
function buildMappingDict(data) {
  const mappingDict = {};
  // Skip header rows and process each mapping row
  data.slice(2).forEach((row) => {
    // Create a compound key from the first two columns
    const schedHeader = ((row[0] ? row[0] + "|" : "") + row[1]).toUpperCase().replace(/\s+/g, "");
    const hapTerm = String(row[3] || "").toUpperCase().trim();
    const section = String(row[2] || "").toUpperCase().trim();
    
    // Add to dictionary
    if (!mappingDict[hapTerm]) mappingDict[hapTerm] = [];
    mappingDict[hapTerm].push([schedHeader, section]);
  });
  return mappingDict;
}

/**
 * Matches terms from the RTF content to mapped terms in the current section
 * @param {string} txt - The text to match
 * @param {string} currentSection - The current section name
 * @param {Object} mappingDict - The mapping dictionary
 * @param {Object} headerMap - The header mapping
 * @returns {Object} Matched values with their corresponding headers
 */
function matchMappedTermsBySection(txt, currentSection, mappingDict, headerMap) {
  const res = {};
  const clean = txt.toLowerCase();
  const sectionClean = deepTrim(currentSection);
  
  // Check each term in the mapping dictionary
  for (const hapKey of Object.keys(mappingDict)) {
    const hapKeyLower = hapKey.toLowerCase().trim();
    if (clean.includes(hapKeyLower)) {
      // Extract the numeric value associated with this term
      const extracted = extractNumberFromContext(txt, hapKey);
      
      // Check all mappings for this term
      for (const [schedHeader, reqSection] of mappingDict[hapKey]) {
        const reqSectionClean = deepTrim(reqSection);
        // If section matches (or no section specified)
        if (!reqSection || sectionClean.includes(reqSectionClean)) {
          if (headerMap[schedHeader] !== undefined) {
            // Standardize power values if needed
            res[schedHeader] = schedHeader.includes("POWER") ? stdPower(parseFloat(extracted)) : extracted;
          }
        }
      }
    }
  }
  return res;
}

/**
 * Extracts a number from text following a keyword
 * @param {string} text - The text to search
 * @param {string} keyword - The keyword to find
 * @returns {string} The extracted number or empty string if none found
 */
function extractNumberFromContext(text, keyword) {
  const index = text.toLowerCase().indexOf(keyword.toLowerCase());
  if (index === -1) return "";
  const after = text.slice(index + keyword.length).trim();

  // Try to match number pairs first
  const pairMatch = after.match(/\d+\.\d+\s*\/\s*\d+\.\d+/);
  if (pairMatch) return pairMatch[0];

  // Then try single numbers
  const numMatch = after.match(/^([-+]?[0-9]*\.?[0-9]+)/);
  if (numMatch) return numMatch[1];

  // Fallback to last number in the text
  const allNums = text.match(/([-+]?[0-9]*\.?[0-9]+)/g);
  return allNums ? allNums[allNums.length - 1] : "";
}

/**
 * Standardizes power values to predefined levels
 * @param {number} value - The power value to standardize
 * @returns {number} The standardized power value
 */
function stdPower(value) {
  const powers = [0, 0.09, 0.19, 0.38, 0.56, 0.75, 1.13, 1.5, 2.25, 3.75, 5.6, 7.5, 11.3, 15, 18.8, 22.5, 30, 37.5, 45, 56.3, 75, 93, 113.5, 150];
  for (const p of powers) if (value <= p) return p;
  return value;
}

/**
 * Writes a row of data to Excel
 * @param {Object} sheet - The Excel worksheet
 * @param {Object} headerMap - The header mapping
 * @param {number} row - The row index to write to
 * @param {Object} data - The data to write
 * @returns {boolean} True if data was written, false if row was empty
 */
function writeRow(sheet, headerMap, row, data) {
  const startCol = headerRange.columnIndex;
  const totalCols = Object.keys(headerMap).length;
  const rowValues = new Array(totalCols).fill("");
  let hasNonEmptyValue = false;

  // Map data to columns based on headerMap
  for (const key in data) {
    const normKey = key.toUpperCase().replace(/\s+/g, '');
    const colIndex = headerMap[normKey];
    if (colIndex !== undefined) {
      const relativeIndex = colIndex - startCol;
      const value = data[key];
      if (value !== "" && value !== null && value !== undefined) {
        rowValues[relativeIndex] = value;
        hasNonEmptyValue = true;
      }
    }
  }

  // Only write if there's data
  if (hasNonEmptyValue) {
    const range = sheet.getRangeByIndexes(row, startCol, 1, rowValues.length);
    range.values = [rowValues];
    console.log(`✅ Writing to row ${row}:`, rowValues);
    return true;
  } else {
    console.log(`⛔ Skipped empty row ${row}`);
    return false;
  }
}

/**
 * Handles logout request by opening email client and cleaning up local state
 */
async function requestLogout() {
  console.log("requestLogout function called.");
  const email = localStorage.getItem("email") || "Unknown User";
  const subject = encodeURIComponent("Logout Request");
  const body = encodeURIComponent(`${email} requests logout from Excel Data Transfer Add‑in.`);
  // Open email client with pre-filled message
  window.open(`mailto:aecoresolutions@gmail.com?subject=${subject}&body=${body}`, "_blank");
  // Clean up local authentication state
  await logoutRequestLocal();
  console.log("logoutRequestLocal completed.");
}
