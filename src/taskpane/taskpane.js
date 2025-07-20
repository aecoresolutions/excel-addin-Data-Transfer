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



let headerRange = null;
let rtfTextContent = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("selectHeader").onclick = selectHeader;
    document.getElementById("uploadRtf").onchange = handleRtfUpload;
    document.getElementById("importRtf").onclick = importRtfToExcel;
  }
});

   function showMessage(msg) {
  const el = document.getElementById("messageBox");
  if (el) el.textContent = msg;
  console.log("MESSAGE:", msg);
}

function deepTrim(str = "") {
  return str.replace(/\s+/g, "").replace(/[\u200B-\u200D\uFEFF]/g, "").toUpperCase();
}

async function selectHeader() {
  try {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["address", "values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
      ctx.trackedObjects.add(range);
      await ctx.sync();

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

function handleRtfUpload(event) {
  const file = event.target.files[0];
  const reader = new FileReader();
  reader.onload = (e) => {
    const content = e.target.result;
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

    const startClearRow = headerRange.rowIndex + headerRange.rowCount;
    const clearRange = sheet.getRangeByIndexes(startClearRow, headerRange.columnIndex, 1000, headerRange.columnCount);
    clearRange.clear(Excel.ClearApplyTo.contents);

    const headerMap = buildHeaderMap(headerRange);
    const mappingDict = buildMappingDict(usedRange.values);

    let currentSection = "";
    let iRow = headerRange.rowIndex + headerRange.rowCount;
    let tempRow = {};
    const writtenUnits = {};
    const unitHeader = getUnitHeader(mappingDict);

    for (let i = 0; i < rtfTextContent.length; i++) {
      const txt = rtfTextContent[i].trim();
      if (!txt) continue;

      if (/Sizing Data|Cooling Coil|Outdoor Ventilation|Air System Information/i.test(txt)) {
        currentSection = txt.toUpperCase();
      }

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

      const matches = matchMappedTermsBySection(txt, currentSection, mappingDict, headerMap);
      if (Object.keys(matches).length > 0) {
        Object.assign(tempRow, matches);
      }
    }

    if (Object.keys(tempRow).length > 1 && tempRow[unitHeader] && !writtenUnits[tempRow[unitHeader]]) {
      if (writeRow(sheet, headerMap, iRow, tempRow)) iRow++;
    }

    await context.sync();
    showMessage("✅ RTF import complete.");
  });
}

function getUnitHeader(mappingDict) {
  if (mappingDict["AIR SYSTEM NAME"]) {
    const pair = mappingDict["AIR SYSTEM NAME"][0];
    return pair[0];
  }
  return "";
}

function extractValueFromLine(text) {
  const pairMatch = text.match(/\d+\.\d+\s*\/\s*\d+\.\d+/);
  if (pairMatch) return pairMatch[0];

  const parts = text.split(/Air System Name/i);
  return parts[1] ? parts[1].trim() : "";
}

function buildHeaderMap(headerRange) {
  const headerMap = {};
  const values = headerRange.values;
  const rowCount = headerRange.rowCount;
  const colCount = headerRange.columnCount;
  for (let c = 0; c < colCount; c++) {
    let top = "", bottom = "";
    for (let r = 0; r < rowCount; r++) if (values[r][c]) { top = values[r][c]; break; }
    for (let r = rowCount - 1; r >= 0; r--) if (values[r][c]) { bottom = values[r][c]; break; }
    let key = (top && bottom && top !== bottom) ? `${top}|${bottom}` : (top || bottom || "");
    key = key.toUpperCase().replace(/\s+/g, "");
    if (key) headerMap[key] = headerRange.columnIndex + c;
  }
  return headerMap;
}

function buildMappingDict(data) {
  const mappingDict = {};
  data.slice(2).forEach((row) => {
    const schedHeader = ((row[0] ? row[0] + "|" : "") + row[1]).toUpperCase().replace(/\s+/g, "");
    const hapTerm = String(row[3] || "").toUpperCase().trim();
    const section = String(row[2] || "").toUpperCase().trim();
    if (!mappingDict[hapTerm]) mappingDict[hapTerm] = [];
    mappingDict[hapTerm].push([schedHeader, section]);
  });
  return mappingDict;
}

function matchMappedTermsBySection(txt, currentSection, mappingDict, headerMap) {
  const res = {};
  const sectionClean = deepTrim(currentSection);
  const txtLower = txt.toLowerCase();

  console.log("➡️ New Line:", txt);
  console.log("📍 Current Section:", sectionClean);

  for (const hapKey of Object.keys(mappingDict)) {
    const hapKeyLower = hapKey.toLowerCase().trim();

    const exactMatchPattern = new RegExp(`^${hapKeyLower.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')}(\\s|\\t)`, 'i');

    if (exactMatchPattern.test(txtLower)) {
      console.log(`✅ Match Found for key: "${hapKey}"`);

      const extracted = extractNumberFromContext(txt, hapKey);
      console.log(`🔢 Extracted Value: "${extracted}"`);

      for (const [schedHeader, reqSection] of mappingDict[hapKey]) {
        const reqSectionClean = deepTrim(reqSection || "");
        console.log(`🧭 Mapping to: "${schedHeader}" in section: "${reqSectionClean}"`);

        if (!reqSection || sectionClean.includes(reqSectionClean)) {
          if (headerMap[schedHeader] !== undefined) {
            res[schedHeader] = schedHeader.includes("POWER")
              ? stdPower(parseFloat(extracted))
              : extracted;

            console.log(`📤 Stored in res: ${schedHeader} = ${res[schedHeader]}`);
          } else {
            console.log(`⚠️ "${schedHeader}" not found in headerMap`);
          }
        } else {
          console.log(`🚫 Section mismatch: required "${reqSectionClean}", current "${sectionClean}"`);
        }
      }
    }
  }

  return res;
}









function extractNumberFromContext(text, keyword) {
  const index = text.toLowerCase().indexOf(keyword.toLowerCase());
  if (index === -1) return "";
  const after = text.slice(index + keyword.length).trim();

  const pairMatch = after.match(/\d+\.\d+\s*\/\s*\d+\.\d+/);
  if (pairMatch) return pairMatch[0];

  const valueMatch = after.match(/^[-+]?\d*\.?\d+/);
  if (valueMatch) return valueMatch[0];

  const allNums = text.match(/([-+]?[0-9]*\.?[0-9]+)/g);
  return allNums ? allNums[allNums.length - 1] : "";
}

function stdPower(value) {
  const powers = [0, 0.09, 0.19, 0.38, 0.56, 0.75, 1.13, 1.5, 2.25, 3.75, 5.6, 7.5, 11.3, 15, 18.8, 22.5, 30, 37.5, 45, 56.3, 75, 93, 113.5, 150];
  for (const p of powers) if (value <= p) return p;
  return value;
}

function writeRow(sheet, headerMap, row, data) {
  const startCol = headerRange.columnIndex;
  const totalCols = Object.keys(headerMap).length;
  const rowValues = new Array(totalCols).fill("");
  let hasNonEmptyValue = false;

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

/* ─── Request Logout (opens mail client) ─── */
async function requestLogout() {
  console.log("requestLogout function called.");
  const email = localStorage.getItem("email") || "Unknown User";
  const subject = encodeURIComponent("Logout Request");
  const body = encodeURIComponent(`${email} requests logout from Excel Data Transfer Add‑in.`);
  // window.location.href = `mailto:aecoresolutions@gmail.com?subject=${subject}&body=${body}`;
  window.open(`mailto:aecoresolutions@gmail.com?subject=${subject}&body=${body}`, "_blank");
  /* local clean‑up */
  // logoutRequestLocal depends on Firebase. If Firebase is not initialized, this won't work.
  await logoutRequestLocal();
  console.log("logoutRequestLocal completed.");
}
