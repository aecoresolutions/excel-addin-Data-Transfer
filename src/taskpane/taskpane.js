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



// When the Office Add-in is ready and hosted in Excel
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // When the user clicks the "run" button, start the main import process
    document.getElementById("run").onclick = fastCWAHU_FromMappingSheet;

    // Clear status message on load
    const s = document.getElementById("status");
    if (s) s.textContent = "";
  }

  // When the user clicks "selectHeader", define a named range "HeaderRange" from the current selection
  document.getElementById("selectHeader").onclick = async () => {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.name = "HeaderRange";
      await ctx.sync();
      updateStatus("✅ HeaderRange defined from selected cells.");
    });
  };
  const logoutBtn = document.getElementById("requestLogout");
  if (logoutBtn) {
    logoutBtn.addEventListener("click", requestLogout);
    console.log("Logout button event listener attached.");
  }
});

// Main function to run the import process from selected RTF files
async function fastCWAHU_FromMappingSheet() {
  try {
    const files = await pickRtfFiles(); // Let user choose RTF files
    if (!files.length) return;

    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet(); // Get current sheet
      const { ws: mapWS, wasCreated } = await ensureMappingSheetExists(ctx); // Ensure mapping sheet exists
      if (wasCreated) updateStatus("✅ Mapping sheet was created.");

      const headers = await buildHeaderMap(ws, ctx); // Build header map from header row
      const mDict = await loadMappingDict(mapWS, ctx); // Load the mapping dictionary

      let nextRow = 6;
      let tempRow = {};
      let curSection = "";

      // Loop through each RTF file
      for (const file of files) {
        const text = await rtfToPlain(file); // Convert RTF to plain text
        const lines = text.replace(/\r\n?/g, "\n").split("\n");

        for (const ln of lines) {
          const txt = ln.trim();
          if (!txt) continue;

          // Detect current section by matching keywords
          if (/Sizing Data|Cooling Coil|Outdoor Ventilation/i.test(txt)) {
            curSection = txt.toUpperCase();
          }

          // Check if line contains "Air System Name"
          const unitHeader = mDict["AIR SYSTEM NAME"]?.[0]?.[0] ?? "";
          if (unitHeader && /Air System Name/i.test(txt)) {
            // If previous row had content, write it
            if (Object.keys(tempRow).length > 1) {
              writeRow(ws, headers, ++nextRow, tempRow);
            }
            tempRow = {};
            tempRow[unitHeader] = extractAfter(txt, "Air System Name");
            continue;
          }

          // Match line against mapping rules and collect data
          const hits = matchMappedTermsBySection(txt, curSection, mDict, headers);
          Object.assign(tempRow, hits);
        }
      }

      // Write the final row if it has values
      if (Object.keys(tempRow).length > 1) {
        writeRow(ws, headers, ++nextRow, tempRow);
      }

      await ctx.sync();
      updateStatus("✅ Import complete.");
    });
  } catch (err) {
    console.error(err);
    updateStatus("❌ " + err.message);
  }
}

// Let user pick multiple RTF files via hidden input
function pickRtfFiles() {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".rtf";
    input.multiple = true;
    input.style.display = "none";
    document.body.appendChild(input);
    input.onchange = () => {
      resolve([...input.files]);
      document.body.removeChild(input);
    };
    input.click();
  });
}

// Update the status text in the HTML
function updateStatus(msg) {
  const el = document.getElementById("status");
  if (el) el.textContent = msg;
}

// Convert RTF file to plain text by stripping RTF formatting
function rtfToPlain(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = () => {
      const rtf = reader.result;
      const txt = rtf
        .replace(/\\par[d]?/g, "\n")
        .replace(/\\'[0-9a-f]{2}/gi, (m) => String.fromCharCode(parseInt(m.substr(2), 16)))
        .replace(/\\[a-z]+\d* ?/gi, "")
        .replace(/[{}]/g, "");
      res(txt);
    };
    reader.onerror = () => rej(reader.error);
    reader.readAsText(file);
  });
}

// Ensure the "Mapping" worksheet exists and populate it with default structure if it doesn't
async function ensureMappingSheetExists(ctx) {
  const sheets = ctx.workbook.worksheets;
  let ws = sheets.getItemOrNullObject("Mapping");
  await ctx.sync();

  let wasCreated = false;
  if (ws.isNullObject) {
    ws = sheets.add("Mapping");
    wasCreated = true;
    await ctx.sync();

    // Header formatting and example data
    ws.getRange("A1:B1").merge(true);
    ws.getRange("A1").values = [["Schedule header"]];
    ws.getRange("C1:D1").merge(true);
    ws.getRange("C1").values = [["Import Data"]];
    ws.getRange("A2:D2").values = [["Section", "Column", "Section", "HAP term"]];

    const s = [
      ["", "", "Supply Fan Sizing Data", "Actual max CFM"],
      ["", "", "Supply Fan Sizing Data", "Fan static"],
      ["", "", "Supply Fan Sizing Data", "Fan motor kW"],
      ["", "", "", ""],
      ["", "", "Return Fan Sizing Data", "Actual max L/s"],
      ["", "", "Return Fan Sizing Data", "Fan motor kW"],
      ["", "", "Return Fan Sizing Data", "Fan static"],
      ["", "", "", ""],
      ["", "", "Outdoor Ventilation Air Data", "CFMJ"],
      ["", "", "", ""],
      ["", "", "Central Cooling Coil Sizing Data", "Total coil load"],
      ["", "", "Central Cooling Coil Sizing Data", "Sensible coil load"],
      ["", "", "", ""],
      ["", "", "Air System Information", "Air System Name"],
      ["", "", "Central Cooling Coil Sizing Data", "CMax block CFM"],
      ["", "", "", ""],
      ["", "", "Supply Fan Sizing Data", "Actual max CFM "],
      ["", "", "Supply Fan Sizing Data", " Standard CFM "],
      ["", "", "Supply Fan Sizing Data", "Calculation Months"]
    ];
    ws.getRange("A3:D17").values = s;

    // Style the mapping table
    const cols = ws.getRange("A:D");
    cols.columnWidth = 140;
    cols.format.horizontalAlignment = "Center";
    cols.format.verticalAlignment = "Center";
    const block = ws.getRange("A1:D17");
    block.format.font.name = "Calibri";
    block.format.font.size = 12;
    block.getCell(0, 0).getResizedRange(1, 3).format.font.bold = true;
    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"].forEach(
      (b) => (block.format.borders.getItem(b).style = "Continuous")
    );

    await ctx.sync();
  }
  return { ws, wasCreated };
}

// Build a map from headers in the defined HeaderRange (column title => column index)
async function buildHeaderMap(ws, ctx) {
  let rangeObj = ws.names.getItemOrNullObject("HeaderRange");
  await ctx.sync();

  if (rangeObj.isNullObject) {
    const selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "values"]);
    await ctx.sync();

    ws.names.add("HeaderRange", selectedRange);
    rangeObj = selectedRange;
  } else {
    rangeObj = rangeObj.getRange();
    rangeObj.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "values"]);
    await ctx.sync();
  }

  const { rowCount, columnCount, columnIndex, values } = rangeObj;
  const map = {};

  for (let c = 0; c < columnCount; c++) {
    let top = "", bottom = "";
    for (let r = 0; r < rowCount && !top; r++) top = (values[r][c] || "").toString().trim();
    for (let r = rowCount - 1; r >= 0 && !bottom; r--) bottom = (values[r][c] || "").toString().trim();

    const key = top.toUpperCase() === bottom.toUpperCase()
      ? top
      : top && bottom
        ? `${top}|${bottom}`
        : top || `|${bottom}`;

    if (key && !map[key.toUpperCase()]) {
      map[key.toUpperCase()] = columnIndex + c + 1;
    }
  }

  return map;
}

// Load the mapping dictionary from the Mapping sheet
async function loadMappingDict(sheet, ctx) {
  const rng = sheet.getUsedRangeOrNullObject();
  await ctx.sync();
  if (rng.isNullObject) return {};

  rng.load("values");
  await ctx.sync();

  const dict = {};
  rng.values.slice(2).forEach((row) => {
    const [schedA, schedB, section, hapTerm] = row.map((v) => (v || "").toString());
    if (!schedB || !hapTerm) return;
    const schedHeader = `${schedA ? schedA + "|" : ""}${schedB}`.toUpperCase();
    (dict[hapTerm.toUpperCase()] = dict[hapTerm.toUpperCase()] || []).push([schedHeader, section.toUpperCase()]);
  });
  return dict;
}

// Match a line of text against the mapping rules and section
function matchMappedTermsBySection(txt, currentSection, mappingDict, headerMap) {
  const res = {};
  const clean = cleanControlChars(txt).trim();

  for (const hapKey of Object.keys(mappingDict)) {
    if (clean.toLowerCase().startsWith(hapKey.toLowerCase())) {
      const nextChar = clean.charAt(hapKey.length) || " ";
      if (" .:-/".includes(nextChar)) {
        const extracted = extractNumber(clean);
        for (const [schedHeader, reqSection] of mappingDict[hapKey]) {
          if (!reqSection || currentSection.includes(reqSection)) {
            if (headerMap[schedHeader.toUpperCase()]) {
              res[schedHeader] = schedHeader.includes("POWER")
                ? stdPower(parseFloat(extracted))
                : extracted;
            }
          }
        }
      }
    }
  }
  return res;
}

// Clean unwanted characters (non-ASCII control characters)
function cleanControlChars(str = "") {
  return [...str].filter((ch) => ch.charCodeAt(0) >= 32 && ch.charCodeAt(0) <= 126).join("");
}

// Extract a numeric string from the end of the line
function extractNumber(input = "") {
  let out = "";
  for (let i = input.length - 1; i >= 0; i--) {
    const ch = input[i];
    if (/[0-9\-\./\s]/.test(ch)) out = ch + out;
    else break;
  }
  return out.trim().replace(/^\/|\/$/g, "");
}

// Normalize power values to standard values
function stdPower(val) {
  const table = [0, 0.09, 0.19, 0.38, 0.56, 0.75, 1.13, 1.5, 2.25, 3.75, 5.6, 7.5, 11.3, 15, 18.8, 22.5, 30, 37.5, 45, 56.3, 75, 93, 113.5, 150];
  for (const p of table) if (p >= val) return p;
  return val;
}

// Extract everything after a specific keyword
function extractAfter(txt, needle) {
  return txt.split(new RegExp(needle, "i"))[1]?.trim() || "";
}

// Write a row of data into the worksheet using the header map
function writeRow(ws, headerMap, row, data) {
  Object.entries(data).forEach(([hdr, val]) => {
    const col = headerMap[hdr.toUpperCase()];
    if (col) ws.getCell(row - 1, col - 1).values = [[val]];
  });
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