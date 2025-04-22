function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Fetch Report")
    .addItem("Run", "fetchReport")
    .addToUi();
}

function fetchReport() {
  const email = Session.getActiveUser().getEmail();
  let parsedData;

  try {
    parsedData = readParamsFromSheet(email);
  } catch (e) {
    SpreadsheetApp.getUi().alert(e.message);
    return;
  }

  if (!parsedData.url) {
    SpreadsheetApp.getUi().alert("URL is missing. Please provide a valid URL in the Developer Settings sheet.");
    return; // Exit without making any changes
  }

  const url = parsedData.url.trim();
  if (!url.startsWith("http://") && !url.startsWith("https://")) {
    SpreadsheetApp.getUi().alert("Invalid URL format.");
    return;
  }

  if (!parsedData.targetTabs || parsedData.targetTabs.length === 0) {
    SpreadsheetApp.getUi().alert("Missing TargetTabs. Please provide TargetTabs.");
    return;
  }

  fetchReportWithParams(parsedData);
}

function fetchReportWithParams(parsedData) {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const runid = [...Array(8)].map(() => (Math.random().toString(36)[2])).join("");
    parsedData.params.runid = runid;

    let url = `${parsedData.url.trim()}?api_key=${parsedData.params.api_key}`;

    let headers = {};

    const username = parsedData.params.data_server_username;
    const password = parsedData.params.data_server_password;

    if (username && password) {
      const basicAuthHeader = "Basic " + Utilities.base64Encode(`${username}:${password}`);
      headers["Authorization"] = basicAuthHeader;
    }

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(parsedData.params),
      headers: headers
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      const errorMessage = `HTTP Error: ${responseCode} - ${response.getContentText()}`;
      Logger.log(errorMessage);
      ui.alert(`Error fetching data: ${errorMessage}`);
      return;
    }

    const rawCsvData = response.getContentText();
    const { confirmationMessage, historyBlock, a1Note, actualCSV } = parseCSVSections(rawCsvData);

    if (historyBlock) {
      appendToHistoryTable(historyBlock, parsedData.targetTabs, parsedData.dataLoadMode, runid);
    }

    if (confirmationMessage) {
      const userResponse = ui.alert(confirmationMessage, ui.ButtonSet.OK_CANCEL);
      if (userResponse !== ui.Button.OK) return;

      parsedData.params.confirm_api_hits = 'no';
      fetchReportWithParams(parsedData);
      return;
    }

    const tabList = parsedData.targetTabs.join(", ");
    const loadModeMessage = parsedData.dataLoadMode === "append"
      ? "Data will be appended to the following tab(s):"
      : "Data will overwrite the existing content in the following tab(s):";

    ui.alert(`${loadModeMessage}\n\n${tabList}`, ui.ButtonSet.OK);

    parsedData.targetTabs.forEach(tabName => {
      let sheet = spreadsheet.getSheetByName(tabName);
      if (!sheet) {
        sheet = spreadsheet.insertSheet(tabName);
      }

      const rowsAdded = populateSheetWithCSV(sheet, actualCSV, parsedData.dataLoadMode, a1Note);

      // Dynamically change the message based on the dataLoadMode
      const loadAction = parsedData.dataLoadMode === "append" ? "appended to" : "overwritten in";
      const noteMessage = a1Note
        ? "Refer the Note of cell A1 for execution summary."
        : "";

      ui.alert(`Data fetched and ${loadAction} ${tabName} tab.\n\n${noteMessage}`);
    });

  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error fetching data: ${error.message}`);
    Logger.log(`Error fetching data: ${error.stack}`);
  }
}

function parseCSVSections(csvData) {
  const extractBlock = (data, startTag, endTag) => {
    const startIdx = data.indexOf(startTag);
    const endIdx = data.indexOf(endTag);
    if (startIdx === -1 || endIdx === -1 || endIdx <= startIdx) return { content: null, stripped: data };
    const content = data.substring(startIdx + startTag.length, endIdx).trim();
    const stripped = (data.substring(0, startIdx) + data.substring(endIdx + endTag.length)).trim();
    return { content, stripped };
  };

  let result = extractBlock(csvData, "#CONFIRMATION_MSG_START#", "#CONFIRMATION_MSG_END#");
  const confirmationMessage = result.content;
  csvData = result.stripped;

  result = extractBlock(csvData, "#EXECUTION_HISTORY_START#", "#EXECUTION_HISTORY_END#");
  const historyBlock = result.content;
  csvData = result.stripped;

  result = extractBlock(csvData, "#A1_NOTE_START#", "#A1_NOTE_END#");
  const a1Note = result.content;
  csvData = result.stripped;

  result = extractBlock(csvData, "#REPORT_DATA_START#", "#REPORT_DATA_END#");
  const actualCSV = result.content || csvData.trim();

  return { confirmationMessage, historyBlock, a1Note, actualCSV };
}

function appendToHistoryTable(historyBlock, targetTabs, dataLoadMode, runid) {
  if (!historyBlock) {
    Logger.log("No history block provided. Skipping history table update.");
    return;
  }

  const historyValues = historyBlock.split(",").map(cell => cell.trim());
  const values = [
    targetTabs.join("; "), // Target Tabs
    dataLoadMode,          // Data Load Mode
    Session.getActiveUser().getEmail(), // Email
    runid,                 // Run ID
    ...historyValues       // Invoked_at, Time Taken, Rows Fetched, API Calls Counts
  ];

  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = sheet.getSheetByName("Execution History");

  if (!historySheet) {
    Logger.log("Creating 'Execution History' tab.");
    historySheet = sheet.insertSheet("Execution History");
    const headers = [
      "Target Tabs",
      "Data Load Mode",
      "Email",
      "Run ID",
      "Invoked at",
      "Time Taken (Seconds)",
      "Rows fetched",
      "API Calls Counts"
    ];
    historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    historySheet.setColumnWidths(1, headers.length, 180);
  }

  const lastRow = historySheet.getLastRow();
  historySheet.getRange(lastRow + 1, 1, 1, values.length).setValues([values]);
  Logger.log("History table updated successfully.");
}

function readParamsFromSheet(email) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Developer Settings");
  const tabNames = spreadsheet.getSheets().map(s => s.getName());

  if (!sheet) throw new Error("Developer Settings sheet not found.");

  const data = sheet.getDataRange().getValues();

  const result = {
    url: null,
    targetTabs: [],
    dataLoadMode: "overwrite",
    params: {}
  };

  for (const [rawKey, rawValue] of data) {
    if (!rawKey) continue;
    const key = rawKey.toString().trim();
    const value = rawValue !== undefined && rawValue !== null ? rawValue.toString().trim() : "";

    if (key.startsWith("Params[")) {
      const paramKey = key.slice(7, -1);

      if (paramKey.toLowerCase() === "data_server_url") {result.url = value; continue;}
      else if (paramKey.toLowerCase() === "target_tabs") {result.targetTabs = value.split(",").map(s => s.trim()).filter(Boolean); continue;}
      else if (paramKey.toLowerCase() === "data_load_mode") {result.dataLoadMode = value.toLowerCase(); continue;}
      else if (paramKey.toLowerCase() === "prompt_for_gemini") {result.params.prompt_for_gemini = value; continue;}

      if (paramKey.endsWith(".params_table") && tabNames.includes(value)) {
        const ui = SpreadsheetApp.getUi();
        const confirm = ui.alert(`Param '${paramKey}' matches tab '${value}'.\nUse its sheet data as param?`, ui.ButtonSet.YES_NO);
        if (confirm === ui.Button.YES) {
          const tabData = spreadsheet.getSheetByName(value).getDataRange().getValues();
          const csvData = tabData.map(row => row.map(cell => `"${cell}"`).join(",")).join("\n"); // Convert to CSV
          result.params[paramKey] = csvData; // Assign CSV string
          continue;
        }
      }
      result.params[paramKey] = value;
    } else {
      result.params[key] = value;
    }
  }

  result.params.email = email;
  return result;
}

function populateSheetWithCSV(sheet, csvData, dataLoadMode, a1Note) {
  const rows = parseCSV(csvData);

  let startRow = 1; // Default start row
  if (dataLoadMode === "append") {
    const lastRow = sheet.getLastRow();
    if (lastRow > 0) {
      // Skip the header row from the CSV data if the sheet already has data
      rows.shift();
      startRow = lastRow + 1; // Append after the last row
    }
  } else if (dataLoadMode === "overwrite") {
    sheet.clear(); // Clear the sheet before populating
  }

  // Ensure the range matches the dimensions of the data
  const numRows = rows.length;
  const numCols = Math.max(...rows.map(row => row.length)); // Find the maximum number of columns in any row

  // Pad rows with fewer columns to match the maximum column count
  const paddedRows = rows.map(row => {
    while (row.length < numCols) {
      row.push(""); // Add empty strings to fill missing columns
    }
    return row;
  });

  // Write the CSV data to the sheet
  sheet.getRange(startRow, 1, numRows, numCols).setValues(paddedRows);

  // Add a note to cell A1 if provided
  if (a1Note) {
    const a1Range = sheet.getRange("A1");
    const existingNote = a1Range.getNote();
    const separator = "-".repeat(50); // 50 dashes
    const updatedNote = `Update Mode: ${dataLoadMode}\n\n${a1Note}\n${separator}\n${existingNote || ""}`.trim();
    a1Range.setNote(updatedNote);
  }

  return numRows; // Return the number of rows added
}

// Custom CSV parser to handle quotes, commas, and empty strings
function parseCSV(csvData) {
  const rows = [];
  let currentRow = [];
  let currentCell = "";
  let insideQuotes = false;

  for (let i = 0; i < csvData.length; i++) {
    const char = csvData[i];
    const nextChar = csvData[i + 1];

    if (char === '"' && insideQuotes && nextChar === '"') {
      // Handle escaped double quotes
      currentCell += '"';
      i++; // Skip the next quote
    } else if (char === '"') {
      // Toggle the insideQuotes flag
      insideQuotes = !insideQuotes;
    } else if (char === ',' && !insideQuotes) {
      // End of a cell
      currentRow.push(currentCell.trim());
      currentCell = "";
    } else if (char === '\n' && !insideQuotes) {
      // End of a row
      currentRow.push(currentCell.trim());
      rows.push(currentRow);
      currentRow = [];
      currentCell = "";
    } else {
      // Add character to the current cell
      currentCell += char;
    }
  }

  // Add the last cell and row if necessary
  if (currentCell || currentRow.length > 0) {
    currentRow.push(currentCell.trim());
    rows.push(currentRow);
  }

  return rows;
}

