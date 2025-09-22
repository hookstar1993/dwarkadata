https://script.google.com/macros/s/AKfycbwe4NdnSrLuHv4qN761N1ebPQ7_nlRT-hBbwTKQkGvb7bM0cKoT4wePMwYuVDswzDBI/exec



/**
 * @license
 * Copyright 2024 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

// --- CONFIGURATION ---
// IMPORTANT: Set your Google Sheet name here.
const SHEET_NAME = "Sheet1"; 
// Columns for which to get unique values for form datalists.
const dynamicDatalistColumns = ['guest_name', 'guest_initials', 'departure_city', 'arrival_city'];

/**
 * Retrieves all data from the sheet and formats it for the web app.
 * This is the primary function for the GET request.
 */
function doGet(e) {
  try {
    const data = getAllData();
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log(error.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles all POST requests from the web app (create, update, delete, etc.).
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Wait for up to 30 seconds for other processes to finish.

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error("Sheet '" + SHEET_NAME + "' not found.");
    }
    
    const action = e.parameter.action;

    // --- ACTION: Create a new record ---
    if (action === 'create') {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(header => e.parameter[header] || "");
      sheet.appendRow(newRow);
    
    // --- ACTION: Update an entire row (legacy, not used by new form) ---
    } else if (action === 'update') {
      const rowIndex = parseInt(e.parameter.rowIndex, 10);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const updatedRow = headers.map(header => e.parameter[header] || "");
      if (rowIndex > 1) { // Ensure we are not overwriting the header
          sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
      } else {
        throw new Error("Invalid row index for update.");
      }
    
    // --- ACTION: Update a single field in a row ---
    } else if (e.parameter.action === 'updateField') {
        const rowIndex = parseInt(e.parameter.rowIndex, 10);
        const fieldName = e.parameter.fieldName;
        const fieldValue = e.parameter.fieldValue;
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const colIndex = headers.indexOf(fieldName) + 1;
        
        if (colIndex > 0 && rowIndex > 1) {
            sheet.getRange(rowIndex, colIndex).setValue(fieldValue);
            // For single field updates, a simple success message is enough.
            // The front-end will fetch all data again upon completion.
            return ContentService.createTextOutput(JSON.stringify({ status: 'SUCCESS', message: `Updated ${fieldName}` })).setMimeType(ContentService.MimeType.JSON);
        } else {
            throw new Error('Field not found or invalid row for updateField action.');
        }

    // --- ACTION: Delete a record ---
    } else if (action === 'delete') {
      const rowIndex = parseInt(e.parameter.rowIndex, 10);
       if (rowIndex > 1) {
          sheet.deleteRow(rowIndex);
      } else {
        throw new Error("Invalid row index for delete.");
      }

    // --- ACTION: Duplicate a record ---
    } else if (action === 'duplicate') {
      const rowIndex = parseInt(e.parameter.rowIndex, 10);
      if (rowIndex > 1) {
        const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues();
        sheet.appendRow(rowValues[0]);
      } else {
         throw new Error("Invalid row index for duplicate.");
      }

    // --- No valid action specified ---
    } else {
      throw new Error(`Invalid action specified: "${action}"`);
    }

    // After a successful CUD operation, return all the latest data.
    const data = getAllData();
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log(error.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", message: `Server-side error during '${e.parameter.action}': ${error.message}` })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * A helper function to fetch and format all data from the sheet.
 * @returns {Object} An object containing headers, rows of data, and unique values for specified columns.
 */
function getAllData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
      throw new Error("Sheet '" + SHEET_NAME + "' not found.");
  }
  const range = sheet.getDataRange();
  const values = range.getValues();

  if (values.length === 0) {
    return { headers: [], rows: [], uniqueValues: {} };
  }

  const headers = values.shift(); // Get headers and remove them from data
  const uniqueValues = {};
  
  // Initialize unique value sets
  dynamicDatalistColumns.forEach(header => {
    if (headers.includes(header)) {
      uniqueValues[header] = new Set();
    }
  });
  
  const rows = values.map((row, index) => {
    const rowObject = {};
    headers.forEach((header, i) => {
      rowObject[header] = row[i];
      // Add to unique value set if the column is tracked
      if (dynamicDatalistColumns.includes(header) && row[i]) {
        uniqueValues[header].add(row[i]);
      }
    });
    rowObject.rowIndex = index + 2; // +2 because Sheets is 1-indexed and we shifted the header row
    return rowObject;
  });

  // Convert sets to sorted arrays
  for (const header in uniqueValues) {
    uniqueValues[header] = Array.from(uniqueValues[header]).sort();
  }

  return { headers: headers, rows: rows, uniqueValues: uniqueValues, status: 'SUCCESS' };
}
