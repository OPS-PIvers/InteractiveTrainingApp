// Server/SheetService.gs

/**
 * Appends a new row to a specified sheet.
 * @param {string} sheetId The ID of the Google Spreadsheet.
 * @param {string} sheetName The name of the sheet (tab) within the spreadsheet.
 * @param {Array<any>} rowData An array of values representing the row to add.
 * @return {void}
 */
function appendRowToSheet(sheetId, sheetName, rowData) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in Spreadsheet ID "${sheetId}".`);
    }
    sheet.appendRow(rowData);
    Logger.log(`Row appended to ${sheetName}: ${rowData.join(', ')}`);
  } catch (e) {
    Logger.log(`Error in appendRowToSheet: ${e.toString()} - SheetID: ${sheetId}, SheetName: ${sheetName}, Data: ${rowData}. Stack: ${e.stack ? e.stack : 'N/A'}`);
    throw new Error(`Failed to append row: ${e.message}`);
  }
}

/**
 * Finds the 1-based row index of the first row matching a value in a specific column.
 * Includes header row in search, so careful if value could be a header.
 * It's often better to search data range only (excluding headers).
 * This version searches the entire column.
 * @param {string} sheetId The ID of the Google Spreadsheet.
 * @param {string} sheetName The name of the sheet (tab) within the spreadsheet.
 * @param {number} columnIndex The 1-based index of the column to search.
 * @param {any} value The value to search for.
 * @return {number|null} The 1-based row index if found, otherwise null.
 */
function findRowIndexByValue(sheetId, sheetName, columnIndex, value) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in Spreadsheet ID "${sheetId}".`);
    }
    const dataRange = sheet.getRange(1, columnIndex, sheet.getLastRow()); // Search entire column
    const values = dataRange.getValues();

    for (let i = 0; i < values.length; i++) {
      if (values[i][0] == value) { // Use == for type flexibility, or === for strict matching
        Logger.log(`Value "${value}" found in column ${columnIndex} at row ${i + 1} in sheet ${sheetName}.`);
        return i + 1; // 1-based row index
      }
    }
    Logger.log(`Value "${value}" not found in column ${columnIndex} in sheet ${sheetName}.`);
    return null;
  } catch (e) {
    Logger.log(`Error in findRowIndexByValue: ${e.toString()} - SheetID: ${sheetId}, SheetName: ${sheetName}, Column: ${columnIndex}, Value: ${value}. Stack: ${e.stack ? e.stack : 'N/A'}`);
    throw new Error(`Failed to find row: ${e.message}`);
  }
}


/**
 * Updates an existing row in a specified sheet.
 * @param {string} sheetId The ID of the Google Spreadsheet.
 * @param {string} sheetName The name of the sheet (tab) within the spreadsheet.
 * @param {number} rowIndex The 1-based index of the row to update.
 * @param {Array<any>} rowData An array of values for the new row content. Must match number of columns.
 * @return {void}
 */
function updateSheetRow(sheetId, sheetName, rowIndex, rowData) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in Spreadsheet ID "${sheetId}".`);
    }
    if (rowIndex <= 0 || rowIndex > sheet.getMaxRows()) {
        throw new Error(`Row index ${rowIndex} is out of bounds for sheet "${sheetName}". Max rows: ${sheet.getMaxRows()}`);
    }
    // Ensure rowData has enough columns; if not, it only updates provided columns.
    // For safety, ensure rowData matches expected number of columns or pad it.
    // For now, assume rowData matches the columns to be updated.
    const numColumns = rowData.length > 0 ? rowData.length : sheet.getLastColumn(); // Or a fixed number of columns
    if (numColumns === 0 && sheet.getLastColumn() === 0) { // Empty sheet
        Logger.log(`Sheet ${sheetName} is empty. Cannot update row ${rowIndex}.`);
        return;
    }
    sheet.getRange(rowIndex, 1, 1, numColumns).setValues([rowData]);
    Logger.log(`Row ${rowIndex} updated in ${sheetName} with data: ${rowData.join(', ')}`);
  } catch (e) {
    Logger.log(`Error in updateSheetRow: ${e.toString()} - SheetID: ${sheetId}, SheetName: ${sheetName}, RowIndex: ${rowIndex}, Data: ${rowData}. Stack: ${e.stack ? e.stack : 'N/A'}`);
    throw new Error(`Failed to update row: ${e.message}`);
  }
}

/**
 * Deletes a row from a specified sheet.
 * @param {string} sheetId The ID of the Google Spreadsheet.
 * @param {string} sheetName The name of the sheet (tab) within the spreadsheet.
 * @param {number} rowIndex The 1-based index of the row to delete.
 * @return {void}
 */
function deleteSheetRow(sheetId, sheetName, rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in Spreadsheet ID "${sheetId}".`);
    }
    if (rowIndex <= 0 || rowIndex > sheet.getLastRow()) { // Check against getLastRow for existing rows
        throw new Error(`Row index ${rowIndex} is out of bounds for deletable rows in sheet "${sheetName}". Last row: ${sheet.getLastRow()}`);
    }
    sheet.deleteRow(rowIndex);
    Logger.log(`Row ${rowIndex} deleted from ${sheetName}.`);
  } catch (e) {
    Logger.log(`Error in deleteSheetRow: ${e.toString()} - SheetID: ${sheetId}, SheetName: ${sheetName}, RowIndex: ${rowIndex}. Stack: ${e.stack ? e.stack : 'N/A'}`);
    throw new Error(`Failed to delete row: ${e.message}`);
  }
}

/**
 * Retrieves all data from a specified sheet.
 * Assumes the first row contains headers. Returns an array of objects if headers exist,
 * otherwise returns a 2D array of values.
 * @param {string} sheetId The ID of the Google Spreadsheet.
 * @param {string} sheetName The name of the sheet (tab) within the spreadsheet.
 * @return {Array<Object>|Array<Array<any>>} Array of objects (if headers) or 2D array of values.
 */
function getAllSheetData(sheetId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in Spreadsheet ID "${sheetId}".`);
    }
    const range = sheet.getDataRange();
    const values = range.getValues();

    if (values.length === 0) {
      Logger.log(`getAllSheetData: Sheet ${sheetName} is empty. Returning [].`);
      return [];
    }
    
    if (values.length === 1) {
      Logger.log(`getAllSheetData: Sheet ${sheetName} has only one row (assumed to be headers). Returning [] as there is no data.`);
      return [];
    }

    // Always treat the first row as headers
    const headers = values[0];
    const dataObjects = [];
    
    for (let i = 1; i < values.length; i++) {
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        // Ensure header is a string and not empty to be a valid key
        const headerKey = (typeof headers[j] === 'string' && headers[j].trim() !== '') ? headers[j].trim() : `_col${j+1}`;
        obj[headerKey] = values[i][j];
      }
      dataObjects.push(obj);
    }
    Logger.log(`getAllSheetData: Retrieved ${dataObjects.length} data objects from ${sheetName} using first row as headers.`);
    return dataObjects;
    
  } catch (e) {
    Logger.log(`Error in getAllSheetData: ${e.toString()} - SheetID: ${sheetId}, SheetName: ${sheetName}. Stack: ${e.stack ? e.stack : 'No stack'}`);
    throw new Error(`Failed to get all data: ${e.message}`);
  }
}

/**
 * Gets an entire row's data by its 1-based index.
 * @param {string} sheetId The ID of the Google Spreadsheet.
 * @param {string} sheetName The name of the sheet (tab) within the spreadsheet.
 * @param {number} rowIndex The 1-based row index.
 * @return {Array<any>|null} An array of values for the row, or null if row index is invalid.
 */
function getSheetRowData(sheetId, sheetName, rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in Spreadsheet ID "${sheetId}".`);
    }
    if (rowIndex <= 0 || rowIndex > sheet.getLastRow()) {
      Logger.log(`Invalid row index ${rowIndex} for sheet ${sheetName}. Last row: ${sheet.getLastRow()}.`);
      return null;
    }
    const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log(`Data for row ${rowIndex} in sheet ${sheetName}: ${rowValues.join(', ')}`);
    return rowValues;
  } catch (e) {
    Logger.log(`Error in getSheetRowData: ${e.toString()} - SheetID: ${sheetId}, SheetName: ${sheetName}, RowIndex: ${rowIndex}. Stack: ${e.stack ? e.stack : 'N/A'}`);
    throw new Error(`Failed to get row data: ${e.message}`);
  }
}