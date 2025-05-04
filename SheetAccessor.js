/**
 * SheetAccessor class provides methods to interact with the Google Spreadsheet
 * Handles reading/writing data and managing sheet structure
 */
class SheetAccessor {
  /**
   * Creates a new SheetAccessor instance
   * * @param {string} spreadsheetId - ID of the spreadsheet to access (optional)
   */
  constructor(spreadsheetId) {
    this.spreadsheetId = spreadsheetId || this.getActiveSpreadsheetId();
    this.spreadsheet = null;
    if (!this.spreadsheetId) {
        throw new Error("Spreadsheet ID could not be determined.");
    }
    this.openSpreadsheet(); // Ensure spreadsheet is opened on instantiation
  }
  
  /**
   * Gets the ID of the active spreadsheet
   * * @return {string|null} Spreadsheet ID or null if none active
   */
  getActiveSpreadsheetId() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        return ss ? ss.getId() : null;
    } catch (e) {
        logWarning("Could not get active spreadsheet ID (may be running in a different context).");
        return null;
    }
  }
  
  /**
   * Opens the spreadsheet with the stored ID
   */
  openSpreadsheet() {
    if (this.spreadsheet) return; // Already open
    
    if (!this.spreadsheetId) {
      throw new Error("Cannot open spreadsheet: No spreadsheet ID provided or available");
    }
    
    try {
      this.spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
      logDebug(`Opened spreadsheet: ${this.spreadsheet.getName()} (ID: ${this.spreadsheetId})`);
    } catch (error) {
      logError(`Failed to open spreadsheet ID ${this.spreadsheetId}: ${error.message}`);
      this.spreadsheet = null; // Ensure it's null on failure
      throw error; // Re-throw the error
    }
  }
  
  /**
   * Creates a new spreadsheet with the specified name
   * * @param {string} name - Name for the new spreadsheet
   * @return {SheetAccessor} A new SheetAccessor instance for the created spreadsheet
   */
  static createSpreadsheet(name) {
    try {
      const newSpreadsheet = SpreadsheetApp.create(name);
      logInfo(`Created new spreadsheet: ${name} (${newSpreadsheet.getId()})`);
      return new SheetAccessor(newSpreadsheet.getId());
    } catch (error) {
      logError(`Failed to create spreadsheet '${name}': ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Gets a sheet by name.
   * * @param {string} sheetName - Name of the sheet.
   * @param {boolean} [createIfNotExist=true] - Whether to create the sheet if it doesn't exist.
   * @return {GoogleAppsScript.Spreadsheet.Sheet|null} Google Sheets Sheet object or null if not found and createIfNotExist is false.
   */
  getSheet(sheetName, createIfNotExist = true) {
      if (!this.spreadsheet) {
          logError(`Spreadsheet not open. Cannot get sheet: ${sheetName}`);
          return null;
      }
      let sheet = this.spreadsheet.getSheetByName(sheetName);
      
      if (!sheet && createIfNotExist) {
          try {
              sheet = this.spreadsheet.insertSheet(sheetName);
              logDebug(`Created new sheet: ${sheetName}`);
          } catch (error) {
               logError(`Failed to create sheet '${sheetName}': ${error.message}`);
               return null;
          }
      }
      
      if (!sheet && !createIfNotExist) {
           logDebug(`Sheet not found and not created: ${sheetName}`);
      }
      
      return sheet;
  }
  
  /**
   * Deletes a sheet by name
   * * @param {string} sheetName - Name of the sheet to delete
   * @return {boolean} True if successfully deleted or if sheet didn't exist. False on error.
   */
  deleteSheet(sheetName) {
    if (!this.spreadsheet) {
        logError(`Spreadsheet not open. Cannot delete sheet: ${sheetName}`);
        return false;
    }
    const sheet = this.spreadsheet.getSheetByName(sheetName);
    
    if (sheet) {
      try {
          this.spreadsheet.deleteSheet(sheet);
          logDebug(`Deleted sheet: ${sheetName}`);
          return true;
      } catch (error) {
           logError(`Failed to delete sheet '${sheetName}': ${error.message}`);
           return false;
      }
    }
    
    logWarning(`Sheet not found for deletion: ${sheetName}`);
    return true; // Considered success if it doesn't exist
  }
  
  /**
   * Gets cell value at specified location
   * * @param {string} sheetName - Name of the sheet
   * @param {number|string} row - Row number (1-based) or A1 notation (e.g., "A1")
   * @param {(number|string)} [column] - Column number (1-based) or column letter (e.g., "B"). Required if row is a number.
   * @return {*} Cell value, or null on error or if sheet/cell not found.
   */
  getCellValue(sheetName, row, column) {
    const sheet = this.getSheet(sheetName, false);
    if (!sheet) return null;
    
    try {
      let range;
      // Handle A1 notation
      if (typeof row === "string" && column === undefined) {
           if (!row.match(/^[A-Z]+[0-9]+$/i)) throw new Error(`Invalid A1 notation: ${row}`);
           range = sheet.getRange(row);
      } 
      // Handle row/column numbers/letters
      else if (typeof row === 'number' && row > 0 && column !== undefined) {
          let colIndex;
          if (typeof column === "string" && column.match(/^[A-Z]+$/i)) {
              colIndex = columnToIndex(column.toUpperCase()) + 1; // Convert to 1-based index
          } else if (typeof column === 'number' && column > 0) {
               colIndex = column;
          } else {
               throw new Error(`Invalid column identifier: ${column}`);
          }
          range = sheet.getRange(row, colIndex);
      } 
      // Invalid arguments
      else {
          throw new Error(`Invalid arguments for getCellValue: row=${row}, column=${column}`);
      }
      
      return range.getValue();
    } catch (error) {
      logError(`Failed to get cell value (${sheetName} ${row}:${column || ''}): ${error.message}`);
      return null;
    }
  }
  
  /**
   * Sets a cell value at specified location
   * * @param {string} sheetName - Name of the sheet
   * @param {number|string} row - Row number (1-based) or A1 notation (e.g., "A1")
   * @param {(number|string)} [column] - Column number (1-based) or column letter (e.g., "B"). Required if row is a number.
   * @param {*} value - Value to set
   * @return {boolean} True if successful, false otherwise.
   */
  setCellValue(sheetName, row, column, value) {
    // Use createIfNotExist = true, as setting a value often implies the sheet should exist.
    const sheet = this.getSheet(sheetName, true); 
    if (!sheet) return false;
    
    try {
      let range;
       // Handle A1 notation
      if (typeof row === "string" && column === undefined) {
           if (!row.match(/^[A-Z]+[0-9]+$/i)) throw new Error(`Invalid A1 notation: ${row}`);
           range = sheet.getRange(row);
      } 
      // Handle row/column numbers/letters
      else if (typeof row === 'number' && row > 0 && column !== undefined) {
          let colIndex;
          if (typeof column === "string" && column.match(/^[A-Z]+$/i)) {
              colIndex = columnToIndex(column.toUpperCase()) + 1; // Convert to 1-based index
          } else if (typeof column === 'number' && column > 0) {
               colIndex = column;
          } else {
               throw new Error(`Invalid column identifier: ${column}`);
          }
          range = sheet.getRange(row, colIndex);
      } 
      // Invalid arguments
      else {
          throw new Error(`Invalid arguments for setCellValue: row=${row}, column=${column}`);
      }

      range.setValue(value);
      // logDebug(`Set cell value (${sheetName} ${row}:${column || ''}) to: ${value}`); // Optional: verbose logging
      return true;
    } catch (error) {
      logError(`Failed to set cell value (${sheetName} ${row}:${column || ''}): ${error.message}`);
      return false;
    }
  }
  
  /**
   * Gets values from a range defined by start/end rows and columns or A1 notation.
   * * @param {string} sheetName - Name of the sheet.
   * @param {number|string} startRowOrA1Notation - Starting row (1-based) or A1 notation string (e.g., "A1:C5").
   * @param {number|string} [startCol] - Starting column (1-based) or column letter. Required if startRowOrA1Notation is a number.
   * @param {number} [numRows] - Number of rows. Required if startRowOrA1Notation is a number.
   * @param {number} [numCols] - Number of columns. Required if startRowOrA1Notation is a number.
   * @return {Array<Array<*>>|null} 2D array of values, or null on error or if sheet not found.
   */
  getRangeValues(sheetName, startRowOrA1Notation, startCol, numRows, numCols) {
    const sheet = this.getSheet(sheetName, false);
    if (!sheet) return null;
    
    try {
      let range;
      if (typeof startRowOrA1Notation === "string") {
          // Assume A1 notation if only one argument is string
           if (!startRowOrA1Notation.match(/^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/i)) throw new Error(`Invalid A1 notation: ${startRowOrA1Notation}`);
           range = sheet.getRange(startRowOrA1Notation);
      } else if (typeof startRowOrA1Notation === 'number' && startCol !== undefined && numRows !== undefined && numCols !== undefined) {
          // Handle column as letter (A, B, C, etc.)
          let startColIndex;
          if (typeof startCol === "string" && startCol.match(/^[A-Z]+$/i)) {
            startColIndex = columnToIndex(startCol.toUpperCase()) + 1; // Convert to 1-based index
          } else if (typeof startCol === 'number') {
               startColIndex = startCol;
          } else {
               throw new Error(`Invalid start column: ${startCol}`);
          }
          if (startRowOrA1Notation <= 0 || startColIndex <= 0 || numRows <= 0 || numCols <= 0) {
               throw new Error(`Invalid range dimensions: ${startRowOrA1Notation},${startColIndex},${numRows},${numCols}`);
          }
          range = sheet.getRange(startRowOrA1Notation, startColIndex, numRows, numCols);
      } else {
           throw new Error("Invalid arguments for getRangeValues.");
      }
      
      return range.getValues();
    } catch (error) {
      logError(`Failed to get range values (${sheetName}, ${startRowOrA1Notation}): ${error.message}`);
      return null;
    }
  }
  
  /**
   * Sets values in a range defined by start row/col or A1 notation.
   * * @param {string} sheetName - Name of the sheet.
   * @param {number|string} startRowOrA1Notation - Starting row (1-based) or A1 notation string (e.g., "A1").
   * @param {(number|string)} [startCol] - Starting column (1-based) or column letter. Required if startRowOrA1Notation is a number.
   * @param {Array<Array<*>>} values - 2D array of values to set. Must not be empty.
   * @return {boolean} True if successful, false otherwise.
   */
  setRangeValues(sheetName, startRowOrA1Notation, startCol, values) {
    // Use createIfNotExist = true
    const sheet = this.getSheet(sheetName, true); 
    if (!sheet) return false;
    
    if (!values || !Array.isArray(values) || values.length === 0 || !Array.isArray(values[0]) || values[0].length === 0) {
        logError(`Invalid or empty values array provided for setRangeValues in ${sheetName}.`);
        return false; // Cannot set empty values
    }

    try {
      const numRows = values.length;
      const numCols = values[0].length;
      let range;

      if (typeof startRowOrA1Notation === "string") {
           // A1 notation provided for the top-left cell
           if (!startRowOrA1Notation.match(/^[A-Z]+[0-9]+$/i)) throw new Error(`Invalid A1 notation for top-left cell: ${startRowOrA1Notation}`);
           range = sheet.getRange(startRowOrA1Notation).offset(0, 0, numRows, numCols);
      } else if (typeof startRowOrA1Notation === 'number' && startCol !== undefined) {
           // Handle column as letter (A, B, C, etc.)
           let startColIndex;
           if (typeof startCol === "string" && startCol.match(/^[A-Z]+$/i)) {
             startColIndex = columnToIndex(startCol.toUpperCase()) + 1; // Convert to 1-based index
           } else if (typeof startCol === 'number') {
                startColIndex = startCol;
           } else {
                throw new Error(`Invalid start column: ${startCol}`);
           }
           if (startRowOrA1Notation <= 0 || startColIndex <= 0) {
                throw new Error(`Invalid start coordinates: ${startRowOrA1Notation},${startColIndex}`);
           }
           range = sheet.getRange(startRowOrA1Notation, startColIndex, numRows, numCols);
      } else {
          throw new Error("Invalid arguments for setRangeValues.");
      }
      
      range.setValues(values);
      // logDebug(`Set range values (${sheetName}, ${startRowOrA1Notation}, ${numRows}x${numCols})`); // Optional: verbose logging
      return true;
    } catch (error) {
      logError(`Failed to set range values (${sheetName}, ${startRowOrA1Notation}): ${error.message}\n${error.stack}`);
      return false;
    }
  }
  
  /**
   * Gets all data from a sheet's data region.
   * * @param {string} sheetName - Name of the sheet.
   * @return {Array<Array<*>>} 2D array of all data, or empty array on error or if sheet not found/empty.
   */
  getSheetData(sheetName) {
    const sheet = this.getSheet(sheetName, false);
    if (!sheet) return [];
    
    try {
      // Check if sheet has data before getting range
      if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
          return []; // Sheet is empty
      }
      const dataRange = sheet.getDataRange();
      return dataRange.getValues();
    } catch (error) {
      logError(`Failed to get sheet data for ${sheetName}: ${error.message}`);
      return [];
    }
  }
  
  /**
   * Appends a row to a sheet.
   * * @param {string} sheetName - Name of the sheet.
   * @param {Array<*>} rowData - Array of values for the row.
   * @return {boolean} True if successful, false otherwise.
   */
  appendRow(sheetName, rowData) {
    // Use createIfNotExist = true
    const sheet = this.getSheet(sheetName, true); 
    if (!sheet) return false;
    
    if (!rowData || !Array.isArray(rowData)) {
        logError(`Invalid rowData provided for appendRow in ${sheetName}.`);
        return false;
    }

    try {
      sheet.appendRow(rowData);
      return true;
    } catch (error) {
      logError(`Failed to append row to ${sheetName}: ${error.message}`);
      return false;
    }
  }

  /**
   * Clears the content and formatting of a specified range.
   * * @param {string} sheetName - Name of the sheet.
   * @param {number|string} startRowOrA1Notation - Starting row (1-based) or A1 notation string (e.g., "A1:C5").
   * @param {number|string} [startCol] - Starting column (1-based) or column letter. Required if startRowOrA1Notation is a number.
   * @param {number} [numRows] - Number of rows. Required if startRowOrA1Notation is a number.
   * @param {number} [numCols] - Number of columns. Required if startRowOrA1Notation is a number.
   * @return {boolean} True if successful, false otherwise.
   */
  clearRange(sheetName, startRowOrA1Notation, startCol, numRows, numCols) {
      const sheet = this.getSheet(sheetName, false); // Don't create if clearing
      if (!sheet) {
          logWarning(`Sheet ${sheetName} not found for clearing range.`);
          return true; // Consider it success if sheet doesn't exist
      }

      try {
          let range;
          if (typeof startRowOrA1Notation === "string") {
              // Assume A1 notation
              if (!startRowOrA1Notation.match(/^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/i)) throw new Error(`Invalid A1 notation: ${startRowOrA1Notation}`);
              range = sheet.getRange(startRowOrA1Notation);
          } else if (typeof startRowOrA1Notation === 'number' && startCol !== undefined && numRows !== undefined && numCols !== undefined) {
              // Handle column as letter
              let startColIndex;
              if (typeof startCol === "string" && startCol.match(/^[A-Z]+$/i)) {
                  startColIndex = columnToIndex(startCol.toUpperCase()) + 1;
              } else if (typeof startCol === 'number') {
                  startColIndex = startCol;
              } else {
                  throw new Error(`Invalid start column: ${startCol}`);
              }
               if (startRowOrA1Notation <= 0 || startColIndex <= 0 || numRows <= 0 || numCols <= 0) {
                   throw new Error(`Invalid range dimensions for clearing: ${startRowOrA1Notation},${startColIndex},${numRows},${numCols}`);
               }
              range = sheet.getRange(startRowOrA1Notation, startColIndex, numRows, numCols);
          } else {
              throw new Error("Invalid arguments for clearRange.");
          }

          range.clear({ contentsOnly: false, formatOnly: false }); // Clear everything: content, formatting, notes, etc.
          logDebug(`Cleared range ${range.getA1Notation()} in ${sheetName}`);
          return true;
      } catch (error) {
          logError(`Failed to clear range (${sheetName}, ${startRowOrA1Notation}): ${error.message}`);
          return false;
      }
  }

  /**
   * Inserts a specified number of rows after a given row index.
   * * @param {string} sheetName - Name of the sheet.
   * @param {number} afterRow - The row index (1-based) after which to insert rows.
   * @param {number} numRows - The number of rows to insert (default: 1).
   * @return {boolean} True if successful, false otherwise.
   */
  insertRows(sheetName, afterRow, numRows = 1) {
      const sheet = this.getSheet(sheetName, true); // Create sheet if it doesn't exist? Maybe false is safer.
      if (!sheet) return false;

      if (typeof afterRow !== 'number' || afterRow < 0 || typeof numRows !== 'number' || numRows <= 0) {
           logError(`Invalid arguments for insertRows: afterRow=${afterRow}, numRows=${numRows}`);
           return false;
      }

      try {
          sheet.insertRowsAfter(afterRow, numRows);
          logDebug(`Inserted ${numRows} rows after row ${afterRow} in ${sheetName}`);
          return true;
      } catch (error) {
          logError(`Failed to insert rows in ${sheetName}: ${error.message}`);
          return false;
      }
  }

  /**
   * Deletes a specified number of rows starting from a given row index.
   * * @param {string} sheetName - Name of the sheet.
   * @param {number} startRow - The starting row index (1-based) to delete.
   * @param {number} numRows - The number of rows to delete (default: 1).
   * @return {boolean} True if successful, false otherwise.
   */
  deleteRows(sheetName, startRow, numRows = 1) {
      const sheet = this.getSheet(sheetName, false); // Don't create if deleting
      if (!sheet) {
           logWarning(`Sheet ${sheetName} not found for deleting rows.`);
           return true; // Consider success if sheet doesn't exist
      }

      if (typeof startRow !== 'number' || startRow <= 0 || typeof numRows !== 'number' || numRows <= 0) {
           logError(`Invalid arguments for deleteRows: startRow=${startRow}, numRows=${numRows}`);
           return false;
      }

      // Avoid deleting all rows if possible
      if (startRow + numRows -1 > sheet.getMaxRows()) {
           numRows = sheet.getMaxRows() - startRow + 1;
           if (numRows <= 0) {
               logDebug(`No rows to delete at startRow ${startRow} in ${sheetName}`);
               return true; // Nothing to delete
           }
      }
      // Prevent deleting the very last row if it's the only one? Optional safeguard.
      // if (sheet.getMaxRows() === 1 && startRow === 1 && numRows === 1) {
      //     logWarning("Attempted to delete the last remaining row. Clearing instead.");
      //     return this.clearRange(sheetName, 1, 1, 1, sheet.getMaxColumns());
      // }


      try {
          sheet.deleteRows(startRow, numRows);
          logDebug(`Deleted ${numRows} rows starting from row ${startRow} in ${sheetName}`);
          return true;
      } catch (error) {
          logError(`Failed to delete rows in ${sheetName}: ${error.message}`);
          return false;
      }
  }
  
  // --- Functions related to specific sheet structures (Project Index, Project Tabs) ---
  // These rely on the structure defined in Config.js

  /**
   * Creates a section header with merged cells and formatting
   * * @param {string} sheetName - Name of the sheet
   * @param {number} row - Starting row (1-based)
   * @param {string|number} startCol - Starting column (letter or 1-based)
   * @param {string|number} endCol - Ending column (letter or 1-based)
   * @param {string} headerText - Text for the header
   * @return {boolean} True if successful
   */
  createSectionHeader(sheetName, row, startCol, endCol, headerText) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return false;
    
    try {
      // Convert column letters to indices if needed
      let startColIndex = (typeof startCol === "string" && startCol.match(/^[A-Z]+$/i)) 
                          ? columnToIndex(startCol.toUpperCase()) + 1 
                          : (typeof startCol === 'number' ? startCol : 1);
      let endColIndex = (typeof endCol === "string" && endCol.match(/^[A-Z]+$/i)) 
                        ? columnToIndex(endCol.toUpperCase()) + 1 
                        : (typeof endCol === 'number' ? endCol : startColIndex);
      
      if (startColIndex <= 0 || endColIndex < startColIndex || row <= 0) {
           throw new Error(`Invalid range for section header: ${row}, ${startCol}, ${endCol}`);
      }

      // Merge cells for header
      const headerRange = sheet.getRange(row, startColIndex, 1, endColIndex - startColIndex + 1);
      headerRange.merge();
      
      // Set header text and formatting
      headerRange.setValue(headerText);
      headerRange.setBackground(DEFAULT_COLORS.SECTION_HEADER_BG);
      headerRange.setFontWeight("bold");
      headerRange.setHorizontalAlignment("center");
      
      return true;
    } catch (error) {
      logError(`Failed to create section header '${headerText}' in ${sheetName}: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Creates a dropdown in a cell with validation
   * * @param {string} sheetName - Name of the sheet
   * @param {number|string} row - Row number (1-based) or A1 notation
   * @param {number|string} column - Column number (1-based) or column letter
   * @param {Array<string>} options - Array of dropdown options
   * @return {boolean} True if successful
   */
  createDropdown(sheetName, row, column, options) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return false;
    
    try {
      let range;
      if (typeof row === "string" && column === undefined) {
           if (!row.match(/^[A-Z]+[0-9]+$/i)) throw new Error(`Invalid A1 notation: ${row}`);
           range = sheet.getRange(row);
      } else if (typeof row === 'number' && column !== undefined) {
          let colIndex;
          if (typeof column === "string" && column.match(/^[A-Z]+$/i)) {
              colIndex = columnToIndex(column.toUpperCase()) + 1;
          } else if (typeof column === 'number') {
              colIndex = column;
          } else {
              throw new Error(`Invalid column identifier: ${column}`);
          }
           if (row <= 0 || colIndex <= 0) throw new Error(`Invalid coordinates: ${row}, ${colIndex}`);
          range = sheet.getRange(row, colIndex);
      } else {
           throw new Error(`Invalid arguments for createDropdown: row=${row}, column=${column}`);
      }
      
      // Create validation rule
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(options, true) // true = show dropdown arrow
        .setAllowInvalid(false) // Disallow values not in the list
        .build();
      
      range.setDataValidation(rule);
      return true;
    } catch (error) {
      logError(`Failed to create dropdown (${sheetName} ${row}:${column || ''}): ${error.message}`);
      return false;
    }
  }
  
  /**
   * Creates a checkbox in a cell
   * * @param {string} sheetName - Name of the sheet
   * @param {number|string} row - Row number (1-based) or A1 notation
   * @param {number|string} column - Column number (1-based) or column letter
   * @param {boolean} [checked=false] - Whether the checkbox is initially checked
   * @return {boolean} True if successful
   */
  createCheckbox(sheetName, row, column, checked = false) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return false;
    
    try {
      let range;
      if (typeof row === "string" && column === undefined) {
           if (!row.match(/^[A-Z]+[0-9]+$/i)) throw new Error(`Invalid A1 notation: ${row}`);
           range = sheet.getRange(row);
      } else if (typeof row === 'number' && column !== undefined) {
          let colIndex;
          if (typeof column === "string" && column.match(/^[A-Z]+$/i)) {
              colIndex = columnToIndex(column.toUpperCase()) + 1;
          } else if (typeof column === 'number') {
              colIndex = column;
          } else {
              throw new Error(`Invalid column identifier: ${column}`);
          }
           if (row <= 0 || colIndex <= 0) throw new Error(`Invalid coordinates: ${row}, ${colIndex}`);
          range = sheet.getRange(row, colIndex);
      } else {
           throw new Error(`Invalid arguments for createCheckbox: row=${row}, column=${column}`);
      }
      
      // Create checkbox using data validation (more robust than insertCheckboxes sometimes)
      range.setDataValidation(SpreadsheetApp.newDataValidation()
          .requireCheckbox()
          .setAllowInvalid(false)
          .build());
      range.setValue(checked); // Set initial state
      
      return true;
    } catch (error) {
      logError(`Failed to create checkbox (${sheetName} ${row}:${column || ''}): ${error.message}`);
      return false;
    }
  }
  
  /**
   * Gets all sheet names in the spreadsheet
   * * @return {Array<string>} Array of sheet names, or empty array on error.
   */
  getSheetNames() {
    if (!this.spreadsheet) {
        logError("Spreadsheet not open. Cannot get sheet names.");
        return [];
    }
    try {
      const sheets = this.spreadsheet.getSheets();
      return sheets.map(sheet => sheet.getName());
    } catch (error) {
      logError(`Failed to get sheet names: ${error.message}`);
      return [];
    }
  }
  
  /**
   * Finds the first empty row in a sheet, starting from a specific column and row.
   * * @param {string} sheetName - Name of the sheet.
   * @param {number|string} column - Column to check (1-based or letter).
   * @param {number} [startRow=1] - Row to start checking from (1-based).
   * @return {number} First empty row number (1-based), or -1 on error.
   */
  findFirstEmptyRow(sheetName, column, startRow = 1) {
    const sheet = this.getSheet(sheetName, false);
    if (!sheet) return -1;
    
    try {
      let colIndex;
      if (typeof column === "string" && column.match(/^[A-Z]+$/i)) {
        colIndex = columnToIndex(column.toUpperCase()) + 1; // Convert to 1-based index
      } else if (typeof column === 'number' && column > 0) {
           colIndex = column;
      } else {
           throw new Error(`Invalid column identifier: ${column}`);
      }
      if (startRow <= 0) startRow = 1;

      const lastRow = sheet.getLastRow();
      
      // If sheet is effectively empty or startRow is beyond last row
      if (lastRow < startRow) {
        return startRow;
      }
      
      // Get all values in the column from startRow downwards
      const values = sheet.getRange(startRow, colIndex, lastRow - startRow + 1, 1).getValues();
      
      // Find first empty cell (considers null, undefined, empty string)
      for (let i = 0; i < values.length; i++) {
        if (values[i][0] === null || values[i][0] === undefined || values[i][0] === '') {
          return startRow + i; // Return the 1-based row index
        }
      }
      
      // If no empty rows found in the existing data, return the next row after the last
      return lastRow + 1;
    } catch (error) {
      logError(`Failed to find first empty row in ${sheetName}, column ${column}: ${error.message}`);
      return -1;
    }
  }
  
  /**
   * Creates a new tab by copying structure and formatting from a template tab.
   * * @param {string} newTabName - Name for the new tab.
   * @param {string} templateTabName - Name of the template tab to copy structure from.
   * @return {boolean} True if successful, false otherwise.
   */
  createTabFromTemplate(newTabName, templateTabName) {
    if (!this.spreadsheet) {
        logError("Spreadsheet not open. Cannot create tab from template.");
        return false;
    }
    try {
      const templateSheet = this.getSheet(templateTabName, false);
      if (!templateSheet) {
        logError(`Template sheet not found: ${templateTabName}`);
        return false;
      }
      
      // Check if sheet already exists
      if (this.getSheet(newTabName, false)) {
           logWarning(`Sheet '${newTabName}' already exists. Skipping creation from template.`);
           return true; // Or false depending on desired behavior
      }

      // Create a new sheet by copying the template
      const newSheet = templateSheet.copyTo(this.spreadsheet).setName(newTabName);
      
      // Optional: Clear specific data ranges if template contains placeholder data
      // Example: newSheet.getRange("E2:G28").clearContent(); // Clear element data values
      
      logInfo(`Created new tab '${newTabName}' from template '${templateTabName}'`);
      return true;
    } catch (error) {
      logError(`Failed to create tab '${newTabName}' from template '${templateTabName}': ${error.message}\n${error.stack}`);
      // Attempt cleanup: delete the partially created sheet if it exists
      const partialSheet = this.getSheet(newTabName, false);
      if (partialSheet) {
          try { this.spreadsheet.deleteSheet(partialSheet); } catch (e) {}
      }
      return false;
    }
  }

  // Removed copyMergedRanges and copyKnownMergedRanges as copyTo() handles merges.

  /**
   * Updates project info section in a project tab based on key-value pairs.
   * Keys should match the keys in SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.
   * * @param {string} sheetName - Name of the project sheet.
   * @param {Object} projectInfoUpdates - Object where keys are field names (e.g., "TITLE", "MODIFIED_AT") and values are the new data.
   * @return {boolean} True if successful, false otherwise.
   */
  updateProjectInfo(sheetName, projectInfoUpdates) {
    const sheet = this.getSheet(sheetName, false);
    if (!sheet) return false;
    
    try {
      const projectInfoFields = SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS;
      let success = true;
      
      // Update each field provided in the updates object
      for (const [fieldKey, newValue] of Object.entries(projectInfoUpdates)) {
        if (projectInfoFields.hasOwnProperty(fieldKey)) {
          const fieldInfo = projectInfoFields[fieldKey];
          if (!this.setCellValue(sheetName, fieldInfo.ROW, fieldInfo.VALUE_COL, newValue)) {
               logWarning(`Failed to update project info field '${fieldKey}' in ${sheetName}`);
               success = false; // Mark as partially failed but continue
          }
        } else {
             logWarning(`Unknown field key '${fieldKey}' provided in updateProjectInfo for ${sheetName}`);
        }
      }
      
      return success;
    } catch (error) {
      logError(`Failed to update project info in ${sheetName}: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Initializes a project index tab with headers and formatting.
   * * @param {string} [sheetName=SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME] - Name for the project index sheet.
   * @return {boolean} True if successful, false otherwise.
   */
  initializeProjectIndexTab(sheetName = SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME) {
    // Use createIfNotExist = true
    const sheet = this.getSheet(sheetName, true); 
    if (!sheet) return false;
    
    try {
      // Check if headers already exist
      if (sheet.getRange(1, 1).getValue() === SHEET_STRUCTURE.PROJECT_INDEX.HEADERS[0]) {
          logDebug(`Project index tab '${sheetName}' already initialized.`);
          return true;
      }

      // Set headers
      const headers = SHEET_STRUCTURE.PROJECT_INDEX.HEADERS;
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground(DEFAULT_COLORS.SECTION_HEADER_BG)
        .setFontWeight("bold");
      
      // Freeze header row
      sheet.setFrozenRows(1);
      
      // Auto-resize columns
      for (let i = 1; i <= headers.length; i++) {
        sheet.autoResizeColumn(i);
      }
      
      logInfo(`Initialized project index tab: ${sheetName}`);
      return true;
    } catch (error) {
      logError(`Failed to initialize project index tab '${sheetName}': ${error.message}`);
      return false;
    }
  }
  
  /**
   * Adds or updates a project entry in the project index.
   * * @param {Object} projectInfo - Project info object { projectId, title, createdAt, modifiedAt, lastAccessed? }. Timestamps should be Date objects.
   * @param {string} [sheetName=SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME] - Name of the project index sheet.
   * @return {boolean} True if successful, false otherwise.
   */
  addProjectToIndex(projectInfo, sheetName = SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME) {
    // Use createIfNotExist = true, and ensure it's initialized
    const sheet = this.getSheet(sheetName, true); 
    if (!sheet) return false;
    if (sheet.getLastRow() === 0) { // Ensure initialized if just created
        this.initializeProjectIndexTab(sheetName);
    }
    
    if (!projectInfo || !projectInfo.projectId || !projectInfo.title) {
         logError("Invalid projectInfo provided to addProjectToIndex.");
         return false;
    }

    try {
      const columns = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS;
      const now = new Date(); // Use for lastAccessed if not provided
      
      // Check if project already exists
      const projectLocation = this.findProjectInIndex(projectInfo.projectId, sheetName);
      
      if (projectLocation.found) {
        // Update existing entry
        const rowIndex = projectLocation.rowIndex;
        // Prepare update values in correct column order
        const updateValues = [];
        updateValues[columns.TITLE] = projectInfo.title;
        // Don't update createdAt
        updateValues[columns.MODIFIED_AT] = projectInfo.modifiedAt instanceof Date ? projectInfo.modifiedAt : now;
        updateValues[columns.LAST_ACCESSED] = projectInfo.lastAccessed instanceof Date ? projectInfo.lastAccessed : now;
        
        // Write updates (more efficient to getRange().setValues() if updating multiple adjacent cells)
        sheet.getRange(rowIndex, columns.TITLE + 1).setValue(updateValues[columns.TITLE]);
        sheet.getRange(rowIndex, columns.MODIFIED_AT + 1).setValue(updateValues[columns.MODIFIED_AT]);
        sheet.getRange(rowIndex, columns.LAST_ACCESSED + 1).setValue(updateValues[columns.LAST_ACCESSED]);

        logDebug(`Updated project in index: ${projectInfo.projectId}`);
        
      } else {
        // Create new entry - ensure order matches HEADERS/COLUMNS
        const newRow = [];
        newRow[columns.PROJECT_ID] = projectInfo.projectId;
        newRow[columns.TITLE] = projectInfo.title;
        newRow[columns.CREATED_AT] = projectInfo.createdAt instanceof Date ? projectInfo.createdAt : now;
        newRow[columns.MODIFIED_AT] = projectInfo.modifiedAt instanceof Date ? projectInfo.modifiedAt : now;
        newRow[columns.LAST_ACCESSED] = projectInfo.lastAccessed instanceof Date ? projectInfo.lastAccessed : now;
        
        // Append the row
        sheet.appendRow(newRow);
        logInfo(`Added project to index: ${projectInfo.projectId}`);
      }
      return true;
    } catch (error) {
      logError(`Failed to add/update project ${projectInfo.projectId} in index: ${error.message}\n${error.stack}`);
      return false;
    }
  }

   /**
   * Deletes a row from the project index based on Project ID.
   * * @param {string} projectId - The ID of the project to delete from the index.
   * @param {string} [sheetName=SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME] - Name of the project index sheet.
   * @return {boolean} True if the row was found and deleted, false otherwise.
   */
  deleteRowFromIndex(projectId, sheetName = SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME) {
      const sheet = this.getSheet(sheetName, false);
      if (!sheet) {
          logWarning(`Project index sheet '${sheetName}' not found for deletion.`);
          return false; // Sheet doesn't exist
      }

      try {
          const projectLocation = this.findProjectInIndex(projectId, sheetName);

          if (projectLocation.found) {
              sheet.deleteRow(projectLocation.rowIndex);
              logInfo(`Deleted project ${projectId} from index sheet '${sheetName}' (row ${projectLocation.rowIndex}).`);
              return true;
          } else {
              logWarning(`Project ${projectId} not found in index sheet '${sheetName}' for deletion.`);
              return false; // Project not found
          }
      } catch (error) {
          logError(`Failed to delete project ${projectId} from index sheet '${sheetName}': ${error.message}\n${error.stack}`);
          return false;
      }
  }
  
  /**
   * Finds a project in the project index by ID.
   * * @param {string} projectId - ID of the project to find.
   * @param {string} [sheetName=SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME] - Name of the project index sheet.
   * @return {Object} Object { found: boolean, rowIndex: number (1-based) or -1 }.
   */
  findProjectInIndex(projectId, sheetName = SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME) {
    const sheet = this.getSheet(sheetName, false);
    if (!sheet || sheet.getLastRow() < 1) { // Check if sheet exists and has rows
        return { found: false, rowIndex: -1 };
    }
    
    try {
      // Get only the Project ID column data for efficiency
      const idColumnIndex = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID + 1;
      const idValues = sheet.getRange(1, idColumnIndex, sheet.getLastRow(), 1).getValues();
      
      // Search for the ID (skip header row index 0)
      for (let i = 1; i < idValues.length; i++) {
        if (idValues[i][0] === projectId) {
          return { found: true, rowIndex: i + 1 }; // +1 for 1-based sheet index
        }
      }
      
      return { found: false, rowIndex: -1 }; // Not found
    } catch (error) {
      // Log error but return standard not found object
      logError(`Error finding project ${projectId} in index '${sheetName}': ${error.message}`);
      return { found: false, rowIndex: -1 };
    }
  }
  
  // Removed createSection and addElementColumn as this logic is better handled
  // by TemplateManager which understands the overall structure and where to insert/find things.
  // SheetAccessor should focus on primitive read/write/clear/format operations.
}
