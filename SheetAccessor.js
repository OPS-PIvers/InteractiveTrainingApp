/**
 * SheetAccessor class provides methods to interact with the Google Spreadsheet
 * Handles reading/writing data and managing sheet structure
 */
class SheetAccessor {
    /**
     * Creates a new SheetAccessor instance
     * 
     * @param {string} spreadsheetId - ID of the spreadsheet to access (optional)
     */
    constructor(spreadsheetId) {
      this.spreadsheetId = spreadsheetId || this.getActiveSpreadsheetId();
      this.spreadsheet = null;
      this.openSpreadsheet();
    }
    
    /**
     * Gets the ID of the active spreadsheet
     * 
     * @return {string} Spreadsheet ID
     */
    getActiveSpreadsheetId() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      return ss ? ss.getId() : null;
    }
    
    /**
     * Opens the spreadsheet with the stored ID
     */
    openSpreadsheet() {
      if (!this.spreadsheetId) {
        throw new Error("No spreadsheet ID provided or available");
      }
      
      try {
        this.spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
        logDebug(`Opened spreadsheet: ${this.spreadsheet.getName()}`);
      } catch (error) {
        logError(`Failed to open spreadsheet: ${error.message}`);
        throw error;
      }
    }
    
    /**
     * Creates a new spreadsheet with the specified name
     * 
     * @param {string} name - Name for the new spreadsheet
     * @return {SheetAccessor} A new SheetAccessor instance for the created spreadsheet
     */
    static createSpreadsheet(name) {
      try {
        const newSpreadsheet = SpreadsheetApp.create(name);
        logInfo(`Created new spreadsheet: ${name} (${newSpreadsheet.getId()})`);
        return new SheetAccessor(newSpreadsheet.getId());
      } catch (error) {
        logError(`Failed to create spreadsheet: ${error.message}`);
        throw error;
      }
    }
    
    /**
     * Gets a sheet by name, creating it if it doesn't exist
     * 
     * @param {string} sheetName - Name of the sheet
     * @param {boolean} createIfNotExist - Whether to create the sheet if it doesn't exist
     * @return {Sheet} Google Sheets Sheet object
     */
    getSheet(sheetName, createIfNotExist = true) {
      let sheet = this.spreadsheet.getSheetByName(sheetName);
      
      if (!sheet && createIfNotExist) {
        sheet = this.spreadsheet.insertSheet(sheetName);
        logDebug(`Created new sheet: ${sheetName}`);
      }
      
      return sheet;
    }
    
    /**
     * Deletes a sheet by name
     * 
     * @param {string} sheetName - Name of the sheet to delete
     * @return {boolean} True if successfully deleted
     */
    deleteSheet(sheetName) {
      const sheet = this.spreadsheet.getSheetByName(sheetName);
      
      if (sheet) {
        this.spreadsheet.deleteSheet(sheet);
        logDebug(`Deleted sheet: ${sheetName}`);
        return true;
      }
      
      logWarning(`Sheet not found for deletion: ${sheetName}`);
      return false;
    }
    
    /**
     * Gets cell value at specified location
     * 
     * @param {string} sheetName - Name of the sheet
     * @param {number|string} row - Row number (1-based) or A1 notation
     * @param {number|string} column - Column number (1-based) or column letter
     * @return {*} Cell value
     */
    getCellValue(sheetName, row, column) {
      const sheet = this.getSheet(sheetName, false);
      if (!sheet) return null;
      
      try {
        // Handle A1 notation
        if (typeof row === "string" && row.match(/^[A-Z]+[0-9]+$/)) {
          return sheet.getRange(row).getValue();
        }
        
        // Handle column as letter (A, B, C, etc.)
        if (typeof column === "string" && column.match(/^[A-Z]+$/)) {
          column = columnToIndex(column) + 1; // Convert to 1-based index
        }
        
        return sheet.getRange(row, column).getValue();
      } catch (error) {
        logError(`Failed to get cell value (${sheetName} ${row}:${column}): ${error.message}`);
        return null;
      }
    }
    
    /**
     * Sets a cell value at specified location
     * 
     * @param {string} sheetName - Name of the sheet
     * @param {number|string} row - Row number (1-based) or A1 notation
     * @param {number|string} column - Column number (1-based) or column letter
     * @param {*} value - Value to set
     * @return {boolean} True if successful
     */
    setCellValue(sheetName, row, column, value) {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;
      
      try {
        // Handle A1 notation
        if (typeof row === "string" && row.match(/^[A-Z]+[0-9]+$/)) {
          sheet.getRange(row).setValue(value);
          return true;
        }
        
        // Handle column as letter (A, B, C, etc.)
        if (typeof column === "string" && column.match(/^[A-Z]+$/)) {
          column = columnToIndex(column) + 1; // Convert to 1-based index
        }
        
        sheet.getRange(row, column).setValue(value);
        return true;
      } catch (error) {
        logError(`Failed to set cell value (${sheetName} ${row}:${column}): ${error.message}`);
        return false;
      }
    }
    
    /**
     * Gets values from a range
     * 
     * @param {string} sheetName - Name of the sheet
     * @param {number} startRow - Starting row (1-based)
     * @param {number|string} startCol - Starting column (1-based) or column letter
     * @param {number} numRows - Number of rows
     * @param {number} numCols - Number of columns
     * @return {Array<Array<*>>} 2D array of values
     */
    getRangeValues(sheetName, startRow, startCol, numRows, numCols) {
      const sheet = this.getSheet(sheetName, false);
      if (!sheet) return null;
      
      try {
        // Handle column as letter (A, B, C, etc.)
        if (typeof startCol === "string" && startCol.match(/^[A-Z]+$/)) {
          startCol = columnToIndex(startCol) + 1; // Convert to 1-based index
        }
        
        return sheet.getRange(startRow, startCol, numRows, numCols).getValues();
      } catch (error) {
        logError(`Failed to get range values: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Sets values in a range
     * 
     * @param {string} sheetName - Name of the sheet
     * @param {number} startRow - Starting row (1-based)
     * @param {number|string} startCol - Starting column (1-based) or column letter
     * @param {Array<Array<*>>} values - 2D array of values
     * @return {boolean} True if successful
     */
    setRangeValues(sheetName, startRow, startCol, values) {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;
      
      try {
        // Handle column as letter (A, B, C, etc.)
        if (typeof startCol === "string" && startCol.match(/^[A-Z]+$/)) {
          startCol = columnToIndex(startCol) + 1; // Convert to 1-based index
        }
        
        const numRows = values.length;
        const numCols = values[0].length;
        
        sheet.getRange(startRow, startCol, numRows, numCols).setValues(values);
        return true;
      } catch (error) {
        logError(`Failed to set range values: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Gets all data from a sheet
     * 
     * @param {string} sheetName - Name of the sheet
     * @return {Array<Array<*>>} 2D array of all data
     */
    getSheetData(sheetName) {
      const sheet = this.getSheet(sheetName, false);
      if (!sheet) return [];
      
      try {
        const dataRange = sheet.getDataRange();
        return dataRange.getValues();
      } catch (error) {
        logError(`Failed to get sheet data: ${error.message}`);
        return [];
      }
    }
    
    /**
     * Appends a row to a sheet
     * 
     * @param {string} sheetName - Name of the sheet
     * @param {Array<*>} rowData - Array of values for the row
     * @return {boolean} True if successful
     */
    appendRow(sheetName, rowData) {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;
      
      try {
        sheet.appendRow(rowData);
        return true;
      } catch (error) {
        logError(`Failed to append row: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Creates a section header with merged cells and formatting
     * 
     * @param {string} sheetName - Name of the sheet
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
        if (typeof startCol === "string" && startCol.match(/^[A-Z]+$/)) {
          startCol = columnToIndex(startCol) + 1; // Convert to 1-based index
        }
        
        if (typeof endCol === "string" && endCol.match(/^[A-Z]+$/)) {
          endCol = columnToIndex(endCol) + 1; // Convert to 1-based index
        }
        
        // Merge cells for header
        const headerRange = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
        headerRange.merge();
        
        // Set header text and formatting
        headerRange.setValue(headerText);
        headerRange.setBackground(DEFAULT_COLORS.SECTION_HEADER_BG);
        headerRange.setFontWeight("bold");
        headerRange.setHorizontalAlignment("center");
        
        return true;
      } catch (error) {
        logError(`Failed to create section header: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Creates a dropdown in a cell with validation
     * 
     * @param {string} sheetName - Name of the sheet
     * @param {number|string} row - Row number (1-based) or A1 notation
     * @param {number|string} column - Column number (1-based) or column letter
     * @param {Array<string>} options - Array of dropdown options
     * @return {boolean} True if successful
     */
    createDropdown(sheetName, row, column, options) {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;
      
      try {
        // Handle A1 notation
        let range;
        if (typeof row === "string" && row.match(/^[A-Z]+[0-9]+$/)) {
          range = sheet.getRange(row);
        } else {
          // Handle column as letter (A, B, C, etc.)
          if (typeof column === "string" && column.match(/^[A-Z]+$/)) {
            column = columnToIndex(column) + 1; // Convert to 1-based index
          }
          
          range = sheet.getRange(row, column);
        }
        
        // Create validation rule
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(options, true)
          .build();
        
        range.setDataValidation(rule);
        return true;
      } catch (error) {
        logError(`Failed to create dropdown: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Creates a checkbox in a cell
     * 
     * @param {string} sheetName - Name of the sheet
     * @param {number|string} row - Row number (1-based) or A1 notation
     * @param {number|string} column - Column number (1-based) or column letter
     * @param {boolean} checked - Whether the checkbox is initially checked
     * @return {boolean} True if successful
     */
    createCheckbox(sheetName, row, column, checked = false) {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;
      
      try {
        // Handle A1 notation
        let range;
        if (typeof row === "string" && row.match(/^[A-Z]+[0-9]+$/)) {
          range = sheet.getRange(row);
        } else {
          // Handle column as letter (A, B, C, etc.)
          if (typeof column === "string" && column.match(/^[A-Z]+$/)) {
            column = columnToIndex(column) + 1; // Convert to 1-based index
          }
          
          range = sheet.getRange(row, column);
        }
        
        // Create checkbox
        range.insertCheckboxes();
        
        // Set initial state
        if (checked) {
          range.setValue(true);
        }
        
        return true;
      } catch (error) {
        logError(`Failed to create checkbox: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Gets all sheet names in the spreadsheet
     * 
     * @return {Array<string>} Array of sheet names
     */
    getSheetNames() {
      try {
        const sheets = this.spreadsheet.getSheets();
        return sheets.map(sheet => sheet.getName());
      } catch (error) {
        logError(`Failed to get sheet names: ${error.message}`);
        return [];
      }
    }
    
    /**
     * Finds the first empty row in a sheet, starting from a specific column
     * 
     * @param {string} sheetName - Name of the sheet
     * @param {number|string} column - Column to check (1-based or letter)
     * @param {number} startRow - Row to start checking from (1-based)
     * @return {number} First empty row (1-based), or -1 if error
     */
    findFirstEmptyRow(sheetName, column, startRow = 1) {
      const sheet = this.getSheet(sheetName, false);
      if (!sheet) return -1;
      
      try {
        // Handle column as letter (A, B, C, etc.)
        if (typeof column === "string" && column.match(/^[A-Z]+$/)) {
          column = columnToIndex(column) + 1; // Convert to 1-based index
        }
        
        const lastRow = sheet.getLastRow();
        
        // If sheet is completely empty
        if (lastRow < startRow) {
          return startRow;
        }
        
        // Get all values in the column from startRow to lastRow
        const values = sheet.getRange(startRow, column, lastRow - startRow + 1, 1).getValues();
        
        // Find first empty row
        for (let i = 0; i < values.length; i++) {
          if (isEmpty(values[i][0])) {
            return startRow + i;
          }
        }
        
        // If all rows have values, return the next row after the last
        return lastRow + 1;
      } catch (error) {
        logError(`Failed to find first empty row: ${error.message}`);
        return -1;
      }
    }
    
    /**
     * Creates a new tab with the structure from the template tab
     * 
     * @param {string} newTabName - Name for the new tab
     * @param {string} templateTabName - Name of the template tab to copy structure from
     * @return {boolean} True if successful
     */
    createTabFromTemplate(newTabName, templateTabName) {
      try {
        const templateSheet = this.getSheet(templateTabName, false);
        if (!templateSheet) {
          logError(`Template sheet not found: ${templateTabName}`);
          return false;
        }
        
        // Create a new sheet
        const newSheet = this.spreadsheet.insertSheet(newTabName);
        
        // Copy template data
        const templateRange = templateSheet.getDataRange();
        const templateValues = templateRange.getValues();
        const templateFormats = templateRange.getNumberFormats();
        const templateValidations = templateRange.getDataValidations();
        const templateBackgrounds = templateRange.getBackgrounds();
        const templateFontColors = templateRange.getFontColors();
        const templateFontWeights = templateRange.getFontWeights();
        
        const numRows = templateValues.length;
        const numCols = templateValues[0].length;
        
        if (numRows > 0 && numCols > 0) {
          const newRange = newSheet.getRange(1, 1, numRows, numCols);
          
          // Set values and formatting
          newRange.setValues(templateValues);
          newRange.setNumberFormats(templateFormats);
          newRange.setDataValidations(templateValidations);
          newRange.setBackgrounds(templateBackgrounds);
          newRange.setFontColors(templateFontColors);
          newRange.setFontWeights(templateFontWeights);
          
          // Copy merged ranges
          const mergedRanges = templateSheet.getMergedRanges();
          for (const mergedRange of mergedRanges) {
            const startRow = mergedRange.getRow();
            const startCol = mergedRange.getColumn();
            const numRows = mergedRange.getNumRows();
            const numCols = mergedRange.getNumColumns();
            
            newSheet.getRange(startRow, startCol, numRows, numCols).merge();
          }
        }
        
        // Copy column widths
        for (let col = 1; col <= templateSheet.getMaxColumns(); col++) {
          const width = templateSheet.getColumnWidth(col);
          newSheet.setColumnWidth(col, width);
        }
        
        // Copy row heights
        for (let row = 1; row <= templateSheet.getMaxRows(); row++) {
          const height = templateSheet.getRowHeight(row);
          newSheet.setRowHeight(row, height);
        }
        
        logInfo(`Created new tab from template: ${newTabName}`);
        return true;
      } catch (error) {
        logError(`Failed to create tab from template: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Updates project info section in a project tab
     * 
     * @param {string} sheetName - Name of the project sheet
     * @param {Object} projectInfo - Project info object with keys matching SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS
     * @return {boolean} True if successful
     */
    updateProjectInfo(sheetName, projectInfo) {
      const sheet = this.getSheet(sheetName, false);
      if (!sheet) return false;
      
      try {
        const projectInfoFields = SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS;
        
        // Update each field in the project info section
        for (const field in projectInfoFields) {
          if (projectInfo.hasOwnProperty(field) && projectInfoFields.hasOwnProperty(field)) {
            const fieldInfo = projectInfoFields[field];
            const row = fieldInfo.ROW;
            const col = fieldInfo.VALUE_COL;
            
            this.setCellValue(sheetName, row, col, projectInfo[field]);
          }
        }
        
        return true;
      } catch (error) {
        logError(`Failed to update project info: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Initializes a project index tab with headers
     * 
     * @param {string} sheetName - Name for the project index sheet (default: ProjectIndex)
     * @return {boolean} True if successful
     */
    initializeProjectIndexTab(sheetName = SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME) {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;
      
      try {
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
        logError(`Failed to initialize project index tab: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Adds a project entry to the project index
     * 
     * @param {Object} projectInfo - Project info with projectId, title, createdAt, modifiedAt
     * @param {string} sheetName - Name of the project index sheet (default: ProjectIndex)
     * @return {boolean} True if successful
     */
    addProjectToIndex(projectInfo, sheetName = SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME) {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;
      
      try {
        const columns = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS;
        
        // Check if project already exists
        const projectExists = this.findProjectInIndex(projectInfo.projectId, sheetName);
        if (projectExists.found) {
          // Update existing entry
          const rowIndex = projectExists.rowIndex;
          sheet.getRange(rowIndex, columns.TITLE + 1).setValue(projectInfo.title);
          sheet.getRange(rowIndex, columns.MODIFIED_AT + 1).setValue(projectInfo.modifiedAt);
          sheet.getRange(rowIndex, columns.LAST_ACCESSED + 1).setValue(new Date());
          
          logDebug(`Updated project in index: ${projectInfo.projectId}`);
          return true;
        }
        
        // Create new entry
        const newRow = [
          projectInfo.projectId,
          projectInfo.title,
          projectInfo.createdAt,
          projectInfo.modifiedAt,
          new Date()
        ];
        
        sheet.appendRow(newRow);
        logInfo(`Added project to index: ${projectInfo.projectId}`);
        return true;
      } catch (error) {
        logError(`Failed to add project to index: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Finds a project in the project index by ID
     * 
     * @param {string} projectId - ID of the project to find
     * @param {string} sheetName - Name of the project index sheet (default: ProjectIndex)
     * @return {Object} Object with found (boolean) and rowIndex (number) properties
     */
    findProjectInIndex(projectId, sheetName = SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME) {
      const sheet = this.getSheet(sheetName, false);
      if (!sheet) return { found: false, rowIndex: -1 };
      
      try {
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        
        // Skip header row
        for (let i = 1; i < values.length; i++) {
          if (values[i][SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID] === projectId) {
            return { found: true, rowIndex: i + 1 }; // +1 because array is 0-based, but sheets are 1-based
          }
        }
        
        return { found: false, rowIndex: -1 };
      } catch (error) {
        logError(`Failed to find project in index: ${error.message}`);
        return { found: false, rowIndex: -1 };
      }
    }
    
    /**
     * Creates a dynamic section in a project tab based on the sheet structure
     * 
     * @param {string} sheetName - Name of the project sheet
     * @param {string} sectionName - Name of the section to create (e.g., "SLIDE_INFO", "ELEMENT_INFO")
     * @param {number} instanceNumber - Instance of the section (for multiple slides, elements, etc.)
     * @param {Object} initial - Initial values for the section fields
     * @return {boolean} True if successful
     */
    createSection(sheetName, sectionName, instanceNumber = 1, initial = {}) {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;
      
      try {
        const sectionStructure = SHEET_STRUCTURE.PROJECT_TAB[sectionName];
        if (!sectionStructure) {
          logError(`Section structure not found: ${sectionName}`);
          return false;
        }
        
        // Determine section location
        let startRow = sectionStructure.SECTION_START_ROW;
        let startCol = sectionStructure.SECTION_START_COL || "A";
        let endCol = sectionStructure.SECTION_END_COL || "B";
        
        // For sections like SLIDE_INFO that can have multiple instances
        if (instanceNumber > 1 && sectionName === "SLIDE_INFO") {
          // Calculate row offset for additional slides
          const rowsPerSlide = sectionStructure.SECTION_END_ROW - sectionStructure.SECTION_START_ROW + 1;
          startRow += (instanceNumber - 1) * rowsPerSlide;
        }
        
        // Create header with instance number if applicable
        let headerText = sectionStructure.HEADER;
        if (sectionName === "SLIDE_INFO") {
          headerText = headerText.replace("1", instanceNumber);
        }
        
        // Create the section header
        this.createSectionHeader(sheetName, startRow, startCol, endCol, headerText);
        
        // Set field labels and initial values
        const fields = sectionStructure.FIELDS;
        for (const field in fields) {
          if (fields.hasOwnProperty(field)) {
            const fieldInfo = fields[field];
            let fieldRow = fieldInfo.ROW;
            const labelCol = fieldInfo.LABEL_COL || startCol;
            const valueCol = fieldInfo.VALUE_COL || (typeof labelCol === "string" ? String.fromCharCode(labelCol.charCodeAt(0) + 1) : labelCol + 1);
            
            // Adjust row for multiple instances
            if (instanceNumber > 1 && sectionName === "SLIDE_INFO") {
              fieldRow += (instanceNumber - 1) * (sectionStructure.SECTION_END_ROW - sectionStructure.SECTION_START_ROW + 1);
            }
            
            // Set field label
            this.setCellValue(sheetName, fieldRow, labelCol, field.replace(/_/g, ' '));
            
            // Set initial value if provided
            if (initial.hasOwnProperty(field)) {
              this.setCellValue(sheetName, fieldRow, valueCol, initial[field]);
            }
            
            // Add validation where needed
            if (field === "FILE_TYPE" && sectionName === "SLIDE_INFO") {
              this.createDropdown(sheetName, fieldRow, valueCol, Object.values(FILE_TYPES));
            } else if (field === "SHOW_CONTROLS" && sectionName === "SLIDE_INFO") {
              this.createCheckbox(sheetName, fieldRow, valueCol, false);
            } else if (field === "TYPE" && sectionName === "ELEMENT_INFO") {
              this.createDropdown(sheetName, fieldRow, valueCol, Object.values(ELEMENT_TYPES).map(type => type.LABEL));
            } else if (field === "INITIALLY_HIDDEN" && sectionName === "ELEMENT_INFO") {
              this.createCheckbox(sheetName, fieldRow, valueCol, false);
            } else if (field === "OUTLINE" && sectionName === "ELEMENT_INFO") {
              this.createCheckbox(sheetName, fieldRow, valueCol, false);
            } else if (field === "SHADOW" && sectionName === "ELEMENT_INFO") {
              this.createCheckbox(sheetName, fieldRow, valueCol, false);
            } else if (field === "FONT" && sectionName === "ELEMENT_INFO") {
              this.createDropdown(sheetName, fieldRow, valueCol, FONTS);
            } else if (field === "TRIGGERS" && sectionName === "ELEMENT_INFO") {
              this.createDropdown(sheetName, fieldRow, valueCol, Object.values(TRIGGER_TYPES));
            } else if (field === "INTERACTION_TYPE" && sectionName === "ELEMENT_INFO") {
              this.createDropdown(sheetName, fieldRow, valueCol, Object.values(INTERACTION_TYPES));
            } else if (field === "TEXT_MODAL" && sectionName === "ELEMENT_INFO") {
              this.createCheckbox(sheetName, fieldRow, valueCol, false);
            } else if (field === "ANIMATION_TYPE" && sectionName === "ELEMENT_INFO") {
              this.createDropdown(sheetName, fieldRow, valueCol, Object.values(ANIMATION_TYPES));
            } else if (field === "ANIMATION_SPEED" && sectionName === "ELEMENT_INFO") {
              this.createDropdown(sheetName, fieldRow, valueCol, Object.values(ANIMATION_SPEEDS));
            } else if (field === "ANIMATION_IN" && sectionName === "TIMELINE") {
              this.createDropdown(sheetName, fieldRow, valueCol, Object.values(ANIMATION_IN_TYPES));
            } else if (field === "ANIMATION_OUT" && sectionName === "TIMELINE") {
              this.createDropdown(sheetName, fieldRow, valueCol, Object.values(ANIMATION_OUT_TYPES));
            } else if (field === "QUESTION_TYPE" && sectionName === "QUIZ") {
              this.createDropdown(sheetName, fieldRow, valueCol, Object.values(QUESTION_TYPES));
            } else if (field === "INCLUDE_FEEDBACK" && sectionName === "QUIZ") {
              this.createCheckbox(sheetName, fieldRow, valueCol, false);
            } else if (field === "PROVIDE_CORRECT_ANSWER" && sectionName === "QUIZ") {
              this.createCheckbox(sheetName, fieldRow, valueCol, false);
            } else if (field === "TRACK_COMPLETION" && sectionName === "USER_TRACKING") {
              this.createCheckbox(sheetName, fieldRow, valueCol, true);
            } else if (field === "REQUIRE_QUIZ_COMPLETION" && sectionName === "USER_TRACKING") {
              this.createCheckbox(sheetName, fieldRow, valueCol, false);
            } else if (field === "SEND_COMPLETION_EMAIL" && sectionName === "USER_TRACKING") {
              this.createCheckbox(sheetName, fieldRow, valueCol, false);
            }
          }
        }
        
        logInfo(`Created ${sectionName} section (instance ${instanceNumber}) in ${sheetName}`);
        return true;
      } catch (error) {
        logError(`Failed to create section ${sectionName}: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Adds an element column to the element info section
     * 
     * @param {string} sheetName - Name of the project sheet
     * @param {number} elementIndex - Index of the element (1-based)
     * @param {Object} initialValues - Initial values for the element
     * @return {boolean} True if successful
     */
    addElementColumn(sheetName, elementIndex, initialValues = {}) {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;
      
      try {
        const sectionStructure = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO;
        
        // Calculate column based on element index (starting from E for first element)
        const startCol = columnToIndex("E") + (elementIndex - 1);
        const colLetter = indexToColumn(startCol);
        
        // Set element header
        this.setCellValue(sheetName, sectionStructure.ELEMENT_HEADERS_ROW, colLetter, `Element ${elementIndex}`);
        
        // Set initial values and create controls
        const fields = sectionStructure.FIELDS;
        for (const field in fields) {
          if (fields.hasOwnProperty(field)) {
            const fieldInfo = fields[field];
            const row = fieldInfo.ROW;
            
            // If there's an initial value, set it
            if (initialValues.hasOwnProperty(field)) {
              this.setCellValue(sheetName, row, colLetter, initialValues[field]);
            }
            
            // Add validation where needed
            if (field === "TYPE") {
              this.createDropdown(sheetName, row, colLetter, Object.values(ELEMENT_TYPES).map(type => type.LABEL));
            } else if (field === "INITIALLY_HIDDEN" || field === "OUTLINE" || field === "SHADOW" || field === "TEXT_MODAL") {
              this.createCheckbox(sheetName, row, colLetter, false);
            } else if (field === "FONT") {
              this.createDropdown(sheetName, row, colLetter, FONTS);
            } else if (field === "TRIGGERS") {
              this.createDropdown(sheetName, row, colLetter, Object.values(TRIGGER_TYPES));
            } else if (field === "INTERACTION_TYPE") {
              this.createDropdown(sheetName, row, colLetter, Object.values(INTERACTION_TYPES));
            } else if (field === "ANIMATION_TYPE") {
              this.createDropdown(sheetName, row, colLetter, Object.values(ANIMATION_TYPES));
            } else if (field === "ANIMATION_SPEED") {
              this.createDropdown(sheetName, row, colLetter, Object.values(ANIMATION_SPEEDS));
            }
          }
        }
        
        logInfo(`Added element column ${elementIndex} in ${sheetName}`);
        return true;
      } catch (error) {
        logError(`Failed to add element column: ${error.message}`);
        return false;
      }
    }
  }