/**
 * TemplateManager class handles the creation and management of project templates
 * and the structure of individual project sheets based on the template.
 * This includes creating the base template, new project sheets, and updating
 * project sheets with data (like slides and elements).
 */
class TemplateManager {
  /**
   * Creates a new TemplateManager instance
   * * @param {SheetAccessor} sheetAccessor - SheetAccessor instance
   */
  constructor(sheetAccessor) {
      if (!sheetAccessor) {
          throw new Error("SheetAccessor instance is required for TemplateManager.");
      }
      this.sheetAccessor = sheetAccessor;
  }

  /**
   * Creates the base template sheet with all defined sections.
   * * @param {string} [templateName="Template"] - Name for the template tab.
   * @return {boolean} True if successful, false otherwise.
   */
  createBaseTemplate(templateName = "Template") {
      try {
          logInfo(`Attempting to create or verify base template: ${templateName}`);
          // Create template sheet if it doesn't exist
          const sheet = this.sheetAccessor.getSheet(templateName, true); // Ensure sheet exists
          if (!sheet) {
              throw new Error(`Failed to get or create template sheet: ${templateName}`);
          }

          // Check if already initialized (e.g., by checking PROJECT_INFO header)
          const projectInfoHeader = this.sheetAccessor.getCellValue(templateName, 
              SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.SECTION_START_ROW, 
              SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_ID.LABEL_COL); // Check label cell
          
          if (projectInfoHeader === SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_ID.LABEL_COL) { // Simple check if label exists
               logDebug(`Base template '${templateName}' seems already initialized.`);
               // Optionally, could add logic here to verify/update existing template sections
               return true;
          }

          // Create sections defined in SHEET_STRUCTURE
          const sections = ["PROJECT_INFO", "SLIDE_INFO", "ELEMENT_INFO", "TIMELINE", "QUIZ", "USER_TRACKING"];
          for (const sectionName of sections) {
              if (!SHEET_STRUCTURE.PROJECT_TAB[sectionName]) {
                  logWarning(`Structure definition missing for section: ${sectionName}. Skipping.`);
                  continue;
              }
              // Use the createSection method which handles headers, labels, and basic formatting/validation
              const success = this._createSheetSectionStructure(templateName, sectionName);
               if (!success) {
                   logWarning(`Failed to create section structure for ${sectionName} in template ${templateName}.`);
                   // Decide whether to continue or fail completely
               }
          }

          // Set default column widths (adjust as needed)
          sheet.setColumnWidth(columnToIndex("A") + 1, 150); // Column A (Labels)
          sheet.setColumnWidth(columnToIndex("B") + 1, 300); // Column B (Values)
          sheet.setColumnWidth(columnToIndex("C") + 1, 20);  // Column C (Spacer)
          sheet.setColumnWidth(columnToIndex("D") + 1, 150); // Column D (Element Labels)
          // Set widths for a few default element columns
          for (let i = 0; i < 5; i++) {
               sheet.setColumnWidth(columnToIndex("E") + 1 + i, 150); // Columns E, F, G, H, I
          }


          logInfo(`Successfully created/verified base template: ${templateName}`);
          return true;
      } catch (error) {
          logError(`Failed to create base template '${templateName}': ${error.message}\n${error.stack}`);
          return false;
      }
  }

   /**
   * Helper to create the structure (headers, labels, controls) for a specific section in a sheet.
   * * @param {string} sheetName - The name of the sheet.
   * @param {string} sectionName - The key of the section in SHEET_STRUCTURE (e.g., "SLIDE_INFO").
   * @return {boolean} True if successful.
   */
  _createSheetSectionStructure(sheetName, sectionName) {
      try {
          const sectionStructure = SHEET_STRUCTURE.PROJECT_TAB[sectionName];
          if (!sectionStructure) throw new Error(`Structure definition not found for section: ${sectionName}`);

          const startRow = sectionStructure.SECTION_START_ROW;
          const startCol = sectionStructure.SECTION_START_COL || "A"; // Default to A if not column-based
          const endCol = sectionStructure.SECTION_END_COL || "B";     // Default to B
          const headerText = sectionStructure.HEADER;

          // Create the main section header
          this.sheetAccessor.createSectionHeader(sheetName, startRow, startCol, endCol, headerText);

          // Set field labels and create controls (dropdowns, checkboxes)
          const fields = sectionStructure.FIELDS;
          for (const [fieldKey, fieldInfo] of Object.entries(fields)) {
              const fieldRow = fieldInfo.ROW;
              const labelCol = fieldInfo.LABEL_COL || startCol; // Column for the label text
              const valueCol = fieldInfo.VALUE_COL || (typeof labelCol === "string" ? String.fromCharCode(labelCol.charCodeAt(0) + 1) : labelCol + 1); // Column for the value/control

              // Set field label text
              // Convert key like PROJECT_ID to "Project Id" or similar for display
              const labelText = fieldKey.split('_').map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join(' ');
              this.sheetAccessor.setCellValue(sheetName, fieldRow, labelCol, labelText);

              // Add validation/controls based on section and field key
              this._applyFieldControl(sheetName, sectionName, fieldKey, fieldRow, valueCol);
          }

           // Special handling for ELEMENT_INFO headers ("Element 1", "Element 2", etc.)
           if (sectionName === "ELEMENT_INFO" && sectionStructure.ELEMENT_HEADERS_ROW) {
               // Add a few default element headers
               const headerRow = sectionStructure.ELEMENT_HEADERS_ROW;
               const firstElementCol = columnToIndex(sectionStructure.SECTION_START_COL) + 2; // E is the first element column
               for (let i = 0; i < 3; i++) { // Add headers for Element 1, 2, 3
                   this.sheetAccessor.setCellValue(sheetName, headerRow, firstElementCol + i, `Element ${i + 1}`);
                   // Optionally format header
                   this.sheetAccessor.getSheet(sheetName).getRange(headerRow, firstElementCol + i).setFontWeight("bold").setHorizontalAlignment("center");
               }
           }

          return true;
      } catch (error) {
          logError(`Error creating structure for section ${sectionName} in ${sheetName}: ${error.message}`);
          return false;
      }
  }

   /**
   * Applies specific controls (dropdown, checkbox) to a field cell based on config.
   * * @param {string} sheetName
   * @param {string} sectionName
   * @param {string} fieldKey (UPPER_SNAKE_CASE from SHEET_STRUCTURE)
   * @param {number} row
   * @param {number|string} valueCol
   */
  _applyFieldControl(sheetName, sectionName, fieldKey, row, valueCol) {
      // Use switch for clarity based on section and field
      switch (`${sectionName}.${fieldKey}`) {
          // SLIDE_INFO Controls
          case "SLIDE_INFO.FILE_TYPE":
              this.sheetAccessor.createDropdown(sheetName, row, valueCol, Object.values(FILE_TYPES));
              break;
          case "SLIDE_INFO.SHOW_CONTROLS":
              this.sheetAccessor.createCheckbox(sheetName, row, valueCol, false);
              break;

          // ELEMENT_INFO Controls
          case "ELEMENT_INFO.TYPE":
              this.sheetAccessor.createDropdown(sheetName, row, valueCol, Object.values(ELEMENT_TYPES).map(type => type.LABEL));
              break;
          case "ELEMENT_INFO.INITIALLY_HIDDEN":
          case "ELEMENT_INFO.OUTLINE":
          case "ELEMENT_INFO.SHADOW":
          case "ELEMENT_INFO.TEXT_MODAL":
              this.sheetAccessor.createCheckbox(sheetName, row, valueCol, false);
              break;
          case "ELEMENT_INFO.FONT":
              this.sheetAccessor.createDropdown(sheetName, row, valueCol, FONTS);
              break;
          case "ELEMENT_INFO.TRIGGERS":
              this.sheetAccessor.createDropdown(sheetName, row, valueCol, Object.values(TRIGGER_TYPES));
              break;
          case "ELEMENT_INFO.INTERACTION_TYPE": // Renamed from TYPE in Config for clarity? Assuming this matches config.
              this.sheetAccessor.createDropdown(sheetName, row, valueCol, Object.values(INTERACTION_TYPES));
              break;
          case "ELEMENT_INFO.ANIMATION_TYPE":
              this.sheetAccessor.createDropdown(sheetName, row, valueCol, Object.values(ANIMATION_TYPES));
              break;
          case "ELEMENT_INFO.ANIMATION_SPEED":
              this.sheetAccessor.createDropdown(sheetName, row, valueCol, Object.values(ANIMATION_SPEEDS));
              break;

          // TIMELINE Controls
           case "TIMELINE.PAUSE_AT":
           case "TIMELINE.SHOW_FOR_DURATION":
               this.sheetAccessor.createCheckbox(sheetName, row, valueCol, false);
               break;
          case "TIMELINE.ANIMATION_IN":
              this.sheetAccessor.createDropdown(sheetName, row, valueCol, Object.values(ANIMATION_IN_TYPES));
              break;
          case "TIMELINE.ANIMATION_OUT":
              this.sheetAccessor.createDropdown(sheetName, row, valueCol, Object.values(ANIMATION_OUT_TYPES));
              break;

          // QUIZ Controls
          case "QUIZ.QUESTION_TYPE":
              this.sheetAccessor.createDropdown(sheetName, row, valueCol, Object.values(QUESTION_TYPES));
              break;
          case "QUIZ.INCLUDE_FEEDBACK":
          case "QUIZ.PROVIDE_CORRECT_ANSWER":
              this.sheetAccessor.createCheckbox(sheetName, row, valueCol, false);
              break;

          // USER_TRACKING Controls
          case "USER_TRACKING.TRACK_COMPLETION":
          case "USER_TRACKING.REQUIRE_QUIZ_COMPLETION":
          case "USER_TRACKING.SEND_COMPLETION_EMAIL":
              this.sheetAccessor.createCheckbox(sheetName, row, valueCol, false);
              break;
      }
  }

  /**
   * Creates a new project tab from the template and initializes it.
   * * @param {string} projectName - Name for the new project.
   * @param {string} [templateName="Template"] - Name of the template tab.
   * @return {string|null} Name of the created project tab, or null if failed.
   */
  createProjectFromTemplate(projectName, templateName = "Template") {
      try {
          // Sanitize project name for use as sheet name
          const projectTabName = sanitizeSheetName(projectName);
          if (!projectTabName) throw new Error("Invalid project name after sanitization.");

          // Create project tab by copying the template
          const success = this.sheetAccessor.createTabFromTemplate(projectTabName, templateName);
          if (!success) {
              throw new Error(`Failed to create project tab from template: ${projectTabName}`);
          }

          // Initialize project with default values
          const projectId = generateUUID();
          const now = new Date();
          const nowTimestamp = now.getTime();

          // Set PROJECT INFO section values
          const projectInfo = {
              PROJECT_ID: projectId,
              PROJECT_WEB_APP_URL: "", // Will be set after web app deployment
              TITLE: projectName, // Use original name here
              CREATED_AT: nowTimestamp,
              MODIFIED_AT: nowTimestamp,
              PROJECT_FOLDER_ID: "" // Will be set after folder creation by ProjectManager
          };
          // Use SheetAccessor to update these specific cells
          this.sheetAccessor.updateProjectInfo(projectTabName, projectInfo);

          // Initialize the first slide section (assuming template has one)
          const firstSlideInfo = {
               SLIDE_ID: generateUUID(),
               SLIDE_TITLE: "Slide 1",
               BACKGROUND_COLOR: DEFAULT_COLORS.BACKGROUND,
               FILE_TYPE: "",
               FILE_URL: "",
               SLIDE_NUMBER: 1,
               SHOW_CONTROLS: false
          };
          this.updateSlideSection(projectTabName, null, firstSlideInfo, 1); // Update based on instance number 1

          logInfo(`Created project sheet from template: ${projectTabName} (ID: ${projectId})`);
          return projectTabName; // Return the tab name, ProjectManager adds to index
      } catch (error) {
          logError(`Failed to create project from template: ${error.message}\n${error.stack}`);
          // Attempt cleanup if tab was created
          this.sheetAccessor.deleteSheet(sanitizeSheetName(projectName));
          return null;
      }
  }

  /**
   * Gets project info from a project tab.
   * * @param {string} projectTabName - Name of the project tab.
   * @return {Object} Project info object (keys are UPPER_SNAKE_CASE). Returns empty object on error.
   */
  getProjectInfo(projectTabName) {
      try {
          const projectInfo = {};
          const projectInfoFields = SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS;

          for (const [fieldKey, fieldInfo] of Object.entries(projectInfoFields)) {
              const value = this.sheetAccessor.getCellValue(
                  projectTabName,
                  fieldInfo.ROW,
                  fieldInfo.VALUE_COL // Column B for values
              );
              // Convert timestamp strings/numbers back to numbers if needed
              if (['CREATED_AT', 'MODIFIED_AT'].includes(fieldKey) && value) {
                  projectInfo[fieldKey] = Number(value);
              } else {
                   projectInfo[fieldKey] = value;
              }
          }
          // Add the tab name itself for reference if needed elsewhere
          projectInfo.projectTabName = projectTabName; 
          return projectInfo;
      } catch (error) {
          logError(`Failed to get project info for tab ${projectTabName}: ${error.message}`);
          return {};
      }
  }

  // --- Slide Management ---

   /**
   * Creates a new slide section in an existing project sheet.
   * * @param {string} projectTabName - Name of the project tab.
   * @param {Object} slideInfo - Initial slide info data (UPPER_SNAKE_CASE keys).
   * @return {boolean} True if successful.
   */
  createNewSlide(projectTabName, slideInfo = {}) {
      try {
          const sheet = this.sheetAccessor.getSheet(projectTabName, false);
          if (!sheet) throw new Error(`Sheet not found: ${projectTabName}`);

          // Count existing slides to determine new slide number and insertion point
          const numExistingSlides = this.countExistingSlides(projectTabName);
          const newSlideNumber = numExistingSlides + 1;

          // Ensure required fields have defaults
          slideInfo.SLIDE_ID = slideInfo.SLIDE_ID || generateUUID();
          slideInfo.SLIDE_NUMBER = slideInfo.SLIDE_NUMBER || newSlideNumber;
          slideInfo.SLIDE_TITLE = slideInfo.SLIDE_TITLE || `Slide ${newSlideNumber}`;
          slideInfo.BACKGROUND_COLOR = slideInfo.BACKGROUND_COLOR || DEFAULT_COLORS.BACKGROUND;
          slideInfo.SHOW_CONTROLS = slideInfo.SHOW_CONTROLS !== undefined ? slideInfo.SHOW_CONTROLS : false;

          // Calculate where to insert the new slide section
          const slideInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO;
          const rowsPerSlide = slideInfoStructure.SECTION_END_ROW - slideInfoStructure.SECTION_START_ROW + 1;
          let insertAfterRow = SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.SECTION_END_ROW; // Default after project info
          if (numExistingSlides > 0) {
              insertAfterRow = slideInfoStructure.SECTION_START_ROW + (numExistingSlides * rowsPerSlide) - 1;
          }

          // Insert rows for the new slide section
          if (!this.sheetAccessor.insertRows(projectTabName, insertAfterRow, rowsPerSlide)) {
               throw new Error("Failed to insert rows for new slide.");
          }
          
          // Create the new slide section structure and populate data
          const newSlideStartRow = insertAfterRow + 1;
          if (!this._createSheetSectionStructure(projectTabName, "SLIDE_INFO")) {
               throw new Error("Failed to create new slide section structure.");
          }
          // Update header text
          const headerRow = newSlideStartRow;
          const headerText = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO.HEADER.replace("1", newSlideNumber);
          this.sheetAccessor.setCellValue(projectTabName, headerRow, "A", headerText);

          // Populate the data for the new slide
          if (!this.updateSlideSection(projectTabName, slideInfo.SLIDE_ID, slideInfo, newSlideNumber)) {
               throw new Error("Failed to populate data for the new slide section.");
          }

          logInfo(`Created new slide ${newSlideNumber} in project: ${projectTabName}`);
          return true;
      } catch (error) {
          logError(`Failed to create new slide in ${projectTabName}: ${error.message}\n${error.stack}`);
          return false;
      }
  }

  /**
   * Updates an existing slide section in a project sheet.
   * * @param {string} projectTabName - Name of the project tab.
   * @param {string|null} slideId - The ID of the slide to update. If null, updates based on instanceNumber.
   * @param {Object} updates - Object with properties to update (keys can be camelCase or UPPER_SNAKE_CASE).
   * @param {number} [instanceNumber] - The 1-based index of the slide section (required if slideId is null).
   * @return {boolean} True if successful.
   */
  updateSlideSection(projectTabName, slideId, updates, instanceNumber) {
      try {
          const sheet = this.sheetAccessor.getSheet(projectTabName, false);
          if (!sheet) throw new Error(`Sheet not found: ${projectTabName}`);

          const slideInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO;
          const rowsPerSlide = slideInfoStructure.SECTION_END_ROW - slideInfoStructure.SECTION_START_ROW + 1;
          let slideStartRow = -1;

          if (slideId) {
              // Find slide section by ID
              const slideIdRowOffset = slideInfoStructure.FIELDS.SLIDE_ID.ROW - slideInfoStructure.SECTION_START_ROW;
              const data = sheet.getRange("B" + slideInfoStructure.FIELDS.SLIDE_ID.ROW + ":B").getValues(); // Read only column B starting from first slide ID row
              for (let i = 0; i < data.length; i++) {
                  if (data[i][0] === slideId) {
                       // Found the ID, calculate the start row of this slide section
                       const actualRowIndex = slideInfoStructure.FIELDS.SLIDE_ID.ROW + i;
                       slideStartRow = actualRowIndex - slideIdRowOffset;
                       break;
                  }
              }
          } else if (instanceNumber && instanceNumber > 0) {
               // Calculate start row based on instance number
               slideStartRow = slideInfoStructure.SECTION_START_ROW + ((instanceNumber - 1) * rowsPerSlide);
          }

          if (slideStartRow === -1) {
               throw new Error(`Slide section not found for ID '${slideId}' or instance number ${instanceNumber}.`);
          }

          // Apply updates
          for (const [key, value] of Object.entries(updates)) {
              // Convert key to UPPER_SNAKE_CASE if needed
              const fieldKey = key.replace(/([A-Z])/g, '_$1').toUpperCase(); 
              if (slideInfoStructure.FIELDS.hasOwnProperty(fieldKey)) {
                  const fieldInfo = slideInfoStructure.FIELDS[fieldKey];
                  const fieldRow = slideStartRow + (fieldInfo.ROW - slideInfoStructure.SECTION_START_ROW);
                  this.sheetAccessor.setCellValue(projectTabName, fieldRow, fieldInfo.VALUE_COL, value);
              } else {
                   logWarning(`Unknown field key '${key}' provided in updateSlideSection for ${projectTabName}`);
              }
          }
          return true;

      } catch (error) {
           logError(`Failed to update slide section (ID: ${slideId}, Instance: ${instanceNumber}) in ${projectTabName}: ${error.message}`);
           return false;
      }
  }

  /**
   * Counts the number of existing slide sections based on headers.
   * * @param {string} projectTabName - Name of the project tab.
   * @return {number} Number of slides found.
   */
  countExistingSlides(projectTabName) {
      try {
          const sheet = this.sheetAccessor.getSheet(projectTabName, false);
          if (!sheet || sheet.getLastRow() === 0) return 0;

          // Read only column A to find headers
          const colAValues = sheet.getRange("A1:A" + sheet.getLastRow()).getValues();
          
          let count = 0;
          const slideHeaderPattern = /^SLIDE \d+ INFO/; // Matches "SLIDE # INFO"
          
          for (const row of colAValues) {
              if (row[0] && typeof row[0] === 'string' && slideHeaderPattern.test(row[0])) {
                  count++;
              }
          }
          return count;
      } catch (error) {
          logError(`Failed to count existing slides in ${projectTabName}: ${error.message}`);
          return 0;
      }
  }

  // --- Element Management ---

  /**
   * Creates a new element column in the element info section.
   * * @param {string} projectTabName - Name of the project tab.
   * @param {Object} elementInfo - Initial element info data (UPPER_SNAKE_CASE keys).
   * @return {boolean} True if successful.
   */
  createNewElement(projectTabName, elementInfo = {}) {
      try {
          const sheet = this.sheetAccessor.getSheet(projectTabName, true); // Ensure sheet exists
          if (!sheet) throw new Error(`Sheet not found: ${projectTabName}`);

          // Find the next available element column (starting from E)
          const elementInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO;
          const elementIdRow = elementInfoStructure.FIELDS.ELEMENT_ID.ROW;
          const firstElementColIndex = columnToIndex("E"); // Column E is index 4
          let nextElementCol = firstElementColIndex + 1; // Start checking from column E (1-based)

          while (true) {
               const cellValue = sheet.getRange(elementIdRow, nextElementCol).getValue();
               if (!cellValue) { // Found an empty ID cell, this is our column
                   break;
               }
               nextElementCol++;
               if (nextElementCol > sheet.getMaxColumns() + 1) { // Prevent infinite loop, allow adding one past max
                    // Optional: Insert a column if needed
                    // sheet.insertColumnAfter(nextElementCol - 1); 
                    // logDebug(`Inserted column ${nextElementCol} in ${projectTabName}`);
                    break; 
               }
          }
          const newElementIndex = nextElementCol - firstElementColIndex; // 1-based index relative to elements

          // Ensure required fields have defaults
          elementInfo.ELEMENT_ID = elementInfo.ELEMENT_ID || generateUUID();
          elementInfo.NICKNAME = elementInfo.NICKNAME || `Element ${newElementIndex}`;
          elementInfo.SEQUENCE = elementInfo.SEQUENCE !== undefined ? elementInfo.SEQUENCE : newElementIndex;
          elementInfo.TYPE = elementInfo.TYPE || ELEMENT_TYPES.RECTANGLE.LABEL; // Use LABEL from Config

          // Set element header
          this.sheetAccessor.setCellValue(projectTabName, elementInfoStructure.ELEMENT_HEADERS_ROW, nextElementCol, `Element ${newElementIndex}`);
           this.sheetAccessor.getSheet(projectTabName).getRange(elementInfoStructure.ELEMENT_HEADERS_ROW, nextElementCol).setFontWeight("bold").setHorizontalAlignment("center");


          // Write data using updateElementData helper
          if (!this.updateElementData(projectTabName, elementInfo.ELEMENT_ID, elementInfo, nextElementCol)) {
               throw new Error("Failed to write initial data for new element.");
          }

          logInfo(`Created new element ${newElementIndex} (Column ${indexToColumn(nextElementCol-1)}) in project: ${projectTabName}`);
          return true;
      } catch (error) {
          logError(`Failed to create new element in ${projectTabName}: ${error.message}\n${error.stack}`);
          return false;
      }
  }

   /**
   * Updates data for a specific element in its column.
   * * @param {string} projectTabName - Name of the project tab.
   * @param {string} elementId - The ID of the element to update.
   * @param {Object} updates - Object with properties to update (keys can be camelCase or UPPER_SNAKE_CASE).
   * @param {number} [knownColIndex] - Optional: The 1-based column index if already known.
   * @return {boolean} True if successful.
   */
  updateElementData(projectTabName, elementId, updates, knownColIndex) {
      try {
          const sheet = this.sheetAccessor.getSheet(projectTabName, false);
          if (!sheet) throw new Error(`Sheet not found: ${projectTabName}`);

          let elementColIndex = knownColIndex;

          // Find the element column if not provided
          if (!elementColIndex) {
               elementColIndex = this._findElementColumn(sheet, elementId);
               if (elementColIndex === -1) {
                    throw new Error(`Element column not found for ID: ${elementId}`);
               }
          }

          // --- Update ELEMENT_INFO Section ---
          const elementInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO;
          for (const [key, value] of Object.entries(updates)) {
               const fieldKey = key.replace(/([A-Z])/g, '_$1').toUpperCase(); // Convert camelCase to UPPER_SNAKE
               if (elementInfoStructure.FIELDS.hasOwnProperty(fieldKey)) {
                   const fieldInfo = elementInfoStructure.FIELDS[fieldKey];
                   this.sheetAccessor.setCellValue(projectTabName, fieldInfo.ROW, elementColIndex, value);
                   // Apply controls if creating new element (knownColIndex implies creation)
                   if (knownColIndex) {
                       this._applyFieldControl(projectTabName, "ELEMENT_INFO", fieldKey, fieldInfo.ROW, elementColIndex);
                   }
               } else if (key === 'timeline' && value && typeof value === 'object') {
                    // Handle timeline data update separately
                    this._updateTimelineData(sheet, elementId, elementColIndex, value);
               } else if (key === 'quiz' && value && typeof value === 'object') {
                    // Handle quiz data update separately
                    this._updateQuizData(sheet, elementId, elementColIndex, value);
               }
               // Ignore unknown keys or keys handled separately (timeline, quiz)
          }

          return true;
      } catch (error) {
           logError(`Failed to update element data for ID ${elementId} in ${projectTabName}: ${error.message}`);
           return false;
      }
  }

   /**
   * Helper to find the column index for a given element ID.
   * * @param {Sheet} sheet - The project sheet object.
   * @param {string} elementId - The element ID to find.
   * @return {number} The 1-based column index, or -1 if not found.
   */
  _findElementColumn(sheet, elementId) {
      const elementInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO;
      const elementIdRow = elementInfoStructure.FIELDS.ELEMENT_ID.ROW;
      const firstElementCol = columnToIndex("E") + 1; // Column E (1-based)
      const lastCol = sheet.getLastColumn();

      if (lastCol < firstElementCol) return -1; // No element columns

      const idRowValues = sheet.getRange(elementIdRow, firstElementCol, 1, lastCol - firstElementCol + 1).getValues()[0];
      const index = idRowValues.findIndex(id => id === elementId);

      return (index !== -1) ? firstElementCol + index : -1;
  }

  /**
   * Helper to update TIMELINE data for a specific element column.
   * * @param {Sheet} sheet - The project sheet object.
   * @param {string} elementId - The element ID.
   * @param {number} elementColIndex - The 1-based column index of the element.
   * @param {Object} timelineUpdates - Object with timeline properties to update (camelCase or UPPER_SNAKE_CASE).
   */
  _updateTimelineData(sheet, elementId, elementColIndex, timelineUpdates) {
       try {
          const timelineStructure = SHEET_STRUCTURE.PROJECT_TAB.TIMELINE;
           // First, ensure the element ID is set in the timeline section for this element
           this.sheetAccessor.setCellValue(sheet.getParent().getName(), timelineStructure.FIELDS.ELEMENT_ID.ROW, elementColIndex, elementId);

          for (const [key, value] of Object.entries(timelineUpdates)) {
              const fieldKey = key.replace(/([A-Z])/g, '_$1').toUpperCase();
              if (timelineStructure.FIELDS.hasOwnProperty(fieldKey) && fieldKey !== 'ELEMENT_ID') { // Don't overwrite ELEMENT_ID label row
                  const fieldInfo = timelineStructure.FIELDS[fieldKey];
                  this.sheetAccessor.setCellValue(sheet.getParent().getName(), fieldInfo.ROW, elementColIndex, value);
                   // Apply controls if needed (e.g., dropdowns for animation)
                   this._applyFieldControl(sheet.getParent().getName(), "TIMELINE", fieldKey, fieldInfo.ROW, elementColIndex);
              }
          }
       } catch(error) {
            logWarning(`Could not update timeline data for element ${elementId} in column ${elementColIndex}: ${error.message}`);
       }
  }

  /**
   * Helper to update QUIZ data for a specific element column.
   * * @param {Sheet} sheet - The project sheet object.
   * @param {string} elementId - The element ID.
   * @param {number} elementColIndex - The 1-based column index of the element.
   * @param {Object} quizUpdates - Object with quiz properties to update (camelCase or UPPER_SNAKE_CASE).
   */
  _updateQuizData(sheet, elementId, elementColIndex, quizUpdates) {
      try {
          const quizStructure = SHEET_STRUCTURE.PROJECT_TAB.QUIZ;
          for (const [key, value] of Object.entries(quizUpdates)) {
              const fieldKey = key.replace(/([A-Z])/g, '_$1').toUpperCase();
              if (quizStructure.FIELDS.hasOwnProperty(fieldKey)) {
                  const fieldInfo = quizStructure.FIELDS[fieldKey];
                  this.sheetAccessor.setCellValue(sheet.getParent().getName(), fieldInfo.ROW, elementColIndex, value);
                   // Apply controls if needed
                   this._applyFieldControl(sheet.getParent().getName(), "QUIZ", fieldKey, fieldInfo.ROW, elementColIndex);
              }
          }
      } catch(error) {
           logWarning(`Could not update quiz data for element ${elementId} in column ${elementColIndex}: ${error.message}`);
      }
  }

  /**
   * Counts the number of existing element columns.
   * * @param {string} projectTabName - Name of the project tab.
   * @return {number} Number of elements found based on non-empty ID cells.
   */
  countExistingElements(projectTabName) {
      try {
          const sheet = this.sheetAccessor.getSheet(projectTabName, false);
          if (!sheet) return 0;

          const elementIdRow = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO.FIELDS.ELEMENT_ID.ROW;
          const firstElementCol = columnToIndex("E") + 1; // Column E (1-based)
          const lastCol = sheet.getLastColumn();

          if (lastCol < firstElementCol) return 0; // No element columns

          const idRowValues = sheet.getRange(elementIdRow, firstElementCol, 1, lastCol - firstElementCol + 1).getValues()[0];
          
          // Count non-empty IDs
          return idRowValues.filter(id => id !== null && id !== '').length;
      } catch (error) {
          logError(`Failed to count existing elements in ${projectTabName}: ${error.message}`);
          return 0;
      }
  }

  // --- Bulk Update Function ---

  /**
   * Updates slides and elements sections from editor data.
   * This function formats the data and writes it back to the sheet.
   * * @param {string} projectTabName - Name of the project tab.
   * @param {Array<Object>} slides - Array of slide objects from editor state.
   * @param {Array<Object>} elements - Array of element objects from editor state.
   * @return {boolean} True if successful.
   */
  updateSlidesAndElements(projectTabName, slides, elements) {
      logInfo(`Starting updateSlidesAndElements for tab: ${projectTabName}`);
      const sheet = this.sheetAccessor.getSheet(projectTabName, false);
      if (!sheet) {
           logError(`Sheet not found for bulk update: ${projectTabName}`);
           return false;
      }

      try {
          // --- 1. Update Slides ---
          logDebug(`Processing ${slides.length} slides...`);
          const slideInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO;
          const rowsPerSlide = slideInfoStructure.SECTION_END_ROW - slideInfoStructure.SECTION_START_ROW + 1;
          const slideFields = Object.keys(slideInfoStructure.FIELDS); // Get field keys like SLIDE_ID, SLIDE_TITLE...

          // Prepare 2D array for all slide data
          const allSlideData = [];
          slides.sort((a, b) => a.slideNumber - b.slideNumber); // Ensure slides are ordered

          for (let i = 0; i < slides.length; i++) {
              const slide = slides[i];
              const slideBlock = [];
              const headerText = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO.HEADER.replace("1", slide.slideNumber || (i + 1));
              
              // Add header row for this slide block
              slideBlock.push([headerText, '']); // Header spans A:B

               // Add data rows for this slide block
               for (let r = 1; r < rowsPerSlide; r++) { // Start from 1 to skip header row offset
                   const targetRowConfig = slideInfoStructure.SECTION_START_ROW + r;
                   let label = '';
                   let value = '';
                   let fieldKeyFound = null;

                   // Find the field key corresponding to this row in the structure
                   for (const [fKey, fInfo] of Object.entries(slideInfoStructure.FIELDS)) {
                        if (fInfo.ROW === targetRowConfig) {
                            fieldKeyFound = fKey;
                            label = fKey.split('_').map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join(' ');
                            // Get value from slide object (convert camelCase key to UPPER_SNAKE_CASE)
                            const camelCaseKey = fKey.toLowerCase().replace(/_([a-z])/g, (match, char) => char.toUpperCase());
                            value = slide[camelCaseKey] !== undefined ? slide[camelCaseKey] : '';
                            break;
                        }
                   }
                   slideBlock.push([label, value]);
               }
               allSlideData.push(...slideBlock); // Add this slide's block to the main array
          }

          // Determine range to write slide data
          const slideStartWriteRow = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO.SECTION_START_ROW;
          const numSlideRowsToWrite = allSlideData.length;
          const slideEndWriteRow = slideStartWriteRow + numSlideRowsToWrite - 1;

          // Clear existing slide area (from first slide start row down to potentially more rows than before)
          const currentMaxRow = sheet.getMaxRows();
          const elementStartRow = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO.SECTION_START_ROW; // Find where elements start
          const clearEndRow = Math.max(slideEndWriteRow, elementStartRow > 0 ? elementStartRow -1 : slideEndWriteRow); // Clear down to just before elements start or end of new data
          if (clearEndRow >= slideStartWriteRow) {
              logDebug(`Clearing slide range: A${slideStartWriteRow}:B${clearEndRow}`);
              this.sheetAccessor.clearRange(projectTabName, slideStartWriteRow, "A", clearEndRow - slideStartWriteRow + 1, 2); // Clear columns A & B
          }

          // Write new slide data
          if (numSlideRowsToWrite > 0) {
              logDebug(`Writing ${numSlideRowsToWrite} rows of slide data starting at A${slideStartWriteRow}`);
              this.sheetAccessor.setRangeValues(projectTabName, slideStartWriteRow, "A", allSlideData);
              // Re-apply formatting/controls? Might be needed if clearRange removed them.
               for (let i = 0; i < slides.length; i++) {
                   const slideStartRowInstance = slideStartWriteRow + (i * rowsPerSlide);
                   this.sheetAccessor.getSheet(projectTabName).getRange(slideStartRowInstance, 1, 1, 2).merge().setBackground(DEFAULT_COLORS.SECTION_HEADER_BG).setFontWeight("bold").setHorizontalAlignment("center");
                   this._applyFieldControl(projectTabName, "SLIDE_INFO", "FILE_TYPE", slideStartRowInstance + (slideInfoStructure.FIELDS.FILE_TYPE.ROW - slideInfoStructure.SECTION_START_ROW), "B");
                   this._applyFieldControl(projectTabName, "SLIDE_INFO", "SHOW_CONTROLS", slideStartRowInstance + (slideInfoStructure.FIELDS.SHOW_CONTROLS.ROW - slideInfoStructure.SECTION_START_ROW), "B");
               }
          }

          // --- 2. Update Elements ---
          logDebug(`Processing ${elements.length} elements...`);
          const elementInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO;
          const timelineStructure = SHEET_STRUCTURE.PROJECT_TAB.TIMELINE;
          const quizStructure = SHEET_STRUCTURE.PROJECT_TAB.QUIZ;
          const elementStartCol = columnToIndex("E") + 1; // Column E (1-based)
          const numElements = elements.length;
          const elementEndCol = elementStartCol + numElements - 1;

          // Determine max row needed for elements, timeline, quiz
          const maxElementRow = elementInfoStructure.SECTION_END_ROW;
          const maxTimelineRow = timelineStructure.SECTION_END_ROW;
          const maxQuizRow = quizStructure.SECTION_END_ROW;
          const maxDataRow = Math.max(maxElementRow, maxTimelineRow, maxQuizRow);

          // Prepare 2D array for all element/timeline/quiz data
          // Rows: From ELEMENT_INFO start row to maxDataRow
          // Cols: Number of elements
          const numDataRows = maxDataRow - elementInfoStructure.SECTION_START_ROW + 1;
          const allElementData = Array(numDataRows).fill(null).map(() => Array(numElements).fill('')); // Initialize with empty strings

          // Map field keys to their 0-based row index within the combined element/timeline/quiz block
          const fieldKeyToRowIndex = {};
           const mapFields = (structure, offset) => {
               for (const [key, info] of Object.entries(structure.FIELDS)) {
                   fieldKeyToRowIndex[key] = info.ROW - offset; // 0-based index relative to start of ELEMENT_INFO
               }
           };
           mapFields(elementInfoStructure, elementInfoStructure.SECTION_START_ROW);
           mapFields(timelineStructure, elementInfoStructure.SECTION_START_ROW); // Use same offset
           mapFields(quizStructure, elementInfoStructure.SECTION_START_ROW);       // Use same offset


          // Populate the 2D array
          for (let j = 0; j < numElements; j++) { // Iterate through elements (columns)
              const element = elements[j];
              for (const [key, value] of Object.entries(element)) {
                  const fieldKey = key.replace(/([A-Z])/g, '_$1').toUpperCase(); // Convert camelCase to UPPER_SNAKE
                  
                  if (fieldKey === 'TIMELINE' && value && typeof value === 'object') {
                       // Populate timeline rows for this element column
                       for (const [tlKey, tlValue] of Object.entries(value)) {
                            const timelineFieldKey = tlKey.replace(/([A-Z])/g, '_$1').toUpperCase();
                            if (fieldKeyToRowIndex.hasOwnProperty(timelineFieldKey)) {
                                 const rowIndex = fieldKeyToRowIndex[timelineFieldKey];
                                 if (rowIndex >= 0 && rowIndex < numDataRows) {
                                     allElementData[rowIndex][j] = tlValue !== null ? tlValue : '';
                                 }
                            }
                       }
                  } else if (fieldKey === 'QUIZ' && value && typeof value === 'object') {
                       // Populate quiz rows for this element column
                       for (const [qKey, qValue] of Object.entries(value)) {
                            const quizFieldKey = qKey.replace(/([A-Z])/g, '_$1').toUpperCase();
                            if (fieldKeyToRowIndex.hasOwnProperty(quizFieldKey)) {
                                 const rowIndex = fieldKeyToRowIndex[quizFieldKey];
                                  if (rowIndex >= 0 && rowIndex < numDataRows) {
                                     allElementData[rowIndex][j] = qValue !== null ? qValue : '';
                                 }
                            }
                       }
                  } else if (fieldKeyToRowIndex.hasOwnProperty(fieldKey)) {
                      // Populate element info rows
                      const rowIndex = fieldKeyToRowIndex[fieldKey];
                       if (rowIndex >= 0 && rowIndex < numDataRows) {
                         allElementData[rowIndex][j] = value !== null ? value : '';
                      }
                  }
              }
               // Ensure ELEMENT_ID is set in the timeline section for this element column
               const timelineElementIdRowIndex = fieldKeyToRowIndex['ELEMENT_ID']; // Reuse ELEMENT_ID row index from element section
               if (timelineElementIdRowIndex !== undefined && timelineStructure.FIELDS.ELEMENT_ID) { // Check if timeline has ELEMENT_ID field defined
                    const timelineRowForId = timelineStructure.FIELDS.ELEMENT_ID.ROW - elementInfoStructure.SECTION_START_ROW;
                     if (timelineRowForId >= 0 && timelineRowForId < numDataRows) {
                         allElementData[timelineRowForId][j] = element.elementId; // Write element ID into the timeline section row
                     }
               }
          }

          // Clear existing element/timeline/quiz columns
          const currentMaxCol = sheet.getMaxColumns();
          if (currentMaxCol >= elementStartCol) {
               logDebug(`Clearing element range: ${indexToColumn(elementStartCol-1)}${elementInfoStructure.SECTION_START_ROW}:${indexToColumn(currentMaxCol-1)}${maxDataRow}`);
               this.sheetAccessor.clearRange(projectTabName, 
                                             elementInfoStructure.SECTION_START_ROW, 
                                             elementStartCol, 
                                             maxDataRow - elementInfoStructure.SECTION_START_ROW + 1, 
                                             currentMaxCol - elementStartCol + 1);
          }

          // Write new element data
          if (numElements > 0) {
               logDebug(`Writing ${numDataRows} rows x ${numElements} cols of element data starting at ${indexToColumn(elementStartCol-1)}${elementInfoStructure.SECTION_START_ROW}`);
               this.sheetAccessor.setRangeValues(projectTabName, elementInfoStructure.SECTION_START_ROW, elementStartCol, allElementData);
               // Re-apply formatting/controls for element headers and potentially all controls?
                for (let j = 0; j < numElements; j++) {
                    const col = elementStartCol + j;
                    this.sheetAccessor.setCellValue(projectTabName, elementInfoStructure.ELEMENT_HEADERS_ROW, col, `Element ${j + 1}`);
                    this.sheetAccessor.getSheet(projectTabName).getRange(elementInfoStructure.ELEMENT_HEADERS_ROW, col).setFontWeight("bold").setHorizontalAlignment("center");
                    // Reapply all controls for this column - could be slow
                    // for (let r = elementInfoStructure.SECTION_START_ROW; r <= maxDataRow; r++) {
                    //      // Find field key for row r and apply control... complex mapping needed
                    // }
                }
          }

          logInfo(`Successfully updated slides and elements for tab: ${projectTabName}`);
          return true;

      } catch (error) {
          logError(`Failed to update slides and elements for ${projectTabName}: ${error.message}\n${error.stack}`);
          return false;
      }
  }

}
