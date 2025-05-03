/**
 * TemplateManager class handles the creation and management of project templates
 * This includes creating the base template and new project templates from it
 */
class TemplateManager {
    /**
     * Creates a new TemplateManager instance
     * 
     * @param {SheetAccessor} sheetAccessor - SheetAccessor instance
     */
    constructor(sheetAccessor) {
      this.sheetAccessor = sheetAccessor;
    }
    
    /**
     * Creates the base template in the spreadsheet
     * 
     * @param {string} templateName - Name for the template tab (default: Template)
     * @return {boolean} True if successful
     */
    createBaseTemplate(templateName = "Template") {
      try {
        // Create template sheet if it doesn't exist
        const sheet = this.sheetAccessor.getSheet(templateName);
        if (!sheet) {
          logError(`Failed to create template sheet: ${templateName}`);
          return false;
        }
        
        // Create PROJECT INFO section
        this.sheetAccessor.createSection(templateName, "PROJECT_INFO");
        
        // Create SLIDE INFO section
        this.sheetAccessor.createSection(templateName, "SLIDE_INFO");
        
        // Create ELEMENT INFO section
        this.sheetAccessor.createSection(templateName, "ELEMENT_INFO");
        
        // Create TIMELINE section
        this.sheetAccessor.createSection(templateName, "TIMELINE");
        
        // Create QUIZ section
        this.sheetAccessor.createSection(templateName, "QUIZ");
        
        // Create USER TRACKING section
        this.sheetAccessor.createSection(templateName, "USER_TRACKING");
        
        // Set default column widths
        sheet.setColumnWidth(1, 150); // Column A
        sheet.setColumnWidth(2, 250); // Column B
        sheet.setColumnWidth(3, 50);  // Column C (spacer)
        sheet.setColumnWidth(4, 150); // Column D
        sheet.setColumnWidth(5, 150); // Column E
        sheet.setColumnWidth(6, 150); // Column F
        sheet.setColumnWidth(7, 150); // Column G
        
        logInfo(`Created base template: ${templateName}`);
        return true;
      } catch (error) {
        logError(`Failed to create base template: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Creates a new project tab from the template
     * 
     * @param {string} projectName - Name for the new project
     * @param {string} templateName - Name of the template tab (default: Template)
     * @return {string} Name of the created project tab, or null if failed
     */
    createProjectFromTemplate(projectName, templateName = "Template") {
      try {
        // Sanitize project name for use as sheet name
        const projectTabName = sanitizeSheetName(projectName);
        
        // Create project tab from template
        const success = this.sheetAccessor.createTabFromTemplate(projectTabName, templateName);
        if (!success) {
          logError(`Failed to create project tab from template: ${projectTabName}`);
          return null;
        }
        
        // Initialize project with default values
        const projectId = generateUUID();
        const now = new Date();
        const nowTimestamp = now.getTime();
        
        // Set PROJECT INFO section values
        const projectInfo = {
          PROJECT_ID: projectId,
          PROJECT_WEB_APP_URL: "", // Will be set after web app deployment
          TITLE: projectName,
          CREATED_AT: nowTimestamp,
          MODIFIED_AT: nowTimestamp,
          PROJECT_FOLDER_ID: "" // Will be set after folder creation
        };
        
        this.sheetAccessor.updateProjectInfo(projectTabName, projectInfo);
        
        // Add project to project index
        this.sheetAccessor.addProjectToIndex({
          projectId: projectId,
          title: projectName,
          createdAt: now,
          modifiedAt: now
        });
        
        logInfo(`Created project from template: ${projectTabName} (${projectId})`);
        return projectTabName;
      } catch (error) {
        logError(`Failed to create project from template: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Updates an existing project tab with new structure elements
     * This is used when the template is updated and existing projects need to be updated
     * 
     * @param {string} projectTabName - Name of the project tab to update
     * @param {string} templateName - Name of the template tab (default: Template)
     * @return {boolean} True if successful
     */
    updateProjectStructure(projectTabName, templateName = "Template") {
      try {
        // Get template and project sheets
        const templateSheet = this.sheetAccessor.getSheet(templateName, false);
        const projectSheet = this.sheetAccessor.getSheet(projectTabName, false);
        
        if (!templateSheet || !projectSheet) {
          logError(`Template or project sheet not found: ${templateName} or ${projectTabName}`);
          return false;
        }
        
        // Get project data to preserve
        const projectInfo = this.getProjectInfo(projectTabName);
        
        // Identify missing sections in the project
        const missingStructures = this.identifyMissingStructures(projectTabName, templateName);
        
        // Add missing sections
        for (const section of missingStructures) {
          if (section === "PROJECT_INFO") {
            // PROJECT_INFO should already exist in all projects
            continue;
          } else if (section === "SLIDE_INFO") {
            // Get existing slides and add any missing structure
            const numExistingSlides = this.countExistingSlides(projectTabName);
            const slideInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO;
            
            // Check if we need to add more slide structure
            if (numExistingSlides === 0) {
              this.sheetAccessor.createSection(projectTabName, "SLIDE_INFO", 1);
            }
          } else {
            // Add other missing sections like ELEMENT_INFO, TIMELINE, QUIZ, USER_TRACKING
            this.sheetAccessor.createSection(projectTabName, section);
          }
        }
        
        logInfo(`Updated project structure: ${projectTabName}`);
        return true;
      } catch (error) {
        logError(`Failed to update project structure: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Counts the number of existing slides in a project
     * 
     * @param {string} projectTabName - Name of the project tab
     * @return {number} Number of slides found
     */
    countExistingSlides(projectTabName) {
      try {
        const sheet = this.sheetAccessor.getSheet(projectTabName, false);
        if (!sheet) return 0;
        
        // Get all data from the sheet
        const data = this.sheetAccessor.getSheetData(projectTabName);
        
        // Count sections that match the slide header pattern
        let count = 0;
        const slideHeaderPattern = /^SLIDE \d+ INFO/;
        
        for (const row of data) {
          if (row[0] && typeof row[0] === 'string' && slideHeaderPattern.test(row[0])) {
            count++;
          }
        }
        
        return count;
      } catch (error) {
        logError(`Failed to count existing slides: ${error.message}`);
        return 0;
      }
    }
    
    /**
     * Gets project info from a project tab
     * 
     * @param {string} projectTabName - Name of the project tab
     * @return {Object} Project info object
     */
    getProjectInfo(projectTabName) {
      try {
        const projectInfo = {};
        const projectInfoFields = SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS;
        
        for (const field in projectInfoFields) {
          if (projectInfoFields.hasOwnProperty(field)) {
            const fieldInfo = projectInfoFields[field];
            const value = this.sheetAccessor.getCellValue(
              projectTabName,
              fieldInfo.ROW,
              fieldInfo.VALUE_COL
            );
            
            projectInfo[field] = value;
          }
        }
        
        return projectInfo;
      } catch (error) {
        logError(`Failed to get project info: ${error.message}`);
        return {};
      }
    }
    
    /**
     * Identifies missing sections in a project compared to the template
     * 
     * @param {string} projectTabName - Name of the project tab
     * @param {string} templateName - Name of the template tab
     * @return {Array<string>} Array of missing section names
     */
    identifyMissingStructures(projectTabName, templateName) {
      try {
        const sections = [
          "PROJECT_INFO",
          "SLIDE_INFO",
          "ELEMENT_INFO",
          "TIMELINE",
          "QUIZ",
          "USER_TRACKING"
        ];
        
        const missingStructures = [];
        
        // Check for each section
        for (const section of sections) {
          const sectionStructure = SHEET_STRUCTURE.PROJECT_TAB[section];
          
          if (!sectionStructure) continue;
          
          // Get section header cell from template and project
          const headerRow = sectionStructure.SECTION_START_ROW;
          let headerCol = sectionStructure.SECTION_START_COL || "A";
          
          const templateHeader = this.sheetAccessor.getCellValue(templateName, headerRow, headerCol);
          const projectHeader = this.sheetAccessor.getCellValue(projectTabName, headerRow, headerCol);
          
          // If section header doesn't exist in the project, add to missing structures
          if (!projectHeader || (typeof templateHeader === 'string' && 
                                 !projectHeader.toString().includes(templateHeader.toString().split(' ')[0]))) {
            missingStructures.push(section);
          }
        }
        
        return missingStructures;
      } catch (error) {
        logError(`Failed to identify missing structures: ${error.message}`);
        return [];
      }
    }
    
    /**
     * Creates a new slide in an existing project
     * 
     * @param {string} projectTabName - Name of the project tab
     * @param {Object} slideInfo - Initial slide info data
     * @return {boolean} True if successful
     */
    createNewSlide(projectTabName, slideInfo = {}) {
      try {
        // Count existing slides
        const numExistingSlides = this.countExistingSlides(projectTabName);
        const newSlideNumber = numExistingSlides + 1;
        
        // Generate slide ID if not provided
        if (!slideInfo.SLIDE_ID) {
          slideInfo.SLIDE_ID = generateUUID();
        }
        
        // Set default slide number if not provided
        if (!slideInfo.SLIDE_NUMBER) {
          slideInfo.SLIDE_NUMBER = newSlideNumber;
        }
        
        // Set default slide title if not provided
        if (!slideInfo.SLIDE_TITLE) {
          slideInfo.SLIDE_TITLE = `Slide ${newSlideNumber}`;
        }
        
        // Set default background color if not provided
        if (!slideInfo.BACKGROUND_COLOR) {
          slideInfo.BACKGROUND_COLOR = DEFAULT_COLORS.BACKGROUND;
        }
        
        // Calculate where to insert the new slide section
        let insertAfterRow = 0;
        
        if (numExistingSlides > 0) {
          // Find the row after the last slide section
          const slideInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO;
          const rowsPerSlide = slideInfoStructure.SECTION_END_ROW - slideInfoStructure.SECTION_START_ROW + 1;
          insertAfterRow = slideInfoStructure.SECTION_START_ROW + (numExistingSlides * rowsPerSlide) - 1;
        } else {
          // Insert after the PROJECT_INFO section
          insertAfterRow = SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.SECTION_END_ROW;
        }
        
        // Insert rows for the new slide section
        const sheet = this.sheetAccessor.getSheet(projectTabName);
        const slideInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO;
        const rowsToInsert = slideInfoStructure.SECTION_END_ROW - slideInfoStructure.SECTION_START_ROW + 1;
        
        sheet.insertRowsAfter(insertAfterRow, rowsToInsert);
        
        // Create the new slide section
        this.sheetAccessor.createSection(projectTabName, "SLIDE_INFO", newSlideNumber, slideInfo);
        
        // Update project modified timestamp
        const now = new Date().getTime();
        this.sheetAccessor.setCellValue(
          projectTabName,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.ROW,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.VALUE_COL,
          now
        );
        
        logInfo(`Created new slide ${newSlideNumber} in project: ${projectTabName}`);
        return true;
      } catch (error) {
        logError(`Failed to create new slide: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Creates a new element column in the element info section
     * 
     * @param {string} projectTabName - Name of the project tab
     * @param {Object} elementInfo - Initial element info data
     * @return {boolean} True if successful
     */
    createNewElement(projectTabName, elementInfo = {}) {
      try {
        // Count existing elements
        const numExistingElements = this.countExistingElements(projectTabName);
        const newElementIndex = numExistingElements + 1;
        
        // Generate element ID if not provided
        if (!elementInfo.ELEMENT_ID) {
          elementInfo.ELEMENT_ID = generateUUID();
        }
        
        // Set default element nickname if not provided
        if (!elementInfo.NICKNAME) {
          elementInfo.NICKNAME = `Element ${newElementIndex}`;
        }
        
        // Set default sequence if not provided
        if (!elementInfo.SEQUENCE) {
          elementInfo.SEQUENCE = newElementIndex;
        }
        
        // Add element column
        this.sheetAccessor.addElementColumn(projectTabName, newElementIndex, elementInfo);
        
        // Update project modified timestamp
        const now = new Date().getTime();
        this.sheetAccessor.setCellValue(
          projectTabName,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.ROW,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.VALUE_COL,
          now
        );
        
        logInfo(`Created new element ${newElementIndex} in project: ${projectTabName}`);
        return true;
      } catch (error) {
        logError(`Failed to create new element: ${error.message}`);
        return false;
      }
    }
    
    /**
     * Counts the number of existing elements in a project
     * 
     * @param {string} projectTabName - Name of the project tab
     * @return {number} Number of elements found
     */
    countExistingElements(projectTabName) {
      try {
        const sheet = this.sheetAccessor.getSheet(projectTabName, false);
        if (!sheet) return 0;
        
        // Get column E, F, G, etc. which contain element data
        // Start from column index 4 (E) which is the first element column
        let count = 0;
        const elementIdRow = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO.FIELDS.ELEMENT_ID.ROW;
        
        for (let col = 5; col <= sheet.getMaxColumns(); col++) {
          const elementId = sheet.getRange(elementIdRow, col).getValue();
          if (elementId) {
            count++;
          } else {
            // Stop at the first empty element ID column
            break;
          }
        }
        
        return count;
      } catch (error) {
        logError(`Failed to count existing elements: ${error.message}`);
        return 0;
      }
    }
  }