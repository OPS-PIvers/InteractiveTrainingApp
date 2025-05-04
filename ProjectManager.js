/**
 * ProjectManager class handles CRUD operations for training projects
 * Centralizes project-related functionality and coordinates between other components
 */
class ProjectManager {
  /**
   * Creates a new ProjectManager instance
   * * @param {SheetAccessor} sheetAccessor - SheetAccessor instance
   * @param {TemplateManager} templateManager - TemplateManager instance
   * @param {DriveManager} driveManager - DriveManager instance
   */
  constructor(sheetAccessor, templateManager, driveManager) {
    this.sheetAccessor = sheetAccessor;
    this.templateManager = templateManager;
    this.driveManager = driveManager;
  }
  
  /**
   * Creates a new project
   * * @param {string} projectName - Name for the new project
   * @return {Object} Result object with project details or error information
   */
  createProject(projectName) {
    try {
      // Validate project name
      if (!projectName || projectName.trim() === '') {
        return { success: false, message: 'Project name cannot be empty' };
      }
      
      // Create project tab from template
      const projectTabName = this.templateManager.createProjectFromTemplate(projectName);
      
      if (!projectTabName) {
        return { success: false, message: 'Failed to create project from template' };
      }
      
      // Get project info (including the generated ID)
      const projectInfo = this.templateManager.getProjectInfo(projectTabName);
      
      if (!projectInfo || !projectInfo.PROJECT_ID) {
        // Attempt to read ID directly if getProjectInfo failed initially
        const idFromSheet = this.sheetAccessor.getCellValue(
            projectTabName, 
            SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_ID.ROW, 
            SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_ID.VALUE_COL
        );
        if (idFromSheet) {
            projectInfo.PROJECT_ID = idFromSheet;
        } else {
            logError(`Failed to get project ID after creating tab: ${projectTabName}`);
            // Consider cleanup: delete the created tab?
            this.sheetAccessor.deleteSheet(projectTabName); 
            return { success: false, message: 'Failed to retrieve project ID after creation' };
        }
      }
      
      // Create project folders in Drive
      const folders = this.driveManager.createProjectFolders(projectInfo.PROJECT_ID, projectName);
      
      if (!folders || !folders.projectFolderId) {
        // Cleanup: delete the created tab if folder creation fails
        this.sheetAccessor.deleteSheet(projectTabName);
        // Also remove from index if it was added
        this.sheetAccessor.deleteRowFromIndex(projectInfo.PROJECT_ID);
        return { success: false, message: 'Failed to create project folders in Drive' };
      }
      
      // Update project info with folder ID
      projectInfo.PROJECT_FOLDER_ID = folders.projectFolderId;
      // Write the folder ID back to the sheet
      this.sheetAccessor.setCellValue(
          projectTabName, 
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_FOLDER_ID.ROW, 
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_FOLDER_ID.VALUE_COL,
          folders.projectFolderId
      );

      // Add to project index *after* successful folder creation and ID retrieval
       const now = new Date();
       this.sheetAccessor.addProjectToIndex({
           PROJECT_ID: projectInfo.PROJECT_ID,
           TITLE: projectName, // Use the original input name
           CREATED_AT: now,
           MODIFIED_AT: now,
           LAST_ACCESSED: now
       });
      
      // Activate the project tab
      const sheet = this.sheetAccessor.getSheet(projectTabName);
      sheet.activate();
      
      logInfo(`Successfully created new project: ${projectName} (${projectInfo.PROJECT_ID})`);
      
      return { 
        success: true, 
        message: 'Project created successfully',
        projectId: projectInfo.PROJECT_ID,
        projectTabName: projectTabName,
        projectFolderId: folders.projectFolderId,
        mediaFolderId: folders.mediaFolderId
      };
    } catch (error) {
      logError(`Failed to create project: ${error.message}\n${error.stack}`);
      // Attempt cleanup if possible (e.g., delete partially created sheet)
      return { success: false, message: `Error: ${error.message}` };
    }
  }
  
  /**
   * Gets a project by ID, including slides and elements.
   * This is the primary function used by the editor and viewer.
   * * @param {string} projectId - ID of the project to retrieve
   * @return {Object} Project object or null if not found
   */
  getProject(projectId) {
    try {
      // Find the project in the index
      const result = this.sheetAccessor.findProjectInIndex(projectId);
      
      if (!result.found) {
        logWarning(`Project not found in index: ${projectId}`);
        return null;
      }
      
      // Get basic project data from index
      const projectData = this.sheetAccessor.getSheetData(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
      const projectRow = projectData[result.rowIndex - 1]; // Adjust for 0-based array
      
      const project = {
        projectId: projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID],
        title: projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE],
        createdAt: projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.CREATED_AT],
        modifiedAt: projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT],
        lastAccessed: projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.LAST_ACCESSED]
      };
      
      // Get project tab name (assuming it matches sanitized title)
      const projectTabName = sanitizeSheetName(project.title);
      const sheetExists = this.sheetAccessor.getSheet(projectTabName, false);

      if (!sheetExists) {
          logError(`Project sheet not found for project ${projectId} with title ${project.title} (expected tab: ${projectTabName})`);
          // Return basic info from index, but indicate sheet issue
          project.error = `Project sheet '${projectTabName}' not found.`;
          project.slides = [];
          project.elements = [];
          return project; 
      }
      
      // Load additional data from project tab using TemplateManager
      const projectInfo = this.templateManager.getProjectInfo(projectTabName);
      
      // Merge project info (like folder ID, web app URL)
      Object.assign(project, projectInfo);
      
      // Get slides using the dedicated method
      project.slides = this.getProjectSlides(projectTabName);
      
      // Get elements using the dedicated method
      project.elements = this.getProjectElements(projectTabName);
      
      // Get folder details (optional, maybe not needed for editor directly)
      // if (project.PROJECT_FOLDER_ID) {
      //   project.folder = this.driveManager.getProjectFolder(projectId);
      // }
      
      logDebug(`Retrieved project ${projectId} with ${project.slides.length} slides and ${project.elements.length} elements.`);
      return project;
    } catch (error) {
      logError(`Failed to get project ${projectId}: ${error.message}\n${error.stack}`);
      return null;
    }
  }
  
  /**
   * Gets all projects (basic info from index)
   * * @param {boolean} includeSlidesAndElements - DEPRECATED: Use getProject for full data. Kept for compatibility.
   * @return {Array<Object>} Array of project objects (basic info)
   */
  getAllProjects(includeSlidesAndElements = false) {
    try {
      const projectIndexData = this.sheetAccessor.getSheetData(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
      
      // Skip header row and filter out any empty rows
      const projects = projectIndexData.slice(1)
        .filter(row => row && row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID] && row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE]) // Ensure projectId and title exist
        .map(row => {
          const project = {
            projectId: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID],
            title: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE],
            createdAt: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.CREATED_AT],
            modifiedAt: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT],
            lastAccessed: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.LAST_ACCESSED]
          };
          
          // NOTE: Including slides/elements here is inefficient. 
          // The caller should use getProject(projectId) for full details.
          if (includeSlidesAndElements) {
            logWarning("getAllProjects: includeSlidesAndElements is inefficient. Use getProject(id) instead.");
            // Optionally fetch full data here if really needed, but be aware of performance.
          }
          
          return project;
        });
        
      logDebug(`Retrieved ${projects.length} projects from index.`);
      return projects;
    } catch (error) {
      logError(`Failed to get all projects: ${error.message}\n${error.stack}`);
      return [];
    }
  }
  
  /**
   * Opens a project for editing (activates sheet and updates timestamp)
   * * @param {string} projectId - ID of the project to open
   * @return {Object} Result object with project details or error information
   */
  openProject(projectId) {
    try {
      // Find the project in the index
      const result = this.sheetAccessor.findProjectInIndex(projectId);
      
      if (!result.found) {
        return { success: false, message: 'Project not found in index' };
      }
      
      // Get project data from index to find the title
      const projectData = this.sheetAccessor.getSheetData(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
      const projectRow = projectData[result.rowIndex - 1]; // Adjust for 0-based array
      const projectTitle = projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE];
      
      // Find project tab by name
      const sanitizedTitle = sanitizeSheetName(projectTitle);
      const sheet = this.sheetAccessor.getSheet(sanitizedTitle, false);
      
      if (!sheet) {
        logError(`Project sheet not found for project ${projectId} with title ${projectTitle} (expected tab: ${sanitizedTitle})`);
        return { success: false, message: `Project sheet '${sanitizedTitle}' not found` };
      }
      
      // Activate the project tab
      sheet.activate();
      
      // Update last accessed timestamp
      const now = new Date();
      this.sheetAccessor.setCellValue(
        SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
        result.rowIndex,
        SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.LAST_ACCESSED + 1, // +1 for 1-based column
        now
      );
      
      logInfo(`Opened project for editing: ${projectTitle} (${projectId})`);
      
      // Return full project data needed by the editor
      const project = this.getProject(projectId);
      
      if (!project) {
          // This shouldn't happen if we found it above, but check anyway
           return { success: false, message: 'Failed to retrieve full project data after opening.' };
      }
      
      return { 
        success: true, 
        message: 'Project opened successfully',
        project: project // Return the full project object
      };
    } catch (error) {
      logError(`Failed to open project ${projectId}: ${error.message}\n${error.stack}`);
      return { success: false, message: `Error: ${error.message}` };
    }
  }
  
  /**
   * Updates a project, including its slides and elements.
   * This function expects the complete `slides` and `elements` arrays 
   * as potentially modified by the editor.
   * * @param {string} projectId - ID of the project to update
   * @param {Object} updates - Object containing properties to update. 
   * Expected to potentially include `title`, `slides`, `elements`.
   * @return {Object} Result object with success flag and message
   */
  updateProject(projectId, updates) {
    let projectTabName; // Define here to be accessible in catch block if needed
    try {
      // Find the project in the index first to get the current title
      const indexResult = this.sheetAccessor.findProjectInIndex(projectId);
      if (!indexResult.found) {
        return { success: false, message: 'Project not found in index' };
      }
      const projectIndexData = this.sheetAccessor.getSheetData(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
      const projectRow = projectIndexData[indexResult.rowIndex - 1];
      const currentTitle = projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE];
      projectTabName = sanitizeSheetName(currentTitle); // Use current title to find the sheet

      const sheet = this.sheetAccessor.getSheet(projectTabName, false);
       if (!sheet) {
          logError(`Project sheet not found for update: ${projectId} (expected tab: ${projectTabName})`);
          return { success: false, message: `Project sheet '${projectTabName}' not found for update.` };
      }

      const now = new Date();
      let updatedFields = [];

      // --- Update Project Info (Title, etc.) ---
      if (updates.title && updates.title !== currentTitle) {
        // Update title in index
        this.sheetAccessor.setCellValue(
          SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
          indexResult.rowIndex,
          SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE + 1,
          updates.title
        );
        
        // Update title in the project sheet itself
         this.sheetAccessor.setCellValue(
            projectTabName, 
            SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.TITLE.ROW, 
            SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.TITLE.VALUE_COL,
            updates.title
        );

        // Rename the sheet tab
        const newSanitizedTitle = sanitizeSheetName(updates.title);
        if (newSanitizedTitle !== projectTabName) {
            try {
                sheet.setName(newSanitizedTitle);
                projectTabName = newSanitizedTitle; // Update tab name for subsequent operations
                logInfo(`Renamed project sheet for ${projectId} to ${projectTabName}`);
            } catch (renameError) {
                logWarning(`Could not rename sheet tab for ${projectId} from ${projectTabName} to ${newSanitizedTitle}: ${renameError.message}`);
                // Continue with the old tab name, but the title is updated internally
            }
        }
        updatedFields.push('title');

         // Rename Drive folder (if applicable and exists)
         const folderId = this.sheetAccessor.getCellValue(
             projectTabName, // Use potentially updated tab name
             SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_FOLDER_ID.ROW,
             SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_FOLDER_ID.VALUE_COL
         );
         if (folderId) {
            try {
               const folder = this.driveManager.getFolder(folderId);
               if (folder) {
                  const newFolderName = `${updates.title} (${projectId})`; // Consistent naming
                  folder.setName(newFolderName);
                  logInfo(`Renamed project folder for ${projectId} to ${newFolderName}`);
               }
            } catch (driveError) {
                logWarning(`Could not rename Drive folder for ${projectId}: ${driveError.message}`);
            }
         }
      }

      // --- Update Slides and Elements ---
      // This is the core part for the editor save
      if (updates.slides || updates.elements) {
          logInfo(`Updating slides/elements for project ${projectId}`);
          // We need a robust way to write this data back.
          // Ideally, SheetAccessor or TemplateManager would have a function like:
          // this.sheetAccessor.writeSlidesAndElements(projectTabName, updates.slides, updates.elements);
          
          // Placeholder for the complex write operation:
          const writeSuccess = this.templateManager.updateSlidesAndElements(
              projectTabName, 
              updates.slides,      // Pass the full array from editor state
              updates.elements     // Pass the full array from editor state
          ); 

          if (!writeSuccess) {
               throw new Error("Failed to write slides and elements data to the sheet.");
          }
          if (updates.slides) updatedFields.push('slides');
          if (updates.elements) updatedFields.push('elements');
          logInfo(`Successfully wrote slides/elements data for project ${projectId}`);
      }
      
      // --- Update Modified Timestamp (always update on any change) ---
      const modifiedTimestamp = now.getTime();
      this.sheetAccessor.setCellValue(
        projectTabName, // Use potentially updated tab name
        SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.ROW,
        SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.VALUE_COL,
        modifiedTimestamp
      );
      
      // Also update in project index
      this.sheetAccessor.setCellValue(
        SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
        indexResult.rowIndex,
        SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT + 1,
        now // Use the Date object here
      );
      updatedFields.push('modifiedAt');
      
      logInfo(`Updated project: ${projectId}. Fields: ${updatedFields.join(', ')}`);
      
      return { 
        success: true, 
        message: 'Project updated successfully',
        projectId: projectId,
        updatedFields: updatedFields
      };
    } catch (error) {
      logError(`Failed to update project ${projectId} (Tab: ${projectTabName || 'unknown'}): ${error.message}\n${error.stack}`);
      return { success: false, message: `Error updating project: ${error.message}` };
    }
  }
  
  /**
   * Deletes a project
   * * @param {string} projectId - ID of the project to delete
   * @param {boolean} deleteDriveFiles - Whether to delete Drive files (default: false)
   * @return {Object} Result object with success flag and message
   */
  deleteProject(projectId, deleteDriveFiles = false) {
    let projectTitle = `Project ${projectId}`; // Default title for logging if fetch fails
    let projectTabName = null;
    let folderId = null;
    try {
      // Find the project in the index to get title and row index
      const result = this.sheetAccessor.findProjectInIndex(projectId);
      
      if (!result.found) {
        // Maybe the project sheet still exists but index is broken? Try finding by ID in sheets?
        // For now, assume if not in index, it's gone or invalid.
        return { success: false, message: 'Project not found in index' };
      }

      const projectIndexData = this.sheetAccessor.getSheetData(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
      const projectRow = projectIndexData[result.rowIndex - 1];
      projectTitle = projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE];
      projectTabName = sanitizeSheetName(projectTitle);

       // Try to get folder ID from the sheet before deleting it
       const sheet = this.sheetAccessor.getSheet(projectTabName, false);
       if (sheet) {
           folderId = this.sheetAccessor.getCellValue(
               projectTabName,
               SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_FOLDER_ID.ROW,
               SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_FOLDER_ID.VALUE_COL
           );
       }
      
      // Delete project tab
      const tabDeleted = this.sheetAccessor.deleteSheet(projectTabName);
      if (!tabDeleted) {
        logWarning(`Could not delete project tab: ${projectTabName}. It might have already been deleted or renamed.`);
      }
      
      // Remove from project index (delete the row)
       const indexSheet = this.sheetAccessor.getSheet(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
       indexSheet.deleteRow(result.rowIndex);
       logInfo(`Removed project ${projectId} from index.`);

      // Handle Drive files if requested
      if (deleteDriveFiles && folderId) {
        try {
          const folder = this.driveManager.getFolder(folderId);
          if (folder) {
            folder.setTrashed(true);
            logInfo(`Moved project folder to trash: ${projectTitle} (${projectId})`);
          } else {
               logWarning(`Drive folder not found for deletion: ${folderId}`);
          }
        } catch (driveError) {
          logWarning(`Could not delete Drive folder ${folderId}: ${driveError.message}`);
        }
      } else if (deleteDriveFiles) {
           logWarning(`Could not delete Drive folder for ${projectId} because folder ID was not found.`);
      }
      
      logInfo(`Deleted project: ${projectTitle} (${projectId})`);
      
      return { 
        success: true, 
        message: 'Project deleted successfully',
        projectId: projectId,
        title: projectTitle,
        driveFilesDeleted: deleteDriveFiles && !!folderId
      };
    } catch (error) {
      logError(`Failed to delete project ${projectId} (${projectTitle}): ${error.message}\n${error.stack}`);
      return { success: false, message: `Error deleting project: ${error.message}` };
    }
  }
  
  /**
   * Gets slides for a project by reading the specific project tab.
   * Assumes slides are stored in contiguous blocks defined by SLIDE_INFO structure.
   * * @param {string} projectTabName - Name of the project tab
   * @return {Array<Object>} Array of slide objects
   */
  getProjectSlides(projectTabName) {
    try {
      const sheet = this.sheetAccessor.getSheet(projectTabName, false);
      if (!sheet) {
          logWarning(`Sheet not found in getProjectSlides: ${projectTabName}`);
          return [];
      }
      
      const slides = [];
      const data = this.sheetAccessor.getSheetData(projectTabName); // Get all data once
      
      const slideInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO;
      const rowsPerSlide = slideInfoStructure.SECTION_END_ROW - slideInfoStructure.SECTION_START_ROW + 1;
      const slideHeaderLabel = slideInfoStructure.HEADER.replace(/\d+/, '(\\d+)'); // Pattern like "SLIDE (\d+) INFO"
      const slideHeaderPattern = new RegExp(`^${slideHeaderLabel}`);
      
      let currentSlide = null;
      
      for (let i = 0; i < data.length; i++) {
          const rowData = data[i];
          const cellA = rowData[0] ? rowData[0].toString().trim() : ''; // Check column A for headers/labels

          const headerMatch = cellA.match(slideHeaderPattern);

          if (headerMatch) {
              // Found a slide header
              if (currentSlide) { // Push the previous slide if exists
                  slides.push(currentSlide);
              }
              // Start a new slide object
              currentSlide = { slideNumber: parseInt(headerMatch[1], 10) };
              // Read the fields for this slide section
              for (let j = 1; j < rowsPerSlide; j++) { // Start from 1 to skip header row
                  const fieldRowIndex = i + j;
                  if (fieldRowIndex >= data.length) break; // Avoid reading past sheet data

                  const fieldRow = data[fieldRowIndex];
                  const fieldLabel = fieldRow[0] ? fieldRow[0].toString().trim() : '';
                  const fieldValue = fieldRow[1]; // Value is in column B

                  // Find the corresponding field in config
                  for (const [configKey, configValue] of Object.entries(slideInfoStructure.FIELDS)) {
                      if (configValue.LABEL_COL === 'A' && configValue.ROW === (slideInfoStructure.SECTION_START_ROW + j)) {
                           // Convert configKey (UPPER_SNAKE) to camelCase
                           const propName = configKey.toLowerCase().replace(/_([a-z])/g, (match, char) => char.toUpperCase());
                           // Assign value, handle boolean for checkbox
                           currentSlide[propName] = (configKey === 'SHOW_CONTROLS') ? !!fieldValue : fieldValue;
                           break; // Found match for this row
                      }
                  }
              }
               // Move main loop index past this slide block
              i += (rowsPerSlide - 1);

          } else if (currentSlide && cellA && cellA.match(/^(ELEMENT INFO|TIMELINE|QUIZ|USER TRACKING)/)) {
               // Found the start of another section, push the last slide and stop looking for slides
               slides.push(currentSlide);
               currentSlide = null;
               break; // Stop searching for slides
          }
      }
      
      // Add the very last slide if loop finished while processing one
      if (currentSlide) {
        slides.push(currentSlide);
      }
      
      // Sort slides by slide number just in case they weren't perfectly ordered
      slides.sort((a, b) => a.slideNumber - b.slideNumber);
      
      logDebug(`Found ${slides.length} slides in tab ${projectTabName}`);
      return slides;
    } catch (error) {
      logError(`Failed to get project slides for tab ${projectTabName}: ${error.message}\n${error.stack}`);
      return [];
    }
  }
  
  /**
   * Gets elements for a project by reading the ELEMENT INFO section.
   * Handles the columnar structure where each column E, F, G... represents an element.
   * * @param {string} projectTabName - Name of the project tab
   * @return {Array<Object>} Array of element objects
   */
  getProjectElements(projectTabName) {
    try {
      const sheet = this.sheetAccessor.getSheet(projectTabName, false);
      if (!sheet) {
          logWarning(`Sheet not found in getProjectElements: ${projectTabName}`);
          return [];
      }
      
      const elements = [];
      const elementInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO;
      const startRow = elementInfoStructure.SECTION_START_ROW;
      const endRow = elementInfoStructure.SECTION_END_ROW;
      const startCol = columnToIndex(elementInfoStructure.SECTION_START_COL) + 1; // +1 because D is the label col
      const numRows = endRow - startRow + 1;
      const maxCols = sheet.getMaxColumns();
      
      if (maxCols < startCol) {
          logDebug(`No element columns found in ${projectTabName} (startCol=${startCol}, maxCols=${maxCols})`);
          return []; // No element columns exist
      }

      // Read the relevant block of data (labels + element values)
      const dataRange = sheet.getRange(startRow, startCol -1, numRows, maxCols - (startCol - 1) + 1);
      const data = dataRange.getValues();

      // Map labels (from the first column of the data read) to their row index within the block
      const labelToRowIndex = {};
       for (const [configKey, configValue] of Object.entries(elementInfoStructure.FIELDS)) {
           const label = this.sheetAccessor.getCellValue(projectTabName, configValue.ROW, configValue.LABEL_COL);
           if (label) {
               labelToRowIndex[label.toString().trim()] = configValue.ROW - startRow; // 0-based index within the data block
           }
       }
      
      // Iterate through columns (starting from the second column of the data read, which is the first element)
      const numElementCols = data[0].length -1; // -1 for the label column
      for (let j = 1; j <= numElementCols; j++) {
          const element = {};
          let elementId = null;

          // Iterate through the known labels and extract values for the current element column
          for (const [label, rowIndex] of Object.entries(labelToRowIndex)) {
               if (rowIndex >= data.length) continue; // Skip if label row is outside data read

               const value = data[rowIndex][j]; // Get value from current element column `j`

               // Find the config key matching the label to convert to camelCase
               let configKey = null;
               for (const [key, config] of Object.entries(elementInfoStructure.FIELDS)) {
                    if (config.ROW === (startRow + rowIndex)) {
                        configKey = key;
                        break;
                    }
               }

               if (configKey) {
                  const propName = configKey.toLowerCase().replace(/_([a-z])/g, (match, char) => char.toUpperCase());
                  // Handle boolean checkboxes
                  if (['initiallyHidden', 'outline', 'shadow', 'textModal'].includes(propName)) {
                      element[propName] = !!value;
                  } else if (propName === 'opacity') { // Ensure opacity is a number
                       element[propName] = (value !== null && value !== '') ? Number(value) : 100;
                  }
                  else {
                      element[propName] = value;
                  }

                  if (propName === 'elementId') {
                      elementId = value;
                  }
               }
          }

          // Only add the element if it has a valid ID
          if (elementId) {
              // --- Fetch Timeline Data (if applicable) ---
              // This requires finding the TIMELINE section and matching elementId
              // Simplified: Assume getTimelineDataForElement exists
              element.timeline = this.getTimelineDataForElement(sheet, elementId, j + startCol -1); // Pass sheet, id, and original column index

              // --- Fetch Quiz Data (if applicable) ---
              // This requires finding the QUIZ section and matching elementId/column
              if (element.interactionType === 'Quiz') {
                   // Simplified: Assume getQuizDataForElement exists
                   element.quiz = this.getQuizDataForElement(sheet, elementId, j + startCol -1); // Pass sheet, id, and original column index
              }

              elements.push(element);
          } else if (data[labelToRowIndex['Element Nickname']][j]) {
               // Log if a column has a nickname but no ID - indicates potential data issue
               logWarning(`Element column ${j + startCol - 1} in ${projectTabName} has data but no Element ID.`);
          }
      }

      logDebug(`Found ${elements.length} elements in tab ${projectTabName}`);
      return elements;
    } catch (error) {
      logError(`Failed to get project elements for tab ${projectTabName}: ${error.message}\n${error.stack}`);
      return [];
    }
  }

   /**
   * Helper to get Timeline data for a specific element column.
   * @param {Sheet} sheet - The project sheet object.
   * @param {string} elementId - The ID of the element.
   * @param {number} elementColIndex - The 1-based column index of the element.
   * @returns {Object|null} Timeline data object or null.
   */
  getTimelineDataForElement(sheet, elementId, elementColIndex) {
      try {
          const timelineStructure = SHEET_STRUCTURE.PROJECT_TAB.TIMELINE;
          const startRow = timelineStructure.SECTION_START_ROW;
          const endRow = timelineStructure.SECTION_END_ROW;
          const labelCol = columnToIndex(timelineStructure.SECTION_START_COL); // 0-based label column index

          // Read the timeline block (labels + element value for this column)
          const dataRange = sheet.getRange(startRow, labelCol + 1, endRow - startRow + 1, elementColIndex - labelCol);
          const data = dataRange.getValues();

          const timeline = {};
          let foundTimelineData = false;

           for (const [configKey, configValue] of Object.entries(timelineStructure.FIELDS)) {
               const rowIndexInData = configValue.ROW - startRow; // 0-based index within data block
               if (rowIndexInData < 0 || rowIndexInData >= data.length) continue;

               const label = data[rowIndexInData][0]; // Label from the first column read
               const value = data[rowIndexInData][elementColIndex - (labelCol + 1)]; // Value from the element's column

               // Check if the timeline entry is for this element ID (optional, maybe structure assumes it?)
               // const timelineElementId = data[timelineStructure.FIELDS.ELEMENT_ID.ROW - startRow][0];
               // if (timelineElementId === elementId) { // This check might be redundant if structure is fixed

                   const propName = configKey.toLowerCase().replace(/_([a-z])/g, (match, char) => char.toUpperCase());
                   // Handle specific types if necessary (e.g., numbers for time)
                   if (['startTime', 'endTime'].includes(propName)) {
                       timeline[propName] = (value !== null && value !== '') ? Number(value) : null;
                   } else if (['pauseAt', 'showForDuration'].includes(propName)) {
                       timeline[propName] = !!value;
                   } else {
                       timeline[propName] = value;
                   }
                   if (value !== null && value !== '') {
                       foundTimelineData = true;
                   }
               // }
           }

          return foundTimelineData ? timeline : null;
      } catch (error) {
          logWarning(`Could not get timeline data for element ${elementId} in column ${elementColIndex}: ${error.message}`);
          return null;
      }
  }

  /**
   * Helper to get Quiz data for a specific element column.
   * @param {Sheet} sheet - The project sheet object.
   * @param {string} elementId - The ID of the element.
   * @param {number} elementColIndex - The 1-based column index of the element.
   * @returns {Object|null} Quiz data object or null.
   */
  getQuizDataForElement(sheet, elementId, elementColIndex) {
     try {
          const quizStructure = SHEET_STRUCTURE.PROJECT_TAB.QUIZ;
          const startRow = quizStructure.SECTION_START_ROW;
          const endRow = quizStructure.SECTION_END_ROW;
          const labelCol = columnToIndex(quizStructure.SECTION_START_COL); // 0-based label column index

          // Read the quiz block (labels + element value for this column)
          const dataRange = sheet.getRange(startRow, labelCol + 1, endRow - startRow + 1, elementColIndex - labelCol);
          const data = dataRange.getValues();

          const quiz = {};
          let foundQuizData = false;

          for (const [configKey, configValue] of Object.entries(quizStructure.FIELDS)) {
               const rowIndexInData = configValue.ROW - startRow; // 0-based index within data block
               if (rowIndexInData < 0 || rowIndexInData >= data.length) continue;

               const label = data[rowIndexInData][0]; // Label from the first column read
               const value = data[rowIndexInData][elementColIndex - (labelCol + 1)]; // Value from the element's column

               const propName = configKey.toLowerCase().replace(/_([a-z])/g, (match, char) => char.toUpperCase());

               // Handle specific types if necessary
               if (['includeFeedback', 'provideCorrectAnswer'].includes(propName)) {
                   quiz[propName] = !!value;
               } else if (['points', 'attempts'].includes(propName)) {
                   quiz[propName] = (value !== null && value !== '') ? Number(value) : null;
               } else {
                   quiz[propName] = value;
               }
               if (value !== null && value !== '') {
                  foundQuizData = true;
               }
          }

          return foundQuizData ? quiz : null;
      } catch (error) {
          logWarning(`Could not get quiz data for element ${elementId} in column ${elementColIndex}: ${error.message}`);
          return null;
      }
  }
  
  // --- Add/Update/Delete for Slides and Elements ---
  // These might be called by the editor OR by other backend processes.
  // The primary editor save mechanism uses updateProject with bulk data.

  /**
   * Adds a new slide to a project (updates sheet structure)
   * * @param {string} projectId - ID of the project
   * @param {Object} slideData - Data for the new slide
   * @return {Object} Result object with slide details or error information
   */
  addSlide(projectId, slideData = {}) {
    let projectTabName;
    try {
      const project = this.getProject(projectId);
      if (!project) return { success: false, message: 'Project not found' };
      projectTabName = sanitizeSheetName(project.title);
      
      const slideInfo = {
        SLIDE_ID: slideData.slideId || generateUUID(),
        SLIDE_TITLE: slideData.title || `Slide ${project.slides.length + 1}`,
        BACKGROUND_COLOR: slideData.backgroundColor || DEFAULT_COLORS.BACKGROUND,
        FILE_TYPE: slideData.fileType || '',
        FILE_URL: slideData.fileUrl || '',
        SLIDE_NUMBER: slideData.slideNumber || (project.slides.length + 1),
        SHOW_CONTROLS: slideData.showControls || false
      };
      
      // Use TemplateManager to handle sheet insertion logic
      const success = this.templateManager.createNewSlide(projectTabName, slideInfo);
      if (!success) throw new Error('TemplateManager failed to create new slide section.');
      
      // Update modified timestamp
      this.updateModifiedTimestamp(projectId, projectTabName);
      
      logInfo(`Added slide to project: ${projectId}, Slide: ${slideInfo.SLIDE_ID}`);
      
      // Convert field names to camelCase for returned object
      const newSlide = {};
      for (const [key, value] of Object.entries(slideInfo)) {
        const camelKey = key.toLowerCase().replace(/_([a-z])/g, (match, char) => char.toUpperCase());
        newSlide[camelKey] = value;
      }
      
      return { 
        success: true, 
        message: 'Slide added successfully',
        slide: newSlide,
        projectId: projectId
      };
    } catch (error) {
      logError(`Failed to add slide to ${projectId} (Tab: ${projectTabName || 'unknown'}): ${error.message}\n${error.stack}`);
      return { success: false, message: `Error adding slide: ${error.message}` };
    }
  }
  
  /**
   * Updates a slide in a project (finds section and updates values)
   * * @param {string} projectId - ID of the project
   * @param {string} slideId - ID of the slide to update
   * @param {Object} updates - Object with properties to update (camelCase keys)
   * @return {Object} Result object with success flag and message
   */
  updateSlide(projectId, slideId, updates) {
     let projectTabName;
     try {
          const project = this.getProject(projectId);
          if (!project) return { success: false, message: 'Project not found' };
          projectTabName = sanitizeSheetName(project.title);

          // Use TemplateManager to handle finding and updating the slide section
          const success = this.templateManager.updateSlideSection(projectTabName, slideId, updates);
          if (!success) throw new Error(`TemplateManager failed to update slide ${slideId}.`);

          // Update modified timestamp
          this.updateModifiedTimestamp(projectId, projectTabName);

          logInfo(`Updated slide in project: ${projectId}, Slide: ${slideId}`);
          
          return { 
              success: true, 
              message: 'Slide updated successfully',
              slideId: slideId,
              projectId: projectId,
              updatedFields: Object.keys(updates)
          };
      } catch (error) {
          logError(`Failed to update slide ${slideId} in ${projectId} (Tab: ${projectTabName || 'unknown'}): ${error.message}\n${error.stack}`);
          return { success: false, message: `Error updating slide: ${error.message}` };
      }
  }
  
  /**
   * Deletes a slide from a project (complex sheet operation, currently marks as deleted)
   * * @param {string} projectId - ID of the project
   * @param {string} slideId - ID of the slide to delete
   * @return {Object} Result object with success flag and message
   */
  deleteSlide(projectId, slideId) {
    // Note: Physically deleting rows/sections and renumbering is complex and error-prone.
    // Marking as deleted or hiding the section is often safer.
    // This implementation marks the title. A better approach might be a hidden "IsDeleted" cell.
    try {
      const project = this.getProject(projectId);
      if (!project) return { success: false, message: 'Project not found' };
      const slide = project.slides.find(s => s.slideId === slideId);
      if (!slide) return { success: false, message: 'Slide not found' };
      
      // Mark as deleted by updating title
      const updateResult = this.updateSlide(projectId, slideId, {
        title: `[DELETED] ${slide.title}` 
        // Consider adding an 'isDeleted: true' field if structure allows
      });

      if (!updateResult.success) {
           throw new Error(`Failed to mark slide ${slideId} as deleted.`);
      }
      
      logInfo(`Marked slide as deleted in project: ${projectId}, Slide: ${slideId}`);
      
      return { 
        success: true, 
        message: 'Slide marked as deleted successfully',
        slideId: slideId,
        projectId: projectId
      };
    } catch (error) {
      logError(`Failed to delete slide ${slideId} in ${projectId}: ${error.message}\n${error.stack}`);
      return { success: false, message: `Error deleting slide: ${error.message}` };
    }
  }
  
  /**
   * Adds a new element to a project (finds or adds a column)
   * * @param {string} projectId - ID of the project
   * @param {string} slideId - ID of the slide to add the element to
   * @param {Object} elementData - Data for the new element (camelCase keys)
   * @return {Object} Result object with element details or error information
   */
  addElement(projectId, slideId, elementData = {}) {
     let projectTabName;
     try {
          const project = this.getProject(projectId);
          if (!project) return { success: false, message: 'Project not found' };
          projectTabName = sanitizeSheetName(project.title);

          // Prepare element data with defaults
          const elementInfo = {
              elementId: elementData.elementId || generateUUID(),
              nickname: elementData.nickname || `Element ${project.elements.length + 1}`,
              slideId: slideId,
              sequence: elementData.sequence || (project.elements.filter(e => e.slideId === slideId).length + 1),
              type: elementData.type || ELEMENT_TYPES.RECTANGLE.LABEL,
              left: elementData.left !== undefined ? elementData.left : 100,
              top: elementData.top !== undefined ? elementData.top : 100,
              width: elementData.width !== undefined ? elementData.width : ELEMENT_TYPES.RECTANGLE.DEFAULT_WIDTH,
              height: elementData.height !== undefined ? elementData.height : ELEMENT_TYPES.RECTANGLE.DEFAULT_HEIGHT,
              angle: elementData.angle !== undefined ? elementData.angle : 0,
              initiallyHidden: elementData.initiallyHidden || false,
              opacity: elementData.opacity !== undefined ? elementData.opacity : 100,
              color: elementData.color || DEFAULT_COLORS.ELEMENT,
              outline: elementData.outline || false,
              outlineWidth: elementData.outlineWidth !== undefined ? elementData.outlineWidth : 1,
              outlineColor: elementData.outlineColor || DEFAULT_COLORS.OUTLINE,
              shadow: elementData.shadow || false,
              text: elementData.text || '',
              font: elementData.font || FONTS[0],
              fontColor: elementData.fontColor || DEFAULT_COLORS.TEXT,
              fontSize: elementData.fontSize !== undefined ? elementData.fontSize : 14,
              triggers: elementData.triggers || TRIGGER_TYPES.CLICK,
              interactionType: elementData.interactionType || INTERACTION_TYPES.REVEAL,
              textModal: elementData.textModal || false,
              textModalMessage: elementData.textModalMessage || '',
              animationType: elementData.animationType || ANIMATION_TYPES.NONE,
              animationSpeed: elementData.animationSpeed || ANIMATION_SPEEDS.MEDIUM,
              // Include timeline and quiz if provided
              timeline: elementData.timeline || null, 
              quiz: elementData.quiz || null      
          };

          // Use TemplateManager to handle finding/adding column and writing data
          const success = this.templateManager.createNewElement(projectTabName, elementInfo);
          if (!success) throw new Error('TemplateManager failed to create new element column/data.');

          // Update modified timestamp
          this.updateModifiedTimestamp(projectId, projectTabName);

          logInfo(`Added element to project: ${projectId}, Element: ${elementInfo.elementId}`);

          return { 
              success: true, 
              message: 'Element added successfully',
              element: elementInfo, // Return the full element data used
              projectId: projectId,
              slideId: slideId
          };
      } catch (error) {
          logError(`Failed to add element to ${projectId} (Tab: ${projectTabName || 'unknown'}): ${error.message}\n${error.stack}`);
          return { success: false, message: `Error adding element: ${error.message}` };
      }
  }
  
  /**
   * Updates an element in a project (finds column and updates values)
   * * @param {string} projectId - ID of the project
   * @param {string} elementId - ID of the element to update
   * @param {Object} updates - Object with properties to update (camelCase keys)
   * @return {Object} Result object with success flag and message
   */
  updateElement(projectId, elementId, updates) {
     let projectTabName;
     try {
          const project = this.getProject(projectId);
          if (!project) return { success: false, message: 'Project not found' };
          projectTabName = sanitizeSheetName(project.title);

          // Use TemplateManager to handle finding element column and updating data
          const success = this.templateManager.updateElementData(projectTabName, elementId, updates);
           if (!success) throw new Error(`TemplateManager failed to update element ${elementId}.`);

          // Update modified timestamp
          this.updateModifiedTimestamp(projectId, projectTabName);

          logInfo(`Updated element in project: ${projectId}, Element: ${elementId}`);
          
          return { 
              success: true, 
              message: 'Element updated successfully',
              elementId: elementId,
              projectId: projectId,
              updatedFields: Object.keys(updates)
          };
      } catch (error) {
          logError(`Failed to update element ${elementId} in ${projectId} (Tab: ${projectTabName || 'unknown'}): ${error.message}\n${error.stack}`);
          return { success: false, message: `Error updating element: ${error.message}` };
      }
  }
  
  /**
   * Deletes an element from a project (complex sheet operation, currently marks as deleted)
   * * @param {string} projectId - ID of the project
   * @param {string} elementId - ID of the element to delete
   * @return {Object} Result object with success flag and message
   */
  deleteElement(projectId, elementId) {
     // Note: Physically deleting columns and shifting is very complex.
     // Marking as deleted is safer.
     try {
          const project = this.getProject(projectId);
          if (!project) return { success: false, message: 'Project not found' };
          const element = project.elements.find(e => e.elementId === elementId);
          if (!element) return { success: false, message: 'Element not found' };

          // Mark as deleted by updating nickname and hiding
          const updateResult = this.updateElement(projectId, elementId, {
              nickname: `[DELETED] ${element.nickname || 'Element'}`,
              opacity: 0, // Make invisible
              initiallyHidden: true // Ensure it stays hidden
              // Consider adding an 'isDeleted: true' field if structure allows
          });

          if (!updateResult.success) {
              throw new Error(`Failed to mark element ${elementId} as deleted.`);
          }

          logInfo(`Marked element as deleted in project: ${projectId}, Element: ${elementId}`);
          
          return { 
              success: true, 
              message: 'Element marked as deleted successfully',
              elementId: elementId,
              projectId: projectId
          };
      } catch (error) {
          logError(`Failed to delete element ${elementId} in ${projectId}: ${error.message}\n${error.stack}`);
          return { success: false, message: `Error deleting element: ${error.message}` };
      }
  }

  /**
   * Helper function to update the modified timestamp in the sheet and index.
   * @param {string} projectId - The project ID.
   * @param {string} projectTabName - The name of the project's sheet tab.
   */
  updateModifiedTimestamp(projectId, projectTabName) {
      const now = new Date();
      const timestamp = now.getTime();

      // Update timestamp in the project sheet
      this.sheetAccessor.setCellValue(
        projectTabName,
        SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.ROW,
        SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.VALUE_COL,
        timestamp
      );
      
      // Update timestamp in the project index
      const result = this.sheetAccessor.findProjectInIndex(projectId);
      if (result.found) {
        this.sheetAccessor.setCellValue(
          SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
          result.rowIndex,
          SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT + 1, // +1 for 1-based column
          now // Use Date object for index sheet formatting
        );
      } else {
          logWarning(`Could not find project ${projectId} in index to update modified timestamp.`);
      }
  }
  
  /**
   * Deploys a project as a web app (updates URL in sheet)
   * * @param {string} projectId - ID of the project to deploy
   * @return {Object} Result object with deployment details or error information
   */
  deployProject(projectId) {
    let projectTabName;
    try {
      const project = this.getProject(projectId);
      if (!project) return { success: false, message: 'Project not found' };
      projectTabName = sanitizeSheetName(project.title);
      
      const scriptUrl = ScriptApp.getService().getUrl();
      // Ensure URL ends with /exec for execution
      const execUrl = scriptUrl.endsWith('/exec') ? scriptUrl : scriptUrl + '/exec';
      const webAppUrl = `${execUrl}?project=${projectId}`; // Parameter for doGet
      
      // Update project with web app URL
      const updateResult = this.updateProject(projectId, { webAppUrl: webAppUrl });
      if (!updateResult.success) {
          throw new Error(`Failed to update project with Web App URL: ${updateResult.message}`);
      }
      
      logInfo(`Deployed project as web app: ${projectId}, URL: ${webAppUrl}`);
      
      return { 
        success: true, 
        message: 'Project deployed successfully. Make sure the script is deployed as a Web App accessible to users.',
        projectId: projectId,
        webAppUrl: webAppUrl,
        scriptId: ScriptApp.getScriptId()
      };
    } catch (error) {
      logError(`Failed to deploy project ${projectId} (Tab: ${projectTabName || 'unknown'}): ${error.message}\n${error.stack}`);
      return { success: false, message: `Error deploying project: ${error.message}` };
    }
  }
}
