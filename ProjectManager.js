/**
 * ProjectManager class handles CRUD operations for training projects
 * Centralizes project-related functionality and coordinates between other components
 */
class ProjectManager {
    /**
     * Creates a new ProjectManager instance
     * 
     * @param {SheetAccessor} sheetAccessor - SheetAccessor instance
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
     * 
     * @param {string} projectName - Name for the new project
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
        
        // Get project info
        const projectInfo = this.templateManager.getProjectInfo(projectTabName);
        
        if (!projectInfo || !projectInfo.PROJECT_ID) {
          return { success: false, message: 'Failed to get project information' };
        }
        
        // Create project folders in Drive
        const folders = this.driveManager.createProjectFolders(projectInfo.PROJECT_ID, projectName);
        
        if (!folders || !folders.projectFolderId) {
          return { success: false, message: 'Failed to create project folders in Drive' };
        }
        
        // Update project info with folder ID
        projectInfo.PROJECT_FOLDER_ID = folders.projectFolderId;
        this.sheetAccessor.updateProjectInfo(projectTabName, projectInfo);
        
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
        logError(`Failed to create project: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
    
    /**
     * Gets a project by ID
     * 
     * @param {string} projectId - ID of the project to retrieve
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
        
        // Get project data from index
        const projectData = this.sheetAccessor.getSheetData(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
        const projectRow = projectData[result.rowIndex - 1]; // Adjust for 0-based array
        
        const project = {
          projectId: projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID],
          title: projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE],
          createdAt: projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.CREATED_AT],
          modifiedAt: projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT],
          lastAccessed: projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.LAST_ACCESSED]
        };
        
        // Get project tab name
        const projectTabName = sanitizeSheetName(project.title);
        
        // Load additional data from project tab
        const projectInfo = this.templateManager.getProjectInfo(projectTabName);
        
        // Merge project info
        Object.assign(project, projectInfo);
        
        // Get slides
        project.slides = this.getProjectSlides(projectTabName);
        
        // Get elements
        project.elements = this.getProjectElements(projectTabName);
        
        // Get folder details
        if (projectInfo.PROJECT_FOLDER_ID) {
          project.folder = this.driveManager.getProjectFolder(projectId);
        }
        
        return project;
      } catch (error) {
        logError(`Failed to get project: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Gets all projects
     * 
     * @param {boolean} includeSlidesAndElements - Whether to include slides and elements data (default: false)
     * @return {Array<Object>} Array of project objects
     */
    getAllProjects(includeSlidesAndElements = false) {
      try {
        const projectIndexData = this.sheetAccessor.getSheetData(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
        
        // Skip header row and filter out any empty rows
        const projects = projectIndexData.slice(1)
          .filter(row => row[0] && row[1]) // Ensure projectId and title exist
          .map(row => {
            const project = {
              projectId: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID],
              title: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE],
              createdAt: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.CREATED_AT],
              modifiedAt: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT],
              lastAccessed: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.LAST_ACCESSED]
            };
            
            if (includeSlidesAndElements) {
              // Get full project data
              const fullProject = this.getProject(project.projectId);
              if (fullProject) {
                Object.assign(project, {
                  slides: fullProject.slides,
                  elements: fullProject.elements,
                  folder: fullProject.folder,
                  webAppUrl: fullProject.PROJECT_WEB_APP_URL,
                  folderId: fullProject.PROJECT_FOLDER_ID
                });
              }
            }
            
            return project;
          });
        
        return projects;
      } catch (error) {
        logError(`Failed to get all projects: ${error.message}`);
        return [];
      }
    }
    
    /**
     * Opens a project for editing
     * 
     * @param {string} projectId - ID of the project to open
     * @return {Object} Result object with project details or error information
     */
    openProject(projectId) {
      try {
        // Find the project in the index
        const result = this.sheetAccessor.findProjectInIndex(projectId);
        
        if (!result.found) {
          return { success: false, message: 'Project not found in index' };
        }
        
        // Get project data from index
        const projectData = this.sheetAccessor.getSheetData(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
        const projectRow = projectData[result.rowIndex - 1]; // Adjust for 0-based array
        const projectTitle = projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE];
        
        // Find project tab by name
        const sanitizedTitle = sanitizeSheetName(projectTitle);
        const sheet = this.sheetAccessor.getSheet(sanitizedTitle, false);
        
        if (!sheet) {
          return { success: false, message: 'Project sheet not found' };
        }
        
        // Activate the project tab
        sheet.activate();
        
        // Update last accessed timestamp
        const now = new Date();
        this.sheetAccessor.setCellValue(
          SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
          result.rowIndex,
          SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.LAST_ACCESSED + 1,
          now
        );
        
        logInfo(`Opened project for editing: ${projectTitle} (${projectId})`);
        
        // Return project data
        const project = this.getProject(projectId);
        
        return { 
          success: true, 
          message: 'Project opened successfully',
          project: project
        };
      } catch (error) {
        logError(`Failed to open project: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
    
    /**
     * Updates a project's properties
     * 
     * @param {string} projectId - ID of the project to update
     * @param {Object} updates - Object with properties to update
     * @return {Object} Result object with success flag and message
     */
    updateProject(projectId, updates) {
      try {
        // Find the project
        const project = this.getProject(projectId);
        
        if (!project) {
          return { success: false, message: 'Project not found' };
        }
        
        const projectTabName = sanitizeSheetName(project.title);
        const now = new Date();
        
        // Update project index if title is changed
        if (updates.title && updates.title !== project.title) {
          // Find the project in the index
          const result = this.sheetAccessor.findProjectInIndex(projectId);
          
          if (result.found) {
            // Update title in index
            this.sheetAccessor.setCellValue(
              SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
              result.rowIndex,
              SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE + 1,
              updates.title
            );
            
            // Update modified timestamp
            this.sheetAccessor.setCellValue(
              SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
              result.rowIndex,
              SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT + 1,
              now
            );
            
            // Rename project sheet (this is complex and might require creating a new sheet)
            // For now, we'll leave the sheet name as is, and just update the title in the sheet
          }
        }
        
        // Update project tab
        const projectInfo = {
          TITLE: updates.title || project.title,
          MODIFIED_AT: now.getTime()
        };
        
        if (updates.webAppUrl) {
          projectInfo.PROJECT_WEB_APP_URL = updates.webAppUrl;
        }
        
        this.sheetAccessor.updateProjectInfo(projectTabName, projectInfo);
        
        // If there's a folder ID and folder was renamed due to title change
        if (project.PROJECT_FOLDER_ID && updates.title && updates.title !== project.title) {
          // Rename the folder
          const folder = this.driveManager.getFolder(project.PROJECT_FOLDER_ID);
          if (folder) {
            const newFolderName = `${updates.title} (${projectId})`;
            folder.setName(newFolderName);
          }
        }
        
        logInfo(`Updated project: ${projectId}`);
        
        return { 
          success: true, 
          message: 'Project updated successfully',
          projectId: projectId,
          updatedFields: Object.keys(updates)
        };
      } catch (error) {
        logError(`Failed to update project: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
    
    /**
     * Deletes a project
     * 
     * @param {string} projectId - ID of the project to delete
     * @param {boolean} deleteDriveFiles - Whether to delete Drive files (default: false)
     * @return {Object} Result object with success flag and message
     */
    deleteProject(projectId, deleteDriveFiles = false) {
      try {
        // Find the project
        const project = this.getProject(projectId);
        
        if (!project) {
          return { success: false, message: 'Project not found' };
        }
        
        const projectTabName = sanitizeSheetName(project.title);
        
        // Delete project tab
        const tabDeleted = this.sheetAccessor.deleteSheet(projectTabName);
        
        if (!tabDeleted) {
          logWarning(`Could not delete project tab: ${projectTabName}`);
        }
        
        // Remove from project index
        const result = this.sheetAccessor.findProjectInIndex(projectId);
        
        if (result.found) {
          // Delete row from index
          const sheet = this.sheetAccessor.getSheet(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
          sheet.deleteRow(result.rowIndex);
        }
        
        // Handle Drive files if requested
        if (deleteDriveFiles && project.PROJECT_FOLDER_ID) {
          try {
            const folder = this.driveManager.getFolder(project.PROJECT_FOLDER_ID);
            if (folder) {
              folder.setTrashed(true);
              logInfo(`Moved project folder to trash: ${project.title} (${projectId})`);
            }
          } catch (driveError) {
            logWarning(`Could not delete Drive folder: ${driveError.message}`);
          }
        }
        
        logInfo(`Deleted project: ${project.title} (${projectId})`);
        
        return { 
          success: true, 
          message: 'Project deleted successfully',
          projectId: projectId,
          title: project.title,
          driveFilesDeleted: deleteDriveFiles
        };
      } catch (error) {
        logError(`Failed to delete project: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
    
    /**
     * Gets slides for a project
     * 
     * @param {string} projectTabName - Name of the project tab
     * @return {Array<Object>} Array of slide objects
     */
    getProjectSlides(projectTabName) {
      try {
        const sheet = this.sheetAccessor.getSheet(projectTabName, false);
        if (!sheet) return [];
        
        const slides = [];
        const data = this.sheetAccessor.getSheetData(projectTabName);
        
        // Process sheet data to find slide sections
        const slideHeaderPattern = /^SLIDE (\d+) INFO/;
        
        let currentSlide = null;
        let inSlideSection = false;
        
        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          
          if (row[0] && typeof row[0] === 'string') {
            const headerMatch = row[0].match(slideHeaderPattern);
            
            if (headerMatch) {
              // Found slide header, start collecting slide info
              if (currentSlide) {
                slides.push(currentSlide);
              }
              
              currentSlide = {
                slideNumber: parseInt(headerMatch[1]),
                slideId: '',
                title: '',
                backgroundColor: '',
                fileType: '',
                fileUrl: '',
                showControls: false
              };
              
              inSlideSection = true;
              continue;
            } else if (row[0] === 'ELEMENT INFO' || row[0] === 'TIMELINE' || row[0] === 'QUIZ' || row[0] === 'USER TRACKING') {
              // Found another section, end slide collection
              inSlideSection = false;
              if (currentSlide) {
                slides.push(currentSlide);
                currentSlide = null;
              }
              continue;
            }
          }
          
          if (inSlideSection && currentSlide && row[0] && row[1] !== undefined) {
            // Process slide fields
            const fieldName = row[0].toString().trim();
            const fieldValue = row[1];
            
            switch (fieldName) {
              case 'SLIDE_ID':
                currentSlide.slideId = fieldValue;
                break;
              case 'SLIDE_TITLE':
                currentSlide.title = fieldValue;
                break;
              case 'BACKGROUND_COLOR':
                currentSlide.backgroundColor = fieldValue;
                break;
              case 'FILE_TYPE':
                currentSlide.fileType = fieldValue;
                break;
              case 'FILE_URL':
                currentSlide.fileUrl = fieldValue;
                break;
              case 'SLIDE_NUMBER':
                currentSlide.slideNumber = fieldValue;
                break;
              case 'SHOW_CONTROLS':
                currentSlide.showControls = !!fieldValue;
                break;
            }
          }
        }
        
        // Add the last slide if we were still processing one
        if (currentSlide) {
          slides.push(currentSlide);
        }
        
        // Sort slides by slide number
        slides.sort((a, b) => a.slideNumber - b.slideNumber);
        
        return slides;
      } catch (error) {
        logError(`Failed to get project slides: ${error.message}`);
        return [];
      }
    }
    
    /**
     * Gets elements for a project
     * 
     * @param {string} projectTabName - Name of the project tab
     * @return {Array<Object>} Array of element objects
     */
    getProjectElements(projectTabName) {
      try {
        const sheet = this.sheetAccessor.getSheet(projectTabName, false);
        if (!sheet) return [];
        
        const elements = [];
        
        // Find ELEMENT INFO section
        const elementInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO;
        
        // Get element header row (which contains element names like "Element 1", "Element 2", etc.)
        const headerRow = elementInfoStructure.ELEMENT_HEADERS_ROW;
        const headerRange = sheet.getRange(headerRow, 5, 1, sheet.getMaxColumns() - 4); // Start from column E (5)
        const headerValues = headerRange.getValues()[0];
        
        // Get field data rows
        const fields = elementInfoStructure.FIELDS;
        const elementColumns = {};
        
        // Find valid element columns (ones with element IDs)
        for (let i = 0; i < headerValues.length; i++) {
          const elementHeaderValue = headerValues[i];
          if (elementHeaderValue) {
            const elementIdCell = sheet.getRange(fields.ELEMENT_ID.ROW, i + 5).getValue();
            if (elementIdCell) {
              elementColumns[i + 5] = elementHeaderValue;
            }
          }
        }
        
        // Process each column with element data
        for (const [colIndex, elementName] of Object.entries(elementColumns)) {
          const col = parseInt(colIndex);
          const element = { name: elementName };
          
          // Extract all field values for this element
          for (const [fieldName, fieldInfo] of Object.entries(fields)) {
            const rowIndex = fieldInfo.ROW;
            const value = sheet.getRange(rowIndex, col).getValue();
            
            // Convert field name to camelCase property name
            const propName = fieldName.toLowerCase().replace(/_(.)/g, (match, char) => char.toUpperCase());
            element[propName] = value;
          }
          
          // Skip elements without IDs
          if (!element.elementId) continue;
          
          // Get timeline data for this element
          const timelineStructure = SHEET_STRUCTURE.PROJECT_TAB.TIMELINE;
          const timeline = {};
          
          for (const [fieldName, fieldInfo] of Object.entries(timelineStructure.FIELDS)) {
            const rowIndex = fieldInfo.ROW;
            
            // Only include timeline entries that have this element ID
            const timelineElementId = sheet.getRange(rowIndex, 4).getValue(); // Element ID in column D
            if (timelineElementId === element.elementId) {
              const value = sheet.getRange(rowIndex, col).getValue();
              
              // Convert field name to camelCase property name
              const propName = fieldName.toLowerCase().replace(/_(.)/g, (match, char) => char.toUpperCase());
              timeline[propName] = value;
            }
          }
          
          if (Object.keys(timeline).length > 0) {
            element.timeline = timeline;
          }
          
          // Get quiz data for this element
          if (element.interactionType === 'Quiz') {
            const quizStructure = SHEET_STRUCTURE.PROJECT_TAB.QUIZ;
            const quiz = {};
            
            for (const [fieldName, fieldInfo] of Object.entries(quizStructure.FIELDS)) {
              const rowIndex = fieldInfo.ROW;
              const value = sheet.getRange(rowIndex, col).getValue();
              
              // Convert field name to camelCase property name
              const propName = fieldName.toLowerCase().replace(/_(.)/g, (match, char) => char.toUpperCase());
              quiz[propName] = value;
            }
            
            if (Object.keys(quiz).length > 0) {
              element.quiz = quiz;
            }
          }
          
          elements.push(element);
        }
        
        return elements;
      } catch (error) {
        logError(`Failed to get project elements: ${error.message}`);
        return [];
      }
    }
    
    /**
     * Adds a new slide to a project
     * 
     * @param {string} projectId - ID of the project
     * @param {Object} slideData - Data for the new slide
     * @return {Object} Result object with slide details or error information
     */
    addSlide(projectId, slideData = {}) {
      try {
        // Find the project
        const project = this.getProject(projectId);
        
        if (!project) {
          return { success: false, message: 'Project not found' };
        }
        
        const projectTabName = sanitizeSheetName(project.title);
        
        // Create new slide
        const slideInfo = {
          SLIDE_ID: slideData.slideId || generateUUID(),
          SLIDE_TITLE: slideData.title || `Slide ${project.slides.length + 1}`,
          BACKGROUND_COLOR: slideData.backgroundColor || DEFAULT_COLORS.BACKGROUND,
          FILE_TYPE: slideData.fileType || '',
          FILE_URL: slideData.fileUrl || '',
          SLIDE_NUMBER: slideData.slideNumber || (project.slides.length + 1),
          SHOW_CONTROLS: slideData.showControls || false
        };
        
        const success = this.templateManager.createNewSlide(projectTabName, slideInfo);
        
        if (!success) {
          return { success: false, message: 'Failed to create new slide' };
        }
        
        // Update project modified timestamp
        const now = new Date();
        this.sheetAccessor.setCellValue(
          projectTabName,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.ROW,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.VALUE_COL,
          now.getTime()
        );
        
        // Also update in project index
        const result = this.sheetAccessor.findProjectInIndex(projectId);
        if (result.found) {
          this.sheetAccessor.setCellValue(
            SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
            result.rowIndex,
            SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT + 1,
            now
          );
        }
        
        logInfo(`Added slide to project: ${projectId}, Slide: ${slideInfo.SLIDE_ID}`);
        
        const newSlide = {
          slideId: slideInfo.SLIDE_ID,
          slideNumber: slideInfo.SLIDE_NUMBER,
          title: slideInfo.SLIDE_TITLE,
          backgroundColor: slideInfo.BACKGROUND_COLOR,
          fileType: slideInfo.FILE_TYPE,
          fileUrl: slideInfo.FILE_URL,
          showControls: slideInfo.SHOW_CONTROLS
        };
        
        return { 
          success: true, 
          message: 'Slide added successfully',
          slide: newSlide,
          projectId: projectId
        };
      } catch (error) {
        logError(`Failed to add slide: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
    
    /**
     * Updates a slide in a project
     * 
     * @param {string} projectId - ID of the project
     * @param {string} slideId - ID of the slide to update
     * @param {Object} updates - Object with properties to update
     * @return {Object} Result object with success flag and message
     */
    updateSlide(projectId, slideId, updates) {
      try {
        // Find the project
        const project = this.getProject(projectId);
        
        if (!project) {
          return { success: false, message: 'Project not found' };
        }
        
        // Find the slide
        const slide = project.slides.find(s => s.slideId === slideId);
        
        if (!slide) {
          return { success: false, message: 'Slide not found' };
        }
        
        const projectTabName = sanitizeSheetName(project.title);
        const sheet = this.sheetAccessor.getSheet(projectTabName);
        
        // Find the slide section in the sheet
        const slideInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.SLIDE_INFO;
        const slideCount = project.slides.length;
        
        // Calculate the row offset for this slide
        const slideIndex = project.slides.findIndex(s => s.slideId === slideId);
        const rowsPerSlide = slideInfoStructure.SECTION_END_ROW - slideInfoStructure.SECTION_START_ROW + 1;
        const slideStartRow = slideInfoStructure.SECTION_START_ROW + (slideIndex * rowsPerSlide);
        
        // Update slide fields
        for (const [field, value] of Object.entries(updates)) {
          let rowIndex;
          
          switch (field) {
            case 'title':
              rowIndex = slideStartRow + (slideInfoStructure.FIELDS.SLIDE_TITLE.ROW - slideInfoStructure.SECTION_START_ROW);
              this.sheetAccessor.setCellValue(projectTabName, rowIndex, 'B', value);
              break;
            case 'backgroundColor':
              rowIndex = slideStartRow + (slideInfoStructure.FIELDS.BACKGROUND_COLOR.ROW - slideInfoStructure.SECTION_START_ROW);
              this.sheetAccessor.setCellValue(projectTabName, rowIndex, 'B', value);
              break;
            case 'fileType':
              rowIndex = slideStartRow + (slideInfoStructure.FIELDS.FILE_TYPE.ROW - slideInfoStructure.SECTION_START_ROW);
              this.sheetAccessor.setCellValue(projectTabName, rowIndex, 'B', value);
              break;
            case 'fileUrl':
              rowIndex = slideStartRow + (slideInfoStructure.FIELDS.FILE_URL.ROW - slideInfoStructure.SECTION_START_ROW);
              this.sheetAccessor.setCellValue(projectTabName, rowIndex, 'B', value);
              break;
            case 'slideNumber':
              rowIndex = slideStartRow + (slideInfoStructure.FIELDS.SLIDE_NUMBER.ROW - slideInfoStructure.SECTION_START_ROW);
              this.sheetAccessor.setCellValue(projectTabName, rowIndex, 'B', value);
              break;
            case 'showControls':
              rowIndex = slideStartRow + (slideInfoStructure.FIELDS.SHOW_CONTROLS.ROW - slideInfoStructure.SECTION_START_ROW);
              this.sheetAccessor.setCellValue(projectTabName, rowIndex, 'B', value);
              break;
          }
        }
        
        // Update project modified timestamp
        const now = new Date();
        this.sheetAccessor.setCellValue(
          projectTabName,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.ROW,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.VALUE_COL,
          now.getTime()
        );
        
        // Also update in project index
        const result = this.sheetAccessor.findProjectInIndex(projectId);
        if (result.found) {
          this.sheetAccessor.setCellValue(
            SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
            result.rowIndex,
            SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT + 1,
            now
          );
        }
        
        logInfo(`Updated slide in project: ${projectId}, Slide: ${slideId}`);
        
        return { 
          success: true, 
          message: 'Slide updated successfully',
          slideId: slideId,
          projectId: projectId,
          updatedFields: Object.keys(updates)
        };
      } catch (error) {
        logError(`Failed to update slide: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
    
    /**
     * Deletes a slide from a project
     * Note: This is complex as it requires shifting rows in the spreadsheet
     * This provides a basic implementation that marks the slide as deleted
     * A full implementation would require restructuring the sheet
     * 
     * @param {string} projectId - ID of the project
     * @param {string} slideId - ID of the slide to delete
     * @return {Object} Result object with success flag and message
     */
    deleteSlide(projectId, slideId) {
      try {
        // Find the project
        const project = this.getProject(projectId);
        
        if (!project) {
          return { success: false, message: 'Project not found' };
        }
        
        // Find the slide
        const slide = project.slides.find(s => s.slideId === slideId);
        
        if (!slide) {
          return { success: false, message: 'Slide not found' };
        }
        
        const projectTabName = sanitizeSheetName(project.title);
        
        // For now, just mark the slide as "deleted" by prefixing its title
        // A full implementation would need to restructure the sheet
        this.updateSlide(projectId, slideId, {
          title: `[DELETED] ${slide.title}`
        });
        
        logInfo(`Marked slide as deleted in project: ${projectId}, Slide: ${slideId}`);
        
        return { 
          success: true, 
          message: 'Slide marked as deleted successfully',
          slideId: slideId,
          projectId: projectId
        };
      } catch (error) {
        logError(`Failed to delete slide: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
    
    /**
     * Adds a new element to a project
     * 
     * @param {string} projectId - ID of the project
     * @param {string} slideId - ID of the slide to add the element to
     * @param {Object} elementData - Data for the new element
     * @return {Object} Result object with element details or error information
     */
    addElement(projectId, slideId, elementData = {}) {
      try {
        // Find the project
        const project = this.getProject(projectId);
        
        if (!project) {
          return { success: false, message: 'Project not found' };
        }
        
        // Find the slide
        const slide = project.slides.find(s => s.slideId === slideId);
        
        if (!slide) {
          return { success: false, message: 'Slide not found' };
        }
        
        const projectTabName = sanitizeSheetName(project.title);
        
        // Prepare element data
        const elementInfo = {
          ELEMENT_ID: elementData.elementId || generateUUID(),
          NICKNAME: elementData.name || `Element ${project.elements.length + 1}`,
          SLIDE_ID: slideId,
          SEQUENCE: elementData.sequence || (project.elements.filter(e => e.slideId === slideId).length + 1),
          TYPE: elementData.type || ELEMENT_TYPES.RECTANGLE.LABEL,
          LEFT: elementData.left || 100,
          TOP: elementData.top || 100,
          WIDTH: elementData.width || ELEMENT_TYPES.RECTANGLE.DEFAULT_WIDTH,
          HEIGHT: elementData.height || ELEMENT_TYPES.RECTANGLE.DEFAULT_HEIGHT,
          ANGLE: elementData.angle || 0,
          INITIALLY_HIDDEN: elementData.initiallyHidden || false,
          OPACITY: elementData.opacity !== undefined ? elementData.opacity : 100,
          COLOR: elementData.color || DEFAULT_COLORS.ELEMENT,
          OUTLINE: elementData.outline || false,
          OUTLINE_WIDTH: elementData.outlineWidth || 1,
          OUTLINE_COLOR: elementData.outlineColor || DEFAULT_COLORS.OUTLINE,
          SHADOW: elementData.shadow || false,
          TEXT: elementData.text || '',
          FONT: elementData.font || FONTS[0],
          FONT_COLOR: elementData.fontColor || DEFAULT_COLORS.TEXT,
          FONT_SIZE: elementData.fontSize || 14,
          TRIGGERS: elementData.triggers || TRIGGER_TYPES.CLICK,
          INTERACTION_TYPE: elementData.interactionType || INTERACTION_TYPES.REVEAL,
          TEXT_MODAL: elementData.textModal || false,
          TEXT_MODAL_MESSAGE: elementData.textModalMessage || '',
          ANIMATION_TYPE: elementData.animationType || ANIMATION_TYPES.NONE,
          ANIMATION_SPEED: elementData.animationSpeed || ANIMATION_SPEEDS.MEDIUM
        };
        
        // Create new element column
        const success = this.templateManager.createNewElement(projectTabName, elementInfo);
        
        if (!success) {
          return { success: false, message: 'Failed to create new element' };
        }
        
        // Update project modified timestamp
        const now = new Date();
        this.sheetAccessor.setCellValue(
          projectTabName,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.ROW,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.VALUE_COL,
          now.getTime()
        );
        
        // Also update in project index
        const result = this.sheetAccessor.findProjectInIndex(projectId);
        if (result.found) {
          this.sheetAccessor.setCellValue(
            SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
            result.rowIndex,
            SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT + 1,
            now
          );
        }
        
        logInfo(`Added element to project: ${projectId}, Element: ${elementInfo.ELEMENT_ID}`);
        
        // Convert field names to camelCase for returned object
        const newElement = {};
        for (const [key, value] of Object.entries(elementInfo)) {
          const camelKey = key.toLowerCase().replace(/_(.)/g, (match, char) => char.toUpperCase());
          newElement[camelKey] = value;
        }
        
        return { 
          success: true, 
          message: 'Element added successfully',
          element: newElement,
          projectId: projectId,
          slideId: slideId
        };
      } catch (error) {
        logError(`Failed to add element: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
    
    /**
     * Updates an element in a project
     * 
     * @param {string} projectId - ID of the project
     * @param {string} elementId - ID of the element to update
     * @param {Object} updates - Object with properties to update
     * @return {Object} Result object with success flag and message
     */
    updateElement(projectId, elementId, updates) {
      try {
        // Find the project
        const project = this.getProject(projectId);
        
        if (!project) {
          return { success: false, message: 'Project not found' };
        }
        
        // Find the element
        const element = project.elements.find(e => e.elementId === elementId);
        
        if (!element) {
          return { success: false, message: 'Element not found' };
        }
        
        const projectTabName = sanitizeSheetName(project.title);
        const sheet = this.sheetAccessor.getSheet(projectTabName);
        
        // Find the element column in the sheet
        const elementInfoStructure = SHEET_STRUCTURE.PROJECT_TAB.ELEMENT_INFO;
        let elementColumn = -1;
        
        // Find which column contains this element
        const elementIdRow = elementInfoStructure.FIELDS.ELEMENT_ID.ROW;
        for (let col = 5; col <= sheet.getMaxColumns(); col++) { // Start from column E (5)
          const cellValue = sheet.getRange(elementIdRow, col).getValue();
          if (cellValue === elementId) {
            elementColumn = col;
            break;
          }
        }
        
        if (elementColumn === -1) {
          return { success: false, message: 'Element column not found in sheet' };
        }
        
        // Update element fields
        for (const [field, value] of Object.entries(updates)) {
          // Convert camelCase field name to UPPER_SNAKE_CASE for spreadsheet
          const sheetField = field.replace(/([A-Z])/g, '_$1').toUpperCase();
          
          if (elementInfoStructure.FIELDS[sheetField]) {
            const rowIndex = elementInfoStructure.FIELDS[sheetField].ROW;
            this.sheetAccessor.setCellValue(projectTabName, rowIndex, elementColumn, value);
          }
        }
        
        // Update project modified timestamp
        const now = new Date();
        this.sheetAccessor.setCellValue(
          projectTabName,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.ROW,
          SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.MODIFIED_AT.VALUE_COL,
          now.getTime()
        );
        
        // Also update in project index
        const result = this.sheetAccessor.findProjectInIndex(projectId);
        if (result.found) {
          this.sheetAccessor.setCellValue(
            SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
            result.rowIndex,
            SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT + 1,
            now
          );
        }
        
        logInfo(`Updated element in project: ${projectId}, Element: ${elementId}`);
        
        return { 
          success: true, 
          message: 'Element updated successfully',
          elementId: elementId,
          projectId: projectId,
          updatedFields: Object.keys(updates)
        };
      } catch (error) {
        logError(`Failed to update element: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
    
    /**
     * Deletes an element from a project
     * Note: This is complex as it requires shifting columns in the spreadsheet
     * This provides a basic implementation that marks the element as deleted
     * A full implementation would require restructuring the sheet
     * 
     * @param {string} projectId - ID of the project
     * @param {string} elementId - ID of the element to delete
     * @return {Object} Result object with success flag and message
     */
    deleteElement(projectId, elementId) {
      try {
        // Find the project
        const project = this.getProject(projectId);
        
        if (!project) {
          return { success: false, message: 'Project not found' };
        }
        
        // Find the element
        const element = project.elements.find(e => e.elementId === elementId);
        
        if (!element) {
          return { success: false, message: 'Element not found' };
        }
        
        // For now, just mark the element as "deleted" by setting opacity to 0 and renaming it
        this.updateElement(projectId, elementId, {
          nickname: `[DELETED] ${element.nickname || 'Element'}`,
          opacity: 0,
          initiallyHidden: true
        });
        
        logInfo(`Marked element as deleted in project: ${projectId}, Element: ${elementId}`);
        
        return { 
          success: true, 
          message: 'Element marked as deleted successfully',
          elementId: elementId,
          projectId: projectId
        };
      } catch (error) {
        logError(`Failed to delete element: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
    
    /**
     * Deploys a project as a web app
     * 
     * @param {string} projectId - ID of the project to deploy
     * @return {Object} Result object with deployment details or error information
     */
    deployProject(projectId) {
      try {
        // Find the project
        const project = this.getProject(projectId);
        
        if (!project) {
          return { success: false, message: 'Project not found' };
        }
        
        // Get current script
        const scriptId = ScriptApp.getScriptId();
        const scriptUrl = ScriptApp.getService().getUrl();
        
        // Create web app URL with project ID parameter
        const webAppUrl = `${scriptUrl}?project=${projectId}`;
        
        // Update project with web app URL
        this.updateProject(projectId, {
          webAppUrl: webAppUrl
        });
        
        logInfo(`Deployed project as web app: ${projectId}, URL: ${webAppUrl}`);
        
        return { 
          success: true, 
          message: 'Project deployed successfully',
          projectId: projectId,
          webAppUrl: webAppUrl,
          scriptId: scriptId
        };
      } catch (error) {
        logError(`Failed to deploy project: ${error.message}`);
        return { success: false, message: `Error: ${error.message}` };
      }
    }
  }