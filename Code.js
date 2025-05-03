/**
 * Interactive Training Projects Web App
 * Main entry point, onOpen trigger, and menu setup
 */

// Global instances of core classes
let sheetAccessor = null;
let templateManager = null;
let driveManager = null;

/**
 * Function that runs when the spreadsheet is opened
 * Sets up custom menu and initializes the application
 * 
 * @param {Object} e - Event object from Google Apps Script
 */
function onOpen(e) {
  try {
    // Create menu
    SpreadsheetApp.getUi()
      .createMenu('Training Projects')
      .addItem('Create New Project', 'showNewProjectDialog')
      .addSeparator()
      .addItem('Edit Project', 'showProjectSelector')
      .addItem('View Project', 'showProjectViewer')
      .addSeparator()
      .addItem('Manage Files', 'showFileManager')
      .addSeparator()
      .addItem('View Analytics', 'showAnalyticsDashboard')
      .addToUi();
    
    // Initialize the application
    initialize();
    
    logInfo('Application initialized on spreadsheet open');
  } catch (error) {
    logError(`Error in onOpen: ${error.message}`);
    showErrorAlert('Failed to initialize the application. Please refresh and try again.');
  }
}

/**
 * Initializes the application by creating the required global instances
 * Creates necessary structures if they don't exist
 */
function initialize() {
  try {
    // Initialize SheetAccessor
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    sheetAccessor = new SheetAccessor(spreadsheetId);
    
    // Check if project index exists, create if not
    const sheets = sheetAccessor.getSheetNames();
    if (!sheets.includes(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME)) {
      sheetAccessor.initializeProjectIndexTab();
    }
    
    // Check if template exists, create if not
    if (!sheets.includes('Template')) {
      // Initialize TemplateManager
      templateManager = new TemplateManager(sheetAccessor);
      templateManager.createBaseTemplate();
    } else {
      // Template exists, just initialize the manager
      templateManager = new TemplateManager(sheetAccessor);
    }
    
    // Initialize DriveManager
    driveManager = new DriveManager();
    
    logInfo('Application successfully initialized');
    return true;
  } catch (error) {
    logError(`Failed to initialize application: ${error.message}`);
    return false;
  }
}

/**
 * Shows a modal dialog with an error message
 * 
 * @param {string} message - Error message to display
 */
function showErrorAlert(message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Error', message, ui.ButtonSet.OK);
}

/**
 * Shows a dialog to create a new project
 */
function showNewProjectDialog() {
  try {
    // Check if the application is initialized
    if (!sheetAccessor || !templateManager || !driveManager) {
      initialize();
    }
    
    // Create HTML template
    const htmlOutput = HtmlService.createHtmlOutput(`
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 20px;
        }
        .form-group {
          margin-bottom: 15px;
        }
        label {
          display: block;
          margin-bottom: 5px;
          font-weight: bold;
        }
        input[type=text] {
          width: 100%;
          padding: 8px;
          box-sizing: border-box;
        }
        .buttons {
          display: flex;
          justify-content: flex-end;
          margin-top: 20px;
        }
        .buttons button {
          padding: 8px 16px;
          margin-left: 10px;
        }
      </style>
      <div>
        <h2>Create New Training Project</h2>
        <div class="form-group">
          <label for="projectName">Project Name:</label>
          <input type="text" id="projectName" placeholder="Enter project name">
        </div>
        <div class="buttons">
          <button onclick="google.script.host.close()">Cancel</button>
          <button onclick="createProject()">Create Project</button>
        </div>
      </div>
      <script>
        function createProject() {
          const projectName = document.getElementById('projectName').value;
          if (!projectName) {
            alert('Please enter a project name');
            return;
          }
          
          google.script.run
            .withSuccessHandler(onSuccess)
            .withFailureHandler(onFailure)
            .createNewProject(projectName);
        }
        
        function onSuccess(result) {
          if (result.success) {
            alert('Project created successfully!');
            google.script.host.close();
          } else {
            alert('Failed to create project: ' + result.message);
          }
        }
        
        function onFailure(error) {
          alert('Error: ' + error.message);
        }
      </script>
    `)
      .setWidth(400)
      .setHeight(250)
      .setTitle('Create New Project');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Create New Project');
  } catch (error) {
    logError(`Failed to show new project dialog: ${error.message}`);
    showErrorAlert('Failed to open the new project dialog. Please try again.');
  }
}

/**
 * Creates a new project with the specified name
 * Called from the new project dialog
 * 
 * @param {string} projectName - Name for the new project
 * @return {Object} Result object with success flag and message
 */
function createNewProject(projectName) {
  try {
    // Validate project name
    if (!projectName || projectName.trim() === '') {
      return { success: false, message: 'Project name cannot be empty' };
    }
    
    // Check if the application is initialized
    if (!sheetAccessor || !templateManager || !driveManager) {
      initialize();
    }
    
    // Create project tab from template
    const projectTabName = templateManager.createProjectFromTemplate(projectName);
    
    if (!projectTabName) {
      return { success: false, message: 'Failed to create project from template' };
    }
    
    // Get project info
    const projectInfo = templateManager.getProjectInfo(projectTabName);
    
    if (!projectInfo || !projectInfo.PROJECT_ID) {
      return { success: false, message: 'Failed to get project information' };
    }
    
    // Create project folders in Drive
    const folders = driveManager.createProjectFolders(projectInfo.PROJECT_ID, projectName);
    
    if (!folders || !folders.projectFolderId) {
      return { success: false, message: 'Failed to create project folders in Drive' };
    }
    
    // Update project info with folder ID
    projectInfo.PROJECT_FOLDER_ID = folders.projectFolderId;
    sheetAccessor.updateProjectInfo(projectTabName, projectInfo);
    
    // Create web app URL if needed
    if (!projectInfo.PROJECT_WEB_APP_URL) {
      // This will be implemented in a later phase when the web app is set up
      // For now, we'll leave it blank
    }
    
    // Activate the project tab
    const sheet = sheetAccessor.getSheet(projectTabName);
    sheet.activate();
    
    logInfo(`Successfully created new project: ${projectName} (${projectInfo.PROJECT_ID})`);
    
    return { 
      success: true, 
      message: 'Project created successfully',
      projectId: projectInfo.PROJECT_ID,
      projectTabName: projectTabName
    };
  } catch (error) {
    logError(`Failed to create new project: ${error.message}`);
    return { success: false, message: `Error: ${error.message}` };
  }
}

/**
 * Shows a dialog to select an existing project for editing
 */
function showProjectSelector() {
  try {
    // Check if the application is initialized
    if (!sheetAccessor || !templateManager || !driveManager) {
      initialize();
    }
    
    // Get list of projects
    const projectIndexSheet = sheetAccessor.getSheet(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME, false);
    if (!projectIndexSheet) {
      showErrorAlert('Project index not found. Please initialize the application first.');
      return;
    }
    
    const projectData = sheetAccessor.getSheetData(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
    
    // Skip header row and filter out any empty rows
    const projects = projectData.slice(1)
      .filter(row => row[0] && row[1]) // Ensure projectId and title exist
      .map(row => {
        return {
          projectId: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID],
          title: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE],
          createdAt: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.CREATED_AT],
          modifiedAt: row[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.MODIFIED_AT]
        };
      });
    
    // Create HTML for project selector
    let projectListHtml = '';
    if (projects.length === 0) {
      projectListHtml = '<p>No projects found. Please create a new project first.</p>';
    } else {
      projectListHtml = '<div class="project-list">';
      projects.forEach(project => {
        const createdDate = project.createdAt ? new Date(project.createdAt).toLocaleDateString() : 'Unknown';
        const modifiedDate = project.modifiedAt ? new Date(project.modifiedAt).toLocaleDateString() : 'Unknown';
        
        projectListHtml += `
          <div class="project-item" data-project-id="${project.projectId}">
            <div class="project-title">${project.title}</div>
            <div class="project-dates">
              Created: ${createdDate} | Last Modified: ${modifiedDate}
            </div>
            <button onclick="editProject('${project.projectId}')">Edit</button>
          </div>
        `;
      });
      projectListHtml += '</div>';
    }
    
    // Create HTML template
    const htmlOutput = HtmlService.createHtmlOutput(`
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 20px;
        }
        h2 {
          margin-bottom: 20px;
        }
        .project-list {
          max-height: 300px;
          overflow-y: auto;
        }
        .project-item {
          border: 1px solid #ddd;
          border-radius: 4px;
          padding: 10px;
          margin-bottom: 10px;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        .project-title {
          font-weight: bold;
          flex: 2;
        }
        .project-dates {
          font-size: 0.8em;
          color: #666;
          flex: 2;
        }
        button {
          padding: 5px 10px;
          background-color: #4285F4;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
        }
        button:hover {
          background-color: #3367D6;
        }
        .footer {
          margin-top: 20px;
          display: flex;
          justify-content: space-between;
        }
      </style>
      <div>
        <h2>Select Project to Edit</h2>
        ${projectListHtml}
        <div class="footer">
          <button onclick="google.script.host.close()">Cancel</button>
          <button onclick="createNewProjectClick()">Create New Project</button>
        </div>
      </div>
      <script>
        function editProject(projectId) {
          google.script.run
            .withSuccessHandler(onSuccess)
            .withFailureHandler(onFailure)
            .openProjectForEditing(projectId);
        }
        
        function onSuccess(result) {
          google.script.host.close();
        }
        
        function onFailure(error) {
          alert('Error: ' + error.message);
        }
        
        function createNewProjectClick() {
          google.script.host.close();
          google.script.run.showNewProjectDialog();
        }
      </script>
    `)
      .setWidth(500)
      .setHeight(400)
      .setTitle('Select Project');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Project');
  } catch (error) {
    logError(`Failed to show project selector: ${error.message}`);
    showErrorAlert('Failed to open the project selector. Please try again.');
  }
}

/**
 * Opens a project for editing
 * Called from the project selector dialog
 * 
 * @param {string} projectId - ID of the project to open
 * @return {Object} Result object with success flag and message
 */
function openProjectForEditing(projectId) {
  try {
    // Check if the application is initialized
    if (!sheetAccessor || !templateManager || !driveManager) {
      initialize();
    }
    
    // Find the project in the index
    const result = sheetAccessor.findProjectInIndex(projectId);
    
    if (!result.found) {
      return { success: false, message: 'Project not found in index' };
    }
    
    // Get project data from index
    const projectData = sheetAccessor.getSheetData(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
    const projectRow = projectData[result.rowIndex - 1]; // Adjust for 0-based array
    const projectTitle = projectRow[SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE];
    
    // Find project tab by name
    const sanitizedTitle = sanitizeSheetName(projectTitle);
    const sheet = sheetAccessor.getSheet(sanitizedTitle, false);
    
    if (!sheet) {
      return { success: false, message: 'Project sheet not found' };
    }
    
    // Activate the project tab
    sheet.activate();
    
    // Update last accessed timestamp
    const now = new Date();
    sheetAccessor.setCellValue(
      SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME,
      result.rowIndex,
      SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.LAST_ACCESSED + 1,
      now
    );
    
    logInfo(`Opened project for editing: ${projectTitle} (${projectId})`);
    
    return { 
      success: true, 
      message: 'Project opened successfully',
      projectId: projectId,
      projectTitle: projectTitle
    };
  } catch (error) {
    logError(`Failed to open project for editing: ${error.message}`);
    return { success: false, message: `Error: ${error.message}` };
  }
}

/**
 * Other functions that will be implemented in later phases
 */
function showProjectViewer() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Coming Soon', 'Project viewer will be implemented in a future phase.', ui.ButtonSet.OK);
}

function showFileManager() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Coming Soon', 'File manager will be implemented in a future phase.', ui.ButtonSet.OK);
}

function showAnalyticsDashboard() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Coming Soon', 'Analytics dashboard will be implemented in a future phase.', ui.ButtonSet.OK);
}