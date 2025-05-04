/**
 * Interactive Training Projects Web App
 * Main entry point, onOpen trigger, and menu setup
 */

// Global instances of core classes
let sheetAccessor = null;
let templateManager = null;
let driveManager = null;
let mediaProcessor = null;
let fileUploader = null;
let projectManager = null;
let authManager = null;
let apiHandler = null;

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
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Project Tools')
        .addItem('Add Slide', 'showAddSlideDialog')
        .addItem('Add Element', 'showAddElementDialog')
        .addItem('Upload Media', 'showFileUploadDialog')
        .addItem('Deploy Project', 'showDeployProjectDialog'))
      .addSeparator()
      .addItem('Manage Files', 'showFileManager')
      .addSeparator()
      .addItem('View Analytics', 'showAnalyticsDashboard')
      .addSeparator()
      .addItem('Locate Root Folder', 'showRootFolderURL')
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
    
    // Initialize Phase 2 components
    mediaProcessor = new MediaProcessor(driveManager);
    fileUploader = new FileUploader(driveManager, mediaProcessor);
    projectManager = new ProjectManager(sheetAccessor, templateManager, driveManager);
    
    // Initialize Phase 3 components
    authManager = new AuthManager(sheetAccessor, driveManager);
    apiHandler = new ApiHandler(projectManager, fileUploader, authManager);
    
    logInfo('Application successfully initialized');
    return true;
  } catch (error) {
    logError(`Failed to initialize application: ${error.message}\n${error.stack}`);
    // Re-throw to ensure the caller knows initialization failed
    throw new Error(`Initialization failed: ${error.message}`);
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
        #status {
          margin-top: 15px;
          padding: 10px;
          border-radius: 4px;
          display: none;
        }
        .success {
          background-color: #d4edda;
          color: #155724;
          border: 1px solid #c3e6cb;
        }
        .error {
          background-color: #f8d7da;
          color: #721c24;
          border: 1px solid #f5c6cb;
        }
        .loading {
          display: inline-block;
          width: 16px;
          height: 16px;
          border: 3px solid rgba(0, 0, 0, 0.1);
          border-radius: 50%;
          border-top-color: #3498db;
          animation: spin 1s ease-in-out infinite;
          margin-right: 10px;
        }
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
        .hidden {
          display: none;
        }
      </style>
      <div>
        <h2>Create New Training Project</h2>
        <div class="form-group">
          <label for="projectName">Project Name:</label>
          <input type="text" id="projectName" placeholder="Enter project name">
        </div>
        <div id="status"></div>
        <div class="buttons">
          <button id="cancelBtn" onclick="cancelDialog()">Cancel</button>
          <button id="createBtn" onclick="createProject()">
            <span id="loadingIcon" class="loading hidden"></span>
            <span id="buttonText">Create Project</span>
          </button>
        </div>
      </div>
      <script>
        function createProject() {
          const projectName = document.getElementById('projectName').value;
          if (!projectName) {
            showStatus('Please enter a project name', 'error');
            return;
          }
          
          // Show loading state
          setLoading(true);
          
          google.script.run
            .withSuccessHandler(onSuccess)
            .withFailureHandler(onFailure)
            .createNewProject(projectName);
        }
        
        function onSuccess(result) {
          setLoading(false);
          
          if (result.success) {
            showStatus('Project created successfully! The dialog will close in 3 seconds.', 'success');
            // Close dialog after a short delay to ensure the message is seen
            setTimeout(function() {
              google.script.host.close();
            }, 3000);
          } else {
            showStatus('Failed to create project: ' + result.message, 'error');
          }
        }
        
        function onFailure(error) {
          setLoading(false);
          showStatus('Error: ' + error.message, 'error');
        }
        
        function cancelDialog() {
          google.script.host.close();
        }
        
        function showStatus(message, type) {
          const statusElement = document.getElementById('status');
          statusElement.textContent = message;
          statusElement.className = type;
          statusElement.style.display = 'block';
        }
        
        function setLoading(isLoading) {
          const loadingIcon = document.getElementById('loadingIcon');
          const buttonText = document.getElementById('buttonText');
          const createBtn = document.getElementById('createBtn');
          const cancelBtn = document.getElementById('cancelBtn');
          
          if (isLoading) {
            loadingIcon.classList.remove('hidden');
            buttonText.textContent = 'Creating...';
            createBtn.disabled = true;
            cancelBtn.disabled = true;
          } else {
            loadingIcon.classList.add('hidden');
            buttonText.textContent = 'Create Project';
            createBtn.disabled = false;
            cancelBtn.disabled = false;
          }
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
    // Check if the application is initialized
    if (!projectManager) {
      initialize();
    }
    
    // Use the ProjectManager to create the project
    const result = projectManager.createProject(projectName);
    
    // If successful, activate the project tab
    if (result.success) {
      const sheet = sheetAccessor.getSheet(result.projectTabName);
      sheet.activate();
    }
    
    return result;
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
    if (!projectManager) {
      initialize();
    }
    
    // Use the ProjectManager to open the project
    return projectManager.openProject(projectId);
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

/**
 * Shows the URL of the root project folder
 */
function showRootFolderURL() {
  try {
    if (!driveManager) {
      initialize();
    }
    
    const rootFolder = driveManager.getRootFolder();
    const folderURL = rootFolder.getUrl();
    const folderName = rootFolder.getName();
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Root Folder Location', 
             `Folder: ${folderName}\n\nURL: ${folderURL}\n\nClick OK to open the folder.`,
             ui.ButtonSet.OK);
    
    // Open the folder in a new tab
    const html = HtmlService.createHtmlOutput(
      `<script>window.open('${folderURL}', '_blank');google.script.host.close();</script>`
    )
    .setWidth(10)
    .setHeight(10);
    
    ui.showModalDialog(html, 'Opening folder...');
    
  } catch (error) {
    logError(`Failed to show root folder URL: ${error.message}`);
    showErrorAlert('Failed to locate the root folder. Please check the logs for details.');
  }
}

/**
 * Processes an API request from the web client
 * Acts as a bridge between the client-side code and the ApiHandler
 * 
 * @param {Object} requestData - Request data from client
 * @return {Object} Response data
 */
function processApiRequest(requestData) {
  try {
    console.log('Processing API request:', requestData);
    
    // Check if the application is initialized
    if (!apiHandler || !projectManager || !authManager) {
      console.log('Components not initialized, initializing...');
      initialize();
    }
    
    // Double-check initialization was successful
    if (!apiHandler || !projectManager || !authManager) {
      throw new Error('Failed to initialize components');
    }
    
    // Get the current user
    const user = Session.getEffectiveUser().getEmail();
    console.log('User making request:', user);
    
    // Handle the request with internal error catching
    let result;
    try {
      result = apiHandler.handleRequest(requestData, user);
      console.log('Request handled successfully, result:', result);
    } catch (handlerError) {
      console.error('Error in request handler:', handlerError);
      return {
        success: false,
        error: `Handler error: ${handlerError.message}`
      };
    }
    
    // If result is already a ContentService TextOutput, return its content
    if (result && typeof result.getContent === 'function') {
      // Extract the JSON content from TextOutput and return as object
      try {
        return JSON.parse(result.getContent());
      } catch (e) {
        return {
          success: true,
          data: result.getContent()
        };
      }
    }
    
    // For normal object results, ensure proper format
    return {
      success: true,
      data: result
    };
  } catch (error) {
    console.error('Error processing API request:', error);
    logError(`Error processing API request: ${error.message}\n${error.stack}`);
    return {
      success: false,
      error: error.message || 'Unknown server error'
    };
  }
}

/**
 * Requests access to a project
 * Sends an email to the project owner
 * 
 * @param {string} projectId - ID of the project
 * @param {string} requesterEmail - Email of the user requesting access
 * @return {Object} Result of the request
 */
function requestProjectAccess(projectId, requesterEmail) {
  try {
    // Check if the application is initialized
    if (!projectManager) {
      initialize();
    }
    
    // Get project info
    const project = projectManager.getProject(projectId);
    
    if (!project) {
      return { 
        success: false, 
        message: 'Project not found'
      };
    }
    
    // Get project owner email
    let ownerEmail = '';
    
    if (project.folderId) {
      try {
        const folder = DriveApp.getFolderById(project.folderId);
        ownerEmail = folder.getOwner().getEmail();
      } catch (folderError) {
        // If can't access folder, use spreadsheet owner
        ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
      }
    } else {
      ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
    }
    
    // Send email
    const subject = `Access Request: ${project.title} Training Project`;
    const body = `
      Hello,
      
      ${requesterEmail} has requested access to your interactive training project "${project.title}".
      
      To grant access, open the project in Google Drive and share it with them.
      
      Project ID: ${projectId}
      
      This is an automated message from the Interactive Training Projects Web App.
    `;
    
    MailApp.sendEmail(ownerEmail, subject, body);
    
    return {
      success: true,
      message: 'Access request sent successfully'
    };
  } catch (error) {
    logError(`Error requesting project access: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}