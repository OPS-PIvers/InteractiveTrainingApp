/**
 * Interactive Training Projects Web App
 * Main entry point, onOpen trigger, and menu setup
 */

// Global instances of core classes - Initialized in initialize()
let sheetAccessor = null;
let templateManager = null;
let driveManager = null;
let mediaProcessor = null; // Assuming defined elsewhere if used
let fileUploader = null;   // Assuming defined elsewhere if used
let projectManager = null;
let authManager = null;    // Assuming defined elsewhere if used
let apiHandler = null;

/**
 * Runs when the spreadsheet is opened. Creates the application menu.
 * @param {Event} e The open event object.
 */
function onOpen(e) {
  try {
    // Create menu only if the context allows UI interaction
    if (e && e.authMode !== ScriptApp.AuthMode.NONE) {
        SpreadsheetApp.getUi()
          .createMenu('Training Projects')
          .addItem('Create New Project', 'showNewProjectDialog')
          .addSeparator()
          .addItem('Edit Project', 'showProjectSelector')
          // .addItem('View Project', 'showProjectViewer') // Placeholder
          .addSeparator()
          // .addSubMenu(SpreadsheetApp.getUi().createMenu('Project Tools') // Placeholder
          //   .addItem('Add Slide', 'showAddSlideDialog')
          //   .addItem('Add Element', 'showAddElementDialog')
          //   .addItem('Upload Media', 'showFileUploadDialog')
          //   .addItem('Deploy Project', 'showDeployProjectDialog'))
          // .addSeparator()
          // .addItem('Manage Files', 'showFileManager') // Placeholder
          // .addSeparator()
          // .addItem('View Analytics', 'showAnalyticsDashboard') // Placeholder
          // .addSeparator()
          .addItem('Locate Root Folder', 'showRootFolderURL')
          .addToUi();
    }
    
    // Initialize the application backend components
    initialize();
    
    logInfo('Application menu created and backend initialized on spreadsheet open.');
  } catch (error) {
    logError(`Error in onOpen: ${error.message}\n${error.stack}`);
    // Avoid showing UI alert here if authMode is NONE (e.g., script running automatically)
    if (e && e.authMode !== ScriptApp.AuthMode.NONE) {
        showErrorAlert('Failed to initialize the application menu. Please refresh and try again.');
    }
  }
}

/**
 * Initializes the application by creating the required global instances.
 * Creates necessary structures (like Template sheet) if they don't exist.
 * Ensures idempotency (can be called multiple times).
 */
function initialize() {
    // Prevent re-initialization if already done
    if (apiHandler && projectManager && templateManager && sheetAccessor && driveManager && authManager) {
        logDebug("Application already initialized.");
        return true;
    }

    logInfo("Initializing application components...");
    try {
        // Initialize SheetAccessor first
        const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
        sheetAccessor = new SheetAccessor(spreadsheetId);
        
        // Initialize TemplateManager (depends on SheetAccessor)
        // This will also ensure the base template sheet exists
        templateManager = new TemplateManager(sheetAccessor);
        templateManager.createBaseTemplate(); // Ensure template exists

        // Initialize DriveManager (independent)
        driveManager = new DriveManager(); // Assuming defined elsewhere

        // Initialize ProjectManager (depends on SheetAccessor, TemplateManager, DriveManager)
        projectManager = new ProjectManager(sheetAccessor, templateManager, driveManager);

        // Initialize AuthManager (depends on SheetAccessor, DriveManager)
        authManager = new AuthManager(sheetAccessor, driveManager); // Assuming defined elsewhere

        // Initialize FileUploader and MediaProcessor (if used, assuming defined elsewhere)
        // mediaProcessor = new MediaProcessor(driveManager);
        // fileUploader = new FileUploader(driveManager, mediaProcessor);
        fileUploader = fileUploader || {}; // Placeholder if not fully implemented yet

        // Initialize ApiHandler (depends on ProjectManager, FileUploader, AuthManager)
        apiHandler = new ApiHandler(projectManager, fileUploader, authManager);
        
        // Check if project index exists, create if not
        const indexSheetName = SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME;
        if (!sheetAccessor.getSheet(indexSheetName, false)) {
            logInfo(`Project index sheet '${indexSheetName}' not found. Initializing...`);
            sheetAccessor.initializeProjectIndexTab(indexSheetName);
        }

        // Verify all critical components were created successfully
        if (!sheetAccessor || !templateManager || !driveManager || !projectManager || !authManager || !apiHandler) {
            throw new Error("One or more core components failed to initialize.");
        }
        
        logInfo('Application successfully initialized.');
        return true;
    } catch (error) {
        logError(`Failed to initialize application: ${error.message}\n${error.stack}`);
        // Reset globals on failure to allow retry?
        sheetAccessor = templateManager = driveManager = projectManager = authManager = apiHandler = null; 
        // Re-throw to ensure the caller knows initialization failed
        throw new Error(`Initialization failed: ${error.message}`);
    }
}

/**
 * Shows a modal dialog with an error message in the Spreadsheet UI.
 * * @param {string} message - Error message to display.
 */
function showErrorAlert(message) {
  try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Error', message, ui.ButtonSet.OK);
  } catch (uiError) {
      logError(`Could not display UI alert: ${uiError.message}`);
      // Fallback logging
      Logger.log(`ALERT (UI unavailable): ${message}`);
  }
}

// ============================================================
// Spreadsheet UI Functions (Dialogs, etc.) - Keep as is or refine
// ============================================================

/**
 * Shows a dialog to create a new project.
 */
function showNewProjectDialog() {
  try {
    initialize(); // Ensure initialized
    const htmlOutput = HtmlService.createTemplateFromFile('NewProjectDialog') // Assuming HTML is in a file
      .evaluate()
      .setWidth(400)
      .setHeight(250);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Create New Project');
  } catch (error) {
    logError(`Failed to show new project dialog: ${error.message}`);
    showErrorAlert('Failed to open the new project dialog. Please ensure the "NewProjectDialog.html" file exists and try again.');
  }
}

/**
 * Server-side function called by the new project dialog's client-side script.
 * * @param {string} projectName - Name for the new project.
 * @return {Object} Result object { success: boolean, message: string, projectId?: string, ... }.
 */
function createNewProject(projectName) {
  try {
    initialize(); // Ensure initialized
    // Use the ProjectManager to create the project
    const result = projectManager.createProject(projectName);
    // If successful, activate the project tab for immediate feedback
    if (result.success && result.projectTabName) {
      try {
          sheetAccessor.getSheet(result.projectTabName, false)?.activate();
      } catch(activateError) {
           logWarning(`Could not activate sheet ${result.projectTabName}: ${activateError.message}`);
      }
    }
    return result; // Return the result object from ProjectManager
  } catch (error) {
    logError(`Failed to create new project via dialog function: ${error.message}\n${error.stack}`);
    return { success: false, message: `Server error: ${error.message}` };
  }
}

/**
 * Shows a dialog to select an existing project for editing.
 */
function showProjectSelector() {
  try {
    initialize(); // Ensure initialized
    // ProjectSelector.html should fetch its own list via API call
    const htmlOutput = HtmlService.createTemplateFromFile('ProjectSelector') // Assuming HTML is in a file
        .evaluate()
        .setWidth(500)
        .setHeight(450); // Adjust size as needed
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Project to Edit');
  } catch (error) {
    logError(`Failed to show project selector: ${error.message}\n${error.stack}`);
    showErrorAlert('Failed to open the project selector. Please ensure the "ProjectSelector.html" file exists and try again.');
  }
}

/**
 * Server-side function called by the project selector dialog to open the editor.
 * This function activates the sheet. The actual editor loading happens via doGet.
 * * @param {string} projectId - ID of the project to open.
 * @return {Object} Result object { success: boolean, message: string }.
 */
function openProjectForEditing(projectId) {
  try {
    initialize(); // Ensure initialized
    // Use the ProjectManager to activate the sheet and update timestamp
    const result = projectManager.openProject(projectId);
    // Note: This function just activates the sheet. The editor UI is loaded via doGet.
    // We might not even need this server function if the "Edit" button in the dialog
    // simply constructs the editor URL and closes the dialog.
    // However, activating the sheet provides feedback to the user in the spreadsheet.
    return result; 
  } catch (error) {
    logError(`Failed to open project for editing (server-side): ${error.message}\n${error.stack}`);
    return { success: false, message: `Server error opening project: ${error.message}` };
  }
}

/**
 * Shows the URL of the root project folder in Drive.
 */
function showRootFolderURL() {
  try {
    initialize(); // Ensure initialized
    const rootFolder = driveManager.getRootFolder(); // Assumes DriveManager handles finding/creating it
    if (!rootFolder) {
         showErrorAlert('Could not find or create the root project folder in Google Drive.');
         return;
    }
    const folderURL = rootFolder.getUrl();
    const folderName = rootFolder.getName();
    
    const ui = SpreadsheetApp.getUi();
    // Show URL in a simple alert first
    ui.alert('Root Folder Location', 
             `The root folder for training projects is:\n\nFolder: ${folderName}\nURL: ${folderURL}`,
             ui.ButtonSet.OK);
             
    // Optional: Try to open the folder (might be blocked by popup blockers)
    // const html = HtmlService.createHtmlOutput(`<script>window.open('${folderURL}', '_blank'); google.script.host.close();</script>`).setWidth(100).setHeight(50);
    // ui.showModalDialog(html, 'Opening Folder...');

  } catch (error) {
    logError(`Failed to show root folder URL: ${error.message}\n${error.stack}`);
    showErrorAlert('Failed to locate the root folder. Please check logs.');
  }
}

// Placeholder functions for menu items not yet implemented
function showProjectViewer() { SpreadsheetApp.getUi().alert('Coming Soon', 'Project viewer will be implemented in a future phase.', SpreadsheetApp.getUi().ButtonSet.OK); }
function showAddSlideDialog() { SpreadsheetApp.getUi().alert('Coming Soon', 'Add Slide dialog will be implemented later.', SpreadsheetApp.getUi().ButtonSet.OK); }
function showAddElementDialog() { SpreadsheetApp.getUi().alert('Coming Soon', 'Add Element dialog will be implemented later.', SpreadsheetApp.getUi().ButtonSet.OK); }
function showFileUploadDialog() { SpreadsheetApp.getUi().alert('Coming Soon', 'File Upload dialog will be implemented later.', SpreadsheetApp.getUi().ButtonSet.OK); }
function showDeployProjectDialog() { SpreadsheetApp.getUi().alert('Coming Soon', 'Deploy Project dialog will be implemented later.', SpreadsheetApp.getUi().ButtonSet.OK); }
function showFileManager() { SpreadsheetApp.getUi().alert('Coming Soon', 'File manager will be implemented in a future phase.', SpreadsheetApp.getUi().ButtonSet.OK); }
function showAnalyticsDashboard() { SpreadsheetApp.getUi().alert('Coming Soon', 'Analytics dashboard will be implemented in a future phase.', SpreadsheetApp.getUi().ButtonSet.OK); }


// ============================================================
// API Request Processing (Called by google.script.run)
// ============================================================

/**
 * Processes an API request from the web client (e.g., Editor, Viewer).
 * Acts as the single entry point for google.script.run calls from the client.
 * * @param {Object} requestData - Request data from client { action: string, ... }
 * @return {Object} Response object { success: boolean, data?: any, error?: string }
 */
function processApiRequest(requestData) {
  let response;
  const startTime = Date.now();
  try {
    logDebug(`Received processApiRequest: ${JSON.stringify(requestData).substring(0, 500)}...`); // Log start and snippet
    
    // Ensure application is initialized
    if (!apiHandler) {
      logInfo('API Handler not initialized in processApiRequest, calling initialize()...');
      initialize();
    }
    
    // Get the current user
    const user = Session.getEffectiveUser().getEmail();
    
    // Handle the request using ApiHandler
    // ApiHandler now returns a standard object { success: boolean, ... }
    const result = apiHandler.handleRequest(requestData, user); 
    
    // Prepare the final response object
    if (result && result.success) {
        response = {
            success: true,
            // Apply safe serialization to the data payload before sending back
            data: result.data !== undefined ? safeSerialize(result.data) : null 
        };
    } else {
         // If ApiHandler returned { success: false, error: ... }
         response = {
             success: false,
             error: result.error || 'An unknown error occurred in the API handler.'
         };
         if (result.requireAuth) { // Propagate auth hint
             response.requireAuth = true;
         }
    }

  } catch (error) {
    // Catch any unexpected errors during initialization or handling
    logError(`Critical error processing API request (${requestData ? requestData.action : 'unknown'}): ${error.message}\n${error.stack}`);
    response = {
      success: false,
      error: `Server error processing request: ${error.message}`
    };
  } 
  
  const duration = Date.now() - startTime;
  logDebug(`processApiRequest completed for action '${requestData ? requestData.action : 'unknown'}' in ${duration}ms. Success: ${response.success}`);
  
  // Return the standard response object - google.script.run handles JSON stringification
  return response; 
}


// ============================================================
// Direct-Callable Functions (Alternative to processApiRequest if needed)
// ============================================================

/**
 * Simple API method to get projects without complex objects.
 * Can be called directly from the client to avoid potential serialization issues
 * with the main processApiRequest if safeSerialize isn't perfect.
 * * @return {Object} { success: boolean, data: Array<{projectId, title, createdAt, modifiedAt}>, error?: string }
 */
function getProjectList() {
  try {
    initialize(); // Ensure initialized
    const user = Session.getEffectiveUser().getEmail();
    logDebug(`Getting direct project list for user: ${user}`);
    
    // Get all projects (basic info)
    const allProjects = projectManager.getAllProjects(); // This already returns basic info
    if (!Array.isArray(allProjects)) {
      throw new Error('Expected array of projects from ProjectManager');
    }
    
    // Filter based on access
    const accessibleProjects = allProjects.filter(project => {
        if (!project || typeof project !== 'object' || !project.projectId) return false;
        return authManager.hasAccess(project.projectId, user);
    });

    // Map to ensure only primitive types are included (safe for direct return)
    const simplifiedProjects = accessibleProjects.map(project => ({
        projectId: project.projectId,
        title: project.title,
        // Convert dates to timestamps (numbers)
        createdAt: project.createdAt instanceof Date ? project.createdAt.getTime() : (typeof project.createdAt === 'number' ? project.createdAt : null),
        modifiedAt: project.modifiedAt instanceof Date ? project.modifiedAt.getTime() : (typeof project.modifiedAt === 'number' ? project.modifiedAt : null)
    }));
    
    logDebug(`Returning ${simplifiedProjects.length} projects via getProjectList.`);
    return { success: true, data: simplifiedProjects };

  } catch (error) {
    logError(`Error in direct getProjectList: ${error.message}\n${error.stack}`);
    return { success: false, error: error.message || 'Unknown error in getProjectList', data: [] };
  }
}

// Removed testApiConnection and requestProjectAccess unless needed for specific UI flows.
// They can be added back if required.

// Server-side function in Code.gs
function getEditorUrl(projectId) {
  console.log("getEditorUrl called for project:", projectId);
  return ScriptApp.getService().getUrl() + "?action=edit&project=" + projectId;
}