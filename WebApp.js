/**
 * WebApp.js - Web request handlers for the Interactive Training Projects Web App
 * Handles doGet and doPost requests and serves HTML content
 */

/**
 * Handles GET requests to the web app
 * Routes requests based on parameters to the appropriate handler
 *
 * @param {Object} e - Event object from Apps Script
 * @return {HtmlOutput} HTML content based on the request
 */
function doGet(e) {
  try {
    // Initialize application if needed
    if (!sheetAccessor || !templateManager || !driveManager || !projectManager || !authManager || !apiHandler) {
      initialize();
    }

    // Get current user
    const userEmail = Session.getEffectiveUser().getEmail();
    console.log(`doGet request from user: ${userEmail}`);

    // Parse query parameters
    const params = e.parameter || {};
    console.log(`Request parameters: ${JSON.stringify(params)}`);

    // Project selector (default view when no parameters)
    if (!params.project && !params.action) {
      console.log("Serving project selector");
      return serveProjectSelector();
    }

    // If project ID is provided, get the project
    if (params.project) {
      const projectId = params.project;
      console.log(`Accessing project: ${projectId}`);

      try {
        // Get project data
        const project = projectManager.getProject(projectId);
        
        if (!project) {
          console.error(`Project not found: ${projectId}`);
          return serveErrorPage(`Project not found: ${projectId}`, 404);
        }

        // Check user access
        const accessLevel = authManager.getAccessLevel(projectId, userEmail);
        
        if (accessLevel === authManager.accessLevels.NONE) {
          console.warn(`Access denied for user ${userEmail} to project ${projectId}`);
          return serveLoginPage(projectId);
        }

        // Based on action parameter, serve appropriate view
        const action = params.action || 'view'; // Default action is view
        
        // Always use combined view but with different default tabs based on action
        console.log(`Serving combined view with action: ${action} for user with access level: ${accessLevel}`);
        return serveCombinedView(project, accessLevel, action);
      } catch (projectError) {
        console.error(`Error accessing project: ${projectError.message}\n${projectError.stack}`);
        logError(`Error accessing project: ${projectError.message}`);
        return serveErrorPage(`Error accessing project: ${projectError.message}`, 500);
      }
    }

    // If we get here, show the project selector
    console.log("No valid action or project ID specified, serving project selector");
    return serveProjectSelector();
  } catch (error) {
    console.error(`Error in doGet: ${error.message}\n${error.stack}`);
    logError(`Error in doGet: ${error.message}`);
    return serveErrorPage(`Server error: ${error.message}`, 500);
  }
}

/**
 * Serves a combined view with both viewer and editor components
 * The components are shown/hidden based on user permissions and requested action
 *
 * @param {Object} project - Project data
 * @param {string} accessLevel - User's access level (owner, editor, viewer)
 * @param {string} defaultAction - Default action/tab to show (view, edit, analytics)
 * @return {HtmlOutput} HTML for the combined view
 */
function serveCombinedView(project, accessLevel, defaultAction = 'view') {
  try {
    // Get the HTML template
    let template = HtmlService.createTemplateFromFile('CombinedView');

    // Add project data to the template
    template.project = project;

    // Add user info and permissions
    template.user = {
      email: Session.getEffectiveUser().getEmail(),
      accessLevel: accessLevel
    };

    // Check if user can edit
    template.canEdit = (accessLevel === authManager.accessLevels.OWNER || 
                         accessLevel === authManager.accessLevels.ADMIN ||
                         accessLevel === authManager.accessLevels.EDITOR);
    
    // Check if user is admin/owner
    template.isAdmin = (accessLevel === authManager.accessLevels.OWNER || 
                        accessLevel === authManager.accessLevels.ADMIN);

    // Set the default action/tab
    template.defaultAction = (defaultAction === 'edit' && template.canEdit) ? 'editor' : 
                            (defaultAction === 'analytics' && template.canEdit) ? 'analytics' : 
                            'viewer';

    // Get analytics data if user is admin/owner or editor and analytics is requested
    if (template.canEdit && defaultAction === 'analytics') {
      // TODO: Get analytics data if needed for initial load
      // This could be moved to client-side API call for better performance
    }

    // Process the template
    let htmlOutput = template.evaluate()
      .setTitle(`${project.title} - Interactive Training`)
      .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png');

    // Set appropriate options for the web app
    htmlOutput = htmlOutput
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return htmlOutput;
  } catch (error) {
    console.error(`Error serving combined view: ${error.message}\n${error.stack}`);
    logError(`Error serving combined view: ${error.message}`);
    return serveErrorPage(`Failed to load training: ${error.message}`, 500);
  }
}

/**
 * Handles POST requests to the web app (primarily for API calls)
 * @param {Object} e - The event object from Apps Script
 * @returns {TextOutput} - JSON response
 */
function doPost(e) {
  try {
    // Parse the request data
    let requestData;
    try {
      requestData = JSON.parse(e.postData.contents);
    } catch (parseError) {
      return createJsonResponse({
        success: false,
        error: "Invalid JSON in request: " + parseError.toString()
      }, 400);
    }
    
    // Get user info
    const userEmail = Session.getEffectiveUser().getEmail();
    const user = { email: userEmail };
    
    // Log incoming request for debugging
    Logger.log("API Request: " + JSON.stringify(requestData));
    
    // Process the API request
    if (!requestData.action) {
      return createJsonResponse({
        success: false,
        error: "No action specified in request"
      }, 400);
    }
    
    // Initialize required components
    const sheetAccessor = new SheetAccessor();
    const templateManager = new TemplateManager(sheetAccessor);
    const driveManager = new DriveManager();
    const projectManager = new ProjectManager(sheetAccessor, templateManager, driveManager);
    const authManager = new AuthManager(sheetAccessor, driveManager);
    const fileUploader = new FileUploader(driveManager, new MediaProcessor(driveManager));
    const apiHandler = new ApiHandler(projectManager, fileUploader, authManager);
    
    // Process the request and get the response
    const response = apiHandler.handleRequest(requestData, user);
    
    // Log the response for debugging
    Logger.log("API Response: " + JSON.stringify(response));
    
    // Return the JSON response
    return createJsonResponse(response);
  } catch (error) {
    // Handle unexpected errors
    Logger.log("Error in doPost: " + error.toString());
    return createJsonResponse({
      success: false,
      error: "Server error: " + error.toString()
    }, 500);
  }
}

/**
 * Creates a JSON response object
 * @param {Object} data - The data to return
 * @param {number} statusCode - HTTP status code (optional, defaults to 200)
 * @returns {TextOutput} - JSON response
 */
function createJsonResponse(data, statusCode = 200) {
  const response = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return response;
}

/**
 * Creates a JSON response with the appropriate content type
 *
 * @param {Object} data - Response data object
 * @param {number} statusCode - HTTP status code (default: 200)
 * @return {TextOutput} Text output with JSON content type
 */
function createJsonResponse(data, statusCode = 200) {
  // Ensure data is an object
  if (typeof data !== 'object' || data === null) {
      data = { success: false, error: 'Invalid response data type from server', dataValue: data };
  }
  // Ensure status is part of the object
  if (!data.hasOwnProperty('status')) {
      data.status = statusCode;
  }

  const response = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);

  return response;
}

/**
 * DEPRECATED: This function is no longer used as we've migrated to CombinedView
 * Kept for reference purposes only.
 */
function serveViewerApp(project) {
  console.warn("serveViewerApp() is deprecated - use serveCombinedView() instead");
  return serveCombinedView(project, "viewer", "viewer");
}

/**
 * DEPRECATED: This function is no longer used as we've migrated to CombinedView
 * Kept for reference purposes only.
 */
function serveEditorView(project) {
  console.warn("serveEditorView() is deprecated - use serveCombinedView() instead");
  return serveCombinedView(project, "editor", "editor");
}

/**
 * Serves the analytics dashboard for a project
 *
 * @param {Object} project - Project data
 * @return {HtmlOutput} HTML for the analytics dashboard
 */
function serveAnalyticsView(project) {
  try {
    // Get the HTML template
    let template = HtmlService.createTemplateFromFile('AnalyticsDashboard');

    // Add project data to the template
    template.project = project;

    // Add user info
    template.user = {
      email: Session.getEffectiveUser().getEmail()
    };

    // Process the template
    let htmlOutput = template.evaluate()
      .setTitle(`Analytics: ${project.title} - Interactive Training`)
      .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png');

    // Set appropriate options for the web app
    htmlOutput = htmlOutput
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return htmlOutput;
  } catch (error) {
    console.error(`Error serving analytics view: ${error.message}\n${error.stack}`);
    logError(`Error serving analytics view: ${error.message}`);
    return serveErrorPage(`Failed to load analytics: ${error.message}`, 500);
  }
}

/**
 * Serves the project selector page
 *
 * @return {HtmlOutput} HTML for the project selector
 */
function serveProjectSelector() {
  try {
    // Get the HTML template
    let template = HtmlService.createTemplateFromFile('ProjectSelector');

    // Add user info
    template.user = {
      email: Session.getEffectiveUser().getEmail()
    };

    // Process the template
    let htmlOutput = template.evaluate()
      .setTitle('Select Project - Interactive Training')
      .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png');

    // Set appropriate options for the web app
    htmlOutput = htmlOutput
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return htmlOutput;
  } catch (error) {
    console.error(`Error serving project selector: ${error.message}\n${error.stack}`);
    logError(`Error serving project selector: ${error.message}`);
    return serveErrorPage(`Failed to load project selector: ${error.message}`, 500);
  }
}

/**
 * Serves a login page when authentication is required
 *
 * @param {string} projectId - ID of the project requiring authentication
 * @return {HtmlOutput} HTML for the login page
 */
function serveLoginPage(projectId) {
  try {
    // Get the HTML template
    let template = HtmlService.createTemplateFromFile('Login');

    // Add data to the template
    template.projectId = projectId;
    template.returnUrl = ScriptApp.getService().getUrl() + '?project=' + projectId;
    template.user = {
      email: Session.getEffectiveUser().getEmail()
    };

    // Process the template
    let htmlOutput = template.evaluate()
      .setTitle('Login Required - Interactive Training')
      .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png');

    // Set appropriate options for the web app
    htmlOutput = htmlOutput
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return htmlOutput;
  } catch (error) {
    console.error(`Error serving login page: ${error.message}\n${error.stack}`);
    logError(`Error serving login page: ${error.message}`);
    return serveErrorPage(`Failed to load login page: ${error.message}`, 500);
  }
}

/**
 * Serves an error page with the specified message and status code
 *
 * @param {string} message - Error message to display
 * @param {number} statusCode - HTTP status code
 * @return {HtmlOutput} HTML for the error page
 */
function serveErrorPage(message, statusCode) {
  try {
    // Get the HTML template
    let template = HtmlService.createTemplateFromFile('Error');

    // Add error data to the template
    template.message = message;
    template.statusCode = statusCode;

    // Process the template
    let htmlOutput = template.evaluate()
      .setTitle('Error - Interactive Training')
      .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png');

    // Set appropriate options for the web app
    htmlOutput = htmlOutput
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return htmlOutput;
  } catch (error) {
    // Last resort error handling if template fails
    console.error(`FATAL: Error serving error page itself! Original: ${message}, Template Error: ${error.message}\n${error.stack}`);
    return HtmlService.createHtmlOutput(`
      <h1>Error</h1>
      <p>Original error: ${message}</p>
      <p>Additionally, failed to load error template: ${error.message}</p>
    `);
  }
}

/**
 * Includes an HTML file in another HTML file
 * This function is called from HTML templates
 *
 * @param {string} filename - Name of the HTML file to include
 * @return {string} Contents of the HTML file
 */
function include(filename) {
  // Simple check to prevent including server-side .js files directly if they exist
  if (filename.endsWith('.js')) {
      console.warn(`Attempted to include a JS file directly via include(): ${filename}. Use <script src="..."> or include HTML wrappers.`);
      return `<!-- Error: Cannot include JS file directly: ${filename} -->`;
  }
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Serves a static file as an HtmlOutput
 * Useful for CSS, JavaScript, and other static files
 *
 * @param {string} filename - Name of the file to serve
 * @return {HtmlOutput} HtmlOutput with the file content
 */
function serveStaticFile(filename) {
  // Added security check: Only allow specific file types or names if needed
  if (!filename.endsWith('.css') && !filename.endsWith('.js') && !filename.endsWith('.html')) {
      console.error(`Attempt to serve potentially unsafe file type: ${filename}`);
      return serveErrorPage("Invalid file request", 400);
  }
  return HtmlService.createHtmlOutputFromFile(filename);
}

/**
 * Gets a list of projects for the current user
 * This function is called directly from the ProjectSelector HTML
 * 
 * @return {Array} Array of project objects with permission info
 */
function getProjectList() {
  try {
    // Initialize application if needed
    if (!sheetAccessor || !templateManager || !driveManager || !projectManager || !authManager || !apiHandler) {
      console.log("API Handler not initialized in getProjectList, calling initialize()...");
      initialize();
    }
    
    // Get current user email
    const userEmail = Session.getEffectiveUser().getEmail();
    console.log(`Getting project list for user: ${userEmail}`);
    
    // Get all projects (basic info)
    const allProjects = projectManager.getAllProjects(false); // Don't need full details
    console.log(`Found ${allProjects.length} total projects`);
    
    // Add access level information for each project
    const projectsWithPermissions = allProjects.map(project => {
      // Get access level for current user
      const accessLevel = authManager.getAccessLevel(project.projectId, userEmail);
      console.log(`Project ${project.title}, Access level: ${accessLevel}`);
      
      // Check if user is a project admin
      const isAdmin = authManager.isProjectAdmin(project.projectId, userEmail);
      console.log(`Project ${project.title}, Is admin: ${isAdmin}`);
      
      // Add permission info to project
      return {
        projectId: project.projectId,
        title: project.title,
        createdAt: project.createdAt,
        modifiedAt: project.modifiedAt, 
        accessLevel: accessLevel,
        isAdmin: isAdmin
      };
    });
    
    console.log(`Returning ${projectsWithPermissions.length} projects with permissions`);
    
    // IMPORTANT: Return the array directly, not wrapped in an object
    return projectsWithPermissions;
  } catch (error) {
    console.error(`Error getting project list: ${error.message}`);
    logError(`Error getting project list: ${error.message}`);
    
    // Return empty array on error
    return [];
  }
}