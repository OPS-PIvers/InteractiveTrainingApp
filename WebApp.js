/**
 * WebApp.js - Web request handlers for the Interactive Training Projects Web App
 * Handles doGet and doPost requests and serves HTML content
 */

/**
 * Handles GET requests to the web app
 * Routes to different views based on parameters
 *
 * @param {Object} e - Event object from Apps Script
 * @return {HtmlOutput} HTML content to display
 */
function doGet(e) {
  try {
    // Initialize application if needed
  if (!sheetAccessor || !templateManager || !driveManager || !projectManager || !authManager || !apiHandler) {
      initialize();
    }

    // Get parameters
    const params = e.parameter || {};
    const mode = params.mode || 'viewer'; // Default to viewer mode
    const projectId = params.project || ''; // Project ID parameter

    // Check if a valid project ID is provided
    if (projectId) {
      // Get the project
      const project = projectManager.getProject(projectId);

      if (!project) {
        return serveErrorPage('Project not found', 404);
      }

      // Check authentication if needed
      if (!authManager.hasAccess(projectId, Session.getEffectiveUser().getEmail())) {
        return serveLoginPage(projectId);
      }

      // Serve the appropriate view based on mode
      switch(mode) {
        case 'editor':
          return serveEditorView(project); // <--- Calls the updated function
        case 'analytics':
          return serveAnalyticsView(project);
        case 'viewer':
        default:
          return serveViewerApp(project);
      }
    } else {
      // No project ID provided, show project selector
      return serveProjectSelector();
    }
  } catch (error) {
    console.error(`Error in doGet: ${error.message}\n${error.stack}`); // Log error
    logError(`Error in doGet: ${error.message}\n${error.stack}`);
    return serveErrorPage(`An error occurred: ${error.message}`, 500);
  }
}

/**
 * Handles POST requests to the web app
 * Processes API calls from the client-side code
 *
 * @param {Object} e - Event object from Apps Script
 * @return {ContentService|TextOutput} Response data in appropriate format
 */
function doPost(e) {
  try {
    // Initialize application if needed
    if (!sheetAccessor || !templateManager || !driveManager || !projectManager || !authManager || !apiHandler) {
      initialize();
    }

    // Parse the request data
    let requestData;
    try {
      requestData = JSON.parse(e.postData.contents);
    } catch (parseError) {
      return createJsonResponse({
        success: false,
        message: 'Invalid JSON data',
        error: parseError.message
      }, 400);
    }

    // *** IMPORTANT: Authentication Check Logic Moved to ApiHandler ***
    // It's better to check auth within the handler based on the specific action

    // Process the API request using ApiHandler
    // Make sure processApiRequest is exposed or handle it differently
    // Assuming processApiRequest is the intended global function to call ApiHandler
    const response = processApiRequest(requestData); // Call the global bridge function

    // processApiRequest should already return the correct structure
    // If it returns ContentService, return it directly
    if (response && typeof response.getContent === 'function') {
        return response;
    }
    // Otherwise, wrap it in a JSON response
    return createJsonResponse(response);

  } catch (error) {
    console.error(`Error in doPost: ${error.message}\n${error.stack}`); // Log error
    logError(`Error in doPost: ${error.message}\n${error.stack}`);
    return createJsonResponse({
      success: false,
      message: 'Server error during POST handling',
      error: error.message
    }, 500);
  }
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
 * Serves the viewer application for a project
 *
 * @param {Object} project - Project data
 * @return {HtmlOutput} HTML for the viewer application
 */
function serveViewerApp(project) {
  try {
    // Get the HTML template
    let template = HtmlService.createTemplateFromFile('ViewerApp');

    // Add project data to the template
    template.project = project;

    // Add user info
    template.user = {
      email: Session.getEffectiveUser().getEmail()
    };

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
    console.error(`Error serving viewer app: ${error.message}\n${error.stack}`); // Log error
    logError(`Error serving viewer app: ${error.message}`);
    return serveErrorPage(`Failed to load viewer: ${error.message}`, 500);
  }
}

/**
 * Serves the editor application for a project
 *
 * @param {Object} project - Project data
 * @return {HtmlOutput} HTML for the editor application
 */
function serveEditorView(project) {
  try {
    // === ADDED: Logging ===
    console.log(`Serving EditorView for project: ${project.projectId} - ${project.title}`);
    try {
      // Attempt to stringify to check for issues early and log sample data
      const projectSample = JSON.stringify(project).substring(0, 500);
      console.log(`Project data sample: ${projectSample}...`);
    } catch (stringifyError) {
      console.error(`Failed to stringify project data before templating: ${stringifyError}`);
      // Consider returning an error page here if stringify fails critically
      return serveErrorPage(`Failed to prepare project data: ${stringifyError.message}`, 500);
    }
    // === END Logging ===

    // Get the HTML template
    let template = HtmlService.createTemplateFromFile('EditorView');

    // Add project data to the template
    template.project = project; // This makes 'project' available in EditorView.html's scriptlets

    // Add user info
    template.user = {
      email: Session.getEffectiveUser().getEmail()
    };

    // Process the template
    let htmlOutput = template.evaluate()
      .setTitle(`Edit: ${project.title} - Interactive Training`)
      .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png');

     // === ADDED: Logging ===
     console.log("Template evaluated successfully for EditorView.");
     // === END Logging ===

    // Set appropriate options for the web app
    htmlOutput = htmlOutput
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return htmlOutput;
  } catch (error) {
    // === UPDATED: Enhanced Logging ===
    console.error(`Error serving editor view: ${error.message}\n${error.stack}`);
    logError(`Error serving editor view: ${error.message}`);
    // === END UPDATE ===
    return serveErrorPage(`Failed to load editor: ${error.message}`, 500);
  }
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
    console.error(`Error serving analytics view: ${error.message}\n${error.stack}`); // Log error
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
    // *** IMPORTANT: Project listing logic might be better handled client-side via API call ***
    // For now, keeping server-side rendering as it was
    // Get all projects (this might fetch more data than needed for just the selector)
    // Consider having a lighter `projectManager.getProjectListSummary()` if performance becomes an issue
    // const projects = projectManager.getAllProjects(); // Assuming this works

    // Get the HTML template
    let template = HtmlService.createTemplateFromFile('ProjectSelector');

    // *** REMOVED: Passing projects directly to template ***
    // template.projects = projects; // Data will now be loaded client-side

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
    console.error(`Error serving project selector: ${error.message}\n${error.stack}`); // Log error
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
    console.error(`Error serving login page: ${error.message}\n${error.stack}`); // Log error
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
    console.error(`FATAL: Error serving error page itself! Original: ${message}, Template Error: ${error.message}\n${error.stack}`); // Log error
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