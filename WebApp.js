/**
 * WebApp.js - Web request handlers for the Interactive Training Projects Web App
 * Handles doGet and doPost requests and serves HTML content
 */

/**
 * Handles GET requests to the web app
 * Routes to different views based on parameters
 * * @param {Object} e - Event object from Apps Script { parameter: { project: '...', mode: '...' } }
 * @return {HtmlOutput} HTML content to display
 */
function doGet(e) {
    try {
      logInfo(`doGet received request: ${JSON.stringify(e.parameter)}`);
      // Initialize application if needed (ensure all managers are available)
      if (!sheetAccessor || !templateManager || !driveManager || !projectManager || !authManager || !apiHandler) {
        logInfo("Managers not initialized, calling initialize()...");
        initialize();
      }
      
      // Get parameters
      const params = e.parameter || {};
      const mode = params.mode || 'viewer'; // Default to viewer mode
      const projectId = params.project || ''; // Project ID parameter
      
      const userEmail = Session.getEffectiveUser().getEmail();
      logInfo(`User: ${userEmail}, Mode: ${mode}, ProjectID: ${projectId}`);

      // Check if a valid project ID is provided
      if (projectId) {
        // Get the project data (should include slides and elements)
        const project = projectManager.getProject(projectId);
        
        if (!project) {
          logWarning(`Project not found for ID: ${projectId}`);
          return serveErrorPage('Project not found', 404);
        }
        if (project.error) {
             logError(`Error fetching project ${projectId}: ${project.error}`);
             return serveErrorPage(`Error loading project: ${project.error}`, 500);
        }
        
        // Check authentication/authorization
        if (!authManager.hasAccess(projectId, userEmail)) {
          logWarning(`Access denied for user ${userEmail} to project ${projectId}`);
          // Consider serving a specific "Access Denied" page instead of login if user is already logged in
          return serveLoginPage(projectId); // Or serveAccessDeniedPage(project);
        }
        
        // Serve the appropriate view based on mode
        logInfo(`Routing to mode: ${mode}`);
        switch(mode) {
          case 'editor':
            return serveEditorView(project, userEmail); // Pass user email
          case 'analytics':
            // Ensure user has permission to view analytics (e.g., owner/editor)
            if (authManager.getAccessLevel(projectId, userEmail) === 'viewer') {
                 logWarning(`User ${userEmail} does not have permission to view analytics for ${projectId}`);
                 return serveErrorPage('Permission denied to view analytics.', 403);
            }
            return serveAnalyticsView(project, userEmail); // Pass user email
          case 'viewer':
          default:
            return serveViewerApp(project, userEmail); // Pass user email
        }
      } else {
        // No project ID provided, show project selector (which handles its own auth check internally)
        logInfo("No project ID provided, serving project selector.");
        return serveProjectSelector(userEmail); // Pass user email
      }
    } catch (error) {
      logError(`Critical error in doGet: ${error.message}\n${error.stack}`);
      return serveErrorPage(`An unexpected server error occurred. Please try again later.`, 500);
    }
  }
  
  /**
   * Handles POST requests to the web app (primarily for API calls)
   * * @param {Object} e - Event object from Apps Script { postData: { contents: '...' } }
   * @return {ContentService.TextOutput} JSON response
   */
  function doPost(e) {
    let response;
    try {
      logInfo("doPost received request.");
      // Initialize application if needed
      if (!apiHandler) { // Only need ApiHandler for POST
        logInfo("API Handler not initialized, calling initialize()...");
        initialize();
      }
      
      // Parse the request data
      let requestData;
      if (!e || !e.postData || !e.postData.contents) {
          throw new Error("Invalid POST request: Missing post data.");
      }
      try {
        requestData = JSON.parse(e.postData.contents);
        logDebug(`doPost parsed request data: Action=${requestData.action}`);
      } catch (parseError) {
         logError(`Invalid JSON in POST request: ${e.postData.contents}`);
         throw new Error(`Invalid request format: ${parseError.message}`);
      }
      
      // Get user (already authenticated by Google)
      const user = Session.getEffectiveUser().getEmail();
      
      // Process the API request using ApiHandler
      // ApiHandler.handleRequest now returns a standard JS object { success: boolean, ... }
      const result = apiHandler.handleRequest(requestData, user); 
      
      // Create the final JSON response for the client
      response = createJsonResponse(result, result.success ? 200 : (result.requireAuth ? 403 : 400)); // Use status hint if available

    } catch (error) {
      logError(`Critical error in doPost: ${error.message}\n${error.stack}`);
      // Create a JSON error response if something went wrong before/after ApiHandler
      response = createJsonResponse({ 
        success: false, 
        error: `Server error: ${error.message}` 
      }, 500);
    }
    return response; // Return the ContentService TextOutput
  }
  
  /**
   * Creates a JSON response object suitable for ContentService output.
   * * @param {Object} data - Response data object { success: boolean, data?: any, error?: string, ... }
   * @param {number} statusCode - Hint for HTTP status code (used internally, not set on response)
   * @return {ContentService.TextOutput} Text output with JSON content type
   */
  function createJsonResponse(data, statusCode = 200) {
    // Ensure the response object includes a status hint
    data.status = statusCode; 
    const jsonString = JSON.stringify(data);
    logDebug(`Sending JSON response (Status hint ${statusCode}): ${jsonString.substring(0, 200)}...`); // Log snippet
    return ContentService.createTextOutput(jsonString)
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  /**
   * Serves the viewer application for a project
   * * @param {Object} project - Full project data object
   * @param {string} userEmail - Email of the current user
   * @return {HtmlOutput} HTML for the viewer application
   */
  function serveViewerApp(project, userEmail) {
    try {
      logInfo(`Serving viewer for project: ${project.projectId}`);
      let template = HtmlService.createTemplateFromFile('ViewerApp');
      
      // Serialize project data to pass safely to the template
      template.projectJson = JSON.stringify(safeSerialize(project)); // Use safeSerialize
      template.userEmail = userEmail; // Pass user email separately
      
      let htmlOutput = template.evaluate()
        .setTitle(`${project.title} - Interactive Training`)
        .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      return htmlOutput;
    } catch (error) {
      logError(`Error serving viewer app for project ${project.projectId}: ${error.message}\n${error.stack}`);
      return serveErrorPage(`Failed to load viewer: ${error.message}`, 500);
    }
  }
  
  /**
   * Serves the editor application for a project
   * * @param {Object} project - Full project data object
   * @param {string} userEmail - Email of the current user
   * @return {HtmlOutput} HTML for the editor application
   */
  function serveEditorView(project, userEmail) {
    try {
      logInfo(`Serving editor for project: ${project.projectId}`);
      let template = HtmlService.createTemplateFromFile('EditorView');
      
      // Serialize project data to pass safely to the template
      // Use safeSerialize from Utils.gs to handle potential Date objects etc.
      template.projectJson = JSON.stringify(safeSerialize(project)); 
      template.userEmail = userEmail; // Pass user email separately
      
      let htmlOutput = template.evaluate()
        .setTitle(`Edit: ${project.title} - Interactive Training`)
        .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png') // Consider a different icon for editor?
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Usually needed for embedding/interaction
      
      return htmlOutput;
    } catch (error) {
      logError(`Error serving editor view for project ${project.projectId}: ${error.message}\n${error.stack}`);
      return serveErrorPage(`Failed to load editor: ${error.message}`, 500);
    }
  }
  
  /**
   * Serves the analytics dashboard for a project
   * * @param {Object} project - Project data object (basic info might suffice)
   * @param {string} userEmail - Email of the current user
   * @return {HtmlOutput} HTML for the analytics dashboard
   */
  function serveAnalyticsView(project, userEmail) {
    try {
      logInfo(`Serving analytics for project: ${project.projectId}`);
      let template = HtmlService.createTemplateFromFile('AnalyticsDashboard');
      
      // Pass only necessary data (maybe just projectId and title)
      template.projectId = project.projectId;
      template.projectTitle = project.title; 
      template.userEmail = userEmail;
      
      let htmlOutput = template.evaluate()
        .setTitle(`Analytics: ${project.title} - Interactive Training`)
        .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      return htmlOutput;
    } catch (error) {
      logError(`Error serving analytics view for project ${project.projectId}: ${error.message}\n${error.stack}`);
      return serveErrorPage(`Failed to load analytics: ${error.message}`, 500);
    }
  }
  
  /**
   * Serves the project selector page
   * * @param {string} userEmail - Email of the current user
   * @return {HtmlOutput} HTML for the project selector
   */
  function serveProjectSelector(userEmail) {
    try {
      logInfo(`Serving project selector for user: ${userEmail}`);
      // Get projects the user can access (handled by listProjects in ApiHandler/ProjectManager)
      // The ProjectSelector.html itself will call the API to list projects.
      
      let template = HtmlService.createTemplateFromFile('ProjectSelector');
      template.userEmail = userEmail; // Pass user email
      
      let htmlOutput = template.evaluate()
        .setTitle('Select Project - Interactive Training')
        .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      return htmlOutput;
    } catch (error) {
      logError(`Error serving project selector: ${error.message}\n${error.stack}`);
      return serveErrorPage(`Failed to load project selector: ${error.message}`, 500);
    }
  }
  
  /**
   * Serves a login page when authentication/authorization fails for a specific project.
   * * @param {string} projectId - ID of the project requiring authentication/access
   * @return {HtmlOutput} HTML for the login/access denied page
   */
  function serveLoginPage(projectId) {
    try {
      logInfo(`Serving login/access denied page for project: ${projectId}`);
      // Consider renaming Login.html to AccessDenied.html if it serves that purpose better
      let template = HtmlService.createTemplateFromFile('Login'); // Assuming Login.html exists
      
      template.projectId = projectId;
      // Provide info needed for requesting access if applicable
      // template.projectTitle = projectManager.getProject(projectId)?.title || 'Unknown Project'; // Fetch title if needed
      template.userEmail = Session.getEffectiveUser().getEmail();
      
      let htmlOutput = template.evaluate()
        .setTitle('Access Required - Interactive Training')
        .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      return htmlOutput;
    } catch (error) {
      logError(`Error serving login page for project ${projectId}: ${error.message}\n${error.stack}`);
      return serveErrorPage(`Failed to load login page: ${error.message}`, 500);
    }
  }
  
  /**
   * Serves a generic error page.
   * * @param {string} message - Error message to display
   * @param {number} statusCode - Hint for HTTP status code
   * @return {HtmlOutput} HTML for the error page
   */
  function serveErrorPage(message, statusCode) {
    try {
      logError(`Serving error page: Status=${statusCode}, Message=${message}`);
      let template = HtmlService.createTemplateFromFile('Error'); // Assuming Error.html exists
      
      template.message = message;
      template.statusCode = statusCode;
      
      let htmlOutput = template.evaluate()
        .setTitle('Error - Interactive Training')
        .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      return htmlOutput;
    } catch (templateError) {
      // Last resort if Error.html itself fails
      logError(`Failed to load Error.html template: ${templateError.message}\n${templateError.stack}`);
      return HtmlService.createHtmlOutput(
        `<!DOCTYPE html><html><head><title>Error</title></head><body>` +
        `<h1>Application Error</h1><p>An error occurred while trying to load the application.</p>` +
        `<p><strong>Original Error:</strong> ${escapeHtml(message)} (Status: ${statusCode})</p>` +
        `<p><strong>Secondary Error:</strong> Failed to load error page template (${escapeHtml(templateError.message)})</p>` +
        `</body></html>`
      );
    }
  }

  /** Basic HTML escaping */
  function escapeHtml(unsafe) {
      if (!unsafe) return '';
      return unsafe
           .replace(/&/g, "&amp;")
           .replace(/</g, "&lt;")
           .replace(/>/g, "&gt;")
           .replace(/"/g, "&quot;")
           .replace(/'/g, "&#039;");
   }
  
  /**
   * Includes an HTML file content within another HTML template.
   * Used via <?!= include('filename'); ?> in HTML templates.
   * * @param {string} filename - Name of the HTML file to include (without .html extension)
   * @return {string} Contents of the HTML file
   */
  function include(filename) {
    // Check if filename already has .html, if not, add it.
    const actualFilename = filename.endsWith('.html') ? filename : filename + '.html';
    try {
        return HtmlService.createHtmlOutputFromFile(actualFilename).getContent();
    } catch (error) {
         logError(`Failed to include file '${actualFilename}': ${error.message}`);
         return `<p style="color:red; font-weight:bold;">Error including file: ${actualFilename}</p>`;
    }
  }
  
  // Note: serveStaticFile is generally not needed in Apps Script web apps
  // as CSS/JS are usually embedded or included directly via `include()`.
