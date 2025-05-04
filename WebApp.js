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
            return serveEditorView(project);
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
      
      // Check authentication
      const user = Session.getEffectiveUser().getEmail();
      if (!authManager.hasAccess(requestData.projectId, user)) {
        return createJsonResponse({ 
          success: false, 
          message: 'Access denied', 
          requireAuth: true 
        }, 403);
      }
      
      // Process the API request using ApiHandler
      return apiHandler.handleRequest(requestData, user);
    } catch (error) {
      logError(`Error in doPost: ${error.message}\n${error.stack}`);
      return createJsonResponse({ 
        success: false, 
        message: 'Server error', 
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
    const response = ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
    
    // Note: Apps Script doesn't support setting HTTP status codes directly
    // We include the status in the response data
    data.status = statusCode;
    
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
      // Get the HTML template
      let template = HtmlService.createTemplateFromFile('EditorView');
      
      // Add project data to the template
      template.project = project;
      
      // Add user info
      template.user = {
        email: Session.getEffectiveUser().getEmail()
      };
      
      // Process the template
      let htmlOutput = template.evaluate()
        .setTitle(`Edit: ${project.title} - Interactive Training`)
        .setFaviconUrl('https://www.google.com/images/icons/product/drive-32.png');
      
      // Set appropriate options for the web app
      htmlOutput = htmlOutput
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      return htmlOutput;
    } catch (error) {
      logError(`Error serving editor view: ${error.message}`);
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
      // Get all projects
      const projects = projectManager.getAllProjects();
      
      // Get the HTML template
      let template = HtmlService.createTemplateFromFile('ProjectSelector');
      
      // Add projects data to the template
      template.projects = projects;
      
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
    return HtmlService.createHtmlOutputFromFile(filename);
  }