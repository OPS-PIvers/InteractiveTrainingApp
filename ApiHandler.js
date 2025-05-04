/**
 * ApiHandler.js - API endpoint handler for the Interactive Training Projects Web App
 * Processes API requests from client-side code
 */

/**
 * ApiHandler class handles API requests and routes them to the appropriate functions
 */
class ApiHandler {
    /**
     * Creates a new ApiHandler instance
     * * @param {ProjectManager} projectManager - ProjectManager instance
     * @param {FileUploader} fileUploader - FileUploader instance
     * @param {AuthManager} authManager - AuthManager instance
     */
    constructor(projectManager, fileUploader, authManager) {
      this.projectManager = projectManager;
      this.fileUploader = fileUploader; // Assuming FileUploader is defined elsewhere
      this.authManager = authManager;
      
      // Define available API endpoints and their handlers
      this.endpoints = {
        // Project endpoints
        'project.get': this.getProject.bind(this),
        'project.list': this.listProjects.bind(this),
        'project.create': this.createProject.bind(this),
        'project.update': this.updateProject.bind(this),
        'project.delete': this.deleteProject.bind(this),
        
        // Slide endpoints
        'slide.add': this.addSlide.bind(this),
        'slide.update': this.updateSlide.bind(this),
        'slide.delete': this.deleteSlide.bind(this),
        
        // Element endpoints
        'element.add': this.addElement.bind(this),
        'element.update': this.updateElement.bind(this),
        'element.delete': this.deleteElement.bind(this),
        
        // Media endpoints
        'media.process': this.processMediaUrl.bind(this),
        
        // User tracking endpoints (Placeholders for now)
        'tracking.saveProgress': this.saveUserProgress.bind(this),
        'tracking.getProgress': this.getUserProgress.bind(this),
        
        // Analytics endpoints (Placeholders for now)
        'analytics.getProjectData': this.getAnalyticsData.bind(this),
        
        // Auth endpoints
        'auth.login': this.login.bind(this),
        'auth.logout': this.logout.bind(this),
        'auth.getStatus': this.getAuthStatus.bind(this)
      };
    }
    
    /**
     * Handles an API request. Called by the main web app handler (e.g., processApiRequest in Code.gs).
     * * @param {Object} requestData - Request data from client { action: '...', projectId: '...', ... }
     * @param {string} user - Email of the user making the request
     * @return {Object} Response object { success: boolean, data: ..., error: ... } - This is NOT a ContentService output.
     */
    handleRequest(requestData, user) {
        try {
          logDebug(`Handling API request: Action=${requestData.action}, ProjectID=${requestData.projectId}, User=${user}`);
          
          // Validate request format
          if (!requestData || typeof requestData !== 'object' || !requestData.action) {
            throw new Error('Invalid request format: Missing action.');
          }
          
          // Find the endpoint handler
          const handler = this.endpoints[requestData.action];
          
          if (!handler) {
            throw new Error(`Unknown action: ${requestData.action}`);
          }
          
          // Check authorization for this endpoint (basic check, AuthManager might do more)
          // Allow listing projects and getting auth status without explicit project access check here
           if (requestData.action !== 'project.list' && requestData.action !== 'auth.getStatus' && requestData.action !== 'project.create') {
               if (!requestData.projectId) {
                   throw new Error(`Missing required parameter: projectId for action ${requestData.action}`);
               }
               if (!this.authManager.hasAccess(requestData.projectId, user)) {
                  logWarning(`Authorization failed: User ${user} cannot access project ${requestData.projectId} for action ${requestData.action}`);
                  // Return specific auth error structure if needed by client
                  return {
                      success: false,
                      error: 'Authorization denied.',
                      requireAuth: true // Flag for client to potentially handle login/request access
                  };
               }
           }
          
          // Call the handler with the request data and user
          const result = handler(requestData, user);
          
          // Ensure the result from the handler is structured correctly (e.g., ProjectManager methods return {success: boolean, ...})
          if (result && typeof result === 'object' && result.success !== undefined) {
              if (result.success) {
                  logDebug(`API Action ${requestData.action} successful.`);
                  // Return only the relevant data part if the handler included extra info
                  return {
                      success: true,
                      // If the result has a specific data payload (like 'project', 'slide', 'element'), return that.
                      // Otherwise, return the whole result object excluding the 'success' flag.
                      data: result.project || result.slide || result.element || (({ success, ...data }) => data)(result) 
                  };
              } else {
                  // Handler indicated failure
                  logError(`API Action ${requestData.action} failed: ${result.message || 'No message provided.'}`);
                  throw new Error(result.message || `Action ${requestData.action} failed.`);
              }
          } else {
               // If the handler returned raw data (like getProject), wrap it
               logDebug(`API Action ${requestData.action} successful (raw data returned).`);
               return {
                  success: true,
                  data: result 
               };
          }

        } catch (error) {
          logError(`API error during action '${requestData ? requestData.action : 'unknown'}': ${error.message}\n${error.stack}`);
          // Return a standard error structure
          return {
            success: false,
            error: `Server error: ${error.message}` 
            // Optionally include error.stack in debug mode
          };
        }
    }
    
    //=========================================================================
    // PROJECT ENDPOINTS
    //=========================================================================
    
    /**
     * Gets a project by ID, ensuring full data (slides/elements) is included.
     * * @param {Object} requestData - { projectId: string }
     * @return {Object} Full project data object. Throws error if not found.
     */
    getProject(requestData) {
      const projectId = requestData.projectId;
      if (!projectId) throw new Error('Missing required parameter: projectId');
      
      // getProject now returns the full object with slides and elements
      const project = this.projectManager.getProject(projectId); 
      
      if (!project) throw new Error(`Project not found: ${projectId}`);
      if (project.error) throw new Error(`Error retrieving project ${projectId}: ${project.error}`); // Handle errors reported by getProject

      // No need to wrap in { success: true, data: ... } here, handleRequest does that.
      return project; 
    }
    
    /**
     * Lists all projects accessible to the user (basic info).
     * * @param {Object} requestData - (Not used currently)
     * @param {string} user - User email
     * @return {Array<Object>} Array of basic project info objects.
     */
    listProjects(requestData, user) {
        try {
            logDebug(`Listing projects for user: ${user}`);
            if (!this.projectManager) throw new Error('Project manager not initialized');
            
            // Get all projects (basic info)
            const allProjects = this.projectManager.getAllProjects();
            if (!Array.isArray(allProjects)) throw new Error('Expected array of projects from ProjectManager');
            
            // Filter projects based on user access
            const accessibleProjects = allProjects.filter(project => {
                if (!project || typeof project !== 'object' || !project.projectId) {
                    logWarning(`Invalid project object encountered during list: ${JSON.stringify(project)}`);
                    return false;
                }
                // Check access using AuthManager
                return this.authManager.hasAccess(project.projectId, user);
            });
            
            logDebug(`Found ${accessibleProjects.length} accessible projects for ${user}.`);
            return accessibleProjects; // Return the array directly
        } catch (error) {
            logError(`Error listing projects for ${user}: ${error.message}\n${error.stack}`);
            return []; // Return empty array on error to avoid breaking client UI
        }
    }
    
    /**
     * Creates a new project.
     * * @param {Object} requestData - { name: string }
     * @param {string} user - User email (Creator)
     * @return {Object} Result object from ProjectManager { success: boolean, projectId: ..., ... }
     */
    createProject(requestData, user) {
      const projectName = requestData.name;
      if (!projectName) throw new Error('Missing required parameter: name');
      
      // Create project
      const result = this.projectManager.createProject(projectName);
      
      // If successful, add the creator as an editor (or owner)
      if (result.success && result.projectId) {
        // Assuming AuthManager has a method to set initial owner/editor
        this.authManager.grantInitialOwnership(result.projectId, user); 
        logInfo(`Granted initial ownership of project ${result.projectId} to ${user}`);
      }
      
      // Return the result object from ProjectManager directly
      return result; 
    }
    
    /**
     * Updates a project, potentially including slides and elements.
     * * @param {Object} requestData - { projectId: string, updates: { title?: string, slides?: [], elements?: [] } }
     * @return {Object} Result object from ProjectManager { success: boolean, ... }
     */
    updateProject(requestData) {
      const projectId = requestData.projectId;
      const updates = requestData.updates || {};
      
      if (!projectId) throw new Error('Missing required parameter: projectId');
      if (Object.keys(updates).length === 0) throw new Error('No updates provided');

      // Log that we received slide/element data if present
      if (updates.slides) logDebug(`Received ${updates.slides.length} slides to update for project ${projectId}`);
      if (updates.elements) logDebug(`Received ${updates.elements.length} elements to update for project ${projectId}`);
      
      // Pass the entire updates object (which might contain slides/elements) to ProjectManager
      return this.projectManager.updateProject(projectId, updates);
    }
    
    /**
     * Deletes a project.
     * * @param {Object} requestData - { projectId: string, deleteDriveFiles?: boolean }
     * @return {Object} Result object from ProjectManager { success: boolean, ... }
     */
    deleteProject(requestData) {
      const projectId = requestData.projectId;
      const deleteDriveFiles = requestData.deleteDriveFiles || false;
      if (!projectId) throw new Error('Missing required parameter: projectId');
      
      return this.projectManager.deleteProject(projectId, deleteDriveFiles);
    }
    
    //=========================================================================
    // SLIDE ENDPOINTS
    //=========================================================================
    
    /**
     * Adds a new slide to a project.
     * * @param {Object} requestData - { projectId: string, slide: object }
     * @return {Object} Result object from ProjectManager { success: boolean, slide: ..., ... }
     */
    addSlide(requestData) {
      const projectId = requestData.projectId;
      const slideData = requestData.slide || {};
      if (!projectId) throw new Error('Missing required parameter: projectId');
      
      return this.projectManager.addSlide(projectId, slideData);
    }
    
    /**
     * Updates a slide.
     * * @param {Object} requestData - { projectId: string, slideId: string, updates: object }
     * @return {Object} Result object from ProjectManager { success: boolean, ... }
     */
    updateSlide(requestData) {
      const projectId = requestData.projectId;
      const slideId = requestData.slideId;
      const updates = requestData.updates || {};
      
      if (!projectId) throw new Error('Missing required parameter: projectId');
      if (!slideId) throw new Error('Missing required parameter: slideId');
      if (Object.keys(updates).length === 0) throw new Error('No updates provided for slide');
      
      return this.projectManager.updateSlide(projectId, slideId, updates);
    }
    
    /**
     * Deletes a slide.
     * * @param {Object} requestData - { projectId: string, slideId: string }
     * @return {Object} Result object from ProjectManager { success: boolean, ... }
     */
    deleteSlide(requestData) {
      const projectId = requestData.projectId;
      const slideId = requestData.slideId;
      if (!projectId) throw new Error('Missing required parameter: projectId');
      if (!slideId) throw new Error('Missing required parameter: slideId');
      
      return this.projectManager.deleteSlide(projectId, slideId);
    }
    
    //=========================================================================
    // ELEMENT ENDPOINTS
    //=========================================================================
    
    /**
     * Adds a new element to a slide.
     * * @param {Object} requestData - { projectId: string, slideId: string, element: object }
     * @return {Object} Result object from ProjectManager { success: boolean, element: ..., ... }
     */
    addElement(requestData) {
      const projectId = requestData.projectId;
      const slideId = requestData.slideId;
      const elementData = requestData.element || {};
      
      if (!projectId) throw new Error('Missing required parameter: projectId');
      if (!slideId) throw new Error('Missing required parameter: slideId');
      
      return this.projectManager.addElement(projectId, slideId, elementData);
    }
    
    /**
     * Updates an element.
     * * @param {Object} requestData - { projectId: string, elementId: string, updates: object }
     * @return {Object} Result object from ProjectManager { success: boolean, ... }
     */
    updateElement(requestData) {
      const projectId = requestData.projectId;
      const elementId = requestData.elementId;
      const updates = requestData.updates || {};
      
      if (!projectId) throw new Error('Missing required parameter: projectId');
      if (!elementId) throw new Error('Missing required parameter: elementId');
      if (Object.keys(updates).length === 0) throw new Error('No updates provided for element');
      
      return this.projectManager.updateElement(projectId, elementId, updates);
    }
    
    /**
     * Deletes an element.
     * * @param {Object} requestData - { projectId: string, elementId: string }
     * @return {Object} Result object from ProjectManager { success: boolean, ... }
     */
    deleteElement(requestData) {
      const projectId = requestData.projectId;
      const elementId = requestData.elementId;
      if (!projectId) throw new Error('Missing required parameter: projectId');
      if (!elementId) throw new Error('Missing required parameter: elementId');
      
      return this.projectManager.deleteElement(projectId, elementId);
    }
    
    //=========================================================================
    // MEDIA ENDPOINTS
    //=========================================================================
    
    /**
     * Processes a media URL (e.g., uploads to Drive if external).
     * * @param {Object} requestData - { projectId: string, url: string, options?: object }
     * @return {Object} Processed media data (e.g., { fileId: ..., downloadUrl: ... })
     */
    processMediaUrl(requestData) {
      const projectId = requestData.projectId;
      const url = requestData.url;
      const options = requestData.options || {}; // e.g., { typeHint: 'image' }
      
      if (!projectId) throw new Error('Missing required parameter: projectId');
      if (!url) throw new Error('Missing required parameter: url');
      
      // Assuming FileUploader handles fetching/uploading and returns result object
      return this.fileUploader.processUrl(projectId, url, options);
    }
    
    //=========================================================================
    // USER TRACKING ENDPOINTS (Placeholders)
    //=========================================================================
    
    saveUserProgress(requestData, user) {
      const projectId = requestData.projectId;
      const progress = requestData.progress || {};
      if (!projectId) throw new Error('Missing required parameter: projectId');
      
      progress.userId = user; // Ensure user is associated
      progress.timestamp = new Date().getTime();
      
      // TODO: Implement actual saving logic (e.g., using UserTracker class)
      logInfo(`Simulating save progress for user ${user}, project ${projectId}`);
      // const tracker = new UserTracker(projectId); // Assuming UserTracker exists
      // return tracker.saveProgress(user, progress);
      return { success: true, message: 'Progress saved (simulated)', timestamp: progress.timestamp };
    }
    
    getUserProgress(requestData, user) {
      const projectId = requestData.projectId;
      if (!projectId) throw new Error('Missing required parameter: projectId');
      
      // TODO: Implement actual retrieval logic
      logInfo(`Simulating get progress for user ${user}, project ${projectId}`);
      // const tracker = new UserTracker(projectId);
      // return tracker.getProgress(user);
      return { projectId: projectId, userId: user, slides: {}, completionStatus: { started: false, completed: false, percentComplete: 0, lastActive: null } };
    }
    
    //=========================================================================
    // ANALYTICS ENDPOINTS (Placeholders)
    //=========================================================================
    
    getAnalyticsData(requestData) {
      const projectId = requestData.projectId;
      if (!projectId) throw new Error('Missing required parameter: projectId');
      
      // TODO: Implement actual analytics retrieval (e.g., using AnalyticsManager)
      logInfo(`Simulating get analytics for project ${projectId}`);
      // const analytics = new AnalyticsManager(projectId);
      // return analytics.getDashboardData();
      return { projectId: projectId, totalViews: 0, uniqueUsers: 0, completionRate: 0, averageTime: 0, slideData: [], recentActivity: [] };
    }
    
    //=========================================================================
    // AUTH ENDPOINTS
    //=========================================================================
    
    /**
     * Handles user login attempt (checks access).
     * * @param {Object} requestData - { projectId: string }
     * @param {string} user - User email from Session
     * @return {Object} Authentication result { success: boolean, user: string, accessLevel: string }
     */
    login(requestData, user) {
      // In Apps Script web apps, the user is implicitly logged in via Google Account.
      // This endpoint primarily checks *authorization* for a specific project.
      const projectId = requestData.projectId;
      if (!projectId) throw new Error('Missing required parameter: projectId for login check');
      
      const hasAccess = this.authManager.hasAccess(projectId, user);
      const accessLevel = hasAccess ? this.authManager.getAccessLevel(projectId, user) : 'none';
      
      logInfo(`Login check for user ${user}, project ${projectId}: Access=${hasAccess}, Level=${accessLevel}`);
      
      return {
        success: hasAccess, // Success means they have *some* level of access
        user: user,
        projectId: projectId,
        accessLevel: accessLevel
      };
    }
    
    /**
     * Handles user logout. (No-op in standard Apps Script flow).
     * * @return {Object} Logout result.
     */
    logout() {
      // Logout is typically handled by Google's sign-out mechanism, not within the script.
      return { success: true, message: 'Logout endpoint called (no server action needed).' };
    }
    
    /**
     * Gets authentication status for the current user and optionally a project.
     * * @param {Object} requestData - { projectId?: string }
     * @param {string} user - User email from Session
     * @return {Object} Auth status { isAuthenticated: boolean, user: string, hasAccess?: boolean, accessLevel?: string }
     */
    getAuthStatus(requestData, user) {
      const projectId = requestData.projectId; // Optional projectId
      let hasAccess = true; // Assume general access if no project specified
      let accessLevel = 'none';
      
      if (projectId) {
        hasAccess = this.authManager.hasAccess(projectId, user);
        accessLevel = hasAccess ? this.authManager.getAccessLevel(projectId, user) : 'none';
      }
      
      return {
        isAuthenticated: true, // User is always authenticated via Google in Apps Script
        user: user,
        projectId: projectId, // Include if provided
        hasAccess: hasAccess, // Specific project access
        accessLevel: accessLevel // Specific project role
      };
    }
  }
