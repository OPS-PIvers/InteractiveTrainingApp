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
     * 
     * @param {ProjectManager} projectManager - ProjectManager instance
     * @param {FileUploader} fileUploader - FileUploader instance
     * @param {AuthManager} authManager - AuthManager instance
     */
    constructor(projectManager, fileUploader, authManager) {
      this.projectManager = projectManager;
      this.fileUploader = fileUploader;
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
        
        // User tracking endpoints
        'tracking.saveProgress': this.saveUserProgress.bind(this),
        'tracking.getProgress': this.getUserProgress.bind(this),
        
        // Analytics endpoints
        'analytics.getProjectData': this.getAnalyticsData.bind(this),
        
        // Auth endpoints
        'auth.login': this.login.bind(this),
        'auth.logout': this.logout.bind(this),
        'auth.getStatus': this.getAuthStatus.bind(this)
      };
    }
    
    /**
     * Handles an API request
     * 
     * @param {Object} requestData - Request data from client
     * @param {string} user - Email of the user making the request
     * @return {TextOutput} Response data as JSON
     */
    handleRequest(requestData, user) {
        try {
          // Validate request format
          if (!requestData.action) {
            return {
              success: false,
              error: 'Missing required parameter: action'
            };
          }
          
          // Find the endpoint handler
          const handler = this.endpoints[requestData.action];
          
          if (!handler) {
            return {
              success: false,
              error: `Unknown action: ${requestData.action}`
            };
          }
          
          // Check authorization for this endpoint
          if (!this.authManager.canPerformAction(requestData.action, requestData.projectId, user)) {
            return {
              success: false,
              error: 'Not authorized to perform this action'
            };
          }
          
          // Call the handler with the request data
          const result = handler(requestData, user);
          
          // Return a properly structured response
          return {
            success: true,
            data: result
          };
        } catch (error) {
          logError(`API error: ${error.message}\n${error.stack}`);
          return {
            success: false,
            error: `Server error: ${error.message}`
          };
        }
    }
    
    /**
     * Creates an error response
     * 
     * @param {string} message - Error message
     * @param {number} statusCode - HTTP status code (default: 400)
     * @return {TextOutput} Error response as JSON
     */
    createErrorResponse(message, statusCode = 400) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: message,
        status: statusCode
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // PROJECT ENDPOINTS
    
    /**
     * Gets a project by ID
     * 
     * @param {Object} requestData - Request data with projectId
     * @return {Object} Project data
     */
    getProject(requestData) {
      const projectId = requestData.projectId;
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      const project = this.projectManager.getProject(projectId);
      
      if (!project) {
        throw new Error(`Project not found: ${projectId}`);
      }
      
      return project;
    }
    
    /**
     * Lists all projects accessible to the user
     * 
     * @param {Object} requestData - Request data
     * @param {string} user - User email
     * @return {Array<Object>} Array of project objects
     */
    listProjects(requestData, user) {
        try {
        console.log('Listing projects for user:', user);
        
        if (!this.projectManager) {
            console.error('Project manager not initialized');
            return [];
        }
        
        // Get all projects from the project manager
        const allProjects = this.projectManager.getAllProjects();
        console.log('All projects found:', allProjects ? allProjects.length : 0);
        
        // Extra validation on return value
        if (!allProjects) {
            console.warn('Project manager returned null or undefined');
            return [];
        }
        
        if (!Array.isArray(allProjects)) {
            console.error('Expected array of projects but got:', typeof allProjects);
            return [];
        }
        
        // Filter projects based on user access
        const accessibleProjects = allProjects.filter(project => {
            // Validate each project object
            if (!project || typeof project !== 'object' || !project.projectId) {
            console.warn('Invalid project object:', project);
            return false;
            }
            return this.authManager.hasAccess(project.projectId, user);
        });
        
        console.log('Accessible projects:', accessibleProjects.length);
        
        return accessibleProjects;
        } catch (error) {
        console.error('Error listing projects:', error);
        // Return empty array instead of throwing to avoid breaking the UI
        return [];
        }
    }
    
    /**
     * Creates a new project
     * 
     * @param {Object} requestData - Request data with project details
     * @param {string} user - User email
     * @return {Object} Result of project creation
     */
    createProject(requestData, user) {
      const projectName = requestData.name;
      
      if (!projectName) {
        throw new Error('Missing required parameter: name');
      }
      
      // Create project
      const result = this.projectManager.createProject(projectName);
      
      // If successful, add the creator as an editor
      if (result.success && result.projectId) {
        this.authManager.addProjectEditor(result.projectId, user);
      }
      
      return result;
    }
    
    /**
     * Updates a project
     * 
     * @param {Object} requestData - Request data with project details
     * @return {Object} Result of project update
     */
    updateProject(requestData) {
      const projectId = requestData.projectId;
      const updates = requestData.updates || {};
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      if (Object.keys(updates).length === 0) {
        throw new Error('No updates provided');
      }
      
      return this.projectManager.updateProject(projectId, updates);
    }
    
    /**
     * Deletes a project
     * 
     * @param {Object} requestData - Request data with projectId
     * @return {Object} Result of project deletion
     */
    deleteProject(requestData) {
      const projectId = requestData.projectId;
      const deleteDriveFiles = requestData.deleteDriveFiles || false;
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      return this.projectManager.deleteProject(projectId, deleteDriveFiles);
    }
    
    // SLIDE ENDPOINTS
    
    /**
     * Adds a new slide to a project
     * 
     * @param {Object} requestData - Request data with projectId and slide details
     * @return {Object} Result of slide addition
     */
    addSlide(requestData) {
      const projectId = requestData.projectId;
      const slideData = requestData.slide || {};
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      return this.projectManager.addSlide(projectId, slideData);
    }
    
    /**
     * Updates a slide
     * 
     * @param {Object} requestData - Request data with projectId, slideId, and updates
     * @return {Object} Result of slide update
     */
    updateSlide(requestData) {
      const projectId = requestData.projectId;
      const slideId = requestData.slideId;
      const updates = requestData.updates || {};
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      if (!slideId) {
        throw new Error('Missing required parameter: slideId');
      }
      
      if (Object.keys(updates).length === 0) {
        throw new Error('No updates provided');
      }
      
      return this.projectManager.updateSlide(projectId, slideId, updates);
    }
    
    /**
     * Deletes a slide
     * 
     * @param {Object} requestData - Request data with projectId and slideId
     * @return {Object} Result of slide deletion
     */
    deleteSlide(requestData) {
      const projectId = requestData.projectId;
      const slideId = requestData.slideId;
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      if (!slideId) {
        throw new Error('Missing required parameter: slideId');
      }
      
      return this.projectManager.deleteSlide(projectId, slideId);
    }
    
    // ELEMENT ENDPOINTS
    
    /**
     * Adds a new element to a slide
     * 
     * @param {Object} requestData - Request data with projectId, slideId, and element details
     * @return {Object} Result of element addition
     */
    addElement(requestData) {
      const projectId = requestData.projectId;
      const slideId = requestData.slideId;
      const elementData = requestData.element || {};
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      if (!slideId) {
        throw new Error('Missing required parameter: slideId');
      }
      
      return this.projectManager.addElement(projectId, slideId, elementData);
    }
    
    /**
     * Updates an element
     * 
     * @param {Object} requestData - Request data with projectId, elementId, and updates
     * @return {Object} Result of element update
     */
    updateElement(requestData) {
      const projectId = requestData.projectId;
      const elementId = requestData.elementId;
      const updates = requestData.updates || {};
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      if (!elementId) {
        throw new Error('Missing required parameter: elementId');
      }
      
      if (Object.keys(updates).length === 0) {
        throw new Error('No updates provided');
      }
      
      return this.projectManager.updateElement(projectId, elementId, updates);
    }
    
    /**
     * Deletes an element
     * 
     * @param {Object} requestData - Request data with projectId and elementId
     * @return {Object} Result of element deletion
     */
    deleteElement(requestData) {
      const projectId = requestData.projectId;
      const elementId = requestData.elementId;
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      if (!elementId) {
        throw new Error('Missing required parameter: elementId');
      }
      
      return this.projectManager.deleteElement(projectId, elementId);
    }
    
    // MEDIA ENDPOINTS
    
    /**
     * Processes a media URL
     * 
     * @param {Object} requestData - Request data with projectId and url
     * @return {Object} Processed media data
     */
    processMediaUrl(requestData) {
      const projectId = requestData.projectId;
      const url = requestData.url;
      const options = requestData.options || {};
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      if (!url) {
        throw new Error('Missing required parameter: url');
      }
      
      return this.fileUploader.processUrl(projectId, url, options);
    }
    
    // USER TRACKING ENDPOINTS
    
    /**
     * Saves user progress
     * 
     * @param {Object} requestData - Request data with progress details
     * @param {string} user - User email
     * @return {Object} Result of saving progress
     */
    saveUserProgress(requestData, user) {
      const projectId = requestData.projectId;
      const progress = requestData.progress || {};
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      // Add user identifier
      progress.userId = user;
      progress.timestamp = new Date().getTime();
      
      // In a complete implementation, this would save to a database or sheet
      // For now, we'll simulate successful saving
      
      return {
        success: true,
        message: 'Progress saved successfully',
        timestamp: progress.timestamp
      };
    }
    
    /**
     * Gets user progress for a project
     * 
     * @param {Object} requestData - Request data with projectId
     * @param {string} user - User email
     * @return {Object} User progress data
     */
    getUserProgress(requestData, user) {
      const projectId = requestData.projectId;
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      // In a complete implementation, this would retrieve from a database or sheet
      // For now, we'll return a placeholder
      
      return {
        projectId: projectId,
        userId: user,
        slides: {},
        completionStatus: {
          started: true,
          completed: false,
          percentComplete: 0,
          lastActive: new Date().getTime()
        }
      };
    }
    
    // ANALYTICS ENDPOINTS
    
    /**
     * Gets analytics data for a project
     * 
     * @param {Object} requestData - Request data with projectId
     * @return {Object} Analytics data
     */
    getAnalyticsData(requestData) {
      const projectId = requestData.projectId;
      
      if (!projectId) {
        throw new Error('Missing required parameter: projectId');
      }
      
      // In a complete implementation, this would retrieve from a database or sheet
      // For now, we'll return a placeholder
      
      return {
        projectId: projectId,
        totalViews: 0,
        uniqueUsers: 0,
        completionRate: 0,
        averageTime: 0,
        slideData: [],
        recentActivity: []
      };
    }
    
    // AUTH ENDPOINTS
    
    /**
     * Handles user login
     * 
     * @param {Object} requestData - Request data with login credentials
     * @return {Object} Authentication result
     */
    login(requestData) {
      const email = Session.getEffectiveUser().getEmail();
      const projectId = requestData.projectId;
      
      // Check if user has access to the project
      const hasAccess = this.authManager.hasAccess(projectId, email);
      
      return {
        success: hasAccess,
        user: email,
        projectId: projectId,
        accessLevel: hasAccess ? this.authManager.getAccessLevel(projectId, email) : 'none'
      };
    }
    
    /**
     * Handles user logout
     * Not needed for Apps Script authentication, but included for API completeness
     * 
     * @return {Object} Logout result
     */
    logout() {
      return {
        success: true,
        message: 'Logged out successfully'
      };
    }
    
    /**
     * Gets authentication status for the current user
     * 
     * @param {Object} requestData - Request data
     * @param {string} user - User email
     * @return {Object} Authentication status
     */
    getAuthStatus(requestData, user) {
      const projectId = requestData.projectId;
      
      return {
        isAuthenticated: true,
        user: user,
        projectId: projectId,
        hasAccess: projectId ? this.authManager.hasAccess(projectId, user) : true,
        accessLevel: projectId ? this.authManager.getAccessLevel(projectId, user) : 'none'
      };
    }
  }