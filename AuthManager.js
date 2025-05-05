/**
 * AuthManager.js - Authentication and permission management for the Interactive Training Projects Web App
 */

/**
 * AuthManager class handles authentication and permission management
 */
class AuthManager {
  /**
   * Creates a new AuthManager instance
   * 
   * @param {SheetAccessor} sheetAccessor - SheetAccessor instance
   * @param {DriveManager} driveManager - DriveManager instance
   */
  constructor(sheetAccessor, driveManager) {
    this.sheetAccessor = sheetAccessor;
    this.driveManager = driveManager;
    
    // Define permission levels
    this.accessLevels = {
      OWNER: 'owner',
      ADMIN: 'admin',    // New admin access level
      EDITOR: 'editor',
      VIEWER: 'viewer',
      NONE: 'none'
    };
    
    // Define required permissions for actions
    this.actionPermissions = {
      // Project actions
      'project.get': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR, this.accessLevels.VIEWER],
      'project.list': [], // Any authenticated user can list projects they have access to
      'project.create': [], // Any authenticated user can create a project
      'project.update': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR],
      'project.delete': [this.accessLevels.OWNER, this.accessLevels.ADMIN],
      
      // Slide actions
      'slide.add': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR],
      'slide.update': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR],
      'slide.delete': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR],
      
      // Element actions
      'element.add': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR],
      'element.update': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR],
      'element.delete': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR],
      
      // Media actions
      'media.process': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR],
      
      // User tracking actions
      'tracking.saveProgress': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR, this.accessLevels.VIEWER],
      'tracking.getProgress': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR, this.accessLevels.VIEWER],
      
      // Analytics actions
      'analytics.getProjectData': [this.accessLevels.OWNER, this.accessLevels.ADMIN, this.accessLevels.EDITOR],
      
      // Admin specific actions
      'admin.addUser': [this.accessLevels.OWNER, this.accessLevels.ADMIN],
      'admin.removeUser': [this.accessLevels.OWNER, this.accessLevels.ADMIN],
      'admin.changePermission': [this.accessLevels.OWNER, this.accessLevels.ADMIN],
      
      // Auth actions
      'auth.login': [], // Anyone can attempt to login
      'auth.logout': [], // Anyone can logout
      'auth.getStatus': [] // Anyone can check their auth status
    };
  }
  
  /**
   * Checks if a user has access to a project
   * 
   * @param {string} projectId - ID of the project
   * @param {string} userEmail - Email of the user
   * @return {boolean} True if the user has access
   */
  hasAccess(projectId, userEmail) {
    try {
      // Validate inputs
      if (!projectId || !userEmail) {
        return false;
      }
      
      // Get access level - if not NONE, then user has access
      const accessLevel = this.getAccessLevel(projectId, userEmail);
      return accessLevel !== this.accessLevels.NONE;
    } catch (error) {
      logError(`Error checking access: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Gets the access level for a user on a project
   * 
   * @param {string} projectId - ID of the project
   * @param {string} userEmail - Email of the user
   * @return {string} Access level (owner, admin, editor, viewer, none)
   */
  getAccessLevel(projectId, userEmail) {
    try {
      // Check if user is a spreadsheet owner (global admin)
      if (this.isSpreadsheetOwner(userEmail)) {
        return this.accessLevels.OWNER;
      }
      
      // Check if user is a project admin
      if (this.isProjectAdmin(projectId, userEmail)) {
        return this.accessLevels.ADMIN;
      }
      
      // Get project info
      const project = this.getProjectInfo(projectId);
      
      if (!project) {
        return this.accessLevels.NONE;
      }
      
      // Check spreadsheet editor status if no project folder
      if (!project.folderId) {
        const isSpreadsheetEditor = this.isSpreadsheetEditor(userEmail);
        return isSpreadsheetEditor ? this.accessLevels.EDITOR : this.accessLevels.NONE;
      }
      
      // Check folder permissions after checking admin status
      return this.getFolderAccessLevel(project.folderId, userEmail);
    } catch (error) {
      logError(`Error getting access level: ${error.message}`);
      return this.accessLevels.NONE;
    }
  }
  
  /**
   * Checks if a user is an admin for a specific project
   * 
   * @param {string} projectId - ID of the project
   * @param {string} userEmail - Email of the user
   * @return {boolean} True if the user is a project admin
   */
  isProjectAdmin(projectId, userEmail) {
    try {
      if (!projectId || !userEmail) {
        return false;
      }
      
      // Get the ProjectIndex sheet
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const indexSheet = ss.getSheetByName(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
      
      if (!indexSheet) {
        logError('Project index sheet not found');
        return false;
      }
      
      // Find the project in the index
      const projectIdCol = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID;
      const adminUsersCol = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.ADMIN_USERS; // Assumed column for admin users
      
      const data = indexSheet.getDataRange().getValues();
      let projectRow = -1;
      
      for (let i = 1; i < data.length; i++) { // Skip header row
        if (data[i][projectIdCol] === projectId) {
          projectRow = i;
          break;
        }
      }
      
      if (projectRow === -1) {
        logError(`Project not found in index: ${projectId}`);
        return false;
      }
      
      // Check if the user is in the admin column
      // The admin column may contain a comma-separated list of admin emails
      const adminUsers = data[projectRow][adminUsersCol] || '';
      
      if (!adminUsers) {
        return false;
      }
      
      // Split the admin users string by commas and check if userEmail is in the list
      const adminUsersList = adminUsers.split(',').map(email => email.trim().toLowerCase());
      return adminUsersList.includes(userEmail.toLowerCase());
    } catch (error) {
      logError(`Error checking project admin status: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Checks if a user can perform a specific action
   * 
   * @param {string} action - Action to check
   * @param {string} projectId - ID of the project (if applicable)
   * @param {string} userEmail - Email of the user
   * @return {boolean} True if the user can perform the action
   */
  canPerformAction(action, projectId, userEmail) {
    try {
      // Check if action requires specific permissions
      const requiredLevels = this.actionPermissions[action];
      
      // If no specific requirements, any authenticated user can perform the action
      if (!requiredLevels || requiredLevels.length === 0) {
        return true;
      }
      
      // For project-specific actions, check project access
      if (projectId) {
        const accessLevel = this.getAccessLevel(projectId, userEmail);
        return requiredLevels.includes(accessLevel);
      }
      
      // For non-project actions, check if user is a spreadsheet owner or editor
      const isSpreadsheetOwner = this.isSpreadsheetOwner(userEmail);
      const isSpreadsheetEditor = this.isSpreadsheetEditor(userEmail);
      
      return isSpreadsheetOwner || isSpreadsheetEditor;
    } catch (error) {
      logError(`Error checking action permission: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Gets folder access level for a user
   * 
   * @param {string} folderId - ID of the Drive folder
   * @param {string} userEmail - Email of the user
   * @return {string} Access level (owner, editor, viewer, none)
   */
  getFolderAccessLevel(folderId, userEmail) {
    try {
      // Try to access the folder
      try {
        const folder = DriveApp.getFolderById(folderId);
        
        // Check owner
        if (folder.getOwner().getEmail() === userEmail) {
          return this.accessLevels.OWNER;
        }
        
        // Check editors
        const editors = folder.getEditors();
        for (let i = 0; i < editors.length; i++) {
          if (editors[i].getEmail() === userEmail) {
            return this.accessLevels.EDITOR;
          }
        }
        
        // Check viewers
        const viewers = folder.getViewers();
        for (let i = 0; i < viewers.length; i++) {
          if (viewers[i].getEmail() === userEmail) {
            return this.accessLevels.VIEWER;
          }
        }
        
        // Check if folder is publicly accessible
        const sharing = folder.getSharingAccess();
        if (sharing === DriveApp.Access.ANYONE || sharing === DriveApp.Access.ANYONE_WITH_LINK) {
          return this.accessLevels.VIEWER;
        }
        
        // No access found
        return this.accessLevels.NONE;
      } catch (folderError) {
        // If user can't access folder at all, they have no permissions
        return this.accessLevels.NONE;
      }
    } catch (error) {
      logError(`Error getting folder access level: ${error.message}`);
      return this.accessLevels.NONE;
    }
  }
  
  /**
   * Checks if a user is the owner of the spreadsheet
   * 
   * @param {string} userEmail - Email of the user
   * @return {boolean} True if the user is the owner
   */
  isSpreadsheetOwner(userEmail) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const owner = spreadsheet.getOwner();
      return owner && owner.getEmail() === userEmail;
    } catch (error) {
      logError(`Error checking spreadsheet owner: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Checks if a user is an editor of the spreadsheet
   * 
   * @param {string} userEmail - Email of the user
   * @return {boolean} True if the user is an editor
   */
  isSpreadsheetEditor(userEmail) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const editors = spreadsheet.getEditors();
      
      for (let i = 0; i < editors.length; i++) {
        if (editors[i].getEmail() === userEmail) {
          return true;
        }
      }
      
      return false;
    } catch (error) {
      logError(`Error checking spreadsheet editors: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Gets basic project info needed for permissions
   * 
   * @param {string} projectId - ID of the project
   * @return {Object} Project info with folder ID
   */
  getProjectInfo(projectId) {
    try {
      // Find the project in the index
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const indexSheet = ss.getSheetByName(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
      
      if (!indexSheet) {
        logError('Project index sheet not found');
        return null;
      }
      
      // Find project row
      const projectIdCol = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID;
      const projectTitleCol = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE;
      
      const data = indexSheet.getDataRange().getValues();
      let projectRow = null;
      
      for (let i = 1; i < data.length; i++) { // Skip header row
        if (data[i][projectIdCol] === projectId) {
          projectRow = data[i];
          break;
        }
      }
      
      if (!projectRow) {
        logError(`Project not found in index: ${projectId}`);
        return null;
      }
      
      const projectTitle = projectRow[projectTitleCol];
      const projectTabName = sanitizeSheetName(projectTitle);
      
      // Get project sheet
      const projectSheet = ss.getSheetByName(projectTabName);
      
      if (!projectSheet) {
        logError(`Project sheet not found: ${projectTabName}`);
        return null;
      }
      
      // Get project folder ID
      const folderIdRow = SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_FOLDER_ID.ROW;
      const folderIdCol = SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_FOLDER_ID.VALUE_COL;
      const folderId = projectSheet.getRange(folderIdRow, columnToIndex(folderIdCol) + 1).getValue();
      
      return {
        projectId: projectId,
        title: projectTitle,
        folderId: folderId
      };
    } catch (error) {
      logError(`Failed to get project info: ${error.message}`);
      return null;
    }
  }
  
  /**
   * Adds a user as a project admin
   * 
   * @param {string} projectId - ID of the project
   * @param {string} userEmail - Email of the user
   * @return {boolean} True if successful
   */
  addProjectAdmin(projectId, userEmail) {
    try {
      // Get the ProjectIndex sheet
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const indexSheet = ss.getSheetByName(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
      
      if (!indexSheet) {
        logError('Project index sheet not found');
        return false;
      }
      
      // Find the project in the index
      const projectIdCol = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID;
      const adminUsersCol = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.ADMIN_USERS;
      
      const data = indexSheet.getDataRange().getValues();
      let projectRow = -1;
      
      for (let i = 1; i < data.length; i++) { // Skip header row
        if (data[i][projectIdCol] === projectId) {
          projectRow = i + 1; // +1 because sheet rows are 1-indexed
          break;
        }
      }
      
      if (projectRow === -1) {
        logError(`Project not found in index: ${projectId}`);
        return false;
      }
      
      // Get current admin users
      const adminUsers = indexSheet.getRange(projectRow, adminUsersCol + 1).getValue() || '';
      
      // Check if user is already an admin
      const adminUsersList = adminUsers.split(',').map(email => email.trim().toLowerCase());
      if (adminUsersList.includes(userEmail.toLowerCase())) {
        return true; // User is already an admin
      }
      
      // Add user to admin list
      let newAdminUsers = adminUsers;
      if (newAdminUsers) {
        newAdminUsers += `,${userEmail}`;
      } else {
        newAdminUsers = userEmail;
      }
      
      // Update admin column
      indexSheet.getRange(projectRow, adminUsersCol + 1).setValue(newAdminUsers);
      
      // Also add as editor to the folder for convenience
      const project = this.getProjectInfo(projectId);
      if (project && project.folderId) {
        this.driveManager.addEditor(project.folderId, userEmail);
      }
      
      return true;
    } catch (error) {
      logError(`Failed to add project admin: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Removes a user as a project admin
   * 
   * @param {string} projectId - ID of the project
   * @param {string} userEmail - Email of the user
   * @return {boolean} True if successful
   */
  removeProjectAdmin(projectId, userEmail) {
    try {
      // Get the ProjectIndex sheet
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const indexSheet = ss.getSheetByName(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
      
      if (!indexSheet) {
        logError('Project index sheet not found');
        return false;
      }
      
      // Find the project in the index
      const projectIdCol = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID;
      const adminUsersCol = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.ADMIN_USERS;
      
      const data = indexSheet.getDataRange().getValues();
      let projectRow = -1;
      
      for (let i = 1; i < data.length; i++) { // Skip header row
        if (data[i][projectIdCol] === projectId) {
          projectRow = i + 1; // +1 because sheet rows are 1-indexed
          break;
        }
      }
      
      if (projectRow === -1) {
        logError(`Project not found in index: ${projectId}`);
        return false;
      }
      
      // Get current admin users
      const adminUsers = indexSheet.getRange(projectRow, adminUsersCol + 1).getValue() || '';
      
      // Remove user from admin list
      const adminUsersList = adminUsers.split(',').map(email => email.trim());
      const filteredAdmins = adminUsersList.filter(email => 
        email.toLowerCase() !== userEmail.toLowerCase());
      
      // Update admin column
      const newAdminUsers = filteredAdmins.join(',');
      indexSheet.getRange(projectRow, adminUsersCol + 1).setValue(newAdminUsers);
      
      return true;
    } catch (error) {
      logError(`Failed to remove project admin: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Adds a user as an editor to a project
   * 
   * @param {string} projectId - ID of the project
   * @param {string} userEmail - Email of the user
   * @return {boolean} True if successful
   */
  addProjectEditor(projectId, userEmail) {
    try {
      // Get project info
      const project = this.getProjectInfo(projectId);
      
      if (!project || !project.folderId) {
        return false;
      }
      
      // Add user as editor to the project folder
      return this.driveManager.addEditor(project.folderId, userEmail);
    } catch (error) {
      logError(`Failed to add project editor: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Adds a user as a viewer to a project
   * 
   * @param {string} projectId - ID of the project
   * @param {string} userEmail - Email of the user
   * @return {boolean} True if successful
   */
  addProjectViewer(projectId, userEmail) {
    try {
      // Get project info
      const project = this.getProjectInfo(projectId);
      
      if (!project || !project.folderId) {
        return false;
      }
      
      // Add user as viewer to the project folder
      return this.driveManager.addViewer(project.folderId, userEmail);
    } catch (error) {
      logError(`Failed to add project viewer: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Removes a user's access to a project
   * 
   * @param {string} projectId - ID of the project
   * @param {string} userEmail - Email of the user
   * @return {boolean} True if successful
   */
  removeProjectAccess(projectId, userEmail) {
    try {
      // Remove admin access first if applicable
      this.removeProjectAdmin(projectId, userEmail);
      
      // Get project info
      const project = this.getProjectInfo(projectId);
      
      if (!project || !project.folderId) {
        return false;
      }
      
      // Remove access from the project folder
      const folder = DriveApp.getFolderById(project.folderId);
      
      try {
        // Try to remove editor access
        folder.removeEditor(userEmail);
      } catch (e) {
        // Ignore error if user is not an editor
      }
      
      try {
        // Try to remove viewer access
        folder.removeViewer(userEmail);
      } catch (e) {
        // Ignore error if user is not a viewer
      }
      
      return true;
    } catch (error) {
      logError(`Failed to remove project access: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Sets project sharing settings
   * 
   * @param {string} projectId - ID of the project
   * @param {boolean} anyoneWithLink - Whether to allow anyone with link to access
   * @param {string} access - Access level: 'VIEW', 'COMMENT', or 'EDIT'
   * @return {boolean} True if successful
   */
  setProjectSharing(projectId, anyoneWithLink, access) {
    try {
      // Get project info
      const project = this.getProjectInfo(projectId);
      
      if (!project || !project.folderId) {
        return false;
      }
      
      // Set sharing for the project folder
      return this.driveManager.setSharing(project.folderId, anyoneWithLink, access);
    } catch (error) {
      logError(`Failed to set project sharing: ${error.message}`);
      return false;
    }
  }
}