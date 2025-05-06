// Server/AdminController.gs

/**
 * Creates a new project:
 * - Generates a unique ProjectID.
 * - Creates a new folder in Google Drive for the project.
 * - Creates an initial project_data.json file in the new folder.
 * - Adds a new row to the "ProjectIndex" sheet.
 * @param {string} projectTitle The title of the project.
 * @return {object} An object indicating success or failure, and new project details.
 */
function createProject(projectTitle) {
    try {
      if (!projectTitle || projectTitle.trim() === "") {
        Logger.log("createProject: Project title was empty.");
        return { success: false, error: "Project title cannot be empty." };
      }
  
      const projectId = Utilities.getUuid();
      Logger.log(`createProject: Starting for title "${projectTitle}", generated ID: ${projectId}`);
  
      // 1. Create a new folder in Google Drive for the project
      // ROOT_PROJECT_FOLDER_ID is defined in Code.gs
      // createDriveFolder is defined in DriveService.gs
      const projectFolderName = `${projectTitle} [${projectId.substring(0, 8)}]`;
      const projectFolderId = createDriveFolder(projectFolderName, ROOT_PROJECT_FOLDER_ID);
      if (!projectFolderId) {
          Logger.log(`createProject: Failed to create project folder for "${projectFolderName}". createDriveFolder returned falsy.`);
          return { success: false, error: "Failed to create project folder in Drive." };
      }
      Logger.log(`createProject: Project folder created: ${projectFolderId} for project ${projectId}`);
  
      // 2. Create an initial project_data.json file
      // PROJECT_DATA_FILENAME is defined in Code.gs
      const initialJsonContent = {
        projectId: projectId,
        title: projectTitle,
        status: "Draft", // Default status
        slides: [],
        createdDate: new Date().toISOString(),
        lastModified: new Date().toISOString()
      };
      const projectDataFileId = saveJsonToDriveFile(
        PROJECT_DATA_FILENAME, 
        JSON.stringify(initialJsonContent, null, 2), // Pretty print JSON
        projectFolderId,
        null // No file to overwrite for a new project
      );
      if (!projectDataFileId) {
          Logger.log(`createProject: Failed to create project data file for project ${projectId}. saveJsonToDriveFile returned falsy.`);
          return { success: false, error: "Failed to create project data file in Drive." };
      }
      Logger.log(`createProject: Project data file created: ${projectDataFileId} for project ${projectId}`);
  
      // 3. Add a new row to the "ProjectIndex" sheet
      // PROJECT_INDEX_SHEET_ID and PROJECT_INDEX_DATA_SHEET_NAME are defined in Code.gs
      // Column order must match your sheet: ProjectID, ProjectTitle, ProjectFolderID, Status, ProjectDataFileID, LastModified, CreatedDate
      const newRowData = [
        projectId,              // COL_PROJECT_ID
        projectTitle,           // COL_PROJECT_TITLE
        projectFolderId,        // COL_PROJECT_FOLDER_ID
        "Draft",                // COL_STATUS
        projectDataFileId,      // COL_PROJECT_DATA_FILE_ID
        new Date().toISOString(), // COL_LAST_MODIFIED
        new Date().toISOString()  // COL_CREATED_DATE
      ];
  
      // Using appendRowToSheet from SheetService.gs
      appendRowToSheet(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, newRowData);
      Logger.log(`createProject: New project row added to sheet for project ${projectId}: ${newRowData.join(', ')}`);
  
      return {
        success: true,
        message: "Project created successfully!",
        projectId: projectId,
        projectFolderId: projectFolderId,
        projectDataFileId: projectDataFileId
      };
  
    } catch (e) {
      Logger.log(`Error in createProject: ${e.toString()} \nStack: ${e.stack ? e.stack : 'No stack available'}`);
      return { success: false, error: `Failed to create project: ${e.message}` };
    }
  }
  
  // --- Placeholder functions for later steps ---
  
  function getAllProjectsForAdmin() {
    Logger.log("getAllProjectsForAdmin called, but not yet implemented.");
    // This will be implemented in Step 4
    // For now, return an empty array or some mock data if needed for UI testing
    return []; 
  }
  
  function saveProjectData(projectId, projectDataJSON) {
    Logger.log(`saveProjectData called for ${projectId}, but not yet implemented.`);
    return { success: false, message: "Not implemented yet."};
  }
  
  function uploadFileToDrive(fileData, projectId, mediaType) {
    Logger.log(`uploadFileToDrive called for ${projectId}, mediaType ${mediaType}, but not yet implemented.`);
    return { success: false, message: "Not implemented yet."};
  }
  
  function updateProjectStatus(projectId, newStatus) {
    Logger.log(`updateProjectStatus called for ${projectId} to ${newStatus}, but not yet implemented.`);
    return { success: false, message: "Not implemented yet."};
  }
  
  function deleteProject(projectId) {
    Logger.log(`deleteProject called for ${projectId}, but not yet implemented.`);
    return { success: false, message: "Not implemented yet."};
  }
  
  function getProjectDataForEditing(projectId) {
    Logger.log(`getProjectDataForEditing called for ${projectId}, but not yet implemented.`);
    return null; // Or an object structure { success: false, error: "Not implemented" }
  }