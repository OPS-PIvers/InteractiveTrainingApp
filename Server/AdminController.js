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
  
  /**
   * Retrieves all projects from the "ProjectIndex" sheet for admin display.
   * @return {Array<Object>} An array of project objects, or an empty array if none/error.
   */
  function getAllProjectsForAdmin() {
    try {
      // PROJECT_INDEX_SHEET_ID and PROJECT_INDEX_DATA_SHEET_NAME are in Code.gs
      // getAllSheetData is in SheetService.gs
      const projectsData = getAllSheetData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME);

      if (!projectsData || !Array.isArray(projectsData)) {
        Logger.log("getAllProjectsForAdmin: No data or invalid data returned from getAllSheetData.");
        return []; // Return empty array if no data or error in fetching
      }
      
      // Assuming getAllSheetData returns an array of objects where keys are header names
      // We need: ProjectID, ProjectTitle, Status (and ProjectFolderID, ProjectDataFileID for later use)
      // The COL_ constants are defined in Code.gs (1-based)
      // If getAllSheetData returns array of arrays (e.g. if no headers) then adapt this mapping
      
      const formattedProjects = projectsData.map(rowObject => {
        // Ensure keys match your sheet headers exactly if getAllSheetData returns objects
        // If it returns array of arrays, use indices: rowObject[COL_PROJECT_ID - 1]
        return {
          projectId: rowObject['ProjectID'] || rowObject[COL_PROJECT_ID - 1], 
          projectTitle: rowObject['ProjectTitle'] || rowObject[COL_PROJECT_TITLE - 1],
          status: rowObject['Status'] || rowObject[COL_STATUS - 1],
          projectFolderId: rowObject['ProjectFolderID'] || rowObject[COL_PROJECT_FOLDER_ID -1], // For Step 5
          projectDataFileId: rowObject['ProjectDataFileID'] || rowObject[COL_PROJECT_DATA_FILE_ID -1] // For Step 5
        };
      }).filter(p => p.projectId); // Ensure project has a projectId

      Logger.log(`getAllProjectsForAdmin: Retrieved ${formattedProjects.length} projects.`);
      return formattedProjects;

    } catch (e) {
      Logger.log(`Error in getAllProjectsForAdmin: ${e.toString()} \nStack: ${e.stack}`);
      // In case of error, return an empty array to prevent client-side errors
      return []; 
    }
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