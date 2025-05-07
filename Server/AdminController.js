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
    const projectFolderName = `${projectTitle} [${projectId.substring(0, 8)}]`;
    const projectFolderId = createDriveFolder(projectFolderName, ROOT_PROJECT_FOLDER_ID);
    if (!projectFolderId) {
        Logger.log(`createProject: Failed to create project folder for "${projectFolderName}". createDriveFolder returned falsy.`);
        return { success: false, error: "Failed to create project folder in Drive." };
    }
    Logger.log(`createProject: Project folder created: ${projectFolderId} for project ${projectId}`);

    // 2. Create an initial project_data.json file
    const initialJsonContent = {
      projectId: projectId,
      title: projectTitle,
      status: "Draft", 
      slides: [],
      createdDate: new Date().toISOString(),
      lastModified: new Date().toISOString()
    };
    const projectDataFileId = saveJsonToDriveFile(
      PROJECT_DATA_FILENAME, 
      JSON.stringify(initialJsonContent, null, 2), 
      projectFolderId,
      null 
    );
    if (!projectDataFileId) {
        Logger.log(`createProject: Failed to create project data file for project ${projectId}. saveJsonToDriveFile returned falsy.`);
        return { success: false, error: "Failed to create project data file in Drive." };
    }
    Logger.log(`createProject: Project data file created: ${projectDataFileId} for project ${projectId}`);

    // 3. Add a new row to the "ProjectIndex" sheet
    const newRowData = [
      projectId,              
      projectTitle,           
      projectFolderId,        
      "Draft",                
      projectDataFileId,      
      new Date().toISOString(), 
      new Date().toISOString()  
    ];

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

/**
 * Retrieves all projects from the "ProjectIndex" sheet for admin display.
 * @return {Array<Object>} An array of project objects, or an empty array if none/error.
 */
function getAllProjectsForAdmin() {
  try {
    const projectsData = getAllSheetData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME);

    if (!projectsData || !Array.isArray(projectsData)) {
      Logger.log("getAllProjectsForAdmin: No data or invalid data returned from getAllSheetData.");
      return []; 
    }
    
    const formattedProjects = projectsData.map(rowObject => {
      return {
        projectId: rowObject['ProjectID'] || rowObject[COL_PROJECT_ID - 1], 
        projectTitle: rowObject['ProjectTitle'] || rowObject[COL_PROJECT_TITLE - 1],
        status: rowObject['Status'] || rowObject[COL_STATUS - 1],
        projectFolderId: rowObject['ProjectFolderID'] || rowObject[COL_PROJECT_FOLDER_ID -1], 
        projectDataFileId: rowObject['ProjectDataFileID'] || rowObject[COL_PROJECT_DATA_FILE_ID -1] 
      };
    }).filter(p => p.projectId); 

    Logger.log(`getAllProjectsForAdmin: Retrieved ${formattedProjects.length} projects.`);
    return formattedProjects;

  } catch (e) {
    Logger.log(`Error in getAllProjectsForAdmin: ${e.toString()} \nStack: ${e.stack}`);
    return []; 
  }
}

/**
 * Helper function to get ProjectFolderID from the ProjectIndex sheet using ProjectID.
 * @param {string} projectId The ID of the project.
 * @return {string|null} The ProjectFolderID if found, otherwise null.
 */
function getProjectFolderIdFromSheet(projectId) {
  const allProjects = getAllSheetData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME);
  if (!allProjects || !Array.isArray(allProjects)) {
    Logger.log(`getProjectFolderIdFromSheet: Could not retrieve project list for projectId "${projectId}".`);
    return null;
  }

  for (let i = 0; i < allProjects.length; i++) {
    const project = allProjects[i];
    const currentProjectId = project['ProjectID'] || (project[COL_PROJECT_ID - 1]);
    if (currentProjectId === projectId) {
      const folderId = project['ProjectFolderID'] || (project[COL_PROJECT_FOLDER_ID - 1]);
      Logger.log(`getProjectFolderIdFromSheet: Found folderId "${folderId}" for projectId "${projectId}".`);
      return folderId;
    }
  }
  Logger.log(`getProjectFolderIdFromSheet: ProjectId "${projectId}" not found in sheet.`);
  return null;
}

/**
 * Uploads a file (image, audio, etc.) to the specified project's folder in Google Drive.
 * @param {Object} fileData An object containing { fileName, mimeType, data (base64 string) }.
 * @param {string} projectId The ID of the project to associate the file with.
 * @param {string} mediaType A string descriptor like 'image' or 'audio' (for logging/future use).
 * @return {Object} An object like { success: true, driveFileId: '...', webContentLink: '...' } or { success: false, error: '...' }.
 */
function uploadFileToDrive(fileData, projectId, mediaType) { // THIS IS THE CORRECT IMPLEMENTATION
  try {
    Logger.log(`uploadFileToDrive: Starting upload for projectId: ${projectId}, mediaType: ${mediaType}, fileName: ${fileData.fileName}`);

    if (!fileData || !fileData.data || !fileData.fileName || !fileData.mimeType) {
      Logger.log("uploadFileToDrive: Invalid fileData received.");
      return { success: false, error: "Invalid file data provided." };
    }
    if (!projectId) {
      Logger.log("uploadFileToDrive: ProjectID is missing.");
      return { success: false, error: "Project ID is required." };
    }

    const projectFolderId = getProjectFolderIdFromSheet(projectId);
    if (!projectFolderId) {
      Logger.log(`uploadFileToDrive: Could not find ProjectFolderID for projectId "${projectId}".`);
      return { success: false, error: `Project folder not found for project ID: ${projectId}.` };
    }

    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    
    const driveFile = createFileInDriveFromBlob(blob, fileData.fileName, projectFolderId);
      
    // MORE DETAILED LOGGING FOR driveFile
    Logger.log(`uploadFileToDrive: Type of driveFile: ${typeof driveFile}`);
    if (driveFile) {
      Logger.log(`uploadFileToDrive: driveFile.getId() exists? ${typeof driveFile.getId === 'function'}`);
      Logger.log(`uploadFileToDrive: driveFile.getName() exists? ${typeof driveFile.getName === 'function'}`);
      Logger.log(`uploadFileToDrive: driveFile.getWebContentLink exists? ${typeof driveFile.getWebContentLink === 'function'}`); // This is the key check
      Logger.log(`uploadFileToDrive: driveFile.getDownloadUrl exists? ${typeof driveFile.getDownloadUrl === 'function'}`);
      Logger.log(`uploadFileToDrive: driveFile properties: ${JSON.stringify(Object.keys(driveFile))}`); // See what methods/props it reports
    } else {
      Logger.log("uploadFileToDrive: driveFile is null or undefined after creation attempt.");
    }
    // END DETAILED LOGGING

    if (!driveFile || !driveFile.getId) { 
        Logger.log(`uploadFileToDrive: Failed to create file in Drive. createFileInDriveFromBlob returned invalid response for ${fileData.fileName}`);
        return { success: false, error: "Failed to create file in Google Drive." };
    }
    
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log(`uploadFileToDrive: File "${driveFile.getName()}" (ID: ${driveFile.getId()}) shared as ANYONE_WITH_LINK.`);
    
    let webContentLink = null; // Initialize to null

    // Check if getWebContentLink method exists AND IS A FUNCTION on the driveFile object
    if (driveFile && typeof driveFile.getWebContentLink === 'function') {
      webContentLink = driveFile.getWebContentLink(); // Call it only if it exists as a function
      Logger.log(`uploadFileToDrive: Called driveFile.getWebContentLink(), result: ${webContentLink}`);
    } else {
      Logger.log(`uploadFileToDrive: driveFile.getWebContentLink is not available or not a function. Object keys: ${JSON.stringify(Object.keys(driveFile || {}))}`);
    }

    if (!webContentLink) { // If it was null from the call, or the function didn't exist/wasn't callable
        Logger.log(`uploadFileToDrive: webContentLink is null or method was not callable. Using fallback for file ID: ${driveFile.getId()}`);
        if (mediaType === 'image') {
          webContentLink = 'https://drive.google.com/uc?export=view&id=' + driveFile.getId();
        } else {
           webContentLink = driveFile.getDownloadUrl().replace("&export=download", "&export=view");
        }
    }

    Logger.log(`uploadFileToDrive: File uploaded successfully. ID: ${driveFile.getId()}, Final WebContentLink: ${webContentLink}`);
    return {
      success: true,
      driveFileId: driveFile.getId(),
      webContentLink: webContentLink,
      fileName: driveFile.getName()
    };

  } catch (e) {
    Logger.log(`Error in uploadFileToDrive: ${e.toString()} - ProjectID: ${projectId}, FileName: ${fileData ? fileData.fileName : 'N/A'} \nStack: ${e.stack}`);
    return { success: false, error: `Failed to upload file: ${e.message}` };
  }
}

// --- Placeholder functions for later steps ---
function saveProjectData(projectId, projectDataJSON) {
  Logger.log(`saveProjectData called for ${projectId}, but not yet implemented.`);
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
  return null; 
}