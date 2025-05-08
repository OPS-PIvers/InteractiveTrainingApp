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
      status: "Draft", // Default status for new projects
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
      initialJsonContent.status, // Use status from initialJsonContent
      projectDataFileId,      
      initialJsonContent.lastModified, 
      initialJsonContent.createdDate  
    ];

    appendRowToSheet(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, newRowData);
    Logger.log(`createProject: New project row added to sheet for project ${projectId}: ${newRowData.join(', ')}`);

    return {
      success: true,
      message: "Project created successfully!",
      projectId: projectId,
      projectFolderId: projectFolderId,
      projectDataFileId: projectDataFileId,
      status: initialJsonContent.status // Return initial status
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
      // Ensure all expected columns are present in rowObject or fallback to array indices
      const projectId = rowObject['ProjectID'] || (rowObject[COL_PROJECT_ID - 1] ? rowObject[COL_PROJECT_ID - 1] : null);
      const projectTitle = rowObject['ProjectTitle'] || (rowObject[COL_PROJECT_TITLE - 1] ? rowObject[COL_PROJECT_TITLE - 1] : null);
      const status = rowObject['Status'] || (rowObject[COL_STATUS - 1] ? rowObject[COL_STATUS - 1] : null);
      const projectFolderId = rowObject['ProjectFolderID'] || (rowObject[COL_PROJECT_FOLDER_ID -1] ? rowObject[COL_PROJECT_FOLDER_ID -1] : null);
      const projectDataFileId = rowObject['ProjectDataFileID'] || (rowObject[COL_PROJECT_DATA_FILE_ID -1] ? rowObject[COL_PROJECT_DATA_FILE_ID -1] : null);
      
      if (!projectId) return null; // Skip if no projectId

      return {
        projectId: projectId, 
        projectTitle: projectTitle,
        status: status,
        projectFolderId: projectFolderId, 
        projectDataFileId: projectDataFileId 
      };
    }).filter(p => p !== null); // Remove null entries from map if projectId was missing

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
function uploadFileToDrive(fileData, projectId, mediaType) { 
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
      
    Logger.log(`uploadFileToDrive: Type of driveFile: ${typeof driveFile}`);
    if (driveFile) {
      Logger.log(`uploadFileToDrive: driveFile.getId() exists? ${typeof driveFile.getId === 'function'}`);
      Logger.log(`uploadFileToDrive: driveFile.getName() exists? ${typeof driveFile.getName === 'function'}`);
      Logger.log(`uploadFileToDrive: driveFile.getWebContentLink exists? ${typeof driveFile.getWebContentLink === 'function'}`); 
      Logger.log(`uploadFileToDrive: driveFile.getDownloadUrl exists? ${typeof driveFile.getDownloadUrl === 'function'}`);
      Logger.log(`uploadFileToDrive: driveFile properties: ${JSON.stringify(Object.keys(driveFile))}`); 
    } else {
      Logger.log("uploadFileToDrive: driveFile is null or undefined after creation attempt.");
    }

    if (!driveFile || !driveFile.getId) { 
        Logger.log(`uploadFileToDrive: Failed to create file in Drive. createFileInDriveFromBlob returned invalid response for ${fileData.fileName}`);
        return { success: false, error: "Failed to create file in Google Drive." };
    }
    
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log(`uploadFileToDrive: File "${driveFile.getName()}" (ID: ${driveFile.getId()}) shared as ANYONE_WITH_LINK.`);
    
    let webContentLink = null; 

    // Try to get webContentLink first
    if (driveFile && typeof driveFile.getWebContentLink === 'function') {
      webContentLink = driveFile.getWebContentLink(); 
      Logger.log(`uploadFileToDrive: Called driveFile.getWebContentLink(), result: ${webContentLink}`);
    } else {
      Logger.log(`uploadFileToDrive: driveFile.getWebContentLink is not available or not a function.`);
    }

    // Fallback if webContentLink is null or unsuitable
    if (!webContentLink) { 
        Logger.log(`uploadFileToDrive: webContentLink is null or wasn't obtained. Using fallback for file ID: ${driveFile.getId()}`);
        if (mediaType === 'image') {
          // Use direct image view URL
          webContentLink = 'https://drive.google.com/uc?export=view&id=' + driveFile.getId();
        } else {
           // For audio/other, construct a view/download URL (may still have issues in <audio> tag)
           const downloadUrl = driveFile.getDownloadUrl(); // Requires Drive API v2 scope potentially if getDownloadUrl fails
           webContentLink = downloadUrl ? downloadUrl.replace("&export=download", "&export=view") : 'https://drive.google.com/uc?id=' + driveFile.getId();
           // Note: This fallback URL might still NOT work reliably in <audio src="">
        }
    }

    Logger.log(`uploadFileToDrive: File uploaded successfully. ID: ${driveFile.getId()}, Final Link (may not work directly for audio): ${webContentLink}`);
    return {
      success: true,
      driveFileId: driveFile.getId(),
      webContentLink: webContentLink, // Return the link, viewer will decide if usable
      fileName: driveFile.getName(),
      mimeType: driveFile.getMimeType() // Return actual mime type
    };

  } catch (e) {
    Logger.log(`Error in uploadFileToDrive: ${e.toString()} - ProjectID: ${projectId}, FileName: ${fileData ? fileData.fileName : 'N/A'} \nStack: ${e.stack}`);
    return { success: false, error: `Failed to upload file: ${e.message}` };
  }
}


/**
 * Retrieves an image file from Drive as a base64 data URI.
 * @param {string} driveFileId The ID of the file in Google Drive.
 * @return {object} An object like { success: true, base64Data: 'data:mime/type;base64,...' } or { success: false, error: '...' }.
 */
function getImageAsBase64(driveFileId) {
  try {
    if (!driveFileId) {
      return { success: false, error: "Drive File ID is required." };
    }
    const file = DriveApp.getFileById(driveFileId);
    const blob = file.getBlob();
    const mimeType = blob.getContentType();
    // Ensure it's an image type before encoding? Optional.
    if (!mimeType || !mimeType.startsWith('image/')) {
       Logger.log(`getImageAsBase64: File ID ${driveFileId} is not an image (MIME: ${mimeType}).`);
       return { success: false, error: `File is not an image (type: ${mimeType})` };
    }
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const dataURI = 'data:' + mimeType + ';base64,' + base64Data;
    
    Logger.log(`getImageAsBase64: Successfully retrieved and encoded file ID: ${driveFileId}. MimeType: ${mimeType}. DataURI length: ${dataURI.length}`);
    return { success: true, base64Data: dataURI, mimeType: mimeType };

  } catch (e) {
     if (e.message.toLowerCase().includes("access denied") || e.message.toLowerCase().includes("not found")) {
          Logger.log(`getImageAsBase64: File not found or access denied for ID ${driveFileId}. Error: ${e.toString()}`);
          return { success: false, error: "Image file not found or access denied." };
      }
    Logger.log(`Error in getImageAsBase64 for file ID ${driveFileId}: ${e.toString()} \nStack: ${e.stack}`);
    return { success: false, error: `Failed to retrieve image as base64: ${e.message}` };
  }
}

/**
 * Saves the project's data (slides, overlays, media links, canvas dimensions) to its JSON file in Drive.
 * Updates the LastModified timestamp in the ProjectIndex sheet.
 * Also ensures the project's status in the sheet matches the status in the projectData.
 * @param {string} projectId The ID of the project.
 * @param {string} projectDataJSON A JSON string representing the entire project data.
 * @return {object} An object indicating success or failure.
 */
function saveProjectData(projectId, projectDataJSON) {
  try {
    Logger.log(`saveProjectData: Attempting to save data for projectId: ${projectId}`);
    if (!projectId || !projectDataJSON) {
      Logger.log("saveProjectData: Missing projectId or projectDataJSON.");
      return { success: false, error: "Project ID and data are required." };
    }

    let projectDataParsed;
    try {
        projectDataParsed = JSON.parse(projectDataJSON);
    } catch (parseError) {
        Logger.log(`saveProjectData: Error parsing projectDataJSON for projectId ${projectId}: ${parseError.toString()}`);
        return { success: false, error: "Invalid project data format. Could not parse JSON." };
    }
    
    const currentStatusInJson = projectDataParsed.status || "Draft"; 

    const allProjects = getAllSheetData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME);
    let projectEntry = null;
    let projectRowIndex = -1; // Initialize to -1 to indicate not found

    if (allProjects && Array.isArray(allProjects)) {
        // Determine if getAllSheetData returned headers
        const headerRow = allProjects.length > 0 ? allProjects[0] : null;
        const looksLikeHeaders = headerRow && headerRow.every(h => typeof h === 'string');

      for (let i = 0; i < allProjects.length; i++) {
        const p = allProjects[i];
        const currentSheetProjectId = typeof p === 'object' ? p['ProjectID'] : p[COL_PROJECT_ID - 1]; // Handle object or array
        if (currentSheetProjectId === projectId) {
          projectEntry = p;
          projectRowIndex = i + (looksLikeHeaders ? 2 : 1); // Adjust index based on headers
          break;
        }
      }
       // Fallback if loop didn't find it but maybe findRowIndexByValue can (e.g., if getAllSheetData failed partially)
       if (!projectEntry) {
            projectRowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
       }
    } else {
        // If getAllSheetData failed, try findRowIndexByValue directly
         projectRowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
    }


    if (projectRowIndex === -1) { // Check if row index was found
      Logger.log(`saveProjectData: Project with ID "${projectId}" not found in ProjectIndex.`);
      return { success: false, error: `Project metadata not found for ID: ${projectId}. Cannot save.` };
    }
    
    // If we only have rowIndex, fetch projectEntry data
    if (!projectEntry) {
        const rowValues = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, projectRowIndex);
        if (rowValues) {
             // Reconstruct projectEntry minimally for folder/file IDs
             projectEntry = {};
             projectEntry['ProjectFolderID'] = rowValues[COL_PROJECT_FOLDER_ID - 1];
             projectEntry['ProjectDataFileID'] = rowValues[COL_PROJECT_DATA_FILE_ID - 1];
        } else {
             Logger.log(`saveProjectData: Found row index ${projectRowIndex} but failed to fetch row data for ${projectId}.`);
             return { success: false, error: `Failed to retrieve project details for ID: ${projectId}. Cannot save.` };
        }
    }


    const projectFolderId = typeof projectEntry === 'object' ? projectEntry['ProjectFolderID'] : projectEntry[COL_PROJECT_FOLDER_ID - 1];
    const projectDataFileIdToOverwrite = typeof projectEntry === 'object' ? projectEntry['ProjectDataFileID'] : projectEntry[COL_PROJECT_DATA_FILE_ID - 1];

    if (!projectFolderId || !projectDataFileIdToOverwrite) {
      Logger.log(`saveProjectData: Missing FolderID or DataFileID for project "${projectId}". FolderID: ${projectFolderId}, FileID: ${projectDataFileIdToOverwrite}`);
      return { success: false, error: "Project folder or data file reference is missing. Cannot save." };
    }
    Logger.log(`saveProjectData: Found FolderID: ${projectFolderId}, DataFileID to overwrite: ${projectDataFileIdToOverwrite} for project ${projectId}.`);

    const savedFileId = saveJsonToDriveFile(
      PROJECT_DATA_FILENAME,
      projectDataJSON, 
      projectFolderId,
      projectDataFileIdToOverwrite 
    );

    if (!savedFileId) {
      Logger.log(`saveProjectData: Failed to save JSON file to Drive for project ${projectId}. saveJsonToDriveFile returned falsy.`);
      return { success: false, error: "Failed to save project data to Google Drive." };
    }
    Logger.log(`saveProjectData: JSON data saved to file ID: ${savedFileId} for project ${projectId}.`);

    if (projectRowIndex) {
      let rowDataArray = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, projectRowIndex);
      if (rowDataArray) {
          rowDataArray[COL_LAST_MODIFIED - 1] = new Date().toISOString(); 
          rowDataArray[COL_STATUS - 1] = currentStatusInJson; 
          updateSheetRow(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, projectRowIndex, rowDataArray);
          Logger.log(`saveProjectData: Updated LastModified and Status ("${currentStatusInJson}") for project ${projectId} at row ${projectRowIndex}.`);
      } else {
            Logger.log(`saveProjectData: Could not retrieve row data for project ${projectId} to update sheet. Row ${projectRowIndex} might be empty or error.`);
      }
    } else {
      Logger.log(`saveProjectData: Could not find row for project ${projectId} to update sheet.`);
    }

    return { success: true, message: "Project data saved successfully." };

  } catch (e) {
    Logger.log(`Error in saveProjectData for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
    return { success: false, error: `Failed to save project data: ${e.message}` };
  }
}

/**
 * Updates the status of a project in the "ProjectIndex" sheet.
 * Also updates the "LastModified" timestamp.
 * @param {string} projectId The ID of the project to update.
 * @param {string} newStatus The new status (e.g., "Active", "Inactive", "Draft").
 * @return {object} An object indicating success or failure.
 */
function updateProjectStatus(projectId, newStatus) {
  try {
    Logger.log(`updateProjectStatus: Attempting to update status for projectId: ${projectId} to "${newStatus}"`);
    if (!projectId || !newStatus) {
      Logger.log("updateProjectStatus: Missing projectId or newStatus.");
      return { success: false, error: "Project ID and new status are required." };
    }

    const validStatuses = ["Draft", "Active", "Inactive"];
    if (validStatuses.indexOf(newStatus) === -1) {
        Logger.log(`updateProjectStatus: Invalid status value "${newStatus}" for projectId: ${projectId}`);
        return { success: false, error: `Invalid status value: ${newStatus}. Must be one of ${validStatuses.join(', ')}.` };
    }

    const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);

    if (!rowIndex) {
      Logger.log(`updateProjectStatus: Project with ID "${projectId}" not found in ProjectIndex.`);
      return { success: false, error: `Project not found for ID: ${projectId}. Cannot update status.` };
    }

    let rowData = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
    if (!rowData) {
        Logger.log(`updateProjectStatus: Could not retrieve current row data for project ${projectId} at row ${rowIndex}.`);
        return { success: false, error: "Could not retrieve project data to update status."};
    }
    
    rowData[COL_STATUS - 1] = newStatus;
    rowData[COL_LAST_MODIFIED - 1] = new Date().toISOString();

    updateSheetRow(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex, rowData);
    Logger.log(`updateProjectStatus: Status for project ${projectId} (row ${rowIndex}) updated to "${newStatus}". LastModified also updated.`);

    // Update status in project_data.json
    const projectFolderId = rowData[COL_PROJECT_FOLDER_ID - 1];
    const projectDataFileId = rowData[COL_PROJECT_DATA_FILE_ID - 1];

    if(projectDataFileId && projectFolderId){
        try {
            let projectJsonContent = readDriveFileContent(projectDataFileId);
            let projectJson = JSON.parse(projectJsonContent);
            projectJson.status = newStatus; // Update status in JSON
            projectJson.lastModified = new Date().toISOString(); // Update lastModified in JSON
            saveJsonToDriveFile(PROJECT_DATA_FILENAME, JSON.stringify(projectJson, null, 2), projectFolderId, projectDataFileId);
            Logger.log(`updateProjectStatus: Status and lastModified in project_data.json for ${projectId} also updated.`);
        } catch (e) {
            Logger.log(`updateProjectStatus: Failed to update status/lastModified in project_data.json for ${projectId}. Error: ${e.toString()}`);
        }
    } else {
        Logger.log(`updateProjectStatus: Could not find folderId or dataFileId for project ${projectId} from sheet to update JSON. FolderID: ${projectFolderId}, FileID: ${projectDataFileId}`);
    }

    return { success: true, message: `Project status successfully updated to ${newStatus}.` };

  } catch (e) {
    Logger.log(`Error in updateProjectStatus for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
    return { success: false, error: `Failed to update project status: ${e.message}` };
  }
}

/**
 * Deletes a project:
 * - Removes the project's row from the "ProjectIndex" sheet.
 * - Deletes the project's folder (and all its contents) from Google Drive.
 * @param {string} projectId The ID of the project to delete.
 * @return {object} An object indicating success or failure.
 */
function deleteProject(projectId) {
  try {
    Logger.log(`deleteProject: Attempting to delete project with ID: ${projectId}`);
    if (!projectId) {
      Logger.log("deleteProject: ProjectID is missing.");
      return { success: false, error: "Project ID is required for deletion." };
    }

    // 1. Find the row and ProjectFolderID from the ProjectIndex sheet
    const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
    let projectFolderId = null;

    if (rowIndex) {
        const rowData = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
        if (rowData) {
            projectFolderId = rowData[COL_PROJECT_FOLDER_ID - 1]; // 0-indexed access to array
        } else {
             Logger.log(`deleteProject: Could not retrieve row data for project ${projectId} at row ${rowIndex}, though index was found.`);
             // Decide whether to proceed with sheet deletion only or fail
             // Let's proceed cautiously and require rowData for folder deletion
             return { success: false, error: "Failed to retrieve project details for deletion." };
        }
        // 2. Delete the project's row from the "ProjectIndex" sheet
        deleteSheetRow(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
        Logger.log(`deleteProject: Row ${rowIndex} for project ${projectId} deleted from ProjectIndex sheet.`);
    } else {
         Logger.log(`deleteProject: Project with ID "${projectId}" not found in ProjectIndex. No row to delete.`);
         // If not in sheet, maybe still try to delete folder if we implement orphan cleanup later?
         // For now, consider it "deleted" from the user's perspective if not in index.
         return { success: true, message: "Project not found in index, assumed already deleted.", deletedProjectId: projectId };
    }


    // 3. Delete the project's folder from Google Drive
    if (projectFolderId) {
      try {
        deleteDriveFolderRecursive(projectFolderId); // From DriveService.gs
        Logger.log(`deleteProject: Project folder ${projectFolderId} for project ${projectId} and its contents have been trashed.`);
      } catch (driveError) {
        Logger.log(`deleteProject: Error while deleting project folder ${projectFolderId} for project ${projectId}. Error: ${driveError.toString()}. Sheet entry was removed.`);
        return { 
            success: true, // Success because removed from list
            message: `Project removed from index. Warning: Error deleting Drive folder: ${driveError.message}`,
            deletedProjectId: projectId 
        };
      }
    } else {
      Logger.log(`deleteProject: No ProjectFolderID found for project ${projectId} (row ${rowIndex}). Cannot delete Drive folder.`);
    }

    return { success: true, message: "Project successfully deleted.", deletedProjectId: projectId };

  } catch (e) {
    Logger.log(`Error in deleteProject for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
    return { success: false, error: `Failed to delete project: ${e.message}` };
  }
}

  /**
   * Retrieves the project's data JSON string from its file in Google Drive.
   * @param {string} projectId The ID of the project.
   * @return {string|null} The JSON string content of the project data file, 
   *                      or null if the project or file is not found or an error occurs.
   */
  function getProjectDataForEditing(projectId) {
    try {
      Logger.log(`getProjectDataForEditing: Attempting to load data for projectId: ${projectId}`);
      if (!projectId) {
        Logger.log("getProjectDataForEditing: ProjectID is missing.");
        return null; 
      }

      const allProjects = getAllSheetData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME);
      let projectEntry = null;
      if (allProjects && Array.isArray(allProjects)) {
        projectEntry = allProjects.find(p => (p['ProjectID'] || p[COL_PROJECT_ID - 1]) === projectId);
      }

      if (!projectEntry) {
        Logger.log(`getProjectDataForEditing: Project with ID "${projectId}" not found in ProjectIndex.`);
        return null; 
      }

      const projectDataFileId = projectEntry['ProjectDataFileID'] || projectEntry[COL_PROJECT_DATA_FILE_ID - 1];

      if (!projectDataFileId) {
        Logger.log(`getProjectDataForEditing: ProjectDataFileID is missing for project "${projectId}".`);
        return null; 
      }
      Logger.log(`getProjectDataForEditing: Found ProjectDataFileID: ${projectDataFileId} for project ${projectId}.`);

      const jsonContent = readDriveFileContent(projectDataFileId);

      Logger.log(`getProjectDataForEditing: Successfully read content for file ID ${projectDataFileId}. Content length: ${jsonContent ? jsonContent.length : 0}`);
      return jsonContent; 

    } catch (e) {
      Logger.log(`Error in getProjectDataForEditing for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
      return null; 
    }
  }

/**
 * Retrieves an audio file from Drive as a base64 data URI.
 * @param {string} driveFileId The ID of the audio file in Google Drive.
 * @return {object} An object like { success: true, base64Data: 'data:audio/mpeg;base64,...', mimeType: 'audio/mpeg' } or { success: false, error: '...' }.
 */
function getAudioAsBase64(driveFileId) {
  try {
    if (!driveFileId) {
      return { success: false, error: "Drive File ID is required." };
    }
    const file = DriveApp.getFileById(driveFileId);
    const blob = file.getBlob();
    const mimeType = blob.getContentType();
    
    // Check if it's an audio type (basic check)
    if (!mimeType || !mimeType.startsWith('audio/')) {
       Logger.log(`getAudioAsBase64: File ID ${driveFileId} is not audio (MIME: ${mimeType}).`);
       return { success: false, error: `File is not audio (type: ${mimeType})` };
    }

    const base64Data = Utilities.base64Encode(blob.getBytes());
    const dataURI = 'data:' + mimeType + ';base64,' + base64Data;
    
    Logger.log(`getAudioAsBase64: Successfully retrieved and encoded audio file ID: ${driveFileId}. MimeType: ${mimeType}. DataURI length: ${dataURI.length}`);
    return { success: true, base64Data: dataURI, mimeType: mimeType };

  } catch (e) {
      if (e.message.toLowerCase().includes("access denied") || e.message.toLowerCase().includes("not found")) {
          Logger.log(`getAudioAsBase64: File not found or access denied for ID ${driveFileId}. Error: ${e.toString()}`);
          return { success: false, error: "Audio file not found or access denied." };
      }
      Logger.log(`Error in getAudioAsBase64 for file ID ${driveFileId}: ${e.toString()} \nStack: ${e.stack}`);
      return { success: false, error: `Failed to retrieve audio as base64: ${e.message}` };
  }
}