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
    const nowISO = new Date().toISOString();
    const initialJsonContent = {
      projectId: projectId,
      title: projectTitle,
      status: "Draft", // Default status for new projects
      slides: [],
      createdDate: nowISO,
      lastModified: nowISO
    };
    const projectDataFileId = saveJsonToDriveFile(
      PROJECT_DATA_FILENAME, 
      JSON.stringify(initialJsonContent, null, 2), 
      projectFolderId,
      null 
    );
    if (!projectDataFileId) {
        Logger.log(`createProject: Failed to create project data file for project ${projectId}. saveJsonToDriveFile returned falsy.`);
        // Consider cleanup: delete folder if file creation fails? Maybe later enhancement.
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
    // Rely on getAllSheetData to handle headers correctly
    const projectsData = getAllSheetData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME);

    if (!projectsData || !Array.isArray(projectsData)) {
      Logger.log("getAllProjectsForAdmin: No data or invalid data returned from getAllSheetData.");
      return []; 
    }
    
    // Check if it returned objects (processed headers) or arrays
    if (projectsData.length > 0 && typeof projectsData[0] === 'object' && projectsData[0] !== null) {
        // Data is already array of objects
        const formattedProjects = projectsData.map(rowObject => {
          // Basic validation within map
          if (!rowObject || !rowObject['ProjectID']) return null;
          return {
            projectId: rowObject['ProjectID'], 
            projectTitle: rowObject['ProjectTitle'] || '(No Title)',
            status: rowObject['Status'] || 'Unknown',
            projectFolderId: rowObject['ProjectFolderID'] || null, 
            projectDataFileId: rowObject['ProjectDataFileID'] || null
          };
        }).filter(p => p !== null); // Remove null entries
        Logger.log(`getAllProjectsForAdmin: Retrieved ${formattedProjects.length} projects (from objects).`);
        return formattedProjects;

    } else if (projectsData.length > 0) {
        // Data is likely array of arrays (or empty)
        // Assume standard column order based on constants
        const headers = ['ProjectID', 'ProjectTitle', 'ProjectFolderID', 'Status', 'ProjectDataFileID', 'LastModified', 'CreatedDate']; // Define expected order
        let dataStartIndex = 0;
        // Simple check if first row looks like *our* headers
        if (typeof projectsData[0][COL_PROJECT_ID-1] === 'string' && projectsData[0][COL_PROJECT_ID-1].toLowerCase().includes('id')) {
             dataStartIndex = 1; // Skip header row if present
        }

        const formattedProjects = [];
        for (let i = dataStartIndex; i < projectsData.length; i++) {
             const rowArray = projectsData[i];
             if (!rowArray || !rowArray[COL_PROJECT_ID - 1]) continue; // Skip if no project ID in the expected column
             formattedProjects.push({
                 projectId: rowArray[COL_PROJECT_ID - 1],
                 projectTitle: rowArray[COL_PROJECT_TITLE - 1] || '(No Title)',
                 status: rowArray[COL_STATUS - 1] || 'Unknown',
                 projectFolderId: rowArray[COL_PROJECT_FOLDER_ID - 1] || null,
                 projectDataFileId: rowArray[COL_PROJECT_DATA_FILE_ID - 1] || null
             });
        }
         Logger.log(`getAllProjectsForAdmin: Retrieved ${formattedProjects.length} projects (from arrays).`);
         return formattedProjects;
    } else {
        // Empty sheet
        Logger.log(`getAllProjectsForAdmin: Sheet appears to be empty.`);
        return [];
    }

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
   try {
        const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
        if (rowIndex) {
            const rowData = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
            if (rowData && rowData[COL_PROJECT_FOLDER_ID - 1]) {
                Logger.log(`getProjectFolderIdFromSheet: Found folderId "${rowData[COL_PROJECT_FOLDER_ID - 1]}" for projectId "${projectId}".`);
                return rowData[COL_PROJECT_FOLDER_ID - 1];
            }
        }
        Logger.log(`getProjectFolderIdFromSheet: ProjectId "${projectId}" or its folder ID not found in sheet.`);
        return null;
    } catch (e) {
        Logger.log(`Error in getProjectFolderIdFromSheet for ${projectId}: ${e.toString()}`);
        return null;
    }
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

    // Fallback if webContentLink is null or unsuitable (especially for audio)
     // Note: For audio, this link might *still* not work directly in <audio src>. Base64 is more reliable.
    if (!webContentLink && mediaType !== 'audio') { 
        Logger.log(`uploadFileToDrive: webContentLink is null or wasn't obtained. Using fallback for file ID: ${driveFile.getId()}`);
        if (mediaType === 'image') {
          webContentLink = 'https://drive.google.com/uc?export=view&id=' + driveFile.getId();
        } else {
           const downloadUrl = driveFile.getDownloadUrl(); 
           webContentLink = downloadUrl ? downloadUrl.replace("&export=download", "&export=view") : 'https://drive.google.com/uc?id=' + driveFile.getId();
        }
    } else if (!webContentLink && mediaType === 'audio') {
         Logger.log(`uploadFileToDrive: webContentLink is null for audio. Relying on client fetching base64 via ID.`);
         // Set webContentLink to null explicitly so client knows to fetch base64
         webContentLink = null; 
    }

    Logger.log(`uploadFileToDrive: File uploaded successfully. ID: ${driveFile.getId()}, Final Link: ${webContentLink}`);
    return {
      success: true,
      driveFileId: driveFile.getId(),
      webContentLink: webContentLink, // May be null for audio if direct link failed
      fileName: driveFile.getName(),
      mimeType: driveFile.getMimeType() 
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
 * Updates the LastModified timestamp and Status in the ProjectIndex sheet.
 * **Simplified Logic:** Uses findRowIndexByValue and getSheetRowData.
 * @param {string} projectId The ID of the project.
 * @param {string} projectDataJSON A JSON string representing the entire project data.
 * @return {object} An object indicating success or failure.
 */
function saveProjectData(projectId, projectDataJSON) {
  try {
    Logger.log(`saveProjectData (Simplified): Attempting to save data for projectId: ${projectId}`);
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
    
    const currentStatusInJson = projectDataParsed.status || "Draft"; // Default if status missing in JSON
    const nowISO = new Date().toISOString(); // Timestamp for modifications

    // 1. Find the row index for the project
    const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
    if (!rowIndex) {
      Logger.log(`saveProjectData: Project with ID "${projectId}" not found in ProjectIndex.`);
      return { success: false, error: `Project metadata not found for ID: ${projectId}. Cannot save.` };
    }

    // 2. Get current row data to find FolderID and FileID
    const rowDataArray = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
     if (!rowDataArray) {
        Logger.log(`saveProjectData: Could not retrieve row data for project ${projectId} at row ${rowIndex}.`);
        return { success: false, error: "Failed to retrieve project details for saving." };
     }
     
    const projectFolderId = rowDataArray[COL_PROJECT_FOLDER_ID - 1];
    const projectDataFileIdToOverwrite = rowDataArray[COL_PROJECT_DATA_FILE_ID - 1];

    if (!projectFolderId || !projectDataFileIdToOverwrite) {
      Logger.log(`saveProjectData: Missing FolderID or DataFileID for project "${projectId}" in row ${rowIndex}. FolderID: ${projectFolderId}, FileID: ${projectDataFileIdToOverwrite}`);
      return { success: false, error: "Project folder or data file reference is missing in index sheet. Cannot save." };
    }
    Logger.log(`saveProjectData: Found FolderID: ${projectFolderId}, DataFileID to overwrite: ${projectDataFileIdToOverwrite} for project ${projectId}.`);

    // 3. Update lastModified in the JSON object itself before saving
    projectDataParsed.lastModified = nowISO; 
    const updatedJsonString = JSON.stringify(projectDataParsed, null, 2);

    // 4. Save/Overwrite the JSON file in Drive
    const savedFileId = saveJsonToDriveFile(
      PROJECT_DATA_FILENAME,
      updatedJsonString, // Use updated string
      projectFolderId,
      projectDataFileIdToOverwrite 
    );

    if (!savedFileId) {
      Logger.log(`saveProjectData: Failed to save JSON file to Drive for project ${projectId}. saveJsonToDriveFile returned falsy.`);
      return { success: false, error: "Failed to save project data to Google Drive." };
    }
    // Optional: Check if savedFileId matches projectDataFileIdToOverwrite. If not, maybe update the sheet? For now, assume overwrite works.
    Logger.log(`saveProjectData: JSON data saved to file ID: ${savedFileId} for project ${projectId}.`);

    // 5. Update the Sheet Row (Status and LastModified)
    rowDataArray[COL_LAST_MODIFIED - 1] = nowISO; 
    rowDataArray[COL_STATUS - 1] = currentStatusInJson; // Sync status from JSON
    updateSheetRow(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex, rowDataArray);
    Logger.log(`saveProjectData: Updated Sheet - LastModified and Status ("${currentStatusInJson}") for project ${projectId} at row ${rowIndex}.`);
   
    return { success: true, message: "Project data saved successfully." };

  } catch (e) {
    Logger.log(`Error in saveProjectData for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
    return { success: false, error: `Failed to save project data: ${e.message}` };
  }
}

/**
 * Updates the status of a project in the "ProjectIndex" sheet.
 * Also updates the "LastModified" timestamp in the sheet and the project JSON file.
 * **Simplified Logic:** Uses findRowIndexByValue and getSheetRowData.
 * @param {string} projectId The ID of the project to update.
 * @param {string} newStatus The new status (e.g., "Active", "Inactive", "Draft").
 * @return {object} An object indicating success or failure.
 */
function updateProjectStatus(projectId, newStatus) {
  try {
    Logger.log(`updateProjectStatus (Simplified): Attempting for projectId: ${projectId} to "${newStatus}"`);
    if (!projectId || !newStatus) {
      Logger.log("updateProjectStatus: Missing projectId or newStatus.");
      return { success: false, error: "Project ID and new status are required." };
    }

    const validStatuses = ["Draft", "Active", "Inactive"];
    if (validStatuses.indexOf(newStatus) === -1) {
        Logger.log(`updateProjectStatus: Invalid status value "${newStatus}" for projectId: ${projectId}`);
        return { success: false, error: `Invalid status value: ${newStatus}. Must be one of ${validStatuses.join(', ')}.` };
    }

    // 1. Find the row index
    const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
    if (!rowIndex) {
      Logger.log(`updateProjectStatus: Project with ID "${projectId}" not found in ProjectIndex.`);
      return { success: false, error: `Project not found for ID: ${projectId}. Cannot update status.` };
    }

    // 2. Get current row data
    let rowData = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
    if (!rowData) {
        Logger.log(`updateProjectStatus: Could not retrieve current row data for project ${projectId} at row ${rowIndex}.`);
        return { success: false, error: "Could not retrieve project data to update status."};
    }
    
    const nowISO = new Date().toISOString();

    // 3. Update Sheet Row Data
    rowData[COL_STATUS - 1] = newStatus;
    rowData[COL_LAST_MODIFIED - 1] = nowISO;
    updateSheetRow(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex, rowData);
    Logger.log(`updateProjectStatus: Sheet status for project ${projectId} (row ${rowIndex}) updated to "${newStatus}". LastModified also updated.`);

    // 4. Update status & lastModified in project_data.json
    const projectFolderId = rowData[COL_PROJECT_FOLDER_ID - 1];
    const projectDataFileId = rowData[COL_PROJECT_DATA_FILE_ID - 1];

    if(projectDataFileId && projectFolderId){
        try {
            let projectJsonContent = readDriveFileContent(projectDataFileId);
            let projectJson = JSON.parse(projectJsonContent);
            projectJson.status = newStatus; // Update status in JSON
            projectJson.lastModified = nowISO; // Update lastModified in JSON
            saveJsonToDriveFile(PROJECT_DATA_FILENAME, JSON.stringify(projectJson, null, 2), projectFolderId, projectDataFileId);
            Logger.log(`updateProjectStatus: Status and lastModified in project_data.json for ${projectId} also updated.`);
        } catch (e) {
            // Log error but don't fail the whole operation if JSON update fails
            Logger.log(`updateProjectStatus: WARNING - Failed to update status/lastModified in project_data.json for ${projectId}. Error: ${e.toString()}`);
        }
    } else {
        Logger.log(`updateProjectStatus: WARNING - Could not find folderId or dataFileId for project ${projectId} from sheet to update JSON. FolderID: ${projectFolderId}, FileID: ${projectDataFileId}`);
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
             return { success: false, error: "Failed to retrieve project details for deletion." };
        }
        // 2. Delete the project's row from the "ProjectIndex" sheet
        deleteSheetRow(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
        Logger.log(`deleteProject: Row ${rowIndex} for project ${projectId} deleted from ProjectIndex sheet.`);
    } else {
         Logger.log(`deleteProject: Project with ID "${projectId}" not found in ProjectIndex. No row to delete.`);
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

      // Find row index first
       const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
       if (!rowIndex) {
            Logger.log(`getProjectDataForEditing: Project with ID "${projectId}" not found in ProjectIndex.`);
            return null; 
       }
        
       // Get row data to find the file ID
       const rowData = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
       if (!rowData) {
            Logger.log(`getProjectDataForEditing: Could not retrieve row data for project ${projectId} at row ${rowIndex}.`);
            return null;
       }

      const projectDataFileId = rowData[COL_PROJECT_DATA_FILE_ID - 1];

      if (!projectDataFileId) {
        Logger.log(`getProjectDataForEditing: ProjectDataFileID is missing for project "${projectId}" in row ${rowIndex}.`);
        return null; 
      }
      Logger.log(`getProjectDataForEditing: Found ProjectDataFileID: ${projectDataFileId} for project ${projectId}.`);

      // Read content using the found file ID
      const jsonContent = readDriveFileContent(projectDataFileId);

      Logger.log(`getProjectDataForEditing: Successfully read content for file ID ${projectDataFileId}. Content length: ${jsonContent ? jsonContent.length : 0}`);
      return jsonContent; 

    } catch (e) {
      Logger.log(`Error in getProjectDataForEditing for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
      // Consider returning specific errors if needed, e.g., differentiate file read error from not found
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

/**
 * Copies a file from Google Drive (source) to the specified project's folder.
 * Sets sharing permissions on the copy to "anyone with link can view".
 * @param {string} sourceFileId The ID of the file to copy from Drive.
 * @param {string} projectId The ID of the project to associate the copied file with.
 * @param {string} mediaType 'image' or 'audio', used for context and potential view settings.
 * @param {string} originalFileName The original name of the file from the picker (optional, helps if source name is weird)
 * @param {string} originalMimeType The original mime type from picker (optional)
 * @return {object} An object like { success: true, driveFileId: '...', fileName: '...', mimeType: '...' } or { success: false, error: '...' }.
 */
function copyDriveFileToProject(sourceFileId, projectId, mediaType, originalFileName, originalMimeType) {
  try {
    Logger.log(`copyDriveFileToProject: Starting copy for sourceFileId: ${sourceFileId}, projectId: ${projectId}, mediaType: ${mediaType}`);

    if (!sourceFileId || !projectId || !mediaType) {
      Logger.log("copyDriveFileToProject: Missing required parameters.");
      return { success: false, error: "Source File ID, Project ID, and Media Type are required." };
    }

    const projectFolderId = getProjectFolderIdFromSheet(projectId);
    if (!projectFolderId) {
      Logger.log(`copyDriveFileToProject: Could not find ProjectFolderID for projectId "${projectId}".`);
      return { success: false, error: `Project folder not found for project ID: ${projectId}.` };
    }
    const projectFolder = DriveApp.getFolderById(projectFolderId);

    const sourceFile = DriveApp.getFileById(sourceFileId);
    const fileNameToUse = originalFileName || sourceFile.getName(); // Prefer Picker name if available
    const mimeTypeToUse = originalMimeType || sourceFile.getMimeType();

    // Check if the file type from picker matches the expected mediaType context
    if (mediaType === 'image' && !mimeTypeToUse.startsWith('image/')) {
        return { success: false, error: `Selected file "${fileNameToUse}" is not an image (type: ${mimeTypeToUse}). Please select an image file.`};
    }
    if (mediaType === 'audio' && !mimeTypeToUse.startsWith('audio/')) {
        return { success: false, error: `Selected file "${fileNameToUse}" is not an audio file (type: ${mimeTypeToUse}). Please select an audio file.`};
    }
    
    // Create a copy in the project folder
    const copiedFile = sourceFile.makeCopy(fileNameToUse, projectFolder);
    if (!copiedFile || !copiedFile.getId()) {
      Logger.log(`copyDriveFileToProject: Failed to copy file ${sourceFileId} to project folder ${projectFolderId}.`);
      return { success: false, error: "Failed to copy file to project folder." };
    }
    
    // Set sharing permissions on the copy
    copiedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log(`copyDriveFileToProject: File "${copiedFile.getName()}" (ID: ${copiedFile.getId()}) copied to project and shared as ANYONE_WITH_LINK.`);
    
    // Get webContentLink (not strictly needed if client fetches base64, but good to have)
    let webContentLink = null;
    try { webContentLink = copiedFile.getWebContentLink(); } catch(e) { Logger.log("Could not get webContentLink for copied file: " + e.toString());}
    
    return {
      success: true,
      driveFileId: copiedFile.getId(),
      fileName: copiedFile.getName(),
      mimeType: copiedFile.getMimeType(),
      webContentLink: webContentLink // Can be null
    };

  } catch (e) {
    Logger.log(`Error in copyDriveFileToProject: ${e.toString()} - SourceFileID: ${sourceFileId}, ProjectID: ${projectId} \nStack: ${e.stack}`);
    return { success: false, error: `Failed to process file from Drive: ${e.message}` };
  }
}