// Server/AdminController.gs

/**
 * Creates a standardized response object.
 * @param {boolean} success - Indicates whether the operation was successful.
 * @param {Object} [data=null] - The data to return on success.
 * @param {string} [error=null] - The error message to return on failure.
 * @return {Object} The response object.
 */
function createResponse(success, data = null, error = null) {
  if (success) {
    return { success: true, data: data };
  }
  return { success: false, error: error || "An unknown error occurred." };
}

/**
 * Creates a new project:
 * - Generates a unique ProjectID.
 * - Creates a new folder in Google Drive for the new project.
 * - Creates an initial project_data.json file in the new folder.
 * - Adds a new row to the "ProjectIndex" sheet.
 * @param {string} projectTitle The title of the project.
 * @return {object} An object indicating success or failure, and new project details.
 */
function createProject(projectTitle) {
  try {
    if (!projectTitle || projectTitle.trim() === "") {
      Logger.log("createProject: Project title was empty.");
      return createResponse(false, null, "Project title cannot be empty.");
    }

    const projectId = Utilities.getUuid();
    Logger.log(`createProject: Starting for title "${projectTitle}", generated ID: ${projectId}`);

    // 1. Create a new folder in Google Drive for the project
    const projectFolderName = `${projectTitle} [${projectId.substring(0, 8)}]`;
    const projectFolderId = createDriveFolder(projectFolderName, ROOT_PROJECT_FOLDER_ID);
    if (!projectFolderId) {
        Logger.log(`createProject: Failed to create project folder for "${projectFolderName}". createDriveFolder returned falsy.`);
        return createResponse(false, null, "Failed to create project folder in Drive.");
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
        return createResponse(false, null, "Failed to create project data file in Drive.");
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

    return createResponse(true, {
      message: "Project created successfully!",
      projectId: projectId,
      projectFolderId: projectFolderId,
      projectDataFileId: projectDataFileId,
      status: initialJsonContent.status
    });

  } catch (e) {
    Logger.log(`Error in createProject: ${e.toString()} \nStack: ${e.stack ? e.stack : 'No stack available'}`);
    return createResponse(false, null, `Failed to create project: ${e.message}`);
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
      return createResponse(true, { projects: [] }); // Return empty list via response object
    }

    // getAllSheetData now always returns an array of objects or an empty array.
    if (projectsData.length === 0) {
        Logger.log("getAllProjectsForAdmin: No projects found or sheet is empty/header-only.");
        return createResponse(true, { projects: [] }); // Return empty list via response object
    }

    // Data is an array of objects
    const formattedProjects = projectsData.map(rowObject => {
      // Basic validation within map
      if (!rowObject || !rowObject['ProjectID']) { // Check for ProjectID specifically
        Logger.log("getAllProjectsForAdmin: Skipping a row object due to missing ProjectID: " + JSON.stringify(rowObject));
        return null;
      }
      return {
        projectId: rowObject['ProjectID'],
        projectTitle: rowObject['ProjectTitle'] || '(No Title)',
        status: rowObject['Status'] || 'Unknown',
        projectFolderId: rowObject['ProjectFolderID'] || null,
        projectDataFileId: rowObject['ProjectDataFileID'] || null
      };
    }).filter(p => p !== null); // Remove null entries that failed validation

    Logger.log(`getAllProjectsForAdmin: Retrieved and formatted ${formattedProjects.length} projects.`);
    return createResponse(true, { projects: formattedProjects });

  } catch (e) {
    Logger.log(`Error in getAllProjectsForAdmin: ${e.toString()} \nStack: ${e.stack ? e.stack : 'No stack available'}`);
    return createResponse(false, null, `Error retrieving projects: ${e.message}`);
  }
}

/**
 * Uploads a file (image, audio, etc.) to the specified project's folder in Google Drive.
 * This version is more robust and self-contained.
 * @param {Object} fileData An object containing { fileName, mimeType, data (base64 string) }.
 * @param {string} projectId The ID of the project to associate the file with.
 * @param {string} mediaType A string descriptor like 'image' or 'audio' (for logging/future use).
 * @return {Object} An object like { success: true, ... } or { success: false, ... }.
 */
function uploadFileToDrive(fileData, projectId, mediaType) {
  try {
    Logger.log(`uploadFileToDrive: Starting upload for projectId: ${projectId}, mediaType: ${mediaType}, fileName: ${fileData.fileName}`);

    if (!fileData || !fileData.data || !fileData.fileName || !fileData.mimeType) {
      throw new Error("Invalid file data provided.");
    }
    if (!projectId) {
      throw new Error("Project ID is required for upload.");
    }

    // Find the project row to get the folder ID directly.
    const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
    if (!rowIndex) {
      throw new Error(`Project with ID "${projectId}" not found in ProjectIndex. Cannot upload file.`);
    }
    
    const rowData = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
    const projectFolderId = rowData ? rowData[COL_PROJECT_FOLDER_ID - 1] : null;

    if (!projectFolderId) {
      throw new Error(`Could not find ProjectFolderID for project "${projectId}".`);
    }
    
    Logger.log(`Found folder ID: ${projectFolderId} for project ${projectId}.`);

    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);

    const driveFile = createFileInDriveFromBlob(blob, fileData.fileName, projectFolderId);

    if (!driveFile || typeof driveFile.getId !== 'function') {
        throw new Error("Failed to create file in Google Drive; received invalid response from service.");
    }

    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Use the direct web view link for images. This is more reliable.
    const webContentLink = `https://drive.google.com/uc?export=view&id=${driveFile.getId()}`;

    Logger.log(`uploadFileToDrive: File uploaded successfully. ID: ${driveFile.getId()}, Link: ${webContentLink}`);
    return createResponse(true, {
      driveFileId: driveFile.getId(),
      webContentLink: webContentLink,
      fileName: driveFile.getName(),
      mimeType: driveFile.getMimeType()
    });

  } catch (e) {
    Logger.log(`Error in uploadFileToDrive: ${e.toString()} \nStack: ${e.stack}`);
    // Use createResponse to return a standardized error object
    return createResponse(false, null, `Failed to upload file: ${e.message}`);
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
      return createResponse(false, null, "Drive File ID is required.");
    }
    const file = DriveApp.getFileById(driveFileId);
    const blob = file.getBlob();
    const mimeType = blob.getContentType();
    if (!mimeType || !mimeType.startsWith('image/')) {
       Logger.log(`getImageAsBase64: File ID ${driveFileId} is not an image (MIME: ${mimeType}).`);
       return createResponse(false, null, `File is not an image (type: ${mimeType})`);
    }
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const dataURI = 'data:' + mimeType + ';base64,' + base64Data;

    Logger.log(`getImageAsBase64: Successfully retrieved and encoded file ID: ${driveFileId}. MimeType: ${mimeType}. DataURI length: ${dataURI.length}`);
    return createResponse(true, { base64Data: dataURI, mimeType: mimeType });

  } catch (e) {
     if (e.message.toLowerCase().includes("access denied") || e.message.toLowerCase().includes("not found")) {
          Logger.log(`getImageAsBase64: File not found or access denied for ID ${driveFileId}. Error: ${e.toString()}`);
          return createResponse(false, null, "Image file not found or access denied.");
      }
    Logger.log(`Error in getImageAsBase64 for file ID ${driveFileId}: ${e.toString()} \nStack: ${e.stack}`);
    return createResponse(false, null, `Failed to retrieve image as base64: ${e.message}`);
  }
}

/**
 * Saves the project's data (slides, overlays, media links, canvas dimensions) to its JSON file in Drive.
 * Updates the LastModified timestamp and Status in the ProjectIndex sheet.
 * @param {string} projectId The ID of the project.
 * @param {string} projectDataJSON A JSON string representing the entire project data.
 * @return {object} An object indicating success or failure.
 */
function saveProjectData(projectId, projectDataJSON) {
  try {
    Logger.log(`saveProjectData: Attempting to save data for projectId: ${projectId}`);
    if (!projectId || !projectDataJSON) {
      Logger.log("saveProjectData: Missing projectId or projectDataJSON.");
      return createResponse(false, null, "Project ID and data are required.");
    }

    let projectDataParsed;
    try {
        projectDataParsed = JSON.parse(projectDataJSON);
    } catch (parseError) {
        Logger.log(`saveProjectData: Error parsing projectDataJSON for projectId ${projectId}: ${parseError.toString()}`);
        return createResponse(false, null, "Invalid project data format. Could not parse JSON.");
    }

    const currentStatusInJson = projectDataParsed.status || "Draft"; // Default if status missing in JSON
    const nowISO = new Date().toISOString(); // Timestamp for modifications

    // Find the row index for the project
    const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
    if (!rowIndex) {
      Logger.log(`saveProjectData: Project with ID "${projectId}" not found in ProjectIndex.`);
      return createResponse(false, null, `Project metadata not found for ID: ${projectId}. Cannot save.`);
    }

    const rowDataArray = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
     if (!rowDataArray) {
        Logger.log(`saveProjectData: Could not retrieve row data for project ${projectId} at row ${rowIndex}.`);
        return createResponse(false, null, "Failed to retrieve project details for saving.");
     }

    const projectFolderId = rowDataArray[COL_PROJECT_FOLDER_ID - 1];
    const projectDataFileIdToOverwrite = rowDataArray[COL_PROJECT_DATA_FILE_ID - 1];

    if (!projectFolderId || !projectDataFileIdToOverwrite) {
      Logger.log(`saveProjectData: Missing FolderID or DataFileID for project "${projectId}" in row ${rowIndex}.`);
      return createResponse(false, null, "Project folder or data file reference is missing in index sheet. Cannot save.");
    }

    projectDataParsed.lastModified = nowISO;
    const updatedJsonString = JSON.stringify(projectDataParsed, null, 2);

    let savedFileId = saveJsonToDriveFile(
      PROJECT_DATA_FILENAME,
      updatedJsonString,
      projectFolderId,
      projectDataFileIdToOverwrite
    );

    if (!savedFileId) {
      throw new Error("saveJsonToDriveFile did not return a valid fileId.");
    }
    
    rowDataArray[COL_LAST_MODIFIED - 1] = nowISO;
    rowDataArray[COL_STATUS - 1] = currentStatusInJson;
    updateSheetRow(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex, rowDataArray);
    
    return createResponse(true, { message: "Project data saved successfully." });

  } catch (e) {
    Logger.log(`Error in saveProjectData for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
    return createResponse(false, null, `Failed to save project data: ${e.message}`);
  }
}


/**
 * Updates the status of a project in the "ProjectIndex" sheet.
 * Also updates the "LastModified" timestamp in the sheet and the project JSON file.
 * @param {string} projectId The ID of the project to update.
 * @param {string} newStatus The new status (e.g., "Active", "Inactive", "Draft").
 * @return {object} An object indicating success or failure.
 */
function updateProjectStatus(projectId, newStatus) {
  try {
    Logger.log(`updateProjectStatus (Hardened): Attempting for projectId: ${projectId} to "${newStatus}"`);
    if (!projectId || !newStatus) {
      throw new Error("Project ID and new status are required.");
    }

    const validStatuses = ["Draft", "Active", "Inactive"];
    if (validStatuses.indexOf(newStatus) === -1) {
        throw new Error(`Invalid status value: ${newStatus}. Must be one of ${validStatuses.join(', ')}.`);
    }

    const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
    if (!rowIndex) {
      throw new Error(`Project not found for ID: ${projectId}. Cannot update status.`);
    }

    const rowData = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
    if (!rowData) {
        throw new Error("Could not retrieve project data to update status.");
    }

    const projectFolderId = rowData[COL_PROJECT_FOLDER_ID - 1];
    const projectDataFileId = rowData[COL_PROJECT_DATA_FILE_ID - 1];

    if (!projectDataFileId || !projectFolderId) {
      throw new Error("Project data file reference or folder is missing from the index. Cannot update status.");
    }
    
    const nowISO = new Date().toISOString();
    let projectJsonContent = readDriveFileContent(projectDataFileId);
    let projectJson = JSON.parse(projectJsonContent);
    
    projectJson.status = newStatus;
    projectJson.lastModified = nowISO;
    const updatedJsonString = JSON.stringify(projectJson, null, 2);
    
    saveJsonToDriveFile(PROJECT_DATA_FILENAME, updatedJsonString, projectFolderId, projectDataFileId);
    
    rowData[COL_STATUS - 1] = newStatus;
    rowData[COL_LAST_MODIFIED - 1] = nowISO;
    updateSheetRow(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex, rowData);
    
    return createResponse(true, { message: `Project status successfully updated to ${newStatus}.` });

  } catch (e) {
    Logger.log(`Error in updateProjectStatus for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack ? e.stack : 'No stack available'}`);
    return createResponse(false, null, `Failed to update project status: ${e.message}`);
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
      return createResponse(false, null, "Project ID is required for deletion.");
    }

    const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
    let projectFolderId = null;

    if (rowIndex) {
        const rowData = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
        if (rowData) {
            projectFolderId = rowData[COL_PROJECT_FOLDER_ID - 1];
        } else {
             Logger.log(`deleteProject: Could not retrieve row data for project ${projectId} at row ${rowIndex}, though index was found.`);
             return createResponse(false, null, "Failed to retrieve project details for deletion.");
        }
        deleteSheetRow(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
        Logger.log(`deleteProject: Row ${rowIndex} for project ${projectId} deleted from ProjectIndex sheet.`);
    } else {
         Logger.log(`deleteProject: Project with ID "${projectId}" not found in ProjectIndex. No row to delete.`);
         return createResponse(true, { message: "Project not found in index, assumed already deleted.", deletedProjectId: projectId });
    }

    if (projectFolderId) {
      try {
        deleteDriveFolderRecursive(projectFolderId);
        Logger.log(`deleteProject: Project folder ${projectFolderId} for project ${projectId} and its contents have been trashed.`);
      } catch (driveError) {
        Logger.log(`deleteProject: Error while deleting project folder ${projectFolderId} for project ${projectId}. Error: ${driveError.toString()}. Sheet entry was removed.`);
        return createResponse(true, {
            message: `Project removed from index. Warning: Error deleting Drive folder: ${driveError.message}`,
            deletedProjectId: projectId
        });
      }
    } else {
      Logger.log(`deleteProject: No ProjectFolderID found for project ${projectId} (row ${rowIndex}). Cannot delete Drive folder.`);
    }

    return createResponse(true, { message: "Project successfully deleted.", deletedProjectId: projectId });

  } catch (e) {
    Logger.log(`Error in deleteProject for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
    return createResponse(false, null, `Failed to delete project: ${e.message}`);
  }
}


  /**
   * Retrieves the project's data JSON string from its file in Google Drive.
   * @param {string} projectId The ID of the project.
   * @return {string} The JSON string content of the project data file.
   */
  function getProjectDataForEditing(projectId) {
    try {
      Logger.log(`getProjectDataForEditing: Attempting to load data for projectId: ${projectId}`);
      if (!projectId) {
        return createResponse(false, null, "ProjectID is missing.");
      }

      const rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectId);
      if (!rowIndex) {
          return createResponse(false, null, `Project index entry not found for ID: ${projectId}.`);
      }

      const rowData = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
      if (!rowData) {
          return createResponse(false, null, "Could not retrieve project metadata from sheet.");
      }

      const projectDataFileId = rowData[COL_PROJECT_DATA_FILE_ID - 1];

      if (!projectDataFileId) {
        return createResponse(false, null, "Project data file reference missing in index.");
      }
      
      const jsonContent = readDriveFileContent(projectDataFileId);

      return createResponse(true, { projectDataJSON: jsonContent });

    } catch (e) {
      Logger.log(`Error in getProjectDataForEditing for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
      return createResponse(false, null, `Failed to get project data for editing: ${e.message}`);
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
      return createResponse(false, null, "Drive File ID is required.");
    }
    const file = DriveApp.getFileById(driveFileId);
    const blob = file.getBlob();
    const mimeType = blob.getContentType();

    if (!mimeType || !mimeType.startsWith('audio/')) {
       Logger.log(`getAudioAsBase64: File ID ${driveFileId} is not audio (MIME: ${mimeType}).`);
       return createResponse(false, null, `File is not audio (type: ${mimeType})`);
    }

    const base64Data = Utilities.base64Encode(blob.getBytes());
    const dataURI = 'data:' + mimeType + ';base64,' + base64Data;

    Logger.log(`getAudioAsBase64: Successfully retrieved and encoded audio file ID: ${driveFileId}. MimeType: ${mimeType}.`);
    return createResponse(true, { base64Data: dataURI, mimeType: mimeType });

  } catch (e) {
      if (e.message.toLowerCase().includes("access denied") || e.message.toLowerCase().includes("not found")) {
          Logger.log(`getAudioAsBase64: File not found or access denied for ID ${driveFileId}. Error: ${e.toString()}`);
          return createResponse(false, null, "Audio file not found or access denied.");
      }
      Logger.log(`Error in getAudioAsBase64 for file ID ${driveFileId}: ${e.toString()} \nStack: ${e.stack}`);
      return createResponse(false, null, `Failed to retrieve audio as base64: ${e.message}`);
  }
}

/**
 * Diagnostic function that can be called from the client to test folder access
 * @return {object} Diagnostic results
 */
function diagnoseFolderAccess() {
  try {
    Logger.log(`diagnoseFolderAccess: Starting diagnostic for ROOT_PROJECT_FOLDER_ID: ${ROOT_PROJECT_FOLDER_ID}`);
    
    const result = testFolderAccess(ROOT_PROJECT_FOLDER_ID);
    
    Logger.log(`diagnoseFolderAccess: Test result: ${JSON.stringify(result)}`);
    return result;
    
  } catch (e) {
    Logger.log(`Error in diagnoseFolderAccess: ${e.toString()}. Stack: ${e.stack ? e.stack : 'N/A'}`);
    return createResponse(false, { step: 'diagnostic wrapper' }, `Diagnostic function failed: ${e.message}`);
  }
}