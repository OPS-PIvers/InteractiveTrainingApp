// Server/DriveService.gs

/**
 * Creates a new folder within a specified parent folder in Google Drive.
 * @param {string} folderName The name for the new folder.
 * @param {string} parentFolderId The ID of the parent folder.
 * @return {string} The ID of the newly created folder.
 */
function createDriveFolder(folderName, parentFolderId) {
    try {
      const parentFolder = DriveApp.getFolderById(parentFolderId);
      const newFolder = parentFolder.createFolder(folderName);
      Logger.log(`Folder "${folderName}" created with ID: ${newFolder.getId()} inside parent ID: ${parentFolderId}`);
      return newFolder.getId();
    } catch (e) {
      Logger.log(`Error in createDriveFolder: ${e.toString()} - FolderName: ${folderName}, ParentID: ${parentFolderId}`);
      throw new Error(`Failed to create folder: ${e.message}`);
    }
}
  
/**
 * Saves or overwrites a JSON string to a file in a specified Drive folder.
 * If fileIdToOverwrite is provided and valid, it overwrites that file. Otherwise, creates a new file.
 * @param {string} fileName The name for the JSON file (e.g., "project_data.json").
 * @param {string} jsonContent The JSON string content to save.
 * @param {string} folderId The ID of the Drive folder where the file will be saved.
 * @param {string|null} fileIdToOverwrite Optional. The ID of an existing file to overwrite.
 * @return {string} The ID of the saved or overwritten file.
 */
function saveJsonToDriveFile(fileName, jsonContent, folderId, fileIdToOverwrite = null) {
  try {
    let file;
    if (fileIdToOverwrite) {
      try {
        file = DriveApp.getFileById(fileIdToOverwrite);
        // Optional: Check if file is in the correct folderId. For simplicity, assume fileIdToOverwrite is correct.
        file.setContent(jsonContent);
        Logger.log(`JSON content overwritten in file ID: ${file.getId()} (${fileName}) in folder ID: ${folderId}. Function: saveJsonToDriveFile.`);
      } catch (e) {
        // If fileIdToOverwrite is invalid or not found, throw an error.
        Logger.log(`Error in saveJsonToDriveFile: Failed to access file for overwrite. File ID: ${fileIdToOverwrite}. Error: ${e.message}. FileName: ${fileName}, FolderID: ${folderId}`);
        throw new Error(`Failed to access file for overwrite. File ID: ${fileIdToOverwrite}. Error: ${e.message}`);
      }
    } else {
      const folder = DriveApp.getFolderById(folderId);
      file = folder.createFile(fileName, jsonContent, MimeType.PLAIN_TEXT); // Or MimeType.JSON
      Logger.log(`New JSON file created: ${file.getId()} (${fileName}) in folder ID: ${folderId}. Function: saveJsonToDriveFile.`);
    }
    return file.getId();
  } catch (e) {
    // Catch any other error, including potential error from DriveApp.getFolderById() or file.createFile()
    Logger.log(`Error in saveJsonToDriveFile: ${e.toString()} - FileName: ${fileName}, FolderID: ${folderId}, FileToOverwrite: ${fileIdToOverwrite}`);
    // Re-throw the original error if it's one of our specific new errors, otherwise wrap it.
    if (e.message.startsWith("Failed to access file for overwrite")) {
      throw e;
    }
    throw new Error(`Failed to save JSON file: ${e.message}`);
  }
}

/**
 * Reads the text content of a file from Drive.
 * @param {string} fileId The ID of the file to read.
 * @return {string} The text content of the file.
 */
function readDriveFileContent(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString();
    Logger.log(`Content read from file ID: ${fileId}`);
    return content;
  } catch (e) {
    Logger.log(`Error in readDriveFileContent: ${e.toString()} - FileID: ${fileId}`);
    // Throw error so the calling function (getProjectDataForEditing) can catch it
    throw new Error(`Failed to read file content: ${e.message}`); 
  }
}

/**
 * Creates a new file in Drive from a blob. Used for media uploads.
 * @param {GoogleAppsScript.Base.Blob} blob The file blob.
 * @param {string} fileName The desired name for the file in Drive.
 * @param {string} folderId The ID of the folder to create the file in.
 * @return {GoogleAppsScript.Drive.File} The newly created Drive File object.
 */
function createFileInDriveFromBlob(blob, fileName, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob.setName(fileName)); // setName ensures the desired filename
    Logger.log(`File "${fileName}" created from blob in folder ID: ${folderId}. File ID: ${file.getId()}`);
    return file; // 'file' here *should* be a DriveApp.File
  } catch (e) {
    Logger.log(`Error in createFileInDriveFromBlob: ${e.toString()} - FileName: ${fileName}, FolderID: ${folderId}`);
    throw new Error(`Failed to create file from blob: ${e.message}`);
  }
}

/**
 * Deletes a folder and all its contents (files and subfolders) from Drive by trashing them.
 * @param {string} folderId The ID of the folder to delete.
 * @return {void}
 */
function deleteDriveFolderRecursive(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);

    let files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      file.setTrashed(true);
      Logger.log(`Trashed file: ${file.getName()} (ID: ${file.getId()}) from folder ${folder.getName()}`);
    }

    let subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      deleteDriveFolderRecursive(subFolder.getId()); // Recursive call
    }

    folder.setTrashed(true);
    Logger.log(`Trashed folder: ${folder.getName()} (ID: ${folder.getId()})`);
  } catch (e) {
    // If folder not found, it might have been already deleted or ID is wrong.
    if (e.message.toLowerCase().includes("not found") || e.message.toLowerCase().includes("no item with id")) {
        Logger.log(`Folder with ID "${folderId}" not found for deletion. It might have been already deleted. Error: ${e.toString()}`);
        return; // Don't throw an error if it's already gone
    }
    Logger.log(`Error in deleteDriveFolderRecursive: ${e.toString()} - FolderID: ${folderId}`);
    throw new Error(`Failed to delete folder recursively: ${e.message}`);
  }
}

