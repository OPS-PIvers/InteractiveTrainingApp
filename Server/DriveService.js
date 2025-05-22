// Server/DriveService.gs

/**
 * Diagnostic function to test folder access and permissions
 * @param {string} folderId The folder ID to test
 * @return {object} Diagnostic information about the folder
 */
function testFolderAccess(folderId) {
  try {
    Logger.log(`Testing access to folder ID: ${folderId}`);
    
    // Test 1: Can we get the folder?
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
      Logger.log(`✓ Successfully retrieved folder: ${folder.getName()}`);
    } catch (e) {
      Logger.log(`✗ Failed to retrieve folder: ${e.toString()}`);
      return {
        success: false,
        error: `Cannot access folder: ${e.message}`,
        step: 'getFolderById'
      };
    }
    
    // Test 2: Can we read folder properties?
    try {
      const folderName = folder.getName();
      const folderId = folder.getId();
      const folderUrl = folder.getUrl();
      Logger.log(`✓ Folder properties - Name: ${folderName}, ID: ${folderId}`);
      Logger.log(`✓ Folder URL: ${folderUrl}`);
    } catch (e) {
      Logger.log(`✗ Failed to read folder properties: ${e.toString()}`);
      return {
        success: false,
        error: `Cannot read folder properties: ${e.message}`,
        step: 'folder properties'
      };
    }
    
    // Test 3: Can we create a test folder?
    try {
      const testFolderName = `TEST_FOLDER_${new Date().getTime()}`;
      const testFolder = folder.createFolder(testFolderName);
      Logger.log(`✓ Successfully created test folder: ${testFolder.getName()} with ID: ${testFolder.getId()}`);
      
      // Clean up - delete the test folder
      try {
        testFolder.setTrashed(true);
        Logger.log(`✓ Test folder cleaned up successfully`);
      } catch (cleanupError) {
        Logger.log(`⚠ Warning: Could not clean up test folder: ${cleanupError.toString()}`);
      }
      
    } catch (e) {
      Logger.log(`✗ Failed to create test folder: ${e.toString()}`);
      return {
        success: false,
        error: `Cannot create folders in this location: ${e.message}`,
        step: 'createFolder'
      };
    }
    
    // Test 4: Check current user and permissions
    try {
      const currentUser = Session.getActiveUser().getEmail();
      Logger.log(`✓ Current user: ${currentUser}`);
    } catch (e) {
      Logger.log(`⚠ Could not get current user: ${e.toString()}`);
    }
    
    return {
      success: true,
      message: 'All folder access tests passed',
      folderName: folder.getName(),
      folderUrl: folder.getUrl()
    };
    
  } catch (e) {
    Logger.log(`Error in testFolderAccess: ${e.toString()}`);
    return {
      success: false,
      error: `Diagnostic test failed: ${e.message}`,
      step: 'general'
    };
  }
}

/**
 * Creates a new folder within a specified parent folder in Google Drive.
 * Enhanced with better error handling and diagnostics.
 * @param {string} folderName The name for the new folder.
 * @param {string} parentFolderId The ID of the parent folder.
 * @return {string} The ID of the newly created folder.
 */
function createDriveFolder(folderName, parentFolderId) {
  try {
    Logger.log(`createDriveFolder: Attempting to create folder "${folderName}" in parent ID: ${parentFolderId}`);
    
    // Enhanced error checking
    if (!folderName || folderName.trim() === '') {
      throw new Error('Folder name cannot be empty');
    }
    
    if (!parentFolderId || parentFolderId.trim() === '') {
      throw new Error('Parent folder ID cannot be empty');
    }
    
    // Step 1: Try to get the parent folder
    let parentFolder;
    try {
      parentFolder = DriveApp.getFolderById(parentFolderId);
      Logger.log(`createDriveFolder: Successfully accessed parent folder: ${parentFolder.getName()}`);
    } catch (e) {
      Logger.log(`createDriveFolder: Failed to access parent folder ID "${parentFolderId}": ${e.toString()}`);
      
      // Provide more specific error messages
      if (e.toString().includes('not found') || e.toString().includes('does not exist')) {
        throw new Error(`Parent folder not found. Please verify the folder ID: ${parentFolderId}`);
      } else if (e.toString().includes('access') || e.toString().includes('permission')) {
        throw new Error(`No permission to access parent folder. Please check sharing settings for folder ID: ${parentFolderId}`);
      } else {
        throw new Error(`Cannot access parent folder (${parentFolderId}): ${e.message}`);
      }
    }
    
    // Step 2: Try to create the new folder
    let newFolder;
    try {
      newFolder = parentFolder.createFolder(folderName);
      Logger.log(`createDriveFolder: Successfully created folder "${folderName}" with ID: ${newFolder.getId()}`);
    } catch (e) {
      Logger.log(`createDriveFolder: Failed to create folder "${folderName}" in parent "${parentFolder.getName()}": ${e.toString()}`);
      
      // Provide more specific error messages
      if (e.toString().includes('quota') || e.toString().includes('limit')) {
        throw new Error(`Drive storage quota exceeded. Cannot create new folder: ${folderName}`);
      } else if (e.toString().includes('permission') || e.toString().includes('access')) {
        throw new Error(`No permission to create folders in: ${parentFolder.getName()}`);
      } else if (e.toString().includes('duplicate') || e.toString().includes('exists')) {
        Logger.log(`createDriveFolder: Folder name might already exist, but continuing...`);
        throw new Error(`Folder "${folderName}" may already exist in: ${parentFolder.getName()}`);
      } else {
        throw new Error(`Cannot create folder "${folderName}": ${e.message}`);
      }
    }
    
    // Step 3: Verify the folder was created successfully
    try {
      const newFolderId = newFolder.getId();
      const newFolderName = newFolder.getName();
      
      if (!newFolderId) {
        throw new Error('Created folder has no ID');
      }
      
      Logger.log(`createDriveFolder: Folder creation verified - Name: "${newFolderName}", ID: ${newFolderId}`);
      return newFolderId;
      
    } catch (e) {
      Logger.log(`createDriveFolder: Failed to verify created folder: ${e.toString()}`);
      throw new Error(`Folder may have been created but verification failed: ${e.message}`);
    }
    
  } catch (e) {
    // Log the full error for debugging
    Logger.log(`Error in createDriveFolder: ${e.toString()} - FolderName: ${folderName}, ParentID: ${parentFolderId}`);
    Logger.log(`Stack trace: ${e.stack || 'No stack trace available'}`);
    
    // Re-throw with the specific error message (don't wrap it again if it's already our custom message)
    if (e.message.startsWith('Parent folder not found') || 
        e.message.startsWith('No permission') || 
        e.message.startsWith('Cannot access parent folder') ||
        e.message.startsWith('Drive storage quota') ||
        e.message.startsWith('Cannot create folder')) {
      throw e; // Re-throw our custom error as-is
    } else {
      throw new Error(`Failed to create folder: ${e.message}`);
    }
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
    Logger.log(`saveJsonToDriveFile: Starting - FileName: ${fileName}, FolderID: ${folderId}, FileToOverwrite: ${fileIdToOverwrite || 'none'}`);
    
    let file;
    if (fileIdToOverwrite) {
      try {
        file = DriveApp.getFileById(fileIdToOverwrite);
        file.setContent(jsonContent);
        Logger.log(`saveJsonToDriveFile: JSON content overwritten in file ID: ${file.getId()} (${fileName})`);
      } catch (e) {
        Logger.log(`saveJsonToDriveFile: Failed to access file for overwrite. File ID: ${fileIdToOverwrite}. Error: ${e.message}`);
        
        if (e.toString().includes('not found') || e.toString().includes('does not exist')) {
          throw new Error(`File to overwrite not found: ${fileIdToOverwrite}`);
        } else if (e.toString().includes('access') || e.toString().includes('permission')) {
          throw new Error(`No permission to overwrite file: ${fileIdToOverwrite}`);
        } else {
          throw new Error(`Failed to access file for overwrite: ${e.message}`);
        }
      }
    } else {
      try {
        const folder = DriveApp.getFolderById(folderId);
        file = folder.createFile(fileName, jsonContent, MimeType.PLAIN_TEXT);
        Logger.log(`saveJsonToDriveFile: New JSON file created: ${file.getId()} (${fileName}) in folder: ${folder.getName()}`);
      } catch (e) {
        Logger.log(`saveJsonToDriveFile: Failed to create new file. Error: ${e.toString()}`);
        
        if (e.toString().includes('not found') || e.toString().includes('does not exist')) {
          throw new Error(`Folder not found: ${folderId}`);
        } else if (e.toString().includes('access') || e.toString().includes('permission')) {
          throw new Error(`No permission to create files in folder: ${folderId}`);
        } else if (e.toString().includes('quota') || e.toString().includes('limit')) {
          throw new Error(`Drive storage quota exceeded. Cannot create file: ${fileName}`);
        } else {
          throw new Error(`Failed to create file: ${e.message}`);
        }
      }
    }
    
    return file.getId();
    
  } catch (e) {
    Logger.log(`Error in saveJsonToDriveFile: ${e.toString()} - FileName: ${fileName}, FolderID: ${folderId}, FileToOverwrite: ${fileIdToOverwrite}`);
    Logger.log(`Stack trace: ${e.stack || 'No stack trace available'}`);
    
    // Re-throw specific errors as-is, wrap others
    if (e.message.startsWith('File to overwrite not found') || 
        e.message.startsWith('No permission') || 
        e.message.startsWith('Folder not found') ||
        e.message.startsWith('Drive storage quota') ||
        e.message.startsWith('Failed to create file') ||
        e.message.startsWith('Failed to access file')) {
      throw e;
    } else {
      throw new Error(`Failed to save JSON file: ${e.message}`);
    }
  }
}

/**
 * Reads the text content of a file from Drive.
 * @param {string} fileId The ID of the file to read.
 * @return {string} The text content of the file.
 */
function readDriveFileContent(fileId) {
  try {
    Logger.log(`readDriveFileContent: Reading file ID: ${fileId}`);
    
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString();
    
    Logger.log(`readDriveFileContent: Successfully read ${content.length} characters from file: ${file.getName()}`);
    return content;
    
  } catch (e) {
    Logger.log(`Error in readDriveFileContent: ${e.toString()} - FileID: ${fileId}`);
    
    if (e.toString().includes('not found') || e.toString().includes('does not exist')) {
      throw new Error(`File not found: ${fileId}`);
    } else if (e.toString().includes('access') || e.toString().includes('permission')) {
      throw new Error(`No permission to read file: ${fileId}`);
    } else {
      throw new Error(`Failed to read file content: ${e.message}`);
    }
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
    Logger.log(`createFileInDriveFromBlob: Creating file "${fileName}" in folder: ${folderId}`);
    
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob.setName(fileName));
    
    Logger.log(`createFileInDriveFromBlob: File "${fileName}" created successfully. File ID: ${file.getId()}`);
    return file;
    
  } catch (e) {
    Logger.log(`Error in createFileInDriveFromBlob: ${e.toString()} - FileName: ${fileName}, FolderID: ${folderId}`);
    
    if (e.toString().includes('not found') || e.toString().includes('does not exist')) {
      throw new Error(`Folder not found: ${folderId}`);
    } else if (e.toString().includes('access') || e.toString().includes('permission')) {
      throw new Error(`No permission to create files in folder: ${folderId}`);
    } else if (e.toString().includes('quota') || e.toString().includes('limit')) {
      throw new Error(`Drive storage quota exceeded. Cannot create file: ${fileName}`);
    } else {
      throw new Error(`Failed to create file from blob: ${e.message}`);
    }
  }
}

/**
 * Deletes a folder and all its contents (files and subfolders) from Drive by trashing them.
 * @param {string} folderId The ID of the folder to delete.
 * @return {void}
 */
function deleteDriveFolderRecursive(folderId) {
  try {
    Logger.log(`deleteDriveFolderRecursive: Starting deletion of folder: ${folderId}`);
    
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();
    
    Logger.log(`deleteDriveFolderRecursive: Processing folder "${folderName}"`);

    let files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      file.setTrashed(true);
      Logger.log(`Trashed file: ${file.getName()} (ID: ${file.getId()}) from folder ${folderName}`);
    }

    let subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      deleteDriveFolderRecursive(subFolder.getId()); // Recursive call
    }

    folder.setTrashed(true);
    Logger.log(`Trashed folder: ${folderName} (ID: ${folder.getId()})`);
    
  } catch (e) {
    Logger.log(`Error in deleteDriveFolderRecursive: ${e.toString()} - FolderID: ${folderId}`);
    
    // If folder not found, it might have been already deleted or ID is wrong.
    if (e.message.toLowerCase().includes("not found") || e.message.toLowerCase().includes("no item with id")) {
      Logger.log(`Folder with ID "${folderId}" not found for deletion. It might have been already deleted.`);
      return; // Don't throw an error if it's already gone
    }
    
    if (e.toString().includes('access') || e.toString().includes('permission')) {
      throw new Error(`No permission to delete folder: ${folderId}`);
    } else {
      throw new Error(`Failed to delete folder recursively: ${e.message}`);
    }
  }
}