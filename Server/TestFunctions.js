// Server/TestFunctions.gs

// --- Test SheetService Functions ---

function storeTestProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

function getTestProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function deleteTestProperty(key) {
  PropertiesService.getScriptProperties().deleteProperty(key);
}

function test_appendRow() {
    const sheetId = PROJECT_INDEX_SHEET_ID; // Uses the constant from Code.gs
    const sheetName = PROJECT_INDEX_DATA_SHEET_NAME; // Uses the constant from Code.gs
    
    // This is your "test set" - a sample row of data
    const testRowData = [
      'testID_' + new Date().getTime(), // Unique ID for testing
      'My Test Project From Script',
      'dummyFolderID_123',
      'Draft',
      'dummyFileID_abc',
      new Date().toISOString(),
      new Date().toISOString()
    ];
    
    try {
      Logger.log(`Attempting to append row: ${testRowData.join(', ')} to sheet: ${sheetName}`);
      appendRowToSheet(sheetId, sheetName, testRowData);
      Logger.log("SUCCESS: test_appendRow completed. Check your sheet!");
    } catch (e) {
      Logger.log(`ERROR in test_appendRow: ${e.toString()}`);
    }
  }
  
  function test_findRow() {
    const sheetId = PROJECT_INDEX_SHEET_ID;
    const sheetName = PROJECT_INDEX_DATA_SHEET_NAME;
    const columnIndexToSearch = COL_PROJECT_ID; // ProjectID column (should be 1)
    
    // IMPORTANT: First, run test_appendRow() to add some data.
    // Then, copy one of the ProjectIDs from your sheet and paste it here:
    const valueToFind = 'testID_1746552194750'; 
  
    if (valueToFind === 'REPLACE_THIS_WITH_A_PROJECT_ID_FROM_YOUR_SHEET') {
      Logger.log("Please edit test_findRow and set 'valueToFind' to an actual ProjectID from your sheet.");
      return;
    }
    
    try {
      Logger.log(`Attempting to find value: "${valueToFind}" in column ${columnIndexToSearch} of sheet: ${sheetName}`);
      const rowIndex = findRowIndexByValue(sheetId, sheetName, columnIndexToSearch, valueToFind);
      if (rowIndex !== null) {
        Logger.log(`SUCCESS: test_findRow found value at row index: ${rowIndex}`);
      } else {
        Logger.log(`INFO: test_findRow did NOT find the value "${valueToFind}". This is okay if it's not there or if you haven't run test_appendRow yet.`);
      }
    } catch (e) {
      Logger.log(`ERROR in test_findRow: ${e.toString()}`);
    }
  }
  
  function test_updateRow() {
    const sheetId = PROJECT_INDEX_SHEET_ID;
    const sheetName = PROJECT_INDEX_DATA_SHEET_NAME;
  
    // IMPORTANT: 
    // 1. Run test_appendRow() if you haven't already.
    // 2. Run test_findRow() to get a rowIndex.
    // 3. Put that rowIndex here:
    const rowIndexToUpdate = 2; // Example: update the second row (the first data row after headers)
                               // MAKE SURE THIS ROW EXISTS!
  
    const updatedData = [
      'updatedID_KeepSame', // Assuming you want to keep the ID the same
      'My Updated Project Title',
      'updatedFolderID',
      'Active', // Changed status
      'updatedFileID',
      new Date().toISOString(), // New last modified
      'originalCreateDate_KeepSame' // Assuming you want to keep original create date
    ];
  
    // If you want to be precise, get the original row first
    const originalRow = getSheetRowData(sheetId, sheetName, rowIndexToUpdate);
    if (!originalRow) {
      Logger.log(`Cannot update row ${rowIndexToUpdate} as it doesn't exist or couldn't be fetched.`);
      return;
    }
    updatedData[0] = originalRow[COL_PROJECT_ID - 1]; // Keep original ProjectID
    updatedData[COL_CREATED_DATE -1] = originalRow[COL_CREATED_DATE -1]; // Keep original CreatedDate
    // Ensure updatedData has values for ALL columns up to the last one you want to change or sheet.getLastColumn().
    // For simplicity, this example assumes 7 columns.
  
    try {
      Logger.log(`Attempting to update row ${rowIndexToUpdate} in sheet: ${sheetName} with data: ${updatedData.join(', ')}`);
      updateSheetRow(sheetId, sheetName, rowIndexToUpdate, updatedData);
      Logger.log(`SUCCESS: test_updateRow completed for row ${rowIndexToUpdate}. Check your sheet!`);
    } catch (e) {
      Logger.log(`ERROR in test_updateRow: ${e.toString()}`);
    }
  }
  
  function test_getAllData() {
    const sheetId = PROJECT_INDEX_SHEET_ID;
    const sheetName = PROJECT_INDEX_DATA_SHEET_NAME;
    try {
      Logger.log(`Attempting to get all data from sheet: ${sheetName}`);
      const allData = getAllSheetData(sheetId, sheetName);
      Logger.log(`SUCCESS: test_getAllData retrieved ${allData.length} items.`);
      Logger.log(JSON.stringify(allData, null, 2)); // Pretty print the data
    } catch (e) {
      Logger.log(`ERROR in test_getAllData: ${e.toString()}`);
    }
  }
  
  
  function test_deleteRow() {
    const sheetId = PROJECT_INDEX_SHEET_ID;
    const sheetName = PROJECT_INDEX_DATA_SHEET_NAME;
  
    // IMPORTANT: Be careful with this!
    // Specify the 1-based row index you want to delete.
    // For example, if you ran test_appendRow once, it added a row. If that's row 2 (after headers), put 2 here.
    const rowIndexToDelete = 2; // MAKE SURE THIS IS THE CORRECT ROW YOU WANT TO DELETE!
    
    try {
      // Optional: Add a confirmation step or get row data first to be sure
      const rowData = getSheetRowData(sheetId, sheetName, rowIndexToDelete);
      if (rowData) {
        Logger.log(`Row to be deleted: ${rowData.join(', ')}`);
        // You could add a Browser.msgBox here for manual confirmation if running from editor directly
        // var ui = SpreadsheetApp.getUi();
        // var response = ui.alert('Confirm Deletion', 'Really delete row ' + rowIndexToDelete + '?', ui.ButtonSet.YES_NO);
        // if (response == ui.Button.NO) {
        //   Logger.log("Deletion cancelled by user.");
        //   return;
        // }
      } else {
        Logger.log(`Row ${rowIndexToDelete} not found. Cannot delete.`);
        return;
      }
  
      Logger.log(`Attempting to delete row ${rowIndexToDelete} from sheet: ${sheetName}`);
      deleteSheetRow(sheetId, sheetName, rowIndexToDelete);
      Logger.log(`SUCCESS: test_deleteRow completed for row ${rowIndexToDelete}. Check your sheet!`);
    } catch (e) {
      Logger.log(`ERROR in test_deleteRow: ${e.toString()}`);
    }
  }
  
  
  function test_createFolderInDrive() {
    const folderName = "My Script Test Folder - " + new Date().getTime();
    const parentFolderId = ROOT_PROJECT_FOLDER_ID; // From Code.gs
    try {
      Logger.log(`Attempting to create folder "${folderName}" in parent "${parentFolderId}"`);
      const createdFolderId = createDriveFolder(folderName, parentFolderId);
      if (createdFolderId) {
        storeTestProperty('testFolderId', createdFolderId); // Store it
        Logger.log(`SUCCESS: test_createFolderInDrive created folder with ID: ${createdFolderId}. ID stored.`);
      } else {
        Logger.log(`FAILURE: test_createFolderInDrive did not return a folder ID.`);
      }
    } catch (e) {
      Logger.log(`ERROR in test_createFolderInDrive: ${e.toString()}`);
    }
  }
  
  function test_saveJsonFileToDrive() {
    const testFolderId = getTestProperty('testFolderId'); // Retrieve it
    if (!testFolderId) {
      Logger.log("Please run test_createFolderInDrive first to create a folder and store its ID.");
      return;
    }
  
    const fileName = "test_project_data.json";
    const jsonContent = JSON.stringify({ 
      message: "Hello from test_saveJsonFileToDrive!", 
      timestamp: new Date().toISOString() 
    }, null, 2);
    const folderId = testFolderId;
    
    try {
      Logger.log(`Attempting to save JSON file "${fileName}" to folder "${folderId}"`);
      const createdFileId = saveJsonToDriveFile(fileName, jsonContent, folderId, null /* no overwrite initially */);
      if (createdFileId) {
        storeTestProperty('testFileId', createdFileId); // Store the file ID
        Logger.log(`SUCCESS: test_saveJsonFileToDrive saved file with ID: ${createdFileId}. ID stored.`);
      } else {
        Logger.log(`FAILURE: test_saveJsonFileToDrive did not return a file ID.`);
      }
    } catch (e) {
      Logger.log(`ERROR in test_saveJsonFileToDrive: ${e.toString()}`);
    }
  }
  
  function test_overwriteFileInDrive() {
    const testFolderId = getTestProperty('testFolderId');
    const testFileId = getTestProperty('testFileId');
  
    if (!testFolderId || !testFileId) {
      Logger.log("Please run test_createFolderInDrive and then test_saveJsonFileToDrive first.");
      return;
    }
  
    const fileName = "test_project_data.json"; // Same filename
    const updatedJsonContent = JSON.stringify({ 
      message: "CONTENT UPDATED by test_overwriteFileInDrive!", 
      timestamp: new Date().toISOString(),
      updateCount: 1 
    }, null, 2);
    const folderId = testFolderId; 
    const fileIdToOverwrite = testFileId; 
  
    try {
      Logger.log(`Attempting to overwrite file ID "${fileIdToOverwrite}" in folder "${folderId}"`);
      const updatedFileId = saveJsonToDriveFile(fileName, updatedJsonContent, folderId, fileIdToOverwrite);
      Logger.log(`SUCCESS: test_overwriteFileInDrive overwrote file. New/Same ID: ${updatedFileId}`);
      if (updatedFileId && updatedFileId !== testFileId) {
          Logger.log("INFO: File ID changed after overwrite. This can happen if original fileId was invalid/deleted. Updating stored ID.");
          storeTestProperty('testFileId', updatedFileId); // Update stored ID if it changed
      } else if (!updatedFileId) {
          Logger.log("FAILURE: test_overwriteFileInDrive did not return a file ID after attempting overwrite.");
      }
    } catch (e) {
      Logger.log(`ERROR in test_overwriteFileInDrive: ${e.toString()}`);
    }
  }
  
  function test_readFileContentFromDrive() {
    const testFileId = getTestProperty('testFileId'); // Retrieve it
    if (!testFileId) {
      Logger.log("Please run test_saveJsonFileToDrive first to create, save a file, and store its ID.");
      return;
    }
    const fileId = testFileId;
    try {
      Logger.log(`Attempting to read content from file ID "${fileId}"`);
      const content = readDriveFileContent(fileId);
      Logger.log(`SUCCESS: test_readFileContentFromDrive read content:`);
      Logger.log(content);
    } catch (e) {
      Logger.log(`ERROR in test_readFileContentFromDrive: ${e.toString()}`);
    }
  }
  
  function test_deleteTestFolderRecursive() {
    const testFolderId = getTestProperty('testFolderId'); // Get the ID from properties
    if (!testFolderId) {
      Logger.log("No test folder ID stored in script properties. Run test_createFolderInDrive first or set the property manually.");
      return;
    }
  
    // Optional: Add UI confirmation if running from a bound script, otherwise rely on careful execution
    // var ui = SpreadsheetApp.getUi(); 
    // var response = ui.alert('Confirm Deletion', `Really delete folder ID "${testFolderId}" and all its contents? This cannot be undone easily.`, ui.ButtonSet.YES_NO);
    // if (response == ui.Button.NO) {
    //   Logger.log("Deletion cancelled by user for folder: " + testFolderId); // Use testFolderId from properties
    //   return;
    // }
  
    try {
      Logger.log(`Attempting to delete folder ID "${testFolderId}" recursively.`); // Use testFolderId from properties
      deleteDriveFolderRecursive(testFolderId); // This is your actual service function, using the ID from properties
      Logger.log(`SUCCESS: test_deleteTestFolderRecursive completed for folder ID: ${testFolderId}. Check your Drive Trash.`); // Use testFolderId from properties
      
      // Clean up stored properties
      deleteTestProperty('testFolderId');
      deleteTestProperty('testFileId');
      Logger.log("Stored testFolderId and testFileId properties have been deleted.");
  
    } catch (e) {
      Logger.log(`ERROR in test_deleteTestFolderRecursive: ${e.toString()} - Attempted to delete FolderID: ${testFolderId}`); // Log the ID being processed
    }
  }