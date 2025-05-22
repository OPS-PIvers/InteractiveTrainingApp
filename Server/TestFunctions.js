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

// --- Tests for Fixed Bugs ---

/**
 * Tests that saveJsonToDriveFile throws an error when trying to overwrite a non-existent file ID
 * and does not create a new file.
 */
function test_saveJsonToDrive_overwriteNonExistentFile() {
  const folderId = getTestProperty('testFolderId'); // Assumes test_createFolderInDrive was run
  if (!folderId) {
    Logger.log("SKIPPING TEST: test_saveJsonToDrive_overwriteNonExistentFile - Run test_createFolderInDrive first.");
    return;
  }

  const fileName = "test_overwrite_nonexistent.json";
  const jsonContent = JSON.stringify({ message: "Test content for overwrite failure", timestamp: new Date().toISOString() });
  const nonExistentFileId = "nonExistentFileId_12345abcde"; // Intentionally invalid ID

  let errorThrown = false;
  try {
    Logger.log(`Attempting to call saveJsonToDriveFile to overwrite non-existent file ID: ${nonExistentFileId}`);
    saveJsonToDriveFile(fileName, jsonContent, folderId, nonExistentFileId);
    // If saveJsonToDriveFile does not throw, it's a failure for this test case
    Logger.log("ERROR: test_saveJsonToDrive_overwriteNonExistentFile - saveJsonToDriveFile DID NOT throw an error as expected.");
  } catch (e) {
    if (e.message.includes("Failed to access file for overwrite") || e.message.includes(nonExistentFileId)) {
      Logger.log(`SUCCESS: test_saveJsonToDrive_overwriteNonExistentFile - saveJsonToDriveFile threw an error as expected: ${e.toString()}`);
      errorThrown = true;
    } else {
      Logger.log(`ERROR: test_saveJsonToDrive_overwriteNonExistentFile - An unexpected error was thrown: ${e.toString()}`);
    }
  }

  if (!errorThrown) {
    Logger.log("FINAL ERROR: test_saveJsonToDrive_overwriteNonExistentFile - Expected an error to be thrown, but it wasn't.");
  }
  // Manual check: Verify in Drive that no file named "test_overwrite_nonexistent.json" was created in the test folder.
  Logger.log("Manual Check Required: Please verify in Google Drive that no file named 'test_overwrite_nonexistent.json' was created in your test folder during this test.");
}

/**
 * Tests AdminController.updateProjectStatus when the underlying JSON file update fails.
 * Expects the function to return success: false and the sheet not to be updated.
 *
 * MANUAL SETUP REQUIRED for first run:
 * 1. Ensure a project exists in your ProjectIndex sheet.
 * 2. Note its ProjectID.
 * 3. For that project, temporarily set its 'ProjectDataFileID' in the sheet to an invalid/fake ID
 *    like "fakeFileIdForTestingJsonUpdateFail".
 * 4. Run this test, providing the ProjectID.
 * 5. After the test, REVERT the 'ProjectDataFileID' in the sheet to its correct value.
 */
function test_updateProjectStatus_jsonUpdateFails() {
  const projectIdToTest = "REPLACE_WITH_PROJECT_ID_FROM_SHEET_WITH_FAKE_DATA_FILE_ID"; // Replace this
  const newStatusToSet = "Active";

  if (projectIdToTest === "REPLACE_WITH_PROJECT_ID_FROM_SHEET_WITH_FAKE_DATA_FILE_ID") {
    Logger.log("SKIPPING TEST: test_updateProjectStatus_jsonUpdateFails - Please edit the test function and provide a valid projectIdToTest and follow setup instructions.");
    return;
  }

  Logger.log(`Attempting to run test_updateProjectStatus_jsonUpdateFails for projectId: ${projectIdToTest}`);
  Logger.log("IMPORTANT: This test assumes you have manually set the 'ProjectDataFileID' for this project to an invalid ID in the ProjectIndex sheet to simulate JSON update failure.");

  // Fetch original status from sheet to compare later
  let originalStatusInSheet = null;
  let rowIndex = null; 
  try {
    rowIndex = findRowIndexByValue(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, COL_PROJECT_ID, projectIdToTest);
    if (rowIndex) {
      const rowData = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
      originalStatusInSheet = rowData ? rowData[COL_STATUS - 1] : "COULD_NOT_FETCH";
    } else {
      Logger.log(`ERROR: test_updateProjectStatus_jsonUpdateFails - ProjectID ${projectIdToTest} not found in sheet. Aborting test.`);
      return;
    }
  } catch(e) {
    Logger.log(`ERROR: test_updateProjectStatus_jsonUpdateFails - Could not fetch original project status. Error: ${e.toString()}`);
    return;
  }

  Logger.log(`Original status for project ${projectIdToTest} in sheet: ${originalStatusInSheet}`);

  try {
    const result = updateProjectStatus(projectIdToTest, newStatusToSet);

    if (result && result.success === false) {
      if (result.error && (result.error.includes("Failed to update project data file") || result.error.includes("Failed to read file content") || result.error.includes("Failed to access file for overwrite"))) {
         Logger.log(`SUCCESS: test_updateProjectStatus_jsonUpdateFails - updateProjectStatus returned success:false as expected. Error: ${result.error}`);
      } else {
         Logger.log(`ERROR: test_updateProjectStatus_jsonUpdateFails - updateProjectStatus returned success:false, but the error message was not as expected. Error: ${result.error}`);
      }
    } else {
      Logger.log(`ERROR: test_updateProjectStatus_jsonUpdateFails - updateProjectStatus did not return {success: false}. Result: ${JSON.stringify(result)}`);
    }

    // Verify sheet status was not changed
    let statusAfterTestInSheet = "VERIFICATION_NEEDED";
    if (rowIndex) { // rowIndex fetched earlier
        const rowDataAfter = getSheetRowData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, rowIndex);
        statusAfterTestInSheet = rowDataAfter ? rowDataAfter[COL_STATUS -1] : "COULD_NOT_FETCH_AFTER";
    }
    
    Logger.log(`Status for project ${projectIdToTest} in sheet AFTER test: ${statusAfterTestInSheet}`);
    if (originalStatusInSheet === statusAfterTestInSheet) {
        Logger.log(`SUCCESS: test_updateProjectStatus_jsonUpdateFails - Project status in the sheet REMAINS UNCHANGED (${originalStatusInSheet}), which is correct.`);
    } else {
        Logger.log(`ERROR: test_updateProjectStatus_jsonUpdateFails - Project status in the sheet CHANGED from ${originalStatusInSheet} to ${statusAfterTestInSheet}, which is INCORRECT for this test.`);
    }

  } catch (e) {
    Logger.log(`ERROR: test_updateProjectStatus_jsonUpdateFails - An unexpected error occurred: ${e.toString()} 
Stack: ${e.stack ? e.stack : 'No stack available'}`);
  } finally {
    Logger.log("Reminder: If you manually changed a ProjectDataFileID for this test, please revert it now to maintain data integrity.");
  }
}

/**
 * Helper function to create a dummy project for testing (Optional, but recommended for more stable tests).
 * This would involve:
 * - Creating a folder in Drive.
 * - Creating a dummy project_data.json file in it.
 * - Appending a row to the ProjectIndex sheet.
 * - Returning the new project's details (ID, folderId, fileId).
 * Requires careful implementation to clean up afterwards (e.g., a corresponding deleteTestProject function).
 */
function setupTestProject(projectTitlePrefix = "TestProject_") {
    // Implementation would go here...
    // createProject() from AdminController could be adapted or called.
    // For now, this is a placeholder.
    Logger.log("Placeholder for setupTestProject. Implement if needed for more robust automated testing.");
    return null; 
}