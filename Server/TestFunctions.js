// Server/TestFunctions.gs

// --- Test SheetService Functions ---

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
  
  
  // --- Test DriveService Functions ---
  let testFolderId_global; // To store folder ID between tests
  let testFileId_global;   // To store file ID between tests
  
  function test_createFolderInDrive() {
    const folderName = "My Script Test Folder - " + new Date().getTime();
    const parentFolderId = ROOT_PROJECT_FOLDER_ID; // From Code.gs
    try {
      Logger.log(`Attempting to create folder "${folderName}" in parent "${parentFolderId}"`);
      testFolderId_global = createDriveFolder(folderName, parentFolderId);
      Logger.log(`SUCCESS: test_createFolderInDrive created folder with ID: ${testFolderId_global}`);
    } catch (e) {
      Logger.log(`ERROR in test_createFolderInDrive: ${e.toString()}`);
    }
  }
  
  function test_saveJsonFileToDrive() {
    if (!testFolderId_global) {
      Logger.log("Please run test_createFolderInDrive first to create a folder.");
      return;
    }
    const fileName = "test_project_data.json";
    const jsonContent = JSON.stringify({ 
      message: "Hello from test_saveJsonFileToDrive!", 
      timestamp: new Date().toISOString() 
    }, null, 2);
    const folderId = testFolderId_global;
    
    try {
      Logger.log(`Attempting to save JSON file "${fileName}" to folder "${folderId}"`);
      testFileId_global = saveJsonToDriveFile(fileName, jsonContent, folderId, null /* no overwrite initially */);
      Logger.log(`SUCCESS: test_saveJsonFileToDrive saved file with ID: ${testFileId_global}`);
    } catch (e) {
      Logger.log(`ERROR in test_saveJsonFileToDrive: ${e.toString()}`);
    }
  }
  
  function test_overwriteFileInDrive() {
    if (!testFolderId_global || !testFileId_global) {
      Logger.log("Please run test_createFolderInDrive and then test_saveJsonFileToDrive first.");
      return;
    }
    const fileName = "test_project_data.json"; // Same filename
    const updatedJsonContent = JSON.stringify({ 
      message: "CONTENT UPDATED by test_overwriteFileInDrive!", 
      timestamp: new Date().toISOString(),
      updateCount: 1 
    }, null, 2);
    const folderId = testFolderId_global; // Folder ID where file exists
    const fileIdToOverwrite = testFileId_global; // ID of the file to overwrite
  
    try {
      Logger.log(`Attempting to overwrite file ID "${fileIdToOverwrite}" in folder "${folderId}"`);
      const updatedFileId = saveJsonToDriveFile(fileName, updatedJsonContent, folderId, fileIdToOverwrite);
      Logger.log(`SUCCESS: test_overwriteFileInDrive overwrote file. New/Same ID: ${updatedFileId}`);
      if (updatedFileId !== fileIdToOverwrite) {
          Logger.log("WARNING: File ID changed after overwrite. This can happen if original fileId was invalid/deleted.");
          testFileId_global = updatedFileId; // Update global if it changed
      }
    } catch (e) {
      Logger.log(`ERROR in test_overwriteFileInDrive: ${e.toString()}`);
    }
  }
  
  function test_readFileContentFromDrive() {
    if (!testFileId_global) {
      Logger.log("Please run test_saveJsonFileToDrive first to create and save a file.");
      return;
    }
    const fileId = testFileId_global;
    try {
      Logger.log(`Attempting to read content from file ID "${fileId}"`);
      const content = readDriveFileContent(fileId);
      Logger.log(`SUCCESS: test_readFileContentFromDrive read content:`);
      Logger.log(content);
    } catch (e)
   {
      Logger.log(`ERROR in test_readFileContentFromDrive: ${e.toString()}`);
    }
  }
  
  function test_deleteTestFolderRecursive() {
    // WARNING: This will delete the folder created by test_createFolderInDrive and all its contents.
    if (!testFolderId_global) {
      Logger.log("No test folder ID stored (testFolderId_global is undefined). Run test_createFolderInDrive first or set the ID manually.");
      // You could prompt user for folder ID if you want to run this standalone
      // const folderIdToDelete = Browser.inputBox("Enter Folder ID to Delete Recursively (BE CAREFUL!):");
      // if (!folderIdToDelete || folderIdToDelete.trim() === "") {
      //   Logger.log("Deletion cancelled or no ID provided.");
      //   return;
      // }
      // testFolderId_global = folderIdToDelete; // for this run
      return;
    }
  
    var ui = SpreadsheetApp.getUi(); // Or DocumentApp.getUi() or FormApp.getUi() if script is bound
                                    // If standalone script, this might not work well without a bound document.
                                    // In that case, just rely on direct execution and careful checking.
    
    // var response = ui.alert(
    //   'Confirm Deletion', 
    //   `Really delete folder ID "${testFolderId_global}" and all its contents? This cannot be undone easily.`,
    //   ui.ButtonSet.YES_NO
    // );
  
    // if (response == ui.Button.NO) {
    //   Logger.log("Deletion cancelled by user for folder: " + testFolderId_global);
    //   return;
    // }
  
  
    try {
      Logger.log(`Attempting to delete folder ID "${testFolderId_global}" recursively.`);
      deleteDriveFolderRecursive(testFolderId_global);
      Logger.log(`SUCCESS: test_deleteTestFolderRecursive completed for folder ID: ${testFolderId_global}. Check your Drive Trash.`);
      testFolderId_global = null; // Clear the global ID as it's now deleted
      testFileId_global = null;   // Clear associated file ID too
    } catch (e) {
      Logger.log(`ERROR in test_deleteTestFolderRecursive: ${e.toString()}`);
    }
  }