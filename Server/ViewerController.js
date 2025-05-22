// Server/ViewerController.gs

/**
 * Retrieves a list of "Active" projects from the "ProjectIndex" sheet.
 * Filters projects based on the Status column.
 * @return {Array<Object>} An array of project objects (e.g., { projectId: '...', projectTitle: '...' })
 *                         or an empty array if no active projects are found or an error occurs.
 */
function getActiveProjectsList() {
  try {
    // PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME, and COL_STATUS are defined in Code.gs
    // getAllSheetData is defined in SheetService.gs
    const allProjectsData = getAllSheetData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME);

    if (!allProjectsData || !Array.isArray(allProjectsData)) {
      Logger.log("getActiveProjectsList: No data or invalid data returned from getAllSheetData.");
      return [];
    }

    const activeProjects = [];
    // getAllSheetData now consistently returns an array of objects.
    allProjectsData.forEach(project => {
      const status = project['Status']; // Directly access by header name
      const projectId = project['ProjectID']; // Directly access by header name
      const projectTitle = project['ProjectTitle']; // Directly access by header name

      if (status === "Active" && projectId && projectTitle) {
        activeProjects.push({
          projectId: projectId,
          projectTitle: projectTitle
        });
      } else if (status === "Active" && (!projectId || !projectTitle)) {
        // Log if an "Active" project is missing essential details
        Logger.log(`getActiveProjectsList: Found an 'Active' project with missing ProjectID or ProjectTitle. Data: ${JSON.stringify(project)}`);
      }
    });

    Logger.log(`getActiveProjectsList: Found ${activeProjects.length} active projects from ${allProjectsData.length} total projects processed.`);
    return activeProjects;

  } catch (e) {
    Logger.log(`Error in getActiveProjectsList: ${e.toString()} \nStack: ${e.stack ? e.stack : 'No stack available'}`);
    return [];
  }
}

// Server/ViewerController.gs

/**
 * Retrieves the project's data JSON string from its file in Google Drive for viewer.
 * Ensures the project is "Active".
 * @param {string} projectId The ID of the project.
 * @return {string} The JSON string content of the project data file if active.
 *                  Throws an error if not found, not active, or other error.
 */
function getProjectViewData(projectId) {
  try {
    Logger.log(`getProjectViewData: Attempting to load data for projectId: ${projectId}`);
    if (!projectId) {
      Logger.log("getProjectViewData: ProjectID is missing.");
      throw new Error("Project ID is required.");
    }

    const allProjects = getAllSheetData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME);
    
    // getAllSheetData now consistently returns an array of objects.
    // Find the project entry using direct property access.
    const projectEntry = allProjects.find(p => p && p['ProjectID'] === projectId);

    if (!projectEntry) {
      Logger.log(`getProjectViewData: Project with ID "${projectId}" not found in ProjectIndex.`);
      throw new Error("Project not found.");
    }

    // Direct property access for status and file ID
    const status = projectEntry['Status'];
    if (status !== "Active") {
      Logger.log(`getProjectViewData: Project with ID "${projectId}" is not active (Status: ${status}).`);
      throw new Error("Project is not currently active.");
    }

    const projectDataFileId = projectEntry['ProjectDataFileID'];
    if (!projectDataFileId) {
      Logger.log(`getProjectViewData: ProjectDataFileID is missing for active project "${projectId}".`);
      throw new Error("Project data file reference is missing.");
    }
    Logger.log(`getProjectViewData: Found ProjectDataFileID: ${projectDataFileId} for active project ${projectId}.`);

    const jsonContent = readDriveFileContent(projectDataFileId); // This can throw

    Logger.log(`getProjectViewData: Successfully read content for file ID ${projectDataFileId}.`);
    return jsonContent; // Return the JSON string directly on success

  } catch (e) {
    Logger.log(`Error in getProjectViewData for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
    // Re-throw the error so it's caught by .withFailureHandler on the client side
    throw new Error(`Server error fetching project data for viewing: ${e.message}`);
  }
}