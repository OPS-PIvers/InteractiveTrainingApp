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
    allProjectsData.forEach(project => {
      const status = project['Status'] || (project[COL_STATUS - 1] ? project[COL_STATUS - 1] : null);
      const projectId = project['ProjectID'] || (project[COL_PROJECT_ID - 1] ? project[COL_PROJECT_ID - 1] : null);
      const projectTitle = project['ProjectTitle'] || (project[COL_PROJECT_TITLE - 1] ? project[COL_PROJECT_TITLE - 1] : null);

      if (status === "Active" && projectId && projectTitle) {
        activeProjects.push({
          projectId: projectId,
          projectTitle: projectTitle
        });
      }
    });

    Logger.log(`getActiveProjectsList: Found ${activeProjects.length} active projects.`);
    return activeProjects;

  } catch (e) {
    Logger.log(`Error in getActiveProjectsList: ${e.toString()} \nStack: ${e.stack ? e.stack : 'No stack available'}`);
    return [];
  }
}

/**
 * Retrieves the project's data JSON string from its file in Google Drive for viewer.
 * Ensures the project is "Active".
 * @param {string} projectId The ID of the project.
 * @return {string|null} The JSON string content of the project data file if active, 
 *                      or null if the project is not found, not active, or an error occurs.
 */
function getProjectViewData(projectId) {
  try {
    Logger.log(`getProjectViewData: Attempting to load data for projectId: ${projectId}`);
    if (!projectId) {
      Logger.log("getProjectViewData: ProjectID is missing.");
      return { success: false, error: "Project ID is required.", data: null };
    }

    // Find the ProjectDataFileID and Status from the ProjectIndex sheet
    const allProjects = getAllSheetData(PROJECT_INDEX_SHEET_ID, PROJECT_INDEX_DATA_SHEET_NAME);
    let projectEntry = null;
    if (allProjects && Array.isArray(allProjects)) {
      projectEntry = allProjects.find(p => (p['ProjectID'] || p[COL_PROJECT_ID - 1]) === projectId);
    }

    if (!projectEntry) {
      Logger.log(`getProjectViewData: Project with ID "${projectId}" not found in ProjectIndex.`);
      return { success: false, error: "Project not found.", data: null };
    }

    const status = projectEntry['Status'] || projectEntry[COL_STATUS - 1];
    if (status !== "Active") {
      Logger.log(`getProjectViewData: Project with ID "${projectId}" is not active (Status: ${status}).`);
      return { success: false, error: "Project is not currently active.", data: null };
    }

    const projectDataFileId = projectEntry['ProjectDataFileID'] || projectEntry[COL_PROJECT_DATA_FILE_ID - 1];

    if (!projectDataFileId) {
      Logger.log(`getProjectViewData: ProjectDataFileID is missing for project "${projectId}".`);
      return { success: false, error: "Project data file reference is missing.", data: null };
    }
    Logger.log(`getProjectViewData: Found ProjectDataFileID: ${projectDataFileId} for active project ${projectId}.`);

    const jsonContent = readDriveFileContent(projectDataFileId); // From DriveService.gs
    
    Logger.log(`getProjectViewData: Successfully read content for file ID ${projectDataFileId}.`);
    return { success: true, data: jsonContent };

  } catch (e) {
    Logger.log(`Error in getProjectViewData for projectId ${projectId}: ${e.toString()} \nStack: ${e.stack}`);
    return { success: false, error: `Server error fetching project data: ${e.message}`, data: null };
  }
}