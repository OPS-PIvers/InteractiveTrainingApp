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
      // Assuming getAllSheetData returns an array of objects where keys are header names
      // If it returns an array of arrays, use indices like project[COL_STATUS - 1]
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
    // Return an empty array in case of error to prevent client-side issues
    return [];
  }
}