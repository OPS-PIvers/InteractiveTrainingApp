// Server/Code.gs

// --- CONSTANTS ---
const PROJECT_INDEX_SHEET_ID = '1_a8qB_Vzy5ItoSTqPPrN25jLLz5wnar5SjQ0GLwFFug';
const ROOT_PROJECT_FOLDER_ID = '1_YNn3PPj0Xn4bcN_C1jqHCBC_v4wuLKU'; // Replace with your actual Root Folder ID
const ADMIN_EMAILS = ['paul.ivers@orono.k12.mn.us']; // Add more admin emails here as needed

const PROJECT_INDEX_DATA_SHEET_NAME = "Projects"; // The name of the sheet/tab within your ProjectIndex Spreadsheet file

// Column indices for the "Projects" sheet (1-based)
const COL_PROJECT_ID = 1;
const COL_PROJECT_TITLE = 2;
const COL_PROJECT_FOLDER_ID = 3;
const COL_STATUS = 4;
// const COL_ADMIN_EMAILS_PER_PROJECT = 5; // This column is optional if using global ADMIN_EMAILS. If used, adjust subsequent column numbers.
const COL_PROJECT_DATA_FILE_ID = 5; // Assuming AdminEmails per project column is omitted
const COL_LAST_MODIFIED = 6;
const COL_CREATED_DATE = 7;

// Default project data filename
const PROJECT_DATA_FILENAME = 'project_data.json';

// --- MAIN WEB APP ENTRY POINT ---
/**
 * Main entry point for the web application.
 * This function will be expanded in later steps to handle routing.
 * @param {Object} e The event parameter.
 * @return {HtmlOutput} The HTML output to serve.
 */
function doGet(e) {
  let htmlOutput;

  if (isAdminUser()) {
    const template = HtmlService.createTemplateFromFile('Client/Admin/AdminView.html');
    // No need to pass mode/projectId for view switching anymore
    htmlOutput = template.evaluate()
        .setTitle('Interactive Training App - Admin Dashboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    const template = HtmlService.createTemplateFromFile('Client/Viewer/ViewerView.html');
    htmlOutput = template.evaluate()
      .setTitle('Interactive Training App - Viewer')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return htmlOutput;
}

// --- USER AUTHENTICATION ---
/**
 * Checks if the current active user is an administrator.
 * @return {boolean} True if the user is an admin, false otherwise.
 */
function isAdminUser() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    return ADMIN_EMAILS.indexOf(currentUserEmail) !== -1;
  } catch (e) {
    // If Session.getActiveUser().getEmail() fails (e.g. script run by anonymous user or in a context without user session)
    Logger.log(`Error getting active user email: ${e.toString()}. Stack: ${e.stack ? e.stack : 'N/A'}`);
    return false;
  }
}

// --- UTILITY FUNCTIONS (Can be expanded or moved later) ---
/**
 * Includes HTML content from one file into another, allowing for templating.
 * @param {string} filename The name of the HTML file to include.
 * @return {string} The content of the HTML file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets the deployed web app's URL.
 * @return {string} The web app URL.
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Called by the client to get initial mode and projectId based on URL parameters
 * that were present when the page initially loaded.
 * @param {object} initialPageParameters The e.parameter object from the initial doGet call, passed by client if needed.
 *                       Alternatively, the client could parse window.location.search itself.
 * @return {object} An object containing mode and projectId.
 */
// Server/Code.gs
function getServerData(initialPageParameters) {
  // initialPageParameters would be like { page: "edit", projectId: "123" }
  const page = initialPageParameters ? initialPageParameters.page : null;
  const projectId = initialPageParameters ? initialPageParameters.projectId : null;

  Logger.log(`getServerData called with received client params: page: ${page}, projectId: ${projectId}`); // ADD THIS LOG

  let modeToReturn = 'list';
  let projectIdToReturn = null; // Important: default to null

  if (page === 'edit' && projectId) { // Check if both are present and page is 'edit'
    modeToReturn = 'edit';
    projectIdToReturn = projectId; // Pass along the projectId
  }
  // If page is not 'edit', or if projectId is missing for 'edit' mode,
  // it defaults to 'list' mode and projectIdToReturn remains null.

  Logger.log(`getServerData returning: mode: ${modeToReturn}, projectId: ${projectIdToReturn}`); // ADD THIS LOG

  return {
    mode: modeToReturn,
    projectId: projectIdToReturn
  };
}