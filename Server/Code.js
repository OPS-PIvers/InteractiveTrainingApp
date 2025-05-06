// Server/Code.gs

// --- CONSTANTS ---
const PROJECT_INDEX_SHEET_ID = '1_a8qB_Vzy5ItoSTqPPrN25jLLz5wnar5SjQ0GLwFFug';
const ROOT_PROJECT_FOLDER_ID = '1_YNn3PPj0Xn4bcN_C1jqHCBC_v4wuLKU';
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
// Server/Code.gs
function doGet(e) { // 'e' is important here
  let htmlOutput;

  // We no longer need to set template.MODE_FROM_SERVER or template.PROJECT_ID_FROM_SERVER here
  // The client will fetch this data using getServerData()

  if (isAdminUser()) {
    const template = HtmlService.createTemplateFromFile('Client/Admin/AdminView.html');
    // We can still determine title based on initial 'e' parameters if needed, or client can set it.
    let title = 'Interactive Training App - Admin Dashboard';
    if (e.parameter.page === 'edit' && e.parameter.projectId) {
        title = 'Interactive Training App - Edit Project';
    }
    htmlOutput = template.evaluate()
        .setTitle(title)
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
    console.error("Error getting active user email: " + e.toString());
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
 * @param {object} params The e.parameter object from the initial doGet call, passed by client if needed.
 *                       Alternatively, the client could parse window.location.search itself.
 * @return {object} An object containing mode and projectId.
 */
function getServerData(initialPageParameters) {
  // initialPageParameters would be like { page: "edit", projectId: "123" } if client sends them
  // For this to work, client needs to parse its own URL and send 'e.parameter' like object
  // OR we rely on the fact that this function is called on page load, and what matters is the initial URL.
  // Let's assume client sends its initial load parameters.

  // To make this robust, the client should parse its current URL and send the relevant params.
  // For now, let's keep your structure, but the values for MODE_FROM_SERVER and PROJECT_ID_FROM_SERVER
  // need to be defined globally or passed into getServerData if they depend on the initial page load 'e' object.

  // Simpler: let the client parse its own URL.
  // The server function can just be a dummy if the client does all the work.
  // But your intention was that the SERVER determines the mode.
  
  // Let's assume the client will pass its initial load's 'page' and 'projectId' parameters
  const page = initialPageParameters ? initialPageParameters.page : null;
  const projectId = initialPageParameters ? initialPageParameters.projectId : null;

  Logger.log(`getServerData called with page: ${page}, projectId: ${projectId}`);

  let modeToReturn = 'list';
  let projectIdToReturn = null;

  if (page === 'edit' && projectId) {
    modeToReturn = 'edit';
    projectIdToReturn = projectId;
  }
  // If page is not 'edit', mode is 'list' and projectId remains null.

  return {
    mode: modeToReturn,
    projectId: projectIdToReturn
  };
}