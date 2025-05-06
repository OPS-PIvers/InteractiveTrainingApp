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
// ...
function doGet(e) {
  const pageParam = e.parameter.page; // Use different names to avoid confusion
  const projectIdParam = e.parameter.projectId;
  let htmlOutput;

  if (isAdminUser()) {
    const template = HtmlService.createTemplateFromFile('Client/Admin/AdminView.html');

    // Explicitly prepare the values for the template
    // Ensure they are either a string or actual JavaScript null
    let modeForClient = 'list'; // Default
    let projectIdForClient = null; // Default

    if (pageParam === 'edit' && projectIdParam) {
      modeForClient = 'edit'; // String
      projectIdForClient = projectIdParam; // String
    }
    // If pageParam is not 'edit', modeForClient remains 'list' and projectIdForClient remains null.

    template.MODE_FROM_SERVER = modeForClient; 
    template.PROJECT_ID_FROM_SERVER = projectIdForClient; 
    
    // Log what's being set on the template object BEFORE evaluate()
    Logger.log(`Server doGet: Setting template.MODE_FROM_SERVER = ${template.MODE_FROM_SERVER} (type: ${typeof template.MODE_FROM_SERVER})`);
    Logger.log(`Server doGet: Setting template.PROJECT_ID_FROM_SERVER = ${template.PROJECT_ID_FROM_SERVER} (type: ${typeof template.PROJECT_ID_FROM_SERVER})`);

    htmlOutput = template.evaluate()
        .setTitle(modeForClient === 'edit' ? 'Interactive Training App - Edit Project' : 'Interactive Training App - Admin Dashboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    const template = HtmlService.createTemplateFromFile('Client/Viewer/ViewerView.html');
    // In case you add similar params for viewer later:
    // template.MODE_FROM_SERVER = e.parameter.page || 'list'; 
    // template.PROJECT_ID_FROM_SERVER = e.parameter.projectId || null;
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

function getServerData() {
  return {
    mode: MODE_FROM_SERVER,
    projectId: PROJECT_ID_FROM_SERVER
  };
}