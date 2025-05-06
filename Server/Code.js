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
function doGet(e) {
  // For now, let's return a simple message. This will be replaced in Step 2.
  // return HtmlService.createHtmlOutput("Welcome to the Interactive Training App. Setup in progress.");

  // Placeholder for Step 2: User Auth/Routing
  if (isAdminUser()) {
    // Later, this will serve AdminView.html
    return HtmlService.createHtmlOutputFromFile('Client/Admin/AdminView.html')
        .setTitle('Interactive Training App - Admin')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Required for Drive Picker API later
  } else {
    // Later, this will serve ViewerView.html
    return HtmlService.createHtmlOutputFromFile('Client/Viewer/ViewerView.html')
        .setTitle('Interactive Training App - Viewer')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
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