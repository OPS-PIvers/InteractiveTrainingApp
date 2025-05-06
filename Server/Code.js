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
  // Placeholder for Step 2: User Auth/Routing
  // This will be expanded in Step 5 for page routing (edit vs list)

  // Determine if accessing edit page
  const page = e.parameter.page;
  const projectId = e.parameter.projectId;

  if (isAdminUser()) {
    if (page === 'edit' && projectId) {
        // Admin wants to edit a specific project
        const template = HtmlService.createTemplateFromFile('Client/Admin/AdminView.html');
        template.projectId = projectId; // Pass projectId to the template
        template.mode = 'edit';
        return template.evaluate()
            .setTitle('Interactive Training App - Admin Edit')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
        // Admin wants the main admin dashboard (project list, create new)
        const template = HtmlService.createTemplateFromFile('Client/Admin/AdminView.html');
        template.projectId = null; // No specific project to edit initially
        template.mode = 'list'; // Default to list view
        return template.evaluate() // <--- USE .evaluate()
            .setTitle('Interactive Training App - Admin')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } else {
    // Viewer view
    // Add similar logic for viewer project view if needed: ?page=view&projectId=...
    const template = HtmlService.createTemplateFromFile('Client/Viewer/ViewerView.html');
    // if (page === 'view' && projectId) {
    //    template.projectId = projectId;
    // }
    return template.evaluate() // <--- USE .evaluate()
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