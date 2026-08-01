/**
 * SGMS Cloud Edition - Main Entry Point
 * This is the main handler for HTTP requests
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    // Handle CORS preflight
    if (e.parameter.method === 'OPTIONS') {
      return createResponse({ success: true }, 200);
    }
    
    // Parse request
    const params = e.parameter || {};
    const postData = e.postData ? JSON.parse(e.postData.contents) : {};
    const action = postData.action || params.action;
    const payload = postData.payload || params.payload || {};
    const token = postData.token || params.token || e.parameter.authorization;
    
    // Route the request
    return routeRequest(action, payload, token);
    
  } catch (error) {
    Logger.log('Error in handleRequest: ' + error.toString());
    return createResponse({
      success: false,
      message: 'Server error: ' + error.toString(),
      data: null
    }, 500);
  }
}

/**
 * Initialize the SGMS system
 * Run this once to set up the database
 */
function initializeSystem() {
  try {
    Logger.log('Starting SGMS initialization...');
    
    // Create database sheets
    createDatabaseSheets();
    
    // Initialize default settings
    initializeSettings();
    
    // Initialize default grading terms
    initializeGradingTerms();
    
    // Create default admin account
    createDefaultAdmin();
    
    Logger.log('SGMS initialization completed successfully!');
    return 'System initialized successfully';
    
  } catch (error) {
    Logger.log('Initialization error: ' + error.toString());
    throw error;
  }
}

/**
 * Get the Web App URL
 * Run this to get your API endpoint URL
 */
function getWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  Logger.log('Web App URL: ' + url);
  return url;
}