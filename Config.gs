/**
 * SGMS Configuration
 * Central configuration for the entire system
 */

const CONFIG = {
  // Spreadsheet Configuration
  SPREADSHEET_NAME: 'SGMS_DATABASE',
  
  // Sheet Names
  SHEETS: {
    STUDENTS: 'STUDENTS',
    GRADING_TERMS: 'GRADING_TERMS',
    ACTIVITIES: 'ACTIVITIES',
    SCORES: 'SCORES',
    QR_TOKENS: 'QR_TOKENS',
    QR_SESSIONS: 'QR_SESSIONS',
    ADMINS: 'ADMINS',
    AUDIT_LOG: 'AUDIT_LOG',
    SETTINGS: 'SETTINGS'
  },
  
  // Default Grading Terms
  DEFAULT_TERMS: [
    { name: 'Midterm Collective', order: 1, weight: 15, passing: 50 },
    { name: 'Other Midterm', order: 2, weight: 10, passing: 50 },
    { name: 'Midterm Exam', order: 3, weight: 25, passing: 50 },
    { name: 'Final Collective Initial', order: 4, weight: 15, passing: 50 },
    { name: 'Final Collective Final', order: 5, weight: 10, passing: 50 },
    { name: 'Final Exam', order: 6, weight: 25, passing: 50 }
  ],
  
  // Default Settings
  DEFAULT_SETTINGS: {
    SCHOOL_NAME: 'School Name',
    CALCULATION_MODE: 'STRICT', // STRICT or PROGRESS
    OVERALL_PASSING_PERCENT: 50,
    SYSTEM_VERSION: '1.0.0',
    THEME: 'purple',
    // Individual component passing scores (per spec requirement #1)
    MIDTERM_COLLECTIVE_PASSING: 50,
    FINAL_COLLECTIVE_INITIAL_PASSING: 50,
    FINAL_COLLECTIVE_FINAL_PASSING: 50,
    MIDTERM_EXAM_PASSING: 50,
    FINAL_EXAM_PASSING: 50,
    // Cumulative stage passing scores
    STAGE1_PASSING: 40, // Midterm Collective + Midterm Exam
    STAGE2_PASSING: 55, // + Final Collective Initial
    STAGE3_PASSING: 65, // + Final Collective Final
    STAGE4_PASSING: 50  // + Final Exam (Overall)
  },
  
  // Session Configuration
  SESSION: {
    EXPIRY_HOURS: 24,
    QR_SESSION_EXPIRY_MINUTES: 5
  },
  
  // Validation Rules
  VALIDATION: {
    MIN_SCORE: 0,
    STUDENT_ID_PREFIX: 'STD',
    ACTIVITY_ID_PREFIX: 'ACT',
    TERM_ID_PREFIX: 'TERM',
    ADMIN_ID_PREFIX: 'ADM',
    QR_TOKEN_LENGTH: 12
  }
};

/**
 * Get the SGMS database spreadsheet
 * Creates it if it doesn't exist
 */
function getDatabase() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let spreadsheetId = scriptProperties.getProperty('SGMS_DATABASE_ID');
  
  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      // Spreadsheet was deleted, create new one
      spreadsheetId = null;
    }
  }
  
  // Create new spreadsheet
  const ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
  spreadsheetId = ss.getId();
  scriptProperties.setProperty('SGMS_DATABASE_ID', spreadsheetId);
  
  Logger.log('Created new SGMS database: ' + spreadsheetId);
  return ss;
}
