/**
 * SGMS Database Service
 * Handles all database operations with Google Sheets
 */

/**
 * Get a sheet by name, create if doesn't exist
 */
function getSheet(sheetName) {
  const ss = getDatabase();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log('Created sheet: ' + sheetName);
  }
  
  return sheet;
}

/**
 * Get all records from a sheet
 */
function getAllRecords(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return [];
  }
  
  const headers = data[0];
  const records = [];
  
  for (let i = 1; i < data.length; i++) {
    const record = {};
    for (let j = 0; j < headers.length; j++) {
      record[headers[j]] = data[i][j];
    }
    records.push(record);
  }
  
  return records;
}

/**
 * Find a record by ID
 */
function findRecordById(sheetName, idField, id) {
  const records = getAllRecords(sheetName);
  // FIX: coerce both sides to string before comparing. Google Sheets can
  //      store the same ID as a number in one cell and a string in another,
  //      and strict === would then fail to match (e.g. 12345 !== '12345').
  //      This caused "Student not found" during score import for some rows.
  var idStr = String(id);
  return records.find(record => String(record[idField]) === idStr) || null;
}

/**
 * Find records by criteria
 */
function findRecords(sheetName, criteria) {
  const records = getAllRecords(sheetName);
  
  // FIX: same string-coercion as findRecordById — see comment there.
  //      This is used by handleSaveScore to find existing scores; strict
  //      === missed matches when the sheet stored IDs as a different type.
  var criteriaStr = {};
  for (let key in criteria) {
    criteriaStr[key] = String(criteria[key]);
  }
  
  return records.filter(record => {
    for (let key in criteriaStr) {
      if (String(record[key]) !== criteriaStr[key]) {
        return false;
      }
    }
    return true;
  });
}

/**
 * Insert a new record
 */
function insertRecord(sheetName, record) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  
  // If sheet is empty, add headers
  if (data.length === 0 || !data[0] || data[0].length === 0) {
    const headers = Object.keys(record);
    sheet.appendRow(headers);
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // FIX: use nullish coalescing (??) instead of || so a valid score of 0
  //      is not silently replaced with an empty string. The old `|| ''`
  //      treated 0, false, and '' all as falsy and dropped them.
  const row = headers.map(header => {
    const val = record[header];
    return val === undefined || val === null ? '' : val;
  });
  
  sheet.appendRow(row);
  return record;
}

/**
 * Update a record
 */
function updateRecord(sheetName, idField, id, updates) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    throw new Error('No records found');
  }
  
  const headers = data[0];
  const idIndex = headers.indexOf(idField);
  
  if (idIndex === -1) {
    throw new Error('ID field not found: ' + idField);
  }
  
  // Find the row
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === id) {
      // Update the row
      const updatedRow = [...data[i]];
      
      for (let key in updates) {
        const colIndex = headers.indexOf(key);
        if (colIndex !== -1) {
          updatedRow[colIndex] = updates[key];
        }
      }
      
      sheet.getRange(i + 1, 1, 1, updatedRow.length).setValues([updatedRow]);
      
      // Return updated record
      const record = {};
      headers.forEach((header, idx) => {
        record[header] = updatedRow[idx];
      });
      
      return record;
    }
  }
  
  throw new Error('Record not found: ' + id);
}

/**
 * Delete a record
 */
function deleteRecord(sheetName, idField, id) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    throw new Error('No records found');
  }
  
  const headers = data[0];
  const idIndex = headers.indexOf(idField);
  
  if (idIndex === -1) {
    throw new Error('ID field not found: ' + idField);
  }
  
  // Find and delete the row
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  
  throw new Error('Record not found: ' + id);
}

/**
 * Generate a unique ID
 */
function generateId(prefix) {
  const timestamp = new Date().getTime().toString().slice(-8);
  const random = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return prefix + timestamp + random;
}

/**
 * Create all database sheets with headers
 */
function createDatabaseSheets() {
  // Students sheet
  createSheetWithHeaders(CONFIG.SHEETS.STUDENTS, [
    'STUDENT_ID', 'THAI_NAME', 'ENGLISH_NAME', 'GRADE_LEVEL', 
    'SECTION_NUMBER', 'CLASS_NUMBER', 'STATUS', 'CREATED_AT', 'UPDATED_AT'
  ]);
  
  // Grading Terms sheet
  createSheetWithHeaders(CONFIG.SHEETS.GRADING_TERMS, [
    'TERM_ID', 'TERM_NAME', 'TERM_ORDER', 'WEIGHT_PERCENT', 
    'PASSING_PERCENT', 'IS_ACTIVE', 'CREATED_AT', 'UPDATED_AT'
  ]);
  
  // Activities sheet
  // FIX: added GRADE_LEVEL column so activities can be grade-level-specific.
  //      Without this column insertRecord() silently drops the GRADE_LEVEL
  //      value and every activity is saved as "All Grades".
  //      ensureSheetColumns() also adds it to EXISTING databases that were
  //      created before this column existed.
  ensureSheetColumns(CONFIG.SHEETS.ACTIVITIES, [
    'ACTIVITY_ID', 'TERM_ID', 'ACTIVITY_NAME', 'ACTIVITY_TYPE',
    'MAX_SCORE', 'GRADE_LEVEL', 'ACTIVITY_ORDER', 'IS_ACTIVE', 'CREATED_AT', 'UPDATED_AT'
  ]);
  
  // Scores sheet
  createSheetWithHeaders(CONFIG.SHEETS.SCORES, [
    'SCORE_ID', 'STUDENT_ID', 'ACTIVITY_ID', 'RAW_SCORE', 
    'RECORDED_BY', 'RECORD_SOURCE', 'CREATED_AT', 'UPDATED_AT'
  ]);
  
  // QR Tokens sheet
  createSheetWithHeaders(CONFIG.SHEETS.QR_TOKENS, [
    'TOKEN_ID', 'TOKEN', 'STUDENT_ID', 'IS_ACTIVE', 
    'CREATED_AT', 'EXPIRES_AT', 'LAST_ACCESSED_AT'
  ]);
  
  // QR Sessions sheet
  // NEW (continuous scanning): SCAN_COUNT / LAST_SCAN_AT / SCANNER_LAST_SEEN
  // let the main device detect *each* new scan (not just the first pairing)
  // and know whether the scanner device is still connected.
  createSheetWithHeaders(CONFIG.SHEETS.QR_SESSIONS, [
    'SESSION_ID', 'CREATED_BY', 'DEVICE_ID', 'STATUS', 'STUDENT_ID',
    'CREATED_AT', 'UPDATED_AT', 'EXPIRES_AT', 'LAST_ACTIVITY',
    'SCAN_COUNT', 'LAST_SCAN_AT', 'SCANNER_LAST_SEEN'
  ]);
  
  // Admins sheet
  createSheetWithHeaders(CONFIG.SHEETS.ADMINS, [
    'ADMIN_ID', 'ADMIN_NAME', 'PASSCODE_HASH', 'IS_ACTIVE', 
    'CREATED_AT', 'UPDATED_AT'
  ]);
  
  // Audit Log sheet
  createSheetWithHeaders(CONFIG.SHEETS.AUDIT_LOG, [
    'LOG_ID', 'TIMESTAMP', 'ADMIN_ID', 'ACTION', 'STUDENT_ID', 
    'ACTIVITY_ID', 'OLD_VALUE', 'NEW_VALUE', 'SOURCE', 'DETAILS'
  ]);
  
  // Settings sheet
  createSheetWithHeaders(CONFIG.SHEETS.SETTINGS, [
    'SETTING_KEY', 'SETTING_VALUE', 'DESCRIPTION'
  ]);
  
  Logger.log('All database sheets created');
}

/**
 * Ensure a sheet has the given columns — appends any that are missing.
 *
 * NEW: existing databases were created before the continuous-scan columns
 * existed. insertRecord()/updateRecord() silently drop fields whose column is
 * absent, so this is called before writing the new QR_SESSIONS fields.
 */
function ensureSheetColumns(sheetName, columns) {
  const sheet = getSheet(sheetName);

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.appendRow(columns);
    sheet.getRange(1, 1, 1, columns.length).setFontWeight('bold');
    return columns;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = columns.filter(function (c) { return headers.indexOf(c) === -1; });

  if (missing.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, headers.length + 1, 1, missing.length).setFontWeight('bold');
  }

  return headers.concat(missing);
}

/**
 * Create a sheet with headers
 */
function createSheetWithHeaders(sheetName, headers) {
  const sheet = getSheet(sheetName);
  
  // Check if headers already exist
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    Logger.log('Added headers to: ' + sheetName);
  }
}