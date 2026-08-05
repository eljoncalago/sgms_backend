/**
 * SGMS Import/Export Service
 * Handles bulk data operations
 */

function handleImportStudents(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);

    // FIX: the frontend (ImportExport.jsx) sends { students, mode } where each
    // row is keyed exactly like the STUDENTS sheet columns (STUDENT_ID,
    // THAI_NAME, ENGLISH_NAME, GRADE_LEVEL, SECTION_NUMBER, CLASS_NUMBER,
    // STATUS) — the same headers as the downloadable Excel template. This
    // used to read payload.csvData (which was never sent — the field never
    // existed) and camelCase keys (thaiName, englishName, ...) that never
    // matched the sheet-style headers actually uploaded, so every import
    // silently produced 0 created / 0 updated records.
    const students = payload.students || payload.csvData;
    const mode = payload.mode;
    
    if (!students || !Array.isArray(students)) {
      return createResponse({ success: false, message: 'Students array is required', data: null });
    }
    
    const results = { created: 0, updated: 0, errors: [] };
    
    students.forEach((row, index) => {
      try {
        const thaiName     = row.THAI_NAME || row.thaiName;
        const englishName  = row.ENGLISH_NAME || row.englishName;
        const gradeLevel   = row.GRADE_LEVEL || row.gradeLevel;
        const sectionNumber= row.SECTION_NUMBER || row.sectionNumber;
        const classNumber  = row.CLASS_NUMBER || row.classNumber;
        const studentId    = row.STUDENT_ID || row.studentId;
        const status       = row.STATUS || row.status || 'Active';
        
        if (!thaiName || !englishName || !gradeLevel) {
          results.errors.push({ index, error: 'Missing required fields (THAI_NAME, ENGLISH_NAME, GRADE_LEVEL)' });
          return;
        }
        
        if (studentId) {
          const existing = findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', studentId);
          if (existing) {
            updateRecord(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', studentId, {
              THAI_NAME: thaiName,
              ENGLISH_NAME: englishName,
              GRADE_LEVEL: gradeLevel,
              SECTION_NUMBER: sectionNumber,
              CLASS_NUMBER: classNumber,
              STATUS: status,
              UPDATED_AT: new Date().toISOString()
            });
            results.updated++;
            return;
          }
        }
        
        // Create new student — keep the STUDENT_ID from the Excel file when
        // provided; only generate one when the row leaves it blank.
        const newStudent = {
          STUDENT_ID: studentId || generateId(CONFIG.VALIDATION.STUDENT_ID_PREFIX),
          THAI_NAME: thaiName,
          ENGLISH_NAME: englishName,
          GRADE_LEVEL: gradeLevel,
          SECTION_NUMBER: sectionNumber || '1',
          CLASS_NUMBER: classNumber || '1',
          STATUS: status,
          CREATED_AT: new Date().toISOString(),
          UPDATED_AT: new Date().toISOString()
        };
        
        insertRecord(CONFIG.SHEETS.STUDENTS, newStudent);
        results.created++;
      } catch (error) {
        results.errors.push({ index, error: error.toString() });
      }
    });
    
    createAuditLog(
      adminId,
      'IMPORT_STUDENTS',
      null,
      null,
      null,
      JSON.stringify(results),
      'web',
      `Imported students: ${results.created} created, ${results.updated} updated`
    );
    
    return createResponse({
      success: true,
      message: 'Import completed',
      data: results
    });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

function handleExportStudents(payload) {
  try {
    const { gradeLevel, section } = payload;
    
    let students = getAllRecords(CONFIG.SHEETS.STUDENTS);
    
    if (gradeLevel) {
      students = students.filter(s => s.GRADE_LEVEL == gradeLevel);
    }
    
    if (section) {
      students = students.filter(s => s.SECTION_NUMBER == section);
    }
    
    return createResponse({
      success: true,
      message: 'Students exported',
      data: students
    });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

function handleImportScores(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);

    // The frontend sends { scores, mode, activityId } where each row is keyed
    // like the score-import Excel template. The template is now GENERATED for a
    // chosen grade level / section / activity, so every row already carries the
    // right ACTIVITY_ID; payload.activityId is used as a fallback when a row
    // leaves that column blank.
    const scores = payload.scores || payload.csvData;
    const fallbackActivityId = payload.activityId || '';

    if (!scores || !Array.isArray(scores)) {
      return createResponse({ success: false, message: 'Scores array is required', data: null });
    }

    const results = { saved: 0, skipped: 0, errors: [] };
    const activityCache = {};

    scores.forEach(function (row, index) {
      try {
        const studentId = String(row.STUDENT_ID || row.studentId || '').trim();
        const activityId = String(row.ACTIVITY_ID || row.activityId || fallbackActivityId || '').trim();
        const rawValue = row.RAW_SCORE !== undefined ? row.RAW_SCORE : row.rawScore;
        const rawText = String(rawValue === undefined || rawValue === null ? '' : rawValue).trim();

        // Blank score = "not graded yet" — skipped, not an error. Teachers hand
        // the same sheet back with only part of the class filled in.
        if (rawText === '') {
          results.skipped++;
          return;
        }

        if (!studentId || !activityId) {
          results.errors.push({ index: index, studentId: studentId, error: 'Missing STUDENT_ID or ACTIVITY_ID' });
          return;
        }

        const rawScore = parseFloat(rawText);
        if (isNaN(rawScore)) {
          results.errors.push({ index: index, studentId: studentId, error: 'RAW_SCORE is not a number' });
          return;
        }

        if (!activityCache[activityId]) {
          activityCache[activityId] = findRecordById(CONFIG.SHEETS.ACTIVITIES, 'ACTIVITY_ID', activityId) || 'MISSING';
        }
        const activity = activityCache[activityId];
        if (activity === 'MISSING') {
          results.errors.push({ index: index, studentId: studentId, error: 'Unknown ACTIVITY_ID: ' + activityId });
          return;
        }

        const maxScore = parseFloat(activity.MAX_SCORE);
        if (rawScore < CONFIG.VALIDATION.MIN_SCORE || rawScore > maxScore) {
          results.errors.push({
            index: index,
            studentId: studentId,
            error: 'Score ' + rawScore + ' is outside 0–' + maxScore + ' for ' + activity.ACTIVITY_NAME
          });
          return;
        }

        const result = handleSaveScore({ studentId: studentId, activityId: activityId, rawScore: rawScore }, token);
        const parsed = JSON.parse(result.getContent());

        if (parsed.success) {
          results.saved++;
        } else {
          results.errors.push({ index: index, studentId: studentId, error: parsed.message });
        }
      } catch (error) {
        results.errors.push({ index: index, error: error.toString() });
      }
    });

    createAuditLog(
      adminId,
      'IMPORT_SCORES',
      null,
      fallbackActivityId || null,
      null,
      JSON.stringify(results),
      'web',
      'Imported ' + results.saved + ' scores (' + results.skipped + ' blank rows skipped)'
    );

    return createResponse({ success: true, message: 'Import completed', data: results });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}
