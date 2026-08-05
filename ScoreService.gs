/**
 * SGMS Score Service
 * Handles all score management operations
 */

/**
 * Handle save score
 */
function handleSaveScore(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);
    const { studentId, activityId, rawScore } = payload;

    // FIX: Excel import can send rawScore as a string (e.g. "85" or even
    //      "85.0"). Coerce to a number so the range check below and the
    //      value written to the sheet are always numeric. Without this,
    //      a string "0" could pass the `=== undefined` guard but then fail
    //      the numeric comparison, or be written as the string "0" which
    //      later reads back as 0 in some contexts and '' in others.
    const numericScore = Number(rawScore);

    if (!studentId || !activityId || rawScore === undefined || rawScore === null || rawScore === '') {
      return createResponse({
        success: false,
        message: 'Student ID, activity ID, and raw score are required',
        data: null
      });
    }

    if (isNaN(numericScore)) {
      return createResponse({
        success: false,
        message: 'Raw score must be a number',
        data: null
      });
    }
    
    // Verify student exists
    const student = findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', studentId);
    if (!student) {
      return createResponse({
        success: false,
        message: 'Student not found',
        data: null
      });
    }
    
    // Verify activity exists
    const activity = findRecordById(CONFIG.SHEETS.ACTIVITIES, 'ACTIVITY_ID', activityId);
    if (!activity) {
      return createResponse({
        success: false,
        message: 'Activity not found',
        data: null
      });
    }
    
    // Validate score
    if (numericScore < CONFIG.VALIDATION.MIN_SCORE || numericScore > Number(activity.MAX_SCORE)) {
      return createResponse({
        success: false,
        message: `Score must be between ${CONFIG.VALIDATION.MIN_SCORE} and ${activity.MAX_SCORE}`,
        data: null
      });
    }
    
    // Check for existing score
    const existingScores = findRecords(CONFIG.SHEETS.SCORES, {
      STUDENT_ID: studentId,
      ACTIVITY_ID: activityId
    });
    
    if (existingScores.length > 0) {
      // Update existing score
      const oldScore = existingScores[0];
      const updates = {
        RAW_SCORE: numericScore,
        RECORDED_BY: adminId,
        UPDATED_AT: new Date().toISOString()
      };
      
      updateRecord(CONFIG.SHEETS.SCORES, 'SCORE_ID', oldScore.SCORE_ID, updates);
      
      createAuditLog(
        adminId,
        'UPDATE_SCORE',
        studentId,
        activityId,
        String(oldScore.RAW_SCORE),
        String(numericScore),
        'web',
        `Updated score for ${student.ENGLISH_NAME} in ${activity.ACTIVITY_NAME}`
      );
      
      return createResponse({
        success: true,
        message: 'Score updated successfully',
        data: { ...oldScore, ...updates }
      });
    } else {
      // Create new score
      const score = {
        SCORE_ID: generateId('SCORE'),
        STUDENT_ID: studentId,
        ACTIVITY_ID: activityId,
        RAW_SCORE: numericScore,
        RECORDED_BY: adminId,
        RECORD_SOURCE: 'web',
        CREATED_AT: new Date().toISOString(),
        UPDATED_AT: new Date().toISOString()
      };
      
      insertRecord(CONFIG.SHEETS.SCORES, score);
      
      createAuditLog(
        adminId,
        'SAVE_SCORE',
        studentId,
        activityId,
        null,
        String(numericScore),
        'web',
        `Saved score for ${student.ENGLISH_NAME} in ${activity.ACTIVITY_NAME}`
      );
      
      return createResponse({
        success: true,
        message: 'Score saved successfully',
        data: score
      });
    }
  } catch (error) {
    return createResponse({
      success: false,
      message: error.toString(),
      data: null
    });
  }
}

/**
 * Handle bulk save scores
 */
function handleBulkSaveScores(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);
    const { scores } = payload;
    
    if (!scores || !Array.isArray(scores)) {
      return createResponse({
        success: false,
        message: 'Scores array is required',
        data: null
      });
    }
    
    const results = [];
    const errors = [];
    
    scores.forEach((scoreData, index) => {
      try {
        const result = handleSaveScore(scoreData, token);
        const parsedResult = JSON.parse(result.getContent());
        
        if (parsedResult.success) {
          results.push(parsedResult.data);
        } else {
          errors.push({ index, error: parsedResult.message });
        }
      } catch (error) {
        errors.push({ index, error: error.toString() });
      }
    });
    
    return createResponse({
      success: true,
      message: `Saved ${results.length} scores, ${errors.length} errors`,
      data: {
        saved: results.length,
        errors: errors
      }
    });
  } catch (error) {
    return createResponse({
      success: false,
      message: error.toString(),
      data: null
    });
  }
}

/**
 * Handle get scores
 */
function handleGetScores(payload) {
  try {
    let scores = getAllRecords(CONFIG.SHEETS.SCORES);
    
    // Filter by activity if specified
    // FIX: string coercion — see findRecordById comment. The activityId
    //      coming from the frontend is a string; the sheet value may not be.
    if (payload.activityId) {
      var actIdStr = String(payload.activityId);
      scores = scores.filter(s => String(s.ACTIVITY_ID) === actIdStr);
    }
    
    return createResponse({
      success: true,
      message: 'Scores retrieved',
      data: scores
    });
  } catch (error) {
    return createResponse({
      success: false,
      message: error.toString(),
      data: null
    });
  }
}

/**
 * Handle get student scores
 */
function handleGetStudentScores(payload) {
  try {
    const { studentId } = payload;
    
    if (!studentId) {
      return createResponse({
        success: false,
        message: 'Student ID is required',
        data: null
      });
    }
    
    const scores = findRecords(CONFIG.SHEETS.SCORES, { STUDENT_ID: studentId });
    
    return createResponse({
      success: true,
      message: 'Student scores retrieved',
      data: scores
    });
  } catch (error) {
    return createResponse({
      success: false,
      message: error.toString(),
      data: null
    });
  }
}
