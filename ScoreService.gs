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
    
    if (!studentId || !activityId || rawScore === undefined) {
      return createResponse({
        success: false,
        message: 'Student ID, activity ID, and raw score are required',
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
    if (rawScore < CONFIG.VALIDATION.MIN_SCORE || rawScore > activity.MAX_SCORE) {
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
        RAW_SCORE: rawScore,
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
        String(rawScore),
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
        RAW_SCORE: rawScore,
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
        String(rawScore),
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
    if (payload.activityId) {
      scores = scores.filter(s => s.ACTIVITY_ID === payload.activityId);
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
