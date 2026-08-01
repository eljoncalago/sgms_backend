/**
 * SGMS Activity Service
 * Handles grading activities and terms management
 */

/**
 * Handle get grading terms
 */
function handleGetGradingTerms() {
  try {
    const terms = getAllRecords(CONFIG.SHEETS.GRADING_TERMS);
    
    // Sort by order
    terms.sort((a, b) => a.TERM_ORDER - b.TERM_ORDER);
    
    return createResponse({
      success: true,
      message: 'Grading terms retrieved',
      data: terms
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
 * Handle update grading term
 */
function handleUpdateGradingTerm(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);
    const { termId, termName, weight, passingPercent } = payload;
    
    if (!termId) {
      return createResponse({
        success: false,
        message: 'Term ID is required',
        data: null
      });
    }
    
    const updates = {
      UPDATED_AT: new Date().toISOString()
    };
    
    if (termName) updates.TERM_NAME = termName;
    if (weight !== undefined) updates.WEIGHT_PERCENT = weight;
    if (passingPercent !== undefined) updates.PASSING_PERCENT = passingPercent;
    
    updateRecord(CONFIG.SHEETS.GRADING_TERMS, 'TERM_ID', termId, updates);
    
    createAuditLog(adminId, 'UPDATE_TERM', null, null, null, JSON.stringify(updates), 'web', 'Updated term: ' + termId);
    
    return createResponse({
      success: true,
      message: 'Grading term updated successfully',
      data: null
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
 * Initialize default grading terms
 */
function initializeGradingTerms() {
  const existingTerms = getAllRecords(CONFIG.SHEETS.GRADING_TERMS);
  
  if (existingTerms.length > 0) {
    Logger.log('Grading terms already initialized');
    return;
  }
  
  CONFIG.DEFAULT_TERMS.forEach(term => {
    const termRecord = {
      TERM_ID: generateId(CONFIG.VALIDATION.TERM_ID_PREFIX),
      TERM_NAME: term.name,
      TERM_ORDER: term.order,
      WEIGHT_PERCENT: term.weight,
      PASSING_PERCENT: term.passing,
      IS_ACTIVE: true,
      CREATED_AT: new Date().toISOString(),
      UPDATED_AT: new Date().toISOString()
    };
    
    insertRecord(CONFIG.SHEETS.GRADING_TERMS, termRecord);
  });
  
  Logger.log('Default grading terms initialized');
}

/**
 * Handle get activities
 */
function handleGetActivities(payload) {
  try {
    let activities = getAllRecords(CONFIG.SHEETS.ACTIVITIES);
    
    // Filter by term if specified
    if (payload.termId) {
      activities = activities.filter(a => a.TERM_ID === payload.termId);
    }
    
    // Filter by active status
    if (payload.activeOnly) {
      activities = activities.filter(a => a.IS_ACTIVE);
    }
    
    // Sort by order
    activities.sort((a, b) => a.ACTIVITY_ORDER - b.ACTIVITY_ORDER);
    
    return createResponse({
      success: true,
      message: 'Activities retrieved',
      data: activities
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
 * Handle get single activity
 */
function handleGetActivity(payload) {
  try {
    const { activityId } = payload;
    
    if (!activityId) {
      return createResponse({
        success: false,
        message: 'Activity ID is required',
        data: null
      });
    }
    
    const activity = findRecordById(CONFIG.SHEETS.ACTIVITIES, 'ACTIVITY_ID', activityId);
    
    if (!activity) {
      return createResponse({
        success: false,
        message: 'Activity not found',
        data: null
      });
    }
    
    return createResponse({
      success: true,
      message: 'Activity retrieved',
      data: activity
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
 * Handle create activity
 */
function handleCreateActivity(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);
    const { termId, activityName, activityType, maxScore } = payload;
    
    if (!termId || !activityName || !maxScore) {
      return createResponse({
        success: false,
        message: 'Term ID, activity name, and max score are required',
        data: null
      });
    }
    
    // Verify term exists
    const term = findRecordById(CONFIG.SHEETS.GRADING_TERMS, 'TERM_ID', termId);
    if (!term) {
      return createResponse({
        success: false,
        message: 'Grading term not found',
        data: null
      });
    }
    
    // Get next order number for this term
    const termActivities = findRecords(CONFIG.SHEETS.ACTIVITIES, { TERM_ID: termId });
    const maxOrder = termActivities.reduce((max, a) => Math.max(max, a.ACTIVITY_ORDER || 0), 0);
    
    const activity = {
      ACTIVITY_ID: generateId(CONFIG.VALIDATION.ACTIVITY_ID_PREFIX),
      TERM_ID: termId,
      ACTIVITY_NAME: activityName,
      ACTIVITY_TYPE: activityType || 'General',
      MAX_SCORE: maxScore,
      ACTIVITY_ORDER: maxOrder + 1,
      IS_ACTIVE: true,
      CREATED_AT: new Date().toISOString(),
      UPDATED_AT: new Date().toISOString()
    };
    
    insertRecord(CONFIG.SHEETS.ACTIVITIES, activity);
    
    createAuditLog(
      adminId,
      'CREATE_ACTIVITY',
      null,
      activity.ACTIVITY_ID,
      null,
      JSON.stringify(activity),
      'web',
      'Created activity: ' + activityName
    );
    
    return createResponse({
      success: true,
      message: 'Activity created successfully',
      data: activity
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
 * Handle update activity
 */
function handleUpdateActivity(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);
    const { activityId, activityName, activityType, maxScore, isActive } = payload;
    
    if (!activityId) {
      return createResponse({
        success: false,
        message: 'Activity ID is required',
        data: null
      });
    }
    
    const updates = {
      UPDATED_AT: new Date().toISOString()
    };
    
    if (activityName) updates.ACTIVITY_NAME = activityName;
    if (activityType) updates.ACTIVITY_TYPE = activityType;
    if (maxScore !== undefined) updates.MAX_SCORE = maxScore;
    if (isActive !== undefined) updates.IS_ACTIVE = isActive;
    
    updateRecord(CONFIG.SHEETS.ACTIVITIES, 'ACTIVITY_ID', activityId, updates);
    
    createAuditLog(
      adminId,
      'UPDATE_ACTIVITY',
      null,
      activityId,
      null,
      JSON.stringify(updates),
      'web',
      'Updated activity: ' + activityId
    );
    
    return createResponse({
      success: true,
      message: 'Activity updated successfully',
      data: null
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
 * Handle delete activity
 */
function handleDeleteActivity(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);
    const { activityId } = payload;
    
    if (!activityId) {
      return createResponse({
        success: false,
        message: 'Activity ID is required',
        data: null
      });
    }
    
    // Soft delete
    const updates = {
      IS_ACTIVE: false,
      UPDATED_AT: new Date().toISOString()
    };
    
    updateRecord(CONFIG.SHEETS.ACTIVITIES, 'ACTIVITY_ID', activityId, updates);
    
    createAuditLog(
      adminId,
      'DELETE_ACTIVITY',
      null,
      activityId,
      null,
      null,
      'web',
      'Deleted activity: ' + activityId
    );
    
    return createResponse({
      success: true,
      message: 'Activity deleted successfully',
      data: null
    });
  } catch (error) {
    return createResponse({
      success: false,
      message: error.toString(),
      data: null
    });
  }
}