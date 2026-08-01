/**
 * SGMS Audit Service
 * Handles system audit logging
 */

function createAuditLog(adminId, action, studentId, activityId, oldValue, newValue, source, details) {
  try {
    const auditRecord = {
      LOG_ID: generateId('LOG'),
      TIMESTAMP: new Date().toISOString(),
      ADMIN_ID: adminId || 'SYSTEM',
      ACTION: action,
      STUDENT_ID: studentId || '',
      ACTIVITY_ID: activityId || '',
      OLD_VALUE: oldValue || '',
      NEW_VALUE: newValue || '',
      SOURCE: source || 'web',
      DETAILS: details || ''
    };
    
    insertRecord(CONFIG.SHEETS.AUDIT_LOG, auditRecord);
    return auditRecord;
  } catch (error) {
    Logger.log('Audit log error: ' + error.toString());
    // Don't throw error - audit logging should not break main operations
    return null;
  }
}

function handleGetAuditLog(payload) {
  try {
    let logs = getAllRecords(CONFIG.SHEETS.AUDIT_LOG);
    
    // Apply filters
    if (payload.adminId) {
      logs = logs.filter(l => l.ADMIN_ID === payload.adminId);
    }
    
    if (payload.action) {
      logs = logs.filter(l => l.ACTION === payload.action);
    }
    
    if (payload.studentId) {
      logs = logs.filter(l => l.STUDENT_ID === payload.studentId);
    }
    
    if (payload.startDate) {
      logs = logs.filter(l => new Date(l.TIMESTAMP) >= new Date(payload.startDate));
    }
    
    if (payload.endDate) {
      logs = logs.filter(l => new Date(l.TIMESTAMP) <= new Date(payload.endDate));
    }
    
    // Sort by timestamp descending (newest first)
    logs.sort((a, b) => new Date(b.TIMESTAMP) - new Date(a.TIMESTAMP));
    
    // Limit results
    const limit = payload.limit || 100;
    logs = logs.slice(0, limit);
    
    return createResponse({
      success: true,
      message: 'Audit log retrieved',
      data: logs
    });
  } catch (error) {
    return createResponse({
      success: false,
      message: error.toString(),
      data: null
    });
  }
}