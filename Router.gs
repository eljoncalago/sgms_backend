/**
 * SGMS API Router
 * Routes all incoming requests to appropriate services.
 *
 * IMPORTANT: This file was originally named "Router,gs" (with a comma) in the
 * source zip — that is a typo. Rename it to exactly "Router.gs" when copying
 * into Google Apps Script.
 */

function routeRequest(action, payload, token) {
  try {
    // Public endpoints (no authentication required)
    // FIX: 'updateQRSession' MUST be public. This action is called from the
    // *second* (unauthenticated) device after it scans the session QR shown
    // on the main/admin device (see QRScanPair.jsx). That device never logs
    // in and has no admin token, so requiring auth here made every secondary
    // -device pairing attempt fail with "No token provided" / 401. The
    // sessionId + studentToken pair themselves act as the authorization for
    // this single, narrow action (handleUpdateQRSession validates both
    // server-side before doing anything).
    // 'heartbeatQRSession' and 'closeQRSession' join updateQRSession as public:
    // all three are called by the unauthenticated scanner device that keeps a
    // continuous QR-lookup session alive and ends it when the page closes.
    const publicEndpoints = [
      'login', 'health', 'validateQR',
      'updateQRSession', 'heartbeatQRSession', 'closeQRSession'
    ];


    if (action === 'health') {
      return createResponse({ success: true, message: 'SGMS API is running', version: '1.0.0' });
    }

    if (action === 'login') {
      return handleLogin(payload);
    }

    // FIX: validateQR is called by mobile QR scanners without an admin token.
    // It was listed in Router but NOT in publicEndpoints, so every scan returned
    // "No token provided". Added it to publicEndpoints here.
    if (action === 'validateQR') {
      return handleValidateQR(payload);
    }

    // Protected endpoints – require authentication
    if (!publicEndpoints.includes(action)) {
      const validation = validateSession(token);
      if (!validation.success) {
        return createResponse(validation, 401);
      }
    }

    // Route to appropriate service
    switch (action) {
      // Student Management
      case 'getStudents':        return handleGetStudents(payload);
      case 'getStudent':         return handleGetStudent(payload);
      case 'createStudent':      return handleCreateStudent(payload, token);
      case 'updateStudent':      return handleUpdateStudent(payload, token);
      case 'deleteStudent':      return handleDeleteStudent(payload, token);
      case 'searchStudents':     return handleSearchStudents(payload);

      // Activity Management
      case 'getActivities':      return handleGetActivities(payload);
      case 'getActivity':        return handleGetActivity(payload);
      case 'createActivity':     return handleCreateActivity(payload, token);
      case 'updateActivity':     return handleUpdateActivity(payload, token);
      case 'deleteActivity':     return handleDeleteActivity(payload, token);

      // Grading Terms
      case 'getGradingTerms':    return handleGetGradingTerms();
      case 'updateGradingTerm':  return handleUpdateGradingTerm(payload, token);

      // Score Management
      case 'saveScore':          return handleSaveScore(payload, token);
      case 'getScores':          return handleGetScores(payload);
      case 'getStudentScores':   return handleGetStudentScores(payload);
      case 'bulkSaveScores':     return handleBulkSaveScores(payload, token);

      // Grade Calculation
      case 'calculateGrades':    return handleCalculateGrades(payload);
      case 'getStudentGrade':    return handleGetStudentGrade(payload);

      // QR System
      case 'generateStudentQR':  return handleGenerateStudentQR(payload, token);
      // validateQR handled above (public)
      case 'createQRSession':    return handleCreateQRSession(payload, token);
      case 'getQRSession':       return handleGetQRSession(payload);
      case 'updateQRSession':    return handleUpdateQRSession(payload, token);
      case 'heartbeatQRSession': return handleHeartbeatQRSession(payload);
      case 'closeQRSession':     return handleCloseQRSession(payload);

      // Reports
      case 'getStudentReport':   return handleGetStudentReport(payload);
      case 'getClassReport':     return handleGetClassReport(payload);
      case 'getPassFailList':    return handleGetPassFailList(payload);

      // Print Reports
      case 'generatePrintReport': return handleGeneratePrintReport(payload, token);

      // Import/Export
      case 'importStudents':     return handleImportStudents(payload, token);
      case 'exportStudents':     return handleExportStudents(payload);
      case 'importScores':       return handleImportScores(payload, token);

      // Settings
      case 'getSettings':        return handleGetSettings();
      case 'updateSettings':     return handleUpdateSettings(payload, token);

      // Audit Log
      case 'getAuditLog':        return handleGetAuditLog(payload);

      // Dashboard
      case 'getDashboardStats':  return handleGetDashboardStats();

      // Admin Management
      case 'getAdmins':          return handleGetAdmins();
      case 'createAdmin':        return handleCreateAdmin(payload, token);
      case 'updateAdmin':        return handleUpdateAdmin(payload, token);
      case 'changePassword':     return handleChangePassword(payload, token);

      default:
        return createResponse({
          success: false,
          message: 'Unknown action: ' + action,
          data: null
        }, 404);
    }

  } catch (error) {
    Logger.log('Router error: ' + error.toString());
    return createResponse({
      success: false,
      message: 'Routing error: ' + error.toString(),
      data: null
    }, 500);
  }
}

/**
 * Create a JSON response.
 * Google Apps Script always returns HTTP 200; the client reads data.success.
 */
function createResponse(data, statusCode) {
  statusCode = statusCode || 200; // Apps Script ignores status codes anyway
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}