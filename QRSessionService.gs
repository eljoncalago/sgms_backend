/**
 * SGMS QR Session Service
 * Handles QR device pairing sessions.
 *
 * CONTINUOUS-SCAN UPGRADE
 * -----------------------
 * Originally a session could pair exactly ONE student: once STATUS became
 * 'FOUND' the session refused any further update. The Grades module now needs
 * a *lookup stream*: the secondary (scanner) device keeps its camera running
 * and sends every student QR it reads; the main device follows along and loads
 * each scanned student in turn.
 *
 * What changed:
 *  - handleUpdateQRSession accepts repeated scans. Each accepted scan bumps
 *    SCAN_COUNT and stamps LAST_SCAN_AT; the main device polls and treats a
 *    higher SCAN_COUNT as "a new student was scanned" (this also handles the
 *    same student being scanned twice in a row).
 *  - Sessions live until they EXPIRE or the scanner device closes the page
 *    (handleCloseQRSession → STATUS 'CLOSED'). A CLOSED/EXPIRED session
 *    rejects further scans.
 *  - handleHeartbeatQRSession lets the scanner device keep the session alive
 *    and lets the main device see that the scanner is still connected.
 *  - New columns (SCAN_COUNT, LAST_SCAN_AT, SCANNER_LAST_SEEN) are added to an
 *    existing QR_SESSIONS sheet automatically via ensureSheetColumns().
 */

const QR_SESSION_COLUMNS = [
  'SESSION_ID', 'CREATED_BY', 'DEVICE_ID', 'STATUS', 'STUDENT_ID',
  'CREATED_AT', 'UPDATED_AT', 'EXPIRES_AT', 'LAST_ACTIVITY',
  'SCAN_COUNT', 'LAST_SCAN_AT', 'SCANNER_LAST_SEEN'
];

/** Session is usable for scanning? */
function isSessionOpen_(session) {
  if (!session) return false;
  if (session.STATUS === 'CLOSED' || session.STATUS === 'EXPIRED') return false;
  return new Date(session.EXPIRES_AT) >= new Date();
}

function findSession_(sessionId) {
  const sessions = findRecords(CONFIG.SHEETS.QR_SESSIONS, { SESSION_ID: sessionId });
  return sessions.length > 0 ? sessions[0] : null;
}

function handleCreateQRSession(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);
    const { deviceId } = payload || {};

    // Make sure the continuous-scan columns exist on older databases.
    ensureSheetColumns(CONFIG.SHEETS.QR_SESSIONS, QR_SESSION_COLUMNS);

    const sessionId = generateSecureToken();
    const expiryTime = new Date();
    expiryTime.setMinutes(expiryTime.getMinutes() + CONFIG.SESSION.QR_SESSION_EXPIRY_MINUTES);

    const now = new Date().toISOString();
    const session = {
      SESSION_ID: sessionId,
      CREATED_BY: adminId,
      DEVICE_ID: deviceId || 'web',
      STATUS: 'WAITING',
      STUDENT_ID: '',
      CREATED_AT: now,
      UPDATED_AT: now,
      EXPIRES_AT: expiryTime.toISOString(),
      LAST_ACTIVITY: now,
      SCAN_COUNT: 0,
      LAST_SCAN_AT: '',
      SCANNER_LAST_SEEN: ''
    };

    insertRecord(CONFIG.SHEETS.QR_SESSIONS, session);

    return createResponse({
      success: true,
      message: 'QR session created',
      data: {
        sessionId: sessionId,
        expiresAt: expiryTime.toISOString(),
        qrUrl: generateQRUrl(`SESSION:${sessionId}`)
      }
    });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

/**
 * Polled by the main device (every few seconds) while a scanning session runs.
 * Always returns the session record — including SCAN_COUNT so the caller can
 * tell a brand-new scan from the one it already loaded.
 */
function handleGetQRSession(payload) {
  try {
    const { sessionId } = payload || {};

    if (!sessionId) {
      return createResponse({ success: false, message: 'Session ID is required', data: null });
    }

    const session = findSession_(sessionId);

    if (!session) {
      return createResponse({ success: false, message: 'Session not found', data: null });
    }

    // Expired sessions are marked once, then reported as expired.
    if (session.STATUS !== 'CLOSED' && new Date(session.EXPIRES_AT) < new Date()) {
      if (session.STATUS !== 'EXPIRED') {
        updateRecord(CONFIG.SHEETS.QR_SESSIONS, 'SESSION_ID', sessionId, {
          STATUS: 'EXPIRED',
          UPDATED_AT: new Date().toISOString()
        });
      }
      session.STATUS = 'EXPIRED';
      return createResponse({
        success: false,
        message: 'Session expired',
        data: { session: session }
      });
    }

    const responseData = {
      session: session,
      scanCount: Number(session.SCAN_COUNT || 0),
      status: session.STATUS
    };

    // Include the most recently scanned student (if any) on every poll.
    if (session.STUDENT_ID) {
      const student = findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', session.STUDENT_ID);
      if (student) responseData.student = student;
    }

    return createResponse({ success: true, message: 'Session retrieved', data: responseData });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

/**
 * Called by the (unauthenticated) scanner device for EVERY student QR it reads.
 * Public endpoint — the sessionId + a valid student token are the authorization.
 */
function handleUpdateQRSession(payload, token) {
  try {
    const { sessionId, studentToken } = payload || {};

    if (!sessionId || !studentToken) {
      return createResponse({ success: false, message: 'Session ID and student token are required', data: null });
    }

    const session = findSession_(sessionId);
    if (!session) {
      return createResponse({ success: false, message: 'Session not found', data: null });
    }

    if (session.STATUS === 'CLOSED') {
      return createResponse({ success: false, message: 'Session closed — ask the main device to start a new one', data: null });
    }

    if (new Date(session.EXPIRES_AT) < new Date()) {
      updateRecord(CONFIG.SHEETS.QR_SESSIONS, 'SESSION_ID', sessionId, {
        STATUS: 'EXPIRED',
        UPDATED_AT: new Date().toISOString()
      });
      return createResponse({ success: false, message: 'Session expired — ask the main device to create a new one', data: null });
    }

    // Validate the student QR token.
    const tokenRecords = findRecords(CONFIG.SHEETS.QR_TOKENS, { TOKEN: studentToken, IS_ACTIVE: true });
    if (tokenRecords.length === 0) {
      return createResponse({ success: false, message: 'Invalid student token', data: null });
    }

    const studentId = tokenRecords[0].STUDENT_ID;
    const student = findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', studentId);
    if (!student) {
      return createResponse({ success: false, message: 'Student not found for this QR code', data: null });
    }

    const now = new Date().toISOString();
    const scanCount = Number(session.SCAN_COUNT || 0) + 1;

    // NOTE: repeated scans are allowed — the session simply switches to the
    // newly scanned student. That is the whole point of the lookup workflow.
    updateRecord(CONFIG.SHEETS.QR_SESSIONS, 'SESSION_ID', sessionId, {
      STATUS: 'FOUND',
      STUDENT_ID: studentId,
      UPDATED_AT: now,
      LAST_ACTIVITY: now,
      SCAN_COUNT: scanCount,
      LAST_SCAN_AT: now,
      SCANNER_LAST_SEEN: now
    });

    return createResponse({
      success: true,
      message: 'Student sent to main device',
      data: { studentId: studentId, student: student, scanCount: scanCount }
    });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

/**
 * Keeps the session alive while the scanner page is open (public).
 * Also used by the main device to show "scanner connected".
 */
function handleHeartbeatQRSession(payload) {
  try {
    const { sessionId } = payload || {};
    if (!sessionId) {
      return createResponse({ success: false, message: 'Session ID is required', data: null });
    }

    const session = findSession_(sessionId);
    if (!session) {
      return createResponse({ success: false, message: 'Session not found', data: null });
    }

    if (!isSessionOpen_(session)) {
      return createResponse({
        success: false,
        message: session.STATUS === 'CLOSED' ? 'Session closed' : 'Session expired',
        data: { status: session.STATUS }
      });
    }

    const now = new Date().toISOString();
    updateRecord(CONFIG.SHEETS.QR_SESSIONS, 'SESSION_ID', sessionId, {
      SCANNER_LAST_SEEN: now,
      LAST_ACTIVITY: now
    });

    return createResponse({
      success: true,
      message: 'Session alive',
      data: { status: session.STATUS, expiresAt: session.EXPIRES_AT }
    });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

/**
 * Ends the session (public) — called when the scanner device closes the page,
 * or by the main device when the teacher stops the scanning session.
 */
function handleCloseQRSession(payload) {
  try {
    const { sessionId } = payload || {};
    if (!sessionId) {
      return createResponse({ success: false, message: 'Session ID is required', data: null });
    }

    const session = findSession_(sessionId);
    if (!session) {
      return createResponse({ success: false, message: 'Session not found', data: null });
    }

    const now = new Date().toISOString();
    updateRecord(CONFIG.SHEETS.QR_SESSIONS, 'SESSION_ID', sessionId, {
      STATUS: 'CLOSED',
      UPDATED_AT: now,
      LAST_ACTIVITY: now
    });

    return createResponse({ success: true, message: 'Session closed', data: { status: 'CLOSED' } });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}
