/**
 * SGMS QR Service
 * Handles QR token generation and validation
 */

function handleGenerateStudentQR(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);
    const { studentId } = payload;
    
    if (!studentId) {
      return createResponse({ success: false, message: 'Student ID is required', data: null });
    }
    
    const student = findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', studentId);
    if (!student) {
      return createResponse({ success: false, message: 'Student not found', data: null });
    }
    
    // Check for existing active token
    const existingTokens = findRecords(CONFIG.SHEETS.QR_TOKENS, { STUDENT_ID: studentId, IS_ACTIVE: true });
    
    if (existingTokens.length > 0) {
      return createResponse({
        success: true,
        message: 'QR token already exists',
        data: { token: existingTokens[0].TOKEN, qrUrl: generateQRUrl(existingTokens[0].TOKEN) }
      });
    }
    
    // Generate new secure token
    const qrToken = generateSecureToken();
    
    const tokenRecord = {
      TOKEN_ID: generateId('QRT'),
      TOKEN: qrToken,
      STUDENT_ID: studentId,
      IS_ACTIVE: true,
      CREATED_AT: new Date().toISOString(),
      EXPIRES_AT: '',
      LAST_ACCESSED_AT: new Date().toISOString()
    };
    
    insertRecord(CONFIG.SHEETS.QR_TOKENS, tokenRecord);
    
    createAuditLog(adminId, 'GENERATE_QR', studentId, null, null, null, 'web', 'Generated QR for: ' + student.ENGLISH_NAME);
    
    return createResponse({
      success: true,
      message: 'QR token generated',
      data: { token: qrToken, qrUrl: generateQRUrl(qrToken) }
    });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

function handleValidateQR(payload) {
  try {
    const { token } = payload;
    
    if (!token) {
      return createResponse({ success: false, message: 'Token is required', data: null });
    }
    
    const tokenRecords = findRecords(CONFIG.SHEETS.QR_TOKENS, { TOKEN: token, IS_ACTIVE: true });
    
    if (tokenRecords.length === 0) {
      return createResponse({ success: false, message: 'Invalid QR token', data: null });
    }
    
    const tokenRecord = tokenRecords[0];
    const student = findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', tokenRecord.STUDENT_ID);
    
    if (!student) {
      return createResponse({ success: false, message: 'Student not found', data: null });
    }
    
    // Update last accessed
    updateRecord(CONFIG.SHEETS.QR_TOKENS, 'TOKEN_ID', tokenRecord.TOKEN_ID, {
      LAST_ACCESSED_AT: new Date().toISOString()
    });
    
    return createResponse({
      success: true,
      message: 'QR token validated',
      data: { studentId: student.STUDENT_ID, studentName: student.ENGLISH_NAME, student: student }
    });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

function generateSecureToken() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let token = '';
  for (let i = 0; i < CONFIG.VALIDATION.QR_TOKEN_LENGTH; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

function generateQRUrl(token) {
  return `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(token)}`;
}