/**
 * SGMS Print Service
 * Generates official Excel/PDF report cards by filling a Google Drive template.
 *
 * Refactor 1 — ONE STUDENT PER A4 PAGE (Student 2 removed):
 *   Each student is written to its own cloned sheet (one A4 page). Only the
 *   Student 1 cell mapping below is used. Cell layout matches the Excel Grade
 *   Sheet Automation spec exactly.
 *
 * Refactor 2 — BATCH PRINTING (no disconnect for ~400 students):
 *   handleGeneratePrintReport now accepts batchIndex / totalBatches / workbookId.
 *   The frontend sends students in small batches; each call clones pages into
 *   ONE accumulating workbook (kept in Drive between calls) and only exports
 *   the final PDF on the last batch. Every call stays well under the Apps
 *   Script ~6-min ceiling, so a 400-student run finishes instead of dropping.
 *
 * Template requirements:
 *  - An Excel file (.xlsx) in Google Drive, saved as Google Sheets.
 *  - First sheet "Grade Card" follows the single-student layout below.
 *
 *   STUDENT INFORMATION
 *     Student ID     C4:E5       Year/Section    I9:J9   (e.g. M3/3)
 *     Thai Name      D9:G9       Class Number    I12:J12
 *     English Name   D12:G12     QR Code         L4:O12
 *
 *   MIDTERM COLLECTIVE (rows 18-31)
 *     Activity names C18:D30  Raw E18:E30  Perfect F18:F30  Weight G16
 *     Totals E31,F31  Equiv G31  Pass/Fail H31
 *
 *   FINAL COLLECTIVE INITIAL (rows 18-23, cols I-N)
 *     Activity names I18:J22  Raw K18:K22  Perfect L18:L22  Weight M16
 *     Totals K23,L23  Equiv M23  Pass/Fail N23
 *
 *   FINAL COLLECTIVE FINAL (rows 27-31, cols I-N)
 *     Activity names I27:J30  Raw K27:K30  Perfect L27:L30  Weight M25
 *     Totals K31,L31  Equiv M31  Pass/Fail N31
 *
 *   MIDTERM EXAMINATION (rows 36-42)
 *     Types C36:D41  Raw E36:E41  Perfect F36:F41  Weight G34
 *     Totals E42,F42  Equiv G42  Pass/Fail H42
 *
 *   FINAL EXAMINATION (rows 36-42, cols I-N)
 *     Types I36:J41  Raw K36:K41  Perfect L36:L41  Weight M34
 *     Totals K42,L42  Equiv M42  Pass/Fail N42
 *
 *   SUMMARY (rows 56-67)
 *     Midterm Coll G56:H56  Midterm Exam G59:H59  Total Final Coll G62:H62
 *     Final Exam G65:H65    Overall G67:H67
 *
 *   SEMESTER GRADE SUMMARY TABLE (col S, rows 18-43)
 *     S18=Midterm Coll  S20=Midterm Exam  S21:S22=Stage1 total  S23=Stage1 PF
 *     S28=Final Coll Init  S29:S30=Stage2 total  S31=Stage2 PF
 *     S33=Final Coll Final  S34:S35=Stage3 total  S36=Stage3 PF
 *     S38=Final Exam  S39:S40=Stage4 total  S41=Stage4 PF  S43=Final cumulative
 */

// ─── Column letter helpers ────────────────────────────────────────────────────

function colLetterToIndex(letter) {
  letter = letter.toUpperCase();
  var idx = 0;
  for (var i = 0; i < letter.length; i++) {
    idx = idx * 26 + (letter.charCodeAt(i) - 64);
  }
  return idx;
}

var COL = {
  C:  colLetterToIndex('C'),
  D:  colLetterToIndex('D'),
  E:  colLetterToIndex('E'),
  F:  colLetterToIndex('F'),
  G:  colLetterToIndex('G'),
  H:  colLetterToIndex('H'),
  I:  colLetterToIndex('I'),
  J:  colLetterToIndex('J'),
  K:  colLetterToIndex('K'),
  L:  colLetterToIndex('L'),
  M:  colLetterToIndex('M'),
  N:  colLetterToIndex('N'),
  O:  colLetterToIndex('O'),
  S:  colLetterToIndex('S'),
  W:  colLetterToIndex('W'),
  X:  colLetterToIndex('X'),
  Y:  colLetterToIndex('Y'),
  Z:  colLetterToIndex('Z'),
  AA: colLetterToIndex('AA'),
  AB: colLetterToIndex('AB'),
  AC: colLetterToIndex('AC'),
  AD: colLetterToIndex('AD'),
  AE: colLetterToIndex('AE'),
  AF: colLetterToIndex('AF'),
  AG: colLetterToIndex('AG'),
  AH: colLetterToIndex('AH'),
  AI: colLetterToIndex('AI'),
  AM: colLetterToIndex('AM'),
};

// ─── Safe cell write ──────────────────────────────────────────────────────────

function writeCell(sheet, row, col, value) {
  if (value === null || value === undefined || value === '') return;
  try {
    sheet.getRange(row, col).setValue(value);
  } catch (e) {
    Logger.log('writeCell error [' + row + ',' + col + ']: ' + e);
  }
}

// ─── Pass/Fail helper ─────────────────────────────────────────────────────────

function passFail(score, threshold) {
  if (score === null || score === undefined || score === '') return '';
  return parseFloat(score) >= parseFloat(threshold || 0) ? 'PASS' : 'FAIL';
}

// ─── QR code insertion ───────────────────────────────────────────────────────

/**
 * Insert a QR code image for a student into L4:O12 (Student 1 only).
 */
function insertQRCode(sheet, student) {
  try {
    var startCol = COL.L;
    var startRow = 4;
    var endCol   = COL.O;
    var endRow   = 12;

    var qrTokens = findRecords(CONFIG.SHEETS.QR_TOKENS, {
      STUDENT_ID: student.STUDENT_ID,
      IS_ACTIVE: true
    });
    if (qrTokens.length === 0) return; // no QR token — skip silently

    var token = qrTokens[0].TOKEN;
    var qrUrl = generateQRUrl(token);

    var qrBlob = UrlFetchApp.fetch(qrUrl).getBlob();
    var qrImage = sheet.insertImage(
      qrBlob,
      startCol,
      startRow,
      endCol - startCol + 1,
      endRow - startRow + 1
    );

    qrImage.setAnchorCell(sheet.getRange(startRow, startCol));
    qrImage.setAnchorCellXOffset(2);
    qrImage.setAnchorCellYOffset(2);
  } catch (e) {
    Logger.log('insertQRCode error for student ' + student.STUDENT_ID + ': ' + e);
  }
}

// ─── Term data builder ────────────────────────────────────────────────────────

/**
 * For a given termId, collect activities (sorted by order) and student scores.
 * Returns { activities: [{name, type, maxScore, rawScore}], weight, passingPercent }
 */
function buildTermData(termId, studentId, studentGradeLevel) {
  var gradeToMatch = String(studentGradeLevel || '').trim();
  var allActivities = getAllRecords(CONFIG.SHEETS.ACTIVITIES);
  var termActivities = allActivities.filter(function(a) {
    if (a.TERM_ID !== termId) return false;
    if (!(a.IS_ACTIVE === true || a.IS_ACTIVE === 'TRUE' || a.IS_ACTIVE === 'true')) return false;
    var actGrade = String(a.GRADE_LEVEL || '').trim();
    if (actGrade !== '' && gradeToMatch !== '' && actGrade !== gradeToMatch) return false;
    return true;
  });
  termActivities.sort(function(a, b) { return (a.ACTIVITY_ORDER || 0) - (b.ACTIVITY_ORDER || 0); });

  var studentScores = findRecords(CONFIG.SHEETS.SCORES, { STUDENT_ID: studentId });

  var activities = termActivities.map(function(act) {
    var score = studentScores.find(function(s) { return s.ACTIVITY_ID === act.ACTIVITY_ID; });
    return {
      name: act.ACTIVITY_NAME || '',
      type: act.ACTIVITY_TYPE || '',
      maxScore: parseFloat(act.MAX_SCORE) || 0,
      rawScore: score ? parseFloat(score.RAW_SCORE) : 0,
    };
  });

  var term = findRecordById(CONFIG.SHEETS.GRADING_TERMS, 'TERM_ID', termId) || {};
  return {
    activities: activities,
    weight: parseFloat(term.WEIGHT_PERCENT) || 0,
    passingPercent: parseFloat(term.PASSING_PERCENT) || 50,
    termName: term.TERM_NAME || termId,
  };
}

/**
 * Calculate equivalent score: (rawTotal / maxTotal) * 100 * (weight/100)
 * Returns blank equiv when maxTotal is 0 to prevent division-by-zero errors.
 */
function calcEquiv(activities, weight) {
  var maxTotal = activities.reduce(function(s, a) { return s + a.maxScore; }, 0);
  var rawTotal = activities.reduce(function(s, a) { return s + a.rawScore; }, 0);
  if (maxTotal === 0) return { rawTotal: 0, maxTotal: 0, equiv: '' };
  var rawPct = (rawTotal / maxTotal) * 100;
  return {
    rawTotal: rawTotal,
    maxTotal: maxTotal,
    equiv: Math.round(rawPct * (weight / 100) * 100) / 100,
  };
}

// ─── Fill one student's page (Student 1 only) ─────────────────────────────────

/**
 * Fill all cells for one student on a sheet (one student per A4 page).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {object} student  — STUDENTS record
 * @param {object} terms    — keyed by: midColl, midExam, finCollInit, finCollFin, finExam
 * @param {object} settings — SETTINGS object
 * @param {number} stageNum — semester stage to generate (1-4); defaults to 4 (all)
 */
function fillStudentColumns(sheet, student, terms, settings, stageNum) {
  stageNum = stageNum || 4;
  var s1 = settings || {};

  // Student 1 column mapping (Student 2 removed — one student per page)
  var cID = COL.C, cThai = COL.D, cEng = COL.D, cYrSec = COL.I, cClsNum = COL.I;

  // Midterm Collective
  var cMC_name = COL.C, cMC_raw = COL.E, cMC_max = COL.F, cMC_wt = COL.G, cMC_eq = COL.G, cMC_pf = COL.H;
  var rowMC_wt = 16, rowMC_st = 18, rowMC_end = 30, rowMC_tot = 31;

  // Final Collective Initial
  var cFCI_name = COL.I, cFCI_raw = COL.K, cFCI_max = COL.L, cFCI_wt = COL.M, cFCI_eq = COL.M, cFCI_pf = COL.N;
  var rowFCI_wt = 16, rowFCI_st = 18, rowFCI_tot = 23;

  // Final Collective Final
  var cFCF_name = COL.I, cFCF_raw = COL.K, cFCF_max = COL.L, cFCF_wt = COL.M, cFCF_eq = COL.M, cFCF_pf = COL.N;
  var rowFCF_wt = 25, rowFCF_st = 27, rowFCF_tot = 31;

  // Midterm Examination
  var cME_name = COL.C, cME_raw = COL.E, cME_max = COL.F, cME_wt = COL.G, cME_eq = COL.G, cME_pf = COL.H;
  var rowME_wt = 34, rowME_st = 36, rowME_end = 41, rowME_tot = 42;

  // Final Examination
  var cFE_name = COL.I, cFE_raw = COL.K, cFE_max = COL.L, cFE_wt = COL.M, cFE_eq = COL.M, cFE_pf = COL.N;
  var rowFE_wt = 34, rowFE_st = 36, rowFE_end = 41, rowFE_tot = 42;

  // Summary + semester table (Student 1)
  var cSum = COL.G;
  var cSGST = COL.S;

  // ── Student info ──
  writeCell(sheet, 4,  cID,     student.STUDENT_ID);
  writeCell(sheet, 9,  cThai,   student.THAI_NAME);
  writeCell(sheet, 12, cEng,    student.ENGLISH_NAME);
  writeCell(sheet, 9,  cYrSec,  'M' + student.GRADE_LEVEL + '/' + student.SECTION_NUMBER);
  writeCell(sheet, 12, cClsNum, student.CLASS_NUMBER);

  // ── QR Code (L4:O12) ──
  insertQRCode(sheet, student);

  // ── Midterm Collective ──
  var mc = terms.midColl;
  writeCell(sheet, rowMC_wt, cMC_wt, mc.weight);
  mc.activities.forEach(function(act, i) {
    var row = rowMC_st + i;
    if (row > rowMC_end) return;
    writeCell(sheet, row, cMC_name, act.name);
    writeCell(sheet, row, cMC_raw,  act.rawScore);
    writeCell(sheet, row, cMC_max,  act.maxScore);
  });
  var mcCalc = calcEquiv(mc.activities, mc.weight);
  writeCell(sheet, rowMC_tot, cMC_raw, mcCalc.rawTotal);
  writeCell(sheet, rowMC_tot, cMC_max, mcCalc.maxTotal);
  writeCell(sheet, rowMC_tot, cMC_eq,  mcCalc.equiv);
  writeCell(sheet, rowMC_tot, cMC_pf,  passFail(mcCalc.equiv, s1.MIDTERM_COLLECTIVE_PASSING || 50));

  // ── Final Collective Initial ──
  var fci = terms.finCollInit;
  writeCell(sheet, rowFCI_wt, cFCI_wt, fci.weight);
  fci.activities.forEach(function(act, i) {
    var row = rowFCI_st + i;
    if (row >= rowFCI_tot) return;
    writeCell(sheet, row, cFCI_name, act.name);
    writeCell(sheet, row, cFCI_raw,  act.rawScore);
    writeCell(sheet, row, cFCI_max,  act.maxScore);
  });
  var fciCalc = calcEquiv(fci.activities, fci.weight);
  writeCell(sheet, rowFCI_tot, cFCI_raw, fciCalc.rawTotal);
  writeCell(sheet, rowFCI_tot, cFCI_max, fciCalc.maxTotal);
  writeCell(sheet, rowFCI_tot, cFCI_eq,  fciCalc.equiv);
  writeCell(sheet, rowFCI_tot, cFCI_pf,  passFail(fciCalc.equiv, s1.FINAL_COLLECTIVE_INITIAL_PASSING || 50));

  // ── Final Collective Final ──
  var fcf = terms.finCollFin;
  writeCell(sheet, rowFCF_wt, cFCF_wt, fcf.weight);
  fcf.activities.forEach(function(act, i) {
    var row = rowFCF_st + i;
    if (row > rowFCF_tot) return;
    writeCell(sheet, row, cFCF_name, act.name);
    writeCell(sheet, row, cFCF_raw,  act.rawScore);
    writeCell(sheet, row, cFCF_max,  act.maxScore);
  });
  var fcfCalc = calcEquiv(fcf.activities, fcf.weight);
  writeCell(sheet, rowFCF_tot, cFCF_raw, fcfCalc.rawTotal);
  writeCell(sheet, rowFCF_tot, cFCF_max, fcfCalc.maxTotal);
  writeCell(sheet, rowFCF_tot, cFCF_eq,  fcfCalc.equiv);
  writeCell(sheet, rowFCF_tot, cFCF_pf,  passFail(fcfCalc.equiv, s1.FINAL_COLLECTIVE_FINAL_PASSING || 50));

  // ── Midterm Examination ──
  var me = terms.midExam;
  writeCell(sheet, rowME_wt, cME_wt, me.weight);
  me.activities.forEach(function(act, i) {
    var row = rowME_st + i;
    if (row > rowME_end) return;
    writeCell(sheet, row, cME_name, act.name || act.type || ('Exam ' + (i + 1)));
    writeCell(sheet, row, cME_raw,  act.rawScore);
    writeCell(sheet, row, cME_max,  act.maxScore);
  });
  var meCalc = calcEquiv(me.activities, me.weight);
  writeCell(sheet, rowME_tot, cME_raw, meCalc.rawTotal);
  writeCell(sheet, rowME_tot, cME_max, meCalc.maxTotal);
  writeCell(sheet, rowME_tot, cME_eq,  meCalc.equiv);
  writeCell(sheet, rowME_tot, cME_pf,  passFail(meCalc.equiv, s1.MIDTERM_EXAM_PASSING || 50));

  // ── Final Examination ──
  var fe = terms.finExam;
  writeCell(sheet, rowFE_wt, cFE_wt, fe.weight);
  fe.activities.forEach(function(act, i) {
    var row = rowFE_st + i;
    if (row > rowFE_end) return;
    writeCell(sheet, row, cFE_name, act.name || act.type || ('Exam ' + (i + 1)));
    writeCell(sheet, row, cFE_raw,  act.rawScore);
    writeCell(sheet, row, cFE_max,  act.maxScore);
  });
  var feCalc = calcEquiv(fe.activities, fe.weight);
  writeCell(sheet, rowFE_tot, cFE_raw, feCalc.rawTotal);
  writeCell(sheet, rowFE_tot, cFE_max, feCalc.maxTotal);
  writeCell(sheet, rowFE_tot, cFE_eq,  feCalc.equiv);
  writeCell(sheet, rowFE_tot, cFE_pf,  passFail(feCalc.equiv, s1.FINAL_EXAM_PASSING || 50));

  // ── Score Summary (rows 56-67) ──
  writeCell(sheet, 56, cSum, mcCalc.equiv);                       // Midterm Collective
  writeCell(sheet, 59, cSum, meCalc.equiv);                       // Midterm Examination
  writeCell(sheet, 62, cSum, fciCalc.equiv + fcfCalc.equiv);      // Total Final Collective
  writeCell(sheet, 65, cSum, feCalc.equiv);                       // Final Examination
  var overall = mcCalc.equiv + meCalc.equiv + fciCalc.equiv + fcfCalc.equiv + feCalc.equiv;
  writeCell(sheet, 67, cSum, Math.round(overall * 100) / 100);   // Overall

  // ── Semester Grade Summary Table (col S, rows 18-43) ──
  writeCell(sheet, 18, cSGST, mcCalc.equiv);   // S18 — Midterm Collective
  writeCell(sheet, 20, cSGST, meCalc.equiv);   // S20 — Midterm Examination

  // Stage 1: Midterm Collective + Midterm Examination
  var st1 = mcCalc.equiv + meCalc.equiv;
  writeCell(sheet, 21, cSGST, st1);
  writeCell(sheet, 22, cSGST, st1);
  writeCell(sheet, 23, cSGST, passFail(st1, s1.STAGE1_PASSING));

  if (stageNum >= 2) {
    writeCell(sheet, 28, cSGST, fciCalc.equiv);   // S28 — Final Collective Initial
    var st2 = st1 + fciCalc.equiv;
    writeCell(sheet, 29, cSGST, st2);
    writeCell(sheet, 30, cSGST, st2);
    writeCell(sheet, 31, cSGST, passFail(st2, s1.STAGE2_PASSING));
  }

  if (stageNum >= 3) {
    writeCell(sheet, 33, cSGST, fcfCalc.equiv);   // S33 — Final Collective Final
    var st3 = st1 + fciCalc.equiv + fcfCalc.equiv;
    writeCell(sheet, 34, cSGST, st3);
    writeCell(sheet, 35, cSGST, st3);
    writeCell(sheet, 36, cSGST, passFail(st3, s1.STAGE3_PASSING));
  }

  if (stageNum >= 4) {
    writeCell(sheet, 38, cSGST, feCalc.equiv);    // S38 — Final Examination
    var st4 = Math.round(overall * 100) / 100;
    writeCell(sheet, 39, cSGST, st4);
    writeCell(sheet, 40, cSGST, st4);
    writeCell(sheet, 41, cSGST, passFail(st4, s1.STAGE4_PASSING || s1.OVERALL_PASSING_PERCENT));
    writeCell(sheet, 43, cSGST, st4);            // S43 — Final cumulative semester score
  }
}

// ─── Main handler ─────────────────────────────────────────────────────────────

function handleGeneratePrintReport(payload, token) {
  try {
    var adminId = getAdminIdFromToken(token);
    var studentIds = payload.studentIds;
    var templateId = payload.templateId;
    var stageNum   = parseInt(payload.stage || 4, 10);

    // Refactor 2 — batch support. Frontend sends students in small batches so
    // each call stays well under the Apps Script ~6-min limit and a 400-student
    // run finishes instead of disconnecting. One accumulating workbook is kept
    // in Drive (workbookId) across batches; PDF is exported only on the last.
    var batchIndex   = payload.batchIndex !== undefined ? parseInt(payload.batchIndex, 10) : 0;
    var totalBatches = payload.totalBatches !== undefined ? parseInt(payload.totalBatches, 10) : 1;
    var workbookId   = payload.workbookId || null;
    var isLastBatch  = (totalBatches <= 1) || (batchIndex >= totalBatches - 1);

    if (!studentIds || !Array.isArray(studentIds) || studentIds.length === 0) {
      return createResponse({ success: false, message: 'Student IDs array is required', data: null });
    }
    if (!templateId) {
      return createResponse({ success: false, message: 'Template file ID is required. Set PRINT_TEMPLATE_ID in Settings.', data: null });
    }

    // ── Settings + grading terms ──
    var settings = getSettingsObject();

    var allTerms = getAllRecords(CONFIG.SHEETS.GRADING_TERMS);
    allTerms.sort(function(a, b) { return a.TERM_ORDER - b.TERM_ORDER; });

    function findTerm(keywords) {
      return allTerms.find(function(t) {
        var n = (t.TERM_NAME || '').toLowerCase();
        return keywords.some(function(k) { return n.indexOf(k.toLowerCase()) !== -1; });
      }) || allTerms[0] || {};
    }

    var termRoles = {
      midColl:    findTerm(['midterm collective', 'mid collective', 'midcoll']),
      finCollInit:findTerm(['final collective initial', 'fin coll init', 'final collective i']),
      finCollFin: findTerm(['final collective final', 'fin coll fin', 'final collective f']),
      midExam:    findTerm(['midterm exam', 'mid exam', 'midterm examination']),
      finExam:    findTerm(['final exam', 'fin exam', 'final examination']),
    };

    function buildAll(student) {
      var sid = student.STUDENT_ID;
      var studentGradeLevel = student.GRADE_LEVEL || '';
      var result = {};
      Object.keys(termRoles).forEach(function(role) {
        var term = termRoles[role];
        result[role] = term.TERM_ID
          ? buildTermData(term.TERM_ID, sid, studentGradeLevel)
          : { activities: [], weight: 0, passingPercent: 50, termName: role };
      });
      return result;
    }

    var copyName = 'SGMS_Report_' + new Date().toISOString().slice(0, 10) + '_Stage' + stageNum;
    var copyId;
    var workbook;
    var templateSheet; // clean template sheet kept for cloning (one student per page)

    if (workbookId) {
      // Continue an existing accumulation workbook (subsequent batches)
      try {
        workbook = SpreadsheetApp.openById(workbookId);
      } catch (e) {
        return createResponse({ success: false, message: 'Cannot reopen workbook: ' + e.toString(), data: null });
      }
      copyId = workbookId;
      templateSheet = workbook.getSheets()[0];
    } else {
      // First batch: make a fresh copy of the template
      var templateFile;
      try {
        templateFile = DriveApp.getFileById(templateId);
      } catch (e) {
        return createResponse({ success: false, message: 'Cannot access template file: ' + e.toString(), data: null });
      }
      var copyFile = templateFile.makeCopy(copyName);
      copyId = copyFile.getId();
      workbook = SpreadsheetApp.openById(copyId);
      templateSheet = workbook.getSheetByName('Grade Card') || workbook.getSheets()[0];
    }

    // ── Fill one student per cloned page (Refactor 1: Student 2 removed) ──
    var errors = [];
    var processed = 0;

    for (var i = 0; i < studentIds.length; i++) {
      var studentId = studentIds[i];
      var student = findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', studentId);
      if (!student) {
        errors.push('Student (' + studentId + '): not found');
        continue;
      }

      // Clone the clean template sheet so every page keeps the merged-cell layout
      var pageSheet = templateSheet.copyTo(workbook);
      pageSheet.setName('Page ' + (workbook.getNumSheets() - 1));

      try {
        fillStudentColumns(pageSheet, student, buildAll(student), settings, stageNum);
        processed++;
      } catch (e) {
        errors.push('Student (' + studentId + '): ' + e.toString());
      }
    }

    SpreadsheetApp.flush();

    var pdfUrl = null;
    var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + copyId + '/edit';

    if (isLastBatch) {
      // Remove the clean template sheet so it is not an empty page in the PDF
      try { workbook.deleteSheet(templateSheet); } catch (e) { /* ignore */ }
      SpreadsheetApp.flush();

      var pdfBlob = workbook.getAs('application/pdf');
      pdfBlob.setName(copyName + '.pdf');
      var pdfFile = DriveApp.getRootFolder().createFile(pdfBlob);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfUrl = 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view';

      createAuditLog(
        adminId,
        'GENERATE_PRINT_REPORT',
        null, null, null,
        JSON.stringify({ stage: stageNum, processed: processed, batches: totalBatches }),
        'web',
        'Generated Stage ' + stageNum + ' report cards (' + totalBatches + ' batch' + (totalBatches !== 1 ? 'es' : '') + ')'
      );
    }

    return createResponse({
      success: true,
      message: isLastBatch
        ? 'Report cards generated for ' + processed + ' student' + (processed !== 1 ? 's' : '') + (errors.length > 0 ? ' (' + errors.length + ' errors)' : '')
        : 'Batch ' + (batchIndex + 1) + ' of ' + totalBatches + ' done (' + processed + ' students)',
      data: {
        processed: processed,
        errors: errors,
        batchIndex: batchIndex,
        totalBatches: totalBatches,
        isLastBatch: isLastBatch,
        workbookId: copyId,
        pdfUrl: pdfUrl,
        spreadsheetUrl: sheetUrl,
        fileUrl: pdfUrl
      }
    });

  } catch (error) {
    Logger.log('PrintService error: ' + error.toString());
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}