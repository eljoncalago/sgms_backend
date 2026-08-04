/**
 * SGMS Print Service
 * Generates official Excel/PDF report cards by filling a Google Drive template.
 *
 * Template requirements:
 *  - An Excel file (.xlsx) stored in Google Drive, converted to Google Sheets format
 *    (File → Save as Google Sheets) so Apps Script can write to it.
 *  - The first sheet must follow the EXACT cell layout specified in the
 *    SGMS Master Specification (Part 5 Appendix — Excel Grade Sheet Automation).
 *  - Two students fit per page, side by side (Student 1 left, Student 2 right).
 *  - Sheet name: "Grade Card" (or first sheet if not found by name)
 *
 * Cell mapping implemented here exactly matches the specification:
 *
 *   STUDENT INFORMATION
 *     Student 1 ID:     C4:E5        Student 2 ID:    W4:Y5
 *     Thai Name S1:     D9:G9        Thai Name S2:    X9:AA9
 *     English Name S1:  D12:G12      English Name S2: X12:AA12
 *     Year/Section S1:  I9:J9        Year/Section S2: AC9:AD9
 *     Class Number S1:  I12:J12      Class Number S2: AC12:AD12
 *     QR Code S1:       L4:O12       QR Code S2:      AF4:AI12
 *
 *   MIDTERM COLLECTIVE (rows 18-31)
 *     Activity names S1:   C18:D30      Activity names S2:  W18:X30
 *     Raw scores S1:       E18:E30      Raw scores S2:      Y18:Y30
 *     Perfect scores S1:   F18:F30      Perfect scores S2:  Z18:Z30
 *     Weight S1:           G16          Weight S2:          AA16
 *     Totals S1:           E31,F31      Totals S2:          Y31,Z31
 *     Equiv score S1:      G31          Equiv score S2:     AA31
 *     Pass/Fail S1:        H31          Pass/Fail S2:       AB31
 *
 *   FINAL COLLECTIVE INITIAL (rows 18-23 cols I-N / AC-AH)
 *     Activity names S1:   I18:J22      Activity names S2:  AC18:AD22
 *     Raw scores S1:       K18:K22      Raw scores S2:      AE18:AE22
 *     Perfect scores S1:   L18:L22      Perfect scores S2:  AF18:AF22
 *     Weight S1:           M16          Weight S2:          AG16
 *     Totals S1:           K23,L23      Totals S2:          AE23,AF23
 *     Equiv score S1:      M23          Equiv score S2:     AG23
 *     Pass/Fail S1:        N23          Pass/Fail S2:       AH23
 *
 *   FINAL COLLECTIVE FINAL (rows 27-31 cols I-N / AC-AH)
 *     Activity names S1:   I27:J30      Activity names S2:  AC27:AD30
 *     Raw scores S1:       K27:K30      Raw scores S2:      AE27:AE30
 *     Perfect scores S1:   L27:L30      Perfect scores S2:  AF27:AF30
 *     Weight S1:           M25          Weight S2:          AG25
 *     Totals S1:           K31,L31      Totals S2:          AE31,AF31
 *     Equiv score S1:      M31          Equiv score S2:     AG31
 *     Pass/Fail S1:        N31          Pass/Fail S2:       AH31
 *
 *   MIDTERM EXAMINATION (rows 36-42)
 *     Types S1:            C36:D41      Types S2:           W36:X41
 *     Raw scores S1:       E36:E41      Raw scores S2:      Y36:Y41
 *     Perfect scores S1:   F36:F41      Perfect scores S2:  Z36:Z41
 *     Weight S1:           G34          Weight S2:          AA34
 *     Totals S1:           E42,F42      Totals S2:          Y42,Z42
 *     Equiv score S1:      G42          Equiv score S2:     AA42
 *     Pass/Fail S1:        H42          Pass/Fail S2:       AB42
 *
 *   FINAL EXAMINATION (rows 36-42 cols I-N / AC-AH)
 *     Types S1:            I36:J41      Types S2:           AC36:AD41
 *     Raw scores S1:       K36:K41      Raw scores S2:      AE36:AE41
 *     Perfect scores S1:   L36:L41      Perfect scores S2:  AF36:AF41
 *     Weight S1:           M34          Weight S2:          AG34
 *     Totals S1:           K42,L42      Totals S2:          AE42,AF42
 *     Equiv score S1:      M42          Equiv score S2:     AG42
 *     Pass/Fail S1:        N42          Pass/Fail S2:       AH42
 *
 *   SUMMARY (rows 56-67)
 *     Midterm Coll S1:     G56:H56     Midterm Coll S2:    AA56:AB56
 *     Midterm Exam S1:     G59:H59     Midterm Exam S2:    AA58:AB58
 *     Total Final Coll S1: G62:H62     Total Final Coll S2:AA62:AB62
 *     Final Exam S1:       G65:H65     Final Exam S2:      AA65:AB65
 *     Overall S1:          G67:H67     Overall S2:         AA67:AB67
 *
 *   SEMESTER GRADE SUMMARY TABLE — STUDENT 1 (col S, rows 18-43)
 *     S18 = midterm coll equiv (G31)
 *     S20 = midterm exam equiv (G42)
 *     S21:S22 = Stage 1 total
 *     S23 = Stage 1 pass/fail
 *     S28 = final coll initial equiv (M23)
 *     S29:S30 = Stage 2 total
 *     S31 = Stage 2 pass/fail
 *     S33 = final coll final equiv (M31)
 *     S34:S35 = Stage 3 total
 *     S36 = Stage 3 pass/fail
 *     S38 = final exam equiv (M42)
 *     S39:S40 = Stage 4 total
 *     S41 = Stage 4 pass/fail
 *     S43 = final cumulative semester score
 *
 *   SEMESTER GRADE SUMMARY TABLE — STUDENT 2 (col AM, same structure)
 */

// ─── Column letter helpers ────────────────────────────────────────────────────

/**
 * Convert a column letter (e.g. "A", "Z", "AA") to a 1-based column index.
 */
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
 * Insert a QR code image for a student into the template sheet.
 * The QR encodes the student's QR token so it can be scanned later.
 * The image is resized to fit entirely inside the assigned merged cell range.
 */
function insertQRCode(sheet, student, isLeft) {
  try {
    // QR cell ranges from the spec:
    //   Student 1: L4:O12  (cols L-O, rows 4-12)
    //   Student 2: AF4:AI12 (cols AF-AI, rows 4-12)
    var startCol = isLeft ? COL.L : COL.AF;
    var startRow = 4;
    var endCol   = isLeft ? COL.O : COL.AI;
    var endRow   = 12;

    // Look up the student's active QR token
    var qrTokens = findRecords(CONFIG.SHEETS.QR_TOKENS, {
      STUDENT_ID: student.STUDENT_ID,
      IS_ACTIVE: true
    });
    if (qrTokens.length === 0) return; // no QR token — skip silently

    var token = qrTokens[0].TOKEN;
    var qrUrl = generateQRUrl(token);

    // Fetch the QR image from the public API
    var qrBlob = UrlFetchApp.fetch(qrUrl).getBlob();
    var qrImage = sheet.insertImage(
      qrBlob,
      startCol,
      startRow,
      endCol - startCol + 1,
      endRow - startRow + 1
    );

    // Offset to center within the merged range
    qrImage.setAnchorCell(sheet.getRange(startRow, startCol));
    // Set a small inset so the QR sits nicely inside the border
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
  // FIX: filter by grade level — include activity if it has no GRADE_LEVEL (shared)
  //      OR its GRADE_LEVEL matches the student's grade level.
  var gradeToMatch = String(studentGradeLevel || '').trim();
  var allActivities = getAllRecords(CONFIG.SHEETS.ACTIVITIES);
  var termActivities = allActivities.filter(function(a) {
    if (a.TERM_ID !== termId) return false;
    if (!(a.IS_ACTIVE === true || a.IS_ACTIVE === 'TRUE' || a.IS_ACTIVE === 'true')) return false;
    var actGrade = String(a.GRADE_LEVEL || '').trim();
    // Include if activity is shared (no grade) OR matches student's grade
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

// ─── Fill one student's column ────────────────────────────────────────────────

/**
 * Fill all cells for one student on the sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {object} student  — STUDENTS record
 * @param {object} terms    — keyed by term role: midColl, midExam, finCollInit, finCollFin, finExam
 * @param {boolean} isLeft  — true = Student 1 (left), false = Student 2 (right)
 * @param {object} settings — SETTINGS object
 * @param {number} stageNum — semester stage to generate (1-4); defaults to 4 (all)
 */
function fillStudentColumns(sheet, student, terms, isLeft, settings, stageNum) {
  stageNum = stageNum || 4;
  var s1 = settings || {};

  // ── Column offsets ──
  var cID     = isLeft ? COL.C  : COL.W;
  var cThai   = isLeft ? COL.D  : COL.X;
  var cEng    = isLeft ? COL.D  : COL.X;
  var cYrSec  = isLeft ? COL.I  : COL.AC;
  var cClsNum = isLeft ? COL.I  : COL.AC;

  // Midterm Collective columns
  var cMC_name  = isLeft ? COL.C  : COL.W;
  var cMC_raw   = isLeft ? COL.E  : COL.Y;
  var cMC_max   = isLeft ? COL.F  : COL.Z;
  var cMC_wt    = isLeft ? COL.G  : COL.AA;
  var cMC_eq    = isLeft ? COL.G  : COL.AA;
  var cMC_pf    = isLeft ? COL.H  : COL.AB;
  var rowMC_wt  = 16;
  var rowMC_st  = 18;
  var rowMC_end = 30;
  var rowMC_tot = 31;

  // Final Collective Initial columns
  var cFCI_name  = isLeft ? COL.I  : COL.AC;
  var cFCI_raw   = isLeft ? COL.K  : COL.AE;
  var cFCI_max   = isLeft ? COL.L  : COL.AF;
  var cFCI_wt    = isLeft ? COL.M  : COL.AG;
  var cFCI_eq    = isLeft ? COL.M  : COL.AG;
  var cFCI_pf    = isLeft ? COL.N  : COL.AH;
  var rowFCI_wt  = 16;
  var rowFCI_st  = 18;
  var rowFCI_tot = 23;

  // Final Collective Final columns (shares col layout with FCI)
  var cFCF_name  = isLeft ? COL.I  : COL.AC;
  var cFCF_raw   = isLeft ? COL.K  : COL.AE;
  var cFCF_max   = isLeft ? COL.L  : COL.AF;
  var cFCF_wt    = isLeft ? COL.M  : COL.AG;
  var cFCF_eq    = isLeft ? COL.M  : COL.AG;
  var cFCF_pf    = isLeft ? COL.N  : COL.AH;
  var rowFCF_wt  = 25;
  var rowFCF_st  = 27;
  var rowFCF_tot = 31;

  // Midterm Exam columns
  var cME_name  = isLeft ? COL.C  : COL.W;
  var cME_raw   = isLeft ? COL.E  : COL.Y;
  var cME_max   = isLeft ? COL.F  : COL.Z;
  var cME_wt    = isLeft ? COL.G  : COL.AA;
  var cME_eq    = isLeft ? COL.G  : COL.AA;
  var cME_pf    = isLeft ? COL.H  : COL.AB;
  var rowME_wt  = 34;
  var rowME_st  = 36;
  var rowME_end = 41;
  var rowME_tot = 42;

  // Final Exam columns
  var cFE_name  = isLeft ? COL.I  : COL.AC;
  var cFE_raw   = isLeft ? COL.K  : COL.AE;
  var cFE_max   = isLeft ? COL.L  : COL.AF;
  var cFE_wt    = isLeft ? COL.M  : COL.AG;
  var cFE_eq    = isLeft ? COL.M  : COL.AG;
  var cFE_pf    = isLeft ? COL.N  : COL.AH;
  var rowFE_wt  = 34;
  var rowFE_st  = 36;
  var rowFE_end = 41;
  var rowFE_tot = 42;

  // Summary columns
  var cSumMC  = isLeft ? COL.G  : COL.AA;
  var cSumME  = isLeft ? COL.G  : COL.AA;
  var cSumFC  = isLeft ? COL.G  : COL.AA;
  var cSumFE  = isLeft ? COL.G  : COL.AA;
  var cSumOv  = isLeft ? COL.G  : COL.AA;

  // Semester grade summary table column
  var cSGST   = isLeft ? COL.S  : COL.AM;

  // ── Student info ──────────────────────────────────────────────────────────
  writeCell(sheet, 4,  cID,     student.STUDENT_ID);
  writeCell(sheet, 9,  cThai,   student.THAI_NAME);
  writeCell(sheet, 12, cEng,    student.ENGLISH_NAME);
  writeCell(sheet, 9,  cYrSec,  'M' + student.GRADE_LEVEL + '/' + student.SECTION_NUMBER);
  writeCell(sheet, 12, cClsNum, student.CLASS_NUMBER);

  // ── QR Code ────────────────────────────────────────────────────────────────
  insertQRCode(sheet, student, isLeft);

  // ── Midterm Collective ────────────────────────────────────────────────────
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

  // ── Final Collective Initial ──────────────────────────────────────────────
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

  // ── Final Collective Final ────────────────────────────────────────────────
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

  // ── Midterm Examination ───────────────────────────────────────────────────
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

  // ── Final Examination ─────────────────────────────────────────────────────
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

  // ── Score Summary (rows 56-67) ────────────────────────────────────────────
  writeCell(sheet, 56, cSumMC, mcCalc.equiv);         // Midterm Collective
  // Student 1 Midterm Exam summary = row 59; Student 2 = row 58 (per spec)
  writeCell(sheet, isLeft ? 59 : 58, cSumME, meCalc.equiv);
  writeCell(sheet, 62, cSumFC, fciCalc.equiv + fcfCalc.equiv);  // Total Final Collective
  writeCell(sheet, 65, cSumFE, feCalc.equiv);         // Final Exam
  var overall = mcCalc.equiv + meCalc.equiv + fciCalc.equiv + fcfCalc.equiv + feCalc.equiv;
  writeCell(sheet, 67, cSumOv, Math.round(overall * 100) / 100);

  // ── Semester Grade Summary Table ──────────────────────────────────────────
  writeCell(sheet, 18, cSGST, mcCalc.equiv);           // S18/AM18 — Midterm Coll
  writeCell(sheet, 20, cSGST, meCalc.equiv);           // S20/AM20 — Midterm Exam

  // Stage 1: Midterm Collective + Midterm Examination
  var st1 = mcCalc.equiv + meCalc.equiv;
  writeCell(sheet, 21, cSGST, st1);
  writeCell(sheet, 22, cSGST, st1);
  writeCell(sheet, 23, cSGST, passFail(st1, s1.STAGE1_PASSING));

  if (stageNum >= 2) {
    writeCell(sheet, 28, cSGST, fciCalc.equiv);        // S28/AM28 — Final Coll Init
    var st2 = st1 + fciCalc.equiv;
    writeCell(sheet, 29, cSGST, st2);
    writeCell(sheet, 30, cSGST, st2);
    writeCell(sheet, 31, cSGST, passFail(st2, s1.STAGE2_PASSING));
  }

  if (stageNum >= 3) {
    writeCell(sheet, 33, cSGST, fcfCalc.equiv);        // S33/AM33 — Final Coll Fin
    var st3 = st1 + fciCalc.equiv + fcfCalc.equiv;
    writeCell(sheet, 34, cSGST, st3);
    writeCell(sheet, 35, cSGST, st3);
    writeCell(sheet, 36, cSGST, passFail(st3, s1.STAGE3_PASSING));
  }

  if (stageNum >= 4) {
    writeCell(sheet, 38, cSGST, feCalc.equiv);         // S38/AM38 — Final Exam
    var st4 = Math.round(overall * 100) / 100;
    writeCell(sheet, 39, cSGST, st4);
    writeCell(sheet, 40, cSGST, st4);
    writeCell(sheet, 41, cSGST, passFail(st4, s1.STAGE4_PASSING || s1.OVERALL_PASSING_PERCENT));
    writeCell(sheet, 43, cSGST, st4);                  // Final cumulative semester score
  }
}

// ─── Main handler ─────────────────────────────────────────────────────────────

function handleGeneratePrintReport(payload, token) {
  try {
    var adminId = getAdminIdFromToken(token);
    var studentIds = payload.studentIds;
    var templateId = payload.templateId;
    var stageNum   = parseInt(payload.stage || 4, 10);

    if (!studentIds || !Array.isArray(studentIds) || studentIds.length === 0) {
      return createResponse({ success: false, message: 'Student IDs array is required', data: null });
    }
    if (!templateId) {
      return createResponse({ success: false, message: 'Template file ID is required. Set PRINT_TEMPLATE_ID in Settings.', data: null });
    }

    // ── Open template ────────────────────────────────────────────────
    var templateFile;
    try {
      templateFile = DriveApp.getFileById(templateId);
    } catch (e) {
      return createResponse({ success: false, message: 'Cannot access template file: ' + e.toString(), data: null });
    }

    // Make a working copy
    var copyName  = 'SGMS_Report_' + new Date().toISOString().slice(0, 10) + '_Stage' + stageNum;
    var copyFile  = templateFile.makeCopy(copyName);
    var copyId    = copyFile.getId();
    var workbook  = SpreadsheetApp.openById(copyId);
    var sheet     = workbook.getSheetByName('Grade Card') || workbook.getSheets()[0];

    // ── Fetch settings ────────────────────────────────────────────────
    var settings = getSettingsObject();

    // ── Load grading terms (sorted by order) ──────────────────────────
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

    // ── Fill pages — 2 students per page ──────────────────────────────
    var errors = [];
    var processed = 0;

    for (var i = 0; i < studentIds.length; i += 2) {
      var id1 = studentIds[i];
      var id2 = studentIds[i + 1] || null;

      var student1 = id1 ? findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', id1) : null;
      var student2 = id2 ? findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', id2) : null;

      if (i > 0) {
        var pageSheet = workbook.duplicateActiveSheet();
        pageSheet.setName('Page ' + (Math.floor(i / 2) + 1));
        sheet = pageSheet;
      }

      // FIX: buildAll now accepts the student object and passes the student's
      //      GRADE_LEVEL to buildTermData so it filters activities correctly.
      function buildAll(student) {
        var studentId = student.STUDENT_ID;
        var studentGradeLevel = student.GRADE_LEVEL || '';
        var result = {};
        Object.keys(termRoles).forEach(function(role) {
          var term = termRoles[role];
          result[role] = term.TERM_ID
            ? buildTermData(term.TERM_ID, studentId, studentGradeLevel)
            : { activities: [], weight: 0, passingPercent: 50, termName: role };
        });
        return result;
      }

      if (student1) {
        try {
          fillStudentColumns(sheet, student1, buildAll(student1), true, settings, stageNum);
          processed++;
        } catch (e) {
          errors.push('Student 1 (' + id1 + '): ' + e.toString());
        }
      }

      if (student2) {
        try {
          fillStudentColumns(sheet, student2, buildAll(student2), false, settings, stageNum);
          processed++;
        } catch (e) {
          errors.push('Student 2 (' + id2 + '): ' + e.toString());
        }
      }
    }

    // ── Flush and export as PDF ───────────────────────────────────────
    SpreadsheetApp.flush();

    var pdfBlob = workbook.getAs('application/pdf');
    pdfBlob.setName(copyName + '.pdf');
    var pdfFolder = DriveApp.getRootFolder();
    var pdfFile = pdfFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var pdfUrl = 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view';

    var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + copyId + '/edit';

    createAuditLog(
      adminId,
      'GENERATE_PRINT_REPORT',
      null, null, null,
      JSON.stringify({ studentIds: studentIds, stage: stageNum, processed: processed }),
      'web',
      'Generated Stage ' + stageNum + ' report cards for ' + processed + ' students'
    );

    return createResponse({
      success: true,
      message: 'Report cards generated for ' + processed + ' student' + (processed !== 1 ? 's' : '') +
               (errors.length > 0 ? ' (' + errors.length + ' errors)' : ''),
      data: {
        processed: processed,
        errors: errors,
        pdfUrl: pdfUrl,
        spreadsheetUrl: sheetUrl,
        fileUrl: pdfUrl,
      }
    });

  } catch (error) {
    Logger.log('PrintService error: ' + error.toString());
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}
