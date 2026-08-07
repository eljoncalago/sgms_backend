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
  try {
    sheet.getRange(row, col).setValue(value === null || value === undefined ? '' : value);
  } catch (e) {
    Logger.log('writeCell error [' + row + ',' + col + ']: ' + e);
  }
}

function writeFormula(sheet, row, col, formula) {
  try {
    sheet.getRange(row, col).setFormula(formula);
  } catch (e) {
    Logger.log('writeFormula error [' + row + ',' + col + ']: ' + e);
  }
}

function setReportSettingsSheet_(workbook, settings) {
  var sheet = workbook.getSheetByName('Report Settings');
  if (!sheet) {
    sheet = workbook.insertSheet('Report Settings');
    sheet.getRange('A1:C1').setValues([['SETTING_KEY', 'SETTING_VALUE', 'DESCRIPTION']]);
  }

  var values = [
    ['MIDTERM_COLLECTIVE_PASSING', parseFloat(settings.MIDTERM_COLLECTIVE_PASSING) || 50, 'Midterm Collective'],
    ['FINAL_COLLECTIVE_INITIAL_PASSING', parseFloat(settings.FINAL_COLLECTIVE_INITIAL_PASSING) || 50, 'Final Collective Initial'],
    ['FINAL_COLLECTIVE_FINAL_PASSING', parseFloat(settings.FINAL_COLLECTIVE_FINAL_PASSING) || 50, 'Final Collective Final'],
    ['MIDTERM_EXAM_PASSING', parseFloat(settings.MIDTERM_EXAM_PASSING) || 50, 'Midterm Examination'],
    ['FINAL_EXAM_PASSING', parseFloat(settings.FINAL_EXAM_PASSING) || 50, 'Final Examination'],
    ['STAGE1_PASSING', parseFloat(settings.STAGE1_PASSING) || 40, 'Stage 1'],
    ['STAGE2_PASSING', parseFloat(settings.STAGE2_PASSING) || 55, 'Stage 2'],
    ['STAGE3_PASSING', parseFloat(settings.STAGE3_PASSING) || 65, 'Stage 3'],
    ['STAGE4_PASSING', parseFloat(settings.STAGE4_PASSING || settings.OVERALL_PASSING_PERCENT) || 50, 'Stage 4']
  ];
  sheet.getRange(2, 1, values.length, 3).setValues(values);
  sheet.setFrozenRows(1);
  sheet.hideSheet();
  return sheet;
}

function reportSettingFormula_(key) {
  var keys = {
    MIDTERM_COLLECTIVE_PASSING: 2,
    FINAL_COLLECTIVE_INITIAL_PASSING: 3,
    FINAL_COLLECTIVE_FINAL_PASSING: 4,
    MIDTERM_EXAM_PASSING: 5,
    FINAL_EXAM_PASSING: 6,
    STAGE1_PASSING: 7,
    STAGE2_PASSING: 8,
    STAGE3_PASSING: 9,
    STAGE4_PASSING: 10
  };
  return "'Report Settings'!$B$" + (keys[key] || 2);
}

function reportJobKey_(jobId, batchIndex) {
  return 'SGMS_PRINT_JOB_' + String(jobId || 'legacy').replace(/[^A-Za-z0-9_-]/g, '_') + '_' + batchIndex;
}

function getCompletedPrintBatch_(jobId, batchIndex) {
  if (!jobId) return null;
  try {
    var value = PropertiesService.getScriptProperties().getProperty(reportJobKey_(jobId, batchIndex));
    return value ? JSON.parse(value) : null;
  } catch (e) {
    Logger.log('getCompletedPrintBatch error: ' + e);
    return null;
  }
}

function saveCompletedPrintBatch_(jobId, batchIndex, result) {
  if (!jobId) return;
  try {
    PropertiesService.getScriptProperties().setProperty(
      reportJobKey_(jobId, batchIndex),
      JSON.stringify(result)
    );
  } catch (e) {
    Logger.log('saveCompletedPrintBatch error: ' + e);
  }
}

function buildPrintContext_() {
  var students = getAllRecords(CONFIG.SHEETS.STUDENTS);
  var activities = getAllRecords(CONFIG.SHEETS.ACTIVITIES);
  var scores = getAllRecords(CONFIG.SHEETS.SCORES);
  var qrTokens = getAllRecords(CONFIG.SHEETS.QR_TOKENS);
  var terms = getAllRecords(CONFIG.SHEETS.GRADING_TERMS);
  var studentById = {};
  var scoresByStudent = {};
  var qrTokensByStudent = {};

  students.forEach(function(student) {
    studentById[String(student.STUDENT_ID).trim()] = student;
  });
  scores.forEach(function(score) {
    var key = String(score.STUDENT_ID).trim();
    if (!scoresByStudent[key]) scoresByStudent[key] = [];
    scoresByStudent[key].push(score);
  });
  qrTokens.forEach(function(qrToken) {
    var active = qrToken.IS_ACTIVE === true || String(qrToken.IS_ACTIVE).toLowerCase() === 'true';
    if (!active) return;
    var key = String(qrToken.STUDENT_ID).trim();
    if (!qrTokensByStudent[key]) qrTokensByStudent[key] = [];
    qrTokensByStudent[key].push(qrToken);
  });

  return {
    students: students,
    activities: activities,
    terms: terms,
    termById: terms.reduce(function(map, term) {
      map[String(term.TERM_ID).trim()] = term;
      return map;
    }, {}),
    scoresByStudent: scoresByStudent,
    qrTokensByStudent: qrTokensByStudent,
    studentById: studentById
  };
}

function applyReportHighlights_(sheet) {
  var inputColor = '#FFF2CC';
  var formulaColor = '#D9EAF7';
  var remarkColor = '#E2F0D9';
  [
    'C4:E5', 'D9:G9', 'D12:G12', 'I9:J9', 'I12:J12', 'L4:O12',
    'C18:F30', 'G16', 'I18:L22', 'M16', 'I27:L30', 'M25',
    'C36:F41', 'G34', 'I36:L41', 'M34'
  ].forEach(function(range) {
    sheet.getRange(range).setBackground(inputColor);
  });
  [
    'E31:F31', 'G31', 'K23:L23', 'M23', 'K31:L31', 'M31',
    'E42:F42', 'G42', 'K42:L42', 'M42',
    'G56:H56', 'G59:H59', 'G62:H62', 'G65:H65', 'G67:H67',
    'S18:S43'
  ].forEach(function(range) {
    sheet.getRange(range).setBackground(formulaColor);
  });
  ['H31', 'N23', 'N31', 'H42', 'N42', 'S23', 'S31', 'S36', 'S41']
    .forEach(function(range) {
      sheet.getRange(range).setBackground(remarkColor);
    });
}

function removeSheetIfExists_(workbook, name) {
  var existing = workbook.getSheetByName(name);
  if (existing && workbook.getSheets().length > 1) {
    workbook.deleteSheet(existing);
  }
}

function createPdfPart_(workbook, pageSheets, copyName, batchIndex) {
  var partBook = SpreadsheetApp.create(copyName + '_Part' + (batchIndex + 1));
  var partDefault = partBook.getSheets()[0];
  var copiedPages = [];
  var settingsSheet = workbook.getSheetByName('Report Settings');

  try {
    if (settingsSheet) {
      var copiedSettings = settingsSheet.copyTo(partBook);
      copiedSettings.setName('Report Settings');
      copiedSettings.hideSheet();
    }
    pageSheets.forEach(function(pageSheet) {
      var copied = pageSheet.copyTo(partBook);
      copied.setName(pageSheet.getName());
      copiedPages.push(copied);
    });

    if (copiedPages.length === 0) {
      throw new Error('No report pages were created for batch ' + (batchIndex + 1));
    }
    if (partDefault && partBook.getSheets().length > 1) {
      partBook.deleteSheet(partDefault);
    }
    SpreadsheetApp.flush();

    var blob = partBook.getAs('application/pdf');
    blob.setName(copyName + '_Part' + (batchIndex + 1) + '.pdf');
    var file = DriveApp.getRootFolder().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {
      part: batchIndex + 1,
      processed: pageSheets.length,
      url: 'https://drive.google.com/file/d/' + file.getId() + '/view'
    };
  } finally {
    try {
      DriveApp.getFileById(partBook.getId()).setTrashed(true);
    } catch (e) {
      Logger.log('Unable to trash temporary PDF workbook: ' + e);
    }
  }
}

function clearRangeContent(sheet, a1Notation) {
  try {
    sheet.getRange(a1Notation).clearContent();
  } catch (e) {
    Logger.log('clearRangeContent error [' + a1Notation + ']: ' + e);
  }
}

function clearImagesInRange(sheet, rangeA1) {
  try {
    var range = sheet.getRange(rangeA1);
    var firstRow = range.getRow();
    var lastRow = firstRow + range.getNumRows() - 1;
    var firstCol = range.getColumn();
    var lastCol = firstCol + range.getNumColumns() - 1;
    sheet.getImages().forEach(function(image) {
      var anchor = image.getAnchorCell();
      if (!anchor) return;
      var row = anchor.getRow();
      var col = anchor.getColumn();
      if (row >= firstRow && row <= lastRow && col >= firstCol && col <= lastCol) {
        image.remove();
      }
    });
  } catch (e) {
    Logger.log('clearImagesInRange error [' + rangeA1 + ']: ' + e);
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
function insertQRCode(sheet, student, qrTokensByStudent) {
  try {
    var startCol = COL.L;
    var startRow = 4;
    var endCol   = COL.O;
    var endRow   = 12;

    var qrTokens = qrTokensByStudent
      ? (qrTokensByStudent[String(student.STUDENT_ID).trim()] || [])
      : findRecords(CONFIG.SHEETS.QR_TOKENS, {
          STUDENT_ID: student.STUDENT_ID,
          IS_ACTIVE: true
        });
    if (qrTokens.length === 0) return; // no QR token — skip silently

    var token = qrTokens[0].TOKEN;
    var qrUrl = generateQRUrl(token);
    var boxWidth = 0;
    var boxHeight = 0;
    for (var column = startCol; column <= endCol; column++) {
      boxWidth += sheet.getColumnWidth(column);
    }
    for (var rowNumber = startRow; rowNumber <= endRow; rowNumber++) {
      boxHeight += sheet.getRowHeight(rowNumber);
    }

    clearImagesInRange(sheet, 'L4:O12');
    var qrBlob = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true }).getBlob();
    var qrImage = sheet.insertImage(qrBlob, startCol, startRow, 0, 0);
    var maxWidth = Math.max(1, boxWidth - 4);
    var maxHeight = Math.max(1, boxHeight - 4);
    var imageWidth = Math.max(1, qrImage.getWidth());
    var imageHeight = Math.max(1, qrImage.getHeight());
    var scale = Math.min(maxWidth / imageWidth, maxHeight / imageHeight);
    var finalWidth = Math.max(1, Math.floor(imageWidth * scale));
    var finalHeight = Math.max(1, Math.floor(imageHeight * scale));

    qrImage.setAnchorCell(sheet.getRange(startRow, startCol));
    qrImage.setWidth(finalWidth);
    qrImage.setHeight(finalHeight);
    qrImage.setAnchorCellXOffset(Math.max(0, Math.floor((boxWidth - finalWidth) / 2)));
    qrImage.setAnchorCellYOffset(Math.max(0, Math.floor((boxHeight - finalHeight) / 2)));
  } catch (e) {
    Logger.log('insertQRCode error for student ' + student.STUDENT_ID + ': ' + e);
  }
}

// ─── Term data builder ────────────────────────────────────────────────────────

/**
 * For a given termId, collect activities (sorted by order) and student scores.
 * Returns { activities: [{name, type, maxScore, rawScore}], weight, passingPercent }
 */
function buildTermData(termId, studentId, studentGradeLevel, context) {
  var gradeToMatch = String(studentGradeLevel || '').trim();
  var allActivities = context && context.activities
    ? context.activities
    : getAllRecords(CONFIG.SHEETS.ACTIVITIES);
  var termActivities = allActivities.filter(function(a) {
    if (String(a.TERM_ID).trim() !== String(termId).trim()) return false;
    if (!(a.IS_ACTIVE === true || a.IS_ACTIVE === 'TRUE' || a.IS_ACTIVE === 'true')) return false;
    var actGrade = String(a.GRADE_LEVEL || '').trim();
    if (actGrade !== '' && gradeToMatch !== '' && actGrade !== gradeToMatch) return false;
    return true;
  });
  termActivities.sort(function(a, b) { return (a.ACTIVITY_ORDER || 0) - (b.ACTIVITY_ORDER || 0); });

  var studentScores = context && context.scoresByStudent
    ? (context.scoresByStudent[String(studentId).trim()] || [])
    : findRecords(CONFIG.SHEETS.SCORES, { STUDENT_ID: studentId });

  var activities = termActivities.map(function(act) {
    var score = studentScores.find(function(s) {
      return String(s.ACTIVITY_ID).trim() === String(act.ACTIVITY_ID).trim();
    });
    return {
      name: act.ACTIVITY_NAME || '',
      type: act.ACTIVITY_TYPE || '',
      maxScore: parseFloat(act.MAX_SCORE) || 0,
      rawScore: score ? parseFloat(score.RAW_SCORE) : 0,
    };
  });

  var term = context && context.termById
    ? (context.termById[String(termId).trim()] || {})
    : findRecordById(CONFIG.SHEETS.GRADING_TERMS, 'TERM_ID', termId) || {};
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
function fillStudentColumns(sheet, student, terms, settings, stageNum, qrTokensByStudent) {
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
  insertQRCode(sheet, student, qrTokensByStudent);

  // ── Midterm Collective ──
  clearRangeContent(sheet, 'C18:D30');
  clearRangeContent(sheet, 'E18:F30');
  clearRangeContent(sheet, 'I18:J22');
  clearRangeContent(sheet, 'K18:L22');
  clearRangeContent(sheet, 'I27:J30');
  clearRangeContent(sheet, 'K27:L30');
  clearRangeContent(sheet, 'C36:D41');
  clearRangeContent(sheet, 'E36:F41');
  clearRangeContent(sheet, 'I36:J41');
  clearRangeContent(sheet, 'K36:L41');
  ['S18', 'S20', 'S21:S22', 'S23', 'S28', 'S29:S30', 'S31',
    'S33', 'S34:S35', 'S36', 'S38', 'S39:S40', 'S41', 'S43']
    .forEach(function(range) { clearRangeContent(sheet, range); });

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
  writeFormula(sheet, rowMC_tot, cMC_raw, '=SUM(E18:E30)');
  writeFormula(sheet, rowMC_tot, cMC_max, '=SUM(F18:F30)');
  writeFormula(sheet, rowMC_tot, cMC_eq, '=IF(OR(F31="",F31=0),"",ROUND(E31/F31*100*G16/100,2))');
  writeFormula(sheet, rowMC_tot, cMC_pf, '=IF(G31="","",IF(G31>=' + reportSettingFormula_('MIDTERM_COLLECTIVE_PASSING') + ',"PASS","FAIL"))');

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
  writeFormula(sheet, rowFCI_tot, cFCI_raw, '=SUM(K18:K22)');
  writeFormula(sheet, rowFCI_tot, cFCI_max, '=SUM(L18:L22)');
  writeFormula(sheet, rowFCI_tot, cFCI_eq, '=IF(OR(L23="",L23=0),"",ROUND(K23/L23*100*M16/100,2))');
  writeFormula(sheet, rowFCI_tot, cFCI_pf, '=IF(M23="","",IF(M23>=' + reportSettingFormula_('FINAL_COLLECTIVE_INITIAL_PASSING') + ',"PASS","FAIL"))');

  // ── Final Collective Final ──
  var fcf = terms.finCollFin;
  writeCell(sheet, rowFCF_wt, cFCF_wt, fcf.weight);
  fcf.activities.forEach(function(act, i) {
    var row = rowFCF_st + i;
    if (row >= rowFCF_tot) return;
    writeCell(sheet, row, cFCF_name, act.name);
    writeCell(sheet, row, cFCF_raw,  act.rawScore);
    writeCell(sheet, row, cFCF_max,  act.maxScore);
  });
  var fcfCalc = calcEquiv(fcf.activities, fcf.weight);
  writeFormula(sheet, rowFCF_tot, cFCF_raw, '=SUM(K27:K30)');
  writeFormula(sheet, rowFCF_tot, cFCF_max, '=SUM(L27:L30)');
  writeFormula(sheet, rowFCF_tot, cFCF_eq, '=IF(OR(L31="",L31=0),"",ROUND(K31/L31*100*M25/100,2))');
  writeFormula(sheet, rowFCF_tot, cFCF_pf, '=IF(M31="","",IF(M31>=' + reportSettingFormula_('FINAL_COLLECTIVE_FINAL_PASSING') + ',"PASS","FAIL"))');

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
  writeFormula(sheet, rowME_tot, cME_raw, '=SUM(E36:E41)');
  writeFormula(sheet, rowME_tot, cME_max, '=SUM(F36:F41)');
  writeFormula(sheet, rowME_tot, cME_eq, '=IF(OR(F42="",F42=0),"",ROUND(E42/F42*100*G34/100,2))');
  writeFormula(sheet, rowME_tot, cME_pf, '=IF(G42="","",IF(G42>=' + reportSettingFormula_('MIDTERM_EXAM_PASSING') + ',"PASS","FAIL"))');

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
  writeFormula(sheet, rowFE_tot, cFE_raw, '=SUM(K36:K41)');
  writeFormula(sheet, rowFE_tot, cFE_max, '=SUM(L36:L41)');
  writeFormula(sheet, rowFE_tot, cFE_eq, '=IF(OR(L42="",L42=0),"",ROUND(K42/L42*100*M34/100,2))');
  writeFormula(sheet, rowFE_tot, cFE_pf, '=IF(M42="","",IF(M42>=' + reportSettingFormula_('FINAL_EXAM_PASSING') + ',"PASS","FAIL"))');

  // ── Score Summary (rows 56-67) ──
  writeFormula(sheet, 56, cSum, '=G31');
  writeFormula(sheet, 59, cSum, '=G42');
  writeFormula(sheet, 62, cSum, '=IF(COUNT(M23,M31)=0,"",SUM(M23,M31))');
  writeFormula(sheet, 65, cSum, '=M42');
  writeFormula(sheet, 67, cSum, '=IF(COUNT(G56,G59,G62,G65)=0,"",ROUND(SUM(G56,G59,G62,G65),2))');

  // ── Semester Grade Summary Table (col S, rows 18-43) ──
  writeFormula(sheet, 18, cSGST, '=G31');
  writeFormula(sheet, 20, cSGST, '=G42');

  // Stage 1: Midterm Collective + Midterm Examination
  writeFormula(sheet, 21, cSGST, '=IF(COUNT(S18,S20)=0,"",SUM(S18,S20))');
  writeFormula(sheet, 23, cSGST, '=IF(S21="","",IF(S21>=' + reportSettingFormula_('STAGE1_PASSING') + ',"PASS","FAIL"))');

  if (stageNum >= 2) {
    writeFormula(sheet, 28, cSGST, '=M23');
    writeFormula(sheet, 29, cSGST, '=IF(COUNT(S21,S28)=0,"",SUM(S21,S28))');
    writeFormula(sheet, 31, cSGST, '=IF(S29="","",IF(S29>=' + reportSettingFormula_('STAGE2_PASSING') + ',"PASS","FAIL"))');
  }

  if (stageNum >= 3) {
    writeFormula(sheet, 33, cSGST, '=M31');
    writeFormula(sheet, 34, cSGST, '=IF(COUNT(S29,S33)=0,"",SUM(S29,S33))');
    writeFormula(sheet, 36, cSGST, '=IF(S34="","",IF(S34>=' + reportSettingFormula_('STAGE3_PASSING') + ',"PASS","FAIL"))');
  }

  if (stageNum >= 4) {
    writeFormula(sheet, 38, cSGST, '=M42');
    writeFormula(sheet, 39, cSGST, '=IF(COUNT(S34,S38)=0,"",SUM(S34,S38))');
    writeFormula(sheet, 41, cSGST, '=IF(S39="","",IF(S39>=' + reportSettingFormula_('STAGE4_PASSING') + ',"PASS","FAIL"))');
    writeFormula(sheet, 43, cSGST, '=S39');
  }

  applyReportHighlights_(sheet);
}

// ─── Main handler ─────────────────────────────────────────────────────────────

function handleGeneratePrintReport(payload, token) {
  try {
    var adminId = getAdminIdFromToken(token);
    var studentIds = payload.studentIds;
    var templateId = payload.templateId;
    var stageNum   = parseInt(payload.stage || 4, 10);
    var jobId      = payload.jobId || '';

    var batchIndex   = payload.batchIndex !== undefined ? parseInt(payload.batchIndex, 10) : 0;
    var totalBatches = payload.totalBatches !== undefined ? parseInt(payload.totalBatches, 10) : 1;
    var workbookId   = payload.workbookId || null;
    var startIndex   = payload.startIndex !== undefined
      ? parseInt(payload.startIndex, 10)
      : batchIndex * (studentIds ? studentIds.length : 0);

    if (!studentIds || !Array.isArray(studentIds) || studentIds.length === 0) {
      return createResponse({ success: false, message: 'Student IDs array is required', data: null });
    }
    if (!templateId) {
      return createResponse({ success: false, message: 'Template file ID is required. Set PRINT_TEMPLATE_ID in Settings.', data: null });
    }

    var completed = getCompletedPrintBatch_(jobId, batchIndex);
    if (completed) {
      return createResponse({ success: true, message: completed.message, data: completed.data });
    }

    // ── Settings + grading terms ──
    var settings = getSettingsObject();
    var context = buildPrintContext_();
    var allTerms = context.terms;
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
          ? buildTermData(term.TERM_ID, sid, studentGradeLevel, context)
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
      setReportSettingsSheet_(workbook, settings);
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
      setReportSettingsSheet_(workbook, settings);
    }

    // ── Fill one student per cloned page (Refactor 1: Student 2 removed) ──
    var errors = [];
    var processed = 0;
    var pageSheets = [];

    for (var i = 0; i < studentIds.length; i++) {
      var studentId = studentIds[i];
      var student = context.studentById[String(studentId).trim()];
      if (!student) {
        errors.push('Student (' + studentId + '): not found');
        continue;
      }

      var pageNumber = startIndex + i + 1;
      var pageName = 'Page ' + pageNumber;
      removeSheetIfExists_(workbook, pageName);
      var pageSheet = templateSheet.copyTo(workbook);
      pageSheet.setName(pageName);

      try {
        fillStudentColumns(
          pageSheet,
          student,
          buildAll(student),
          settings,
          stageNum,
          context.qrTokensByStudent
        );
        pageSheets.push(pageSheet);
        processed++;
      } catch (e) {
        errors.push('Student (' + studentId + '): ' + e.toString());
        try { workbook.deleteSheet(pageSheet); } catch (deleteError) { Logger.log(deleteError); }
      }
    }

    SpreadsheetApp.flush();

    var pdfPart = pageSheets.length > 0
      ? createPdfPart_(workbook, pageSheets, copyName + '_' + (jobId || 'run'), batchIndex)
      : null;
    var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + copyId + '/edit';
    var isLastBatch  = (totalBatches <= 1) || (batchIndex >= totalBatches - 1);

    if (isLastBatch) {
      // Remove the clean template sheet so it is not an empty page in the PDF
      try { workbook.deleteSheet(templateSheet); } catch (e) { /* ignore */ }
      SpreadsheetApp.flush();

      createAuditLog(
        adminId,
        'GENERATE_PRINT_REPORT',
        null, null, null,
        JSON.stringify({ stage: stageNum, processed: processed, batches: totalBatches, jobId: jobId }),
        'web',
        'Generated Stage ' + stageNum + ' report cards (' + totalBatches + ' batch' + (totalBatches !== 1 ? 'es' : '') + ')'
      );
    }

    var responseMessage = isLastBatch
      ? 'Report cards generated for ' + processed + ' student' + (processed !== 1 ? 's' : '') + (errors.length > 0 ? ' (' + errors.length + ' errors)' : '')
      : 'Batch ' + (batchIndex + 1) + ' of ' + totalBatches + ' done (' + processed + ' students)';
    var responseData = {
      processed: processed,
      errors: errors,
      batchIndex: batchIndex,
      totalBatches: totalBatches,
      isLastBatch: isLastBatch,
      workbookId: copyId,
      pdfUrl: pdfPart ? pdfPart.url : null,
      pdfPart: pdfPart,
      pdfParts: pdfPart ? [pdfPart] : [],
      spreadsheetUrl: sheetUrl,
      fileUrl: pdfPart ? pdfPart.url : null
    };
    saveCompletedPrintBatch_(jobId, batchIndex, {
      message: responseMessage,
      data: responseData
    });
    return createResponse({ success: true, message: responseMessage, data: responseData });

  } catch (error) {
    Logger.log('PrintService error: ' + error.toString());
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}