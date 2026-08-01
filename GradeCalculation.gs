/**
 * SGMS Grade Calculation Engine
 * Handles all grade calculations with STRICT and PROGRESS modes.
 *
 * FIX: findRecords() uses strict === comparison, but Google Sheets may store
 * numbers as numbers while the payload sends strings (or vice-versa).
 * calculateClassStatistics() now uses a custom loose-equality filter for
 * GRADE_LEVEL and SECTION_NUMBER to avoid silent empty results.
 */

function handleCalculateGrades(payload) {
  try {
    const { studentId } = payload;
    if (!studentId) {
      return createResponse({ success: false, message: 'Student ID is required', data: null });
    }
    const grades = calculateStudentGrades(studentId);
    return createResponse({ success: true, message: 'Grades calculated', data: grades });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

function handleGetStudentGrade(payload) {
  try {
    const { studentId } = payload;
    if (!studentId) {
      return createResponse({ success: false, message: 'Student ID is required', data: null });
    }
    const grades = calculateStudentGrades(studentId);
    return createResponse({ success: true, message: 'Student grade retrieved', data: grades });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

/**
 * Calculate grades for a single student.
 * @param {string} studentId
 * @returns {object}
 */
function calculateStudentGrades(studentId) {
  const student = findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', studentId);
  if (!student) throw new Error('Student not found');

  const terms = getAllRecords(CONFIG.SHEETS.GRADING_TERMS);
  terms.sort(function(a, b) { return a.TERM_ORDER - b.TERM_ORDER; });

  const allActivities = getAllRecords(CONFIG.SHEETS.ACTIVITIES);
  const studentScores = findRecords(CONFIG.SHEETS.SCORES, { STUDENT_ID: studentId });

  const settings = getSettingsObject();
  const calculationMode = settings.CALCULATION_MODE || 'STRICT';
  const overallPassingPercent = parseFloat(settings.OVERALL_PASSING_PERCENT) || 50;

  const termGrades = [];
  var totalWeightedScore = 0;

  terms.forEach(function(term) {
    // IS_ACTIVE may be boolean true or the string "TRUE" depending on the sheet
    var termActivities = allActivities.filter(function(a) {
      return a.TERM_ID === term.TERM_ID && (a.IS_ACTIVE === true || a.IS_ACTIVE === 'TRUE' || a.IS_ACTIVE === 'true');
    });

    if (termActivities.length === 0) {
      termGrades.push({
        termId: term.TERM_ID,
        termName: term.TERM_NAME,
        weight: term.WEIGHT_PERCENT,
        rawPercentage: 0,
        weightedScore: 0,
        passed: false,
        activities: []
      });
      return;
    }

    var totalMaxScore = 0;
    var totalRawScore = 0;
    var activityDetails = [];

    termActivities.forEach(function(activity) {
      var score = studentScores.find(function(s) { return s.ACTIVITY_ID === activity.ACTIVITY_ID; });

      if (calculationMode === 'STRICT') {
        totalMaxScore += parseFloat(activity.MAX_SCORE);
        totalRawScore += score ? parseFloat(score.RAW_SCORE) : 0;
        activityDetails.push({
          activityId: activity.ACTIVITY_ID,
          activityName: activity.ACTIVITY_NAME,
          maxScore: activity.MAX_SCORE,
          rawScore: score ? score.RAW_SCORE : 0,
          hasScore: !!score
        });
      } else {
        // PROGRESS mode – only count activities that have scores
        if (score) {
          totalMaxScore += parseFloat(activity.MAX_SCORE);
          totalRawScore += parseFloat(score.RAW_SCORE);
          activityDetails.push({
            activityId: activity.ACTIVITY_ID,
            activityName: activity.ACTIVITY_NAME,
            maxScore: activity.MAX_SCORE,
            rawScore: score.RAW_SCORE,
            hasScore: true
          });
        }
      }
    });

    var rawPercentage = totalMaxScore > 0 ? (totalRawScore / totalMaxScore) * 100 : 0;
    var weightedScore = (rawPercentage * parseFloat(term.WEIGHT_PERCENT)) / 100;
    totalWeightedScore += weightedScore;

    var passingPercent = parseFloat(term.PASSING_PERCENT) || 50;
    termGrades.push({
      termId: term.TERM_ID,
      termName: term.TERM_NAME,
      weight: term.WEIGHT_PERCENT,
      rawPercentage: Math.round(rawPercentage * 100) / 100,
      weightedScore: Math.round(weightedScore * 100) / 100,
      passed: rawPercentage >= passingPercent,
      activities: activityDetails,
      totalMaxScore: totalMaxScore,
      totalRawScore: totalRawScore
    });
  });

  var overallGrade = Math.round(totalWeightedScore * 100) / 100;
  return {
    studentId: studentId,
    studentName: student.ENGLISH_NAME,
    calculationMode: calculationMode,
    termGrades: termGrades,
    overallGrade: overallGrade,
    overallPassed: overallGrade >= overallPassingPercent,
    overallPassingPercent: overallPassingPercent
  };
}

/**
 * Calculate class statistics.
 *
 * FIX: use loose == comparison for GRADE_LEVEL / SECTION_NUMBER so that
 * numeric values stored in Sheets match string parameters from the frontend.
 */
function calculateClassStatistics(gradeLevel, section) {
  var allStudents = getAllRecords(CONFIG.SHEETS.STUDENTS);

  // eslint-disable-next-line eqeqeq -- intentional loose comparison (type coercion)
  var students = allStudents.filter(function(s) {
    return s.GRADE_LEVEL == gradeLevel &&
           s.SECTION_NUMBER == section &&
           s.STATUS === 'Active';
  });

  if (students.length === 0) {
    return { totalStudents: 0, averageGrade: 0, highestGrade: 0, lowestGrade: 0,
             passingCount: 0, failingCount: 0, passingRate: 0 };
  }

  var grades = [];
  var passingCount = 0;

  students.forEach(function(student) {
    try {
      var gradeData = calculateStudentGrades(student.STUDENT_ID);
      grades.push(gradeData.overallGrade);
      if (gradeData.overallPassed) passingCount++;
    } catch (error) {
      Logger.log('Error calculating grades for ' + student.STUDENT_ID + ': ' + error);
    }
  });

  var averageGrade = grades.length > 0 ? grades.reduce(function(s, g) { return s + g; }, 0) / grades.length : 0;
  return {
    totalStudents: students.length,
    averageGrade: Math.round(averageGrade * 100) / 100,
    highestGrade: grades.length > 0 ? Math.max.apply(null, grades) : 0,
    lowestGrade: grades.length > 0 ? Math.min.apply(null, grades) : 0,
    passingCount: passingCount,
    failingCount: students.length - passingCount,
    passingRate: Math.round((passingCount / students.length) * 10000) / 100
  };
}
