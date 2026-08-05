/**
 * SGMS Report Service
 * Handles report generation.
 *
 * FIX: handleGetClassReport and handleGetPassFailList now use a custom loose-
 * equality filter for GRADE_LEVEL / SECTION_NUMBER instead of the strict
 * findRecords() helper.  Google Sheets stores numbers as numbers; the frontend
 * sends them as numbers after Number(), but if any mismatch exists the strict
 * comparison in findRecords() would return an empty student list.
 */

function handleGetStudentReport(payload) {
  try {
    var studentId = payload.studentId;
    if (!studentId) {
      return createResponse({ success: false, message: 'Student ID is required', data: null });
    }

    var student = findRecordById(CONFIG.SHEETS.STUDENTS, 'STUDENT_ID', studentId);
    if (!student) {
      return createResponse({ success: false, message: 'Student not found', data: null });
    }

    var grades = calculateStudentGrades(studentId);
    var scores = findRecords(CONFIG.SHEETS.SCORES, { STUDENT_ID: studentId });
    var activities = getAllRecords(CONFIG.SHEETS.ACTIVITIES);

    // FIX: only show activities that apply to this student's grade level.
    //      An activity with no GRADE_LEVEL is shared across all grades; an
    //      activity with a specific GRADE_LEVEL only applies to that grade.
    var studentGradeLevel = String(student.GRADE_LEVEL || '').trim();
    var applicableActivities = activities.filter(function(a) {
      var actGrade = String(a.GRADE_LEVEL || '').trim();
      if (actGrade === '') return true;
      return actGrade === studentGradeLevel;
    });

    var scoreDetails = scores.map(function(score) {
      var activity = applicableActivities.find(function(a) { return a.ACTIVITY_ID === score.ACTIVITY_ID; });
      if (!activity) return null; // score belongs to an activity for a different grade
      return Object.assign({}, score, {
        activityName: activity.ACTIVITY_NAME,
        activityType: activity.ACTIVITY_TYPE,
        maxScore: activity.MAX_SCORE
      });
    }).filter(function(d) { return d !== null; });

    return createResponse({
      success: true,
      message: 'Student report generated',
      data: {
        student: student,
        grades: grades,
        scores: scoreDetails,
        generatedAt: new Date().toISOString()
      }
    });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

function handleGetClassReport(payload) {
  try {
    var gradeLevel = payload.gradeLevel;
    var section    = payload.section;

    if (!gradeLevel || !section) {
      return createResponse({ success: false, message: 'Grade level and section are required', data: null });
    }

    // FIX: loose == comparison so "1" == 1 matches regardless of storage type.
    var allStudents = getAllRecords(CONFIG.SHEETS.STUDENTS);
    var students = allStudents.filter(function(s) {
      // eslint-disable-next-line eqeqeq
      return s.GRADE_LEVEL == gradeLevel && s.SECTION_NUMBER == section && s.STATUS === 'Active';
    });

    var statistics = calculateClassStatistics(gradeLevel, section);

    var studentGrades = students.map(function(student) {
      try {
        var g = calculateStudentGrades(student.STUDENT_ID);
        return {
          studentId: student.STUDENT_ID,
          studentName: student.ENGLISH_NAME,
          overallGrade: g.overallGrade,
          passed: g.overallPassed
        };
      } catch (error) {
        return { studentId: student.STUDENT_ID, studentName: student.ENGLISH_NAME,
                 overallGrade: 0, passed: false, error: error.toString() };
      }
    });

    studentGrades.sort(function(a, b) { return b.overallGrade - a.overallGrade; });

    return createResponse({
      success: true,
      message: 'Class report generated',
      data: {
        gradeLevel: gradeLevel,
        section: section,
        statistics: statistics,
        students: studentGrades,
        generatedAt: new Date().toISOString()
      }
    });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}

function handleGetPassFailList(payload) {
  try {
    var gradeLevel = payload.gradeLevel;
    var section    = payload.section;
    var status     = payload.status;

    if (!gradeLevel || !section) {
      return createResponse({ success: false, message: 'Grade level and section are required', data: null });
    }

    // FIX: loose == comparison (same reason as above)
    var allStudents = getAllRecords(CONFIG.SHEETS.STUDENTS);
    var students = allStudents.filter(function(s) {
      // eslint-disable-next-line eqeqeq
      return s.GRADE_LEVEL == gradeLevel && s.SECTION_NUMBER == section && s.STATUS === 'Active';
    });

    var passingStudents = [];
    var failingStudents = [];

    students.forEach(function(student) {
      try {
        var grades = calculateStudentGrades(student.STUDENT_ID);
        var studentData = {
          studentId: student.STUDENT_ID,
          thaiName: student.THAI_NAME,
          englishName: student.ENGLISH_NAME,
          classNumber: student.CLASS_NUMBER,
          overallGrade: grades.overallGrade
        };
        if (grades.overallPassed) {
          passingStudents.push(studentData);
        } else {
          failingStudents.push(studentData);
        }
      } catch (error) {
        Logger.log('Error processing ' + student.STUDENT_ID + ': ' + error);
      }
    });

    passingStudents.sort(function(a, b) { return a.classNumber - b.classNumber; });
    failingStudents.sort(function(a, b) { return a.classNumber - b.classNumber; });

    var data = {};
    if (!status || status === 'pass') data.passingStudents = passingStudents;
    if (!status || status === 'fail') data.failingStudents = failingStudents;

    return createResponse({ success: true, message: 'Pass/Fail list generated', data: data });
  } catch (error) {
    return createResponse({ success: false, message: error.toString(), data: null });
  }
}
