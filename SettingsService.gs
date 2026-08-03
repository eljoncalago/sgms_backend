/**
 * SGMS Settings Service
 * Handles system settings management
 */

function initializeSettings() {
  const existingSettings = getAllRecords(CONFIG.SHEETS.SETTINGS);
  
  if (existingSettings.length > 0) {
    Logger.log('Settings already initialized');
    return;
  }
  
  // Create default settings
  for (let key in CONFIG.DEFAULT_SETTINGS) {
    const setting = {
      SETTING_KEY: key,
      SETTING_VALUE: String(CONFIG.DEFAULT_SETTINGS[key]),
      DESCRIPTION: getSettingDescription(key)
    };
    
    insertRecord(CONFIG.SHEETS.SETTINGS, setting);
  }
  
  Logger.log('Default settings initialized');
}

function getSettingDescription(key) {
  const descriptions = {
    SCHOOL_NAME: 'Name of the school',
    CALCULATION_MODE: 'Grade calculation mode: STRICT or PROGRESS',
    OVERALL_PASSING_PERCENT: 'Overall passing percentage',
    SYSTEM_VERSION: 'SGMS system version',
    THEME: 'UI theme: purple, midnight, amber, or forest',
    MIDTERM_COLLECTIVE_PASSING: 'Passing score for Midterm Collective component',
    FINAL_COLLECTIVE_INITIAL_PASSING: 'Passing score for Final Collective Initial component',
    FINAL_COLLECTIVE_FINAL_PASSING: 'Passing score for Final Collective Final component',
    MIDTERM_EXAM_PASSING: 'Passing score for Midterm Examination component',
    FINAL_EXAM_PASSING: 'Passing score for Final Examination component',
    STAGE1_PASSING: 'Passing score for Stage 1 (Midterm Collective + Midterm Exam)',
    STAGE2_PASSING: 'Passing score for Stage 2 (+ Final Collective Initial)',
    STAGE3_PASSING: 'Passing score for Stage 3 (+ Final Collective Final)',
    STAGE4_PASSING: 'Passing score for Stage 4 (+ Final Exam / Overall)'
  };
  
  return descriptions[key] || '';
}

function handleGetSettings() {
  try {
    const settings = getAllRecords(CONFIG.SHEETS.SETTINGS);
    
    // Convert to object format
    const settingsObj = {};
    settings.forEach(setting => {
      settingsObj[setting.SETTING_KEY] = setting.SETTING_VALUE;
    });
    
    return createResponse({
      success: true,
      message: 'Settings retrieved',
      data: settingsObj
    });
  } catch (error) {
    return createResponse({
      success: false,
      message: error.toString(),
      data: null
    });
  }
}

function handleUpdateSettings(payload, token) {
  try {
    const adminId = getAdminIdFromToken(token);
    const { settings } = payload;
    
    if (!settings || typeof settings !== 'object') {
      return createResponse({
        success: false,
        message: 'Settings object is required',
        data: null
      });
    }
    
    const updated = [];
    
    for (let key in settings) {
      try {
        const existingSettings = findRecords(CONFIG.SHEETS.SETTINGS, { SETTING_KEY: key });
        
        if (existingSettings.length > 0) {
          // Update existing setting
          const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
          const data = sheet.getDataRange().getValues();
          const headers = data[0];
          const keyIndex = headers.indexOf('SETTING_KEY');
          const valueIndex = headers.indexOf('SETTING_VALUE');
          
          for (let i = 1; i < data.length; i++) {
            if (data[i][keyIndex] === key) {
              sheet.getRange(i + 1, valueIndex + 1).setValue(String(settings[key]));
              updated.push(key);
              break;
            }
          }
        } else {
          // Create new setting
          insertRecord(CONFIG.SHEETS.SETTINGS, {
            SETTING_KEY: key,
            SETTING_VALUE: String(settings[key]),
            DESCRIPTION: getSettingDescription(key)
          });
          updated.push(key);
        }
      } catch (error) {
        Logger.log(`Error updating setting ${key}: ${error}`);
      }
    }
    
    createAuditLog(
      adminId,
      'UPDATE_SETTINGS',
      null,
      null,
      null,
      JSON.stringify(settings),
      'web',
      `Updated ${updated.length} settings`
    );
    
    return createResponse({
      success: true,
      message: 'Settings updated',
      data: { updated: updated }
    });
  } catch (error) {
    return createResponse({
      success: false,
      message: error.toString(),
      data: null
    });
  }
}

function getSettingsObject() {
  const settings = getAllRecords(CONFIG.SHEETS.SETTINGS);
  const settingsObj = {};
  settings.forEach(setting => {
    settingsObj[setting.SETTING_KEY] = setting.SETTING_VALUE;
  });
  return settingsObj;
}

function handleGetDashboardStats() {
  try {
    const students = getAllRecords(CONFIG.SHEETS.STUDENTS);
    const activeStudents = students.filter(s => s.STATUS === 'Active');
    const activities = getAllRecords(CONFIG.SHEETS.ACTIVITIES);
    const activeActivities = activities.filter(a => a.IS_ACTIVE);
    const scores = getAllRecords(CONFIG.SHEETS.SCORES);
    
    // Calculate some basic statistics
    const uniqueGrades = [...new Set(activeStudents.map(s => s.GRADE_LEVEL))];
    const uniqueClasses = [...new Set(activeStudents.map(s => `${s.GRADE_LEVEL}/${s.SECTION_NUMBER}`))];
    
    // Calculate average grade (sample from active students)
    let totalGrade = 0;
    let gradedStudents = 0;
    let passingCount = 0;
    
    // Sample up to 50 students for performance
    const sampleStudents = activeStudents.slice(0, 50);
    
    sampleStudents.forEach(student => {
      try {
        const gradeData = calculateStudentGrades(student.STUDENT_ID);
        totalGrade += gradeData.overallGrade;
        gradedStudents++;
        if (gradeData.overallPassed) {
          passingCount++;
        }
      } catch (error) {
        Logger.log(`Error calculating grade for ${student.STUDENT_ID}`);
      }
    });
    
    const averageGrade = gradedStudents > 0 ? totalGrade / gradedStudents : 0;
    const passingRate = gradedStudents > 0 ? (passingCount / gradedStudents) * 100 : 0;
    
    const stats = {
      totalStudents: activeStudents.length,
      totalClasses: uniqueClasses.length,
      totalActivities: activeActivities.length,
      averageGrade: Math.round(averageGrade * 100) / 100,
      passingRate: Math.round(passingRate * 100) / 100,
      totalScores: scores.length,
      gradeLevels: uniqueGrades.sort()
    };
    
    return createResponse({
      success: true,
      message: 'Dashboard statistics retrieved',
      data: stats
    });
  } catch (error) {
    return createResponse({
      success: false,
      message: error.toString(),
      data: null
    });
  }
}
