// ============================================================================
// GOOGLE APPS SCRIPT - QA KPI LOGGING TOOL
// ============================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('UI')
    .setTitle('QA KPI Logging Tool')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================================
// CONFIGURATION
// ============================================================================

const ROSTER_SHEET_ID = '10plhrAaWj-gLj_pCBDberwW1Dr45DFCSriuXsdt1x7s';
const ROSTER_TAB_NAME = 'Global_QA_Roster';

const RESPONSES_SHEET_ID = '1eDSLrN-we-1J57F7LuI80QSVKEUGYRw3X3LKfazQAE0';
const ACCURACY_TAB_NAME = 'Accuracy % to Goal';
const PRODUCTIVITY_TAB_NAME = 'Productivity';
const QUALITY_TAB_NAME = 'Quality % to Goal';
const VOC_TAB_NAME = 'VOC % to Goal';
const TEAM_DEVELOPMENT_TAB_NAME = 'Team Development';
const PPD_TAB_NAME = 'PPD';

// ============================================================================
// ROSTER LOOKUP FUNCTIONS
// ============================================================================

/**
 * Returns detailed roster information for a specific WDID
 * Looks up by Column B (WDID) and returns:
 * - Name → Column C
 * - Manager → Column D
 * - Sr Manager → Column E
 * - Role → Column G
 * - Program → Column I
 * - Email → Column K
 * - Region → Column N
 */
function getRosterDetailsByWDID(wdid) {
  if (!wdid) return null;

  try {
    const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    const sheet = ss.getSheetByName(ROSTER_TAB_NAME);
    if (!sheet) throw new Error('Sheet Global_QA_Roster not found');

    const values = sheet.getDataRange().getValues();

    // Find the row where column B (WDID) matches
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowWdid = String(row[1]).trim(); // Column B (index 1)

      if (rowWdid === String(wdid).trim()) {
        return {
          wdid: rowWdid,
          name: String(row[2]).trim(),           // Column C (index 2)
          manager: String(row[3]).trim(),        // Column D (index 3)
          sr_manager: String(row[4]).trim(),     // Column E (index 4)
          role: String(row[6]).trim(),           // Column G (index 6)
          program: String(row[8]).trim(),        // Column I (index 8)
          email: String(row[10]).trim(),         // Column K (index 10)
          country: String(row[11]).trim(),       // Column L (index 11)
          region: String(row[13]).trim()         // Column N (index 13)
        };

      }
    }

    return null; // WDID not found
  } catch (error) {
    console.error('Error in getRosterDetailsByWDID:', error);
    throw error;
  }
}

/**
 * Returns all team members under a given TL WDID.
 * Looks up supervisor in column D (format: "Name (WDID)")
 */
function getTeamMembersByTlWdid(tlWdid) {
  if (!tlWdid) return [];

  try {
    const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    const sheet = ss.getSheetByName(ROSTER_TAB_NAME);
    if (!sheet) throw new Error('Sheet Global_QA_Roster not found');

    const values = sheet.getDataRange().getValues();
    const team = [];

    // Start at row 1 (skip header row 0)
    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const tmWdid = row[1];       // Column B (WDID)
      const tmName = row[2];       // Column C (Name)
      const supervisor = row[3];   // Column D (Supervisor - "Name (WDID)")

      if (!supervisor || !tmWdid || !tmName) continue;

      // Check if supervisor field contains the TL WDID
      if (String(supervisor).includes(`(${tlWdid})`)) {
        team.push({
          wdid: String(tmWdid).trim(),
          name: String(tmName).trim()
        });
      }
    }

    return team;
  } catch (error) {
    console.error('Error in getTeamMembersByTlWdid:', error);
    throw error;
  }
}

/**
 * Combined function: Validates TL WDID, gets TL details, and loads team
 */
function loadTLAndTeam(tlWdid) {
  try {
    // Get TL details
    const tlDetails = getRosterDetailsByWDID(tlWdid);
    if (!tlDetails) {
      throw new Error(`WDID ${tlWdid} not found in the Global QA Roster`);
    }

    // Get team members
    const teamMembers = getTeamMembersByTlWdid(tlWdid);

    return {
      success: true,
      tlDetails: tlDetails,
      teamMembers: teamMembers
    };
  } catch (error) {
    console.error('Error in loadTLAndTeam:', error);
    return {
      success: false,
      message: error.message || 'An error occurred while loading team data'
    };
  }
}

// ============================================================================
// SUBMISSION FUNCTIONS
// ============================================================================

/**
 * Submit Accuracy % to Goal KPIs
 * Appends rows to "Accuracy % to Goal" tab starting at row 3
 */
function submitAccuracy(payload) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(ACCURACY_TAB_NAME);
    if (!sheet) throw new Error('Accuracy % to Goal sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || ''; // Use TL email from roster

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
  // Skip if exempt
  if (member.exempt?.accuracy) {
    Logger.log(`Skipping ${member.wdid} - marked as exempt for Accuracy`);
    return;
  }

  const rosterDetails = getRosterDetailsByWDID(member.wdid);

  if (!rosterDetails) {
    throw new Error(`WDID ${member.wdid} not found in roster`);
  }


      // Build row according to required structure starting in column B (18 columns)
      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        payload.month,                                // D: report_month
        payload.year,                                 // E: report_year
        payload.tlWdid || '',                         // F: tl_wdid
        payload.tlName || '',                         // G: tl_name
        (getRosterDetailsByWDID(payload.tlWdid) || {}).manager || '',     // H: tl_manager
        (getRosterDetailsByWDID(payload.tlWdid) || {}).sr_manager || '',  // I: tl_sr_manager
        member.wdid,                                  // J: tm_wdid
        rosterDetails.name,                           // K: tm_name
        rosterDetails.program,                        // L: program
        rosterDetails.country,                        // M: country
        rosterDetails.region,                         // N: region
        member.accuracy.calAccurate || 0,             // O: cal_accurate
        member.accuracy.calPerformed || 0,            // P: cal_performed
        member.accuracy.validDisputes || 0,           // Q: valid_disputes
        member.accuracy.totalAudits || 0,             // R: total_completed_evals
        member.accuracy.accurateAudits || 0,          // S: accurate_evals
        member.accuracy.notes || ''                   // T: notes
      ];



      rowsToAppend.push(row);
    });

    // Find next available row (must be at least row 3)
    const startRow = getNextAppendRow_(sheet, 3, 2); // row 3, column B


    // Append rows starting from column B (column index 2)
    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 19).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} rows to Accuracy % to Goal starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow,
      message: `Successfully submitted ${rowsToAppend.length} Accuracy records`
    };

  } catch (error) {
    Logger.log('Error in submitAccuracy: ' + error.toString());
    console.error('Error in submitAccuracy:', error);
    throw error;
  }
}

/**
 * Submit Productivity KPIs
 * Appends rows to "Productivity" tab starting at row 3
 */
function submitProductivity(payload) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(PRODUCTIVITY_TAB_NAME);
    if (!sheet) throw new Error('Productivity sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || ''; // Use TL email from roster

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
  // Skip if exempt
  if (member.exempt?.productivity) {
    Logger.log(`Skipping ${member.wdid} - marked as exempt for Productivity`);
    return;
  }

  const rosterDetails = getRosterDetailsByWDID(member.wdid);

  if (!rosterDetails) {
    throw new Error(`WDID ${member.wdid} not found in roster`);
  }


      // Build row according to column mapping (B through T = 19 columns)
      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        payload.month,                                // D: report_month
        payload.year,                                 // E: report_year
        payload.tlWdid || '',                         // F: tl_wdid
        payload.tlName || '',                         // G: tl_name
        (getRosterDetailsByWDID(payload.tlWdid) || {}).manager || '',     // H: tl_manager
        (getRosterDetailsByWDID(payload.tlWdid) || {}).sr_manager || '',  // I: tl_sr_manager
        member.wdid,                                  // J: tm_wdid
        rosterDetails.name,                           // K: tm_name
        rosterDetails.program,                        // L: program
        rosterDetails.country,                        // M: country
        rosterDetails.region,                         // N: region
        member.productivity.qaCompleted || 0,         // O: qa_evals_completed
        member.productivity.qaTarget || 0,            // P: qa_evals_target
        member.productivity.vocCompleted || 0,        // Q: voc_evals_completed
        member.productivity.vocTarget || 0,           // R: voc_evals_target
        member.productivity.calCompleted || 0,        // S: calibrations_completed
        member.productivity.calTarget || 0,           // T: calibrations_target
        member.productivity.notes || ''               // U: notes
      ];


      rowsToAppend.push(row);
    });

    // Find next available row (must be at least row 3)
    const startRow = getNextAppendRow_(sheet, 3, 2); // row 3, column B


    // Append rows starting from column B (column index 2)
    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 20).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} rows to Productivity starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow,
      message: `Successfully submitted ${rowsToAppend.length} Productivity records`
    };

  } catch (error) {
    Logger.log('Error in submitProductivity: ' + error.toString());
    console.error('Error in submitProductivity:', error);
    throw error;
  }
}

/**
 * Submit Quality % to Goal KPIs
 * Appends rows to "Quality % to Goal" tab starting at row 3
 */
function submitQuality(payload) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(QUALITY_TAB_NAME);
    if (!sheet) throw new Error('Quality % to Goal sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
  // Skip if exempt
  if (member.exempt?.quality) {
    Logger.log(`Skipping ${member.wdid} - marked as exempt for Quality`);
    return;
  }

  const rosterDetails = getRosterDetailsByWDID(member.wdid);

  if (!rosterDetails) {
    throw new Error(`WDID ${member.wdid} not found in roster`);
  }


      // Build row according to required structure (A:O, 15 columns starting at B)
      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        payload.month,                                // D: report_month
        payload.year,                                 // E: report_year
        payload.tlWdid || '',                         // F: tl_wdid
        payload.tlName || '',                         // G: tl_name
        (getRosterDetailsByWDID(payload.tlWdid) || {}).manager || '',     // H: tl_manager
        (getRosterDetailsByWDID(payload.tlWdid) || {}).sr_manager || '',  // I: tl_sr_manager
        member.wdid,                                  // J: tm_wdid
        rosterDetails.name,                           // K: tm_name
        rosterDetails.program,                        // L: program
        rosterDetails.country,                        // M: country
        rosterDetails.region,                         // N: region
        member.quality.qualityPctToGoal || 0,         // O: quality_pct_to_goal
        member.quality.notes || ''                    // P: notes
      ];


      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 3, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 15).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} rows to Quality % to Goal starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow,
      message: `Successfully submitted ${rowsToAppend.length} Quality records`
    };

  } catch (error) {
    Logger.log('Error in submitQuality: ' + error.toString());
    console.error('Error in submitQuality:', error);
    throw error;
  }
}

/**
 * Submit VOC % to Goal KPIs
 * Appends rows to "VOC % to Goal" tab starting at row 3
 */
function submitVOC(payload) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(VOC_TAB_NAME);
    if (!sheet) throw new Error('VOC % to Goal sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
  // Skip if exempt
  if (member.exempt?.voc) {
    Logger.log(`Skipping ${member.wdid} - marked as exempt for VOC`);
    return;
  }

  const rosterDetails = getRosterDetailsByWDID(member.wdid);

  if (!rosterDetails) {
    throw new Error(`WDID ${member.wdid} not found in roster`);
  }


      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        payload.month,                                // D: report_month
        payload.year,                                 // E: report_year
        payload.tlWdid || '',                         // F: tl_wdid
        payload.tlName || '',                         // G: tl_name
        (getRosterDetailsByWDID(payload.tlWdid) || {}).manager || '',     // H: tl_manager
        (getRosterDetailsByWDID(payload.tlWdid) || {}).sr_manager || '',  // I: tl_sr_manager
        member.wdid,                                  // J: tm_wdid
        rosterDetails.name,                           // K: tm_name
        rosterDetails.program,                        // L: program
        rosterDetails.country,                        // M: country
        rosterDetails.region,                         // N: region
        member.voc.vocPctToGoal || 0,                 // O: voc_pct_to_goal
        member.voc.notes || ''                        // P: notes
      ];


      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 3, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 15).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} rows to VOC % to Goal starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow,
      message: `Successfully submitted ${rowsToAppend.length} VOC records`
    };

  } catch (error) {
    Logger.log('Error in submitVOC: ' + error.toString());
    console.error('Error in submitVOC:', error);
    throw error;
  }
}

/**
 * Submit Team Development KPIs
 * Appends rows to "Team Development" tab starting at row 3
 */
function submitTeamDevelopment(payload) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(TEAM_DEVELOPMENT_TAB_NAME);
    if (!sheet) throw new Error('Team Development sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
  // Skip if exempt
  if (member.exempt?.teamDevelopment) {
    Logger.log(`Skipping ${member.wdid} - marked as exempt for Team Development`);
    return;
  }

  const rosterDetails = getRosterDetailsByWDID(member.wdid);

  if (!rosterDetails) {
    throw new Error(`WDID ${member.wdid} not found in roster`);
  }


      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        payload.month,                                // D: report_month
        payload.year,                                 // E: report_year
        payload.tlWdid || '',                         // F: tl_wdid
        payload.tlName || '',                         // G: tl_name
        (getRosterDetailsByWDID(payload.tlWdid) || {}).manager || '',     // H: tl_manager
        (getRosterDetailsByWDID(payload.tlWdid) || {}).sr_manager || '',  // I: tl_sr_manager
        member.wdid,                                  // J: tm_wdid
        rosterDetails.name,                           // K: tm_name
        rosterDetails.program,                        // L: program
        rosterDetails.country,                        // M: country
        rosterDetails.region,                         // N: region
        member.teamDevelopment.workshopsDelivered || 0,  // O: workshops_delivered
        member.teamDevelopment.workshopsRequired || 0,   // P: workshops_required
        member.teamDevelopment.notes || ''               // Q: notes
      ];


      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 3, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 16).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} rows to Team Development starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow,
      message: `Successfully submitted ${rowsToAppend.length} Team Development records`
    };

  } catch (error) {
    Logger.log('Error in submitTeamDevelopment: ' + error.toString());
    console.error('Error in submitTeamDevelopment:', error);
    throw error;
  }
}

/**
 * Submit PPD KPIs
 * Appends rows to "PPD" tab starting at row 3
 */
function submitPPD(payload) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(PPD_TAB_NAME);
    if (!sheet) throw new Error('PPD sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
  // Skip if exempt
  if (member.exempt?.ppd) {
    Logger.log(`Skipping ${member.wdid} - marked as exempt for PPD`);
    return;
  }

  const rosterDetails = getRosterDetailsByWDID(member.wdid);

  if (!rosterDetails) {
    throw new Error(`WDID ${member.wdid} not found in roster`);
  }


      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        payload.month,                                // D: report_month
        payload.year,                                 // E: report_year
        payload.tlWdid || '',                         // F: tl_wdid
        payload.tlName || '',                         // G: tl_name
        (getRosterDetailsByWDID(payload.tlWdid) || {}).manager || '',     // H: tl_manager
        (getRosterDetailsByWDID(payload.tlWdid) || {}).sr_manager || '',  // I: tl_sr_manager
        member.wdid,                                  // J: tm_wdid
        rosterDetails.name,                           // K: tm_name
        rosterDetails.program,                        // L: program
        rosterDetails.country,                        // M: country
        rosterDetails.region,                         // N: region
        member.ppd.coursesCompleted || 0,             // O: courses_completed
        member.ppd.coursesRequired || 0,              // P: courses_required
        member.ppd.notes || ''                        // Q: notes
      ];


      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 3, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 16).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} rows to PPD starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow,
      message: `Successfully submitted ${rowsToAppend.length} PPD records`
    };

  } catch (error) {
    Logger.log('Error in submitPPD: ' + error.toString());
    console.error('Error in submitPPD:', error);
    throw error;
  }
}


/**
 * Master submit function that handles all 6 KPI types
 */
function submitKPIs(payload) {
  try {
    Logger.log('Starting KPI submission for TL WDID: ' + payload.tlWdid);

    // Get TL details from roster (including email)
    const tlDetails = getRosterDetailsByWDID(payload.tlWdid);
    if (!tlDetails) {
      throw new Error(`TL WDID ${payload.tlWdid} not found in roster`);
    }

    // Add TL name and email to payload
    payload.tlName = tlDetails.name;
    payload.tlEmail = tlDetails.email;

    Logger.log(`Submitting KPIs for ${payload.teamMembers.length} team members`);
    Logger.log(`Submitter email: ${payload.tlEmail}`);

    // Submit all 6 KPI types
    const accuracyResult = submitAccuracy(payload);
    const productivityResult = submitProductivity(payload);
    const qualityResult = submitQuality(payload);
    const vocResult = submitVOC(payload);
    const teamDevResult = submitTeamDevelopment(payload);
    const ppdResult = submitPPD(payload);

    Logger.log('Submission completed successfully');

    return {
      success: true,
      message: `Successfully submitted all KPIs for ${payload.teamMembers.length} team members`,
      details: {
        accuracy: accuracyResult,
        productivity: productivityResult,
        quality: qualityResult,
        voc: vocResult,
        teamDevelopment: teamDevResult,
        ppd: ppdResult
      }
    };

  } catch (error) {
    Logger.log('Error in submitKPIs: ' + error.toString());
    console.error('Error in submitKPIs:', error);
    return {
      success: false,
      message: error.message || 'An error occurred during submission'
    };
  }
}


function getNextAppendRow_(sheet, startRow, col) {
  const last = sheet.getLastRow();
  if (last < startRow) return startRow;

  const numRows = last - startRow + 1;
  const values = sheet.getRange(startRow, col, numRows, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '' && values[i][0] !== null) {
      return startRow + i + 1; // next row after last non-empty
    }
  }
  return startRow;
}

