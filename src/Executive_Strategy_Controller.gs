// ============================================================================
// GOOGLE APPS SCRIPT - QA KPI LOGGING TOOL (UNIFIED)
// ============================================================================
function doGet(e) {
  const mode = e.parameter.mode || 'qa'; // Default to QA TL mode
  
  if (mode === 'manager') {
    return HtmlService.createHtmlOutputFromFile('UI_QATL')
      .setTitle('Monthly QA TL KPI Submission | Manager View');
  } else {
    return HtmlService.createHtmlOutputFromFile('UI_QA')
      .setTitle('Monthly CX QA KPI Submission | Americas');
  }
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

function getTLsByManagerWdid(managerWdid) {
  if (!managerWdid) return [];

  try {
    const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    const sheet = ss.getSheetByName(ROSTER_TAB_NAME);
    if (!sheet) throw new Error('Sheet Global_QA_Roster not found');

    const values = sheet.getDataRange().getValues();
    const tls = [];

    Logger.log(`Looking for TLs reporting to Manager WDID: ${managerWdid}`);

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const tlWdid = row[1];       // Column B
      const tlName = row[2];       // Column C
      const supervisor = row[3];   // Column D
      const role = row[6];         // Column G
      const program = row[8];      // Column I
      const country = row[11];     // Column L
      const region = row[13];      // Column N

      if (!supervisor || !tlWdid || !tlName || !role) continue;

      // Check if supervisor field contains the Manager WDID
      if (String(supervisor).includes(`(${managerWdid})`)) {
        const roleLower = String(role).toLowerCase();
        
        // Filter: role must contain "lead" or "leader" but NOT "manager"
        if ((roleLower.includes('lead') || roleLower.includes('leader')) && 
            !roleLower.includes('manager')) {
          
          Logger.log(`✓ Found TL: ${tlName} (${tlWdid}) - Role: ${role}`);
          
          tls.push({
            wdid: String(tlWdid).trim(),
            name: String(tlName).trim(),
            role: String(role).trim(),
            program: String(program).trim(),
            country: String(country).trim(),
            region: String(region).trim()
          });
        }
      }
    }

    Logger.log(`Total TLs found: ${tls.length}`);
    return tls;

  } catch (error) {
    Logger.log('Error in getTLsByManagerWdid: ' + error.toString());
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

/**
 * Enhanced loadTLAndTeam that includes role-based UI routing info
 */
function loadTLAndTeamWithRole(tlWdid) {
  try {
    const tlDetails = getRosterDetailsByWDID(tlWdid);
    if (!tlDetails) {
      throw new Error(`WDID ${tlWdid} not found in the Global QA Roster`);
    }

    // Determine if this is a Manager role
    const isManager = tlDetails.role.toLowerCase().includes('manager');

    let teamMembers = [];

    if (isManager) {
      // Manager: Get their TL direct reports
      teamMembers = getTLsByManagerWdid(tlWdid);
    } else {
      // TL: Get their QA Analyst direct reports
      teamMembers = getTeamMembersByTlWdid(tlWdid);
    }

    return {
      success: true,
      tlDetails: tlDetails,
      teamMembers: teamMembers,
      isManager: isManager,
      uiType: isManager ? 'TL' : 'QA'
    };
  } catch (error) {
    console.error('Error in loadTLAndTeamWithRole:', error);
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

// ============================================================================
// TL KPI CONFIGURATION
// ============================================================================

const TL_KPI_SHEET_ID = '1kY5bFV2zb10yOjZw4GaGWvbW6YGD_BtWandt4JXvWsY';
const TL_PRODUCTIVITY_TAB = 'Productivity';
const TL_SERVICE_DELIVERY_TAB = 'Service Delivery';
const TL_PPD_TAB = 'PPD';
const TL_INNOVATION_TAB = 'Innovation';
const TL_ROLLUP_TAB = 'TL roll-up';

/**
 * Get TL's prior submitted roll-up values from TL roll-up sheet
 * Used in Service Delivery section to show read-only Quality and VOC averages
 */
function getTLPriorRollup(tlWdid) {
  try {
    const ss = SpreadsheetApp.openById(TL_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(TL_ROLLUP_TAB);
    if (!sheet) {
      Logger.log('TL roll-up sheet not found');
      return { quality_avg: '', voc_avg: '' };
    }

    const values = sheet.getDataRange().getValues();

    // Find row where column A (WDID) matches tlWdid (starting from row 3)
    for (let i = 2; i < values.length; i++) { // i=2 is row 3
      const rowWdid = String(values[i][0]).trim(); // Column A
      
      if (rowWdid === String(tlWdid).trim()) {
        return {
          quality_avg: values[i][7] || '', // Column H (index 7)
          voc_avg: values[i][10] || ''     // Column K (index 10)
        };
      }
    }

    // Not found
    return { quality_avg: '', voc_avg: '' };

  } catch (error) {
    Logger.log('Error in getTLPriorRollup: ' + error.toString());
    return { quality_avg: '', voc_avg: '' };
  }
}

// ============================================================================
// TL-LEVEL KPI SUBMISSION FUNCTIONS
// ============================================================================

/**
 * Submit TL Productivity KPIs
 * Writes to TL KPI workbook → Productivity tab
 * Range: B:T (19 columns)
 */
function submitTLProductivity(payload) {
  try {
    const ss = SpreadsheetApp.openById(TL_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(TL_PRODUCTIVITY_TAB);
    if (!sheet) throw new Error('TL Productivity sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.submitterEmail || '';

    const rowsToAppend = [];

    payload.tlList.forEach(tl => {
      // Skip if exempt = Yes and no data entered
      // But we ALWAYS write the row (even if exempt)
      
      const row = [
        timestamp,                          // B: timestamp
        submitterEmail,                     // C: submitter_email
        payload.month,                      // D: report_month
        payload.year,                       // E: report_year
        tl.wdid,                            // F: tl_wdid
        tl.name,                            // G: tl_name
        tl.manager || '',                   // H: tl_manager
        tl.sr_manager || '',                // I: tl_sr_manager
        tl.program || '',                   // J: program
        tl.country || '',                   // K: country
        tl.region || '',                    // L: region
        tl.productivity.qa_count || 0,      // M: qa_count
        tl.productivity.audits_completed || 0,    // N: audits_completed
        tl.productivity.audits_required || 0,     // O: audits_required
        tl.productivity.coachings_completed || 0, // P: coachings_completed
        tl.productivity.coachings_required || 0,  // Q: coachings_required
        tl.productivity.exempt || 'No',     // R: exempt
        tl.productivity.exception_reason || '', // S: exception_reason
        tl.productivity.notes || ''         // T: notes
      ];

      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 3, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 19).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} TL Productivity rows starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow
    };

  } catch (error) {
    Logger.log('Error in submitTLProductivity: ' + error.toString());
    throw error;
  }
}

/**
 * Submit TL Service Delivery KPIs
 * Writes to TL KPI workbook → Service Delivery tab
 * Range: B:P (15 columns)
 */
function submitTLServiceDelivery(payload) {
  try {
    const ss = SpreadsheetApp.openById(TL_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(TL_SERVICE_DELIVERY_TAB);
    if (!sheet) throw new Error('TL Service Delivery sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.submitterEmail || '';

    const rowsToAppend = [];

    payload.tlList.forEach(tl => {
      const row = [
        timestamp,                                  // B: timestamp
        submitterEmail,                             // C: submitter_email
        payload.month,                              // D: report_month
        payload.year,                               // E: report_year
        tl.wdid,                                    // F: tl_wdid
        tl.name,                                    // G: tl_name (formula-driven in sheet, but keep structure)
        tl.serviceDelivery.exempt || 'No',          // H: exempt
        tl.serviceDelivery.exemption_reason || '',  // I: exemption_reason
        tl.serviceDelivery.notes || '',             // J: notes
        tl.manager || '',                           // K: tl_manager (formula-driven)
        tl.sr_manager || '',                        // L: tl_sr_manager (formula-driven)
        tl.program || '',                           // M: program (formula-driven)
        tl.country || '',                           // N: country (formula-driven)
        tl.region || '',                            // O: region (formula-driven)
        tl.serviceDelivery.reporting_compliance || 'No' // P: reporting_compliance
      ];

      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 3, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 15).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} TL Service Delivery rows starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow
    };

  } catch (error) {
    Logger.log('Error in submitTLServiceDelivery: ' + error.toString());
    throw error;
  }
}

/**
 * Submit TL PPD KPIs
 * Writes to TL KPI workbook → PPD tab
 * Range: B:R (17 columns)
 */
function submitTLPPD(payload) {
  try {
    const ss = SpreadsheetApp.openById(TL_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(TL_PPD_TAB);
    if (!sheet) throw new Error('TL PPD sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.submitterEmail || '';

    const rowsToAppend = [];

    payload.tlList.forEach(tl => {
      const row = [
        timestamp,                          // B: timestamp
        submitterEmail,                     // C: submitter_email
        payload.month,                      // D: report_month
        payload.year,                       // E: report_year
        tl.wdid,                            // F: tl_wdid
        tl.name,                            // G: tl_name
        tl.ppd.exempt || 'No',              // H: exempt
        tl.ppd.exemption_reason || '',      // I: exemption_reason
        tl.ppd.notes || '',                 // J: notes
        tl.manager || '',                   // K: tl_manager
        tl.sr_manager || '',                // L: tl_sr_manager
        tl.program || '',                   // M: program
        tl.country || '',                   // N: country
        tl.region || '',                    // O: region
        tl.ppd.ppd_required_courses || 0,   // P: ppd_required_courses
        tl.ppd.ppd_mandatory_completed || 0,// Q: ppd_mandatory_completed
        tl.ppd.ppd_extra_completed || 0     // R: ppd_extra_completed
      ];

      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 3, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 17).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} TL PPD rows starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow
    };

  } catch (error) {
    Logger.log('Error in submitTLPPD: ' + error.toString());
    throw error;
  }
}

/**
 * Submit TL Innovation KPIs
 * Writes to TL KPI workbook → Innovation tab
 * Range: B:R (17 columns)
 */
function submitTLInnovation(payload) {
  try {
    const ss = SpreadsheetApp.openById(TL_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(TL_INNOVATION_TAB);
    if (!sheet) throw new Error('TL Innovation sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.submitterEmail || '';

    const rowsToAppend = [];

    payload.tlList.forEach(tl => {
      const row = [
        timestamp,                              // B: timestamp
        submitterEmail,                         // C: submitter_email
        payload.month,                          // D: report_month
        payload.year,                           // E: report_year
        tl.wdid,                                // F: tl_wdid
        tl.name,                                // G: tl_name
        tl.innovation.exempt || 'No',           // H: exempt
        tl.innovation.exemption_reason || '',   // I: exemption_reason
        tl.innovation.notes || '',              // J: notes
        tl.manager || '',                       // K: tl_manager
        tl.sr_manager || '',                    // L: tl_sr_manager
        tl.program || '',                       // M: program
        tl.country || '',                       // N: country
        tl.region || '',                        // O: region
        tl.innovation.innovation_high_count || 0,  // P: innovation_high_count
        tl.innovation.innovation_med_count || 0,   // Q: innovation_med_count
        tl.innovation.innovation_low_count || 0    // R: innovation_low_count
      ];

      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 3, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 17).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} TL Innovation rows starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow
    };

  } catch (error) {
    Logger.log('Error in submitTLInnovation: ' + error.toString());
    throw error;
  }
}

/**
 * Master submit function for TL-level KPIs
 */
function submitKPIs_TL(payload) {
  try {
    Logger.log('Starting TL KPI submission for Manager WDID: ' + payload.managerWdid);

    // Get manager details
    const managerDetails = getRosterDetailsByWDID(payload.managerWdid);
    if (!managerDetails) {
      throw new Error(`Manager WDID ${payload.managerWdid} not found in roster`);
    }

    payload.submitterEmail = managerDetails.email;

    Logger.log(`Submitting TL KPIs for ${payload.tlList.length} TL(s)`);

    // Submit all 4 TL KPI types
    const productivityResult = submitTLProductivity(payload);
    const serviceDeliveryResult = submitTLServiceDelivery(payload);
    const ppdResult = submitTLPPD(payload);
    const innovationResult = submitTLInnovation(payload);

    Logger.log('TL KPI submission completed successfully');

    return {
      success: true,
      message: `Successfully submitted TL KPIs for ${payload.tlList.length} Team Leader(s)`,
      details: {
        productivity: productivityResult,
        serviceDelivery: serviceDeliveryResult,
        ppd: ppdResult,
        innovation: innovationResult
      }
    };

  } catch (error) {
    Logger.log('Error in submitKPIs_TL: ' + error.toString());
    console.error('Error in submitKPIs_TL:', error);
    return {
      success: false,
      message: error.message || 'An error occurred during TL KPI submission'
    };
  }
}


function testManagerLookup() {
  const managerWdid = '10017175'; // Manager WDID to search for
  
  const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  const sheet = ss.getSheetByName(ROSTER_TAB_NAME);
  const values = sheet.getDataRange().getValues();
  
  Logger.log('=== Testing Manager Lookup (FULL ROSTER) ===');
  Logger.log('Looking for Manager WDID: ' + managerWdid);
  Logger.log('Total rows in roster: ' + values.length);
  Logger.log('');
  
  let foundCount = 0;
  let tlCount = 0;
  
  // Check ALL rows (not just first 20)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const wdid = row[1];
    const name = row[2];
    const supervisor = row[3];
    const role = row[6];
    
    if (!wdid || !name) continue; // Skip empty rows
    
    // Check if this person reports to the Manager
    if (String(supervisor).includes('(' + managerWdid + ')')) {
      foundCount++;
      const roleLower = String(role).toLowerCase();
      const isLead = roleLower.includes('lead');
      
      Logger.log(`✓ FOUND: ${name} (${wdid})`);
      Logger.log(`  Supervisor: "${supervisor}"`);
      Logger.log(`  Role: "${role}"`);
      Logger.log(`  Contains "lead"? ${isLead}`);
      
      if (isLead) {
        tlCount++;
        Logger.log(`  ✓✓ QUALIFIED AS TL`);
      } else {
        Logger.log(`  ✗ NOT A LEAD (excluded)`);
      }
      Logger.log('---');
    }
  }
  
  Logger.log('');
  Logger.log('=== SUMMARY ===');
  Logger.log(`Total direct reports found: ${foundCount}`);
  Logger.log(`Team Leaders (with "lead" in role): ${tlCount}`);
  Logger.log('===============');
  
  if (foundCount === 0) {
    Logger.log('');
    Logger.log('⚠️ NO DIRECT REPORTS FOUND!');
    Logger.log('Possible reasons:');
    Logger.log('1. Manager WDID 10017175 is not in anyone\'s Supervisor field (Column D)');
    Logger.log('2. The Supervisor field format is different than expected');
    Logger.log('3. This WDID might not be a Manager in the roster');
    Logger.log('');
    Logger.log('Checking if this WDID exists in the roster...');
    
    const managerDetails = getRosterDetailsByWDID(managerWdid);
    if (managerDetails) {
      Logger.log('✓ Manager WDID found in roster:');
      Logger.log('  Name: ' + managerDetails.name);
      Logger.log('  Role: ' + managerDetails.role);
      Logger.log('  Manager: ' + managerDetails.manager);
    } else {
      Logger.log('✗ Manager WDID NOT FOUND in roster!');
    }
  }
}

