// ============================================================================
// GOOGLE APPS SCRIPT - QA KPI LOGGING TOOL (UNIFIED)
// ============================================================================
function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || 'qa';

  if (mode === 'manager') {
    return HtmlService.createHtmlOutputFromFile('UI_QATL')
      .setTitle('Monthly QA TL KPI Submission | Manager View')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (mode === 'director') {
    return HtmlService.createHtmlOutputFromFile('UI_QAMNGR')
      .setTitle('Monthly QA Manager KPI Submission | Director View')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (mode === 'regional') {
    return HtmlService.createHtmlOutputFromFile('UI_REGIONAL')
      .setTitle('Regional QA Lead KPI Submission | CX & T&S Quality')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (mode === 'directorself') {
    return HtmlService.createHtmlOutputFromFile('UI_DIRECTOR')
      .setTitle('CX & T&S Quality Director KPI Submission')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (mode === 'globalprime') {
    return HtmlService.createHtmlOutputFromFile('UI_GLOBALPRIME')
      .setTitle('Global Prime KPI Submission')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutputFromFile('UI_QA')
      .setTitle('Monthly CX QA KPI Submission | Americas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ============================================================================
// CONFIGURATION
// ============================================================================

// ✅ COST % TO REVENUE FEATURE FLAG
// Set to true to enable Cost % to Revenue section in Director form
// Set to false to hide it (e.g., during most of the year)
const ENABLE_COST_TO_REVENUE = true;  // Change to false to disable

const ROSTER_SHEET_ID = '10plhrAaWj-gLj_pCBDberwW1Dr45DFCSriuXsdt1x7s';
const ROSTER_TAB_NAME = 'Global_QA_Roster';

const RESPONSES_SHEET_ID = '1eDSLrN-we-1J57F7LuI80QSVKEUGYRw3X3LKfazQAE0';
const ACCURACY_TAB_NAME = 'Accuracy % to Goal';
const PRODUCTIVITY_TAB_NAME = 'Productivity';
const QUALITY_TAB_NAME = 'Quality % to Goal';
const VOC_TAB_NAME = 'VOC % to Goal';
const TEAM_DEVELOPMENT_TAB_NAME = 'Team Development';
const PPD_TAB_NAME = 'PPD';

// Manager KPI Sheet Configuration
const MANAGER_KPI_SHEET_ID = '1wHo7bKOwBLwBRrrhVYKO6eI9G05NReoj6I6N4B7i3F8';
const MNGR_INNOVATION_TAB = 'Innovation';
const MNGR_COST_TO_REVENUE_TAB = 'Cost_to_Revenue';
const MNGR_PPD_TAB = 'PPD';

// Service Delivery Configuration (Read-Only for Managers)
const SERVICE_DELIVERY_SHEET_ID = '1OJnP40k-cpHXXpFLKq-vrrTlxGzGqwKnVxYzNbTPYNI';
const SERVICE_DELIVERY_TAB = 'Service Delivery';

// ============================================================================
// GLOBAL PRIME KPI CONFIGURATION
// ============================================================================

// Global Prime KPI Sheet Configuration
const GLOBAL_PRIME_SHEET_ID = '1h45mDFXSCGcHDGf07cbDdF6s08k_Zpj1FAsjrGhIASI';
const GLOBAL_PRIME_TAB = 'Global_Prime_KPIs';

// Hardcoded reporting period (update monthly)
const GLOBAL_PRIME_REPORT_MONTH = 'Jan';
const GLOBAL_PRIME_REPORT_YEAR = '2026';

// ============================================================================
// QA ROSTER SNAPSHOT CONFIGURATION (NEW DESTINATION)
// ============================================================================
const QA_SNAPSHOT_SHEET_ID = '1lQ4rirUuSzNH8l0hQJ18Qu_NjHI9Xzg9vYRaBYXJJnM';
const QA_SNAPSHOT_TAB_NAME = 'QA_Snapshot';

// ============================================================================
// ROSTER LOOKUP FUNCTIONS
// ============================================================================

/**
 * Returns detailed roster information for a specific WDID
 * Uses cached roster index for O(1) lookup (instead of O(n) sheet read)
 * 
 * Looks up by Column B (WDID) and returns:
 * - Name → Column C
 * - Manager → Column E
 * - Sr Manager → Column D
 * - Department Head → Column F
 * - Role → Column G
 * - Program → Column I
 * - Email → Column K
 * - Country → Column L
 * - Region → Column N
 */
function getRosterDetailsByWDID(wdid) {
  if (!wdid) return null;

  try {
    // ✅ Load roster index (cached per execution)
    const rosterIndex = getRosterIndex_();
    const details = rosterIndex.get(String(wdid).trim());

    if (!details) {
      Logger.log(`WDID ${wdid} not found in roster`);
      return null;
    }

    return details;

  } catch (error) {
    Logger.log('Error in getRosterDetailsByWDID: ' + error.toString());
    throw error;
  }
}

/**
 * Load entire roster into a Map (once per execution)
 * Returns: Map(wdid → detailsObject)
 * 
 * Details object includes:
 * - wdid, name, email, country, region, program, role
 * - sup_l1 (Column D), sup_l2 (Column E), sup_l3 (Column F) — neutral level names
 * - manager, sr_manager, department_head — DEPRECATED (kept for backward compat)
 * 
 * @returns {Map} Map of WDID → roster details
 */
function getRosterIndex_() {
  try {
    const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    const sheet = ss.getSheetByName(ROSTER_TAB_NAME);
    if (!sheet) throw new Error('Sheet Global_QA_Roster not found');

    const values = sheet.getDataRange().getValues();
    const index = new Map();

    // Start at row 1 (skip header row 0)
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowWdid = String(row[1]).trim(); // Column B (index 1)

      if (!rowWdid) continue; // Skip empty rows

      const details = {
        wdid: rowWdid,
        name: String(row[2]).trim(),           // Column C (index 2)
        // ✅ NEW: Neutral level names (reflecting actual roster columns)
        sup_l1: String(row[3]).trim(),         // Column D (index 3) = immediate supervisor
        sup_l2: String(row[4]).trim(),         // Column E (index 4) = Sr Manager
        sup_l3: String(row[5]).trim(),         // Column F (index 5) = Department Head
        // ✅ DEPRECATED: Kept for backward compatibility only
        manager: String(row[4]).trim(),        // Column E (index 4)
        sr_manager: String(row[3]).trim(),     // Column D (index 3)
        department_head: String(row[5]).trim(), // Column F (index 5)
        role: String(row[6]).trim(),           // Column G (index 6)
        program: String(row[8]).trim(),        // Column I (index 8)
        email: String(row[10]).trim(),         // Column K (index 10)
        country: String(row[11]).trim(),       // Column L (index 11)
        region: String(row[13]).trim()         // Column N (index 13)
      };

      index.set(rowWdid, details);
    }

    Logger.log(`Loaded roster index: ${index.size} entries`);
    return index;

  } catch (error) {
    Logger.log('Error in getRosterIndex_: ' + error.toString());
    throw error;
  }
}

/**
 * Map hierarchy levels to role-relative titles
 * Interprets sup_l1/sup_l2/sup_l3 based on context (who is being reported on)
 * 
 * @param {Object} personDetails - Roster details object (from getRosterIndex_)
 * @param {string} context - One of: "QA_SHEETS", "TL_SHEETS", "MNGR_SHEETS"
 * @returns {Object} { immediate, sr, deptHead } — role-relative hierarchy
 * 
 * MAPPING LOGIC:
 * - QA_SHEETS (QA being reported on):
 *   immediate = sup_l1 (their direct leader/TL)
 *   sr = sup_l2 (Manager above them)
 *   deptHead = sup_l3 (Sr Manager above them)
 * 
 * - TL_SHEETS (TL being reported on):
 *   immediate = sup_l1 (their direct manager)
 *   sr = sup_l2 (Sr Manager above them)
 *   deptHead = sup_l3 (Department Head above them)
 * 
 * - MNGR_SHEETS (Manager being reported on):
 *   immediate = sup_l1 (their direct manager)
 *   sr = sup_l2 (Sr Manager above them)
 *   deptHead = sup_l3 (Department Head above them)
 */
function mapHierarchy_(personDetails, context) {
  if (!personDetails) {
    return { immediate: '', sr: '', deptHead: '' };
  }

  // All contexts map the same way (neutral interpretation of levels)
  // sup_l1 = immediate supervisor (one level up)
  // sup_l2 = two levels up
  // sup_l3 = three levels up
  return {
    immediate: personDetails.sup_l1 || '',
    sr: personDetails.sup_l2 || '',
    deptHead: personDetails.sup_l3 || ''
  };
}

/**
 * Returns all team members under a given TL WDID.
 * Uses cached roster index for efficiency.
 * Looks up supervisor in column D (format: "Name (WDID)")
 * 
 * @param {string} tlWdid - Team Leader's WDID
 * @returns {Array} Array of team members with wdid, name, program
 */
function getTeamMembersByTlWdid(tlWdid) {
  if (!tlWdid) return [];

  try {
    const rosterIndex = getRosterIndex_();
    const team = [];
    const tlWdidStr = String(tlWdid).trim();

    // Iterate through all roster entries
    rosterIndex.forEach((details, wdid) => {
      // Check if this person's supervisor field contains the TL WDID
      // Supervisor is in Column D (stored as "Name (WDID)")
      const supervisor = String(details.sr_manager).trim(); // Column D

      if (supervisor.includes(`(${tlWdidStr})`)) {
        team.push({
          wdid: wdid,
          name: details.name,
          program: details.program || 'N/A'
        });
      }
    });

    Logger.log(`Found ${team.length} team members for TL ${tlWdid}`);
    return team;

  } catch (error) {
    Logger.log('Error in getTeamMembersByTlWdid: ' + error.toString());
    throw error;
  }
}

/**
 * Returns all TL direct reports under a Manager WDID
 * Uses cached roster index for efficiency
 * Filters by role (must contain "lead" but not "manager")
 * 
 * @param {string} managerWdid - Manager's WDID
 * @returns {Array} Array of TL objects
 */
function getTLsByManagerWdid(managerWdid) {
  if (!managerWdid) return [];

  try {
    const rosterIndex = getRosterIndex_();
    const tls = [];
    const managerWdidStr = String(managerWdid).trim();

    Logger.log(`Looking for TLs reporting to Manager WDID: ${managerWdidStr}`);

    rosterIndex.forEach((details, wdid) => {
      const supervisor = String(details.sr_manager).trim(); // Column D
      const role = String(details.role).toLowerCase();

      // Check if this person reports to the Manager
      if (!supervisor.includes(`(${managerWdidStr})`)) return;

      // Filter: role must contain "lead" or "leader" but NOT "manager"
      if ((role.includes('lead') || role.includes('leader')) && !role.includes('manager')) {
        Logger.log(`✓ Found TL: ${details.name} (${wdid}) - Role: ${details.role}`);

        tls.push({
          wdid: wdid,
          name: details.name,
          role: details.role,
          program: details.program,
          country: details.country,
          region: details.region
        });
      }
    });

    Logger.log(`Total TLs found: ${tls.length}`);
    return tls;

  } catch (error) {
    Logger.log('Error in getTLsByManagerWdid: ' + error.toString());
    throw error;
  }
}

/**
 * Returns all QA Managers reporting to a Director or Sr Manager
 * Uses cached roster index for efficiency
 * Filters by role (must contain "manager")
 * 
 * @param {string} directorWdid - Director's WDID
 * @returns {Array} Array of Manager objects
 */
function getManagersByDirectorWdid(directorWdid) {
  if (!directorWdid) return [];

  try {
    const rosterIndex = getRosterIndex_();
    const managers = [];
    const directorWdidStr = String(directorWdid).trim();

    Logger.log(`Looking for QA Managers reporting to Director WDID: ${directorWdidStr}`);

    rosterIndex.forEach((details, wdid) => {
      const supervisor = String(details.sr_manager).trim(); // Column D
      const role = String(details.role).toLowerCase();

      // Check if this person reports to the Director
      if (!supervisor.includes(`(${directorWdidStr})`)) return;

      // Filter: role must contain "manager"
      if (!role.includes('manager')) {
        Logger.log(`  ✗ Skipped ${details.name} (${wdid}) - Role "${details.role}" doesn't contain "manager"`);
        return;
      }

      Logger.log(`✓ Found Manager: ${details.name} (${wdid}) - Role: ${details.role}`);

      managers.push({
        wdid: wdid,
        name: details.name,
        manager: details.sr_manager,      // Column D → manager field
        sr_manager: details.manager,      // Column E → sr_manager field
        department_head: details.department_head, // Column F
        role: details.role,
        program: details.program,
        country: details.country,
        region: details.region
      });
    });

    Logger.log(`Total QA Managers found: ${managers.length}`);
    return managers;

  } catch (error) {
    Logger.log('Error in getManagersByDirectorWdid: ' + error.toString());
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

/**
 * Enhanced role detection for Directors/Sr Managers
 * Detects three levels:
 * - QA TL (role contains "lead" but not "manager")
 * - Manager (role contains "manager" but not "sr" or "director")
 * - Director/Sr Manager (role contains "sr" or "director")
 */
function loadDirectorAndTeamWithRole(wdid) {
  try {
    const userDetails = getRosterDetailsByWDID(wdid);
    if (!userDetails) {
      throw new Error(`WDID ${wdid} not found in the Global QA Roster`);
    }

    const roleLower = userDetails.role.toLowerCase();

    Logger.log(`User: ${userDetails.name} (${wdid})`);
    Logger.log(`Role: ${userDetails.role}`);
    Logger.log(`Role (lowercase): ${roleLower}`);

    // Determine role level
    const isDirector = roleLower.includes('sr') || roleLower.includes('director');
    const isManager = roleLower.includes('manager') && !isDirector;
    const isTL = (roleLower.includes('lead') || roleLower.includes('leader')) && !isManager && !isDirector;

    Logger.log(`isDirector: ${isDirector}`);
    Logger.log(`isManager: ${isManager}`);
    Logger.log(`isTL: ${isTL}`);

    let teamMembers = [];
    let uiType = 'QA'; // default

    if (isDirector) {
      // Director/Sr Manager: Get their Manager direct reports
      Logger.log('Loading managers for Director/Sr Manager...');
      teamMembers = getManagersByDirectorWdid(wdid);
      uiType = 'MNGR';
    } else if (isManager) {
      // Manager: Get their TL direct reports
      Logger.log('Loading TLs for Manager...');
      teamMembers = getTLsByManagerWdid(wdid);
      uiType = 'TL';
    } else {
      // TL: Get their QA Analyst direct reports
      Logger.log('Loading team members for TL...');
      teamMembers = getTeamMembersByTlWdid(wdid);
      uiType = 'QA';
    }

    Logger.log(`Team members found: ${teamMembers.length}`);

    return {
      success: true,
      userDetails: userDetails,
      teamMembers: teamMembers,
      isDirector: isDirector,
      isManager: isManager,
      isTL: isTL,
      uiType: uiType
    };

  } catch (error) {
    Logger.log('Error in loadDirectorAndTeamWithRole: ' + error.toString());
    console.error('Error in loadDirectorAndTeamWithRole:', error);
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
 * Appends rows to "Accuracy % to Goal" tab starting at row 2
 */
function submitAccuracy(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(ACCURACY_TAB_NAME);
    if (!sheet) throw new Error('Accuracy % to Goal sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
      const rosterDetails = rosterIndex.get(String(member.wdid).trim());  // ✅ Fast lookup

      if (!rosterDetails) {
        throw new Error(`WDID ${member.wdid} not found in roster`);
      }

      // ✅ Calculate applicable components count
      const componentNA = member.componentNA?.accuracy || {};
      const totalComponents = 3; // calibrations, disputes, audits
      const naCount = (componentNA.calibrations ? 1 : 0) +
        (componentNA.disputes ? 1 : 0) +
        (componentNA.audits ? 1 : 0);
      const applicableCount = totalComponents - naCount;

       const tlDetails = rosterIndex.get(String(payload.tlWdid).trim());  // ✅ Use passed index
       const tlHierarchy = mapHierarchy_(tlDetails, "QA_SHEETS"); 

      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        getMonthAbbreviation_(payload.month),         // D: report_month
        payload.year,                                 // E: report_year
        payload.tlWdid || '',                         // F: tl_wdid
        payload.tlName || '',                         // G: tl_name
        tlHierarchy.immediate || '',
        tlHierarchy.sr || '',
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
        member.accuracy.notes || '',                  // T: notes
        'Yes',                                        // U: kpi_applicable ✅ MOVED
        member.exempt?.accuracy ? 'Yes' : 'No',       // V: exempt ✅ MOVED
        member.exempt?.accuracy ? (member.accuracy.notes || '') : '',  // W: reason ✅ MOVED
        componentNA.calibrations ? 'No' : 'Yes',      // X: calibrations_na ✅ MOVED
        componentNA.disputes ? 'No' : 'Yes',          // Y: disputes_na ✅ MOVED
        componentNA.audits ? 'No' : 'Yes'            // Z: audits_na ✅ MOVED
      ];

      rowsToAppend.push(row);
    });

    // Find next available row (must be at least row 2)
    const startRow = getNextAppendRow_(sheet, 2, 1);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 1, rowsToAppend.length, 25).setValues(rowsToAppend);  // (A:Y)
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
 * Appends rows to "Productivity" tab starting at row 2
 */
function submitProductivity(payload, rosterIndex) {
  try {

     // ✅ ADD DIAGNOSTIC LOG:
    Logger.log(`submitProductivity called with ${payload.teamMembers.length} members`);
    Logger.log(`First member productivity: ${JSON.stringify(payload.teamMembers[0]?.productivity)}`);
    
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(PRODUCTIVITY_TAB_NAME);
    if (!sheet) throw new Error('Productivity sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
      const rosterDetails = rosterIndex.get(String(member.wdid).trim());  // ✅ Fast lookup

      if (!rosterDetails) {
        throw new Error(`WDID ${member.wdid} not found in roster`);
      }

      // ✅ Calculate applicable components count
      const componentNA = member.componentNA?.productivity || {};
      const totalComponents = 3; // qa_evals, voc, calibrations
      const naCount = (componentNA.qa_evals ? 1 : 0) +
        (componentNA.voc ? 1 : 0) +
        (componentNA.calibrations ? 1 : 0);
      const applicableCount = totalComponents - naCount;

       const tlDetails = rosterIndex.get(String(payload.tlWdid).trim());  // ✅ Use passed index
       const tlHierarchy = mapHierarchy_(tlDetails, "QA_SHEETS"); 

      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        getMonthAbbreviation_(payload.month),         // D: report_month
        payload.year,                                 // E: report_year
        payload.tlWdid || '',                         // F: tl_wdid
        payload.tlName || '',                         // G: tl_name
        tlHierarchy.immediate || '',  // ✅ H: tl_manager (Column D)
        tlHierarchy.sr || '',
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
        member.productivity.notes || '',              // U: notes
        'Yes',                                        // V: kpi_applicable ✅ MOVED
        member.exempt?.productivity ? 'Yes' : 'No',   // W: exempt ✅ MOVED
        member.exempt?.productivity ? (member.productivity.notes || '') : '',  // X: reason ✅ MOVED
        componentNA.qa_evals ? 'No' : 'Yes',          // Y: qa_evals_na ✅ MOVED
        componentNA.voc ? 'No' : 'Yes',               // Z: voc_na ✅ MOVED
        componentNA.calibrations ? 'No' : 'Yes'      // AA: calibrations_na ✅ MOVED
      ];

      rowsToAppend.push(row);
    });

    // Find next available row (must be at least row 2)
    const startRow = getNextAppendRow_(sheet, 2, 1);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 1, rowsToAppend.length, 26).setValues(rowsToAppend);  // (A:Z)
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
 * Appends rows to "Quality % to Goal" tab starting at row 2
 */
function submitQuality(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(QUALITY_TAB_NAME);
    if (!sheet) throw new Error('Quality % to Goal sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
      const rosterDetails = rosterIndex.get(String(member.wdid).trim());  // ✅ Fast lookup

      if (!rosterDetails) {
        throw new Error(`WDID ${member.wdid} not found in roster`);
      }

      Logger.log(`Quality data for ${member.wdid}:`);
      Logger.log(`  qualityPctToGoal (actual): ${member.quality.qualityPctToGoal}`);
      Logger.log(`  qualityProgramTarget: ${member.quality.qualityProgramTarget}`);
      Logger.log(`  Type of qualityProgramTarget: ${typeof member.quality.qualityProgramTarget}`);

       const tlDetails = rosterIndex.get(String(payload.tlWdid).trim());  // ✅ Use passed index
       const tlHierarchy = mapHierarchy_(tlDetails, "QA_SHEETS"); 

      const row = [
        timestamp,                                    // B
        submitterEmail,                               // C
        getMonthAbbreviation_(payload.month),         // D
        payload.year,                                 // E
        payload.tlWdid || '',                         // F
        payload.tlName || '',                         // G
        tlHierarchy.immediate || '',  // ✅ H: tl_manager (Column D)
        tlHierarchy.sr || '',
        member.wdid,                                  // J
        rosterDetails.name,                           // K
        rosterDetails.program,                        // L
        rosterDetails.country,                        // M
        rosterDetails.region,                         // N
        member.quality.qualityPctToGoal || 0,         // O: quality_score_actual
        member.quality.qualityProgramTarget || 0,     // P: quality_program_target
        member.quality.notes || '',                   // Q: notes
        member.notApplicable?.quality ? 'No' : 'Yes', // R: kpi_applicable ✅ MOVED
        member.exempt?.quality ? 'Yes' : 'No',        // S: exempt ✅ MOVED
        (member.notApplicable?.quality || member.exempt?.quality) ? (member.quality.notes || '') : ''  // T: reason ✅ MOVED
      ];



      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 2, 1);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 1, rowsToAppend.length, 19).setValues(rowsToAppend);  // (A:S)
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
 * Appends rows to "VOC % to Goal" tab starting at row 2
 */
function submitVOC(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(VOC_TAB_NAME);
    if (!sheet) throw new Error('VOC % to Goal sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
      const rosterDetails = rosterIndex.get(String(member.wdid).trim());  // ✅ Fast lookup

      if (!rosterDetails) {
        throw new Error(`WDID ${member.wdid} not found in roster`);
      }

      Logger.log(`VOC data for ${member.wdid}:`);
      Logger.log(`  vocPctToGoal (actual): ${member.voc.vocPctToGoal}`);
      Logger.log(`  vocProgramTarget: ${member.voc.vocProgramTarget}`);
      Logger.log(`  Type of vocProgramTarget: ${typeof member.voc.vocProgramTarget}`);

       const tlDetails = rosterIndex.get(String(payload.tlWdid).trim());  // ✅ Use passed index
       const tlHierarchy = mapHierarchy_(tlDetails, "QA_SHEETS"); 

      const row = [
        timestamp,                                    // B
        submitterEmail,                               // C
        getMonthAbbreviation_(payload.month),         // D
        payload.year,                                 // E
        payload.tlWdid || '',                         // F
        payload.tlName || '',                         // G
        tlHierarchy.immediate || '',  // ✅ H: tl_manager (Column D)
        tlHierarchy.sr || '',
        member.wdid,                                  // J
        rosterDetails.name,                           // K
        rosterDetails.program,                        // L
        rosterDetails.country,                        // M
        rosterDetails.region,                         // N
        member.voc.vocPctToGoal || 0,                 // O: voc_score_actual
        member.voc.vocProgramTarget || 0,             // P: voc_program_target
        member.voc.notes || '',                       // Q: notes
        member.notApplicable?.voc ? 'No' : 'Yes',     // R: kpi_applicable ✅ MOVED
        member.exempt?.voc ? 'Yes' : 'No',            // S: exempt ✅ MOVED
        (member.notApplicable?.voc || member.exempt?.voc) ? (member.voc.notes || '') : ''  // T: reason ✅ MOVED
      ];



      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 2, 1);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 1, rowsToAppend.length, 19).setValues(rowsToAppend);  // (A:S)
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
 * Appends rows to "Team Development" tab starting at row 2
 */
function submitTeamDevelopment(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(TEAM_DEVELOPMENT_TAB_NAME);
    if (!sheet) throw new Error('Team Development sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
      const rosterDetails = rosterIndex.get(String(member.wdid).trim());  // ✅ Fast lookup

      if (!rosterDetails) {
        throw new Error(`WDID ${member.wdid} not found in roster`);
      }

 const tlDetails = rosterIndex.get(String(payload.tlWdid).trim());  // ✅ Use passed index
 const tlHierarchy = mapHierarchy_(tlDetails, "QA_SHEETS"); 

      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        getMonthAbbreviation_(payload.month),         // D: report_month
        payload.year,                                 // E: report_year
        payload.tlWdid || '',                         // F: tl_wdid
        payload.tlName || '',                         // G: tl_name
        tlHierarchy.immediate || '',  // ✅ H: tl_manager (Column D)
        tlHierarchy.sr || '',
        member.wdid,                                  // J: tm_wdid
        rosterDetails.name,                           // K: tm_name
        rosterDetails.program,                        // L: program
        rosterDetails.country,                        // M: country
        rosterDetails.region,                         // N: region
        member.teamDevelopment.workshopsDelivered || 0,  // O: workshops_delivered
        member.teamDevelopment.workshopsRequired || 0,   // P: workshops_required
        member.teamDevelopment.notes || '',              // Q: notes
        'Yes',                                        // R: kpi_applicable ✅ MOVED
        member.exempt?.teamDevelopment ? 'Yes' : 'No', // S: exempt ✅ MOVED
        member.exempt?.teamDevelopment ? (member.teamDevelopment.notes || '') : ''  // T: reason ✅ MOVED
      ];


      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 2, 1);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 1, rowsToAppend.length, 19).setValues(rowsToAppend);  // (A:S)
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
 * Appends rows to "PPD" tab starting at row 2
 */
function submitPPD(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName(PPD_TAB_NAME);
    if (!sheet) throw new Error('PPD sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const rowsToAppend = [];

    payload.teamMembers.forEach(member => {
      const rosterDetails = rosterIndex.get(String(member.wdid).trim());  // ✅ Fast lookup

      if (!rosterDetails) {
        throw new Error(`WDID ${member.wdid} not found in roster`);
      }

 const tlDetails = rosterIndex.get(String(payload.tlWdid).trim());  // ✅ Use passed index
 const tlHierarchy = mapHierarchy_(tlDetails, "QA_SHEETS"); 

      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        getMonthAbbreviation_(payload.month),         // D: report_month
        payload.year,                                 // E: report_year
        payload.tlWdid || '',                         // F: tl_wdid
        payload.tlName || '',                         // G: tl_name
        tlHierarchy.immediate || '',  // ✅ H: tl_manager (Column D)
        tlHierarchy.sr || '',
        member.wdid,                                  // J: tm_wdid
        rosterDetails.name,                           // K: tm_name
        rosterDetails.program,                        // L: program
        rosterDetails.country,                        // M: country
        rosterDetails.region,                         // N: region
        member.ppd.coursesCompleted || 0,             // O: courses_completed
        member.ppd.coursesRequired || 0,              // P: courses_required
        member.ppd.notes || '',                       // Q: notes
        'Yes',                                        // R: kpi_applicable 
        member.exempt?.ppd ? 'Yes' : 'No',            // S: exempt 
        member.exempt?.ppd ? (member.ppd.notes || '') : ''  // T: reason 
      ];


      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 2, 1);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 1, rowsToAppend.length, 19).setValues(rowsToAppend);  // (A:S)
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
 * Submit TL Self-Report Productivity KPIs
 * Appends rows to "QA_TL_Productivity" tab starting at row 2
 */
function submitTLSelfProductivity(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(RESPONSES_SHEET_ID);
    const sheet = ss.getSheetByName('QA_TL_Productivity');
    if (!sheet) throw new Error('QA_TL_Productivity sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';

    const tlSelf = payload.tlSelfReport;
    const tlDetails = rosterIndex.get(String(tlSelf.wdid).trim());
    const tlHierarchy = mapHierarchy_(tlDetails, "TL_SHEETS");  // ✅ TL context

    // Build row according to column mapping (B:R = 17 columns)
    const row = [
      timestamp,                                    // B: timestamp
      submitterEmail,                               // C: submitter_email
      getMonthAbbreviation_(payload.month),         // D: report_month
      payload.year,                                 // E: report_year
      tlSelf.wdid,                                  // F: tl_wdid
      tlSelf.name,                                  // G: tl_name
      tlHierarchy.immediate || '',                  // H: tl_manager
      tlHierarchy.sr || '',                         // I: tl_sr_manager
      tlSelf.program || '',                         // J: program
      tlSelf.country || '',                         // K: country
      tlSelf.region || '',                          // L: region
      tlSelf.tlSelfProductivity.qaCount || 0,       // M: qa_count
      tlSelf.tlSelfProductivity.auditsCompleted || 0, // N: audits_completed
      tlSelf.tlSelfProductivity.auditsRequired || 0,  // O: audits_required
      tlSelf.exempt ? 'Yes' : 'No',                 // P: exempt
      tlSelf.exempt ? (tlSelf.tlSelfProductivity.notes || '') : '', // Q: exception_reason
      tlSelf.tlSelfProductivity.notes || ''         // R: notes
    ];

    const startRow = getNextAppendRow_(sheet, 2, 1);

    sheet.getRange(startRow, 2, 1, 17).setValues([row]);

    Logger.log(`Submitted TL Self-Report Productivity to row ${startRow}`);

    return {
      success: true,
      rowsWritten: 1,
      startRow: startRow,
      message: 'Successfully submitted TL Self-Report Productivity'
    };

  } catch (error) {
    Logger.log('Error in submitTLSelfProductivity: ' + error.toString());
    console.error('Error in submitTLSelfProductivity:', error);
    throw error;
  }
}
/**
 * Master submit function that handles all 6 KPI types + Roster Snapshot
 * ✅ Loads roster index ONCE and passes to all sub-functions
 */
function submitKPIs(payload) {
  try {
    Logger.log('Starting KPI submission for TL WDID: ' + payload.tlWdid);

    // ✅ LOAD ROSTER INDEX ONCE (shared across all functions)
    const rosterIndex = getRosterIndex_();
    Logger.log(`Roster index loaded: ${rosterIndex.size} entries`);

    // Get TL details from roster (including email)
    const tlDetails = rosterIndex.get(String(payload.tlWdid).trim());
    if (!tlDetails) {
      throw new Error(`TL WDID ${payload.tlWdid} not found in roster`);
    }

    // Add TL name and email to payload
    payload.tlName = tlDetails.name;
    payload.tlEmail = tlDetails.email;

    Logger.log(`Submitting KPIs for ${payload.teamMembers.length} team members`);
    Logger.log(`Submitter email: ${payload.tlEmail}`);

    // Submit all 6 team member KPI types
    const accuracyResult = submitAccuracy(payload, rosterIndex);
    const productivityResult = submitProductivity(payload, rosterIndex);
    const qualityResult = submitQuality(payload, rosterIndex);
    const vocResult = submitVOC(payload, rosterIndex);
    const teamDevResult = submitTeamDevelopment(payload, rosterIndex);
    const ppdResult = submitPPD(payload, rosterIndex);

    // ✅ Submit TL self-report
    const tlSelfResult = submitTLSelfProductivity(payload, rosterIndex);

    // ✅ NEW: Write roster snapshot (using original roster from payload)
    let rosterSnapshotResult = { success: true, rowsWritten: 0 };

    if (payload.teamRoster && payload.teamRoster.length > 0) {
      rosterSnapshotResult = writeRosterSnapshot_(payload, payload.teamRoster, rosterIndex);

      // Log warning if roster snapshot failed (but don't fail the whole submission)
      if (!rosterSnapshotResult.success) {
        Logger.log('⚠️ WARNING: Roster snapshot failed but KPI submission succeeded');
      }
    } else {
      Logger.log('⚠️ WARNING: No teamRoster in payload, roster snapshot skipped');
    }

    Logger.log('Submission completed successfully');

    return {
      success: true,
      message: `Successfully submitted all KPIs for ${payload.teamMembers.length} team members + TL self-report`,
      rosterSnapshotWarning: !rosterSnapshotResult.success ? rosterSnapshotResult.message : null,
      details: {
        accuracy: accuracyResult,
        productivity: productivityResult,
        quality: qualityResult,
        voc: vocResult,
        teamDevelopment: teamDevResult,
        ppd: ppdResult,
        tlSelfProductivity: tlSelfResult,
        rosterSnapshot: rosterSnapshotResult
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
// Manager Roll-up Configuration (Read-Only for Managers)
const MNGR_ROLLUP_TAB = 'MNGR_roll_up';
const MNGR_ROLLUP_BY_PROGRAM_TAB = 'MNGR_roll_up_by_program';
// QA Roster Snapshot Configuration
const TL_PPD_TAB = 'PPD';
const TL_INNOVATION_TAB = 'Innovation';

// ============================================================================
// TL-LEVEL KPI SUBMISSION FUNCTIONS
// ============================================================================

/**
 * Submit TL PPD KPIs
 * Writes to TL KPI workbook → PPD tab
 * Range: B:Q (16 columns)
 */
function submitTLPPD(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(TL_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(TL_PPD_TAB);
    if (!sheet) throw new Error('TL PPD sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.submitterEmail || '';

    const rowsToAppend = [];

    payload.tlList.forEach(tl => {
      const tlDetails = rosterIndex.get(String(tl.wdid).trim());  // ✅ Use cache
      const tlHierarchy = mapHierarchy_(tlDetails, "TL_SHEETS");  // ✅ TL context

      const row = [
        timestamp,                          // B: timestamp
        submitterEmail,                     // C: submitter_email
        getMonthAbbreviation_(payload.month),  // D: report_month
        payload.year,                       // E: report_year
        tl.wdid,                            // F: tl_wdid
        tl.name,                            // G: tl_name
        tlHierarchy.immediate || '',        // H: tl_manager
        tlHierarchy.sr || '',               // I: tl_sr_manager
        tl.program || '',                   // J: program
        tl.country || '',                   // K: country
        tl.region || '',                    // L: region
        tl.ppd.ppd_mandatory_completed || 0,   // M: ppd_mandatory_completed 
        tl.ppd.ppd_required_courses || 0,      // N: ppd_required_courses
        tl.ppd.exempt || 'No',              // O: exempt
        tl.ppd.exemption_reason || '',      // P: exemption_reason
        tl.ppd.notes || ''                  // Q: notes
      ];

      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 2, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 16).setValues(rowsToAppend);
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
function submitTLInnovation(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(TL_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(TL_INNOVATION_TAB);
    if (!sheet) throw new Error('TL Innovation sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.submitterEmail || '';

    const rowsToAppend = [];

    payload.tlList.forEach(tl => {
      const tlDetails = rosterIndex.get(String(tl.wdid).trim());  // ✅ Use cache
      const tlHierarchy = mapHierarchy_(tlDetails, "TL_SHEETS");  // ✅ TL context

      const row = [
        timestamp,                              // B: timestamp
        submitterEmail,                         // C: submitter_email
        getMonthAbbreviation_(payload.month),   // D: report_month
        payload.year,                           // E: report_year
        tl.wdid,                                // F: tl_wdid
        tl.name,                                // G: tl_name
        tlHierarchy.immediate || '',            // H: tl_manager
        tlHierarchy.sr || '',                   // I: tl_sr_manager
        tl.program || '',                       // J: program
        tl.country || '',                       // K: country
        tl.region || '',                        // L: region
        tl.innovation.innovation_high_count || 0,  // M: innovation_high_count
        tl.innovation.innovation_med_count || 0,   // N: innovation_med_count
        tl.innovation.innovation_low_count || 0,   // O: innovation_low_count
        tl.innovation.exempt || 'No',           // P: exempt
        tl.innovation.exemption_reason || '',   // Q: exemption_reason
        tl.innovation.notes || ''               // R: notes
      ];

      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 2, 2);

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
 * Write QA Roster Snapshot to new destination
 * Uses cached roster index for efficiency
 * 
 * @param {Object} payload - Submission payload with TL and team info
 * @param {Array} teamRoster - Array of team members from the UI (already loaded)
 * @returns {Object} Result object with success status
 */
function writeRosterSnapshot_(payload, teamRoster, rosterIndex) {
  try {
    Logger.log('Starting roster snapshot write...');

    if (!teamRoster || teamRoster.length === 0) {
      Logger.log('No team roster provided, skipping snapshot');
      return { success: true, rowsWritten: 0, message: 'No roster to snapshot' };
    }

    const ss = SpreadsheetApp.openById(QA_SNAPSHOT_SHEET_ID);
    const sheet = ss.getSheetByName(QA_SNAPSHOT_TAB_NAME);
    if (!sheet) throw new Error(
      `Sheet "${QA_SNAPSHOT_TAB_NAME}" not found in workbook ${QA_SNAPSHOT_SHEET_ID}`
    );

    const timestamp = new Date();
    const submitterEmail = payload.tlEmail || '';
    const tlWdid = payload.tlWdid;
    const tlName = payload.tlName;
    const tlDetails = rosterIndex.get(String(tlWdid).trim()) || {};
    const tlManager = tlDetails.sr_manager || '';
    const tlSrManager = tlDetails.manager || '';

    // ── STEP 1: Read existing keys from QA_SNAPSHOT sheet ──────────────────
    const existingKeys = new Set();
    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      sheet.getRange(2, 14, lastRow - 1, 1).getValues()
        .forEach(r => { const k = String(r[0]).trim(); if (k) existingKeys.add(k); });
    }
    Logger.log(`Found ${existingKeys.size} existing roster snapshot keys`);

    // ── STEP 2: Build rows (new keys only) ──────────────────────────────────
    const rowsToAppend = [];
    let skippedCount = 0;

    teamRoster.forEach(qa => {
      const qaDetails = rosterIndex.get(String(qa.wdid).trim());  // ✅ Use cache
      if (!qaDetails) {
        Logger.log(`Warning: QA WDID ${qa.wdid} not found in roster, skipping`);
        return;
      }

      const monthAbbr = getMonthAbbreviation_(payload.month);
      const key = `${payload.year}|${monthAbbr}|TL${tlWdid}|QA${qa.wdid}|${qaDetails.program}`;

      if (existingKeys.has(key)) { skippedCount++; return; }

      rowsToAppend.push([
        timestamp, submitterEmail, String(payload.year), monthAbbr,
        String(tlWdid), tlName, tlManager, tlSrManager,
        String(qa.wdid), qaDetails.name, qaDetails.program,
        qaDetails.country, qaDetails.region, key
      ]);
    });

    // ── STEP 3: Write to QA_SNAPSHOT sheet only ────────────────────────────
    if (rowsToAppend.length > 0) {
      const startRow = getNextAppendRow_(sheet, 2, 1);
      sheet.getRange(startRow, 1, rowsToAppend.length, 14).setValues(rowsToAppend);
      Logger.log(`Wrote ${rowsToAppend.length} rows to QA_SNAPSHOT_SHEET_ID at row ${startRow}`);
    }

    if (skippedCount > 0) Logger.log(`Skipped ${skippedCount} duplicate roster entries`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      skippedDuplicates: skippedCount,
      message: `Roster snapshot saved: ${rowsToAppend.length} new, ${skippedCount} duplicates skipped`
    };

  } catch (error) {
    Logger.log('Error in writeRosterSnapshot_: ' + error.toString());
    return { success: false, message: 'Roster snapshot failed: ' + error.message };
  }
}

/**
 * Helper: Convert month name to 3-letter abbreviation
 * @param {string} monthName - Full month name (e.g., "December")
 * @returns {string} 3-letter abbreviation (e.g., "Dec")
 */
function getMonthAbbreviation_(monthName) {
  const monthMap = {
    'january': 'Jan', 'february': 'Feb', 'march': 'Mar', 'april': 'Apr',
    'may': 'May', 'june': 'Jun', 'july': 'Jul', 'august': 'Aug',
    'september': 'Sep', 'october': 'Oct', 'november': 'Nov', 'december': 'Dec'
  };

  const monthLower = String(monthName).toLowerCase().trim();
  return monthMap[monthLower] || monthName.substring(0, 3);
}

/**
 * Master submit function for TL-level KPIs
 */
function submitKPIs_TL(payload) {
  try {
    Logger.log('Starting TL KPI submission for Manager WDID: ' + payload.managerWdid);

    // ✅ Load roster index once
    const rosterIndex = getRosterIndex_();

    // Get manager details
    const managerDetails = rosterIndex.get(String(payload.managerWdid).trim());
    if (!managerDetails) {
      throw new Error(`Manager WDID ${payload.managerWdid} not found in roster`);
    }

    payload.submitterEmail = managerDetails.email;

    Logger.log(`Submitting TL KPIs for ${payload.tlList.length} TL(s)`);

    // ✅ normalizes missing KPI objects so .exempt never crashes
    payload.tlList = (payload.tlList || []).map(tl => ({
      ...tl,
      ppd: tl?.ppd || {},
      innovation: tl?.innovation || {}
    }));

    // ✅ Pass rosterIndex to both functions
    const ppdResult = submitTLPPD(payload, rosterIndex);
    const innovationResult = submitTLInnovation(payload, rosterIndex);

    Logger.log('TL KPI submission completed successfully');

    return {
      success: true,
      message: `Successfully submitted TL KPIs for ${payload.tlList.length} Team Leader(s)`,
      details: {
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

// ========================================================================
// MANAGER KPI SUBMISSION FUNCTIONS
// ========================================================================

/**
 * Submit Manager Innovation KPIs
 * Writes to Manager KPI workbook → Innovation tab
 * Range: B:R (17 columns)
 */
function submitManagerInnovation(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(MANAGER_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(MNGR_INNOVATION_TAB);
    if (!sheet) throw new Error('Manager Innovation sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.submitterEmail || '';

    const rowsToAppend = [];

    payload.managerList.forEach(mngr => {
      const managerDetails = rosterIndex.get(String(mngr.wdid).trim());  // ✅ Use cache
      const mngrHierarchy = mapHierarchy_(managerDetails, "MNGR_SHEETS");  // ✅ Manager context

      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        getMonthAbbreviation_(payload.month),         // D: report_month
        payload.year,                                 // E: report_year
        mngr.wdid,                                    // F: mngr_wdid
        mngr.name,                                    // G: mngr_name
        mngrHierarchy.immediate || '',                // H: mngr_immediate_manager
        mngrHierarchy.sr || '',                       // I: mngr_sr_manager
        mngrHierarchy.deptHead || '',                 // J: mngr_department_head
        mngr.program || '',                           // K: program
        mngr.country || '',                           // L: country
        mngr.region || '',                            // M: region
        mngr.innovation.innovation_high_count || 0,   // N: innovation_high_count
        mngr.innovation.innovation_med_count || 0,    // O: innovation_med_count
        mngr.innovation.innovation_low_count || 0,    // P: innovation_low_count
        mngr.innovation.exempt || 'No',               // Q: exempt
        mngr.innovation.exception_reason || '',       // R: exception_reason
        mngr.innovation.notes || ''                   // S: notes
      ];

      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 2, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 18).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} Manager Innovation rows starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow
    };

  } catch (error) {
    Logger.log('Error in submitManagerInnovation: ' + error.toString());
    throw error;
  }
}

/**
 * Submit Manager Ratio_Alignment KPIs
 * Writes to Manager KPI workbook → Ratio_Alignment tab
 * Range: B:Q (16 columns)
 */
function submitManagerCostToRevenue(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(MANAGER_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(MNGR_COST_TO_REVENUE_TAB);
    if (!sheet) throw new Error('Manager Cost_to_Revenue sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.submitterEmail || '';

    const rowsToAppend = [];

    payload.managerList.forEach(mngr => {
      const managerDetails = rosterIndex.get(String(mngr.wdid).trim());  // ✅ Use cache
      const mngrHierarchy = mapHierarchy_(managerDetails, "MNGR_SHEETS");  // ✅ Manager context

      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        getMonthAbbreviation_(payload.month),         // D: report_month
        payload.year,                                 // E: report_year
        mngr.wdid,                                    // F: mngr_wdid
        mngr.name,                                    // G: mngr_name
        mngrHierarchy.immediate || '',                // H: mngr_immediate_manager
        mngrHierarchy.sr || '',                       // I: mngr_sr_manager
        mngrHierarchy.deptHead || '',                 // J: mngr_department_head
        mngr.program || '',                           // K: program
        mngr.country || '',                           // L: country
        mngr.region || '',                            // M: region
        mngr.cost_to_revenue.actual_revenue || 0,     // N: actual_revenue
        mngr.cost_to_revenue.projected_yearly_revenue || 0,    // O: projected_yearly_revenue
        mngr.cost_to_revenue.notes || ''              // P: notes
      ];

      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 2, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 15).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} Manager Cost_to_Revenue rows starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow
    };

  } catch (error) {
    Logger.log('Error in submitManagerCostToRevenue: ' + error.toString());
    throw error;
  }
}

/**
 * Submit Manager PPD KPIs
 * Writes to Manager KPI workbook → PPD tab
 * Range: B:Q (16 columns)
 */
function submitManagerPPD(payload, rosterIndex) {
  try {
    const ss = SpreadsheetApp.openById(MANAGER_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(MNGR_PPD_TAB);
    if (!sheet) throw new Error('Manager PPD sheet not found');

    const timestamp = new Date();
    const submitterEmail = payload.submitterEmail || '';

    const rowsToAppend = [];

    payload.managerList.forEach(mngr => {
      const managerDetails = rosterIndex.get(String(mngr.wdid).trim());  // ✅ Use cache
      const mngrHierarchy = mapHierarchy_(managerDetails, "MNGR_SHEETS");  // ✅ Manager context

      const row = [
        timestamp,                                    // B: timestamp
        submitterEmail,                               // C: submitter_email
        getMonthAbbreviation_(payload.month),         // D: report_month
        payload.year,                                 // E: report_year
        mngr.wdid,                                    // F: mngr_wdid
        mngr.name,                                    // G: mngr_name
        mngrHierarchy.immediate || '',                // H: mngr_immediate_manager
        mngrHierarchy.sr || '',                       // I: mngr_sr_manager
        mngrHierarchy.deptHead || '',                 // J: mngr_department_head
        mngr.program || '',                           // K: program
        mngr.country || '',                           // L: country
        mngr.region || '',                            // M: region
        mngr.ppd.ppd_completed_courses || 0,          // N: ppd_completed_courses
        mngr.ppd.ppd_required_courses || 0,           // O: ppd_required_courses
        mngr.ppd.exempt || 'No',                      // P: exempt
        mngr.ppd.exception_reason || '',              // Q: exception_reason
        mngr.ppd.notes || ''                          // R: notes
      ];

      rowsToAppend.push(row);
    });

    const startRow = getNextAppendRow_(sheet, 2, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 17).setValues(rowsToAppend);
    }

    Logger.log(`Submitted ${rowsToAppend.length} Manager PPD rows starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow
    };

  } catch (error) {
    Logger.log('Error in submitManagerPPD: ' + error.toString());
    throw error;
  }
}

/**
 * Master submit function for Manager-level KPIs
 */
function submitKPIs_MNGR(payload) {
  try {
    Logger.log('Starting Manager KPI submission for Director WDID: ' + payload.directorWdid);

    // ✅ Load roster index once
    const rosterIndex = getRosterIndex_();

    // Get director details
    const directorDetails = rosterIndex.get(String(payload.directorWdid).trim());
    if (!directorDetails) {
      throw new Error(`Director WDID ${payload.directorWdid} not found in roster`);
    }

    payload.submitterEmail = directorDetails.email;

    Logger.log(`Submitting Manager KPIs for ${payload.managerList.length} Manager(s)`);

    // ✅ Pass rosterIndex to all three functions
    const innovationResult = submitManagerInnovation(payload, rosterIndex);
    const ppdResult = submitManagerPPD(payload, rosterIndex);
    const costToRevenueResult = submitManagerCostToRevenue(payload, rosterIndex);

    Logger.log('Manager KPI submission completed successfully');

    return {
      success: true,
      message: `Successfully submitted Manager KPIs for ${payload.managerList.length} QA Manager(s)`,
      details: {
        innovation: innovationResult,
        ppd: ppdResult,
        costToRevenue: costToRevenueResult
      }
    };

  } catch (error) {
    Logger.log('Error in submitKPIs_MNGR: ' + error.toString());
    console.error('Error in submitKPIs_MNGR:', error);
    return {
      success: false,
      message: error.message || 'An error occurred during Manager KPI submission'
    };
  }
}

function testAndreaLookup() {
  const directorWdid = '10017245'; // Andrea's WDID

  Logger.log('=== TESTING ANDREA LOOKUP ===');
  Logger.log(`Looking for WDID: ${directorWdid}`);

  const ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  const sheet = ss.getSheetByName(ROSTER_TAB_NAME);
  const values = sheet.getDataRange().getValues();

  Logger.log(`Total rows in roster: ${values.length}`);

  let foundCount = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const supervisor = String(row[3]).trim(); // Column D

    if (supervisor.includes(`(${directorWdid})`)) {
      foundCount++;
      Logger.log(`\n--- Found Match #${foundCount} ---`);
      Logger.log(`  WDID: ${row[1]}`);
      Logger.log(`  Name: ${row[2]}`);
      Logger.log(`  Supervisor: ${row[3]}`);
      Logger.log(`  Manager: ${row[4]}`);
      Logger.log(`  Sr Manager: ${row[5]}`);
      Logger.log(`  Job Profile: ${row[6]}`);
    }
  }

  Logger.log(`\n=== SUMMARY ===`);
  Logger.log(`Total people reporting to Andrea: ${foundCount}`);

  // Now test the actual function
  Logger.log(`\n=== TESTING getManagersByDirectorWdid() ===`);
  const managers = getManagersByDirectorWdid(directorWdid);
  Logger.log(`Managers returned: ${managers.length}`);
  managers.forEach(m => {
    Logger.log(`  - ${m.name} (${m.wdid}) - ${m.role}`);
  });
}

// ============================================================================
// PROGRAM-LEVEL RATIO ALIGNMENT FUNCTIONS
// ============================================================================

/**
 * Get list of programs from WD FAs sheet
 * Returns sorted, unique, trimmed array of program names
 */
function getProgramsList() {
  try {
    const ss = SpreadsheetApp.openById(MANAGER_KPI_SHEET_ID);
    const sheet = ss.getSheetByName('WD FAs');
    if (!sheet) throw new Error('WD FAs sheet not found');

    const values = sheet.getRange('A2:A').getValues();
    const programs = [];

    values.forEach(row => {
      const program = String(row[0]).trim();
      if (program && program !== '' && !programs.includes(program)) {
        programs.push(program);
      }
    });

    // Sort alphabetically
    programs.sort((a, b) => a.localeCompare(b));

    Logger.log(`Found ${programs.length} unique programs`);
    return programs;

  } catch (error) {
    Logger.log('Error in getProgramsList: ' + error.toString());
    throw error;
  }
}


function submitCostToRevenueProgramLevel(payload) {
  try {
    // Validate payload
    if (!payload.directorWdid) {
      throw new Error('Director WDID is required');
    }
    if (!payload.managerWdid) {
      throw new Error('Manager WDID is required');
    }
    if (!payload.programs || payload.programs.length === 0) {
      throw new Error('At least one program must be selected');
    }

    // ✅ Load roster index once
    const rosterIndex = getRosterIndex_();

    // Get director and manager details
    const directorDetails = rosterIndex.get(String(payload.directorWdid).trim());
    const managerDetails = rosterIndex.get(String(payload.managerWdid).trim());

    if (!directorDetails) {
      throw new Error(`Director WDID ${payload.directorWdid} not found in roster`);
    }
    if (!managerDetails) {
      throw new Error(`Manager WDID ${payload.managerWdid} not found in roster`);
    }

    const submitterEmail = directorDetails.email || '';
    const timestamp = new Date();

    // ✅ Use mapHierarchy_ for manager hierarchy
    const mngrHierarchy = mapHierarchy_(managerDetails, "MNGR_SHEETS");

    Logger.log(`Manager: ${managerDetails.name} (${payload.managerWdid})`);
    Logger.log(`  Immediate Manager: ${mngrHierarchy.immediate}`);
    Logger.log(`  Sr Manager: ${mngrHierarchy.sr}`);
    Logger.log(`  Department Head: ${mngrHierarchy.deptHead}`);

    // Validate each program entry
    payload.programs.forEach((prog, index) => {
      if (!prog.program || prog.program.trim() === '') {
        throw new Error(`Program ${index + 1}: Program name is required`);
      }

      const exempt = prog.exempt || 'No';

      if (exempt === 'Yes') {
        if (!prog.exception_reason || prog.exception_reason.trim() === '') {
          throw new Error(`Program "${prog.program}": Exception reason is required when exempt = Yes`);
        }
      } else {
        // Validate numeric fields
        const projectedRevenue = parseFloat(prog.projected_yearly_revenue);
        const actualRevenue = parseFloat(prog.actual_revenue);

        if (isNaN(projectedRevenue) || projectedRevenue < 0) {
          throw new Error(`Program "${prog.program}": Year end actual Operations revenue must be a valid number (0 or greater)`);
        }
        if (isNaN(actualRevenue) || actualRevenue < 0) {
          throw new Error(`Program "${prog.program}": Quality year end actual S&B must be a valid number (0 or greater)`);
        }
      }
    });

    // Open sheet
    const ss = SpreadsheetApp.openById(MANAGER_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(MNGR_COST_TO_REVENUE_TAB);
    if (!sheet) throw new Error('Cost_to_Revenue sheet not found');

    const rowsToAppend = [];

    // Build one row per program
    payload.programs.forEach(prog => {
      const row = [
        timestamp,                                      // B: timestamp
        submitterEmail,                                 // C: submitter_email
        getMonthAbbreviation_(payload.month),           // D: report_month
        payload.year,                                   // E: report_year
        payload.managerWdid,                            // F: mngr_wdid
        payload.managerName || managerDetails.name,     // G: mngr_name
        mngrHierarchy.immediate || '',                  // H: mngr_immediate_manager
        mngrHierarchy.sr || '',                         // I: mngr_sr_manager
        mngrHierarchy.deptHead || '',                   // J: mngr_department_head
        prog.program,                                   // K: program
        managerDetails.country || '',                   // L: country
        managerDetails.region || '',                    // M: region
        parseFloat(prog.actual_revenue) || 0,            // N: actual_revenue
        parseFloat(prog.projected_yearly_revenue) || 0,  // O: projected_yearly_revenue
        prog.notes || ''                                // P: notes 
      ];

      rowsToAppend.push(row);
    });

    // Append all rows
    const startRow = getNextAppendRow_(sheet, 2, 2);

    if (rowsToAppend.length > 0) {
      sheet.getRange(startRow, 2, rowsToAppend.length, 15).setValues(rowsToAppend);

      // ✅ Format currency columns (N & O)
      formatCostToRevenueCurrency_(startRow, rowsToAppend.length);
    }

    Logger.log(`Submitted ${rowsToAppend.length} program-level Cost_to_Revenue rows starting at row ${startRow}`);

    return {
      success: true,
      rowsWritten: rowsToAppend.length,
      startRow: startRow,
      message: `Successfully submitted ${rowsToAppend.length} program(s) for ${payload.managerName}`
    };

  } catch (error) {
    Logger.log('Error in submitCostToRevenueProgramLevel: ' + error.toString());
    throw error;
  }
}

/**
 * Format currency columns in Cost_to_Revenue sheet
 * Formats columns N (projected_yearly_revenue) and O (actual_revenue) as currency
 * 
 * @param {number} startRow - Starting row number
 * @param {number} numRows - Number of rows to format
 */
function formatCostToRevenueCurrency_(startRow, numRows) {
  try {
    const ss = SpreadsheetApp.openById(MANAGER_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(MNGR_COST_TO_REVENUE_TAB);
    if (!sheet) return;

    // Format column N (projected_yearly_revenue) - column 14 from column A, or column 13 from column B
    // Since data starts at column B, column N is position 13 (B=1, C=2... N=13)
    const projectedRange = sheet.getRange(startRow, 14, numRows, 1); // Column N
    projectedRange.setNumberFormat('$#,##0.00');

    // Format column O (actual_revenue) - position 14 from column B
    const actualRange = sheet.getRange(startRow, 15, numRows, 1); // Column O
    actualRange.setNumberFormat('$#,##0.00');

    Logger.log(`Formatted ${numRows} rows with currency format (columns N & O)`);

  } catch (error) {
    Logger.log('Error formatting currency: ' + error.toString());
    // Don't throw - formatting is non-critical
  }
}

// ============================================================================
// SERVICE DELIVERY READ-ONLY RETRIEVAL (FOR MANAGER VIEW)
// ============================================================================

/**
 * Get Service Delivery rows for a Manager's TLs
 * Used in Manager UI to display read-only Service Delivery results
 * 
 * @param {string} managerWdid - Manager's WDID
 * @param {string} reportMonth - Report month (e.g., "December")
 * @param {string} reportYear - Report year (e.g., "2025")
 * @returns {Array} Array of Service Delivery objects for matching TLs
 */
function getServiceDeliveryRowsForManager(managerWdid, reportMonth, reportYear) {
  try {
    Logger.log(`Loading Service Delivery for Manager ${managerWdid}, ${reportMonth} ${reportYear}`);

    if (!managerWdid || !reportMonth || !reportYear) {
      return [];
    }

    // Get manager's TLs first
    const tls = getTLsByManagerWdid(managerWdid);
    if (tls.length === 0) {
      Logger.log('No TLs found for this manager');
      return [];
    }

    const tlWdids = tls.map(tl => String(tl.wdid).trim());
    Logger.log(`Looking for Service Delivery data for ${tlWdids.length} TLs: ${tlWdids.join(', ')}`);

    // Open Service Delivery sheet
    const ss = SpreadsheetApp.openById(SERVICE_DELIVERY_SHEET_ID);
    const sheet = ss.getSheetByName(SERVICE_DELIVERY_TAB);
    if (!sheet) {
      Logger.log('Service Delivery sheet not found');
      return [];
    }

    // Get all data (columns B:Q)
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data in Service Delivery sheet');
      return [];
    }

    const values = sheet.getRange(2, 2, lastRow - 1, 16).getValues(); // Start row 2, column B, 16 columns (B:Q)

    const results = [];
    const monthAbbr = getMonthAbbreviation_(reportMonth).toLowerCase(); // ✅ Convert to 3-letter abbreviation
    const yearStr = String(reportYear).trim(); // ✅ Ensure year is string

    Logger.log(`Filtering for month: "${monthAbbr}", year: "${yearStr}"`);

    for (let i = 0; i < values.length; i++) {
      const row = values[i];

      // ✅ Robust type conversion for all comparison fields
      const rowYear = String(row[0]).trim();        // Column B (index 0) - report_year
      const rowMonth = String(row[1]).toLowerCase().trim(); // Column C (index 1) - report_month
      const tlWdid = String(row[2]).trim();         // Column D (index 2) - tl_wdid

      // Skip empty rows
      if (!rowYear || !rowMonth || !tlWdid) continue;

      // ✅ Match year (as string), month (as abbreviation), and TL WDID
      if (rowYear === yearStr && rowMonth === monthAbbr && tlWdids.includes(tlWdid)) {
        results.push({
          tl_wdid: tlWdid,                                    // D
          tl_name: String(row[3]).trim(),                     // E
          tl_manager: String(row[4]).trim(),                  // F
          tl_sr_manager: String(row[5]).trim(),               // G
          program: String(row[6]).trim(),                     // H
          country: String(row[7]).trim(),                     // I
          region: String(row[8]).trim(),                      // J
          quality_ptg_avg: row[9],                            // K
          voc_ptg_avg: row[10],                               // L
          team_scores_ptg: row[11],                           // M
          reporting_ptg: row[12],                             // N
          service_delivery_pct_to_goal: row[13],              // O
          service_delivery_weight_earned: row[14],            // P
          service_delivery_status: String(row[15]).trim()     // Q
        });
      }
    }

    Logger.log(`Found ${results.length} Service Delivery records`);

    // Sort by program, then TL name
    results.sort((a, b) => {
      if (a.program !== b.program) {
        return a.program.localeCompare(b.program);
      }
      return a.tl_name.localeCompare(b.tl_name);
    });

    return results;

  } catch (error) {
    Logger.log('Error in getServiceDeliveryRowsForManager: ' + error.toString());
    console.error('Error in getServiceDeliveryRowsForManager:', error);
    return [];
  }
}

// ============================================================================
// MANAGER ROLL-UP READ-ONLY RETRIEVAL (FOR MANAGER VIEW)
// ============================================================================

/**
 * Get Manager's overall roll-up data
 * Returns single object from MNGR_roll_up tab
 * 
 * @param {string} managerWdid - Manager's WDID
 * @param {string} reportMonth - Report month (e.g., "December")
 * @param {string} reportYear - Report year (e.g., "2025")
 * @returns {Object|null} Manager roll-up object or null if not found
 */
function getManagerRollup(managerWdid, reportMonth, reportYear) {
  try {
    Logger.log(`Loading Manager Roll-up for Manager ${managerWdid}, ${reportMonth} ${reportYear}`);

    if (!managerWdid || !reportMonth || !reportYear) {
      return null;
    }

    const ss = SpreadsheetApp.openById(MANAGER_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(MNGR_ROLLUP_TAB);

    if (!sheet) {
      Logger.log('MNGR_roll_up sheet not found');
      return null;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data in MNGR_roll_up sheet');
      return null;
    }

    const values = sheet.getRange(2, 1, lastRow - 1, 17).getValues(); // A:Q (17 columns)

    const monthAbbr = getMonthAbbreviation_(reportMonth).toLowerCase();
    const yearStr = String(reportYear).trim();

    Logger.log(`Filtering for month: "${monthAbbr}", year: "${yearStr}", manager: "${managerWdid}"`);

    for (let i = 0; i < values.length; i++) {
      const row = values[i];

      const rowYear = String(row[0]).trim();        // A: report_year
      const rowMonth = String(row[1]).toLowerCase().trim(); // B: report_month
      const mngrWdid = String(row[2]).trim();       // C: mngr_wdid

      if (!rowYear || !rowMonth || !mngrWdid) continue;

      // ✅ FLEXIBLE MONTH MATCHING: Accept both "dec" and "december"
      const monthMatches = rowMonth === monthAbbr || rowMonth === reportMonth.toLowerCase().trim();

      // Match year, month, and manager WDID
      if (rowYear === yearStr && monthMatches && mngrWdid === String(managerWdid).trim()) {
        Logger.log('✓ Found matching roll-up row');

        return {
          mngr_wdid: mngrWdid,                      // C
          mngr_name: String(row[3]).trim(),         // D
          country: String(row[4]).trim(),           // E
          region: String(row[5]).trim(),            // F
          tl_count: row[6],                         // G
          qa_count: row[7],                         // H
          quality_ptg_avg: row[8],                  // I
          voc_ptg_avg: row[9],                      // J
          accuracy_ptg_avg: row[10],                // K
          quality_weight_earned: row[11],           // L
          voc_weight_earned: row[12],               // M
          accuracy_weight_earned: row[13],          // N
          quality_status: String(row[14]).trim(),   // O
          voc_status: String(row[15]).trim(),       // P
          accuracy_status: String(row[16]).trim()   // Q
        };
      }
    }

    Logger.log('No matching roll-up data found');
    return null;

  } catch (error) {
    Logger.log('Error in getManagerRollup: ' + error.toString());
    console.error('Error in getManagerRollup:', error);
    return null;
  }
}

/**
 * Get Manager's roll-up data by program
 * Returns array of objects from MNGR_roll_up_by_program tab
 * ✅ UPDATED: Now aggregates TL-level rows by program
 * 
 * @param {string} managerWdid - Manager's WDID
 * @param {string} reportMonth - Report month (e.g., "December")
 * @param {string} reportYear - Report year (e.g., "2025")
 * @returns {Array} Array of program roll-up objects
 */
function getManagerRollupByProgram(managerWdid, reportMonth, reportYear) {
  try {
    Logger.log(`Loading Manager Roll-up by Program for Manager ${managerWdid}, ${reportMonth} ${reportYear}`);

    if (!managerWdid || !reportMonth || !reportYear) {
      return [];
    }

    const ss = SpreadsheetApp.openById(MANAGER_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(MNGR_ROLLUP_BY_PROGRAM_TAB);

    if (!sheet) {
      Logger.log('MNGR_roll_up_by_program sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data in MNGR_roll_up_by_program sheet');
      return [];
    }

    // ✅ Read all 20 columns (A:T) to match new layout
    const values = sheet.getRange(2, 1, lastRow - 1, 20).getValues();

    const monthAbbr = getMonthAbbreviation_(reportMonth).toLowerCase();
    const yearStr = String(reportYear).trim();

    Logger.log(`Filtering for month: "${monthAbbr}", year: "${yearStr}", manager: "${managerWdid}"`);

    // ✅ STEP 1: Collect all matching TL-level rows
    const matchingRows = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];

      const rowYear = String(row[0]).trim();        // A: report_year
      const rowMonth = String(row[1]).toLowerCase().trim(); // B: report_month
      const mngrWdid = String(row[2]).trim();       // C: mngr_wdid

      if (!rowYear || !rowMonth || !mngrWdid) continue;

      // Flexible month matching
      const monthMatches = rowMonth === monthAbbr || rowMonth === reportMonth.toLowerCase().trim();

      if (rowYear === yearStr && monthMatches && mngrWdid === String(managerWdid).trim()) {
        matchingRows.push({
          mngr_wdid: mngrWdid,                      // C
          mngr_name: String(row[3]).trim(),         // D
          tl_wdid: String(row[4]).trim(),           // E
          tl_name: String(row[5]).trim(),           // F
          program: String(row[6]).trim(),           // G 
          country: String(row[7]).trim(),           // H 
          region: String(row[8]).trim(),            // I 
          tl_count: row[9],                         // J 
          qa_count: row[10],                        // K 
          quality_ptg_avg: row[11],                 // L 
          voc_ptg_avg: row[12],                     // M 
          accuracy_ptg_avg: row[13],                // N 
          quality_weight_earned: row[14],           // O 
          voc_weight_earned: row[15],               // P 
          accuracy_weight_earned: row[16],          // Q 
          quality_status: String(row[17]).trim(),   // R 
          voc_status: String(row[18]).trim(),       // S 
          accuracy_status: String(row[19]).trim()   // T 
        });
      }
    }

    Logger.log(`Found ${matchingRows.length} TL-level rows to aggregate`);

    if (matchingRows.length === 0) {
      return [];
    }

    // ✅ STEP 2: Aggregate by program
    const programMap = {};

    matchingRows.forEach(row => {
      const prog = row.program;

      if (!programMap[prog]) {
        programMap[prog] = {
          program: prog,
          country: row.country,
          region: row.region,
          tl_wdids: [],
          tl_names: [],
          qa_counts: [],
          quality_ptg_avgs: [],
          voc_ptg_avgs: [],
          accuracy_ptg_avgs: [],
          quality_weight_earneds: [],
          voc_weight_earneds: [],
          accuracy_weight_earneds: [],
          quality_statuses: [],
          voc_statuses: [],
          accuracy_statuses: []
        };
      }

      // Collect values for aggregation
      programMap[prog].tl_wdids.push(row.tl_wdid);
      programMap[prog].tl_names.push(row.tl_name);

      if (row.qa_count !== '' && row.qa_count !== null && !isNaN(row.qa_count)) {
        programMap[prog].qa_counts.push(parseFloat(row.qa_count));
      }

      if (row.quality_ptg_avg !== '' && row.quality_ptg_avg !== null && !isNaN(row.quality_ptg_avg)) {
        programMap[prog].quality_ptg_avgs.push(parseFloat(row.quality_ptg_avg));
      }

      if (row.voc_ptg_avg !== '' && row.voc_ptg_avg !== null && !isNaN(row.voc_ptg_avg)) {
        programMap[prog].voc_ptg_avgs.push(parseFloat(row.voc_ptg_avg));
      }

      if (row.accuracy_ptg_avg !== '' && row.accuracy_ptg_avg !== null && !isNaN(row.accuracy_ptg_avg)) {
        programMap[prog].accuracy_ptg_avgs.push(parseFloat(row.accuracy_ptg_avg));
      }

      if (row.quality_weight_earned !== '' && row.quality_weight_earned !== null && !isNaN(row.quality_weight_earned)) {
        programMap[prog].quality_weight_earneds.push(parseFloat(row.quality_weight_earned));
      }

      if (row.voc_weight_earned !== '' && row.voc_weight_earned !== null && !isNaN(row.voc_weight_earned)) {
        programMap[prog].voc_weight_earneds.push(parseFloat(row.voc_weight_earned));
      }

      if (row.accuracy_weight_earned !== '' && row.accuracy_weight_earned !== null && !isNaN(row.accuracy_weight_earned)) {
        programMap[prog].accuracy_weight_earneds.push(parseFloat(row.accuracy_weight_earned));
      }

      // Collect statuses (we'll pick the most common or worst)
      if (row.quality_status) programMap[prog].quality_statuses.push(row.quality_status);
      if (row.voc_status) programMap[prog].voc_statuses.push(row.voc_status);
      if (row.accuracy_status) programMap[prog].accuracy_statuses.push(row.accuracy_status);
    });

    // ✅ STEP 3: Calculate aggregated values per program
    const results = [];

    Object.keys(programMap).forEach(progName => {
      const prog = programMap[progName];

      // Helper: Calculate average
      const avg = (arr) => {
        if (!arr || arr.length === 0) return null;
        const sum = arr.reduce((a, b) => a + b, 0);
        return sum / arr.length;
      };

      // Helper: Pick most severe status (IR > Developing > Achieving > Excelling)
      const pickWorstStatus = (statuses) => {
        if (!statuses || statuses.length === 0) return '';
        if (statuses.some(s => s.toLowerCase().includes('ir') || s.toLowerCase().includes('improvement'))) return 'IR';
        if (statuses.some(s => s.toLowerCase().includes('developing'))) return 'Developing';
        if (statuses.some(s => s.toLowerCase().includes('achieving'))) return 'Achieving';
        if (statuses.some(s => s.toLowerCase().includes('excelling'))) return 'Excelling';
        return statuses[0]; // Fallback to first status
      };

      // ✅ Get unique TL names (in case same TL appears multiple times)
      const uniqueTLNames = [...new Set(prog.tl_names)];
      const tlNamesDisplay = uniqueTLNames.join(', ');

      results.push({
        mngr_wdid: managerWdid,
        mngr_name: matchingRows[0].mngr_name, // Same for all rows
        program: progName,
        country: prog.country,
        region: prog.region,
        tl_count: prog.tl_wdids.length,  // Count unique TLs
        tl_names: tlNamesDisplay,
        qa_count: prog.qa_counts.reduce((a, b) => a + b, 0), // Sum QA counts
        quality_ptg_avg: avg(prog.quality_ptg_avgs),
        voc_ptg_avg: avg(prog.voc_ptg_avgs),
        accuracy_ptg_avg: avg(prog.accuracy_ptg_avgs),
        quality_weight_earned: avg(prog.quality_weight_earneds),
        voc_weight_earned: avg(prog.voc_weight_earneds),
        accuracy_weight_earned: avg(prog.accuracy_weight_earneds),
        quality_status: pickWorstStatus(prog.quality_statuses),
        voc_status: pickWorstStatus(prog.voc_statuses),
        accuracy_status: pickWorstStatus(prog.accuracy_statuses)
      });
    });

    Logger.log(`Aggregated into ${results.length} program roll-up records`);

    // Sort by program name (A-Z)
    results.sort((a, b) => a.program.localeCompare(b.program));

    return results;

  } catch (error) {
    Logger.log('Error in getManagerRollupByProgram: ' + error.toString());
    console.error('Error in getManagerRollupByProgram:', error);
    return [];
  }
}

/**
 * Get Manager's roll-up data by TL
 * Returns array of TL-level objects from MNGR_roll_up_by_program tab
 * Each TL-program combination is returned as a separate row (no deduplication)
 * 
 * @param {string} managerWdid - Manager's WDID
 * @param {string} reportMonth - Report month (e.g., "December")
 * @param {string} reportYear - Report year (e.g., "2025")
 * @returns {Array} Array of TL roll-up objects
 */
function getManagerRollupByTL(managerWdid, reportMonth, reportYear) {
  try {
    Logger.log(`Loading Manager Roll-up by TL for Manager ${managerWdid}, ${reportMonth} ${reportYear}`);

    if (!managerWdid || !reportMonth || !reportYear) {
      return [];
    }

    const ss = SpreadsheetApp.openById(MANAGER_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(MNGR_ROLLUP_BY_PROGRAM_TAB);

    if (!sheet) {
      Logger.log('MNGR_roll_up_by_program sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data in MNGR_roll_up_by_program sheet');
      return [];
    }

    // Read columns A:T (20 columns)
    const values = sheet.getRange(2, 1, lastRow - 1, 20).getValues();

    const results = [];
    const monthAbbr = getMonthAbbreviation_(reportMonth).toLowerCase();
    const yearStr = String(reportYear).trim();

    Logger.log(`Filtering for month: "${monthAbbr}", year: "${yearStr}", manager: "${managerWdid}"`);

    for (let i = 0; i < values.length; i++) {
      const row = values[i];

      const rowYear = String(row[0]).trim();        // A: report_year
      const rowMonth = String(row[1]).toLowerCase().trim(); // B: report_month
      const mngrWdid = String(row[2]).trim();       // C: mngr_wdid

      if (!rowYear || !rowMonth || !mngrWdid) continue;

      // Flexible month matching
      const monthMatches = rowMonth === monthAbbr || rowMonth === reportMonth.toLowerCase().trim();

      if (rowYear === yearStr && monthMatches && mngrWdid === String(managerWdid).trim()) {
        // ✅ Return each row as-is (no deduplication)
        results.push({
          tl_wdid: String(row[4]).trim(),           // E: tl_wdid
          tl_name: String(row[5]).trim(),           // F: tl_name
          program: String(row[6]).trim(),           // G: program
          country: String(row[7]).trim(),           // H: country
          region: String(row[8]).trim(),            // I: region
          tl_count: row[9],                         // J: tl_count
          qa_count: row[10],                        // K: qa_count
          quality_ptg_avg: row[11],                 // L: quality_ptg_avg
          voc_ptg_avg: row[12],                     // M: voc_ptg_avg
          accuracy_ptg_avg: row[13],                // N: accuracy_ptg_avg
          quality_weight_earned: row[14],           // O: quality_weight_earned
          voc_weight_earned: row[15],               // P: voc_weight_earned
          accuracy_weight_earned: row[16],          // Q: accuracy_weight_earned
          quality_status: String(row[17]).trim(),   // R: quality_status
          voc_status: String(row[18]).trim(),       // S: voc_status
          accuracy_status: String(row[19]).trim()   // T: accuracy_status
        });
      }
    }

    Logger.log(`Found ${results.length} TL roll-up records`);

    // Sort by TL name, then program
    results.sort((a, b) => {
      if (a.tl_name !== b.tl_name) {
        return a.tl_name.localeCompare(b.tl_name);
      }
      return a.program.localeCompare(b.program);
    });

    return results;

  } catch (error) {
    Logger.log('Error in getManagerRollupByTL: ' + error.toString());
    console.error('Error in getManagerRollupByTL:', error);
    return [];
  }
}

// ============================================================================
// REGIONAL LEADS KPI FUNCTIONS
// ============================================================================

// Regional Leads Configuration
const REGIONAL_LEADS_SHEET_ID = '1ZERxb4bV6-KyzJiGIgdy4AHQQmBFqyx1Nc38SXzLxNU';
const REGIONAL_LEADS_TAB = 'Regional_Leads_KPIs';

// Allowed Regional Lead WDIDs (access control list)
const REGIONAL_LEADS_ALLOWED = ['10053071', '10017324'];

/**
 * Validate if WDID is allowed to submit Regional Lead KPIs
 * AND retrieve their details from the Global QA Roster
 * @param {string} wdid - Workday ID to validate
 * @return {Object} - { allowed: boolean, details: object }
 */
function validateWdidAllowed(wdid) {
  try {
    const normalizedWdid = String(wdid).trim();

    // Check if WDID is in allowed list
    if (!REGIONAL_LEADS_ALLOWED.includes(normalizedWdid)) {
      return {
        allowed: false,
        details: null
      };
    }

    // Get details from roster
    const rosterDetails = getRosterDetailsByWDID(normalizedWdid);

    if (!rosterDetails) {
      Logger.log(`Regional Lead WDID ${normalizedWdid} is in allowed list but not found in roster`);
      return {
        allowed: false,
        details: null,
        message: 'WDID not found in Global QA Roster'
      };
    }

    // Return roster details
    return {
      allowed: true,
      details: {
        name: rosterDetails.name,
        email: rosterDetails.email,
        region: rosterDetails.region,
        role: rosterDetails.role,
        program: rosterDetails.program,
        country: rosterDetails.country,
        manager: rosterDetails.manager,
        sr_manager: rosterDetails.sr_manager
      }
    };

  } catch (error) {
    Logger.log('Error in validateWdidAllowed: ' + error.toString());
    return {
      allowed: false,
      details: null,
      error: error.message
    };
  }
}

// ============================================================================
// DIRECTOR SELF-REPORTING KPI CONFIGURATION
// ============================================================================

// Director KPI Sheet Configuration
const DIRECTOR_KPI_SHEET_ID = '1CtzvbhvgExcC4rex-bQSPUcwq5dPwt8e_3cOgBejhuA';
const DIRECTOR_KPI_TAB = 'QA_Director_KPIs';

// Allowed Director WDID (access control)
const DIRECTOR_ALLOWED_WDIDS = ['10089598', 'test'];

// Hardcoded reporting period (update monthly)
const DIRECTOR_REPORT_MONTH = 'January';
const DIRECTOR_REPORT_YEAR = '2026';

// ============================================================================
// DIRECTOR SELF-REPORTING KPI FUNCTIONS
// ============================================================================

/**
 * Validate if WDID is allowed to submit Director KPIs
 * AND retrieve their details from the Global QA Roster
 * @param {string} wdid - Workday ID to validate
 * @return {Object} - { allowed: boolean, details: object }
 */
function validateDirectorWdid(wdid) {
  try {
    const normalizedWdid = String(wdid).trim();

    // Check if WDID is in the allowed list
    if (!DIRECTOR_ALLOWED_WDIDS.includes(normalizedWdid)) {
      return {
        allowed: false,
        details: null
      };
    }

    // Handle "test" WDID with mock data
    if (normalizedWdid === 'test') {
      return {
        allowed: true,
        details: {
          name: 'Test Director',
          email: 'test.director@telus.com',
          region: 'Americas',
          role: 'CX & T&S Quality Director',
          program: 'Test Program',
          country: 'Test Country'
        }
      };
    }

    // Get details from roster for real WDIDs
    const rosterDetails = getRosterDetailsByWDID(normalizedWdid);

    if (!rosterDetails) {
      Logger.log(`Director WDID ${normalizedWdid} is allowed but not found in roster`);
      return {
        allowed: false,
        details: null,
        message: 'WDID not found in Global QA Roster'
      };
    }

    // Return roster details
    return {
      allowed: true,
      details: {
        name: rosterDetails.name,
        email: rosterDetails.email,
        region: rosterDetails.region,
        role: rosterDetails.role || 'CX & T&S Quality Director',
        program: rosterDetails.program,
        country: rosterDetails.country
      }
    };

  } catch (error) {
    Logger.log('Error in validateDirectorWdid: ' + error.toString());
    return {
      allowed: false,
      details: null,
      error: error.message
    };
  }
}
/**
 * Get Director details by WDID from Global QA Roster
 * (Wrapper around getRosterDetailsByWDID for consistency)
 * @param {string} wdid - Workday ID
 * @return {Object} - Director details or null
 */
function getDirectorDetailsByWdid_(wdid) {
  try {
    const normalizedWdid = String(wdid).trim();

    // Check if WDID is allowed
    if (normalizedWdid !== DIRECTOR_ALLOWED_WDID) {
      return null;
    }

    // Get from roster
    const rosterDetails = getRosterDetailsByWDID(normalizedWdid);

    if (!rosterDetails) {
      return null;
    }

    return {
      name: rosterDetails.name,
      email: rosterDetails.email,
      region: rosterDetails.region,
      role: rosterDetails.role || 'CX & T&S Quality Director'
    };

  } catch (error) {
    Logger.log('Error in getDirectorDetailsByWdid_: ' + error.toString());
    return null;
  }
}

/**
 * Submit Director Self-Reporting KPIs
 * Writes to QA_Director_KPIs sheet
 * Populates ONLY columns B:K, O:P, and T (same row)
 * Does NOT write to column A
 * 
 * @param {Object} payload - KPI data from form
 * @return {Object} - { success: boolean, message: string }
 */
function submitDirectorKpis(payload) {
  try {
    // Validate WDID and get roster details
    const validation = validateDirectorWdid(payload.director_wdid);

    if (!validation.allowed) {
      return {
        success: false,
        message: validation.message || 'Access denied. Your WDID is not authorized to submit Director KPIs.'
      };
    }

    // Enrich payload with roster data (ensure consistency)
    const enrichedPayload = {
      ...payload,
      director_name: validation.details.name,
      submitter_email: validation.details.email,
      region: validation.details.region,
      role: validation.details.role,
      report_month: DIRECTOR_REPORT_MONTH,  // ✅ Hardcoded
      report_year: DIRECTOR_REPORT_YEAR     // ✅ Hardcoded
    };

    // ✅ Convert month to 3-letter abbreviation
    const monthAbbr = getMonthAbbreviation_(enrichedPayload.report_month);

    Logger.log('Submitting Director KPIs:', {
      wdid: enrichedPayload.director_wdid,
      name: enrichedPayload.director_name,
      email: enrichedPayload.submitter_email,
      region: enrichedPayload.region,
      role: enrichedPayload.role,
      month: monthAbbr,
      year: enrichedPayload.report_year,
      sei_score: enrichedPayload.sei_score,
      sei_target: enrichedPayload.sei_target
    });

    // Generate timestamp
    const timestamp = new Date();

    // Open sheet
    const ss = SpreadsheetApp.openById(DIRECTOR_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(DIRECTOR_KPI_TAB);

    if (!sheet) {
      return {
        success: false,
        message: `Sheet "${DIRECTOR_KPI_TAB}" not found in spreadsheet.`
      };
    }

    // ✅ Find next empty row in column B (starting from row 2)
    const lastRowWithData = getNextAppendRow_(sheet, 2, 2);
    const targetRow = lastRowWithData;

    Logger.log(`Target row for submission: ${targetRow}`);

    // ✅ Convert SEI values to decimal (85 becomes 0.85 for percentage formatting)
    const seiScoreDecimal = enrichedPayload.sei_score / 100;
    const seiTargetDecimal = enrichedPayload.sei_target / 100;

    // ============================================
    // WRITE DATA TO SPECIFIC COLUMN RANGES
    // ============================================

    // ✅ RANGE 1: Columns B:K (10 columns) - Metadata + SEI Score/Target
    const metadataAndSEI = [
      timestamp,                             // B: timestamp
      enrichedPayload.submitter_email,       // C: submitter_email
      monthAbbr,                             // D: report_month (✅ ABBREVIATED)
      enrichedPayload.report_year,           // E: report_year
      enrichedPayload.region,                // F: region
      enrichedPayload.role,                  // G: role
      enrichedPayload.director_wdid,         // H: director_wdid
      enrichedPayload.director_name,         // I: director_name
      seiScoreDecimal,                       // J: sei_score (✅ AS DECIMAL 0.85)
      seiTargetDecimal                       // K: sei_target (✅ AS DECIMAL 0.85)
    ];

    sheet.getRange(targetRow, 2, 1, 10).setValues([metadataAndSEI]);

    // ✅ RANGE 2: Columns O:P (2 columns) - Revenue Actual/Projected
    const revenueData = [
      enrichedPayload.actual_revenue,        // O: actual_revenue
      enrichedPayload.projected_yearly_revenue // P: projected_yearly_revenue
    ];

    sheet.getRange(targetRow, 15, 1, 2).setValues([revenueData]);

    // ✅ RANGE 3: Column T (1 column) - Comments
    const commentsData = [
      enrichedPayload.comments || ''         // T: comments
    ];

    sheet.getRange(targetRow, 20, 1, 1).setValues([commentsData]);

    // ✅ Format currency columns (O and P)
    formatDirectorCurrency_(targetRow);

    // ✅ Format SEI columns as percentage (J and K)
    formatSEIPercentages_(DIRECTOR_KPI_SHEET_ID, DIRECTOR_KPI_TAB, targetRow);

    Logger.log('Director KPIs submitted successfully:', {
      wdid: enrichedPayload.director_wdid,
      name: enrichedPayload.director_name,
      month: monthAbbr,
      year: enrichedPayload.report_year,
      row: targetRow
    });

    return {
      success: true,
      message: `KPIs submitted successfully for ${monthAbbr} ${enrichedPayload.report_year}.`
    };

  } catch (error) {
    Logger.log('Error in submitDirectorKpis: ' + error.toString());
    return {
      success: false,
      message: 'Error submitting KPIs: ' + error.message
    };
  }
}

/**
 * Format currency columns in QA_Director_KPIs sheet
 * Formats columns O (actual_revenue) and P (projected_yearly_revenue) as currency
 * 
 * @param {number} rowNumber - Row number to format
 */
function formatDirectorCurrency_(rowNumber) {
  try {
    const ss = SpreadsheetApp.openById(DIRECTOR_KPI_SHEET_ID);
    const sheet = ss.getSheetByName(DIRECTOR_KPI_TAB);
    if (!sheet) return;

    // Format column O (actual_revenue) - column 15
    const actualRange = sheet.getRange(rowNumber, 15, 1, 1); // Column O
    actualRange.setNumberFormat('$#,##0.00');

    // Format column P (projected_yearly_revenue) - column 16
    const projectedRange = sheet.getRange(rowNumber, 16, 1, 1); // Column P
    projectedRange.setNumberFormat('$#,##0.00');

    Logger.log(`Formatted row ${rowNumber} with currency format (columns O & P)`);

  } catch (error) {
    Logger.log('Error formatting currency: ' + error.toString());
    // Don't throw - formatting is non-critical
  }
}

/**
 * Get lead details by WDID from Global QA Roster
 * (This is now just a wrapper around getRosterDetailsByWDID for consistency)
 * @param {string} wdid - Workday ID
 * @return {Object} - Lead details or null
 */
function getLeadDetailsByWdid(wdid) {
  try {
    const normalizedWdid = String(wdid).trim();

    // Check if WDID is allowed
    if (!REGIONAL_LEADS_ALLOWED.includes(normalizedWdid)) {
      return null;
    }

    // Get from roster
    const rosterDetails = getRosterDetailsByWDID(normalizedWdid);

    if (!rosterDetails) {
      return null;
    }

    return {
      name: rosterDetails.name,
      email: rosterDetails.email,
      region: rosterDetails.region,
      role: rosterDetails.role,
      program: rosterDetails.program,
      country: rosterDetails.country,
      manager: rosterDetails.manager,
      sr_manager: rosterDetails.sr_manager
    };

  } catch (error) {
    Logger.log('Error in getLeadDetailsByWdid: ' + error.toString());
    return null;
  }
}

/**
 * Submit Regional Lead KPIs
 * 
 * CORRECT Column mapping:
 * B:K = metadata + SEI score/target
 * O = sei_comments ✅
 * P:Q = revenue actual/projected
 * U = ctr_comments (revenue comments) ✅
 * V:W = PPD courses completed/required
 * AA = ppd_comments ✅
 * 
 * @param {Object} payload - KPI data from form
 * @return {Object} - { success: boolean, message: string }
 */
function submitRegionalLeadKpis(payload) {
  try {
    // Validate WDID and get roster details
    const validation = validateWdidAllowed(payload.lead_wdid);

    if (!validation.allowed) {
      return {
        success: false,
        message: validation.message || 'Access denied. Your WDID is not authorized to submit Regional Lead KPIs.'
      };
    }

    // Enrich payload with roster data (ensure consistency)
    const enrichedPayload = {
      ...payload,
      lead_name: validation.details.name,
      submitter_email: validation.details.email,
      region: validation.details.region,
      role: validation.details.role
    };

    // ✅ Convert month to 3-letter abbreviation
    const monthAbbr = getMonthAbbreviation_(enrichedPayload.report_month);

    Logger.log('Submitting Regional Lead KPIs:', {
      wdid: enrichedPayload.lead_wdid,
      name: enrichedPayload.lead_name,
      email: enrichedPayload.submitter_email,
      region: enrichedPayload.region,
      role: enrichedPayload.role,
      month: monthAbbr,
      year: enrichedPayload.report_year,
      sei_score: enrichedPayload.sei_score,
      sei_target: enrichedPayload.sei_target
    });

    // Get current timestamp
    const timestamp = new Date();

    // Open sheet
    const ss = SpreadsheetApp.openById(REGIONAL_LEADS_SHEET_ID);
    const sheet = ss.getSheetByName(REGIONAL_LEADS_TAB);

    if (!sheet) {
      return {
        success: false,
        message: `Sheet "${REGIONAL_LEADS_TAB}" not found in spreadsheet.`
      };
    }

    // ✅ Find next empty row in column B (starting from row 2)
    const lastRowWithData = getNextAppendRow_(sheet, 2, 2);
    const targetRow = lastRowWithData;

    Logger.log(`Target row for submission: ${targetRow}`);

    // ✅ Convert SEI values to decimal (85 becomes 0.85 for percentage formatting)
    const seiScoreDecimal = enrichedPayload.sei_score / 100;
    const seiTargetDecimal = enrichedPayload.sei_target / 100;

    // ============================================
    // WRITE DATA TO SPECIFIC COLUMN RANGES
    // ============================================

    // ✅ RANGE 1: Columns B:K (10 columns) - Metadata + SEI Score/Target
    const metadataAndSEI = [
      timestamp,                             // B: timestamp
      enrichedPayload.submitter_email,       // C: submitter_email
      monthAbbr,                             // D: report_month (✅ ABBREVIATED)
      enrichedPayload.report_year,           // E: report_year
      enrichedPayload.region,                // F: region
      enrichedPayload.role,                  // G: role
      enrichedPayload.lead_wdid,             // H: lead_wdid
      enrichedPayload.lead_name,             // I: lead_name
      seiScoreDecimal,                       // J: sei_score (✅ AS DECIMAL 0.85)
      seiTargetDecimal                       // K: sei_target (✅ AS DECIMAL 0.85)
    ];

    sheet.getRange(targetRow, 2, 1, 10).setValues([metadataAndSEI]);

    // ✅ RANGE 2: Column O (1 column) - SEI Comments
    const seiCommentsData = [
      enrichedPayload.sei_comments || ''     // O: sei_comments ✅
    ];

    sheet.getRange(targetRow, 15, 1, 1).setValues([seiCommentsData]);

    // ✅ RANGE 3: Columns P:Q (2 columns) - Revenue Actual/Projected
    const revenueData = [
      enrichedPayload.actual_revenue,        // P: actual_revenue
      enrichedPayload.projected_yearly_revenue // Q: projected_yearly_revenue
    ];

    sheet.getRange(targetRow, 16, 1, 2).setValues([revenueData]);

    // ✅ RANGE 4: Column U (1 column) - Revenue Comments
    const revenueCommentsData = [
      enrichedPayload.revenue_comments || '' // U: ctr_comments ✅
    ];

    sheet.getRange(targetRow, 21, 1, 1).setValues([revenueCommentsData]);

    // ✅ RANGE 5: Columns V:W (2 columns) - PPD Courses
    const ppdData = [
      enrichedPayload.courses_completed,     // V: courses_completed
      enrichedPayload.courses_required       // W: courses_required
    ];

    sheet.getRange(targetRow, 22, 1, 2).setValues([ppdData]);

    // ✅ RANGE 6: Column AA (1 column) - PPD Comments
    const ppdCommentsData = [
      enrichedPayload.ppd_comments || ''     // AA: ppd_comments ✅
    ];

    sheet.getRange(targetRow, 27, 1, 1).setValues([ppdCommentsData]);

    // ✅ Format currency columns (P and Q)
    formatRegionalLeadsCurrency_(targetRow);

    // ✅ Format SEI columns as percentage (J and K)
    formatSEIPercentages_(REGIONAL_LEADS_SHEET_ID, REGIONAL_LEADS_TAB, targetRow);

    Logger.log('Regional Lead KPIs submitted successfully:', {
      wdid: enrichedPayload.lead_wdid,
      name: enrichedPayload.lead_name,
      month: monthAbbr,
      year: enrichedPayload.report_year,
      row: targetRow
    });

    return {
      success: true,
      message: `KPIs submitted successfully for ${monthAbbr} ${enrichedPayload.report_year}.`
    };

  } catch (error) {
    Logger.log('Error in submitRegionalLeadKpis: ' + error.toString());
    return {
      success: false,
      message: 'Error submitting KPIs: ' + error.message
    };
  }
}

/**
 * Format currency columns in Regional_Leads_KPIs sheet
 * Formats columns P (actual_revenue) and Q (projected_yearly_revenue) as currency
 * 
 * @param {number} rowNumber - Row number to format
 */
function formatRegionalLeadsCurrency_(rowNumber) {
  try {
    const ss = SpreadsheetApp.openById(REGIONAL_LEADS_SHEET_ID);
    const sheet = ss.getSheetByName(REGIONAL_LEADS_TAB);
    if (!sheet) return;

    // Format column P (actual_revenue) - column 16
    const actualRange = sheet.getRange(rowNumber, 16, 1, 1); // Column P 
    actualRange.setNumberFormat('$#,##0.00');

    // Format column Q (projected_yearly_revenue) - column 17
    const projectedRange = sheet.getRange(rowNumber, 17, 1, 1); // Column Q
    projectedRange.setNumberFormat('$#,##0.00');

    Logger.log(`Formatted row ${rowNumber} with currency format (columns P & Q)`);

  } catch (error) {
    Logger.log('Error formatting currency: ' + error.toString());
    // Don't throw - formatting is non-critical
  }
}

/**
 * Format SEI columns as percentages (UNIFIED)
 * Works for any sheet - formats columns J (sei_score) and K (sei_target) as percentages
 * 
 * @param {string} sheetId - Spreadsheet ID
 * @param {string} tabName - Tab/sheet name
 * @param {number} rowNumber - Row number to format
 */
function formatSEIPercentages_(sheetId, tabName, rowNumber) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      Logger.log(`Sheet "${tabName}" not found in ${sheetId}`);
      return;
    }

    // Format column J (sei_score) - column 10
    const scoreRange = sheet.getRange(rowNumber, 10, 1, 1); // Column J
    scoreRange.setNumberFormat('0.00%');

    // Format column K (sei_target) - column 11
    const targetRange = sheet.getRange(rowNumber, 11, 1, 1); // Column K
    targetRange.setNumberFormat('0.00%');

    Logger.log(`Formatted row ${rowNumber} with percentage format (columns J & K) in ${tabName}`);

  } catch (error) {
    Logger.log('Error formatting SEI percentages: ' + error.toString());
    // Don't throw - formatting is non-critical
  }
}

// ============================================================================
// GLOBAL PRIME KPI FUNCTIONS
// ============================================================================

/**
 * Validate Prime WDID and retrieve details from roster
 * @param {string} wdid - Workday ID to validate
 * @return {Object} - { allowed: boolean, details: object }
 */
function validatePrimeWdid(wdid) {
  try {
    const normalizedWdid = String(wdid).trim();

    // Get details from roster
    const rosterDetails = getRosterDetailsByWDID(normalizedWdid);

    if (!rosterDetails) {
      Logger.log(`Prime WDID ${normalizedWdid} not found in roster`);
      return {
        allowed: false,
        details: null,
        message: 'WDID not found in Global QA Roster'
      };
    }

    // Return roster details
    return {
      allowed: true,
      details: {
        name: rosterDetails.name,
        email: rosterDetails.email,
        region: rosterDetails.region,
        role: 'Global Prime',
        program: rosterDetails.program,
        country: rosterDetails.country
      }
    };

  } catch (error) {
    Logger.log('Error in validatePrimeWdid: ' + error.toString());
    return {
      allowed: false,
      details: null,
      error: error.message
    };
  }
}

/**
 * Get Prime details by WDID from Global QA Roster
 * (Wrapper for consistency)
 * @param {string} wdid - Workday ID
 * @return {Object} - Prime details or null
 */
function getPrimeDetailsByWdid_(wdid) {
  try {
    const normalizedWdid = String(wdid).trim();

    // Get from roster
    const rosterDetails = getRosterDetailsByWDID(normalizedWdid);

    if (!rosterDetails) {
      return null;
    }

    return {
      name: rosterDetails.name,
      email: rosterDetails.email,
      region: rosterDetails.region,
      role: 'Global Prime'
    };

  } catch (error) {
    Logger.log('Error in getPrimeDetailsByWdid_: ' + error.toString());
    return null;
  }
}

/**
 * Submit Global Prime KPIs
 * Writes to Global_Prime_KPIs sheet
 * Populates columns: B:K, O, P, X, Y:Z, AD (same row)
 * 
 * @param {Object} payload - KPI data from form
 * @return {Object} - { success: boolean, message: string }
 */
function submitGlobalPrimeKpis(payload) {
  try {
    // Validate WDID and get roster details
    const validation = validatePrimeWdid(payload.prime_wdid);

    if (!validation.allowed) {
      return {
        success: false,
        message: validation.message || 'Access denied. Your WDID is not authorized to submit Global Prime KPIs.'
      };
    }

    // Enrich payload with roster data + hardcoded month/year
    const enrichedPayload = {
      ...payload,
      prime_name: validation.details.name,
      submitter_email: validation.details.email,
      region: validation.details.region,
      role: 'Global Prime',
      report_month: GLOBAL_PRIME_REPORT_MONTH,  // ✅ Hardcoded
      report_year: GLOBAL_PRIME_REPORT_YEAR     // ✅ Hardcoded
    };

    Logger.log('Submitting Global Prime KPIs:', {
      wdid: enrichedPayload.prime_wdid,
      name: enrichedPayload.prime_name,
      email: enrichedPayload.submitter_email,
      region: enrichedPayload.region,
      month: enrichedPayload.report_month,
      year: enrichedPayload.report_year
    });

    // Generate timestamp
    const timestamp = new Date();

    // Open sheet
    const ss = SpreadsheetApp.openById(GLOBAL_PRIME_SHEET_ID);
    const sheet = ss.getSheetByName(GLOBAL_PRIME_TAB);

    if (!sheet) {
      return {
        success: false,
        message: `Sheet "${GLOBAL_PRIME_TAB}" not found in spreadsheet.`
      };
    }

    // Find next empty row in column B
    const lastRowWithData = getNextAppendRow_(sheet, 2, 2);
    const targetRow = lastRowWithData;

    Logger.log(`Target row for submission: ${targetRow}`);

    // ✅ Convert SEI values to decimal (85 becomes 0.85)
    const seiScoreDecimal = enrichedPayload.sei_score / 100;
    const seiTargetDecimal = enrichedPayload.sei_target / 100;

    // ============================================
    // WRITE DATA TO SPECIFIC COLUMN RANGES
    // ============================================

    // ✅ RANGE 1: Columns B:K (10 columns) - Metadata + SEI
    const metadataAndSEI = [
      timestamp,                             // B: timestamp
      enrichedPayload.submitter_email,       // C: submitter_email
      enrichedPayload.report_month,          // D: report_month (already abbreviated)
      enrichedPayload.report_year,           // E: report_year
      enrichedPayload.region,                // F: region
      enrichedPayload.role,                  // G: role
      enrichedPayload.prime_wdid,            // H: prime_wdid
      enrichedPayload.prime_name,            // I: prime_name
      seiScoreDecimal,                       // J: sei_score (as decimal)
      seiTargetDecimal                       // K: sei_target (as decimal)
    ];

    sheet.getRange(targetRow, 2, 1, 10).setValues([metadataAndSEI]);

    // ✅ RANGE 2: Column O (1 column) - SEI Comments
    sheet.getRange(targetRow, 15, 1, 1).setValues([[enrichedPayload.gp_sei_comments || '']]);

    // ✅ RANGE 3: Column P (1 column) - Customer Sentiment
    sheet.getRange(targetRow, 16, 1, 1).setValues([[enrichedPayload.gp_cs_sentiment || '']]);

    // ✅ RANGE 4: Column X (1 column) - Customer Sentiment Comments
    sheet.getRange(targetRow, 24, 1, 1).setValues([[enrichedPayload.gp_cs_comments || '']]);

    // ✅ RANGE 5: Columns Y:Z (2 columns) - Revenue
    const revenueData = [
      enrichedPayload.actual_revenue,
      enrichedPayload.projected_yearly_revenue
    ];
    sheet.getRange(targetRow, 25, 1, 2).setValues([revenueData]);

    // ✅ RANGE 6: Column AD (1 column) - Revenue Comments
    sheet.getRange(targetRow, 30, 1, 1).setValues([[enrichedPayload.gp_ctr_comments || '']]);

    // ✅ Format SEI columns as percentage (J and K)
    formatSEIPercentages_(GLOBAL_PRIME_SHEET_ID, GLOBAL_PRIME_TAB, targetRow);

    // ✅ Format currency columns (Y and Z)
    formatGlobalPrimeCurrency_(targetRow);

    Logger.log('Global Prime KPIs submitted successfully:', {
      wdid: enrichedPayload.prime_wdid,
      name: enrichedPayload.prime_name,
      month: enrichedPayload.report_month,
      year: enrichedPayload.report_year,
      row: targetRow
    });

    return {
      success: true,
      message: `KPIs submitted successfully for ${enrichedPayload.report_month} ${enrichedPayload.report_year}.`
    };

  } catch (error) {
    Logger.log('Error in submitGlobalPrimeKpis: ' + error.toString());
    return {
      success: false,
      message: 'Error submitting KPIs: ' + error.message
    };
  }
}

/**
 * Format currency columns in Global_Prime_KPIs sheet
 * Formats columns Y (actual_revenue) and Z (projected_yearly_revenue) as currency
 * 
 * @param {number} rowNumber - Row number to format
 */
function formatGlobalPrimeCurrency_(rowNumber) {
  try {
    const ss = SpreadsheetApp.openById(GLOBAL_PRIME_SHEET_ID);
    const sheet = ss.getSheetByName(GLOBAL_PRIME_TAB);
    if (!sheet) return;

    // Format column Y (actual_revenue) - column 25
    const actualRange = sheet.getRange(rowNumber, 25, 1, 1);
    actualRange.setNumberFormat('$#,##0.00');

    // Format column Z (projected_yearly_revenue) - column 26
    const projectedRange = sheet.getRange(rowNumber, 26, 1, 1);
    projectedRange.setNumberFormat('$#,##0.00');

    Logger.log(`Formatted row ${rowNumber} with currency format (columns Y & Z)`);

  } catch (error) {
    Logger.log('Error formatting currency: ' + error.toString());
  }
}
// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * TEST FUNCTION: Debug Roster Snapshot Logic
 */
function testRosterSnapshot() {
  const testTlWdid = '10081381';
  const testMonth = 'December';
  const testYear = '2025';

  Logger.log('=== ROSTER SNAPSHOT DEBUG TEST ===');
  Logger.log(`TL WDID: ${testTlWdid}`);
  Logger.log(`Month: ${testMonth}`);
  Logger.log(`Year: ${testYear}`);

  const tlDetails = getRosterDetailsByWDID(testTlWdid);
  if (!tlDetails) {
    Logger.log(`❌ ERROR: TL WDID ${testTlWdid} not found in roster`);
    return;
  }

  Logger.log(`✓ TL Found: ${tlDetails.name}`);
  const teamRoster = getTeamMembersByTlWdid(testTlWdid);
  Logger.log(`✓ Found ${teamRoster.length} team members`);
}

/**
 * TEST FUNCTION: Debug Manager Roll-up Data Retrieval
 */
function testManagerRollupRetrieval() {
  const testManagerWdid = '10019678';
  const testMonth = 'December';
  const testYear = '2025';

  Logger.log('=== MANAGER ROLL-UP DEBUG TEST ===');
  Logger.log(`Manager WDID: ${testManagerWdid}`);
  Logger.log(`Month: ${testMonth}, Year: ${testYear}`);

  const monthAbbr = getMonthAbbreviation_(testMonth);
  Logger.log(`Month abbreviation: ${monthAbbr}`);
}

/**
 * TEST FUNCTION: Verify both Program and TL roll-up retrieval
 */
function testBothRollupFunctions() {
  const testManagerWdid = '10019678';
  const testMonth = 'December';
  const testYear = '2025';

  Logger.log('=== TESTING BOTH ROLL-UP FUNCTIONS ===');
  Logger.log(`Manager: ${testManagerWdid}`);
  Logger.log(`Period: ${testMonth} ${testYear}`);

  const programData = getManagerRollupByProgram(testManagerWdid, testMonth, testYear);
  Logger.log(`Programs found: ${programData.length}`);

  const tlData = getManagerRollupByTL(testManagerWdid, testMonth, testYear);
  Logger.log(`TLs found: ${tlData.length}`);
}

/**
 * TEST FUNCTION: Test Regional Leads validation
 */
function testRegionalLeadsValidation() {
  Logger.log('=== TESTING REGIONAL LEADS VALIDATION ===');

  // Test allowed WDID
  const test1 = validateWdidAllowed('10053071');
  Logger.log('Test 1 (10053071):');
  Logger.log(`  Allowed: ${test1.allowed}`);
  if (test1.allowed) {
    Logger.log(`  Name: ${test1.details.name}`);
    Logger.log(`  Email: ${test1.details.email}`);
    Logger.log(`  Region: ${test1.details.region}`);
    Logger.log(`  Role: ${test1.details.role}`);
  }
  Logger.log('');

  // Test another allowed WDID
  const test2 = validateWdidAllowed('10017324');
  Logger.log('Test 2 (10017324):');
  Logger.log(`  Allowed: ${test2.allowed}`);
  if (test2.allowed) {
    Logger.log(`  Name: ${test2.details.name}`);
    Logger.log(`  Email: ${test2.details.email}`);
    Logger.log(`  Region: ${test2.details.region}`);
    Logger.log(`  Role: ${test2.details.role}`);
  }
  Logger.log('');

  // Test disallowed WDID
  const test3 = validateWdidAllowed('99999999');
  Logger.log('Test 3 (99999999 - should be denied):');
  Logger.log(`  Allowed: ${test3.allowed}`);
  Logger.log('');

  Logger.log('=== TEST COMPLETE ===');
}

