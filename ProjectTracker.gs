/**
 * PROJECT TRACKER AUTOMATION 2026
 * Features: 
 * 1. Hourly Batched Email Replies on Status/Env/Blocker changes.
 *    - All edits within 1 hour are grouped into 1 reply per task thread.
 * 2. Threading based on "Internal Mail Subject".
 * 3. Daily Summary Reports (6 AM / 6 PM).
 * 4. Gmail Label Management (Parent: "A Assessment" with nested status labels)
 * 5. Google Tasks sync (main task = parent, sub-tasks = children under parent)
 *    - Parent moves to the Google Task list matching main Status (column I)
 *    - Subtasks: skip ONLY when Status + Production matches Completed, Closed/Cancelled,
 *      Closed–Billable, or Closed–Not Proceeding; all other combos add (remove if skip applies)
 * 
 * SETUP INSTRUCTIONS:
 * 1. Go to Extensions > Apps Script
 * 2. Paste ALL .gs files into the same project (see FILE LAYOUT below)
 * 3. Run createAllStatusLabels() once manually to create all labels
 * 4. Google Tasks: add Services (+) > Google Tasks API once, then run
 *    setupAllTrackersComplete() — permissions (Allow popup) are requested automatically.
 * 5. Edit ProjectTrackerUserConfig.gs — your email, sheet tab names, BCC, labels
 * 6. Dev Tracker: column U = Gmail Thread ID, column V = Google Task ID
 *    Dev Tracker Sub: column V = Google Subtask ID, column W = Main Google Task ID,
 *    column X = Google Nest Status (optional — run refreshSubNestStatusColumn())
 * 7. Create Google Task lists with the same names as CONFIG.STATUS_LABELS
 * 8. Run verifyGoogleTaskLists() then syncAllGoogleTasksFromDevTracker() once
 * 9. Set up triggers: run setupAllTriggers() in ProjectTrackerTriggers.gs
 *
 * FILE LAYOUT:
 *   ProjectTrackerUserConfig.gs       — YOUR email, sheet names, BCC (edit this file)
 *   ProjectTracker.gs                  — column layout, Gmail labels, sheet helpers
 *   ProjectTrackerGoogleTasks.gs       — Google Tasks sync (main + sub, bulk, reset)
 *   ProjectTrackerNotifications.gs   — email, hourly batch, daily/weekly summaries
 *   ProjectTrackerTriggers.gs          — onEdit + time-driven trigger install
 *   ProjectTrackerTests.gs             — manual tests + Google Tasks test suite
 *   TestAllTrackers.gs                 — unified setup + full test suite
 *
 * Google Tasks setup: see ProjectTrackerGoogleTasks.gs — verifyGoogleTaskLists(),
 *   syncExistingTasksToGoogle(), getGoogleTasksLastSyncLog()
 *
 * BCC: Edit BCC_EMAILS in ProjectTrackerUserConfig.gs.
 *
 * USER VALUES: edit ProjectTrackerUserConfig.gs (email, tab names, parent label, batch interval).
 * CONFIG below: column numbers, status labels, Google Tasks tuning only.
 */

const CONFIG = {
  // Dev Tracker columns (row 1 headers):
  // … R=Internal Mail Subject, U=Gmail Thread ID, V=Google Task ID
  MAIN_PROJECT_COL: 1,        // Column A: Project Name
  MAIN_TASK_NAME_COL: 3,      // Column C: Task
  MAIN_ASSIGNED_COL: 5,       // Column E: Assigned to
  MAIN_STATUS_COL: 9,         // Column I: Status
  MAIN_ENV_COL: 10,           // Column J: Env
  MAIN_INFORMED_COL: 13,      // Column M: Informed
  MAIN_BLOCKER_COL: 14,       // Column N: Support Needed / Blocker
  MAIN_DEPENDENCY_COL: 15,    // Column O: Dependency Detail
  MAIN_PAGE_URL_COL: 16,      // Column P: Page Url
  MAIN_SUBJECT_COL: 18,       // Column R: Internal Mail Subject
  MAIN_THREAD_ID_COL: 21,     // Column U: Gmail Thread ID (header resolved at runtime)
  MAIN_GOOGLE_TASK_ID_COL: 22, // Column V: Google Task ID (header resolved at runtime)
  
  SUB_LINK_COL: 2,            // Column B: Main Task Name (must match Dev Tracker column C Task)
  SUB_TASK_NAME_COL: 4,       // Column D: Task
  SUB_ASSIGNED_COL: 8,        // Column H: Assigned to
  SUB_STATUS_COL: 9,          // Column I: Status
  SUB_ENV_COL: 10,            // Column J: Env
  SUB_INFORMED_COL: 14,       // Column N: Informed
  SUB_BLOCKER_COL: 15,        // Column O: Support Needed / Blocker
  SUB_DEPENDENCY_COL: 16,     // Column P: Dependency Detail
  SUB_THREAD_ID_COL: 21,      // Column U: Gmail Thread ID (header resolved at runtime)
  SUB_GOOGLE_TASK_ID_COL: 22, // Column V: Google Subtask ID (header resolved at runtime)
  /** Column W: Main Google Task ID — from Dev Tracker column V; column B is parent source of truth. */
  SUB_PARENT_GOOGLE_TASK_ID_COL: 23,
  /** Column X: Google Nest Status (OK / ORPHAN / WRONG PARENT — run refreshSubNestStatusColumn()). */
  SUB_NEST_STATUS_COL: 24,

  // Main Dev Tracker — columns to monitor for batched email updates:
  // I=Status, J=Env, M=Informed, N=Blocker, O=Dependency
  MONITORED_COLS: [9, 10, 13, 14, 15],

  // Dev Tracker Sub — I=Status, J=Env, N=Informed, O=Blocker, P=Dependency; U/V/W/X = IDs + nest status
  SUB_MONITORED_COLS: [9, 10, 14, 15, 16],

  /** Status names — must match Google Task lists and Gmail labels under PARENT_LABEL in ProjectTrackerUserConfig.gs */
  STATUS_LABELS: [
    'Pipeline / To Do',
    'Requirements Not Clear',
    'In Progress',
    'On Hold / Blocked',
    'Waiting for Client Input',
    'Under Review & Testing',
    'Changes Requested / Rework',
    'Ready for Execution',
    'Completed',
    'Closed / Cancelled',
    'Closed – Billable',
    'Closed – Not Proceeding'
  ],

  /** Delay between rows during bulk Google Tasks sync (ms). Lower = faster but closer to quota limits. */
  GOOGLE_TASKS_SYNC_SLEEP_MS: 50,

  /**
   * Max main OR sub rows per execution (one "batch"). 0 = no row limit (time limit only).
   * Example: 300 → sync 300 subs, pause, auto-continue next 300 until done (safe for 2000–5000 subs).
   */
  GOOGLE_TASKS_SYNC_BATCH_SIZE: 300,

  /**
   * Stop each batch before Apps Script 6-minute hard timeout (ms). Default 4.5 min leaves save/trigger buffer.
   */
  GOOGLE_TASKS_SYNC_MAX_RUNTIME_MS: 270000,

  /**
   * When a batch pauses, schedule continueGoogleTasksBulkSync_() automatically (~1 min later).
   * Set false to continue manually with syncExistingTasksToGoogle().
   */
  GOOGLE_TASKS_AUTO_CONTINUE_BULK: true,

  /** Delay before auto-continue trigger fires (ms). Minimum ~60s for Apps Script time triggers. */
  GOOGLE_TASKS_AUTO_CONTINUE_DELAY_MS: 60000,

  /**
   * Dev Tracker Sub bulk scan order: bottom → row 2 (newest rows at bottom are processed first).
   * Bulk sync order: Phase 1 all Dev Tracker mains, Phase 2 all subs (main tasks always parents).
   */
  SUB_SHEET_SCAN_BOTTOM_UP: true,

  /**
   * Bulk sync: skip Dev Tracker rows with empty column C and blank sub rows entirely
   * (builds row lists — does not loop row 2…getLastRow()).
   */
  GOOGLE_TASKS_SKIP_EMPTY_ROWS: true,

  /**
   * Main Status (column I) edit:
   *   false — re-sync only sub rows linked to that main (TARGETED index, not full sheet loop)
   *   true  — sync main Google Task only; subs update on sub-row edit or bulk sync (lightest)
   */
  GOOGLE_TASKS_LIGHT_MAIN_STATUS_EDIT: false,

  /**
   * Google Tasks subtask title suffix (sheet column D is unchanged — suffix is Google Tasks only).
   *   'duplicate' — append [Main · #row] only when column D name appears on 2+ sub rows
   *   'always'    — every subtask gets [Main · #row] (safest; avoids collision with main task names)
   */
  GOOGLE_TASKS_SUBTASK_TITLE_SUFFIX: 'duplicate',

  /** Fallback background when status label is unknown / custom. */
  STATUS_COLOR_UNKNOWN: '#f0f0f0'
};

/**
 * @param {string|Array} raw From TRACKER_USER_CONFIG.BCC_EMAILS or script property
 * @returns {string[]}
 */
function parseTrackerBccEmails_(raw) {
  if (!raw) {
    return [];
  }
  if (Array.isArray(raw)) {
    return raw.map(function (s) { return (s || '').toString().trim(); })
      .filter(function (s) { return s !== ''; });
  }
  return raw.toString()
    .split(',')
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s !== ''; });
}

/** @returns {string} */
function getMainSheetName_() {
  return (TRACKER_USER_CONFIG.MAIN_SHEET || '').toString().trim();
}

/** @returns {string} */
function getSubSheetName_() {
  return (TRACKER_USER_CONFIG.SUB_SHEET || '').toString().trim();
}

/** @returns {string} */
function getPendingSheetName_() {
  return (TRACKER_USER_CONFIG.PENDING_SHEET || '').toString().trim();
}

/** @returns {string} */
function getCompletedSheetName_() {
  return (TRACKER_USER_CONFIG.COMPLETED_SHEET || '').toString().trim();
}

/** @returns {string} */
function getParentLabelName_() {
  return (TRACKER_USER_CONFIG.PARENT_LABEL || '').toString().trim();
}

/**
 * @param {string} sheetName
 * @returns {boolean}
 */
function isMainTrackerSheetName_(sheetName) {
  return (sheetName || '').toString().trim().toLowerCase() ===
    getMainSheetName_().toLowerCase();
}

/**
 * @param {string} sheetName
 * @returns {boolean}
 */
function isSubTrackerSheetName_(sheetName) {
  return (sheetName || '').toString().trim().toLowerCase() ===
    getSubSheetName_().toLowerCase();
}

/**
 * Script property keys for optional overrides (advanced). Normal edits: ProjectTrackerUserConfig.gs
 */
const SCRIPT_PROP_PROJECT_MY_EMAIL = "PROJECT_MY_EMAIL";
const SCRIPT_PROP_PROJECT_BCC_EMAILS = "PROJECT_BCC_EMAILS";
const SCRIPT_PROP_PROJECT_BATCH_INTERVAL_HOURS = "PROJECT_BATCH_INTERVAL_HOURS";

/** Whole-hour intervals Apps Script accepts for newTrigger().everyHours(). */
const ALLOWED_BATCH_INTERVAL_HOURS = [1, 2, 4, 6, 8, 12];

/**
 * Resolves the hourly batch interval.
 *   1. Script property PROJECT_BATCH_INTERVAL_HOURS (if set)
 *   2. TRACKER_USER_CONFIG.BATCH_INTERVAL_HOURS
 *   3. Falls back to 1 when not one of 1, 2, 4, 6, 8, 12
 * @returns {number}
 */
function getBatchIntervalHours_() {
  let value = parseInt(TRACKER_USER_CONFIG.BATCH_INTERVAL_HOURS, 10) || 1;

  try {
    const raw = PropertiesService
      .getScriptProperties()
      .getProperty(SCRIPT_PROP_PROJECT_BATCH_INTERVAL_HOURS);
    if (raw && raw.toString().trim() !== '') {
      const parsed = parseInt(raw.toString().trim(), 10);
      if (!isNaN(parsed) && parsed > 0) {
        value = parsed;
      }
    }
  } catch (e) {
    // Fall through to CONFIG default
  }

  if (ALLOWED_BATCH_INTERVAL_HOURS.indexOf(value) === -1) {
    Logger.log(
      'BATCH_INTERVAL_HOURS=' + value + ' is not allowed by Apps Script ' +
      '(must be one of ' + ALLOWED_BATCH_INTERVAL_HOURS.join(', ') + '). ' +
      'Falling back to 1 hour.'
    );
    return 1;
  }

  return value;
}

/**
 * Primary sender: script property → ProjectTrackerUserConfig.gs → active Google account.
 * @returns {string}
 */
function getProjectMyEmail_() {
  try {
    const raw = PropertiesService
      .getScriptProperties()
      .getProperty(SCRIPT_PROP_PROJECT_MY_EMAIL);
    if (raw && raw.toString().trim() !== '') {
      return raw.toString().trim();
    }
  } catch (e) {
    // Fall through
  }

  const configured = (TRACKER_USER_CONFIG.MY_EMAIL || '').toString().trim();
  if (configured) {
    return configured;
  }

  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    Logger.log('MY_EMAIL not set — edit MY_EMAIL in ProjectTrackerUserConfig.gs');
    return '';
  }
}

/**
 * BCC list: script property → ProjectTrackerUserConfig.gs → none.
 * @returns {string[]}
 */
function getProjectBccEmails_() {
  try {
    const raw = PropertiesService
      .getScriptProperties()
      .getProperty(SCRIPT_PROP_PROJECT_BCC_EMAILS);
    if (raw && raw.toString().trim() !== '') {
      const parsed = parseTrackerBccEmails_(raw);
      if (parsed.length > 0) {
        return parsed;
      }
    }
  } catch (e) {
    // Fall through
  }

  return parseTrackerBccEmails_(TRACKER_USER_CONFIG.BCC_EMAILS);
}

/**
 * Active spreadsheet, or the bound tracker ID saved by setupAllTrackers().
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet|null}
 */
function getTrackerSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      return active;
    }
  } catch (e) {
    // Fall through to script property
  }

  try {
    const id = PropertiesService.getScriptProperties().getProperty('LEAVE_TRACKER_SPREADSHEET_ID');
    if (id && id.toString().trim() !== '') {
      return SpreadsheetApp.openById(id.toString().trim());
    }
  } catch (error) {
    Logger.log('getTrackerSpreadsheet_ Error: ' + error.toString());
  }

  return null;
}

/**
 * Resolves the Dev Tracker sheet (exact name, then case-insensitive trim match).
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getMainTrackerSheet_() {
  const ss = getTrackerSpreadsheet_();
  if (!ss) {
    Logger.log(
      'No spreadsheet open. Open your Project Tracker Google Sheet in the browser, then run again.'
    );
    return null;
  }

  const exact = ss.getSheetByName(getMainSheetName_());
  if (exact) {
    return exact;
  }

  const target = getMainSheetName_().trim().toLowerCase();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const name = (sheets[i].getName() || '').toString().trim();
    if (name.toLowerCase() === target) {
      Logger.log('Dev Tracker sheet matched: "' + name + '"');
      return sheets[i];
    }
  }

  const names = [];
  for (let j = 0; j < sheets.length; j++) {
    names.push('"' + sheets[j].getName() + '"');
  }
  Logger.log(
    'Sheet "' + getMainSheetName_() + '" not found. Tabs in this file: ' + names.join(', ')
  );
  return null;
}

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getSubTrackerSheet_() {
  const ss = getTrackerSpreadsheet_();
  if (!ss) {
    return null;
  }

  const exact = ss.getSheetByName(getSubSheetName_());
  if (exact) {
    return exact;
  }

  const target = getSubSheetName_().trim().toLowerCase();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const name = (sheets[i].getName() || '').toString().trim();
    if (name.toLowerCase() === target) {
      return sheets[i];
    }
  }
  return null;
}

/** Column caches (projectThreadIdColCache_, projectGoogle*ColCache_) live in ProjectTrackerGoogleTasks.gs. */

/**
 * Finds "Gmail Thread ID" / "Thread ID" column from row 1, else column U (21).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
 * @returns {number}
 */
function findProjectThreadIdColumn_(mainSheet) {
  ensureColumnCachesFresh_(mainSheet, null);
  if (projectThreadIdColCache_) {
    return projectThreadIdColCache_;
  }
  if (!mainSheet) {
    return CONFIG.MAIN_THREAD_ID_COL;
  }
  const lastCol = mainSheet.getLastColumn();
  const headers = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let c = 0; c < headers.length; c++) {
    const h = (headers[c] || '').toString().toLowerCase().trim();
    if (
      h.indexOf('gmail thread') >= 0 ||
      h === 'thread id' ||
      h.indexOf('thread id') >= 0
    ) {
      projectThreadIdColCache_ = c + 1;
      Logger.log('Thread ID column from header: ' + projectThreadIdColCache_ + ' ("' + headers[c] + '")');
      return projectThreadIdColCache_;
    }
  }
  projectThreadIdColCache_ = CONFIG.MAIN_THREAD_ID_COL;
  return projectThreadIdColCache_;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
 * @param {number} row
 * @returns {string}
 */
function getProjectThreadIdForRow_(mainSheet, row) {
  const col = findProjectThreadIdColumn_(mainSheet);
  return (mainSheet.getRange(row, col).getValue() || '').toString().trim();
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
 * @param {number} row
 * @param {string} threadId
 */
function setProjectThreadIdForRow_(mainSheet, row, threadId) {
  const col = findProjectThreadIdColumn_(mainSheet);
  mainSheet.getRange(row, col).setValue(threadId);
}

/**
 * Run from Apps Script to debug one row (column U vs header, reply vs new thread).
 * @param {number} row Defaults to 2
 */
function moveMainRowToCompletedTrackerIfNeeded_(mainSheet, row) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return;

  const props = PropertiesService.getDocumentProperties();
  const guardKey = 'MOVE_TO_COMPLETED_IN_PROGRESS';
  try {
    if (props.getProperty(guardKey) === '1') return;

    const status = (mainSheet.getRange(row, CONFIG.MAIN_STATUS_COL).getValue() || '').toString().trim();
    if (!isCompletedOrClosedStatus_(status)) return;

    props.setProperty(guardKey, '1');

    const subject = (mainSheet.getRange(row, CONFIG.MAIN_SUBJECT_COL).getValue() || '').toString().trim();
    if (subject) {
      purgePendingChangesForSubject_(subject);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let completedSheet = ss.getSheetByName(getCompletedSheetName_());
    if (!completedSheet) {
      completedSheet = ss.insertSheet(getCompletedSheetName_());
    }

    const lastCol = mainSheet.getLastColumn();
    const header = mainSheet.getRange(1, 1, 1, lastCol).getValues();
    const movedAtHeader = 'Moved At';
    const targetCols = lastCol + 1;
    const headerWithMovedAt = [header[0].slice().concat([movedAtHeader])];

    if (completedSheet.getLastRow() === 0) {
      completedSheet.getRange(1, 1, 1, targetCols).setValues(headerWithMovedAt);
      completedSheet.setFrozenRows(1);
    } else if (completedSheet.getLastRow() === 1 && completedSheet.getLastColumn() === 1 && completedSheet.getRange(1, 1).getValue() === '') {
      completedSheet.getRange(1, 1, 1, targetCols).setValues(headerWithMovedAt);
      completedSheet.setFrozenRows(1);
    }

    const rowValues = mainSheet.getRange(row, 1, 1, lastCol).getValues()[0];
    const valuesToAppend = rowValues.concat([new Date()]);
    const targetRow = completedSheet.getLastRow() + 1;

    completedSheet.getRange(targetRow, 1, 1, valuesToAppend.length).setValues([valuesToAppend]);

    const taskNameIdx = CONFIG.MAIN_TASK_NAME_COL - 1;
    const written = completedSheet.getRange(targetRow, 1, 1, valuesToAppend.length).getValues()[0];
    const sourceTaskName = (rowValues[taskNameIdx] || '').toString().trim();
    const writtenTaskName = (written[taskNameIdx] || '').toString().trim();
    if (sourceTaskName && writtenTaskName !== sourceTaskName) {
      throw new Error(
        'Append verification failed for Dev Tracker row ' + row +
        ' — row not deleted from main sheet.'
      );
    }

    const parentGoogleTaskId = (rowValues[CONFIG.MAIN_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
    fillSubParentGoogleTaskIdForLinkedSubsIfEmpty_(sourceTaskName, parentGoogleTaskId);

    mainSheet.deleteRow(row);

    Logger.log('Moved row ' + row + ' to "' + getCompletedSheetName_() + '" due to status: ' + status);
  } catch (error) {
    Logger.log('moveMainRowToCompletedTrackerIfNeeded_ Error: ' + error.toString());
  } finally {
    props.deleteProperty(guardKey);
    try { lock.releaseLock(); } catch (e) {}
  }
}


function purgePendingChangesForSubject_(subject) {
  try {
    if (!subject || subject.toString().trim() === '') return;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pendingSheet = ss.getSheetByName(getPendingSheetName_());
    if (!pendingSheet) return;

    const values = pendingSheet.getDataRange().getValues();
    if (values.length <= 1) return;

    const toDelete = [];
    for (let i = 1; i < values.length; i++) {
      if ((values[i][1] || '').toString().trim() === subject.toString().trim()) {
        toDelete.push(i + 1); // sheet row number
      }
    }

    // delete bottom-up to preserve indices
    for (let i = toDelete.length - 1; i >= 0; i--) {
      pendingSheet.deleteRow(toDelete[i]);
    }
  } catch (error) {
    Logger.log('purgePendingChangesForSubject_ Error: ' + error.toString());
  }
}

/**
 * HELPER: Immediately sync Gmail labels for an existing main thread
 * Triggered on main sheet Status edits so label add/remove is automatic.
 */
function updateLabelsForMainRow(mainSheet, row) {
  try {
    const threadId = getProjectThreadIdForRow_(mainSheet, row);
    if (!threadId || threadId.toString().trim() === '') {
      Logger.log('No thread ID for row ' + row + '. Label sync skipped.');
      return;
    }

    const status = mainSheet.getRange(row, CONFIG.MAIN_STATUS_COL).getValue();
    manageThreadLabels(threadId, status);
    Logger.log('Immediate label sync done for row ' + row + ' with status: ' + status);
  } catch (error) {
    Logger.log('updateLabelsForMainRow Error: ' + error.toString());
  }
}

/** Per-execution cache for Gmail user labels — resets between trigger runs. */
let gmailUserLabelMapCache_ = null;

/**
 * Per-execution cache for Gmail user labels (avoids getUserLabels() on every status change).
 * @returns {Object<string, GmailLabel>}
 */
function getGmailUserLabelMap_() {
  if (!gmailUserLabelMapCache_) {
    gmailUserLabelMapCache_ = {};
    GmailApp.getUserLabels().forEach(function (label) {
      gmailUserLabelMapCache_[label.getName()] = label;
    });
  }
  return gmailUserLabelMapCache_;
}

/**
 * @param {string} labelName
 * @param {Object} labelMap
 * @returns {GmailLabel}
 */
function getOrCreateGmailLabel_(labelName, labelMap) {
  if (labelMap[labelName]) {
    return labelMap[labelName];
  }
  const created = GmailApp.createLabel(labelName);
  labelMap[labelName] = created;
  Logger.log('Created Gmail label: ' + labelName);
  return created;
}

/**
 * NEW FUNCTION: Manage Gmail labels with parent/child structure
 * Removes old status labels and adds new status label under "A Assessment"
 */
function manageThreadLabels(threadId, newStatus) {
  try {
    if (!threadId || threadId.toString().trim() === "") {
      Logger.log('No thread ID provided for label management');
      return;
    }
    
    const thread = GmailApp.getThreadById(threadId);
    if (!thread) {
      Logger.log('Thread not found: ' + threadId);
      return;
    }
    
    const labelMap = getGmailUserLabelMap_();
    
    const parentLabelName = getParentLabelName_();

    // STEP 1: Remove all existing status labels (children of parent label)
    CONFIG.STATUS_LABELS.forEach(function(statusName) {
      const nestedLabelName = parentLabelName + '/' + statusName;
      if (labelMap[nestedLabelName]) {
        thread.removeLabel(labelMap[nestedLabelName]);
        Logger.log('Removed old nested label: ' + nestedLabelName);
      }
    });
    
    // STEP 2: Add the parent label if not already present
    const parentLabel = getOrCreateGmailLabel_(parentLabelName, labelMap);
    thread.addLabel(parentLabel);
    
    // STEP 3: Add new status label (nested under parent)
    if (newStatus && newStatus.toString().trim() !== "") {
      const newStatusName = newStatus.toString().trim();
      const nestedLabelName = parentLabelName + '/' + newStatusName;
      const newLabel = getOrCreateGmailLabel_(nestedLabelName, labelMap);
      thread.addLabel(newLabel);
      Logger.log('Added new nested label: ' + nestedLabelName);
    }
    
  } catch (error) {
    Logger.log('manageThreadLabels Error: ' + error.toString());
  }
}

/**
 * SETUP HELPER: Create all nested labels at once (run once during setup)
 */
function createAllStatusLabels() {
  try {
    // Get all existing labels
    const allLabels = GmailApp.getUserLabels();
    const labelMap = {};
    allLabels.forEach(function(label) {
      labelMap[label.getName()] = label;
    });
    
    const parentLabelName = getParentLabelName_();

    // Create parent label
    if (!labelMap[parentLabelName]) {
      GmailApp.createLabel(parentLabelName);
      Logger.log('Created parent label: ' + parentLabelName);
    }
    
    // Create all nested labels
    CONFIG.STATUS_LABELS.forEach(function(statusName) {
      const nestedLabelName = parentLabelName + '/' + statusName;
      if (!labelMap[nestedLabelName]) {
        GmailApp.createLabel(nestedLabelName);
        Logger.log('Created nested label: ' + nestedLabelName);
      }
    });
    
    Logger.log('✅ All status labels created successfully!');
    Logger.log('Total labels created: ' + (CONFIG.STATUS_LABELS.length + 1));
    
  } catch (error) {
    Logger.log('createAllStatusLabels Error: ' + error.toString());
  }
}
