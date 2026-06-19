/************************************************************
 * UNIFIED TEST + TRIGGER SETUP  ← START HERE
 * ----------------------------------------------------------
 * ONE-CLICK PRODUCTION SETUP (run once in Apps Script):
 *   setupAllTrackersComplete()  — permissions + Gmail labels + Google Tasks + all triggers
 *
 * APPS SCRIPT FILES (paste all into the same project):
 *   ProjectTrackerUserConfig.gs       — YOUR email, sheet tab names, BCC
 *   ProjectTracker.gs                  — CONFIG, sheets, Gmail labels
 *   ProjectTrackerGoogleTasks.gs       — Google Tasks sync
 *   ProjectTrackerNotifications.gs     — email, hourly batch, summaries
 *   ProjectTrackerTriggers.gs          — onEdit + triggers
 *   ProjectTrackerTests.gs             — tests + runFullGoogleTasksTestSuite()
 *   TestAllTrackers.gs                 — this file (unified setup + full test suite)
 *   LeaveTracker.gs                — leave tracker (if used)
 *   (No need to run authorizeGoogleTasksAccess() separately — setup calls it for you.)
 *
 *   syncExistingTasksToGoogle()   — bulk sync: all Dev Tracker mains first, then subs bottom-up
 *   runFullGoogleTasksTestSuite() — ONE-CLICK test: reset + backfill + sync loop + diagnose (ProjectTrackerTests.gs)
 *   fixAllOrphanSubtasks()        — repair top-level orphan subtasks only (rows with column V ID)
 *   resetAllGoogleTasksAndRecreate() — DELETE all tracker-list tasks + clear column V + full recreate
 *
 * TRIGGERS ONLY (does NOT import existing tasks — use syncExistingTasksToGoogle for that):
 *   setupAllTrackers()            — permissions + install triggers
 *   removeDuplicateTrackerTriggers() — cleanup only (no new triggers)
 *
 * GOOGLE TASKS: add Services (+) → Google Tasks API once in the editor, then run
 *   setupAllTrackersComplete() — the Allow popup appears automatically during setup.
 *   Dev Tracker column U → "Gmail Thread ID"
 *   Dev Tracker column V → "Google Task ID"
 *   Dev Tracker Sub column U → "Gmail Thread ID"
 *   Dev Tracker Sub column V → "Google Subtask ID"
 *   Dev Tracker Sub column W → "Main Google Task ID" (auto from Dev Tracker column V when W empty)
 *   Dev Tracker Sub column X → "Google Nest Status" (run refreshSubNestStatusColumn() to fill)
 *   Google Tasks lists named same as CONFIG.STATUS_LABELS in ProjectTracker.gs
 *
 * TEST NOW (does not wait for 4 AM / 6 PM / hourly triggers):
 *   runAllTrackerTestsNow()         — full test immediately (fastest)
 *   runQuickTaskAndBatchTest()      — task edit + batch email only (no 1 hr wait)
 *   scheduleTestAllTrackersNow()    — one-time trigger ~1 min from now (background)
 *   setupAndTestAllTrackers()       — reinstall production triggers, then test
 *
 * EMAIL THREAD RESET (Gmail trash + clear sheet Thread IDs):
 *   previewResetAllTrackerEmailThreads() — counts only (safe)
 *   resetAllTrackerEmailThreads()        — set CONFIRM_DELETE_GMAIL_THREADS = true first
 *   recoverProjectThreadIdsFromGmail()   — search Gmail + create new threads if missing (column U)
 *
 * TRIGGERS INSTALLED BY setupAllTrackers() / setupAllTrackersComplete():
 *   onEditHandler              — Leave Tracker on sheet edit
 *   projectOnEditHandler       — Project Gmail labels + Google Tasks + pending queue
 *   sendHourlyBatch            — batched project emails (every 1 hr by default)
 *   sendDailyLeaveReport       — 4 AM and 4 PM daily
 *   sendDailySummary           — 6 AM and 6 PM daily
 *   sendWeeklySummary          — Friday 6 PM
 *   autoRecoverProjectThreadIds — daily 5 AM (refill column U Thread IDs)
 ************************************************************/

/**
 * Search Gmail (including Trash) for existing threads; if none, send a starter email and save new Thread ID to column U.
 */
const THREAD_RECOVERY_CONFIG = {
  SEARCH_IN_TRASH: true,
  CREATE_NEW_THREAD_IF_NOT_FOUND: true,
  /** 0 = all data rows; or set e.g. 2 and 10 to do rows 2–10 only (avoids 6 min timeout) */
  START_ROW: 2,
  END_ROW: 0,
  MS_BETWEEN_ROWS: 3500
};

/**
 * Set CONFIRM_DELETE_GMAIL_THREADS to true before resetAllTrackerEmailThreads().
 * Moves matching Gmail threads to Trash and clears Dev Tracker column U (Thread ID).
 */
const THREAD_RESET_CONFIG = {
  CONFIRM_DELETE_GMAIL_THREADS: false,
  TRASH_LEAVE_THREADS: true,
  TRASH_PROJECT_THREADS_FROM_SHEET: true,
  CLEAR_PROJECT_THREAD_IDS: true,
  CLEAR_PENDING_CHANGES: true
};

const ALL_TRACKER_TRIGGER_HANDLERS = [
  'sendDailyLeaveReport',
  'onEditHandler',
  'processPendingLeaveReport',
  'projectOnEditHandler',
  'onEdit',
  'sendHourlyBatch',
  'sendDailySummary',
  'sendWeeklySummary',
  'autoRecoverProjectThreadIds'
];

/** One-time “run tests soon” triggers only — not removed by setupAllTrackers() */
const QUICK_TEST_TRIGGER_HANDLERS = [
  'runAllTrackerTestsNow',
  'testAllTrackerFunctionality'
];

const TEST_ALL_CONFIG = {
  PROJECT_TEST_ROW: 2,
  /** Environment column (J) — avoids Completed-row move on Status (I) */
  PROJECT_ONEDIT_TEST_COL: 10,
  STEP_DELAY_MS: 2000,
  QUICK_STEP_DELAY_MS: 500
};

/**
 * Options for setupAllTrackersComplete() — one-click production setup.
 * Set SYNC_ALL_GOOGLE_TASKS to false on first run if you have many rows (avoid timeout).
 */
const SETUP_ALL_CONFIG = {
  REQUEST_ALL_PERMISSIONS: true,
  CREATE_GMAIL_LABELS: true,
  VERIFY_GOOGLE_TASK_LISTS: true,
  SYNC_ALL_GOOGLE_TASKS: true,
  INSTALL_TRIGGERS: true
};

/**
 * Requests OAuth for Gmail + Google Tasks in one pass (Allow popups if shown).
 * Called automatically by setupAllTrackersComplete() and setupAllTrackers().
 * You do not need to run authorizeGoogleTasksAccess() separately.
 * @returns {{ gmail: boolean, googleTasks: boolean }}
 */
function requestAllSetupPermissions_() {
  Logger.log('--- Permissions: Gmail + Google Tasks (click Allow if popups appear) ---');

  let gmailOk = false;
  let googleTasksOk = false;

  try {
    GmailApp.getUserLabels();
    gmailOk = true;
    Logger.log('Gmail access: OK');
  } catch (error) {
    Logger.log('Gmail access pending or denied: ' + error.toString());
  }

  googleTasksOk = authorizeGoogleTasksAccess();

  if (googleTasksOk) {
    Logger.log('Google Tasks access: OK');
  }

  return { gmail: gmailOk, googleTasks: googleTasksOk };
}

/**
 * Removes every installable trigger for Leave + Project handlers (fixes duplicates).
 * @returns {number} Count of triggers deleted
 */
function removeDuplicateTrackerTriggers() {
  const removed = removeAllTrackerTriggers_();
  Logger.log('Removed ' + removed + ' duplicate tracker trigger(s).');
  logTrackerTriggerInventory_('After cleanup');
  return removed;
}

/**
 * ONE-CLICK SETUP — run this once from Apps Script to enable everything.
 * 0. Request permissions (Gmail + Google Tasks — click Allow if popups appear)
 * 1. Gmail status labels (A Assessment / …)
 * 2. Verify Google Task lists match CONFIG.STATUS_LABELS
 * 3. Sync all Dev Tracker + Dev Tracker Sub rows to Google Tasks (column V)
 * 4. Install all Leave + Project triggers (deduped)
 *
 * Prerequisite (one-time, manual): Services (+) → Google Tasks API → Add
 */
function setupAllTrackersComplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== SETUP ALL TRACKERS (COMPLETE) ===');
  Logger.log('Spreadsheet: ' + ss.getName() + ' (' + ss.getId() + ')');
  Logger.log('User config: edit ProjectTrackerUserConfig.gs for email, sheet names, BCC');

  PropertiesService.getScriptProperties().setProperty(
    SCRIPT_PROP_SPREADSHEET_ID,
    ss.getId()
  );

  if (SETUP_ALL_CONFIG.REQUEST_ALL_PERMISSIONS) {
    requestAllSetupPermissions_();
  } else {
    Logger.log('--- Step 0: Permission requests skipped ---');
  }

  if (SETUP_ALL_CONFIG.CREATE_GMAIL_LABELS) {
    Logger.log('--- Step 1: Gmail status labels ---');
    createAllStatusLabels();
  } else {
    Logger.log('--- Step 1: Gmail labels skipped (SETUP_ALL_CONFIG.CREATE_GMAIL_LABELS = false) ---');
  }

  if (SETUP_ALL_CONFIG.VERIFY_GOOGLE_TASK_LISTS) {
    Logger.log('--- Step 2: Verify Google Task lists ---');
    if (isGoogleTasksApiEnabled_()) {
      verifyGoogleTaskLists();
    } else {
      Logger.log('Google Task lists not verified — add Services (+) > Google Tasks API, then run setup again.');
    }
  } else {
    Logger.log('--- Step 2: Google Task list verify skipped ---');
  }

  if (SETUP_ALL_CONFIG.SYNC_ALL_GOOGLE_TASKS && isGoogleTasksApiEnabled_()) {
    Logger.log('--- Step 3: Sync all rows to Google Tasks (column V) ---');
    syncAllGoogleTasksFromDevTracker();
  } else if (SETUP_ALL_CONFIG.SYNC_ALL_GOOGLE_TASKS) {
    Logger.log('--- Step 3: Google Tasks bulk sync skipped (add Services (+) > Google Tasks API first) ---');
  } else {
    Logger.log('--- Step 3: Google Tasks bulk sync skipped (edit sheet to sync per row) ---');
  }

  if (SETUP_ALL_CONFIG.INSTALL_TRIGGERS) {
    Logger.log('--- Step 4: Install all triggers ---');
    installAllTrackerTriggers_(ss);
  } else {
    Logger.log('--- Step 4: Trigger install skipped ---');
  }

  Logger.log('=== SETUP COMPLETE ===');
  Logger.log('Check Apps Script > Triggers (clock icon) — you should see ~9 enabled triggers.');
  Logger.log('Google Tasks sync runs automatically on edit via projectOnEditHandler.');
  Logger.log('Optional test: runAllTrackerTestsNow() or runQuickTaskAndBatchTest()');
}

/** Import/repair ALL Dev Tracker rows — re-nest subtasks under parents, remove orphan top-level duplicates. */
function syncExistingTasksToGoogle() {
  if (!authorizeGoogleTasksAccess()) {
    Logger.log('Add Services (+) > Google Tasks API, then run syncExistingTasksToGoogle() again.');
    return;
  }
  syncAllGoogleTasksFromDevTracker();
}

/** Alias for setupAllTrackersComplete() — same one-click setup. */
function enableAllTrackers() {
  setupAllTrackersComplete();
}

/**
 * One-time production setup: permissions + dedupe + full trigger set.
 * Prefer setupAllTrackersComplete() for first-time setup (labels + tasks + triggers).
 */
function setupAllTrackers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  PropertiesService.getScriptProperties().setProperty(
    SCRIPT_PROP_SPREADSHEET_ID,
    ss.getId()
  );

  requestAllSetupPermissions_();
  installAllTrackerTriggers_(ss);
}

/**
 * Creates all Leave + Project installable triggers (after dedupe).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function installAllTrackerTriggers_(ss) {
  const removed = removeAllTrackerTriggers_();
  if (removed > 0) {
    Logger.log('Removed ' + removed + ' old trigger(s) before reinstall.');
  }

  /************************************************
   * LEAVE TRACKER
   ************************************************/
  ScriptApp.newTrigger('sendDailyLeaveReport')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();

  ScriptApp.newTrigger('sendDailyLeaveReport')
    .timeBased()
    .everyDays(1)
    .atHour(16)
    .create();

  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  /************************************************
   * PROJECT TRACKER
   ************************************************/
  ScriptApp.newTrigger('projectOnEditHandler')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  const intervalHours = getBatchIntervalHours_();
  ScriptApp.newTrigger('sendHourlyBatch')
    .timeBased()
    .everyHours(intervalHours)
    .create();

  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(18)
    .everyWeeks(1)
    .create();

  // Daily safety net: refill any column-U cell that lost its Thread ID due
  // to Gmail index lag during a previous send. Runs at 5 AM, an hour before
  // the 6 AM daily summary, so IDs are correct for the next email cycle.
  ScriptApp.newTrigger('autoRecoverProjectThreadIds')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();

  Logger.log(
    'All tracker triggers installed (batch interval: ' + intervalHours + ' hour(s); ' +
    'auto-recovery sweep daily at 5 AM).'
  );
  Logger.log('projectOnEditHandler also syncs Gmail labels + Google Tasks on sheet edit.');
  logTrackerTriggerInventory_('After setup');
}

/**
 * Setup triggers, then run the full test suite.
 */
function setupAndTestAllTrackers() {
  setupAllTrackersComplete();
  Utilities.sleep(1000);
  runAllTrackerTestsNow();
}

/**
 * Run ALL tests immediately — no waiting for scheduled triggers (4 AM, hourly, etc.).
 * Select this in Apps Script and click Run.
 */
function runAllTrackerTestsNow() {
  Logger.log('=== QUICK TEST NOW (immediate) — calling all report/email functions directly ===');
  withQuickTestDelays_(function () {
    testAllTrackerFunctionality();
  });
}

/**
 * Quick test: main/sub task edit (onEdit) → pending queue → sendHourlyBatch immediately.
 * Does NOT wait for the hourly production trigger.
 */
function runQuickTaskAndBatchTest() {
  Logger.log('=== QUICK TASK + BATCH (onEdit → queue → email now) ===');
  const started = new Date();
  const results = [];

  ensureLeaveTrackerSpreadsheetIdForTest_();

  runTestStep_(results, 'Main task change (onEdit → pending queue)', function () {
    testProjectMainTaskOnEdit_();
  });

  runTestStep_(results, 'Sub-task change (onEdit → pending queue)', function () {
    testProjectSubTaskOnEdit_();
  });

  runTestStep_(results, 'Batch email now (sendHourlyBatch — skips 1 hour wait)', function () {
    sendHourlyBatch();
  });

  runTestStep_(results, 'Gmail labels for current status (if thread ID)', function () {
    testProjectStatusLabelsIfPossible_();
  });

  runTestStep_(results, 'Google Tasks sync (main + sub row, if Tasks API enabled)', function () {
    testGoogleTasksSyncIfPossible_();
  });

  logTestSummary_(results, started);
}

function withQuickTestDelays_(fn) {
  const savedDelay = TEST_ALL_CONFIG.STEP_DELAY_MS;
  TEST_ALL_CONFIG.STEP_DELAY_MS = TEST_ALL_CONFIG.QUICK_STEP_DELAY_MS;
  try {
    fn();
  } finally {
    TEST_ALL_CONFIG.STEP_DELAY_MS = savedDelay;
  }
}

/**
 * Schedules a one-time trigger to run tests in ~1 minute (background).
 * Use if the editor times out on a long run. Removes any prior quick-test trigger first.
 */
function scheduleTestAllTrackersNow() {
  removeQuickTestTriggers_();
  ScriptApp.newTrigger('runAllTrackerTestsNow')
    .timeBased()
    .after(60 * 1000)
    .create();
  Logger.log(
    'Quick test trigger created — runAllTrackerTestsNow will run in ~1 minute. ' +
    'Check Apps Script > Executions.'
  );
}

/**
 * Removes one-time quick-test triggers only (production triggers stay).
 */
function removeQuickTestTriggers() {
  const removed = removeQuickTestTriggers_();
  Logger.log('Removed ' + removed + ' quick-test trigger(s).');
}

/**
 * Single entry point: run all tracker tests in sequence.
 * View results in Apps Script > Executions > Logs.
 */
function testAllTrackerFunctionality() {
  const started = new Date();
  const results = [];

  Logger.log('=== TRACKER TEST RUN START: ' + started.toLocaleString() + ' ===');

  ensureLeaveTrackerSpreadsheetIdForTest_();

  runTestStep_(results, 'Leave Tracker: daily report email', function () {
    sendDailyLeaveReport();
  });

  runTestStep_(results, 'Project Tracker: create Gmail status labels', function () {
    createAllStatusLabels();
  });

  runTestStep_(results, 'Project Tracker: verify Google Task lists', function () {
    verifyGoogleTaskLists();
  });

  runTestStep_(results, 'Project Tracker: daily summary email', function () {
    sendDailySummary();
  });

  sleepBetweenSteps_();

  runTestStep_(results, 'Project Tracker: weekly summary email', function () {
    sendWeeklySummary();
  });

  sleepBetweenSteps_();

  runTestStep_(results, 'Project: main task edit → pending (onEdit, no 1 hr wait)', function () {
    testProjectMainTaskOnEdit_();
  });

  runTestStep_(results, 'Project: sub-task edit → pending (onEdit)', function () {
    testProjectSubTaskOnEdit_();
  });

  runTestStep_(results, 'Project: batch email now (sendHourlyBatch)', function () {
    sendHourlyBatch();
  });

  runTestStep_(results, 'Project: Gmail labels for status (current row)', function () {
    testProjectStatusLabelsIfPossible_();
  });

  runTestStep_(results, 'Project: Gmail label swap test (In Progress → Completed)', function () {
    testProjectLabelChangeIfPossible_();
  });

  runTestStep_(results, 'Project: Google Tasks sync (main + sub test row)', function () {
    testGoogleTasksSyncIfPossible_();
  });

  runTestStep_(results, 'Leave Tracker: onEdit (leave status column)', function () {
    testLeaveOnEditHandler_();
  });

  logTestSummary_(results, started);
}

/************************************************************
 * TRIGGER HELPERS
 ************************************************************/

/**
 * @returns {number}
 */
function removeQuickTestTriggers_() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    const handler = trigger.getHandlerFunction();
    if (QUICK_TEST_TRIGGER_HANDLERS.indexOf(handler) === -1) {
      return;
    }
    ScriptApp.deleteTrigger(trigger);
    removed++;
  });
  return removed;
}

/**
 * @returns {number}
 */
function removeAllTrackerTriggers_() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    const handler = trigger.getHandlerFunction();
    if (ALL_TRACKER_TRIGGER_HANDLERS.indexOf(handler) === -1) {
      return;
    }
    Logger.log(
      'Deleting trigger: ' + handler + ' — ' + describeTrackerTrigger_(trigger)
    );
    ScriptApp.deleteTrigger(trigger);
    removed++;
  });
  return removed;
}

function describeTrackerTrigger_(trigger) {
  try {
    if (trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
      return 'time-driven';
    }
    if (trigger.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS) {
      return 'spreadsheet ' + (trigger.getTriggerSourceId() || '');
    }
  } catch (e) {
    return 'unknown source';
  }
  return 'unknown';
}

function logTrackerTriggerInventory_(label) {
  const lines = [];
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    const handler = trigger.getHandlerFunction();
    if (ALL_TRACKER_TRIGGER_HANDLERS.indexOf(handler) === -1) {
      return;
    }
    lines.push('  • ' + handler + ' (' + describeTrackerTrigger_(trigger) + ')');
  });

  Logger.log('--- Trigger inventory: ' + label + ' (' + lines.length + ') ---');
  if (lines.length === 0) {
    Logger.log('  (none)');
  } else {
    for (let i = 0; i < lines.length; i++) {
      Logger.log(lines[i]);
    }
  }
}

/************************************************************
 * TEST HELPERS
 ************************************************************/

function runTestStep_(results, label, fn) {
  try {
    Logger.log('[RUN] ' + label);
    fn();
    results.push({ label: label, ok: true });
    Logger.log('[OK]  ' + label);
  } catch (err) {
    const message = err && err.toString ? err.toString() : String(err);
    results.push({ label: label, ok: false, error: message });
    Logger.log('[FAIL] ' + label + ' — ' + message);
  }
}

function sleepBetweenSteps_() {
  if (TEST_ALL_CONFIG.STEP_DELAY_MS > 0) {
    Utilities.sleep(TEST_ALL_CONFIG.STEP_DELAY_MS);
  }
}

function ensureLeaveTrackerSpreadsheetIdForTest_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(SCRIPT_PROP_SPREADSHEET_ID)) {
    props.setProperty(
      SCRIPT_PROP_SPREADSHEET_ID,
      SpreadsheetApp.getActiveSpreadsheet().getId()
    );
    Logger.log('Leave Tracker spreadsheet ID saved to script properties.');
  }
}

function buildMockEditEvent_(sheetName, row, col) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }
  if (row < 2) {
    throw new Error('Test row must be >= 2 (data row): ' + row);
  }
  return {
    source: ss,
    range: sheet.getRange(row, col)
  };
}

/**
 * Simulates editing a monitored column on Dev Tracker (same as real onEdit).
 * Uses Environment (J), not Status (I), so the row is not moved to Completed Tracker.
 */
function testProjectMainTaskOnEdit_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  if (!mainSheet) {
    throw new Error('Main sheet missing: ' + getMainSheetName_());
  }
  const row = TEST_ALL_CONFIG.PROJECT_TEST_ROW;
  const taskName = mainSheet.getRange(row, CONFIG.MAIN_TASK_NAME_COL).getValue();
  if (!taskName || taskName.toString().trim() === '') {
    throw new Error('Row ' + row + ' has no task name — pick a row with data.');
  }
  projectOnEditHandler(buildMockEditEvent_(
    getMainSheetName_(),
    row,
    TEST_ALL_CONFIG.PROJECT_ONEDIT_TEST_COL
  ));
  Logger.log('Main task onEdit simulated (col ' + TEST_ALL_CONFIG.PROJECT_ONEDIT_TEST_COL + ') → pending queue.');
}

/**
 * Simulates sub-sheet task edit → queues parent main row for batch email.
 */
function testProjectSubTaskOnEdit_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subSheet = ss.getSheetByName(getSubSheetName_());
  if (!subSheet) {
    Logger.log('Sub sheet not found — skipped.');
    return;
  }
  const row = TEST_ALL_CONFIG.PROJECT_TEST_ROW;
  const link = subSheet.getRange(row, CONFIG.SUB_LINK_COL).getValue();
  if (!link || link.toString().trim() === '') {
    Logger.log('Sub row ' + row + ' has no main-task link — skipped.');
    return;
  }
  projectOnEditHandler(buildMockEditEvent_(
    getSubSheetName_(),
    row,
    TEST_ALL_CONFIG.PROJECT_ONEDIT_TEST_COL
  ));
  Logger.log('Sub-task onEdit simulated → parent task queued for batch.');
}

/** Same as production when Status column changes (labels only; does not edit Status). */
function testProjectStatusLabelsIfPossible_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  if (!mainSheet) {
    throw new Error('Main sheet missing: ' + getMainSheetName_());
  }
  const row = TEST_ALL_CONFIG.PROJECT_TEST_ROW;
  const threadId = mainSheet.getRange(row, CONFIG.MAIN_THREAD_ID_COL).getValue();
  if (!threadId || threadId.toString().trim() === '') {
    Logger.log('No thread ID on row ' + row + ' — status labels skipped.');
    return;
  }
  updateLabelsForMainRow(mainSheet, row);
  Logger.log('Status labels applied from current row ' + row + ' status.');
}

function testProjectLabelChangeIfPossible_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  if (!mainSheet) {
    throw new Error('Main sheet missing: ' + getMainSheetName_());
  }
  const row = TEST_ALL_CONFIG.PROJECT_TEST_ROW;
  const threadId = mainSheet.getRange(row, CONFIG.MAIN_THREAD_ID_COL).getValue();
  if (!threadId || threadId.toString().trim() === '') {
    Logger.log('No thread ID on row ' + row + ' — label swap skipped (run after a threaded email exists).');
    return;
  }
  manageThreadLabels(threadId, 'In Progress');
  Utilities.sleep(TEST_ALL_CONFIG.STEP_DELAY_MS);
  manageThreadLabels(threadId, 'Completed');
}

/**
 * Syncs Google Tasks for test row on main + sub sheets (same as real onEdit).
 */
function testGoogleTasksSyncIfPossible_() {
  if (typeof isGoogleTasksApiEnabled_ === 'function' && !isGoogleTasksApiEnabled_()) {
    Logger.log('Google Tasks API not enabled — skipped. Add Services (+) > Google Tasks API.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  if (!mainSheet) {
    throw new Error('Main sheet missing: ' + getMainSheetName_());
  }

  const row = TEST_ALL_CONFIG.PROJECT_TEST_ROW;
  const taskName = mainSheet.getRange(row, CONFIG.MAIN_TASK_NAME_COL).getValue();
  if (!taskName || taskName.toString().trim() === '') {
    Logger.log('Row ' + row + ' has no task name — Google Tasks test skipped.');
    return;
  }

  syncGoogleTaskForMainRow(mainSheet, row);
  const mainTaskId = getGoogleTaskIdForMainRow_(mainSheet, row);
  Logger.log('Main Google Task ID (column V): ' + (mainTaskId || '(not saved — check list names / API)'));

  const subSheet = ss.getSheetByName(getSubSheetName_());
  if (!subSheet) {
    Logger.log('Sub sheet not found — subtask sync skipped.');
    return;
  }

  const link = subSheet.getRange(row, CONFIG.SUB_LINK_COL).getValue();
  if (!link || link.toString().trim() === '') {
    Logger.log('Sub row ' + row + ' has no main-task link — subtask sync skipped.');
    return;
  }

  syncGoogleSubtaskForSubRow(subSheet, row);
  const subTaskId = getGoogleSubtaskIdForSubRow_(subSheet, row);
  Logger.log('Sub Google Task ID (column V): ' + (subTaskId || '(not saved — check parent link / subtask name)'));
}

function testLeaveOnEditHandler_() {
  const sheet = getLeaveTrackerSheet_();
  if (!sheet) {
    throw new Error('Leave sheet missing: ' + SHEET_NAME);
  }
  const cols = getLeaveColumnIndexes_(sheet);
  const leaveStatusCol = cols.leaveStatus + 1;
  const row = TEST_ALL_CONFIG.PROJECT_TEST_ROW;
  const statusValue = sheet.getRange(row, leaveStatusCol).getValue();
  if (!statusValue || statusValue.toString().trim() === '') {
    Logger.log('Row ' + row + ' has no Leave Status — onEdit test skipped.');
    return;
  }

  // 1) Fire the onEdit handler — this now SCHEDULES a deferred send
  //    rather than sending immediately (debounce/coalescing).
  onEditHandler(buildMockEditEvent_(SHEET_NAME, row, leaveStatusCol));

  // 2) Verify a deferred trigger was actually scheduled.
  const triggers = ScriptApp.getProjectTriggers();
  const scheduled = triggers.some(function (t) {
    return t.getHandlerFunction() === 'processPendingLeaveReport';
  });
  if (!scheduled) {
    throw new Error(
      'onEditHandler did not schedule a deferred leave report (processPendingLeaveReport trigger missing).'
    );
  }

  // 3) Fast-forward the debounce window so the test doesn't wait
  //    ~30s — invoke the deferred handler synchronously. It will
  //    clean up its own trigger and send the report once.
  processPendingLeaveReport();
}

function logTestSummary_(results, started) {
  const passed = results.filter(function (r) { return r.ok; }).length;
  const failed = results.length - passed;

  Logger.log('');
  Logger.log('=== TRACKER TEST SUMMARY ===');
  for (let i = 0; i < results.length; i++) {
    const r = results[i];
    Logger.log((r.ok ? '✅' : '❌') + ' ' + r.label + (r.error ? ' — ' + r.error : ''));
  }
  Logger.log('Passed: ' + passed + ' | Failed: ' + failed + ' | Total: ' + results.length);
  Logger.log('Finished: ' + new Date().toLocaleString() + ' (started ' + started.toLocaleString() + ')');
  Logger.log('=== END ===');
}

/************************************************************
 * EMAIL THREAD RESET + CONFLICT DIAGNOSIS
 ************************************************************/

/**
 * Logs why Leave vs Project mail may have mixed (no changes made).
 */
function diagnoseTrackerConflicts() {
  Logger.log('=== TRACKER CONFLICT AUDIT ===');
  auditTrackerConflicts_();
}

function diagnoseEmailThreadConflict() {
  diagnoseTrackerConflicts();
}

function auditTrackerConflicts_() {
  Logger.log('=== EMAIL THREAD DIAGNOSIS ===');
  Logger.log(
    'OK: Leave uses sendLeaveTrackerEmail_(), Project uses sendProjectTrackerEmail_() ' +
    '(old duplicate sendThreadedEmail was removed).'
  );
  Logger.log(
    'OK: Project edit handler is projectOnEditHandler (not onEdit) to avoid simple+installable double-run.'
  );
  Logger.log(
    'OK: Leave edit handler is onEditHandler — separate from project.'
  );
  Logger.log('Leave thread: Gmail subject "' + getLeaveEmailSubjectForReset_() + '" (no sheet column).');
  Logger.log('Project thread: column R = Internal Mail Subject, column U = Thread ID (Dev Tracker).');
  Logger.log('Google Tasks: column V = Google Task ID (main + sub); synced via projectOnEditHandler.');
  Logger.log('Do NOT put "' + getLeaveEmailSubjectForReset_() + '" in Dev Tracker column R.');

  const preview = previewResetAllTrackerEmailThreads_();
  Logger.log('Leave threads found: ' + preview.leaveThreads);
  Logger.log('Project thread IDs on sheet: ' + preview.projectThreadIds);
  Logger.log('Pending queue rows: ' + preview.pendingRows);

  Logger.log('--- Trigger check (run setupAllTrackersComplete once to enable all) ---');
  logTrackerTriggerInventory_('Current');

  Logger.log('--- Setup overlap (avoid running all three) ---');
  Logger.log('Use setupAllTrackersComplete() or enableAllTrackers() — not setupLeaveTracker() + setupAllTriggers() separately.');

  Logger.log('--- Data rules ---');
  Logger.log('Never use subject "' + getLeaveEmailSubjectForReset_() + '" in Dev Tracker column R.');

  Logger.log('=== END CONFLICT AUDIT ===');
}

/**
 * Safe preview — counts threads/rows that would be reset (nothing deleted).
 */
function previewResetAllTrackerEmailThreads() {
  const preview = previewResetAllTrackerEmailThreads_();
  Logger.log('=== RESET PREVIEW (no changes) ===');
  Logger.log('Leave Gmail threads to trash: ' + preview.leaveThreads);
  Logger.log('Project Gmail threads to trash (from column U): ' + preview.projectThreads);
  Logger.log('Dev Tracker Thread ID cells to clear: ' + preview.projectThreadIds);
  Logger.log('Pending Changes rows to clear: ' + preview.pendingRows);
  Logger.log('Set THREAD_RESET_CONFIG.CONFIRM_DELETE_GMAIL_THREADS = true then run resetAllTrackerEmailThreads()');
}

/**
 * Trash old Gmail threads, clear column U, clear pending queue — fresh start.
 * Requires THREAD_RESET_CONFIG.CONFIRM_DELETE_GMAIL_THREADS = true
 */
function resetAllTrackerEmailThreads() {
  if (!THREAD_RESET_CONFIG.CONFIRM_DELETE_GMAIL_THREADS) {
    Logger.log(
      'SAFETY: Set THREAD_RESET_CONFIG.CONFIRM_DELETE_GMAIL_THREADS = true in TestAllTrackers.gs, then run again.'
    );
    previewResetAllTrackerEmailThreads();
    return;
  }

  Logger.log('=== RESET ALL TRACKER EMAIL THREADS ===');
  const summary = {
    leaveTrashed: 0,
    projectTrashed: 0,
    threadIdsCleared: 0,
    pendingCleared: 0
  };

  if (THREAD_RESET_CONFIG.TRASH_LEAVE_THREADS) {
    summary.leaveTrashed = trashAllLeaveTrackerGmailThreads_();
  }

  if (THREAD_RESET_CONFIG.TRASH_PROJECT_THREADS_FROM_SHEET) {
    summary.projectTrashed = trashProjectThreadsFromSheetIds_();
  }

  if (THREAD_RESET_CONFIG.CLEAR_PROJECT_THREAD_IDS) {
    summary.threadIdsCleared = clearProjectThreadIdColumn_();
  }

  if (THREAD_RESET_CONFIG.CLEAR_PENDING_CHANGES) {
    summary.pendingCleared = clearPendingChangesSheet_();
  }

  Logger.log('Leave threads trashed: ' + summary.leaveTrashed);
  Logger.log('Project threads trashed: ' + summary.projectTrashed);
  Logger.log('Thread ID cells cleared (col U): ' + summary.threadIdsCleared);
  Logger.log('Pending queue rows cleared: ' + summary.pendingCleared);
  Logger.log('Next Leave email → new "' + getLeaveEmailSubjectForReset_() + '" thread.');
  Logger.log('Next Project email per task → new thread (ID saved to column U again).');
  Logger.log('Column U was cleared on purpose. Run recoverProjectThreadIdsFromGmail() to refill from Gmail.');
  Logger.log('=== RESET COMPLETE ===');
}

/**
 * Refill column U: find existing Gmail threads OR create a new thread per row.
 * Run after reset when column U is empty and old mail was trashed.
 */
function recoverProjectThreadIdsFromGmail() {
  recoverOrCreateProjectThreadIds_();
}

/** Same as recoverProjectThreadIdsFromGmail — explicit name */
function recoverOrCreateProjectThreadIds() {
  recoverOrCreateProjectThreadIds_();
}

/**
 * TRIGGER: Daily safety net for orphaned Thread IDs.
 *
 * If sendProjectTrackerEmail_'s post-send save ever loses a race against
 * Gmail's index lag (column U stays empty even though the email went out),
 * subsequent batches would spawn yet another fresh thread every hour. This
 * sweep re-binds those rows to the Gmail thread that actually exists.
 *
 * Unlike recoverOrCreateProjectThreadIds(), this never CREATES a new thread
 * on a miss — that would risk daily spam if Gmail can't find an existing
 * thread temporarily. It only fills in IDs for threads it can confirm.
 */
/**
 * Extended thread recovery with THREAD_RECOVERY_CONFIG — delegates to ProjectTracker.gs
 * autoRecoverProjectThreadIds() for the daily trigger; use this for manual full recovery.
 */
function recoverProjectThreadIdsFull() {
  const previousCreateFlag = THREAD_RECOVERY_CONFIG.CREATE_NEW_THREAD_IF_NOT_FOUND;
  THREAD_RECOVERY_CONFIG.CREATE_NEW_THREAD_IF_NOT_FOUND = false;
  try {
    Logger.log('=== AUTO RECOVERY (read-only, no new threads) ===');
    recoverOrCreateProjectThreadIds_();
  } catch (err) {
    Logger.log('recoverProjectThreadIdsFull Error: ' + err.toString());
  } finally {
    THREAD_RECOVERY_CONFIG.CREATE_NEW_THREAD_IF_NOT_FOUND = previousCreateFlag;
  }
}

function recoverOrCreateProjectThreadIds_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  if (!mainSheet) {
    Logger.log('Dev Tracker sheet not found.');
    return;
  }

  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows on Dev Tracker.');
    return;
  }

  const startRow = Math.max(2, THREAD_RECOVERY_CONFIG.START_ROW || 2);
  const endRow = THREAD_RECOVERY_CONFIG.END_ROW > 0
    ? Math.min(THREAD_RECOVERY_CONFIG.END_ROW, lastRow)
    : lastRow;

  let foundExisting = 0;
  let createdNew = 0;
  let skippedHasId = 0;
  let skippedNoSubject = 0;
  let failed = 0;

  Logger.log('=== RECOVER / CREATE THREAD IDs (column U), rows ' + startRow + '–' + endRow + ' ===');
  Logger.log('CREATE_NEW_THREAD_IF_NOT_FOUND=' + THREAD_RECOVERY_CONFIG.CREATE_NEW_THREAD_IF_NOT_FOUND);

  for (let row = startRow; row <= endRow; row++) {
    const existingId = (mainSheet.getRange(row, CONFIG.MAIN_THREAD_ID_COL).getValue() || '').toString().trim();
    if (existingId) {
      skippedHasId++;
      continue;
    }

    const subjectInfo = getProjectSubjectForRow_(mainSheet, row);
    if (!subjectInfo) {
      skippedNoSubject++;
      Logger.log('Row ' + row + ': skipped — no subject in column R and no task name');
      continue;
    }

    let threadId = findProjectGmailThreadIdForSubject_(subjectInfo.gmailSubject, subjectInfo.taskName);

    if (threadId) {
      mainSheet.getRange(row, CONFIG.MAIN_THREAD_ID_COL).setValue(threadId);
      foundExisting++;
      Logger.log('Row ' + row + ': found existing thread ' + threadId);
    } else if (THREAD_RECOVERY_CONFIG.CREATE_NEW_THREAD_IF_NOT_FOUND) {
      threadId = createNewProjectThreadForRow_(mainSheet, row, subjectInfo);
      if (threadId) {
        createdNew++;
        Logger.log('Row ' + row + ': NEW thread created ' + threadId + ' — ' + subjectInfo.gmailSubject);
      } else {
        failed++;
        Logger.log('Row ' + row + ': failed to create thread — ' + subjectInfo.gmailSubject);
      }
    } else {
      failed++;
      Logger.log('Row ' + row + ': not found (create disabled) — ' + subjectInfo.gmailSubject);
    }

    if (THREAD_RECOVERY_CONFIG.MS_BETWEEN_ROWS > 0) {
      Utilities.sleep(THREAD_RECOVERY_CONFIG.MS_BETWEEN_ROWS);
    }
  }

  Logger.log('Found in Gmail: ' + foundExisting + ' | New threads sent: ' + createdNew +
    ' | Already had ID: ' + skippedHasId + ' | No subject: ' + skippedNoSubject + ' | Failed: ' + failed);
  Logger.log('=== DONE ===');
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
 * @param {number} row
 * @returns {{gmailSubject: string, taskName: string}|null}
 */
function getProjectSubjectForRow_(mainSheet, row) {
  const subjectRaw = (mainSheet.getRange(row, CONFIG.MAIN_SUBJECT_COL).getValue() || '').toString().trim();
  const taskName = (mainSheet.getRange(row, CONFIG.MAIN_TASK_NAME_COL).getValue() || '').toString().trim();
  let gmailSubject = subjectRaw;
  if (!gmailSubject && taskName) {
    gmailSubject = 'TASK UPDATE: ' + taskName;
  }
  if (!gmailSubject) {
    return null;
  }
  if (typeof ensureProjectEmailSubject_ === 'function') {
    gmailSubject = ensureProjectEmailSubject_(gmailSubject);
  }
  return { gmailSubject: gmailSubject, taskName: taskName || gmailSubject };
}

/**
 * Sends one starter email and saves Thread ID to column U (via sendProjectTrackerEmail_).
 * @returns {string|null}
 */
function createNewProjectThreadForRow_(mainSheet, row, subjectInfo) {
  const status = (mainSheet.getRange(row, 9).getValue() || '').toString().trim();
  const taskLabel = subjectInfo.taskName || subjectInfo.gmailSubject;
  const plain = 'Project Tracker: Gmail thread re-started for task — ' + taskLabel;
  const html = '<html><body style="font-family: Arial, sans-serif;">' +
    '<p>This message <strong>starts a new Gmail thread</strong> for Dev Tracker task:</p>' +
    '<p><strong>' + (typeof escapeHtml === 'function' ? escapeHtml(taskLabel) : taskLabel) + '</strong></p>' +
    '<p style="color:#7f8c8d;font-size:12px;">Future task updates will reply in this thread. ' +
    'Auto-generated by Project Tracker (thread recovery).</p></body></html>';

  sendProjectTrackerEmail_(
    subjectInfo.gmailSubject,
    plain,
    html,
    mainSheet,
    row,
    status
  );

  Utilities.sleep(2000);
  const id = (mainSheet.getRange(row, CONFIG.MAIN_THREAD_ID_COL).getValue() || '').toString().trim();
  return id || null;
}

/**
 * @param {string} sheetSubject
 * @param {string} taskName
 * @returns {string|null} Gmail thread ID
 */
function findProjectGmailThreadIdForSubject_(sheetSubject, taskName) {
  const subjectsToTry = buildProjectSubjectSearchList_(sheetSubject, taskName);
  const queries = buildGmailSearchQueriesForSubjects_(subjectsToTry);
  const triedThreads = {};

  for (let q = 0; q < queries.length; q++) {
    let threads;
    try {
      threads = GmailApp.search(queries[q], 0, 20);
    } catch (e) {
      continue;
    }

    for (let t = 0; t < threads.length; t++) {
      const thread = threads[t];
      const id = thread.getId();
      if (triedThreads[id]) {
        continue;
      }
      triedThreads[id] = true;

      if (!isLikelyProjectTrackerGmailThread_(thread, sheetSubject, subjectsToTry, taskName)) {
        continue;
      }

      return id;
    }
  }

  return null;
}

function stripDevTrackerPrefix_(subject) {
  return (subject || '').toString().replace(/^\[Dev Tracker\]\s*/i, '').trim();
}

function buildProjectSubjectSearchList_(sheetSubject, taskName) {
  const list = [];
  const seen = {};
  function add(s) {
    const t = (s || '').toString().trim();
    if (!t || seen[t]) {
      return;
    }
    seen[t] = true;
    list.push(t);
  }

  const raw = (sheetSubject || '').toString().trim();
  add(raw);
  add(stripDevTrackerPrefix_(raw));
  add(normalizeEmailSubject_(raw));
  add(normalizeEmailSubject_(stripDevTrackerPrefix_(raw)));

  if (typeof ensureProjectEmailSubject_ === 'function') {
    add(ensureProjectEmailSubject_(raw));
    add(ensureProjectEmailSubject_(stripDevTrackerPrefix_(raw)));
  }

  if (raw.indexOf('[Dev Tracker]') !== 0) {
    add('[Dev Tracker] ' + raw);
  }

  const tn = (taskName || '').toString().trim();
  if (tn) {
    add('TASK UPDATE: ' + tn);
    add('[Dev Tracker] TASK UPDATE: ' + tn);
    add('Project Update: ' + tn);
  }

  return list;
}

function buildGmailSearchQueriesForSubjects_(subjectsToTry) {
  const queries = [];
  const seenQ = {};
  function addQuery(q) {
    if (!q || seenQ[q]) {
      return;
    }
    seenQ[q] = true;
    queries.push(q);
  }

  for (let s = 0; s < subjectsToTry.length; s++) {
    const esc = subjectsToTry[s].replace(/"/g, '\\"');
    addQuery('subject:"' + esc + '"');
    if (THREAD_RECOVERY_CONFIG.SEARCH_IN_TRASH) {
      addQuery('in:anywhere subject:"' + esc + '"');
    }
    const core = stripDevTrackerPrefix_(subjectsToTry[s]);
    if (core.indexOf('TASK UPDATE:') >= 0) {
      const after = core.substring(core.indexOf('TASK UPDATE:')).replace(/"/g, '');
      addQuery('subject:"' + after + '"');
      if (THREAD_RECOVERY_CONFIG.SEARCH_IN_TRASH) {
        addQuery('in:anywhere subject:"' + after + '"');
      }
    }
  }

  return queries;
}

function normalizeEmailSubject_(subject) {
  return (subject || '').toString().replace(/^(Re:\s*)+/gi, '').trim();
}

/**
 * @param {GoogleAppsScript.Gmail.GmailThread} thread
 * @param {string} sheetSubject
 * @param {string[]} subjectCandidates
 * @returns {boolean}
 */
function isLikelyProjectTrackerGmailThread_(thread, sheetSubject, subjectCandidates, taskName) {
  if (typeof isLeaveTrackerThread_ === 'function' && isLeaveTrackerThread_(thread)) {
    return false;
  }

  const leaveSubject = getLeaveEmailSubjectForReset_();
  const firstSub = normalizeEmailSubject_((thread.getFirstMessageSubject() || '').trim());
  if (firstSub === leaveSubject) {
    return false;
  }

  const normalizedSheet = normalizeEmailSubject_(stripDevTrackerPrefix_(sheetSubject));
  const normalizedTask = (taskName || '').toString().trim().toLowerCase();
  const messages = thread.getMessages();

  for (let m = 0; m < messages.length; m++) {
    const msgSub = normalizeEmailSubject_((messages[m].getSubject() || '').trim());
    const msgSubLower = msgSub.toLowerCase();

    for (let c = 0; c < subjectCandidates.length; c++) {
      const cand = normalizeEmailSubject_(subjectCandidates[c]);
      if (msgSub === cand || msgSub === normalizedSheet) {
        return true;
      }
      if (cand.length > 10 && msgSub.indexOf(cand) >= 0) {
        return true;
      }
    }

    if (normalizedTask && msgSubLower.indexOf(normalizedTask) >= 0 &&
        (msgSubLower.indexOf('task update') >= 0 || msgSubLower.indexOf('project update') >= 0)) {
      return true;
    }

    const body = (messages[m].getPlainBody() || '') + (messages[m].getBody() || '');
    if (body.indexOf('Auto-generated by Project Tracker') >= 0) {
      return true;
    }
    if (normalizedTask && body.toLowerCase().indexOf(normalizedTask.toLowerCase()) >= 0 &&
        body.indexOf('TASK UPDATE') >= 0) {
      return true;
    }
  }

  const firstLower = firstSub.toLowerCase();
  if (firstLower.indexOf('task update') >= 0 || firstLower.indexOf('project update') >= 0) {
    if (!normalizedTask || firstLower.indexOf(normalizedTask) >= 0) {
      return true;
    }
  }

  return false;
}

function getLeaveEmailSubjectForReset_() {
  if (typeof EMAIL_SUBJECT === 'string' && EMAIL_SUBJECT) {
    return EMAIL_SUBJECT;
  }
  return 'Team Leave Tracker Report';
}

function previewResetAllTrackerEmailThreads_() {
  return {
    leaveThreads: countLeaveTrackerGmailThreads_(),
    projectThreads: countProjectThreadsFromSheetIds_(),
    projectThreadIds: countProjectThreadIdCells_(),
    pendingRows: countPendingChangeRows_()
  };
}

function countLeaveTrackerGmailThreads_() {
  const threads = GmailApp.search(
    'subject:"' + getLeaveEmailSubjectForReset_().replace(/"/g, '\\"') + '"',
    0,
    100
  );
  let n = 0;
  for (let i = 0; i < threads.length; i++) {
    if (typeof isLeaveTrackerThread_ === 'function' && isLeaveTrackerThread_(threads[i])) {
      n++;
    }
  }
  return n;
}

function trashAllLeaveTrackerGmailThreads_() {
  const threads = GmailApp.search(
    'subject:"' + getLeaveEmailSubjectForReset_().replace(/"/g, '\\"') + '"',
    0,
    100
  );
  let trashed = 0;
  for (let i = 0; i < threads.length; i++) {
    if (typeof isLeaveTrackerThread_ === 'function' && isLeaveTrackerThread_(threads[i])) {
      threads[i].moveToTrash();
      trashed++;
      Utilities.sleep(200);
    }
  }
  return trashed;
}

function collectProjectThreadIdsFromSheet_() {
  const ids = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  if (!mainSheet) {
    return ids;
  }
  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) {
    return ids;
  }
  const col = CONFIG.MAIN_THREAD_ID_COL;
  const values = mainSheet.getRange(2, col, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const id = (values[i][0] || '').toString().trim();
    if (id) {
      ids[id] = true;
    }
  }
  return ids;
}

function countProjectThreadsFromSheetIds_() {
  const ids = collectProjectThreadIdsFromSheet_();
  let n = 0;
  for (const k in ids) {
    if (ids.hasOwnProperty(k)) {
      n++;
    }
  }
  return n;
}

function countProjectThreadIdCells_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  if (!mainSheet || mainSheet.getLastRow() < 2) {
    return 0;
  }
  const values = mainSheet.getRange(2, CONFIG.MAIN_THREAD_ID_COL, mainSheet.getLastRow() - 1, 1).getValues();
  let cells = 0;
  for (let i = 0; i < values.length; i++) {
    if ((values[i][0] || '').toString().trim()) {
      cells++;
    }
  }
  return cells;
}

function trashProjectThreadsFromSheetIds_() {
  const ids = collectProjectThreadIdsFromSheet_();
  let trashed = 0;
  for (const threadId in ids) {
    if (!ids.hasOwnProperty(threadId)) {
      continue;
    }
    try {
      const thread = GmailApp.getThreadById(threadId);
      if (thread) {
        thread.moveToTrash();
        trashed++;
        Utilities.sleep(200);
      }
    } catch (e) {
      Logger.log('Could not trash project thread ' + threadId + ': ' + e);
    }
  }
  return trashed;
}

function clearProjectThreadIdColumn_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  if (!mainSheet) {
    return 0;
  }
  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }
  const range = mainSheet.getRange(2, CONFIG.MAIN_THREAD_ID_COL, lastRow - 1, 1);
  const before = range.getValues();
  let cleared = 0;
  for (let i = 0; i < before.length; i++) {
    if ((before[i][0] || '').toString().trim()) {
      cleared++;
    }
  }
  range.clearContent();
  return cleared;
}

function countPendingChangeRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pendingSheet = ss.getSheetByName(getPendingSheetName_());
  if (!pendingSheet) {
    return 0;
  }
  const n = pendingSheet.getLastRow();
  return n > 1 ? n - 1 : 0;
}

function clearPendingChangesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pendingSheet = ss.getSheetByName(getPendingSheetName_());
  if (!pendingSheet) {
    return 0;
  }
  const rowCount = pendingSheet.getLastRow();
  if (rowCount <= 1) {
    return 0;
  }
  const cleared = rowCount - 1;
  pendingSheet.deleteRows(2, cleared);
  return cleared;
}
