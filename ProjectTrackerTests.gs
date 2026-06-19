/**
 * PROJECT TRACKER — TESTS & DIAGNOSTICS
 * Manual test runners and orphan sub-task diagnostics.
 * Full unified tests: TestAllTrackers.gs
 * Google Tasks sync code: ProjectTrackerGoogleTasks.gs
 *
 * ONE-CLICK GOOGLE TASKS TEST (open spreadsheet first):
 *   runFullGoogleTasksTestSuite()     — sequential steps only (lock + one-at-a-time)
 *   testGoogleTasksFullPipeline()     — alias
 */
function diagnoseOrphanSubRows() {
  const subSheet = getSubTrackerSheet_();
  if (!subSheet) {
    Logger.log('Sub sheet not found: ' + getSubSheetName_());
    return;
  }

  const mainSheet = getMainTrackerSheet_();
  const mainData = mainSheet ? mainSheet.getDataRange().getValues() : [];
  const mainLookup = buildMainTaskNameIndex_(mainData);

  const lastRow = subSheet.getLastRow();
  let orphanCount = 0;

  for (let r = lastRow; r >= 2; r--) {
    const lastCol = Math.max(subSheet.getLastColumn(), CONFIG.SUB_GOOGLE_TASK_ID_COL);
    const row = subSheet.getRange(r, 1, 1, lastCol).getValues()[0];
    const linkedMain = (row[CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
    const subTaskName = (row[CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();
    const taskId = (row[CONFIG.SUB_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();

    if (!subTaskName && !linkedMain) {
      continue;
    }

    if (linkedMain && mainLookup[linkedMain]) {
      continue;
    }

    orphanCount++;
    Logger.log(
      'ROW ' + r + ' — ORPHAN (no matching main task in Dev Tracker column C)' +
      '\n  Sub task (col D): ' + (subTaskName || '(empty)') +
      '\n  Linked main (col B): "' + linkedMain + '"' +
      '\n  Column V Task ID: ' + (taskId || '(empty)') +
      (linkedMain.indexOf('\t') !== -1 || linkedMain.length > 120
        ? '\n  WARNING: column B looks like multiple fields merged — it must be ONLY the exact Dev Tracker task name.'
        : '')
    );
  }

  Logger.log('Diagnosis complete. Orphan sub rows: ' + orphanCount);
  return orphanCount;
}

/** Alias for diagnoseOrphanSubRows(). */
function diagnoseOrphanSubRow() {
  diagnoseOrphanSubRows();
}

/**
 * Removes Google Tasks for sub rows with no valid main-task link (column B ≠ any Dev Tracker column C).
 * Clears column V and deletes top-level orphans by sub-task title across all status lists.
 * Safe to run anytime; does not touch rows that have a valid parent link.
 */
function removeOrphanSubtasksWithNoMainTask() {
  if (!isGoogleTasksApiEnabled_()) {
    Logger.log('Google Tasks API not enabled — add Services (+) > Google Tasks API first.');
    return;
  }

  const mainSheet = getMainTrackerSheet_();
  const subSheet = getSubTrackerSheet_();
  if (!mainSheet || !subSheet) {
    Logger.log('Sheets not found');
    return;
  }

  refreshGoogleTaskListCache_();
  taskIdToListIdCache_ = {};

  const mainData = mainSheet.getDataRange().getValues();
  const mainLookup = buildMainTaskNameIndex_(mainData);
  const subLast = subSheet.getLastRow();

  let cleaned = 0;
  let skipped = 0;

  for (let r = subLast; r >= 2; r--) {
    const lastCol = Math.max(subSheet.getLastColumn(), CONFIG.SUB_GOOGLE_TASK_ID_COL);
    const row = subSheet.getRange(r, 1, 1, lastCol).getValues()[0];
    const linkedMain = (row[CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
    const subTaskName = (row[CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();

    if (!subTaskName && !linkedMain) {
      skipped++;
      continue;
    }

    if (linkedMain && mainLookup[linkedMain]) {
      skipped++;
      continue;
    }

    const hadTaskId = !!getGoogleSubtaskIdForSubRow_(subSheet, r);
    Logger.log(
      'Row ' + r + ' (' + (subTaskName || 'no name') + '): no matching main for "' + linkedMain + '"' +
      (hadTaskId ? ' — removing stale Google Task' : ' — scanning for title orphans')
    );

    cleanupGoogleSubtaskOnSyncSkip_(subSheet, r, row);
    cleaned++;
    Utilities.sleep(200);
  }

  Logger.log('=== Cleanup complete ===');
  Logger.log('Orphan rows cleaned: ' + cleaned);
  Logger.log('Skipped (valid link or empty row): ' + skipped);
}

/**
 * Repairs subtasks that exist in Google Tasks as top-level orphans (wrong nest).
 * Only processes rows that already have a Google Subtask ID in column V.
 * Run from Apps Script after syncExistingTasksToGoogle() if some subtasks still appear as main tasks.
 */
function testDailySummary() {
  sendDailySummary();
  Logger.log('Test summary sent. Check your email: ' + getProjectMyEmail_());
}

/**
 * TEST FUNCTION: Run this to test instant email for a specific row
 */
function testInstantEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  
  // Test with row 2 (first data row)
  handleMainUpdate(mainSheet, 2);
  Logger.log('Test email sent for row 2. Check your email: ' + getProjectMyEmail_());
}

/**
 * TEST FUNCTION: Test nested label functionality
 */
function testNestedLabels() {
  Logger.log('=== TESTING LABEL MANAGEMENT ===');
  
  // First create all labels
  Logger.log('Step 1: Creating all labels...');
  createAllStatusLabels();
  
  // Test with a specific row
  Logger.log('Step 2: Testing with row 2...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  handleMainUpdate(mainSheet, 2);
  
  // Run the hourly batch
  Logger.log('Step 3: Running hourly batch...');
  Utilities.sleep(2000);
  sendHourlyBatch();
  
  Logger.log('=== TEST COMPLETE ===');
  Logger.log('Check Gmail - thread should have:');
  Logger.log('- Parent label: A Assessment');
  Logger.log('- Child label: A Assessment/[Current Status from row 2]');
}

/**
 * TEST FUNCTION: Test label change when status updates
 */
function testLabelChange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  
  // Get the thread ID from row 2
  const threadId = getProjectThreadIdForRow_(mainSheet, 2);
  
  if (!threadId || threadId.toString().trim() === "") {
    Logger.log('No thread ID found in row 2. Run testNestedLabels() first to create a thread.');
    return;
  }
  
  Logger.log('Testing label change on existing thread: ' + threadId);
  
  // Change the label to "In Progress"
  manageThreadLabels(threadId, "In Progress");
  Logger.log('Changed to: In Progress');
  
  Utilities.sleep(2000);
  
  // Change the label to "Completed"
  manageThreadLabels(threadId, "Completed");
  Logger.log('Changed to: Completed');
  
  Logger.log('Test complete! Check Gmail - should see only "Completed" label now.');
}

/**
 * Clears Dev Tracker Sub column W (Main Google Task ID) on all data rows — optional test prep.
 */
function clearSubParentGoogleTaskIdColumnForTest_() {
  const subSheet = getSubTrackerSheet_();
  if (!subSheet) {
    return;
  }
  const last = subSheet.getLastRow();
  if (last < 2) {
    return;
  }
  const parentCol = findSubParentGoogleTaskIdColumn_(subSheet);
  subSheet.getRange(2, parentCol, last, parentCol).clearContent();
  Logger.log('Cleared Dev Tracker Sub Main Google Task ID column (rows 2–' + last + ').');
}

/**
 * Logs main V, sub V, and sub W for one row after sync.
 * @param {number} [row=2]
 */
function logSampleRowGoogleTaskIds_(row) {
  row = row || 2;
  const mainSheet = getMainTrackerSheet_();
  const subSheet = getSubTrackerSheet_();

  if (!mainSheet) {
    Logger.log('Dev Tracker sheet not found.');
    return;
  }

  const mainName = (mainSheet.getRange(row, CONFIG.MAIN_TASK_NAME_COL).getValue() || '').toString().trim();
  const mainV = getGoogleTaskIdForMainRow_(mainSheet, row);
  Logger.log(
    'Row ' + row + ' MAIN "' + (mainName || '(no name)') + '" → Dev Tracker column V: ' + (mainV || '(empty)')
  );

  if (!subSheet || row > subSheet.getLastRow()) {
    Logger.log('Row ' + row + ' — no Dev Tracker Sub row to check.');
    return;
  }

  const subV = getGoogleSubtaskIdForSubRow_(subSheet, row);
  const subW = getSubParentGoogleTaskIdForSubRow_(subSheet, row);
  const nestStatus = (subSheet.getRange(row, findSubNestStatusColumn_(subSheet)).getValue() || '').toString().trim();
  const subName = (subSheet.getRange(row, CONFIG.SUB_TASK_NAME_COL).getValue() || '').toString().trim();
  Logger.log(
    'Row ' + row + ' SUB "' + (subName || '(no name)') + '" → column V: ' + (subV || '(empty)') +
    ', column W: ' + (subW || '(empty)') +
    ', column X (nest): ' + (nestStatus || '(empty)')
  );

  if (mainV && subV && mainV === subV) {
    Logger.log('WARNING row ' + row + ': sub column V equals main column V — sub should have its own ID.');
  } else if (mainV && subW && mainV !== subW) {
    Logger.log('WARNING row ' + row + ': column W does not match main column V — run backfillSubParentGoogleTaskIdsIfEmpty().');
  } else if (mainV && subW && mainV === subW) {
    Logger.log('OK row ' + row + ': column W matches main column V' + (subV ? '; sub column V is the nested task id.' : '.'));
  }
}

/**
 * Runs one test-suite step synchronously — the next step does not start until this returns.
 * @param {number} stepNum
 * @param {number} totalSteps
 * @param {string} label
 * @param {function(): *} fn
 * @returns {*}
 */
function runTestSuiteStep_(stepNum, totalSteps, label, fn) {
  Logger.log('');
  Logger.log('============================================================');
  Logger.log('STEP ' + stepNum + '/' + totalSteps + ' START: ' + label);
  Logger.log('(sequential — next step waits until this one finishes)');
  Logger.log('============================================================');

  const t0 = Date.now();
  let result;
  try {
    result = fn();
  } catch (err) {
    Logger.log('STEP ' + stepNum + ' FAILED: ' + err.toString());
    throw err;
  }

  try {
    SpreadsheetApp.flush();
  } catch (flushErr) {
    // non-fatal
  }

  Utilities.sleep(300);
  Logger.log(
    'STEP ' + stepNum + '/' + totalSteps + ' DONE (' + (Date.now() - t0) + ' ms): ' + label
  );
  return result;
}

/**
 * Runs the full Google Tasks test pipeline in order — strictly one step at a time.
 *
 * Uses a script lock so only ONE test run can execute at once.
 * onEdit Google Tasks sync is paused while this runs (no parallel sync).
 *
 * Open the spreadsheet in your browser before running.
 *
 * Options (all optional):
 *   fullReset: true    — delete tracker-list Google Tasks + clear column V (default true)
 *   clearColumnW: false — also clear column W before backfill (default false)
 *   sampleRow: 2      — row to inspect at the end
 *   maxSyncRuns: 50    — max bulk sync loops when sheet is large (5 min pause per run)
 *
 * Example: runFullGoogleTasksTestSuite({ fullReset: true, clearColumnW: true, sampleRow: 2 })
 */
function runFullGoogleTasksTestSuite(options) {
  options = options || {};
  const fullReset = options.fullReset !== false;
  const clearColumnW = options.clearColumnW === true;
  const sampleRow = options.sampleRow || 2;
  const maxSyncRuns = options.maxSyncRuns || 50;
  const totalSteps = fullReset ? 9 : 8;
  const started = new Date();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log(
      'ABORT — another Google Tasks job is already running. ' +
      'Wait for it to finish, then run runFullGoogleTasksTestSuite() again.'
    );
    return;
  }

  Logger.log('=== GOOGLE TASKS FULL TEST SUITE START (sequential, one step at a time) ===');
  Logger.log('Started: ' + started.toLocaleString());
  Logger.log('Do not run this function again until you see TEST SUITE COMPLETE in the log.');

  let orphans = 0;
  let aborted = false;

  try {
    setGoogleTasksTestSuiteRunning_(true);

    runTestSuiteStep_(1, totalSteps, 'Authorize Google Tasks API', function () {
      if (!authorizeGoogleTasksAccess()) {
        throw new Error('Google Tasks API not enabled — add Services (+) > Google Tasks API');
      }
    });

    runTestSuiteStep_(2, totalSteps, 'Clear bulk sync resume state', function () {
      clearGoogleTasksSyncResumeState_();
    });

    runTestSuiteStep_(3, totalSteps, 'Verify Google Task lists (CONFIG.STATUS_LABELS)', function () {
      verifyGoogleTaskLists();
    });

    if (fullReset) {
      runTestSuiteStep_(4, totalSteps, 'Delete tracker-list tasks + clear column V', function () {
        const deleted = deleteAllTasksFromTrackerStatusLists_();
        clearAllGoogleTaskIdColumns_();
        Logger.log('Deleted ' + deleted + ' Google Task(s) from tracker status lists.');
        if (clearColumnW) {
          clearSubParentGoogleTaskIdColumnForTest_();
        }
      });
    }

    const stepBackfill = fullReset ? 5 : 4;
    runTestSuiteStep_(stepBackfill, totalSteps, 'Backfill / correct column W from Dev Tracker column V', function () {
      backfillSubParentGoogleTaskIdsIfEmpty();
    });

    const stepSync = fullReset ? 6 : 5;
    runTestSuiteStep_(stepSync, totalSteps, 'Bulk sync — mains first, then subs (batched)', function () {
      return runBulkGoogleTasksSyncUntilComplete_(1);
    });

    if (isGoogleTasksBulkSyncPaused_()) {
      const status = getGoogleTasksBulkSyncStatus();
      Logger.log(
        'Bulk sync paused after first batch (normal for 2000+ subs). Phase=' + status.phase +
        ' row=' + status.resumeRow + ' subTotal=' + status.subTotal
      );
      if (CONFIG.GOOGLE_TASKS_AUTO_CONTINUE_BULK !== false) {
        Logger.log('Auto-continue is ON — next batch runs in ~1 min via continueGoogleTasksBulkSync_().');
        Logger.log('Check progress anytime: getGoogleTasksBulkSyncStatus() or getGoogleTasksLastSyncLog()');
      } else {
        Logger.log('Run syncExistingTasksToGoogle() repeatedly until getGoogleTasksBulkSyncStatus().paused is false.');
        aborted = true;
        return;
      }
    } else {
      Logger.log('Bulk sync finished in one batch (all rows synced).');
    }

    const stepSummary = fullReset ? 7 : 6;
    runTestSuiteStep_(stepSummary, totalSteps, 'Last sync summary', function () {
      Logger.log(getGoogleTasksLastSyncLog());
    });

    const stepDiagnose = fullReset ? 8 : 7;
    orphans = runTestSuiteStep_(stepDiagnose, totalSteps, 'Diagnose orphan sub rows', function () {
      return diagnoseOrphanSubRows();
    });

    const stepNest = fullReset ? 9 : 8;
    if (isGoogleTasksBulkSyncPaused_()) {
      Logger.log(
        'STEP ' + stepNest + '/' + totalSteps + ' SKIPPED: refresh column X — bulk sync still paused. ' +
        'After sync finishes, run refreshSubNestStatusColumn().'
      );
    } else {
      runTestSuiteStep_(stepNest, totalSteps, 'Refresh column X (Google Nest Status)', function () {
        const stats = refreshSubNestStatusColumn({ batchSize: 300 });
        Logger.log(
          'Column X batch: OK=' + stats.ok + ' ORPHAN=' + stats.orphan +
          ' WRONG=' + stats.wrong + ' other=' + stats.other
        );
        if (stats.paused) {
          Logger.log('Column X paused — run refreshSubNestStatusColumn({ startRow: ... }) from log to continue.');
        }
      });
    }

    Logger.log('');
    Logger.log('============================================================');
    Logger.log('STEP (extra): Sample row ' + sampleRow + ' check');
    Logger.log('============================================================');
    logSampleRowGoogleTaskIds_(sampleRow);

  } catch (err) {
    Logger.log('TEST SUITE ERROR: ' + err.toString());
    aborted = true;
  } finally {
    setGoogleTasksTestSuiteRunning_(false);
    setGoogleTasksBulkSyncInProgress_(false);
    lock.releaseLock();
  }

  if (aborted) {
    Logger.log('=== GOOGLE TASKS TEST SUITE ABORTED ===');
    return;
  }

  Logger.log('');
  Logger.log('============================================================');
  Logger.log('GOOGLE TASKS TEST SUITE COMPLETE');
  Logger.log('Started:  ' + started.toLocaleString());
  Logger.log('Finished: ' + new Date().toLocaleString());
  Logger.log('Orphan sub rows: ' + orphans + (orphans === 0 ? ' (OK)' : ' (review log above)'));
  Logger.log('Filter Dev Tracker Sub column X: OK = nested; ORPHAN / WRONG PARENT = needs fix.');
  Logger.log('Next: open Google Tasks app — mains top-level, subs nested under mains.');
  Logger.log('============================================================');
}

/** Alias — run from Apps Script: testGoogleTasksFullPipeline() */
function testGoogleTasksFullPipeline() {
  runFullGoogleTasksTestSuite();
}