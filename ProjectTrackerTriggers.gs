/**
 * PROJECT TRACKER — TRIGGERS
 * Installable on-edit handler and time-driven trigger setup.
 */
function projectOnEditHandler(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (!sheet) {
      Logger.log('projectOnEditHandler: could not resolve edited sheet from event.');
      return;
    }

    const sheetName = sheet.getName();
    const col = e.range.getColumn();
    const row = e.range.getRow();
    
    if (row < 2) return; // Skip header row

    const googleSyncBlocked = isGoogleTasksSyncBlockedForOnEdit_();
    if (googleSyncBlocked) {
      Logger.log(
        '[GoogleTasks] onEdit row ' + row + ' — Google Tasks sync deferred (bulk/test suite running one step at a time)'
      );
    }

    if (isMainTrackerSheetName_(sheetName)) {
      const mainGoogleSyncCols = [CONFIG.MAIN_TASK_NAME_COL, CONFIG.MAIN_STATUS_COL].concat(CONFIG.MONITORED_COLS);
      if (!googleSyncBlocked && mainGoogleSyncCols.includes(col)) {
        syncGoogleTaskForMainRow(sheet, row);
      }

      if (col === CONFIG.MAIN_STATUS_COL) {
        updateLabelsForMainRow(sheet, row);
        if (!googleSyncBlocked && !CONFIG.GOOGLE_TASKS_LIGHT_MAIN_STATUS_EDIT) {
          syncGoogleSubtasksForMainRow_(sheet, row);
        } else if (googleSyncBlocked) {
          Logger.log(
            '[GoogleTasks:ROW] main Status edit row ' + row +
            ' — linked subs deferred until bulk/test suite finishes'
          );
        } else {
          Logger.log(
            '[GoogleTasks:ROW] main Status edit row ' + row +
            ' — linked subs deferred (GOOGLE_TASKS_LIGHT_MAIN_STATUS_EDIT=true); run bulk or edit sub rows'
          );
        }
        moveMainRowToCompletedTrackerIfNeeded_(sheet, row);
        // If moved to completed, do not queue/hourly-email this task anymore
        return;
      }

      if (CONFIG.MONITORED_COLS.includes(col)) {
        handleMainUpdate(sheet, row);
      }
    } else if (isSubTrackerSheetName_(sheetName)) {
      const subGoogleSyncCols = [CONFIG.SUB_LINK_COL, CONFIG.SUB_TASK_NAME_COL, CONFIG.SUB_STATUS_COL,
        findSubParentGoogleTaskIdColumn_(sheet)]
        .concat(CONFIG.SUB_MONITORED_COLS);
      if (!googleSyncBlocked && subGoogleSyncCols.includes(col)) {
        syncGoogleSubtaskForSubRow(sheet, row);
      }

      if (CONFIG.SUB_MONITORED_COLS.includes(col)) {
        handleSubUpdate(sheet, row);
      }
    }
  } catch (error) {
    Logger.log('projectOnEditHandler Error: ' + error.toString());
  }
}

/**
 * HELPER: Move main task row to Completed Tracker when status is Completed/Closed
 * Runs only on Status edits (column I).
 */
function setupHourlyTrigger() {
  // Remove any existing sendHourlyBatch triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendHourlyBatch') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create a new time-driven trigger using the configured interval
  const intervalHours = getBatchIntervalHours_();
  ScriptApp.newTrigger('sendHourlyBatch')
    .timeBased()
    .everyHours(intervalHours)
    .create();

  Logger.log('Batch trigger created for sendHourlyBatch — every ' + intervalHours + ' hour(s).');
}

/**
 * Daily safety net: fill empty column U from existing Gmail threads (read-only — never sends mail).
 * TestAllTrackers.gs has an extended version with THREAD_RECOVERY_CONFIG; this entry point
 * ensures the setupAllTriggers daily job always resolves even if only ProjectTracker.gs is reviewed.
 */
function setupAllTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Only installable triggers can be deleted — not simple triggers named onEdit.
  const handlers = [
    'projectOnEditHandler',
    'sendHourlyBatch',
    'sendDailySummary',
    'sendWeeklySummary',
    'autoRecoverProjectThreadIds'
  ];

  // Remove existing triggers for these handlers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (handlers.indexOf(handler) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Installable on-edit trigger (use projectOnEditHandler — not "onEdit", which also runs as a simple trigger)
  ScriptApp.newTrigger('projectOnEditHandler')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Batch trigger (interval from ProjectTrackerUserConfig.gs or script property)
  const intervalHours = getBatchIntervalHours_();
  ScriptApp.newTrigger('sendHourlyBatch')
    .timeBased()
    .everyHours(intervalHours)
    .create();

  // Daily summary triggers (6 AM and 6 PM)
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

  // Weekly summary trigger (Friday 6 PM)
  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(18)
    .everyWeeks(1)
    .create();

  // Daily safety net for orphaned Thread IDs (runs before the 6 AM summary).
  ScriptApp.newTrigger('autoRecoverProjectThreadIds')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();

  Logger.log(
    '✅ All triggers created: projectOnEditHandler, batch (every ' + intervalHours +
    ' hr), daily (6/18), weekly (Fri 18), auto-recovery (daily 5 AM).'
  );
}

/**
 * Task names (Dev Tracker column C) still on the main sheet. Sub-rows whose link
 * does not match are orphans after the parent moved to Completed Tracker.
 */
