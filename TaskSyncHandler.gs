/**
 * Installable onEdit handler — main/sub task linking, field sync, and effort rollup.
 * All team sheets resolved dynamically from the Matching Sheet.
 */
const TaskSyncHandler = (function () {
  'use strict';

  const LOCK_TIMEOUT_MS = 30000;

  const MAIN_TRIGGER_COLUMNS = [SheetColumns.MAIN.TASK];
  const SUB_TRIGGER_COLUMNS = [
    SheetColumns.SUB.MAIN_TASK_NAME,
    SheetColumns.SUB.ESTIMATED_EFFORT,
    SheetColumns.SUB.ACTUAL_EFFORT,
    SheetColumns.SUB.PROJECT_NAME,
  ];

  function rangeTouchesColumns_(startCol, endCol, columns) {
    for (var i = 0; i < columns.length; i++) {
      if (SpreadsheetUtils.columnRangeOverlaps(startCol, endCol, columns[i])) {
        return true;
      }
    }
    return false;
  }

  function shouldProcessEdit_(e) {
    if (!e || !e.range) {
      return false;
    }

    const sheet = e.range.getSheet();
    if (!sheet) {
      return false;
    }

    const sheetInfo = MatchingSheetModule.getSheetRole(sheet.getName());
    if (!sheetInfo) {
      return false;
    }

    if (sheetInfo.role !== 'main' && sheetInfo.role !== 'sub') {
      return false;
    }

    const startCol = e.range.getColumn();
    const endCol = startCol + e.range.getNumColumns() - 1;
    const triggerColumns =
      sheetInfo.role === 'main' ? MAIN_TRIGGER_COLUMNS : SUB_TRIGGER_COLUMNS;

    if (!rangeTouchesColumns_(startCol, endCol, triggerColumns)) {
      return false;
    }

    if (e.range.getRow() <= AppConfig.getHeaderRow()) {
      return false;
    }

    return true;
  }

  function processMainEdit_(sheet, mapping, startRow, numRows) {
    const results = [];

    for (var i = 0; i < numRows; i++) {
      const rowIndex = startRow + i;
      if (rowIndex <= AppConfig.getHeaderRow()) {
        continue;
      }
      try {
        const syncResult = SubTaskSyncModule.syncSubTasksForMainRow(sheet, rowIndex, mapping);
        results.push({ type: 'main', rowIndex: rowIndex, sync: syncResult });
      } catch (err) {
        console.error('TaskSyncHandler.processMainEdit_: row failed', {
          sheet: sheet.getName(),
          rowIndex: rowIndex,
          error: err,
        });
      }
    }

    return results;
  }

  function processSubEdit_(sheet, mapping, startRow, numRows) {
    const results = [];

    for (var j = 0; j < numRows; j++) {
      const rowIndex = startRow + j;
      if (rowIndex <= AppConfig.getHeaderRow()) {
        continue;
      }
      try {
        const syncResult = SubTaskSyncModule.syncSingleSubRowFromMain(sheet, rowIndex, mapping);
        results.push({ type: 'sub', rowIndex: rowIndex, sync: syncResult });
      } catch (err) {
        console.error('TaskSyncHandler.processSubEdit_: row failed', {
          sheet: sheet.getName(),
          rowIndex: rowIndex,
          error: err,
        });
      }
    }

    return results;
  }

  function processEdit_(e) {
    const sheet = e.range.getSheet();
    const sheetInfo = MatchingSheetModule.getSheetRole(sheet.getName());
    if (!sheetInfo || (sheetInfo.role !== 'main' && sheetInfo.role !== 'sub')) {
      return [];
    }

    MatchingSheetModule.refreshMappings();

    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();

    if (sheetInfo.role === 'main') {
      return processMainEdit_(sheet, sheetInfo.mapping, startRow, numRows);
    }

    return processSubEdit_(sheet, sheetInfo.mapping, startRow, numRows);
  }

  return {
    shouldProcessEdit: shouldProcessEdit_,
    processEdit: processEdit_,
    LOCK_TIMEOUT_MS: LOCK_TIMEOUT_MS,
  };
})();

/**
 * Entry point for installable onEdit trigger. Must remain a top-level function.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function handleSpreadsheetEdit_(e) {
  if (AppConfig.isTesting()) {
    return handleSpreadsheetEditTest_(e);
  }
  return handleSpreadsheetEditProduction_(e);
}

/**
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function handleSpreadsheetEditProduction_(e) {
  const lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    hasLock = lock.tryLock(TaskSyncHandler.LOCK_TIMEOUT_MS || 30000);
    if (!hasLock) {
      console.error('handleSpreadsheetEditProduction_: could not acquire lock');
      return;
    }

    if (CompletionTaskHandler.shouldProcessEdit(e)) {
      CompletionTaskHandler.processEdit(e);
    }

    if (TaskSyncHandler.shouldProcessEdit(e)) {
      TaskSyncHandler.processEdit(e);
    }

    if (NotificationHandler.shouldProcessEdit(e)) {
      NotificationHandler.processEdit(e);
    }
  } catch (err) {
    console.error('handleSpreadsheetEditProduction_: handler failed', { error: err });
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function handleSpreadsheetEditTest_(e) {
  handleSpreadsheetEditProduction_(e);
}

/**
 * Manual run: sync sub tasks for a main task row (sheet name from Matching Sheet).
 * @param {string} mainSheetName Main sheet tab name as listed in Matching Sheet
 * @param {number} rowIndex Row index on the main sheet (must be >= 2)
 */
function runSubTaskSyncForRow(mainSheetName, rowIndex) {
  return SubTaskSyncModule.syncSubTasksForMainSheetRow(mainSheetName, rowIndex);
}

/**
 * Manual run: sync sub tasks and rollup effort by main task name.
 * @param {string} mainSheetName Main sheet tab name as listed in Matching Sheet
 * @param {string} mainTaskName Value from the Task column on the main sheet
 */
function runSubTaskSyncForTaskName(mainSheetName, mainTaskName) {
  const syncResult = SubTaskSyncModule.syncAllSubTasksForMainTaskName(mainSheetName, mainTaskName);
  const rollupResult = TaskEffortRollupModule.rollupEffortForMainSheetName(mainSheetName, mainTaskName);
  return { sync: syncResult, rollup: rollupResult };
}

/**
 * Manual run: rollup Estimated/Actual Effort from sub sheet to main sheet by task name.
 * @param {string} mainSheetName Main sheet tab name from Matching Sheet
 * @param {string} mainTaskName Task column value on the main sheet
 */
function runEffortRollupForTaskName(mainSheetName, mainTaskName) {
  return TaskEffortRollupModule.rollupEffortForMainSheetName(mainSheetName, mainTaskName);
}

/**
 * Manual run: sync and rollup for a sub-task row (sheet name from Matching Sheet).
 * @param {string} subSheetName Sub sheet tab name from Matching Sheet
 * @param {number} subRowIndex Row index on the sub sheet (must be >= 2)
 */
function runSubRowSyncAndRollup(subSheetName, subRowIndex) {
  return SubTaskSyncModule.syncSubRowForMainSheet(subSheetName, subRowIndex);
}

/** @deprecated Use runSubTaskSyncForRow(mainSheetName, rowIndex) */
function runUmbracoSubTaskSyncForRow(rowIndex) {
  const mappings = MatchingSheetModule.getMappings();
  var umbracoMapping = null;
  for (var i = 0; i < mappings.length; i++) {
    if (SpreadsheetUtils.normalizeText(mappings[i].mainSheet).indexOf('umbraco') >= 0) {
      umbracoMapping = mappings[i];
      break;
    }
  }
  if (!umbracoMapping) {
    return { success: false, error: 'No Umbraco mapping found on Matching Sheet' };
  }
  return runSubTaskSyncForRow(umbracoMapping.mainSheet, rowIndex);
}

/** @deprecated Use runSubTaskSyncForTaskName(mainSheetName, mainTaskName) */
function runUmbracoSubTaskSyncForTaskName(mainTaskName) {
  const mappings = MatchingSheetModule.getMappings();
  var umbracoMapping = null;
  for (var i = 0; i < mappings.length; i++) {
    if (SpreadsheetUtils.normalizeText(mappings[i].mainSheet).indexOf('umbraco') >= 0) {
      umbracoMapping = mappings[i];
      break;
    }
  }
  if (!umbracoMapping) {
    return { success: false, error: 'No Umbraco mapping found on Matching Sheet' };
  }
  return runSubTaskSyncForTaskName(umbracoMapping.mainSheet, mainTaskName);
}
