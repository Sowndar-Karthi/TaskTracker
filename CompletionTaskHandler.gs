/**
 * Handles status edits and scheduled completion moves.
 */
const CompletionTaskHandler = (function () {
  'use strict';

  function getStatusColumnForRole_(sheetRole) {
    return sheetRole === 'sub' ? SheetColumns.SUB.STATUS : SheetColumns.MAIN.STATUS;
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

    const statusCol = getStatusColumnForRole_(sheetInfo.role);
    const startCol = e.range.getColumn();
    const endCol = startCol + e.range.getNumColumns() - 1;

    if (!SpreadsheetUtils.columnRangeOverlaps(startCol, endCol, statusCol)) {
      return false;
    }

    if (e.range.getRow() <= AppConfig.getHeaderRow()) {
      return false;
    }

    return true;
  }

  function handleStatusChange_(sheet, rowIndex, sheetInfo) {
    const statusCol = getStatusColumnForRole_(sheetInfo.role);
    const rowValues = SpreadsheetUtils.getRowValues(
      sheet,
      rowIndex,
      sheetInfo.role === 'sub' ? SheetColumns.SUB.LAST_COLUMN : SheetColumns.MAIN.LAST_COLUMN
    );
    const statusValue = rowValues[statusCol - 1];
    const rowSignature = CompletionQueueModule.buildRowSignature(sheet, rowIndex, sheetInfo.role);

    if (TaskStatusModule.isCompletionStatus(statusValue)) {
      return CompletionQueueModule.scheduleRow(
        sheet,
        rowIndex,
        sheetInfo.role,
        sheetInfo.mapping,
        statusValue
      );
    }

    return CompletionQueueModule.cancelRow(sheet.getName(), rowSignature);
  }

  function processEdit_(e) {
    const sheet = e.range.getSheet();
    const sheetInfo = MatchingSheetModule.getSheetRole(sheet.getName());
    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();
    const results = [];

    CompletionQueueModule.ensureQueueSheet();

    for (var i = 0; i < numRows; i++) {
      const rowIndex = startRow + i;
      if (rowIndex <= AppConfig.getHeaderRow()) {
        continue;
      }
      try {
        results.push(handleStatusChange_(sheet, rowIndex, sheetInfo));
      } catch (err) {
        console.error('CompletionTaskHandler.processEdit_: row failed', {
          sheet: sheet.getName(),
          rowIndex: rowIndex,
          error: err,
        });
      }
    }

    return results;
  }

  return {
    shouldProcessEdit: shouldProcessEdit_,
    processEdit: processEdit_,
  };
})();

/**
 * Hourly time-driven trigger — moves rows queued for completion after 24 hours.
 */
function processScheduledCompletions_() {
  if (AppConfig.isTesting()) {
    return processScheduledCompletionsTest_();
  }
  return processScheduledCompletionsProduction_();
}

function processScheduledCompletionsProduction_() {
  const lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    hasLock = lock.tryLock(30000);
    if (!hasLock) {
      console.error('processScheduledCompletionsProduction_: could not acquire lock');
      return;
    }
    CompletionQueueModule.ensureQueueSheet();
    return CompletionMoveModule.processDueEntries();
  } catch (err) {
    console.error('processScheduledCompletionsProduction_: failed', { error: err });
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

function processScheduledCompletionsTest_() {
  return processScheduledCompletionsProduction_();
}

/**
 * Manual run: process all due completion queue entries immediately.
 */
function runProcessCompletionQueue() {
  return CompletionMoveModule.processDueEntries();
}

/**
 * Manual run: schedule a row for testing (uses normal 24h / test delay rules).
 * @param {string} sheetName
 * @param {number} rowIndex
 */
function runScheduleCompletionForRow(sheetName, rowIndex) {
  const spreadsheet = AppConfig.getSpreadsheet();
  const sheet = SpreadsheetUtils.getSheetByName(spreadsheet, sheetName);
  const sheetInfo = MatchingSheetModule.getSheetRole(sheetName);
  if (!sheet || !sheetInfo) {
    return { success: false, error: 'Sheet or mapping not found' };
  }
  const statusCol = sheetInfo.role === 'sub' ? SheetColumns.SUB.STATUS : SheetColumns.MAIN.STATUS;
  const statusValue = sheet.getRange(rowIndex, statusCol).getValue();
  return CompletionQueueModule.scheduleRow(sheet, rowIndex, sheetInfo.role, sheetInfo.mapping, statusValue);
}
