/**
 * Moves task rows from main/sub sheets to completed sheets.
 */
const CompletionMoveModule = (function () {
  'use strict';

  function mapSubRowToCompletedRow_(subRowValues) {
    const mainRow = new Array(SheetColumns.MAIN.LAST_COLUMN).fill('');

    mainRow[SheetColumns.MAIN.PROJECT_NAME - 1] = subRowValues[SheetColumns.SUB.PROJECT_NAME - 1];
    mainRow[SheetColumns.MAIN.TASK - 1] = subRowValues[SheetColumns.SUB.TASK - 1];
    mainRow[SheetColumns.MAIN.PRIORITY - 1] = subRowValues[SheetColumns.SUB.PRIORITY - 1];
    mainRow[SheetColumns.MAIN.ASSIGN - 1] = subRowValues[SheetColumns.SUB.ASSIGN - 1];
    mainRow[SheetColumns.MAIN.START_DATE - 1] = subRowValues[SheetColumns.SUB.START_DATE - 1];
    mainRow[SheetColumns.MAIN.END_DATE - 1] = subRowValues[SheetColumns.SUB.END_DATE - 1];
    mainRow[SheetColumns.MAIN.AS_PER_WO - 1] = subRowValues[SheetColumns.SUB.AS_PER_WO - 1];
    mainRow[SheetColumns.MAIN.STATUS - 1] = subRowValues[SheetColumns.SUB.STATUS - 1];
    mainRow[SheetColumns.MAIN.ENV - 1] = subRowValues[SheetColumns.SUB.ENV - 1];
    mainRow[SheetColumns.MAIN.ESTIMATED_EFFORT - 1] = subRowValues[SheetColumns.SUB.ESTIMATED_EFFORT - 1];
    mainRow[SheetColumns.MAIN.ACTUAL_EFFORT - 1] = subRowValues[SheetColumns.SUB.ACTUAL_EFFORT - 1];
    mainRow[SheetColumns.MAIN.INFORMED - 1] = subRowValues[SheetColumns.SUB.INFORMED - 1];
    mainRow[SheetColumns.MAIN.SUPPORT_BLOCKER - 1] = subRowValues[SheetColumns.SUB.SUPPORT_BLOCKER - 1];
    mainRow[SheetColumns.MAIN.PAGE_URL - 1] = subRowValues[SheetColumns.SUB.PAGE_URL - 1];
    mainRow[SheetColumns.MAIN.FUNCTIONALITY - 1] = subRowValues[SheetColumns.SUB.FUNCTIONALITY - 1];
    mainRow[SheetColumns.MAIN.MAIL_SUBJECT - 1] = subRowValues[SheetColumns.SUB.MAIL_SUBJECT - 1];
    mainRow[SheetColumns.MAIN.INTERNAL_MAIL_SUBJECT - 1] =
      subRowValues[SheetColumns.SUB.INTERNAL_MAIL_SUBJECT - 1];
    mainRow[SheetColumns.MAIN.GMAIL_THREAD_ID - 1] = subRowValues[SheetColumns.SUB.GMAIL_THREAD_ID - 1];
    mainRow[SheetColumns.MAIN.GOOGLE_TASK_ID - 1] = subRowValues[SheetColumns.SUB.GOOGLE_TASK_ID - 1];

    const metaParts = [];
    const mainTaskName = subRowValues[SheetColumns.SUB.MAIN_TASK_NAME - 1];
    const ticketId = subRowValues[SheetColumns.SUB.TICKET_ID - 1];
    if (mainTaskName) {
      metaParts.push('Main Task: ' + mainTaskName);
    }
    if (ticketId) {
      metaParts.push('Ticket Id: ' + ticketId);
    }

    const existingHistory = subRowValues[SheetColumns.SUB.LATEST_UPDATE - 1];
    if (metaParts.length > 0) {
      mainRow[SheetColumns.MAIN.LATEST_UPDATE - 1] = existingHistory
        ? metaParts.join(' | ') + ' | ' + existingHistory
        : metaParts.join(' | ');
    } else {
      mainRow[SheetColumns.MAIN.LATEST_UPDATE - 1] = existingHistory;
    }

    return mainRow;
  }

  function findSourceRowBySignature_(sourceSheet, sheetRole, rowSignature) {
    const lastRow = SpreadsheetUtils.getLastDataRow(
      sourceSheet,
      sheetRole === 'sub' ? SheetColumns.SUB.MAIN_TASK_NAME : SheetColumns.MAIN.TASK
    );
    if (lastRow < AppConfig.getDataStartRow()) {
      return -1;
    }

    for (var rowIndex = AppConfig.getDataStartRow(); rowIndex <= lastRow; rowIndex++) {
      const signature = CompletionQueueModule.buildRowSignature(sourceSheet, rowIndex, sheetRole);
      if (SpreadsheetUtils.normalizeText(signature) === SpreadsheetUtils.normalizeText(rowSignature)) {
        return rowIndex;
      }
    }
    return -1;
  }

  function appendCompletedRow_(completedSheet, rowValues) {
    const nextRow = Math.max(completedSheet.getLastRow() + 1, AppConfig.getDataStartRow());
    completedSheet.getRange(nextRow, 1, 1, SheetColumns.MAIN.LAST_COLUMN).setValues([rowValues]);
    return nextRow;
  }

  function moveRowToCompleted_(sourceSheet, sourceRowIndex, sheetRole, completedSheetName) {
    if (!sourceSheet || sourceRowIndex < AppConfig.getDataStartRow()) {
      return { success: false, error: 'Invalid source row' };
    }

    const spreadsheet = sourceSheet.getParent();
    const completedSheet = SpreadsheetUtils.getSheetByName(spreadsheet, completedSheetName);
    if (!completedSheet) {
      console.error('CompletionMoveModule.moveRowToCompleted_: completed sheet not found', {
        completedSheetName: completedSheetName,
      });
      return { success: false, error: 'Completed sheet not found' };
    }

    try {
      var rowValues;
      if (sheetRole === 'sub') {
        const subValues = SpreadsheetUtils.getRowValues(
          sourceSheet,
          sourceRowIndex,
          SheetColumns.SUB.LAST_COLUMN
        );
        rowValues = mapSubRowToCompletedRow_(subValues);
      } else {
        rowValues = SpreadsheetUtils.getRowValues(
          sourceSheet,
          sourceRowIndex,
          SheetColumns.MAIN.LAST_COLUMN
        );
      }

      const statusValue = rowValues[SheetColumns.MAIN.STATUS - 1];
      if (!TaskStatusModule.isCompletionStatus(statusValue)) {
        return { success: false, skipped: true, reason: 'Status is no longer a completion status' };
      }

      const completedRow = appendCompletedRow_(completedSheet, rowValues);
      sourceSheet.deleteRow(sourceRowIndex);

      console.log('CompletionMoveModule: row moved to completed sheet', {
        sourceSheet: sourceSheet.getName(),
        sourceRowIndex: sourceRowIndex,
        completedSheet: completedSheetName,
        completedRow: completedRow,
        status: statusValue,
      });

      return {
        success: true,
        sourceSheet: sourceSheet.getName(),
        completedSheet: completedSheetName,
        completedRow: completedRow,
        status: statusValue,
      };
    } catch (err) {
      console.error('CompletionMoveModule.moveRowToCompleted_: move failed', {
        sourceSheet: sourceSheet.getName(),
        sourceRowIndex: sourceRowIndex,
        completedSheetName: completedSheetName,
        error: err,
      });
      return { success: false, error: String(err) };
    }
  }

  function processQueueEntry_(entry) {
    if (!entry || !entry.sourceSheet || !entry.completedSheet) {
      return { success: false, error: 'Invalid queue entry' };
    }

    const spreadsheet = AppConfig.getSpreadsheet();
    const sourceSheet = SpreadsheetUtils.getSheetByName(spreadsheet, entry.sourceSheet);
    if (!sourceSheet) {
      return { success: false, error: 'Source sheet not found' };
    }

    const sourceRowIndex = findSourceRowBySignature_(sourceSheet, entry.sheetRole, entry.rowSignature);
    if (sourceRowIndex < AppConfig.getDataStartRow()) {
      return { success: false, skipped: true, reason: 'Source row no longer found' };
    }

    return moveRowToCompleted_(sourceSheet, sourceRowIndex, entry.sheetRole, entry.completedSheet);
  }

  function processDueEntries_() {
    const dueEntries = CompletionQueueModule.getDueEntries();
    const results = [];

    dueEntries
      .sort(function (a, b) {
        return b.queueRowIndex - a.queueRowIndex;
      })
      .forEach(function (entry) {
        try {
          const result = processQueueEntry_(entry);
          results.push(result);
          if (result && result.success) {
            CompletionQueueModule.removeQueueEntry(entry.queueRowIndex);
          } else if (result && result.skipped) {
            CompletionQueueModule.removeQueueEntry(entry.queueRowIndex);
          }
        } catch (err) {
          console.error('CompletionMoveModule.processDueEntries_: entry failed', {
            entry: entry,
            error: err,
          });
        }
      });

    return {
      processedCount: results.length,
      movedCount: results.filter(function (result) {
        return result && result.success;
      }).length,
      results: results,
    };
  }

  return {
    moveRowToCompleted: moveRowToCompleted_,
    processDueEntries: processDueEntries_,
    mapSubRowToCompletedRow: mapSubRowToCompletedRow_,
  };
})();
