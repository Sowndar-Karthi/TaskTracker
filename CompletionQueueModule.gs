/**
 * Queue for delayed row moves to completed task sheets.
 */
const CompletionQueueModule = (function () {
  'use strict';

  const QUEUE_HEADERS = [
    'Source Sheet',
    'Sheet Role',
    'Row Signature',
    'Status',
    'Scheduled At',
    'Completed Sheet',
  ];

  function getQueueSheet_(spreadsheet) {
    if (!spreadsheet) {
      return null;
    }

    const queueName = AppConfig.getCompletionQueueSheetName();
    var sheet = SpreadsheetUtils.getSheetByName(spreadsheet, queueName);
    if (sheet) {
      return sheet;
    }

    try {
      sheet = spreadsheet.insertSheet(queueName);
      sheet.getRange(1, 1, 1, QUEUE_HEADERS.length).setValues([QUEUE_HEADERS]);
      sheet.setFrozenRows(1);
      spreadsheet.hideSheet(sheet);
      return sheet;
    } catch (err) {
      ErrorLogModule.error('CompletionQueueModule', 'getQueueSheet_', 'Create queue sheet failed', {
        error: err,
      });
      return null;
    }
  }

  function buildRowSignature_(sheet, rowIndex, sheetRole) {
    if (!sheet || rowIndex < AppConfig.getDataStartRow()) {
      return '';
    }

    if (sheetRole === 'sub') {
      const values = SpreadsheetUtils.getRowValues(sheet, rowIndex, SheetColumns.SUB.LAST_COLUMN);
      return [
        values[SheetColumns.SUB.PROJECT_NAME - 1],
        values[SheetColumns.SUB.MAIN_TASK_NAME - 1],
        values[SheetColumns.SUB.TICKET_ID - 1],
        values[SheetColumns.SUB.TASK - 1],
      ]
        .map(function (part) {
          return SpreadsheetUtils.normalizeText(part);
        })
        .join('||');
    }

    const values = SpreadsheetUtils.getRowValues(sheet, rowIndex, SheetColumns.MAIN.LAST_COLUMN);
    return [
      values[SheetColumns.MAIN.PROJECT_NAME - 1],
      values[SheetColumns.MAIN.TASK - 1],
      values[SheetColumns.MAIN.ASSIGN - 1],
      values[SheetColumns.MAIN.START_DATE - 1],
    ]
      .map(function (part) {
        return SpreadsheetUtils.normalizeText(part);
      })
      .join('||');
  }

  function findQueueRowIndex_(queueSheet, sourceSheetName, rowSignature) {
    const lastRow = SpreadsheetUtils.getLastDataRow(queueSheet, 1);
    if (lastRow < AppConfig.getDataStartRow()) {
      return -1;
    }

    const rowCount = lastRow - AppConfig.getHeaderRow();
    const values = queueSheet.getRange(AppConfig.getDataStartRow(), 1, rowCount, QUEUE_HEADERS.length).getValues();
    const normalizedSource = SpreadsheetUtils.normalizeText(sourceSheetName);
    const normalizedSignature = SpreadsheetUtils.normalizeText(rowSignature);

    for (var i = 0; i < values.length; i++) {
      if (
        SpreadsheetUtils.normalizeText(values[i][0]) === normalizedSource &&
        SpreadsheetUtils.normalizeText(values[i][2]) === normalizedSignature
      ) {
        return AppConfig.getDataStartRow() + i;
      }
    }
    return -1;
  }

  function scheduleRow_(sourceSheet, rowIndex, sheetRole, mapping, statusValue) {
    if (!sourceSheet || !mapping || !mapping.completedSheet) {
      return { success: false, error: 'Invalid schedule input' };
    }

    const spreadsheet = sourceSheet.getParent();
    const queueSheet = getQueueSheet_(spreadsheet);
    if (!queueSheet) {
      return { success: false, error: 'Queue sheet unavailable' };
    }

    const rowSignature = buildRowSignature_(sourceSheet, rowIndex, sheetRole);
    if (!rowSignature || rowSignature === '||||' || rowSignature === '|||') {
      return { success: false, error: 'Could not build row signature' };
    }

    try {
      const existingRow = findQueueRowIndex_(queueSheet, sourceSheet.getName(), rowSignature);
      const queueRow = [
        sourceSheet.getName(),
        sheetRole,
        rowSignature,
        statusValue,
        new Date(),
        mapping.completedSheet,
      ];

      if (existingRow > 0) {
        queueSheet.getRange(existingRow, 1, 1, QUEUE_HEADERS.length).setValues([queueRow]);
      } else {
        queueSheet.appendRow(queueRow);
      }

      return {
        success: true,
        sourceSheet: sourceSheet.getName(),
        rowSignature: rowSignature,
        scheduledAt: queueRow[4],
        completedSheet: mapping.completedSheet,
      };
    } catch (err) {
      ErrorLogModule.error('CompletionQueueModule', 'scheduleRow_', 'Schedule failed', {
        sourceSheet: sourceSheet.getName(),
        rowIndex: rowIndex,
        error: err,
      });
      return { success: false, error: String(err) };
    }
  }

  function cancelRow_(sourceSheetName, rowSignature) {
    const spreadsheet = AppConfig.getSpreadsheet();
    const queueSheet = getQueueSheet_(spreadsheet);
    if (!queueSheet) {
      return { success: false };
    }

    const queueRowIndex = findQueueRowIndex_(queueSheet, sourceSheetName, rowSignature);
    if (queueRowIndex < AppConfig.getDataStartRow()) {
      return { success: true, removed: false };
    }

    try {
      queueSheet.deleteRow(queueRowIndex);
      return { success: true, removed: true };
    } catch (err) {
      ErrorLogModule.error('CompletionQueueModule', 'cancelRow_', 'Cancel failed', { error: err });
      return { success: false, error: String(err) };
    }
  }

  function getDueEntries_() {
    const spreadsheet = AppConfig.getSpreadsheet();
    const queueSheet = getQueueSheet_(spreadsheet);
    if (!queueSheet) {
      return [];
    }

    const lastRow = SpreadsheetUtils.getLastDataRow(queueSheet, 1);
    if (lastRow < AppConfig.getDataStartRow()) {
      return [];
    }

    const now = new Date().getTime();
    const delayMs = AppConfig.getCompletionDelayMs();
    const rowCount = lastRow - AppConfig.getHeaderRow();
    const values = queueSheet.getRange(AppConfig.getDataStartRow(), 1, rowCount, QUEUE_HEADERS.length).getValues();
    const dueEntries = [];

    values.forEach(function (row, index) {
      const scheduledAt = row[4];
      if (!scheduledAt) {
        return;
      }
      const scheduledTime = scheduledAt instanceof Date ? scheduledAt.getTime() : new Date(scheduledAt).getTime();
      if (isNaN(scheduledTime)) {
        return;
      }
      if (now - scheduledTime >= delayMs) {
        dueEntries.push({
          queueRowIndex: AppConfig.getDataStartRow() + index,
          sourceSheet: row[0],
          sheetRole: row[1],
          rowSignature: row[2],
          status: row[3],
          scheduledAt: scheduledAt,
          completedSheet: row[5],
        });
      }
    });

    return dueEntries;
  }

  function removeQueueEntry_(queueRowIndex) {
    const spreadsheet = AppConfig.getSpreadsheet();
    const queueSheet = getQueueSheet_(spreadsheet);
    if (!queueSheet || queueRowIndex < AppConfig.getDataStartRow()) {
      return false;
    }
    try {
      queueSheet.deleteRow(queueRowIndex);
      return true;
    } catch (err) {
      ErrorLogModule.error('CompletionQueueModule', 'removeQueueEntry_', 'Remove queue entry failed', {
        queueRowIndex: queueRowIndex,
        error: err,
      });
      return false;
    }
  }

  return {
    scheduleRow: scheduleRow_,
    cancelRow: cancelRow_,
    getDueEntries: getDueEntries_,
    removeQueueEntry: removeQueueEntry_,
    buildRowSignature: buildRowSignature_,
    ensureQueueSheet: function () {
      return getQueueSheet_(AppConfig.getSpreadsheet());
    },
  };
})();
