/**
 * Syncs sub-task rows when a main task is added or updated.
 * Matches sub tasks by Main Task Name === main sheet Task column.
 */
const SubTaskSyncModule = (function () {
  'use strict';

  const SYNC_FIELD_MAP = [
    { mainCol: SheetColumns.MAIN.PROJECT_NAME, subCol: SheetColumns.SUB.PROJECT_NAME },
    { mainCol: SheetColumns.MAIN.PAGE_URL, subCol: SheetColumns.SUB.PAGE_URL },
    { mainCol: SheetColumns.MAIN.PRIORITY, subCol: SheetColumns.SUB.PRIORITY },
    { mainCol: SheetColumns.MAIN.ASSIGN, subCol: SheetColumns.SUB.ASSIGN },
    { mainCol: SheetColumns.MAIN.START_DATE, subCol: SheetColumns.SUB.START_DATE },
    { mainCol: SheetColumns.MAIN.END_DATE, subCol: SheetColumns.SUB.END_DATE },
    { mainCol: SheetColumns.MAIN.AS_PER_WO, subCol: SheetColumns.SUB.AS_PER_WO },
    { mainCol: SheetColumns.MAIN.ENV, subCol: SheetColumns.SUB.ENV },
    { mainCol: SheetColumns.MAIN.INFORMED, subCol: SheetColumns.SUB.INFORMED },
    { mainCol: SheetColumns.MAIN.SUPPORT_BLOCKER, subCol: SheetColumns.SUB.SUPPORT_BLOCKER },
    { mainCol: SheetColumns.MAIN.FUNCTIONALITY, subCol: SheetColumns.SUB.FUNCTIONALITY },
    { mainCol: SheetColumns.MAIN.MAIL_SUBJECT, subCol: SheetColumns.SUB.MAIL_SUBJECT },
    { mainCol: SheetColumns.MAIN.INTERNAL_MAIL_SUBJECT, subCol: SheetColumns.SUB.INTERNAL_MAIL_SUBJECT },
  ];

  function getMainTaskName_(mainRowValues) {
    if (!mainRowValues || mainRowValues.length < SheetColumns.MAIN.TASK) {
      return '';
    }
    const taskName = mainRowValues[SheetColumns.MAIN.TASK - 1];
    if (taskName === null || taskName === undefined || String(taskName).trim() === '') {
      return '';
    }
    return String(taskName).trim();
  }

  function findMatchingSubRowIndexes_(subRows, mainTaskName) {
    const matches = [];
    const normalizedMainTask = SpreadsheetUtils.normalizeText(mainTaskName);

    subRows.forEach(function (row, index) {
      if (!row || row.length < SheetColumns.SUB.MAIN_TASK_NAME) {
        return;
      }
      const subMainTaskName = row[SheetColumns.SUB.MAIN_TASK_NAME - 1];
      if (SpreadsheetUtils.normalizeText(subMainTaskName) === normalizedMainTask) {
        matches.push(index);
      }
    });

    return matches;
  }

  function applyMainValuesToSubRow_(subRow, mainRowValues) {
    const updatedRow = subRow.slice();
    SYNC_FIELD_MAP.forEach(function (fieldMap) {
      const mainValue = mainRowValues[fieldMap.mainCol - 1];
      updatedRow[fieldMap.subCol - 1] = mainValue;
    });
    return updatedRow;
  }

  function syncSubTasksForMainRow_(mainSheet, mainRowIndex, mapping) {
    if (!mainSheet || !mapping || !mapping.subSheet || mainRowIndex < AppConfig.getDataStartRow()) {
      return { success: false, matchedCount: 0, error: 'Invalid input' };
    }

    const mainRowValues = SpreadsheetUtils.getRowValues(
      mainSheet,
      mainRowIndex,
      SheetColumns.MAIN.LAST_COLUMN
    );
    const mainTaskName = getMainTaskName_(mainRowValues);
    if (!mainTaskName) {
      return { success: true, matchedCount: 0, skipped: true, reason: 'Empty main task name' };
    }

    const spreadsheet = mainSheet.getParent();
    const subSheet = SpreadsheetUtils.getSheetByName(spreadsheet, mapping.subSheet);
    if (!subSheet) {
      ErrorLogModule.error('SubTaskSyncModule', 'syncSubTasksForMainRow_', 'Sub sheet not found', {
        subSheet: mapping.subSheet,
      });
      return { success: false, matchedCount: 0, error: 'Sub sheet not found' };
    }

    const lastRow = SpreadsheetUtils.getLastDataRow(subSheet, SheetColumns.SUB.MAIN_TASK_NAME);
    if (lastRow < AppConfig.getDataStartRow()) {
      return { success: true, matchedCount: 0, reason: 'No sub-task rows' };
    }

    try {
      const rowCount = lastRow - AppConfig.getHeaderRow();
      const subDataRange = subSheet.getRange(
        AppConfig.getDataStartRow(),
        1,
        rowCount,
        SheetColumns.SUB.LAST_COLUMN
      );
      const subRows = subDataRange.getValues();
      const matchIndexes = findMatchingSubRowIndexes_(subRows, mainTaskName);

      if (matchIndexes.length === 0) {
        console.log('SubTaskSyncModule: no matching sub tasks', {
          mainSheet: mainSheet.getName(),
          mainTaskName: mainTaskName,
          subSheet: mapping.subSheet,
        });
        return { success: true, matchedCount: 0, mainTaskName: mainTaskName };
      }

      matchIndexes.forEach(function (matchIndex) {
        subRows[matchIndex] = applyMainValuesToSubRow_(subRows[matchIndex], mainRowValues);
      });

      subDataRange.setValues(subRows);

      const rollupResult = TaskEffortRollupModule.rollupEffortForMainTaskName(mapping, mainTaskName);

      console.log('SubTaskSyncModule: synced sub tasks', {
        mainSheet: mainSheet.getName(),
        mainRowIndex: mainRowIndex,
        mainTaskName: mainTaskName,
        subSheet: mapping.subSheet,
        matchedCount: matchIndexes.length,
      });

      return {
        success: true,
        matchedCount: matchIndexes.length,
        mainTaskName: mainTaskName,
        subSheet: mapping.subSheet,
        rollup: rollupResult,
      };
    } catch (err) {
      ErrorLogModule.error('SubTaskSyncModule', 'syncSubTasksForMainRow_', 'Sync failed', {
        mainSheet: mainSheet.getName(),
        mainRowIndex: mainRowIndex,
        subSheet: mapping.subSheet,
        error: err,
      });
      return { success: false, matchedCount: 0, error: String(err) };
    }
  }

  function syncSubTasksForMainSheetRow_(mainSheetName, mainRowIndex) {
    const mapping = MatchingSheetModule.getMappingByMainSheet(mainSheetName);
    if (!mapping) {
      return { success: false, matchedCount: 0, error: 'No mapping for main sheet' };
    }

    const spreadsheet = AppConfig.getSpreadsheet();
    const mainSheet = SpreadsheetUtils.getSheetByName(spreadsheet, mainSheetName);
    if (!mainSheet) {
      return { success: false, matchedCount: 0, error: 'Main sheet not found' };
    }

    return syncSubTasksForMainRow_(mainSheet, mainRowIndex, mapping);
  }

  function syncAllSubTasksForMainTaskName_(mainSheetName, mainTaskName) {
    if (!mainSheetName || !mainTaskName) {
      return { success: false, matchedCount: 0, error: 'Missing sheet or task name' };
    }

    const spreadsheet = AppConfig.getSpreadsheet();
    const mainSheet = SpreadsheetUtils.getSheetByName(spreadsheet, mainSheetName);
    const mapping = MatchingSheetModule.getMappingByMainSheet(mainSheetName);
    if (!mainSheet || !mapping) {
      return { success: false, matchedCount: 0, error: 'Sheet or mapping not found' };
    }

    const lastRow = SpreadsheetUtils.getLastDataRow(mainSheet, SheetColumns.MAIN.TASK);
    const normalizedTarget = SpreadsheetUtils.normalizeText(mainTaskName);
    let totalMatched = 0;

    for (var row = AppConfig.getDataStartRow(); row <= lastRow; row++) {
      const rowValues = SpreadsheetUtils.getRowValues(mainSheet, row, SheetColumns.MAIN.TASK);
      if (SpreadsheetUtils.normalizeText(getMainTaskName_(rowValues)) !== normalizedTarget) {
        continue;
      }
      const result = syncSubTasksForMainRow_(mainSheet, row, mapping);
      if (result && result.matchedCount) {
        totalMatched += result.matchedCount;
      }
    }

    return { success: true, matchedCount: totalMatched, mainTaskName: mainTaskName };
  }

  function syncSingleSubRowFromMain_(subSheet, subRowIndex, mapping) {
    if (!subSheet || !mapping || subRowIndex < AppConfig.getDataStartRow()) {
      return { success: false, error: 'Invalid input' };
    }

    const subRowValues = SpreadsheetUtils.getRowValues(subSheet, subRowIndex, SheetColumns.SUB.LAST_COLUMN);
    const mainTaskName = subRowValues[SheetColumns.SUB.MAIN_TASK_NAME - 1];
    if (!mainTaskName || String(mainTaskName).trim() === '') {
      return { success: true, skipped: true, reason: 'Empty Main Task Name on sub row' };
    }

    const spreadsheet = subSheet.getParent();
    const mainSheet = SpreadsheetUtils.getSheetByName(spreadsheet, mapping.mainSheet);
    if (!mainSheet) {
      return { success: false, error: 'Main sheet not found' };
    }

    const mainRowIndexes = TaskEffortRollupModule.findMainRowIndexesByTaskName(
      mainSheet,
      String(mainTaskName).trim()
    );
    if (mainRowIndexes.length === 0) {
      return {
        success: true,
        skipped: true,
        reason: 'No matching main task',
        mainTaskName: String(mainTaskName).trim(),
      };
    }

    try {
      const mainRowValues = SpreadsheetUtils.getRowValues(
        mainSheet,
        mainRowIndexes[0],
        SheetColumns.MAIN.LAST_COLUMN
      );
      const updatedSubRow = applyMainValuesToSubRow_(subRowValues, mainRowValues);
      subSheet.getRange(subRowIndex, 1, 1, SheetColumns.SUB.LAST_COLUMN).setValues([updatedSubRow]);

      const rollupResult = TaskEffortRollupModule.rollupEffortForMainTaskName(
        mapping,
        String(mainTaskName).trim()
      );

      return {
        success: true,
        mainTaskName: String(mainTaskName).trim(),
        mainRowIndex: mainRowIndexes[0],
        rollup: rollupResult,
      };
    } catch (err) {
      ErrorLogModule.error('SubTaskSyncModule', 'syncSingleSubRowFromMain_', 'Sync failed', {
        subSheet: subSheet.getName(),
        subRowIndex: subRowIndex,
        error: err,
      });
      return { success: false, error: String(err) };
    }
  }

  function syncSubRowForMainSheet_(subSheetName, subRowIndex) {
    const mapping = MatchingSheetModule.getMappingBySubSheet(subSheetName);
    if (!mapping) {
      return { success: false, error: 'No mapping for sub sheet' };
    }

    const spreadsheet = AppConfig.getSpreadsheet();
    const subSheet = SpreadsheetUtils.getSheetByName(spreadsheet, subSheetName);
    if (!subSheet) {
      return { success: false, error: 'Sub sheet not found' };
    }

    return syncSingleSubRowFromMain_(subSheet, subRowIndex, mapping);
  }

  return {
    syncSubTasksForMainRow: syncSubTasksForMainRow_,
    syncSubTasksForMainSheetRow: syncSubTasksForMainSheetRow_,
    syncAllSubTasksForMainTaskName: syncAllSubTasksForMainTaskName_,
    syncSingleSubRowFromMain: syncSingleSubRowFromMain_,
    syncSubRowForMainSheet: syncSubRowForMainSheet_,
  };
})();
