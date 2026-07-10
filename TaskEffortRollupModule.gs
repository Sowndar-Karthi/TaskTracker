/**
 * Matches main Task (col C) to sub Main Task Name (col B) and rolls up effort totals.
 * Works for any team row on the Matching Sheet (including future teams).
 */
const TaskEffortRollupModule = (function () {
  'use strict';

  function parseEffort_(value) {
    if (value === null || value === undefined || value === '') {
      return 0;
    }
    if (typeof value === 'number') {
      return isNaN(value) ? 0 : value;
    }
    const parsed = parseFloat(String(value).replace(/[^\d.-]/g, ''));
    return isNaN(parsed) ? 0 : parsed;
  }

  function getMainTaskNameFromRow_(rowValues) {
    if (!rowValues || rowValues.length < SheetColumns.MAIN.TASK) {
      return '';
    }
    const taskName = rowValues[SheetColumns.MAIN.TASK - 1];
    if (taskName === null || taskName === undefined || String(taskName).trim() === '') {
      return '';
    }
    return String(taskName).trim();
  }

  function getSubMainTaskNameFromRow_(rowValues) {
    if (!rowValues || rowValues.length < SheetColumns.SUB.MAIN_TASK_NAME) {
      return '';
    }
    const taskName = rowValues[SheetColumns.SUB.MAIN_TASK_NAME - 1];
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
      if (SpreadsheetUtils.normalizeText(row[SheetColumns.SUB.MAIN_TASK_NAME - 1]) === normalizedMainTask) {
        matches.push(index);
      }
    });

    return matches;
  }

  function findMainRowIndexesByTaskName_(mainSheet, mainTaskName) {
    const matches = [];
    const normalizedTask = SpreadsheetUtils.normalizeText(mainTaskName);
    const lastRow = SpreadsheetUtils.getLastDataRow(mainSheet, SheetColumns.MAIN.TASK);

    for (var rowIndex = AppConfig.getDataStartRow(); rowIndex <= lastRow; rowIndex++) {
      const rowValues = SpreadsheetUtils.getRowValues(mainSheet, rowIndex, SheetColumns.MAIN.TASK);
      if (SpreadsheetUtils.normalizeText(getMainTaskNameFromRow_(rowValues)) === normalizedTask) {
        matches.push(rowIndex);
      }
    }

    return matches;
  }

  function sumEffortFromSubRows_(subRows, matchIndexes) {
    var totalEstimated = 0;
    var totalActual = 0;

    matchIndexes.forEach(function (index) {
      const row = subRows[index];
      if (!row) {
        return;
      }
      totalEstimated += parseEffort_(row[SheetColumns.SUB.ESTIMATED_EFFORT - 1]);
      totalActual += parseEffort_(row[SheetColumns.SUB.ACTUAL_EFFORT - 1]);
    });

    return {
      estimatedEffort: totalEstimated,
      actualEffort: totalActual,
    };
  }

  function effortsMatch_(currentValue, totalValue) {
    return parseEffort_(currentValue) === parseEffort_(totalValue);
  }

  function rollupEffortForMainTaskName_(mapping, mainTaskName) {
    if (!mapping || !mapping.mainSheet || !mapping.subSheet || !mainTaskName) {
      return { success: false, error: 'Invalid rollup input' };
    }

    const spreadsheet = AppConfig.getSpreadsheet();
    const mainSheet = SpreadsheetUtils.getSheetByName(spreadsheet, mapping.mainSheet);
    const subSheet = SpreadsheetUtils.getSheetByName(spreadsheet, mapping.subSheet);

    if (!mainSheet || !subSheet) {
      return { success: false, error: 'Main or sub sheet not found' };
    }

    const mainRowIndexes = findMainRowIndexesByTaskName_(mainSheet, mainTaskName);
    if (mainRowIndexes.length === 0) {
      return { success: true, skipped: true, reason: 'No matching main task row', mainTaskName: mainTaskName };
    }

    const subLastRow = SpreadsheetUtils.getLastDataRow(subSheet, SheetColumns.SUB.MAIN_TASK_NAME);
    if (subLastRow < AppConfig.getDataStartRow()) {
      return { success: true, skipped: true, reason: 'No sub-task rows', mainTaskName: mainTaskName };
    }

    try {
      const subRowCount = subLastRow - AppConfig.getHeaderRow();
      const subRows = subSheet
        .getRange(AppConfig.getDataStartRow(), 1, subRowCount, SheetColumns.SUB.LAST_COLUMN)
        .getValues();
      const matchIndexes = findMatchingSubRowIndexes_(subRows, mainTaskName);

      if (matchIndexes.length === 0) {
        return {
          success: true,
          skipped: true,
          reason: 'No matching sub tasks',
          mainTaskName: mainTaskName,
        };
      }

      const totals = sumEffortFromSubRows_(subRows, matchIndexes);
      const updatedMainRows = [];

      mainRowIndexes.forEach(function (mainRowIndex) {
        const estimatedCell = mainSheet.getRange(mainRowIndex, SheetColumns.MAIN.ESTIMATED_EFFORT);
        const actualCell = mainSheet.getRange(mainRowIndex, SheetColumns.MAIN.ACTUAL_EFFORT);
        const currentEstimated = estimatedCell.getValue();
        const currentActual = actualCell.getValue();

        if (!effortsMatch_(currentEstimated, totals.estimatedEffort)) {
          estimatedCell.setValue(totals.estimatedEffort);
        }
        if (!effortsMatch_(currentActual, totals.actualEffort)) {
          actualCell.setValue(totals.actualEffort);
        }

        updatedMainRows.push({
          rowIndex: mainRowIndex,
          estimatedEffort: totals.estimatedEffort,
          actualEffort: totals.actualEffort,
        });
      });

      console.log('TaskEffortRollupModule: effort rolled up to main', {
        mainSheet: mapping.mainSheet,
        mainTaskName: mainTaskName,
        subSheet: mapping.subSheet,
        matchedSubCount: matchIndexes.length,
        totals: totals,
        updatedMainRows: updatedMainRows,
      });

      return {
        success: true,
        mainTaskName: mainTaskName,
        matchedSubCount: matchIndexes.length,
        totals: totals,
        updatedMainRows: updatedMainRows,
      };
    } catch (err) {
      ErrorLogModule.error('TaskEffortRollupModule', 'rollupEffortForMainTaskName_', 'Rollup failed', {
        mainSheet: mapping.mainSheet,
        mainTaskName: mainTaskName,
        error: err,
      });
      return { success: false, error: String(err) };
    }
  }

  function rollupEffortForMainRow_(mainSheet, mainRowIndex, mapping) {
    const mainRowValues = SpreadsheetUtils.getRowValues(
      mainSheet,
      mainRowIndex,
      SheetColumns.MAIN.LAST_COLUMN
    );
    const mainTaskName = getMainTaskNameFromRow_(mainRowValues);
    if (!mainTaskName) {
      return { success: true, skipped: true, reason: 'Empty main task name' };
    }
    return rollupEffortForMainTaskName_(mapping, mainTaskName);
  }

  function rollupEffortForSubRow_(subSheet, subRowIndex, mapping) {
    const subRowValues = SpreadsheetUtils.getRowValues(subSheet, subRowIndex, SheetColumns.SUB.LAST_COLUMN);
    const mainTaskName = getSubMainTaskNameFromRow_(subRowValues);
    if (!mainTaskName) {
      return { success: true, skipped: true, reason: 'Empty sub Main Task Name' };
    }
    return rollupEffortForMainTaskName_(mapping, mainTaskName);
  }

  function rollupEffortForMainSheetName_(mainSheetName, mainTaskName) {
    const mapping = MatchingSheetModule.getMappingByMainSheet(mainSheetName);
    if (!mapping) {
      return { success: false, error: 'No mapping for main sheet' };
    }
    return rollupEffortForMainTaskName_(mapping, mainTaskName);
  }

  return {
    rollupEffortForMainTaskName: rollupEffortForMainTaskName_,
    rollupEffortForMainRow: rollupEffortForMainRow_,
    rollupEffortForSubRow: rollupEffortForSubRow_,
    rollupEffortForMainSheetName: rollupEffortForMainSheetName_,
    findMainRowIndexesByTaskName: findMainRowIndexesByTaskName_,
    parseEffort: parseEffort_,
  };
})();
