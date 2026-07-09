/**
 * Reusable spreadsheet read/write helpers.
 */
const SpreadsheetUtils = (function () {
  'use strict';

  function normalizeText_(value) {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim().toLowerCase();
  }

  function getSheetByName_(spreadsheet, sheetName) {
    if (!spreadsheet || !sheetName) {
      console.error('SpreadsheetUtils.getSheetByName_: missing spreadsheet or sheetName', {
        sheetName: sheetName,
      });
      return null;
    }
    try {
      return spreadsheet.getSheetByName(sheetName);
    } catch (err) {
      console.error('SpreadsheetUtils.getSheetByName_: failed', { sheetName: sheetName, error: err });
      return null;
    }
  }

  function getLastDataRow_(sheet, columnIndex) {
    if (!sheet) {
      return 0;
    }
    const col = columnIndex || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= AppConfig.getHeaderRow()) {
      return AppConfig.getHeaderRow();
    }
    return lastRow;
  }

  function getRowValues_(sheet, rowIndex, lastColumn) {
    if (!sheet || rowIndex < 1 || lastColumn < 1) {
      return [];
    }
    try {
      return sheet.getRange(rowIndex, 1, 1, lastColumn).getValues()[0];
    } catch (err) {
      console.error('SpreadsheetUtils.getRowValues_: read failed', {
        rowIndex: rowIndex,
        lastColumn: lastColumn,
        error: err,
      });
      return [];
    }
  }

  function columnRangeOverlaps_(rangeStartCol, rangeEndCol, targetCol) {
    return targetCol >= rangeStartCol && targetCol <= rangeEndCol;
  }

  return {
    normalizeText: normalizeText_,
    getSheetByName: getSheetByName_,
    getLastDataRow: getLastDataRow_,
    getRowValues: getRowValues_,
    columnRangeOverlaps: columnRangeOverlaps_,
  };
})();
