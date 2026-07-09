/**
 * Reads main → sub → completed mappings from the Matching Sheet.
 * Team sheet names are never hardcoded — always loaded from the spreadsheet.
 */
const MatchingSheetModule = (function () {
  'use strict';

  var cachedMappings_ = null;
  var cacheSpreadsheetId_ = null;

  function clearCache_() {
    cachedMappings_ = null;
    cacheSpreadsheetId_ = null;
  }

  function trimMappingValue_(value) {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim();
  }

  function normalizeMappingRow_(row, index) {
    return {
      rowNumber: AppConfig.getDataStartRow() + index,
      id: row[SheetColumns.MATCHING.ID - 1],
      mainSheet: trimMappingValue_(row[SheetColumns.MATCHING.MAIN_SHEET - 1]),
      subSheet: trimMappingValue_(row[SheetColumns.MATCHING.SUB_SHEET - 1]),
      completedSheet: trimMappingValue_(row[SheetColumns.MATCHING.COMPLETED_SHEET - 1]),
    };
  }

  function isValidMapping_(mapping) {
    return !!(mapping && mapping.mainSheet && mapping.subSheet && mapping.completedSheet);
  }

  function readMappings_(forceRefresh) {
    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      console.error('MatchingSheetModule.readMappings_: spreadsheet unavailable');
      return [];
    }

    const spreadsheetId = spreadsheet.getId();
    if (!forceRefresh && cachedMappings_ && cacheSpreadsheetId_ === spreadsheetId) {
      return cachedMappings_;
    }

    const matchingSheetName = AppConfig.getMatchingSheetName();
    const sheet = SpreadsheetUtils.getSheetByName(spreadsheet, matchingSheetName);
    if (!sheet) {
      console.error('MatchingSheetModule.readMappings_: matching sheet not found', {
        matchingSheetName: matchingSheetName,
      });
      return [];
    }

    const lastRow = SpreadsheetUtils.getLastDataRow(sheet, SheetColumns.MATCHING.MAIN_SHEET);
    if (lastRow < AppConfig.getDataStartRow()) {
      cachedMappings_ = [];
      cacheSpreadsheetId_ = spreadsheetId;
      return [];
    }

    try {
      const rowCount = lastRow - AppConfig.getHeaderRow();
      const values = sheet
        .getRange(AppConfig.getDataStartRow(), 1, rowCount, SheetColumns.MATCHING.LAST_COLUMN)
        .getValues();

      const mappings = values
        .map(normalizeMappingRow_)
        .filter(isValidMapping_);

      cachedMappings_ = mappings;
      cacheSpreadsheetId_ = spreadsheetId;
      return mappings;
    } catch (err) {
      console.error('MatchingSheetModule.readMappings_: read failed', { error: err });
      return [];
    }
  }

  function findMapping_(sheetName, fieldName) {
    if (!sheetName) {
      return null;
    }
    const normalized = SpreadsheetUtils.normalizeText(sheetName);
    const mappings = readMappings_();
    for (var i = 0; i < mappings.length; i++) {
      if (SpreadsheetUtils.normalizeText(mappings[i][fieldName]) === normalized) {
        return mappings[i];
      }
    }
    return null;
  }

  function getMappingByMainSheet_(mainSheetName) {
    return findMapping_(mainSheetName, 'mainSheet');
  }

  function getMappingBySubSheet_(subSheetName) {
    return findMapping_(subSheetName, 'subSheet');
  }

  function getMappingByCompletedSheet_(completedSheetName) {
    return findMapping_(completedSheetName, 'completedSheet');
  }

  function getSheetRole_(sheetName) {
    if (!sheetName) {
      return null;
    }

    const mainMapping = getMappingByMainSheet_(sheetName);
    if (mainMapping) {
      return { mapping: mainMapping, role: 'main' };
    }

    const subMapping = getMappingBySubSheet_(sheetName);
    if (subMapping) {
      return { mapping: subMapping, role: 'sub' };
    }

    const completedMapping = getMappingByCompletedSheet_(sheetName);
    if (completedMapping) {
      return { mapping: completedMapping, role: 'completed' };
    }

    return null;
  }

  function isTrackedSheet_(sheetName) {
    return !!getSheetRole_(sheetName);
  }

  function isMainOrSubSheet_(sheetName) {
    const sheetInfo = getSheetRole_(sheetName);
    return !!(sheetInfo && (sheetInfo.role === 'main' || sheetInfo.role === 'sub'));
  }

  function getAllMainSheetNames_() {
    return readMappings_().map(function (mapping) {
      return mapping.mainSheet;
    });
  }

  function getAllSubSheetNames_() {
    return readMappings_().map(function (mapping) {
      return mapping.subSheet;
    });
  }

  function getAllCompletedSheetNames_() {
    return readMappings_().map(function (mapping) {
      return mapping.completedSheet;
    });
  }

  function validateMappingsExist_() {
    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      return { valid: false, errors: ['Spreadsheet unavailable'] };
    }

    const errors = [];
    readMappings_(true).forEach(function (mapping) {
      ['mainSheet', 'subSheet', 'completedSheet'].forEach(function (field) {
        const sheetName = mapping[field];
        if (!SpreadsheetUtils.getSheetByName(spreadsheet, sheetName)) {
          errors.push('Row ' + mapping.rowNumber + ': tab not found — ' + sheetName + ' (' + field + ')');
        }
      });
    });

    if (errors.length > 0) {
      console.error('MatchingSheetModule.validateMappingsExist_: validation failed', { errors: errors });
    }

    return { valid: errors.length === 0, errors: errors };
  }

  return {
    getMappings: readMappings_,
    refreshMappings: function () {
      return readMappings_(true);
    },
    clearCache: clearCache_,
    getMappingByMainSheet: getMappingByMainSheet_,
    getMappingBySubSheet: getMappingBySubSheet_,
    getMappingByCompletedSheet: getMappingByCompletedSheet_,
    getSheetRole: getSheetRole_,
    isTrackedSheet: isTrackedSheet_,
    isMainOrSubSheet: isMainOrSubSheet_,
    getAllMainSheetNames: getAllMainSheetNames_,
    getAllSubSheetNames: getAllSubSheetNames_,
    getAllCompletedSheetNames: getAllCompletedSheetNames_,
    validateMappingsExist: validateMappingsExist_,
  };
})();
