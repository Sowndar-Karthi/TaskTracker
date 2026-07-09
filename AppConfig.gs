/**
 * Application configuration — environment flags and spreadsheet IDs.
 */
const AppConfig = (function () {
  'use strict';

  const IS_TESTING = false;
  const PRODUCTION_SPREADSHEET_ID = ''; // Set your Team Project Tracker spreadsheet ID
  const TEST_SPREADSHEET_ID = ''; // Set your test copy spreadsheet ID

  const HEADER_ROW = 1;
  const DATA_START_ROW = 2;
  const COMPLETION_DELAY_HOURS = 24;
  const COMPLETION_QUEUE_SHEET = '_Completion Queue';
  /** Bootstrap tab name — team sheet names live on this sheet, not in code. */
  const MATCHING_SHEET_NAME = 'Matching Sheet';
  const EMAIL_SETTINGS_SHEET_NAME = '📧 Email Settings';

  function getActiveSpreadsheetId_() {
    return IS_TESTING ? TEST_SPREADSHEET_ID : PRODUCTION_SPREADSHEET_ID;
  }

  function getSpreadsheet_() {
    const spreadsheetId = getActiveSpreadsheetId_();
    if (spreadsheetId) {
      return SpreadsheetApp.openById(spreadsheetId);
    }
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  return {
    isTesting: function () {
      return IS_TESTING;
    },
    getSpreadsheetId: getActiveSpreadsheetId_,
    getSpreadsheet: getSpreadsheet_,
    getHeaderRow: function () {
      return HEADER_ROW;
    },
    getDataStartRow: function () {
      return DATA_START_ROW;
    },
    getCompletionDelayMs: function () {
      if (IS_TESTING) {
        return 60 * 1000;
      }
      return COMPLETION_DELAY_HOURS * 60 * 60 * 1000;
    },
    getCompletionQueueSheetName: function () {
      return COMPLETION_QUEUE_SHEET;
    },
    getMatchingSheetName: function () {
      return MATCHING_SHEET_NAME;
    },
    getEmailSettingsSheetName: function () {
      return EMAIL_SETTINGS_SHEET_NAME;
    },
  };
})();
