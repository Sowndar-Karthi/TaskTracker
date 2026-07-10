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
  const ERROR_SHEET_NAME = 'Error';
  const RENEWAL_TRACKER_SHEET_NAME = 'Renewal Tracker';
  const RENEWAL_REMINDER_HOUR = 9;
  /** Max sub/main rows per scheduled digest email (Gmail size limits). */
  const DIGEST_MAX_ROWS_PER_EMAIL = 100;

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
    getErrorSheetName: function () {
      return ERROR_SHEET_NAME;
    },
    getDigestMaxRowsPerEmail: function () {
      return DIGEST_MAX_ROWS_PER_EMAIL;
    },
    getRenewalTrackerSheetName: function () {
      return RENEWAL_TRACKER_SHEET_NAME;
    },
    getRenewalReminderHour: function () {
      return RENEWAL_REMINDER_HOUR;
    },
  };
})();
