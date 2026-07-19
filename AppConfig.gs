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

  const TRAIN_ROUTE_SHEET_NAME = 'Train Route';
  const TRAIN_BOOKING_HISTORY_SHEET_NAME = 'Train Booking History';
  const TRAIN_COMPLETED_JOURNEYS_SHEET_NAME = 'Train Completed Journeys';
  /** Optional — maps IRCTC station text to canonical labels (column A). */
  const TRAIN_MATCH_PLACE_SHEET_NAME = 'Match Place';
  const IRCTC_EMAIL_FROM = 'ticketadmin@irctc.co.in';
  const IRCTC_MAX_THREADS_PER_RUN = 40;
  const IRCTC_MARK_READ_AFTER_IMPORT = true;
  /** Forwarded IRCTC confirmations from these accounts are also imported. */
  const IRCTC_FORWARD_FROM_EMAILS = ['narutokarthi00m@gmail.com', 'dragonkarthi00m@gmail.com'];
  /** Override Gmail search via Script Property key IRCTC_GMAIL_QUERY if set. */
  const IRCTC_GMAIL_QUERY_PROPERTY = 'IRCTC_GMAIL_QUERY';
  const TRAIN_DATE_TIMEZONE = 'Asia/Kolkata';
  /** IRCTC Advance Reservation Period (days before journey). */
  const TRAIN_ARP_DAYS = 60;
  /**
   * How far ahead (travel dates) the missing-booking list looks.
   * Must be ≥ ARP so Status can show “Opens in N day(s)” before the window opens.
   */
  const TRAIN_REMINDER_LOOKAHEAD_DAYS = 90;
  /** Usual IRCTC general booking open hour (IST), informational for subjects. */
  const TRAIN_BOOKING_OPEN_HOUR = 8;


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
    getTrainRouteSheetName: function () {
      return TRAIN_ROUTE_SHEET_NAME;
    },
    getTrainBookingHistorySheetName: function () {
      return TRAIN_BOOKING_HISTORY_SHEET_NAME;
    },
    getTrainCompletedJourneysSheetName: function () {
      return TRAIN_COMPLETED_JOURNEYS_SHEET_NAME;
    },
    getTrainMatchPlaceSheetName: function () {
      return TRAIN_MATCH_PLACE_SHEET_NAME;
    },
    getIrctcEmailFrom: function () {
      return IRCTC_EMAIL_FROM;
    },
    getIrctcMaxThreadsPerRun: function () {
      return IRCTC_MAX_THREADS_PER_RUN;
    },
    getIrctcMarkReadAfterImport: function () {
      return IRCTC_MARK_READ_AFTER_IMPORT;
    },
    getIrctcForwardFromEmails: function () {
      return IRCTC_FORWARD_FROM_EMAILS.slice();
    },
    getIrctcGmailQueryPropertyKey: function () {
      return IRCTC_GMAIL_QUERY_PROPERTY;
    },
    getTrainDateTimezone: function () {
      return TRAIN_DATE_TIMEZONE;
    },
    getTrainArpDays: function () {
      return TRAIN_ARP_DAYS;
    },
    getTrainReminderLookaheadDays: function () {
      return TRAIN_REMINDER_LOOKAHEAD_DAYS;
    },
    getTrainBookingOpenHour: function () {
      return TRAIN_BOOKING_OPEN_HOUR;
    },
  };
})();
