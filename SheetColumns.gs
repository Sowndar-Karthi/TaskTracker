/**
 * Column indices for Team Project Tracker sheets (1-based).
 */
const SheetColumns = (function () {
  'use strict';

  const MAIN = {
    PROJECT_NAME: 1,
    WEBSITE_URL: 2,
    TASK: 3,
    PRIORITY: 4,
    ASSIGN: 5,
    START_DATE: 6,
    END_DATE: 7,
    AS_PER_WO: 8,
    STATUS: 9,
    ENV: 10,
    ESTIMATED_EFFORT: 11,
    ACTUAL_EFFORT: 12,
    INFORMED: 13,
    SUPPORT_BLOCKER: 14,
    LATEST_UPDATE: 15,
    PAGE_URL: 16,
    FUNCTIONALITY: 17,
    MAIL_SUBJECT: 18,
    INTERNAL_MAIL_SUBJECT: 19,
    GMAIL_THREAD_ID: 22,
    GOOGLE_TASK_ID: 23,
    LAST_COLUMN: 23,
  };

  const SUB = {
    PROJECT_NAME: 1,
    MAIN_TASK_NAME: 2,
    TICKET_ID: 3,
    TASK: 4,
    PRIORITY: 5,
    ASSIGN: 6,
    START_DATE: 7,
    END_DATE: 8,
    AS_PER_WO: 9,
    STATUS: 10,
    ENV: 11,
    ACCOUNT: 12,
    ESTIMATED_EFFORT: 13,
    ACTUAL_EFFORT: 14,
    INFORMED: 15,
    SUPPORT_BLOCKER: 16,
    LATEST_UPDATE: 17,
    PAGE_URL: 18,
    FUNCTIONALITY: 19,
    MAIL_SUBJECT: 20,
    INTERNAL_MAIL_SUBJECT: 21,
    GMAIL_THREAD_ID: 24,
    GOOGLE_TASK_ID: 25,
    LAST_COLUMN: 25,
  };

  const MATCHING = {
    ID: 1,
    MAIN_SHEET: 2,
    SUB_SHEET: 3,
    COMPLETED_SHEET: 4,
    LAST_COLUMN: 4,
  };

  const EMAIL = {
    TO: 1,
    CC: 2,
    BCC: 3,
    DATA_SHEET_NAME: 4,
    TIMEZONE: 5,
    ALERT_TIME: 6,
    NOTIFICATION_TYPE: 7,
    LAST_COLUMN: 7,
  };

  const ERROR = {
    LOGGED_AT: 1,
    SEVERITY: 2,
    MODULE: 3,
    FUNCTION: 4,
    SHEET_NAME: 5,
    ROW_INDEX: 6,
    TASK_KEY: 7,
    MESSAGE: 8,
    DETAILS: 9,
    TRIGGER_SOURCE: 10,
    RESOLVED: 11,
    FIX_NOTES: 12,
    LAST_COLUMN: 12,
  };

  const RENEWAL = {
    ITEM_NAME: 1,
    CATEGORY: 2,
    RECURRENCE_INTERVAL: 3,
    RECURRENCE_UNIT: 4,
    ANCHOR_TYPE: 5,
    ANCHOR_DATE: 6,
    DUE_DAY_OF_MONTH: 7,
    MINIMUM_DUE_AMOUNT: 8,
    TOTAL_DUE_AMOUNT: 9,
    LAST_PAID_DATE: 10,
    NEXT_DUE_DATE: 11,
    REMINDER_DAYS_BEFORE: 12,
    REMINDER_EMAIL: 13,
    STATUS: 14,
    LAST_COLUMN: 14,
  };

  /**
   * Train Route — weekly + one-time travel templates.
   * Order must match the spreadsheet header row exactly.
   */
  const TRAIN_ROUTE = {
    ROUTE_NAME: 1,
    FROM_STATION: 2,
    TO_STATION: 3,
    DAY_OF_WEEK: 4,
    RECURRENCE: 5,
    ACTIVE: 6,
    PREFERRED_CLASS: 7,
    TRAVEL_DATE: 8,
    PREFERRED_TRAIN_NO: 9,
    PREFERRED_TRAIN_NAME: 10,
    REMINDER_DAYS_BEFORE: 11,
    ALERT_TIME: 12,
    REMINDER_EMAIL: 13,
    NOTES: 14,
    LAST_COLUMN: 14,
  };

  /**
   * Train Booking History / Train Completed Journeys —
   * IRCTC import columns (same layout on both sheets).
   */
  const TRAIN_BOOKING = {
    FROM: 1,
    TO: 2,
    DATE_OF_JOURNEY: 3,
    NAME: 4,
    STATUS: 5,
    COACH: 6,
    SEAT_BERTH: 7,
    CLASS: 8,
    PNR: 9,
    TRAIN_NO: 10,
    SCHEDULED_DEPARTURE: 11,
    DATE_OF_BOARDING: 12,
    TRANSACTION_ID: 13,
    BOOKING_DATETIME: 14,
    USER_ID: 15,
    PASSENGER_MOBILE: 16,
    GMAIL_THREAD_ID: 17,
    LAST_COLUMN: 17,
  };

  return {
    MAIN: MAIN,
    SUB: SUB,
    MATCHING: MATCHING,
    EMAIL: EMAIL,
    ERROR: ERROR,
    RENEWAL: RENEWAL,
    TRAIN_ROUTE: TRAIN_ROUTE,
    TRAIN_BOOKING: TRAIN_BOOKING,
  };
})();
