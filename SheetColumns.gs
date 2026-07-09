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
    ESTIMATED_EFFORT: 12,
    ACTUAL_EFFORT: 13,
    INFORMED: 14,
    SUPPORT_BLOCKER: 15,
    LATEST_UPDATE: 16,
    PAGE_URL: 17,
    FUNCTIONALITY: 18,
    MAIL_SUBJECT: 19,
    INTERNAL_MAIL_SUBJECT: 20,
    GMAIL_THREAD_ID: 23,
    GOOGLE_TASK_ID: 24,
    LAST_COLUMN: 24,
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

  return {
    MAIN: MAIN,
    SUB: SUB,
    MATCHING: MATCHING,
    EMAIL: EMAIL,
  };
})();
