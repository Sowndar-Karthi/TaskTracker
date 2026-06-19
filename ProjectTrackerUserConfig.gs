/**
 * PROJECT TRACKER — USER CONFIG
 * =================================
 * Edit YOUR values here only. Paste this file into the same Apps Script project
 * as ProjectTracker.gs (Extensions → Apps Script → + Add file).
 *
 * Other .gs files read these values via getMainSheetName_(), getProjectMyEmail_(), etc.
 * You do NOT need to edit ProjectTracker.gs for email or sheet name changes.
 *
 * Optional overrides (advanced, rarely needed):
 *   Script property PROJECT_MY_EMAIL
 *   Script property PROJECT_BCC_EMAILS  (comma-separated)
 *   Script property PROJECT_BATCH_INTERVAL_HOURS
 */
const TRACKER_USER_CONFIG = {
  /** Primary sender for hourly updates, summaries, and test emails */
  MY_EMAIL: 'sowndar@vajraglobal.com',

  /**
   * BCC on every tracker email. Use [] or '' for none.
   * Or a comma-separated string: 'person1@x.com, person2@y.com'
   */
  BCC_EMAILS: [],

  /** Tab names — must match your spreadsheet exactly (case-sensitive) */
  MAIN_SHEET: 'Dev Tracker',
  SUB_SHEET: 'Dev Tracker Sub',
  PENDING_SHEET: 'Pending Changes',
  COMPLETED_SHEET: 'Completed Tracker',

  /** Gmail parent label; status labels are created as PARENT_LABEL/Status Name */
  PARENT_LABEL: 'A Assessment',

  /**
   * How often sendHourlyBatch runs (whole hours only).
   * Allowed: 1, 2, 4, 6, 8, or 12 — re-run setupAllTrackers() after changing.
   */
  BATCH_INTERVAL_HOURS: 1
};
