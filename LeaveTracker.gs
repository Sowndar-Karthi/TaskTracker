/************************************************************
 * GOOGLE LEAVE TRACKER AUTOMATION
 * ----------------------------------------------------------
 * Spreadsheet : Project Tracker
 * Sheet       : Leave Tracker
 *
 * FEATURES
 * ----------------------------------------------------------
 * ✅ Auto Trigger Setup
 * ✅ 4 AM Daily Mail
 * ✅ 4 PM Daily Mail
 * ✅ Same Gmail Subject
 * ✅ Same Gmail Thread
 * ✅ TO / CC / BCC
 * ✅ Ignore Invalid Emails
 * ✅ Ignore Empty Emails
 * ✅ Responsive HTML Email (mobile cards + scroll)
 * ✅ Validation Error Table
 * ✅ Row Number Error Reporting
 * ✅ Upcoming Leaves
 * ✅ Ongoing Leaves
 * ✅ Completed Leaves (Last 5 Days)
 * ✅ Exclude Completed > 7 Days
 * ✅ Trigger ONLY after Leave Status Update
 * ✅ Avoid Mail Spam
 * ✅ Full Null Checks
 * ✅ Invalid Date Handling
 * ✅ Any Deliver On This Week Support
 *
 * ----------------------------------------------------------
 * IMPORTANT
 * ----------------------------------------------------------
 * AFTER PASTING THIS CODE:
 *
 * RUN (TestAllTrackers.gs):
 *    setupAllTrackersComplete()    — labels + Google Tasks + all triggers (preferred)
 *    setupAllTrackers()            — triggers only
 *    removeDuplicateTrackerTriggers() — cleanup only
 *    testAllTrackerFunctionality() — test all flows
 *
 * Legacy: setupLeaveTracker() — use setupAllTrackersComplete() instead
 *
 ************************************************************/

/************************************************************
 * CONFIGURATION
 ************************************************************/

const SHEET_NAME = "Leave Tracker";

const EMAIL_SUBJECT =
  "Team Leave Tracker Report";

/************************************************************
 * EMAIL CONFIG
 ************************************************************/

const TO_EMAILS = [
  "sowndar@vajraglobal.com"
];

const CC_EMAILS = [
  
];

const BCC_EMAILS = [
  "dragonkarthi00m@gmail.com"
];

/************************************************************
 * COLUMN INDEXES
 ************************************************************/

const COL_EMPLOYEE_NAME = 0;
const COL_LEAVE_TYPE = 1;
const COL_FROM_DATE = 2;
const COL_TO_DATE = 3;
const COL_ACTUAL_LEAVE = 4;
const COL_WHICH_ACTUAL_DAYS = 5;
const COL_WHICH_TOTAL_DAYS = 6;
const COL_TOTAL_DAYS = 7;
const COL_HALF_FULL = 8;
const COL_STATUS = 9;
const COL_COMP_OFF_DATE = 10;
const COL_REASON = 11;
const COL_ANY_DELIVER_THIS_WEEK = 12;
const COL_EMERGENCY = 13;
const COL_APPLIED_DATE = 14;
const COL_LEAVE_STATUS = 15;
const COL_APPROVED_BY = 16;
const COL_APPROVAL_DATE = 17;
const COL_REJECTION_REASON = 18;

const SCRIPT_PROP_SPREADSHEET_ID =
  "LEAVE_TRACKER_SPREADSHEET_ID";

/**
 * Script property keys for recipient overrides. Set any of these via:
 *   PropertiesService.getScriptProperties().setProperty('LEAVE_TO_EMAILS', 'a@x.com,b@y.com');
 * When the property is unset or empty, the hardcoded TO_EMAILS / CC_EMAILS /
 * BCC_EMAILS arrays above are used as defaults. Keeps the script runnable
 * out-of-the-box while allowing per-environment overrides without code edits.
 */
const SCRIPT_PROP_LEAVE_TO_EMAILS = "LEAVE_TO_EMAILS";
const SCRIPT_PROP_LEAVE_CC_EMAILS = "LEAVE_CC_EMAILS";
const SCRIPT_PROP_LEAVE_BCC_EMAILS = "LEAVE_BCC_EMAILS";

/**
 * ON-EDIT DEBOUNCE
 * --------------------------------------------------------------
 * onEditHandler does NOT send the report directly. It schedules a
 * one-shot time-based trigger N seconds in the future. Any further
 * edits within that window are coalesced into the same scheduled
 * send, so editing 20 cells in 30 seconds = 1 email, not 20.
 *
 * Apps Script's `after()` is "at-or-after", so actual fire time
 * can drift a few seconds — fine for our coalescing purpose.
 */
const LEAVE_EDIT_DEBOUNCE_SECONDS = 30;
const PROP_LEAVE_REPORT_PENDING = "LEAVE_REPORT_PENDING";
const PROP_LEAVE_REPORT_SCHEDULED_AT = "LEAVE_REPORT_SCHEDULED_AT";
const LEAVE_DEFERRED_TRIGGER_HANDLER = "processPendingLeaveReport";

const LEAVE_TRIGGER_HANDLERS = [
  "sendDailyLeaveReport",
  "onEditHandler",
  LEAVE_DEFERRED_TRIGGER_HANDLER
];

/************************************************************
 * SETUP FUNCTION
 * RUN ONLY ONE TIME
 ************************************************************/

function setupLeaveTracker() {

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    PropertiesService
      .getScriptProperties()
      .setProperty(
        SCRIPT_PROP_SPREADSHEET_ID,
        ss.getId()
      );

    removeExistingTriggers();

    /************************************************
     * 4 AM DAILY TRIGGER
     ************************************************/

    ScriptApp.newTrigger(
      "sendDailyLeaveReport"
    )
      .timeBased()
      .everyDays(1)
      .atHour(4)
      .create();

    /************************************************
     * 4 PM DAILY TRIGGER
     ************************************************/

    ScriptApp.newTrigger(
      "sendDailyLeaveReport"
    )
      .timeBased()
      .everyDays(1)
      .atHour(16)
      .create();

    /************************************************
     * ON EDIT TRIGGER
     ************************************************/

    ScriptApp.newTrigger(
      "onEditHandler"
    )
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    Logger.log(
      "All triggers created successfully"
    );

  } catch (err) {

    Logger.log(
      "Setup Error: " + err
    );

  }

}

/************************************************************
 * REMOVE EXISTING TRIGGERS
 ************************************************************/

function removeExistingTriggers() {

  ScriptApp.getProjectTriggers().forEach((trigger) => {

    const handler =
      trigger.getHandlerFunction();

    if (
      LEAVE_TRIGGER_HANDLERS.indexOf(
        handler
      ) !== -1
    ) {
      ScriptApp.deleteTrigger(trigger);
    }

  });

}

/************************************************************
 * ON EDIT HANDLER
 ************************************************************/

function onEditHandler(e) {

  try {

    const sheet =
      e.source.getActiveSheet();

    if (
      sheet.getName() !== SHEET_NAME
    ) {
      return;
    }

    const row =
      e.range.getRow();

    const col =
      e.range.getColumn();

    // Ignore Header
    if (row === 1) {
      return;
    }

    /************************************************
     * LEAVE STATUS COLUMN
     ************************************************/

    const cols =
      getLeaveColumnIndexes_(sheet);

    const leaveStatusCol =
      cols.leaveStatus + 1;

    // Trigger only after Leave Status update
    if (col !== leaveStatusCol) {
      return;
    }

    /************************************************
     * COALESCE RAPID EDITS
     * --------------------------------------------
     * Instead of sleeping and sending one email per
     * edit, schedule a single deferred send. Any
     * other edits in the next LEAVE_EDIT_DEBOUNCE_SECONDS
     * piggy-back on the same scheduled send.
     ************************************************/

    scheduleDeferredLeaveReport_();

  } catch (err) {

    Logger.log(
      "onEditHandler Error: " + err
    );

  }

}

/************************************************************
 * SCHEDULE DEFERRED LEAVE REPORT
 * ------------------------------------------------------------
 * Sets a "pending" flag, then creates a one-shot time-based
 * trigger N seconds out — but only if one isn't already
 * scheduled. This is the debounce / coalescing primitive.
 ************************************************************/

function scheduleDeferredLeaveReport_() {

  try {

    const props = PropertiesService.getScriptProperties();
    props.setProperty(PROP_LEAVE_REPORT_PENDING, "1");

    // If a scheduled trigger already exists, this edit will be
    // picked up when that trigger fires — nothing else to do.
    const existing = ScriptApp.getProjectTriggers();
    for (let i = 0; i < existing.length; i++) {
      if (existing[i].getHandlerFunction() === LEAVE_DEFERRED_TRIGGER_HANDLER) {
        Logger.log(
          "Leave report send already scheduled; coalescing this edit."
        );
        return;
      }
    }

    ScriptApp.newTrigger(LEAVE_DEFERRED_TRIGGER_HANDLER)
      .timeBased()
      .after(LEAVE_EDIT_DEBOUNCE_SECONDS * 1000)
      .create();

    props.setProperty(
      PROP_LEAVE_REPORT_SCHEDULED_AT,
      new Date().toISOString()
    );

    Logger.log(
      "Leave report send scheduled ~" +
        LEAVE_EDIT_DEBOUNCE_SECONDS +
        "s out."
    );

  } catch (err) {

    Logger.log(
      "scheduleDeferredLeaveReport_ Error: " + err
    );

  }

}

/************************************************************
 * PROCESS PENDING LEAVE REPORT (trigger handler)
 * ------------------------------------------------------------
 * Fires once after the debounce window. Cleans up its own
 * one-shot trigger, clears the pending flag, then sends the
 * report ONCE for the entire burst of recent edits.
 ************************************************************/

function processPendingLeaveReport() {

  try {

    // Self-cleanup: delete every one-shot trigger pointing at us
    // (defensive — usually there is exactly one).
    ScriptApp.getProjectTriggers().forEach((t) => {
      if (t.getHandlerFunction() === LEAVE_DEFERRED_TRIGGER_HANDLER) {
        try {
          ScriptApp.deleteTrigger(t);
        } catch (_) {
          // Ignore: trigger may already be gone.
        }
      }
    });

    const props = PropertiesService.getScriptProperties();
    const wasPending = props.getProperty(PROP_LEAVE_REPORT_PENDING) === "1";
    props.deleteProperty(PROP_LEAVE_REPORT_PENDING);
    props.deleteProperty(PROP_LEAVE_REPORT_SCHEDULED_AT);

    if (!wasPending) {
      Logger.log(
        "processPendingLeaveReport: no pending edits to report."
      );
      return;
    }

    sendDailyLeaveReport();

  } catch (err) {

    Logger.log(
      "processPendingLeaveReport Error: " + err
    );

  }

}

/************************************************************
 * MAIN REPORT FUNCTION
 ************************************************************/

function sendDailyLeaveReport() {

  try {

    const sheet = getLeaveTrackerSheet_();
    if (!sheet) {
      return;
    }

    const cols = getLeaveColumnIndexes_(sheet);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const displayValues = dataRange.getDisplayValues();
    const dataRows = values.length > 1 ? values.slice(1) : [];
    const today = new Date();

    const report = collectLeaveReportData_(
      dataRows,
      displayValues,
      cols,
      today
    );

    const html = buildResponsiveTemplate(
      report.filteredRows,
      report.validationErrors
    );

    sendLeaveTrackerEmail_(html);

  } catch (err) {

    Logger.log("sendDailyLeaveReport Error: " + err);

  }

}

/**
 * Iterates every data row and produces both the summary rows (top of mail)
 * and the validation error list (bottom of mail). Per-row failures are
 * caught so a single bad row never aborts the whole report.
 * @returns {{filteredRows: Object[], validationErrors: Object[]}}
 */
function collectLeaveReportData_(dataRows, displayValues, cols, today) {

  const filteredRows = [];
  const validationErrors = [];

  dataRows.forEach(function (row, index) {

    const sheetRow = index + 2;

    try {

      const displayRow = displayValues[sheetRow - 1] || [];
      const outcome = processOneLeaveRow_(row, displayRow, cols, today, sheetRow);

      if (outcome.errorEntry) {
        validationErrors.push(outcome.errorEntry);
      }

      if (outcome.summaryRow) {
        filteredRows.push(outcome.summaryRow);
      }

    } catch (err) {

      validationErrors.push({
        row: sheetRow,
        employee: "-",
        errors: err.toString()
      });

    }

  });

  return { filteredRows: filteredRows, validationErrors: validationErrors };

}

/**
 * Validates one leave row and decides whether it joins the summary table,
 * the validation error table, or both. Returns null halves where the row
 * doesn't contribute.
 * @returns {{summaryRow: Object|null, errorEntry: Object|null}}
 */
function processOneLeaveRow_(row, displayRow, cols, today, sheetRow) {

  // Skip fully empty rows (no employee / from / to)
  if (isLeaveRowFullyEmpty_(row, displayRow, cols)) {
    return { summaryRow: null, errorEntry: null };
  }

  // Skip fragment rows that have data elsewhere but no employee name
  const employeePreview = getCellValue_(row, displayRow, cols.employee);
  if (!employeePreview) {
    return { summaryRow: null, errorEntry: null };
  }

  const fields = extractLeaveRowFields_(row, displayRow, cols);

  const rowErrors = collectLeaveRowValidationErrors_({
    employee: fields.employee,
    leaveType: fields.leaveType,
    fromDate: fields.fromDate,
    toDate: fields.toDate,
    totalDays: fields.totalDays,
    halfFull: fields.halfFull,
    status: fields.status,
    anyDeliverText: fields.anyDeliverThisWeek,
    anyDeliverRaw: getCellRaw_(row, displayRow, cols.anyDeliver),
    anyDeliverDisplay: displayRow[cols.anyDeliver],
    leaveStatus: fields.leaveStatus
  });

  const errorEntry = rowErrors.length > 0
    ? { row: sheetRow, employee: fields.employee, errors: rowErrors.join(", ") }
    : null;

  const completedDays = getDaysDifference(today, fields.toDate);
  const include = shouldIncludeLeaveInSummary_(fields.status, completedDays);

  // Always show error rows on top too so they're visible in both tables
  if (!include && rowErrors.length === 0) {
    return { summaryRow: null, errorEntry: errorEntry };
  }

  return { summaryRow: fields, errorEntry: errorEntry };

}

function isLeaveRowFullyEmpty_(row, displayRow, cols) {

  return (
    isEmpty(getCellValue_(row, displayRow, cols.employee)) &&
    isEmpty(getCellValue_(row, displayRow, cols.fromDate)) &&
    isEmpty(getCellValue_(row, displayRow, cols.toDate))
  );

}

/**
 * Pulls all leave columns out of one sheet row in display-friendly form.
 * Returned shape matches what buildLeaveDesktopTable_ / buildLeaveMobileCards_
 * expect.
 */
function extractLeaveRowFields_(row, displayRow, cols) {

  return {
    employee: getCellValue_(row, displayRow, cols.employee),
    leaveType: getCellValue_(row, displayRow, cols.leaveType),
    fromDate: parseDate(getCellRaw_(row, displayRow, cols.fromDate)),
    toDate: parseDate(getCellRaw_(row, displayRow, cols.toDate)),
    actualLeave: getCellValue_(row, displayRow, cols.actualLeave),
    whichActualDays: getCellValue_(row, displayRow, cols.whichActualDays),
    whichTotalDays: getCellValue_(row, displayRow, cols.whichTotalDays),
    totalDays: getCellValue_(row, displayRow, cols.totalDays),
    halfFull: getCellValue_(row, displayRow, cols.halfFull),
    status: getCellValue_(row, displayRow, cols.status),
    anyDeliverThisWeek: getCellValue_(row, displayRow, cols.anyDeliver),
    leaveStatus: getCellValue_(row, displayRow, cols.leaveStatus),
    approvedBy: getCellValue_(row, displayRow, cols.approvedBy)
  };

}

/**
 * Pure business rule: which leave records appear in the summary table.
 *  - Upcoming / Ongoing: always
 *  - Completed: only within the last 5 days
 *  - Completed older than 7 days: never
 * @returns {boolean}
 */
function shouldIncludeLeaveInSummary_(status, completedDays) {

  const s = (status || "").toLowerCase();

  if (s === "upcoming" || s === "ongoing") {
    return true;
  }

  if (s === "completed") {
    if (completedDays > 7) {
      return false;
    }
    return completedDays <= 5;
  }

  return false;

}

/************************************************************
 * SAME THREAD EMAIL (never reply into Dev Tracker threads)
 ************************************************************/

/**
 * True if this Gmail thread is only for Team Leave Tracker Report.
 */
function isLeaveTrackerThread_(thread) {

  if (!thread) {
    return false;
  }

  try {

    const messages = thread.getMessages();
    if (!messages || messages.length === 0) {
      return false;
    }

    const firstSub = (messages[0].getSubject() || "").trim();
    if (firstSub !== EMAIL_SUBJECT) {
      return false;
    }

    for (let m = 0; m < messages.length; m++) {
      const sub = (messages[m].getSubject() || "").trim();
      if (sub.indexOf("[Dev Tracker]") === 0) {
        return false;
      }
    }

    return true;

  } catch (e) {
    return false;
  }

}

function findLeaveTrackerThread_() {

  const threads = GmailApp.search(
    'subject:"' + EMAIL_SUBJECT.replace(/"/g, '\\"') + '"',
    0,
    25
  );

  if (!threads) {
    return null;
  }

  // Gmail caps a single conversation at 100 messages. The 101st reply lands
  // in a fresh conversation in the inbox UI even though `thread.reply` keeps
  // the References header intact. Roll over to a new thread at 95 so the
  // user never sees Gmail's split mid-stream.
  const MAX_MESSAGES_PER_THREAD = 95;

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    if (!isLeaveTrackerThread_(thread)) {
      continue;
    }

    let messageCount = 0;
    try {
      messageCount = thread.getMessageCount();
    } catch (countErr) {
      messageCount = 0;
    }

    if (messageCount >= MAX_MESSAGES_PER_THREAD) {
      Logger.log(
        'Leave thread ' + thread.getId() +
        ' has ' + messageCount + ' messages — Gmail will split at 100. ' +
        'Starting a fresh Leave Tracker thread.'
      );
      // Most recent Leave Tracker thread is at/over the cap — caller will
      // start a brand new thread (next search returns this fresh thread).
      return null;
    }

    return thread;
  }

  return null;

}

/**
 * Reads a recipient list override from script properties (comma-separated)
 * and falls back to the in-file defaults when no override is set.
 * @param {string} propertyKey
 * @param {string[]} fallbackArray
 * @returns {string[]}
 */
function getLeaveRecipients_(propertyKey, fallbackArray) {

  try {

    const raw = PropertiesService
      .getScriptProperties()
      .getProperty(propertyKey);

    if (raw && raw.toString().trim() !== "") {

      const parsed = raw
        .toString()
        .split(",")
        .map(function (s) { return s.trim(); })
        .filter(function (s) { return s !== ""; });

      if (parsed.length > 0) {
        return parsed;
      }

    }

  } catch (e) {
    // Fall through to defaults
  }

  return fallbackArray || [];

}

function sendLeaveTrackerEmail_(
  htmlBody
) {

  try {

    const toEmails =
      cleanEmails(
        getLeaveRecipients_(
          SCRIPT_PROP_LEAVE_TO_EMAILS,
          TO_EMAILS
        )
      );

    const ccEmails =
      cleanEmails(
        getLeaveRecipients_(
          SCRIPT_PROP_LEAVE_CC_EMAILS,
          CC_EMAILS
        )
      );

    const bccEmails =
      cleanEmails(
        getLeaveRecipients_(
          SCRIPT_PROP_LEAVE_BCC_EMAILS,
          BCC_EMAILS
        )
      );

    /************************************************
     * VALIDATE TO EMAILS
     ************************************************/

    if (!toEmails) {

      Logger.log(
        "No valid TO emails"
      );

      return;

    }

    /************************************************
     * FIND EXISTING THREAD (exact subject, no Dev Tracker mail)
     ************************************************/

    const leaveThread = findLeaveTrackerThread_();

    /************************************************
     * REPLY TO SAME THREAD
     ************************************************/

    if (leaveThread) {

      leaveThread.reply(
        "HTML Email",
        {

          htmlBody:
            htmlBody,

          cc:
            ccEmails || "",

          bcc:
            bccEmails || ""

        }
      );

    } else {

      /**********************************************
       * FIRST EMAIL
       **********************************************/

      GmailApp.sendEmail(

        toEmails,

        EMAIL_SUBJECT,

        "HTML Email",

        {

          htmlBody:
            htmlBody,

          cc:
            ccEmails || "",

          bcc:
            bccEmails || "",

          name:
            "Leave Tracker Bot"

        }

      );

    }

  } catch (err) {

    Logger.log(
      "Email Error: " + err
    );

  }

}

/************************************************************
 * RESPONSIVE EMAIL TEMPLATE
 ************************************************************/

function buildEmailMobileStyles_() {

  return `
    <style type="text/css">
      body, table, td {
        -webkit-text-size-adjust: 100%;
        -ms-text-size-adjust: 100%;
      }
      table {
        border-collapse: collapse;
        mso-table-lspace: 0;
        mso-table-rspace: 0;
      }
      img {
        border: 0;
        outline: none;
        text-decoration: none;
      }
      .table-scroll {
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
        width: 100%;
        max-width: 100%;
      }
      .leave-data-table {
        min-width: 1280px;
      }
      .error-data-table {
        min-width: 520px;
      }
      .mobile-cards-wrap {
        display: none;
        max-height: 0;
        overflow: hidden;
        mso-hide: all;
      }
      .desktop-table-wrap {
        display: block;
        width: 100%;
      }
      .scroll-hint {
        display: none;
        font-size: 12px;
        color: #666;
        margin: 0 0 8px 0;
      }
      @media only screen and (max-width: 620px) {
        .email-body {
          padding: 10px !important;
        }
        .email-inner {
          padding: 14px !important;
        }
        .email-header {
          padding: 16px !important;
        }
        .email-header h2 {
          font-size: 20px !important;
        }
        .desktop-table-wrap {
          display: block !important;
          width: 100% !important;
        }
        .scroll-hint {
          display: block !important;
        }
        .mobile-cards-wrap {
          display: block !important;
          max-height: none !important;
          overflow: visible !important;
        }
        .dt-cell, .dt-head {
          padding: 8px 6px !important;
          font-size: 12px !important;
        }
        .leave-card {
          font-size: 14px !important;
        }
      }
      @media only screen and (max-width: 480px) {
        .desktop-table-wrap {
          display: none !important;
          max-height: 0 !important;
          overflow: hidden !important;
          mso-hide: all !important;
        }
        .scroll-hint {
          display: none !important;
        }
        .mobile-cards-wrap {
          display: block !important;
          max-height: none !important;
          overflow: visible !important;
        }
      }
    </style>
  `;

}

function wrapScrollableTable_(
  tableHtml,
  tableClass
) {

  return `
    <div class="table-scroll" style="
      overflow-x:auto;
      -webkit-overflow-scrolling:touch;
      width:100%;
      max-width:100%;
    ">
      <table
        class="${tableClass} data-table"
        role="presentation"
        cellpadding="0"
        cellspacing="0"
        style="
          width:100%;
          border-collapse:collapse;
          mso-table-lspace:0;
          mso-table-rspace:0;
        ">
        ${tableHtml}
      </table>
    </div>
  `;

}

function buildMobileCardRow_(
  label,
  value
) {

  return `
    <tr>
      <td style="
        padding:4px 8px 4px 0;
        color:#555;
        font-size:13px;
        vertical-align:top;
        white-space:nowrap;
        width:42%;
      ">
        <strong>${label}</strong>
      </td>
      <td style="
        padding:4px 0;
        color:#222;
        font-size:13px;
        vertical-align:top;
        word-break:break-word;
      ">
        ${displayCell(value)}
      </td>
    </tr>
  `;

}

function buildLeaveMobileCards_(rows) {

  let cards = "";

  rows.forEach((r) => {

    cards += `
      <div class="leave-card" style="
        border:1px solid #d0d7de;
        border-radius:8px;
        padding:14px;
        margin:0 0 14px 0;
        background:#f8fafc;
      ">
        <table
          role="presentation"
          width="100%"
          cellpadding="0"
          cellspacing="0"
          style="border-collapse:collapse;"
        >
          ${buildMobileCardRow_(
            "Employee",
            r.employee
          )}
          ${buildMobileCardRow_(
            "Leave Type",
            r.leaveType
          )}
          ${buildMobileCardRow_(
            "From",
            formatDate(r.fromDate)
          )}
          ${buildMobileCardRow_(
            "To",
            formatDate(r.toDate)
          )}
          ${buildMobileCardRow_(
            "Actual Leave",
            r.actualLeave
          )}
          ${buildMobileCardRow_(
            "Which Actual Days",
            r.whichActualDays
          )}
          ${buildMobileCardRow_(
            "Which Total Days",
            r.whichTotalDays
          )}
          ${buildMobileCardRow_(
            "Total Days",
            r.totalDays
          )}
          ${buildMobileCardRow_(
            "Half / Full",
            r.halfFull
          )}
          ${buildMobileCardRow_(
            "Status",
            r.status
          )}
          ${buildMobileCardRow_(
            "Any Deliver This Week",
            r.anyDeliverThisWeek
          )}
          ${buildMobileCardRow_(
            "Approval",
            r.leaveStatus
          )}
          ${buildMobileCardRow_(
            "Approved By",
            r.approvedBy
          )}
        </table>
      </div>
    `;

  });

  return `
    <div class="mobile-cards-wrap">
      ${cards}
    </div>
  `;

}

function buildLeaveDesktopTable_(rows) {

  let tableRows = "";

  rows.forEach((r) => {

    tableRows += `
      <tr>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.employee)}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.leaveType)}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(formatDate(r.fromDate))}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(formatDate(r.toDate))}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.actualLeave)}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.whichActualDays)}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.whichTotalDays)}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.totalDays)}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.halfFull)}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.status)}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.anyDeliverThisWeek)}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.leaveStatus)}
        </td>
        <td class="dt-cell" style="${cellStyle()}">
          ${displayCell(r.approvedBy)}
        </td>
      </tr>
    `;

  });

  const tableInner = `
    <thead>
      <tr style="background:#d9eaf7;">
        <th class="dt-head" style="${headerStyle()}">Employee</th>
        <th class="dt-head" style="${headerStyle()}">Leave Type</th>
        <th class="dt-head" style="${headerStyle()}">From</th>
        <th class="dt-head" style="${headerStyle()}">To</th>
        <th class="dt-head" style="${headerStyle()}">Actual Leave</th>
        <th class="dt-head" style="${headerStyle()}">Which Actual Days</th>
        <th class="dt-head" style="${headerStyle()}">Which Total Days</th>
        <th class="dt-head" style="${headerStyle()}">Total Days</th>
        <th class="dt-head" style="${headerStyle()}">Half / Full</th>
        <th class="dt-head" style="${headerStyle()}">Status</th>
        <th class="dt-head" style="${headerStyle()}">Any Deliver This Week</th>
        <th class="dt-head" style="${headerStyle()}">Approval</th>
        <th class="dt-head" style="${headerStyle()}">Approved By</th>
      </tr>
    </thead>
    <tbody>
      ${tableRows}
    </tbody>
  `;

  return wrapScrollableTable_(
    tableInner,
    "leave-data-table"
  );

}

function buildResponsiveTemplate(
  rows,
  validationErrors
) {

  const hasRows = !!(rows && rows.length > 0);

  const summarySection = buildLeaveSummarySection_(rows, hasRows);
  const errorSection = buildLeaveValidationErrorSection_(validationErrors, hasRows);

  return buildLeaveEmailDocument_(summarySection, errorSection);

}

/**
 * Builds the "Leave Summary" section (heading + responsive table + mobile cards).
 * Falls back to an empty-state card when there are no rows to show.
 * @returns {string} HTML fragment
 */
function buildLeaveSummarySection_(rows, hasRows) {

  if (hasRows) {

    return `
      <h3 style="
        color:#1f4e78;
        margin:0 0 8px 0;
        font-size:18px;
      ">
        Leave Summary
      </h3>

      <p class="scroll-hint" style="
        display:none;
        font-size:12px;
        color:#666;
        margin:0 0 8px 0;
      ">
        Swipe the table left/right to see all columns.
      </p>

      <div class="desktop-table-wrap">
        ${buildLeaveDesktopTable_(rows)}
      </div>

      ${buildLeaveMobileCards_(rows)}
    `;

  }

  return `
      <h3 style="
        color:#1f4e78;
        margin:0 0 12px 0;
        font-size:18px;
      ">
        Leave Summary
      </h3>

      <p style="
        margin:0;
        padding:16px;
        background:#f4f6f9;
        border-radius:6px;
        color:#555;
        font-size:14px;
      ">
        No valid leave entries in this report.
      </p>
    `;

}

/**
 * Builds the "Validation Errors" section. Always present so reviewers can
 * see at a glance that no errors exist.
 * @returns {string} HTML fragment
 */
function buildLeaveValidationErrorSection_(validationErrors, hasSummaryRows) {

  const errorTopMargin = hasSummaryRows ? "40px" : "24px";

  if (!validationErrors || validationErrors.length === 0) {

    return `

      <h3 style="
        color:#d9534f;
        margin-top:${errorTopMargin};
      ">
        Validation Errors
      </h3>

      <p style="
        margin:10px 0 0 0;
        padding:16px;
        background:#f4f6f9;
        border-radius:6px;
        color:#555;
      ">
        No validation errors found.
      </p>

    `;

  }

  let errorRows = "";

  validationErrors.forEach(function (e) {

    errorRows += `
        <tr>
          <td class="dt-cell" style="${cellStyle()}">
            ${displayCell(e.row)}
          </td>
          <td class="dt-cell" style="${cellStyle()}">
            ${displayCell(e.employee)}
          </td>
          <td class="dt-cell" style="${cellStyle()} word-break:break-word;">
            ${displayCell(e.errors)}
          </td>
        </tr>
      `;

  });

  const errorTableInner = `
        <thead>
          <tr style="background:#f8d7da;">
            <th class="dt-head" style="${headerStyle()}">Row Number</th>
            <th class="dt-head" style="${headerStyle()}">Employee</th>
            <th class="dt-head" style="${headerStyle()}">Error</th>
          </tr>
        </thead>
        <tbody>
          ${errorRows}
        </tbody>
    `;

  return `

      <h3 style="
        color:#d9534f;
        margin-top:${errorTopMargin};
        font-size:18px;
      ">
        Validation Errors
      </h3>

      ${wrapScrollableTable_(
        errorTableInner,
        "error-data-table"
      )}

    `;

}

/**
 * Wraps the two body sections in the full responsive email shell (head,
 * styles, branded header, inner card).
 * @returns {string} Complete HTML document
 */
function buildLeaveEmailDocument_(summarySection, errorSection) {

  return `
  <!DOCTYPE html>
  <html lang="en">

    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <meta http-equiv="X-UA-Compatible" content="IE=edge">
      <title>${EMAIL_SUBJECT}</title>
      ${buildEmailMobileStyles_()}
    </head>

    <body class="email-body" style="
      margin:0;
      padding:20px;
      background:#f4f6f9;
      font-family:Arial, Helvetica, sans-serif;
      font-size:15px;
      line-height:1.4;
      color:#222;
    ">

      <div style="
        max-width:1200px;
        margin:0 auto;
        background:#ffffff;
        border-radius:10px;
        overflow:visible;
      ">

        <div class="email-header" style="
          background:#1f4e78;
          color:#ffffff;
          padding:20px;
        ">

          <h2 style="
            margin:0 0 8px 0;
            font-size:22px;
          ">
            Team Leave Tracker Report
          </h2>

          <p style="margin:0;font-size:14px;">
            Auto Generated Leave Summary
          </p>

        </div>

        <div class="email-inner" style="
          padding:20px;
          overflow:visible;
        ">

          ${summarySection}

          ${errorSection}

        </div>

      </div>

    </body>

  </html>
  `;

}

/************************************************************
 * SPREADSHEET + COLUMN HELPERS
 ************************************************************/

function getLeaveSpreadsheet_() {

  const storedId =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        SCRIPT_PROP_SPREADSHEET_ID
      );

  if (storedId) {
    return SpreadsheetApp.openById(storedId);
  }

  return SpreadsheetApp.getActiveSpreadsheet();

}

function getLeaveTrackerSheet_() {

  const ss =
    getLeaveSpreadsheet_();

  if (!ss) {
    return null;
  }

  return ss.getSheetByName(SHEET_NAME);

}

function findHeaderColumn_(
  headers,
  candidates,
  fallback
) {

  const normalized =
    headers.map((h) =>
      (h || "")
        .toString()
        .trim()
        .toLowerCase()
    );

  for (let i = 0; i < candidates.length; i++) {

    const name =
      candidates[i].toLowerCase();

    const index =
      normalized.indexOf(name);

    if (index !== -1) {
      return index;
    }

  }

  return fallback;

}

function getLeaveColumnIndexes_(sheet) {

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0];

  return {
    employee: findHeaderColumn_(
      headers,
      ["employee name"],
      COL_EMPLOYEE_NAME
    ),
    leaveType: findHeaderColumn_(
      headers,
      ["leave type"],
      COL_LEAVE_TYPE
    ),
    fromDate: findHeaderColumn_(
      headers,
      ["from date"],
      COL_FROM_DATE
    ),
    toDate: findHeaderColumn_(
      headers,
      ["to date"],
      COL_TO_DATE
    ),
    actualLeave: findHeaderColumn_(
      headers,
      ["actual leave"],
      COL_ACTUAL_LEAVE
    ),
    whichActualDays: findHeaderColumn_(
      headers,
      [
        "which actual days"
      ],
      COL_WHICH_ACTUAL_DAYS
    ),
    whichTotalDays: findHeaderColumn_(
      headers,
      [
        "which total days"
      ],
      COL_WHICH_TOTAL_DAYS
    ),
    totalDays: findHeaderColumn_(
      headers,
      ["total days"],
      COL_TOTAL_DAYS
    ),
    halfFull: findHeaderColumn_(
      headers,
      [
        "half day / full day",
        "half / full day"
      ],
      COL_HALF_FULL
    ),
    status: findHeaderColumn_(
      headers,
      ["status"],
      COL_STATUS
    ),
    anyDeliver: findHeaderColumn_(
      headers,
      ["any deliver on this week"],
      COL_ANY_DELIVER_THIS_WEEK
    ),
    leaveStatus: findHeaderColumn_(
      headers,
      ["leave status"],
      COL_LEAVE_STATUS
    ),
    approvedBy: findHeaderColumn_(
      headers,
      ["approved by"],
      COL_APPROVED_BY
    )
  };

}

function getCellRaw_(
  row,
  displayRow,
  colIndex
) {

  if (
    colIndex === null ||
    colIndex === undefined ||
    colIndex < 0
  ) {
    return "";
  }

  const raw =
    row[colIndex];

  if (!isBlankCell_(raw)) {
    return raw;
  }

  const shown =
    displayRow[colIndex];

  if (!isBlankCell_(shown)) {
    return shown;
  }

  return "";

}

function getCellValue_(row, displayRow, colIndex) {

  return normalizeCellText_(
    getCellRaw_(row, displayRow, colIndex)
  );

}

function isBlankCell_(value) {

  if (value === null || value === undefined) {
    return true;
  }

  if (typeof value === "boolean") {
    return false;
  }

  if (
    typeof value === "number" &&
    !isNaN(value)
  ) {
    return false;
  }

  return (
    value
      .toString()
      .trim() === ""
  );

}

function hasAnyDeliverAnswer_(value) {

  if (value === null || value === undefined) {
    return false;
  }

  if (typeof value === "boolean") {
    return true;
  }

  if (
    typeof value === "number" &&
    !isNaN(value)
  ) {
    return true;
  }

  return (
    value
      .toString()
      .trim() !== ""
  );

}

function normalizeCellText_(value) {

  if (value === null || value === undefined) {
    return "";
  }

  if (value === true) {
    return "Yes";
  }

  if (value === false) {
    return "No";
  }

  return value.toString().trim();

}

function collectLeaveRowValidationErrors_(
  fields
) {

  const errors = [];

  if (!fields.employee) {
    errors.push("Employee Name Missing");
  }

  if (!fields.leaveType) {
    errors.push("Leave Type Missing");
  }

  if (!fields.fromDate) {
    errors.push("Invalid From Date");
  }

  if (!fields.toDate) {
    errors.push("Invalid To Date");
  }

  if (!fields.totalDays) {
    errors.push("Total Days Missing");
  }

  if (!fields.halfFull) {
    errors.push("Half Day / Full Day Missing");
  }

  if (!fields.status) {
    errors.push("Status Missing");
  }

  if (
    !normalizeCellText_(
      fields.anyDeliverText
    ) &&
    !hasAnyDeliverAnswer_(
      fields.anyDeliverRaw
    ) &&
    !hasAnyDeliverAnswer_(
      fields.anyDeliverDisplay
    )
  ) {
    errors.push(
      "Any Deliver On This Week Missing"
    );
  }

  if (!fields.leaveStatus) {
    errors.push("Leave Status Missing");
  }

  return errors;

}

function displayCell(value) {

  if (value === null || value === undefined) {
    return "—";
  }

  if (
    typeof value === "number" &&
    !isNaN(value)
  ) {
    return escapeHtml_(String(value));
  }

  const text =
    normalizeCellText_(value);

  if (text === "") {
    return "—";
  }

  return escapeHtml_(text);

}

function escapeHtml_(text) {

  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

}

/************************************************************
 * EMAIL CLEANER
 ************************************************************/

function cleanEmails(
  emails
) {

  try {

    if (
      !emails ||
      !Array.isArray(
        emails
      )
    ) {

      return "";

    }

    return emails
      .filter((email) => {

        if (!email) {
          return false;
        }

        const value =
          email
            .toString()
            .trim();

        if (
          value === ""
        ) {
          return false;
        }

        const regex =
          /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

        return regex.test(
          value
        );

      })
      .join(",");

  } catch (err) {

    return "";

  }

}

/************************************************************
 * SAFE VALUE
 ************************************************************/

function safeValue(value) {

  try {

    if (
      value === null ||
      value === undefined
    ) {

      return "";

    }

    return value.toString();

  } catch (err) {

    return "";

  }

}

/************************************************************
 * EMPTY CHECK
 ************************************************************/

function isEmpty(value) {

  if (value === null || value === undefined) {
    return true;
  }

  if (typeof value === "boolean") {
    return false;
  }

  return (
    value
      .toString()
      .trim() === ""
  );

}

/************************************************************
 * DATE PARSER
 ************************************************************/

function parseDate(value) {

  try {

    if (!value) {
      return null;
    }

    const date =
      new Date(value);

    if (
      isNaN(
        date.getTime()
      )
    ) {

      return null;

    }

    return date;

  } catch (err) {

    return null;

  }

}

/************************************************************
 * FORMAT DATE
 ************************************************************/

function formatDate(date) {

  try {

    const parsed =
      parseDate(date);

    if (!parsed) {

      return "";

    }

    return Utilities.formatDate(

      parsed,

      Session
        .getScriptTimeZone(),

      "dd-MMM-yyyy"

    );

  } catch (err) {

    return "";

  }

}

/************************************************************
 * DATE DIFFERENCE
 ************************************************************/

function getDaysDifference(
  currentDate,
  oldDate
) {

  try {

    return Math.floor(

      (
        currentDate -
        oldDate
      ) /

      (
        1000 *
        60 *
        60 *
        24
      )

    );

  } catch (err) {

    return 9999;

  }

}

/************************************************************
 * STYLE HELPERS
 ************************************************************/

function headerStyle() {

  return `
    padding:10px 8px;
    border:1px solid #ccc;
    text-align:left;
    font-weight:bold;
    font-size:13px;
    white-space:nowrap;
  `;

}

function cellStyle() {

  return `
    padding:8px;
    border:1px solid #ddd;
    font-size:13px;
    word-break:break-word;
  `;

}