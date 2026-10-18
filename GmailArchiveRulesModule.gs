/**
 * Sheet-driven Gmail labeling/archive:
 * Reads 📧 Gmail Archive Rules and applies labels to matching threads older than N days.
 * Add a new sheet row to support another subject tag / label — no code change needed.
 */
const GmailArchiveRulesModule = (function () {
  'use strict';

  const HEADERS = [
    'Active',
    'Subject Contains',
    'Label Name',
    'Older Than Days',
    'Archive Inbox',
    'Notes',
  ];

  const DEFAULT_ROWS = [
    ['Yes', '[TaskTracker]', 'TaskTracker/Older-3d', 3, 'Yes', 'All task digests + alerts'],
    ['Yes', '[Train]', 'Train/Older-3d', 3, 'Yes', 'Train reminder digests'],
    ['Yes', '[Renewal]', 'Renewal/Older-3d', 3, 'Yes', 'Renewal reminders'],
  ];

  function isYes_(value) {
    const n = SpreadsheetUtils.normalizeText(value);
    return n === 'yes' || n === 'y' || n === 'true' || n === '1' || n === 'active';
  }

  function parseDays_(value) {
    const n = parseInt(String(value || '').trim(), 10);
    if (isNaN(n) || n < 1) {
      return 0;
    }
    return n;
  }

  function getRulesSheet_() {
    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      return null;
    }
    return SpreadsheetUtils.getSheetByName(spreadsheet, AppConfig.getGmailArchiveRulesSheetName());
  }

  /**
   * Create sheet with headers + starter rows if missing.
   */
  function ensureRulesSheet_() {
    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      return { success: false, error: 'Spreadsheet unavailable' };
    }

    const sheetName = AppConfig.getGmailArchiveRulesSheetName();
    var sheet = SpreadsheetUtils.getSheetByName(spreadsheet, sheetName);
    var created = false;

    try {
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
        created = true;
      }

      const lastCol = Math.max(sheet.getLastColumn(), SheetColumns.GMAIL_ARCHIVE.LAST_COLUMN);
      const headerRange = sheet.getRange(1, 1, 1, SheetColumns.GMAIL_ARCHIVE.LAST_COLUMN);
      const existingHeaders = headerRange.getValues()[0];
      const headersMissing =
        !existingHeaders[0] ||
        SpreadsheetUtils.normalizeText(existingHeaders[0]) !== SpreadsheetUtils.normalizeText(HEADERS[0]);

      if (headersMissing) {
        headerRange.setValues([HEADERS]);
        sheet.setFrozenRows(1);
      }

      if (created || sheet.getLastRow() < AppConfig.getDataStartRow()) {
        sheet
          .getRange(AppConfig.getDataStartRow(), 1, DEFAULT_ROWS.length, SheetColumns.GMAIL_ARCHIVE.LAST_COLUMN)
          .setValues(DEFAULT_ROWS);
      }

      return { success: true, created: created, sheetName: sheetName };
    } catch (err) {
      console.error('GmailArchiveRulesModule.ensureRulesSheet_ failed', {
        error: err,
        stack: err && err.stack,
      });
      ErrorLogModule.error('GmailArchiveRulesModule', 'ensureRulesSheet_', String(err), {
        sheetName: sheetName,
      });
      return { success: false, error: String(err) };
    }
  }

  function normalizeRule_(row, rowIndex) {
    return {
      rowNumber: rowIndex,
      active: isYes_(row[SheetColumns.GMAIL_ARCHIVE.ACTIVE - 1]),
      subjectContains: String(row[SheetColumns.GMAIL_ARCHIVE.SUBJECT_CONTAINS - 1] || '').trim(),
      labelName: String(row[SheetColumns.GMAIL_ARCHIVE.LABEL_NAME - 1] || '').trim(),
      olderThanDays: parseDays_(row[SheetColumns.GMAIL_ARCHIVE.OLDER_THAN_DAYS - 1]),
      archiveInbox: isYes_(row[SheetColumns.GMAIL_ARCHIVE.ARCHIVE_INBOX - 1]),
      notes: String(row[SheetColumns.GMAIL_ARCHIVE.NOTES - 1] || '').trim(),
    };
  }

  function isValidRule_(rule) {
    return !!(
      rule &&
      rule.active &&
      rule.subjectContains &&
      rule.labelName &&
      rule.olderThanDays > 0
    );
  }

  function readActiveRules_() {
    ensureRulesSheet_();
    const sheet = getRulesSheet_();
    if (!sheet) {
      return [];
    }

    const lastRow = SpreadsheetUtils.getLastDataRow(sheet, SheetColumns.GMAIL_ARCHIVE.SUBJECT_CONTAINS);
    if (lastRow < AppConfig.getDataStartRow()) {
      return [];
    }

    try {
      const rowCount = lastRow - AppConfig.getHeaderRow();
      const values = sheet
        .getRange(AppConfig.getDataStartRow(), 1, rowCount, SheetColumns.GMAIL_ARCHIVE.LAST_COLUMN)
        .getValues();

      return values
        .map(function (row, index) {
          return normalizeRule_(row, AppConfig.getDataStartRow() + index);
        })
        .filter(isValidRule_);
    } catch (err) {
      console.error('GmailArchiveRulesModule.readActiveRules_ failed', {
        error: err,
        stack: err && err.stack,
      });
      ErrorLogModule.error('GmailArchiveRulesModule', 'readActiveRules_', String(err), {
        sheetName: AppConfig.getGmailArchiveRulesSheetName(),
      });
      return [];
    }
  }

  function ensureLabel_(labelName) {
    if (!labelName) {
      return null;
    }
    var label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      label = GmailApp.createLabel(labelName);
    }
    return label;
  }

  function threadHasLabel_(thread, label) {
    if (!thread || !label) {
      return false;
    }
    const labels = thread.getLabels();
    const target = SpreadsheetUtils.normalizeText(label.getName());
    for (var i = 0; i < labels.length; i++) {
      if (SpreadsheetUtils.normalizeText(labels[i].getName()) === target) {
        return true;
      }
    }
    return false;
  }

  function buildSearchQuery_(rule) {
    const subject = rule.subjectContains;
    // Quote when subject tag has brackets/spaces so Gmail search stays exact
    const quoted =
      subject.indexOf(' ') >= 0 || subject.indexOf('[') >= 0
        ? '"' + subject.replace(/"/g, '') + '"'
        : subject;
    return 'subject:' + quoted + ' older_than:' + rule.olderThanDays + 'd';
  }

  function processRule_(rule, summary) {
    const label = ensureLabel_(rule.labelName);
    if (!label) {
      summary.errors++;
      return;
    }

    const query = buildSearchQuery_(rule);
    const maxThreads = AppConfig.getGmailArchiveMaxThreadsPerRule();
    var threads = [];

    try {
      threads = GmailApp.search(query, 0, maxThreads);
    } catch (err) {
      summary.errors++;
      console.error('GmailArchiveRulesModule.processRule_: search failed', {
        row: rule.rowNumber,
        query: query,
        error: err,
      });
      ErrorLogModule.error('GmailArchiveRulesModule', 'processRule_', String(err), {
        details: query,
      });
      return;
    }

    threads.forEach(function (thread) {
      try {
        if (threadHasLabel_(thread, label)) {
          summary.skippedAlreadyLabeled++;
          return;
        }
        thread.addLabel(label);
        summary.labeled++;
        if (rule.archiveInbox) {
          thread.moveToArchive();
          summary.archived++;
        }
      } catch (err) {
        summary.errors++;
        console.error('GmailArchiveRulesModule.processRule_: thread failed', {
          row: rule.rowNumber,
          error: err,
        });
      }
    });

    console.log('GmailArchiveRulesModule.processRule_: done', {
      row: rule.rowNumber,
      query: query,
      label: rule.labelName,
      matched: threads.length,
    });
  }

  /**
   * Process all Active rules from the sheet.
   */
  function processAllRules_() {
    const summary = {
      rules: 0,
      labeled: 0,
      archived: 0,
      skippedAlreadyLabeled: 0,
      errors: 0,
    };

    const rules = readActiveRules_();
    summary.rules = rules.length;

    rules.forEach(function (rule) {
      try {
        processRule_(rule, summary);
      } catch (err) {
        summary.errors++;
        console.error('GmailArchiveRulesModule.processAllRules_: rule failed', {
          row: rule && rule.rowNumber,
          error: err,
          stack: err && err.stack,
        });
      }
    });

    console.log('GmailArchiveRulesModule.processAllRules_', summary);
    return summary;
  }

  return {
    ensureRulesSheet: ensureRulesSheet_,
    readActiveRules: readActiveRules_,
    processAllRules: processAllRules_,
  };
})();
