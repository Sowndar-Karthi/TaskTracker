/**
 * Reads renewal tracker rows from the Renewal Tracker sheet.
 */
const RenewalTrackerModule = (function () {
  'use strict';

  const REMINDABLE_STATUSES = ['active', 'due soon'];
  const DATE_TIMEZONE = 'Asia/Kolkata';
  const DATE_DISPLAY_FORMAT = 'd-MMM-yyyy';
  const DATE_PARSE_FORMATS = ['d-MMM-yyyy', 'dd-MMM-yyyy', 'd/MMM/yyyy', 'dd/MMM/yyyy'];

  function startOfDay_(date) {
    const copy = new Date(date);
    copy.setHours(0, 0, 0, 0);
    return copy;
  }

  function parseSheetSerialDate_(serial) {
    if (typeof serial !== 'number' || isNaN(serial)) {
      return null;
    }
    const utcDays = Math.floor(serial - 25569);
    const parsed = new Date(utcDays * 86400 * 1000);
    return isNaN(parsed.getTime()) ? null : startOfDay_(parsed);
  }

  function parseDate_(value) {
    if (value === null || value === undefined || value === '') {
      return null;
    }
    if (value instanceof Date && !isNaN(value.getTime())) {
      return startOfDay_(value);
    }
    if (typeof value === 'number') {
      return parseSheetSerialDate_(value);
    }

    const text = String(value).trim();
    if (!text) {
      return null;
    }

    for (var i = 0; i < DATE_PARSE_FORMATS.length; i++) {
      try {
        const parsed = Utilities.parseDate(text, DATE_TIMEZONE, DATE_PARSE_FORMATS[i]);
        if (parsed && !isNaN(parsed.getTime())) {
          return startOfDay_(parsed);
        }
      } catch (err) {
        // try next format
      }
    }

    const fallback = new Date(text);
    return isNaN(fallback.getTime()) ? null : startOfDay_(fallback);
  }

  function formatDate_(date) {
    if (!date) {
      return '';
    }
    return Utilities.formatDate(date, DATE_TIMEZONE, DATE_DISPLAY_FORMAT);
  }

  function parseNumber_(value, defaultValue) {
    if (value === null || value === undefined || value === '') {
      return defaultValue === undefined ? null : defaultValue;
    }
    if (typeof value === 'number') {
      return isNaN(value) ? defaultValue : value;
    }
    const parsed = parseFloat(String(value).replace(/[^\d.-]/g, ''));
    return isNaN(parsed) ? defaultValue : parsed;
  }

  function normalizeAnchorType_(value) {
    return SpreadsheetUtils.normalizeText(value);
  }

  function normalizeStatus_(value) {
    return SpreadsheetUtils.normalizeText(value);
  }

  function parseEmailList_(rawValue) {
    if (typeof EmailSettingsModule !== 'undefined' && EmailSettingsModule.parseEmailList) {
      return EmailSettingsModule.parseEmailList(rawValue);
    }
    if (!rawValue || String(rawValue).trim() === '') {
      return [];
    }
    return String(rawValue)
      .split(',')
      .map(function (email) {
        return String(email).trim();
      })
      .filter(function (email) {
        return email !== '' && email.indexOf('@') > 0;
      });
  }

  function parseReminderDays_(rawValue) {
    if (!rawValue || String(rawValue).trim() === '') {
      return [];
    }
    return String(rawValue)
      .split(',')
      .map(function (part) {
        return parseInt(String(part).trim(), 10);
      })
      .filter(function (days) {
        return !isNaN(days) && days >= 0;
      });
  }

  function normalizeItemRow_(row, rowIndex) {
    const recurrenceUnit = String(row[SheetColumns.RENEWAL.RECURRENCE_UNIT - 1] || '').trim();
    const recurrenceInterval = parseNumber_(row[SheetColumns.RENEWAL.RECURRENCE_INTERVAL - 1], 1);

    return {
      rowNumber: rowIndex,
      itemName: String(row[SheetColumns.RENEWAL.ITEM_NAME - 1] || '').trim(),
      category: String(row[SheetColumns.RENEWAL.CATEGORY - 1] || '').trim(),
      recurrenceInterval: recurrenceInterval > 0 ? recurrenceInterval : 1,
      recurrenceUnit: recurrenceUnit,
      anchorType: String(row[SheetColumns.RENEWAL.ANCHOR_TYPE - 1] || '').trim(),
      anchorTypeKey: normalizeAnchorType_(row[SheetColumns.RENEWAL.ANCHOR_TYPE - 1]),
      anchorDate: parseDate_(row[SheetColumns.RENEWAL.ANCHOR_DATE - 1]),
      dueDayOfMonth: parseNumber_(row[SheetColumns.RENEWAL.DUE_DAY_OF_MONTH - 1], null),
      minimumDueAmount: parseNumber_(row[SheetColumns.RENEWAL.MINIMUM_DUE_AMOUNT - 1], null),
      totalDueAmount: parseNumber_(row[SheetColumns.RENEWAL.TOTAL_DUE_AMOUNT - 1], null),
      lastPaidDate: parseDate_(row[SheetColumns.RENEWAL.LAST_PAID_DATE - 1]),
      nextDueDate: parseDate_(row[SheetColumns.RENEWAL.NEXT_DUE_DATE - 1]),
      reminderDaysBefore: parseReminderDays_(row[SheetColumns.RENEWAL.REMINDER_DAYS_BEFORE - 1]),
      reminderEmails: parseEmailList_(row[SheetColumns.RENEWAL.REMINDER_EMAIL - 1]),
      status: String(row[SheetColumns.RENEWAL.STATUS - 1] || '').trim(),
      statusKey: normalizeStatus_(row[SheetColumns.RENEWAL.STATUS - 1]),
    };
  }

  function isValidItem_(item) {
    return !!(
      item &&
      item.itemName &&
      item.recurrenceUnit &&
      item.nextDueDate &&
      item.reminderDaysBefore.length > 0 &&
      item.reminderEmails.length > 0 &&
      REMINDABLE_STATUSES.indexOf(item.statusKey) >= 0
    );
  }

  function getRenewalSheet_(spreadsheet) {
    return SpreadsheetUtils.getSheetByName(spreadsheet, AppConfig.getRenewalTrackerSheetName());
  }

  function readAllItems_() {
    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      ErrorLogModule.error('RenewalTrackerModule', 'readAllItems_', 'Spreadsheet unavailable', {});
      return [];
    }

    const sheet = getRenewalSheet_(spreadsheet);
    if (!sheet) {
      return [];
    }

    const lastRow = SpreadsheetUtils.getLastDataRow(sheet, SheetColumns.RENEWAL.ITEM_NAME);
    if (lastRow < AppConfig.getDataStartRow()) {
      return [];
    }

    try {
      const rowCount = lastRow - AppConfig.getHeaderRow();
      const values = sheet
        .getRange(AppConfig.getDataStartRow(), 1, rowCount, SheetColumns.RENEWAL.LAST_COLUMN)
        .getValues();

      return values
        .map(function (row, index) {
          return normalizeItemRow_(row, AppConfig.getDataStartRow() + index);
        })
        .filter(function (item) {
          return item.itemName !== '';
        });
    } catch (err) {
      ErrorLogModule.error('RenewalTrackerModule', 'readAllItems_', 'Read failed', { error: err });
      return [];
    }
  }

  function readRemindableItems_() {
    return readAllItems_().filter(isValidItem_);
  }

  function readItemByRow_(rowIndex) {
    const spreadsheet = AppConfig.getSpreadsheet();
    const sheet = getRenewalSheet_(spreadsheet);
    if (!sheet || rowIndex < AppConfig.getDataStartRow()) {
      return null;
    }

    const rowValues = SpreadsheetUtils.getRowValues(sheet, rowIndex, SheetColumns.RENEWAL.LAST_COLUMN);
    const item = normalizeItemRow_(rowValues, rowIndex);
    return item.itemName ? item : null;
  }

  function writeNextDueDate_(rowIndex, nextDueDate) {
    const spreadsheet = AppConfig.getSpreadsheet();
    const sheet = getRenewalSheet_(spreadsheet);
    if (!sheet || rowIndex < AppConfig.getDataStartRow() || !nextDueDate) {
      return false;
    }

    try {
      sheet.getRange(rowIndex, SheetColumns.RENEWAL.NEXT_DUE_DATE).setValue(nextDueDate);
      return true;
    } catch (err) {
      ErrorLogModule.error('RenewalTrackerModule', 'writeNextDueDate_', 'Write failed', {
        rowIndex: rowIndex,
        error: err,
      });
      return false;
    }
  }

  function updateStatus_(rowIndex, status) {
    const spreadsheet = AppConfig.getSpreadsheet();
    const sheet = getRenewalSheet_(spreadsheet);
    if (!sheet || rowIndex < AppConfig.getDataStartRow()) {
      return false;
    }

    try {
      sheet.getRange(rowIndex, SheetColumns.RENEWAL.STATUS).setValue(status);
      return true;
    } catch (err) {
      ErrorLogModule.error('RenewalTrackerModule', 'updateStatus_', 'Write failed', {
        rowIndex: rowIndex,
        error: err,
      });
      return false;
    }
  }

  function validateSheetExists_() {
    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      return { valid: false, error: 'Spreadsheet unavailable' };
    }

    const sheet = getRenewalSheet_(spreadsheet);
    if (!sheet) {
      return {
        valid: false,
        error: 'Renewal Tracker sheet not found — create tab "' + AppConfig.getRenewalTrackerSheetName() + '"',
      };
    }

    return { valid: true };
  }

  return {
    readAllItems: readAllItems_,
    readRemindableItems: readRemindableItems_,
    readItemByRow: readItemByRow_,
    writeNextDueDate: writeNextDueDate_,
    updateStatus: updateStatus_,
    validateSheetExists: validateSheetExists_,
    parseDate: parseDate_,
    formatDate: formatDate_,
    isRemindableStatus: function (statusKey) {
      return REMINDABLE_STATUSES.indexOf(statusKey) >= 0;
    },
  };
})();
