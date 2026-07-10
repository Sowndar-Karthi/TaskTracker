/**
 * Scheduled and on-edit handlers for Renewal Tracker reminders.
 */
const RenewalReminderHandler = (function () {
  'use strict';

  const RECALC_TRIGGER_COLUMNS = [
    SheetColumns.RENEWAL.LAST_PAID_DATE,
    SheetColumns.RENEWAL.RECURRENCE_INTERVAL,
    SheetColumns.RENEWAL.RECURRENCE_UNIT,
    SheetColumns.RENEWAL.ANCHOR_TYPE,
    SheetColumns.RENEWAL.ANCHOR_DATE,
    SheetColumns.RENEWAL.DUE_DAY_OF_MONTH,
  ];

  function isRenewalSheet_(sheetName) {
    return (
      SpreadsheetUtils.normalizeText(sheetName) ===
      SpreadsheetUtils.normalizeText(AppConfig.getRenewalTrackerSheetName())
    );
  }

  function rangeTouchesColumns_(startCol, endCol, columns) {
    for (var i = 0; i < columns.length; i++) {
      if (SpreadsheetUtils.columnRangeOverlaps(startCol, endCol, columns[i])) {
        return true;
      }
    }
    return false;
  }

  function shouldProcessEdit_(e) {
    if (!e || !e.range) {
      return false;
    }

    const sheet = e.range.getSheet();
    if (!sheet || !isRenewalSheet_(sheet.getName())) {
      return false;
    }

    if (e.range.getRow() <= AppConfig.getHeaderRow()) {
      return false;
    }

    const startCol = e.range.getColumn();
    const endCol = startCol + e.range.getNumColumns() - 1;
    return rangeTouchesColumns_(startCol, endCol, RECALC_TRIGGER_COLUMNS);
  }

  function processEdit_(e) {
    const sheet = e.range.getSheet();
    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();
    const results = [];

    for (var i = 0; i < numRows; i++) {
      const rowIndex = startRow + i;
      if (rowIndex <= AppConfig.getHeaderRow()) {
        continue;
      }

      try {
        const item = RenewalTrackerModule.readItemByRow(rowIndex);
        if (!item) {
          continue;
        }

        if (item.anchorTypeKey === 'manual only') {
          results.push({ rowIndex: rowIndex, skipped: true, reason: 'Manual only' });
          continue;
        }

        const result = RenewalDueDateModule.recalculateItem(item);
        if (result && result.success && item.lastPaidDate) {
          RenewalTrackerModule.updateStatus(rowIndex, 'Active');
        }
        results.push(result);
      } catch (err) {
        ErrorLogModule.error('RenewalReminderHandler', 'processEdit_', 'Row failed', {
          rowIndex: rowIndex,
          error: err,
        });
      }
    }

    return results;
  }

  function processScheduledReminders_() {
    return RenewalReminderModule.processDueReminders(new Date());
  }

  return {
    shouldProcessEdit: shouldProcessEdit_,
    processEdit: processEdit_,
    processScheduledReminders: processScheduledReminders_,
  };
})();

/**
 * Daily trigger — sends renewal reminders when Reminder Days Before matches today.
 */
function processScheduledRenewalReminders_() {
  ErrorLogModule.setTriggerSource('scheduled-renewal');
  const lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    hasLock = lock.tryLock(30000);
    if (!hasLock) {
      ErrorLogModule.warning(
        'RenewalReminderHandler',
        'processScheduledRenewalReminders_',
        'Could not acquire lock',
        { triggerSource: 'scheduled-renewal' }
      );
      return;
    }
    return RenewalReminderHandler.processScheduledReminders();
  } catch (err) {
    ErrorLogModule.error('RenewalReminderHandler', 'processScheduledRenewalReminders_', 'Failed', {
      error: err,
      triggerSource: 'scheduled-renewal',
    });
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
    ErrorLogModule.clearTriggerSource();
  }
}

/**
 * Manual run: send renewal reminders now (ignores daily schedule).
 */
function runRenewalRemindersNow() {
  return RenewalReminderModule.processDueReminders(new Date());
}

/**
 * Manual run: recalculate Next Due Date for all renewal rows.
 */
function runRecalculateRenewalDueDates() {
  return RenewalDueDateModule.recalculateAll();
}

/**
 * Manual run: recalculate Next Due Date for one renewal row.
 * @param {number} rowIndex
 */
function runRecalculateRenewalRow(rowIndex) {
  const item = RenewalTrackerModule.readItemByRow(rowIndex);
  if (!item) {
    return { success: false, error: 'Row not found' };
  }
  return RenewalDueDateModule.recalculateItem(item);
}
