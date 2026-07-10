/**
 * Scheduled and on-edit handlers for Renewal Tracker reminders.
 */
const RenewalReminderHandler = (function () {
  'use strict';

  function assertSheetColumns_() {
    if (typeof SheetColumns === 'undefined') {
      throw new Error(
        'SheetColumns is not defined. Open SheetColumns.gs in Apps Script and paste the FULL file from your PC ' +
          '(must start with "const SheetColumns = (function () {" and include RENEWAL). ' +
          'Run debugWhatIsBroken() for a full checklist.'
      );
    }
    if (!SheetColumns.RENEWAL || SheetColumns.RENEWAL.LAST_COLUMN !== 14) {
      throw new Error(
        'SheetColumns.gs is incomplete — RENEWAL section missing or truncated. ' +
          'Paste the full 121-line SheetColumns.gs from your PC.'
      );
    }
  }

  function getRecalcTriggerColumns_() {
    assertSheetColumns_();
    return [
      SheetColumns.RENEWAL.LAST_PAID_DATE,
      SheetColumns.RENEWAL.RECURRENCE_INTERVAL,
      SheetColumns.RENEWAL.RECURRENCE_UNIT,
      SheetColumns.RENEWAL.ANCHOR_TYPE,
      SheetColumns.RENEWAL.ANCHOR_DATE,
      SheetColumns.RENEWAL.DUE_DAY_OF_MONTH,
    ];
  }

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
    return rangeTouchesColumns_(startCol, endCol, getRecalcTriggerColumns_());
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
  if (typeof SheetColumns === 'undefined') {
    const message =
      'SheetColumns is not defined — SheetColumns.gs is missing or incomplete in Apps Script. ' +
      'Paste the full file from your PC, then run debugWhatIsBroken().';
    Logger.log('RenewalReminderHandler: ' + message);
    try {
      ErrorLogModule.error('RenewalReminderHandler', 'processScheduledRenewalReminders_', message, {
        triggerSource: 'scheduled-renewal',
        fix: 'Paste full SheetColumns.gs (121 lines) into Apps Script',
      });
    } catch (logErr) {
      // ErrorLogModule may also be missing
    }
    return;
  }

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
