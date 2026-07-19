/**
 * Scheduled / manual entry points for Train Route reminders.
 */
const TrainReminderHandler = (function () {
  'use strict';

  const LOCK_TIMEOUT_MS = 30000;

  function processScheduledReminders_() {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
      console.log('TrainReminderHandler: could not acquire lock');
      return { error: 'Lock timeout' };
    }

    try {
      ErrorLogModule.setTriggerSource('scheduled');
      return TrainReminderModule.processDueReminders();
    } catch (err) {
      console.error('TrainReminderHandler.processScheduledReminders_ failed', {
        error: err,
        stack: err && err.stack,
      });
      ErrorLogModule.error('TrainReminderHandler', 'processScheduledReminders_', String(err), {
        triggerSource: 'scheduled',
      });
      return { error: String(err) };
    } finally {
      ErrorLogModule.clearTriggerSource();
      lock.releaseLock();
    }
  }

  function runRemindersNow_() {
    ErrorLogModule.setTriggerSource('manual');
    try {
      const summary = TrainReminderModule.processDueReminders();
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Train reminders — digest(s):' +
            (summary.digestsSent || 0) +
            ' book:' +
            (summary.bookSent || 0) +
            ' travel:' +
            (summary.travelSent || 0) +
            ' missing:' +
            (summary.pendingListed || 0),
          'Train Tracker',
          8
        );
      } catch (uiErr) {
        // toast may be unavailable
      }
      return summary;
    } finally {
      ErrorLogModule.clearTriggerSource();
    }
  }

  return {
    processScheduledReminders: processScheduledReminders_,
    runRemindersNow: runRemindersNow_,
  };
})();

/**
 * Installable time-driven trigger — every 30 minutes (alert-time windows).
 * Do not rename (wired in TriggerManager).
 */
function processScheduledTrainReminders_() {
  return TrainReminderHandler.processScheduledReminders();
}

/**
 * Manual run from Apps Script editor.
 */
function runTrainRemindersNow() {
  return TrainReminderHandler.runRemindersNow();
}
