/**
 * Scheduled / manual entry points for sheet-driven Gmail archive rules.
 */
const GmailArchiveRulesHandler = (function () {
  'use strict';

  const LOCK_TIMEOUT_MS = 30000;

  function processScheduledArchiveRules_() {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
      console.log('GmailArchiveRulesHandler: could not acquire lock');
      return { error: 'Lock timeout' };
    }

    try {
      ErrorLogModule.setTriggerSource('scheduled');
      GmailArchiveRulesModule.ensureRulesSheet();
      return GmailArchiveRulesModule.processAllRules();
    } catch (err) {
      console.error('GmailArchiveRulesHandler.processScheduledArchiveRules_ failed', {
        error: err,
        stack: err && err.stack,
      });
      ErrorLogModule.error('GmailArchiveRulesHandler', 'processScheduledArchiveRules_', String(err), {
        triggerSource: 'scheduled',
      });
      return { error: String(err) };
    } finally {
      ErrorLogModule.clearTriggerSource();
      lock.releaseLock();
    }
  }

  function runArchiveRulesNow_() {
    ErrorLogModule.setTriggerSource('manual');
    try {
      GmailArchiveRulesModule.ensureRulesSheet();
      const summary = GmailArchiveRulesModule.processAllRules();
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Gmail archive — rules:' +
            (summary.rules || 0) +
            ' labeled:' +
            (summary.labeled || 0) +
            ' archived:' +
            (summary.archived || 0),
          'Gmail Archive Rules',
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
    processScheduledArchiveRules: processScheduledArchiveRules_,
    runArchiveRulesNow: runArchiveRulesNow_,
  };
})();

/**
 * Installable daily trigger — processes 📧 Gmail Archive Rules.
 * Do not rename (wired in TriggerManager).
 */
function processScheduledGmailArchiveRules_() {
  return GmailArchiveRulesHandler.processScheduledArchiveRules();
}

/**
 * Manual run from Apps Script editor.
 */
function runGmailArchiveRulesNow() {
  return GmailArchiveRulesHandler.runArchiveRulesNow();
}
