/**
 * Scheduled / manual archive: Booking History → Train Completed Journeys.
 */
const TrainJourneyArchiveHandler = (function () {
  'use strict';

  const LOCK_TIMEOUT_MS = 30000;

  function processScheduledArchive_() {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
      console.log('TrainJourneyArchiveHandler: could not acquire lock');
      return { error: 'Lock timeout' };
    }

    try {
      ErrorLogModule.setTriggerSource('scheduled');
      return TrainJourneyArchiveModule.archivePastJourneys();
    } catch (err) {
      console.error('TrainJourneyArchiveHandler.processScheduledArchive_ failed', {
        error: err,
        stack: err && err.stack,
      });
      ErrorLogModule.error(
        'TrainJourneyArchiveHandler',
        'processScheduledArchive_',
        String(err),
        { triggerSource: 'scheduled' }
      );
      return { error: String(err) };
    } finally {
      ErrorLogModule.clearTriggerSource();
      lock.releaseLock();
    }
  }

  function runArchiveNow_() {
    ErrorLogModule.setTriggerSource('manual');
    try {
      const summary = TrainJourneyArchiveModule.archivePastJourneys();
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Archived ' +
            (summary.moved || 0) +
            ' journey(s) to ' +
            AppConfig.getTrainCompletedJourneysSheetName() +
            ' (skipped ' +
            (summary.skipped || 0) +
            ')',
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
    processScheduledArchive: processScheduledArchive_,
    runArchiveNow: runArchiveNow_,
  };
})();

/**
 * Installable daily trigger handler — do not rename.
 */
function processScheduledTrainArchive_() {
  return TrainJourneyArchiveHandler.processScheduledArchive();
}

/**
 * Manual run from Apps Script editor.
 */
function runArchiveTrainJourneysNow() {
  return TrainJourneyArchiveHandler.runArchiveNow();
}
