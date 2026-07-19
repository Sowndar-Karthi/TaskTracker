/**
 * Creates and removes installable triggers for Team Project Tracker.
 */
const TriggerManager = (function () {
  'use strict';

  const HANDLERS = {
    SPREADSHEET_EDIT: 'handleSpreadsheetEdit_',
    SCHEDULED_COMPLETIONS: 'processScheduledCompletions_',
    SCHEDULED_NOTIFICATIONS: 'processScheduledNotifications_',
    SCHEDULED_RENEWAL_REMINDERS: 'processScheduledRenewalReminders_',
    SCHEDULED_IRCTC_IMPORT: 'processScheduledIrctcImport_',
    SCHEDULED_TRAIN_REMINDERS: 'processScheduledTrainReminders_',
    SCHEDULED_TRAIN_ARCHIVE: 'processScheduledTrainArchive_',
  };

  function removeAllProjectTriggers_() {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
  }

  function createSpreadsheetEditTrigger_(spreadsheetId) {
    if (!spreadsheetId) {
      throw new Error('TriggerManager: spreadsheetId is required');
    }
    ScriptApp.newTrigger(HANDLERS.SPREADSHEET_EDIT)
      .forSpreadsheet(spreadsheetId)
      .onEdit()
      .create();
  }

  function createHourlyCompletionTrigger_() {
    ScriptApp.newTrigger(HANDLERS.SCHEDULED_COMPLETIONS).timeBased().everyHours(1).create();
  }

  function createNotificationTrigger_() {
    ScriptApp.newTrigger(HANDLERS.SCHEDULED_NOTIFICATIONS).timeBased().everyMinutes(30).create();
  }

  function createRenewalReminderTrigger_() {
    ScriptApp.newTrigger(HANDLERS.SCHEDULED_RENEWAL_REMINDERS)
      .timeBased()
      .everyDays(1)
      .atHour(AppConfig.getRenewalReminderHour())
      .create();
  }

  function createIrctcImportTrigger_() {
    ScriptApp.newTrigger(HANDLERS.SCHEDULED_IRCTC_IMPORT).timeBased().everyHours(1).create();
  }

  function createTrainReminderTrigger_() {
    ScriptApp.newTrigger(HANDLERS.SCHEDULED_TRAIN_REMINDERS).timeBased().everyMinutes(30).create();
  }

  function createTrainArchiveTrigger_() {
    ScriptApp.newTrigger(HANDLERS.SCHEDULED_TRAIN_ARCHIVE)
      .timeBased()
      .everyDays(1)
      .atHour(AppConfig.getRenewalReminderHour())
      .create();
  }

  function listProjectTriggers_() {
    return ScriptApp.getProjectTriggers().map(function (trigger) {
      return {
        handler: trigger.getHandlerFunction(),
        eventType: trigger.getEventType(),
        sourceId: trigger.getTriggerSourceId(),
      };
    });
  }

  function initializeTriggers_() {
    const spreadsheetId = AppConfig.getSpreadsheetId() || SpreadsheetApp.getActiveSpreadsheet().getId();
    const environment = AppConfig.isTesting() ? 'TEST' : 'PRODUCTION';

    console.log('TriggerManager.initializeTriggers: starting in ' + environment, {
      spreadsheetId: spreadsheetId,
    });

    removeAllProjectTriggers_();
    createSpreadsheetEditTrigger_(spreadsheetId);
    createHourlyCompletionTrigger_();
    createNotificationTrigger_();
    createRenewalReminderTrigger_();
    createIrctcImportTrigger_();
    createTrainReminderTrigger_();
    createTrainArchiveTrigger_();

    console.log('TriggerManager.initializeTriggers: complete', {
      triggers: listProjectTriggers_(),
    });

    return listProjectTriggers_();
  }

  return {
    initializeTriggers: initializeTriggers_,
    removeAllTriggers: removeAllProjectTriggers_,
    listTriggers: listProjectTriggers_,
    HANDLERS: HANDLERS,
  };
})();
