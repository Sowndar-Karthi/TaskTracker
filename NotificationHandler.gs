/**
 * Scheduled and immediate task notification handlers.
 */
const NotificationHandler = (function () {
  'use strict';

  function processScheduledAlerts_() {
    const settings = EmailSettingsModule.readAllSettings();
    const results = [];
    const timezoneDefault = 'Asia/Kolkata';

    settings.forEach(function (config) {
      try {
        const alertCheck = TaskNotificationModule.isAlertTimeNow(config.alertTimes, config.timezone);
        if (!alertCheck.due) {
          results.push({ configRow: config.rowNumber, skipped: true, reason: 'Not alert time' });
          return;
        }

        const timezone = alertCheck.timezone || timezoneDefault;
        const dateKey = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');

        if (TaskNotificationModule.wasScheduledAlertSent(config.rowNumber, alertCheck.slot, dateKey)) {
          results.push({
            configRow: config.rowNumber,
            skipped: true,
            reason: 'Already sent for slot',
            slot: alertCheck.slot,
          });
          return;
        }

        const tasks = TaskNotificationModule.findMatchingTasks(config);
        const sendResult = NotificationService.sendScheduledDigest(config, tasks);

        if (sendResult && sendResult.success && !sendResult.skipped) {
          TaskNotificationModule.markScheduledAlertSent(config.rowNumber, alertCheck.slot, dateKey);
        }

        results.push({
          configRow: config.rowNumber,
          dataSheetName: config.dataSheetName,
          slot: alertCheck.slot,
          taskCount: tasks.length,
          send: sendResult,
        });
      } catch (err) {
        console.error('NotificationHandler.processScheduledAlerts_: config failed', {
          configRow: config.rowNumber,
          error: err,
        });
        results.push({ configRow: config.rowNumber, success: false, error: String(err) });
      }
    });

    return { processedCount: results.length, results: results };
  }

  function getStatusColumnForRole_(sheetRole) {
    return sheetRole === 'sub' ? SheetColumns.SUB.STATUS : SheetColumns.MAIN.STATUS;
  }

  function getPriorityColumnForRole_(sheetRole) {
    return sheetRole === 'sub' ? SheetColumns.SUB.PRIORITY : SheetColumns.MAIN.PRIORITY;
  }

  function shouldProcessEdit_(e) {
    if (!e || !e.range) {
      return false;
    }

    const sheet = e.range.getSheet();
    if (!sheet) {
      return false;
    }

    const sheetName = sheet.getName();
    const configs = EmailSettingsModule.getSettingsForDataSheet(sheetName);
    if (!configs || configs.length === 0) {
      return false;
    }

    const sheetInfo = MatchingSheetModule.getSheetRole(sheetName);
    const sheetRole = sheetInfo && sheetInfo.role === 'sub' ? 'sub' : 'main';

    const startCol = e.range.getColumn();
    const endCol = startCol + e.range.getNumColumns() - 1;
    const statusCol = getStatusColumnForRole_(sheetRole);
    const priorityCol = getPriorityColumnForRole_(sheetRole);

    if (
      !SpreadsheetUtils.columnRangeOverlaps(startCol, endCol, statusCol) &&
      !SpreadsheetUtils.columnRangeOverlaps(startCol, endCol, priorityCol)
    ) {
      return false;
    }

    if (e.range.getRow() <= AppConfig.getHeaderRow()) {
      return false;
    }

    return true;
  }

  function processEdit_(e) {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const configs = EmailSettingsModule.getSettingsForDataSheet(sheetName);
    const sheetInfo = MatchingSheetModule.getSheetRole(sheetName);
    const sheetRole = sheetInfo && sheetInfo.role === 'sub' ? 'sub' : 'main';
    const lastColumn = sheetRole === 'sub' ? SheetColumns.SUB.LAST_COLUMN : SheetColumns.MAIN.LAST_COLUMN;
    const results = [];

    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();

    for (var i = 0; i < numRows; i++) {
      const rowIndex = startRow + i;
      if (rowIndex <= AppConfig.getHeaderRow()) {
        continue;
      }

      if (TaskNotificationModule.wasImmediateAlertSentRecently(sheetName, rowIndex)) {
        continue;
      }

      const rowValues = SpreadsheetUtils.getRowValues(sheet, rowIndex, lastColumn);
      const taskInfo = TaskNotificationModule.buildTaskInfoFromRow(
        rowValues,
        sheetRole,
        rowIndex,
        sheetName
      );

      if (!taskInfo) {
        continue;
      }

      configs.forEach(function (config) {
        try {
          if (!NotificationTypeModule.taskMatchesNotificationTypes(taskInfo, config.notificationTypes)) {
            return;
          }

          const matchedBy = TaskNotificationModule.getMatchedTypeForTask(taskInfo, config.notificationTypes);
          const sendResult = NotificationService.sendImmediateAlert(config, taskInfo, matchedBy);

          if (sendResult && sendResult.success) {
            TaskNotificationModule.markImmediateAlertSent(sheetName, rowIndex);
          }

          results.push({
            rowIndex: rowIndex,
            configRow: config.rowNumber,
            matchedBy: matchedBy,
            send: sendResult,
          });
        } catch (err) {
          console.error('NotificationHandler.processEdit_: row failed', {
            sheetName: sheetName,
            rowIndex: rowIndex,
            error: err,
          });
        }
      });
    }

    return results;
  }

  return {
    processScheduledAlerts: processScheduledAlerts_,
    shouldProcessEdit: shouldProcessEdit_,
    processEdit: processEdit_,
  };
})();

/**
 * Time-driven entry — runs every 30 minutes; sends digest when Alert Time matches.
 */
function processScheduledNotifications_() {
  const lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    hasLock = lock.tryLock(30000);
    if (!hasLock) {
      console.error('processScheduledNotifications_: could not acquire lock');
      return;
    }
    return NotificationHandler.processScheduledAlerts();
  } catch (err) {
    console.error('processScheduledNotifications_: failed', { error: err });
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Manual run: send scheduled digest now (ignores alert time window).
 */
function runSendNotificationDigestNow() {
  const settings = EmailSettingsModule.readAllSettings();
  const results = [];

  settings.forEach(function (config) {
    const tasks = TaskNotificationModule.findMatchingTasks(config);
    results.push({
      config: config.dataSheetName,
      taskCount: tasks.length,
      send: NotificationService.sendScheduledDigest(config, tasks),
    });
  });

  return results;
}

/**
 * Manual run: test notification type parsing.
 * @param {string} rawValue Raw Notification Type cell value
 */
function runParseNotificationTypes(rawValue) {
  return NotificationTypeModule.parseNotificationTypes(rawValue);
}
