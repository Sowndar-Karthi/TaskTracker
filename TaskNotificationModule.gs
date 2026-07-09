/**
 * Finds tasks matching notification criteria and builds alert content.
 */
const TaskNotificationModule = (function () {
  'use strict';

  const TIMEZONE_ALIASES = {
    ist: 'Asia/Kolkata',
    utc: 'Etc/UTC',
    gmt: 'Etc/GMT',
    est: 'America/New_York',
    pst: 'America/Los_Angeles',
  };

  function resolveTimezone_(timezoneValue) {
    if (!timezoneValue) {
      return 'Asia/Kolkata';
    }
    const trimmed = String(timezoneValue).trim();
    if (trimmed.indexOf('/') > 0) {
      return trimmed;
    }
    const alias = TIMEZONE_ALIASES[SpreadsheetUtils.normalizeText(trimmed)];
    return alias || 'Asia/Kolkata';
  }

  function parseTimeToMinutes_(timeText) {
    if (!timeText) {
      return -1;
    }

    const value = String(timeText).trim().toUpperCase();
    const match12 = value.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/);
    if (match12) {
      var hours = parseInt(match12[1], 10);
      const minutes = parseInt(match12[2], 10);
      const meridiem = match12[3];
      if (meridiem === 'PM' && hours < 12) {
        hours += 12;
      }
      if (meridiem === 'AM' && hours === 12) {
        hours = 0;
      }
      return hours * 60 + minutes;
    }

    const match24 = value.match(/^(\d{1,2}):(\d{2})$/);
    if (match24) {
      return parseInt(match24[1], 10) * 60 + parseInt(match24[2], 10);
    }

    return -1;
  }

  function formatMinutesAsSlot_(minutes) {
    const hours = Math.floor(minutes / 60);
    const mins = minutes % 60;
    return (hours < 10 ? '0' : '') + hours + ':' + (mins < 10 ? '0' : '') + mins;
  }

  function isAlertTimeNow_(alertTimes, timezoneValue) {
    if (!alertTimes || alertTimes.length === 0) {
      return { due: false };
    }

    const timezone = resolveTimezone_(timezoneValue);
    const now = new Date();
    const currentMinutes = parseTimeToMinutes_(Utilities.formatDate(now, timezone, 'H:mm'));

    for (var i = 0; i < alertTimes.length; i++) {
      const alertMinutes = parseTimeToMinutes_(alertTimes[i]);
      if (alertMinutes < 0) {
        continue;
      }
      if (Math.abs(currentMinutes - alertMinutes) <= 15) {
        return {
          due: true,
          slot: formatMinutesAsSlot_(alertMinutes),
          timezone: timezone,
        };
      }
    }

    return { due: false };
  }

  function buildTaskInfoFromRow_(rowValues, sheetRole, rowIndex, sheetName) {
    if (!rowValues || rowValues.length === 0) {
      return null;
    }

    if (sheetRole === 'sub') {
      return {
        rowIndex: rowIndex,
        sheetName: sheetName,
        sheetRole: sheetRole,
        projectName: rowValues[SheetColumns.SUB.PROJECT_NAME - 1],
        taskName: rowValues[SheetColumns.SUB.TASK - 1],
        mainTaskName: rowValues[SheetColumns.SUB.MAIN_TASK_NAME - 1],
        status: rowValues[SheetColumns.SUB.STATUS - 1],
        priority: rowValues[SheetColumns.SUB.PRIORITY - 1],
        assign: rowValues[SheetColumns.SUB.ASSIGN - 1],
        endDate: rowValues[SheetColumns.SUB.END_DATE - 1],
      };
    }

    return {
      rowIndex: rowIndex,
      sheetName: sheetName,
      sheetRole: sheetRole || 'main',
      projectName: rowValues[SheetColumns.MAIN.PROJECT_NAME - 1],
      taskName: rowValues[SheetColumns.MAIN.TASK - 1],
      mainTaskName: rowValues[SheetColumns.MAIN.TASK - 1],
      status: rowValues[SheetColumns.MAIN.STATUS - 1],
      priority: rowValues[SheetColumns.MAIN.PRIORITY - 1],
      assign: rowValues[SheetColumns.MAIN.ASSIGN - 1],
      endDate: rowValues[SheetColumns.MAIN.END_DATE - 1],
    };
  }

  function getSheetRoleForDataSheet_(sheetName) {
    const sheetInfo = MatchingSheetModule.getSheetRole(sheetName);
    if (!sheetInfo) {
      return { role: 'main', mapping: null };
    }
    return sheetInfo;
  }

  function findMatchingTasks_(config) {
    if (!config || !config.dataSheetName) {
      return [];
    }

    const spreadsheet = AppConfig.getSpreadsheet();
    const dataSheet = SpreadsheetUtils.getSheetByName(spreadsheet, config.dataSheetName);
    if (!dataSheet) {
      console.error('TaskNotificationModule.findMatchingTasks_: data sheet not found', {
        dataSheetName: config.dataSheetName,
      });
      return [];
    }

    const sheetInfo = getSheetRoleForDataSheet_(config.dataSheetName);
    const sheetRole = sheetInfo.role === 'sub' ? 'sub' : 'main';
    const lastColumn = sheetRole === 'sub' ? SheetColumns.SUB.LAST_COLUMN : SheetColumns.MAIN.LAST_COLUMN;
    const statusColumn = sheetRole === 'sub' ? SheetColumns.SUB.STATUS : SheetColumns.MAIN.STATUS;
    const lastRow = SpreadsheetUtils.getLastDataRow(dataSheet, statusColumn);
    const matches = [];

    for (var rowIndex = AppConfig.getDataStartRow(); rowIndex <= lastRow; rowIndex++) {
      const rowValues = SpreadsheetUtils.getRowValues(dataSheet, rowIndex, lastColumn);
      const taskInfo = buildTaskInfoFromRow_(rowValues, sheetRole, rowIndex, config.dataSheetName);
      if (!taskInfo || !taskInfo.taskName) {
        continue;
      }
      if (NotificationTypeModule.taskMatchesNotificationTypes(taskInfo, config.notificationTypes)) {
        matches.push(taskInfo);
      }
    }

    return matches;
  }

  function buildDigestHtml_(config, tasks) {
    var rowsHtml = tasks
      .map(function (task) {
        return (
          '<tr>' +
          '<td>' +
          escapeHtml_(task.projectName) +
          '</td>' +
          '<td>' +
          escapeHtml_(task.taskName) +
          '</td>' +
          '<td>' +
          escapeHtml_(task.status) +
          '</td>' +
          '<td>' +
          escapeHtml_(task.priority) +
          '</td>' +
          '<td>' +
          escapeHtml_(task.assign) +
          '</td>' +
          '<td>' +
          escapeHtml_(task.endDate) +
          '</td>' +
          '</tr>'
        );
      })
      .join('');

    return (
      '<p><strong>Team Project Tracker</strong> — scheduled alert for <em>' +
      escapeHtml_(config.dataSheetName) +
      '</em></p>' +
      '<p>Tasks matching notification types: <strong>' +
      escapeHtml_(config.notificationTypes.join(', ')) +
      '</strong></p>' +
      '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
      '<thead><tr><th>Project</th><th>Task</th><th>Status</th><th>Priority</th><th>Assign</th><th>End Date</th></tr></thead>' +
      '<tbody>' +
      rowsHtml +
      '</tbody></table>'
    );
  }

  function buildImmediateHtml_(config, task, matchedBy) {
    return (
      '<p><strong>Team Project Tracker</strong> — task alert for <em>' +
      escapeHtml_(config.dataSheetName) +
      '</em></p>' +
      '<p>A task now matches your notification rule (<strong>' +
      escapeHtml_(matchedBy) +
      '</strong>).</p>' +
      '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
      '<tr><th align="left">Project</th><td>' +
      escapeHtml_(task.projectName) +
      '</td></tr>' +
      '<tr><th align="left">Task</th><td>' +
      escapeHtml_(task.taskName) +
      '</td></tr>' +
      '<tr><th align="left">Status</th><td>' +
      escapeHtml_(task.status) +
      '</td></tr>' +
      '<tr><th align="left">Priority</th><td>' +
      escapeHtml_(task.priority) +
      '</td></tr>' +
      '<tr><th align="left">Assign</th><td>' +
      escapeHtml_(task.assign) +
      '</td></tr>' +
      '<tr><th align="left">End Date</th><td>' +
      escapeHtml_(task.endDate) +
      '</td></tr>' +
      '</table>'
    );
  }

  function escapeHtml_(value) {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function getMatchedTypeForTask_(taskInfo, notificationTypes) {
    for (var i = 0; i < notificationTypes.length; i++) {
      const type = notificationTypes[i];
      if (NotificationTypeModule.isPriorityType(type) && SpreadsheetUtils.normalizeText(taskInfo.priority) === SpreadsheetUtils.normalizeText(type)) {
        return type;
      }
      if (TaskStatusModule.statusesMatch(taskInfo.status, type)) {
        return type;
      }
    }
    return '';
  }

  function wasScheduledAlertSent_(configRow, slot, dateKey) {
    const key = 'notif_sent_' + configRow + '_' + dateKey + '_' + slot;
    return PropertiesService.getScriptProperties().getProperty(key) === '1';
  }

  function markScheduledAlertSent_(configRow, slot, dateKey) {
    const key = 'notif_sent_' + configRow + '_' + dateKey + '_' + slot;
    PropertiesService.getScriptProperties().setProperty(key, '1');
  }

  function wasImmediateAlertSentRecently_(sheetName, rowIndex) {
    const key = 'notif_immediate_' + SpreadsheetUtils.normalizeText(sheetName) + '_' + rowIndex;
    const lastSent = PropertiesService.getScriptProperties().getProperty(key);
    if (!lastSent) {
      return false;
    }
    const elapsed = Date.now() - parseInt(lastSent, 10);
    return elapsed < 5 * 60 * 1000;
  }

  function markImmediateAlertSent_(sheetName, rowIndex) {
    const key = 'notif_immediate_' + SpreadsheetUtils.normalizeText(sheetName) + '_' + rowIndex;
    PropertiesService.getScriptProperties().setProperty(key, String(Date.now()));
  }

  return {
    resolveTimezone: resolveTimezone_,
    isAlertTimeNow: isAlertTimeNow_,
    findMatchingTasks: findMatchingTasks_,
    buildDigestHtml: buildDigestHtml_,
    buildImmediateHtml: buildImmediateHtml_,
    buildTaskInfoFromRow: buildTaskInfoFromRow_,
    getMatchedTypeForTask: getMatchedTypeForTask_,
    wasScheduledAlertSent: wasScheduledAlertSent_,
    markScheduledAlertSent: markScheduledAlertSent_,
    wasImmediateAlertSentRecently: wasImmediateAlertSentRecently_,
    markImmediateAlertSent: markImmediateAlertSent_,
  };
})();
