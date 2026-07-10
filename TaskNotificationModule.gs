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
        mainTaskName: rowValues[SheetColumns.SUB.MAIN_TASK_NAME - 1],
        ticketId: rowValues[SheetColumns.SUB.TICKET_ID - 1],
        taskName: rowValues[SheetColumns.SUB.TASK - 1],
        status: rowValues[SheetColumns.SUB.STATUS - 1],
        priority: rowValues[SheetColumns.SUB.PRIORITY - 1],
        assign: rowValues[SheetColumns.SUB.ASSIGN - 1],
        endDate: rowValues[SheetColumns.SUB.END_DATE - 1],
        account: rowValues[SheetColumns.SUB.ACCOUNT - 1],
        supportBlocker: rowValues[SheetColumns.SUB.SUPPORT_BLOCKER - 1],
        latestUpdate: rowValues[SheetColumns.SUB.LATEST_UPDATE - 1],
        functionality: rowValues[SheetColumns.SUB.FUNCTIONALITY - 1],
      };
    }

    return {
      rowIndex: rowIndex,
      sheetName: sheetName,
      sheetRole: sheetRole || 'main',
      projectName: rowValues[SheetColumns.MAIN.PROJECT_NAME - 1],
      mainTaskName: rowValues[SheetColumns.MAIN.TASK - 1],
      taskName: rowValues[SheetColumns.MAIN.TASK - 1],
      status: rowValues[SheetColumns.MAIN.STATUS - 1],
      priority: rowValues[SheetColumns.MAIN.PRIORITY - 1],
      assign: rowValues[SheetColumns.MAIN.ASSIGN - 1],
      endDate: rowValues[SheetColumns.MAIN.END_DATE - 1],
      supportBlocker: rowValues[SheetColumns.MAIN.SUPPORT_BLOCKER - 1],
      latestUpdate: rowValues[SheetColumns.MAIN.LATEST_UPDATE - 1],
      functionality: rowValues[SheetColumns.MAIN.FUNCTIONALITY - 1],
    };
  }

  function isValidTaskForNotification_(taskInfo) {
    if (!taskInfo) {
      return false;
    }
    if (taskInfo.sheetRole === 'sub') {
      return !!(taskInfo.taskName || taskInfo.mainTaskName);
    }
    return !!taskInfo.taskName;
  }

  function getDigestColumns_(sheetRole) {
    if (sheetRole === 'sub') {
      return [
        { key: 'projectName', label: 'Project' },
        { key: 'mainTaskName', label: 'Main Task' },
        { key: 'taskName', label: 'Sub Task' },
        { key: 'ticketId', label: 'Ticket Id' },
        { key: 'account', label: 'Account' },
        { key: 'status', label: 'Status' },
        { key: 'priority', label: 'Priority' },
        { key: 'assign', label: 'Assign' },
        { key: 'supportBlocker', label: 'Support / Blocker' },
        { key: 'latestUpdate', label: 'Latest Update' },
        { key: 'functionality', label: 'Functionality' },
        { key: 'endDate', label: 'End Date' },
      ];
    }

    return [
      { key: 'projectName', label: 'Project' },
      { key: 'taskName', label: 'Task' },
      { key: 'status', label: 'Status' },
      { key: 'priority', label: 'Priority' },
      { key: 'assign', label: 'Assign' },
      { key: 'supportBlocker', label: 'Support / Blocker' },
      { key: 'latestUpdate', label: 'Latest Update' },
      { key: 'functionality', label: 'Functionality' },
      { key: 'endDate', label: 'End Date' },
    ];
  }

  /** Sub digest grouped by parent — Main Task shown as section header, not repeated per row. */
  function getGroupedSubDigestColumns_() {
    return [
      { key: 'projectName', label: 'Project' },
      { key: 'taskName', label: 'Sub Task' },
      { key: 'ticketId', label: 'Ticket Id' },
      { key: 'account', label: 'Account' },
      { key: 'status', label: 'Status' },
      { key: 'priority', label: 'Priority' },
      { key: 'assign', label: 'Assign' },
      { key: 'supportBlocker', label: 'Support / Blocker' },
      { key: 'latestUpdate', label: 'Latest Update' },
      { key: 'functionality', label: 'Functionality' },
      { key: 'endDate', label: 'End Date' },
    ];
  }

  function getParentKeyFromTask_(task) {
    if (!task) {
      return '';
    }
    const parentName = task.mainTaskName || task.taskName || '';
    return SpreadsheetUtils.normalizeText(parentName) || '(no parent task)';
  }

  function getParentLabelFromTask_(task) {
    if (!task) {
      return '(No parent task)';
    }
    const parentName = task.mainTaskName || task.taskName;
    if (!parentName || String(parentName).trim() === '') {
      return '(No parent task)';
    }
    return String(parentName).trim();
  }

  function groupSubTasksByParent_(tasks) {
    const groups = [];
    const groupMap = {};

    tasks.forEach(function (task) {
      const key = getParentKeyFromTask_(task);
      if (!groupMap[key]) {
        groupMap[key] = {
          parentKey: key,
          parentLabel: getParentLabelFromTask_(task),
          tasks: [],
        };
        groups.push(groupMap[key]);
      }
      groupMap[key].tasks.push(task);
    });

    return groups;
  }

  function chunkArray_(items, chunkSize) {
    const chunks = [];
    if (!items || items.length === 0 || chunkSize < 1) {
      return chunks;
    }
    for (var i = 0; i < items.length; i += chunkSize) {
      chunks.push(items.slice(i, i + chunkSize));
    }
    return chunks;
  }

  /**
   * Splits digest tasks into email batches. Sub sheets are grouped by parent;
   * a single parent with 300 subs becomes 3 emails of 100 rows each.
   */
  function planDigestBatches_(tasks, sheetRole, maxRowsPerEmail) {
    if (!tasks || tasks.length === 0) {
      return [];
    }

    const maxRows = maxRowsPerEmail || AppConfig.getDigestMaxRowsPerEmail();
    const batches = [];

    if (sheetRole !== 'sub') {
      chunkArray_(tasks, maxRows).forEach(function (batchTasks, index) {
        batches.push({
          batchNumber: index + 1,
          totalBatches: 0,
          tasks: batchTasks,
          rowCount: batchTasks.length,
          groupedByParent: false,
          parentGroups: null,
        });
      });
      batches.forEach(function (batch) {
        batch.totalBatches = batches.length;
      });
      return batches;
    }

    const parentGroups = groupSubTasksByParent_(tasks);
    var currentBatch = {
      batchNumber: 0,
      totalBatches: 0,
      tasks: [],
      rowCount: 0,
      groupedByParent: true,
      parentGroups: [],
    };

    function flushCurrentBatch_() {
      if (currentBatch.rowCount === 0) {
        return;
      }
      batches.push(currentBatch);
      currentBatch = {
        batchNumber: 0,
        totalBatches: 0,
        tasks: [],
        rowCount: 0,
        groupedByParent: true,
        parentGroups: [],
      };
    }

    parentGroups.forEach(function (group) {
      if (group.tasks.length > maxRows) {
        flushCurrentBatch_();
        const groupChunks = chunkArray_(group.tasks, maxRows);
        groupChunks.forEach(function (chunkTasks, chunkIndex) {
          batches.push({
            batchNumber: 0,
            totalBatches: 0,
            tasks: chunkTasks.slice(),
            rowCount: chunkTasks.length,
            groupedByParent: true,
            parentGroups: [
              {
                parentKey: group.parentKey,
                parentLabel: group.parentLabel,
                tasks: chunkTasks,
                partNumber: chunkIndex + 1,
                totalParts: groupChunks.length,
              },
            ],
          });
        });
        return;
      }

      if (currentBatch.rowCount > 0 && currentBatch.rowCount + group.tasks.length > maxRows) {
        flushCurrentBatch_();
      }

      currentBatch.tasks = currentBatch.tasks.concat(group.tasks);
      currentBatch.rowCount += group.tasks.length;
      currentBatch.parentGroups.push({
        parentKey: group.parentKey,
        parentLabel: group.parentLabel,
        tasks: group.tasks,
        partNumber: 1,
        totalParts: 1,
      });
    });

    flushCurrentBatch_();

    batches.forEach(function (batch, index) {
      batch.batchNumber = index + 1;
      batch.totalBatches = batches.length;
    });

    return batches;
  }

  function findMatchingSubTasksForParent_(config, mainTaskName) {
    if (!config || !mainTaskName) {
      return [];
    }
    const normalizedParent = SpreadsheetUtils.normalizeText(mainTaskName);
    return findMatchingTasks_(config).filter(function (task) {
      return task.sheetRole === 'sub' && getParentKeyFromTask_(task) === normalizedParent;
    });
  }

  function getImmediateFields_(task) {
    const sheetRole = task && task.sheetRole === 'sub' ? 'sub' : 'main';
    const columns = getDigestColumns_(sheetRole);
    return columns.map(function (column) {
      return { label: column.label, value: task[column.key] };
    });
  }

  function buildTableHeaderHtml_(columns) {
    return (
      '<thead><tr>' +
      columns
        .map(function (column) {
          return '<th>' + escapeHtml_(column.label) + '</th>';
        })
        .join('') +
      '</tr></thead>'
    );
  }

  function buildDigestRowHtml_(task, columns) {
    return (
      '<tr>' +
      columns
        .map(function (column) {
          return '<td>' + escapeHtml_(task[column.key]) + '</td>';
        })
        .join('') +
      '</tr>'
    );
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
      ErrorLogModule.error('TaskNotificationModule', 'findMatchingTasks_', 'Data sheet not found', {
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
      if (!isValidTaskForNotification_(taskInfo)) {
        continue;
      }
      if (NotificationTypeModule.taskMatchesNotificationTypes(taskInfo, config.notificationTypes)) {
        matches.push(taskInfo);
      }
    }

    return matches;
  }

  function buildDigestIntroHtml_(config, sheetRole, batch) {
    const sheetTypeLabel = sheetRole === 'sub' ? 'sub-task sheet' : 'main task sheet';
    var intro =
      '<p><strong>Team Project Tracker</strong> — scheduled alert for <em>' +
      escapeHtml_(config.dataSheetName) +
      '</em> (' +
      sheetTypeLabel +
      ')</p>' +
      '<p>Tasks matching notification types: <strong>' +
      escapeHtml_(config.notificationTypes.join(', ')) +
      '</strong></p>';

    if (batch && batch.totalBatches > 1) {
      intro +=
        '<p><strong>Email ' +
        batch.batchNumber +
        ' of ' +
        batch.totalBatches +
        '</strong> — showing ' +
        batch.rowCount +
        ' task row(s) in this message.</p>';
    }

    if (sheetRole === 'sub') {
      intro += '<p>Sub-tasks are <strong>grouped by parent main task</strong> below.</p>';
    }

    return intro;
  }

  function buildGroupedSubDigestBodyHtml_(parentGroups) {
    const columns = getGroupedSubDigestColumns_();
    return parentGroups
      .map(function (group) {
        var heading = escapeHtml_(group.parentLabel);
        if (group.totalParts > 1) {
          heading += ' (part ' + group.partNumber + ' of ' + group.totalParts + ')';
        }
        const rowsHtml = group.tasks
          .map(function (task) {
            return buildDigestRowHtml_(task, columns);
          })
          .join('');

        return (
          '<h3 style="margin:18px 0 8px;">Main Task: ' +
          heading +
          ' <span style="font-weight:normal;">(' +
          group.tasks.length +
          ' sub-task(s))</span></h3>' +
          '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;margin-bottom:12px;">' +
          buildTableHeaderHtml_(columns) +
          '<tbody>' +
          rowsHtml +
          '</tbody></table>'
        );
      })
      .join('');
  }

  function buildDigestHtml_(config, tasks, batch) {
    const sheetRole =
      tasks && tasks.length > 0 && tasks[0].sheetRole === 'sub' ? 'sub' : 'main';

    if (sheetRole === 'sub') {
      const parentGroups =
        batch && batch.parentGroups && batch.parentGroups.length > 0
          ? batch.parentGroups
          : groupSubTasksByParent_(tasks);
      return buildDigestIntroHtml_(config, sheetRole, batch) + buildGroupedSubDigestBodyHtml_(parentGroups);
    }

    const columns = getDigestColumns_(sheetRole);
    const rowsHtml = tasks
      .map(function (task) {
        return buildDigestRowHtml_(task, columns);
      })
      .join('');

    return (
      buildDigestIntroHtml_(config, sheetRole, batch) +
      '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
      buildTableHeaderHtml_(columns) +
      '<tbody>' +
      rowsHtml +
      '</tbody></table>'
    );
  }

  function buildImmediateGroupedSubHtml_(config, parentLabel, tasks, matchedBy) {
    const parentGroups = [
      {
        parentLabel: parentLabel,
        tasks: tasks,
        totalParts: 1,
        partNumber: 1,
      },
    ];

    return (
      '<p><strong>Team Project Tracker</strong> — sub-task alert for <em>' +
      escapeHtml_(config.dataSheetName) +
      '</em></p>' +
      '<p>Parent task <strong>' +
      escapeHtml_(parentLabel) +
      '</strong> has <strong>' +
      tasks.length +
      '</strong> matching sub-task(s) for rule (<strong>' +
      escapeHtml_(matchedBy) +
      '</strong>).</p>' +
      buildGroupedSubDigestBodyHtml_(parentGroups)
    );
  }

  function buildImmediateHtml_(config, task, matchedBy) {
    const fields = getImmediateFields_(task);
    const sheetTypeLabel = task.sheetRole === 'sub' ? 'sub-task' : 'main task';

    return (
      '<p><strong>Team Project Tracker</strong> — ' +
      sheetTypeLabel +
      ' alert for <em>' +
      escapeHtml_(config.dataSheetName) +
      '</em></p>' +
      '<p>This row now matches your notification rule (<strong>' +
      escapeHtml_(matchedBy) +
      '</strong>).</p>' +
      '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
      fields
        .map(function (field) {
          return (
            '<tr><th align="left">' +
            escapeHtml_(field.label) +
            '</th><td>' +
            escapeHtml_(field.value) +
            '</td></tr>'
          );
        })
        .join('') +
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

  function wasImmediateAlertSentRecently_(sheetName, rowIndex, parentKey) {
    if (parentKey) {
      const parentCooldownKey =
        'notif_immediate_parent_' + SpreadsheetUtils.normalizeText(sheetName) + '_' + parentKey;
      const parentLastSent = PropertiesService.getScriptProperties().getProperty(parentCooldownKey);
      if (parentLastSent) {
        const parentElapsed = Date.now() - parseInt(parentLastSent, 10);
        if (parentElapsed < 5 * 60 * 1000) {
          return true;
        }
      }
    }

    const key = 'notif_immediate_' + SpreadsheetUtils.normalizeText(sheetName) + '_' + rowIndex;
    const lastSent = PropertiesService.getScriptProperties().getProperty(key);
    if (!lastSent) {
      return false;
    }
    const elapsed = Date.now() - parseInt(lastSent, 10);
    return elapsed < 5 * 60 * 1000;
  }

  function markImmediateAlertSent_(sheetName, rowIndex, parentKey) {
    const now = String(Date.now());
    if (parentKey) {
      const parentCooldownKey =
        'notif_immediate_parent_' + SpreadsheetUtils.normalizeText(sheetName) + '_' + parentKey;
      PropertiesService.getScriptProperties().setProperty(parentCooldownKey, now);
    }
    const key = 'notif_immediate_' + SpreadsheetUtils.normalizeText(sheetName) + '_' + rowIndex;
    PropertiesService.getScriptProperties().setProperty(key, now);
  }

  function getSheetRoleForTasks_(tasks) {
    if (!tasks || tasks.length === 0) {
      return 'main';
    }
    return tasks[0].sheetRole === 'sub' ? 'sub' : 'main';
  }

  return {
    resolveTimezone: resolveTimezone_,
    isAlertTimeNow: isAlertTimeNow_,
    findMatchingTasks: findMatchingTasks_,
    findMatchingSubTasksForParent: findMatchingSubTasksForParent_,
    planDigestBatches: planDigestBatches_,
    groupSubTasksByParent: groupSubTasksByParent_,
    buildDigestHtml: buildDigestHtml_,
    buildImmediateHtml: buildImmediateHtml_,
    buildImmediateGroupedSubHtml: buildImmediateGroupedSubHtml_,
    buildTaskInfoFromRow: buildTaskInfoFromRow_,
    getMatchedTypeForTask: getMatchedTypeForTask_,
    getParentKeyFromTask: getParentKeyFromTask_,
    getParentLabelFromTask: getParentLabelFromTask_,
    wasScheduledAlertSent: wasScheduledAlertSent_,
    markScheduledAlertSent: markScheduledAlertSent_,
    wasImmediateAlertSentRecently: wasImmediateAlertSentRecently_,
    markImmediateAlertSent: markImmediateAlertSent_,
    getSheetRoleForTasks: getSheetRoleForTasks_,
  };
})();
