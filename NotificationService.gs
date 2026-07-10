/**
 * Sends email notifications for task status and priority alerts.
 */
const NotificationService = (function () {
  'use strict';

  function sendEmail_(config, subject, htmlBody) {
    if (!config || !config.toEmails || config.toEmails.length === 0) {
      return { success: false, error: 'No recipients configured' };
    }

    try {
      const options = {
        htmlBody: htmlBody,
        name: 'Team Project Tracker',
      };

      if (config.ccEmails && config.ccEmails.length > 0) {
        options.cc = config.ccEmails.join(',');
      }
      if (config.bccEmails && config.bccEmails.length > 0) {
        options.bcc = config.bccEmails.join(',');
      }

      GmailApp.sendEmail(config.toEmails.join(','), subject, '', options);

      console.log('NotificationService.sendEmail_: sent', {
        subject: subject,
        to: config.toEmails,
        dataSheetName: config.dataSheetName,
      });

      return { success: true, subject: subject };
    } catch (err) {
      ErrorLogModule.error('NotificationService', 'sendEmail_', 'Send failed', {
        subject: subject,
        error: err,
      });
      return { success: false, error: String(err) };
    }
  }

  function buildDigestSubject_(config, tasks, batch) {
    const sheetRole = TaskNotificationModule.getSheetRoleForTasks(tasks);
    const rowCount = batch ? batch.rowCount : tasks.length;
    var subject =
      'Team Project Tracker Alert — ' + config.dataSheetName + ' (' + rowCount + ' task(s))';

    if (batch && batch.totalBatches > 1) {
      subject += ' — part ' + batch.batchNumber + ' of ' + batch.totalBatches;
    }

    if (sheetRole === 'sub' && batch && batch.parentGroups && batch.parentGroups.length === 1) {
      const group = batch.parentGroups[0];
      if (group.parentLabel) {
        subject += ' — ' + group.parentLabel;
        if (group.totalParts > 1) {
          subject += ' (sub-tasks ' + group.partNumber + '/' + group.totalParts + ')';
        }
      }
    }

    return subject;
  }

  function sendScheduledDigest_(config, tasks) {
    if (!tasks || tasks.length === 0) {
      return { success: true, skipped: true, reason: 'No matching tasks' };
    }

    const sheetRole = TaskNotificationModule.getSheetRoleForTasks(tasks);
    const batches = TaskNotificationModule.planDigestBatches(
      tasks,
      sheetRole,
      AppConfig.getDigestMaxRowsPerEmail()
    );
    const sendResults = [];

    batches.forEach(function (batch) {
      const subject = buildDigestSubject_(config, batch.tasks, batch);
      const htmlBody = TaskNotificationModule.buildDigestHtml(config, batch.tasks, batch);
      sendResults.push({
        batchNumber: batch.batchNumber,
        totalBatches: batch.totalBatches,
        rowCount: batch.rowCount,
        send: sendEmail_(config, subject, htmlBody),
      });
    });

    const failed = sendResults.filter(function (result) {
      return !result.send || !result.send.success;
    });

    if (failed.length > 0) {
      return {
        success: false,
        emailCount: sendResults.length,
        sentCount: sendResults.length - failed.length,
        failedCount: failed.length,
        batches: sendResults,
        error: failed[0].send ? failed[0].send.error : 'Send failed',
      };
    }

    return {
      success: true,
      emailCount: sendResults.length,
      batches: sendResults,
      taskCount: tasks.length,
    };
  }

  function sendImmediateAlert_(config, task, matchedBy, relatedTasks) {
    const isGroupedSub =
      task &&
      task.sheetRole === 'sub' &&
      relatedTasks &&
      relatedTasks.length > 0;

    if (!isGroupedSub) {
      const taskLabel =
        task.sheetRole === 'sub'
          ? task.taskName || task.mainTaskName || 'Sub task'
          : task.taskName || 'Task';
      const subject =
        'Task Alert: ' + taskLabel + ' — ' + matchedBy + ' [' + config.dataSheetName + ']';
      const htmlBody = TaskNotificationModule.buildImmediateHtml(config, task, matchedBy);
      return sendEmail_(config, subject, htmlBody);
    }

    const parentLabel = TaskNotificationModule.getParentLabelFromTask(task);
    const maxRows = AppConfig.getDigestMaxRowsPerEmail();
    const chunks = [];
    for (var i = 0; i < relatedTasks.length; i += maxRows) {
      chunks.push(relatedTasks.slice(i, i + maxRows));
    }

    const sendResults = [];
    chunks.forEach(function (chunkTasks, index) {
      var subject =
        'Task Alert: ' +
        parentLabel +
        ' (' +
        chunkTasks.length +
        ' sub-task(s)) — ' +
        matchedBy +
        ' [' +
        config.dataSheetName +
        ']';
      if (chunks.length > 1) {
        subject += ' — part ' + (index + 1) + ' of ' + chunks.length;
      }
      const htmlBody = TaskNotificationModule.buildImmediateGroupedSubHtml(
        config,
        parentLabel,
        chunkTasks,
        matchedBy
      );
      sendResults.push(sendEmail_(config, subject, htmlBody));
    });

    const failed = sendResults.filter(function (result) {
      return !result || !result.success;
    });

    if (failed.length > 0) {
      return {
        success: false,
        emailCount: sendResults.length,
        sentCount: sendResults.length - failed.length,
        failedCount: failed.length,
        error: failed[0].error || 'Send failed',
      };
    }

    return {
      success: true,
      emailCount: sendResults.length,
      groupedByParent: true,
      taskCount: relatedTasks.length,
    };
  }

  return {
    sendScheduledDigest: sendScheduledDigest_,
    sendImmediateAlert: sendImmediateAlert_,
  };
})();
