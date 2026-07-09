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
      console.error('NotificationService.sendEmail_: failed', {
        subject: subject,
        error: err,
      });
      return { success: false, error: String(err) };
    }
  }

  function sendScheduledDigest_(config, tasks) {
    if (!tasks || tasks.length === 0) {
      return { success: true, skipped: true, reason: 'No matching tasks' };
    }

    const subject =
      'Team Project Tracker Alert — ' + config.dataSheetName + ' (' + tasks.length + ' task(s))';
    const htmlBody = TaskNotificationModule.buildDigestHtml(config, tasks);
    return sendEmail_(config, subject, htmlBody);
  }

  function sendImmediateAlert_(config, task, matchedBy) {
    const subject =
      'Task Alert: ' + task.taskName + ' — ' + matchedBy + ' [' + config.dataSheetName + ']';
    const htmlBody = TaskNotificationModule.buildImmediateHtml(config, task, matchedBy);
    return sendEmail_(config, subject, htmlBody);
  }

  return {
    sendScheduledDigest: sendScheduledDigest_,
    sendImmediateAlert: sendImmediateAlert_,
  };
})();
