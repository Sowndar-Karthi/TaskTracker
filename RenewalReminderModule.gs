/**
 * Builds and tracks renewal reminder emails.
 */
const RenewalReminderModule = (function () {
  'use strict';

  function formatDate_(date) {
    return RenewalTrackerModule.formatDate(date);
  }

  function formatAmount_(amount) {
    if (amount === null || amount === undefined || amount === '') {
      return '';
    }
    return String(amount);
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

  function getDueDateKey_(dueDate) {
    return Utilities.formatDate(dueDate, 'Asia/Kolkata', 'yyyy-MM-dd');
  }

  function getReminderSentKey_(rowNumber, dueDate, daysBefore) {
    return (
      'renewal_reminder_' +
      rowNumber +
      '_' +
      getDueDateKey_(dueDate) +
      '_' +
      daysBefore
    );
  }

  function wasReminderSent_(rowNumber, dueDate, daysBefore) {
    const key = getReminderSentKey_(rowNumber, dueDate, daysBefore);
    return PropertiesService.getScriptProperties().getProperty(key) === '1';
  }

  function markReminderSent_(rowNumber, dueDate, daysBefore) {
    const key = getReminderSentKey_(rowNumber, dueDate, daysBefore);
    PropertiesService.getScriptProperties().setProperty(key, '1');
  }

  function getMatchingReminderDay_(item, today) {
    if (!item || !item.nextDueDate || item.reminderDaysBefore.length === 0) {
      return null;
    }

    const daysUntilDue = RenewalDueDateModule.daysUntilDue(item.nextDueDate, today);
    if (daysUntilDue === null) {
      return null;
    }

    for (var i = 0; i < item.reminderDaysBefore.length; i++) {
      const daysBefore = item.reminderDaysBefore[i];
      if (daysUntilDue === daysBefore) {
        return daysBefore;
      }
    }

    return null;
  }

  function getAmountLabel_(item) {
    const categoryKey = SpreadsheetUtils.normalizeText(item.category);
    if (categoryKey === 'credit card') {
      if (item.minimumDueAmount !== null && item.totalDueAmount !== null) {
        return (
          'Minimum Due: ' +
          formatAmount_(item.minimumDueAmount) +
          ' | Total Due: ' +
          formatAmount_(item.totalDueAmount)
        );
      }
      if (item.totalDueAmount !== null) {
        return 'Total Due: ' + formatAmount_(item.totalDueAmount);
      }
      if (item.minimumDueAmount !== null) {
        return 'Minimum Due: ' + formatAmount_(item.minimumDueAmount);
      }
    }

    if (item.totalDueAmount !== null) {
      return 'Amount: ' + formatAmount_(item.totalDueAmount);
    }

    return '';
  }

  function getReminderSubject_(item, daysBefore) {
    var timing = '';
    if (daysBefore === 0) {
      timing = 'Due Today';
    } else if (daysBefore === 1) {
      timing = 'Due Tomorrow';
    } else {
      timing = 'Due in ' + daysBefore + ' days';
    }

    return 'Renewal Reminder: ' + item.itemName + ' — ' + timing;
  }

  function buildReminderHtml_(item, daysBefore) {
    const amountLabel = getAmountLabel_(item);
    const daysUntilDue = RenewalDueDateModule.daysUntilDue(item.nextDueDate);

    return (
      '<p><strong>Personal Renewal Tracker</strong></p>' +
      '<p><strong>' +
      escapeHtml_(item.itemName) +
      '</strong> (' +
      escapeHtml_(item.category) +
      ')</p>' +
      '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
      '<tr><th align="left">Next Due Date</th><td>' +
      escapeHtml_(formatDate_(item.nextDueDate)) +
      '</td></tr>' +
      '<tr><th align="left">Reminder</th><td>' +
      (daysBefore === 0 ? 'Due today' : 'Due in ' + daysBefore + ' day(s)') +
      '</td></tr>' +
      (amountLabel
        ? '<tr><th align="left">Amount</th><td>' + escapeHtml_(amountLabel) + '</td></tr>'
        : '') +
      '<tr><th align="left">Days Until Due</th><td>' +
      escapeHtml_(daysUntilDue) +
      '</td></tr>' +
      '<tr><th align="left">Recurrence</th><td>Every ' +
      escapeHtml_(item.recurrenceInterval) +
      ' ' +
      escapeHtml_(item.recurrenceUnit) +
      '</td></tr>' +
      '</table>' +
      '<p style="color:#666;font-size:12px;">Update <em>Last Paid Date</em> after payment to roll the next due date forward.</p>'
    );
  }

  function sendReminder_(item, daysBefore) {
    if (!item || item.reminderEmails.length === 0) {
      return { success: false, error: 'No reminder email configured' };
    }

    if (wasReminderSent_(item.rowNumber, item.nextDueDate, daysBefore)) {
      return { success: true, skipped: true, reason: 'Already sent' };
    }

    const subject = getReminderSubject_(item, daysBefore);
    const htmlBody = buildReminderHtml_(item, daysBefore);

    try {
      GmailApp.sendEmail(item.reminderEmails.join(','), subject, '', {
        htmlBody: htmlBody,
        name: 'Renewal Tracker',
      });

      markReminderSent_(item.rowNumber, item.nextDueDate, daysBefore);

      if (daysBefore > 0 && daysBefore <= 7 && item.statusKey === 'active') {
        RenewalTrackerModule.updateStatus(item.rowNumber, 'Due Soon');
      }

      return { success: true, subject: subject, daysBefore: daysBefore };
    } catch (err) {
      ErrorLogModule.error('RenewalReminderModule', 'sendReminder_', 'Send failed', {
        itemName: item.itemName,
        rowNumber: item.rowNumber,
        error: err,
      });
      return { success: false, error: String(err) };
    }
  }

  function processDueReminders_(today) {
    const items = RenewalTrackerModule.readAllItems()
      .filter(function (item) {
        return (
          RenewalTrackerModule.isRemindableStatus(item.statusKey) &&
          item.itemName &&
          item.reminderEmails.length > 0 &&
          item.reminderDaysBefore.length > 0
        );
      })
      .map(function (item) {
        if (!item.nextDueDate && item.anchorTypeKey !== 'manual only') {
          RenewalDueDateModule.recalculateItem(item);
          return RenewalTrackerModule.readItemByRow(item.rowNumber) || item;
        }
        return item;
      })
      .filter(function (item) {
        return item && item.nextDueDate;
      });

    const results = [];

    items.forEach(function (item) {
      try {
        const daysBefore = getMatchingReminderDay_(item, today);
        if (daysBefore === null) {
          results.push({
            rowNumber: item.rowNumber,
            itemName: item.itemName,
            skipped: true,
            reason: 'No reminder due today',
          });
          return;
        }

        results.push({
          rowNumber: item.rowNumber,
          itemName: item.itemName,
          daysBefore: daysBefore,
          send: sendReminder_(item, daysBefore),
        });
      } catch (err) {
        ErrorLogModule.error('RenewalReminderModule', 'processDueReminders_', 'Item failed', {
          itemName: item.itemName,
          error: err,
        });
      }
    });

    return {
      checkedCount: items.length,
      results: results,
    };
  }

  return {
    getMatchingReminderDay: getMatchingReminderDay_,
    buildReminderHtml: buildReminderHtml_,
    sendReminder: sendReminder_,
    processDueReminders: processDueReminders_,
    wasReminderSent: wasReminderSent_,
  };
})();
