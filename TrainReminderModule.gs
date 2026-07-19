/**
 * Train Route reminders:
 * - BOOK: ARP booking-open window (travel date − 60 days), skipped if already booked
 * - TRAVEL: Reminder Days Before travel date (still sent after booking)
 * - PENDING: unbooked travel dates in the next lookahead days (default 90), grouped by route
 *   Status: Booking opens TODAY / Opens in N day(s) / Open — not booked
 * One digest per recipient email (routes fan out to each address listed on the row).
 */
const TrainReminderModule = (function () {
  'use strict';

  const CARD_FIELDS = [
    { key: 'daysToGo', label: 'Countdown' },
    { key: 'type', label: 'Type' },
    { key: 'routeName', label: 'Route' },
    { key: 'fromStation', label: 'From' },
    { key: 'toStation', label: 'To' },
    { key: 'travelDate', label: 'Travel date' },
    { key: 'bookingOpens', label: 'Booking opens' },
    { key: 'recurrence', label: 'Recurrence' },
    { key: 'dayOfWeek', label: 'Day of week' },
    { key: 'trainNo', label: 'Train no' },
    { key: 'trainName', label: 'Train name' },
    { key: 'travelClass', label: 'Class' },
    { key: 'daysBefore', label: 'Days before' },
    { key: 'booked', label: 'Already booked' },
    { key: 'notes', label: 'Notes' },
  ];

  function getTimezone_() {
    return AppConfig.getTrainDateTimezone() || 'Asia/Kolkata';
  }

  function today_() {
    return TrainRouteModule.startOfDay(new Date());
  }

  function daysUntil_(targetDate, fromDate) {
    if (!targetDate) {
      return null;
    }
    const from = fromDate || today_();
    const ms = targetDate.getTime() - from.getTime();
    return Math.round(ms / (24 * 60 * 60 * 1000));
  }

  function bookingOpenDate_(travelDate) {
    if (!travelDate) {
      return null;
    }
    const open = new Date(travelDate.getTime());
    open.setDate(open.getDate() - AppConfig.getTrainArpDays());
    return TrainRouteModule.startOfDay(open);
  }

  function dateKey_(date) {
    return Utilities.formatDate(date, getTimezone_(), 'yyyy-MM-dd');
  }

  function matchingReminderDay_(daysUntil, reminderDaysBefore) {
    if (daysUntil === null || daysUntil === undefined) {
      return null;
    }
    for (var i = 0; i < reminderDaysBefore.length; i++) {
      if (reminderDaysBefore[i] === daysUntil) {
        return reminderDaysBefore[i];
      }
    }
    return null;
  }

  function escapeHtml_(value) {
    return String(value === null || value === undefined ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function reminderSentKey_(type, rowNumber, targetDate, daysBefore, slot) {
    return (
      'train_rem_' +
      type +
      '_' +
      rowNumber +
      '_' +
      dateKey_(targetDate) +
      '_' +
      daysBefore +
      '_' +
      String(slot || 'na')
    );
  }

  function wasSent_(type, rowNumber, targetDate, daysBefore, slot) {
    const key = reminderSentKey_(type, rowNumber, targetDate, daysBefore, slot);
    return PropertiesService.getScriptProperties().getProperty(key) === '1';
  }

  function markSent_(type, rowNumber, targetDate, daysBefore, slot) {
    const key = reminderSentKey_(type, rowNumber, targetDate, daysBefore, slot);
    PropertiesService.getScriptProperties().setProperty(key, '1');
  }

  /** One digest per email address (not per Reminder Email list). */
  function recipientKey_(email) {
    return SpreadsheetUtils.normalizeText(email || '');
  }

  function typeLabel_(kind, daysBefore) {
    if (kind === 'PENDING') {
      return 'PENDING — not booked yet';
    }
    if (kind === 'BOOK') {
      if (daysBefore === 0) {
        return 'BOOK — OPEN today';
      }
      if (daysBefore === 1) {
        return 'BOOK — opens tomorrow';
      }
      return 'BOOK — opens in ' + daysBefore + 'd';
    }
    if (daysBefore === 0) {
      return 'TRAVEL — today';
    }
    if (daysBefore === 1) {
      return 'TRAVEL — tomorrow';
    }
    return 'TRAVEL — in ' + daysBefore + 'd';
  }

  /**
   * Human countdown e.g. "2 days to go", "Today", "1 day to go".
   */
  function daysToGoLabel_(kind, daysBefore) {
    if (kind === 'PENDING') {
      if (daysBefore === null || daysBefore === undefined) {
        return 'Booking still missing';
      }
      if (daysBefore < 0) {
        return 'Booking window already open — still not booked';
      }
      if (daysBefore === 0) {
        return 'Today — book now (still pending)';
      }
      if (daysBefore === 1) {
        return '1 day to go until booking opens';
      }
      return daysBefore + ' days to go until booking opens';
    }
    if (daysBefore === 0) {
      return kind === 'BOOK' ? 'Today — book now' : 'Today — travel day';
    }
    if (daysBefore === 1) {
      return kind === 'BOOK' ? '1 day to go — book tomorrow' : '1 day to go';
    }
    return daysBefore + ' days to go';
  }

  function entryToCardData_(entry, index) {
    const route = entry.route;
    const openHour = AppConfig.getTrainBookingOpenHour();
    const openLabel = entry.bookingOpen
      ? TrainRouteModule.formatDate(entry.bookingOpen) + ' ~' + openHour + ':00 AM IST'
      : '—';

    var accent = '#2b6cb0';
    var background = '#f0f7ff';
    if (entry.kind === 'BOOK') {
      accent = '#b7791f';
      background = '#fffaf0';
    } else if (entry.kind === 'PENDING') {
      accent = '#c53030';
      background = '#fff5f5';
    }

    const countdownDays =
      entry.kind === 'PENDING'
        ? entry.daysUntilOpen
        : entry.daysBefore;

    return {
      index: index + 1,
      kind: entry.kind,
      accent: accent,
      background: background,
      type: typeLabel_(entry.kind, entry.kind === 'PENDING' ? entry.daysUntilOpen : entry.daysBefore),
      daysToGo: daysToGoLabel_(entry.kind, countdownDays),
      routeName: route.routeName || '—',
      fromStation: route.fromStation || '—',
      toStation: route.toStation || '—',
      travelDate: TrainRouteModule.formatDate(entry.travelDate) || '—',
      bookingOpens: openLabel,
      recurrence: route.recurrence || '—',
      dayOfWeek: route.dayOfWeek || '—',
      trainNo: route.preferredTrainNo || '—',
      trainName: route.preferredTrainName || '—',
      travelClass: route.preferredClass || '—',
      daysBefore:
        entry.kind === 'PENDING'
          ? String(entry.daysUntilOpen != null ? entry.daysUntilOpen : '—')
          : String(entry.daysBefore),
      booked: entry.alreadyBooked ? 'Yes' : 'No',
      notes: route.notes || '—',
    };
  }

  /**
   * Mobile-first stacked cards (email tables wider than ~360px are hard on phones).
   */
  function buildCardHtml_(card) {
    var fieldsHtml = '';
    CARD_FIELDS.forEach(function (field) {
      const value = card[field.key];
      fieldsHtml +=
        '<tr>' +
        '<td style="padding:8px 10px;border-top:1px solid #e2e8f0;width:38%;vertical-align:top;' +
        'font-size:12px;line-height:1.35;color:#4a5568;font-weight:bold">' +
        escapeHtml_(field.label) +
        '</td>' +
        '<td style="padding:8px 10px;border-top:1px solid #e2e8f0;width:62%;vertical-align:top;' +
        'font-size:14px;line-height:1.4;color:#1a202c;word-break:break-word">' +
        escapeHtml_(value) +
        '</td>' +
        '</tr>';
    });

    return (
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
      'style="width:100%;max-width:600px;border-collapse:separate;border-spacing:0;' +
      'margin:0 0 14px 0;background:' +
      card.background +
      ';border:1px solid #cbd5e0;border-left:5px solid ' +
      card.accent +
      ';border-radius:8px">' +
      '<tr><td style="padding:12px 12px 4px 12px">' +
      '<div style="font-size:11px;letter-spacing:0.04em;text-transform:uppercase;color:#718096;margin:0 0 4px 0">' +
      'Item ' +
      card.index +
      '</div>' +
      '<div style="font-size:16px;font-weight:bold;line-height:1.3;color:' +
      card.accent +
      ';margin:0">' +
      escapeHtml_(card.type) +
      '</div>' +
      '<div style="display:inline-block;margin:8px 0 0 0;padding:6px 10px;border-radius:999px;' +
      'background:' +
      card.accent +
      ';color:#ffffff;font-size:14px;font-weight:bold;line-height:1.2">' +
      escapeHtml_(card.daysToGo) +
      '</div>' +
      '<div style="font-size:15px;font-weight:bold;line-height:1.35;color:#1a202c;margin:8px 0 0 0;' +
      'word-break:break-word">' +
      escapeHtml_(card.routeName) +
      '</div>' +
      '<div style="font-size:13px;line-height:1.4;color:#2d3748;margin:6px 0 8px 0;word-break:break-word">' +
      escapeHtml_(card.fromStation) +
      ' → ' +
      escapeHtml_(card.toStation) +
      '</div>' +
      '</td></tr>' +
      '<tr><td style="padding:0">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse">' +
      fieldsHtml +
      '</table>' +
      '</td></tr></table>'
    );
  }

  function buildDigestCardsHtml_(entries) {
    var html = '';
    entries.forEach(function (entry, index) {
      html += buildCardHtml_(entryToCardData_(entry, index));
    });
    return html;
  }

  /** Status for missing-date rows based on days until IRCTC booking opens. */
  function pendingStatusLabel_(daysUntilOpen) {
    if (daysUntilOpen === null || daysUntilOpen === undefined) {
      return 'Not booked';
    }
    if (daysUntilOpen < 0) {
      return 'Open — not booked';
    }
    if (daysUntilOpen === 0) {
      return 'Booking opens TODAY';
    }
    if (daysUntilOpen === 1) {
      return 'Opens in 1 day';
    }
    return 'Opens in ' + daysUntilOpen + ' day(s)';
  }

  /**
   * Unbooked travel dates in lookahead window, grouped by route (From → To / Route Name).
   */
  function buildPendingDatesTableHtml_(pendingEntries) {
    const openHour = AppConfig.getTrainBookingOpenHour();
    const groups = {};
    const groupOrder = [];

    pendingEntries.forEach(function (entry) {
      const route = entry.route;
      const key =
        SpreadsheetUtils.normalizeText(route.fromStation) +
        '|' +
        SpreadsheetUtils.normalizeText(route.toStation) +
        '|' +
        SpreadsheetUtils.normalizeText(route.routeName);
      if (!groups[key]) {
        groups[key] = {
          routeName: route.routeName || 'Route',
          fromStation: route.fromStation || '',
          toStation: route.toStation || '',
          entries: [],
        };
        groupOrder.push(key);
      }
      groups[key].entries.push(entry);
    });

    groupOrder.sort(function (a, b) {
      const ga = groups[a];
      const gb = groups[b];
      const na = SpreadsheetUtils.normalizeText(ga.routeName + ' ' + ga.fromStation + ' ' + ga.toStation);
      const nb = SpreadsheetUtils.normalizeText(gb.routeName + ' ' + gb.fromStation + ' ' + gb.toStation);
      if (na < nb) {
        return -1;
      }
      if (na > nb) {
        return 1;
      }
      return 0;
    });

    const th =
      'border:1px solid #999;padding:8px;background:#9b2c2c;color:#fff;font-size:11px;text-align:left';

    var html = '';
    var globalIndex = 0;

    groupOrder.forEach(function (key) {
      const group = groups[key];
      group.entries.sort(function (a, b) {
        return a.travelDate.getTime() - b.travelDate.getTime();
      });

      html +=
        '<div style="margin:0 0 16px 0;border:1px solid #feb2b2;border-radius:8px;overflow:hidden">' +
        '<div style="background:#c53030;color:#fff;padding:10px 12px;font-family:Arial,Helvetica,sans-serif">' +
        '<div style="font-size:15px;font-weight:bold;line-height:1.3;word-break:break-word">' +
        escapeHtml_(group.routeName) +
        '</div>' +
        '<div style="font-size:13px;line-height:1.35;margin-top:4px;word-break:break-word;opacity:0.95">' +
        escapeHtml_(group.fromStation) +
        ' → ' +
        escapeHtml_(group.toStation) +
        '</div>' +
        '<div style="font-size:12px;margin-top:4px;opacity:0.9">' +
        group.entries.length +
        ' missing date(s)</div></div>' +
        '<div style="width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch;background:#fff">' +
        '<table cellpadding="0" cellspacing="0" border="0" ' +
        'style="width:100%;min-width:480px;border-collapse:collapse;font-family:Arial,Helvetica,sans-serif">' +
        '<thead><tr>' +
        '<th style="' +
        th +
        '">#</th>' +
        '<th style="' +
        th +
        '">Travel date</th>' +
        '<th style="' +
        th +
        '">From</th>' +
        '<th style="' +
        th +
        '">To</th>' +
        '<th style="' +
        th +
        '">Booking opens</th>' +
        '<th style="' +
        th +
        '">Travel in</th>' +
        '<th style="' +
        th +
        '">Status</th>' +
        '</tr></thead><tbody>';

      group.entries.forEach(function (entry, index) {
        globalIndex++;
        const route = entry.route;
        const travelLabel = TrainRouteModule.formatDate(entry.travelDate);
        const openLabel = entry.bookingOpen
          ? TrainRouteModule.formatDate(entry.bookingOpen) + ' ' + openHour + ':00 AM'
          : '—';
        const travelIn =
          entry.daysUntilTravel === 0
            ? 'Today'
            : entry.daysUntilTravel === 1
              ? '1 day'
              : entry.daysUntilTravel + ' days';
        const statusLabel = pendingStatusLabel_(entry.daysUntilOpen);
        const isOpen = entry.daysUntilOpen != null && entry.daysUntilOpen < 0;
        const statusColor = isOpen ? '#c53030' : '#b7791f';
        const bg = index % 2 === 0 ? '#fff5f5' : '#ffffff';

        html +=
          '<tr style="background:' +
          bg +
          '">' +
          '<td style="border:1px solid #e2e8f0;padding:8px;font-size:12px;text-align:center">' +
          globalIndex +
          '</td>' +
          '<td style="border:1px solid #e2e8f0;padding:8px;font-size:13px;font-weight:bold;white-space:nowrap">' +
          escapeHtml_(travelLabel) +
          '</td>' +
          '<td style="border:1px solid #e2e8f0;padding:8px;font-size:12px;word-break:break-word">' +
          escapeHtml_(route.fromStation || '—') +
          '</td>' +
          '<td style="border:1px solid #e2e8f0;padding:8px;font-size:12px;word-break:break-word">' +
          escapeHtml_(route.toStation || '—') +
          '</td>' +
          '<td style="border:1px solid #e2e8f0;padding:8px;font-size:12px;white-space:nowrap">' +
          escapeHtml_(openLabel) +
          '</td>' +
          '<td style="border:1px solid #e2e8f0;padding:8px;font-size:12px;text-align:center;font-weight:bold;color:#c53030">' +
          escapeHtml_(travelIn) +
          '</td>' +
          '<td style="border:1px solid #e2e8f0;padding:8px;font-size:12px;text-align:center;color:' +
          statusColor +
          '"><b>' +
          escapeHtml_(statusLabel) +
          '</b></td>' +
          '</tr>';
      });

      html += '</tbody></table></div></div>';
    });

    return html;
  }

  function buildDigestHtml_(dueEntries, pendingEntries, slot) {
    const due = dueEntries || [];
    const pending = pendingEntries || [];
    const bookCount = due.filter(function (e) {
      return e.kind === 'BOOK';
    }).length;
    const travelCount = due.filter(function (e) {
      return e.kind === 'TRAVEL';
    }).length;
    const todayLabel = Utilities.formatDate(new Date(), getTimezone_(), 'dd-MMM-yyyy HH:mm');

    var bodySections = '';

    if (due.length > 0) {
      bodySections +=
        '<div style="font-size:15px;font-weight:bold;color:#1a202c;margin:0 0 8px 0">Due today (countdown)</div>' +
        buildDigestCardsHtml_(due);
    }

    if (pending.length > 0) {
      const lookaheadDays =
        typeof AppConfig.getTrainReminderLookaheadDays === 'function'
          ? AppConfig.getTrainReminderLookaheadDays()
          : 90;
      bodySections +=
        '<div style="font-size:15px;font-weight:bold;color:#c53030;margin:16px 0 8px 0">' +
        'Missing dates — next ' +
        lookaheadDays +
        ' days, still not booked (' +
        pending.length +
        ')</div>' +
        '<p style="margin:0 0 10px 0;font-size:12px;line-height:1.45;color:#718096">' +
        'Grouped by route (From → To). Status shows when booking opens (ARP ' +
        AppConfig.getTrainArpDays() +
        ' days before travel). Booking one weekly trip does not clear other dates.</p>' +
        buildPendingDatesTableHtml_(pending);
    }

    return (
      '<div style="margin:0;padding:0;background:#f7fafc">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
      'style="width:100%;border-collapse:collapse;background:#f7fafc">' +
      '<tr><td align="center" style="padding:12px 10px">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
      'style="width:100%;max-width:600px;border-collapse:collapse;background:#ffffff;' +
      'border:1px solid #e2e8f0;border-radius:10px">' +
      '<tr><td style="padding:16px 14px 8px 14px;font-family:Arial,Helvetica,sans-serif;color:#1a202c">' +
      '<h1 style="margin:0 0 8px 0;font-size:20px;line-height:1.3">Train Route reminders</h1>' +
      '<p style="margin:0 0 6px 0;font-size:13px;line-height:1.45;color:#4a5568">' +
      escapeHtml_(todayLabel) +
      ' IST · Slot ' +
      escapeHtml_(slot || '') +
      '</p>' +
      '<p style="margin:0 0 14px 0;font-size:14px;line-height:1.45;color:#2d3748">' +
      'Due <b>' +
      due.length +
      '</b> (book ' +
      bookCount +
      ' / travel ' +
      travelCount +
      ') · Unbooked (next ' +
      (typeof AppConfig.getTrainReminderLookaheadDays === 'function'
        ? AppConfig.getTrainReminderLookaheadDays()
        : 90) +
      'd) <b>' +
      pending.length +
      '</b></p>' +
      '<p style="margin:0 0 12px 0;font-size:12px;line-height:1.45;color:#718096">' +
      'Gold/blue cards = today’s countdown. Red table = all open windows still not booked.</p>' +
      '</td></tr>' +
      '<tr><td style="padding:0 10px 16px 10px;font-family:Arial,Helvetica,sans-serif">' +
      bodySections +
      '</td></tr>' +
      '</table></td></tr></table></div>'
    );
  }

  function buildDigestSubject_(dueEntries, pendingEntries) {
    const due = dueEntries || [];
    const pending = pendingEntries || [];
    const bookCount = due.filter(function (e) {
      return e.kind === 'BOOK';
    }).length;
    const travelCount = due.filter(function (e) {
      return e.kind === 'TRAVEL';
    }).length;
    return (
      'Train reminders — due:' +
      due.length +
      ' (book:' +
      bookCount +
      ', travel:' +
      travelCount +
      ') · missing:' +
      pending.length
    );
  }

  function pendingDigestSentKey_(recipientKey, slot) {
    return 'train_rem_PENDING_DIGEST_' + recipientKey + '_' + dateKey_(today_()) + '_' + String(slot || 'na');
  }

  function wasPendingDigestSent_(recipientKey, slot) {
    return PropertiesService.getScriptProperties().getProperty(pendingDigestSentKey_(recipientKey, slot)) === '1';
  }

  function markPendingDigestSent_(recipientKey, slot) {
    PropertiesService.getScriptProperties().setProperty(pendingDigestSentKey_(recipientKey, slot), '1');
  }

  function sendDigest_(emails, dueEntries, pendingEntries, slot, recipientKey) {
    const due = dueEntries || [];
    const pending = pendingEntries || [];
    if (!emails || emails.length === 0 || (due.length === 0 && pending.length === 0)) {
      return { success: false, error: 'Nothing to send' };
    }

    const subject = buildDigestSubject_(due, pending);
    const html = buildDigestHtml_(due, pending, slot);

    try {
      GmailApp.sendEmail(emails.join(','), subject, '', {
        htmlBody: html,
        name: 'Train Tracker',
      });

      due.forEach(function (entry) {
        const target = entry.kind === 'BOOK' ? entry.bookingOpen : entry.travelDate;
        markSent_(entry.kind, entry.route.rowNumber, target, entry.daysBefore, slot);
      });

      if (pending.length > 0 && recipientKey) {
        markPendingDigestSent_(recipientKey, slot);
      }

      console.log('TrainReminderModule.sendDigest_: sent', {
        to: emails,
        due: due.length,
        pending: pending.length,
        subject: subject,
      });
      return {
        success: true,
        subject: subject,
        dueCount: due.length,
        pendingCount: pending.length,
      };
    } catch (err) {
      console.error('TrainReminderModule.sendDigest_ failed', {
        error: err,
        stack: err && err.stack,
      });
      ErrorLogModule.error('TrainReminderModule', 'sendDigest_', String(err), {
        sheetName: AppConfig.getTrainRouteSheetName(),
        details: subject,
      });
      return { success: false, error: String(err) };
    }
  }

  function ensureGroup_(map, key, emails, slot) {
    if (!map[key]) {
      map[key] = {
        emails: (emails || []).slice(),
        slot: slot,
        entries: [],
        pending: [],
        inAlertWindow: false,
      };
    }
    if (slot) {
      map[key].slot = slot;
    }
    return map[key];
  }

  /**
   * Build digests keyed by individual email.
   * Each person gets one mail with only routes where their address appears in Reminder Email.
   * Example: Row1 → A,B | Row2 → A,B,C | Row3 → A  ⇒  A gets 1+2+3, B gets 1+2, C gets 2 only.
   */
  function collectDueAndPending_(routes, today, summary) {
    const byRecipient = {};
    const lookaheadDays =
      typeof AppConfig.getTrainReminderLookaheadDays === 'function'
        ? AppConfig.getTrainReminderLookaheadDays()
        : AppConfig.getTrainArpDays();

    routes.forEach(function (route) {
      try {
        const alertCheck = TaskNotificationModule.isAlertTimeNow(route.alertTimes, getTimezone_());
        if (!alertCheck.due) {
          summary.skippedNoAlertTime++;
          return;
        }
        const slot = alertCheck.slot || 'slot';
        const routeDue = [];
        const routePending = [];

        const travelDate = TrainRouteModule.resolveTravelDate(route, today);
        if (!travelDate) {
          summary.skippedNotDue++;
        } else {
          const bookingOpen = bookingOpenDate_(travelDate);
          const alreadyBooked = TrainBookingHistoryModule.hasMatchingBooking(
            route.fromStation,
            route.toStation,
            travelDate
          );

          const daysUntilOpen = daysUntil_(bookingOpen, today);
          const bookDay = matchingReminderDay_(daysUntilOpen, route.reminderDaysBefore);
          const daysUntilTravel = daysUntil_(travelDate, today);
          const travelDay = matchingReminderDay_(daysUntilTravel, route.reminderDaysBefore);

          if (bookDay !== null) {
            if (alreadyBooked) {
              summary.skippedBooked++;
            } else if (wasSent_('BOOK', route.rowNumber, bookingOpen, bookDay, slot)) {
              summary.skippedAlreadySent++;
            } else {
              routeDue.push({
                kind: 'BOOK',
                route: route,
                travelDate: travelDate,
                bookingOpen: bookingOpen,
                daysBefore: bookDay,
                alreadyBooked: false,
              });
            }
          }

          if (travelDay !== null) {
            if (wasSent_('TRAVEL', route.rowNumber, travelDate, travelDay, slot)) {
              summary.skippedAlreadySent++;
            } else {
              routeDue.push({
                kind: 'TRAVEL',
                route: route,
                travelDate: travelDate,
                bookingOpen: bookingOpen,
                daysBefore: travelDay,
                alreadyBooked: alreadyBooked,
              });
            }
          } else if (bookDay === null && alreadyBooked) {
            summary.skippedNotDue++;
          }
        }

        // PENDING: every unbooked upcoming date in lookahead (incl. before window opens)
        const upcomingDates = TrainRouteModule.listUpcomingTravelDates(route, today, lookaheadDays);
        upcomingDates.forEach(function (tripDate) {
          const openDate = bookingOpenDate_(tripDate);
          if (!openDate) {
            return;
          }
          if (tripDate.getTime() < today.getTime()) {
            return;
          }
          const booked = TrainBookingHistoryModule.hasMatchingBooking(
            route.fromStation,
            route.toStation,
            tripDate
          );
          if (booked) {
            return;
          }

          const daysUntilOpenTrip = daysUntil_(openDate, today);
          const daysUntilTravelTrip = daysUntil_(tripDate, today);
          routePending.push({
            kind: 'PENDING',
            route: route,
            travelDate: tripDate,
            bookingOpen: openDate,
            daysBefore: daysUntilOpenTrip,
            daysUntilOpen: daysUntilOpenTrip,
            daysUntilTravel: daysUntilTravelTrip,
            alreadyBooked: false,
          });
        });

        const emails = route.reminderEmails || [];
        if (emails.length === 0) {
          console.error('TrainReminderModule.collectDueAndPending_: route has no Reminder Email', {
            route: route.routeName,
            row: route.rowNumber,
          });
          return;
        }

        emails.forEach(function (email) {
          const key = recipientKey_(email);
          if (!key) {
            return;
          }
          const group = ensureGroup_(byRecipient, key, [String(email).trim()], slot);
          group.inAlertWindow = true;
          routeDue.forEach(function (entry) {
            group.entries.push(entry);
          });
          routePending.forEach(function (entry) {
            group.pending.push(entry);
          });
        });
      } catch (err) {
        summary.errors++;
        console.error('TrainReminderModule.collectDueAndPending_: route failed', {
          route: route && route.routeName,
          error: err,
          stack: err && err.stack,
        });
      }
    });

    return byRecipient;
  }

  /**
   * Process all active routes; send digest with due items + all missing (unbooked) trips.
   */
  function processDueReminders_() {
    const routes = TrainRouteModule.readActiveRoutes();
    const today = today_();
    const summary = {
      routes: routes.length,
      digestsSent: 0,
      bookSent: 0,
      travelSent: 0,
      pendingListed: 0,
      skippedBooked: 0,
      skippedNotDue: 0,
      skippedAlreadySent: 0,
      skippedNoAlertTime: 0,
      errors: 0,
    };

    const byRecipient = collectDueAndPending_(routes, today, summary);

    Object.keys(byRecipient).forEach(function (key) {
      const group = byRecipient[key];
      if (!group.inAlertWindow) {
        return;
      }

      const due = group.entries || [];
      var pending = group.pending || [];

      // Sort pending by travel date soonest first
      pending = pending.slice().sort(function (a, b) {
        return a.travelDate.getTime() - b.travelDate.getTime();
      });

      if (due.length === 0 && pending.length === 0) {
        return;
      }

      // Avoid re-sending pending-only digests every 30 minutes
      if (due.length === 0 && wasPendingDigestSent_(key, group.slot)) {
        summary.skippedAlreadySent++;
        return;
      }

      const result = sendDigest_(group.emails, due, pending, group.slot, key);
      if (result.success) {
        summary.digestsSent++;
        summary.pendingListed += pending.length;
        due.forEach(function (entry) {
          if (entry.kind === 'BOOK') {
            summary.bookSent++;
          } else if (entry.kind === 'TRAVEL') {
            summary.travelSent++;
          }
        });
      } else {
        summary.errors++;
      }
    });

    console.log('TrainReminderModule.processDueReminders_', summary);
    return summary;
  }

  return {
    processDueReminders: processDueReminders_,
    bookingOpenDate: bookingOpenDate_,
    daysUntil: daysUntil_,
  };
})();
