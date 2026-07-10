/**
 * Calculates next due dates for Renewal Tracker items.
 */
const RenewalDueDateModule = (function () {
  'use strict';

  const ANCHOR_LAST_PAYMENT = 'last payment date';
  const ANCHOR_FIXED_DATE = 'fixed anchor date';
  const ANCHOR_FIXED_DUE_DAY = 'fixed due day';
  const ANCHOR_MANUAL = 'manual only';

  function startOfDay_(date) {
    const copy = new Date(date);
    copy.setHours(0, 0, 0, 0);
    return copy;
  }

  function normalizeUnit_(unit) {
    return SpreadsheetUtils.normalizeText(unit);
  }

  function addInterval_(date, interval, unitKey) {
    const result = new Date(date);
    const amount = interval > 0 ? interval : 1;

    if (unitKey === 'days') {
      result.setDate(result.getDate() + amount);
      return result;
    }
    if (unitKey === 'months') {
      result.setMonth(result.getMonth() + amount);
      return result;
    }
    if (unitKey === 'years') {
      result.setFullYear(result.getFullYear() + amount);
      return result;
    }

    result.setMonth(result.getMonth() + amount);
    return result;
  }

  function clampDayInMonth_(year, monthIndex, day) {
    const lastDay = new Date(year, monthIndex + 1, 0).getDate();
    return Math.min(day, lastDay);
  }

  function buildDateWithDay_(year, monthIndex, day) {
    const safeDay = clampDayInMonth_(year, monthIndex, day);
    return new Date(year, monthIndex, safeDay);
  }

  function advanceFromAnchor_(anchorDate, interval, unitKey, today) {
    if (!anchorDate) {
      return null;
    }

    var next = startOfDay_(anchorDate);
    const todayStart = startOfDay_(today || new Date());

    if (next.getTime() >= todayStart.getTime()) {
      return next;
    }

    var guard = 0;
    while (next.getTime() < todayStart.getTime() && guard < 5000) {
      next = startOfDay_(addInterval_(next, interval, unitKey));
      guard++;
    }

    return guard >= 5000 ? null : next;
  }

  function calculateFixedDueDay_(item, today) {
    const dueDay = parseInt(item.dueDayOfMonth, 10);
    if (!dueDay || dueDay < 1 || dueDay > 31) {
      return null;
    }

    const todayStart = startOfDay_(today || new Date());
    var year = todayStart.getFullYear();
    var month = todayStart.getMonth();
    var candidate = buildDateWithDay_(year, month, dueDay);

    if (item.lastPaidDate) {
      const lastPaidStart = startOfDay_(item.lastPaidDate);
      while (candidate.getTime() <= lastPaidStart.getTime()) {
        month += 1;
        candidate = buildDateWithDay_(year, month, dueDay);
      }
      return candidate;
    }

    if (candidate.getTime() < todayStart.getTime()) {
      month += 1;
      candidate = buildDateWithDay_(year, month, dueDay);
    }

    return candidate;
  }

  function calculateNextDueDate_(item, today) {
    if (!item) {
      return null;
    }

    const anchorType = item.anchorTypeKey || normalizeUnit_(item.anchorType);
    const interval = item.recurrenceInterval || 1;
    const unitKey = normalizeUnit_(item.recurrenceUnit);

    if (anchorType === ANCHOR_MANUAL) {
      return item.nextDueDate ? startOfDay_(item.nextDueDate) : null;
    }

    if (anchorType === ANCHOR_FIXED_DUE_DAY) {
      return calculateFixedDueDay_(item, today);
    }

    if (anchorType === ANCHOR_LAST_PAYMENT) {
      const baseDate = item.lastPaidDate || item.anchorDate;
      if (!baseDate) {
        return item.nextDueDate ? startOfDay_(item.nextDueDate) : null;
      }
      var next = startOfDay_(addInterval_(baseDate, interval, unitKey));
      const todayStart = startOfDay_(today || new Date());
      var guard = 0;
      while (next.getTime() < todayStart.getTime() && guard < 5000) {
        next = startOfDay_(addInterval_(next, interval, unitKey));
        guard++;
      }
      return guard >= 5000 ? null : next;
    }

    if (anchorType === ANCHOR_FIXED_DATE) {
      return advanceFromAnchor_(item.anchorDate, interval, unitKey, today);
    }

    return item.nextDueDate ? startOfDay_(item.nextDueDate) : null;
  }

  function recalculateItem_(item, today) {
    const nextDueDate = calculateNextDueDate_(item, today);
    if (!nextDueDate) {
      return { success: false, skipped: true, reason: 'Could not calculate next due date' };
    }

    const wrote = RenewalTrackerModule.writeNextDueDate(item.rowNumber, nextDueDate);
    return {
      success: wrote,
      rowNumber: item.rowNumber,
      itemName: item.itemName,
      nextDueDate: nextDueDate,
    };
  }

  function recalculateAll_() {
    const items = RenewalTrackerModule.readAllItems();
    const results = [];

    items.forEach(function (item) {
      if (item.anchorTypeKey === ANCHOR_MANUAL) {
        results.push({
          rowNumber: item.rowNumber,
          skipped: true,
          reason: 'Manual only',
        });
        return;
      }

      try {
        results.push(recalculateItem_(item));
      } catch (err) {
        ErrorLogModule.error('RenewalDueDateModule', 'recalculateAll_', 'Row failed', {
          rowNumber: item.rowNumber,
          error: err,
        });
      }
    });

    return results;
  }

  function daysUntilDue_(nextDueDate, today) {
    if (!nextDueDate) {
      return null;
    }
    const due = startOfDay_(nextDueDate);
    const now = startOfDay_(today || new Date());
    const diffMs = due.getTime() - now.getTime();
    return Math.round(diffMs / (24 * 60 * 60 * 1000));
  }

  return {
    calculateNextDueDate: calculateNextDueDate_,
    recalculateItem: recalculateItem_,
    recalculateAll: recalculateAll_,
    daysUntilDue: daysUntilDue_,
  };
})();
