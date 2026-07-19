/**
 * Reads Active rows from the Train Route sheet (weekly + one-time).
 */
const TrainRouteModule = (function () {
  'use strict';

  const DAY_NAME_TO_INDEX = {
    sunday: 0,
    monday: 1,
    tuesday: 2,
    wednesday: 3,
    thursday: 4,
    friday: 5,
    saturday: 6,
  };

  function getTimezone_() {
    return AppConfig.getTrainDateTimezone() || 'Asia/Kolkata';
  }

  function startOfDay_(date) {
    const copy = new Date(date.getTime());
    copy.setHours(0, 0, 0, 0);
    return copy;
  }

  function parseDate_(value) {
    if (value === null || value === undefined || value === '') {
      return null;
    }
    if (value instanceof Date && !isNaN(value.getTime())) {
      return startOfDay_(value);
    }
    const text = String(value)
      .replace(/\s+Scheduled\s+Departure[\s\S]*$/i, '')
      .trim();
    if (!text) {
      return null;
    }

    const formats = ['d-MMM-yyyy', 'dd-MMM-yyyy', 'd/MMM/yyyy', 'dd/MMM/yyyy', 'd-MMMM-yyyy', 'dd-MMMM-yyyy'];
    for (var i = 0; i < formats.length; i++) {
      try {
        const parsed = Utilities.parseDate(text, getTimezone_(), formats[i]);
        if (parsed && !isNaN(parsed.getTime())) {
          return startOfDay_(parsed);
        }
      } catch (err) {
        // try next
      }
    }

    const fallback = new Date(text);
    return isNaN(fallback.getTime()) ? null : startOfDay_(fallback);
  }

  function formatDate_(date) {
    if (!date) {
      return '';
    }
    return Utilities.formatDate(date, getTimezone_(), 'dd-MMM-yyyy');
  }

  function parseReminderDays_(raw) {
    if (raw === null || raw === undefined || String(raw).trim() === '') {
      return [];
    }
    return String(raw)
      .split(',')
      .map(function (part) {
        return parseInt(String(part).trim(), 10);
      })
      .filter(function (n) {
        return !isNaN(n) && n >= 0;
      });
  }

  function parseAlertTimes_(raw) {
    if (typeof EmailSettingsModule !== 'undefined' && EmailSettingsModule.parseAlertTimes) {
      return EmailSettingsModule.parseAlertTimes(raw);
    }
    if (!raw || String(raw).trim() === '') {
      return [];
    }
    return String(raw)
      .split(',')
      .map(function (t) {
        return String(t).trim();
      })
      .filter(function (t) {
        return t !== '';
      });
  }

  function parseEmails_(raw) {
    if (typeof EmailSettingsModule !== 'undefined' && EmailSettingsModule.parseEmailList) {
      return EmailSettingsModule.parseEmailList(raw);
    }
    return [];
  }

  function isActive_(value) {
    const n = SpreadsheetUtils.normalizeText(value);
    return n === 'yes' || n === 'y' || n === 'true' || n === '1' || n === 'active';
  }

  function normalizeRecurrence_(value) {
    const n = SpreadsheetUtils.normalizeText(value);
    if (n === 'one-time' || n === 'onetime' || n === 'once' || n === 'one time') {
      return 'one-time';
    }
    if (n === 'weekly' || n === 'week') {
      return 'weekly';
    }
    return n;
  }

  function dayOfWeekIndex_(value) {
    const n = SpreadsheetUtils.normalizeText(value);
    if (n in DAY_NAME_TO_INDEX) {
      return DAY_NAME_TO_INDEX[n];
    }
    return -1;
  }

  /**
   * Next travel date on/after today for a weekly day name.
   */
  function nextWeeklyTravelDate_(dayName, fromDate) {
    const targetDow = dayOfWeekIndex_(dayName);
    if (targetDow < 0) {
      return null;
    }
    const start = startOfDay_(fromDate || new Date());
    const currentDow = start.getDay();
    var add = (targetDow - currentDow + 7) % 7;
    const result = new Date(start.getTime());
    result.setDate(result.getDate() + add);
    return result;
  }

  function resolveTravelDate_(route, today) {
    if (!route) {
      return null;
    }
    if (route.recurrence === 'one-time') {
      if (!route.travelDate) {
        return null;
      }
      if (route.travelDate.getTime() < today.getTime()) {
        return null;
      }
      return route.travelDate;
    }
    if (route.recurrence === 'weekly') {
      return nextWeeklyTravelDate_(route.dayOfWeek, today);
    }
    // Fallback: prefer explicit travel date, else weekly day
    if (route.travelDate && route.travelDate.getTime() >= today.getTime()) {
      return route.travelDate;
    }
    return nextWeeklyTravelDate_(route.dayOfWeek, today);
  }

  function normalizeRow_(row, rowIndex) {
    const cols = SheetColumns.TRAIN_ROUTE;
    const recurrence = normalizeRecurrence_(row[cols.RECURRENCE - 1]);
    return {
      rowNumber: rowIndex,
      routeName: String(row[cols.ROUTE_NAME - 1] || '').trim(),
      fromStation: String(row[cols.FROM_STATION - 1] || '').trim(),
      toStation: String(row[cols.TO_STATION - 1] || '').trim(),
      dayOfWeek: String(row[cols.DAY_OF_WEEK - 1] || '').trim(),
      recurrence: recurrence,
      active: isActive_(row[cols.ACTIVE - 1]),
      preferredClass: String(row[cols.PREFERRED_CLASS - 1] || '').trim(),
      travelDate: parseDate_(row[cols.TRAVEL_DATE - 1]),
      preferredTrainNo: String(row[cols.PREFERRED_TRAIN_NO - 1] || '').trim(),
      preferredTrainName: String(row[cols.PREFERRED_TRAIN_NAME - 1] || '').trim(),
      reminderDaysBefore: parseReminderDays_(row[cols.REMINDER_DAYS_BEFORE - 1]),
      alertTimes: parseAlertTimes_(row[cols.ALERT_TIME - 1]),
      reminderEmails: parseEmails_(row[cols.REMINDER_EMAIL - 1]),
      notes: String(row[cols.NOTES - 1] || '').trim(),
    };
  }

  function isValidRoute_(route) {
    if (!route || !route.active) {
      return false;
    }
    if (!route.routeName || !route.fromStation || !route.toStation) {
      return false;
    }
    if (!route.reminderEmails || route.reminderEmails.length === 0) {
      return false;
    }
    if (!route.reminderDaysBefore || route.reminderDaysBefore.length === 0) {
      return false;
    }
    if (!route.alertTimes || route.alertTimes.length === 0) {
      return false;
    }
    if (route.recurrence === 'weekly' && dayOfWeekIndex_(route.dayOfWeek) < 0) {
      return false;
    }
    if (route.recurrence === 'one-time' && !route.travelDate) {
      return false;
    }
    return true;
  }

  function readActiveRoutes_() {
    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      return [];
    }
    const sheet = SpreadsheetUtils.getSheetByName(spreadsheet, AppConfig.getTrainRouteSheetName());
    if (!sheet) {
      ErrorLogModule.warning('TrainRouteModule', 'readActiveRoutes_', 'Train Route sheet not found', {});
      return [];
    }

    const lastRow = SpreadsheetUtils.getLastDataRow(sheet, SheetColumns.TRAIN_ROUTE.ROUTE_NAME);
    if (lastRow < AppConfig.getDataStartRow()) {
      return [];
    }

    const rowCount = lastRow - AppConfig.getHeaderRow();
    const values = sheet
      .getRange(AppConfig.getDataStartRow(), 1, rowCount, SheetColumns.TRAIN_ROUTE.LAST_COLUMN)
      .getValues();

    const routes = [];
    for (var i = 0; i < values.length; i++) {
      try {
        const route = normalizeRow_(values[i], AppConfig.getDataStartRow() + i);
        if (isValidRoute_(route)) {
          routes.push(route);
        }
      } catch (err) {
        console.error('TrainRouteModule.readActiveRoutes_: row failed', {
          row: AppConfig.getDataStartRow() + i,
          error: err,
        });
      }
    }
    return routes;
  }

  function validateSheetExists_() {
    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      return { valid: false, error: 'Spreadsheet unavailable' };
    }
    const sheet = SpreadsheetUtils.getSheetByName(spreadsheet, AppConfig.getTrainRouteSheetName());
    if (!sheet) {
      return { valid: false, error: 'Sheet not found: ' + AppConfig.getTrainRouteSheetName() };
    }
    return { valid: true, error: '' };
  }

  /**
   * Upcoming travel dates from today through today + horizonDays (inclusive).
   * Weekly → every matching weekday in range; One-time → that date if in range.
   */
  function listUpcomingTravelDates_(route, today, horizonDays) {
    const dates = [];
    if (!route) {
      return dates;
    }
    const start = startOfDay_(today || new Date());
    const horizon = typeof horizonDays === 'number' ? horizonDays : 60;
    const end = new Date(start.getTime());
    end.setDate(end.getDate() + horizon);

    if (route.recurrence === 'one-time') {
      if (
        route.travelDate &&
        route.travelDate.getTime() >= start.getTime() &&
        route.travelDate.getTime() <= end.getTime()
      ) {
        dates.push(route.travelDate);
      }
      return dates;
    }

    // Weekly (and fallback): step week by week from next matching day
    var cursor = nextWeeklyTravelDate_(route.dayOfWeek, start);
    if (!cursor) {
      return dates;
    }
    while (cursor.getTime() <= end.getTime()) {
      dates.push(new Date(cursor.getTime()));
      cursor = new Date(cursor.getTime());
      cursor.setDate(cursor.getDate() + 7);
      cursor = startOfDay_(cursor);
    }
    return dates;
  }

  return {
    readActiveRoutes: readActiveRoutes_,
    resolveTravelDate: resolveTravelDate_,
    listUpcomingTravelDates: listUpcomingTravelDates_,
    parseDate: parseDate_,
    formatDate: formatDate_,
    nextWeeklyTravelDate: nextWeeklyTravelDate_,
    validateSheetExists: validateSheetExists_,
    startOfDay: startOfDay_,
  };
})();
