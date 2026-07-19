/**
 * Train Booking History sheet — headers, append, dedupe, sort.
 * Train Completed Journeys uses the same column layout.
 */
const TrainBookingHistoryModule = (function () {
  'use strict';

  const HEADERS = [
    'From',
    'To',
    'Date of Journey',
    'Name',
    'Status',
    'Coach',
    'Seat / Berth',
    'Class',
    'PNR',
    'Train No',
    'Scheduled Departure',
    'Date of Boarding',
    'Transaction ID',
    'Date & Time of Booking',
    'User Id',
    'Passenger Mobile',
    'Gmail Thread ID',
  ];

  function getHeaders_() {
    return HEADERS.slice();
  }

  function getSpreadsheet_() {
    return AppConfig.getSpreadsheet();
  }

  function getBookingSheet_() {
    const spreadsheet = getSpreadsheet_();
    if (!spreadsheet) {
      console.error('TrainBookingHistoryModule.getBookingSheet_: spreadsheet unavailable');
      return null;
    }
    let sheet = SpreadsheetUtils.getSheetByName(
      spreadsheet,
      AppConfig.getTrainBookingHistorySheetName()
    );
    if (!sheet) {
      sheet = spreadsheet.insertSheet(AppConfig.getTrainBookingHistorySheetName());
    }
    ensureHeaders_(sheet);
    return sheet;
  }

  function validateSheetExists_() {
    const spreadsheet = getSpreadsheet_();
    if (!spreadsheet) {
      return { valid: false, error: 'Spreadsheet unavailable' };
    }
    const sheet = SpreadsheetUtils.getSheetByName(
      spreadsheet,
      AppConfig.getTrainBookingHistorySheetName()
    );
    if (!sheet) {
      return {
        valid: false,
        error: 'Sheet not found: ' + AppConfig.getTrainBookingHistorySheetName(),
      };
    }
    return { valid: true, error: '' };
  }

  function ensureHeaders_(sheet) {
    if (!sheet) {
      console.error('TrainBookingHistoryModule.ensureHeaders_: sheet is null');
      return;
    }

    try {
      const cols = SheetColumns.TRAIN_BOOKING.LAST_COLUMN;
      if (sheet.getLastRow() > 0) {
        const first = sheet.getRange(1, 1, 1, cols).getValues()[0];
        if (String(first[0] || '').trim() === HEADERS[0]) {
          return;
        }
      }
      sheet
        .getRange(1, 1, 1, HEADERS.length)
        .setValues([HEADERS])
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    } catch (err) {
      console.error('TrainBookingHistoryModule.ensureHeaders_ failed', {
        error: err,
        stack: err && err.stack,
      });
    }
  }

  function rowExists_(sheet, parsed) {
    if (!sheet || !parsed) {
      return false;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < AppConfig.getDataStartRow()) {
      return false;
    }

    const cols = SheetColumns.TRAIN_BOOKING;
    const data = sheet
      .getRange(AppConfig.getDataStartRow(), 1, lastRow - AppConfig.getHeaderRow(), cols.LAST_COLUMN)
      .getValues();

    const pnrNeed = String(parsed.pnr || '').trim();
    const txnNeed = String(parsed.transactionId || '').trim();
    const threadNeed = String(parsed.threadId || '').trim();
    const fromNeed = String(parsed.from || '').trim().toLowerCase();
    const toNeed = String(parsed.to || '').trim().toLowerCase();
    const dateNeed = String(parsed.dateOfJourney || '').trim().toLowerCase();

    for (var i = 0; i < data.length; i++) {
      const rowPnr = String(data[i][cols.PNR - 1] || '').trim();
      const rowTxn = String(data[i][cols.TRANSACTION_ID - 1] || '').trim();
      const rowThread = String(data[i][cols.GMAIL_THREAD_ID - 1] || '').trim();
      const rowFrom = String(data[i][cols.FROM - 1] || '').trim().toLowerCase();
      const rowTo = String(data[i][cols.TO - 1] || '').trim().toLowerCase();
      const rowDate = String(data[i][cols.DATE_OF_JOURNEY - 1] || '').trim().toLowerCase();

      if (pnrNeed && rowPnr && rowPnr === pnrNeed) {
        return true;
      }
      if (txnNeed && rowTxn && rowTxn === txnNeed) {
        return true;
      }
      if (threadNeed && rowThread && rowThread === threadNeed) {
        return true;
      }
      if (
        pnrNeed &&
        fromNeed &&
        toNeed &&
        dateNeed &&
        rowPnr === pnrNeed &&
        rowFrom === fromNeed &&
        rowTo === toNeed &&
        rowDate === dateNeed
      ) {
        return true;
      }
    }
    return false;
  }

  function appendParsedRow_(sheet, parsed) {
    if (!sheet || !parsed) {
      console.error('TrainBookingHistoryModule.appendParsedRow_: invalid args', {
        hasSheet: !!sheet,
        hasParsed: !!parsed,
      });
      return false;
    }

    const cols = SheetColumns.TRAIN_BOOKING;
    const row = new Array(cols.LAST_COLUMN);
    for (var i = 0; i < row.length; i++) {
      row[i] = '';
    }

    row[cols.FROM - 1] = parsed.from || '';
    row[cols.TO - 1] = parsed.to || '';
    row[cols.DATE_OF_JOURNEY - 1] = parsed.dateOfJourney || '';
    row[cols.NAME - 1] = parsed.name || '';
    row[cols.STATUS - 1] = parsed.status || '';
    row[cols.COACH - 1] = parsed.coach || '';
    row[cols.SEAT_BERTH - 1] = parsed.seatBerth || '';
    row[cols.CLASS - 1] = parsed.travelClass || '';
    row[cols.PNR - 1] = parsed.pnr || '';
    row[cols.TRAIN_NO - 1] = parsed.trainNo || '';
    row[cols.SCHEDULED_DEPARTURE - 1] = parsed.scheduledDeparture || '';
    row[cols.DATE_OF_BOARDING - 1] = parsed.dateOfBoarding || '';
    row[cols.TRANSACTION_ID - 1] = parsed.transactionId || '';
    row[cols.BOOKING_DATETIME - 1] = parsed.bookingDateTime || '';
    row[cols.USER_ID - 1] = parsed.userId || '';
    row[cols.PASSENGER_MOBILE - 1] = parsed.passengerMobile || '';
    row[cols.GMAIL_THREAD_ID - 1] = parsed.threadId || '';

    try {
      sheet.appendRow(row);
      return true;
    } catch (err) {
      console.error('TrainBookingHistoryModule.appendParsedRow_ failed', {
        pnr: parsed.pnr,
        error: err,
        stack: err && err.stack,
      });
      ErrorLogModule.error('TrainBookingHistoryModule', 'appendParsedRow_', String(err), {
        pnr: parsed.pnr,
        taskKey: parsed.pnr || parsed.transactionId || '',
      });
      return false;
    }
  }

  /**
   * Sort Booking History by journey/travel date ascending (chronological).
   * Handles Date cells and text formats like 26-Jul-2026 / 19-July-2026.
   */
  function sortByJourneyDate_(sheet) {
    const targetSheet = sheet || getBookingSheet_();
    if (!targetSheet || targetSheet.getLastRow() < AppConfig.getDataStartRow()) {
      return { success: true, sorted: 0 };
    }

    try {
      const lastRow = targetSheet.getLastRow();
      const lastCol = Math.max(targetSheet.getLastColumn(), SheetColumns.TRAIN_BOOKING.LAST_COLUMN);
      const headers = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0];

      var dateColIndex = findHeaderIndex_(headers, ['Date of Journey', 'Travel Date']);
      if (dateColIndex < 0) {
        dateColIndex = SheetColumns.TRAIN_BOOKING.DATE_OF_JOURNEY - 1;
      }

      const numRows = lastRow - AppConfig.getHeaderRow();
      const data = targetSheet.getRange(AppConfig.getDataStartRow(), 1, numRows, lastCol).getValues();

      const decorated = data.map(function (row, index) {
        var parsed = null;
        if (typeof TrainRouteModule !== 'undefined' && TrainRouteModule.parseDate) {
          parsed = TrainRouteModule.parseDate(row[dateColIndex]);
        } else if (row[dateColIndex] instanceof Date && !isNaN(row[dateColIndex].getTime())) {
          parsed = row[dateColIndex];
        }
        return {
          row: row,
          sortKey: parsed ? parsed.getTime() : Number.MAX_SAFE_INTEGER,
          originalIndex: index,
        };
      });

      decorated.sort(function (a, b) {
        if (a.sortKey !== b.sortKey) {
          return a.sortKey - b.sortKey;
        }
        return a.originalIndex - b.originalIndex;
      });

      const sorted = decorated.map(function (item) {
        return item.row;
      });

      targetSheet.getRange(AppConfig.getDataStartRow(), 1, numRows, lastCol).setValues(sorted);

      console.log('TrainBookingHistoryModule.sortByJourneyDate_: sorted ' + numRows + ' row(s)', {
        dateColumn: dateColIndex + 1,
      });
      return { success: true, sorted: numRows };
    } catch (err) {
      console.error('TrainBookingHistoryModule.sortByJourneyDate_ failed', {
        error: err,
        stack: err && err.stack,
      });
      ErrorLogModule.error('TrainBookingHistoryModule', 'sortByJourneyDate_', String(err), {
        sheetName: AppConfig.getTrainBookingHistorySheetName(),
      });
      return { success: false, sorted: 0, error: String(err) };
    }
  }

  function stationCode_(value) {
    const m = String(value || '').match(/\(\s*([A-Za-z0-9]+)\s*\)/);
    if (m) {
      return m[1].toUpperCase();
    }
    return SpreadsheetUtils.normalizeText(value);
  }

  function datesMatch_(a, b) {
    if (!a || !b) {
      return false;
    }
    const tz = AppConfig.getTrainDateTimezone();
    function toKey(value) {
      if (value instanceof Date && !isNaN(value.getTime())) {
        return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
      }
      if (typeof TrainRouteModule !== 'undefined' && TrainRouteModule.parseDate) {
        const parsed = TrainRouteModule.parseDate(value);
        if (parsed) {
          return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
        }
      }
      return SpreadsheetUtils.normalizeText(
        String(value).replace(/\s+Scheduled\s+Departure[\s\S]*$/i, '')
      );
    }
    return toKey(a) === toKey(b);
  }
  function findHeaderIndex_(headers, candidates) {
    for (var i = 0; i < headers.length; i++) {
      const h = SpreadsheetUtils.normalizeText(headers[i]);
      for (var c = 0; c < candidates.length; c++) {
        if (h === SpreadsheetUtils.normalizeText(candidates[c])) {
          return i;
        }
      }
    }
    return -1;
  }

  /**
   * True if Booking History already has a trip matching from/to/travel date.
   * Direction-agnostic: TBM→KLT and KLT→TBM on the same date both count as booked.
   * Supports IRCTC headers or History-style headers (From Station / Travel Date).
   */
  function hasMatchingBooking_(fromStation, toStation, travelDate) {
    if (!fromStation || !toStation || !travelDate) {
      return false;
    }

    const sheet = getBookingSheet_();
    if (!sheet || sheet.getLastRow() < AppConfig.getDataStartRow()) {
      return false;
    }

    try {
      const lastCol = Math.max(sheet.getLastColumn(), SheetColumns.TRAIN_BOOKING.LAST_COLUMN);
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var fromIdx = findHeaderIndex_(headers, ['From', 'From Station']);
      var toIdx = findHeaderIndex_(headers, ['To', 'To Station']);
      var dateIdx = findHeaderIndex_(headers, ['Date of Journey', 'Travel Date']);

      if (fromIdx < 0) {
        fromIdx = SheetColumns.TRAIN_BOOKING.FROM - 1;
      }
      if (toIdx < 0) {
        toIdx = SheetColumns.TRAIN_BOOKING.TO - 1;
      }
      if (dateIdx < 0) {
        dateIdx = SheetColumns.TRAIN_BOOKING.DATE_OF_JOURNEY - 1;
      }

      const numRows = sheet.getLastRow() - AppConfig.getHeaderRow();
      const data = sheet.getRange(AppConfig.getDataStartRow(), 1, numRows, lastCol).getValues();
      const fromCode = stationCode_(fromStation);
      const toCode = stationCode_(toStation);

      for (var i = 0; i < data.length; i++) {
        const rowFrom = stationCode_(data[i][fromIdx]);
        const rowTo = stationCode_(data[i][toIdx]);
        const rowDate = data[i][dateIdx];
        if (!rowFrom || !rowTo || !datesMatch_(rowDate, travelDate)) {
          continue;
        }
        const sameDirection = rowFrom === fromCode && rowTo === toCode;
        const reverseDirection = rowFrom === toCode && rowTo === fromCode;
        if (sameDirection || reverseDirection) {
          return true;
        }
      }
    } catch (err) {
      console.error('TrainBookingHistoryModule.hasMatchingBooking_ failed', {
        error: err,
        stack: err && err.stack,
      });
    }
    return false;
  }

  return {
    HEADERS: HEADERS,
    getHeaders: getHeaders_,
    getBookingSheet: getBookingSheet_,
    ensureHeaders: ensureHeaders_,
    validateSheetExists: validateSheetExists_,
    rowExists: rowExists_,
    appendParsedRow: appendParsedRow_,
    sortByJourneyDate: sortByJourneyDate_,
    hasMatchingBooking: hasMatchingBooking_,
  };
})();

/**
 * Manual: sort Train Booking History by Date of Journey / Travel Date (oldest first).
 */
function runSortTrainBookingHistory() {
  const result = TrainBookingHistoryModule.sortByJourneyDate();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      result.success
        ? 'Sorted ' + (result.sorted || 0) + ' row(s) by journey date'
        : 'Sort failed: ' + (result.error || 'unknown'),
      'Train Tracker',
      6
    );
  } catch (uiErr) {
    // toast may be unavailable
  }
  return result;
}
