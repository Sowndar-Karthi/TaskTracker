/**
 * Archives past journeys from Train Booking History → Train Completed Journeys.
 * Runs after the travel date has passed (travel date &lt; today).
 */
const TrainJourneyArchiveModule = (function () {
  'use strict';

  function getTimezone_() {
    return AppConfig.getTrainDateTimezone() || 'Asia/Kolkata';
  }

  function today_() {
    if (typeof TrainRouteModule !== 'undefined' && TrainRouteModule.startOfDay) {
      return TrainRouteModule.startOfDay(new Date());
    }
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    return d;
  }

  function parseJourneyDate_(value) {
    if (typeof TrainRouteModule !== 'undefined' && TrainRouteModule.parseDate) {
      return TrainRouteModule.parseDate(value);
    }
    if (value instanceof Date && !isNaN(value.getTime())) {
      const d = new Date(value.getTime());
      d.setHours(0, 0, 0, 0);
      return d;
    }
    return null;
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

  function getCompletedSheet_() {
    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      return null;
    }
    let sheet = SpreadsheetUtils.getSheetByName(
      spreadsheet,
      AppConfig.getTrainCompletedJourneysSheetName()
    );
    if (!sheet) {
      sheet = spreadsheet.insertSheet(AppConfig.getTrainCompletedJourneysSheetName());
    }
    return sheet;
  }

  function ensureCompletedHeaders_(completedSheet, bookingSheet) {
    if (!completedSheet || !bookingSheet) {
      return;
    }

    try {
      const bookingLastCol = Math.max(
        bookingSheet.getLastColumn(),
        SheetColumns.TRAIN_BOOKING.LAST_COLUMN
      );
      const bookingHeaders = bookingSheet.getRange(1, 1, 1, bookingLastCol).getValues()[0];

      if (completedSheet.getLastRow() > 0) {
        const first = String(completedSheet.getRange(1, 1).getValue() || '').trim();
        if (first && first === String(bookingHeaders[0] || '').trim()) {
          return;
        }
      }

      completedSheet
        .getRange(1, 1, 1, bookingHeaders.length)
        .setValues([bookingHeaders])
        .setFontWeight('bold');
      completedSheet.setFrozenRows(1);
    } catch (err) {
      console.error('TrainJourneyArchiveModule.ensureCompletedHeaders_ failed', {
        error: err,
        stack: err && err.stack,
      });
      // Fallback to IRCTC headers
      if (typeof TrainBookingHistoryModule !== 'undefined') {
        TrainBookingHistoryModule.ensureHeaders(completedSheet);
      }
    }
  }

  function completedHasPnr_(completedSheet, pnr, pnrColIndex) {
    if (!pnr || pnrColIndex < 0 || !completedSheet) {
      return false;
    }
    const lastRow = completedSheet.getLastRow();
    if (lastRow < AppConfig.getDataStartRow()) {
      return false;
    }
    const values = completedSheet
      .getRange(AppConfig.getDataStartRow(), pnrColIndex + 1, lastRow - AppConfig.getHeaderRow(), 1)
      .getValues();
    const need = String(pnr).trim();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === need) {
        return true;
      }
    }
    return false;
  }

  /**
   * Move Booking History rows whose journey date is before today.
   * @returns {{moved: number, skipped: number, errors: number}}
   */
  function archivePastJourneys_() {
    const summary = { moved: 0, skipped: 0, errors: 0 };

    const bookingSheet =
      typeof TrainBookingHistoryModule !== 'undefined'
        ? TrainBookingHistoryModule.getBookingSheet()
        : null;
    if (!bookingSheet) {
      ErrorLogModule.error(
        'TrainJourneyArchiveModule',
        'archivePastJourneys_',
        'Train Booking History unavailable',
        {}
      );
      return summary;
    }

    const completedSheet = getCompletedSheet_();
    if (!completedSheet) {
      ErrorLogModule.error(
        'TrainJourneyArchiveModule',
        'archivePastJourneys_',
        'Train Completed Journeys unavailable',
        {}
      );
      return summary;
    }

    ensureCompletedHeaders_(completedSheet, bookingSheet);

    const lastRow = bookingSheet.getLastRow();
    if (lastRow < AppConfig.getDataStartRow()) {
      return summary;
    }

    const lastCol = Math.max(bookingSheet.getLastColumn(), 1);
    const headers = bookingSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var dateIdx = findHeaderIndex_(headers, ['Date of Journey', 'Travel Date']);
    if (dateIdx < 0) {
      dateIdx = SheetColumns.TRAIN_BOOKING.DATE_OF_JOURNEY - 1;
    }
    var pnrIdx = findHeaderIndex_(headers, ['PNR', 'PNR No']);
    if (pnrIdx < 0) {
      pnrIdx = SheetColumns.TRAIN_BOOKING.PNR - 1;
    }

    const today = today_();
    const completedLastCol = Math.max(completedSheet.getLastColumn(), lastCol);

    // Bottom-up so deletes don't shift unprocessed rows
    for (var rowIndex = lastRow; rowIndex >= AppConfig.getDataStartRow(); rowIndex--) {
      try {
        const rowValues = bookingSheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
        const journeyDate = parseJourneyDate_(rowValues[dateIdx]);

        if (!journeyDate) {
          summary.skipped++;
          continue;
        }

        // Journey finished = travel calendar day is fully in the past
        if (journeyDate.getTime() >= today.getTime()) {
          summary.skipped++;
          continue;
        }

        const pnr = pnrIdx >= 0 ? rowValues[pnrIdx] : '';
        if (pnr && completedHasPnr_(completedSheet, pnr, pnrIdx)) {
          bookingSheet.deleteRow(rowIndex);
          summary.moved++;
          console.log('TrainJourneyArchiveModule: removed duplicate past booking row', {
            rowIndex: rowIndex,
            pnr: pnr,
          });
          continue;
        }

        const outRow = rowValues.slice();
        while (outRow.length < completedLastCol) {
          outRow.push('');
        }

        const nextRow = Math.max(completedSheet.getLastRow() + 1, AppConfig.getDataStartRow());
        completedSheet.getRange(nextRow, 1, 1, outRow.length).setValues([outRow]);
        bookingSheet.deleteRow(rowIndex);
        summary.moved++;

        console.log('TrainJourneyArchiveModule: archived journey', {
          rowIndex: rowIndex,
          journeyDate: Utilities.formatDate(journeyDate, getTimezone_(), 'dd-MMM-yyyy'),
          pnr: pnr,
        });
      } catch (err) {
        summary.errors++;
        console.error('TrainJourneyArchiveModule.archivePastJourneys_: row failed', {
          rowIndex: rowIndex,
          error: err,
          stack: err && err.stack,
        });
        ErrorLogModule.error(
          'TrainJourneyArchiveModule',
          'archivePastJourneys_',
          'Row archive failed: ' + err,
          {
            sheetName: AppConfig.getTrainBookingHistorySheetName(),
            rowIndex: rowIndex,
          }
        );
      }
    }

    console.log('TrainJourneyArchiveModule.archivePastJourneys_', summary);

    if (summary.moved > 0) {
      try {
        TrainBookingHistoryModule.sortByJourneyDate(bookingSheet);
      } catch (sortErr) {
        console.error('TrainJourneyArchiveModule: re-sort after archive failed', {
          error: sortErr,
        });
      }
    }

    return summary;
  }

  return {
    archivePastJourneys: archivePastJourneys_,
    getCompletedSheet: getCompletedSheet_,
  };
})();
