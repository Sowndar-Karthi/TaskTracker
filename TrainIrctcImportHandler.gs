/**
 * Scheduled and manual entry points for IRCTC → Train Booking History import.
 */
const TrainIrctcImportHandler = (function () {
  'use strict';

  function processScheduledImport_() {
    try {
      const result = TrainIrctcImportModule.extractEmailsWithLock();
      console.log('TrainIrctcImportHandler.processScheduledImport_', result);
      return result;
    } catch (err) {
      console.error('TrainIrctcImportHandler.processScheduledImport_ failed', {
        error: err,
        stack: err && err.stack,
      });
      ErrorLogModule.error(
        'TrainIrctcImportHandler',
        'processScheduledImport_',
        String(err),
        { triggerSource: 'scheduled' }
      );
      return { appended: 0, skipped: 0, error: String(err) };
    }
  }

  function runImportNow_() {
    ErrorLogModule.setTriggerSource('manual');
    try {
      const result = TrainIrctcImportModule.extractEmails();
      console.log('TrainIrctcImportHandler.runImportNow_', result);

      try {
        if (result && result.appended > 0) {
          SpreadsheetApp.getActiveSpreadsheet().toast(
            'Imported ' +
              result.appended +
              ' IRCTC booking(s) into ' +
              AppConfig.getTrainBookingHistorySheetName(),
            'Train Tracker',
            8
          );
        } else if (result && !result.error) {
          SpreadsheetApp.getActiveSpreadsheet().toast(
            'No new unread IRCTC bookings found',
            'Train Tracker',
            5
          );
        }
      } catch (uiErr) {
        // UI toast may be unavailable in some contexts
      }

      return result;
    } finally {
      ErrorLogModule.clearTriggerSource();
    }
  }

  return {
    processScheduledImport: processScheduledImport_,
    runImportNow: runImportNow_,
  };
})();

/**
 * Installable time-driven trigger handler — do not rename (wired in TriggerManager).
 */
function processScheduledIrctcImport_() {
  return TrainIrctcImportHandler.processScheduledImport();
}

/**
 * Manual run from Apps Script editor — import unread IRCTC emails now.
 */
function runExtractIrctcEmails() {
  return TrainIrctcImportHandler.runImportNow();
}

/**
 * Manual test — parse sample IRCTC bodies (no Gmail required).
 */
function testParseSampleIrctcEmail() {
  const tabSample =
    'From: ticketadmin@irctc.co.in\n' +
    'To: dragonkarthi00m@gmail.com\n' +
    'From 01st July 2025, Aadhaar authentication of IRCTC user profile is mandatory.\n' +
    'Ticket Confirmation\n' +
    'Dear SOWNDAR(User Id: Sowndar92),\n' +
    'PNR No. :\t4138727671\tTrain No. / Name :\t16160 / MAQ CHENNAI EXP\tQuota :\tGENERAL\n' +
    'Transaction ID :\t100006604757184\tDate & Time of Booking :\t27-May-2026 08:47:55 AM HRS\tClass :\tSLEEPER CLASS\n' +
    'From :\tKULITALAI (KLT)\tDate of Journey :\t26-Jul-2026\tTo :\tTAMBARAM (TBM)\n' +
    'Boarding At :\tKLT\tDate Of Boarding :\t26-Jul-2026\tScheduled Departure* :\t26-Jul-2026 20:25\n' +
    'Passenger Mobile No :\t9994476164\tDistance :\t349KM\n' +
    'Passenger Details\n' +
    'Sl. No.\tName\tAge\tGender\tCatering Service Option\tStatus\tCoach\tSeat / Berth / WL No\n' +
    '1\tSOWNDAR\t33\tMale\tN/A\tCNF\tS4 \t40 \n' +
    'Fare Details (Inclusive of GST)\n';

  const htmlLikeSample =
    'Ticket Confirmation\n' +
    'PNR No. : 4138727671 Train No. / Name : 16160 / MAQ CHENNAI EXP\n' +
    'Transaction ID : 100006604757184 Date & Time of Booking : 27-May-2026 08:47:55 AM HRS Class : SLEEPER CLASS\n' +
    'From : KULITALAI (KLT) Date of Journey : 26-Jul-2026 To : TAMBARAM (TBM)\n' +
    'Boarding At : KLT Date Of Boarding : 26-Jul-2026 Scheduled Departure* : 26-Jul-2026 20:25\n' +
    'Passenger Mobile No : 9994476164 Distance : 349KM\n' +
    'Passenger Details\n' +
    '1 SOWNDAR 33 Male N/A CNF S4 40\n' +
    'Fare Details\n';

  const msg = {
    getDate: function () {
      return new Date();
    },
  };

  Logger.log('Tab body:\n' + JSON.stringify(TrainIrctcParseModule.parseEmailBody(tabSample, msg), null, 2));
  Logger.log(
    'HTML-like body:\n' + JSON.stringify(TrainIrctcParseModule.parseEmailBody(htmlLikeSample, msg), null, 2)
  );
}
