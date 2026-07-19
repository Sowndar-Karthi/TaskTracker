/**
 * Standalone dependency checks — no imports, safe to run when other files fail to load.
 * Run debugWhatIsBroken() from the script editor when logs only show "(anonymous) @".
 */

function debugWhatIsBroken() {
  var lines = [];
  lines.push('=== What is broken? ===');
  lines.push('');

  if (typeof SheetColumns === 'undefined') {
    lines.push('ROOT CAUSE: SheetColumns is not defined');
    lines.push('');
    lines.push('Why the log is vague:');
    lines.push('  Apps Script often shows only "(anonymous) @" because code runs inside');
    lines.push('  module wrappers. The real failure is RenewalReminderHandler.gs using');
    lines.push('  SheetColumns when renewal due-date columns are read.');
    lines.push('');
    lines.push('FIX:');
    lines.push('  1. In Apps Script, open SheetColumns.gs');
    lines.push('  2. Delete everything in that file');
    lines.push('  3. Paste the FULL SheetColumns.gs from your PC (121 lines)');
    lines.push('  4. First line must be: const SheetColumns = (function () {');
    lines.push('  5. Last lines must include RENEWAL and end with: })();');
  lines.push('  6. Confirm all project .gs files exist in the project sidebar');
  lines.push('  7. Run debugRunFullDiagnostics() — expect OVERALL: PASS');
  } else if (!SheetColumns.RENEWAL || SheetColumns.RENEWAL.LAST_COLUMN !== 14) {
    lines.push('ROOT CAUSE: SheetColumns.gs is incomplete (RENEWAL section missing or wrong)');
    lines.push('');
    lines.push('FIX: Replace SheetColumns.gs with the full file from your PC.');
    lines.push('Expected: SheetColumns.RENEWAL.LAST_COLUMN === 14');
    lines.push(
      'Actual: ' + (SheetColumns.RENEWAL ? SheetColumns.RENEWAL.LAST_COLUMN : '(RENEWAL missing)')
    );
  } else if (!SheetColumns.TRAIN_BOOKING || SheetColumns.TRAIN_BOOKING.LAST_COLUMN !== 17) {
    lines.push('ROOT CAUSE: SheetColumns.gs missing TRAIN_BOOKING (train IRCTC import)');
    lines.push('FIX: Paste latest SheetColumns.gs from your PC (includes TRAIN_ROUTE + TRAIN_BOOKING).');
  } else {
    lines.push('OK: SheetColumns, RENEWAL, and TRAIN_BOOKING look correct');
  }

  lines.push('');
  lines.push('Other required modules:');

  var requiredModules = [
    { name: 'AppConfig', file: 'AppConfig.gs' },
    { name: 'SpreadsheetUtils', file: 'SpreadsheetUtils.gs' },
    { name: 'RenewalTrackerModule', file: 'RenewalTrackerModule.gs' },
    { name: 'RenewalDueDateModule', file: 'RenewalDueDateModule.gs' },
    { name: 'RenewalReminderModule', file: 'RenewalReminderModule.gs' },
    { name: 'RenewalReminderHandler', file: 'RenewalReminderHandler.gs' },
    { name: 'TrainBookingHistoryModule', file: 'TrainBookingHistoryModule.gs' },
    { name: 'TrainIrctcParseModule', file: 'TrainIrctcParseModule.gs' },
    { name: 'TrainIrctcImportModule', file: 'TrainIrctcImportModule.gs' },
    { name: 'TrainIrctcImportHandler', file: 'TrainIrctcImportHandler.gs' },
    { name: 'ScriptDiagnosticsModule', file: 'ScriptDiagnostics.gs' },
  ];

  var missingCount = 0;
  requiredModules.forEach(function (item) {
    var ok = false;
    try {
      ok = typeof eval(item.name) !== 'undefined';
    } catch (err) {
      ok = false;
    }
    if (!ok) {
      missingCount++;
      lines.push('  FAIL  ' + item.name + '  — add or fix ' + item.file);
    } else {
      lines.push('  OK    ' + item.name);
    }
  });

  if (missingCount > 0) {
    lines.push('');
    lines.push('Missing modules usually mean the .gs file is not in Apps Script or failed to load.');
  }

  lines.push('');
  lines.push('Next step: run debugRunFullDiagnostics() for the full project-file check.');

  var report = lines.join('\n');
  Logger.log(report);

  try {
    SpreadsheetApp.getUi().alert('Dependency check', report, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (uiErr) {
    Logger.log('(Open the spreadsheet and run again to see a popup alert.)');
  }

  return report;
}
