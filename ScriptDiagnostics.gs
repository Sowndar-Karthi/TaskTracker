/**
 * Manual diagnostic tools — verify all project files, modules, exports, and config.
 */
const ScriptDiagnosticsModule = (function () {
  'use strict';

  const PROJECT_FILES = [
  {
    file: 'AppConfig.gs',
    module: 'AppConfig',
    exports: ['getSpreadsheet', 'getSpreadsheetId', 'getRenewalTrackerSheetName', 'getTrainBookingHistorySheetName', 'isTesting'],
  },
  {
    file: 'SheetColumns.gs',
    module: 'SheetColumns',
    exports: ['MAIN', 'SUB', 'MATCHING', 'EMAIL', 'ERROR', 'RENEWAL', 'TRAIN_ROUTE', 'TRAIN_BOOKING'],
  },
  {
    file: 'SpreadsheetUtils.gs',
    module: 'SpreadsheetUtils',
    exports: ['getSheetByName', 'getRowValues', 'normalizeText'],
  },
  {
    file: 'ErrorLogModule.gs',
    module: 'ErrorLogModule',
    exports: ['ensureErrorSheet', 'error', 'warning'],
  },
  {
    file: 'MatchingSheetModule.gs',
    module: 'MatchingSheetModule',
    exports: ['getMappings', 'getSheetRole', 'validateMappingsExist'],
  },
  {
    file: 'SubTaskSyncModule.gs',
    module: 'SubTaskSyncModule',
    exports: ['syncSubTasksForMainRow', 'syncSubRowForMainSheet'],
  },
  {
    file: 'TaskEffortRollupModule.gs',
    module: 'TaskEffortRollupModule',
    exports: ['rollupEffortForMainTaskName', 'rollupEffortForMainRow'],
  },
  {
    file: 'TaskStatusModule.gs',
    module: 'TaskStatusModule',
    exports: ['isCompletionStatus', 'ALL_STATUSES'],
  },
  {
    file: 'TaskSyncHandler.gs',
    module: 'TaskSyncHandler',
    exports: ['shouldProcessEdit', 'processEdit'],
    functions: ['handleSpreadsheetEdit_'],
  },
  {
    file: 'CompletionQueueModule.gs',
    module: 'CompletionQueueModule',
    exports: ['scheduleRow', 'getDueEntries', 'ensureQueueSheet'],
  },
  {
    file: 'CompletionMoveModule.gs',
    module: 'CompletionMoveModule',
    exports: ['processDueEntries', 'moveRowToCompleted'],
  },
  {
    file: 'CompletionTaskHandler.gs',
    module: 'CompletionTaskHandler',
    exports: ['shouldProcessEdit', 'processEdit'],
    functions: ['processScheduledCompletions_'],
  },
  {
    file: 'EmailSettingsModule.gs',
    module: 'EmailSettingsModule',
    exports: ['readAllSettings', 'parseEmailList'],
  },
  {
    file: 'NotificationTypeModule.gs',
    module: 'NotificationTypeModule',
    exports: ['parseNotificationTypes', 'taskMatchesNotificationTypes'],
  },
  {
    file: 'TaskNotificationModule.gs',
    module: 'TaskNotificationModule',
    exports: ['findMatchingTasks', 'isAlertTimeNow', 'planDigestBatches'],
  },
  {
    file: 'NotificationService.gs',
    module: 'NotificationService',
    exports: ['sendScheduledDigest', 'sendImmediateAlert'],
  },
  {
    file: 'NotificationHandler.gs',
    module: 'NotificationHandler',
    exports: ['processScheduledAlerts', 'processEdit'],
    functions: ['processScheduledNotifications_'],
  },
  {
    file: 'TriggerManager.gs',
    module: 'TriggerManager',
    exports: ['initializeTriggers', 'listTriggers', 'removeAllTriggers'],
  },
  {
    file: 'Setup.gs',
    module: null,
    functions: ['setupApp', 'setupTestTriggers', 'teardownApp'],
  },
  {
    file: 'RenewalTrackerModule.gs',
    module: 'RenewalTrackerModule',
    exports: ['readAllItems', 'readRemindableItems', 'validateSheetExists', 'writeNextDueDate'],
  },
  {
    file: 'RenewalDueDateModule.gs',
    module: 'RenewalDueDateModule',
    exports: ['calculateNextDueDate', 'recalculateAll', 'recalculateItem'],
  },
  {
    file: 'RenewalReminderModule.gs',
    module: 'RenewalReminderModule',
    exports: ['processDueReminders', 'sendReminder'],
  },
  {
    file: 'RenewalReminderHandler.gs',
    module: 'RenewalReminderHandler',
    exports: ['processScheduledReminders', 'shouldProcessEdit', 'processEdit'],
    functions: ['processScheduledRenewalReminders_', 'runRenewalRemindersNow', 'runRecalculateRenewalDueDates'],
  },
  {
    file: 'TrainBookingHistoryModule.gs',
    module: 'TrainBookingHistoryModule',
    exports: ['getBookingSheet', 'appendParsedRow', 'rowExists', 'validateSheetExists', 'hasMatchingBooking', 'sortByJourneyDate'],
    functions: ['runSortTrainBookingHistory'],
  },
  {
    file: 'TrainIrctcParseModule.gs',
    module: 'TrainIrctcParseModule',
    exports: ['parseEmailBody', 'getMessageBody', 'clearMatchPlaceCache'],
  },
  {
    file: 'TrainIrctcImportModule.gs',
    module: 'TrainIrctcImportModule',
    exports: ['extractEmails', 'extractEmailsWithLock', 'getGmailQuery'],
  },
  {
    file: 'TrainIrctcImportHandler.gs',
    module: 'TrainIrctcImportHandler',
    exports: ['processScheduledImport', 'runImportNow'],
    functions: ['processScheduledIrctcImport_', 'runExtractIrctcEmails', 'testParseSampleIrctcEmail'],
  },
  {
    file: 'TrainRouteModule.gs',
    module: 'TrainRouteModule',
    exports: ['readActiveRoutes', 'resolveTravelDate', 'listUpcomingTravelDates', 'validateSheetExists'],
  },
  {
    file: 'TrainReminderModule.gs',
    module: 'TrainReminderModule',
    exports: ['processDueReminders', 'bookingOpenDate'],
  },
  {
    file: 'TrainReminderHandler.gs',
    module: 'TrainReminderHandler',
    exports: ['processScheduledReminders', 'runRemindersNow'],
    functions: ['processScheduledTrainReminders_', 'runTrainRemindersNow'],
  },
  {
    file: 'TrainJourneyArchiveModule.gs',
    module: 'TrainJourneyArchiveModule',
    exports: ['archivePastJourneys', 'getCompletedSheet'],
  },
  {
    file: 'TrainJourneyArchiveHandler.gs',
    module: 'TrainJourneyArchiveHandler',
    exports: ['processScheduledArchive', 'runArchiveNow'],
    functions: ['processScheduledTrainArchive_', 'runArchiveTrainJourneysNow'],
  },
  {
    file: 'ScriptDiagnostics.gs',
    module: 'ScriptDiagnosticsModule',
    exports: ['runFullDiagnostics', 'validateProjectFiles'],
    functions: ['debugRunFullDiagnostics', 'debugValidateProjectFiles'],
  },
];

  const EXPECTED_SHEET_COLUMNS = {
    MAIN: 23,
    SUB: 25,
    MATCHING: 4,
    EMAIL: 7,
    ERROR: 12,
    RENEWAL: 14,
    TRAIN_ROUTE: 14,
    TRAIN_BOOKING: 17,
  };

  const EXPECTED_TRIGGERS = [
    'handleSpreadsheetEdit_',
    'processScheduledCompletions_',
    'processScheduledNotifications_',
    'processScheduledRenewalReminders_',
    'processScheduledIrctcImport_',
    'processScheduledTrainReminders_',
    'processScheduledTrainArchive_',
  ];

  function isDefined_(name) {
    try {
      return typeof eval(name) !== 'undefined';
    } catch (err) {
      return false;
    }
  }

  function isFunction_(name) {
    try {
      return typeof eval(name) === 'function';
    } catch (err) {
      return false;
    }
  }

  function hasExport_(moduleName, exportName) {
    if (!isDefined_(moduleName)) {
      return false;
    }
    try {
      const moduleRef = eval(moduleName);
      if (!moduleRef) {
        return false;
      }
      const value = moduleRef[exportName];
      return value !== undefined && value !== null;
    } catch (err) {
      return false;
    }
  }

  function validateProjectFiles_() {
    const fileResults = [];
    var passCount = 0;
    var failCount = 0;

    PROJECT_FILES.forEach(function (spec) {
      const result = {
        file: spec.file,
        module: spec.module || '',
        ok: true,
        issues: [],
      };

      if (spec.module) {
        if (!isDefined_(spec.module)) {
          result.ok = false;
          result.issues.push('Module not loaded — add or fix ' + spec.file);
        } else if (spec.exports) {
          spec.exports.forEach(function (exportName) {
            if (!hasExport_(spec.module, exportName)) {
              result.ok = false;
              result.issues.push('Missing export: ' + spec.module + '.' + exportName);
            }
          });
        }
      }

      if (spec.functions) {
        spec.functions.forEach(function (functionName) {
          if (!isFunction_(functionName)) {
            result.ok = false;
            result.issues.push('Missing function: ' + functionName + '() in ' + spec.file);
          }
        });
      }

      if (result.ok) {
        passCount++;
      } else {
        failCount++;
      }

      fileResults.push(result);
    });

    return {
      totalFiles: PROJECT_FILES.length,
      passCount: passCount,
      failCount: failCount,
      allOk: failCount === 0,
      files: fileResults,
    };
  }

  function validateSheetColumnsStructure_() {
    const checks = [];
    var allOk = true;

    if (!isDefined_('SheetColumns')) {
      return {
        ok: false,
        message: 'SheetColumns is not defined — SheetColumns.gs is missing or broken',
        checks: checks,
      };
    }

    Object.keys(EXPECTED_SHEET_COLUMNS).forEach(function (sectionName) {
      const expectedLastColumn = EXPECTED_SHEET_COLUMNS[sectionName];
      const section = SheetColumns[sectionName];
      const ok = !!(section && section.LAST_COLUMN === expectedLastColumn);
      if (!ok) {
        allOk = false;
      }
      checks.push({
        section: sectionName,
        ok: ok,
        expectedLastColumn: expectedLastColumn,
        actualLastColumn: section ? section.LAST_COLUMN : null,
        message: ok
          ? sectionName + '.LAST_COLUMN = ' + expectedLastColumn
          : sectionName + ' missing or wrong — expected LAST_COLUMN ' + expectedLastColumn,
      });
    });

    return {
      ok: allOk,
      message: allOk ? 'SheetColumns structure OK' : 'SheetColumns.gs incomplete — paste full file from PC',
      checks: checks,
    };
  }

  function validateScriptModules_() {
    const missing = [];
    const found = [];

    PROJECT_FILES.forEach(function (spec) {
      if (!spec.module) {
        return;
      }
      if (isDefined_(spec.module)) {
        found.push(spec.module);
      } else {
        missing.push({
          module: spec.module,
          file: spec.file,
        });
      }
    });

    const sheetColumnsCheck = validateSheetColumnsStructure_();

    return {
      found: found,
      missing: missing,
      sheetColumns: sheetColumnsCheck,
      allLoaded: missing.length === 0 && sheetColumnsCheck.ok,
    };
  }

  function validateTriggers_() {
    if (!isDefined_('TriggerManager')) {
      return { ok: false, error: 'TriggerManager not loaded', installed: [], missing: EXPECTED_TRIGGERS };
    }

    const installed = TriggerManager.listTriggers().map(function (trigger) {
      return trigger.handler;
    });
    const missing = EXPECTED_TRIGGERS.filter(function (handlerName) {
      return installed.indexOf(handlerName) < 0;
    });

    return {
      ok: missing.length === 0,
      installed: installed,
      missing: missing,
      triggers: TriggerManager.listTriggers(),
    };
  }

  function validateSpreadsheetConfig_() {
    const result = {
      spreadsheetId: '',
      spreadsheetName: '',
      spreadsheetAvailable: false,
      tabs: [],
      matchingSheet: null,
      emailSettingsSheet: null,
      renewalTrackerSheet: null,
      errorSheet: null,
      renewalItems: null,
    };

    if (!isDefined_('AppConfig')) {
      result.error = 'AppConfig not loaded';
      return result;
    }

    result.spreadsheetId = AppConfig.getSpreadsheetId() || '';

    try {
      const spreadsheet = AppConfig.getSpreadsheet();
      if (!spreadsheet) {
        result.error = 'Spreadsheet unavailable';
        return result;
      }

      result.spreadsheetAvailable = true;
      result.spreadsheetName = spreadsheet.getName();
      result.spreadsheetId = result.spreadsheetId || spreadsheet.getId();
      result.tabs = spreadsheet.getSheets().map(function (sheet) {
        return sheet.getName();
      });

      if (isDefined_('MatchingSheetModule')) {
        const mappingValidation = MatchingSheetModule.validateMappingsExist();
        result.matchingSheet = {
          ok: mappingValidation.valid,
          mappingCount: MatchingSheetModule.getMappings().length,
          errors: mappingValidation.errors || [],
        };
      }

      if (isDefined_('RenewalTrackerModule')) {
        const renewalValidation = RenewalTrackerModule.validateSheetExists();
        const items = renewalValidation.valid ? RenewalTrackerModule.readAllItems() : [];
        const remindable = renewalValidation.valid ? RenewalTrackerModule.readRemindableItems() : [];
        result.renewalTrackerSheet = {
          ok: renewalValidation.valid,
          itemCount: items.length,
          remindableCount: remindable.length,
          error: renewalValidation.error || '',
        };
        result.renewalItems = items.map(function (item) {
          return {
            row: item.rowNumber,
            name: item.itemName,
            nextDueDate: RenewalTrackerModule.formatDate(item.nextDueDate),
            status: item.status,
            remindable: RenewalTrackerModule.isRemindableStatus(item.statusKey),
            hasEmail: item.reminderEmails.length > 0,
          };
        });
      }

      if (isDefined_('SpreadsheetUtils')) {
        result.emailSettingsSheet = {
          ok: !!SpreadsheetUtils.getSheetByName(spreadsheet, AppConfig.getEmailSettingsSheetName()),
          expectedName: AppConfig.getEmailSettingsSheetName(),
        };
        result.errorSheet = {
          ok: !!SpreadsheetUtils.getSheetByName(spreadsheet, AppConfig.getErrorSheetName()),
          expectedName: AppConfig.getErrorSheetName(),
        };
      }
    } catch (err) {
      result.error = String(err);
    }

    return result;
  }

  function logProjectFiles_(result) {
    Logger.log('=== Project files check (' + result.totalFiles + ' files) ===');
    Logger.log('PASS: ' + result.passCount + '  FAIL: ' + result.failCount);

    result.files.forEach(function (fileResult) {
      if (fileResult.ok) {
        Logger.log('  OK   ' + fileResult.file);
      } else {
        Logger.log('  FAIL ' + fileResult.file);
        fileResult.issues.forEach(function (issue) {
          Logger.log('       - ' + issue);
        });
      }
    });
  }

  function logModuleValidation_(result) {
    Logger.log('=== Module load check ===');
    Logger.log('Loaded: ' + result.found.length + ' modules');

    if (result.missing.length > 0) {
      Logger.log('MISSING modules:');
      result.missing.forEach(function (item) {
        Logger.log('  - ' + item.file + '  (' + item.module + ')');
      });
    }

    if (result.sheetColumns) {
      Logger.log('SheetColumns: ' + result.sheetColumns.message);
      result.sheetColumns.checks.forEach(function (check) {
        Logger.log('  ' + (check.ok ? 'OK' : 'FAIL') + '  ' + check.message);
      });
    }
  }

  function logSpreadsheetConfig_(result) {
    Logger.log('=== Spreadsheet check ===');
    Logger.log('Spreadsheet ID: ' + (result.spreadsheetId || '(empty — uses active spreadsheet)'));

    if (!result.spreadsheetAvailable) {
      Logger.log('ERROR: ' + (result.error || 'Spreadsheet not available'));
      return;
    }

    Logger.log('Spreadsheet: ' + result.spreadsheetName);
    Logger.log('Tabs (' + result.tabs.length + '): ' + result.tabs.join(', '));

    if (result.matchingSheet) {
      Logger.log(
        'Matching Sheet: ' + (result.matchingSheet.ok ? 'OK' : 'ISSUES') +
          ' (' + result.matchingSheet.mappingCount + ' mappings)'
      );
      result.matchingSheet.errors.forEach(function (validationError) {
        Logger.log('  - ' + validationError);
      });
    }

    if (result.emailSettingsSheet) {
      Logger.log(
        'Email Settings: ' + (result.emailSettingsSheet.ok ? 'OK' : 'NOT FOUND') +
          ' [' + result.emailSettingsSheet.expectedName + ']'
      );
    }

    if (result.renewalTrackerSheet) {
      Logger.log(
        'Renewal Tracker: ' + (result.renewalTrackerSheet.ok ? 'OK' : 'NOT FOUND') +
          ' — ' + result.renewalTrackerSheet.itemCount + ' item(s), ' +
          result.renewalTrackerSheet.remindableCount + ' remindable'
      );
      if (result.renewalTrackerSheet.error) {
        Logger.log('  - ' + result.renewalTrackerSheet.error);
      }
    }

    if (result.renewalItems && result.renewalItems.length > 0) {
      Logger.log('Renewal rows:');
      result.renewalItems.forEach(function (item) {
        Logger.log(
          '  Row ' + item.row + ': ' + item.name +
            ' | due ' + (item.nextDueDate || '(empty)') +
            ' | ' + item.status +
            (item.remindable && item.hasEmail ? '' : ' | NOT READY FOR REMINDERS')
        );
      });
    }

    if (result.errorSheet) {
      Logger.log(
        'Error sheet: ' + (result.errorSheet.ok ? 'OK' : 'NOT FOUND') +
          ' [' + result.errorSheet.expectedName + ']'
      );
    }
  }

  function logTriggers_(result) {
    Logger.log('=== Trigger check ===');
    if (result.error) {
      Logger.log('ERROR: ' + result.error);
      return;
    }

    if (result.installed.length === 0) {
      Logger.log('No triggers — run setupApp()');
    } else {
      result.triggers.forEach(function (trigger) {
        Logger.log('  OK  ' + trigger.handler + ' (' + trigger.eventType + ')');
      });
    }

    if (result.missing.length > 0) {
      Logger.log('MISSING triggers — run setupApp():');
      result.missing.forEach(function (handlerName) {
        Logger.log('  - ' + handlerName);
      });
    }
  }

  function logSummary_(results) {
    const allPass =
      results.projectFiles.allOk &&
      results.modules.allLoaded &&
      results.triggers.ok &&
      results.spreadsheet.spreadsheetAvailable;

    Logger.log('========================================');
    Logger.log(allPass ? 'OVERALL: PASS' : 'OVERALL: FAIL — fix items above');
    Logger.log('========================================');
  }

  function runFullDiagnostics_() {
    const projectFiles = validateProjectFiles_();
    logProjectFiles_(projectFiles);

    const modules = validateScriptModules_();
    logModuleValidation_(modules);

    const spreadsheet = validateSpreadsheetConfig_();
    logSpreadsheetConfig_(spreadsheet);

    const triggers = validateTriggers_();
    logTriggers_(triggers);

    const results = {
      projectFiles: projectFiles,
      modules: modules,
      spreadsheet: spreadsheet,
      triggers: triggers,
    };

    logSummary_(results);
    return results;
  }

  return {
    validateProjectFiles: validateProjectFiles_,
    validateSheetColumnsStructure: validateSheetColumnsStructure_,
    validateScriptModules: validateScriptModules_,
    validateSpreadsheetConfig: validateSpreadsheetConfig_,
    validateTriggers: validateTriggers_,
    runFullDiagnostics: runFullDiagnostics_,
    logProjectFiles: logProjectFiles_,
    logModuleValidation: logModuleValidation_,
    logSpreadsheetConfig: logSpreadsheetConfig_,
    logTriggers: logTriggers_,
    PROJECT_FILES: PROJECT_FILES,
  };
})();

/**
 * Checks all project files, module exports, and entry-point functions.
 */
function debugValidateProjectFiles() {
  const result = ScriptDiagnosticsModule.validateProjectFiles();
  ScriptDiagnosticsModule.logProjectFiles(result);
  return result;
}

/**
 * Checks SheetColumns MAIN/SUB/RENEWAL etc. structure.
 */
function debugValidateSheetColumns() {
  const result = ScriptDiagnosticsModule.validateSheetColumnsStructure();
  Logger.log('SheetColumns: ' + result.message);
  result.checks.forEach(function (check) {
    Logger.log((check.ok ? 'OK' : 'FAIL') + '  ' + check.message);
  });
  return result;
}

/**
 * Lists missing modules (e.g. SheetColumns is not defined).
 */
function debugValidateScriptModules() {
  const result = ScriptDiagnosticsModule.validateScriptModules();
  ScriptDiagnosticsModule.logModuleValidation(result);
  return result;
}

/**
 * Lists spreadsheet tabs and key sheets.
 */
function debugValidateSpreadsheetConfig() {
  const result = ScriptDiagnosticsModule.validateSpreadsheetConfig();
  ScriptDiagnosticsModule.logSpreadsheetConfig(result);
  return result;
}

/**
 * Lists installed triggers vs expected handlers.
 */
function debugListTriggers() {
  const result = ScriptDiagnosticsModule.validateTriggers();
  ScriptDiagnosticsModule.logTriggers(result);
  return result;
}

/**
 * Full diagnostic — run this first. Checks files, code, sheets, triggers, renewal rows.
 */
function debugRunFullDiagnostics() {
  return ScriptDiagnosticsModule.runFullDiagnostics();
}
