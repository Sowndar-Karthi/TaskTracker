/**
 * Admin entry points for one-time or repeat project setup.
 */

/**
 * Touch services that need OAuth so Google shows the consent screen
 * when scopes are missing or were revoked. Safe to re-run anytime.
 *
 * @returns {{ok: boolean, checks: Object, error: string}}
 */
function authorizeApp() {
  const checks = {
    spreadsheet: false,
    gmail: false,
    script: false,
  };

  try {
    const ss = AppConfig.getSpreadsheet();
    if (!ss) {
      throw new Error(
        'Spreadsheet unavailable — check AppConfig spreadsheet ID / open from the bound sheet'
      );
    }
    ss.getName();
    checks.spreadsheet = true;

    // Triggers scope (ScriptApp)
    ScriptApp.getProjectTriggers();
    checks.script = true;

    // Gmail send + search (IRCTC import, digests, train reminders)
    GmailApp.getInboxThreads(0, 1);
    checks.gmail = true;

    Logger.log('authorizeApp: all required permissions OK ' + JSON.stringify(checks));
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Permissions OK (Sheets + Gmail + Triggers)',
        'Task Tracker setup',
        6
      );
    } catch (uiErr) {
      // toast may be unavailable
    }
    return { ok: true, checks: checks, error: '' };
  } catch (err) {
    const message = String(err && err.message ? err.message : err);
    console.error('authorizeApp: permission check failed', {
      checks: checks,
      error: err,
      stack: err && err.stack,
    });
    Logger.log(
      'authorizeApp FAILED — re-run this function and click "Review permissions" / Allow. Detail: ' +
        message
    );
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Permission needed — run authorizeApp again and Allow access',
        'Task Tracker setup',
        10
      );
    } catch (uiErr) {
      // ignore
    }
    return { ok: false, checks: checks, error: message };
  }
}

/**
 * Run once by a project admin to install all installable triggers.
 * Safe to re-run — removes existing project triggers before recreating them.
 * If Gmail/Sheets permissions are missing, Google’s consent dialog appears; Allow, then run again.
 */
function setupApp() {
  const environment = AppConfig.isTesting() ? 'TEST' : 'PRODUCTION';
  Logger.log('setupApp running in: ' + environment);

  const auth = authorizeApp();
  if (!auth.ok) {
    Logger.log(
      'setupApp stopped — grant permissions via authorizeApp / Review permissions, then re-run setupApp'
    );
    return {
      success: false,
      error: 'Missing permissions — run authorizeApp (or setupApp again) and Allow access',
      auth: auth,
    };
  }

  const validation = MatchingSheetModule.validateMappingsExist();
  if (!validation.valid) {
    Logger.log('WARNING: Matching Sheet validation failed: ' + validation.errors.join('; '));
    validation.errors.forEach(function (validationError) {
      ErrorLogModule.warning('Setup', 'setupApp', validationError, { triggerSource: 'manual' });
    });
  } else {
    Logger.log('Matching Sheet OK — ' + MatchingSheetModule.getMappings().length + ' team mapping(s) loaded');
  }

  CompletionQueueModule.ensureQueueSheet();
  ErrorLogModule.ensureErrorSheet();

  const renewalValidation = RenewalTrackerModule.validateSheetExists();
  if (!renewalValidation.valid) {
    Logger.log('WARNING: Renewal Tracker — ' + renewalValidation.error);
    ErrorLogModule.warning('Setup', 'setupApp', renewalValidation.error, { triggerSource: 'manual' });
  } else {
    Logger.log('Renewal Tracker OK — ' + RenewalTrackerModule.readAllItems().length + ' item(s) loaded');
  }

  const bookingSheet = TrainBookingHistoryModule.getBookingSheet();
  if (!bookingSheet) {
    Logger.log('WARNING: Could not create/open Train Booking History');
    ErrorLogModule.warning('Setup', 'setupApp', 'Train Booking History unavailable', {
      triggerSource: 'manual',
    });
  } else {
    Logger.log('Train Booking History OK — headers ensured');
  }

  const completedSheet = TrainJourneyArchiveModule.getCompletedSheet();
  if (completedSheet && bookingSheet) {
    Logger.log('Train Completed Journeys OK');
  }

  const trainRouteValidation = TrainRouteModule.validateSheetExists();
  if (!trainRouteValidation.valid) {
    Logger.log('WARNING: Train Route — ' + trainRouteValidation.error);
    ErrorLogModule.warning('Setup', 'setupApp', trainRouteValidation.error, { triggerSource: 'manual' });
  } else {
    Logger.log(
      'Train Route OK — ' + TrainRouteModule.readActiveRoutes().length + ' active route(s)'
    );
  }

  return TriggerManager.initializeTriggers();
}

/**
 * Run to install triggers against the test spreadsheet (set IS_TESTING = true in AppConfig first).
 */
function setupTestTriggers() {
  if (!AppConfig.isTesting()) {
    throw new Error('Set IS_TESTING = true in AppConfig.gs before running setupTestTriggers()');
  }
  return setupApp();
}

/**
 * Removes all installable triggers for this script project.
 */
function teardownApp() {
  TriggerManager.removeAllTriggers();
  Logger.log('teardownApp: all project triggers removed');
}
