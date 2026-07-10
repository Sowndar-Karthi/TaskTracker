/**
 * Admin entry points for one-time or repeat project setup.
 */

/**
 * Run once by a project admin to install all installable triggers.
 * Safe to re-run — removes existing project triggers before recreating them.
 */
function setupApp() {
  const environment = AppConfig.isTesting() ? 'TEST' : 'PRODUCTION';
  Logger.log('setupApp running in: ' + environment);

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
