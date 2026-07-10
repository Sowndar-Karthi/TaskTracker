/**
 * Writes automation errors and warnings to the Error sheet for review and fixes.
 */
const ErrorLogModule = (function () {
  'use strict';

  const SEVERITY = {
    ERROR: 'ERROR',
    WARNING: 'WARNING',
    INFO: 'INFO',
  };

  const HEADERS = [
    'Logged At',
    'Severity',
    'Module',
    'Function',
    'Sheet Name',
    'Row #',
    'Task / Key',
    'Message',
    'Details',
    'Trigger Source',
    'Resolved',
    'Fix Notes',
  ];

  var currentTriggerSource_ = 'system';
  var MAX_DETAILS_LENGTH = 45000;

  function setTriggerSource_(source) {
    currentTriggerSource_ = source || 'system';
  }

  function clearTriggerSource_() {
    currentTriggerSource_ = 'system';
  }

  function getTriggerSource_() {
    return currentTriggerSource_;
  }

  function serializeDetails_(context) {
    if (!context) {
      return '';
    }

    try {
      const payload = {};
      Object.keys(context).forEach(function (key) {
        const value = context[key];
        if (value instanceof Error) {
          payload[key] = { message: value.message, stack: value.stack };
        } else {
          payload[key] = value;
        }
      });
      const json = JSON.stringify(payload);
      if (json.length > MAX_DETAILS_LENGTH) {
        return json.substring(0, MAX_DETAILS_LENGTH) + '…[truncated]';
      }
      return json;
    } catch (err) {
      return String(context);
    }
  }

  function extractTaskKey_(context) {
    if (!context) {
      return '';
    }
    if (context.mainTaskName) {
      return String(context.mainTaskName);
    }
    if (context.taskName) {
      return String(context.taskName);
    }
    if (context.rowSignature) {
      return String(context.rowSignature);
    }
    return '';
  }

  function ensureErrorSheet_(spreadsheet) {
    if (!spreadsheet) {
      return null;
    }

    const sheetName = AppConfig.getErrorSheetName();
    var sheet = SpreadsheetUtils.getSheetByName(spreadsheet, sheetName);

    if (!sheet) {
      try {
        sheet = spreadsheet.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      } catch (err) {
        console.error('ErrorLogModule.ensureErrorSheet_: could not create sheet', { error: err });
        return null;
      }
    } else if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    }

    return sheet;
  }

  function write_(entry) {
    try {
      const spreadsheet = AppConfig.getSpreadsheet();
      const sheet = ensureErrorSheet_(spreadsheet);
      if (!sheet) {
        return { success: false, error: 'Error sheet unavailable' };
      }

      const row = [
        entry.loggedAt || new Date(),
        entry.severity || SEVERITY.ERROR,
        entry.module || '',
        entry.functionName || '',
        entry.sheetName || '',
        entry.rowIndex || '',
        entry.taskKey || '',
        entry.message || '',
        entry.details || '',
        entry.triggerSource || getTriggerSource_(),
        '',
        '',
      ];

      sheet.appendRow(row);
      return { success: true };
    } catch (err) {
      console.error('ErrorLogModule.write_: failed to write error log', { error: err, entry: entry });
      return { success: false, error: String(err) };
    }
  }

  function log_(severity, module, functionName, message, context) {
    const safeContext = context || {};
    const logMessage = (module || 'Unknown') + '.' + (functionName || 'unknown') + ': ' + message;

    if (severity === SEVERITY.ERROR) {
      console.error(logMessage, safeContext);
    } else if (severity === SEVERITY.WARNING) {
      console.warn(logMessage, safeContext);
    } else {
      console.log(logMessage, safeContext);
    }

    return write_({
      loggedAt: new Date(),
      severity: severity,
      module: module || '',
      functionName: functionName || '',
      sheetName: safeContext.sheetName || safeContext.sheet || safeContext.dataSheetName || '',
      rowIndex: safeContext.rowIndex || safeContext.row || '',
      taskKey: extractTaskKey_(safeContext),
      message: message || '',
      details: serializeDetails_(safeContext),
      triggerSource: safeContext.triggerSource || getTriggerSource_(),
    });
  }

  function error_(module, functionName, message, context) {
    return log_(SEVERITY.ERROR, module, functionName, message, context);
  }

  function warning_(module, functionName, message, context) {
    return log_(SEVERITY.WARNING, module, functionName, message, context);
  }

  function info_(module, functionName, message, context) {
    return log_(SEVERITY.INFO, module, functionName, message, context);
  }

  return {
    SEVERITY: SEVERITY,
    HEADERS: HEADERS,
    setTriggerSource: setTriggerSource_,
    clearTriggerSource: clearTriggerSource_,
    getTriggerSource: getTriggerSource_,
    ensureErrorSheet: function () {
      return ensureErrorSheet_(AppConfig.getSpreadsheet());
    },
    log: log_,
    error: error_,
    warning: warning_,
    info: info_,
  };
})();

/**
 * Manual run: create Error sheet with headers (safe to re-run).
 */
function runEnsureErrorSheet() {
  ErrorLogModule.ensureErrorSheet();
  Logger.log('Error sheet ready: ' + AppConfig.getErrorSheetName());
}

/**
 * Manual run: write a test row to the Error sheet.
 */
function runTestErrorLog() {
  return ErrorLogModule.error('ErrorLogModule', 'runTestErrorLog', 'Test error log entry', {
    triggerSource: 'manual',
    note: 'If you see this row, error logging works.',
  });
}
