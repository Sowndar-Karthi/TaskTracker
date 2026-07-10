/**
 * Reads email alert configuration from the Email Settings sheet.
 */
const EmailSettingsModule = (function () {
  'use strict';

  function parseEmailList_(rawValue) {
    if (!rawValue || String(rawValue).trim() === '') {
      return [];
    }
    return String(rawValue)
      .split(',')
      .map(function (email) {
        return String(email).trim();
      })
      .filter(function (email) {
        return email !== '' && email.indexOf('@') > 0;
      });
  }

  function parseAlertTimes_(rawValue) {
    if (!rawValue || String(rawValue).trim() === '') {
      return [];
    }
    return String(rawValue)
      .split(',')
      .map(function (timeValue) {
        return String(timeValue).trim();
      })
      .filter(function (timeValue) {
        return timeValue !== '';
      });
  }

  function normalizeConfigRow_(row, rowIndex) {
    const notificationTypes = NotificationTypeModule.parseNotificationTypes(
      row[SheetColumns.EMAIL.NOTIFICATION_TYPE - 1]
    );

    return {
      rowNumber: rowIndex,
      toEmails: parseEmailList_(row[SheetColumns.EMAIL.TO - 1]),
      ccEmails: parseEmailList_(row[SheetColumns.EMAIL.CC - 1]),
      bccEmails: parseEmailList_(row[SheetColumns.EMAIL.BCC - 1]),
      dataSheetName: String(row[SheetColumns.EMAIL.DATA_SHEET_NAME - 1] || '').trim(),
      timezone: String(row[SheetColumns.EMAIL.TIMEZONE - 1] || 'IST').trim(),
      alertTimes: parseAlertTimes_(row[SheetColumns.EMAIL.ALERT_TIME - 1]),
      notificationTypes: notificationTypes,
      notificationTypesRaw: String(row[SheetColumns.EMAIL.NOTIFICATION_TYPE - 1] || '').trim(),
    };
  }

  function isValidConfig_(config) {
    return !!(
      config &&
      config.toEmails.length > 0 &&
      config.dataSheetName &&
      config.alertTimes.length > 0 &&
      config.notificationTypes.length > 0
    );
  }

  function readAllSettings_() {
    const spreadsheet = AppConfig.getSpreadsheet();
    if (!spreadsheet) {
      ErrorLogModule.error('EmailSettingsModule', 'readAllSettings_', 'Spreadsheet unavailable', {});
      return [];
    }

    const sheet = SpreadsheetUtils.getSheetByName(spreadsheet, AppConfig.getEmailSettingsSheetName());
    if (!sheet) {
      ErrorLogModule.error('EmailSettingsModule', 'readAllSettings_', 'Email settings sheet not found', {});
      return [];
    }

    const lastRow = SpreadsheetUtils.getLastDataRow(sheet, SheetColumns.EMAIL.DATA_SHEET_NAME);
    if (lastRow < AppConfig.getDataStartRow()) {
      return [];
    }

    try {
      const rowCount = lastRow - AppConfig.getHeaderRow();
      const values = sheet
        .getRange(AppConfig.getDataStartRow(), 1, rowCount, SheetColumns.EMAIL.LAST_COLUMN)
        .getValues();

      return values
        .map(function (row, index) {
          return normalizeConfigRow_(row, AppConfig.getDataStartRow() + index);
        })
        .filter(isValidConfig_);
    } catch (err) {
      ErrorLogModule.error('EmailSettingsModule', 'readAllSettings_', 'Read failed', { error: err });
      return [];
    }
  }

  function getSettingsForDataSheet_(dataSheetName) {
    if (!dataSheetName) {
      return [];
    }
    const normalized = SpreadsheetUtils.normalizeText(dataSheetName);
    return readAllSettings_().filter(function (config) {
      return SpreadsheetUtils.normalizeText(config.dataSheetName) === normalized;
    });
  }

  return {
    readAllSettings: readAllSettings_,
    getSettingsForDataSheet: getSettingsForDataSheet_,
    parseEmailList: parseEmailList_,
    parseAlertTimes: parseAlertTimes_,
  };
})();
