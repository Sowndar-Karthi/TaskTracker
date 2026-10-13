/**
 * Prunes old dedupe keys from Script Properties (notifications, train, renewals).
 * Safe to run manually — never deletes IRCTC_GMAIL_QUERY or unknown custom keys.
 */
const ScriptPropertiesMaintenanceModule = (function () {
  'use strict';

  const PROTECTED_KEYS = ['IRCTC_GMAIL_QUERY'];
  const KNOWN_PREFIXES = [
    'train_rem_',
    'notif_sent_',
    'notif_immediate_',
    'notif_immediate_parent_',
    'renewal_reminder_',
  ];

  function startOfDay_(date) {
    const d = new Date(date.getTime());
    d.setHours(0, 0, 0, 0);
    return d;
  }

  function parseDateFromKey_(key) {
    if (!key) {
      return null;
    }
    var match = String(key).match(/(\d{4}-\d{2}-\d{2})/);
    if (match) {
      const parsed = new Date(match[1] + 'T00:00:00');
      return isNaN(parsed.getTime()) ? null : startOfDay_(parsed);
    }
    match = String(key).match(/_(\d{8})_/);
    if (match) {
      const y = match[1].substring(0, 4);
      const m = match[1].substring(4, 6);
      const d = match[1].substring(6, 8);
      const parsed = new Date(y + '-' + m + '-' + d + 'T00:00:00');
      return isNaN(parsed.getTime()) ? null : startOfDay_(parsed);
    }
    return null;
  }

  function isKnownDedupeKey_(key) {
    if (!key || PROTECTED_KEYS.indexOf(key) >= 0) {
      return false;
    }
    for (var i = 0; i < KNOWN_PREFIXES.length; i++) {
      if (String(key).indexOf(KNOWN_PREFIXES[i]) === 0) {
        return true;
      }
    }
    return false;
  }

  /**
   * Remove dedupe keys older than retentionDays (default 90).
   * @param {number} retentionDays
   * @returns {{removed: number, kept: number, errors: number}}
   */
  function cleanupOldDedupeKeys_(retentionDays) {
    const days = typeof retentionDays === 'number' && retentionDays > 0 ? retentionDays : 90;
    const cutoff = startOfDay_(new Date());
    cutoff.setDate(cutoff.getDate() - days);

    const props = PropertiesService.getScriptProperties();
    const allKeys = props.getKeys();
    var removed = 0;
    var kept = 0;
    var errors = 0;

    allKeys.forEach(function (key) {
      if (!isKnownDedupeKey_(key)) {
        return;
      }

      try {
        var shouldRemove = false;

        if (key.indexOf('notif_immediate_') === 0) {
          const raw = props.getProperty(key);
          const ts = parseInt(raw, 10);
          if (!isNaN(ts)) {
            shouldRemove = ts < cutoff.getTime();
          }
        } else {
          const keyDate = parseDateFromKey_(key);
          if (keyDate) {
            shouldRemove = keyDate.getTime() < cutoff.getTime();
          }
        }

        if (shouldRemove) {
          props.deleteProperty(key);
          removed++;
        } else {
          kept++;
        }
      } catch (err) {
        errors++;
        console.error('ScriptPropertiesMaintenanceModule.cleanupOldDedupeKeys_: key failed', {
          key: key,
          error: err,
        });
      }
    });

    const summary = { removed: removed, kept: kept, errors: errors, retentionDays: days };
    console.log('ScriptPropertiesMaintenanceModule.cleanupOldDedupeKeys_', summary);
    return summary;
  }

  return {
    cleanupOldDedupeKeys: cleanupOldDedupeKeys_,
  };
})();

/**
 * Manual: prune notification/train/renewal dedupe keys older than 90 days.
 * @param {number} retentionDays Optional override (default 90)
 */
function runCleanupScriptProperties(retentionDays) {
  return ScriptPropertiesMaintenanceModule.cleanupOldDedupeKeys(retentionDays);
}
