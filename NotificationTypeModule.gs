/**
 * Parses notification types (statuses and priorities) from Email Settings.
 */
const NotificationTypeModule = (function () {
  'use strict';

  const PRIORITY_TYPES = ['High', 'Medium', 'Low'];

  function normalizeToken_(value) {
    return SpreadsheetUtils.normalizeText(value);
  }

  function findKnownMatch_(text, knownList) {
    const normalized = normalizeToken_(text);
    if (!normalized) {
      return null;
    }
    for (var i = 0; i < knownList.length; i++) {
      if (normalizeToken_(knownList[i]) === normalized) {
        return knownList[i];
      }
    }
    return null;
  }

  function getAllKnownTypes_() {
    return TaskStatusModule.ALL_STATUSES.concat(PRIORITY_TYPES);
  }

  /**
   * Parses comma-separated notification types, merging fragments for multi-word statuses.
   * e.g. "Under,Review & Testing" → "Under Review & Testing"
   */
  function parseNotificationTypes_(rawValue) {
    if (!rawValue || String(rawValue).trim() === '') {
      return [];
    }

    const knownTypes = getAllKnownTypes_();
    const parts = String(rawValue)
      .split(',')
      .map(function (part) {
        return String(part).trim();
      })
      .filter(function (part) {
        return part !== '';
      });

    const parsed = [];
    var index = 0;

    while (index < parts.length) {
      var candidate = parts[index];
      var match = findKnownMatch_(candidate, knownTypes);

      while (!match && index + 1 < parts.length) {
        index++;
        candidate = candidate + ' ' + parts[index];
        match = findKnownMatch_(candidate, knownTypes);
      }

      if (match) {
        parsed.push(match);
      } else if (candidate) {
        parsed.push(candidate);
      }

      index++;
    }

    return parsed;
  }

  function isPriorityType_(value) {
    return !!findKnownMatch_(value, PRIORITY_TYPES);
  }

  function isStatusType_(value) {
    return !!findKnownMatch_(value, TaskStatusModule.ALL_STATUSES);
  }

  function taskMatchesNotificationTypes_(taskInfo, notificationTypes) {
    if (!taskInfo || !notificationTypes || notificationTypes.length === 0) {
      return false;
    }

    for (var i = 0; i < notificationTypes.length; i++) {
      const type = notificationTypes[i];
      if (isPriorityType_(type) && normalizeToken_(taskInfo.priority) === normalizeToken_(type)) {
        return true;
      }
      if (isStatusType_(type) && TaskStatusModule.statusesMatch(taskInfo.status, type)) {
        return true;
      }
      if (
        !isPriorityType_(type) &&
        !isStatusType_(type) &&
        (TaskStatusModule.statusesMatch(taskInfo.status, type) ||
          normalizeToken_(taskInfo.priority) === normalizeToken_(type))
      ) {
        return true;
      }
    }

    return false;
  }

  return {
    PRIORITY_TYPES: PRIORITY_TYPES,
    parseNotificationTypes: parseNotificationTypes_,
    taskMatchesNotificationTypes: taskMatchesNotificationTypes_,
    isPriorityType: isPriorityType_,
    isStatusType: isStatusType_,
  };
})();
