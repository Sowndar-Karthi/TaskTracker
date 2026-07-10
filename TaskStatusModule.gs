/**
 * Task status values and completion-status detection.
 */
const TaskStatusModule = (function () {
  'use strict';

  const ALL_STATUSES = [
    'Pipeline / To Do',
    'Requirements Not Clear',
    'In Progress',
    'On Hold / Blocked',
    'Waiting for Client Input',
    'Under Review & Testing',
    'Changes Requested / Rework',
    'Ready for Execution',
    'Completed',
    'Closed / Cancelled',
    'Closed – Billable',
    'Closed – Not Proceeding',
  ];

  const COMPLETION_STATUSES = [
    'Completed',
    'Closed / Cancelled',
    'Closed – Billable',
    'Closed – Not Proceeding',
  ];

  function normalizeStatus_(value) {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value)
      .trim()
      .replace(/\u2013|\u2014/g, '-')
      .replace(/\s+/g, ' ')
      .toLowerCase();
  }

  function isCompletionStatus_(statusValue) {
    const normalized = normalizeStatus_(statusValue);
    if (!normalized) {
      return false;
    }
    for (var i = 0; i < COMPLETION_STATUSES.length; i++) {
      if (normalizeStatus_(COMPLETION_STATUSES[i]) === normalized) {
        return true;
      }
    }
    return false;
  }

  function statusesMatch_(valueA, valueB) {
    return normalizeStatus_(valueA) === normalizeStatus_(valueB);
  }

  return {
    ALL_STATUSES: ALL_STATUSES,
    COMPLETION_STATUSES: COMPLETION_STATUSES,
    normalizeStatus: normalizeStatus_,
    isCompletionStatus: isCompletionStatus_,
    statusesMatch: statusesMatch_,
  };
})();
