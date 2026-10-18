/**
 * Task status values and completion-status detection.
 */
const TaskStatusModule = (function () {
  'use strict';

  const ALL_STATUSES = [
    'Estimate Pending',
    'Pipeline / To Do',
    'In Progress',
    'Internal QA / Review',
    'Client UAT',
    'Waiting for Client Approval',
    'Deployed – Awaiting Verification',
    'Completed',
    'Requirements Not Clear',
    'Waiting for Client Input',
    'On Hold / Blocked',
    'Changes Requested / Rework',
    'Re-opened',
    'Closed / Cancelled',
    "Duplicate / Won't Fix",
    'Closed – Billable',
  ];

  const COMPLETION_STATUSES = [
    'Completed',
    'Closed / Cancelled',
    "Duplicate / Won't Fix",
    'Closed – Billable',
  ];

  function normalizeStatus_(value) {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value)
      .trim()
      .replace(/\u2013|\u2014/g, '-')
      .replace(/\u2019/g, "'")
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
