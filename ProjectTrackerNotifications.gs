/**
 * PROJECT TRACKER — EMAIL & NOTIFICATIONS
 * Hourly batched task updates, daily/weekly summaries, pending queue, Gmail threading.
 * Depends on CONFIG and helpers in ProjectTracker.gs.
 */
function diagnoseProjectRowEmailThread(row) {
  row = row || 2;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  if (!mainSheet) {
    Logger.log('Dev Tracker not found');
    return;
  }
  const threadCol = findProjectThreadIdColumn_(mainSheet);
  const threadId = getProjectThreadIdForRow_(mainSheet, row);
  const subjectR = (mainSheet.getRange(row, CONFIG.MAIN_SUBJECT_COL).getValue() || '').toString().trim();
  const taskName = (mainSheet.getRange(row, CONFIG.MAIN_TASK_NAME_COL).getValue() || '').toString().trim();

  Logger.log('=== DIAGNOSE ROW ' + row + ' ===');
  Logger.log('Thread ID column: ' + threadCol + ' (' + columnToLetter_(threadCol) + ')');
  Logger.log('Thread ID value: ' + (threadId || '(empty — script will send NEW email)'));
  Logger.log('Internal Mail Subject (R): ' + subjectR);
  Logger.log('Task (C): ' + taskName);

  if (!threadId) {
    Logger.log('WHY new thread: no Thread ID in column ' + columnToLetter_(threadCol));
    return;
  }

  try {
    const thread = GmailApp.getThreadById(threadId);
    if (!thread) {
      Logger.log('WHY new thread: Gmail cannot find this ID (trashed/invalid). Script will clear and resend.');
      return;
    }
    const firstSub = thread.getFirstMessageSubject();
    Logger.log('Gmail thread FOUND — first message subject: "' + firstSub + '"');
    Logger.log('Next send will REPLY in this thread. Inbox shows: Re: ' + firstSub);
    Logger.log('(Re: uses the FIRST email subject, not column R — that is normal Gmail behavior.)');
    if (normalizeProjectEmailSubject_(firstSub) !== normalizeProjectEmailSubject_(subjectR)) {
      Logger.log('NOTE: Column R subject differs from first Gmail subject — thread is still the same if ID is valid.');
    }
  } catch (e) {
    Logger.log('WHY new thread: getThreadById failed — ' + e);
  }
}

/**
 * Fix one task when inbox shows "Re: Project Update: …" but column R is TASK UPDATE.
 * Clears Thread ID, sets column R from task name, sends one new first email (new subject line).
 *
 * Apps Script Run button cannot pass arguments — use a wrapper below, e.g. resetProjectThreadRow10()
 * @param {number} row Sheet row on Dev Tracker (must be >= 2)
 */
function resetProjectThreadForRow(row) {
  const n = Number(row);
  if (!n || n < 2 || n !== Math.floor(n)) {
    Logger.log('Invalid row. Apps Script cannot pass row from the Run menu.');
    Logger.log('Use: resetProjectThreadRow10()  for Open data form Phase - 2 (sheet row 10)');
    Logger.log('Or: resetProjectThreadRow2()   for Demo Page (sheet row 2)');
    Logger.log('Or edit RESET_PROJECT_THREAD_ROW below and run resetProjectThreadForConfiguredRow()');
    return;
  }
  resetProjectThreadForRow_(n);
}

/** Change this row, then run resetProjectThreadForConfiguredRow from the editor */
const RESET_PROJECT_THREAD_ROW = 10;

function resetProjectThreadForConfiguredRow() {
  const rowNum = parseInt(RESET_PROJECT_THREAD_ROW, 10);
  if (!rowNum || rowNum < 2) {
    throw new Error(
      'RESET_PROJECT_THREAD_ROW must be an integer >= 2 (currently: ' + RESET_PROJECT_THREAD_ROW + ')'
    );
  }
  resetProjectThreadForRow_(rowNum);
}

/** Open data form Phase - 2 — Dev Tracker row 10 */
function resetProjectThreadRow10() {
  resetProjectThreadForRow_(10);
}

/** Demo Page — Dev Tracker row 2 */
function resetProjectThreadRow2() {
  resetProjectThreadForRow_(2);
}

function resetProjectThreadForRow_(row) {
  const rowNum = parseInt(row, 10);
  if (!rowNum || rowNum < 2) {
    throw new Error('resetProjectThreadForRow_: row must be an integer >= 2 (got "' + row + '")');
  }
  row = rowNum;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(getMainSheetName_());
  if (!mainSheet) {
    Logger.log('Dev Tracker not found');
    return;
  }

  const taskName = (mainSheet.getRange(row, CONFIG.MAIN_TASK_NAME_COL).getValue() || '').toString().trim();
  if (!taskName) {
    Logger.log('Row ' + row + ' has no task name in column C');
    return;
  }

  const newSubject = buildDefaultInternalMailSubject_(taskName);
  const oldId = getProjectThreadIdForRow_(mainSheet, row);

  setProjectThreadIdForRow_(mainSheet, row, '');
  mainSheet.getRange(row, CONFIG.MAIN_SUBJECT_COL).setValue(newSubject);

  Logger.log('Row ' + row + ' (' + taskName + '): cleared thread ' + (oldId || '(none)'));
  Logger.log('Column R set to: ' + newSubject);

  const status = (mainSheet.getRange(row, CONFIG.MAIN_STATUS_COL).getValue() || '').toString().trim();
  const plain = 'New Gmail thread started for: ' + taskName;
  const html = '<html><body style="font-family: Arial, sans-serif;"><p>New thread for <b>' +
    (typeof escapeHtml === 'function' ? escapeHtml(taskName) : taskName) +
    '</b>. Future updates will use subject: ' + newSubject + '</p></body></html>';

  sendProjectTrackerEmail_(newSubject, plain, html, mainSheet, row, status);

  Utilities.sleep(1500);
  const newId = getProjectThreadIdForRow_(mainSheet, row);
  Logger.log('Done. New Thread ID in column U: ' + (newId || '(check logs — may need a few seconds)'));
  Logger.log('Next emails should show: Re: TASK UPDATE: ' + taskName);
}

function columnToLetter_(column) {
  let temp = '';
  let col = column;
  while (col > 0) {
    const rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}

/**
 * GmailApp expects BCC as a comma-separated string.
 * @returns {string}
 */
function getBccRecipientsForGmail_() {
  const raw = getProjectBccEmails_();
  if (!raw && raw !== 0) return '';
  const parts = [];
  if (Array.isArray(raw)) {
    for (let i = 0; i < raw.length; i++) {
      const s = (raw[i] || '').toString().trim();
      if (s) parts.push(s);
    }
  } else {
    const str = raw.toString().trim();
    if (str) {
      str.split(',').forEach(function (piece) {
        const t = piece.trim();
        if (t) parts.push(t);
      });
    }
  }
  return parts.join(',');
}

/**
 * @param {Object} base Options for GmailApp.sendEmail / GmailMessage.reply
 * @returns {Object} Shallow copy with bcc set when BCC is configured in ProjectTrackerUserConfig.gs
 */
function mergeEmailOptionsWithBcc_(base) {
  const out = {};
  const b = base || {};
  for (const k in b) {
    if (Object.prototype.hasOwnProperty.call(b, k)) {
      out[k] = b[k];
    }
  }
  const bcc = getBccRecipientsForGmail_();
  if (bcc) {
    out.bcc = bcc;
  }
  return out;
}

/**
 * TRIGGER: Runs instantly when a cell is edited
 */
function handleMainUpdate(sheet, row) {
  try {
    const data = sheet.getRange(row, 1, 1, 21).getValues()[0];
    const taskName = data[CONFIG.MAIN_TASK_NAME_COL - 1];
    let subject = data[CONFIG.MAIN_SUBJECT_COL - 1];
    const status = (data[CONFIG.MAIN_STATUS_COL - 1] || '').toString().trim();

    if (isCompletedOrClosedStatus_(status)) {
      const purgeSubject = (subject || '').toString().trim() ||
        (taskName ? buildDefaultInternalMailSubject_(taskName) : '');
      if (purgeSubject) {
        purgePendingChangesForSubject_(purgeSubject);
      }
      Logger.log('Skipping pending log for completed/closed task at row ' + row + ' (status: ' + status + ')');
      return;
    }
    
    // BACKUP LOGIC: If R is empty, use C (must match TASK UPDATE format used in column R)
    if (!subject || subject.toString().trim() === "") {
      if (taskName && taskName.toString().trim() !== "") {
        subject = buildDefaultInternalMailSubject_(taskName);
        sheet.getRange(row, CONFIG.MAIN_SUBJECT_COL).setValue(subject);
        Logger.log('Empty subject in Row ' + row + '. Used Task Name as backup.');
      } else {
        Logger.log('Both Subject and Task Name are empty for row ' + row + '. Skipping.');
        return;
      }
    }

    // Log to pending queue — hourly batch will send the email
    logPendingChange(subject, row);
    Logger.log('Change logged for: ' + subject + ' (row ' + row + ')');
    
  } catch (error) {
    Logger.log('handleMainUpdate Error: ' + error.toString());
  }
}

/**
 * Sort rank for Related Sub-Tasks in email: workflow statuses first (CONFIG.STATUS_LABELS
 * positions 0–7), then unknown/other, then Completed/Closed family last.
 * Within the same status, larger sheet row = nearer top (reverse sheet order).
 */
function getSubtaskEmailStatusRank_(statusRaw) {
  const s = (statusRaw || '').toString().trim();
  if (!s || s === 'N/A') {
    return 4000;
  }

  const workflowOrder = CONFIG.STATUS_LABELS.slice(0, 8);
  const exact = workflowOrder.indexOf(s);
  if (exact !== -1) {
    return exact;
  }

  const sl = s.toLowerCase();
  if (sl.indexOf('changes requested') !== -1 && sl.indexOf('rework') !== -1) {
    return workflowOrder.indexOf('Changes Requested / Rework');
  }

  const completedFamily = CONFIG.STATUS_LABELS.slice(8);
  for (let k = 0; k < completedFamily.length; k++) {
    if (completedFamily[k] === s) {
      return 1000 + k;
    }
  }
  return 500;
}

/**
 * Same ordering as Related Sub-Tasks: workflow status rank, then higher sheet row first.
 */
function sortSubTaskRowsForSummaryEmail_(entries) {
  entries.sort(function (a, b) {
    const ra = getSubtaskEmailStatusRank_(a.data[CONFIG.SUB_STATUS_COL - 1]);
    const rb = getSubtaskEmailStatusRank_(b.data[CONFIG.SUB_STATUS_COL - 1]);
    if (ra !== rb) {
      return ra - rb;
    }
    return b.sheetRow - a.sheetRow;
  });
  return entries;
}

/**
 * NEW FUNCTION: Get all sub-tasks related to a main task
 */
function getRelatedSubTasks(mainTaskName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const subSheet = ss.getSheetByName(getSubSheetName_());
    
    if (!subSheet) {
      return '<p style="color: #7f8c8d; font-style: italic;">No sub-task sheet found.</p>';
    }
    
    const subData = subSheet.getDataRange().getValues();
    const relatedTasks = [];
    
    // Find all sub-tasks that link to this main task
    for (let i = 1; i < subData.length; i++) {
      const linkedTaskName = subData[i][CONFIG.SUB_LINK_COL - 1]; // Column B
      
      if (linkedTaskName === mainTaskName) {
        relatedTasks.push({
          _sheetRow: i + 1,
          task: subData[i][CONFIG.SUB_TASK_NAME_COL - 1] || 'N/A',
          assigned: subData[i][CONFIG.SUB_ASSIGNED_COL - 1] || 'N/A',
          status: subData[i][CONFIG.SUB_STATUS_COL - 1] || 'N/A',
          env: subData[i][CONFIG.SUB_ENV_COL - 1] || 'N/A',
          blocker: subData[i][CONFIG.SUB_BLOCKER_COL - 1] || 'None',
          dependency: subData[i][CONFIG.SUB_DEPENDENCY_COL - 1] || 'None'
        });
      }
    }

    relatedTasks.sort(function (a, b) {
      const ra = getSubtaskEmailStatusRank_(a.status);
      const rb = getSubtaskEmailStatusRank_(b.status);
      if (ra !== rb) {
        return ra - rb;
      }
      return b._sheetRow - a._sheetRow;
    });
    
    if (relatedTasks.length === 0) {
      return `<div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-top: 20px;">
  <h3 style="color: #34495e; margin-top: 0;">Related Sub-Tasks</h3>
  <p style="color: #7f8c8d; font-style: italic;">No sub-tasks found for this main task.</p>
</div>`;
    }
    
    // Build HTML table
    let table = `<div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-top: 20px;">
  <h3 style="color: #34495e; margin-top: 0;">Related Sub-Tasks (${relatedTasks.length})</h3>
  <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; background-color: white;">
    <thead>
      <tr style="background-color: #2ecc71; color: white;">
        <th style="text-align: left; padding: 10px;">Task</th>
        <th style="text-align: left; padding: 10px;">Assigned</th>
        <th style="text-align: left; padding: 10px;">Status</th>
        <th style="text-align: left; padding: 10px;">Environment</th>
        <th style="text-align: left; padding: 10px;">Support Needed / Blocker</th>
        <th style="text-align: left; padding: 10px;">Dependency Detail</th>
      </tr>
    </thead>
    <tbody>`;
    
    for (let i = 0; i < relatedTasks.length; i++) {
      const task = relatedTasks[i];
      const rowColor = i % 2 === 0 ? '#f8f9fa' : 'white';
      const statusColor = getStatusColor(task.status, task.env);
      
      table += `<tr style="background-color: ${rowColor};">
        <td style="padding: 8px;">${escapeHtml(task.task)}</td>
        <td style="padding: 8px;">${escapeHtml(task.assigned)}</td>
        <td style="padding: 8px; background-color: ${statusColor}; font-weight: bold;">${escapeHtml(task.status)}</td>
        <td style="padding: 8px;">${escapeHtml(task.env)}</td>
        <td style="padding: 8px; color: ${task.blocker !== 'None' ? '#e74c3c' : '#7f8c8d'};">
          ${escapeHtml(task.blocker)}
        </td>
        <td style="padding: 8px; color: ${task.dependency !== 'None' ? '#34495e' : '#7f8c8d'};">
          ${escapeHtml(task.dependency)}
        </td>
      </tr>`;
    }
    
    table += `</tbody>
  </table>
</div>`;
    
    return table;
    
  } catch (error) {
    Logger.log('getRelatedSubTasks Error: ' + error.toString());
    return '<p style="color: #e74c3c;">Error loading sub-tasks.</p>';
  }
}

/**
 * LOGIC: Process updates from the Sub Sheet (Links back to Main)
 */
function handleSubUpdate(subSheet, row) {
  try {
    const lastCol = Math.max(subSheet.getLastColumn(), CONFIG.SUB_GOOGLE_TASK_ID_COL);
    const subData = subSheet.getRange(row, 1, 1, lastCol).getValues()[0];
    const linkedTaskName = subData[CONFIG.SUB_LINK_COL - 1];  // Column B: Main Task Name
    
    if (!linkedTaskName || linkedTaskName.toString().trim() === "") {
      Logger.log('No linked task name found for sub-task row ' + row);
      return;
    }
    
    const mainSheet = getMainTrackerSheet_();
    if (!mainSheet) {
      Logger.log('Main sheet not found: ' + getMainSheetName_());
      return;
    }
    
    const mainData = mainSheet.getDataRange().getValues();
    
    let mainRowIndex = -1;
    let subject = "";

    // Find the parent row in Main Sheet to get the correct Subject/Thread
    for (let i = 1; i < mainData.length; i++) {
      if (mainData[i][CONFIG.MAIN_TASK_NAME_COL - 1] === linkedTaskName) {
        mainRowIndex = i + 1;
        subject = mainData[i][CONFIG.MAIN_SUBJECT_COL - 1];
        
        // BACKUP: If subject is empty, use task name
        if (!subject || subject.toString().trim() === "") {
          subject = buildDefaultInternalMailSubject_(linkedTaskName);
        }
        break;
      }
    }

    if (mainRowIndex === -1) {
      Logger.log('Could not find matching main task for: ' + linkedTaskName);
      return;
    }
    
    // Log to pending queue — hourly batch will send the email
    logPendingChange(subject, mainRowIndex);
    Logger.log('Sub-task change logged for: ' + subject + ' (main row ' + mainRowIndex + ')');
    
  } catch (error) {
    Logger.log('handleSubUpdate Error: ' + error.toString());
  }
}

/**
 * Default Internal Mail Subject (column R) — keep one format so Gmail stays on one thread.
 * @param {string} taskName
 * @returns {string}
 */
function buildDefaultInternalMailSubject_(taskName) {
  const name = (taskName || '').toString().trim();
  if (!name) {
    return '';
  }
  return 'TASK UPDATE: ' + name;
}

/**
 * Normalize subjects for matching (Re:, [Dev Tracker], etc.)
 * @param {string} subject
 * @returns {string}
 */
function normalizeProjectEmailSubject_(subject) {
  let s = (subject || '').toString().trim();
  s = s.replace(/^(\s*(re|fwd?|fw)\s*:\s*)+/gi, '');
  s = s.replace(/^\[dev tracker\]\s*/gi, '');
  return s.trim().toLowerCase();
}

/**
 * Avoid project mail using the same subject as Leave Tracker (Gmail would merge threads).
 * @param {string} subject
 * @returns {string}
 */
function ensureProjectEmailSubject_(subject) {
  const s = (subject || '').toString().trim();
  const leaveSubject = 'Team Leave Tracker Report';
  if (!s) {
    return s;
  }
  if (s === leaveSubject || s.indexOf(leaveSubject) === 0) {
    Logger.log('Subject matched Leave Tracker — prefixed with [Dev Tracker]');
    return '[Dev Tracker] ' + s;
  }
  return s;
}

/**
 * Backoff schedule (ms) for finding our just-sent thread in Gmail. Gmail's
 * indexing of newly-sent mail can lag well past 3s; with the wider schedule
 * below the worst case is ~26s before we give up and warn.
 */
const PROJECT_THREAD_LOOKUP_DELAYS_MS = [6000, 8000, 12000];

/**
 * @param {number} mainRowIndex
 * @returns {string}
 */
function buildDevTrackerRowMarker_(mainRowIndex) {
  return '[Dev Tracker Row: ' + mainRowIndex + ']';
}

/**
 * @param {GoogleAppsScript.Gmail.GmailMessage} message
 * @param {number} mainRowIndex
 * @returns {boolean}
 */
function messageContainsDevTrackerRow_(message, mainRowIndex) {
  if (!mainRowIndex || !message) {
    return false;
  }
  const marker = buildDevTrackerRowMarker_(mainRowIndex).toLowerCase();
  const body = ((message.getPlainBody() || '') + ' ' + (message.getBody() || '')).toLowerCase();
  return body.indexOf(marker) >= 0;
}

/**
 * @param {GoogleAppsScript.Gmail.GmailThread} thread
 * @param {string} wantSub Normalized subject
 * @returns {boolean}
 */
function threadMatchesNormalizedSubject_(thread, wantSub) {
  const messages = thread.getMessages();
  const lastMessage = messages[messages.length - 1];
  const lastSub = normalizeProjectEmailSubject_(lastMessage.getSubject());
  const firstSub = normalizeProjectEmailSubject_(thread.getFirstMessageSubject());
  return lastSub === wantSub || firstSub === wantSub;
}

/**
 * Picks the best Gmail thread for a subject. Prefers row-marker match, then timestamp proximity / recency.
 * @param {GoogleAppsScript.Gmail.GmailThread[]} threads
 * @param {string} wantSub
 * @param {Object} [options]
 * @returns {GoogleAppsScript.Gmail.GmailThread|null}
 */
function pickBestGmailThreadForSubject_(threads, wantSub, options) {
  options = options || {};
  const mainRowIndex = options.mainRowIndex;
  const sentAtMs = options.sentAtMs;
  const matchWindowMs = options.matchWindowMs || 30000;
  const preferRecent = options.preferRecent === true;

  let bestThread = null;
  let bestScore = -1;

  for (let t = 0; t < threads.length; t++) {
    const thread = threads[t];
    if (!threadMatchesNormalizedSubject_(thread, wantSub)) {
      continue;
    }

    const messages = thread.getMessages();
    const lastMessage = messages[messages.length - 1];
    let score = 0;

    if (mainRowIndex && messageContainsDevTrackerRow_(lastMessage, mainRowIndex)) {
      score += 100000;
    }

    if (sentAtMs) {
      const delta = Math.abs(lastMessage.getDate().getTime() - sentAtMs);
      if (delta <= matchWindowMs) {
        score += Math.max(0, matchWindowMs - delta);
      } else if (!preferRecent) {
        continue;
      }
    }

    if (preferRecent) {
      score += lastMessage.getDate().getTime() / 1000;
    }

    if (score > bestScore) {
      bestScore = score;
      bestThread = thread;
    }
  }

  return bestThread;
}

/**
 * After GmailApp.sendEmail starts a new thread, find that thread and write
 * its ID into the row's Thread ID column.
 *
 * Robust version of the previous inline "sleep 3s → single search → silently
 * give up" pattern, which was racing Gmail's send-indexing lag and leaving
 * column U empty on slow indexing days. That orphaned the row and made every
 * subsequent send create yet another fresh thread.
 *
 * @param {string} [newStatus]
 * @param {number} [sentAtMs] Epoch ms when the email was sent (for precise thread matching).
 * @returns {string|null} The new thread ID, or null if we couldn't find it.
 */
function findAndSaveProjectThreadIdAfterSend_(mainSheet, mainRowIndex, subject, newStatus, sentAtMs) {
  const escaped = subject.replace(/"/g, '\\"');
  const wantSub = normalizeProjectEmailSubject_(subject);
  const sentAt = sentAtMs || Date.now();
  const MATCH_WINDOW_MS = 30000;

  for (let attempt = 0; attempt < PROJECT_THREAD_LOOKUP_DELAYS_MS.length; attempt++) {
    Utilities.sleep(PROJECT_THREAD_LOOKUP_DELAYS_MS[attempt]);

    const windowMinutes = attempt === 0 ? 2 : (attempt === 1 ? 5 : 10);

    const queries = [
      'in:sent subject:"' + escaped + '" newer_than:' + windowMinutes + 'm',
      'subject:"' + escaped + '" newer_than:' + windowMinutes + 'm'
    ];

    let bestThread = null;
    let bestDelta = MATCH_WINDOW_MS + 1;

    for (let q = 0; q < queries.length; q++) {
      let threads;
      try {
        threads = GmailApp.search(queries[q], 0, 10);
      } catch (searchErr) {
        Logger.log('Thread lookup search failed (' + queries[q] + '): ' + searchErr);
        continue;
      }

      const picked = pickBestGmailThreadForSubject_(threads, wantSub, {
        mainRowIndex: mainRowIndex,
        sentAtMs: sentAt,
        matchWindowMs: MATCH_WINDOW_MS
      });

      if (picked) {
        const lastMsg = picked.getMessages()[picked.getMessageCount() - 1];
        const delta = Math.abs(lastMsg.getDate().getTime() - sentAt);
        if (delta < bestDelta) {
          bestDelta = delta;
          bestThread = picked;
        }
      }
    }

    if (bestThread) {
      const savedThreadId = bestThread.getId();
      setProjectThreadIdForRow_(mainSheet, mainRowIndex, savedThreadId);
      Logger.log(
        'Thread ID saved on attempt ' + (attempt + 1) + ' (delta ' + bestDelta + 'ms): ' + savedThreadId
      );
      if (newStatus) {
        manageThreadLabels(savedThreadId, newStatus);
      }
      return savedThreadId;
    }
  }

  Logger.log(
    'WARNING: Could not save Thread ID after ' + PROJECT_THREAD_LOOKUP_DELAYS_MS.length +
    ' attempts for subject "' + subject + '" (row ' + mainRowIndex + '). ' +
    'Next email for this row will create yet another new thread. ' +
    'Run autoRecoverProjectThreadIds() or diagnoseProjectRowEmailThread(' + mainRowIndex + ') to recover.'
  );
  return null;
}

/**
 * CORE LOGIC: Handles Gmail threading with HTML support + Label Management
 * Uses GmailMessage.reply() to guarantee replies stay in the same thread.
 * @returns {boolean} True when the email was sent successfully.
 */
function sendProjectTrackerEmail_(subject, plainTextBody, htmlBody, mainSheet, mainRowIndex, newStatus) {
  subject = ensureProjectEmailSubject_(subject);
  try {
    const threadCol = findProjectThreadIdColumn_(mainSheet);
    const existingThreadId = getProjectThreadIdForRow_(mainSheet, mainRowIndex);

    Logger.log(
      'sendProjectTrackerEmail_ row=' + mainRowIndex +
      ' threadCol=' + columnToLetter_(threadCol) +
      ' threadId=' + (existingThreadId || '(empty)') +
      ' subject=' + subject
    );

    if (!existingThreadId || existingThreadId.toString().trim() === "") {
      const sentAt = Date.now();
      GmailApp.sendEmail(getProjectMyEmail_(), subject, plainTextBody, mergeEmailOptionsWithBcc_({
        htmlBody: htmlBody || plainTextBody
      }));

      findAndSaveProjectThreadIdAfterSend_(mainSheet, mainRowIndex, subject, newStatus, sentAt);
      return true;
    }

    try {
      const thread = GmailApp.getThreadById(existingThreadId);
      if (!thread) throw new Error('Thread not found for ID: ' + existingThreadId);

      const MAX_MESSAGES_PER_THREAD = 95;
      let messageCount = 0;
      try {
        messageCount = thread.getMessageCount();
      } catch (countErr) {
        messageCount = 0;
      }
      if (messageCount >= MAX_MESSAGES_PER_THREAD) {
        throw new Error(
          'Thread ' + existingThreadId + ' has ' + messageCount +
          ' messages — at Gmail 100-msg cap. Starting new thread.'
        );
      }

      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      lastMessage.reply(plainTextBody, mergeEmailOptionsWithBcc_({
        htmlBody: htmlBody || plainTextBody
      }));
      Logger.log('Reply sent in existing thread: ' + existingThreadId);

      if (newStatus) {
        manageThreadLabels(existingThreadId, newStatus);
      }
      return true;

    } catch (threadError) {
      Logger.log('Thread reply failed, starting new thread: ' + threadError.toString());
      setProjectThreadIdForRow_(mainSheet, mainRowIndex, '');

      const sentAt = Date.now();
      GmailApp.sendEmail(getProjectMyEmail_(), subject, plainTextBody, mergeEmailOptionsWithBcc_({
        htmlBody: htmlBody || plainTextBody
      }));

      findAndSaveProjectThreadIdAfterSend_(mainSheet, mainRowIndex, subject, newStatus, sentAt);
      return true;
    }
  } catch (error) {
    Logger.log('sendProjectTrackerEmail_ Error: ' + error.toString());
    try {
      const sentAt = Date.now();
      GmailApp.sendEmail(getProjectMyEmail_(), subject, plainTextBody, mergeEmailOptionsWithBcc_({
        htmlBody: htmlBody || plainTextBody
      }));
      findAndSaveProjectThreadIdAfterSend_(mainSheet, mainRowIndex, subject, newStatus, sentAt);
      return true;
    } catch (fallbackError) {
      Logger.log('Fallback email also failed: ' + fallbackError.toString());
      return false;
    }
  }
}

/**
 * Serializes pending-queue writes/clears across concurrent onEdit + hourly triggers.
 * @param {function(): *} fn
 * @returns {*}
 */
function withPendingQueueLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('Pending queue lock not acquired — operation skipped.');
    return null;
  }
  try {
    return fn();
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/**
 * QUEUE: Logs a change to the Pending Changes sheet instead of sending immediately
 */
function logPendingChange(subject, mainRowIndex) {
  withPendingQueueLock_(function () {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let pendingSheet = ss.getSheetByName(getPendingSheetName_());

      if (!pendingSheet) {
        pendingSheet = ss.insertSheet(getPendingSheetName_());
        pendingSheet.hideSheet();
        pendingSheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Subject', 'MainRowIndex']]);
        pendingSheet.setFrozenRows(1);
      }

      pendingSheet.appendRow([new Date(), subject, mainRowIndex]);
    } catch (error) {
      Logger.log('logPendingChange Error: ' + error.toString());
    }
  });
}

/**
 * TRIGGER: Runs every hour — reads all pending changes, sends 1 batched email per task, clears queue
 * Set up: Triggers > Add Trigger > sendHourlyBatch > Time-driven > Hour timer > Every hour
 */
function sendHourlyBatch() {
  const snapshot = withPendingQueueLock_(function () {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const pendingSheet = ss.getSheetByName(getPendingSheetName_());
      if (!pendingSheet) {
        return null;
      }

      const allData = pendingSheet.getDataRange().getValues();
      if (allData.length <= 1) {
        return null;
      }

      const groups = groupPendingChangesByTask_(allData);
      clearPendingChangesQueueUnlocked_(pendingSheet);

      return {
        allData: allData,
        groups: groups,
        pendingSheet: pendingSheet,
        mainSheet: ss.getSheetByName(getMainSheetName_())
      };
    } catch (error) {
      Logger.log('sendHourlyBatch snapshot Error: ' + error.toString());
      return null;
    }
  });

  if (!snapshot) {
    Logger.log('No pending changes to send.');
    return;
  }

  if (!snapshot.mainSheet) {
    Logger.log('Main sheet not found — restoring pending queue.');
    withPendingQueueLock_(function () {
      const allKeys = {};
      for (const key in snapshot.groups) {
        if (Object.prototype.hasOwnProperty.call(snapshot.groups, key)) {
          allKeys[key] = true;
        }
      }
      reappendPendingQueueRowsByKeys_(snapshot.pendingSheet, snapshot.allData, allKeys);
    });
    return;
  }

  const keysToRetry = {};
  let sentCount = 0;
  let retryCount = 0;

  for (const key in snapshot.groups) {
    if (!Object.prototype.hasOwnProperty.call(snapshot.groups, key)) {
      continue;
    }
    try {
      const outcome = processOneHourlyBatchTask_(snapshot.groups[key], snapshot.mainSheet);
      if (outcome === 'sent' || outcome === 'clear') {
        if (outcome === 'sent') {
          sentCount++;
        }
      } else {
        keysToRetry[key] = true;
        retryCount++;
      }
    } catch (taskError) {
      keysToRetry[key] = true;
      retryCount++;
      Logger.log('sendHourlyBatch task error (' + snapshot.groups[key].subject + '): ' + taskError.toString());
    }
  }

  withPendingQueueLock_(function () {
    reappendPendingQueueRowsByKeys_(snapshot.pendingSheet, snapshot.allData, keysToRetry);
  });

  Logger.log(
    'Hourly batch complete. Sent: ' + sentCount +
    ', re-queued for retry: ' + retryCount
  );
}

/**
 * Deduplicates rapid-fire edits on the same task so we send one email per
 * (subject, mainRowIndex) pair instead of one per edit.
 * @returns {Object<string,{subject: string, mainRowIndex: *, changeCount: number}>}
 */
function groupPendingChangesByTask_(allData) {
  const groups = {};
  for (let i = 1; i < allData.length; i++) {
    const subject = allData[i][1];
    const mainRowIndex = allData[i][2];
    if (!subject || !mainRowIndex) continue;
    const key = subject + '||' + mainRowIndex;
    if (!groups[key]) {
      groups[key] = { subject: subject, mainRowIndex: mainRowIndex, changeCount: 0 };
    }
    groups[key].changeCount++;
  }
  return groups;
}

/**
 * Pipeline for one grouped pending task: resolve row → validate → build → send.
 * @returns {'sent'|'clear'|'retry'}
 */
function processOneHourlyBatchTask_(group, mainSheet) {
  const subject = group.subject;
  const mainRowIndex = group.mainRowIndex;
  const changeCount = group.changeCount;

  const resolution = resolveMainRowForPendingTask_(mainSheet, subject, mainRowIndex);
  if (!resolution) {
    Logger.log('Skipping hourly email; task not found in main sheet for subject: ' + subject);
    return 'clear';
  }

  const rowData = resolution.rowData;
  const resolvedRowIndex = resolution.resolvedRowIndex;

  const taskName = rowData[CONFIG.MAIN_TASK_NAME_COL - 1];
  const currentStatus = rowData[CONFIG.MAIN_STATUS_COL - 1];

  if (isCompletedOrClosedStatus_(currentStatus)) {
    Logger.log('Skipping hourly email for completed/closed task: ' + (subject || taskName));
    return 'clear';
  }

  if (!mainRowHasMeaningfulTaskData_(rowData)) {
    Logger.log('Skipping hourly email for empty row data: ' + subject);
    return 'clear';
  }

  let subjectToSend = (rowData[CONFIG.MAIN_SUBJECT_COL - 1] || '').toString().trim();
  if (!subjectToSend && taskName) {
    subjectToSend = buildDefaultInternalMailSubject_(taskName);
  }
  if (!subjectToSend) {
    subjectToSend = subject;
  }

  const bodies = buildHourlyBatchEmailBodies_(taskName, rowData, changeCount, resolvedRowIndex);

  const sent = sendProjectTrackerEmail_(
    subjectToSend,
    bodies.plain,
    bodies.html,
    mainSheet,
    resolvedRowIndex,
    currentStatus
  );

  if (!sent) {
    throw new Error('Email send failed for subject: ' + subjectToSend);
  }

  Logger.log('Hourly batch sent for: ' + subjectToSend + ' (' + bodies.changeLabel + ')');
  return 'sent';
}

/**
 * Row indices in the pending queue may have shifted (tasks moved or deleted).
 * First try the indexed row, then fall back to a subject scan of the main
 * sheet. Returns null when the task can't be located (likely moved to the
 * Completed Tracker).
 * @returns {{rowData: any[], resolvedRowIndex: number}|null}
 */
function resolveMainRowForPendingTask_(mainSheet, subject, mainRowIndex) {
  let resolvedRowIndex = Number(mainRowIndex);
  const normalizedTarget = subject ? subject.toString().trim() : '';

  try {
    const candidate = mainSheet.getRange(resolvedRowIndex, 1, 1, 21).getValues()[0];
    const candidateSubject = (candidate[CONFIG.MAIN_SUBJECT_COL - 1] || '').toString().trim();
    if (candidateSubject && normalizedTarget && candidateSubject === normalizedTarget) {
      return { rowData: candidate, resolvedRowIndex: resolvedRowIndex };
    }
  } catch (e) {
    // Fall through to a full scan
  }

  const mainData = mainSheet.getDataRange().getValues();
  for (let i = 1; i < mainData.length; i++) {
    const s = (mainData[i][CONFIG.MAIN_SUBJECT_COL - 1] || '').toString().trim();
    if (s && normalizedTarget && s === normalizedTarget) {
      return { rowData: mainData[i], resolvedRowIndex: i + 1 };
    }
  }

  return null;
}

/**
 * Guardrail: don't send a "task update" email when every monitored field is
 * blank — usually means the row was cleared mid-edit.
 */
function mainRowHasMeaningfulTaskData_(rowData) {
  const taskName = rowData[CONFIG.MAIN_TASK_NAME_COL - 1];
  return (
    (taskName && taskName.toString().trim() !== '') ||
    (rowData[CONFIG.MAIN_STATUS_COL - 1] && rowData[CONFIG.MAIN_STATUS_COL - 1].toString().trim() !== '') ||
    (rowData[CONFIG.MAIN_ENV_COL - 1] && rowData[CONFIG.MAIN_ENV_COL - 1].toString().trim() !== '') ||
    (rowData[CONFIG.MAIN_INFORMED_COL - 1] && rowData[CONFIG.MAIN_INFORMED_COL - 1].toString().trim() !== '') ||
    (rowData[CONFIG.MAIN_BLOCKER_COL - 1] && rowData[CONFIG.MAIN_BLOCKER_COL - 1].toString().trim() !== '') ||
    (rowData[CONFIG.MAIN_DEPENDENCY_COL - 1] && rowData[CONFIG.MAIN_DEPENDENCY_COL - 1].toString().trim() !== '') ||
    (rowData[CONFIG.MAIN_PAGE_URL_COL - 1] && rowData[CONFIG.MAIN_PAGE_URL_COL - 1].toString().trim() !== '')
  );
}

/**
 * Builds the HTML and plain-text bodies for the hourly batched task update.
 * @returns {{html: string, plain: string, changeLabel: string}}
 */
function buildHourlyBatchEmailBodies_(taskName, rowData, changeCount, mainRowIndex) {
  const changeLabel = changeCount === 1 ? '1 change' : changeCount + ' changes';
  const subTasksTable = getRelatedSubTasks(taskName);
  const status = rowData[CONFIG.MAIN_STATUS_COL - 1];
  const env = rowData[CONFIG.MAIN_ENV_COL - 1];
  const informed = rowData[CONFIG.MAIN_INFORMED_COL - 1];
  const blocker = rowData[CONFIG.MAIN_BLOCKER_COL - 1];
  const dependency = rowData[CONFIG.MAIN_DEPENDENCY_COL - 1];
  const pageUrl = rowData[CONFIG.MAIN_PAGE_URL_COL - 1];
  const statusColor = getStatusColor(status, env);
  const rowMarker = mainRowIndex ? buildDevTrackerRowMarker_(mainRowIndex) : '';

  const html = `<html><body style="font-family: Arial, sans-serif; padding: 20px;">
<div style="border-left: 4px solid #3498db; padding-left: 15px; margin-bottom: 20px;">
  <h2 style="color: #2c3e50; margin: 0;">TASK UPDATE</h2>
  <p style="color: #7f8c8d; margin: 5px 0 0 0; font-size: 13px;">
    ${escapeHtml(changeLabel)} in the last hour
  </p>
</div>

<table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
  <tr style="background-color: #ecf0f1;">
    <td style="padding: 10px; font-weight: bold; width: 30%;">Task:</td>
    <td style="padding: 10px;">${escapeHtml(taskName)}</td>
  </tr>
  <tr>
    <td style="padding: 10px; font-weight: bold;">Status:</td>
    <td style="padding: 10px; background-color: ${statusColor}; font-weight: bold;">${escapeHtml(status || 'N/A')}</td>
  </tr>
  <tr style="background-color: #ecf0f1;">
    <td style="padding: 10px; font-weight: bold;">Environment:</td>
    <td style="padding: 10px;">${escapeHtml(env || 'N/A')}</td>
  </tr>
  <tr>
    <td style="padding: 10px; font-weight: bold;">Informed:</td>
    <td style="padding: 10px;">${escapeHtml(informed || 'N/A')}</td>
  </tr>
  <tr style="background-color: #ecf0f1;">
    <td style="padding: 10px; font-weight: bold;">Support Needed/Blocker:</td>
    <td style="padding: 10px; color: ${blocker ? '#e74c3c' : '#7f8c8d'}; font-weight: ${blocker ? 'bold' : 'normal'};">
      ${escapeHtml(blocker || 'None')}
    </td>
  </tr>
  <tr>
    <td style="padding: 10px; font-weight: bold;">Dependency Detail:</td>
    <td style="padding: 10px;">${escapeHtml(dependency || 'None')}</td>
  </tr>
  <tr style="background-color: #ecf0f1;">
    <td style="padding: 10px; font-weight: bold;">Page URL:</td>
    <td style="padding: 10px;">${renderProjectUrlCell_(pageUrl)}</td>
  </tr>
</table>

${subTasksTable}

<hr style="border: none; border-top: 1px solid #bdc3c7; margin: 20px 0;">
<p style="color: #95a5a6; font-size: 12px;">
  Batch sent at: ${new Date().toLocaleString()} &nbsp;|&nbsp; ${changeLabel} batched<br>
  Auto-generated by Project Tracker
</p>
${rowMarker ? '<!-- ' + escapeHtml(rowMarker) + ' -->' : ''}
</body></html>`;

  const plain = `--- TASK UPDATE (${changeLabel} batched) ---
Task: ${taskName || 'N/A'}
Status: ${status || 'N/A'}
Environment: ${env || 'N/A'}
Informed: ${informed || 'N/A'}
Support Needed/Blocker: ${blocker || 'None'}
Dependency Detail: ${dependency || 'None'}
Page URL: ${pageUrl || 'N/A'}

Batch sent at: ${new Date().toLocaleString()}${rowMarker ? '\n' + rowMarker : ''}`;

  return { html: html, plain: plain, changeLabel: changeLabel };
}

/**
 * Removes every data row from the pending queue (keeps the header). Caller must hold the lock.
 */
function clearPendingChangesQueueUnlocked_(pendingSheet) {
  const rowCount = pendingSheet.getLastRow();
  if (rowCount > 1) {
    pendingSheet.deleteRows(2, rowCount - 1);
  }
}

/**
 * Re-appends pending rows whose batch send failed. Caller must hold the lock.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} pendingSheet
 * @param {Array} allData Snapshot taken before the queue was cleared
 * @param {Object<string, boolean>} keysToRetry
 */
function reappendPendingQueueRowsByKeys_(pendingSheet, allData, keysToRetry) {
  if (!pendingSheet || !allData || allData.length <= 1) {
    return;
  }

  for (let i = 1; i < allData.length; i++) {
    const subject = allData[i][1];
    const mainRowIndex = allData[i][2];
    if (!subject || !mainRowIndex) {
      continue;
    }
    const key = subject + '||' + mainRowIndex;
    if (keysToRetry[key]) {
      pendingSheet.appendRow([allData[i][0], subject, mainRowIndex]);
    }
  }
}

/**
 * Removes pending queue rows for tasks that sent successfully or are safe to drop.
 * Failed sends remain for the next hourly run.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} pendingSheet
 * @param {Array} allData
 * @param {Object<string, boolean>} keysToClear
 */
function removePendingQueueRowsByKeys_(pendingSheet, allData, keysToClear) {
  if (!pendingSheet || !allData || allData.length <= 1) {
    return;
  }

  for (let i = allData.length - 1; i >= 1; i--) {
    const subject = allData[i][1];
    const mainRowIndex = allData[i][2];
    if (!subject || !mainRowIndex) {
      continue;
    }
    const key = subject + '||' + mainRowIndex;
    if (keysToClear[key]) {
      pendingSheet.deleteRow(i + 1);
    }
  }
}

/**
 * Removes every data row from the pending queue (keeps the header).
 * Prefer removePendingQueueRowsByKeys_ after hourly batch — this is for manual resets.
 */
function clearPendingChangesQueue_(pendingSheet) {
  withPendingQueueLock_(function () {
    clearPendingChangesQueueUnlocked_(pendingSheet);
  });
}

/**
 * SETUP HELPER: Run this once manually to create the hourly trigger automatically
 */
function autoRecoverProjectThreadIds() {
  try {
    Logger.log('=== autoRecoverProjectThreadIds (read-only) ===');
    const mainSheet = getMainTrackerSheet_();
    if (!mainSheet) {
      Logger.log('Dev Tracker sheet not found.');
      return;
    }

    const lastRow = mainSheet.getLastRow();
    let recovered = 0;

    for (let row = 2; row <= lastRow; row++) {
      const existingId = getProjectThreadIdForRow_(mainSheet, row);
      if (existingId) {
        continue;
      }

      let subject = (mainSheet.getRange(row, CONFIG.MAIN_SUBJECT_COL).getValue() || '').toString().trim();
      const taskName = (mainSheet.getRange(row, CONFIG.MAIN_TASK_NAME_COL).getValue() || '').toString().trim();
      if (!subject && taskName) {
        subject = buildDefaultInternalMailSubject_(taskName);
      }
      if (!subject) {
        continue;
      }

      const escaped = subject.replace(/"/g, '\\"');
      const wantSub = normalizeProjectEmailSubject_(subject);
      let threads;
      try {
        threads = GmailApp.search(
          'subject:"' + escaped + '" newer_than:180d',
          0,
          15
        );
      } catch (searchErr) {
        Logger.log('Row ' + row + ' search failed: ' + searchErr);
        continue;
      }

      const bestThread = pickBestGmailThreadForSubject_(threads, wantSub, {
        mainRowIndex: row,
        preferRecent: true
      });

      if (bestThread) {
        setProjectThreadIdForRow_(mainSheet, row, bestThread.getId());
        recovered++;
        Logger.log('Row ' + row + ': recovered thread ' + bestThread.getId());
      }
    }

    Logger.log('autoRecoverProjectThreadIds: recovered ' + recovered + ' thread ID(s).');
  } catch (error) {
    Logger.log('autoRecoverProjectThreadIds Error: ' + error.toString());
  }
}

/**
 * SETUP HELPER: Run this once manually to create ALL required triggers
 * Creates/refreshes:
 * - projectOnEditHandler (installable spreadsheet trigger)
 * - sendHourlyBatch (every hour, or CONFIG.BATCH_INTERVAL_HOURS)
 * - sendDailySummary (6 AM and 6 PM)
 * - sendWeeklySummary (Friday 6 PM)
 * - autoRecoverProjectThreadIds (daily 5 AM safety net for orphaned Thread IDs)
 */
function getActiveMainTaskNameSet_(mainData) {
  const set = {};
  const col = CONFIG.MAIN_TASK_NAME_COL - 1;
  for (let i = 1; i < mainData.length; i++) {
    const n = (mainData[i][col] || '').toString().trim();
    if (n) set[n] = true;
  }
  return set;
}

function subRowLinksToActiveMain_(subRow, activeMainNames) {
  const link = (subRow[CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
  return link !== '' && activeMainNames[link] === true;
}

/**
 * TRIGGER: Sends the Daily Summary at 6 AM and 6 PM
 * To set up: Go to Triggers > Add Trigger > sendDailySummary > Time-driven > Day timer > 6am-7am and 6pm-7pm
 */
function sendDailySummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = getMainTrackerSheet_();
    const subSheet = ss.getSheetByName(getSubSheetName_());

    if (!mainSheet || !subSheet) {
      Logger.log('Error: Could not find required sheets');
      return;
    }

    const mainData = mainSheet.getDataRange().getValues();
    const subData = subSheet.getDataRange().getValues();

    const htmlBody = buildDailySummaryHtml_(mainData, subData);
    const emailSubject = buildDailySummarySubject_();

    GmailApp.sendEmail(
      getProjectMyEmail_(),
      emailSubject,
      "Please view this email in HTML format.",
      mergeEmailOptionsWithBcc_({ htmlBody: htmlBody })
    );

    Logger.log('Daily summary sent successfully');
  } catch (error) {
    Logger.log('sendDailySummary Error: ' + error.toString());
  }
}

/**
 * TRIGGER: Sends the Weekly Summary on Friday 6 PM
 * Includes:
 * - All current rows in Dev Tracker (main + sub sheets)
 * - Only last 7 days of Completed/Closed tasks from Completed Tracker
 *
 * Set up: Triggers > Add Trigger > sendWeeklySummary > Time-driven > Week timer > Friday > 6pm-7pm
 */
function sendWeeklySummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = getMainTrackerSheet_();
    const subSheet = ss.getSheetByName(getSubSheetName_());
    const completedSheet = ss.getSheetByName(getCompletedSheetName_());

    if (!mainSheet || !subSheet) {
      Logger.log('Error: Could not find required sheets for weekly summary');
      return;
    }

    const now = new Date();
    const lookbackDays = 7;

    const mainData = mainSheet.getDataRange().getValues();
    const subData = subSheet.getDataRange().getValues();

    const htmlBody = buildWeeklySummaryHtml_(mainData, subData, completedSheet, lookbackDays, now);
    const emailSubject = buildWeeklySummarySubject_(now);

    GmailApp.sendEmail(
      getProjectMyEmail_(),
      emailSubject,
      "Please view this email in HTML format.",
      mergeEmailOptionsWithBcc_({ htmlBody: htmlBody })
    );

    Logger.log('Weekly summary sent successfully');
  } catch (error) {
    Logger.log('sendWeeklySummary Error: ' + error.toString());
  }
}

/**
 * Builds the full HTML body for the daily summary email.
 * @returns {string}
 */
function buildDailySummaryHtml_(mainData, subData) {
  const mainTable = buildSummaryMainProjectsTable_(mainData, 'Main Projects', 'Total Main Tasks');
  const subTable = buildSummarySubTasksTable_(subData, mainData, 'Sub-Task Status', 'Total Sub-Tasks');

  return `<html><body style="font-family: Arial, sans-serif;">
<h2 style="color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px;">
  Project Tracker Daily Summary
</h2>${mainTable}<br>${subTable}<hr style="margin-top: 30px; border: none; border-top: 1px solid #bdc3c7;">
<p style="color: #95a5a6; font-size: 11px;">
  Generated automatically at: ${new Date().toLocaleString()}<br>
  Report Type: Daily Summary
</p>
</body></html>`;
}

/**
 * Builds the full HTML body for the weekly summary email (current main + sub
 * tables, plus the last-7-days completed/closed table).
 * @returns {string}
 */
function buildWeeklySummaryHtml_(mainData, subData, completedSheet, lookbackDays, now) {
  const mainTable = buildSummaryMainProjectsTable_(
    mainData,
    'Current Main Tasks (Dev Tracker)',
    'Total Current Main Tasks'
  );
  const subTable = buildSummarySubTasksTable_(
    subData,
    mainData,
    'Current Sub-Tasks (Dev Tracker Sub)',
    'Total Current Sub-Tasks'
  );

  const recentCompleted = getRecentlyMovedCompletedRows_(completedSheet, lookbackDays, now);
  const completedSection = `<h3 style="color: #34495e; margin-top: 20px;">Completed/Closed in Last ${lookbackDays} Days (Completed Tracker)</h3>${recentCompleted.html}<p style="color: #7f8c8d; font-size: 12px;">Total Recent Completed/Closed: ${recentCompleted.count}</p>`;

  return `<html><body style="font-family: Arial, sans-serif;">
<h2 style="color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px;">
  Project Tracker Weekly Summary (Friday)
</h2>
<p style="color: #7f8c8d; margin-top: 0;">
  Generated: ${escapeHtml(now.toLocaleString())}<br>
  Completed/Closed included: last ${lookbackDays} days only
</p>${mainTable}${subTable}${completedSection}<hr style="margin-top: 30px; border: none; border-top: 1px solid #bdc3c7;">
<p style="color: #95a5a6; font-size: 11px;">
  Report Type: Weekly Summary (Friday 6 PM)
</p>
</body></html>`;
}

/**
 * Renders the main projects table used by both daily and weekly summaries.
 * The heading and the "total" label differ between reports, so they are
 * passed in.
 * @returns {string}
 */
function buildSummaryMainProjectsTable_(mainData, heading, totalLabel) {
  let bodyRows = '';
  let mainCount = 0;

  for (let i = 1; i < mainData.length; i++) {
    const taskName = (mainData[i][CONFIG.MAIN_TASK_NAME_COL - 1] || '').toString().trim();
    if (!taskName) {
      continue;
    }

    const statusColor = getStatusColor(
      mainData[i][CONFIG.MAIN_STATUS_COL - 1],
      mainData[i][CONFIG.MAIN_ENV_COL - 1]
    );
    bodyRows += `<tr style="background-color: ${i % 2 === 0 ? '#ecf0f1' : 'white'};">
        <td>${escapeHtml(mainData[i][CONFIG.MAIN_PROJECT_COL - 1])}</td>
        <td>${escapeHtml(taskName)}</td>
        <td style="background-color: ${statusColor}; font-weight: bold;">${escapeHtml(mainData[i][CONFIG.MAIN_STATUS_COL - 1])}</td>
        <td>${escapeHtml(mainData[i][CONFIG.MAIN_ENV_COL - 1])}</td>
        <td>${escapeHtml(mainData[i][CONFIG.MAIN_ASSIGNED_COL - 1])}</td>
        <td>${escapeHtml(mainData[i][CONFIG.MAIN_BLOCKER_COL - 1])}</td>
        <td>${escapeHtml(mainData[i][CONFIG.MAIN_DEPENDENCY_COL - 1])}</td>
      </tr>`;
    mainCount++;
  }

  return `<h3 style="color: #34495e; margin-top: 20px;">${heading}</h3>
<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #bdc3c7;">
  <thead>
    <tr style="background-color: #3498db; color: white;">
      <th>Project</th>
      <th>Task</th>
      <th>Status</th>
      <th>Environment</th>
      <th>Assigned</th>
      <th>Blockers</th>
      <th>Dependency</th>
    </tr>
  </thead>
  <tbody>${bodyRows}</tbody></table>
<p style="color: #7f8c8d; font-size: 12px;">${totalLabel}: ${mainCount}</p>`;
}

/**
 * Renders the sub-tasks table used by both daily and weekly summaries.
 * Filters out orphan sub-rows whose parent task is no longer on the Dev
 * Tracker (parent moved to Completed Tracker).
 * @returns {string}
 */
function buildSummarySubTasksTable_(subData, mainData, heading, totalLabel) {
  const activeMainTaskNames = getActiveMainTaskNameSet_(mainData);

  const subEntries = [];
  for (let j = 1; j < subData.length; j++) {
    if (!subData[j][CONFIG.SUB_LINK_COL - 1] || subData[j][CONFIG.SUB_LINK_COL - 1].toString().trim() === '') {
      continue;
    }
    if (!subRowLinksToActiveMain_(subData[j], activeMainTaskNames)) continue;
    subEntries.push({ sheetRow: j + 1, data: subData[j] });
  }
  sortSubTaskRowsForSummaryEmail_(subEntries);

  let bodyRows = '';
  for (let k = 0; k < subEntries.length; k++) {
    const row = subEntries[k].data;
    const statusColor = getStatusColor(
      row[CONFIG.SUB_STATUS_COL - 1],
      row[CONFIG.SUB_ENV_COL - 1]
    );
    bodyRows += `<tr style="background-color: ${k % 2 === 0 ? '#ecf0f1' : 'white'};">
        <td>${escapeHtml(row[CONFIG.SUB_LINK_COL - 1])}</td>
        <td>${escapeHtml(row[CONFIG.SUB_TASK_NAME_COL - 1])}</td>
        <td style="background-color: ${statusColor}; font-weight: bold;">${escapeHtml(row[CONFIG.SUB_STATUS_COL - 1])}</td>
        <td>${escapeHtml(row[CONFIG.SUB_ENV_COL - 1])}</td>
        <td>${escapeHtml(row[CONFIG.SUB_ASSIGNED_COL - 1])}</td>
      </tr>`;
  }

  return `<h3 style="color: #34495e; margin-top: 20px;">${heading}</h3>
<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #bdc3c7;">
  <thead>
    <tr style="background-color: #2ecc71; color: white;">
      <th>Main Task Link</th>
      <th>Task</th>
      <th>Status</th>
      <th>Environment</th>
      <th>Assigned</th>
    </tr>
  </thead>
  <tbody>${bodyRows}</tbody></table>
<p style="color: #7f8c8d; font-size: 12px;">${totalLabel}: ${subEntries.length}</p>`;
}

function buildDailySummarySubject_() {
  return 'Project Summary Report: ' + new Date().toLocaleDateString('en-US', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
}

function buildWeeklySummarySubject_(now) {
  return 'Weekly Project Summary (Fri): ' + now.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
}

function getRecentlyMovedCompletedRows_(completedSheet, lookbackDays, now) {
  try {
    if (!completedSheet) {
      return {
        html: '<p style="color: #7f8c8d; font-style: italic;">No Completed Tracker sheet found.</p>',
        count: 0
      };
    }

    const data = completedSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        html: '<p style="color: #7f8c8d; font-style: italic;">No completed/closed tasks found.</p>',
        count: 0
      };
    }

    const header = data[0].map(h => (h || '').toString().trim());
    const movedAtIdx = header.findIndex(h => h.toLowerCase() === 'moved at');
    if (movedAtIdx === -1) {
      return {
        html: '<p style="color: #7f8c8d; font-style: italic;">Completed Tracker has no "Moved At" column yet (new moves will include it).</p>',
        count: 0
      };
    }

    const cutoff = new Date(now.getTime() - lookbackDays * 24 * 60 * 60 * 1000);
    const recent = [];
    for (let i = 1; i < data.length; i++) {
      const movedAt = data[i][movedAtIdx];
      if (!(movedAt instanceof Date)) continue;
      if (movedAt >= cutoff && movedAt <= now) {
        recent.push(data[i]);
      }
    }

    if (recent.length === 0) {
      return {
        html: '<p style="color: #7f8c8d; font-style: italic;">No completed/closed tasks in the last 7 days.</p>',
        count: 0
      };
    }

    // Completed Tracker uses the same column layout as Dev Tracker when rows are moved.
    const colProject = CONFIG.MAIN_PROJECT_COL - 1;
    const colTask = CONFIG.MAIN_TASK_NAME_COL - 1;
    const colStatus = CONFIG.MAIN_STATUS_COL - 1;
    const colEnv = CONFIG.MAIN_ENV_COL - 1;
    const colAssigned = CONFIG.MAIN_ASSIGNED_COL - 1;
    const colBlocker = CONFIG.MAIN_BLOCKER_COL - 1;
    const colDependency = CONFIG.MAIN_DEPENDENCY_COL - 1;

    let html = `<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #bdc3c7;">
  <thead>
    <tr style="background-color: #9b59b6; color: white;">
      <th>Project</th>
      <th>Task</th>
      <th>Status</th>
      <th>Environment</th>
      <th>Assigned</th>
      <th>Blockers</th>
      <th>Dependency</th>
      <th>Moved At</th>
    </tr>
  </thead>
  <tbody>`;

    for (let i = 0; i < recent.length; i++) {
      const r = recent[i];
      const statusColor = getStatusColor(r[colStatus], r[colEnv]);
      const movedAt = r[movedAtIdx];
      html += `<tr style="background-color: ${i % 2 === 0 ? '#ecf0f1' : 'white'};">
        <td>${escapeHtml(r[colProject])}</td>
        <td>${escapeHtml(r[colTask])}</td>
        <td style="background-color: ${statusColor}; font-weight: bold;">${escapeHtml(r[colStatus])}</td>
        <td>${escapeHtml(r[colEnv])}</td>
        <td>${escapeHtml(r[colAssigned])}</td>
        <td>${escapeHtml(r[colBlocker])}</td>
        <td>${escapeHtml(r[colDependency])}</td>
        <td>${escapeHtml(movedAt.toLocaleString())}</td>
      </tr>`;
    }

    html += `</tbody></table>`;
    return { html, count: recent.length };
  } catch (error) {
    Logger.log('getRecentlyMovedCompletedRows_ Error: ' + error.toString());
    return { html: '<p style="color: #e74c3c;">Error loading recent completed tasks.</p>', count: 0 };
  }
}

/**
 * Exact status label → background color for email/summary tables.
 * @returns {Object<string, string>}
 */
function getStatusColorMap_() {
  return {
    'Pipeline / To Do': '#d6eaf8',
    'Requirements Not Clear': '#fdebd0',
    'In Progress': '#fff4cc',
    'On Hold / Blocked': '#fadbd8',
    'Waiting for Client Input': '#d6eaf8',
    'Under Review & Testing': '#e8daef',
    'Changes Requested / Rework': '#fdebd0',
    'Ready for Execution': '#d5f4e6',
    'Completed': '#d5f4e6',
    'Closed / Cancelled': '#eaecee',
    'Closed – Billable': '#d5f4e6',
    'Closed – Not Proceeding': '#eaecee'
  };
}

/**
 * HELPER: Get color based on status
 */
function getStatusColor(status, environment) {
  const unknownColor = CONFIG.STATUS_COLOR_UNKNOWN || '#f0f0f0';
  const envLower = (environment || '').toString().trim().toLowerCase();
  if (envLower === 'production') {
    return '#d5f4e6';
  }

  const label = (status || '').toString().trim();
  if (!label) {
    return unknownColor;
  }

  const map = getStatusColorMap_();
  if (map[label]) {
    return map[label];
  }

  const normalized = normalizeGoogleTaskListName_(label).toLowerCase();
  for (let i = 0; i < CONFIG.STATUS_LABELS.length; i++) {
    const configLabel = CONFIG.STATUS_LABELS[i];
    if (normalizeGoogleTaskListName_(configLabel).toLowerCase() === normalized) {
      return map[configLabel] || unknownColor;
    }
  }

  return unknownColor;
}

/**
 * HELPER: Escape HTML to prevent rendering issues
 */
function escapeHtml(text) {
  if (!text) return '';
  return text.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Renders a Page URL cell value safely:
 *  - Only http:// and https:// values become clickable links, blocking
 *    javascript:, data:, vbscript: and other URI scheme injections that
 *    `escapeHtml` alone wouldn't stop (Gmail strips them, but forwarded
 *    copies and other mail clients may not).
 *  - Any other non-empty value renders as escaped plain text (no link).
 *  - Empty / null becomes "N/A".
 * @param {*} value
 * @returns {string} HTML fragment
 */
function renderProjectUrlCell_(value) {
  if (value === null || value === undefined) {
    return 'N/A';
  }
  const text = value.toString().trim();
  if (text === '') {
    return 'N/A';
  }
  if (/^https?:\/\//i.test(text)) {
    try {
      const url = new URL(text);
      if (url.protocol !== 'http:' && url.protocol !== 'https:') {
        return escapeHtml(text);
      }
      if (!url.hostname) {
        return escapeHtml(text);
      }
      const safe = escapeHtml(url.href);
      return '<a href="' + safe + '">' + safe + '</a>';
    } catch (urlError) {
      return escapeHtml(text);
    }
  }
  return escapeHtml(text);
}

/**
 * TEST FUNCTION: Run this manually to test the daily summary
 */
