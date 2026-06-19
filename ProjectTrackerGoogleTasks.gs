  /**
  * PROJECT TRACKER — GOOGLE TASKS
  * Sync Dev Tracker rows as parent Google Tasks; Dev Tracker Sub rows as nested subtasks.
  * Requires CONFIG and sheet helpers from ProjectTracker.gs (same Apps Script project).
  *
  * PUBLIC ENTRY POINTS:
  *   syncExistingTasksToGoogle()     — in TestAllTrackers.gs (calls syncAllGoogleTasksFromDevTracker)
  *   syncGoogleTaskForMainRow(2)    — on-edit / single main row (pass row number)
  *   syncGoogleSubtaskForSubRow()    — on-edit / single sub row
  *   backfillSubParentGoogleTaskIdsIfEmpty()
  *   resetAllGoogleTasksAndRecreate()
  *   verifyGoogleTaskLists()
  *   getGoogleTasksLastSyncLog()
  *   getGoogleTasksBulkSyncStatus()   — paused? phase, row, sub total
  *   continueGoogleTasksBulkSync_()   — auto-scheduled between batches
  *   refreshSubNestStatusColumn()  — fill column X (nest check vs column W)
  *   runFullGoogleTasksTestSuite()   — in ProjectTrackerTests.gs
  *
  * LARGE SHEETS (2000–5000 subs): CONFIG.GOOGLE_TASKS_SYNC_BATCH_SIZE = 300 per run;
  *   auto-continues every ~1 min until finished (GOOGLE_TASKS_AUTO_CONTINUE_BULK).
  *
  * SYNC MODES (log prefix [GoogleTasks:…]):
  *   ROW | TARGETED | BULK-MAIN | BULK-SUB | BULK
  */
  const SCRIPT_PROP_GOOGLE_SYNC_RESUME_ROW = "GOOGLE_TASKS_SYNC_RESUME_ROW";
  const SCRIPT_PROP_GOOGLE_SYNC_RESUME_PHASE = "GOOGLE_TASKS_SYNC_RESUME_PHASE";
  const SCRIPT_PROP_GOOGLE_TASKS_LAST_SYNC_LOG = "GOOGLE_TASKS_LAST_SYNC_LOG";
  const SCRIPT_PROP_GOOGLE_TASKS_BULK_IN_PROGRESS = "GOOGLE_TASKS_BULK_IN_PROGRESS";
  const SCRIPT_PROP_GOOGLE_TASKS_TEST_SUITE_RUNNING = "GOOGLE_TASKS_TEST_SUITE_RUNNING";
  const SCRIPT_PROP_GOOGLE_SYNC_SUB_TOTAL = "GOOGLE_TASKS_SYNC_SUB_TOTAL";

  /** Per-sheet header fingerprints used to invalidate column caches mid-execution. */
  let projectMainColumnCacheHeaderHash_ = null;
  let projectSubColumnCacheHeaderHash_ = null;

  /** Per-run Google Tasks sync mode for logging: ROW | TARGETED | BULK-MAIN | BULK-SUB | BULK */
  let googleTasksSyncContext_ = null;

  /**
  * Sync modes (see Executions log or getGoogleTasksLastSyncLog()):
  *   ROW              — on-edit, one main or sub row only
  *   TARGETED         — main Status edit → linked sub rows only
  *   BULK-MAIN        — full sheet bulk phase 1: all Dev Tracker main rows (parents first)
  *   BULK-SUB         — full sheet bulk phase 2: sub rows bottom-up
  *   BULK             — entire bulk run wrapper
  */
  function beginGoogleTasksSync_(mode, detail) {
    googleTasksSyncContext_ = {
      mode: mode || 'ROW',
      detail: detail || '',
      started: Date.now()
    };
    logGoogleTasksSync_('START ' + (detail || ''));
  }

  function logGoogleTasksSync_(message) {
    const mode = googleTasksSyncContext_ ? googleTasksSyncContext_.mode : 'ROW';
    Logger.log('[GoogleTasks:' + mode + '] ' + message);
  }

  function endGoogleTasksSync_(summary) {
    const ctx = googleTasksSyncContext_;
    const mode = ctx ? ctx.mode : 'ROW';
    const elapsed = ctx ? Date.now() - ctx.started : 0;
    const line = '[GoogleTasks:' + mode + '] DONE in ' + elapsed + 'ms — ' + (summary || '');
    Logger.log(line);
    try {
      PropertiesService.getScriptProperties().setProperty(SCRIPT_PROP_GOOGLE_TASKS_LAST_SYNC_LOG, line);
    } catch (e) {
      // non-fatal
    }
    googleTasksSyncContext_ = null;
  }

  /**
  * Returns the last Google Tasks sync summary (from script properties).
  * Run from Apps Script editor after an edit or bulk sync.
  * @returns {string}
  */
  function getGoogleTasksLastSyncLog() {
    try {
      return PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GOOGLE_TASKS_LAST_SYNC_LOG) ||
        '(no sync logged yet — edit a row or run syncExistingTasksToGoogle())';
    } catch (e) {
      return '(could not read sync log)';
    }
  }

  /**
  * True while bulk sync or the full test suite is running — onEdit Google Tasks sync is skipped.
  * @returns {boolean}
  */
  function isGoogleTasksSyncBlockedForOnEdit_() {
    try {
      const props = PropertiesService.getScriptProperties();
      return props.getProperty(SCRIPT_PROP_GOOGLE_TASKS_BULK_IN_PROGRESS) === 'true' ||
        props.getProperty(SCRIPT_PROP_GOOGLE_TASKS_TEST_SUITE_RUNNING) === 'true';
    } catch (e) {
      return false;
    }
  }

  /** @param {boolean} active */
  function setGoogleTasksBulkSyncInProgress_(active) {
    try {
      const props = PropertiesService.getScriptProperties();
      if (active) {
        props.setProperty(SCRIPT_PROP_GOOGLE_TASKS_BULK_IN_PROGRESS, 'true');
      } else {
        props.deleteProperty(SCRIPT_PROP_GOOGLE_TASKS_BULK_IN_PROGRESS);
      }
    } catch (e) {
      // non-fatal
    }
  }

  /**
  * @returns {boolean} True while a bulk sync execution holds the in-progress lock.
  */
  function isGoogleTasksBulkSyncInProgress_() {
    try {
      return PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GOOGLE_TASKS_BULK_IN_PROGRESS) === 'true';
    } catch (e) {
      return false;
    }
  }

  /** @param {boolean} active */
  function setGoogleTasksTestSuiteRunning_(active) {
    try {
      const props = PropertiesService.getScriptProperties();
      if (active) {
        props.setProperty(SCRIPT_PROP_GOOGLE_TASKS_TEST_SUITE_RUNNING, 'true');
      } else {
        props.deleteProperty(SCRIPT_PROP_GOOGLE_TASKS_TEST_SUITE_RUNNING);
      }
    } catch (e) {
      // non-fatal
    }
  }

  function ensureGoogleTasksRowSyncContext_(kind, row) {
    if (!googleTasksSyncContext_) {
      beginGoogleTasksSync_('ROW', kind + ' row ' + row);
    }
  }

  /**
  * Stable fingerprint of row-1 headers (lowercase, trimmed). Returns '' when sheet is null/missing.
  * @param {GoogleAppsScript.Spreadsheet.Sheet|null} sheet
  * @returns {string}
  */
  function getSheetHeaderFingerprint_(sheet) {
    if (!sheet) {
      return '';
    }
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      return '';
    }
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    return headers.map(function (h) {
      return (h || '').toString().toLowerCase().trim();
    }).join('|');
  }

  /**
  * Clears column-index caches when row-1 headers change during a long sync run.
  * Main and sub fingerprints are tracked independently so calls like
  * ensureColumnCachesFresh_(mainSheet, null) do not wipe sub column caches (and vice versa).
  * Either sheet may be omitted/null (e.g. Sub sheet missing during bulk main-only sync).
  * @param {GoogleAppsScript.Spreadsheet.Sheet|null} [mainSheet]
  * @param {GoogleAppsScript.Spreadsheet.Sheet|null} [subSheet]
  */
  function ensureColumnCachesFresh_(mainSheet, subSheet) {
    if (mainSheet) {
      const mainFp = getSheetHeaderFingerprint_(mainSheet);
      if (projectMainColumnCacheHeaderHash_ && projectMainColumnCacheHeaderHash_ !== mainFp) {
        projectThreadIdColCache_ = null;
        projectGoogleTaskIdColCache_ = null;
        Logger.log('Main sheet column caches invalidated — header row changed during this execution.');
      }
      projectMainColumnCacheHeaderHash_ = mainFp;
    }

    if (subSheet) {
      const subFp = getSheetHeaderFingerprint_(subSheet);
      if (projectSubColumnCacheHeaderHash_ && projectSubColumnCacheHeaderHash_ !== subFp) {
        projectGoogleSubtaskIdColCache_ = null;
        projectGoogleSubParentIdColCache_ = null;
        projectGoogleSubNestStatusColCache_ = null;
        Logger.log('Sub sheet column caches invalidated — header row changed during this execution.');
      }
      projectSubColumnCacheHeaderHash_ = subFp;
    }
  }

  /**
  * Clears all per-execution column and task-list lookup caches (e.g. after nuclear reset).
  */
  function clearProjectColumnCaches_() {
    projectMainColumnCacheHeaderHash_ = null;
    projectSubColumnCacheHeaderHash_ = null;
    projectThreadIdColCache_ = null;
    projectGoogleTaskIdColCache_ = null;
    projectGoogleSubtaskIdColCache_ = null;
    projectGoogleSubParentIdColCache_ = null;
    projectGoogleSubNestStatusColCache_ = null;
  }
  function isCompletedOrClosedStatus_(status) {
    const s = normalizeGoogleTaskListName_(status).toLowerCase();
    return s === 'completed' || s.indexOf('closed') === 0;
  }

  /**
  * Exact subtask statuses that must NOT sync to Google Tasks.
  * @returns {string[]}
  */
  function getGoogleTasksExcludedSubtaskStatuses_() {
    return CONFIG.STATUS_LABELS.slice(8);
  }

  /**
  * True when subtask Status (column I) is one of:
  * Completed, Closed / Cancelled, Closed – Billable, Closed – Not Proceeding
  * @param {string} status
  * @returns {boolean}
  */
  function isGoogleTasksExcludedSubtaskStatus_(status) {
    const normalized = normalizeGoogleTaskListName_(status).toLowerCase();
    if (!normalized) {
      return false;
    }

    const excludedStatuses = getGoogleTasksExcludedSubtaskStatuses_();
    for (let i = 0; i < excludedStatuses.length; i++) {
      if (normalizeGoogleTaskListName_(excludedStatuses[i]).toLowerCase() === normalized) {
        return true;
      }
    }

    return false;
  }

  /**
  * @param {string} env
  * @returns {boolean}
  */
  function isProductionEnv_(env) {
    return (env || '').toString().trim().toLowerCase() === 'production';
  }

  /**
  * Subtask is NOT added to Google Tasks only when BOTH are true:
  *   • Status is Completed, Closed / Cancelled, Closed – Billable, or Closed – Not Proceeding
  *   • AND Environment (column J) is Production
  *
  * All other combinations are added (e.g. Closed/Cancelled + Staging, In Progress + Production).
  * When skip applies and a task already exists, it is removed from Google Tasks.
  *
  * @param {string} status
  * @param {string} env
  * @returns {boolean}
  */
  function shouldSkipSubtaskGoogleSync_(status, env) {
    return isProductionEnv_(env) && isGoogleTasksExcludedSubtaskStatus_(status);
  }

  /**
  * Deletes the Google subtask for a Dev Tracker Sub row and clears column V.
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  */
  function removeGoogleSubtaskForRowIfExists_(subSheet, row) {
    if (!isGoogleTasksApiEnabled_() || !subSheet) {
      return;
    }

    const taskId = getGoogleSubtaskIdForSubRow_(subSheet, row);
    if (!taskId) {
      return;
    }

    const listId = findTaskListIdContainingTask_(taskId);
    if (listId) {
      try {
        Tasks.Tasks.remove(listId, taskId);
        invalidateTaskListIdCacheForTask_(taskId);
        Logger.log('Removed Google subtask for Dev Tracker Sub row ' + row + ': ' + taskId);
      } catch (error) {
        Logger.log('removeGoogleSubtaskForRowIfExists_ Error (row ' + row + '): ' + error.toString());
      }
    } else {
      removeGoogleTaskFromAnyList_(taskId);
    }

    setGoogleSubtaskIdForSubRow_(subSheet, row, '');
  }

  /**
  * Removes stray top-level Google Tasks whose title matches a sub-task name across all status lists.
  * Used when a sub row has no valid main-task link (orphans from earlier failed syncs).
  * @param {string} baseTitle Plain column D name
  * @param {string} [keepTaskId] Optional task ID to keep (e.g. a correctly nested child)
  * @param {string} [googleTitle] Suffixed Google Tasks title (optional)
  */
  function removeOrphanTopLevelTasksByTitleAcrossAllStatusLists_(baseTitle, keepTaskId, googleTitle) {
    const base = (baseTitle || '').toString().trim();
    if (!base || !isGoogleTasksApiEnabled_()) {
      return;
    }

    refreshGoogleTaskListCache_();
    for (let i = 0; i < CONFIG.STATUS_LABELS.length; i++) {
      const listId = getTaskListIdForStatus_(CONFIG.STATUS_LABELS[i]);
      if (!listId) {
        continue;
      }
      removeOrphanTopLevelDuplicateSubtasks_(listId, '', base, googleTitle || '', keepTaskId || '', true);
    }
  }

  /**
  * Removes Google Tasks for a sub row that must not sync (empty link, missing parent, excluded status, etc.).
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  * @param {Array} [subData]
  */
  function cleanupGoogleSubtaskOnSyncSkip_(subSheet, row, subData) {
    if (!subSheet || row < 2) {
      return;
    }

    if (!subData) {
      const lastCol = Math.max(subSheet.getLastColumn(), CONFIG.SUB_GOOGLE_TASK_ID_COL);
      subData = subSheet.getRange(row, 1, 1, lastCol).getValues()[0];
    }

    const subTaskName = (subData[CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();
    const linkedMainName = (subData[CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
    const titleDuplicateIndex = buildSubTaskTitleDuplicateIndex_(subSheet.getDataRange().getValues());
    const titles = getSubGoogleTaskTitlesForRow_(subTaskName, row, linkedMainName, titleDuplicateIndex);

    removeGoogleSubtaskForRowIfExists_(subSheet, row);
    if (titles.baseTitle) {
      removeOrphanTopLevelTasksByTitleAcrossAllStatusLists_(titles.baseTitle, '', titles.googleTitle);
    }
  }

  /** Cached map: normalized list name -> Google Task list ID (per execution). */
  let googleTaskListCache_ = null;
  /** Per-execution taskId → listId map; reset at bulk sync start; entries validated on read. */
  let taskIdToListIdCache_ = null;

  /** Cached 1-based column index for main Google Task ID (per execution). */
  let projectGoogleTaskIdColCache_ = null;

  /** Cached 1-based column index for sub Google Subtask ID (per execution). */
  let projectGoogleSubtaskIdColCache_ = null;
  let projectGoogleSubParentIdColCache_ = null;
  let projectGoogleSubNestStatusColCache_ = null;
  /** Cached 1-based column index for Gmail Thread ID on Dev Tracker (per execution). */
  let projectThreadIdColCache_ = null;

  /**
  * Returns true when the Google Tasks advanced service is wired in (Services +).
  * This is NOT the same as OAuth permission — see authorizeGoogleTasksAccess().
  * @returns {boolean}
  */
  function isGoogleTasksApiEnabled_() {
    try {
      if (typeof Tasks === 'undefined' || !Tasks.Tasklists) {
        return false;
      }
      return true;
    } catch (error) {
      Logger.log('isGoogleTasksApiEnabled_ Error: ' + error.toString());
      return false;
    }
  }

  /**
  * Logs why OAuth did not appear yet, or calls the Tasks API to trigger Allow popup.
  * Called automatically by setupAllTrackersComplete() / setupAllTrackers().
  * You can also run this alone for troubleshooting.
  * @returns {boolean}
  */
  function authorizeGoogleTasksAccess() {
    if (!isGoogleTasksApiEnabled_()) {
      Logger.log('Google Tasks service missing — add Services (+) > Google Tasks API, then run setup again.');
      Logger.log('After adding the service, setupAllTrackersComplete() will show the Allow popup automatically.');
      return false;
    }

    try {
      const response = Tasks.Tasklists.list({ maxResults: 5 });
      const items = response.items || [];
      Logger.log('Google Tasks API reachable. Lists visible: ' + items.length);
      for (let i = 0; i < items.length; i++) {
        Logger.log('  • ' + items[i].title);
      }
      return true;
    } catch (error) {
      Logger.log('authorizeGoogleTasksAccess Error: ' + error.toString());
      Logger.log(
        'Click Allow when the permission popup appears, then run setupAllTrackersComplete() again. ' +
        'If "Access Not Configured", open Project Settings → Google Cloud Platform → enable Tasks API.'
      );
      return false;
    }
  }

  /**
  * Normalize list/status names so en-dash vs hyphen mismatches still match.
  * @param {string} name
  * @returns {string}
  */
  function normalizeGoogleTaskListName_(name) {
    return (name || '')
      .toString()
      .trim()
      .replace(/\u2013/g, '-')
      .replace(/\s+/g, ' ');
  }

  /**
  * Loads all Google Task lists and caches them by normalized name.
  * Cache is per execution only (Apps Script globals reset between trigger runs).
  * @returns {Object<string, string>}
  */
  function refreshGoogleTaskListCache_() {
    const map = {};
    if (!isGoogleTasksApiEnabled_()) {
      return map;
    }

    let pageToken = null;
    do {
      const response = pageToken
        ? Tasks.Tasklists.list({ maxResults: 100, pageToken: pageToken })
        : Tasks.Tasklists.list({ maxResults: 100 });

      const items = response.items || [];
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (item && item.title && item.id) {
          map[normalizeGoogleTaskListName_(item.title)] = item.id;
        }
      }
      pageToken = response.nextPageToken || null;
    } while (pageToken);

    googleTaskListCache_ = map;
    return map;
  }

  /**
  * @returns {Object<string, string>}
  */
  function getGoogleTaskListMap_() {
    if (!googleTaskListCache_) {
      return refreshGoogleTaskListCache_();
    }
    return googleTaskListCache_;
  }

  /**
  * Resolves a Google Task list ID from a Dev Tracker status value.
  * @param {string} status
  * @returns {string|null}
  */
  function getTaskListIdForStatus_(status) {
    const statusName = (status || '').toString().trim();
    if (!statusName) {
      return null;
    }

    const map = getGoogleTaskListMap_();
    const normalized = normalizeGoogleTaskListName_(statusName);
    if (map[normalized]) {
      return map[normalized];
    }

    for (let i = 0; i < CONFIG.STATUS_LABELS.length; i++) {
      const label = CONFIG.STATUS_LABELS[i];
      if (normalizeGoogleTaskListName_(label) === normalized) {
        return map[normalizeGoogleTaskListName_(label)] || null;
      }
    }

    return null;
  }

  /**
  * SETUP HELPER: Verify Google Task lists exist for every CONFIG.STATUS_LABELS entry.
  */
  function verifyGoogleTaskLists() {
    try {
      if (!isGoogleTasksApiEnabled_()) {
        Logger.log('Google Tasks API not wired in. Add Services (+) > Google Tasks API, then run setupAllTrackersComplete().');
        return;
      }

      const map = refreshGoogleTaskListCache_();
      let missing = 0;

      CONFIG.STATUS_LABELS.forEach(function (statusName) {
        const key = normalizeGoogleTaskListName_(statusName);
        if (map[key]) {
          Logger.log('OK  "' + statusName + '" -> ' + map[key]);
        } else {
          missing++;
          Logger.log('MISSING Google Task list: "' + statusName + '"');
        }
      });

      if (missing === 0) {
        Logger.log('All ' + CONFIG.STATUS_LABELS.length + ' Google Task lists found.');
      } else {
        Logger.log(missing + ' list(s) missing. Create them in Google Tasks with exact names.');
      }
    } catch (error) {
      Logger.log('verifyGoogleTaskLists Error: ' + error.toString());
    }
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
  * @param {number} defaultCol
  * @param {string} headerText
  * @param {string[]} headerMatchers
  * @returns {number}
  */
  function findGoogleTaskIdColumn_(sheet, defaultCol, headerText, headerMatchers) {
    if (!sheet) {
      return defaultCol;
    }

    const lastCol = Math.max(sheet.getLastColumn(), defaultCol);
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    for (let c = 0; c < headers.length; c++) {
      const h = (headers[c] || '').toString().toLowerCase().trim();
      for (let m = 0; m < headerMatchers.length; m++) {
        if (h.indexOf(headerMatchers[m]) >= 0) {
          return c + 1;
        }
      }
    }

    if (!headers[defaultCol - 1] || headers[defaultCol - 1].toString().trim() === '') {
      sheet.getRange(1, defaultCol).setValue(headerText);
    }

    return defaultCol;
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
  * @returns {number}
  */
  function findMainGoogleTaskIdColumn_(mainSheet) {
    ensureColumnCachesFresh_(mainSheet, null);
    if (projectGoogleTaskIdColCache_) {
      return projectGoogleTaskIdColCache_;
    }

    projectGoogleTaskIdColCache_ = findGoogleTaskIdColumn_(
      mainSheet,
      CONFIG.MAIN_GOOGLE_TASK_ID_COL,
      'Google Task ID',
      ['google task id', 'google task']
    );
    return projectGoogleTaskIdColCache_;
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @returns {number}
  */
  function findSubGoogleTaskIdColumn_(subSheet) {
    ensureColumnCachesFresh_(null, subSheet);
    if (projectGoogleSubtaskIdColCache_) {
      return projectGoogleSubtaskIdColCache_;
    }

    projectGoogleSubtaskIdColCache_ = findGoogleTaskIdColumn_(
      subSheet,
      CONFIG.SUB_GOOGLE_TASK_ID_COL,
      'Google Subtask ID',
      ['google subtask id', 'google subtask', 'google task id']
    );
    return projectGoogleSubtaskIdColCache_;
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @returns {number}
  */
  function findSubParentGoogleTaskIdColumn_(subSheet) {
    ensureColumnCachesFresh_(null, subSheet);
    if (projectGoogleSubParentIdColCache_) {
      return projectGoogleSubParentIdColCache_;
    }

    projectGoogleSubParentIdColCache_ = findGoogleTaskIdColumn_(
      subSheet,
      CONFIG.SUB_PARENT_GOOGLE_TASK_ID_COL,
      'Main Google Task ID',
      ['main google task id', 'parent google task id']
    );
    return projectGoogleSubParentIdColCache_;
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  * @returns {string}
  */
  function getSubParentGoogleTaskIdForSubRow_(subSheet, row) {
    const col = findSubParentGoogleTaskIdColumn_(subSheet);
    return (subSheet.getRange(row, col).getValue() || '').toString().trim();
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  * @param {string} taskId
  */
  function setSubParentGoogleTaskIdForSubRow_(subSheet, row, taskId) {
    const col = findSubParentGoogleTaskIdColumn_(subSheet);
    subSheet.getRange(row, col).setValue(taskId);
  }

  /**
  * @param {Array} subData
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @returns {string}
  */
  function getSubParentGoogleTaskIdFromRowData_(subData, subSheet) {
    if (!subData || !subSheet) {
      return '';
    }
    const col = findSubParentGoogleTaskIdColumn_(subSheet);
    return (subData[col - 1] || '').toString().trim();
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @returns {number}
  */
  function findSubNestStatusColumn_(subSheet) {
    ensureColumnCachesFresh_(null, subSheet);
    if (projectGoogleSubNestStatusColCache_) {
      return projectGoogleSubNestStatusColCache_;
    }

    projectGoogleSubNestStatusColCache_ = findGoogleTaskIdColumn_(
      subSheet,
      CONFIG.SUB_NEST_STATUS_COL,
      'Google Nest Status',
      ['google nest status', 'nest status', 'subtask nest']
    );
    return projectGoogleSubNestStatusColCache_;
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  * @param {string} status
  */
  function setSubNestStatusForRow_(subSheet, row, status) {
    if (!subSheet || row < 2) {
      return;
    }
    const col = findSubNestStatusColumn_(subSheet);
    subSheet.getRange(row, col).setValue(status || '');
  }

  /**
  * Expected parent for nest audit: column W first, else column B → Dev Tracker column V.
  * @param {Array} subData
  * @param {string} linkedName
  * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
  * @param {Object} mainLookup
  * @param {Array} mainData
  * @returns {string}
  */
  function getNestExpectedParentId_(subData, linkedName, mainSheet, mainLookup, mainData, subSheet) {
    const w = getSubParentGoogleTaskIdFromRowData_(subData, subSheet);
    if (w) {
      return w;
    }
    return getExpectedParentGoogleTaskIdForSubRow_(linkedName, mainSheet, mainLookup, mainData);
  }

  /**
  * @param {string} subTaskId
  * @param {string} [hintListId]
  * @returns {{found: boolean, listId: string, parentId: string}}
  */
  function fetchGoogleTaskNestInfo_(subTaskId, hintListId) {
    if (!subTaskId) {
      return { found: false, listId: '', parentId: '' };
    }

    const listId = hintListId || findTaskListIdContainingTask_(subTaskId);
    if (!listId) {
      return { found: false, listId: '', parentId: '' };
    }

    try {
      const task = Tasks.Tasks.get(listId, subTaskId);
      return {
        found: true,
        listId: listId,
        parentId: (task.parent || '').toString().trim()
      };
    } catch (error) {
      return { found: false, listId: listId, parentId: '' };
    }
  }

  /**
  * Computes column X label: is sub column V nested under column W (or B→main V)?
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  * @param {Array} [subData]
  * @param {Object} [mainLookup]
  * @param {Array} [mainData]
  * @returns {string}
  */
  function computeSubNestStatusLabel_(subSheet, row, subData, mainLookup, mainData) {
    if (!subSheet || row < 2) {
      return '';
    }

    const lastCol = Math.max(
      subSheet.getLastColumn(),
      CONFIG.SUB_NEST_STATUS_COL,
      CONFIG.SUB_PARENT_GOOGLE_TASK_ID_COL,
      CONFIG.SUB_GOOGLE_TASK_ID_COL
    );
    if (!subData) {
      subData = subSheet.getRange(row, 1, 1, lastCol).getValues()[0];
    }

    if (isSubSheetDataRowEmpty_(subData)) {
      return '';
    }

    const linkedName = (subData[CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
    const subTaskName = (subData[CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();
    const subTaskId = (subData[CONFIG.SUB_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
    const parentW = getSubParentGoogleTaskIdFromRowData_(subData, subSheet);
    const subStatus = (subData[CONFIG.SUB_STATUS_COL - 1] || '').toString().trim();
    const subEnv = (subData[CONFIG.SUB_ENV_COL - 1] || '').toString().trim();

    if (!subTaskId) {
      return 'NO SUB ID';
    }

    if (subTaskName && shouldSkipSubtaskGoogleSync_(subStatus, subEnv)) {
      return 'SKIP';
    }

    const mainSheet = getMainTrackerSheet_();
    if (!mainLookup && mainSheet) {
      mainData = mainSheet.getDataRange().getValues();
      mainLookup = buildMainTaskNameIndex_(mainData);
    }

    const expectedParentId = getNestExpectedParentId_(subData, linkedName, mainSheet, mainLookup, mainData, subSheet);
    if (!expectedParentId) {
      return 'NO PARENT ID';
    }

    const mainVExpected = linkedName
      ? getExpectedParentGoogleTaskIdForSubRow_(linkedName, mainSheet, mainLookup, mainData)
      : '';
    const wMismatch = parentW && mainVExpected && parentW !== mainVExpected;

    const nestInfo = fetchGoogleTaskNestInfo_(subTaskId);
    if (!nestInfo.found) {
      return wMismatch ? 'W≠MAIN V | NOT IN GOOGLE' : 'NOT IN GOOGLE';
    }

    let label;
    if (!nestInfo.parentId) {
      label = 'ORPHAN (top-level)';
    } else if (nestInfo.parentId === expectedParentId) {
      label = 'OK';
    } else {
      label = 'WRONG PARENT: ' + nestInfo.parentId;
    }

    if (wMismatch) {
      label = label === 'OK' ? 'OK (W≠MAIN V)' : 'W≠MAIN V | ' + label;
    }

    return label;
  }

  /**
  * Updates column X for one Dev Tracker Sub row (1 Google Tasks API read when column V is set).
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  * @param {Array} [subData]
  * @param {Object} [mainLookup]
  * @param {Array} [mainData]
  * @returns {string}
  */
  function refreshSubNestStatusForRow_(subSheet, row, subData, mainLookup, mainData) {
    const label = computeSubNestStatusLabel_(subSheet, row, subData, mainLookup, mainData);
    setSubNestStatusForRow_(subSheet, row, label);
    return label;
  }

  /**
  * Fills column X on Dev Tracker Sub for all rows with sub data.
  * Checks whether column V is nested under column W in Google Tasks.
  *
  * Run from Apps Script (open spreadsheet first):
  *   refreshSubNestStatusColumn()
  *
  * Options: { batchSize: 300, maxRuntimeMs: 270000, startRow: 0 }
  * Re-run if log says paused — processes next batch until done.
  *
  * @param {Object} [options]
  * @returns {{ok: number, orphan: number, wrong: number, other: number, processed: number, paused: boolean}}
  */
  function refreshSubNestStatusColumn(options) {
    options = options || {};
    if (!isGoogleTasksApiEnabled_()) {
      Logger.log('Google Tasks API not enabled — add Services (+) > Google Tasks API first.');
      return { ok: 0, orphan: 0, wrong: 0, other: 0, processed: 0, paused: false };
    }

    const subSheet = getSubTrackerSheet_();
    const mainSheet = getMainTrackerSheet_();
    if (!subSheet || !mainSheet) {
      Logger.log('Sheets not found');
      return { ok: 0, orphan: 0, wrong: 0, other: 0, processed: 0, paused: false };
    }

    refreshGoogleTaskListCache_();
    const mainData = mainSheet.getDataRange().getValues();
    const mainLookup = buildMainTaskNameIndex_(mainData);
    const subData = subSheet.getDataRange().getValues();
    const subLast = CONFIG.GOOGLE_TASKS_SKIP_EMPTY_ROWS !== false
      ? getSubLastDataRow_(subData)
      : subSheet.getLastRow();
    const rowList = getSubRowNumbersWithDataBottomUp_(subData, subLast);

    const batchSize = options.batchSize || CONFIG.GOOGLE_TASKS_SYNC_BATCH_SIZE || 300;
    const maxRuntime = options.maxRuntimeMs || CONFIG.GOOGLE_TASKS_SYNC_MAX_RUNTIME_MS || 270000;
    const startRow = options.startRow || 0;
    const startTime = Date.now();
    const stats = { ok: 0, orphan: 0, wrong: 0, other: 0, processed: 0, paused: false };
    const sleepMs = CONFIG.GOOGLE_TASKS_SYNC_SLEEP_MS || 50;

    beginGoogleTasksSync_('BULK', 'refresh column X nest status (rows with sub data)');

    for (let i = startRow; i < rowList.length; i++) {
      if (stats.processed >= batchSize || Date.now() - startTime >= maxRuntime) {
        stats.paused = i < rowList.length;
        logGoogleTasksSync_(
          'Nest status batch paused at index ' + i + '/' + rowList.length +
          ' — run refreshSubNestStatusColumn({ startRow: ' + i + ' }) to continue'
        );
        break;
      }

      const row = rowList[i];
      const label = refreshSubNestStatusForRow_(
        subSheet,
        row,
        subData[row - 1],
        mainLookup,
        mainData
      );

      if (label === 'OK' || label.indexOf('OK') === 0) {
        stats.ok++;
      } else if (label.indexOf('ORPHAN') >= 0) {
        stats.orphan++;
      } else if (label.indexOf('WRONG PARENT') >= 0) {
        stats.wrong++;
      } else if (label) {
        stats.other++;
      }

      stats.processed++;
      if (sleepMs > 0) {
        Utilities.sleep(sleepMs);
      }
    }

    endGoogleTasksSync_(
      'column X: processed ' + stats.processed + ', OK=' + stats.ok +
      ', ORPHAN=' + stats.orphan + ', WRONG=' + stats.wrong + ', other=' + stats.other +
      (stats.paused ? ' (paused — run again to continue)' : ' (complete)')
    );
    return stats;
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
  * @param {number} row
  * @returns {string}
  */
  function getGoogleTaskIdForMainRow_(mainSheet, row) {
    const col = findMainGoogleTaskIdColumn_(mainSheet);
    return (mainSheet.getRange(row, col).getValue() || '').toString().trim();
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
  * @param {number} row
  * @param {string} taskId
  */
  function setGoogleTaskIdForMainRow_(mainSheet, row, taskId) {
    const col = findMainGoogleTaskIdColumn_(mainSheet);
    mainSheet.getRange(row, col).setValue(taskId);
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  * @returns {string}
  */
  function getGoogleSubtaskIdForSubRow_(subSheet, row) {
    const col = findSubGoogleTaskIdColumn_(subSheet);
    return (subSheet.getRange(row, col).getValue() || '').toString().trim();
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  * @param {string} taskId
  */
  function setGoogleSubtaskIdForSubRow_(subSheet, row, taskId) {
    const col = findSubGoogleTaskIdColumn_(subSheet);
    subSheet.getRange(row, col).setValue(taskId);
  }

  /**
  * @param {string} taskId
  * @param {string} [hintListId] Try this list first (e.g. parent list from sheet status).
  * @returns {string|null}
  */
  function findTaskListIdContainingTask_(taskId, hintListId) {
    if (!taskId) {
      return null;
    }

    if (!taskIdToListIdCache_) {
      taskIdToListIdCache_ = {};
    }
    if (taskIdToListIdCache_[taskId]) {
      const cachedListId = taskIdToListIdCache_[taskId];
      try {
        Tasks.Tasks.get(cachedListId, taskId);
        return cachedListId;
      } catch (error) {
        delete taskIdToListIdCache_[taskId];
      }
    }

    if (hintListId) {
      try {
        Tasks.Tasks.get(hintListId, taskId);
        taskIdToListIdCache_[taskId] = hintListId;
        return hintListId;
      } catch (error) {
        // Not in hinted list — fall through to full search.
      }
    }

    const map = getGoogleTaskListMap_();
    const listIds = {};
    Object.keys(map).forEach(function (key) {
      listIds[map[key]] = true;
    });

    for (const listId in listIds) {
      if (!Object.prototype.hasOwnProperty.call(listIds, listId)) {
        continue;
      }
      try {
        Tasks.Tasks.get(listId, taskId);
        taskIdToListIdCache_[taskId] = listId;
        return listId;
      } catch (error) {
        // Task is not in this list.
      }
    }

    return null;
  }

  /**
  * @param {string} taskId
  * @param {string} listId
  */
  function cacheTaskListIdForTask_(taskId, listId) {
    if (!taskId || !listId) {
      return;
    }
    if (!taskIdToListIdCache_) {
      taskIdToListIdCache_ = {};
    }
    taskIdToListIdCache_[taskId] = listId;
  }

  /**
  * Drops a stale task→list mapping (after move, delete, or 404).
  * @param {string} taskId
  */
  function invalidateTaskListIdCacheForTask_(taskId) {
    if (!taskId || !taskIdToListIdCache_) {
      return;
    }
    delete taskIdToListIdCache_[taskId];
  }

  /**
  * @param {Array} mainData
  * @returns {Object<string, {row: number, data: Array}>}
  */
  function buildMainTaskNameIndex_(mainData) {
    const index = {};
    if (!mainData) {
      return index;
    }
    for (let i = 1; i < mainData.length; i++) {
      const name = (mainData[i][CONFIG.MAIN_TASK_NAME_COL - 1] || '').toString().trim();
      if (name) {
        index[name] = { row: i + 1, data: mainData[i] };
      }
    }
    return index;
  }

  /**
  * Dev Tracker column V → main row. Used to ensure main task IDs are never treated as subtask IDs.
  * @param {Array} mainData
  * @returns {Object<string, {row: number, data: Array}>}
  */
  function buildMainGoogleTaskIdIndex_(mainData) {
    const index = {};
    if (!mainData) {
      return index;
    }
    for (let i = 1; i < mainData.length; i++) {
      const id = (mainData[i][CONFIG.MAIN_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
      if (id && !index[id]) {
        index[id] = { row: i + 1, data: mainData[i] };
      }
    }
    return index;
  }

  /**
  * Sub column V must not hold a Dev Tracker main Google Task ID — mains are always parents, never subs.
  * Clears the cell only (does not delete the main Google Task).
  * @returns {boolean}
  */
  function clearSubColumnVIfDevTrackerMainTaskId_(subSheet, row, mainIdIndex) {
    if (!subSheet || row < 2 || !mainIdIndex) {
      return false;
    }

    const subId = getGoogleSubtaskIdForSubRow_(subSheet, row);
    if (!subId || !mainIdIndex[subId]) {
      return false;
    }

    const mainHit = mainIdIndex[subId];
    logGoogleTasksSync_(
      'Sub row ' + row + ': column V has Dev Tracker MAIN task id ' + subId +
      ' (main row ' + mainHit.row + ') — cleared; a new subtask id will be created on sync'
    );
    setGoogleSubtaskIdForSubRow_(subSheet, row, '');
    return true;
  }

  /**
  * Maps main task name (Dev Tracker col C) → linked sub rows.
  * One sub sheet read; used instead of scanning every sub row on main Status edit.
  * @param {Array} subData Full sub sheet including header row
  * @returns {Object<string, Array<{sheetRow: number, data: Array}>>}
  */
  function buildSubTaskIndexByMainName_(subData) {
    const index = {};
    if (!subData) {
      return index;
    }
    for (let i = 1; i < subData.length; i++) {
      const link = (subData[i][CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
      if (!link) {
        continue;
      }
      if (!index[link]) {
        index[link] = [];
      }
      index[link].push({ sheetRow: i + 1, data: subData[i] });
    }
    return index;
  }

  /**
  * Dev Tracker Sub: newest rows are at the bottom — bulk sync, audit, and repair scan bottom → row 2.
  * @param {number} subLast Last data row on Dev Tracker Sub
  * @param {number} [resumeRow] 0 or omitted = start at bottom; saved row when resuming a paused bulk run
  * @returns {number}
  */
  function getSubSheetScanStartRow_(subLast, resumeRow) {
    if (subLast < 2) {
      return 2;
    }
    if (resumeRow && resumeRow >= 2 && resumeRow <= subLast) {
      return resumeRow;
    }
    return subLast;
  }

  /**
  * @param {number} subLast
  * @param {number} startRow
  * @param {function(number): void} fn Sheet row numbers from startRow down to 2
  */
  function eachSubSheetRowBottomUp_(subLast, startRow, fn) {
    if (subLast < 2 || typeof fn !== 'function') {
      return;
    }
    const start = getSubSheetScanStartRow_(subLast, startRow);
    for (let r = start; r >= 2; r--) {
      fn(r);
    }
  }

  /**
  * Sort linked sub entries so higher sheet rows (latest at bottom) run first.
  * @param {Array<{sheetRow: number}>} entries
  */
  function sortSubTaskEntriesBottomUp_(entries) {
    entries.sort(function (a, b) {
      return b.sheetRow - a.sheetRow;
    });
    return entries;
  }

  /**
  * True when a sub row has no link, name, or Google Task IDs — safe to skip in bulk sync.
  * @param {Array} subData One row from sub sheet getDataRange()
  * @returns {boolean}
  */
  function isSubSheetDataRowEmpty_(subData) {
    if (!subData) {
      return true;
    }
    const link = (subData[CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
    const name = (subData[CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();
    const subId = (subData[CONFIG.SUB_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
    const parentW = (subData[CONFIG.SUB_PARENT_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
    return !link && !name && !subId && !parentW;
  }

  /**
  * Sheet row numbers (2-based) where Dev Tracker column C (task name) is not empty.
  * Bulk sync uses this list so empty rows in the middle are never scanned.
  * @param {Array} mainData
  * @returns {number[]}
  */
  function getMainRowNumbersWithTaskNames_(mainData) {
    const rows = [];
    if (!mainData) {
      return rows;
    }
    const col = CONFIG.MAIN_TASK_NAME_COL - 1;
    for (let i = 1; i < mainData.length; i++) {
      if ((mainData[i][col] || '').toString().trim() !== '') {
        rows.push(i + 1);
      }
    }
    return rows;
  }

  /**
  * Sheet row numbers for sub rows with data (B/D/V/W), highest row first (bottom-up).
  * @param {Array} subData
  * @param {number} [maxRow] Cap scan at this row (from getSubLastDataRow_)
  * @returns {number[]}
  */
  function getSubRowNumbersWithDataBottomUp_(subData, maxRow) {
    const rows = [];
    if (!subData || subData.length <= 1) {
      return rows;
    }
    const lastIndex = subData.length - 1;
    const startIndex =
      maxRow && maxRow >= 2 ? Math.min(maxRow, subData.length) - 1 : lastIndex;
    for (let i = Math.min(startIndex, lastIndex); i >= 1; i--) {
      if (!isSubSheetDataRowEmpty_(subData[i])) {
        rows.push(i + 1);
      }
    }
    return rows;
  }

  /**
  * Last sheet row with a task name in column C (or 1 if none). Faster than getLastRow() when blank rows exist below data.
  * @param {Array} mainData From getDataRange() including header row
  * @returns {number}
  */
  function getMainLastDataRow_(mainData) {
    if (!mainData || mainData.length <= 1) {
      return 1;
    }
    const col = CONFIG.MAIN_TASK_NAME_COL - 1;
    for (let i = mainData.length - 1; i >= 1; i--) {
      if ((mainData[i][col] || '').toString().trim() !== '') {
        return i + 1;
      }
    }
    return 1;
  }

  /**
  * Last sheet row with sub task data (B, D, V, or W). Ignores trailing blank rows.
  * @param {Array} subData From getDataRange() including header row
  * @returns {number}
  */
  function getSubLastDataRow_(subData) {
    if (!subData || subData.length <= 1) {
      return 1;
    }
    for (let i = subData.length - 1; i >= 1; i--) {
      if (!isSubSheetDataRowEmpty_(subData[i])) {
        return i + 1;
      }
    }
    return 1;
  }

  /**
  * @param {Array} data
  * @param {number} row
  * @returns {string}
  */
  function buildMainGoogleTaskNotes_(data, row) {
    const lines = [
      'Dev Tracker row: ' + row,
      'Status: ' + ((data[CONFIG.MAIN_STATUS_COL - 1] || 'N/A').toString().trim() || 'N/A'),
      'Environment: ' + ((data[CONFIG.MAIN_ENV_COL - 1] || 'N/A').toString().trim() || 'N/A'),
      'Informed: ' + ((data[CONFIG.MAIN_INFORMED_COL - 1] || 'N/A').toString().trim() || 'N/A'),
      'Support Needed / Blocker: ' + ((data[CONFIG.MAIN_BLOCKER_COL - 1] || 'None').toString().trim() || 'None'),
      'Dependency: ' + ((data[CONFIG.MAIN_DEPENDENCY_COL - 1] || 'None').toString().trim() || 'None')
    ];
    return lines.join('\n');
  }

  /**
  * @param {Array} data
  * @param {number} row
  * @param {string} mainTaskName
  * @returns {string}
  */
  function buildSubGoogleTaskNotes_(data, row, mainTaskName) {
    const lines = [
      'Dev Tracker Sub row: ' + row,
      'Sheet task name (column D): ' + ((data[CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim() || '(empty)'),
      'Main task: ' + mainTaskName,
      'Status: ' + ((data[CONFIG.SUB_STATUS_COL - 1] || 'N/A').toString().trim() || 'N/A'),
      'Assigned: ' + ((data[CONFIG.SUB_ASSIGNED_COL - 1] || 'N/A').toString().trim() || 'N/A'),
      'Environment: ' + ((data[CONFIG.SUB_ENV_COL - 1] || 'N/A').toString().trim() || 'N/A'),
      'Informed: ' + ((data[CONFIG.SUB_INFORMED_COL - 1] || 'N/A').toString().trim() || 'N/A'),
      'Support Needed / Blocker: ' + ((data[CONFIG.SUB_BLOCKER_COL - 1] || 'None').toString().trim() || 'None'),
      'Dependency: ' + ((data[CONFIG.SUB_DEPENDENCY_COL - 1] || 'None').toString().trim() || 'None')
    ];
    return lines.join('\n');
  }

  /**
  * Counts how many sub rows share the same column D task name (case-insensitive).
  * @param {Array} subData Full sub sheet data including header row
  * @returns {Object<string, number>}
  */
  function buildSubTaskTitleDuplicateIndex_(subData) {
    const counts = {};
    if (!subData || subData.length <= 1) {
      return counts;
    }

    for (let i = 1; i < subData.length; i++) {
      const name = (subData[i][CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();
      if (!name) {
        continue;
      }
      const key = name.toLowerCase();
      counts[key] = (counts[key] || 0) + 1;
    }
    return counts;
  }

  /**
  * @param {string} text
  * @param {number} maxLen
  * @returns {string}
  */
  function shortenTextForGoogleTaskTitle_(text, maxLen) {
    const s = (text || '').toString().trim();
    if (!s || s.length <= maxLen) {
      return s;
    }
    return s.substring(0, Math.max(0, maxLen - 1)).trim() + '…';
  }

  /**
  * True when this sub row should get a disambiguating Google Tasks title suffix.
  * @param {string} subTaskName
  * @param {Object<string, number>} titleDuplicateIndex
  * @returns {boolean}
  */
  function shouldSuffixGoogleSubtaskTitle_(subTaskName, titleDuplicateIndex) {
    const mode = (CONFIG.GOOGLE_TASKS_SUBTASK_TITLE_SUFFIX || 'duplicate').toString().toLowerCase();
    if (mode === 'always') {
      return true;
    }
    const key = (subTaskName || '').toString().trim().toLowerCase();
    return !!(titleDuplicateIndex && key && titleDuplicateIndex[key] > 1);
  }

  /**
  * Google Tasks title for a sub row. Sheet column D stays the plain name; suffix avoids duplicate
  * top-level orphans that look like main tasks when nesting fails.
  * @param {string} subTaskName Column D
  * @param {number} row Sub sheet row
  * @param {string} linkedMainName Column B
  * @param {Object<string, number>} [titleDuplicateIndex]
  * @returns {string}
  */
  function formatGoogleSubtaskTitleForGoogle_(subTaskName, row, linkedMainName, titleDuplicateIndex) {
    const base = (subTaskName || '').toString().trim();
    if (!base) {
      return base;
    }
    if (!shouldSuffixGoogleSubtaskTitle_(base, titleDuplicateIndex)) {
      return base;
    }

    const mainShort = shortenTextForGoogleTaskTitle_(linkedMainName, 36);
    const suffix = mainShort ? ' [' + mainShort + ' · #' + row + ']' : ' [Sub #' + row + ']';
    const maxLen = 200;
    if (base.length + suffix.length <= maxLen) {
      return base + suffix;
    }
    return shortenTextForGoogleTaskTitle_(base, maxLen - suffix.length) + suffix;
  }

  /**
  * @param {string} subTaskName
  * @param {number} row
  * @param {string} linkedMainName
  * @param {Object<string, number>} [titleDuplicateIndex]
  * @returns {{baseTitle: string, googleTitle: string}}
  */
  function getSubGoogleTaskTitlesForRow_(subTaskName, row, linkedMainName, titleDuplicateIndex) {
    const baseTitle = (subTaskName || '').toString().trim();
    return {
      baseTitle: baseTitle,
      googleTitle: formatGoogleSubtaskTitleForGoogle_(baseTitle, row, linkedMainName, titleDuplicateIndex)
    };
  }

  /**
  * True when a top-level Google Task title matches a sub row's base or suffixed title.
  * @param {string} taskTitle
  * @param {string} baseTitle
  * @param {string} [googleTitle]
  * @returns {boolean}
  */
  function isOrphanTopLevelTitleMatchForSub_(taskTitle, baseTitle, googleTitle) {
    const t = (taskTitle || '').toString().trim();
    const base = (baseTitle || '').toString().trim();
    const google = (googleTitle || '').toString().trim();
    if (!t || !base) {
      return false;
    }
    if (t === base || (google && t === google)) {
      return true;
    }
    return t.indexOf(base + ' [') === 0;
  }

  /**
  * Builds the Google Tasks taskResource for one sub row (unique title + notes).
  * @param {Array} subData
  * @param {number} row
  * @param {string} linkedMainName
  * @param {Object<string, number>} [titleDuplicateIndex]
  * @returns {{title: string, notes: string, baseTitle: string, googleTitle: string}}
  */
  function buildSubGoogleTaskResource_(subData, row, linkedMainName, titleDuplicateIndex) {
    const subTaskName = (subData[CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();
    const titles = getSubGoogleTaskTitlesForRow_(subTaskName, row, linkedMainName, titleDuplicateIndex);
    return {
      title: titles.googleTitle,
      notes: buildSubGoogleTaskNotes_(subData, row, linkedMainName),
      baseTitle: titles.baseTitle,
      googleTitle: titles.googleTitle
    };
  }

  /**
  * @param {string} taskListId
  * @param {string} taskId
  * @param {Object} taskResource
  */
  function patchGoogleTask_(taskListId, taskId, taskResource) {
    const patch = {};
    if (taskResource && taskResource.title !== undefined) {
      patch.title = taskResource.title;
    }
    if (taskResource && taskResource.notes !== undefined) {
      patch.notes = taskResource.notes;
    }
    if (taskResource && taskResource.status !== undefined) {
      patch.status = taskResource.status;
    }
    if (Object.keys(patch).length === 0) {
      return;
    }
    // Apps Script signature: patch(resource, tasklist, task) — not (tasklist, task, resource)
    Tasks.Tasks.patch(patch, taskListId, taskId);
  }

  /**
  * Reopens a parent in Google Tasks only when the sheet says it is still active.
  * Never reopens completed/closed main tasks — that corrupts the user's Completed list.
  * @param {string} listId
  * @param {string} parentTaskId
  * @param {string} [parentSheetStatus]
  * @returns {boolean} False when nesting must not proceed.
  */
  function ensureGoogleParentTaskOpenForSubtasks_(listId, parentTaskId, parentSheetStatus) {
    if (isCompletedOrClosedStatus_(parentSheetStatus || '')) {
      Logger.log(
        'Parent Google Task ' + parentTaskId + ' is completed/closed on sheet — not reopening for subtasks.'
      );
      return false;
    }

    try {
      const parent = Tasks.Tasks.get(listId, parentTaskId);
      if (parent.status === 'completed') {
        patchGoogleTask_(listId, parentTaskId, { status: 'needsAction' });
        Logger.log('Reopened parent Google Task ' + parentTaskId + ' (sheet status is still open).');
      }
      return true;
    } catch (error) {
      Logger.log('ensureGoogleParentTaskOpenForSubtasks_ Error: ' + error.toString());
      return false;
    }
  }

  /**
  * @param {string} listId
  * @param {string} taskId
  * @param {string} parentTaskId
  * @param {string} [knownListId] List the task was just inserted into (skip index lag).
  * @returns {boolean}
  */
  function verifyGoogleTaskNestedUnderParent_(listId, taskId, parentTaskId, knownListId) {
    const actualListId = knownListId || findTaskListIdContainingTask_(taskId, listId) || listId;
    if (!actualListId || !taskId || !parentTaskId) {
      return false;
    }

    try {
      const task = Tasks.Tasks.get(actualListId, taskId);
      return task.parent === parentTaskId;
    } catch (error) {
      Logger.log('verifyGoogleTaskNestedUnderParent_ Error: ' + error.toString());
      return false;
    }
  }

  /**
  * Removes a task from whichever list currently contains it.
  * @param {string} taskId
  */
  function removeGoogleTaskFromAnyList_(taskId) {
    if (!taskId) {
      return;
    }

    const listId = findTaskListIdContainingTask_(taskId);
    if (!listId) {
      return;
    }

    try {
      Tasks.Tasks.remove(listId, taskId);
      invalidateTaskListIdCacheForTask_(taskId);
      Logger.log('Removed Google Task ' + taskId + ' from list ' + listId);
    } catch (error) {
      Logger.log('removeGoogleTaskFromAnyList_ Error: ' + error.toString());
    }
  }

  /**
  * Attempts move + verify. Uses destinationTasklist when the subtask is in a different list than its parent.
  * @param {string} parentListId
  * @param {string} taskId
  * @param {string} parentTaskId
  * @param {string} [knownListId] Use when the task was just inserted (avoids index-lag list lookup).
  * @param {string} [parentSheetStatus]
  * @returns {boolean}
  */
  function tryNestGoogleTaskUnderParent_(parentListId, taskId, parentTaskId, knownListId, parentSheetStatus) {
    const currentListId = knownListId || findTaskListIdContainingTask_(taskId) || parentListId;
    if (!currentListId) {
      Logger.log('tryNestGoogleTaskUnderParent_: task ' + taskId + ' not found in any list.');
      return false;
    }

    if (verifyGoogleTaskNestedUnderParent_(parentListId, taskId, parentTaskId, currentListId)) {
      cacheTaskListIdForTask_(taskId, parentListId);
      Logger.log('Subtask ' + taskId + ' already nested under parent ' + parentTaskId);
      return true;
    }

    if (!ensureGoogleParentTaskOpenForSubtasks_(parentListId, parentTaskId, parentSheetStatus)) {
      return false;
    }

    // Apps Script advanced service maps this object to REST query params (parent, previous, destinationTasklist).
    const moveOpts = { parent: parentTaskId };
    if (currentListId !== parentListId) {
      moveOpts.destinationTasklist = parentListId;
      Logger.log(
        'Moving subtask ' + taskId + ' from list ' + currentListId +
        ' to parent list ' + parentListId + ' under parent ' + parentTaskId
      );
    } else {
      Logger.log('Re-parenting subtask ' + taskId + ' under parent ' + parentTaskId + ' in list ' + currentListId);
    }

    try {
      Tasks.Tasks.move(currentListId, taskId, moveOpts);
      Utilities.sleep(500);
    } catch (error) {
      Logger.log('tryNestGoogleTaskUnderParent_ move Error: ' + error.toString());
      return false;
    }

    let nested = verifyGoogleTaskNestedUnderParent_(parentListId, taskId, parentTaskId, parentListId);
    if (nested) {
      cacheTaskListIdForTask_(taskId, parentListId);
      Logger.log('Nest via move verified for subtask ' + taskId + ' under parent ' + parentTaskId);
      return true;
    }

    Logger.log('Re-parent verification failed for task ' + taskId + ' — retrying move after 1s.');
    Utilities.sleep(1000);

    try {
      Tasks.Tasks.move(parentListId, taskId, { parent: parentTaskId });
      Utilities.sleep(500);
    } catch (retryError) {
      Logger.log('tryNestGoogleTaskUnderParent_ retry move Error: ' + retryError.toString());
      return false;
    }

    nested = verifyGoogleTaskNestedUnderParent_(parentListId, taskId, parentTaskId, parentListId);
    if (nested) {
      cacheTaskListIdForTask_(taskId, parentListId);
      Logger.log('Nest via move verified on retry for subtask ' + taskId + ' under parent ' + parentTaskId);
    } else {
      Logger.log(
        'Re-parent verification failed for task ' + taskId +
        ' after retry — still not nested under parent ' + parentTaskId +
        ' (recreate fallback will run if applicable)'
      );
    }
    return nested;
  }

  /**
  * Forces a task to be a child of parentTaskId. Verifies after move; does not recreate.
  * @param {string} listId
  * @param {string} taskId
  * @param {string} parentTaskId
  * @returns {Object|null}
  */
  function nestGoogleTaskUnderParent_(listId, taskId, parentTaskId) {
    if (tryNestGoogleTaskUnderParent_(listId, taskId, parentTaskId)) {
      const actualListId = findTaskListIdContainingTask_(taskId) || listId;
      try {
        return Tasks.Tasks.get(actualListId, taskId);
      } catch (error) {
        Logger.log('nestGoogleTaskUnderParent_ get after nest Error: ' + error.toString());
      }
    }
    return null;
  }

  /**
  * Ensures a subtask exists nested under parent. Tries move+verify first; if that fails
  * (common for old top-level tasks), deletes the stale task and recreates as a child.
  * @param {Object} taskResource
  * @param {string} parentListId
  * @param {string} parentTaskId
  * @param {string} [existingSubTaskId]
  * @param {string} [parentSheetStatus]
  * @returns {{taskId: string, listId: string, recreated: boolean}|null}
  */
  function ensureGoogleSubtaskUnderParent_(taskResource, parentListId, parentTaskId, existingSubTaskId, parentSheetStatus) {
    if (isCompletedOrClosedStatus_(parentSheetStatus || '')) {
      Logger.log('ensureGoogleSubtaskUnderParent_: parent main task is completed/closed — skipping subtask nest.');
      return null;
    }

    if (existingSubTaskId) {
      const mainSheet = getMainTrackerSheet_();
      if (mainSheet) {
        const mainIdIndex = buildMainGoogleTaskIdIndex_(mainSheet.getDataRange().getValues());
        if (mainIdIndex[existingSubTaskId]) {
          Logger.log(
            'ensureGoogleSubtaskUnderParent_: id ' + existingSubTaskId +
            ' is a Dev Tracker MAIN task (row ' + mainIdIndex[existingSubTaskId].row + ') — will create new subtask instead'
          );
          existingSubTaskId = null;
        }
      }
    }

    if (!ensureGoogleParentTaskOpenForSubtasks_(parentListId, parentTaskId, parentSheetStatus)) {
      return null;
    }

    const subTitle = (taskResource.title || '').toString().trim();
    const baseTitle = (taskResource.baseTitle || subTitle).toString().trim();

    if (existingSubTaskId) {
      if (tryNestGoogleTaskUnderParent_(parentListId, existingSubTaskId, parentTaskId, null, parentSheetStatus)) {
        const listId = findTaskListIdContainingTask_(existingSubTaskId, parentListId) || parentListId;
        if (verifyGoogleTaskNestedUnderParent_(parentListId, existingSubTaskId, parentTaskId, listId)) {
          removeOrphanTopLevelDuplicateSubtasks_(
            parentListId,
            parentTaskId,
            baseTitle,
            subTitle,
            existingSubTaskId,
            false
          );
          return { taskId: existingSubTaskId, listId: listId, recreated: false };
        }
        Logger.log(
          'Subtask ' + existingSubTaskId + ' move reported success but is still top-level — deleting and recreating.'
        );
      } else {
        Logger.log(
          'Re-parent failed for subtask ' + existingSubTaskId +
          ' — deleting and recreating under parent ' + parentTaskId
        );
      }
      removeGoogleTaskFromAnyList_(existingSubTaskId);
    }

    removeOrphanTopLevelDuplicateSubtasks_(
      parentListId,
      parentTaskId,
      baseTitle,
      subTitle,
      '',
      true
    );

    const created = insertGoogleSubtaskUnderParent_(taskResource, parentListId, parentTaskId, parentSheetStatus);
    if (!created) {
      return null;
    }

    if (!verifyGoogleTaskNestedUnderParent_(parentListId, created.id, parentTaskId, parentListId)) {
      Logger.log(
        'Newly created subtask ' + created.id +
        ' is not nested under parent ' + parentTaskId + ' — deleting orphan.'
      );
      removeGoogleTaskFromAnyList_(created.id);
      return null;
    }

    return { taskId: created.id, listId: parentListId, recreated: true };
  }

  /**
  * Inserts a subtask and guarantees it is nested under the parent (not top-level).
  * @param {Object} taskResource
  * @param {string} parentListId
  * @param {string} parentTaskId
  * @param {string} [parentSheetStatus]
  * @returns {Object|null}
  */
  function insertGoogleSubtaskUnderParent_(taskResource, parentListId, parentTaskId, parentSheetStatus) {
    if (!ensureGoogleParentTaskOpenForSubtasks_(parentListId, parentTaskId, parentSheetStatus)) {
      return null;
    }

    const created = Tasks.Tasks.insert(
      { title: taskResource.title, notes: taskResource.notes },
      parentListId,
      { parent: parentTaskId }
    );

    if (created.parent === parentTaskId) {
      Logger.log('Insert returned nested subtask ' + created.id + ' under parent ' + parentTaskId);
      return created;
    }

    Logger.log(
      'Insert did not nest subtask ' + created.id + ' (parent=' + (created.parent || 'none') + ') — retrying move.'
    );

    if (tryNestGoogleTaskUnderParent_(parentListId, created.id, parentTaskId, parentListId, parentSheetStatus)) {
      if (verifyGoogleTaskNestedUnderParent_(parentListId, created.id, parentTaskId, parentListId)) {
        try {
          return Tasks.Tasks.get(parentListId, created.id);
        } catch (getErr) {
          Logger.log('insertGoogleSubtaskUnderParent_ get after nest Error: ' + getErr.toString());
          return created;
        }
      }
      Logger.log('insertGoogleSubtaskUnderParent_: nest retry reported success but verify failed.');
    }

    Utilities.sleep(1000);
    if (tryNestGoogleTaskUnderParent_(parentListId, created.id, parentTaskId, parentListId, parentSheetStatus)) {
      if (verifyGoogleTaskNestedUnderParent_(parentListId, created.id, parentTaskId, parentListId)) {
        try {
          return Tasks.Tasks.get(parentListId, created.id);
        } catch (getErr2) {
          Logger.log('insertGoogleSubtaskUnderParent_ get after retry Error: ' + getErr2.toString());
        }
      }
    }

    Logger.log(
      'FAILED to nest subtask ' + created.id + ' under parent ' + parentTaskId +
      ' after insert + retries — deleting top-level orphan.'
    );
    removeGoogleTaskFromAnyList_(created.id);
    return null;
  }

  /**
  * Ensures the main Google Task exists, is in the correct status list, and can accept subtasks.
  * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
  * @param {number} parentRow
  * @returns {{parentTaskId: string, parentListId: string}|null}
  */
  function prepareParentGoogleTaskForSubtask_(mainSheet, parentRow) {
    if (!mainSheet) {
      Logger.log('prepareParentGoogleTaskForSubtask_: main sheet is missing.');
      return null;
    }

    const parentRowNum = parseInt(parentRow, 10);
    if (!parentRowNum || parentRowNum < 2) {
      Logger.log('prepareParentGoogleTaskForSubtask_: invalid parent row "' + parentRow + '".');
      return null;
    }
    parentRow = parentRowNum;

    syncGoogleTaskForMainRow(mainSheet, parentRow);

    const parentTaskId = getGoogleTaskIdForMainRow_(mainSheet, parentRow);
    if (!parentTaskId) {
      return null;
    }

    const lastCol = Math.max(mainSheet.getLastColumn(), 21);
    const parentData = mainSheet.getRange(parentRow, 1, 1, lastCol).getValues()[0];
    const parentStatus = (parentData[CONFIG.MAIN_STATUS_COL - 1] || '').toString().trim() || 'Pipeline / To Do';

    if (isCompletedOrClosedStatus_(parentStatus)) {
      Logger.log(
        'prepareParentGoogleTaskForSubtask_: parent row ' + parentRow +
        ' is completed/closed — subtasks cannot nest under it.'
      );
      return null;
    }

    const targetListId = getTaskListIdForStatus_(parentStatus);
    if (!targetListId) {
      return null;
    }

    let listId = findTaskListIdContainingTask_(parentTaskId);
    if (!listId) {
      return null;
    }

    if (listId !== targetListId) {
      Tasks.Tasks.move(listId, parentTaskId, { destinationTasklist: targetListId });
      listId = targetListId;
      cacheTaskListIdForTask_(parentTaskId, targetListId);
      Logger.log('Moved parent task ' + parentTaskId + ' to list for status: ' + parentStatus);
    }

    ensureGoogleParentTaskOpenForSubtasks_(listId, parentTaskId, parentStatus);

    return { parentTaskId: parentTaskId, parentListId: listId, parentStatus: parentStatus };
  }

  /**
  * Removes stray top-level tasks with the same title as a subtask (from earlier buggy syncs).
  * Keeps the task whose ID is stored in Dev Tracker Sub column V.
  * @param {string} listId
  * @param {string} parentTaskId
  * @param {string} baseTitle Plain column D name
  * @param {string} [googleTitle] Suffixed Google Tasks title
  * @param {string} [keepTaskId]
  * @param {boolean} [removeKeepIdIfTopLevel] When true, also removes keepTaskId if it is still top-level (pre-insert cleanup).
  */
  function removeOrphanTopLevelDuplicateSubtasks_(
    listId,
    parentTaskId,
    baseTitle,
    googleTitle,
    keepTaskId,
    removeKeepIdIfTopLevel
  ) {
    const base = (baseTitle || '').toString().trim();
    const google = (googleTitle || '').toString().trim();
    if (!listId || !base) {
      return;
    }

    try {
      let pageToken = null;
      do {
        const opts = { showCompleted: true, showHidden: true, maxResults: 100 };
        if (pageToken) {
          opts.pageToken = pageToken;
        }
        const response = Tasks.Tasks.list(listId, opts);
        const items = response.items || [];

        for (let i = 0; i < items.length; i++) {
          const t = items[i];
          if (!t || !t.id) {
            continue;
          }
          if (t.id === parentTaskId) {
            continue;
          }
          if (t.parent) {
            continue;
          }
          if (!isOrphanTopLevelTitleMatchForSub_((t.title || '').toString(), base, google)) {
            continue;
          }
          if (keepTaskId && t.id === keepTaskId && !removeKeepIdIfTopLevel) {
            continue;
          }

          Tasks.Tasks.remove(listId, t.id);
          invalidateTaskListIdCacheForTask_(t.id);
          Logger.log('Removed orphan top-level duplicate: "' + (t.title || base) + '" (id ' + t.id + ')');
        }

        pageToken = response.nextPageToken || null;
      } while (pageToken);
    } catch (error) {
      Logger.log('removeOrphanTopLevelDuplicateSubtasks_ Error: ' + error.toString());
    }
  }

  /**
  * @param {string} listId
  * @param {string} taskId
  * @param {boolean} shouldBeCompleted
  */
  function setGoogleTaskCompletion_(listId, taskId, shouldBeCompleted) {
    const task = Tasks.Tasks.get(listId, taskId);
    const isCompleted = task.status === 'completed';

    if (shouldBeCompleted && !isCompleted) {
      patchGoogleTask_(listId, taskId, { status: 'completed' });
    } else if (!shouldBeCompleted && isCompleted) {
      patchGoogleTask_(listId, taskId, { status: 'needsAction' });
    }
  }

  /**
  * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
  * @param {string} taskName
  * @returns {{row: number, data: Array}|null}
  */
  function findMainRowByTaskName_(mainSheet, taskName, mainLookup) {
    const target = (taskName || '').toString().trim();
    if (!target) {
      return null;
    }

    if (mainLookup && mainLookup[target]) {
      return mainLookup[target];
    }

    if (!mainSheet) {
      return null;
    }

    const mainData = mainSheet.getDataRange().getValues();
    for (let i = 1; i < mainData.length; i++) {
      const name = (mainData[i][CONFIG.MAIN_TASK_NAME_COL - 1] || '').toString().trim();
      if (name === target) {
        return { row: i + 1, data: mainData[i] };
      }
    }

    return null;
  }

  /**
  * Finds a main task on Completed Tracker by Google Task ID (column V) or task name (column C).
  * @returns {{row: number, data: Array, matchedBy: string, onCompletedSheet: boolean}|null}
  */
  function findMainRowOnCompletedTracker_(linkedName, parentTaskId) {
    const ss = getTrackerSpreadsheet_();
    if (!ss) {
      return null;
    }
    const sheet = ss.getSheetByName(getCompletedSheetName_());
    if (!sheet || sheet.getLastRow() < 2) {
      return null;
    }

    const data = sheet.getDataRange().getValues();

    if (parentTaskId) {
      for (let i = 1; i < data.length; i++) {
        const id = (data[i][CONFIG.MAIN_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
        if (id === parentTaskId) {
          return {
            row: i + 1,
            data: data[i],
            matchedBy: 'completed-parent-google-task-id',
            onCompletedSheet: true
          };
        }
      }
    }

    const target = (linkedName || '').toString().trim();
    if (target) {
      for (let i = 1; i < data.length; i++) {
        const name = (data[i][CONFIG.MAIN_TASK_NAME_COL - 1] || '').toString().trim();
        if (name === target) {
          return {
            row: i + 1,
            data: data[i],
            matchedBy: 'completed-main-task-name',
            onCompletedSheet: true
          };
        }
      }
    }

    return null;
  }

  /**
  * Expected parent Google Task ID for a sub row from column B → main Dev Tracker column V
  * (then Completed Tracker if main was moved). Does not read sub column W.
  * @returns {string}
  */
  function getExpectedParentGoogleTaskIdForSubRow_(linkedName, mainSheet, mainLookup, mainData) {
    const name = (linkedName || '').toString().trim();
    if (!name) {
      return '';
    }

    const byName = findMainRowByTaskName_(mainSheet, name, mainLookup);
    if (byName && byName.data) {
      return (byName.data[CONFIG.MAIN_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
    }

    const onCompleted = findMainRowOnCompletedTracker_(name, '');
    if (onCompleted && onCompleted.data) {
      return (onCompleted.data[CONFIG.MAIN_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
    }

    return '';
  }

  /**
  * When sub column W is empty, copy Dev Tracker column V (main Google Task ID) into sub column W.
  * Never overwrites W if already set.
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} subRow
  * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
  * @param {{row: number, data?: Array}} parentInfo
  * @returns {boolean}
  */
  function fillSubColumnWFromMainVIfEmpty_(subSheet, subRow, mainSheet, parentInfo) {
    if (!subSheet || subRow < 2 || !mainSheet || !parentInfo || parentInfo.row < 2) {
      return false;
    }

    let mainV = getGoogleTaskIdForMainRow_(mainSheet, parentInfo.row);
    if (!mainV && parentInfo.data) {
      mainV = (parentInfo.data[CONFIG.MAIN_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
    }
    if (!mainV) {
      return false;
    }

    return refreshSubParentGoogleTaskIdIfEmpty_(subSheet, subRow, mainV);
  }

  /**
  * Column B (main task name on Dev Tracker) is the source of truth for parent matching.
  * Fills empty W from main V; corrects W when it does not match the linked main's column V.
  * @returns {boolean} True if column W was written or corrected
  */
  function reconcileSubColumnWToMainV_(subSheet, subRow, mainSheet, parentInfo) {
    if (!subSheet || subRow < 2 || !mainSheet || !parentInfo || parentInfo.onCompletedSheet) {
      return false;
    }

    const linkedName = (subSheet.getRange(subRow, CONFIG.SUB_LINK_COL).getValue() || '').toString().trim();
    if (!linkedName) {
      return fillSubColumnWFromMainVIfEmpty_(subSheet, subRow, mainSheet, parentInfo);
    }

    let mainV = getGoogleTaskIdForMainRow_(mainSheet, parentInfo.row);
    if (!mainV && parentInfo.data) {
      mainV = (parentInfo.data[CONFIG.MAIN_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
    }
    if (!mainV) {
      return false;
    }

    const col = findSubParentGoogleTaskIdColumn_(subSheet);
    const currentW = (subSheet.getRange(subRow, col).getValue() || '').toString().trim();
    if (currentW === mainV) {
      return false;
    }

    if (!currentW) {
      subSheet.getRange(subRow, col).setValue(mainV);
      logGoogleTasksSync_('Sub row ' + subRow + ': column W was empty — copied main column V → ' + mainV);
      return true;
    }

    subSheet.getRange(subRow, col).setValue(mainV);
    logGoogleTasksSync_(
      'Sub row ' + subRow + ': corrected column W ' + currentW + ' → ' + mainV +
      ' (Dev Tracker main via column B is always the parent)'
    );
    return true;
  }

  /**
  * Sets column W only when it is empty. Never overwrites an existing Main Google Task ID.
  * @returns {boolean} True if column W was written
  */
  function refreshSubParentGoogleTaskIdIfEmpty_(subSheet, subRow, parentTaskId) {
    if (!subSheet || subRow < 2 || !parentTaskId) {
      return false;
    }

    const col = findSubParentGoogleTaskIdColumn_(subSheet);
    const existing = (subSheet.getRange(subRow, col).getValue() || '').toString().trim();
    if (existing) {
      return false;
    }

    subSheet.getRange(subRow, col).setValue(parentTaskId);
    logGoogleTasksSync_('Sub row ' + subRow + ': column W was empty — copied main column V → ' + parentTaskId);
    return true;
  }

  /**
  * Master / bulk audit: compare sub column W with main column V (via column B link).
  * - W empty + main has V → fill W
  * - W set but ≠ main V → log MISMATCH (does not change W)
  * - W matches → OK
  * @returns {{ok: number, filled: number, mismatch: number, missingMain: number}}
  */
  function auditSubParentGoogleTaskIds_(subSheet, subData, mainLookup, mainData, fillEmpty) {
    const stats = { ok: 0, filled: 0, mismatch: 0, missingMain: 0 };
    if (!subSheet || !subData || subData.length <= 1) {
      return stats;
    }

    const mainSheet = getMainTrackerSheet_();
    const shouldFill = fillEmpty !== false;
    const parentCol = findSubParentGoogleTaskIdColumn_(subSheet);

    logGoogleTasksSync_('Auditing Main Google Task ID (column W) vs Dev Tracker column V (bottom → top)…');

    const subLast = CONFIG.GOOGLE_TASKS_SKIP_EMPTY_ROWS !== false
      ? getSubLastDataRow_(subData)
      : subSheet.getLastRow();
    eachSubSheetRowBottomUp_(subLast, subLast, function (subRow) {
      const i = subRow - 1;
      if (i < 1 || i >= subData.length) {
        return;
      }

      const linkedName = (subData[i][CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
      const subTaskName = (subData[i][CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();
      const currentW = (subData[i][parentCol - 1] || '').toString().trim();

      if (!linkedName && !subTaskName) {
        return;
      }

      if (!linkedName) {
        logGoogleTasksSync_('Sub row ' + subRow + ': column B (main link) is empty — cannot audit column W');
        return;
      }

      const expected = getExpectedParentGoogleTaskIdForSubRow_(linkedName, mainSheet, mainLookup, mainData);

      if (!expected) {
        stats.missingMain++;
        logGoogleTasksSync_(
          'Sub row ' + subRow + ' ("' + (subTaskName || 'no name') + '"): main "' + linkedName +
          '" has no Google Task ID on Dev Tracker — sync main row first (column V empty)'
        );
        return;
      }

      if (!currentW) {
        if (shouldFill) {
          subSheet.getRange(subRow, parentCol).setValue(expected);
          subData[i][parentCol - 1] = expected;
          stats.filled++;
          logGoogleTasksSync_('Sub row ' + subRow + ': filled empty column W → ' + expected);
        } else {
          logGoogleTasksSync_('Sub row ' + subRow + ': column W empty (expected ' + expected + ')');
        }
      } else if (currentW !== expected) {
        stats.mismatch++;
        if (shouldFill) {
          subSheet.getRange(subRow, parentCol).setValue(expected);
          subData[i][parentCol - 1] = expected;
          stats.filled++;
          logGoogleTasksSync_('Sub row ' + subRow + ': corrected column W → ' + expected + ' (matches Dev Tracker main V)');
        } else {
          logGoogleTasksSync_(
            'MISMATCH sub row ' + subRow + ' ("' + (subTaskName || 'no name') + '"): column W=' + currentW +
            ' but main "' + linkedName + '" column V=' + expected + ' — run backfillSubParentGoogleTaskIdsIfEmpty() to fix'
          );
        }
      } else {
        stats.ok++;
      }
    });

    logGoogleTasksSync_(
      'Parent ID audit done: OK=' + stats.ok + ', filled(empty W)=' + stats.filled +
      ', MISMATCH=' + stats.mismatch + ', missing main V=' + stats.missingMain
    );
    return stats;
  }

  /**
  * When a main row moves to Completed Tracker, fill column W on linked subs only if W is still empty.
  * @param {string} taskName
  * @param {string} parentGoogleTaskId
  */
  function fillSubParentGoogleTaskIdForLinkedSubsIfEmpty_(taskName, parentGoogleTaskId) {
    if (!parentGoogleTaskId || !taskName) {
      return;
    }

    const subSheet = getSubTrackerSheet_();
    if (!subSheet) {
      return;
    }

    const subData = subSheet.getDataRange().getValues();
    const linked = sortSubTaskEntriesBottomUp_(buildSubTaskIndexByMainName_(subData)[taskName] || []);

    for (let i = 0; i < linked.length; i++) {
      refreshSubParentGoogleTaskIdIfEmpty_(subSheet, linked[i].sheetRow, parentGoogleTaskId);
    }
  }

  /**
  * Resolves the main task row for a sub row. Lookup order:
  *   1. Column B — main task name on Dev Tracker (column C) — always first when set
  *   2. Column W — Main Google Task ID (matches Dev Tracker column V) — only when B is empty
  *   3. Column B or W on Completed Tracker (parent moved — sync will skip nesting)
  * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
  * @param {Array} subData
  * @param {Object} [mainLookup]
  * @param {Array} [mainData]
  * @param {GoogleAppsScript.Spreadsheet.Sheet} [subSheet]
  * @returns {{row: number, data: Array, matchedBy: string}|null}
  */
  function findMainRowForSub_(mainSheet, subData, mainLookup, mainData, subSheet) {
    if (!subData) {
      return null;
    }

    const parentTaskId = subSheet
      ? getSubParentGoogleTaskIdFromRowData_(subData, subSheet)
      : (subData[CONFIG.SUB_PARENT_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
    const linkedName = (subData[CONFIG.SUB_LINK_COL - 1] || '').toString().trim();

    if (!mainSheet) {
      mainSheet = getMainTrackerSheet_();
    }
    if (!mainSheet) {
      return null;
    }

    if (!mainData) {
      mainData = mainSheet.getDataRange().getValues();
    }
    if (!mainLookup) {
      mainLookup = buildMainTaskNameIndex_(mainData);
    }

    if (linkedName) {
      const byName = findMainRowByTaskName_(mainSheet, linkedName, mainLookup);
      if (byName) {
        byName.matchedBy = 'main-task-name';
        return byName;
      }
    }

    if (parentTaskId) {
      for (let i = 1; i < mainData.length; i++) {
        const mainId = (mainData[i][CONFIG.MAIN_GOOGLE_TASK_ID_COL - 1] || '').toString().trim();
        if (mainId === parentTaskId) {
          return { row: i + 1, data: mainData[i], matchedBy: 'parent-google-task-id' };
        }
      }
    }

    const onCompleted = findMainRowOnCompletedTracker_(linkedName, parentTaskId);
    if (onCompleted) {
      return onCompleted;
    }

    return null;
  }

  /**
  * After a main task syncs, set column W on linked subs only where W is still empty.
  * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
  * @param {number} mainRow
  */
  function refreshSubParentMappingsForMainRow_(mainSheet, mainRow) {
    const subSheet = getSubTrackerSheet_();
    if (!mainSheet || !subSheet || mainRow < 2) {
      return;
    }

    const parentTaskId = getGoogleTaskIdForMainRow_(mainSheet, mainRow);
    if (!parentTaskId) {
      return;
    }

    const taskName = (mainSheet.getRange(mainRow, CONFIG.MAIN_TASK_NAME_COL).getValue() || '')
      .toString()
      .trim();
    if (!taskName) {
      return;
    }

    const subData = subSheet.getDataRange().getValues();
    const linked = sortSubTaskEntriesBottomUp_(buildSubTaskIndexByMainName_(subData)[taskName] || []);

    for (let i = 0; i < linked.length; i++) {
      refreshSubParentGoogleTaskIdIfEmpty_(subSheet, linked[i].sheetRow, parentTaskId);
    }
  }

  /**
  * One-time repair: fill or correct column W from Dev Tracker column V (via column B link).
  * Wrong W values are corrected to match the linked main's column V.
  */
  function backfillSubParentGoogleTaskIdsIfEmpty() {
    beginGoogleTasksSync_('BULK', 'backfill Main Google Task ID (column W) from Dev Tracker column V where empty');
    const mainSheet = getMainTrackerSheet_();
    const subSheet = getSubTrackerSheet_();
    if (!mainSheet || !subSheet) {
      Logger.log('Sheets not found');
      endGoogleTasksSync_('aborted');
      return;
    }

    const mainData = mainSheet.getDataRange().getValues();
    const mainLookup = buildMainTaskNameIndex_(mainData);
    const subData = subSheet.getDataRange().getValues();
    const stats = auditSubParentGoogleTaskIds_(subSheet, subData, mainLookup, mainData, true);
    endGoogleTasksSync_(
      'filled ' + stats.filled + ', mismatch ' + stats.mismatch + ', missing main V ' + stats.missingMain
    );
  }

  /** @deprecated Use backfillSubParentGoogleTaskIdsIfEmpty() */
  function backfillSubParentMappingColumns() {
    backfillSubParentGoogleTaskIdsIfEmpty();
  }

  /**
  * Validates whether a Dev Tracker Sub row should sync to Google Tasks.
  *
  * SKIP (do not add; remove if already in Google Tasks) ONLY when BOTH match:
  *   • Status = Completed, Closed / Cancelled, Closed – Billable, or Closed – Not Proceeding
  *   • AND Environment = Production
  *
  * ADD for all other combinations (e.g. Closed/Cancelled + Staging, In Progress + Production).
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  * @param {Array} [subData] Optional row data (columns A…)
  * @param {Object<string, {row: number, data: Array}>} [mainLookup] Pre-built main sheet index
  * @returns {{ok: boolean, message: string, parentInfo: Object|null, subData: Array|null}}
  */
  function validateSubtaskGoogleSync_(subSheet, row, subData, mainLookup) {
    const fail = function (message) {
      return { ok: false, message: message, parentInfo: null, subData: subData || null };
    };

    if (!subSheet || row < 2) {
      return fail('SKIP sub-task row ' + row + ': invalid row.');
    }

    if (!subData) {
      const lastCol = Math.max(subSheet.getLastColumn(), CONFIG.SUB_GOOGLE_TASK_ID_COL);
      subData = subSheet.getRange(row, 1, 1, lastCol).getValues()[0];
    }

    const linkedMainName = (subData[CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
    const subTaskName = (subData[CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();
    const subStatus = (subData[CONFIG.SUB_STATUS_COL - 1] || '').toString().trim();
    const subEnv = (subData[CONFIG.SUB_ENV_COL - 1] || '').toString().trim();
    const parentIdCol = getSubParentGoogleTaskIdFromRowData_(subData, subSheet);

    if (!linkedMainName && !parentIdCol) {
      return fail(
        'SKIP sub-task row ' + row + ': no main link — fill column B (task name) or run backfillSubParentGoogleTaskIdsIfEmpty().'
      );
    }

    if (!subTaskName) {
      return fail('SKIP sub-task row ' + row + ': column D (sub-task name) is empty.');
    }

    if (shouldSkipSubtaskGoogleSync_(subStatus, subEnv)) {
      return fail(
        'SKIP sub-task row ' + row + ' ("' + subTaskName + '"): status "' + subStatus +
        '" + Environment "' + subEnv + '" — excluded combo (Completed/Closed + Production only); ' +
        'removed from Google Tasks if it existed.'
      );
    }

    const mainSheet = getMainTrackerSheet_();
    if (!mainSheet) {
      return fail('SKIP sub-task row ' + row + ': Dev Tracker sheet not found.');
    }

    const parentInfo = findMainRowForSub_(mainSheet, subData, mainLookup, null, subSheet);
    if (!parentInfo) {
      return fail(
        'SKIP sub-task row ' + row + ' ("' + subTaskName + '"): main task not found ' +
        '(check column B or W — run backfillSubParentGoogleTaskIdsIfEmpty()).'
      );
    }

    if (parentInfo.onCompletedSheet) {
      return fail(
        'SKIP sub-task row ' + row + ' ("' + subTaskName + '"): main task is on Completed Tracker ' +
        '(matched via ' + parentInfo.matchedBy + ') — subtasks cannot nest under a completed/closed parent. ' +
        'Column W kept for reference.'
      );
    }

    return {
      ok: true,
      message: '',
      parentInfo: parentInfo,
      subData: subData,
      matchedBy: parentInfo.matchedBy || 'main-task-name'
    };
  }

  /**
  * Creates or updates the parent Google Task for one Dev Tracker main row.
  *
  * Call as:
  *   syncGoogleTaskForMainRow(mainSheet, 2)  — from onEdit / internal code
  *   syncGoogleTaskForMainRow(2)             — from Apps Script editor (row only)
  *   syncExistingTasksToGoogle()             — bulk sync (all rows)
  *
  * @param {GoogleAppsScript.Spreadsheet.Sheet|number} mainSheetOrRow
  * @param {number} [row]
  */
  function syncGoogleTaskForMainRow(mainSheetOrRow, row) {
    try {
      if (!isGoogleTasksApiEnabled_()) {
        return;
      }

      if (arguments.length === 0) {
        Logger.log(
          'syncGoogleTaskForMainRow: pass a row number (e.g. syncGoogleTaskForMainRow(2)). ' +
          'For bulk sync use syncExistingTasksToGoogle() or syncAllGoogleTasksFromDevTracker().'
        );
        return;
      }

      let mainSheet = mainSheetOrRow;

      // Run from editor: syncGoogleTaskForMainRow(2) — only row number passed
      if (arguments.length === 1 && (typeof mainSheetOrRow === 'number' || typeof mainSheetOrRow === 'string')) {
        const asNum = parseInt(mainSheetOrRow, 10);
        if (!isNaN(asNum)) {
          row = asNum;
          mainSheet = null;
        }
      }

      if (!mainSheet || typeof mainSheet.getRange !== 'function') {
        mainSheet = getMainTrackerSheet_();
        if (!mainSheet) {
          return;
        }
      }

      const rowNum = parseInt(row, 10);
      if (!rowNum || rowNum < 2) {
        Logger.log(
          'syncGoogleTaskForMainRow: invalid row "' + row + '" — skipping. ' +
          'Use syncAllGoogleTasksFromDevTracker() or syncExistingTasksToGoogle() for bulk sync.'
        );
        return;
      }
      row = rowNum;

      ensureGoogleTasksRowSyncContext_('main', row);

      const lastCol = Math.max(mainSheet.getLastColumn(), 21);
      const data = mainSheet.getRange(row, 1, 1, lastCol).getValues()[0];
      const taskName = (data[CONFIG.MAIN_TASK_NAME_COL - 1] || '').toString().trim();
      if (!taskName) {
        return;
      }

      const status = (data[CONFIG.MAIN_STATUS_COL - 1] || '').toString().trim() || 'Pipeline / To Do';
      const listId = getTaskListIdForStatus_(status);
      if (!listId) {
        Logger.log('No Google Task list for status "' + status + '" on row ' + row);
        return;
      }

      const notes = buildMainGoogleTaskNotes_(data, row);
      const taskResource = {
        title: taskName,
        notes: notes
      };

      let taskId = getGoogleTaskIdForMainRow_(mainSheet, row);
      let activeListId = listId;

      if (!taskId) {
        const created = Tasks.Tasks.insert(taskResource, listId);
        taskId = created.id;
        activeListId = listId;
        setGoogleTaskIdForMainRow_(mainSheet, row, taskId);
        cacheTaskListIdForTask_(taskId, listId);
        Logger.log('Created Google Task for row ' + row + ': ' + taskId);
      } else {
        const currentListId = findTaskListIdContainingTask_(taskId, listId);
        if (!currentListId) {
          const recreated = Tasks.Tasks.insert(taskResource, listId);
          setGoogleTaskIdForMainRow_(mainSheet, row, recreated.id);
          setGoogleTaskCompletion_(listId, recreated.id, isCompletedOrClosedStatus_(status));
          Logger.log('Recreated missing Google Task for row ' + row);
          return;
        }

        activeListId = currentListId;
        if (currentListId !== listId) {
          Tasks.Tasks.move(currentListId, taskId, { destinationTasklist: listId });
          activeListId = listId;
          cacheTaskListIdForTask_(taskId, listId);
          Logger.log('Moved Google Task for row ' + row + ' to list for: ' + status);
        }

        patchGoogleTask_(activeListId, taskId, taskResource);
      }

      setGoogleTaskCompletion_(activeListId, taskId, isCompletedOrClosedStatus_(status));
      cacheTaskListIdForTask_(taskId, activeListId);
      logGoogleTasksSync_('main task synced row ' + row + ' id=' + taskId);
      refreshSubParentMappingsForMainRow_(mainSheet, row);
    } catch (error) {
      Logger.log('syncGoogleTaskForMainRow Error (row ' + row + '): ' + error.toString());
      clearInvalidGoogleTaskIdIfNeeded_(mainSheet, row, null, error);
    } finally {
      if (googleTasksSyncContext_ && googleTasksSyncContext_.mode === 'ROW') {
        endGoogleTasksSync_('main row ' + row);
      }
    }
  }

  /**
  * Creates or updates one Google subtask under its parent main task.
  * Validation runs first — stale column V IDs are removed when the main link is invalid.
  * Skips when: main task missing from Dev Tracker, subtask Completed/Closed, or empty link/name.
  * @param {GoogleAppsScript.Spreadsheet.Sheet} subSheet
  * @param {number} row
  * @param {{ok: boolean, parentInfo: Object, subData: Array}|null} [preValidated]
  * @param {Object<string, number>} [titleDuplicateIndex] Pre-built column D duplicate counts (bulk sync)
  */
  function syncGoogleSubtaskForSubRow(subSheet, row, preValidated, titleDuplicateIndex) {
    try {
      if (!isGoogleTasksApiEnabled_()) {
        return;
      }

      if (!preValidated) {
        ensureGoogleTasksRowSyncContext_('sub', row);
      }

      const lastCol = Math.max(subSheet.getLastColumn(), CONFIG.SUB_GOOGLE_TASK_ID_COL);
      let subData = subSheet.getRange(row, 1, 1, lastCol).getValues()[0];

      const mainSheetForGuard = getMainTrackerSheet_();
      const mainDataForGuard = mainSheetForGuard ? mainSheetForGuard.getDataRange().getValues() : [];
      const mainIdIndex = buildMainGoogleTaskIdIndex_(mainDataForGuard);
      if (clearSubColumnVIfDevTrackerMainTaskId_(subSheet, row, mainIdIndex)) {
        subData = subSheet.getRange(row, 1, 1, lastCol).getValues()[0];
      }

      if (!titleDuplicateIndex) {
        titleDuplicateIndex = buildSubTaskTitleDuplicateIndex_(subSheet.getDataRange().getValues());
      }

      const validation = preValidated || validateSubtaskGoogleSync_(subSheet, row, subData);
      if (!validation.ok) {
        cleanupGoogleSubtaskOnSyncSkip_(subSheet, row, subData);
        if (!preValidated) {
          Logger.log(validation.message);
        }
        return;
      }

      const parentInfo = validation.parentInfo;
      const linkedMainName = (subData[CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
      const subTaskName = (subData[CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();

      const mainSheet = getMainTrackerSheet_();

      reconcileSubColumnWToMainV_(subSheet, row, mainSheet, parentInfo);
      if (validation.matchedBy) {
        logGoogleTasksSync_('parent resolved via ' + validation.matchedBy + ' → Dev Tracker main row ' + parentInfo.row);
      }

      const parentCtx = prepareParentGoogleTaskForSubtask_(mainSheet, parentInfo.row);
      if (!parentCtx) {
        Logger.log(
          'SKIP sub-task row ' + row + ' ("' + subTaskName + '"): parent Google Task not ready for main task "' +
          linkedMainName + '" (row ' + parentInfo.row + '). Ensure Dev Tracker column V has a valid main Task ID.'
        );
        return;
      }

      reconcileSubColumnWToMainV_(subSheet, row, mainSheet, parentInfo);

      const parentTaskId = parentCtx.parentTaskId;
      const parentListId = parentCtx.parentListId;
      const parentSheetStatus = parentCtx.parentStatus || '';

      const taskResource = buildSubGoogleTaskResource_(subData, row, linkedMainName, titleDuplicateIndex);
      if (taskResource.googleTitle !== taskResource.baseTitle) {
        logGoogleTasksSync_(
          'sub row ' + row + ' Google title: "' + taskResource.googleTitle + '" (sheet column D unchanged)'
        );
      }

      const priorSubTaskId = getGoogleSubtaskIdForSubRow_(subSheet, row);
      const subResult = ensureGoogleSubtaskUnderParent_(
        taskResource,
        parentListId,
        parentTaskId,
        priorSubTaskId || null,
        parentSheetStatus
      );

      if (!subResult) {
        setGoogleSubtaskIdForSubRow_(subSheet, row, '');
        removeOrphanTopLevelTasksByTitleAcrossAllStatusLists_(
          taskResource.baseTitle,
          '',
          taskResource.googleTitle
        );
        Logger.log(
          'SKIP sub-task row ' + row + ' ("' + subTaskName + '"): could not nest under parent — ' +
          'cleared column V and removed any top-level orphan with this title.'
        );
        return;
      }

      const subTaskId = subResult.taskId;
      if (subResult.recreated || subTaskId !== priorSubTaskId) {
        setGoogleSubtaskIdForSubRow_(subSheet, row, subTaskId);
        cacheTaskListIdForTask_(subTaskId, subResult.listId);
      }

      if (!subResult.recreated) {
        patchGoogleTask_(subResult.listId, subTaskId, taskResource);
      }

      removeOrphanTopLevelTasksByTitleAcrossAllStatusLists_(
        taskResource.baseTitle,
        subTaskId,
        taskResource.googleTitle
      );
      logGoogleTasksSync_(
        'sub row ' + row + ' id=' + subTaskId + ' under parent ' + parentTaskId +
        (subResult.recreated ? ' (recreated)' : '')
      );
    } catch (error) {
      Logger.log('syncGoogleSubtaskForSubRow Error (row ' + row + '): ' + error.toString());
      clearInvalidGoogleTaskIdIfNeeded_(null, row, subSheet, error);
    } finally {
      try {
        if (subSheet && row >= 2 && isGoogleTasksApiEnabled_()) {
          refreshSubNestStatusForRow_(subSheet, row);
        }
      } catch (nestErr) {
        Logger.log('refreshSubNestStatusForRow_ Error (row ' + row + '): ' + nestErr.toString());
      }
      if (!preValidated && googleTasksSyncContext_ && googleTasksSyncContext_.mode === 'ROW') {
        endGoogleTasksSync_('sub row ' + row);
      }
    }
  }

  /**
  * Clears a stale Google Task ID from column V when the API rejects it (wrong/corrupt ID).
  * @param {GoogleAppsScript.Spreadsheet.Sheet|null} mainSheet
  * @param {number} row
  * @param {GoogleAppsScript.Spreadsheet.Sheet|null} subSheet
  * @param {Object} error
  */
  function clearInvalidGoogleTaskIdIfNeeded_(mainSheet, row, subSheet, error) {
    const msg = (error && error.toString ? error.toString() : String(error)).toLowerCase();
    if (msg.indexOf('invalid json') === -1 && msg.indexOf('not found') === -1 && msg.indexOf('404') === -1) {
      return;
    }

    if (mainSheet) {
      setGoogleTaskIdForMainRow_(mainSheet, row, '');
      Logger.log('Cleared invalid Google Task ID on Dev Tracker row ' + row + ' — will recreate on next sync.');
    }
    if (subSheet) {
      setGoogleSubtaskIdForSubRow_(subSheet, row, '');
      Logger.log('Cleared invalid Google Subtask ID on Dev Tracker Sub row ' + row + ' — will recreate on next sync.');
    }
  }

  /**
  * Re-sync sub rows linked to one main task (sub→main index; not a full-sheet API loop).
  * Called on main Status edit unless GOOGLE_TASKS_LIGHT_MAIN_STATUS_EDIT is true.
  * @param {GoogleAppsScript.Spreadsheet.Sheet} mainSheet
  * @param {number} row
  * @param {Array} [subData] Optional pre-loaded sub sheet data
  * @param {Object} [mainLookup] Optional pre-built main index
  */
  function syncGoogleSubtasksForMainRow_(mainSheet, row, subData, mainLookup) {
    try {
      if (!mainSheet || !row || row < 2) {
        return;
      }

    const subSheet = getSubTrackerSheet_();
    if (!subSheet) {
      return;
    }

      const taskName = (mainSheet.getRange(row, CONFIG.MAIN_TASK_NAME_COL).getValue() || '')
        .toString()
        .trim();
      if (!taskName) {
        return;
      }

      if (!subData) {
        subData = subSheet.getDataRange().getValues();
      }
      if (!mainLookup) {
        const mainData = mainSheet.getDataRange().getValues();
        mainLookup = buildMainTaskNameIndex_(mainData);
      }

      const subIndex = buildSubTaskIndexByMainName_(subData);
      const linked = sortSubTaskEntriesBottomUp_(subIndex[taskName] || []);

      beginGoogleTasksSync_('TARGETED', 'main row ' + row + ' "' + taskName + '" → ' + linked.length + ' sub(s)');
      logGoogleTasksSync_('sub→main index lookup (1 sheet read, not per-row scan)');

      let synced = 0;
      for (let i = 0; i < linked.length; i++) {
        const entry = linked[i];
        const validation = validateSubtaskGoogleSync_(subSheet, entry.sheetRow, entry.data, mainLookup);
        if (validation.ok) {
          syncGoogleSubtaskForSubRow(subSheet, entry.sheetRow, validation);
          synced++;
        }
      }

      endGoogleTasksSync_('main row ' + row + ': ' + synced + '/' + linked.length + ' linked sub(s) synced');
    } catch (error) {
      Logger.log('syncGoogleSubtasksForMainRow_ Error (row ' + row + '): ' + error.toString());
      endGoogleTasksSync_('ERROR row ' + row);
    }
  }

  /**
  * Run from Apps Script to sync one main row: syncGoogleTaskForMainRowByNumber(2)
  * @param {number} row Dev Tracker data row (>= 2)
  */
  function syncGoogleTaskForMainRowByNumber(row) {
    syncGoogleTaskForMainRow(row);
  }

  /**
  * SETUP HELPER: Create/sync all main Google Tasks and related subtasks from both sheets.
  * Run once after enabling Google Tasks API and creating the status lists.
  */
  function syncAllGoogleTasksFromDevTracker() {
    let bulkFlagSet = false;
    try {
      if (!isGoogleTasksApiEnabled_()) {
        return;
      }

      const resumingPausedBulk = isGoogleTasksBulkSyncPaused_();
      if (isGoogleTasksBulkSyncInProgress_() && !resumingPausedBulk) {
        Logger.log(
          'syncAllGoogleTasksFromDevTracker: bulk sync already in progress — skipping overlapping run.'
        );
        return;
      }

      setGoogleTasksBulkSyncInProgress_(true);
      bulkFlagSet = true;

      beginGoogleTasksSync_('BULK', 'full sheet — Dev Tracker mains first, then subs');
      logGoogleTasksSync_('=== Syncing all Google Tasks (main parents first, then subs bottom-up) ===');
      refreshGoogleTaskListCache_();
      taskIdToListIdCache_ = {};

      const mainSheet = getMainTrackerSheet_();
      const subSheet = getSubTrackerSheet_();

      if (!mainSheet) {
        Logger.log('Dev Tracker sheet not found.');
        endGoogleTasksSync_('aborted — main sheet missing');
        return;
      }

      ensureColumnCachesFresh_(mainSheet, subSheet);

      const mainData = mainSheet.getDataRange().getValues();
      const mainLookup = buildMainTaskNameIndex_(mainData);
      const mainIdIndex = buildMainGoogleTaskIdIndex_(mainData);
      const subData = subSheet ? subSheet.getDataRange().getValues() : null;
      const syncSleep = CONFIG.GOOGLE_TASKS_SYNC_SLEEP_MS || 50;
      const maxRuntime = CONFIG.GOOGLE_TASKS_SYNC_MAX_RUNTIME_MS || 270000;
      const batchLimit = getGoogleTasksBulkBatchLimit_();
      const syncStartedAt = Date.now();

      if (batchLimit > 0) {
        logGoogleTasksSync_('Batch size: up to ' + batchLimit + ' rows per run (auto-continues until finished)');
      }

      const props = PropertiesService.getScriptProperties();
      let resumePhase = props.getProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_PHASE) || 'main';
      // Legacy resume keys from older bulk orders
      if (resumePhase === 'main-orphans') {
        resumePhase = 'main';
      }
      let resumeRow = parseInt(props.getProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_ROW) || '0', 10);

      const skipEmpty = CONFIG.GOOGLE_TASKS_SKIP_EMPTY_ROWS !== false;
      const subLastForScan = subSheet && subData
        ? (skipEmpty ? getSubLastDataRow_(subData) : subSheet.getLastRow())
        : 1;

      let mainRowsToSync = [];
      if (skipEmpty) {
        mainRowsToSync = getMainRowNumbersWithTaskNames_(mainData);
        if (resumePhase === 'main' && resumeRow >= 2) {
          mainRowsToSync = mainRowsToSync.filter(function (r) {
            return r >= resumeRow;
          });
        }
        logGoogleTasksSync_(
          'Dev Tracker: sync ' + mainRowsToSync.length + ' row(s) with task names only' +
          ' (empty column C rows skipped — sheet last row ' + mainSheet.getLastRow() + ')'
        );
        logGoogleTasksSync_(
          'Dev Tracker Sub: data through row ' + subLastForScan +
          ' (empty sub rows skipped)'
        );
      } else {
        const mainEnd = mainSheet.getLastRow();
        const mainStart = resumePhase === 'main' && resumeRow >= 2 ? resumeRow : 2;
        for (let r = mainStart; r <= mainEnd; r++) {
          mainRowsToSync.push(r);
        }
      }

      let subEmptySkipped = 0;

      let mainCreated = 0;
      let mainSkipped = 0;
      let mainErrors = 0;
      let mainSynced = 0;

      let subCreated = 0;
      let subSkipped = 0;

      // Phase 1: Dev Tracker rows that have a task name (column C) — empty rows never scanned when skipEmpty
      if (resumePhase === 'main') {
        googleTasksSyncContext_.mode = 'BULK-MAIN';
        logGoogleTasksSync_(
          'Phase 1/2: Dev Tracker — ' + mainRowsToSync.length + ' main row(s) to sync (parents first)'
        );

        let mainProcessedThisChunk = 0;
        for (let mi = 0; mi < mainRowsToSync.length; mi++) {
          const r = mainRowsToSync[mi];
          const stopReason = shouldStopBulkSyncChunk_(syncStartedAt, maxRuntime, mainProcessedThisChunk);
          if (stopReason) {
            pauseBulkGoogleTasksSync_(
              props,
              'main',
              r,
              stopReason,
              'main batch: ' + mainProcessedThisChunk + ' synced this run, ' +
              (mainRowsToSync.length - mi) + ' main row(s) left — next at row ' + r
            );
            return;
          }

          if (!skipEmpty) {
            const taskName = (mainData[r - 1][CONFIG.MAIN_TASK_NAME_COL - 1] || '').toString().trim();
            if (!taskName) {
              mainSkipped++;
              continue;
            }
          }

          const beforeId = getGoogleTaskIdForMainRow_(mainSheet, r);
          syncGoogleTaskForMainRow(mainSheet, r);
          const afterId = getGoogleTaskIdForMainRow_(mainSheet, r);

          if (afterId) {
            mainSynced++;
            if (!beforeId) {
              mainCreated++;
            }
          } else {
            mainErrors++;
          }

          mainProcessedThisChunk++;
          Utilities.sleep(syncSleep);
        }

        resumeRow = 0;
        resumePhase = 'sub';
      }

      // Phase 2: Dev Tracker Sub rows bottom-up (newest at bottom first)
      if (resumePhase === 'sub' && subSheet && subData) {
        let subRowsToSync;
        if (skipEmpty) {
          subRowsToSync = getSubRowNumbersWithDataBottomUp_(subData, subLastForScan);
          if (resumeRow >= 2) {
            subRowsToSync = subRowsToSync.filter(function (r) {
              return r <= resumeRow;
            });
          }
        } else {
          subRowsToSync = [];
          const subScanStartLegacy = getSubSheetScanStartRow_(subLastForScan, resumeRow);
          for (let r = subScanStartLegacy; r >= 2; r--) {
            subRowsToSync.push(r);
          }
        }

        const isFreshSubPass = resumeRow === 0 || resumeRow > subLastForScan;

        if (isFreshSubPass) {
          props.setProperty(SCRIPT_PROP_GOOGLE_SYNC_SUB_TOTAL, String(subRowsToSync.length));
          auditSubParentGoogleTaskIds_(subSheet, subData, mainLookup, mainData, true);
        }

        const subTotal = parseInt(
          props.getProperty(SCRIPT_PROP_GOOGLE_SYNC_SUB_TOTAL) || String(subRowsToSync.length),
          10
        );

        googleTasksSyncContext_.mode = 'BULK-SUB';
        logGoogleTasksSync_(
          'Phase 2/2: Dev Tracker Sub — ' + subRowsToSync.length + ' sub row(s) this phase' +
          (subTotal ? ' (total ~' + subTotal + ')' : '') + ', bottom-up'
        );

        const subTitleDuplicateIndex = buildSubTaskTitleDuplicateIndex_(subData);

        let subProcessedThisChunk = 0;
        for (let si = 0; si < subRowsToSync.length; si++) {
          const r = subRowsToSync[si];
          const stopReason = shouldStopBulkSyncChunk_(syncStartedAt, maxRuntime, subProcessedThisChunk);
          if (stopReason) {
            const remaining = subRowsToSync.length - si;
            pauseBulkGoogleTasksSync_(
              props,
              'sub',
              r,
              stopReason,
              'sub batch: ' + subProcessedThisChunk + ' synced this run, ' + remaining +
              ' sub row(s) left — next at row ' + r + ' (~' + subTotal + ' total subs)'
            );
            return;
          }

          const rowData = subData[r - 1];
          if (!skipEmpty && isSubSheetDataRowEmpty_(rowData)) {
            subEmptySkipped++;
            continue;
          }

          clearSubColumnVIfDevTrackerMainTaskId_(subSheet, r, mainIdIndex);

          const validation = validateSubtaskGoogleSync_(subSheet, r, rowData, mainLookup);
          if (!validation.ok) {
            subSkipped++;
            if (!isSubSheetDataRowEmpty_(rowData)) {
              cleanupGoogleSubtaskOnSyncSkip_(subSheet, r, validation.subData);
            }
            continue;
          }

          const beforeSubId = getGoogleSubtaskIdForSubRow_(subSheet, r);
          syncGoogleSubtaskForSubRow(subSheet, r, validation, subTitleDuplicateIndex);
          const afterSubId = getGoogleSubtaskIdForSubRow_(subSheet, r);
          if (afterSubId && !beforeSubId) {
            subCreated++;
          }

          subProcessedThisChunk++;
          Utilities.sleep(syncSleep);
        }
      }

      props.deleteProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_ROW);
      props.deleteProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_PHASE);
      props.deleteProperty(SCRIPT_PROP_GOOGLE_SYNC_SUB_TOTAL);
      cancelGoogleTasksBulkSyncContinueTriggers_();

      const summary =
        'mains synced ' + mainSynced + ' (created ' + mainCreated + '), mains skipped ' + mainSkipped +
        ', main errors ' + mainErrors +
        '; subs created ' + subCreated + ', subs skipped ' + subSkipped +
        (subEmptySkipped ? ', empty sub rows skipped ' + subEmptySkipped : '');
      logGoogleTasksSync_('=== Google Tasks sync complete ===');
      logGoogleTasksSync_(summary);
      endGoogleTasksSync_(summary);
    } catch (error) {
      Logger.log('syncAllGoogleTasksFromDevTracker Error: ' + error.toString());
      endGoogleTasksSync_('ERROR — ' + error.toString());
    } finally {
      if (bulkFlagSet && !isGoogleTasksBulkSyncPaused_()) {
        setGoogleTasksBulkSyncInProgress_(false);
      }
    }
  }

  /**
  * Lists every Dev Tracker Sub row whose column B does not match any Dev Tracker column C task.
  * Run from Apps Script editor → View → Executions to read the log.
  */
  function fixAllOrphanSubtasks() {
    if (!isGoogleTasksApiEnabled_()) {
      Logger.log('Google Tasks API not enabled — add Services (+) > Google Tasks API first.');
      return;
    }

    const mainSheet = getMainTrackerSheet_();
    const subSheet = getSubTrackerSheet_();

    if (!mainSheet || !subSheet) {
      Logger.log('Sheets not found');
      return;
    }

    refreshGoogleTaskListCache_();
    taskIdToListIdCache_ = {};

    const mainData = mainSheet.getDataRange().getValues();
    const mainLookup = buildMainTaskNameIndex_(mainData);
    const subAllData = subSheet.getDataRange().getValues();
    const subTitleDuplicateIndex = buildSubTaskTitleDuplicateIndex_(subAllData);
    const subLast = subSheet.getLastRow();

    let fixed = 0;
    let skipped = 0;
    let failed = 0;

    for (let r = subLast; r >= 2; r--) {
      const subData = subSheet
        .getRange(r, 1, 1, Math.max(subSheet.getLastColumn(), CONFIG.SUB_GOOGLE_TASK_ID_COL))
        .getValues()[0];
      const subTaskId = getGoogleSubtaskIdForSubRow_(subSheet, r);
      const subTaskName = (subData[CONFIG.SUB_TASK_NAME_COL - 1] || '').toString().trim();

      if (!subTaskId || !subTaskName) {
        skipped++;
        continue;
      }

      const validation = validateSubtaskGoogleSync_(subSheet, r, subData, mainLookup);
      if (!validation.ok) {
        cleanupGoogleSubtaskOnSyncSkip_(subSheet, r, subData);
        Logger.log('Row ' + r + ': ' + validation.message);
        skipped++;
        continue;
      }

      const parentInfo = validation.parentInfo;
      const parentCtx = prepareParentGoogleTaskForSubtask_(mainSheet, parentInfo.row);

      if (!parentCtx) {
        Logger.log('Row ' + r + ' (' + subTaskName + '): parent not ready — skipped');
        skipped++;
        continue;
      }

      const alreadyNested = verifyGoogleTaskNestedUnderParent_(
        parentCtx.parentListId,
        subTaskId,
        parentCtx.parentTaskId,
        parentCtx.parentListId
      );

      const linkedMainName = (subData[CONFIG.SUB_LINK_COL - 1] || '').toString().trim();
      const taskResource = buildSubGoogleTaskResource_(subData, r, linkedMainName, subTitleDuplicateIndex);

      if (alreadyNested) {
        patchGoogleTask_(parentCtx.parentListId, subTaskId, taskResource);
        removeOrphanTopLevelDuplicateSubtasks_(
          parentCtx.parentListId,
          parentCtx.parentTaskId,
          taskResource.baseTitle,
          taskResource.googleTitle,
          subTaskId,
          false
        );
        Logger.log('Row ' + r + ' (' + subTaskName + '): already nested — title/notes refreshed');
        skipped++;
        continue;
      }

      Logger.log('Row ' + r + ' (' + subTaskName + '): fixing orphan...');

      const result = ensureGoogleSubtaskUnderParent_(
        taskResource,
        parentCtx.parentListId,
        parentCtx.parentTaskId,
        subTaskId,
        parentCtx.parentStatus
      );

      if (result) {
        if (result.taskId !== subTaskId) {
          setGoogleSubtaskIdForSubRow_(subSheet, r, result.taskId);
          cacheTaskListIdForTask_(result.taskId, result.listId);
          Logger.log('Row ' + r + ': recreated as ' + result.taskId);
        } else {
          cacheTaskListIdForTask_(result.taskId, result.listId);
          Logger.log('Row ' + r + ': moved under parent successfully');
        }
        removeOrphanTopLevelDuplicateSubtasks_(
          parentCtx.parentListId,
          parentCtx.parentTaskId,
          taskResource.baseTitle,
          taskResource.googleTitle,
          result.taskId,
          false
        );
        fixed++;
      } else {
        Logger.log('Row ' + r + ' (' + subTaskName + '): FAILED to fix');
        failed++;
      }

      Utilities.sleep(300);
    }

    Logger.log('=== Fix complete ===');
    Logger.log('Fixed: ' + fixed + ' | Skipped: ' + skipped + ' | Failed: ' + failed);
  }

  /**
  * Clears bulk-sync resume state from script properties.
  */
  function clearGoogleTasksSyncResumeState_() {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_ROW);
    props.deleteProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_PHASE);
    props.deleteProperty(SCRIPT_PROP_GOOGLE_SYNC_SUB_TOTAL);
    cancelGoogleTasksBulkSyncContinueTriggers_();
  }

  /**
  * @returns {number} Max rows per batch (0 = unlimited except time).
  */
  function getGoogleTasksBulkBatchLimit_() {
    const n = CONFIG.GOOGLE_TASKS_SYNC_BATCH_SIZE;
    if (n === undefined || n === null || n === '') {
      return 300;
    }
    const parsed = parseInt(n, 10);
    return isNaN(parsed) || parsed < 0 ? 300 : parsed;
  }

  /**
  * @returns {'batch'|'time'|null}
  */
  function shouldStopBulkSyncChunk_(syncStartedAt, maxRuntime, processedThisChunk) {
    const batchLimit = getGoogleTasksBulkBatchLimit_();
    if (batchLimit > 0 && processedThisChunk >= batchLimit) {
      return 'batch';
    }
    if (Date.now() - syncStartedAt > maxRuntime) {
      return 'time';
    }
    return null;
  }

  /**
  * Saves resume state, ends sync log, optionally schedules the next batch.
  */
  function pauseBulkGoogleTasksSync_(props, phase, resumeRow, reason, detail) {
    props.setProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_ROW, String(resumeRow));
    props.setProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_PHASE, phase);
    const msg = detail || ('paused ' + phase + ' at row ' + resumeRow + ' (' + reason + ')');
    logGoogleTasksSync_(msg);
    logGoogleTasksSync_(
      'Run syncExistingTasksToGoogle() again, or wait for auto-continue (~1 min) if GOOGLE_TASKS_AUTO_CONTINUE_BULK is true.'
    );
    endGoogleTasksSync_('paused — ' + reason);
    scheduleGoogleTasksBulkSyncContinue_();
  }

  function cancelGoogleTasksBulkSyncContinueTriggers_() {
    try {
      ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === 'continueGoogleTasksBulkSync_') {
          ScriptApp.deleteTrigger(trigger);
        }
      });
    } catch (e) {
      // non-fatal
    }
  }

  /**
  * Schedules the next bulk batch (~1 min). Handler: continueGoogleTasksBulkSync_()
  */
  function scheduleGoogleTasksBulkSyncContinue_() {
    if (CONFIG.GOOGLE_TASKS_AUTO_CONTINUE_BULK === false) {
      logGoogleTasksSync_('Auto-continue off — run syncExistingTasksToGoogle() for the next batch.');
      return;
    }
    try {
      cancelGoogleTasksBulkSyncContinueTriggers_();
      const delayMs = Math.max(60000, CONFIG.GOOGLE_TASKS_AUTO_CONTINUE_DELAY_MS || 60000);
      ScriptApp.newTrigger('continueGoogleTasksBulkSync_')
        .timeBased()
        .after(delayMs)
        .create();
      Logger.log(
        '[GoogleTasks] Next batch scheduled in ~' + Math.round(delayMs / 1000) + 's (continueGoogleTasksBulkSync_)'
      );
    } catch (error) {
      Logger.log('scheduleGoogleTasksBulkSyncContinue_ Error: ' + error.toString());
      logGoogleTasksSync_('Could not schedule auto-continue — run syncExistingTasksToGoogle() manually.');
    }
  }

  /**
  * Time-driven entry: resume bulk sync next batch when a prior run paused.
  */
  function continueGoogleTasksBulkSync_() {
    if (!isGoogleTasksBulkSyncPaused_()) {
      Logger.log('continueGoogleTasksBulkSync_: nothing to resume.');
      cancelGoogleTasksBulkSyncContinueTriggers_();
      return;
    }
    Logger.log('continueGoogleTasksBulkSync_: starting next batch…');
    syncAllGoogleTasksFromDevTracker();
  }

  /**
  * Progress snapshot for large bulk runs (2000+ subs).
  * @returns {{paused: boolean, phase: string, resumeRow: string, subTotal: string, lastLog: string}}
  */
  function getGoogleTasksBulkSyncStatus() {
    const props = PropertiesService.getScriptProperties();
    return {
      paused: isGoogleTasksBulkSyncPaused_(),
      phase: props.getProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_PHASE) || '',
      resumeRow: props.getProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_ROW) || '',
      subTotal: props.getProperty(SCRIPT_PROP_GOOGLE_SYNC_SUB_TOTAL) || '',
      lastLog: getGoogleTasksLastSyncLog()
    };
  }

  /**
  * @returns {boolean} True when a prior bulk sync paused mid-sheet and saved resume state.
  */
  function isGoogleTasksBulkSyncPaused_() {
    try {
      return !!PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GOOGLE_SYNC_RESUME_PHASE);
    } catch (e) {
      return false;
    }
  }

  /**
  * Runs bulk sync in chunks. Each call = one batch (avoids 6-minute timeout).
  * With GOOGLE_TASKS_AUTO_CONTINUE_BULK, the next batch is scheduled automatically.
  * @param {number} [maxRuns=50] Only loops in one execution when auto-continue is false
  * @returns {number} Batches executed this call
  */
  function runBulkGoogleTasksSyncUntilComplete_(maxRuns) {
    maxRuns = maxRuns || 50;
    if (!isGoogleTasksApiEnabled_()) {
      return 0;
    }

    let runs = 0;
    for (let i = 0; i < maxRuns; i++) {
      runs++;
      Logger.log('[BulkSyncLoop] Batch ' + runs + ' starting…');
      syncAllGoogleTasksFromDevTracker();
      if (!isGoogleTasksBulkSyncPaused_()) {
        Logger.log('[BulkSyncLoop] All batches complete after ' + runs + ' run(s) this session.');
        cancelGoogleTasksBulkSyncContinueTriggers_();
        return runs;
      }
      Logger.log('[BulkSyncLoop] Batch ' + runs + ' paused — progress saved.');
      if (CONFIG.GOOGLE_TASKS_AUTO_CONTINUE_BULK !== false) {
        Logger.log('[BulkSyncLoop] Auto-continue scheduled — do not re-run manually unless it stalls.');
        return runs;
      }
      Utilities.sleep(2000);
    }

    Logger.log('[BulkSyncLoop] Still paused after ' + runs + ' batch(es). Run syncExistingTasksToGoogle() again.');
    return runs;
  }

  /**
  * Deletes every task (open + completed) in one Google Task list.
  * @param {string} listId
  * @returns {number}
  */
  function deleteAllTasksInGoogleList_(listId) {
    let deleted = 0;
    let pageToken = null;

    do {
      const opts = { showCompleted: true, showHidden: true, maxResults: 100 };
      if (pageToken) {
        opts.pageToken = pageToken;
      }

      let response;
      try {
        response = Tasks.Tasks.list(listId, opts);
      } catch (error) {
        Logger.log('deleteAllTasksInGoogleList_ list Error (' + listId + '): ' + error.toString());
        break;
      }

      const items = response.items || [];
      for (let i = 0; i < items.length; i++) {
        const task = items[i];
        if (!task || !task.id) {
          continue;
        }
        try {
          Tasks.Tasks.remove(listId, task.id);
          deleted++;
        } catch (removeErr) {
          Logger.log('Failed to delete task ' + task.id + ': ' + removeErr.toString());
        }
      }

      pageToken = response.nextPageToken || null;
    } while (pageToken);

    return deleted;
  }

  /**
  * Deletes all tasks from every list named in CONFIG.STATUS_LABELS (tracker lists only).
  * @returns {number}
  */
  function deleteAllTasksFromTrackerStatusLists_() {
    const map = refreshGoogleTaskListCache_();
    const listIds = {};
    Object.keys(map).forEach(function (key) {
      listIds[map[key]] = true;
    });

    let totalDeleted = 0;
    for (const listId in listIds) {
      if (!Object.prototype.hasOwnProperty.call(listIds, listId)) {
        continue;
      }
      const deleted = deleteAllTasksInGoogleList_(listId);
      totalDeleted += deleted;
      Logger.log('Deleted ' + deleted + ' task(s) from list ' + listId);
      Utilities.sleep(200);
    }

    return totalDeleted;
  }

  /**
  * Clears Google Task ID column V on Dev Tracker and Dev Tracker Sub (data rows only).
  */
  function clearAllGoogleTaskIdColumns_() {
    const mainSheet = getMainTrackerSheet_();
    const subSheet = getSubTrackerSheet_();

    clearProjectColumnCaches_();
    taskIdToListIdCache_ = {};

    if (mainSheet) {
      const col = findMainGoogleTaskIdColumn_(mainSheet);
      const last = mainSheet.getLastRow();
      if (last >= 2) {
        mainSheet.getRange(2, col, last - 1, 1).clearContent();
      }
    }

    if (subSheet) {
      const col = findSubGoogleTaskIdColumn_(subSheet);
      const last = subSheet.getLastRow();
      if (last >= 2) {
        subSheet.getRange(2, col, last - 1, 1).clearContent();
      }
      const nestCol = findSubNestStatusColumn_(subSheet);
      if (last >= 2) {
        subSheet.getRange(2, nestCol, last - 1, 1).clearContent();
      }
    }

    Logger.log('Cleared Google Task ID columns (V) and Nest Status (X) on Dev Tracker Sub; Dev Tracker column V.');
  }

  /**
  * DESTRUCTIVE — wipes all tasks from tracker status lists, clears column V, recreates from the sheet.
  *
  * Deletes tasks only from lists matching CONFIG.STATUS_LABELS (e.g. Pipeline / To Do, In Progress…).
  * Does NOT delete tasks in your personal @default list or other unrelated lists.
  *
  * Run from Apps Script (open spreadsheet first):
  *   resetAllGoogleTasksAndRecreate()
  */
  function resetAllGoogleTasksAndRecreate() {
    if (!authorizeGoogleTasksAccess()) {
      Logger.log('Add Services (+) > Google Tasks API, then run resetAllGoogleTasksAndRecreate() again.');
      return;
    }

    Logger.log('=== RESET ALL GOOGLE TASKS (tracker status lists) ===');
    clearGoogleTasksSyncResumeState_();

    const deleted = deleteAllTasksFromTrackerStatusLists_();
    clearAllGoogleTaskIdColumns_();

    Logger.log('Deleted ' + deleted + ' task(s). Creating fresh tasks from spreadsheet...');
    runBulkGoogleTasksSyncUntilComplete_();
    Logger.log('=== RESET COMPLETE — check Google Tasks and column V ===');
  }
