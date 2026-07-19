# Team Project Tracker — Google Apps Script

Automation and development project for the **Team Project Tracker** Google Spreadsheet. This repository holds the Apps Script source, development rules, and documentation so the team and AI assistants can work consistently.

**Spreadsheet:** Team Project Tracker  
**Schema reference:** [`SPREADSHEET_SCHEMA.md`](SPREADSHEET_SCHEMA.md)

---

## Who This Is For

| Role | Start here |
|------|------------|
| **Admin / owner** | [First-time setup](#first-time-setup) → run `setupApp()` |
| **Developer** | [Project structure](#project-structure) + [`.cursorrules`](.cursorrules) |
| **Anyone editing the spreadsheet** | [`SPREADSHEET_SCHEMA.md`](SPREADSHEET_SCHEMA.md) — what each sheet and column means |
| **BA / Product / Umbraco team leads** | Task tracker sheets + [Sheet overview](#spreadsheet-sheets) below |

---

## Spreadsheet Sheets

The **Team Project Tracker** workbook contains these tabs:

### Task trackers (active work)

| Sheet | Team |
|-------|------|
| BA Task Tracker | Business Analysis |
| BA Sub Task Tracker | BA sub-tasks |
| Product Task Tracker | Product |
| Product Sub Task Tracker | Product sub-tasks |
| Umbraco Task Tracker | Umbraco |
| Umbraco Sub Task Tracker | Umbraco sub-tasks |

### Completed archives

| Sheet | Team |
|-------|------|
| BA Completed Task | Completed BA tasks |
| Product Completed Task | Completed Product tasks |
| Umbraco Completed Task | Completed Umbraco tasks |

### Configuration

| Sheet | Purpose |
|-------|---------|
| Matching Sheet | Links each team's main sheet → sub-task sheet → completed sheet |
| 📧 Email Settings | TO/CC/BCC recipients, alert schedule, and which task sheet to monitor |

All **main task** and **completed** sheets share one column layout. **Sub-task** sheets use a different layout (includes `Main Task Name` and `Ticket Id`). Full details in [`SPREADSHEET_SCHEMA.md`](SPREADSHEET_SCHEMA.md).

### Matching Sheet — single source of truth for team tabs

All team sheet names are read from the **Matching Sheet** at runtime. **No team tab names are hardcoded in script files.**

| # | MainSheet | SubSheet | Completed Task |
|---|-----------|----------|----------------|
| 1 | BA Task Tracker | BA Sub Task Tracker | BA Completed Task |
| 2 | Product Task Tracker | Product Sub Task Tracker | Product Completed Task |
| 3 | Umbraco Task Tracker | Umbraco Sub Task Tracker | Umbraco Completed Task |

**If a user renames a sheet:** update the matching row on the Matching Sheet — no code change needed.  
**If the Matching Sheet tab itself is renamed:** update `MATCHING_SHEET_NAME` in `AppConfig.gs` only.

`setupApp()` validates that every tab listed on the Matching Sheet exists in the workbook.

### Sub-task link, field sync & effort rollup (all teams via Matching Sheet)

**Match rule:** Main `Task` (col **C**) = Sub `Main Task Name` (col **B**)

#### Main sheet — when `Task` column is edited

1. Finds matching sub rows on the paired sub sheet (from Matching Sheet).
2. Syncs parent fields to sub tasks (Project Name, Page Url, Priority, Assign, dates, etc.).
3. **Rolls up** total `Estimated Effort` and `Actual Effort` from all matching subs → main sheet.

#### Sub sheet — when `Main Task Name`, `Estimated Effort`, `Actual Effort`, or `Project Name` is edited

1. Finds the main row where `Task` matches `Main Task Name`.
2. If no main match → **no update**.
3. If matched → corrects sub fields from main, then rolls up effort totals to main.

**Example:** Main task `Diriyah Season Handover` + sub with `Main Task Name` = `Diriyah Season Handover`, Est=1, Actual=1 → main sheet gets Est=1, Actual=1.

**Not synced to subs:** Status, Estimated Effort, Actual Effort (subs keep their own values; main gets the sum).

**Future teams:** Add a row to the Matching Sheet — sync and rollup work automatically.

**Manual test functions** (use exact tab names from your Matching Sheet):
- `runSubTaskSyncForRow('Umbraco Task Tracker', 2)` — sync + rollup for main row 2
- `runSubTaskSyncForTaskName('Umbraco Task Tracker', 'Diriyah Season Handover')` — by task name
- `runEffortRollupForTaskName('Umbraco Task Tracker', 'Diriyah Season Handover')` — effort only
- `runSubRowSyncAndRollup('Umbraco Sub Task Tracker', 2)` — sync + rollup from sub row
- `MatchingSheetModule.getMappings()` — view all loaded team mappings
- `MatchingSheetModule.validateMappingsExist()` — verify all tabs exist

### Delayed move to completed sheet (Status column I / J)

When a user selects a **completion status** on **Umbraco Task Tracker** (column **I**) or **Umbraco Sub Task Tracker** (column **J**):

| Completion statuses (trigger move after 24 hours) |
|---------------------------------------------------|
| Completed |
| Closed / Cancelled |
| Closed – Billable |
| Closed – Not Proceeding |

**Flow:**
1. User selects a completion status → row is queued (hidden `_Completion Queue` sheet).
2. After **24 hours** (if status is still a completion value), the row moves to the team's **Completed Task** sheet (read from Matching Sheet).
3. All cell values are preserved; dropdown status values copy exactly.
4. If the user changes status away from a completion value before 24 hours, the queue entry is cancelled.

Works for **every team row** on the Matching Sheet.

**Triggers:** `onEdit` queues the row; hourly time-driven trigger runs `processScheduledCompletions_`.

**Manual test functions:**
- `runScheduleCompletionForRow('Umbraco Task Tracker', 2)` — queue row 2 (use tab name from Matching Sheet)
- `runProcessCompletionQueue()` — process all due rows now
- Set `IS_TESTING = true` in `AppConfig.gs` to use a **1-minute** delay instead of 24 hours

### Matching Sheet (how teams connect)

> Tab names below are **examples**. The script reads the live values from the Matching Sheet — not from this README.

| # | Main Sheet | Sub Sheet | Completed Sheet |
|---|------------|-----------|-----------------|
| 1 | BA Task Tracker | BA Sub Task Tracker | BA Completed Task |
| 2 | Product Task Tracker | Product Sub Task Tracker | Product Completed Task |
| 3 | Umbraco Task Tracker | Umbraco Sub Task Tracker | Umbraco Completed Task |

### Email alerts (status & priority)

Configured on **📧 Email Settings** (one row per monitored sheet):

| Setting | Current value |
|---------|---------------|
| TO | dragonkarthi00m@gmail.com |
| BCC | narutokarthi00m@gmail.com, sowndar@vajraglobal.com |
| Data sheet | Umbraco Task Tracker |
| Timezone | IST |
| Alert times | 9:30 AM, 13:00, 15:00 |
| Notification types | Requirements Not Clear, On Hold / Blocked, Waiting for Client Input, Under Review & Testing, Changes Requested / Rework, **High** |

**When emails are sent:**

1. **Scheduled digest** — at each alert time, lists all tasks on the data sheet whose **Status** or **Priority** matches Notification Type.
2. **Immediate** — when a user edits Status or Priority to a matching value.

**Fields included in email (main task sheet):**

Project, Task, Status, Priority, Assign, **Support Needed / Blocker**, **Latest Update / History**, **Functionality**, End Date

**Fields included in email (sub-task sheet):**

Project, **Main Task Name**, Sub Task, **Ticket Id**, **Account**, Status, Priority, Assign, **Support / Blocker**, **Latest Update**, **Functionality**, End Date

**Sub-task monitoring:** Add a **second row** on 📧 Email Settings with `DATA SHEET NAME` = `Umbraco Sub Task Tracker` (or BA/Product sub sheets). Each row monitors one sheet independently.

**Manual tests:**
- `runSendNotificationDigestNow()` — send digest immediately
- `runParseNotificationTypes('Requirements Not Clear,On Hold / Blocked,High')` — test parsing

Add more rows on 📧 Email Settings to monitor BA/Product sheets with different rules. Re-run **`setupApp()`** to install the 30-minute notification trigger.

### Error log sheet

All automation **errors** and **warnings** are written to the **Error** sheet (auto-created on `setupApp()`).

| Column | Purpose |
|--------|---------|
| Logged At | Date & time of the error |
| Severity | ERROR / WARNING / INFO |
| Module + Function | Where it happened in code |
| Sheet Name + Row # | Which sheet/row to fix |
| Message + Details | What went wrong (Details has full JSON context) |
| Resolved + Fix Notes | **You fill in** after fixing |

**Manual tests:** `runEnsureErrorSheet()` · `runTestErrorLog()`

Routine skips (no task match, empty cells) are **not** logged — only real errors and warnings.

---

## First-Time Setup

1. **Open the Apps Script editor**  
   In your Google Sheet: **Extensions → Apps Script**, or open the bound script project directly.

2. **Copy or sync project files**  
   Add the `.gs` files from this repo into the script project (via clasp or manual copy).

3. **Configure environment**  
   In `AppConfig.gs`, set:
   - `PRODUCTION_SPREADSHEET_ID`
   - `TEST_SPREADSHEET_ID` (for safe testing)
   - `IS_TESTING` — keep `false` for live use

4. **Confirm spreadsheet structure**  
   The **Team Project Tracker** sheets and columns should match [`SPREADSHEET_SCHEMA.md`](SPREADSHEET_SCHEMA.md).  
   If you add a new sheet or column in the spreadsheet, **update that file too**.

5. **Run setup (admin only, once)**  
   In the script editor, select `setupApp` and click **Run**.  
   This installs **installable triggers** so automation runs for all users without repeated permission prompts.

6. **Authorize the script**  
   The admin who runs `setupApp()` must approve the requested scopes (Spreadsheet, Triggers, etc.).

---

## Project Structure

```
Google Scrip/
├── README.md
├── SPREADSHEET_SCHEMA.md
├── .cursorrules
├── AppConfig.gs
├── SheetColumns.gs
├── SpreadsheetUtils.gs
├── MatchingSheetModule.gs
├── SubTaskSyncModule.gs
├── TaskEffortRollupModule.gs
├── EmailSettingsModule.gs
├── NotificationTypeModule.gs
├── TaskNotificationModule.gs
├── NotificationService.gs
├── NotificationHandler.gs
├── ErrorLogModule.gs
├── TaskStatusModule.gs
├── CompletionQueueModule.gs
├── CompletionMoveModule.gs
├── CompletionTaskHandler.gs
├── TaskSyncHandler.gs
├── RenewalReminderHandler.gs
├── TrainBookingHistoryModule.gs
├── TrainIrctcParseModule.gs
├── TrainIrctcImportModule.gs
├── TrainIrctcImportHandler.gs
├── TriggerManager.gs
└── Setup.gs
```

### Train tracking (IRCTC email import)

| Sheet | Role |
|-------|------|
| **Train Route** | Your recurring templates (reminders later) |
| **Train Booking History** | IRCTC confirmations land here |
| **Train Completed Journeys** | Archive (manual for now) |
| **Match Place** *(optional)* | Canonical station labels in column A |

**How to run**
1. Ensure **Train Booking History** headers match schema (or run `setupApp()` / `runExtractIrctcEmails()` — headers are auto-created).
2. Re-run **`setupApp()`** to install the hourly `processScheduledIrctcImport_` trigger (needs Gmail permission).
3. Manual test: `testParseSampleIrctcEmail()` then `runExtractIrctcEmails()` (imports **unread** matching IRCTC mail).

Override Gmail search with Script Property `IRCTC_GMAIL_QUERY` if needed.

**Train route reminders**
- Re-run **`setupApp()`** to install `processScheduledTrainReminders_` (every 30 minutes).
- Manual: **`runTrainRemindersNow()`** (must be within an Alert Time window ±15 min, e.g. 9:30 / 13:00 / 15:00).
- BOOK reminders are skipped when From/To/date already exist on **Train Booking History**.

**Completed journeys archive**
- Daily: `processScheduledTrainArchive_` moves rows whose **Date of Journey / Travel Date** is before today from **Train Booking History** → **Train Completed Journeys**.
- Manual: **`runArchiveTrainJourneysNow()`**.

---

## Important Concepts

### Installable triggers (not simple triggers)

This project uses **programmatic installable triggers** (`ScriptApp.newTrigger()`), not basic `onEdit` / `onOpen` simple triggers. That way:

- Automation runs under the admin who ran `setupApp()`
- Other spreadsheet users are not asked to authorize the script on every edit

Re-run `setupApp()` after adding new trigger types in code.

### Test vs production

- Set `IS_TESTING = true` in `AppConfig.gs` only when developing against the **test** spreadsheet.
- Test triggers and test code must **never** write to production data.
- See `.cursorrules` §3 for full rules.

### Namespace modules (no `import` / `export`)

Google Apps Script does not support ES modules. Code is organized as global namespace objects:

```javascript
const NotificationService = (function () {
  'use strict';
  // ...
  return { send: send_ };
})();
```

---

## Spreadsheet Documentation (Required)

**[`SPREADSHEET_SCHEMA.md`](SPREADSHEET_SCHEMA.md)** is the official record of:

- Every sheet (tab) name and purpose  
- Every column: index, name, type, required or not, which script uses it  
- Change history when structure evolves  

### When you must update it

| You do this… | Also do this… |
|--------------|----------------|
| Add a new sheet tab | Add a section in `SPREADSHEET_SCHEMA.md` |
| Add a new column | Add a row to that sheet’s column table |
| Rename a sheet or column | Update the doc + add a Change History entry |
| Add a column index in code | Match the index in `SPREADSHEET_SCHEMA.md` |

Cursor AI is configured (via `.cursorrules`) to update this file whenever sheets or fields change in code.

---

## Development Guidelines (Summary)

Full rules live in [`.cursorrules`](.cursorrules). Highlights:

| Topic | Rule |
|-------|------|
| File names | PascalCase, purpose-driven (`NotificationService.gs`, not `utils.gs`) |
| Features | Group by domain; don’t scatter notification logic across unrelated files |
| Errors | Null-check inputs; `try/catch`; log with `console.error()`; don’t crash the main flow |
| Performance | Batch `getValues()` / `setValues()`; respect 6-minute execution limit |
| Reuse | Search existing code before adding new modules |

---

## Common Tasks

### Add a new feature

1. Read `SPREADSHEET_SCHEMA.md` and existing `.gs` files for overlap.  
2. Add or update sheets/columns in the spreadsheet **and** in `SPREADSHEET_SCHEMA.md`.  
3. Create feature files with a shared prefix (e.g. `NotificationService.gs`).  
4. Wire triggers in `TriggerManager.gs` and call `setupApp()` again.  
5. Test with `IS_TESTING = true` on the test spreadsheet first.

### Add a new sheet or field

1. Create the tab/headers in Google Sheets.  
2. Document in [`SPREADSHEET_SCHEMA.md`](SPREADSHEET_SCHEMA.md) (columns table + Change History).  
3. Add column constants in code if needed (e.g. `SheetColumns.gs`).  
4. Implement logic using those constants, not magic numbers without documentation.

### Debug a failed run

1. **Executions:** Apps Script editor → **Executions** (or Cloud logs for `console.error`).  
2. Confirm triggers exist: run `TriggerManager` helpers or check **Triggers** in the editor.  
3. Confirm sheet/column names match `SPREADSHEET_SCHEMA.md` exactly.

---

## Code Review Checklist

Before merging or deploying:

- [ ] `setupApp()` run (or re-run) if triggers changed  
- [ ] `SPREADSHEET_SCHEMA.md` updated for any structural change  
- [ ] Tested on test spreadsheet when `IS_TESTING = true`  
- [ ] No simple triggers used for shared automation  
- [ ] Errors logged; primary workflow still completes on partial failures  

---

## Links

- [Google Apps Script documentation](https://developers.google.com/apps-script)
- [Installable triggers guide](https://developers.google.com/apps-script/guides/triggers/installable)
- [Spreadsheet service reference](https://developers.google.com/apps-script/reference/spreadsheet)

---

## Questions?

- **What does each column mean?** → `SPREADSHEET_SCHEMA.md`  
- **How should code be written?** → `.cursorrules`  
- **How do I set up the project?** → [First-time setup](#first-time-setup) above
