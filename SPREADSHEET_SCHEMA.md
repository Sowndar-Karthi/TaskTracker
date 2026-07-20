# Spreadsheet Schema — Team Project Tracker

> **Living document** — Update this file whenever you add, rename, or remove a sheet or field.  
> Code and spreadsheet structure must stay in sync. See `.cursorrules` §9 for rules.

**Spreadsheet name:** Team Project Tracker  
**Last updated:** 2026-07-19  
**Production spreadsheet ID:** _(set in `AppConfig.gs`)_  
**Test spreadsheet ID:** _(set in `AppConfig.gs`)_

---

## How to Use This File

1. **Before adding a field in code** — Check if the field already exists here.
2. **When creating a new sheet or column** — Add it to this file in the same change as the spreadsheet/code update.
3. **When renaming** — Update the table and log the old name in [Change History](#change-history).
4. **Column index** — Always 1-based (column A = 1, B = 2, …) to match Apps Script `getRange(row, col)`.
5. **Tab names are case-sensitive** — Must match Google Sheets exactly (including emoji in `📧 Email Settings`).

---

## Sheets Overview

| Tab Name | Purpose | Header Row | Status |
|----------|---------|------------|--------|
| Umbraco Task Tracker | Main Umbraco team tasks | 1 | Active |
| Umbraco Sub Task Tracker | Umbraco sub-tasks linked to main tasks | 1 | Active |
| BA Task Tracker | Main BA team tasks | 1 | Active |
| BA Sub Task Tracker | BA sub-tasks linked to main tasks | 1 | Active |
| Product Task Tracker | Main Product team tasks | 1 | Active |
| Product Sub Task Tracker | Product sub-tasks linked to main tasks | 1 | Active |
| Umbraco Completed Task | Archived completed Umbraco tasks | 1 | Active |
| BA Completed Task | Archived completed BA tasks | 1 | Active |
| Product Completed Task | Archived completed Product tasks | 1 | Active |
| Matching Sheet | Maps main → sub → completed sheets per team | 1 | Active |
| 📧 Email Settings | Email alerts: recipients, data sheet, schedule | 1 | Active |
| _Completion Queue | Hidden — delayed completion move queue (auto-created by script) | 1 | System |
| Error | Automation error & warning log for debugging and fixes | 1 | System |
| Train Route | Recurring train travel templates (day, stations, reminders) | 1 | Active |
| Train Booking History | Bookings imported from IRCTC email + manual | 1 | Active |
| Train Completed Journeys | Archived completed train journeys (same columns as Booking History) | 1 | Active |
| Match Place | Optional — canonical station labels for IRCTC From/To matching | 1 | Optional |

---

## Domain Groups (Matching Sheet)

The **Matching Sheet** defines how main task trackers, sub-task trackers, and completed-task archives relate:

| # | MainSheet | SubSheet | Completed Task |
|---|-----------|----------|----------------|
| 1 | BA Task Tracker | BA Sub Task Tracker | BA Completed Task |
| 2 | Product Task Tracker | Product Sub Task Tracker | Product Completed Task |
| 3 | Umbraco Task Tracker | Umbraco Sub Task Tracker | Umbraco Completed Task |

Scripts **must** read all team sheet names from this table at runtime. Do **not** hardcode team tab names in `.gs` files. If a user renames a task sheet, update the matching row here — no code changes required.

**Code entry point:** `MatchingSheetModule.getMappings()`  
**Bootstrap config only:** `AppConfig.getMatchingSheetName()` → default `'Matching Sheet'` (change in `AppConfig.gs` only if this config tab itself is renamed)

### Features driven by Matching Sheet

| Feature | How it uses Matching Sheet |
|---------|---------------------------|
| Sub-task sync | `MainSheet` → `SubSheet` when Task column edited |
| Effort rollup | Sums sub `Estimated`/`Actual` Effort → main sheet by task name match |
| Completion move | `MainSheet` / `SubSheet` → `Completed Task` after 24h |
| Trigger routing | `getSheetRole()` detects if edited tab is main, sub, or completed |

Run `setupApp()` after changing Matching Sheet rows — it validates all tab names exist.

---

## Shared Column Layout — Main Task Trackers

The following **6 main task sheets** and **3 completed archive sheets** share the **same column headers** (row 1):

**Main task sheets:**
- `Umbraco Task Tracker`
- `BA Task Tracker`
- `Product Task Tracker`

**Completed archive sheets:**
- `Umbraco Completed Task`
- `BA Completed Task`
- `Product Completed Task`

**Header row:** 1  
**Data starts at row:** 2

### Columns

| Col | Index | Field Name | Type | Required | Description | Used By |
|-----|-------|------------|------|----------|-------------|---------|
| A | 1 | Project Name | Text | Yes | Name of the project | Task modules |
| B | 2 | Website Url | URL / Text | No | Primary website URL for the project | Task modules |
| C | 3 | Task | Text | Yes | Task title or description | Task modules |
| D | 4 | Priority | Text | No | Task priority level | Task modules |
| E | 5 | Assign | Text | No | Assigned team member(s) | Task modules, alerts |
| F | 6 | Start Date | Date | No | Task start date | Task modules |
| G | 7 | End Date | Date | No | Task due / end date | Task modules |
| H | 8 | As Per WO | Text / Boolean | No | Whether task aligns with work order | Task modules |
| I | 9 | Status | Text | Yes | Current task status (dropdown) | Task modules, completion flow |
| J | 10 | Env | Text | No | Environment (e.g. Dev, Staging, Prod) | Task modules |
| K | 11 | Estimated Effort | Number / Text | No | Planned effort estimate | Task modules |
| L | 12 | Actual Effort | Number / Text | No | Actual effort spent | Task modules |
| M | 13 | Informed | Text | No | Stakeholders informed about the task | Task modules |
| N | 14 | Support Needed / Blocker | Text | No | Blockers or support requests | Task modules, alerts |
| O | 15 | Latest Update / History | Text | No | Chronological updates and notes | Task modules |
| P | 16 | Page Url | URL / Text | No | Page-level URL for the task | Task modules, SubTaskSyncModule |
| Q | 17 | Functionality | Text | No | Functional area or feature name | Task modules |
| R | 18 | Mail Subject | Text | No | External email subject line | NotificationService |
| S | 19 | Internal Mail Subject | Text | No | Internal email subject line | NotificationService |
| T | 20 | _(reserved)_ | — | No | Empty column in current layout | — |
| U | 21 | _(reserved)_ | — | No | Empty column in current layout | — |
| V | 22 | Gmail Thread ID | Text | No | Linked Gmail thread identifier | Email integration |
| W | 23 | Google Task ID | Text | No | Linked Google Task identifier | Google Tasks integration |

### Notes — Main Task Trackers
- **Completed sheets** (`Umbraco Completed Task`, `BA Completed Task`, `Product Completed Task`) use this **same column layout** as main task trackers.
- **Website Url vs Page Url:** Column B is the project website URL; column P is the specific page URL for the task.
- **Empty columns:** Columns T and U appear blank between `Internal Mail Subject` and `Gmail Thread ID` in the current sheet layout.
- **Sub-task link key:** The `Task` column (C) on main sheets matches `Main Task Name` (B) on sub-task sheets. See [Sub Task Tracker layout](#shared-column-layout--sub-task-trackers).
- **Completion move:** Selecting a completion status queues the row for move to the completed sheet after 24 hours. See [Status values & completion flow](#status-values--completion-flow).
- Do not insert columns without updating this document and `SheetColumns.gs`.

---

## Status Values & Completion Flow

### Allowed status values (dropdown)

Used on **Status** column for main task sheets (col **I**) and sub-task sheets (col **J**):

| Status |
|--------|
| Pipeline / To Do |
| Requirements Not Clear |
| In Progress |
| On Hold / Blocked |
| Waiting for Client Input |
| Under Review & Testing |
| Changes Requested / Rework |
| Ready for Execution |
| Completed |
| Closed / Cancelled |
| Closed – Billable |
| Closed – Not Proceeding |

### Completion statuses (auto-move after 24 hours)

When any of these statuses is selected, the row is queued and moved to the team's **Completed Task** sheet (from Matching Sheet) **24 hours later** if the status is still a completion value:

- `Completed`
- `Closed / Cancelled`
- `Closed – Billable`
- `Closed – Not Proceeding`

| Source sheet | Status column | Completed destination |
|--------------|---------------|------------------------|
| Umbraco Task Tracker | I (9) | Umbraco Completed Task |
| Umbraco Sub Task Tracker | J (10) | Umbraco Completed Task |
| BA Task Tracker | I (9) | BA Completed Task |
| BA Sub Task Tracker | J (10) | BA Completed Task |
| Product Task Tracker | I (9) | Product Completed Task |
| Product Sub Task Tracker | J (10) | Product Completed Task |

### Move rules
- Row values are copied exactly (including dropdown status text).
- Main sheet rows copy 1:1 to the completed sheet (same column layout).
- Sub-task rows are mapped to the completed sheet layout; `Main Task Name` and `Ticket Id` are preserved in `Latest Update / History`.
- Source row is **deleted** after a successful move.
- Queue is stored on hidden sheet `_Completion Queue`.
- Changing status away from a completion value **cancels** the scheduled move.
- Re-selecting a completion status **resets** the 24-hour timer.

### Internal queue sheet: `_Completion Queue`

| Col | Field | Description |
|-----|-------|-------------|
| A | Source Sheet | Origin tab name |
| B | Sheet Role | `main` or `sub` |
| C | Row Signature | Unique row identifier |
| D | Status | Completion status when queued |
| E | Scheduled At | Timestamp when status was set |
| F | Completed Sheet | Destination tab name |

---

## Sheet: Error

**Purpose:** Central log of automation **errors** and **warnings** so the team can review, fix sheet issues, and mark items resolved.

**Tab name:** `Error` (auto-created by `setupApp()` or `runEnsureErrorSheet()` if missing)  
**Header row:** 1  
**Log rows start at:** 2

### Columns

| Col | Index | Field Name | Type | Required | Description | Filled by |
|-----|-------|------------|------|----------|-------------|-----------|
| A | 1 | Logged At | Date/Time | Yes | When the error occurred | Script (auto) |
| B | 2 | Severity | Text | Yes | `ERROR`, `WARNING`, or `INFO` | Script (auto) |
| C | 3 | Module | Text | Yes | Source module (e.g. `SubTaskSyncModule`) | Script (auto) |
| D | 4 | Function | Text | Yes | Function name where error occurred | Script (auto) |
| E | 5 | Sheet Name | Text | No | Task sheet involved, if any | Script (auto) |
| F | 6 | Row # | Number | No | Row number on the sheet, if applicable | Script (auto) |
| G | 7 | Task / Key | Text | No | Task name or row signature for context | Script (auto) |
| H | 8 | Message | Text | Yes | Short description of the issue | Script (auto) |
| I | 9 | Details | Text | No | JSON context (error message, stack, parameters) | Script (auto) |
| J | 10 | Trigger Source | Text | No | `onEdit`, `scheduled-completion`, `scheduled-notification`, `manual`, `system` | Script (auto) |
| K | 11 | Resolved | Text | No | Mark `Yes` when fixed | **User** |
| L | 12 | Fix Notes | Text | No | What was done to fix the issue | **User** |

### What gets logged

| Severity | Examples |
|----------|----------|
| **ERROR** | Sync failed, email send failed, sheet not found, move to completed failed |
| **WARNING** | Matching Sheet tab missing, lock not acquired, validation failed on setup |
| **INFO** | Optional informational entries (rare) |

### What is NOT logged

Routine skips (no match, empty task name, not alert time) are **not** logged — those are normal operation, not errors.

### How to use

1. Check **Error** sheet when automation seems broken.
2. Read **Message** + **Details** to find root cause.
3. Fix the sheet/config issue.
4. Set **Resolved** = `Yes` and add **Fix Notes**.

### Manual test

```javascript
runEnsureErrorSheet();  // create sheet + headers
runTestErrorLog();      // write one test row
```

---

The following **3 sub-task sheets** share a **different column layout** from main task sheets:

- `Umbraco Sub Task Tracker`
- `BA Sub Task Tracker`
- `Product Sub Task Tracker`

**Header row:** 1  
**Data starts at row:** 2

### Columns

| Col | Index | Field Name | Type | Required | Description | Used By |
|-----|-------|------------|------|----------|-------------|---------|
| A | 1 | Project Name | Text | Yes | Project name (synced from main task) | SubTaskSyncModule |
| B | 2 | Main Task Name | Text | Yes | **Match key** — must equal main sheet `Task` column | SubTaskSyncModule |
| C | 3 | Ticket Id | Text | No | Ticket / issue identifier | Task modules |
| D | 4 | Task | Text | Yes | Sub-task title (not overwritten by sync) | Task modules |
| E | 5 | Priority | Text | No | Synced from main task when matched | SubTaskSyncModule |
| F | 6 | Assign | Text | No | Synced from main task when matched | SubTaskSyncModule |
| G | 7 | Start Date | Date | No | Synced from main task when matched | SubTaskSyncModule |
| H | 8 | End Date | Date | No | Synced from main task when matched | SubTaskSyncModule |
| I | 9 | As Per WO | Text / Boolean | No | Synced from main task when matched | SubTaskSyncModule |
| J | 10 | Status | Text | No | Sub-task status; completion trigger | CompletionMoveModule |
| K | 11 | Env | Text | No | Synced from main task when matched | SubTaskSyncModule |
| L | 12 | Account | Text | No | Account name (sub-task only, not on main sheet) | Task modules |
| M | 13 | Estimated Effort | Number / Text | No | Rolled up to main sheet (not synced main → sub) | TaskEffortRollupModule |
| N | 14 | Actual Effort | Number / Text | No | Rolled up to main sheet (not synced main → sub) | TaskEffortRollupModule |
| O | 15 | Informed | Text | No | Synced from main task when matched | SubTaskSyncModule |
| P | 16 | Support Needed / Blocker | Text | No | Synced from main task when matched | SubTaskSyncModule |
| Q | 17 | Latest Update / History | Text | No | Sub-task notes (not overwritten by sync) | Task modules |
| R | 18 | Page Url | URL / Text | No | Synced from main task `Page Url` (col P) | SubTaskSyncModule |
| S | 19 | Functionality | Text | No | Synced from main task when matched | SubTaskSyncModule |
| T | 20 | Mail Subject | Text | No | Synced from main task when matched | SubTaskSyncModule |
| U | 21 | Internal Mail Subject | Text | No | Synced from main task when matched | SubTaskSyncModule |
| V | 22 | _(reserved)_ | — | No | Empty column in current layout | — |
| W | 23 | _(reserved)_ | — | No | Empty column in current layout | — |
| X | 24 | Gmail Thread ID | Text | No | Linked Gmail thread (not overwritten by sync) | Email integration |
| Y | 25 | Google Task ID | Text | No | Linked Google Task (not overwritten by sync) | Google Tasks integration |

### Sub-Task Matching Rule

**Match key:** Main sheet `Task` (col **C**) = Sub sheet `Main Task Name` (col **B**) — case-insensitive.

Works for **every team** on the Matching Sheet (including new teams added in the future).

#### When main sheet `Task` column is edited

1. Find all sub rows where `Main Task Name` matches.
2. Sync parent fields **main → sub** (Project Name, Page Url, Priority, Assign, dates, Env, etc.).
3. **Sum** `Estimated Effort` and `Actual Effort` from all matching sub rows.
4. Write totals to main sheet `Estimated Effort` (col **K**) and `Actual Effort` (col **L**).

#### When sub sheet row is edited (`Main Task Name`, `Estimated Effort`, `Actual Effort`, `Project Name`)

1. Find main row where `Task` = sub `Main Task Name`.
2. If **no main match** → skip (no update).
3. If matched → correct sub fields from main (except effort and status).
4. Roll up effort totals from **all** matching subs → main sheet.

#### Effort rollup example

| Main Task (col C) | Sub Main Task Name (col B) | Sub Est. | Sub Actual | → Main Est. | → Main Actual |
|-------------------|----------------------------|----------|------------|-------------|---------------|
| Diriyah Season Handover | Diriyah Season Handover | 1 | 1 | **1** | **1** |
| Diriyah Season Handover | Diriyah Season Handover | 2 | 3 | **3** | **4** (sum of both subs) |

#### Not overwritten on sub tasks

`Main Task Name`, `Ticket Id`, `Account`, sub-task `Task`, `Status`, `Estimated Effort`, `Actual Effort`, `Latest Update / History`, `Gmail Thread ID`, `Google Task ID`

#### Not synced main → sub

`Website Url` (main col B), `Status`, `Estimated Effort`, `Actual Effort`

#### If no sub match

Main effort columns are **not** updated. No rollup runs.

### Notes — Sub Task Trackers
- Sub tasks can exist **before** the main task is created (pre-filled `Main Task Name`); they link when the main `Task` is added or when the sub row is edited.
- Add a new row on the **Matching Sheet** to onboard a new team — no code changes required.

---

## Sheet: Matching Sheet

**Purpose:** Configuration table that links each team's main task sheet, sub-task sheet, and completed-task archive. Used by scripts to route tasks across sheets.

**Header row:** 1  
**Data starts at row:** 2

### Columns

| Col | Index | Field Name | Type | Required | Description | Used By |
|-----|-------|------------|------|----------|-------------|---------|
| A | 1 | # | Number | Yes | Row identifier (1, 2, 3, …) | MatchingModule |
| B | 2 | MainSheet | Text | Yes | Tab name of the main task tracker | MatchingModule |
| C | 3 | SubSheet | Text | Yes | Tab name of the sub-task tracker | MatchingModule |
| D | 4 | Completed Task | Text | Yes | Tab name of the completed-task archive | MatchingModule |

### Current Data

| # | MainSheet | SubSheet | Completed Task |
|---|-----------|----------|----------------|
| 1 | BA Task Tracker | BA Sub Task Tracker | BA Completed Task |
| 2 | Product Task Tracker | Product Sub Task Tracker | Product Completed Task |
| 3 | Umbraco Task Tracker | Umbraco Sub Task Tracker | Umbraco Completed Task |

### Notes
- Tab names in columns B–D must exactly match the sheet tab names in this workbook.
- Add a new row here when onboarding a new team domain (do not hard-code only three teams in script logic).

---

## Sheet: 📧 Email Settings

**Purpose:** Configures automated email alerts: recipients, which task sheet to monitor, timezone, and alert schedule.

**Header row:** 1  
**Config value row:** 2 _(single config row in current layout)_

### Columns

| Col | Index | Field Name | Type | Required | Description | Used By |
|-----|-------|------------|------|----------|-------------|---------|
| A | 1 | TO EMAILS | Text | Yes | Primary recipient(s), comma-separated | NotificationService |
| B | 2 | CC | Text | No | CC recipient(s), comma-separated | NotificationService |
| C | 3 | BCC EMAILS | Text | No | BCC recipient(s), comma-separated | NotificationService |
| D | 4 | DATA SHEET NAME | Text | Yes | Task sheet tab name to read for alerts | NotificationService |
| E | 5 | TIMEZONE | Text | No | Timezone for scheduled alerts (e.g. `Asia/Kolkata`) | TriggerManager |
| F | 6 | Alert Time | Text | Yes | Comma-separated alert times (e.g. `9:30 AM, 13:00, 15:00`) | NotificationHandler |
| G | 7 | Notification Type | Text | Yes | Comma-separated statuses and/or priorities to alert on | NotificationTypeModule |

### Current Config Values (row 2)

| Field | Current Value |
|-------|---------------|
| TO EMAILS | `dragonkarthi00m@gmail.com` |
| CC | _(empty)_ |
| BCC EMAILS | `narutokarthi00m@gmail.com,sowndar@vajraglobal.com` |
| DATA SHEET NAME | `Umbraco Task Tracker` |
| TIMEZONE | `IST` |
| Alert Time | `9:30 AM, 13:00, 15:00` |
| Notification Type | `Requirements Not Clear, On Hold / Blocked, Waiting for Client Input, Under Review & Testing, Changes Requested / Rework, High` |

### Notification Type rules

- Comma-separated list of **status** values and/or **priority** values (`High`, `Medium`, `Low`).
- A task triggers an alert when its **Status** OR **Priority** matches any entry.
- Multi-word statuses with commas in the cell (e.g. `Under,Review & Testing`) are auto-merged to `Under Review & Testing`.
- Recommended format: `Requirements Not Clear, On Hold / Blocked, High`

### Alert behaviour

| Trigger | When |
|---------|------|
| **Scheduled digest** | At each Alert Time (in TIMEZONE), sends one email listing all matching tasks on DATA SHEET NAME |
| **Immediate alert** | When Status or Priority is edited to a matching value on DATA SHEET NAME |

### Email content by sheet type

**Main task sheet** (`* Task Tracker`): Project, Task, Status, Priority, Assign, Support Needed / Blocker, Latest Update / History, Functionality, End Date

**Sub-task sheet** (`* Sub Task Tracker`): Project, Main Task Name, Sub Task, Ticket Id, Account, Status, Priority, Assign, Support / Blocker, Latest Update, Functionality, End Date

Add a **separate Email Settings row** per sheet to monitor sub tasks (e.g. `Umbraco Sub Task Tracker`).

### Notes
- Tab name includes the emoji prefix: `📧 Email Settings` (exact match required in `getSheetByName()`).
- `DATA SHEET NAME` must match a task tracker tab name exactly.
- `Alert Time` supports multiple times separated by commas; scripts should parse and create time-driven triggers accordingly.
- Set `TIMEZONE` to avoid ambiguity for scheduled alerts (`IST` → `Asia/Kolkata`).
- Add multiple config rows to monitor different sheets with different notification rules.
- Re-run `setupApp()` after changing Alert Time schedule.

---

## Column Constants Reference

Defined in `SheetColumns.gs`.

### Main task sheets & completed sheets

| Constant | Index | Col | Field | Applies To |
|----------|-------|-----|-------|------------|
| `SheetColumns.MAIN.PROJECT_NAME` | 1 | A | Project Name | Main + completed sheets |
| `SheetColumns.MAIN.WEBSITE_URL` | 2 | B | Website Url | Main + completed sheets |
| `SheetColumns.MAIN.TASK` | 3 | C | Task | Main + completed sheets |
| `SheetColumns.MAIN.PRIORITY` | 4 | D | Priority | Main + completed sheets |
| `SheetColumns.MAIN.ASSIGN` | 5 | E | Assign | Main + completed sheets |
| `SheetColumns.MAIN.START_DATE` | 6 | F | Start Date | Main + completed sheets |
| `SheetColumns.MAIN.END_DATE` | 7 | G | End Date | Main + completed sheets |
| `SheetColumns.MAIN.AS_PER_WO` | 8 | H | As Per WO | Main + completed sheets |
| `SheetColumns.MAIN.STATUS` | 9 | I | Status | Main + completed sheets |
| `SheetColumns.MAIN.ENV` | 10 | J | Env | Main + completed sheets |
| `SheetColumns.MAIN.ESTIMATED_EFFORT` | 11 | K | Estimated Effort | Main + completed sheets |
| `SheetColumns.MAIN.ACTUAL_EFFORT` | 12 | L | Actual Effort | Main + completed sheets |
| `SheetColumns.MAIN.INFORMED` | 13 | M | Informed | Main + completed sheets |
| `SheetColumns.MAIN.SUPPORT_BLOCKER` | 14 | N | Support Needed / Blocker | Main + completed sheets |
| `SheetColumns.MAIN.LATEST_UPDATE` | 15 | O | Latest Update / History | Main + completed sheets |
| `SheetColumns.MAIN.PAGE_URL` | 16 | P | Page Url | Main + completed sheets |
| `SheetColumns.MAIN.FUNCTIONALITY` | 17 | Q | Functionality | Main + completed sheets |
| `SheetColumns.MAIN.MAIL_SUBJECT` | 18 | R | Mail Subject | Main + completed sheets |
| `SheetColumns.MAIN.INTERNAL_MAIL_SUBJECT` | 19 | S | Internal Mail Subject | Main + completed sheets |
| `SheetColumns.MAIN.GMAIL_THREAD_ID` | 22 | V | Gmail Thread ID | Main + completed sheets |
| `SheetColumns.MAIN.GOOGLE_TASK_ID` | 23 | W | Google Task ID | Main + completed sheets |

### Sub task sheets

| Constant | Index | Col | Field | Applies To |
|----------|-------|-----|-------|------------|
| `SheetColumns.SUB.PROJECT_NAME` | 1 | A | Project Name | Sub task sheets |
| `SheetColumns.SUB.MAIN_TASK_NAME` | 2 | B | Main Task Name | Sub task sheets |
| `SheetColumns.SUB.TICKET_ID` | 3 | C | Ticket Id | Sub task sheets |
| `SheetColumns.SUB.TASK` | 4 | D | Task | Sub task sheets |
| `SheetColumns.SUB.PRIORITY` | 5 | E | Priority | Sub task sheets |
| `SheetColumns.SUB.ACCOUNT` | 12 | L | Account | Sub task sheets |
| `SheetColumns.SUB.ESTIMATED_EFFORT` | 13 | M | Estimated Effort | Sub task sheets |
| `SheetColumns.SUB.ACTUAL_EFFORT` | 14 | N | Actual Effort | Sub task sheets |
| `SheetColumns.SUB.PAGE_URL` | 18 | R | Page Url | Sub task sheets |
| `SheetColumns.SUB.LATEST_UPDATE` | 17 | Q | Latest Update / History | Sub task sheets |
| `SheetColumns.SUB.GMAIL_THREAD_ID` | 24 | X | Gmail Thread ID | Sub task sheets |
| `SheetColumns.SUB.GOOGLE_TASK_ID` | 25 | Y | Google Task ID | Sub task sheets |

### Matching Sheet & Email Settings

| Constant | Index | Col | Field | Applies To |
|----------|-------|-----|-------|------------|
| `SheetColumns.MATCHING.ID` | 1 | A | # | Matching Sheet |
| `SheetColumns.MATCHING.MAIN_SHEET` | 2 | B | MainSheet | Matching Sheet |
| `SheetColumns.MATCHING.SUB_SHEET` | 3 | C | SubSheet | Matching Sheet |
| `SheetColumns.MATCHING.COMPLETED_SHEET` | 4 | D | Completed Task | Matching Sheet |
| `SheetColumns.EMAIL.TO` | 1 | A | TO EMAILS | 📧 Email Settings |
| `SheetColumns.EMAIL.CC` | 2 | B | CC | 📧 Email Settings |
| `SheetColumns.EMAIL.BCC` | 3 | C | BCC EMAILS | 📧 Email Settings |
| `SheetColumns.EMAIL.DATA_SHEET_NAME` | 4 | D | DATA SHEET NAME | 📧 Email Settings |
| `SheetColumns.EMAIL.TIMEZONE` | 5 | E | TIMEZONE | 📧 Email Settings |
| `SheetColumns.EMAIL.ALERT_TIME` | 6 | F | Alert Time | 📧 Email Settings |
| `SheetColumns.EMAIL.NOTIFICATION_TYPE` | 7 | G | Notification Type | 📧 Email Settings |
| `SheetColumns.ERROR.LOGGED_AT` | 1 | A | Logged At | Error |
| `SheetColumns.ERROR.SEVERITY` | 2 | B | Severity | Error |
| `SheetColumns.ERROR.MESSAGE` | 8 | H | Message | Error |
| `SheetColumns.ERROR.RESOLVED` | 11 | K | Resolved | Error |
| `SheetColumns.TRAIN_ROUTE.ROUTE_NAME` | 1 | A | Route Name | Train Route |
| `SheetColumns.TRAIN_ROUTE.RECURRENCE` | 5 | E | Recurrence | Train Route |
| `SheetColumns.TRAIN_ROUTE.TRAVEL_DATE` | 8 | H | Travel Date | Train Route |
| `SheetColumns.TRAIN_ROUTE.SKIP_DATES` | 14 | N | Skip Dates | Train Route |
| `SheetColumns.TRAIN_ROUTE.NOTES` | 15 | O | Notes | Train Route |
| `SheetColumns.TRAIN_BOOKING.FROM` | 1 | A | From | Train Booking History |
| `SheetColumns.TRAIN_BOOKING.PNR` | 9 | I | PNR | Train Booking History |
| `SheetColumns.TRAIN_BOOKING.GMAIL_THREAD_ID` | 17 | Q | Gmail Thread ID | Train Booking History |

### System config (code only — not team sheet names)

| Config | Location | Purpose |
|--------|----------|---------|
| `AppConfig.getMatchingSheetName()` | `AppConfig.gs` | Tab name of the Matching Sheet itself |
| `AppConfig.getEmailSettingsSheetName()` | `AppConfig.gs` | Tab name of 📧 Email Settings |
| `AppConfig.getErrorSheetName()` | `AppConfig.gs` | Tab name of Error log (`Error`) |
| `AppConfig.getCompletionQueueSheetName()` | `AppConfig.gs` | Hidden queue tab (`_Completion Queue`) |
| `AppConfig.getTrainRouteSheetName()` | `AppConfig.gs` | Train Route templates tab |
| `AppConfig.getTrainBookingHistorySheetName()` | `AppConfig.gs` | IRCTC import target tab |
| `AppConfig.getTrainCompletedJourneysSheetName()` | `AppConfig.gs` | Completed journeys archive tab |
| `AppConfig.getTrainMatchPlaceSheetName()` | `AppConfig.gs` | Optional station label map |

**Team sheet names (BA, Product, Umbraco, etc.) are NOT in code.** They live only in the Matching Sheet data rows above.

---

## Relationships

| From Sheet | Field / Logic | To Sheet | Field / Target | Relationship |
|------------|---------------|----------|----------------|--------------|
| Matching Sheet | `MainSheet` | `* Task Tracker` | — | Config lookup — defines main sheet per team |
| Matching Sheet | `SubSheet` | `* Sub Task Tracker` | — | Config lookup — defines sub-task sheet per team |
| Matching Sheet | `Completed Task` | `* Completed Task` | — | Config lookup — defines archive sheet per team |
| `* Task Tracker` | `Task` (col C) | `* Sub Task Tracker` | `Main Task Name` (col B) | SubTaskSyncModule — auto-sync on main task add/edit |
| `* Task Tracker` / `* Sub Task Tracker` | Completion status | `* Completed Task` | Archived row | CompletionMoveModule — after 24 hours |
| `* Task Tracker` | Task row | `* Sub Task Tracker` | Sub-task rows | One main task → many sub-tasks |
| `* Task Tracker` | Status = complete | `* Completed Task` | Archive row | Task moved/copied on completion |
| 📧 Email Settings | `DATA SHEET NAME` | Named task tracker | — | Alert source sheet |
| Gmail (IRCTC) | Unread confirmation | Train Booking History | Booking row | TrainIrctcImportModule — hourly / manual |
| Train Booking History | Completed / past journey | Train Completed Journeys | Archived row | Manual for now (auto-archive TBD) |
| Match Place | Station label (col A) | Train Booking History | From / To | Canonical label mapping on import |

---

## Train Route

**Tab name:** `Train Route`  
**Purpose:** Weekly and one-time travel templates (stations, train prefs, reminder schedule). Not written by IRCTC import.  
**Header row:** 1 | **Data starts:** 2

| Col | Index | Field Name | Type | Required | Description | Used By |
|-----|-------|------------|------|----------|-------------|---------|
| A | 1 | Route Name | Text | Yes | Template label | Future reminders / matching |
| B | 2 | From Station | Text | Yes | e.g. `TAMBARAM ( TBM)` | Future matching |
| C | 3 | To Station | Text | Yes | e.g. `KULITALAI (KLT)` | Future matching |
| D | 4 | Day of Week | Text | Conditional | Required when Recurrence = Weekly | Future journey generation |
| E | 5 | Recurrence | Text | Yes | `Weekly` or `One-time` | Future reminders |
| F | 6 | Active | Text | Yes | Yes / No | Filter active templates |
| G | 7 | Preferred Class | Text | No | Seat class (SL, 3A) — not Weekly/One-time | Manual / future booking UX |
| H | 8 | Travel Date | Date | Conditional | Required when Recurrence = One-time (e.g. `11-Oct-2026`) | Future reminders / ARP |
| I | 9 | Preferred Train No | Text | No | Preferred train number | Manual / future booking UX |
| J | 10 | Preferred Train Name | Text | No | Preferred train name | Manual / future booking UX |
| K | 11 | Reminder Days Before | Text | No | Comma list e.g. `0,1,2` | Future reminders |
| L | 12 | Alert Time | Text | No | Comma list e.g. `9:30 AM, 13:00` | Future reminders |
| M | 13 | Reminder Email | Text | No | Comma-separated emails | TrainReminderModule |
| N | 14 | Skip Dates | Text | No | Comma-separated travel dates to ignore e.g. `26-Jul-2026, 09-Aug-2026` | TrainReminderModule / TrainRouteModule |
| O | 15 | Notes | Text | No | Free text | — |

**Rules**
- **Weekly** → fill Day of Week; leave Travel Date empty  
- **One-time** → fill Travel Date; leave Day of Week empty  
- **Skip Dates** → optional; those exact travel dates are excluded from missing list, BOOK, and TRAVEL reminders (weekly keeps other weeks)

**Code constants:** `SheetColumns.TRAIN_ROUTE.*` · `AppConfig.getTrainRouteSheetName()`

---

## Train Booking History

**Tab name:** `Train Booking History`  
**Purpose:** Concrete bookings — primarily filled by IRCTC Gmail import (`TrainIrctcImportModule`).  
**Header row:** 1 | **Data starts:** 2

| Col | Index | Field Name | Type | Required | Description | Used By |
|-----|-------|------------|------|----------|-------------|---------|
| A | 1 | From | Text | Yes | From station (resolved via Match Place if present) | TrainIrctcImportModule |
| B | 2 | To | Text | Yes | To station | TrainIrctcImportModule |
| C | 3 | Date of Journey | Text / Date | Yes | e.g. `26-Jul-2026` | TrainIrctcImportModule |
| D | 4 | Name | Text | No | Passenger name | TrainIrctcImportModule |
| E | 5 | Status | Text | No | CONFIRMED / RAC / WL / BOOKED | TrainIrctcImportModule |
| F | 6 | Coach | Text | No | Coach code | TrainIrctcImportModule |
| G | 7 | Seat / Berth | Text | No | Seat or berth | TrainIrctcImportModule |
| H | 8 | Class | Text | No | Travel class | TrainIrctcImportModule |
| I | 9 | PNR | Text | Yes* | 10-digit PNR (*or Transaction ID) | TrainIrctcImportModule (dedupe) |
| J | 10 | Train No | Text | No | Train number | TrainIrctcImportModule |
| K | 11 | Scheduled Departure | Text | No | Departure datetime | TrainIrctcImportModule |
| L | 12 | Date of Boarding | Text / Date | No | Boarding date | TrainIrctcImportModule |
| M | 13 | Transaction ID | Text | Yes* | IRCTC transaction id | TrainIrctcImportModule (dedupe) |
| N | 14 | Date & Time of Booking | Text | No | Booking timestamp | TrainIrctcImportModule |
| O | 15 | User Id | Text | No | IRCTC user id | TrainIrctcImportModule |
| P | 16 | Passenger Mobile | Text | No | 10-digit mobile | TrainIrctcImportModule |
| Q | 17 | Gmail Thread ID | Text | No | Source Gmail thread | TrainIrctcImportModule (dedupe) |

**Triggers / entry points**
- Hourly: `processScheduledIrctcImport_`
- Manual: `runExtractIrctcEmails()`, `testParseSampleIrctcEmail()`

**Gmail:** Unread mail from `ticketadmin@irctc.co.in` or configured forward addresses with IRCTC subjects. Optional Script Property `IRCTC_GMAIL_QUERY` overrides the search string.

**Reminders (Train Route)**
- Every 30 minutes: `processScheduledTrainReminders_`
- Manual: `runTrainRemindersNow()`
- **BOOK** — when today is Reminder Days Before the IRCTC open date (travel − 60 days); **skipped** if a matching booking exists on Train Booking History
- **TRAVEL** — when today is Reminder Days Before the travel date; still sent after booking
- **Missing list** — unbooked travel dates in the next **90 days** (lookahead); Status = `Booking opens TODAY` / `Opens in N day(s)` / `Open — not booked`. ARP for open-date math stays **60** days. Dates listed in **Skip Dates** are excluded.
- One digest **per recipient email** (each address listed on Reminder Email gets only the routes that include them)
- Fires only inside **Alert Time** windows (±15 minutes), same style as task digests

**Code constants:** `SheetColumns.TRAIN_BOOKING.*` · `AppConfig.getTrainBookingHistorySheetName()`

---

## Train Completed Journeys

**Tab name:** `Train Completed Journeys`  
**Purpose:** Archive of finished journeys — same column layout as Train Booking History.  
**Header row:** 1 | **Data starts:** 2  

Columns: identical to [Train Booking History](#train-booking-history).  
**Code:** `AppConfig.getTrainCompletedJourneysSheetName()` · `SheetColumns.TRAIN_BOOKING`  
**Note:** Past journeys (travel date before today) are auto-moved daily by `processScheduledTrainArchive_` from Train Booking History → Train Completed Journeys. Manual: `runArchiveTrainJourneysNow()`.

---

## Match Place (optional)

**Tab name:** `Match Place`  
**Purpose:** Column A lists canonical station labels (e.g. `TAMBARAM ( TBM)`). IRCTC From/To text is mapped to these labels by station code or name.  
**Used by:** `TrainIrctcParseModule`  
If the sheet is missing, email station text is stored as-is.

---

## Change History

| Date | Author / Context | Summary | Sheets / Fields Affected |
|------|------------------|---------|--------------------------|
| 2026-07-09 | Project setup | Initial schema template created | All (template only) |
| 2026-07-09 | User: Team Project Tracker | Documented all live sheets, Matching Sheet data, Email Settings config, and shared task column layout | All active sheets |
| 2026-07-09 | Feature: SubTaskSync | Sub-task sheets use separate column layout; Main Task Name matches main Task column; SubTaskSyncModule added | Sub task sheets, Umbraco/BA/Product Sub Task Tracker |
| 2026-07-09 | User: column rename | Renamed duplicate `Website Url` → `Page Url` on main/completed (col P) and sub (col Q) sheets | Umbraco Task Tracker, Umbraco Sub Task Tracker, Umbraco Completed Task |
| 2026-07-09 | Feature: CompletionMove | Status-based delayed move to completed sheets; 24h queue; status dropdown list documented | Status col, Completed sheets, `_Completion Queue` |
| 2026-07-09 | Refactor: dynamic matching | Removed hardcoded team sheet names from code; all routing via Matching Sheet | MatchingSheetModule, SheetColumns.gs, AppConfig.gs |
| 2026-07-09 | Feature: EffortRollup | Sum sub Estimated/Actual Effort → main; match Task↔Main Task Name; future teams via Matching Sheet | TaskEffortRollupModule, SubTaskSyncModule |
| 2026-07-09 | Feature: EmailNotifications | Status/Priority alerts via Notification Type; scheduled + immediate email | 📧 Email Settings col G, NotificationHandler |
| 2026-07-09 | User: sub column add | Added `Account` column (L) on sub-task sheets; shifted columns L–Y | All Sub Task Tracker sheets, SheetColumns.gs |
| 2026-07-09 | Feature: ErrorLog | Added `Error` sheet for automation error/warning logging | Error sheet, ErrorLogModule.gs |
| 2026-07-19 | Feature: TrainIrctcImport | Train Route templates; IRCTC email → Train Booking History; optional archive to Train Completed Journeys | Train Route, Train Booking History, Train Completed Journeys, Match Place |
| 2026-07-19 | User: Train Route columns | Added Recurrence + Travel Date; reordered Active / Preferred Class / train fields | Train Route, SheetColumns.TRAIN_ROUTE |
| 2026-07-19 | Feature: TrainReminders | BOOK (ARP 60-day open) + TRAVEL reminders; skip BOOK if Booking History match | Train Route, TrainReminder*, TriggerManager |
| 2026-07-19 | Feature: TrainArchive | Auto-move past travel dates Booking History → Train Completed Journeys (daily) | TrainJourneyArchive*, TriggerManager |
| 2026-07-19 | Update: TrainReminder digest | Per-recipient digests; 90-day missing list; Status booking-open countdown; From/To columns | AppConfig, TrainReminderModule |
| 2026-07-19 | Fix: booking match | Train reminder match accepts reverse From/To on same journey date | TrainBookingHistoryModule |
| 2026-07-19 | Feature: Skip Dates | Train Route col N Skip Dates — exclude specific weeks from weekly reminders | Train Route, SheetColumns, TrainRouteModule |

### How to log changes

When you add or change structure, append a row:

```markdown
| 2026-07-10 | Feature: BA alerts | Added Alert Time for BA Task Tracker | 📧 Email Settings |
```

**Change types to record:**
- New sheet
- New column
- Renamed sheet/column (include old → new name)
- Removed sheet/column (mark deprecated; keep history)
- Type or validation change
- Config value change in Email Settings or Matching Sheet

---

## Quick Checklist (copy when editing)

- [ ] Tab name matches Google Sheets exactly (case-sensitive, including `📧`)
- [ ] Column indices match code constants in `SheetColumns.gs`
- [ ] New fields have Type, Required, and Used By filled in
- [ ] Matching Sheet updated if a new team domain is added
- [ ] Email Settings updated if alert target sheet or schedule changes
- [ ] Change History row added
- [ ] README.md updated if user-facing sheet behavior changed
