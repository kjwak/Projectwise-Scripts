# ProjectWise Audit Trail Architecture — From Directory Polling to Event-Driven Processing

## 1. Executive Summary

This document evaluates replacing the current recursive directory polling model with a ProjectWise audit-trail-driven event architecture. The analysis covers available `pwps_dab` audit cmdlets, the `dms_audt` table structure, integration points with the existing watcher/queue/worker pipeline, proposed database schemas, migration phases, and operational risk assessment.

**Recommendation**: Adopt a hybrid architecture — audit trail polling as the primary trigger source with periodic directory reconciliation as a safety net. This reduces ProjectWise API calls by 80-95% during steady state while preserving the correctness guarantees of the current system.

---

## 2. Current Architecture Analysis

### 2.1 Pipeline Overview

```text
┌─────────────────────────────────────────────────────────────────┐
│                    Watch-QCTrigger.ps1                           │
│                                                                 │
│  PW watchList roots ──► Find-PWSheetsFoldersUnderRoot           │
│       │                    └── Get-PWImmediateChildFolders       │
│       ▼                                                         │
│  For each PW folder:                                            │
│    ├─ STATUS_SET_GEN: Get-StatusSetPWFolderState (all docs)     │
│    │    └─ Test-StatusSetWatcherShouldEnqueue (manifest diff)   │
│    │    └─ New-QCJobObject → Add-QCQueueJob                    │
│    │                                                            │
│    └─ QC_PREPEND: Get-PWDocumentsInFolder (all docs)            │
│         └─ filter *.pdf with QC_Archivist description tag       │
│         └─ New-QCJobObject → Add-QCQueueJob                    │
│                                                                 │
│  Local watchFolders ──► Get-ChildItem -Recurse                  │
│       └─ filter/trigger/hash → enqueue                          │
│                                                                 │
│  Cache: local-file-cache.json (skip re-evaluation)              │
└─────────────────────────────────────────────────────────────────┘
          │ JSON queue: pending/ running/ succeeded/ failed/
          ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Run-QCProcessor.ps1                           │
│                                                                 │
│  Get-NextQCJob → Lock → Invoke-QCProcessorByType → Move        │
│    ├─ QC_PREPEND: export/overlay/history/tag-clear              │
│    ├─ STATUS_SET_GEN: pair/merge/writeback                     │
│    └─ QC_REPORTING_SCAN: attribute snapshot                    │
└─────────────────────────────────────────────────────────────────┘
```

### 2.2 Current Cost per Watch Tick

For each tick, the watcher must:

| Operation | API Calls | Notes |
|---|---|---|
| Enumerate roots → sheets folders | ~3-10 `Get-PWFolders`, `Get-PWFolderView` per root | Recursive depth expansion |
| STATUS_SET_GEN per folder | 1 `Get-PWDocumentsInFolder` per sheets folder (all docs) | Returns every document to build state hash |
| QC_PREPEND per folder | 1 `Get-PWDocumentsInFolder` per sheets folder (same call or separate) | Scans all docs, filters for PDF + QC_Archivist tag |
| Local filesystem scan | `Get-ChildItem -Recurse` per watchFolder | Full recursive dir walk |

With ~50 project folders at depth 1 expanding to ~200+ sheets/discipline folders, each tick issues **200-400+ PW API calls** to re-read folder contents regardless of whether anything changed.

### 2.3 What Works Well (Preserve)

- **JSON queue**: robust, atomic, lock-aware, PID-validated. No reason to replace.
- **Processor dispatch**: clean `Invoke-QCProcessorByType` with per-type modes. Unchanged by trigger source.
- **Trigger/filter/dedupe pipeline**: `QC.Filters`, `QC.Triggers`, `QC.JobFactory` are trigger-source-agnostic — they evaluate candidates without caring how they were discovered.
- **StatusSet manifest gate**: `Test-StatusSetWatcherShouldEnqueue` already prevents unnecessary rebuilds. Audit events feed this gate instead of replacing it.
- **Worker pool / dashboard**: `Start-QCPipelineDashboard` orchestration is unaffected.

---

## 3. ProjectWise Audit Trail Capabilities

### 3.1 Available Cmdlets (pwps_dab)

| Cmdlet | Type | Purpose |
|---|---|---|
| `Get-PWDocumentAuditTrailRecords` | Cmdlet | Audit records for specific document(s) |
| `Get-PWFolderAuditTrailRecords` | Cmdlet | Audit records for specific folder(s); supports `-IncludeSubFolders` `-IncludeDocuments` |
| `Get-PWUserAuditTrailRecords` | Cmdlet | Audit records for specific user(s) in date range; includes email column |
| `Get-PWAuditTrailRecordsFromPreviousDays` | Function | Convenience wrapper; extracts recent records |
| `Export-PWAuditTrailToSQLite` | Cmdlet | Bulk export audit trail to SQLite; supports `-StartDate`, `-EndDate`, `-IncludeTime`, `-IncludeFullObjectNames` |
| `Export-PWAuditTrailNotificationReport` | Cmdlet | Formatted audit notification report |
| `Get-PWAuditTrailSettings` | Cmdlet | Query current audit trail configuration; returns secondary table name if configured |
| `Set-PWAuditTrailSetting` | Cmdlet | Configure individual audit setting |
| `Set-PWAuditTrailSettingsAll` | Cmdlet | Configure all audit settings |
| `Clear-PWAuditTrailSetting` | Cmdlet | Clear individual audit setting |
| `Clear-PWAuditTrailSettingsAll` | Cmdlet | Clear all audit settings |
| `Move-PWLogInOutAuditTrailRecords` | Cmdlet | Move login/logout records (maintenance) |

### 3.2 Direct SQL Access via `Select-PWSQL`

The `dms_audt` table can be queried directly using `Select-PWSQL`. This is the most flexible approach for incremental polling.

**Table structure** (`dms_audt`):

| Column | Alias | Description |
|---|---|---|
| `o_objtype` | TYPE | Object type ID (1=Folder, 2=Document, 3=Set, etc.) |
| `o_objguid` | OBJECT_GUID | GUID of affected object |
| `o_objno` | OBJECT_NUMBER | Numeric ID of affected object |
| `o_action` | ACTION_ID | Action type numeric code |
| `o_acttime` | DATE_TIME | Timestamp of the action |
| `o_userno` | USER_NUMBER | User ID who performed action |
| `o_comments` | COMMENTS | Optional comment text |
| `o_numparam1` | PARAM_01 | Numeric parameter 1 |
| `o_numparam2` | PARAM_02 | Numeric parameter 2 |
| `o_textparam` | TEXT_PARAM | Text parameter |
| `o_guidparam` | GUID_PARAM | GUID parameter |
| `o_itemname` | ITEM_NAME | Name of affected item |
| `o_itemdesc` | ITEM_DESCRIPTION | Description of affected item |
| `o_parentguid` | PARENT_GUID | Parent GUID (containing folder, etc.) |

### 3.3 Audit Trail Action Codes — Events Relevant to QC Pipeline

#### Document Events (o_objtype = 2)

| Code | Name | Relevance to QC |
|---|---|---|
| 1001 | `DOCUMENT_CREATE` | New document in sheets folder → STATUS_SET_GEN candidate |
| 1002 | `DOCUMENT_MODIFY` | Attribute/property change → potential QC_PREPEND if description tag added |
| 1003 | `DOCUMENT_ATTR` | Environment attribute change → QC workflow attribute monitoring |
| 1004 | `DOCUMENT_FILE_ADD` | File content added |
| 1006 | `DOCUMENT_FILE_REP` | File content replaced (check-in with new content) → STATUS_SET_GEN rebuild |
| 1007 | `DOCUMENT_CIN` | **Check-in** → primary trigger for STATUS_SET_GEN and QC_PREPEND |
| 1009 | `DOCUMENT_CHOUT` | Check-out → track in-progress work |
| 1010 | `DOCUMENT_CPOUT` | Copy-out |
| 1012 | `DOCUMENT_STATE` | **Workflow state change** → QC workflow tracking, notifications |
| 1015 | `DOCUMENT_VERSION` | New version created → STATUS_SET_GEN rebuild trigger |
| 1016 | `DOCUMENT_MOVE` | Document moved → path-based filter re-evaluation |
| 1020 | `DOCUMENT_DELETE` | Document deleted → STATUS_SET_GEN rebuild trigger |
| 1022 | `DOCUMENT_FREE` | Free/undo checkout |

#### Folder Events (o_objtype = 1)

| Code | Name | Relevance |
|---|---|---|
| 1 | `FOLDER_CREATE` | New project folder → watch path expansion |
| 3 | `FOLDER_WFLOW` | Folder workflow assignment change |
| 5 | `FOLDER_STATE` | Folder state change |

### 3.4 Events That Map to Current Triggers

| Current Trigger | Current Detection | Audit Event(s) |
|---|---|---|
| QC_PREPEND (description tag) | Scan all docs → find `QC_Archivist` in description | `DOCUMENT_MODIFY (1002)` or `DOCUMENT_ATTR (1003)` on the doc |
| STATUS_SET_GEN (folder change) | Scan all docs → hash folder state → compare manifest | `DOCUMENT_CIN (1007)`, `DOCUMENT_CREATE (1001)`, `DOCUMENT_DELETE (1020)`, `DOCUMENT_VERSION (1015)`, `DOCUMENT_FILE_REP (1006)` on any doc in the folder |
| QC workflow state change | Not yet implemented (future) | `DOCUMENT_STATE (1012)` |
| QC attribute change | Not yet implemented (future) | `DOCUMENT_ATTR (1003)` |

### 3.5 Key Capabilities and Limitations

**Capabilities:**
- Audit trail includes timestamps (`o_acttime`) suitable for watermark/cursor polling
- Object GUIDs and numeric IDs link events to specific documents/folders
- Parent GUIDs connect document events to containing folders
- Action codes are stable and well-documented
- Both primary and secondary (backup) audit trail tables are queryable
- `Select-PWSQL` allows arbitrary SQL against the datasource database

**Limitations:**
- Audit trail records do **not** include the old/new values for attribute changes — only that an attribute change occurred. A follow-up `Get-PWDocumentEAttributes` call is needed to read current values.
- Description field changes (`DOCUMENT_MODIFY`) do not distinguish which property was modified. The description must be re-read.
- No built-in "event stream" or push notification — polling is required.
- Parameter sets for the built-in cmdlets are not exposed in offline introspection (require active PW connection). Direct SQL via `Select-PWSQL` is more reliable for automation.
- `dms_audt` can grow large; heavy queries impact SQL server performance. Use narrow time windows and `TOP N` limits.
- Audit trail retention is administrator-configured; old records may be purged or moved to secondary tables.

---

## 4. Proposed Architecture: Hybrid Audit-Driven Pipeline

### 4.1 Architecture Diagram

```text
┌────────────────────────────────────────────────────────────────────────┐
│                  PW.AuditPoller (new module)                           │
│                                                                        │
│  Every N seconds (configurable, default 60s):                          │
│    1. Read watermark from poll_state.json                              │
│    2. SELECT from dms_audt WHERE o_acttime > @watermark                │
│       AND o_action IN (1001,1002,1003,1006,1007,1012,1015,1020)       │
│       AND o_objtype = 2                                                │
│       ORDER BY o_acttime ASC                                           │
│    3. For each event:                                                  │
│       a. Map event → candidate type (QC_PREPEND, STATUS_SET_GEN, etc) │
│       b. Resolve document/folder context via parent GUID               │
│       c. Deduplicate: coalesce same-document events within window      │
│       d. Append to audit_events log (database or JSON-lines)          │
│    4. Feed candidates into existing trigger/filter/enqueue pipeline    │
│    5. Update watermark                                                 │
│                                                                        │
│  Periodic reconciliation (every 6-24 hours):                           │
│    Full directory scan (current watcher logic) as safety net           │
│    Compare against audit-derived state                                │
│    Log discrepancies                                                   │
└────────────────────────────────────────────────────────────────────────┘
          │
          ▼ (same JSON queue — unchanged)
┌────────────────────────────────────────────────────────────────────────┐
│              Existing pipeline (unchanged)                             │
│  QC.Filters → QC.Triggers → QC.JobFactory → QC.Queue.Json            │
│  → QC.Processors → QC.Worker                                         │
└────────────────────────────────────────────────────────────────────────┘
          │
          ▼ (future addition)
┌────────────────────────────────────────────────────────────────────────┐
│              Event Log Database (SQLite or future RDBMS)              │
│                                                                        │
│  audit_events          — raw audit trail records                      │
│  document_activity     — enriched per-document activity               │
│  transition_events     — detected state/attribute transitions         │
│  poll_runs             — poller health and performance tracking       │
│  processing_jobs       — job lifecycle (mirrors queue state)          │
│  document_state_history — time-series document state                  │
└────────────────────────────────────────────────────────────────────────┘
          │
          ▼ (future)
┌────────────────────────────────────────────────────────────────────────┐
│  Dashboard / Reporting / Notifications                                │
│  - QC workflow aging metrics                                          │
│  - Reviewer/designer activity timelines                               │
│  - Status set rebuild frequency                                       │
│  - Operational analytics                                              │
└────────────────────────────────────────────────────────────────────────┘
```

### 4.2 Event Flow Detail

```text
dms_audt record (DOCUMENT_CIN on sheet.pdf in CADD\Sheets\Structures)
    │
    ├─ o_parentguid → resolve to Documents\AZDOT\PROJ\CADD\Sheets\Structures
    ├─ o_objguid    → resolve to specific document
    ├─ o_acttime    → 2026-05-26T13:45:12
    ├─ o_action     → 1007 (DOCUMENT_CIN)
    │
    ▼
PW.AuditPoller classifies:
    ├─ Folder matches watchList root "Documents\AZDOT" with sheetsPathFromProject "CADD\Sheets"
    ├─ Document is *.pdf in a Sheets folder → STATUS_SET_GEN candidate (folder-level)
    ├─ Check description for QC_Archivist → if present, also QC_PREPEND candidate
    │
    ▼
Feed into existing pipeline:
    ├─ Test-QCPathAllowed (filters/blacklist)
    ├─ Test-StatusSetWatcherShouldEnqueue (manifest gate — still runs)
    ├─ New-QCJobObject (job creation with dedupe key)
    ├─ Test-QCDuplicateJob (queue dedupe)
    ├─ Add-QCQueueJob (enqueue)
```

### 4.3 Watermark / Cursor Approach

The poller maintains a persistent watermark — the `o_acttime` of the last successfully processed audit record.

```powershell
# Pseudocode: PW.AuditPoller core loop

function Invoke-PollTick {
    param([hashtable]$Config, [hashtable]$PollState)

    $watermark = $PollState.lastProcessedTimestamp
    if (-not $watermark) {
        $watermark = (Get-Date).AddHours(-2).ToString('yyyy-MM-dd HH:mm:ss')
    }

    $relevantActions = '1001,1002,1003,1006,1007,1012,1015,1020'

    $sql = @"
SELECT TOP 500
    o_acttime, o_action, o_objtype, o_objno, o_objguid,
    o_parentguid, o_userno, o_itemname, o_itemdesc, o_textparam
FROM dms_audt
WHERE o_acttime > '$watermark'
  AND o_objtype = 2
  AND o_action IN ($relevantActions)
ORDER BY o_acttime ASC
"@

    $results = Select-PWSQL -SQLSelectStatement $sql
    if ($results.Rows.Count -eq 0) { return @{ newEvents = 0 } }

    $candidates = @()
    $maxTimestamp = $watermark

    foreach ($row in $results.Rows) {
        $ts = [string]$row.o_acttime
        if ($ts -gt $maxTimestamp) { $maxTimestamp = $ts }

        $actionName = $AuditActionMap[[int]$row.o_action]
        $parentGuid = [string]$row.o_parentguid
        $docGuid    = [string]$row.o_objguid
        $docName    = [string]$row.o_itemname

        # Resolve parent GUID to folder path (cached)
        $folderPath = Resolve-FolderPathFromGuid -Guid $parentGuid -Cache $FolderCache

        # Classify event
        $candidate = @{
            eventType     = $actionName
            eventTime     = $ts
            documentGuid  = $docGuid
            documentName  = $docName
            folderPath    = $folderPath
            action        = [int]$row.o_action
        }

        # Route to appropriate trigger type
        switch ([int]$row.o_action) {
            { $_ -in @(1001, 1006, 1007, 1015, 1020) } {
                # Content change in sheets folder → STATUS_SET_GEN
                $candidate.triggerType = 'STATUS_SET_GEN'
            }
            { $_ -in @(1002, 1003) } {
                # Attribute/property change → check for QC_PREPEND tag
                $candidate.triggerType = 'CHECK_DESCRIPTION'
            }
            1012 {
                # State change → QC workflow tracking
                $candidate.triggerType = 'STATE_CHANGE'
            }
        }

        $candidates += $candidate
    }

    # Coalesce: if multiple events hit the same folder within this tick,
    # emit one STATUS_SET_GEN candidate per folder
    $folderGroups = $candidates |
        Where-Object { $_.triggerType -eq 'STATUS_SET_GEN' } |
        Group-Object folderPath

    foreach ($group in $folderGroups) {
        # Feed one candidate per folder into existing trigger pipeline
        Invoke-ExistingTriggerPipeline -FolderPath $group.Name -Config $Config
    }

    # Update watermark
    $PollState.lastProcessedTimestamp = $maxTimestamp
    Save-PollState -PollState $PollState

    return @{
        newEvents = $results.Rows.Count
        candidateFolders = $folderGroups.Count
        newWatermark = $maxTimestamp
    }
}
```

### 4.4 Duplicate Event Suppression

Multiple audit events can fire for the same logical change (e.g., check-in generates `DOCUMENT_CIN` + `DOCUMENT_FILE_REP` + potentially `DOCUMENT_ATTR`). The poller should:

1. **Coalesce by document within a poll tick** — if the same document has multiple events in one batch, emit one candidate.
2. **Coalesce by folder for STATUS_SET_GEN** — any document change in a folder emits one folder-level rebuild candidate per tick.
3. **Rely on existing dedupe** — `Test-QCDuplicateJob` and the `dedupeKey` mechanism in `QC.JobFactory` already prevent duplicate queue entries. The `folderStateHash` in STATUS_SET_GEN jobs ensures that only actual content changes create new work.

### 4.5 QC_PREPEND Trigger via Audit

Currently QC_PREPEND relies on scanning all documents and checking the `Description` field for `QC_Archivist`. With audit polling:

1. Watch for `DOCUMENT_MODIFY (1002)` events on documents in watched folders.
2. When detected, re-read the document description with `Get-PWDocumentsBySearch -DocumentName $name -FolderPath $folder -JustThisFolder`.
3. If description now contains `QC_Archivist`, feed into `QC_PREPEND` trigger pipeline.

This replaces scanning all documents in a folder (potentially hundreds) with a targeted read of only the changed document.

---

## 5. Proposed Module Structure

### 5.1 New Module: `PW.AuditPoller.psm1`

```
modules/
├── PW.AuditPoller.psm1          # New: audit trail polling and event classification
├── PW.AuditPoller.Schema.psm1   # New: database schema management (SQLite)
├── PW.Connection.psm1           # Existing: add Select-PWSQL wrapper
├── PW.Discovery.psm1            # Existing: unchanged (used by reconciliation)
├── QC.Filters.psm1              # Existing: unchanged
├── QC.Triggers.psm1             # Existing: unchanged
├── QC.JobFactory.psm1           # Existing: unchanged
├── QC.Queue.Json.psm1           # Existing: unchanged
├── QC.Processors.psm1           # Existing: unchanged
└── Orchestrator.Pipeline.psm1   # Existing: add Invoke-QCAuditPollTick
```

### 5.2 Exports from `PW.AuditPoller.psm1`

```powershell
# Core polling
Invoke-PWAuditPollTick          # One poll cycle: read events, classify, emit candidates
Get-PWAuditWatermark            # Read current watermark
Set-PWAuditWatermark            # Persist watermark after successful processing

# Event classification
ConvertTo-QCAuditCandidate      # Map raw audit row → pipeline candidate
Group-QCAuditCandidates         # Coalesce by folder/document
Test-QCAuditEventRelevant       # Filter to watched folders and relevant actions

# Folder resolution
Resolve-PWFolderPathFromGuid    # GUID → folder path (cached)
Test-PWFolderInWatchScope       # Is this folder under a configured watch root?

# Reconciliation
Invoke-QCAuditReconcile         # Full scan comparison (safety net)
Compare-QCAuditVsDirectoryScan  # Diff detected events vs full scan results
```

### 5.3 Configuration Additions to `appsettings.json`

```jsonc
{
  "auditPoller": {
    "enabled": true,
    "pollIntervalSeconds": 60,
    "maxEventsPerPoll": 500,
    "lookbackOnStartupHours": 2,
    "relevantDocumentActions": [1001, 1002, 1003, 1006, 1007, 1012, 1015, 1020],
    "relevantFolderActions": [1, 3, 4, 5],
    "coalesceWindowSeconds": 30,
    "stateFilePath": "C:\\QC_E2E_RealRun\\audit_poller\\poll_state.json",
    "eventLogPath": "C:\\QC_E2E_RealRun\\audit_poller\\events",
    "reconciliation": {
      "enabled": true,
      "intervalHours": 12,
      "fullScanOnStartup": true
    },
    "database": {
      "enabled": false,
      "provider": "SQLite",
      "connectionString": "Data Source=C:\\QC_E2E_RealRun\\audit_poller\\qc_events.sqlite"
    },
    "performance": {
      "folderGuidCacheMaxEntries": 1000,
      "folderGuidCacheTtlMinutes": 60,
      "skipSecondaryAuditTrail": true
    }
  }
}
```

---

## 6. Proposed Database Schema

### 6.1 SQLite Schema (Phase 1)

```sql
-- Raw audit events captured from dms_audt
CREATE TABLE IF NOT EXISTS audit_events (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    captured_at_utc     TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now')),
    poll_run_id         INTEGER,
    pw_acttime          TEXT NOT NULL,           -- o_acttime from dms_audt
    pw_action           INTEGER NOT NULL,        -- o_action code
    pw_action_name      TEXT,                    -- human-readable action name
    pw_objtype          INTEGER NOT NULL,        -- o_objtype (1=folder, 2=document)
    pw_objno            INTEGER,                 -- o_objno
    pw_objguid          TEXT,                    -- o_objguid
    pw_parentguid       TEXT,                    -- o_parentguid (containing folder)
    pw_userno           INTEGER,                 -- o_userno
    pw_itemname         TEXT,                    -- o_itemname
    pw_itemdesc         TEXT,                    -- o_itemdesc
    pw_textparam        TEXT,                    -- o_textparam
    resolved_folder     TEXT,                    -- resolved folder path
    candidate_type      TEXT,                    -- QC_PREPEND, STATUS_SET_GEN, STATE_CHANGE, etc.
    processed           INTEGER NOT NULL DEFAULT 0,
    enqueued_job_id     TEXT                     -- job ID if event led to enqueue
);

CREATE INDEX IF NOT EXISTS idx_audit_events_acttime ON audit_events(pw_acttime);
CREATE INDEX IF NOT EXISTS idx_audit_events_objguid ON audit_events(pw_objguid);
CREATE INDEX IF NOT EXISTS idx_audit_events_folder ON audit_events(resolved_folder);
CREATE INDEX IF NOT EXISTS idx_audit_events_processed ON audit_events(processed);
CREATE INDEX IF NOT EXISTS idx_audit_events_poll_run ON audit_events(poll_run_id);

-- Document activity summary (enriched, deduplicated)
CREATE TABLE IF NOT EXISTS document_activity (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    document_guid       TEXT NOT NULL,
    document_name       TEXT,
    folder_path         TEXT,
    last_action         TEXT,
    last_action_code    INTEGER,
    last_action_time    TEXT,
    last_action_user    INTEGER,
    total_events        INTEGER NOT NULL DEFAULT 0,
    first_seen_utc      TEXT,
    last_seen_utc       TEXT,
    UNIQUE(document_guid)
);

CREATE INDEX IF NOT EXISTS idx_doc_activity_folder ON document_activity(folder_path);
CREATE INDEX IF NOT EXISTS idx_doc_activity_lastaction ON document_activity(last_action_time);

-- Document state/attribute history (time-series)
CREATE TABLE IF NOT EXISTS document_state_history (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    document_guid       TEXT NOT NULL,
    document_name       TEXT,
    folder_path         TEXT,
    captured_at_utc     TEXT NOT NULL,
    event_type          TEXT,                    -- STATE_CHANGE, ATTR_CHANGE, etc.
    source_audit_id     INTEGER,                 -- FK to audit_events.id
    old_value           TEXT,                    -- previous state/attribute value (if known)
    new_value           TEXT,                    -- new state/attribute value
    field_name          TEXT,                    -- which field changed
    changed_by_user     INTEGER
);

CREATE INDEX IF NOT EXISTS idx_state_hist_docguid ON document_state_history(document_guid);
CREATE INDEX IF NOT EXISTS idx_state_hist_captured ON document_state_history(captured_at_utc);
CREATE INDEX IF NOT EXISTS idx_state_hist_type ON document_state_history(event_type);

-- Detected transitions (business-level events)
CREATE TABLE IF NOT EXISTS transition_events (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    detected_at_utc     TEXT NOT NULL,
    document_guid       TEXT NOT NULL,
    document_name       TEXT,
    folder_path         TEXT,
    transition_type     TEXT NOT NULL,            -- QC_STAGE_CHANGE, CHECKIN, STATUS_SET_REBUILD, etc.
    from_value          TEXT,
    to_value            TEXT,
    trigger_audit_id    INTEGER,                  -- FK to audit_events.id
    job_id              TEXT,                      -- generated job ID (if any)
    job_type            TEXT,                      -- QC_PREPEND, STATUS_SET_GEN, etc.
    notification_sent   INTEGER NOT NULL DEFAULT 0,
    notification_id     TEXT
);

CREATE INDEX IF NOT EXISTS idx_transition_docguid ON transition_events(document_guid);
CREATE INDEX IF NOT EXISTS idx_transition_type ON transition_events(transition_type);
CREATE INDEX IF NOT EXISTS idx_transition_detected ON transition_events(detected_at_utc);

-- Poll run tracking (operational health)
CREATE TABLE IF NOT EXISTS poll_runs (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    started_at_utc      TEXT NOT NULL,
    completed_at_utc    TEXT,
    watermark_before    TEXT,
    watermark_after     TEXT,
    events_fetched      INTEGER NOT NULL DEFAULT 0,
    events_relevant     INTEGER NOT NULL DEFAULT 0,
    candidates_created  INTEGER NOT NULL DEFAULT 0,
    jobs_enqueued       INTEGER NOT NULL DEFAULT 0,
    duration_ms         INTEGER,
    error_message       TEXT,
    is_reconciliation   INTEGER NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_poll_runs_started ON poll_runs(started_at_utc);

-- Processing jobs (mirrors queue state for dashboards)
CREATE TABLE IF NOT EXISTS processing_jobs (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id              TEXT NOT NULL UNIQUE,
    job_type            TEXT NOT NULL,
    source_path         TEXT,
    source_folder       TEXT,
    created_at_utc      TEXT NOT NULL,
    started_at_utc      TEXT,
    completed_at_utc    TEXT,
    status              TEXT NOT NULL DEFAULT 'pending',   -- pending, running, succeeded, failed
    trigger_source      TEXT,                              -- audit_poll, directory_scan, reconciliation, manual
    trigger_audit_id    INTEGER,
    dedupe_key          TEXT,
    attempt_count       INTEGER NOT NULL DEFAULT 0,
    duration_ms         INTEGER,
    error_code          TEXT,
    error_message       TEXT,
    result_data         TEXT                                -- JSON blob for processor output
);

CREATE INDEX IF NOT EXISTS idx_jobs_status ON processing_jobs(status);
CREATE INDEX IF NOT EXISTS idx_jobs_type ON processing_jobs(job_type);
CREATE INDEX IF NOT EXISTS idx_jobs_created ON processing_jobs(created_at_utc);
CREATE INDEX IF NOT EXISTS idx_jobs_folder ON processing_jobs(source_folder);
```

### 6.2 Views for Dashboard Queries

```sql
-- Active QC cycle aging
CREATE VIEW IF NOT EXISTS v_qc_cycle_aging AS
SELECT
    document_guid,
    document_name,
    folder_path,
    MIN(CASE WHEN event_type = 'STATE_CHANGE' AND new_value LIKE '%QC Received%' THEN captured_at_utc END) AS qc_received_at,
    MAX(captured_at_utc) AS last_activity_at,
    ROUND((julianday('now') - julianday(MIN(CASE WHEN event_type = 'STATE_CHANGE' AND new_value LIKE '%QC Received%' THEN captured_at_utc END))) * 24, 1) AS hours_in_qc
FROM document_state_history
GROUP BY document_guid;

-- Folder activity heatmap
CREATE VIEW IF NOT EXISTS v_folder_activity AS
SELECT
    resolved_folder AS folder_path,
    COUNT(*) AS event_count,
    COUNT(DISTINCT pw_objguid) AS unique_documents,
    MAX(pw_acttime) AS last_activity,
    MIN(pw_acttime) AS first_activity
FROM audit_events
WHERE pw_acttime > datetime('now', '-7 days')
GROUP BY resolved_folder
ORDER BY event_count DESC;

-- Poller health
CREATE VIEW IF NOT EXISTS v_poller_health AS
SELECT
    id,
    started_at_utc,
    duration_ms,
    events_fetched,
    events_relevant,
    jobs_enqueued,
    CASE WHEN error_message IS NOT NULL THEN 'ERROR' ELSE 'OK' END AS status,
    watermark_after
FROM poll_runs
ORDER BY started_at_utc DESC
LIMIT 100;
```

---

## 7. Migration Path

### Phase 1: Instrument and Observe (1-2 weeks)

**Goal**: Prove audit trail data is reliable and complete without changing trigger behavior.

1. Add `PW.AuditPoller.psm1` as a **read-only observer** alongside the existing watcher.
2. On each watcher tick, also poll `dms_audt` for events since last tick.
3. Log both:
   - What the directory scan would trigger.
   - What the audit poller would trigger.
4. Compare results. Track:
   - Events that audit catches but directory scan misses (timing gap).
   - Events that directory scan catches but audit misses (audit coverage gap — should be rare/none).
   - Latency between PW action and audit record appearance.

**Deliverables**:
- `PW.AuditPoller.psm1` with observe-only mode.
- Comparison logging in `Watch-QCTrigger.ps1`.
- Confidence report: "audit trail covers N% of real triggers over M days".

**Risk**: Zero — read-only; existing pipeline unchanged.

### Phase 2: Audit-Primary with Directory Fallback (2-4 weeks)

**Goal**: Audit poller becomes primary trigger source; directory scan runs as periodic reconciliation.

1. Configure `auditPoller.enabled = true` in `appsettings.json`.
2. `Start-QCPipelineDashboard` runs audit poll ticks instead of (or interleaved with) full directory scans.
3. Full directory scan runs every N hours (configurable) as reconciliation.
4. All candidates flow through existing `QC.Filters → QC.Triggers → QC.JobFactory → QC.Queue.Json` pipeline unchanged.
5. Enable SQLite event logging for operational visibility.

**Deliverables**:
- Updated `Orchestrator.Pipeline.psm1` with `Invoke-QCAuditPollTick`.
- Updated `Start-QCPipelineDashboard.ps1` to support audit-primary mode.
- Reconciliation scheduler.
- SQLite event database.

**Risk**: Low — fallback reconciliation catches anything audit misses.

### Phase 3: Database-Backed Dashboard and Analytics (4-8 weeks)

**Goal**: Event database powers dashboards, notifications, and workflow analytics.

1. Enable full event logging to SQLite.
2. Build dashboard queries against `v_qc_cycle_aging`, `v_folder_activity`, etc.
3. Wire `QC.Notifications` to fire from `transition_events` instead of/in addition to job completion.
4. Build aging/SLA metrics from `document_state_history`.
5. Evaluate migration to SQL Server or PostgreSQL if SQLite limits are reached.

**Deliverables**:
- Dashboard integration.
- Notification triggers from transition events.
- Operational analytics reports.

### Phase 4: Optimize and Scale (ongoing)

1. Reduce reconciliation frequency based on Phase 2 confidence data.
2. Tune poll interval based on measured latency/throughput.
3. Add audit trail health monitoring (detect gaps, purges, secondary table failover).
4. Consider direct `dms_audt` table trigger or Change Data Capture if RDBMS supports it (SQL Server only).

---

## 8. Comparison: Directory Polling vs. Audit-Driven

| Dimension | Directory Polling (Current) | Audit-Driven (Proposed) |
|---|---|---|
| **PW API calls per tick** | 200-400+ (enumerate all folders and documents) | 1 SQL query (`Select-PWSQL`) + 0-5 targeted reads |
| **Latency to detection** | Tick interval (currently minutes) | Poll interval (configurable, 30-60s typical) |
| **SQL server impact** | Medium-high (many folder/doc queries) | Low (single indexed query on `dms_audt`) |
| **Can detect attribute changes** | Only `Description` field (manual scan) | Yes — `DOCUMENT_ATTR (1003)` events |
| **Can detect state changes** | No | Yes — `DOCUMENT_STATE (1012)` events |
| **Can detect check-in/out** | No | Yes — `DOCUMENT_CIN`, `DOCUMENT_CHOUT` |
| **Scales with folder count** | Linear degradation | Constant (event volume, not folder count) |
| **Missed event risk** | Low (full scan every tick) | Low with reconciliation; medium without |
| **Startup recovery** | Immediate (full scan) | Lookback window + reconciliation |
| **Event history** | None (fire-and-forget) | Full history in database |
| **Dashboard support** | Limited (snapshot metrics only) | Rich (time-series, aging, activity heatmaps) |
| **Future notification support** | Requires separate mechanism | Native — transitions trigger notifications |

---

## 9. Operational Concerns and Mitigations

### 9.1 SQL Load

**Concern**: Polling `dms_audt` could impact the PW SQL server.

**Mitigation**:
- Use `TOP N` (default 500) to cap result size.
- Use the `o_acttime` watermark in WHERE clause — this hits the timestamp index.
- Poll at 60s intervals, not continuous. One query per minute is negligible compared to normal PW Explorer traffic.
- `dms_audt` is regularly queried by PW Administrator for reporting; the table is designed for this.
- Monitor query execution time in `poll_runs` table.

### 9.2 Missed Events

**Concern**: If the poller stops or the audit trail is purged, events could be missed.

**Mitigation**:
- On startup, look back `lookbackOnStartupHours` (default 2 hours) to recover events from downtime.
- Periodic reconciliation (full directory scan) catches anything missed.
- Watermark is persisted to disk; survives process restart.
- If watermark is older than audit trail retention, log a warning and trigger immediate reconciliation.
- Monitor `poll_runs` table for gaps.

### 9.3 Duplicate Processing

**Concern**: Same logical change could generate multiple audit records.

**Mitigation**:
- Coalesce events by document and folder within each poll tick.
- Existing `dedupeKey` in `QC.JobFactory` prevents duplicate queue entries.
- `folderStateHash` in STATUS_SET_GEN ensures rebuilds only when content actually changed.
- `Test-StatusSetWatcherShouldEnqueue` manifest comparison still runs — false triggers are a no-op.

### 9.4 Audit Trail Retention and Secondary Tables

**Concern**: PW administrators may configure audit trail purging or backup-table rotation.

**Mitigation**:
- Query `Get-PWAuditTrailSettings` on startup to confirm configuration.
- Use `-SkipSecondary` (or `skipSecondaryAuditTrail` config) by default for polling to avoid double-counting.
- If secondary table is configured, query it during reconciliation for historical completeness.
- Alert if watermark is older than retention period.

### 9.5 Folder GUID Resolution Performance

**Concern**: Resolving `o_parentguid` to folder paths could be expensive.

**Mitigation**:
- Cache GUID → folder path mappings with configurable TTL (default 60 minutes).
- Pre-warm cache with watch root folders on startup.
- Watched folder count is bounded by configuration; cache size is predictable.
- Use `Get-PWFoldersByGUIDs` for batch resolution when available.

### 9.6 Clock Skew and Timestamp Ordering

**Concern**: `o_acttime` might not be perfectly monotonic across PW server restarts or timezone changes.

**Mitigation**:
- Store watermark as the maximum `o_acttime` seen, not the current clock time.
- On each poll, overlap the watermark by a small margin (e.g., subtract 5 seconds) to catch boundary events.
- Accept that occasional duplicate processing is harmless (dedupe handles it).

---

## 10. What Can Be Eliminated vs. Retained

### 10.1 Can Be Eliminated (Steady-State)

| Component | Current Role | Replacement |
|---|---|---|
| Per-folder `Get-PWDocumentsInFolder` for QC_PREPEND | Scan all docs, check description | Audit event `DOCUMENT_MODIFY` → targeted read |
| Per-folder `Get-StatusSetPWFolderState` on every tick | Hash all docs to detect changes | Audit event `DOCUMENT_CIN/CREATE/DELETE` triggers targeted re-hash |
| `Find-PWSheetsFoldersUnderRoot` on every tick | Enumerate folder tree | Cache folder tree; refresh on `FOLDER_CREATE` events |
| `Get-PWImmediateChildFolders` on every tick | Expand oneLevelDeep | Cache; refresh on folder events |
| `Get-ChildItem -Recurse` local scan on every tick | Find new/changed local files | Not replaced by audit (local-only); keep for local watchFolders |

### 10.2 Must Retain

| Component | Reason |
|---|---|
| `Test-StatusSetWatcherShouldEnqueue` | Manifest comparison is the correctness gate; audit tells us to check, manifest confirms whether rebuild is needed |
| `QC.Filters → QC.Triggers → QC.JobFactory` | Trigger-source-agnostic; audit candidates feed into the same pipeline |
| `QC.Queue.Json` | Queue mechanism is orthogonal to trigger source |
| `QC.Processors` | Unaffected by trigger source |
| Local filesystem watcher | Audit trail is ProjectWise-only; local `watchFolders` still need `Get-ChildItem` |
| Periodic reconciliation (full scan) | Safety net for missed events, audit retention gaps, GUID cache staleness |

### 10.3 Recommended as Fallback Reconciliation

| Component | Reconciliation Role |
|---|---|
| Full directory scan (current watcher logic) | Run every 6-24 hours; compare detected changes with audit-derived state; log discrepancies |
| Local watcher cache (`local-file-cache.json`) | Continue for local filesystem paths only |

---

## 11. Future Capabilities Enabled

### 11.1 PM Dashboards

With `document_state_history` and `transition_events`, dashboards can show:
- Documents currently in each QC state, with aging.
- Time-in-state distributions per project/folder.
- Reviewer and designer activity timelines.
- Status set rebuild frequency and timing.

### 11.2 QC Workflow Tracking

`DOCUMENT_STATE (1012)` events provide real-time state change detection:
- Red → Green → Blue cycle tracking.
- Automatic `QC.Notifications` triggering on state transitions.
- Loopback detection (Green → Red re-review cycles).
- Escalation on stale documents (aging threshold breach).

### 11.3 Reviewer/Designer Notifications

`transition_events` can drive notifications:
- "Your document has entered QC Received" (to reviewer).
- "Corrections have been submitted" (to reviewer for backcheck).
- "Document has been in Redlines Issued for 5 days" (aging alert).

### 11.4 Operational Analytics

`poll_runs` and `audit_events` provide:
- System activity patterns (peak hours, heavy users).
- Automation health monitoring (poller latency, event throughput).
- Capacity planning (event volume trends).

---

## 12. Unknowns and Risks Requiring Testing

| # | Unknown | Test Plan | Risk Level |
|---|---|---|---|
| 1 | `o_acttime` monotonicity and timezone behavior | Run poller for 1 week; check for out-of-order or duplicate timestamps | Low |
| 2 | `o_parentguid` reliability for folder resolution | Verify against known document/folder pairs; check for null parent GUIDs | Low |
| 3 | Audit coverage of `Description` field changes | Manually add QC_Archivist tag to a document; verify `DOCUMENT_MODIFY` fires | Medium |
| 4 | Audit trail retention configuration on production datasource | Run `Get-PWAuditTrailSettings`; confirm retention period | Low |
| 5 | `Select-PWSQL` availability and permissions | Verify the service account can run `Select-PWSQL` against `dms_audt` | Medium |
| 6 | Performance of `dms_audt` query at scale | Measure query time with 1M+ records in table; confirm index usage | Low |
| 7 | Whether attribute changes (`DOCUMENT_ATTR`) reliably fire for QC environment attributes | Test `Set-PWDocumentEAttributes` and verify audit record appears | Medium |
| 8 | Whether `o_itemdesc` includes the new description after `DOCUMENT_MODIFY` | Check if the description is in the audit record or requires follow-up read | Low |
| 9 | Behavior when PW connection drops during poll | Test reconnection handling; verify watermark is not advanced past unprocessed events | Low |
| 10 | Interaction between audit polling and concurrent watcher processes | Verify dedupe prevents double-enqueue when both are running in Phase 2 | Low |

### Discovery Script for Risk Items 3, 5, 7, 8

```powershell
# tools/discovery/Test-PWAuditTrailCapabilities.ps1
# Run with active PW connection to validate audit trail assumptions

# 1. Check Select-PWSQL availability
$sqlCmd = Get-Command -Name Select-PWSQL -ErrorAction SilentlyContinue
Write-Host "Select-PWSQL available: $([bool]$sqlCmd)"

# 2. Check audit trail settings
$settings = Get-PWAuditTrailSettings -ErrorAction SilentlyContinue
$settings | Format-List

# 3. Sample recent audit records
$sql = "SELECT TOP 20 o_acttime, o_action, o_objtype, o_objguid, o_parentguid, o_itemname, o_itemdesc FROM dms_audt ORDER BY o_acttime DESC"
$results = Select-PWSQL -SQLSelectStatement $sql -ErrorAction SilentlyContinue
$results | Format-Table -AutoSize

# 4. Check action distribution in last 24 hours
$since = (Get-Date).AddHours(-24).ToString('yyyy-MM-dd HH:mm:ss')
$dist = Select-PWSQL "SELECT o_action, COUNT(*) AS cnt FROM dms_audt WHERE o_acttime > '$since' GROUP BY o_action ORDER BY cnt DESC"
$dist | Format-Table -AutoSize

# 5. Check for DOCUMENT_MODIFY events (description changes)
$mods = Select-PWSQL "SELECT TOP 10 o_acttime, o_itemname, o_itemdesc, o_textparam FROM dms_audt WHERE o_action = 1002 AND o_acttime > '$since' ORDER BY o_acttime DESC"
$mods | Format-Table -AutoSize

# 6. Check for DOCUMENT_ATTR events (attribute changes)
$attrs = Select-PWSQL "SELECT TOP 10 o_acttime, o_itemname, o_textparam, o_objguid FROM dms_audt WHERE o_action = 1003 AND o_acttime > '$since' ORDER BY o_acttime DESC"
$attrs | Format-Table -AutoSize
```

---

## 13. Summary of Recommendations

1. **Adopt hybrid architecture**: audit trail polling as primary trigger + periodic directory reconciliation as safety net.

2. **Use `Select-PWSQL` for polling**: Direct SQL queries against `dms_audt` with watermark/cursor approach give maximum control over query shape, filtering, and performance.

3. **Preserve existing pipeline**: `QC.Filters → QC.Triggers → QC.JobFactory → QC.Queue.Json → QC.Processors` is trigger-source-agnostic and should not change.

4. **Start with Phase 1 (observe-only)**: Run audit poller alongside existing watcher for 1-2 weeks to build confidence in audit coverage and latency.

5. **Introduce SQLite event database**: Low-friction, embedded, no additional infrastructure. Sufficient for Phase 1-3; migrate to SQL Server/PostgreSQL only if needed.

6. **Don't eliminate directory scanning entirely**: Keep as reconciliation at reduced frequency (every 6-24 hours vs. every tick).

7. **Run discovery script first**: Validate `Select-PWSQL` availability, audit trail settings, and event coverage before writing production code.

8. **Event coalescing is critical**: Multiple audit records per logical change are normal; coalesce before feeding into trigger pipeline.

9. **Local filesystem watching remains unchanged**: Audit trail is ProjectWise-only. `Get-ChildItem -Recurse` for local `watchFolders` stays.

10. **Future wins are significant**: State change detection, attribute monitoring, notification triggering, and workflow analytics all become straightforward once the event database exists.
