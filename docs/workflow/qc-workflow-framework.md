# QC Workflow Framework

**Settings reference:** [`appsettings-reference.md` § qcWorkflow](appsettings-reference.md#qcworkflow) — every `appsettings.json` key, defaults, and enablement order.

## Purpose

This framework adds an optional ProjectWise document-attribute and workflow-state writeback layer on top of the existing `QC_PREPEND` PDF processing flow.

- **`QC_PREPEND`** remains the PDF history, prepend, and overlay engine. Production prepend runs through **`scripts/processing/Invoke-QCPrependPw.ps1`** when `qcPrepend.mode` is `projectWise` (committed default). `legacy/prepend_qc.ps1` is a deprecated forwarder only.
- **ProjectWise document state** is the QC lifecycle source of truth (do not duplicate lifecycle in `QC_Stage` or `stageMap`).
- **`QC_Process_Type`** identifies the active QC lane: `Production`, `Review`, or `Check` (canonical attribute key).
- **`QC_Review_Type`** mirrors the lane program via `reviewTypes` / `processTypes` config (same three values in TYPSA deployments).
- Role emails (`QC_Designer_Email`, `QC_Reviewer_Email`, `QC_Checker_Email`) drive assignment; the designer is the default corrector/originator.

The framework runs after successful prepend processing and returns structured warnings/results without failing the job unless `qcWorkflow.strictMode` is enabled.

## Three-lane QC model (current)

Each sheet stem can have up to three lane QC PDFs in ProjectWise:

| Lane | `QC_Process_Type` | PDF suffix |
|------|-------------------|------------|
| Production | `Production` | `*-prod.pdf` |
| Review | `Review` | `*-rev.pdf` |
| Check | `Check` | `*-chk.pdf` |

Lane states are **independent** by default (`QCProcess.EnableLegacySiblingStateSync: false`). Post-initial-prepend lane split is handled by `Sync-PWPostInitialPrependLaneStates` in `PW.Discovery.psm1`.

### Fast `DOCUMENT_STATE` prepend enqueue

The watcher stays on the QC server and remains the only enqueue path. With `qcPrepend.fastAuditEnqueue` (default **false** in committed `appsettings.json`):

1. Identify a Sheets-folder stem PDF `DOCUMENT_STATE` candidate (not a lane PDF, DGN, status-set output, or automation actor).
2. Read live workflow state once (`Get-PWDocumentWorkflowStateName`).
3. If the state is **Initiate Origination** (`qcInitiated`) or **Initiate Verification** (`qcFinalizing`), enqueue `QC_PREPEND` immediately. Do not wait for sibling maps, attr sync, or live process-type.
4. After every candidate in the tick has been considered, run `Sync-PWAssociatedSheetWorkflowState` as a second pass so `sheet_index` / telemetry still update.
5. The worker (`Invoke-QCPrependProcessor`) confirms live state, DGN pair, and lane **before** spawning overlay **when live ProjectWise state is readable**. A stale or echo audit with a real non-trigger state becomes a successful skip (`QC_PREPEND_SKIPPED_NOT_QC_INITIATED`, `QC_PREPEND_SKIPPED_NO_DGN_PAIR`, `QC_PREPEND_SKIPPED_NO_PROCESS_TYPE`). Empty/unreadable live state (typical on remote processor hosts, which do not `Open-PWConnection` in the parent) must **not** skip-succeed — the prepend child opens PW and is the gate. Writeback-only resume (`prepend_complete` / `writeback_running`) does not re-run that preflight.

Enable the flag on the QC server overlay after soak; do not flip the committed default without operator sign-off.

When legacy sibling sync is off, stem/DGN `DOCUMENT_STATE` audit handling must **not** copy the stem state onto lane QC PDFs (`*-prod/-chk/-rev`). `Sync-PWAssociatedSheetWorkflowState` filters lane members out of the sync set and skips any remaining lane write attempts (`WATCH_SHEET_STATE_SYNC_LANE_SKIPPED`). This prevents a successful post-prepend **Originated** lane state from being overwritten when a delayed stem audit still shows **In Development**.

**Lane-independent prepend telemetry** (`Invoke-QCSheetGroupWorkflowTransition` with `laneIndependentInitialPrepend`):

- Only the **active** lane PDF and (on initial prepend) stem/DGN reference members receive `transition_events` / `qc_workflow_events`.
- Inactive sibling lane PDFs (for example `-prod.pdf` during a review prepend) must not record state transitions — PW writeback already skips them; telemetry mirrors that rule.
- `qc_workflow_events.qc_review_type` for lane PDFs is resolved from the **filename** (`*-prod/-rev/-chk`), not from shared prepend job context (stem `QC_Process_Type`).

### Legacy compatibility (`*-qc.pdf`)

The single-lane `*-qc.pdf` naming pattern is **legacy**. Code still supports it as a bridge (for example standalone `Invoke-QCPrependPw.ps1` without strict lane params, normalization helpers, and older `sheet_index` rows). **Do not** treat `*-qc.pdf` as the current authoritative lane document. New work should use `*-prod.pdf` / `*-rev.pdf` / `*-chk.pdf`.

SQL telemetry registers lane PDFs in `sheet_package_qc_pdfs` (see [`database-telemetry.md`](../data/database-telemetry.md)).

## TYPSA QC lifecycle states

Committed `appsettings.json` uses the **TYPSA QC** workflow. State names must match ProjectWise labels exactly.

| Config key | TYPSA state | Meaning |
|------------|-------------|---------|
| `production` | `In Development` | Normal sheet production before QC intake |
| `qcInitiated` | `Initiate Origination` | Designer requested QC intake; triggers `QC_PREPEND` |
| `qcReceived` / `readyForQc` | `Originated` | Automation completed intake; reviewer/checker may begin |
| `redlinesReceived` | `Redlines Received` | Review complete; designer owns response |
| `qcFinalizing` | `Initiate Verification` | Designer marked corrections complete; verification pending |
| `readyForVerification` | `Ready for Verification` | Awaiting reviewer/checker verification |
| `complete` | `Verified` | QC cycle finished |
| `error` | `Error Needs Attention` | Automation or process issue |

Recommended transition path:

```text
In Development
-> Initiate Origination
-> Originated
-> Redlines Received
-> Initiate Verification
-> Ready for Verification
-> Verified
```

Loopback: `Ready for Verification` → `Redlines Received` when verification fails.

Reopen: `Verified` → `In Development` when a new cycle starts.

Error path: any active QC state → `Error Needs Attention`.

### Legacy state names (historical)

Pre-TYPSA deployments used labels such as `In Production`, `QC Initiated`, `QC Received`, `Ready for QC`, `Review In Progress`, `Corrections In Progress`, `Verification In Progress`, and `QC Complete`. Config keys like `readyForQc` and `qcReceived` are retained for backward compatibility but map to TYPSA names in committed config. Disabled notification event keys (`QC Received`, `Ready for QC`, `QC Complete`) remain in `appsettings.json` for environments that have not migrated.

## Process types and assignment

`Resolve-QCWorkflowAssignee` sets `QC_Assigned_To` from **state + process type** (`QC_Process_Type` / `QC_Review_Type`):

| State | Production / Review | Check |
|-------|-------------------|-------|
| In Development | Designer | Designer |
| Originated | Reviewer | Checker |
| Redlines Received | Designer | Designer |
| Initiate Verification | Reviewer | Checker |
| Ready for Verification | Reviewer | Checker |
| Verified | — | — |

### Legacy review type labels (historical)

Older configs used `Production QC`, `Peer Review`, and `Independent Check` as `QC_Review_Type` values. TYPSA config maps these config keys to `Production`, `Review`, and `Check` respectively.

## Recommended document attributes

Map logical keys in `qcWorkflow.attributeMap` to your ProjectWise environment names:

- `QC_Active`, `QC_Process_Type`, `QC_Review_Type`, `QC_Cycle_ID`, `QC_Cycle_Number`
- `QC_Designer_Email`, `QC_Reviewer_Email`, `QC_Checker_Email`
- `QC_Assigned_To`, `QC_Last_Action_By`, `QC_Last_Action_Date`, `QC_Status`
- `QC_History_PDF_Path`, `QC_Latest_Overlay_PDF_Path`, `QC_Source_Document_Path`
- `QC_Automation_Last_Run`, `QC_Automation_Result`, `QC_Automation_Error`

Do **not** use `QC_Stage` or red/green/blue `stageMap` — they are deprecated and ignored.

## Configuration example (TYPSA)

Excerpt aligned with committed `appsettings.json`:

```json
{
  "qcWorkflow": {
    "enabled": true,
    "strictMode": false,
    "dryRunWriteback": false,
    "mode": "StateAndAttributes",
    "workflowName": "TYPSA QC",
    "expectedWorkflowName": "TYPSA QC",
    "states": {
      "production": "In Development",
      "qcInitiated": "Initiate Origination",
      "qcReceived": "Originated",
      "readyForQc": "Originated",
      "redlinesReceived": "Redlines Received",
      "qcFinalizing": "Initiate Verification",
      "readyForVerification": "Ready for Verification",
      "complete": "Verified",
      "error": "Error Needs Attention"
    },
    "processTypes": {
      "production": "Production",
      "check": "Check",
      "review": "Review"
    },
    "reviewTypes": {
      "productionQc": "Production",
      "peerReview": "Review",
      "independentCheck": "Check"
    },
    "defaultProcessType": "production",
    "defaultReviewType": "Production",
    "defaultStateAfterPrepend": "Originated",
    "stateAfterSuccessfulPrepend": "Originated",
    "stateAfterFailedPrepend": "Error Needs Attention",
    "defaultPrependTrigger": "initialQcPdf",
    "stateAfterPrependByTrigger": {
      "initialQcPdf": "Originated",
      "finalQcComplete": "Ready for Verification"
    },
    "autoSetState": true,
    "autoWriteAttributes": true,
    "attributeMap": {
      "processType": "QC_Process_Type",
      "reviewType": "QC_Review_Type"
    }
  }
}
```

### Deprecated configuration (backward compatible)

Legacy keys still load with warnings (non-fatal unless `strictMode`):

- `productionStateName`, `receivedStateName`, etc. → mapped into `states.*` when `states` is omitted
- `stageMap` (red/green/blue) — ignored
- `attributeMap.stage` / `reviewer` — not written; use state + role emails

Prefer `states.*`, `processTypes.*`, and `reviewTypes.*` when present.

## Operational modes

### AttributesOnly

Writes QC attributes only. Does not change ProjectWise document state.

### StateAndAttributes (TYPSA production)

Writes attributes first, then optionally sets state when `autoSetState = true`, using `Resolve-QCWorkflowStateAfterPrepend` (trigger/context), `stateAfterFailedPrepend` on failure, or explicit `Context.targetState`.

### Post-prepend state by trigger

Successful `QC_PREPEND` target state depends on `metadata.prependTrigger` on the job (or `Context.prependTrigger`):

| Trigger key | Typical use | Default target state (TYPSA) |
|-------------|-------------|------------------------------|
| `initialQcPdf` | Initial lane QC PDF creation (e.g. `Initiate Origination` intake) | `Originated` |
| `finalQcComplete` | Final production prepend before verification | `Ready for Verification` |

Configure targets in `qcWorkflow.stateAfterPrependByTrigger`. Unknown triggers fall back to `stateAfterSuccessfulPrepend` (`Originated` in committed config).

Explicit flags on job metadata (alternative to `prependTrigger` string): `finalQcComplete=true`, `reviewerRedlineUpdate=true`, `correctionComplete=true`.

## Capability discovery before enabling writeback

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\discovery\Test-QCWorkflowCapabilities.ps1 -Pretty
```

Pilot writeback:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\discovery\Test-QCWorkflowWriteback.ps1 -ConfirmWrites
```

## PowerShell API (`modules/Workflow/QC.Workflow.psm1`)

| Function | Role |
| --- | --- |
| `Get-QCWorkflowSettings` | Normalize config (states, process types, review types, attribute map) |
| `Get-QCWorkflowDeprecationWarnings` | List deprecated keys in raw config |
| `Resolve-QCWorkflowAssignee` | State + process type → assignee email |
| `Get-QCWorkflowStateName` | Resolve configured state label by key |
| `Test-QCWorkflowConfig` | Validate enabled workflow config |
| `Invoke-QCWorkflowWriteback` | Post-prepend attribute/state writeback |
| `Set-PWQCAttributes` / `Set-PWQCWorkflowState` | Low-level PW writes |

## Rollout checklist

1. Confirm PW attributes and TYPSA workflow states exist; align `attributeMap` and `states.*` names.
2. `dryRunWriteback: true` — inspect planned writes in logs.
3. Pilot `dryRunWriteback: false` on one project.
4. Validate lane PDF naming (`*-prod/-rev/-chk`) before relying on lane-independent state split.

### Expanding QC to more folders (sheet_index already populated)

Status-set reconciliation may already have rows in `sheet_index` for the new watch roots. Baseline suppression only skips telemetry when **`pw_state_name` was empty** and PW still shows **In Development** (`auditPoller.workflowTriggers.suppressBaselineIndexStateTransition`, default **true**).

To avoid replaying historical `DOCUMENT_STATE` from the audit backlog when you add roots:

1. Set `auditPoller.workflowTriggers.processingGoLiveUtc` to the UTC time you enable QC.
2. Add the new `projectWise.watchList.roots` paths; do not reset `audit-capture-watermark.txt` unless you intend a deliberate backfill.
3. Optionally run one scheduled full reconciliation so `pw_state_name` on existing rows matches PW before go-live.
4. After `audit_events` with `processed = 0` drains, clear `processingGoLiveUtc` (empty string).

Log codes: `WATCH_AUDIT_BASELINE_STATE_SUPPRESSED`, `WATCH_AUDIT_STATE_SKIPPED_GO_LIVE`.

## Tests

- `test/powershell/test_qc_workflow.ps1` — settings, assignment, deprecation warnings, attribute mapping
- `test/python/test_qc_workflow_config_defaults.py` — repo defaults and module shape
