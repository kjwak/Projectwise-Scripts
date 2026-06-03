# QC Workflow Framework

**Settings reference:** [`appsettings-reference.md` § qcWorkflow](appsettings-reference.md#qcworkflow) — every `appsettings.json` key, defaults, and enablement order. Repo default: `"qcWorkflow": { "enabled": false, … }`.

## Purpose

This framework adds an optional ProjectWise document-attribute writeback layer, with optional document-state integration, on top of the existing `QC_PREPEND` PDF processing flow.

- **`QC_PREPEND`** remains the PDF history, prepend, and overlay engine.
- **ProjectWise document state** is the QC lifecycle source of truth (do not duplicate lifecycle in `QC_Stage` or `stageMap`).
- **`QC_Review_Type`** identifies the review program (`Production QC`, `Peer Review`, `Independent Check`).
- Role emails (`QC_Designer_Email`, `QC_Reviewer_Email`, `QC_Checker_Email`) drive assignment; the designer is the default corrector/originator.

The framework runs after successful prepend processing and returns structured warnings/results without failing the job unless `qcWorkflow.strictMode` is enabled.

## Shared QC lifecycle states

| State | Meaning | Typical assignee (see review type) |
| --- | --- | --- |
| `In Production` | Normal sheet production before QC. | Designer |
| `QC Initiated` | Designer requested QC intake; triggers `QC_PREPEND` / `QC_RENDITION`. | Designer |
| `QC Received` | Automation completed intake (overlay/history); triggers reviewer email. | Reviewer or checker |
| `Ready for QC` | Optional: reviewer may start review (if used in PW). | Reviewer or checker |
| `Review In Progress` | Active QC review. | Reviewer or checker |
| `Redlines Issued` | Review complete; comments/redlines delivered; designer owns response before corrections start. | Designer |
| `Corrections In Progress` | Designer/producer actively addressing comments. | Designer |
| `Verification In Progress` | Reviewer/checker verifying corrections. | Reviewer or checker |
| `QC Complete` | QC cycle finished. | No active assignee |
| `Error Needs Attention` | Automation or process issue. | Manual follow-up |

Recommended transition path:

```text
In Production
-> QC Initiated
-> QC Received
-> Ready for QC (optional)
-> Review In Progress
-> Redlines Issued
-> Corrections In Progress
-> Verification In Progress
-> QC Complete
```

`Redlines Issued` separates **review activity** from **correction activity** and supports future duration metrics (review time, designer response time, correction time, verification time).

Loopback: `Verification In Progress` → `Corrections In Progress` when verification fails.

Reopen: `QC Complete` → `In Production` when a new cycle starts.

Error path: any active QC state → `Error Needs Attention`.

State automation is optional (`qcWorkflow.mode = StateAndAttributes` and `autoSetState = true`). By default only attributes are written (`AttributesOnly`, `autoSetState = false`).

## Review types and assignment

`Resolve-QCWorkflowAssignee` sets `QC_Assigned_To` from **state + `QC_Review_Type`**:

| State | Production QC / Peer Review | Independent Check |
| --- | --- | --- |
| In Production | Designer | Designer |
| Ready for QC | Reviewer | Checker |
| Review In Progress | Reviewer | Checker |
| Redlines Issued | Designer | Designer |
| Corrections In Progress | Designer | Designer |
| Verification In Progress | Reviewer | Checker |
| QC Complete | — | — |

## Recommended document attributes

Map logical keys in `qcWorkflow.attributeMap` to your ProjectWise environment names:

- `QC_Active`, `QC_Review_Type`, `QC_Cycle_ID`, `QC_Cycle_Number`
- `QC_Designer_Email`, `QC_Reviewer_Email`, `QC_Checker_Email`
- `QC_Assigned_To`, `QC_Last_Action_By`, `QC_Last_Action_Date`, `QC_Status`
- `QC_History_PDF_Path`, `QC_Latest_Overlay_PDF_Path`, `QC_Source_Document_Path`
- `QC_Automation_Last_Run`, `QC_Automation_Result`, `QC_Automation_Error`

Do **not** use `QC_Stage` or red/green/blue `stageMap` — they are deprecated and ignored.

## Configuration example

```json
{
  "qcWorkflow": {
    "enabled": false,
    "strictMode": false,
    "dryRunWriteback": true,
    "mode": "AttributesOnly",
    "workflowName": "",
    "expectedWorkflowName": "",
    "states": {
      "production": "In Production",
      "readyForQc": "Ready for QC",
      "reviewInProgress": "Review In Progress",
      "redlinesIssued": "Redlines Issued",
      "correctionsInProgress": "Corrections In Progress",
      "verificationInProgress": "Verification In Progress",
      "complete": "QC Complete",
      "error": "Error Needs Attention"
    },
    "reviewTypes": {
      "productionQc": "Production QC",
      "peerReview": "Peer Review",
      "independentCheck": "Independent Check"
    },
    "defaultReviewType": "Production QC",
    "defaultStateAfterPrepend": "Ready for QC",
    "stateAfterSuccessfulPrepend": "Ready for QC",
    "stateAfterFailedPrepend": "Error Needs Attention",
    "defaultPrependTrigger": "initialQcPdf",
    "stateAfterPrependByTrigger": {
      "initialQcPdf": "Ready for QC",
      "reviewerRedlineUpdate": "Redlines Issued",
      "designerCorrectionComplete": "Verification In Progress"
    },
    "autoSetState": false,
    "autoWriteAttributes": true,
    "attributeMap": {
      "qcActive": "QC_Active",
      "reviewType": "QC_Review_Type",
      "cycleId": "QC_Cycle_ID",
      "cycleNumber": "QC_Cycle_Number",
      "designerEmail": "QC_Designer_Email",
      "reviewerEmail": "QC_Reviewer_Email",
      "checkerEmail": "QC_Checker_Email",
      "assignedTo": "QC_Assigned_To",
      "lastActionBy": "QC_Last_Action_By",
      "lastActionDate": "QC_Last_Action_Date",
      "status": "QC_Status",
      "historyPdfPath": "QC_History_PDF_Path",
      "latestOverlayPdfPath": "QC_Latest_Overlay_PDF_Path",
      "sourceDocumentPath": "QC_Source_Document_Path",
      "automationLastRun": "QC_Automation_Last_Run",
      "automationResult": "QC_Automation_Result",
      "automationError": "QC_Automation_Error"
    }
  }
}
```

### Deprecated configuration (backward compatible)

Legacy keys still load with warnings (non-fatal unless `strictMode`):

- `productionStateName`, `receivedStateName`, `correctionsInProgressStateName`, `backcheckInProgressStateName`, `errorStateName` → mapped into `states.*` when `states` is omitted
- `stageMap` (red/green/blue) — ignored
- `attributeMap.stage` / `reviewer` — not written; use state + role emails

Prefer `states.*` and `reviewTypes.*` when present.

## Operational modes

### AttributesOnly (default)

Writes QC attributes only. Does not change ProjectWise document state.

### StateAndAttributes (optional)

Writes attributes first, then optionally sets state when `autoSetState = true`, using `Resolve-QCWorkflowStateAfterPrepend` (trigger/context), `stateAfterFailedPrepend` on failure, or explicit `Context.targetState`.

### Post-prepend state by trigger

Successful `QC_PREPEND` does **not** always move documents to `Redlines Issued`. Set `metadata.prependTrigger` on the job (or `Context.prependTrigger`):

| Trigger key | Typical use | Default target state |
| --- | --- | --- |
| `initialQcPdf` | Initial QC PDF creation / rendition (e.g. `QC_Initiated` intake) | `Ready for QC` |
| `reviewerRedlineUpdate` | Reviewer issued redlines / comment overlay update | `Redlines Issued` |
| `designerCorrectionComplete` | Designer explicitly marked corrections complete | `Verification In Progress` |

Configure targets in `qcWorkflow.stateAfterPrependByTrigger`. Unknown triggers fall back to `stateAfterSuccessfulPrepend` (`Ready for QC` by default).

Explicit flags on job metadata (alternative to `prependTrigger` string): `reviewerRedlineUpdate=true`, `correctionComplete=true` / `designerCorrectionComplete=true`.

## Capability discovery before enabling writeback

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\discovery\Test-QCWorkflowCapabilities.ps1 -Pretty
```

Pilot writeback:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\discovery\Test-QCWorkflowWriteback.ps1 -ConfirmWrites
```

## PowerShell API (`modules/QC.Workflow.psm1`)

| Function | Role |
| --- | --- |
| `Get-QCWorkflowSettings` | Normalize config (states, review types, attribute map) |
| `Get-QCWorkflowDeprecationWarnings` | List deprecated keys in raw config |
| `Resolve-QCWorkflowAssignee` | State + review type → assignee email |
| `Get-QCWorkflowStateName` | Resolve configured state label by key |
| `Test-QCWorkflowConfig` | Validate enabled workflow config |
| `Invoke-QCWorkflowWriteback` | Post-prepend attribute/state writeback |
| `Set-PWQCAttributes` / `Set-PWQCWorkflowState` | Low-level PW writes |

## Rollout checklist

1. `enabled: false` — production default.
2. `enabled: true`, `dryRunWriteback: true` — inspect planned writes in logs.
3. Confirm PW attributes exist; align `attributeMap` names.
4. Pilot `dryRunWriteback: false` on one project.
5. Optionally enable `StateAndAttributes` + `autoSetState` after workflow transitions are validated.

## Tests

- `test/test_qc_workflow.ps1` — settings, assignment, deprecation warnings, attribute mapping
- `tests/test_qc_workflow_config_defaults.py` — repo defaults and module shape
