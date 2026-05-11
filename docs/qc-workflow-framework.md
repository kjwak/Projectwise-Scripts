# QC Workflow Framework

## Purpose

This framework adds an optional ProjectWise document-attribute writeback layer, with optional document-state integration, on top of the existing `QC_PREPEND` PDF processing flow. QC lifecycle values are an attribute-first overlay on the existing project workflow and existing `CADD/Sheets` folder structure.

The framework is designed to support the red / green / blue QC process:

- **Red**: reviewer marks comments or issues.
- **Green**: designer or producer completes changes and responds.
- **Blue**: reviewer verifies corrections and closes the cycle.

`QC_PREPEND` remains responsible for PDF history, prepend, and overlay processing. The workflow framework only runs after successful processing and returns structured warnings/results without failing the job unless `qcWorkflow.strictMode` is enabled.

## Expected ProjectWise Administrator setup

QC PDFs and status sets remain in the existing ProjectWise `CADD/Sheets` folder structure. The automation does **not** create dedicated QC folders, does **not** assign workflows to documents, and does **not** replace the project workflow. ProjectWise Administrator setup should focus first on QC environment attributes; optional state integration can be enabled later within the existing workflow.

Recommended setup tasks:

1. Create or expose QC document attributes in the ProjectWise environment used by sheet PDFs.
2. Confirm `Update-PWDocumentAttributes` parameter behavior in the target datasource.
3. Optionally identify existing workflow states that can represent QC milestones without replacing the project workflow.
4. Confirm `Set-PWDocumentState` and `Get-PWWorkflowStateLinks` behavior before enabling state-and-attribute mode.
5. Pilot with dry-run writeback before enabling real writes.

## Initial ProjectWise QC workflow model

The initial ProjectWise model keeps QC PDFs and status sets in the existing `CADD/Sheets` folder structure. The ProjectWise workflow state is optional secondary context; QC attributes remain the primary control layer used by prepend processing, reporting, and future dashboard logic.

| State | Meaning | Primary owner/responsibility |
| --- | --- | --- |
| `In Production` | Normal sheet production. This is the starting/default state before QC begins. | Designer/producer owns sheet production. |
| `QC Received` | QC set/file has been created or received and is ready for review. | Reviewer/QC coordinator owns review intake. |
| `Redlines Issued` | Reviewer has completed redline comments. Ownership moves to the designer/producer. | Designer/producer owns response and correction planning. |
| `Corrections In Progress` | Designer/producer is actively addressing reviewer comments. | Designer/producer owns corrections. |
| `Corrections Complete` | Designer/producer has marked responses/corrections complete. Ownership moves back to reviewer. | Reviewer owns backcheck readiness. |
| `Backcheck In Progress` | Reviewer is verifying the corrected work. | Reviewer owns verification. |
| `Verified Closed` | Reviewer has verified corrections. QC cycle is complete. | Reviewer/QC coordinator owns closure. |
| `Error Needs Attention` | Automation or process issue requires manual attention. | QC coordinator or automation owner investigates and restores normal processing. |

Recommended transition path:

```text
In Production
-> QC Received
-> Redlines Issued
-> Corrections In Progress
-> Corrections Complete
-> Backcheck In Progress
-> Verified Closed
```

Recommended loopback:

```text
Backcheck In Progress -> Corrections In Progress
```

Recommended reopening path:

```text
Verified Closed -> In Production
```

Recommended error path:

```text
Any active QC state -> Error Needs Attention
```

Do not automatically move documents from `In Production` to `QC Received` during normal prepend processing yet. Prepend processing should continue to write QC attributes only by default.

State automation is intentionally optional because ProjectWise workflows may already encode project production controls, review permissions, or client-specific gates. The QC framework must not make ProjectWise state the source of truth unless a project explicitly opts into state writeback. State changes only occur when `qcWorkflow.mode = StateAndAttributes` and `qcWorkflow.autoSetState = true`.

Attributes remain the primary control layer because they are less disruptive than workflow state changes, can be reported consistently across projects, can coexist with existing production workflows, and allow dry-run/pilot validation before any ProjectWise state writes are enabled.

## Recommended document attributes

The default attribute map uses example names only. Rename them in `qcWorkflow.attributeMap` to match final ProjectWise environment attribute names.

- `QC_Cycle_ID`
- `QC_Stage`
- `QC_Reviewer`
- `QC_Assigned_To`
- `QC_Last_Action_By`
- `QC_Last_Action_Date`
- `QC_Due_Date`
- `QC_Status`
- `QC_History_PDF_Path`
- `QC_Latest_Overlay_PDF_Path`
- `QC_Source_Document_Path`
- `QC_Automation_Last_Run`
- `QC_Automation_Result`
- `QC_Automation_Error`

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
    "productionStateName": "In Production",
    "receivedStateName": "QC Received",
    "correctionsInProgressStateName": "Corrections In Progress",
    "backcheckInProgressStateName": "Backcheck In Progress",
    "errorStateName": "Error Needs Attention",
    "defaultStateAfterPrepend": "QC Received",
    "stateAfterSuccessfulPrepend": "Redlines Issued",
    "stateAfterFailedPrepend": "Error Needs Attention",
    "autoSetState": false,
    "autoWriteAttributes": true,
    "attributeMap": {
      "qcActive": "QC_Active",
      "cycleId": "QC_Cycle_ID",
      "stage": "QC_Stage",
      "reviewer": "QC_Reviewer",
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
    },
    "stageMap": {
      "red": {
        "stageValue": "Red",
        "statusValue": "Open",
        "optionalStateName": "Redlines Issued"
      },
      "green": {
        "stageValue": "Green",
        "statusValue": "Pending Backcheck",
        "optionalStateName": "Corrections Complete"
      },
      "blue": {
        "stageValue": "Blue",
        "statusValue": "Closed",
        "optionalStateName": "Verified Closed"
      }
    }
  }
}
```



## Operational modes

### AttributesOnly (default)

`AttributesOnly` writes configured QC attributes only and does not attempt to change ProjectWise document state. This is the recommended and safest deployment mode.

### StateAndAttributes (optional)

`StateAndAttributes` still writes QC attributes first, then optionally validates and changes document state when `qcWorkflow.autoSetState = true`. It validates target state existence and transition links where possible using `Get-PWWorkflowStateLinks`, then writes with `Set-PWDocumentState`. It never changes workflow assignment.

## Capability discovery before enabling writeback

Before setting `qcWorkflow.dryRunWriteback` to `false`, run the read-only capability discovery script in the same PowerShell host/profile that will run the QC worker:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\discovery\Test-QCWorkflowCapabilities.ps1 -Pretty
```

The script attempts to import `pwps` and `pwps_dab`, reports loaded module versions, checks candidate ProjectWise cmdlets, records parameter/parameter-set metadata, and emits JSON shaped like:

```json
{
  "pwpsLoaded": true,
  "pwpsDabLoaded": true,
  "availableCmdlets": [],
  "missingCmdlets": [],
  "candidateWriteCmdlets": [],
  "readCapabilities": {},
  "writeCapabilities": {},
  "warnings": []
}
```

The discovery script is intentionally read-only. It does **not** assign workflows, change states, or write attributes. If a ProjectWise connection is already open, you may optionally provide a test document for read-only inspection of exposed workflow/state properties and environment attributes:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\discovery\Test-QCWorkflowCapabilities.ps1 `
  -FolderPath "Project\CADD\Sheets" `
  -DocumentName "example-qc.pdf" `
  -Pretty
```

Review these capability areas before enabling writeback:

- Workflow catalog/state reads: `Get-PWWorkflows`, `Get-PWWorkflowStateLinks`.
- Document lookup reads: `Get-PWDocumentsBySearch`, `Get-PWDocumentsBySearchExtended`, `Get-PWDocumentsBySearchWithReturnColumns`.
- Environment attribute reads: `Get-PWDocumentEAttributes`, `Get-PWEnvironmentColumns`.
- Confirmed document write cmdlets: `Set-PWDocumentState`, `Update-PWDocumentAttributes`.
- Confirmed workflow-state cmdlets for optional state integration: `Get-PWWorkflowStateLinks`, `Set-PWDocumentState`.
- Confirmed reporting/search cmdlets: `Get-PWFolderTreeDocumentStateCount`, `Get-PWDocumentsBySearchExtended`, `Get-PWDocumentsBySearchWithReturnColumns`, `Get-PWDocumentEAttributes`, `Get-PWEnvironmentColumns`.
- Confirmed missing direct document workflow cmdlet: `Set-PWDocumentWorkflow`; the QC framework never changes workflow assignment.

Do not enable real workflow/state/attribute writes until the discovery output confirms the exact parameter sets for the target ProjectWise environment.


## Controlled writeback testing

After capability discovery confirms the target environment exposes `Set-PWDocumentState` and `Update-PWDocumentAttributes`, use the controlled writeback harness for a single pilot document. The script requires both `-ConfirmWrites` and `-TestDocumentPath`, prints planned dry-run operations first, and supports state-only, attribute-only, or combined tests:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\discovery\Test-QCWorkflowWriteback.ps1 `
  -ConfirmWrites `
  -TestDocumentPath "Project\CADD\Sheets\example-qc.pdf" `
  -Mode StateOnly `
  -TargetState "Redlines Issued" `
  -ExpectedWorkflowName "QC Review Workflow" `
  -Pretty
```

Rollback is best-effort because ProjectWise workflow rules may prevent returning to the prior state and because environment attribute value shapes vary by datasource:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\discovery\Test-QCWorkflowWriteback.ps1 `
  -ConfirmWrites `
  -Rollback `
  -TestDocumentPath "Project\CADD\Sheets\example-qc.pdf" `
  -Mode Combined `
  -TargetState "Redlines Issued"
```

Never run the controlled writeback script against production documents until the workflow, valid state transitions, environment attributes, and rollback expectations are validated with the ProjectWise administrator.

## Dry-run testing steps

1. Confirm `qcWorkflow.enabled` is `false` for the initial deployment.
2. Enable the framework in a non-production or pilot configuration:
   - Set `qcWorkflow.enabled` to `true`.
   - Keep `qcWorkflow.dryRunWriteback` set to `true`.
   - Keep global `dryRun` as needed for broader processor dry-run testing.
3. Run `QC_PREPEND` against representative QC PDFs.
4. Review structured workflow writeback results attached to the processor result and log events:
   - `QC_WORKFLOW_DRYRUN`
   - `QC_WORKFLOW_VALIDATED` / `QC_WORKFLOW_EXPECTED_MISSING`
   - `QC_WORKFLOW_STATE_TRANSITION_VALID` / `QC_WORKFLOW_STATE_TRANSITION_INVALID`
   - `QC_WORKFLOW_STATE_PLANNED`
   - `QC_WORKFLOW_ATTRIBUTES_PLANNED`
   - `QC_WORKFLOW_STATE_WRITE_SUCCESS` and `QC_WORKFLOW_ATTRIBUTE_WRITE_SUCCESS` only after real writeback is enabled
5. Confirm planned workflow, state, and attribute values with the ProjectWise administrator.
6. Only after validation, set `qcWorkflow.dryRunWriteback` to `false` for a pilot.

## Rollout plan

- **Phase 1**: Framework added, disabled by default.
- **Phase 2**: Dry-run against real QC PDFs.
- **Phase 3**: Create workflow and attributes in ProjectWise Administrator.
- **Phase 4**: Enable writeback for a pilot project.
- **Phase 5**: Enable notifications and state-based reporting.
- **Phase 6**: Add dashboard/database integration.

## Risks and limitations

- The final ProjectWise state and attribute names are configurable. QC attributes are authoritative; workflow states are optional secondary context.
- `pwps` / `pwps_dab` cmdlet parameter sets can vary by installation and datasource. The module isolates attribute and optional state writes in wrapper functions and never reassigns workflows.
- Missing workflow/state/attribute support logs warnings and returns non-fatal results by default.
- `strictMode` should only be enabled after configuration and ProjectWise cmdlet behavior are validated.
- Dry-run mode must be used before any production writeback is enabled.
