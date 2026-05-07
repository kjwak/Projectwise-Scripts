# QC Workflow Framework

## Purpose

This framework adds an optional ProjectWise workflow, state, and document-attribute writeback layer on top of the existing `QC_PREPEND` PDF processing flow. It is intentionally disabled by default because the final ProjectWise workflow, states, and environment attributes have not been finalized.

The framework is designed to support the red / green / blue QC process:

- **Red**: reviewer marks comments or issues.
- **Green**: designer or producer completes changes and responds.
- **Blue**: reviewer verifies corrections and closes the cycle.

`QC_PREPEND` remains responsible for PDF history, prepend, and overlay processing. The workflow framework only runs after successful processing and returns structured warnings/results without failing the job unless `qcWorkflow.strictMode` is enabled.

## Expected ProjectWise Administrator setup

A likely future setup is a dedicated ProjectWise workflow for QC PDF files, created and maintained in ProjectWise Administrator. The automation does **not** create this workflow or any environment attributes.

Recommended setup tasks:

1. Create a dedicated workflow for QC PDFs.
2. Create the desired workflow states.
3. Create or expose document attributes in the ProjectWise environment used by QC PDFs.
4. Confirm which `pwps` / `pwps_dab` cmdlets and parameters are approved for workflow assignment, state transitions, and attribute updates in the target datasource.
5. Pilot with dry-run writeback before enabling real writes.

## Recommended workflow states

The default configuration uses example names only. Rename them in `qcWorkflow` to match the final ProjectWise setup.

- `QC Received`
- `Redlines Issued`
- `Corrections In Progress`
- `Corrections Complete`
- `Backcheck In Progress`
- `Verified Closed`
- `Superseded`
- `Error / Needs Attention`

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
    "workflowName": "QC Review Workflow",
    "defaultStateAfterPrepend": "QC Received",
    "stateAfterSuccessfulPrepend": "Redlines Issued",
    "stateAfterFailedPrepend": "Error / Needs Attention",
    "autoAssignWorkflow": true,
    "autoSetState": true,
    "autoWriteAttributes": true,
    "attributeMap": {
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
        "stateName": "Redlines Issued",
        "statusValue": "Open"
      },
      "green": {
        "stageValue": "Green",
        "stateName": "Corrections Complete",
        "statusValue": "Pending Backcheck"
      },
      "blue": {
        "stageValue": "Blue",
        "stateName": "Verified Closed",
        "statusValue": "Closed"
      }
    }
  }
}
```


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
- Candidate write cmdlets only after confirmation: `Set-PWDocumentState`, `Set-PWDocumentWorkflow`, `Update-PWDocumentAttributes`.

Do not implement or enable real workflow/state/attribute writes until the discovery output confirms the exact cmdlets and parameter sets for the target ProjectWise environment.

## Dry-run testing steps

1. Confirm `qcWorkflow.enabled` is `false` for the initial deployment.
2. Enable the framework in a non-production or pilot configuration:
   - Set `qcWorkflow.enabled` to `true`.
   - Keep `qcWorkflow.dryRunWriteback` set to `true`.
   - Keep global `dryRun` as needed for broader processor dry-run testing.
3. Run `QC_PREPEND` against representative QC PDFs.
4. Review structured workflow writeback results attached to the processor result and log events:
   - `QC_WORKFLOW_DRYRUN`
   - `QC_WORKFLOW_ASSIGN_PLANNED`
   - `QC_WORKFLOW_STATE_PLANNED`
   - `QC_WORKFLOW_ATTRIBUTES_PLANNED`
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

- The final ProjectWise workflow, state, and attribute names are not known yet; all names must remain configurable.
- `pwps` / `pwps_dab` workflow cmdlet availability and signatures can vary by installation and datasource. The module isolates these calls in wrapper functions with TODO comments for site-specific confirmation.
- Missing workflow/state/attribute support logs warnings and returns non-fatal results by default.
- `strictMode` should only be enabled after configuration and ProjectWise cmdlet behavior are validated.
- Dry-run mode must be used before any production writeback is enabled.
