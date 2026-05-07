# QC Reporting

## Purpose

QC reporting is a read-only, attribute-first reporting layer for QC PDFs and status sets that remain in the existing ProjectWise `CADD/Sheets` folder structure. It does not require dedicated QC folders and does not treat workflow state as the source of truth for QC lifecycle.

## Architecture

Reporting scans existing project sheet folders and aggregates normalized QC values from ProjectWise document attributes first. Workflow state is optional secondary context only.

Primary data source:

- `QC_Active`
- `QC_Cycle_ID`
- `QC_Stage`
- `QC_Status`
- `QC_Last_Action_Date`
- `QC_Automation_Result`
- `QC_Automation_Error`

Optional secondary data source:

- Current document workflow/state properties returned by ProjectWise searches.
- `Get-PWFolderTreeDocumentStateCount` or state values from search results for dashboard context.

## Module

`modules/QC.Reporting.psm1` is read-only and provides:

- `Get-QCReportingSettings`
- `Get-QCReportingDocuments`
- `ConvertTo-QCReportingDocument`
- `New-QCReportingSnapshot`
- `Write-QCReportingSnapshot`
- `Invoke-QCReportingScan`
- `New-QCReportingScanJob`

## Scheduled job type

`QC_REPORTING_SCAN` is an independent scheduled/reporting job type. It should be enqueued by a scheduler using a dedupe key that prevents duplicate pending/running scans for the same project/time bucket. The current implementation writes JSON snapshots first; database ingestion can be added later.

Snapshot path:

```text
metrics/qc/<timestamp>/<project>.json
```

## Metrics

Each snapshot includes:

- `qcActiveCount`
- `qcOpenCount`
- `qcPendingBackcheckCount`
- `qcClosedCount`
- `qcErrorCount`
- `staleQcCount`
- `avgQcCycleDays`

## Limitations

- Reporting depends on configured QC attributes being populated consistently.
- Workflow state counts are secondary and may not match the QC lifecycle when project workflows are used for non-QC purposes.
- Average cycle days requires cycle start and last action dates to be populated.
- ProjectWise search return-column support can vary by environment; the reporting module falls back to broader read-only searches when needed.
