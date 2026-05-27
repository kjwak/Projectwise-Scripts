# QC Reporting

## Purpose

QC reporting is a read-only, attribute-first reporting layer for QC PDFs and status sets that remain in the existing ProjectWise `CADD/Sheets` folder structure. It does not require dedicated QC folders and does not treat workflow state as the source of truth for QC lifecycle.

## Architecture

Reporting scans existing project sheet folders and aggregates normalized QC values from ProjectWise document attributes first. Workflow state is optional secondary context only. Reports should prefer QC attributes first, then use workflow state as a fallback, cross-check, or dashboard grouping when attributes are missing or when state writeback has been explicitly enabled.

### Data Sources

**Primary — QC Attributes:**

- `QC_Active`
- `QC_Cycle_ID`
- `QC_Stage`
- `QC_Status`
- `QC_Last_Action_Date`
- `QC_Automation_Result`
- `QC_Automation_Error`

**Secondary — Workflow State:**

- Current document workflow/state properties returned by ProjectWise searches.
- `WorkflowState` property from `Get-PWDocumentsBySearchWithReturnColumns` results.
- `Get-PWFolderTreeDocumentStateCount` or state values from search results for dashboard context.

**Database — `sheet_index` Table (operational):**

The `sheet_index` table in SQL Server provides a live index of all documents in watched `CADD\Sheets` folders. It is populated during both audit-trail scans and full reconciliation scans in `Watch-QCTrigger.ps1`. This table includes:

- Document metadata (GUID, name, folder, extension)
- QC PDF pairing (`qc_pdf_guid`, `qc_pdf_name`)
- Ownership (`designer_email`, `reviewer_email` from `EM_Designer_Email` / `EM_Reviewer_Email` attributes)
- Workflow state (`pw_state_name`)
- QC lifecycle fields (`qc_stage`, `qc_status`)

The `v_sheet_status` view provides a project status overview. See `docs/database-telemetry.md` for the full schema.

**Database — `processing_jobs` Table:**

Job outcomes (type, status, duration, errors) are written to `processing_jobs` by `Write-QCJobTelemetry`. The `v_job_summary` view provides aggregated counts by type and status. See `docs/database-telemetry.md`.

**Power BI:**

Power BI connects directly to the `QC_Pipeline` database via SQL Server connector for dashboards and reporting visualizations.

## Reporting Buckets

The finalized initial ProjectWise QC workflow model defines these reporting buckets:

| Metric | Preferred source | Workflow state context |
| --- | --- | --- |
| `inProductionCount` | QC attributes indicating no active QC cycle, or production status when configured. | `In Production` |
| `qcReceivedCount` | QC active attributes with received/intake status when configured. | `QC Received` |
| `redlinesIssuedCount` | `QC_Stage = Red` or `QC_Status = Open`. | `Redlines Issued` |
| `correctionsInProgressCount` | QC attributes indicating designer correction ownership when configured. | `Corrections In Progress` |
| `correctionsCompleteCount` | `QC_Stage = Green` or `QC_Status = Pending Backcheck`. | `Corrections Complete` |
| `backcheckInProgressCount` | QC attributes indicating reviewer backcheck ownership when configured. | `Backcheck In Progress` |
| `verifiedClosedCount` | `QC_Stage = Blue` or `QC_Status = Closed`. | `Verified Closed` |
| `errorNeedsAttentionCount` | `QC_Automation_Error`, failed automation result, or configured error status. | `Error Needs Attention` |
| `staleOpenQcCount` | Active, non-closed QC documents older than `qcReporting.staleDays`. | Any non-closed active QC state |

Workflow state counts can be useful for validating ProjectWise adoption, but reports should not depend on state values alone. A document can remain in the project production workflow while still carrying authoritative QC attributes.

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
- `inProductionCount`
- `qcReceivedCount`
- `redlinesIssuedCount`
- `correctionsInProgressCount`
- `correctionsCompleteCount`
- `backcheckInProgressCount`
- `verifiedClosedCount`
- `errorNeedsAttentionCount`
- `staleOpenQcCount`
- `avgQcCycleDays`

## Limitations

- Reporting depends on configured QC attributes being populated consistently.
- Workflow state counts are secondary and may not match the QC lifecycle when project workflows are used for non-QC purposes or when state writeback remains disabled.
- Average cycle days requires cycle start and last action dates to be populated.
- ProjectWise search return-column support can vary by environment; the reporting module falls back to broader read-only searches when needed.
- `sheet_index` email and state fields depend on `Get-PWDocumentsBySearchWithReturnColumns` returning populated `.Attributes` bags. Not all documents have these attributes set.

## Related Documentation

- `docs/database-telemetry.md` — SQL Server schema, tables, views
- `docs/hybrid-polling.md` — how sheet_index is populated during watcher ticks
- `docs/pw-environment-email-attributes.md` — email attribute extraction details
