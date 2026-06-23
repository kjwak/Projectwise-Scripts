# QC Reporting

## Purpose

QC reporting is a read-only, attribute-first reporting layer for QC PDFs and status sets that remain in the existing ProjectWise `CADD/Sheets` folder structure. It does not require dedicated QC folders and does not treat workflow state as the sole source of truth for QC lifecycle detail.

## Architecture

Reporting scans existing project sheet folders and aggregates normalized QC values from ProjectWise document attributes first. **ProjectWise workflow state** is the preferred lifecycle dimension for bucket counts; QC attributes (`QC_Process_Type`, `QC_Review_Type`, role emails, automation fields) provide detail. Do not use deprecated `QC_Stage` or red/green/blue stage maps.

### Three-lane QC PDFs (current)

Lane QC PDFs use `*-prod.pdf`, `*-rev.pdf`, and `*-chk.pdf` suffixes keyed by `QC_Process_Type`. Legacy `*-qc.pdf` rows may still appear in older `sheet_index` data.

### Data Sources

**Primary — QC Attributes:**

- `QC_Active`, `QC_Process_Type`, `QC_Cycle_ID`, `QC_Review_Type`
- `QC_Status`, `QC_Last_Action_Date`
- `QC_Automation_Result`, `QC_Automation_Error`

**Secondary — Workflow State:**

- Current document workflow/state properties from ProjectWise searches.
- `WorkflowState` from `Get-PWDocumentsBySearchWithReturnColumns`.
- Folder state counts for dashboard context.

**Database — `sheet_index` and `sheet_package_qc_pdfs`:**

The `sheet_index` table indexes documents in watched `CADD\Sheets` folders. Lane registry rows live in `sheet_package_qc_pdfs` (canonical for new logic). Legacy columns `qc_pdf_guid` / `qc_pdf_name` on `sheet_index` may still be populated for backward compatibility.

The `v_sheet_status` and `v_sheet_package_status` views provide project status overviews. See `docs/data/database-telemetry.md`.

**Database — `processing_jobs`:**

Job outcomes are written by `Write-QCJobTelemetry`. See `v_job_summary` in `docs/data/database-telemetry.md`.

**Power BI:**

Power BI connects to the `QC_Pipeline` database via SQL Server connector.

## Reporting Buckets (TYPSA)

| Metric | Preferred source | TYPSA workflow state |
| --- | --- | --- |
| `inProductionCount` | Documents in production workflow state. | `In Development` |
| `readyForQcCount` | Documents ready for QC intake. | `Originated` |
| `redlinesReceivedCount` | Review complete; redlines returned to designer. | `Redlines Received` |
| `correctionsReceivedCount` | Designer returned corrections; verification pending. | `Initiate Verification` |
| `qcFinalizingCount` | Verification gate / final prepend pending. | `Ready for Verification` |
| `qcCompleteCount` | Completed QC cycles. | `Verified` |
| `errorNeedsAttentionCount` | Automation/process errors. | `Error Needs Attention` |
| `staleOpenQcCount` | Active QC documents older than `qcReporting.staleDays`, not in `Verified`. | Non-complete active states |

**Legacy bucket labels** (`In Production`, `Ready for QC`, `QC Finalizing`, `QC Complete`) map to the TYPSA states above in migrated deployments.

## Module

`modules/Reporting/QC.Reporting.psm1` is read-only and provides:

- `Get-QCReportingSettings`, `Get-QCReportingDocuments`, `ConvertTo-QCReportingDocument`
- `New-QCReportingSnapshot`, `Write-QCReportingSnapshot`, `Invoke-QCReportingScan`
- `Get-QCPackageReportingRows` — queries SQL `v_sheet_package_status` (not the unwired in-memory `QC.Package*` modules)

## Scheduled job type

`QC_REPORTING_SCAN` writes JSON snapshots under `metrics/qc/<timestamp>/<project>.json`.

## Limitations

- Reporting depends on configured QC attributes being populated consistently.
- Workflow state counts are secondary when state writeback is disabled.
- `sheet_index` email and state fields depend on PW search return columns.
- Lane-level metrics require `sheet_package_qc_pdfs` to be populated (post-prepend lane registry sync).

## Related Documentation

- `docs/data/database-telemetry.md` — SQL Server schema, tables, views
- `docs/architecture/hybrid-polling.md` — how sheet_index is populated during watcher ticks
- `docs/workflow/pw-environment-email-attributes.md` — email attribute extraction details
