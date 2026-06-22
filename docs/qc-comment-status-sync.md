# QC Comment-Status Sync

## Overview

`QC_COMMENT_STATUS_SYNC` extends the existing **watcher → queue → worker → processor** pipeline. When a ProjectWise **lane QC PDF** (`*-prod.pdf`, `*-rev.pdf`, `*-chk.pdf`) is updated, the automation exports the PDF, extracts comments, decides the workflow state, persists telemetry, optionally updates PW state, and sends notifications.

### Legacy trigger (`*-qc.pdf`)

Comment-sync trigger rules still match legacy `*-qc.pdf` filenames for backward compatibility. New deployments should use lane PDF suffixes.

## Module boundaries

| Module | Role |
|--------|------|
| `QC.PdfExport.psm1` | PW export to staging only |
| `QC.CommentExtract.psm1` + `overlay/qc_pdf_comments.py` | Parser adapter (Bluebeam logic isolated in Python) |
| `QC.CommentStatusDecision.psm1` | **Pure** `Resolve-QCCommentWorkflowState` |
| `QC.CommentSync.Database.psm1` | SQL persistence + `qc_workflow_events` |
| `QC.CommentSync.State.psm1` | Thin wrapper over `Set-PWQCWorkflowState` |
| `QC.CommentSync.Notifications.psm1` | Email routing |
| `QC.CommentStatusProcessor.psm1` | Thin orchestrator (six steps only) |

## Orchestrator steps

1. export → 2. extract → 3. decide → 4. persist → 5. update state → 6. notify

## Configuration (`qcCommentSync`)

- `enabled`, `processorVersion`, `stagingRoot`, `retainTempFiles`
- `pythonExecutable`, `commentExtractScript`, `parserVersion`
- `auditActions` — default: `DOCUMENT_MODIFY`, `DOCUMENT_FILE_REP`, `DOCUMENT_VERSION`
- `reviewerAuthorPatterns`, `statusMappings`, `targetStates` — use TYPSA state names (`Redlines Received`, `Ready for Verification`, etc.)
- Per watch root: `enableQcCommentSync` (defaults to `enableQcPrepend` when omitted)

## Trigger rule

- **id:** `qc-comment-status-pw`
- **jobType:** `QC_COMMENT_STATUS_SYNC`
- **match:** `*.pdf` + filename `(?i)-qc\.pdf$` (legacy); lane suffix rules also apply via watcher audit path

Dedupe includes file hash (pseudo-hash at watch time; content SHA256 after export).

## `processing_jobs` telemetry

Queue jobs remain type `QC_COMMENT_STATUS_SYNC`. When the worker records outcomes via `Write-QCJobTelemetry`, `processing_jobs.job_type` is stored as **`QC_STATE`** (see `Get-QCProcessingJobType` in `Core.Database.psm1`).

## Dry-run matrix

| Action | Default when `dryRun: true` |
|--------|---------------------------|
| PW state | Planned only (`Set-PWQCWorkflowState -DryRun`) |
| Email | Planned / Mock (`notifications.dryRun`) |
| Database | No writes unless `database.allowWritesInDryRun: true` |

## Bluebeam / parser calibration

Standard PyMuPDF `annot.info` does not expose all Bluebeam review statuses. The Python parser also reads IRT-linked `/State` keys. **Calibrate against real project PDFs** before production trust.

## Database (schema 1.2.0)

- `qc_comment_runs` — one row per processor run
- `qc_comments` — annotation snapshot
- `qc_comment_status_history` — status transitions
- `qc_workflow_events` — workflow/audit events

## Tests

- `test/test_qc_comment_trigger.ps1`
- `test/test_qc_comment_dedupe.ps1`
- `test/test_qc_comment_status_decision.ps1`
- `test/test_qc_comment_notification_routing.ps1`
- `test/test_qc_comment_sync_orchestrator.ps1`
- `tests/test_qc_pdf_comments.py`
