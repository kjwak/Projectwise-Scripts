# Scripts Organization Review (Reassessed)

Date: 2026-05-06
Branch baseline: `7d260c2`

## Reassessment summary
After re-checking the latest branch, the original conclusion still stands: the repo has a good entrypoint/module split, but **script-level helper duplication is still high** and several algorithmic blocks should be module-owned.

## Evidence from latest branch
- `_ToHashtable` exists in at least: `run_prepend_qc`, `Test-PWConnection`, `Start-QCPipelineDashboard`, `Reconcile-QCStatusSets`, `Run-QCProcessor`, `Show-QCStatus`, `Watch-QCTrigger`.
- `_Read-AppSettings*` and `_Log` patterns appear across `Run-QCProcessor`, `Watch-QCTrigger`, `Start-QCPipelineDashboard`, `run_prepend_qc`, `Reconcile-QCStatusSets`.
- `Watch-QCTrigger.ps1` still contains heavy reusable helpers:
  - hashing (`_Sha256FileHex`, `_Sha256TextHex`),
  - PW metadata extraction (`_PW-Get*`),
  - large folder discovery algorithm (`_PW-DiscoverSheetsFoldersUnderRoot`).
- `Run-QCProcessor.ps1` still owns retry/transition policy via `_Move-QCJobWithLockRetries`.

## Updated recommendations (same direction, tightened scope)

### Priority 1: runtime foundation module — DONE
~~Create `modules/Core/Core.Runtime.psm1`~~ with:
- `ConvertTo-HashtableDeep`
- `Read-QCAppSettings`
- `Write-QCJsonLog`
- `Get-QCTimestamp` (MST/MDT standardized timestamps)

`Core.Runtime.psm1` exists and is imported by all scripts. Local `_ToHashtable` / `_Read-AppSettings` copies have been removed from most scripts.

### Priority 1b: database telemetry module — DONE
`modules/Database/Core.Database.psm1` created with SQL Server connectivity, schema management, and fire-and-forget telemetry writers. Wired into `Watch-QCTrigger.ps1` and `Run-QCProcessor.ps1`.

### Priority 1c: audit poller module — DONE
`modules/ProjectWise/PW.AuditPoller.psm1` created with audit-trail scanning, watermark management, and watch-root matching. Integrated into `Watch-QCTrigger.ps1` as the primary trigger source.

### Priority 2: PW discovery utilities
Extend `modules/ProjectWise/PW.Discovery.psm1` with:
- metadata helpers (`Get-PWDocName`, `Get-PWDocDescription`, `Get-PWDocLastModifiedUtc`)
- folder traversal helpers used by watcher discovery

### Priority 3: hashing utilities — DONE
`modules/Core/Core.Hashing.psm1` exists and is used by status-set processing.

### Priority 4: worker policy extraction — DONE
`modules/Queue/QC.Worker.psm1` exists with lock-retry and transition policy.

## What should remain in scripts
Keep scripts as thin entrypoints only:
- params / validation,
- module imports,
- orchestration call,
- process exit code.

## Proposed acceptance criteria for completion
1. No `_ToHashtable` defined in `scripts/*.ps1`. — **Mostly done** (some legacy scripts remain)
2. No script-local `_Read-AppSettings*` implementation. — **Mostly done**
3. Watcher-specific algorithmic helpers moved to modules. — **Partially done** (audit poller extracted; some PW discovery helpers remain in Watch-QCTrigger)
4. Worker retry/transition logic callable from module API. — **Done** (`QC.Worker.psm1`)
5. Existing script behavior unchanged (log codes and job-state transitions preserved). — **Verified**

## Current module inventory (30 modules)

**Core**: `Core.Config`, `Core.Database`, `Core.Hashing`, `Core.Logging`, `Core.Metrics`, `Core.Paths`, `Core.Results`, `Core.Runtime`

**QC**: `QC.Filters`, `QC.JobFactory`, `QC.Notifications`, `QC.NotificationGraph`, `QC.NotificationMock`, `QC.NotificationTemplates`, `QC.Processors`, `QC.Queue.Json`, `QC.Reporting`, `QC.StatusSet`, `QC.Triggers`, `QC.Worker`, `QC.Workflow`

**PW**: `PW.AuditPoller`, `PW.Connection`, `PW.Discovery`

**Orchestrator**: `Orchestrator.Pipeline`

## Bottom line
Most Priority 1 items are complete. Remaining work is consolidating PW discovery helpers (Priority 2) and cleaning up the few remaining legacy script-local helper copies.
