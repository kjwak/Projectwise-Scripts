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

### Priority 1: runtime foundation module
Create `modules/Core.Runtime.psm1` with:
- `ConvertTo-HashtableDeep`
- `Read-QCAppSettings`
- `Write-QCJsonLog`

Then replace local copies in all scripts.

### Priority 2: PW discovery utilities
Extend `modules/PW.Discovery.psm1` with:
- metadata helpers (`Get-PWDocName`, `Get-PWDocDescription`, `Get-PWDocLastModifiedUtc`)
- folder traversal helpers used by watcher discovery

### Priority 3: hashing utilities
Create `modules/Core.Hashing.psm1` and move SHA helpers there.

### Priority 4: worker policy extraction
Add `modules/QC.Worker.psm1` (or extend queue module) for lock-retry + transition policy currently in `Run-QCProcessor.ps1`.

## What should remain in scripts
Keep scripts as thin entrypoints only:
- params / validation,
- module imports,
- orchestration call,
- process exit code.

## Proposed acceptance criteria for completion
1. No `_ToHashtable` defined in `scripts/*.ps1`.
2. No script-local `_Read-AppSettings*` implementation.
3. Watcher-specific algorithmic helpers moved to modules.
4. Worker retry/transition logic callable from module API.
5. Existing script behavior unchanged (log codes and job-state transitions preserved).

## Bottom line
Reassessment confirms the previous guidance is still correct on the latest branch; the key next step is execution: consolidate duplicated helpers into modules and keep scripts as thin entrypoints.
