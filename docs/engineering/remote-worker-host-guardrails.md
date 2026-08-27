# Remote worker host — guardrails

**Status:** In force for the modelling-PC clone  
**Related:** [`remote-worker-architecture-intent.md`](remote-worker-architecture-intent.md)

This clone on the drainage modelling workstation (AZTEC002799) is a **processor host**, not a second QC coordinator.

## What may run here

| Allowed | Not allowed |
|---------|-------------|
| `Run-QCProcessor.ps1` (after host-aware locks are live on the **QC server**) | `Watch-QCTrigger.ps1` |
| `Start-QCRemoteWorkerHost.ps1` (processor-only supervisor) | `Start-QCPipelineDashboard.ps1` (full pipeline) |
| Local `qcPrepend` temp/output on fast disk (E: / F:) | Audit polling, enqueue, reconciliation sweeps |
| Read-only diagnostics (`Show-QCQueueDiag.ps1`) | UNC claims without `-AllowUncQueue` / `workers.remoteHost.allowUncQueue` |

The supervisor and logon console (`Watch-QCRemoteWorkerHostConsole.ps1`) take **in-flight work from `queue\running`** for this host (`machineName`). Heartbeat `idle` / `busy=N <document>` is that snapshot, not JSONL. `QUEUE_RUNNING` prints when a new running file appears and JSONL never logged `WORKER_SELECTED`. JSONL is still tailed for stage/success/failure (`RW*` by default; `-ShowAllWorkers` / `-AllWorkers` includes QC-server jobs). Pass `-AllWorkers` on the logon console to include every host in `running\`.

Restrict this host with gitignored `workers.enabledJobTypes` (modelling PC: `["QC_PREPEND"]` so `STATUS_SET_GEN` stays on the QC server). Empty/omitted `enabledJobTypes` means all types — do not set that in committed `appsettings.json`.

## Stamp assets (QC_PREPEND)

Lane PDF stamps are **not** optional for check/review prepend on this host. `QCProcess.StampAssets` paths are relative to the repo root unless `QCProcess.stampsRoot` is set in `appsettings.local.json`.

1. Copy the `stamps\` folder from the QC server install root into `<repo>\stamps\`, **or**
2. Copy stamp PDFs to e.g. `C:\PW_QC_LOCAL\stamps\` and set `"QCProcess": { "stampsRoot": "C:\\PW_QC_LOCAL\\stamps" }` in `appsettings.local.json`.

`Start-QCRemoteWorkerHost.ps1` warns at startup when configured stamp files are missing. Production lane prepend continues without a stamp when the template is absent; check/review prepend fails fast once stamp resolution fails.

## Stuck job in `running\`

A prepend job stays in `running\` until the worker finishes **post-prepend workflow writeback** (lane state change + `QC_NOTIFICATION` enqueue) and `Run-QCProcessor.ps1` moves it to `succeeded\`. The notification job is created during writeback, not after the move — but writeback must complete first.

The worker waits **only for the prepend PowerShell child PID**. Bentley Connection Client / other descendants that inherit redirected stdout/stderr must not keep the job in `running\`. Progress lines: `QC_PREPEND_PW_CHILD_START`, `QC_PREPEND_PW_CHILD_PID`, `QC_PREPEND_PW_CHILD_DONE`, `QC_PREPEND_PW_RECONNECT_START`, `QC_PREPEND_WORKFLOW_WRITEBACK_*`.

Durable queue checkpoints (`job.checkpoint`, mirrored to `processing_jobs.checkpoint`):

| Checkpoint | Meaning | Recovery |
|---|---|---|
| *(absent)* | Prepend child has not succeeded | Retry may rerun prepend |
| `prepend_complete` / `writeback_running` | Lane PDF created; writeback not finished | Resume **writeback only** — never rerun prepend |
| `writeback_complete` | Writeback finished | Skip prepend and writeback; archive the job |

**Unstick a hung job:**

1. If `checkpoint` is `prepend_complete` or `writeback_running`, stop the hung worker and let stale recovery requeue it. The next worker resumes writeback only.
2. If there is no checkpoint, stopping the worker may rerun prepend — do not requeue blindly after the lane PDF already exists. Persist `prepend_complete` first, or invoke writeback only.
3. On the modelling PC, stale recovery detects a dead local PID within ~30s (or wait for `queue.recover.staleSeconds`, default 900s, without a heartbeat).

Do not manually move the job to `succeeded\` unless prepend **and** workflow writeback already completed in ProjectWise — otherwise notification will not fire and lane state may be wrong.

## Config

- **Do not** change committed `appsettings.json` defaults to enable remote workers, UNC queue, or extra PW writes.
- Per-machine paths and SQL overrides go in gitignored `appsettings.local.json`.
- Copy [`appsettings.remote-worker.example.json`](../../appsettings.remote-worker.example.json) as a starting overlay. Never commit `pw_cred.txt` or connection secrets.

`Read-QCAppSettings` already merges `appsettings.json` → `appsettings.local.json` → `appsettings.secrets.json`.

## Host resource throttle

When `workers.remoteHost.throttle.enabled` is true (modelling PC `appsettings.local.json` only), the supervisor samples **machine-wide** CPU, physical memory, and matched modelling process trees across all sessions. Results go to `queue\_remote_worker.{host}.throttle.json` (includes `matchedProcessCpuPercent` / `matchedProcessMemoryMb`). Each write deletes and recreates that file so UNC share ACLs stay inherited. If the file becomes unreadable (ops throttle frozen), delete it on the QC-server queue root; the next sample recreates it.

- **Busy:** `recommendedSlots=busyRecommendedSlots` (default **1**). Supervisor stops spawning above that count. Extra children stay alive but skip `Get-NextQCJob` (lowest `RW*` labels keep claiming). `pauseNewClaims` is true only when `busyRecommendedSlots` is **0**. After a worker lease-exits, the next spawn reuses the lowest free `RW*` label (do not keep incrementing to `RW102`).
- **In-flight work always finishes** (prepend, reconnect, writeback, notification).
- **Fail-open:** missing, malformed, unreadable, stale (`sampledAtUtc` older than `max(3 * sampleSeconds, 30s)`), `reason=sample_error`, or the **first** process-CPU sample (no prior snapshot to compare) does **not** pause claims.
- Process-name patterns identify CAD/hydro apps (case-insensitive wildcards vs `Win32_Process.Name`, `.exe` optional). When `processCpuPercent` is set (typical **15**), a matching app sitting idle (HEC-RAS open on `dteam` with no model running) is **not** busy — the matched process **plus descendants** (solver children) must use CPU over `sampleSeconds`. `processMemoryMb` is optional and usually left at `0`. When both process gates are `0`, a name match alone still counts as busy. QC supervisor/worker PIDs and their descendant tree are excluded; other PowerShell processes are not.
- Leave machine `cpuPercent` / `memoryPercent` at `0` on this host so QC prepend does not throttle itself.
- Committed `appsettings.json` must keep this **off**. Do not add SQL `qc_workers` here — this is local host control only.

If you use Process Lasso (or similar) on this PC, exclude the QC remote-host supervisor and `Run-QCProcessor` PowerShell processes there. That is an ops setting, not QC code.

## Lock safety

Queue lock files must include `machineName` plus `pid`. The same fields are stamped onto the job JSON at claim (`running/`) and kept when the file moves to `succeeded/` or `failed/`. Recover-to-pending clears them. SQL `processing_jobs` mirrors `worker_machine_name` / `worker_pid` (schema 1.22.0). Local PID-only liveness will steal (or fail to recover) jobs across machines if the QC server is not running the same host-aware lock code.

See `modules/Queue/QC.Queue.Json.psm1` (`_QCQJ-IsLockOwnerDead`, `Recover-QCStaleJobs`). Job `machineName` is attribution only — lock files remain the ownership primitive.

## Accounts

Do not run QC processors under the shared interactive D-Team login. Use a dedicated scheduled-task account when the PC is shared; a personal account (e.g. `jflint`) is acceptable for a single-user modelling PC.

**Boot without login:** register a scheduled task (elevated):

```powershell
.\scripts\deployment\Register-QCRemoteWorkerHostTask.ps1 -AllowUncQueue
```

Runs `Start-QCRemoteWorkerHost.ps1` at startup with a short delay, **whether the user is logged on or not**. You will be prompted for the run-as password (stored with the task). Unregister with `-Unregister`.

The QC server (`PXBENTLEY01`) uses a different registrar — `Register-QCPipelineDashboardTask.ps1` — do not run that here.

**Log console at logon:** read-only tail in a PowerShell window (does not start a second supervisor):

```powershell
.\scripts\deployment\Register-QCRemoteWorkerHostLogonConsole.ps1
```

## ProjectWise modules on this host

`Open-PWConnection` comes from **pwps**. Folder/document cmdlets come from **pwps_dab**. The QC server usually has both on `PSModulePath`, so `Import-Module pwps_dab` is enough there. Explorer workstations often have `pwps_dab` in `WindowsPowerShell\Modules` while `pwps` is a separate module (or only under the Bentley install path). Workers spawn with `-NoProfile`, so a user profile that auto-imports `pwps` never runs.

`Test-PWConnection.ps1` already loads both via `PW.Connection` (`Import-PWCmdletModules`). Production prepend (`Invoke-QCPrependPw.ps1`) must use that same loader — do not import `pwps_dab` alone. Imports from inside `PW.Connection` use `-Global` so `Get-PWDocumentsBySearch` is visible in the prepend script after connect.

The modelling PC does **not** need to clone the QC server. It needs the same **cmdlet surface in the prepend child**: MTA, `pwps` (`Open-PWConnection`), and `pwps_dab` (`Get-PWDocumentsBySearch`, export/update). Watcher, dashboard, and server hardware stay on the coordinator.
