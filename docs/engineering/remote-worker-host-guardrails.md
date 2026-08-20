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

Restrict this host with gitignored `workers.enabledJobTypes` (modelling PC: `["QC_PREPEND"]` so `STATUS_SET_GEN` stays on the QC server). Empty/omitted `enabledJobTypes` means all types — do not set that in committed `appsettings.json`.

## Config

- **Do not** change committed `appsettings.json` defaults to enable remote workers, UNC queue, or extra PW writes.
- Per-machine paths and SQL overrides go in gitignored `appsettings.local.json`.
- Copy [`appsettings.remote-worker.example.json`](../../appsettings.remote-worker.example.json) as a starting overlay. Never commit `pw_cred.txt` or connection secrets.

`Read-QCAppSettings` already merges `appsettings.json` → `appsettings.local.json` → `appsettings.secrets.json`.

## Lock safety

Queue lock files must include `machineName` plus `pid`. The same fields are stamped onto the job JSON at claim (`running/`) and kept when the file moves to `succeeded/` or `failed/`. Recover-to-pending clears them. SQL `processing_jobs` mirrors `worker_machine_name` / `worker_pid` (schema 1.22.0). Local PID-only liveness will steal (or fail to recover) jobs across machines if the QC server is not running the same host-aware lock code.

See `modules/Queue/QC.Queue.Json.psm1` (`_QCQJ-IsLockOwnerDead`, `Recover-QCStaleJobs`). Job `machineName` is attribution only — lock files remain the ownership primitive.

## Accounts

Do not run QC processors under the shared interactive D-Team login. Use a dedicated scheduled-task account (later).
