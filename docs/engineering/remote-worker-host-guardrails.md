# Remote worker host — guardrails

**Status:** In force for the modelling-PC clone  
**Related:** [`remote-worker-architecture-intent.md`](remote-worker-architecture-intent.md)

This clone on the drainage modelling workstation (AZTEC002799) is a **processor host**, not a second QC coordinator.

## What may run here

| Allowed | Not allowed |
|---------|-------------|
| `Run-QCProcessor.ps1` (after host-aware locks are live on the **QC server**) | `Watch-QCTrigger.ps1` |
| Future `Start-QCRemoteWorkerHost.ps1` supervisor | `Start-QCPipelineDashboard.ps1` (full pipeline) |
| Local `qcPrepend` temp/output on fast disk (E: / F:) | Audit polling, enqueue, reconciliation sweeps |
| Read-only diagnostics (`Show-QCQueueDiag.ps1`) | Treating `\\192.168.22.90\QC_Queue` as a write target until slice 1 is deployed on the server |

## Config

- **Do not** change committed `appsettings.json` defaults to enable remote workers, UNC queue, or extra PW writes.
- Per-machine paths and SQL overrides go in gitignored `appsettings.local.json`.
- Copy [`appsettings.remote-worker.example.json`](../../appsettings.remote-worker.example.json) as a starting overlay. Never commit `pw_cred.txt` or connection secrets.

`Read-QCAppSettings` already merges `appsettings.json` → `appsettings.local.json` → `appsettings.secrets.json`.

## Lock safety

Queue lock files must include `machineName` plus `pid`. Do **not** point this host at the live UNC queue until the QC server is running the same host-aware lock code. Local PID-only liveness will steal (or fail to recover) jobs across machines.

See `modules/Queue/QC.Queue.Json.psm1` (`_QCQJ-IsLockOwnerDead`, `Recover-QCStaleJobs`).

## Accounts

Do not run QC processors under the shared interactive D-Team login. Use a dedicated scheduled-task account (later).
