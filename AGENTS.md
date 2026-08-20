# Agent instructions — ProjectWise QC Pipeline

Guidance for AI agents (Cursor, Codex, etc.) working in this repository.

## Production constraints

- **Windows + PowerShell** is the production runtime. Service entrypoints live under `scripts/service/`.
- **Do not assume Python is installed** on production worker hosts. Python is used for overlay **build** (`tools/overlay/`, PyInstaller → `dist/qc_overlay_prepend/`) and **developer tests** (`test/python/`). Production `QC_PREPEND` invokes the built overlay executable — not `python` on the worker.
- **ProjectWise** (`pwps`, `pwps_dab`) is required for live PW operations. Cloud sandboxes cannot validate PW connectivity without the Bentley stack and TYPSA network.
- **Secrets** live in gitignored files (`appsettings.secrets.json`, `pw_cred.txt`, `appsettings.local.json`) — never commit them.

## Canonical service entrypoints

| Script | Purpose |
|--------|---------|
| `scripts/service/Watch-QCTrigger.ps1` | Watcher: audit/reconcile, filters, triggers, enqueue |
| `scripts/service/Run-QCProcessor.ps1` | Worker: dequeue, lock, dispatch processor, transition job state |
| `scripts/service/Start-QCPipelineDashboard.ps1` | Dashboard: watcher + worker pool + terminal UI |

Pass `-AppSettingsPath` to override config (e.g. `.\appsettings.test.json` for safe local runs).

There are **no** watcher/worker scripts at the repository root.

## Domain model (current production)

### Lane QC PDFs

Authoritative lane documents per sheet stem:

- `*-prod.pdf` — Production (`QC_Process_Type` = Production)
- `*-rev.pdf` — Review
- `*-chk.pdf` — Check

Legacy `*-qc.pdf` is a **compatibility bridge** only (older prepend paths, normalization helpers, historical `sheet_index` rows). **Do not** document or configure new work around `*-qc.pdf` as the primary lane trigger.

### Execution vs telemetry

| Layer | Role | Source of truth |
|-------|------|-----------------|
| **JSON queue** | Job scheduling and worker dispatch | `queue.rootDir` → `pending/`, `running/`, `succeeded/`, `failed/`, `locks/` (`modules/Queue/QC.Queue.Json.psm1`) |
| **SQL (`QC_Pipeline`)** | Audit ingest, package index, mirrors, reporting | Durable telemetry — **not** a second job scheduler |

**Invariants:**

- Workers read jobs **only** from the JSON queue.
- `processing_jobs` in SQL mirrors queue outcomes for dashboards; it does not drive execution.
- `sheet_packages` / `sheet_package_qc_pdfs` are canonical for lane/document relationships, not for job state.

See `docs/data/database-telemetry.md` and `docs/architecture/qc-package-model.md`.

### Audit watermark

When `database.enabled` is true:

1. **Primary:** `watcher_state` table (`audit_watermark_utc` / related keys) in SQL — see `modules/Database/Core.Database.psm1`, `modules/ProjectWise/PW.AuditPoller.psm1`.
2. **Mirror / fallback:** `queue/_watcher/audit-capture-watermark.txt` and `poll_runs.watermark_after`.

Steady-state discovery is **audit-driven** (`dms_audt` → `audit_events`). Full folder scans are reconciliation-only (scheduled or hybrid downtime). See `docs/architecture/hybrid-polling.md`.

### Prepend mode

Committed production uses `qcPrepend.mode: "projectWise"`. `QC_PREPEND` jobs are orchestrated by `Invoke-QCPrependProcessor` in `modules/Processing/QC.Processors.psm1`, which spawns `scripts/processing/Invoke-QCPrependPw.ps1` (PW export, overlay merge, lane PDF upload). `legacy/prepend_qc.ps1` is a deprecated compatibility forwarder only.

`qcPrepend.mode: "local"` runs a disk-only path inside `QC.Processors.psm1` (no PW file I/O). It is for tests and staging — not production. `legacyPw` and `pw` are accepted aliases for `projectWise`. See `docs/engineering/phase-2-native-prepend-parity-plan.md` for future in-process PW parity work.

### Status-set local history

Each workspace under `statusSet.localRoot` keeps rollback copies in `_history` (`_StatusSet_yyyyMMdd_HHmmss.pdf`) and manifest `.bak_*` files. These are **not** a second scheduler. Thinning runs during `Invoke-StatusSetReconcile` and once per scheduled full-scan slot (`statusSet.historyRetention`: recent copies + daily/weekly calendar tiers + `maxGbPerFolder`). Do not delete `_history` from `STATUS_SET_GEN` itself (AV file-churn). Live `_StatusSet.pdf` and the current manifest are never deleted.

## Documentation map

Start at `docs/README.md`. Key paths:

| Topic | Path |
|-------|------|
| Config reference | `docs/reference/appsettings-reference.md` |
| Hybrid polling / watcher | `docs/architecture/hybrid-polling.md` |
| SQL schema | `docs/data/database-telemetry.md` |
| QC workflow writeback | `docs/workflow/qc-workflow-framework.md` |
| Notifications | `docs/workflow/qc-notifications.md` |
| Testing profile | `docs/reference/testing-config.md` |
| Dry-run policy | `docs/reference/dry-run-policy.md` |
| Module conventions | `docs/reference/module-contracts.md` |
| Phase migration history | `docs/archive/phase/` (historical — see below) |

### Archive documents (`docs/archive/`)

Files under `docs/archive/` are **historical references**. They may describe older paths, workflow states, or architecture and **must not** be treated as current implementation guidance.

- Prefer `reference/`, `architecture/`, `workflow/`, `data/`, and `engineering/` for how the system works today.
- Consult `docs/archive/` only for migration context, decision history, or audit trail of refactors.
- If an archive doc disagrees with code or a non-archive doc, **trust code and the current doc**.
- Treat an archive document as authoritative **only** when another **current** document explicitly links to it for that purpose.

## When you change behavior — update docs

If your change affects any of the following, update the **relevant** docs in the same PR (or follow-up immediately):

- Runtime behavior, triggers, or processor logic
- `appsettings.json` keys, defaults, or merge chain
- File paths, queue layout, or service entrypoints
- Workflow states, lane PDF rules, or notification routing
- SQL schema, telemetry tables, or watermark persistence
- Operational assumptions (dry-run flags, reconciliation schedule, fallback behavior)

Also update `AGENTS.md` if agent-facing assumptions change.

## Laptop queue/log share (read-only diagnostics)

Live execution queue and service logs are exposed read-only at `\\192.168.22.90\QC_Queue` (logs under `_logs`). This is **not** production `queue.rootDir` (`C:\QC_E2E_RealRun\queue` on the worker) and must not be used as a write target.

- Resolver: `scripts/diagnostics/Resolve-QCDebugLocations.ps1` — precedence: `-QueueRoot` → `$env:QC_DEBUG_QUEUE_ROOT` → `config/qc-debug-locations.local.json` → default UNC.
- Snapshot: `scripts/diagnostics/Show-QCQueueDiag.ps1` (optional `-NoLogs`).
- Copy `config/qc-debug-locations.example.json` → `config/qc-debug-locations.local.json` (gitignored) for laptop overrides.

## Verification honesty

- **Do not claim** end-to-end verification of ProjectWise writeback, SQL telemetry, Graph email, or prepend output unless you ran it on a Windows worker with live PW/SQL/Graph access.
- Code-only and static-test validation is sufficient for many PRs — state what was and was not validated.
- For PW workflow debugging on a connected machine, prefer MCP `pw-qc-debug` tools before deep code speculation (see `.cursor/rules/projectwise-debugging.mdc`).

## Remote worker host (modelling PC)

This clone may become an optional **processor host**. It must not run the watcher or the full dashboard. Per-machine queue UNC, SQL, and temp roots go in gitignored `appsettings.local.json` (start from `appsettings.remote-worker.example.json`). Do **not** change committed `appsettings.json` defaults to enable remote workers.

Do not claim jobs from the live UNC queue until **host-aware locks** (`machineName` + PID) are deployed on the QC server. See [`docs/engineering/remote-worker-host-guardrails.md`](docs/engineering/remote-worker-host-guardrails.md).

## Safe change boundaries (unless explicitly requested)

- Do not remove legacy compatibility code or flat module shims without a tracked deprecation plan.
- Do not change production `appsettings.json` defaults to enable PW writes, notifications, native prepend, or remote workers without operator sign-off.
- Do not treat SQL as a job queue replacement without an explicit design change.
- Do not run `Watch-QCTrigger.ps1` or `Start-QCPipelineDashboard.ps1` on the modelling PC as a second coordinator.
