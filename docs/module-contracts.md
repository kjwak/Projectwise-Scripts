# Module Contracts and Conventions

## Standard Result Object Schema
All public module functions should return a result object with this shape:

- `IsSuccess` `[bool]`: Indicates success/failure of the function call.
- `Code` `[string]`: Machine-readable outcome code (success, validation, or error category).
- `Message` `[string]`: Human-readable summary message.
- `Data` `[object]`: Optional payload with function-specific details.

Helper constructors are defined in `modules/Core.Results.psm1`:
- `New-QCResult`
- `New-QCSuccessResult`
- `New-QCFailureResult`

## Error Code Naming Convention
Recommended pattern:

`<AREA>_<CATEGORY>_<DETAIL>`

Examples:
- `CONFIG_VALIDATION_MISSING_KEY`
- `QUEUE_DUPLICATE_DETECTED`
- `PW_SESSION_UNAVAILABLE`
- `TRIGGER_NO_MATCH`

Guidelines:
- Use uppercase snake case.
- Keep stable codes for dashboarding and alert rules.
- Prefer specific, deterministic codes over generic ones.

## Public/Private Function Naming
- Public function naming: approved Verb-Noun cmdlet style.
- Internal/private helper naming recommendation:
  - Prefix private helpers with underscore (for example: `_Resolve-InternalRule`).
  - Avoid exporting private helpers.

## Module Export Policy Recommendation
Current scaffolding exports all functions using:

`Export-ModuleMember -Function *`

Recommendation for implementation phase:
- Keep broad export during rapid prototyping.
- Move to curated exports once public/private split is clear.
- Add module-level tests to enforce export boundaries.

## ProjectWise Safety Rule
- No ProjectWise write operations are allowed without explicit approval.
- Read-only discovery/session checks are permitted.
- Any future write-capable function must be reviewed separately and clearly documented.

## Database Telemetry Convention
- All database writes use the fire-and-forget pattern via `_QDB-SafeWrite` in `Core.Database.psm1`.
- Pipeline execution must never fail because telemetry fails.
- Database availability is checked via `Test-QCDatabaseEnabled` before attempting writes.
- Schema initialization is idempotent and version-tracked via `schema_version` table.
- See `docs/database-telemetry.md` for the full schema and API reference.

## Timestamp Convention
- All timestamps are stored internally as UTC (`DATETIMEOFFSET` in SQL Server, ISO 8601 in JSON).
- Display timestamps are standardized to Mountain Time (MST/MDT) via `Get-QCTimestamp` from `Core.Runtime.psm1`.
- Modules should use `Get-QCTimestamp` for any user-facing timestamp, not `Get-Date` directly.

## Module Inventory

| Module | Responsibility |
|--------|---------------|
| `Core.Config` | Configuration loading and validation |
| `Core.Database` | SQL Server connectivity, schema, telemetry writers |
| `Core.Hashing` | SHA-256 hashing utilities |
| `Core.Logging` | Structured logging |
| `Core.Metrics` | Performance metrics collection |
| `Core.Paths` | Path normalization (PW URIs, local paths) |
| `Core.Results` | Standard result object constructors |
| `Core.Runtime` | Shared runtime utilities (`Read-QCAppSettings`, `Get-QCTimestamp`, `Write-QCJsonLog`) |
| `PW.AuditPoller` | Audit-trail scanning, watermark, watch-root matching |
| `PW.Connection` | ProjectWise session management |
| `PW.Discovery` | PW folder/document discovery utilities |
| `QC.Filters` | Path and document filtering |
| `QC.JobFactory` | Job object creation and dedupe keys |
| `QC.Notifications` | Notification orchestration and dedupe |
| `QC.NotificationGraph` | Microsoft Graph stub (future) |
| `QC.NotificationMock` | Mock notification provider |
| `QC.NotificationTemplates` | Email subject/body templates |
| `QC.Processors` | Job processor dispatch |
| `QC.Queue.Json` | JSON file-based queue (atomic, lock-aware) |
| `QC.Reporting` | QC reporting snapshots |
| `QC.StatusSet` | Status set pairing, merge, writeback |
| `QC.Triggers` | Trigger rule evaluation |
| `QC.Worker` | Worker lifecycle, lock-retry, transitions |
| `QC.Workflow` | QC workflow state/attribute writeback |
| `Orchestrator.Pipeline` | Pipeline orchestration (dashboard integration) |
