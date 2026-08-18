# Status Set AV File-Churn Notes

## Current file-operation pattern

The active native status-set processor keeps PowerShell responsible for ProjectWise enumeration, export, upload, queue orchestration, and logging. Local PDF processing is still performed by PowerShell calling `qpdf.exe`.

High-churn areas identified:

- ProjectWise exports write sheet PDFs into a per-job `_export_<job>` staging folder, then PowerShell renames each export to a sequence-prefixed filename to avoid same-name collisions.
- Fresh exports are copied into the persistent `_sheet_cache` so later cycles can reuse unchanged sheets.
- The merged status set used to be written directly to the persistent `_StatusSet.pdf` path.
- The processor aggressively cleaned every `_export_*` scratch folder after each rebuild, which caused rapid file deletes immediately after rapid file writes.
- Manifest writes use a temporary JSON file and replace the manifest after a throttle/backoff.

## Hash usage

Hashes are still used for stable identifiers, not for volatile output names:

- `Get-StatusSetWorkspaceDirectory` hashes the normalized ProjectWise/local sheets path to derive a deterministic workspace key and avoid collisions between folders.
- `_SSS-StatusSetSheetCachePath` hashes the sheet key to create collision-safe sheet-cache filenames.
- Job IDs and dedupe keys elsewhere may remain hashed, but status-set output remains the configured human-readable `_StatusSet.pdf`.

## Implemented low-risk mitigation

This patch follows the minimal/medium path: keep PowerShell as orchestrator and ProjectWise owner, but reduce rapid local mutation.

- Incremental mode is explicit and defaults to enabled.
- Sheet-cache reuse remains enabled so unchanged PW sheets are not re-exported when the manifest proves the cache is current.
- The merged PDF is rendered to staging first, then installed to the persistent `_StatusSet.pdf` path with atomic replacement enabled by default.
- Existing `_StatusSet.pdf` files are preserved in `_history` during replacement.
- Export scratch folders are no longer deleted at the start/end of every job by default.
- Cleanup is retention-based and disabled by default; when enabled, only expired staging/render files are removed.
- `_history` PDFs and manifest `.bak_*` copies are thinned during reconcile / scheduled full-scan (`statusSet.historyRetention`), not after every rebuild.
- Dry-run results include an operation report listing planned downloads, writes, replacements, deletes, and skips.

## Longer-term recommendation

A whitelisted compiled status-set EXE is still the better long-term architecture if AV continues to flag PowerShell behavior. Keep PowerShell for ProjectWise download/upload and queue orchestration, but move the local manifest diff, PDF page planning, merge, atomic install, history, and cleanup planning into the EXE. The EXE should emit a JSON result/operation report consumed by the PowerShell worker.
