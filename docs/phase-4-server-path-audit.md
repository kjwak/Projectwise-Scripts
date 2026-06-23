# Phase 4 — Server path audit

**Branch:** `phase-4/server-path-audit`  
**Base branch:** `phase-4/integration`  
**Status:** In-repo inventory complete — production server sign-off pending

## Purpose

Phase 4 moved scripts and modules behind compatibility wrappers. **External callers** (Task Scheduler, ProjectWise Rules Engine, operator shortcuts, worker publish layout, per-machine `appsettings.local.json`) may still reference old paths.

This audit documents:

1. Every **in-repo** path surface that production or operators might depend on.
2. What **`Publish-QCPipelineCode.ps1`** actually deploys to worker roots.
3. A **manual checklist** to run on each worker / automation host before `phase-4/service-scripts` or `phase-4/shim-removal` (4H).

No repo behavior changes. Production commands are not run from this branch.

## Gate criteria

| Next branch | Blocked until |
|-------------|---------------|
| `phase-4/service-scripts` | Publish copy list updated for any moved service paths; worker sign-off below |
| `phase-4/shim-removal` (4H) | 1-week production soak **and** all external callers confirmed on shim or new paths |
| `phase-4/publish-email-assets` | Independent — closes documented `email/` publish gap |

---

## 1. Worker install layout (expected)

Documented example worker root (from `Publish-QCPipelineCode.ps1` help):

```text
D:\QC_Pipeline\Prepend PDF QC\
  appsettings.json          # NOT copied by publish — must exist separately on worker
  appsettings.local.json    # optional per-machine overrides (gitignored)
  email\                    # NOT copied by publish — required for Graph HTML notifications
  dist\qc_overlay_prepend\  # NOT copied by publish — overlay exe for prepend/review stamp
  modules\                  # copied (full tree)
  scripts\
    Watch-QCTrigger.ps1
    Run-QCProcessor.ps1
    Import-QCScriptModules.ps1
    Start-QCPipelineDashboard.ps1
    Reset-QCFolderWorkflow.ps1
    … (wrappers + subfolders if full git clone)
  queue\                    # runtime data — not in publish
  legacy\                   # NOT in publish copy plan — needed for legacyPw / -Legacy paths
```

**Publish copies today** (`scripts/Publish-QCPipelineCode.ps1`):

| Source | Destination on worker |
|--------|------------------------|
| `modules/` (recursive) | `modules/` |
| `scripts/Watch-QCTrigger.ps1` | same relative path |
| `scripts/Run-QCProcessor.ps1` | same |
| `scripts/Import-QCScriptModules.ps1` | same |
| `scripts/Start-QCPipelineDashboard.ps1` | same |
| `scripts/Reset-QCFolderWorkflow.ps1` | same |

**Not copied by publish** (worker must already have these from git clone, manual copy, or local build):

| Path | Why it matters |
|------|----------------|
| `email/` | `notifications.email.templatePath` → `email/templates/qc_notification.html` |
| `appsettings.json` (+ local/secrets) | All runtime config |
| `dist/qc_overlay_prepend/` | Default `qcPrepend.overlayExePath` / review stamp |
| `legacy/` | `run_prepend_qc -Legacy`, `statusSet.mode=legacy`, deprecated prepend shim |
| `scripts/Stop-QCPipeline.ps1` | Used by publish `-ConfirmRestart` only |
| `scripts/Restore-QCModuleExports.ps1` | Service bootstrap — present on full clone, not published |
| Diagnostic/maintenance wrappers | Operator tools — not published |

---

## 2. Root compatibility shims (repo root → `scripts/`)

These **must remain** until 4H after external validation.

| Root shim | Forwards to |
|-----------|-------------|
| `Start-QCPipelineDashboard.ps1` | `scripts\Start-QCPipelineDashboard.ps1` |
| `Watch-QCTrigger.ps1` | `scripts\Watch-QCTrigger.ps1` |
| `Run-QCProcessor.ps1` | `scripts\Run-QCProcessor.ps1` |
| `run_prepend_qc.ps1` | `scripts\run_prepend_qc.ps1` |

**Risk:** Task Scheduler or operator shortcuts that use repo-root paths (e.g. `...\Prepend PDF QC\Start-QCPipelineDashboard.ps1`) depend on these shims.

---

## 3. Service internal spawn paths (hardcoded in repo)

These paths are **inside** production entrypoints — moving service scripts requires updating them **and** publish.

| Caller | Spawn target | Notes |
|--------|--------------|-------|
| `scripts/Start-QCPipelineDashboard.ps1` | `scripts\Watch-QCTrigger.ps1`, `scripts\Run-QCProcessor.ps1` | Uses `$repoRoot` + `scripts\` |
| `scripts/run_prepend_qc.ps1` (default) | `scripts\Watch-QCTrigger.ps1`, `scripts\Run-QCProcessor.ps1` | Uses `$PSScriptRoot` (scripts folder) |
| `scripts/run_prepend_qc.ps1` (default dashboard) | `scripts\Start-QCPipelineDashboard.ps1` | |
| `scripts/run_prepend_qc.ps1` (`-Legacy`) | `legacy\prepend_qc_on_trigger.ps1`, `legacy\watchlist.json` | True-legacy monolith path |
| `scripts/Publish-QCPipelineCode.ps1` (`-ConfirmRestart`) | `scripts\Stop-QCPipeline.ps1`, `scripts\Start-QCPipelineDashboard.ps1` | Stop matches process **command-line substrings**, not file paths |
| `scripts/Stop-QCPipeline.ps1` | (none — kills by name) | Matches `Start-QCPipelineDashboard`, `Watch-QCTrigger`, `Run-QCProcessor` in command line |

`Stop-QCPipeline.ps1` does **not** match file paths — only script name tokens in `Win32_Process.CommandLine`. Renaming scripts would break stop/restart until patterns are updated.

---

## 4. `scripts/` compatibility wrappers (old flat path → subfolder)

Validated by existing wrapper tests:

| Area | Wrapper count | Test |
|------|-----------------|------|
| Diagnostics | 23 | `test/test_diagnostic_script_wrappers.ps1` |
| Maintenance | 16 | `test/test_maintenance_script_wrappers.ps1` |
| Processing / deployment | 4 | `test/test_processing_deployment_script_wrappers.ps1` |

**Cross-wrapper call:** `maintenance/Invoke-QCDatabaseRetention.ps1` spawns `scripts\maintenance\Remove-QCAuditEvents.ps1` via **wrapper path** `Join-Path $PSScriptRoot 'Remove-QCAuditEvents.ps1'` (relative to `maintenance/` — resolves to implementation, not `scripts/Remove-QCAuditEvents.ps1`). External callers using `scripts\Remove-QCAuditEvents.ps1` still hit the flat wrapper.

---

## 5. Module flat shims (`modules/*.psm1` → `modules/<Area>/`)

41 flat shims forward to folder implementations (Phase 4E). Production entrypoints import **folder paths** directly; flat shims exist for backward compatibility.

**Risk:** Any external script or old worker code doing `Import-Module modules\QC.Queue.Json.psm1` still works via shim. Removing shims (4H) breaks those callers.

Test: `test/test_module_folder_shims.ps1`

---

## 6. Processor / worker dispatch paths (runtime spawns)

From `modules/Processing/QC.Processors.psm1`:

| Job / mode | Default script / asset | Override key |
|------------|------------------------|--------------|
| `QC_PREPEND` (`legacyPw` / `projectWise` / `pw`) | `scripts\processing\Invoke-QCPrependPw.ps1` | `qcPrepend.projectWiseScriptPath`, `qcPrepend.legacyScriptPath` |
| `legacy/prepend_qc.ps1` shim | Forwards to `Invoke-QCPrependPw.ps1` with deprecation warning | Used if override points at legacy path |
| `STATUS_SET_GEN` (`statusSet.mode=legacy`) | `legacy\combine_status_set.ps1` | `statusSet.legacyScriptPath` |
| Overlay merge | `dist\qc_overlay_prepend\qc_overlay_prepend.exe` | `qcPrepend.overlayExePath` |
| Email HTML | `email/templates/qc_notification.html` (repo-relative) | `notifications.email.templatePath` |

Production uses `statusSet.mode = native` — legacy combine path is fallback only.

---

## 7. MCP / developer tooling (not worker pipeline)

`.cursor/mcp.json` pins **machine-specific** absolute paths:

- Python venv + `tools/pw-qc-mcp/server.py`
- `PWQC_REPO_ROOT`, `PWQC_APPSETTINGS`, SQL env vars

`server.py` defaults `PWQC_REPO_ROOT` to repo root (parent of `tools/`). MCP is **out of scope** for worker publish; do not change during service moves unless coordinating with developers.

---

## 8. In-repo finding: no scheduled tasks or Rules Engine definitions

Searched the repository for `Register-ScheduledTask`, `schtasks`, and Rules Engine callback registrations — **none found**. Task Scheduler and ProjectWise Rules Engine configuration lives **outside** this repo.

---

## 9. Manual production checklist (per worker / automation host)

Run on each machine that runs QC pipeline or invokes QC scripts. Record results in the sign-off table (Section 10).

### 9.1 Running processes (reveals actual script paths in use)

```powershell
$patterns = 'Start-QCPipelineDashboard|Watch-QCTrigger|Run-QCProcessor'
Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" |
  Where-Object { $_.CommandLine -match $patterns } |
  Select-Object ProcessId, @{n='Cmd';e={
    if ($_.CommandLine.Length -gt 200) { $_.CommandLine.Substring(0,200) + '...' } else { $_.CommandLine }
  }}
```

Note whether paths use **repo root** shims vs `scripts\` directly.

### 9.2 Task Scheduler

```powershell
Get-ScheduledTask |
  Where-Object { $_.Actions.Execute -match 'powershell|pwsh' } |
  ForEach-Object {
    $a = $_.Actions
    [PSCustomObject]@{
      Task = $_.TaskName
      Execute = $a.Execute
      Arguments = $a.Arguments
    }
  } |
  Where-Object { $_.Arguments -match 'QC|Prepend|Watch-QCTrigger|Run-QCProcessor|Start-QCPipeline' }
```

Also review Task Scheduler GUI on hosts where QC was set up manually.

### 9.3 Per-machine config overrides

On the worker root, inspect merged settings (paths may override defaults):

```powershell
$root = 'D:\QC_Pipeline\Prepend PDF QC'   # adjust per site
@(
  'qcPrepend.legacyScriptPath'
  'qcPrepend.projectWiseScriptPath'
  'qcPrepend.overlayExePath'
  'statusSet.legacyScriptPath'
  'notifications.email.templatePath'
) | ForEach-Object { "Check appsettings for: $_" }

Get-ChildItem -LiteralPath $root -Filter 'appsettings*.json' | Select-Object Name, LastWriteTime
```

### 9.4 ProjectWise Rules Engine

In ProjectWise Administrator / Rules Engine UI for each datasource:

- Search rules for `prepend_qc`, `combine_status_set`, `Watch-QCTrigger`, `Run-QCProcessor`, `Start-QCPipeline`, or absolute paths under the worker root.
- Document rule name, trigger, and full script path if any rule invokes PowerShell.

### 9.5 Publish vs clone drift

On worker:

```powershell
$root = 'D:\QC_Pipeline\Prepend PDF QC'
@(
  'email\templates\qc_notification.html'
  'scripts\Restore-QCModuleExports.ps1'
  'scripts\Stop-QCPipeline.ps1'
  'legacy\prepend_qc.ps1'
  'dist\qc_overlay_prepend\qc_overlay_prepend.exe'
) | ForEach-Object {
  $p = Join-Path $root $_
  [PSCustomObject]@{ Path = $_; Exists = (Test-Path -LiteralPath $p) }
} | Format-Table -AutoSize
```

### 9.6 Operator docs and shortcuts

Search shared runbooks / desktop shortcuts for:

- `legacy\prepend_qc.ps1`
- Root-level `Start-QCPipelineDashboard.ps1`
- Old paths before Phase 4C/4D moves (`scripts\Show-QCStatus.ps1` still valid via wrapper)

---

## 10. Production sign-off (fill in manually)

| Site / host | Worker root path | Task Scheduler paths found | Rules Engine rules found | `appsettings.local` path overrides | Running process paths | Reviewer | Date |
|-------------|------------------|----------------------------|--------------------------|-----------------------------------|----------------------|----------|------|
| *(example)* | `D:\QC_Pipeline\Prepend PDF QC` | | | | | | |

**4H shim removal** requires every row populated with “no unexpected hardcoded paths” or documented migration plan.

---

## 11. Static in-repo validation

```powershell
./test/test_server_path_audit.ps1
./test/test_diagnostic_script_wrappers.ps1
./test/test_maintenance_script_wrappers.ps1
./test/test_processing_deployment_script_wrappers.ps1
./test/test_notification_template_paths.ps1
```

---

## 12. Recommended next steps

| Priority | Branch | Rationale |
|----------|--------|-----------|
| 1 | Complete Section 10 on production hosts | Only gap blocking 4H / service moves |
| 2 | `phase-4/publish-email-assets` | Close `email/` publish gap (independent, low risk) |
| 3 | `phase-4/service-scripts` | Move service entrypoints to `scripts/service/` after sign-off + publish update |
| 4 | `phase-4/shim-removal` | After soak + sign-off |

---

## Related documents

- [`docs/phase-4-module-script-organization-plan.md`](phase-4-module-script-organization-plan.md) — Section 8 external checklist, Appendix L
- [`docs/dead-code-and-deprecation-audit.md`](dead-code-and-deprecation-audit.md) — Rules Engine open question
- [`scripts/Publish-QCPipelineCode.ps1`](../scripts/Publish-QCPipelineCode.ps1)
