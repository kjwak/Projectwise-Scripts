# Dashboard Upgrade Options (Reassessed)

Date: 2026-05-06
Branch baseline: `7d260c2`

## Reassessment summary
On the latest branch, `scripts/Start-QCPipelineDashboard.ps1` is still a PowerShell-host rendered dashboard. The best path remains:
1. short-term: PowerShell ANSI diff rendering,
2. long-term: backend/frontend separation with a dedicated TUI.

## Current-state constraints observed
- Dashboard rendering is still terminal-host dependent, so flicker varies by host (ConsoleHost vs Windows Terminal).
- Script contains substantial formatting/render helpers, which makes feature evolution harder without a dedicated render abstraction.

## Updated option ranking

### 1) **Recommended now**: ANSI diff renderer in existing script
Implement a dual mode:
- `FullRedraw` (current fallback)
- `DiffAnsi` (new default where VT is supported)

Add:
- frame builder (`Get-FrameLines`),
- previous-frame cache,
- row-addressed updates only on changed lines,
- optional `-MaxFps` throttling.

Why first:
- No architecture split required.
- Lowest-risk path to perceptibly smoother UX.

### 2) **Recommended target**: PowerShell engine + dedicated TUI app
Keep orchestration in PowerShell; expose normalized event stream for a separate UI process.

Why second:
- Best no-flicker behavior and richer controls.
- Clean separation of concerns.

### 3) C# host replacement
Valid if team prefers .NET and wants native terminal performance, but higher replacement cost than option 1.

### 4) Web dashboard
Best for remote/multi-operator observability, but heavier footprint than needed for local terminal UX only.

## Concrete implementation checklist (Phase 1)
1. [x] Add params to `Start-QCPipelineDashboard.ps1`:
   - `-RenderMode FullRedraw|DiffAnsi`
   - `-MaxFps` (default 8)
2. [x] Refactor rendering to return full logical frame lines (no direct writes inside formatters).
3. [x] Add ANSI cursor + hide/show cursor wrappers.
4. [x] Diff previous/current frame and repaint changed rows only.
5. [x] Auto-fallback to `FullRedraw` when VT unsupported.

## Phase 1 implementation notes
- `DiffAnsi` is now the default render mode; use `-RenderMode FullRedraw` to force the legacy clear-and-redraw behavior.
- `-MaxFps` defaults to 8 and throttles dashboard paints independently of watcher/worker polling.
- The first ANSI render initializes the screen, then subsequent frames repaint only changed rows and clear those rows with ANSI line-clear sequences.
- Hosts that do not advertise VT/ANSI support are automatically downgraded to `FullRedraw`.

## Phase 1 dashboard layout follow-up
- Dashboard content is now grouped into bordered sections for at-a-glance scanning: summary, watcher/ProjectWise scan, workers, recent jobs, processor activity, and warnings/errors.
- Status-oriented rows are colorized so failures/warnings, active work, completed work, and neutral frame chrome are visually distinct.
- Workers now maintain a per-slot stage message from `WORKER_STAGE` events, making each worker show whether it is polling, locking, running a processor, moving results, or exiting.

## Success criteria
- No `Clear-Host` per frame in `DiffAnsi` mode.
- Stable layout with changed-line repaint only.
- Reduced flicker in Windows Terminal during high event throughput.
- No regression in watcher/worker orchestration.

## Bottom line
Reassessment confirms prior recommendation: implement `DiffAnsi` in the existing dashboard first, then migrate to a dedicated TUI frontend when ready.
