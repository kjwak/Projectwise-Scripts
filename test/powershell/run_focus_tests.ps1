$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$tests = @(
    'test\powershell\test_lock_steal_dead_pid.ps1',
    'test\powershell\test_host_aware_locks.ps1',
    'test\powershell\test_recover_orphan_pid.ps1',
    'test\powershell\test_get_next_excludes.ps1',
    'test\powershell\test_enabled_job_types.ps1',
    'test\powershell\test_host_throttle.ps1',
    'test\powershell\test_queue_json.ps1',
    'test\powershell\test_qc_prepend_child_wait_and_checkpoint.ps1',
    'test\powershell\test_fast_prepend_enqueue.ps1',
    'test\powershell\test_orphan_lock_recovery.ps1',
    'test\powershell\test_merge_statusset_qpdf.ps1',
    'test\powershell\test_statusset_workspace_dir.ps1',
    'test\powershell\test_statusset_history_retention.ps1',
    'test\powershell\test_blacklist_pw_uri.ps1',
    'test\powershell\test_statusset_processor_default.ps1',
    'test\powershell\test_pw_writeback_failure.ps1',
    'test\powershell\test_qc_workflow.ps1',
    'test\powershell\test_move_qcjob_with_job.ps1',
    'test\powershell\test_statusset_reconcile.ps1',
    'test\powershell\test_statusset_doc_filter.ps1',
    'test\powershell\test_statusset_job_in_flight.ps1',
    'test\powershell\test_paired_sheets_array.ps1',
    'test\powershell\test_audit_poll_window.ps1',
    'test\powershell\test_audit_watch_match.ps1',
    'test\powershell\test_audit_events_db.ps1',
    'test\powershell\test_audit_workflow_triggers.ps1',
    'test\powershell\test_watcher_module_bootstrap.ps1'
)
$fail = 0
foreach ($t in $tests) {
    $path = Join-Path $repoRoot $t
    Write-Host ("`n=== " + $t) -ForegroundColor Cyan
    & powershell -NoProfile -ExecutionPolicy Bypass -File $path
    if ($LASTEXITCODE -ne 0) { Write-Host ('FAIL: ' + $t) -ForegroundColor Red; $fail++ }
    else { Write-Host ('PASS: ' + $t) -ForegroundColor Green }
}
Write-Host ''
if ($fail -gt 0) { Write-Host ("Failures: $fail") -ForegroundColor Red; exit 1 }
Write-Host 'ALL PASSED' -ForegroundColor Green
