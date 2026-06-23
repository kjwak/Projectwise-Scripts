$tests = @(
    'test\test_lock_steal_dead_pid.ps1',
    'test\test_recover_orphan_pid.ps1',
    'test\test_get_next_excludes.ps1',
    'test\test_queue_json.ps1',
    'test\test_orphan_lock_recovery.ps1',
    'test\test_merge_statusset_qpdf.ps1',
    'test\test_statusset_workspace_dir.ps1',
    'test\test_blacklist_pw_uri.ps1',
    'test\test_statusset_processor_default.ps1',
    'test\test_pw_writeback_failure.ps1',
    'test\test_qc_workflow.ps1',
    'test\test_move_qcjob_with_job.ps1',
    'test\test_statusset_reconcile.ps1',
    'test\test_statusset_doc_filter.ps1',
    'test\test_statusset_job_in_flight.ps1',
    'test\test_paired_sheets_array.ps1',
    'test\test_audit_poll_window.ps1',
    'test\test_audit_watch_match.ps1',
    'test\test_audit_events_db.ps1',
    'test\test_audit_workflow_triggers.ps1',
    'test\test_watcher_module_bootstrap.ps1'
)
$fail = 0
foreach ($t in $tests) {
    Write-Host ("`n=== " + $t) -ForegroundColor Cyan
    & powershell -NoProfile -ExecutionPolicy Bypass -File $t
    if ($LASTEXITCODE -ne 0) { Write-Host ('FAIL: ' + $t) -ForegroundColor Red; $fail++ }
    else { Write-Host ('PASS: ' + $t) -ForegroundColor Green }
}
Write-Host ''
if ($fail -gt 0) { Write-Host ("Failures: $fail") -ForegroundColor Red; exit 1 }
Write-Host 'ALL PASSED' -ForegroundColor Green
