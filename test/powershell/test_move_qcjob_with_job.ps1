$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Core\Core.Paths.psm1') -Force
Import-Module (Join-Path $root 'modules\Queue\QC.Queue.Json.psm1') -Force

$failures = 0
function _Assert([bool]$ok, [string]$msg) {
    if (-not $ok) { Write-Host "  FAIL: $msg" -ForegroundColor Red; $script:failures++ }
    else          { Write-Host "  ok:   $msg" -ForegroundColor Green }
}

$qroot = Join-Path $env:TEMP ("qcq_test_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $qroot -Force | Out-Null
foreach ($s in 'pending','running','succeeded','failed','locks') {
    New-Item -ItemType Directory -Path (Join-Path $qroot $s) -Force | Out-Null
}
$config = @{ queue = @{ rootDir = $qroot } }

try {
    Write-Host "Test: Move-QCJob -Job persists in-memory hashtable to dest, single lock cycle" -ForegroundColor Cyan

    $jobId = 'qc_test_movejob_001'
    $job = @{
        id = $jobId; type = 'STATUS_SET_GEN'; status = 'running'
        attempts = 0; sourceFolder = 'documents\foo\bar\cadd\sheets'
        createdAt = ([datetime]::UtcNow.ToString('o'))
    }
    $runPath = Join-Path $qroot ("running\$jobId.json")
    $job | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $runPath -Encoding UTF8

    $job['result'] = @{
        code = 'STATUS_SET_OK'; message = 'ok'
        completedAtUtc = ([datetime]::UtcNow.ToString('o'))
        data = @{ pwUpload = 'UPDATED'; needsFullRebuild = $true; changedCount = 0 }
    }
    $job['status'] = 'succeeded'

    $mv = Move-QCJob -JobId $jobId -FromState 'running' -ToState 'succeeded' -Config $config -Job $job
    _Assert ($mv.IsSuccess)                                    "Move-QCJob -Job succeeded"
    _Assert ($mv.Code -eq 'QUEUE_JOB_MOVED')                   "Move-QCJob returned QUEUE_JOB_MOVED"

    $sucPath = Join-Path $qroot ("succeeded\$jobId.json")
    _Assert (Test-Path -LiteralPath $sucPath)                  "Job file landed in succeeded\"
    _Assert (-not (Test-Path -LiteralPath $runPath))           "Job file removed from running\"

    $written = Get-Content -LiteralPath $sucPath -Raw | ConvertFrom-Json
    _Assert ($written.status -eq 'succeeded')                  "Stamped status=succeeded"
    _Assert ($null -ne $written.result)                        "result block persisted"
    _Assert ($written.result.code -eq 'STATUS_SET_OK')         "result.code persisted"
    _Assert ($written.result.data.pwUpload -eq 'UPDATED')      "result.data.pwUpload persisted (the diagnostic that was being lost)"
    _Assert ([bool]$written.result.data.needsFullRebuild)      "result.data.needsFullRebuild persisted"
    _Assert ($null -ne $written.updatedAtUtc)                  "updatedAtUtc stamped"

    Write-Host ""
    Write-Host "Test: Move-QCJob idempotent when job already in succeeded\" -ForegroundColor Cyan
    $jobId2 = 'qc_test_already_succeeded'
    $sucOnly = Join-Path $qroot ("succeeded\$jobId2.json")
    @{ id = $jobId2; type = 'STATUS_SET_GEN'; status = 'succeeded' } | ConvertTo-Json | Set-Content -LiteralPath $sucOnly -Encoding UTF8
    $job2 = @{
        id = $jobId2; type = 'STATUS_SET_GEN'; status = 'succeeded'
        result = @{ code = 'STATUS_SET_OK'; message = 'ok'; data = @{ pwUpload = 'YES' } }
    }
    $mv2 = Move-QCJob -JobId $jobId2 -FromState 'running' -ToState 'succeeded' -Config $config -Job $job2
    _Assert ($mv2.IsSuccess) "Idempotent move from running when file is already in succeeded\"
    _Assert ($mv2.Code -eq 'QUEUE_JOB_ALREADY_MOVED') "Idempotent code is QUEUE_JOB_ALREADY_MOVED"
    $w2 = Get-Content -LiteralPath $sucOnly -Raw | ConvertFrom-Json
    _Assert ($w2.result.data.pwUpload -eq 'YES') "Idempotent move refreshes result payload on disk"

    Write-Host ""
    Write-Host "Test: Move-QCJob -Job rejects mismatched Job.id" -ForegroundColor Cyan
    $jobBad = @{ id = 'qc_some_other_id'; status = 'running' }
    $rmRun = Join-Path $qroot ("running\qc_test_mismatch.json")
    @{ id='qc_test_mismatch'; status='running' } | ConvertTo-Json | Set-Content -LiteralPath $rmRun -Encoding UTF8
    $mvBad = Move-QCJob -JobId 'qc_test_mismatch' -FromState 'running' -ToState 'succeeded' -Config $config -Job $jobBad
    _Assert (-not $mvBad.IsSuccess)                            "Mismatched Job.id is rejected"
    _Assert ($mvBad.Code -eq 'QUEUE_JOB_ID_MISMATCH')          "Mismatch code is QUEUE_JOB_ID_MISMATCH"
    _Assert (Test-Path -LiteralPath $rmRun)                    "Source file untouched after rejection"

    Write-Host ""
    Write-Host "Test: Move-QCJob (no -Job) preserves backward-compatible behavior" -ForegroundColor Cyan
    $jobLegacy = @{ id = 'qc_legacy_001'; type = 'QC_PREPEND'; status = 'running' }
    $legacyRun = Join-Path $qroot ("running\qc_legacy_001.json")
    $jobLegacy | ConvertTo-Json | Set-Content -LiteralPath $legacyRun -Encoding UTF8
    $mvLegacy = Move-QCJob -JobId 'qc_legacy_001' -FromState 'running' -ToState 'succeeded' -Config $config
    _Assert ($mvLegacy.IsSuccess)                              "Legacy call (no -Job) still succeeds"
    $legacySuc = Join-Path $qroot ("succeeded\qc_legacy_001.json")
    $j = Get-Content -LiteralPath $legacySuc -Raw | ConvertFrom-Json
    _Assert ($j.status -eq 'succeeded')                        "Legacy call stamps status=succeeded"

    Write-Host ""
    Write-Host "Test: Worker success path uses single Move-QCJob -Job call (no Update-QCJob)" -ForegroundColor Cyan
    $rqp = Join-Path $root 'scripts\service\Run-QCProcessor.ps1'
    $wsrc = Get-Content -LiteralPath $rqp -Raw

    $idxSucc = $wsrc.IndexOf("if (`$proc.IsSuccess)")
    $idxFailHeader = $wsrc.IndexOf("`$Job['lastError']")
    _Assert (($idxSucc -gt 0) -and ($idxFailHeader -gt $idxSucc)) "Success block found before failure block"
    $succBlock = $wsrc.Substring($idxSucc, $idxFailHeader - $idxSucc)
    _Assert ($succBlock -notmatch [regex]::Escape('Update-QCJob -Job $Job -Config $Config')) "Success path no longer invokes Update-QCJob"
    _Assert ($succBlock -match [regex]::Escape("Move-QCJobWithLockRetries -JobId `$jobId -FromState 'running' -ToState 'succeeded' -Config `$Config -Job `$Job")) "Success path uses Move-QCJobWithLockRetries (Move-QCJob -Job under one lock cycle)"

    $idxFmv = $wsrc.IndexOf("Move-QCJobWithLockRetries -JobId `$jobId -FromState 'running' -ToState `$target -Config `$Config -Job `$Job")
    _Assert ($idxFmv -gt 0)                                      "Failure path uses Move-QCJob -Job to target"
    _Assert ($wsrc -notmatch [regex]::Escape("Set-QCJobStatus -JobId `$jobId -Status `$target")) "Failure path no longer calls Set-QCJobStatus before move"

    if ($failures -gt 0) { Write-Host "`nFAILED ($failures)" -ForegroundColor Red; exit 1 }
    Write-Host "`nPASSED" -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue
}
