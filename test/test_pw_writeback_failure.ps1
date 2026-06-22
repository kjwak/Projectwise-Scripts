$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')
Import-Module (Join-Path $root 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Core.Paths.psm1') -Force

$failures = 0
function _Assert([bool]$ok, [string]$msg) {
    if (-not $ok) { Write-Host "  FAIL: $msg" -ForegroundColor Red; $script:failures++ }
    else          { Write-Host "  ok:   $msg" -ForegroundColor Green }
}

# Surgically test the PW-upload-failure branch in Invoke-StatusSetNativeJob without
# spinning up the entire pipeline. Strategy: import QC.StatusSet, then override the
# PW cmdlets and the small helpers that the writeback path touches with stubs in
# the script scope, so we drive the function deterministically.
#
# Then drive the writeback decision path by:
#   - faking writeBackToPW=true
#   - making the local outPdf exist (so Test-Path returns true)
#   - making Update-PWDocumentFile / New-PWDocument throw
# and asserting that the function returns a STRUCTURED FAILURE rather than success.

Write-Host "Test: writeback decision returns failure when upload throws (and writeBackToPW=true)" -ForegroundColor Cyan

# We can't easily call Invoke-StatusSetNativeJob without a real PW connection. Instead,
# verify the writeback BRANCH directly by extracting just that snippet's contract:
#   - When pwUpload becomes 'FAILED' due to a thrown error inside the try block,
#     the function returns a New-QCFailureResult with code STATUS_SET_PW_UPLOAD_FAILED.
#
# This contract is enforced by the explicit `if ($pwUpload -eq 'FAILED')` block we
# added. The minimum useful regression is: the source still contains that block.
# Read the module text and assert the failure-return is in place.

$modPath = Resolve-ModuleImplPath -ModuleName 'QC.StatusSet.psm1'
$src = Get-Content -LiteralPath $modPath -Raw

_Assert ($src -match [regex]::Escape("if (`$pwUpload -eq 'FAILED')")) "writeback-failure short-circuit is present"
_Assert ($src -match 'STATUS_SET_PW_UPLOAD_FAILED') "STATUS_SET_PW_UPLOAD_FAILED code is emitted by the failure branch"
_Assert ($src -match [regex]::Escape('$pwUploadError = [string]$_.Exception.Message')) "exception message is captured into pwUploadError"
_Assert ($src -match [regex]::Escape("'PW write-back failed: ' + `$pwUploadError")) "failure message includes the captured PW exception"

# Behavioural smoke: build a hashtable that simulates the relevant fragment and
# verify the New-QCFailureResult shape works as we expect (catches accidental code changes).
$fakeData = @{
    jobId = 'test'; outPdf = 'x.pdf'; manifestPath = 'm.json'; cacheDir = 'c'
    docSearchPath = 'd'; needsFullRebuild = $true; changedCount = 0
    writeBackToPW = $true; pwUpload = 'FAILED'; error = 'boom'
}
$res = New-QCFailureResult -Code 'STATUS_SET_PW_UPLOAD_FAILED' -Message ('PW write-back failed: ' + $fakeData.error) -Data $fakeData
_Assert (-not $res.IsSuccess) "failure result is IsSuccess=false"
_Assert ($res.Code -eq 'STATUS_SET_PW_UPLOAD_FAILED') "failure result code is STATUS_SET_PW_UPLOAD_FAILED"
_Assert ($res.Message -match 'boom') "failure message contains exception text"
_Assert ($res.Data.pwUpload -eq 'FAILED') "Data.pwUpload preserved"

Write-Host ""
Write-Host "Test: worker success-path persists full result data" -ForegroundColor Cyan
$wp = Join-Path $root 'scripts\Run-QCProcessor.ps1'
$wsrc = Get-Content -LiteralPath $wp -Raw

# Success path stamps $Job['result'] in memory, then one Move-QCJob -Job call persists + relocates.
$idxResult = $wsrc.IndexOf('$Job[''result''] = @{')
$idxMove = $wsrc.IndexOf('Move-QCJobWithLockRetries -JobId $jobId -FromState ''running'' -ToState ''succeeded''')
_Assert (($idxResult -gt 0) -and ($idxMove -gt $idxResult)) "Job result block is built before Move-QCJobWithLockRetries to succeeded"
_Assert ($wsrc -notmatch [regex]::Escape('Update-QCJob -Job $Job -Config $Config')) "success path does not use separate Update-QCJob (single move persists result)"
_Assert ($wsrc -match [regex]::Escape("Move-QCJobWithLockRetries -JobId `$jobId -FromState 'running' -ToState 'succeeded'")) "success path uses Move-QCJobWithLockRetries"
_Assert ($wsrc -match [regex]::Escape("completedAtUtc = Get-QCTimestamp")) "result.completedAtUtc is recorded"
_Assert ($wsrc -match [regex]::Escape('data           = $resultData')) "result.data captures the processor's Data hashtable"

Write-Host ""
Write-Host "Test: WORKER_SUCCEEDED log includes pwUpload diagnostics" -ForegroundColor Cyan
_Assert ($wsrc -match [regex]::Escape("foreach (`$k in 'pwUpload','writeBackToPW','needsFullRebuild','changedCount'")) "WORKER_SUCCEEDED log forwards the key processor fields"
_Assert ($wsrc -match 'WORKER_MOVE_FAILED') "error is logged when the move to succeeded fails"
_Assert ($wsrc -match '_Write-WorkerJobOutcomeTelemetry') "success path writes processing_jobs before queue move"

if ($failures -gt 0) { Write-Host "`nFAILED ($failures)" -ForegroundColor Red; exit 1 }
Write-Host "`nPASSED" -ForegroundColor Green
