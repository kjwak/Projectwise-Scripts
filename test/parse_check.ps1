$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$files = @(
    'modules\Queue\QC.Queue.Json.psm1',
    'modules\Processing\QC.StatusSet.psm1',
    'modules\Processing\QC.Processors.psm1',
    'scripts\service\Start-QCPipelineDashboard.ps1',
    'scripts\diagnostics\Show-QCQueueDiag.ps1',
    'scripts\service\Run-QCProcessor.ps1'
)
$bad = 0
foreach ($f in $files) {
    $full = Join-Path $repoRoot $f
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($full, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        $bad++
        Write-Host ("FAIL: {0} ({1} errors)" -f $f, $errors.Count) -ForegroundColor Red
        $errors | Select-Object -First 5 | ForEach-Object {
            Write-Host ("  L{0}: {1}" -f $_.Extent.StartLineNumber, $_.Message) -ForegroundColor Yellow
        }
    } else {
        Write-Host ("OK:   {0}" -f $f) -ForegroundColor Green
    }
}
exit $bad
