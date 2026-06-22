$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$paths = @(
    'scripts\Watch-QCTrigger.ps1'
    'scripts\Reconcile-QCStatusSets.ps1'
    'scripts\Run-QCProcessor.ps1'
    'modules\Processing\QC.StatusSet.psm1'
    'modules\Queue\QC.Queue.Json.psm1'
)
$failed = $false
foreach ($p in $paths) {
    $full = Join-Path $root $p
    try {
        [System.Management.Automation.Language.Parser]::ParseFile($full, [ref]$null, [ref]$null) | Out-Null
        Write-Host ("OK:   " + $p)
    } catch {
        Write-Host ("FAIL: " + $p + " :: " + $_.Exception.Message) -ForegroundColor Red
        $failed = $true
    }
}
if ($failed) { exit 1 }
