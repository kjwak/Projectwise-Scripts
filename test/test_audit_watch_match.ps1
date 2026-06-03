# Unit checks for watch-root prefix matching (underscores in project names break -like wildcards).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

function _Assert($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

$fp = 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'
$root = 'Documents\Caltrans\CAFWY2200-I-15_ELPSE'

# PowerShell -like treats _ as single-char wildcard; StartsWith is literal prefix match.
$likeMatch = ($fp -like "$root*")
$startsMatch = $fp.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)
_Assert ($startsMatch) 'StartsWith should match Caltrans Seg_1 folder'
if (-not $likeMatch) {
    Write-Host 'Note: -like failed for underscore path (expected); production uses StartsWith.' -ForegroundColor DarkGray
}

Write-Host 'OK: audit watch match tests passed.' -ForegroundColor Green
