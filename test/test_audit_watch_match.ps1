# Unit checks for watch-root prefix matching and pw_batch action code resolution.
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

# pw_batch rows only have o_action; [int]$row.pw_action silently becomes 0 (regression guard).
$pwBatchRow = @{ o_action = 1012 }
$legacyCode = 0
try { $legacyCode = [int]$pwBatchRow.pw_action } catch { }
_Assert ($legacyCode -eq 0) 'missing pw_action must not throw; old code treated this as action 0'
$fixedCode = 0
try {
    if ($pwBatchRow -is [hashtable] -and $pwBatchRow.ContainsKey('pw_action') -and $null -ne $pwBatchRow['pw_action']) {
        $fixedCode = [int]$pwBatchRow['pw_action']
    }
} catch { $fixedCode = 0 }
if ($fixedCode -eq 0) { $fixedCode = [int]$pwBatchRow.o_action }
_Assert ($fixedCode -eq 1012) 'trigger action must fall back to o_action for pw_batch rows'

Write-Host 'OK: audit watch match tests passed.' -ForegroundColor Green
