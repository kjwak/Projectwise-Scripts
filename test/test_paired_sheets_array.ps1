# Regression: pairedSheets from foreach must be wrapped in @() so one row is still an array.
$ErrorActionPreference = 'Stop'

$one = foreach ($r in @([pscustomobject]@{ Stem = 'a'; Dir = 'd' })) {
    @{ stem = $r.Stem; dir = $r.Dir }
}
if ($one -is [array]) { throw 'bare foreach should yield a single hashtable, not [array]' }

$oneFixed = @(foreach ($r in @([pscustomobject]@{ Stem = 'a'; Dir = 'd' })) {
    @{ stem = $r.Stem; dir = $r.Dir }
})
if (-not ($oneFixed -is [array])) { throw '@(foreach) must yield [array] for one row' }
if ($oneFixed.Count -ne 1) { throw "expected Count 1, got $($oneFixed.Count)" }

$orderKey = (@($oneFixed | ForEach-Object { ($_.dir + '|' + $_.stem) }) -join "`n")
if ($orderKey -ne 'd|a') { throw "expected d|a, got: $orderKey" }

Write-Host 'OK: pairedSheets array regression passed.' -ForegroundColor Green
