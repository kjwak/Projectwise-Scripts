# Unit test: extension suffix filter + empty-array parameter binding (PS mandatory [object[]]).
$ErrorActionPreference = 'Stop'

function _Test-AcceptsEmptyDocs {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Docs
    )
    if (-not $Docs -or @($Docs).Count -eq 0) { return @() }
    return @($Docs)
}
$empty = @(_Test-AcceptsEmptyDocs -Docs @())
if ($empty.Count -ne 0) { throw 'AllowEmptyCollection regression: empty Docs must bind' }


function _Test-FilterByExt($Docs, $Extensions) {
    $suffixes = @($Extensions | ForEach-Object {
        $e = [string]$_
        if (-not $e.StartsWith('.')) { $e = '.' + $e }
        $e.ToLowerInvariant()
    })
    $out = @()
    foreach ($d in @($Docs)) {
        $name = [string]$d.Name
        $lower = $name.ToLowerInvariant()
        foreach ($suf in $suffixes) {
            if ($lower.EndsWith($suf)) { $out += $d; break }
        }
    }
    return $out
}

$docs = @(
    [pscustomobject]@{ Name = 'sheet.pdf' },
    [pscustomobject]@{ Name = 'sheet.dgn' },
    [pscustomobject]@{ Name = 'readme.txt' }
)
$filtered = @(_Test-FilterByExt -Docs $docs -Extensions @('.pdf', '.dgn'))
if ($filtered.Count -ne 2) { throw "expected 2 docs, got $($filtered.Count)" }

Write-Host 'OK: status set doc filter tests passed.' -ForegroundColor Green
