$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Processing\QC.StatusSet.psm1') -Force

function Assert-True([bool]$Cond, [string]$Msg) {
    if (-not $Cond) { throw $Msg }
}

$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('qc_wsg_' + [Guid]::NewGuid().ToString('n'))
$config = @{
    statusSet = @{
        localRoot = $tmpRoot
        manifestFileName = '_statusset.manifest.json'
        statusSetPdfName = '_statusset.pdf'
        forceRebuild = $false
    }
}

try {
    New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
    $sheetsTmp = Join-Path $tmpRoot 'sheets'
    New-Item -ItemType Directory -Path $sheetsTmp -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $sheetsTmp 'a.dgn') -Value 'x' -Encoding ascii -NoNewline
    Set-Content -LiteralPath (Join-Path $sheetsTmp 'a.pdf') -Value 'x' -Encoding ascii -NoNewline

    $state = Get-StatusSetLocalFolderState -RootFolder $sheetsTmp
    # No workspace yet -> manifest missing -> should enqueue
    $g1 = Test-StatusSetWatcherShouldEnqueue -Config $config -SourceFolder $sheetsTmp -FolderState $state
    Assert-True $g1.IsSuccess 'gate 1 success'
    Assert-True ([bool]$g1.Data.shouldEnqueue) 'no manifest should enqueue'

    $ws = [string](Get-StatusSetWorkspaceDirectory -LocalRoot $tmpRoot -SheetsFolderPath $sheetsTmp).Data.workspaceDir
    New-Item -ItemType Directory -Path $ws -Force | Out-Null
    $man = New-StatusSetManifestObject -State $state -SheetsFolderDisplay $sheetsTmp
    $man | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $ws '_statusset.manifest.json') -Encoding UTF8
    $ssPdf = Join-Path $ws '_statusset.pdf'
    Set-Content -LiteralPath $ssPdf -Value '%PDF' -Encoding ascii -NoNewline
    [System.IO.File]::SetLastWriteTimeUtc($ssPdf, [datetime]::UtcNow.AddMinutes(10))

    $g2 = Test-StatusSetWatcherShouldEnqueue -Config $config -SourceFolder $sheetsTmp -FolderState $state
    Assert-True $g2.IsSuccess 'gate 2 success'
    Assert-True (-not [bool]$g2.Data.shouldEnqueue) 'manifest + fresh status pdf should skip enqueue'
    Assert-True ([string]$g2.Data.gateReason -eq 'ALREADY_CURRENT') 'gate reason'

    'ok'
} finally {
    Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
}
