$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\QC.StatusSet.psm1') -Force

function Assert-True([bool]$Cond, [string]$Msg) {
    if (-not $Cond) { throw $Msg }
}

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qc_ss_test_' + [Guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    Set-Content -LiteralPath (Join-Path $tmp 'a.dgn') -Value 'x' -Encoding ascii -NoNewline
    Set-Content -LiteralPath (Join-Path $tmp 'a.pdf') -Value 'x' -Encoding ascii -NoNewline

    $st = Get-StatusSetLocalFolderState -RootFolder $tmp
    Assert-True ($st.pairedCount -eq 1) 'expected one pair'

    $man = New-StatusSetManifestObject -State $st -SheetsFolderDisplay 'C:\dummy'
    $cmpSelf = Test-StatusSetRebuildNeeded -Manifest $man -CurrentState $st -ForceRebuild:$false
    Assert-True ($cmpSelf.Data.needsRebuild -eq $false) 'manifest vs same state should not rebuild'

    $man2 = @{}
    foreach ($k in $man.Keys) { $man2[$k] = $man[$k] }
    $man2['folderStateHash'] = 'deadbeef'
    $cmpDrift = Test-StatusSetRebuildNeeded -Manifest $man2 -CurrentState $st -ForceRebuild:$false
    Assert-True ($cmpDrift.Data.needsRebuild -eq $true) 'hash mismatch should rebuild'

    # _statusset.pdf newer than all paired sheet PDFs → skip rebuild despite hash drift (mtime rule)
    $ssOut = Join-Path $tmp '_statusset.pdf'
    Set-Content -LiteralPath $ssOut -Value '%PDF-1.4 minimal' -Encoding ascii -NoNewline
    [System.IO.File]::SetLastWriteTimeUtc($ssOut, [datetime]::UtcNow.AddMinutes(30))
    $cmpSkipHashOnly = Test-StatusSetRebuildNeeded -Manifest $man2 -CurrentState $st -ForceRebuild:$false -StatusSetPdfPath $ssOut
    Assert-True ($cmpSkipHashOnly.Data.needsRebuild -eq $false) 'fresh status set PDF should skip rebuild on hash-only drift'

    # Stale _statusset.pdf → rebuild even when manifest hash still matches current state
    [System.IO.File]::SetLastWriteTimeUtc($ssOut, [datetime]::UtcNow.AddDays(-1))
    $cmpStale = Test-StatusSetRebuildNeeded -Manifest $man -CurrentState $st -ForceRebuild:$false -StatusSetPdfPath $ssOut
    Assert-True ($cmpStale.Data.needsRebuild -eq $true) 'stale status set PDF vs sheet PDFs should rebuild'
    Assert-True ($cmpStale.Data.reasons -contains 'STATUS_SET_PDF_OLDER_THAN_SHEET_PDF') 'stale rule should set reason'

    'ok'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
