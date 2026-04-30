$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\QC.StatusSet.psm1') -Force

function Assert([bool]$Cond, [string]$Msg) {
    if (-not $Cond) { throw $Msg }
}

$iso1 = '2024-06-01T12:00:00.0000000Z'
$iso2 = '2024-06-02T12:00:00.0000000Z'

$rowA = @{
    stem = 'a'
    dir = 'documents\x'
    pdf = @{ length = '100'; lastWriteTimeUtc = $iso1; documentId = 'p1' }
    dgn = @{ length = '50'; lastWriteTimeUtc = '2020-01-01T00:00:00Z'; documentId = 'd1' }
}
$rowBDate = @{
    stem = 'a'
    dir = 'documents\x'
    pdf = @{ length = '100'; lastWriteTimeUtc = $iso2; documentId = 'p1' }
    dgn = @{ length = '50'; lastWriteTimeUtc = '2020-01-01T00:00:00Z'; documentId = 'd1' }
}
$rowBSize = @{
    stem = 'a'
    dir = 'documents\x'
    pdf = @{ length = '101'; lastWriteTimeUtc = $iso1; documentId = 'p1' }
    dgn = @{ length = '50'; lastWriteTimeUtc = '2020-01-01T00:00:00Z'; documentId = 'd1' }
}

Assert (Test-StatusSetSheetPdfTimestampMatch -ManifestSheetRow $rowA -CurrentPairedRow $rowA) 'same row timestamp match'
Assert (-not (Test-StatusSetSheetPdfTimestampMatch -ManifestSheetRow $rowA -CurrentPairedRow $rowBDate)) 'pdf date drift'
Assert (-not (Test-StatusSetSheetPdfTimestampMatch -ManifestSheetRow $rowA -CurrentPairedRow $rowBSize)) 'pdf size drift'

$sig1 = Get-StatusSetPairedSheetSignature -Row $rowA
$sig2 = Get-StatusSetPairedSheetSignature -Row $rowBDate
Assert ($sig1 -eq $sig2) 'ids+sizes equal => signature ignores pdf listing time'
$sig3 = Get-StatusSetPairedSheetSignature -Row $rowBSize
Assert ($sig1 -ne $sig3) 'pdf size change breaks signature'

'ok'
