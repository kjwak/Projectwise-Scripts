$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Core\Core.Paths.psm1') -Force
Import-Module (Join-Path $root 'modules\Processing\QC.StatusSet.psm1') -Force

$failures = 0
function _Assert([bool]$ok, [string]$msg) {
    if (-not $ok) { Write-Host "  FAIL: $msg" -ForegroundColor Red; $script:failures++ }
    else          { Write-Host "  ok:   $msg" -ForegroundColor Green }
}

$localRoot = Join-Path $env:TEMP ("sshist_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $localRoot -Force | Out-Null

function _WriteBytes([string]$Path, [int]$Size) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllBytes($Path, (New-Object byte[] $Size))
}

function _PdfName([datetime]$dt) {
    return ('_StatusSet_{0}.pdf' -f $dt.ToString('yyyyMMdd_HHmmss'))
}

function _BakName([datetime]$dt) {
    return ('_statusset.manifest.json.bak_{0}' -f $dt.ToString('yyyyMMddHHmmss'))
}

try {
    $now = [TimeZoneInfo]::ConvertTimeFromUtc(
        [datetime]::UtcNow,
        [TimeZoneInfo]::FindSystemTimeZoneById('US Mountain Standard Time')
    )
    $ws = Join-Path $localRoot 'aaaaaaaaaaaaaaaa'
    $hist = Join-Path $ws '_history'
    New-Item -ItemType Directory -Path $hist -Force | Out-Null

    _WriteBytes (Join-Path $ws '_StatusSet.pdf') 50
    _WriteBytes (Join-Path $ws '_statusset.manifest.json') 40

    # Four copies today; keepRecentCount=2 should keep the newest two.
    $today = @(
        $now.Date.AddHours(10).AddMinutes(0)
        $now.Date.AddHours(12).AddMinutes(0)
        $now.Date.AddHours(14).AddMinutes(0)
        $now.Date.AddHours(16).AddMinutes(0)
    )
    foreach ($dt in $today) {
        _WriteBytes (Join-Path $hist (_PdfName $dt)) 100
        _WriteBytes (Join-Path $ws (_BakName $dt)) 20
    }

    # One copy per previous calendar day in the daily window.
    $dailyKept = @()
    for ($d = 1; $d -le 7; $d++) {
        $dt = $now.Date.AddDays(-1 * $d).AddHours(15)
        $dailyKept += $dt
        _WriteBytes (Join-Path $hist (_PdfName $dt)) 100
        _WriteBytes (Join-Path $ws (_BakName $dt)) 20
    }

    # Two copies in the same ISO week outside the daily window; keep newest only.
    $weekA = $now.Date.AddDays(-14).AddHours(9)
    $weekB = $now.Date.AddDays(-14).AddHours(18)
    _WriteBytes (Join-Path $hist (_PdfName $weekA)) 100
    _WriteBytes (Join-Path $hist (_PdfName $weekB)) 100
    _WriteBytes (Join-Path $ws (_BakName $weekA)) 20
    _WriteBytes (Join-Path $ws (_BakName $weekB)) 20

    # Older than 26 weeks: delete.
    $ancient = $now.Date.AddDays(-200).AddHours(8)
    _WriteBytes (Join-Path $hist (_PdfName $ancient)) 100
    _WriteBytes (Join-Path $ws (_BakName $ancient)) 20

    $cfg = @{
        statusSet = @{
            localRoot = $localRoot
            historyRetention = @{
                enabled = $true
                keepRecentCount = 2
                dailyDays = 7
                weeklyWeeks = 26
                maxGbPerFolder = 10
            }
        }
    }

    Write-Host "Test: dry-run reports deletes without removing files" -ForegroundColor Cyan
    $dry = Invoke-StatusSetHistoryRetention -Config $cfg -DryRun
    _Assert ($dry.IsSuccess) "dry-run succeeds"
    _Assert ($dry.Code -eq 'STATUS_SET_HISTORY_RETENTION_DONE') "dry-run code"
    _Assert ([bool]$dry.Data.dryRun) "dryRun flag set"
    _Assert ([int]$dry.Data.deleted -gt 0) "dry-run would delete files"
    $pdfCount = @(Get-ChildItem -LiteralPath $hist -Filter '*.pdf').Count
    $bakCount = @(Get-ChildItem -LiteralPath $ws -Filter '_statusset.manifest.json.bak_*').Count
    _Assert ($pdfCount -eq 14) "dry-run left all 14 history PDFs"
    _Assert ($bakCount -eq 14) "dry-run left all 14 manifest backups"
    _Assert (Test-Path -LiteralPath (Join-Path $ws '_StatusSet.pdf')) "live PDF remains after dry-run"
    _Assert (Test-Path -LiteralPath (Join-Path $ws '_statusset.manifest.json')) "live manifest remains after dry-run"

    Write-Host "Test: live run keeps recent + daily + weekly, deletes the rest" -ForegroundColor Cyan
    $live = Invoke-StatusSetHistoryRetention -Config $cfg
    _Assert ($live.IsSuccess) "live run succeeds"
    _Assert (-not [bool]$live.Data.dryRun) "live run is not dry-run"

    $remainPdf = @(Get-ChildItem -LiteralPath $hist -Filter '*.pdf' | ForEach-Object { $_.Name })
    $remainBak = @(Get-ChildItem -LiteralPath $ws -Filter '_statusset.manifest.json.bak_*' | ForEach-Object { $_.Name })

    $expectPdfKeep = @(
        (_PdfName $today[3])
        (_PdfName $today[2])
    ) + @($dailyKept | ForEach-Object { _PdfName $_ }) + @(_PdfName $weekB)
    $expectBakKeep = @(
        (_BakName $today[3])
        (_BakName $today[2])
    ) + @($dailyKept | ForEach-Object { _BakName $_ }) + @(_BakName $weekB)

    foreach ($n in $expectPdfKeep) {
        _Assert ($remainPdf -contains $n) ("kept PDF {0}" -f $n)
    }
    foreach ($n in @((_PdfName $today[0]), (_PdfName $today[1]), (_PdfName $weekA), (_PdfName $ancient))) {
        _Assert ($remainPdf -notcontains $n) ("deleted PDF {0}" -f $n)
    }
    foreach ($n in $expectBakKeep) {
        _Assert ($remainBak -contains $n) ("kept manifest bak {0}" -f $n)
    }
    foreach ($n in @((_BakName $today[0]), (_BakName $today[1]), (_BakName $weekA), (_BakName $ancient))) {
        _Assert ($remainBak -notcontains $n) ("deleted manifest bak {0}" -f $n)
    }
    _Assert ($remainPdf.Count -eq $expectPdfKeep.Count) ("PDF keep count {0}" -f $remainPdf.Count)
    _Assert ($remainBak.Count -eq $expectBakKeep.Count) ("manifest keep count {0}" -f $remainBak.Count)
    _Assert (Test-Path -LiteralPath (Join-Path $ws '_StatusSet.pdf')) "live PDF never deleted"
    _Assert (Test-Path -LiteralPath (Join-Path $ws '_statusset.manifest.json')) "live manifest never deleted"

    Write-Host "Test: per-folder byte cap drops oldest non-recent copies" -ForegroundColor Cyan
    $wsCap = Join-Path $localRoot 'bbbbbbbbbbbbbbbb'
    $histCap = Join-Path $wsCap '_history'
    New-Item -ItemType Directory -Path $histCap -Force | Out-Null
    for ($d = 0; $d -le 4; $d++) {
        $dt = $now.Date.AddDays(-1 * $d).AddHours(11)
        _WriteBytes (Join-Path $histCap (_PdfName $dt)) 100
    }
    $cfgCap = @{
        statusSet = @{
            localRoot = $wsCap
            historyRetention = @{
                enabled = $true
                keepRecentCount = 1
                dailyDays = 7
                weeklyWeeks = 0
                maxBytesPerFolder = 250
            }
        }
    }
    # Invoke with WorkspaceDir so we don't re-process the first workspace.
    $cap = Invoke-StatusSetHistoryRetention -Config $cfgCap -WorkspaceDir $wsCap
    _Assert ($cap.IsSuccess) "cap run succeeds"
    $capRemain = @(Get-ChildItem -LiteralPath $histCap -Filter '*.pdf')
    _Assert ($capRemain.Count -le 2) ("cap left {0} PDFs (<=2)" -f $capRemain.Count)
    $capBytes = ($capRemain | Measure-Object -Property Length -Sum).Sum
    _Assert ($capBytes -le 250) ("cap kept {0} bytes" -f $capBytes)
    $newestCap = _PdfName ($now.Date.AddHours(11))
    _Assert (@($capRemain | ForEach-Object { $_.Name }) -contains $newestCap) "cap kept newest (protected recent)"

    Write-Host "Test: enabled=false is a no-op" -ForegroundColor Cyan
    $wsSkip = Join-Path $localRoot 'cccccccccccccccc'
    $histSkip = Join-Path $wsSkip '_history'
    New-Item -ItemType Directory -Path $histSkip -Force | Out-Null
    _WriteBytes (Join-Path $histSkip (_PdfName $ancient)) 100
    $cfgSkip = @{
        statusSet = @{
            localRoot = $localRoot
            historyRetention = @{ enabled = $false }
        }
    }
    $skip = Invoke-StatusSetHistoryRetention -Config $cfgSkip
    _Assert ($skip.Code -eq 'STATUS_SET_HISTORY_RETENTION_SKIPPED') "disabled code"
    _Assert (Test-Path -LiteralPath (Join-Path $histSkip (_PdfName $ancient))) "disabled run left ancient PDF"

    if ($failures -gt 0) { Write-Host "`nFAILED ($failures)" -ForegroundColor Red; exit 1 }
    Write-Host "`nPASSED" -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $localRoot -Recurse -Force -ErrorAction SilentlyContinue
}
