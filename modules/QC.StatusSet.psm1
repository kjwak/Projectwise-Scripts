<#
QC.StatusSet.psm1

Native StatusSet implementation that matches legacy/combine_status_set.ps1 method:
  - manifest path naming:   LocalRoot\status_set_manifest_<safe>.json
  - cache dir:              LocalRoot\status_set_cache\<safe>\*.pdf
  - manifest schema:        v2 (bump to invalidate old logic)
  - pairing:                PDF must have matching DGN/DWG (or CAD doc without extension) base name in same folder listing
  - ordering:               alphabetical by PDF filename (same as legacy)
  - native PW exports:     one `_export_<jobId>\\` folder; each PDF renamed to
                             `NNN_<originalname>.pdf` after export so names stay
                             unique without one folder per sheet; scratch cleanup
                             deletes files first then the directory (AV-friendlier
                             than `Remove-Item -Recurse` on many trees)
  - write-back:             optional _SSS-UpdatePWDocumentFileFromDisk (pwps_dab 24+: -InputDocuments/-NewFilePathName) / New-PWDocument
#>

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Paths.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'PW.Connection.psm1') -Force

$script:StatusSetManifestSchemaVersion = 2
$script:StatusSetOutputName = '_StatusSet.pdf'

# Per-process throttle (milliseconds) inserted between PDF/cache file operations
# (Move/Remove batches) so AV scanners (Fortinet, etc.) don't flag rapid temp-PDF
# churn. Default 2000 ms; override via top-level Config.fileOpThrottleMs in
# appsettings.json. Set to 0 to disable.
$script:_SSS_FsThrottleMs = 2000

# Per-PW-export throttle (milliseconds). The legacy combine_status_set.ps1 sleeps
# 400 ms after every Export-PWDocumentsSimple because Fortinet briefly opens each
# freshly-written PDF for scanning, and starting the next export immediately
# results in file-handle contention / partial writes / occasional corruption.
# Default 400 ms (matches legacy). Override via Config.statusSet.pwExportThrottleMs
# or top-level Config.pwExportThrottleMs. Set to 0 to disable.
$script:_SSS_PwExportThrottleMs = 400

function _SSS-FsThrottle {
    if ($script:_SSS_FsThrottleMs -and $script:_SSS_FsThrottleMs -gt 0) {
        Start-Sleep -Milliseconds $script:_SSS_FsThrottleMs
    }
}

function _SSS-PwExportThrottle {
    if ($script:_SSS_PwExportThrottleMs -and $script:_SSS_PwExportThrottleMs -gt 0) {
        Start-Sleep -Milliseconds $script:_SSS_PwExportThrottleMs
    }
}

function _SSS-UpdatePWDocumentFileFromDisk {
    <#
    .SYNOPSIS
    Wrapper for pwps_dab Update-PWDocumentFile with the 24.x parameter names.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputDocument,
        [Parameter(Mandatory)]
        [string]$LocalFilePath
    )
    # pwps_dab 24+: Update-PWDocumentFile [-InputDocuments] ... [-NewFilePathName] ...
    # (Older samples used -InputDocument / -FilePath; those do not exist and throw
    # "parameter cannot be found".)
    $docs = @($InputDocument | Where-Object { $_ })
    Update-PWDocumentFile -InputDocuments $docs -NewFilePathName $LocalFilePath -ErrorAction Stop | Out-Null
}

function _SSS-ApplyFsThrottleConfig([hashtable]$Config) {
    if (-not $Config) { return }
    try {
        if ($Config.ContainsKey('fileOpThrottleMs') -and $null -ne $Config['fileOpThrottleMs']) {
            $script:_SSS_FsThrottleMs = [int]$Config['fileOpThrottleMs']
        } elseif ($Config.ContainsKey('projectWise') -and $Config['projectWise']) {
            $pw = $Config['projectWise']
            if ($pw -is [hashtable] -and $pw.ContainsKey('fileOpThrottleMs') -and $null -ne $pw['fileOpThrottleMs']) {
                $script:_SSS_FsThrottleMs = [int]$pw['fileOpThrottleMs']
            } elseif ($pw.PSObject -and $pw.PSObject.Properties['fileOpThrottleMs'] -and $null -ne $pw.PSObject.Properties['fileOpThrottleMs'].Value) {
                $script:_SSS_FsThrottleMs = [int]$pw.PSObject.Properties['fileOpThrottleMs'].Value
            }
        }
        if ($Config.ContainsKey('pwExportThrottleMs') -and $null -ne $Config['pwExportThrottleMs']) {
            $script:_SSS_PwExportThrottleMs = [int]$Config['pwExportThrottleMs']
        }
        if ($Config.ContainsKey('statusSet') -and $Config['statusSet']) {
            $ss = $Config['statusSet']
            if ($ss -is [hashtable]) {
                if ($ss.ContainsKey('pwExportThrottleMs') -and $null -ne $ss['pwExportThrottleMs']) {
                    $script:_SSS_PwExportThrottleMs = [int]$ss['pwExportThrottleMs']
                }
            } elseif ($ss.PSObject -and $ss.PSObject.Properties['pwExportThrottleMs']) {
                $v = $ss.PSObject.Properties['pwExportThrottleMs'].Value
                if ($null -ne $v) { $script:_SSS_PwExportThrottleMs = [int]$v }
            }
        }
    } catch { }
}

function _SSS-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _SSS-EnsurePWDiscoveryCmdlets {
    [CmdletBinding()]
    param()

    if (Get-Command -Name Ensure-PWDiscoveryCmdlets -ErrorAction SilentlyContinue) {
        return Ensure-PWDiscoveryCmdlets
    }

    if ((Get-Command -Name Get-PWFolderView -ErrorAction SilentlyContinue) -or (Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue)) {
        return New-QCSuccessResult -Code 'PW_DISCOVERY_READY' -Message 'ProjectWise discovery cmdlets are available.' -Data @{}
    }

    try { Import-Module pwps_dab -Force -ErrorAction Stop | Out-Null } catch { }
    if ((Get-Command -Name Get-PWFolderView -ErrorAction SilentlyContinue) -or (Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue)) {
        return New-QCSuccessResult -Code 'PW_DISCOVERY_READY' -Message 'ProjectWise discovery cmdlets are available after re-import.' -Data @{}
    }

    return New-QCFailureResult -Code 'STATUS_SET_PW_DISCOVERY_INCOMPLETE' -Message 'ProjectWise discovery cmdlets are missing (Get-PWFolderView/Get-PWDocumentsBySearch); cannot distinguish an empty datasource from an incomplete pwps_dab runspace.' -Data @{ missingDiscoveryCmdlets = @('Get-PWFolderView','Get-PWDocumentsBySearch'); psModulePath = $env:PSModulePath }
}

function _SSS-TestPWDiscoveryCmdlets {
    [CmdletBinding()]
    param()

    if (Get-Command -Name Test-PWDiscoveryCmdlets -ErrorAction SilentlyContinue) {
        return [bool](Test-PWDiscoveryCmdlets)
    }
    return [bool]((Get-Command -Name Get-PWFolderView -ErrorAction SilentlyContinue) -or (Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue))
}

function _SSS-EnsureDir([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}


function _SSS-GetHashtableBool([hashtable]$Map, [string]$Name, [bool]$Default) {
    if (-not $Map -or -not $Map.ContainsKey($Name) -or $null -eq $Map[$Name]) { return $Default }
    try { return [bool]$Map[$Name] } catch { return $Default }
}

function _SSS-GetHashtableInt([hashtable]$Map, [string]$Name, [int]$Default) {
    if (-not $Map -or -not $Map.ContainsKey($Name) -or $null -eq $Map[$Name]) { return $Default }
    try { return [int]$Map[$Name] } catch { return $Default }
}

function _SSS-NewStatusSetOperationReport {
    param(
        [Parameter(Mandatory)]
        [string]$WorkspaceDir,
        [Parameter(Mandatory)]
        [string]$OutputPdf,
        [Parameter(Mandatory)]
        [string]$ManifestPath,
        [Parameter(Mandatory)]
        [hashtable]$Options
    )
    return [ordered]@{
        workspaceDir = $WorkspaceDir
        outputPdf = $OutputPdf
        manifestPath = $ManifestPath
        options = $Options
        downloads = @()
        writes = @()
        replaces = @()
        deletes = @()
        skips = @()
    }
}

function _SSS-AddStatusSetOperation {
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Report,
        [Parameter(Mandatory)]
        [ValidateSet('downloads','writes','replaces','deletes','skips')]
        [string]$Kind,
        [Parameter(Mandatory)]
        [hashtable]$Operation
    )
    $Report[$Kind] = @($Report[$Kind]) + @($Operation)
}

function _SSS-CleanupExpiredStatusSetStaging {
    param(
        [Parameter(Mandatory)]
        [string]$WorkspaceDir,
        [Parameter(Mandatory)]
        [int]$RetentionDays,
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$OperationReport,
        [Parameter(Mandatory = $false)]
        [switch]$DryRun
    )
    if ($RetentionDays -lt 1) { $RetentionDays = 1 }
    if (-not (Test-Path -LiteralPath $WorkspaceDir)) { return }
    $cutoff = (Get-QCWallClockNow).AddDays(-1 * $RetentionDays)
    $patterns = @('_export_*', '_render', '_render_*', '_pw_status_chunk_*.pdf', '*.next.*.pdf')
    foreach ($pattern in $patterns) {
        $items = @(Get-ChildItem -LiteralPath $WorkspaceDir -Filter $pattern -Force -ErrorAction SilentlyContinue)
        foreach ($item in $items) {
            $mtime = $item.LastWriteTimeUtc
            if ($mtime -gt $cutoff) {
                _SSS-AddStatusSetOperation -Report $OperationReport -Kind 'skips' -Operation @{ action='retention-skip'; path=[string]$item.FullName; lastWriteTimeUtc=ConvertTo-QCTimestamp $mtime; cutoffUtc=ConvertTo-QCTimestamp $cutoff }
                continue
            }
            _SSS-AddStatusSetOperation -Report $OperationReport -Kind 'deletes' -Operation @{ action='retention-delete'; path=[string]$item.FullName; lastWriteTimeUtc=ConvertTo-QCTimestamp $mtime; cutoffUtc=ConvertTo-QCTimestamp $cutoff }
            if (-not $DryRun) {
                try {
                    if ($item.PSIsContainer) { _SSS-RemoveExportDirContentsFileFirst -DirPath $item.FullName -RemoveEmptyDir $true }
                    else { Remove-Item -LiteralPath $item.FullName -Force -ErrorAction SilentlyContinue }
                } catch { }
                _SSS-FsThrottle
            }
        }
    }
}

function _SSS-InstallStatusSetPdf {
    param(
        [Parameter(Mandatory)]
        [string]$SourcePdf,
        [Parameter(Mandatory)]
        [string]$OutputPdf,
        [Parameter(Mandatory)]
        [bool]$AtomicReplace,
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$OperationReport,
        [Parameter(Mandatory = $false)]
        [string]$HistoryDir = ''
    )
    if ($SourcePdf -eq $OutputPdf) {
        _SSS-AddStatusSetOperation -Report $OperationReport -Kind 'skips' -Operation @{ action='install-output'; reason='source-is-output'; path=$OutputPdf }
        return New-QCSuccessResult -Code 'STATUS_SET_INSTALL_OK' -Message 'Output already written in place.' -Data @{ outputPdf=$OutputPdf; atomicReplace=$false }
    }
    _SSS-EnsureDir (Split-Path -Parent $OutputPdf)
    $backupPath = $null
    if (Test-Path -LiteralPath $OutputPdf) {
        if (_SSS-IsNullOrWhiteSpace $HistoryDir) { $HistoryDir = Join-Path (Split-Path -Parent $OutputPdf) '_history' }
        _SSS-EnsureDir $HistoryDir
        $stamp = Get-QCTimestampShort
        $backupPath = Join-Path $HistoryDir (([System.IO.Path]::GetFileNameWithoutExtension($OutputPdf)) + '_' + $stamp + [System.IO.Path]::GetExtension($OutputPdf))
    }
    _SSS-AddStatusSetOperation -Report $OperationReport -Kind 'replaces' -Operation @{ action='atomic-output-replace'; source=$SourcePdf; destination=$OutputPdf; backup=$backupPath; atomicReplace=$AtomicReplace }
    try {
        if ($AtomicReplace -and $backupPath -and (Test-Path -LiteralPath $OutputPdf)) {
            try {
                [System.IO.File]::Replace($SourcePdf, $OutputPdf, $backupPath, $true)
            } catch {
                # File.Replace is the least noisy path, but it can fail if staging and
                # output are on different volumes. Preserve rollback and fall back to a
                # throttled copy+move rather than deleting/recreating the final PDF.
                Copy-Item -LiteralPath $OutputPdf -Destination $backupPath -Force -ErrorAction Stop
                _SSS-FsThrottle
                Move-Item -LiteralPath $SourcePdf -Destination $OutputPdf -Force -ErrorAction Stop
            }
        } elseif ($backupPath -and (Test-Path -LiteralPath $OutputPdf)) {
            Copy-Item -LiteralPath $OutputPdf -Destination $backupPath -Force -ErrorAction Stop
            _SSS-FsThrottle
            Move-Item -LiteralPath $SourcePdf -Destination $OutputPdf -Force -ErrorAction Stop
        } else {
            Move-Item -LiteralPath $SourcePdf -Destination $OutputPdf -Force -ErrorAction Stop
        }
        return New-QCSuccessResult -Code 'STATUS_SET_INSTALL_OK' -Message 'Status set PDF installed.' -Data @{ outputPdf=$OutputPdf; backupPath=$backupPath; atomicReplace=$AtomicReplace }
    } catch {
        return New-QCFailureResult -Code 'STATUS_SET_INSTALL_FAILED' -Message 'Failed to install status set PDF.' -Data @{ source=$SourcePdf; outputPdf=$OutputPdf; backupPath=$backupPath; errorMessage=$_.Exception.Message }
    }
}

function _SSS-NormalizeStatusSetExportLeaf {
    <#
    Strip leading NNN_ prefix from export filenames so retries do not build 002_002_name.pdf.
    #>
    param([string]$Leaf)
    if ([string]::IsNullOrWhiteSpace($Leaf)) { return $Leaf }
    if ($Leaf -match '^\d{3}_(.+)$') { return $Matches[1] }
    return $Leaf
}

function _SSS-MoveStatusSetExportToUniqueName {
    <#
    Rename exported PDF to NNN_<basename>.pdf inside the job export folder.
    Removes a stale destination file when the same job is retried (folder is reused per job id).
    Retries Move/Remove when AV briefly locks the freshly exported PDF.
    #>
    param(
        [Parameter(Mandatory)][string]$LocalPath,
        [Parameter(Mandatory)][string]$ExportWorkDir,
        [Parameter(Mandatory)][int]$SequenceIndex
    )

    $leaf = _SSS-NormalizeStatusSetExportLeaf -Leaf ([System.IO.Path]::GetFileName($LocalPath))
    $uniqueLeaf = ('{0:000}_{1}' -f $SequenceIndex, $leaf)
    $uniquePath = Join-Path $ExportWorkDir $uniqueLeaf

    try {
        $resolvedLocal = (Resolve-Path -LiteralPath $LocalPath -ErrorAction Stop).Path
        $resolvedUnique = (Resolve-Path -LiteralPath $uniquePath -ErrorAction SilentlyContinue).Path
        if ($resolvedUnique -and ($resolvedLocal -eq $resolvedUnique)) {
            return $uniquePath
        }
    } catch { }

    if ($LocalPath -eq $uniquePath) { return $uniquePath }

    _SSS-FsThrottle

    $lastEx = $null
    for ($a = 1; $a -le 22; $a++) {
        try {
            if (Test-Path -LiteralPath $uniquePath) {
                Remove-Item -LiteralPath $uniquePath -Force -ErrorAction Stop
                _SSS-FsThrottle
            }
            Move-Item -LiteralPath $LocalPath -Destination $uniquePath -Force -ErrorAction Stop
            _SSS-FsThrottle
            return $uniquePath
        } catch {
            $lastEx = $_
        }
        if ($a -ge 22) { break }
        Start-Sleep -Milliseconds ([Math]::Min(3000, 200 + ($a * 150)))
    }
    throw $lastEx
}

function _SSS-RemoveExportDirContentsFileFirst {
    <#
    .SYNOPSIS
    Delete files inside an export scratch dir one-by-one (friendlier to AV than -Recurse on trees).
    Optionally removes the now-empty directory; failures are non-terminating.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DirPath,
        [Parameter(Mandatory = $false)]
        [bool]$RemoveEmptyDir = $true
    )
    if (-not (Test-Path -LiteralPath $DirPath)) { return }
    try {
        $files = @(Get-ChildItem -LiteralPath $DirPath -File -ErrorAction SilentlyContinue)
        $n = 0
        foreach ($file in $files) {
            try {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            } catch { }
            $n++
            if ($script:_SSS_FsThrottleMs -gt 0 -and (($n % 2) -eq 0)) {
                Start-Sleep -Milliseconds ([Math]::Min(2000, $script:_SSS_FsThrottleMs))
            }
        }
        if ($RemoveEmptyDir) {
            Start-Sleep -Milliseconds 150
            try {
                Remove-Item -LiteralPath $DirPath -Force -ErrorAction Stop
            } catch {
                try {
                    Get-ChildItem -LiteralPath $DirPath -Recurse -Force -ErrorAction SilentlyContinue |
                        Remove-Item -Force -ErrorAction SilentlyContinue
                } catch { }
            }
        }
    } catch { }
}

function _SSS-CleanAllExportScratchDirsInWorkspace {
    <#
    .SYNOPSIS
    Removes every workspace\_export_* directory using file-first cleanup (handles legacy
    per-sheet subdirs and the newer single-folder layout).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$WorkspaceDir
    )
    if (-not (Test-Path -LiteralPath $WorkspaceDir)) { return }
    $dirs = @(Get-ChildItem -LiteralPath $WorkspaceDir -Directory -Filter '_export_*' -ErrorAction SilentlyContinue)
    foreach ($d in $dirs) {
        # Legacy layout: nested _export_* subdirs under a parent _export_* 
        $nested = @(Get-ChildItem -LiteralPath $d.FullName -Directory -Filter '_export_*' -ErrorAction SilentlyContinue)
        foreach ($nd in $nested) {
            _SSS-RemoveExportDirContentsFileFirst -DirPath $nd.FullName -RemoveEmptyDir $true
        }
        _SSS-RemoveExportDirContentsFileFirst -DirPath $d.FullName -RemoveEmptyDir $true
        if ($script:_SSS_FsThrottleMs -gt 0) { _SSS-FsThrottle }
    }
}

function _SSS-ParseIsoDateTime([object]$Value) {
    if ($null -eq $Value) { return $null }
    try {
        if ($Value -is [DateTime]) { return ([DateTime]$Value).ToUniversalTime() }
        if ($Value -is [DateTimeOffset]) { return ([DateTimeOffset]$Value).UtcDateTime }
        $s = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($s)) { return $null }
        return [DateTime]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
    } catch {
        return $null
    }
}

function _SSS-NormalizeFileSize([object]$Value) {
    if ($null -eq $Value) { return '' }
    try {
        if ($Value -is [int] -or $Value -is [long]) { return ([int64]$Value).ToString([System.Globalization.CultureInfo]::InvariantCulture) }
        if ($Value -is [double] -or $Value -is [decimal]) { return ([int64][math]::Round([double]$Value)).ToString([System.Globalization.CultureInfo]::InvariantCulture) }
        $s = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($s)) { return '' }
        # PW sometimes returns sizes formatted as "12,345" or "12345.0"
        $s2 = ($s -replace '[,\\s]', '')
        $n = 0L
        if ([int64]::TryParse($s2, [ref]$n)) { return $n.ToString([System.Globalization.CultureInfo]::InvariantCulture) }
        $d = 0.0
        if ([double]::TryParse($s2, [ref]$d)) { return ([int64][math]::Round($d)).ToString([System.Globalization.CultureInfo]::InvariantCulture) }
        return $s2
    } catch {
        return ''
    }
}

function _SSS-Sha256TextHex([string]$Text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash($bytes) } finally { $sha.Dispose() }
    return (([BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant())
}

function _SSS-Sha256HexOfPath([string]$Path) {
    # Stable hex digest of a normalized path. Used by Get-StatusSetWorkspaceDirectory
    # to derive a deterministic per-folder workspace key so cached manifests survive
    # across runs and different folders never collide. Returns $null if the input
    # cannot be normalized (caller treats that as STATUS_SET_PATH_NORMALIZE_FAILED).
    if (_SSS-IsNullOrWhiteSpace $Path) { return $null }
    $norm = $null
    try {
        $r = Normalize-QCPath -Path $Path
        if ($r -and $r.IsSuccess -and $r.Data -and $r.Data.path) { $norm = [string]$r.Data.path }
    } catch { $norm = $null }
    if (_SSS-IsNullOrWhiteSpace $norm) {
        # Fallback to a manual normalization so the workspace can still be derived
        # even if Core.Paths failed to load (e.g. import-order edge cases).
        $norm = ([string]$Path).Trim() -replace '/', '\' -replace '\\{2,}', '\'
        $norm = $norm.TrimEnd('\').ToLowerInvariant()
    }
    if (_SSS-IsNullOrWhiteSpace $norm) { return $null }
    return (_SSS-Sha256TextHex -Text $norm)
}

function _SSS-GetFileHashSha256([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
    } finally {
        $sha.Dispose()
    }
    return ([BitConverter]::ToString($hash) -replace '-', '')
}

function _SSS-AssertCommand([string]$ExePath) {
    $cmd = Get-Command $ExePath -ErrorAction SilentlyContinue
    if (-not $cmd -and -not (Test-Path -LiteralPath $ExePath)) {
        throw "Required executable not found: '$ExePath'. Install it or configure statusSet.qpdfExe."
    }
}

function _SSS-GetPdfPageCount([string]$Path, [string]$QpdfExe) {
    try {
        $info = & $QpdfExe --show-npages $Path 2>&1
        $line = ($info | Select-Object -First 1) -as [string]
        if ($line -match '^\d+$') { return [int]$line }
    } catch { }
    return 1
}

function _SSS-MergePdfs([string[]]$PdfPaths, [string]$OutPath, [string]$QpdfExe) {
    if ($PdfPaths.Count -eq 0) { throw 'No PDFs to merge.' }
    if ($PdfPaths.Count -eq 1) {
        Copy-Item -LiteralPath $PdfPaths[0] -Destination $OutPath -Force
        return
    }
    $maxPerBatch = 100
    $chunkDir = Split-Path -Parent $OutPath
    if (-not $chunkDir) { $chunkDir = $env:TEMP }
    _SSS-EnsureDir $chunkDir
    if ($PdfPaths.Count -le $maxPerBatch) {
        $allArgs = @('--empty', '--pages') + @($PdfPaths) + @('--', $OutPath)
        & $QpdfExe @allArgs | Out-Null
        if (-not (Test-Path -LiteralPath $OutPath)) { throw "qpdf failed to create output: $OutPath" }
        return
    }
    $chunkTemp = @()
    try {
        for ($i = 0; $i -lt $PdfPaths.Count; $i += $maxPerBatch) {
            $end = [Math]::Min($i + $maxPerBatch - 1, $PdfPaths.Count - 1)
            $batch = @($PdfPaths[$i..$end])
            $chunkOut = Join-Path $chunkDir ("_pw_status_chunk_{0}.pdf" -f [guid]::NewGuid().ToString('N'))
            $chunkTemp += $chunkOut
            _SSS-MergePdfs -PdfPaths $batch -OutPath $chunkOut -QpdfExe $QpdfExe
        }
        _SSS-MergePdfs -PdfPaths $chunkTemp -OutPath $OutPath -QpdfExe $QpdfExe
    } finally {
        foreach ($t in $chunkTemp) { Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue }
    }
    if (-not (Test-Path -LiteralPath $OutPath)) { throw "qpdf failed to create output: $OutPath" }
}

function _SSS-ReplacePdfPages([string]$CombinedPath, [string]$NewPdfPath, [int]$PageStart, [int]$PageEnd, [string]$OutPath, [string]$QpdfExe) {
    $totalPages = _SSS-GetPdfPageCount -Path $CombinedPath -QpdfExe $QpdfExe
    if ($PageStart -lt 1 -or $PageEnd -lt $PageStart) { throw "Replace-PdfPages: invalid range $PageStart-$PageEnd." }
    if ($PageEnd -gt $totalPages -or $PageStart -gt $totalPages) { throw "Replace-PdfPages: range $PageStart-$PageEnd is outside combined PDF ($totalPages page(s))." }
    $pageArgs = @()
    if ($PageStart -gt 1) { $pageArgs += $CombinedPath; $pageArgs += "1-$($PageStart - 1)" }
    $pageArgs += $NewPdfPath
    if ($PageEnd -lt $totalPages) { $pageArgs += $CombinedPath; $pageArgs += "$($PageEnd + 1)-z" }
    $allArgs = @($CombinedPath, '--pages') + $pageArgs + @('--', $OutPath)
    & $QpdfExe @allArgs | Out-Null
    if (-not (Test-Path -LiteralPath $OutPath)) { throw "qpdf failed to replace pages: $OutPath" }
}

function _SSS-PWGetProp([object]$Obj, [string]$Name) {
    try {
        if ($null -eq $Obj) { return $null }
        if ($Obj.PSObject -and $Obj.PSObject.Properties[$Name]) { return $Obj.$Name }
    } catch { }
    return $null
}

function _SSS-GetDocName([object]$Doc) {
    foreach ($n in @('Name','DocumentName')) {
        $v = _SSS-PWGetProp -Obj $Doc -Name $n
        if ($v) { return [string]$v }
    }
    try { return [System.IO.Path]::GetFileName([string](_SSS-PWGetProp -Obj $Doc -Name 'FullPath')) } catch { }
    return ''
}

function _SSS-PWGetDocName([object]$Doc) {
    # Back-compat shim: older code referenced _SSS-PWGetDocName; it maps to the same name extraction.
    return _SSS-GetDocName $Doc
}

function _SSS-GetDocLastModified([object]$Doc) {
    # Prefer FileUpdateDateUtc (stable for "reprint"/new version detection).
    # Fall back to other date columns for older PW return shapes.
    foreach ($n in @(
        'FileUpdateDateUtc',
        'FileUpdatedDateUtc',
        'FileUpdateDate',
        'FileUpdatedDate',
        'DocumentUpdateDate',
        'VersionModifiedDate',
        'Version Modified Date'
    )) {
        $v = _SSS-PWGetProp -Obj $Doc -Name $n
        $dt = _SSS-ParseIsoDateTime $v
        if ($dt) { return $dt }
    }
    return $null
}

function _SSS-PWGetDocLastModifiedUtcIso([object]$Doc) {
    # Back-compat shim for older watcher/state code: return UTC ISO string (or '').
    $dt = _SSS-GetDocLastModified $Doc
    if ($dt) { return ConvertTo-QCTimestamp $dt }
    return ''
}

function _SSS-GetPwFolderPath([string]$Path) {
    $p = ($Path -as [string]).Trim().TrimEnd('\')
    return ($p -replace '^Documents\\', '')
}

function _SSS-GetPwSheetsTryPaths([string]$SheetsFolderPath) {
    $trimmed = ($SheetsFolderPath -as [string]).Trim().TrimEnd('\')
    if (-not $trimmed) { return @() }
    $pw = _SSS-GetPwFolderPath $trimmed
    $withDoc = "Documents\\$pw"
    $ordered = @()
    if ($pw) { $ordered += $pw }
    if ($withDoc -and $ordered -notcontains $withDoc) { $ordered += $withDoc }
    if ($trimmed -match '^Documents\\' -and $trimmed -ne $withDoc -and $ordered -notcontains $trimmed) { $ordered += $trimmed }
    return $ordered
}

function _SSS-PWListDocsInFolder([string]$FolderPath, [string[]]$DateCols) {
    # Legacy order: search-with-columns -> plain search -> folder view
    try {
        $cmd = Get-Command -Name Get-PWDocumentsBySearchWithReturnColumns -ErrorAction SilentlyContinue
        if ($cmd) {
            $withCols = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -ColumnsToReturn $DateCols -PopulatePath -ErrorAction SilentlyContinue
            if ($withCols -and @($withCols).Count -gt 0) { return @($withCols) }
        }
    } catch { }
    try {
        $plain = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath -ErrorAction SilentlyContinue)
        if ($plain.Count -gt 0) { return @($plain) }
    } catch { }
    try {
        $folder = Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction SilentlyContinue
        if ($folder) {
            $view = $folder | Get-PWFolderView -ErrorAction SilentlyContinue
            if ($view -and $view.Documents) { return @($view.Documents) }
            if ($view -and $view.Children) { return @($view.Children | Where-Object { $_.DocumentID }) }
        }
    } catch { }
    return @()
}

function _SSS-FilterDocsByExtensions {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Docs,
        [Parameter(Mandatory)][string[]]$Extensions
    )
    if (-not $Docs -or @($Docs).Count -eq 0) { return @() }
    $suffixes = @($Extensions | Where-Object { $_ } | ForEach-Object {
        $e = [string]$_
        if (-not $e.StartsWith('.')) { $e = '.' + $e }
        $e.ToLowerInvariant()
    } | Select-Object -Unique)
    if ($suffixes.Count -eq 0) { return @($Docs) }

    $out = @()
    foreach ($d in @($Docs)) {
        $name = _SSS-PWGetDocName -Doc $d
        if (-not $name) { continue }
        $lower = $name.ToLowerInvariant()
        foreach ($suf in $suffixes) {
            if ($lower.EndsWith($suf)) { $out += $d; break }
        }
    }
    return @($out)
}

function _SSS-DedupePwDocRows {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Docs
    )
    if (-not $Docs -or @($Docs).Count -eq 0) { return @() }
    $seen = @{}
    $out = @()
    foreach ($d in @($Docs)) {
        if (-not $d) { continue }
        $key = $null
        try {
            $g = _SSS-PWGetProp -Obj $d -Name 'DocumentGUID'
            if ($g) { $key = ('guid|' + [string]$g).ToLowerInvariant() }
        } catch { }
        if (-not $key) {
            $n = _SSS-PWGetDocName -Doc $d
            if ($n) { $key = ('name|' + $n).ToLowerInvariant() }
        }
        if (-not $key) { $key = ('obj|' + $d.GetHashCode()) }
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $out += $d
    }
    return @($out)
}

function _SSS-PWListDocsInFolderViaDiscovery {
    param([Parameter(Mandatory)][string]$FolderPath)
    try {
        $discPath = Join-Path $PSScriptRoot 'PW.Discovery.psm1'
        if (-not (Get-Module -Name 'PW.Discovery' -ErrorAction SilentlyContinue)) {
            Import-Module $discPath -Force -ErrorAction Stop
        }
        return @(Get-PWDocumentsInFolder -FolderPath $FolderPath)
    } catch {
        return @()
    }
}

function _SSS-PWListDocsInFolderByExtensions {
    <#
    .SYNOPSIS
    Targeted PW document listing for status-set fingerprinting.
    .DESCRIPTION
    Attempts wildcard search per extension first. Many PW environments return empty for
    DocumentName=*.pdf even when the folder has documents; we then fall back to full-folder
    listing (same strategies as QC prepend / Get-PWDocumentsInFolder).
    Sets $script:_SSS_LastDocListingMethod: wildcard | folder_full | discovery_folder.
    #>
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string[]]$Extensions,
        [string[]]$DateCols
    )
    if (-not $DateCols -or $DateCols.Count -eq 0) {
        $DateCols = @('Name','DocumentID','DocumentGUID','FileUpdatedDate','FileUpdateDate','DocumentUpdateDate','VersionModifiedDate','Version Modified Date','FileSize','Size','StateName')
    }

    $script:_SSS_LastDocListingMethod = 'wildcard'

    $cmdCols = Get-Command -Name Get-PWDocumentsBySearchWithReturnColumns -ErrorAction SilentlyContinue
    $cmdPlain = Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue
    if (-not $cmdCols -and -not $cmdPlain) {
        $script:_SSS_LastDocListingMethod = 'folder_full'
        return @(_SSS-FilterDocsByExtensions -Docs @(_SSS-PWListDocsInFolder -FolderPath $FolderPath -DateCols $DateCols) -Extensions $Extensions)
    }

    $all = @()
    $patternFailed = $false
    foreach ($ext in @($Extensions | Where-Object { $_ })) {
        $pattern = "*$ext"
        $gotRows = $false
        try {
            if ($cmdCols) {
                $rows = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $pattern -ColumnsToReturn $DateCols -PopulatePath -ErrorAction SilentlyContinue
                if ($rows -and @($rows).Count -gt 0) { $all += @($rows); $gotRows = $true }
            }
        } catch { $patternFailed = $true }
        if (-not $gotRows) {
            try {
                if ($cmdPlain) {
                    $rows2 = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $pattern -PopulatePath -ErrorAction SilentlyContinue)
                    if ($rows2.Count -gt 0) { $all += @($rows2); $gotRows = $true }
                }
            } catch { $patternFailed = $true }
        }
        if (-not $gotRows -and -not $cmdCols -and -not $cmdPlain) { $patternFailed = $true }
    }

    if ($patternFailed -and $all.Count -eq 0) {
        $script:_SSS_LastDocListingMethod = 'folder_full'
        return @(_SSS-FilterDocsByExtensions -Docs @(_SSS-PWListDocsInFolder -FolderPath $FolderPath -DateCols $DateCols) -Extensions $Extensions)
    }

    if (@($all).Count -gt 0) {
        $all = @(_SSS-DedupePwDocRows -Docs $all)
        if ($all.Count -gt 0) { return $all }
    }

    $script:_SSS_LastDocListingMethod = 'folder_full'
    $fromFolder = @(_SSS-FilterDocsByExtensions -Docs @(_SSS-PWListDocsInFolder -FolderPath $FolderPath -DateCols $DateCols) -Extensions $Extensions)
    if (@($fromFolder).Count -gt 0) {
        $fromFolder = @(_SSS-DedupePwDocRows -Docs $fromFolder)
        if ($fromFolder.Count -gt 0) { return $fromFolder }
    }

    $script:_SSS_LastDocListingMethod = 'discovery_folder'
    $fromDisc = @(_SSS-FilterDocsByExtensions -Docs @(_SSS-PWListDocsInFolderViaDiscovery -FolderPath $FolderPath) -Extensions $Extensions)
    if (@($fromDisc).Count -gt 0) {
        return @(_SSS-DedupePwDocRows -Docs $fromDisc)
    }
    return @()
}

function Get-StatusSetManifestPathLegacy([string]$FolderPath, [string]$LocalRoot) {
    $safe = (($FolderPath -replace '[\\/:]', '_').Trim())
    if (-not $safe) { $safe = 'default' }
    return Join-Path $LocalRoot ("status_set_manifest_{0}.json" -f $safe)
}

function Get-StatusSetCacheDirLegacy([string]$FolderPath, [string]$LocalRoot) {
    $safe = (($FolderPath -replace '[\\/:]', '_').Trim())
    if (-not $safe) { $safe = 'default' }
    return Join-Path (Join-Path $LocalRoot 'status_set_cache') $safe
}

function Read-StatusSetManifestLegacy([string]$ManifestPath) {
    if (-not (Test-Path -LiteralPath $ManifestPath)) { return $null }
    try { return (Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}

function Write-StatusSetManifestLegacy([string]$ManifestPath, [hashtable]$ManifestObj) {
    _SSS-EnsureDir (Split-Path -Parent $ManifestPath)
    $ManifestObj | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8
}

function _SSS-SelectPairedPdfDocsLegacy([object[]]$AllDocs) {
    $hasCad = @{}
    $nonCadExt = '\.(pdf|xlsx|xls|doc|docx|txt|zip|jpg|jpeg|png|gif|bmp|tif|tiff|log|xml|json|csv)$'
    foreach ($doc in @($AllDocs)) {
        $name = _SSS-GetDocName $doc
        if (-not $name) { continue }
        if ($name -match '\.pdf$') { continue }
        if ($name -match $nonCadExt) { continue }
        if ($name -match '\.(dgn|dwg)$') { $base = ($name -replace '\.(dgn|dwg)$', '').ToLowerInvariant() }
        else { $base = $name.ToLowerInvariant() }
        $hasCad[$base] = $true
    }
    $pdfDocs = @()
    foreach ($doc in @($AllDocs)) {
        $name = _SSS-GetDocName $doc
        if (-not $name -or $name -notmatch '\.pdf$') { continue }
        if ($name -match '-qc\.pdf$') { continue }
        $base = ($name -replace '\.pdf$', '').ToLowerInvariant()
        if (-not $hasCad[$base]) { continue }
        $pdfDocs += $doc
    }
    return @($pdfDocs | Sort-Object { _SSS-GetDocName $_ })
}

function Get-StatusSetLocalFolderState {
    <#
    .SYNOPSIS
    Watcher helper: stable fingerprint for local folder (legacy pairing: PDF requires matching DGN/DWG base).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RootFolder
    )

    $pdfs = @(Get-ChildItem -LiteralPath $RootFolder -Recurse -File -Filter '*.pdf' -ErrorAction SilentlyContinue)
    $cads = @()
    $cads += @(Get-ChildItem -LiteralPath $RootFolder -Recurse -File -Filter '*.dgn' -ErrorAction SilentlyContinue)
    $cads += @(Get-ChildItem -LiteralPath $RootFolder -Recurse -File -Filter '*.dwg' -ErrorAction SilentlyContinue)

    $cadByDir = @{}
    foreach ($c in $cads) {
        $dir = [string]$c.DirectoryName
        if (-not $cadByDir.ContainsKey($dir)) { $cadByDir[$dir] = @{} }
        $stem = [System.IO.Path]::GetFileNameWithoutExtension([string]$c.Name)
        if ($stem) { $cadByDir[$dir][$stem.ToLowerInvariant()] = $c }
    }

    $lines = @()
    $pairedSheets = @()
    $paired = 0
    foreach ($p in $pdfs) {
        $name = [string]$p.Name
        if (-not $name) { continue }
        if ($name -match '(?i)-qc\.pdf$') { continue }
        if ($name -match '(?i)_statusset\.pdf$') { continue }
        if ($name -match '(?i)_statusset\.manifest\.json$') { continue }

        $dir = [string]$p.DirectoryName
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($name)
        if (-not $stem) { continue }
        $stemKey = $stem.ToLowerInvariant()

        if (-not ($cadByDir.ContainsKey($dir) -and $cadByDir[$dir].ContainsKey($stemKey))) { continue }
        $c = $cadByDir[$dir][$stemKey]

        $paired++
        $pairedSheets += @{
            stem = $stemKey
            dir = ([string]$dir).ToLowerInvariant()
            pdf = @{
                name = [string]$p.Name
                fullName = [string]$p.FullName
                length = [int64]$p.Length
                lastWriteTimeUtc = ConvertTo-QCTimestamp $p.LastWriteTimeUtc
            }
            cad = @{
                name = [string]$c.Name
                fullName = [string]$c.FullName
                length = [int64]$c.Length
                lastWriteTimeUtc = ConvertTo-QCTimestamp $c.LastWriteTimeUtc
            }
        }
        $lines += (@(
            $stemKey,
            ([string]$p.Length),
            (ConvertTo-QCTimestamp $p.LastWriteTimeUtc),
            ([string]$c.Length),
            (ConvertTo-QCTimestamp $c.LastWriteTimeUtc),
            ([string]$dir).ToLowerInvariant()
        ) -join '|')
    }

    $lines = @($lines | Sort-Object)
    $stable = ($lines -join "`n")
    # Use text hash, not file hash
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($stable)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $h = $sha.ComputeHash($bytes) } finally { $sha.Dispose() }
    $hash = ([BitConverter]::ToString($h) -replace '-', '').ToLowerInvariant()

    $pairedSheets = @($pairedSheets | Sort-Object -Property @{ Expression = { $_.dir } }, @{ Expression = { $_.stem } })
    $orderKey = (@($pairedSheets | ForEach-Object { ($_.dir + '|' + $_.stem) }) -join "`n")

    return @{
        folderStateHash = $hash
        pairedCount = $paired
        stableInput = $stable
        orderKey = $orderKey
        pairedSheets = $pairedSheets
    }
}

function Get-StatusSetPWFolderState {
    <#
    .SYNOPSIS
    Watcher helper: stable fingerprint for PW folder using legacy pairing (PDF requires matching DGN/DWG base).
    .PARAMETER FolderPath
    PW folder path WITHOUT leading Documents\ (same convention as other watcher PW calls).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath,
        [Parameter(Mandatory)]
        [bool]$OneLevelDeep
    )

    $paths = @([string]$FolderPath)
    if ($OneLevelDeep) {
        try {
            $kids = @(Get-PWImmediateChildFolders -FolderPath $FolderPath)
            foreach ($k in $kids) {
                $kp = _SSS-PWGetProp -Obj $k -Name 'FolderPath'
                if ($kp) { $paths += [string]$kp }
            }
        } catch { }
    }

    $dateCols = @('Name','DocumentID','FileUpdatedDate','FileUpdateDate','DocumentUpdateDate','VersionModifiedDate','Version Modified Date','FileSize','Size')
    $docs = @()
    foreach ($p in @($paths | Select-Object -Unique)) {
        $docs += @(_SSS-PWListDocsInFolder -FolderPath $p -DateCols $dateCols)
    }

    $hasCad = @{}
    $pdfDocs = @()
    $nonCadExt = '\.(pdf|xlsx|xls|doc|docx|txt|zip|jpg|jpeg|png|gif|bmp|tif|tiff|log|xml|json|csv)$'
    foreach ($doc in @($docs)) {
        $name = _SSS-GetDocName $doc
        if (-not $name) { continue }
        if ($name -match '\.pdf$') { continue }
        if ($name -match $nonCadExt) { continue }
        if ($name -match '\.(dgn|dwg)$') { $base = ($name -replace '\.(dgn|dwg)$','').ToLowerInvariant() } else { $base = $name.ToLowerInvariant() }
        $hasCad[$base] = $true
    }
    foreach ($doc in @($docs)) {
        $name = _SSS-GetDocName $doc
        if (-not $name -or $name -notmatch '\.pdf$') { continue }
        if ($name -match '-qc\.pdf$') { continue }
        if ($name -match '(?i)_statusset\.pdf$') { continue }
        $base = ($name -replace '\.pdf$','').ToLowerInvariant()
        if (-not $hasCad[$base]) { continue }
        $pdfDocs += $doc
    }
    $pdfDocs = @($pdfDocs | Sort-Object { _SSS-GetDocName $_ })

    $lines = @()
    $pairedSheets = @()
    foreach ($p in $pdfDocs) {
        $pn = _SSS-GetDocName $p
        $stemKey = ([System.IO.Path]::GetFileNameWithoutExtension($pn)).ToLowerInvariant()
        $mod = _SSS-GetDocLastModified $p
        $modIso = if ($mod) { ConvertTo-QCTimestamp $mod } else { '' }
        $sz = _SSS-PWGetProp -Obj $p -Name 'FileSize'
        if (-not $sz) { $sz = _SSS-PWGetProp -Obj $p -Name 'Size' }
        $sz = _SSS-NormalizeFileSize $sz
        $dirKey = ([string]$FolderPath).ToLowerInvariant()
        $pairedSheets += @{ stem = $stemKey; dir = $dirKey; pdf = @{ name=$pn; lastWriteTimeUtc=$modIso; length=$sz } }
        $lines += (@($stemKey, ([string]$sz), $modIso, $dirKey) -join '|')
    }

    $lines = @($lines | Sort-Object)
    $stable = ($lines -join "`n")
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($stable)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $h = $sha.ComputeHash($bytes) } finally { $sha.Dispose() }
    $hash = ([BitConverter]::ToString($h) -replace '-', '').ToLowerInvariant()
    $pairedSheets = @($pairedSheets | Sort-Object -Property @{ Expression = { $_.dir } }, @{ Expression = { $_.stem } })
    $orderKey = (@($pairedSheets | ForEach-Object { ($_.dir + '|' + $_.stem) }) -join "`n")

    return @{
        folderStateHash = $hash
        pairedCount = $pdfDocs.Count
        pdfCount = $pdfDocs.Count
        stableInput = $stable
        orderKey = $orderKey
        pairedSheets = $pairedSheets
    }
}

function _SSS-FindStatusSetDocLegacy([string[]]$TryPaths, [string]$OutputName, [string[]]$DateCols) {
    foreach ($tp in $TryPaths) {
        try {
            $doc = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $tp -JustThisFolder -DocumentName $OutputName -ColumnsToReturn $DateCols -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($doc) { return $doc }
        } catch { }
        try {
            $doc2 = Get-PWDocumentsBySearch -FolderPath $tp -JustThisFolder -DocumentName $OutputName -PopulatePath -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($doc2) { return $doc2 }
        } catch { }
    }
    return $null
}

function Invoke-StatusSetNativeJob {
    <#
    .SYNOPSIS
    STATUS_SET_GEN implementation matching legacy/combine_status_set.ps1 method (one-shot).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $ss = @{}
    if ($Config.ContainsKey('statusSet') -and $Config.statusSet) {
        $raw = $Config.statusSet
        if ($raw -is [hashtable]) { $ss = $raw } elseif ($raw.PSObject) { foreach ($p in $raw.PSObject.Properties) { $ss[$p.Name] = $p.Value } }
    }

    $localRoot = if ($ss.ContainsKey('localRoot') -and $ss.localRoot) { [string]$ss.localRoot } else { 'C:\PW_QC_LOCAL' }
    $qpdfExe = if ($ss.ContainsKey('qpdfExe') -and $ss.qpdfExe) { [string]$ss.qpdfExe } else { '' }
    if (_SSS-IsNullOrWhiteSpace $qpdfExe) {
        if ($Config.ContainsKey('qcPrepend') -and $Config.qcPrepend) {
            $qc = $Config.qcPrepend
            if ($qc -is [hashtable] -and $qc.ContainsKey('qpdfExePath')) { $qpdfExe = [string]$qc['qpdfExePath'] }
            elseif ($qc.PSObject.Properties['qpdfExePath']) { $qpdfExe = [string]$qc.qpdfExePath }
        }
    }
    $repoRoot = Split-Path -Parent $PSScriptRoot
    if (-not (_SSS-IsNullOrWhiteSpace $qpdfExe)) {
        try {
            if (-not [System.IO.Path]::IsPathRooted($qpdfExe)) {
                $qpdfExe = Join-Path $repoRoot $qpdfExe
            }
        } catch { }
    }
    if (_SSS-IsNullOrWhiteSpace $qpdfExe) { $qpdfExe = Join-Path $repoRoot 'tools\qpdf\bin\qpdf.exe' }
    _SSS-AssertCommand $qpdfExe

    $forceRebuild = $false
    if ($ss.ContainsKey('forceRebuild')) { try { $forceRebuild = [bool]$ss.forceRebuild } catch { $forceRebuild = $false } }
    $writeBackToPW = $false
    if ($ss.ContainsKey('writeBackToPW')) { try { $writeBackToPW = [bool]$ss.writeBackToPW } catch { $writeBackToPW = $false } }

    $sourceFolder = if ($Job.ContainsKey('sourceFolder') -and $Job.sourceFolder) { [string]$Job.sourceFolder } else { '' }
    if (_SSS-IsNullOrWhiteSpace $sourceFolder) {
        return New-QCFailureResult -Code 'STATUS_SET_MISSING_SOURCE_FOLDER' -Message 'Job.sourceFolder is required.' -Data @{ jobId = [string]$Job.id }
    }

    $pwCfg = @{}
    if ($Config.ContainsKey('projectWise')) {
        $p = $Config.projectWise
        if ($p -is [hashtable]) { $pwCfg = $p } elseif ($p.PSObject) { foreach ($x in $p.PSObject.Properties) { $pwCfg[$x.Name] = $x.Value } }
    }
    $credPath = if ($pwCfg.ContainsKey('credentialPath') -and $pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
    $ds = if ($pwCfg.ContainsKey('datasourceName') -and $pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { 'typsa-us-pw.bentley.com:typsa-us-pw-03' }
    if ($Job.ContainsKey('metadata') -and $Job.metadata -is [hashtable] -and $Job.metadata.ContainsKey('candidate')) {
        $cand = $Job.metadata.candidate
        if ($cand -is [hashtable] -and $cand.ContainsKey('datasourceName') -and $cand.datasourceName) { $ds = [string]$cand.datasourceName }
    }

    $isDryRun = $false
    if ($Config.ContainsKey('dryRun')) { try { $isDryRun = [bool]$Config.dryRun } catch { $isDryRun = $false } }

    # Local-folder mode is supported (same merge + manifest, but without PW export).
    $normRes = Normalize-QCPath -Path $sourceFolder
    $norm = if ($normRes.IsSuccess) { [string]$normRes.Data.path } else { $sourceFolder }
    $isPwLogical = $norm.StartsWith('documents\\', [System.StringComparison]::OrdinalIgnoreCase) -or ($sourceFolder -match '(?i)^pw:\\\\')
    $useLocalFs = (-not $isPwLogical) -and (Test-Path -LiteralPath $sourceFolder -PathType Container)

    $pwSheetsPath = if ($useLocalFs) { $sourceFolder } else { _SSS-GetPwFolderPath $sourceFolder }
    $tryPaths = if ($useLocalFs) { @() } else { _SSS-GetPwSheetsTryPaths $sourceFolder }

    $manifestPath = Get-StatusSetManifestPathLegacy -FolderPath $pwSheetsPath -LocalRoot $localRoot
    $cacheDir = Get-StatusSetCacheDirLegacy -FolderPath $pwSheetsPath -LocalRoot $localRoot
    _SSS-EnsureDir $cacheDir

    # Per-folder + per-job temp dir so concurrent STATUS_SET_GEN workers never
    # collide on the same _StatusSet.pdf or status_set_replace_*.pdf.
    $tempSafe = (($pwSheetsPath -replace '[\\/:]', '_').Trim())
    if (-not $tempSafe) { $tempSafe = 'default' }
    $jobIdSafe = [string]$Job.id
    if (-not $jobIdSafe) { $jobIdSafe = [guid]::NewGuid().ToString('N') }
    $tempWorkDir = Join-Path $env:TEMP ("PW_QC_StatusSet\{0}_{1}" -f $tempSafe, $jobIdSafe)
    _SSS-EnsureDir $tempWorkDir
    $outPdf = Join-Path $tempWorkDir $script:StatusSetOutputName

    $manifest = Read-StatusSetManifestLegacy -ManifestPath $manifestPath
    $manifestSchemaStale = $false
    if ($manifest) {
        try {
            $mv = $manifest.manifestSchemaVersion
            if ($null -eq $mv -or [int]$mv -lt $script:StatusSetManifestSchemaVersion) {
                $manifestSchemaStale = $true
                $manifest = $null
            }
        } catch { $manifest = $null; $manifestSchemaStale = $true }
    }
    if ($forceRebuild) { $manifest = $null }

    if ($useLocalFs) {
        # Local: use all *.pdf not -qc + has matching *.dgn/*.dwg base (same pairing semantics).
        $allFiles = @(Get-ChildItem -LiteralPath $sourceFolder -File -ErrorAction SilentlyContinue)
        $hasCad = @{}
        foreach ($f in $allFiles) {
            $n = [string]$f.Name
            if (-not $n) { continue }
            if ($n -match '\.pdf$') { continue }
            if ($n -match '\.(dgn|dwg)$') { $base = ($n -replace '\.(dgn|dwg)$','').ToLowerInvariant() } else { $base = ($n -replace '\.[^.]+$','').ToLowerInvariant() }
            $hasCad[$base] = $true
        }
        $pdfs = @()
        foreach ($f in $allFiles) {
            $n = [string]$f.Name
            if ($n -notmatch '\.pdf$') { continue }
            if ($n -match '-qc\.pdf$') { continue }
            $base = ($n -replace '\.pdf$','').ToLowerInvariant()
            if (-not $hasCad[$base]) { continue }
            $pdfs += $f.FullName
        }
        $localPdfPaths = @($pdfs | Sort-Object { [System.IO.Path]::GetFileName($_) })
        if ($localPdfPaths.Count -eq 0) {
            return New-QCFailureResult -Code 'STATUS_SET_NO_PAIRS' -Message 'No matching PDF+DGN/DWG pairs found on disk.' -Data @{ folder = $sourceFolder }
        }

        $needsFullRebuild = $true
        $changed = @()
        $expectedTotalPages = 0
        foreach ($lp in $localPdfPaths) { $expectedTotalPages += _SSS-GetPdfPageCount -Path $lp -QpdfExe $qpdfExe }

        if (-not $forceRebuild -and $manifest -and $manifest.sources) {
            try {
                $manifestSources = @($manifest.sources)
                if ($manifestSources.Count -eq $localPdfPaths.Count) {
                    $needsFullRebuild = $false
                    $pageEnd = 0
                    for ($i = 0; $i -lt $localPdfPaths.Count; $i++) {
                        $localPath = $localPdfPaths[$i]
                        $name = [System.IO.Path]::GetFileName($localPath)
                        $hash = _SSS-GetFileHashSha256 $localPath
                        $pageCount = _SSS-GetPdfPageCount -Path $localPath -QpdfExe $qpdfExe
                        $pageStart = $pageEnd + 1
                        $pageEnd = $pageStart + $pageCount - 1
                        $old = $manifestSources | Where-Object { $_.name -eq $name } | Select-Object -First 1
                        if (-not $old -or $old.hash -ne $hash) { $changed += @{ Index=$i; Name=$name; PageStart=$pageStart; PageEnd=$pageEnd; LocalPath=$localPath } }
                    }
                }
            } catch { $needsFullRebuild = $true }
        }

        if ($isDryRun) {
            return New-QCSuccessResult -Code 'STATUS_SET_DRYRUN' -Message 'Dry-run: would build StatusSet (legacy method, local filesystem).' -Data @{ jobId=[string]$Job.id; outPdf=$outPdf; manifestPath=$manifestPath; cacheDir=$cacheDir; needsFullRebuild=$needsFullRebuild; changedCount=$changed.Count }
        }

        if ($needsFullRebuild) {
            _SSS-MergePdfs -PdfPaths $localPdfPaths -OutPath $outPdf -QpdfExe $qpdfExe
        } elseif ($changed.Count -gt 0 -and (Test-Path -LiteralPath $outPdf)) {
            $actualPages = _SSS-GetPdfPageCount -Path $outPdf -QpdfExe $qpdfExe
            if ($actualPages -ne $expectedTotalPages) {
                _SSS-MergePdfs -PdfPaths $localPdfPaths -OutPath $outPdf -QpdfExe $qpdfExe
                $needsFullRebuild = $true
            } else {
                $sorted = $changed | Sort-Object { $_.PageStart } -Descending
                $workPdf = $outPdf
                foreach ($ch in $sorted) {
                    $tmpOut = Join-Path $tempWorkDir ("status_set_replace_{0}.pdf" -f [guid]::NewGuid().ToString('N').Substring(0,8))
                    _SSS-ReplacePdfPages -CombinedPath $workPdf -NewPdfPath $ch.LocalPath -PageStart $ch.PageStart -PageEnd $ch.PageEnd -OutPath $tmpOut -QpdfExe $qpdfExe
                    if ($workPdf -ne $outPdf) { Remove-Item -LiteralPath $workPdf -Force -ErrorAction SilentlyContinue }
                    $workPdf = $tmpOut
                }
                if ($workPdf -ne $outPdf) { Move-Item -LiteralPath $workPdf -Destination $outPdf -Force }
            }
        } else {
            if (-not (Test-Path -LiteralPath $outPdf)) { _SSS-MergePdfs -PdfPaths $localPdfPaths -OutPath $outPdf -QpdfExe $qpdfExe }
        }

        # manifest update
        if ($needsFullRebuild -or $changed.Count -gt 0 -or -not $manifest) {
            $sources = @()
            $pageEnd = 0
            foreach ($lp in $localPdfPaths) {
                $name = [System.IO.Path]::GetFileName($lp)
                $hash = _SSS-GetFileHashSha256 $lp
                $pageCount = _SSS-GetPdfPageCount -Path $lp -QpdfExe $qpdfExe
                $pageStart = $pageEnd + 1
                $pageEnd = $pageStart + $pageCount - 1
                $sources += @{ name = $name; hash = $hash; pageStart = $pageStart; pageEnd = $pageEnd }
            }
            $manifestObj = @{ manifestSchemaVersion = $script:StatusSetManifestSchemaVersion; folderPath = $pwSheetsPath; pinnedStatusSetLastModified = $null; sources = $sources }
            Write-StatusSetManifestLegacy -ManifestPath $manifestPath -ManifestObj $manifestObj
        }

        return New-QCSuccessResult -Code 'STATUS_SET_OK' -Message 'Status set generated (legacy method, local filesystem).' -Data @{ jobId=[string]$Job.id; outPdf=$outPdf; manifestPath=$manifestPath; cacheDir=$cacheDir; writeBackToPW=$false }
    }

    # --- ProjectWise legacy-method path ---
    Import-Module (Join-Path $PSScriptRoot 'PW.Connection.psm1') -Force
    $credRes = Get-PWCredentialFromFile -CredentialPath $credPath
    if (-not $credRes.IsSuccess) { return $credRes }
    $connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
    if (-not $connRes.IsSuccess) { return $connRes }

    try {
        $dateCols = @('Name','DocumentID','FileUpdatedDate','FileUpdateDate','DocumentUpdateDate','VersionModifiedDate','Version Modified Date')
        $allDocs = @()
        $docSearchPath = $pwSheetsPath
        foreach ($tp in $tryPaths) {
            $allDocs = @(_SSS-PWListDocsInFolder -FolderPath $tp -DateCols $dateCols)
            if ($allDocs.Count -gt 0) { $docSearchPath = $tp; break }
        }
        if ($allDocs.Count -eq 0) {
            return New-QCFailureResult -Code 'STATUS_SET_PW_NO_DOCS' -Message 'No documents returned from PW for sheets folder.' -Data @{ folder = $sourceFolder; tryPaths = $tryPaths }
        }

        $pdfDocs = @(_SSS-SelectPairedPdfDocsLegacy -AllDocs $allDocs)
        if ($pdfDocs.Count -eq 0) {
            return New-QCFailureResult -Code 'STATUS_SET_NO_PAIRS' -Message 'No matching sheet PDFs (need paired DGN/DWG base name).' -Data @{ folder = $sourceFolder; docCount = $allDocs.Count }
        }

        # cutoff logic (PW StatusSet last modified or pinned manifest)
        $statusSetDateCols = @($dateCols + @('FileUpdatedDate','FileUpdateDate')) | Select-Object -Unique
        $statusSetDoc = _SSS-FindStatusSetDocLegacy -TryPaths $tryPaths -OutputName $script:StatusSetOutputName -DateCols $statusSetDateCols
        $statusSetLastModified = if ($statusSetDoc) { _SSS-GetDocLastModified $statusSetDoc } else { $null }
        $manifestPinnedCutoff = $null
        if ($manifest -and $manifest.pinnedStatusSetLastModified) { $manifestPinnedCutoff = _SSS-ParseIsoDateTime ([string]$manifest.pinnedStatusSetLastModified) }
        $exportCutoff = if ($statusSetLastModified) { $statusSetLastModified } else { $manifestPinnedCutoff }

        # export loop with cache + manifest skip rules
        $localPdfPaths = @()
        $sourcePwLastModByName = @{}
        foreach ($doc in $pdfDocs) {
            $docName = _SSS-GetDocName $doc
            $cachedPath = Join-Path $cacheDir $docName
            $docLastMod = _SSS-GetDocLastModified $doc
            $sourcePwLastModByName[$docName] = $docLastMod

            $shouldExport = $true
            if ($exportCutoff -and $docLastMod) {
                if ($docLastMod -le $exportCutoff -and (Test-Path -LiteralPath $cachedPath)) { $shouldExport = $false }
            }
            if ($shouldExport -and $manifest -and $manifest.sources) {
                $oldSrc = @($manifest.sources) | Where-Object { $_.name -eq $docName } | Select-Object -First 1
                if ($oldSrc) {
                    if ($oldSrc.pwLastModified) {
                        $stored = _SSS-ParseIsoDateTime ([string]$oldSrc.pwLastModified)
                        if ($stored -and $docLastMod -and ($docLastMod -le $stored) -and (Test-Path -LiteralPath $cachedPath)) { $shouldExport = $false }
                    }
                    if ($shouldExport -and $oldSrc.hash -and (Test-Path -LiteralPath $cachedPath)) {
                        $h = _SSS-GetFileHashSha256 $cachedPath
                        if ($h -eq $oldSrc.hash) {
                            $stored = $null
                            if ($oldSrc.pwLastModified) { $stored = _SSS-ParseIsoDateTime ([string]$oldSrc.pwLastModified) }
                            if ($docLastMod -and $stored -and ($docLastMod -gt $stored)) {
                                # refresh from PW
                            } else {
                                $shouldExport = $false
                            }
                        }
                    }
                }
            }

            if ($shouldExport) {
                if ($isDryRun) {
                    # no-op
                } else {
                    Export-PWDocumentsSimple -InputDocuments $doc -TargetFolder $cacheDir -ErrorAction Stop | Out-Null
                    Start-Sleep -Milliseconds 400
                }
            }

            $localPath = $cachedPath
            $copied = _SSS-PWGetProp -Obj $doc -Name 'CopiedOutLocalFileName'
            if (-not (Test-Path -LiteralPath $localPath) -and $copied -and (Test-Path -LiteralPath $copied)) {
                Copy-Item -LiteralPath $copied -Destination $localPath -Force -ErrorAction SilentlyContinue
            }
            if (-not (Test-Path -LiteralPath $localPath)) {
                $found = Get-ChildItem -LiteralPath $cacheDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq $docName } | Select-Object -First 1
                if ($found) {
                    $localPath = [string]$found.FullName
                    if ($localPath -ne $cachedPath) { Copy-Item -LiteralPath $localPath -Destination $cachedPath -Force -ErrorAction SilentlyContinue; $localPath = $cachedPath }
                }
            }
            if (Test-Path -LiteralPath $localPath) { $localPdfPaths += $localPath }
        }

        if ($localPdfPaths.Count -eq 0) {
            return New-QCFailureResult -Code 'STATUS_SET_EXPORT_FAILED' -Message 'No PDFs were exported successfully.' -Data @{ cacheDir = $cacheDir }
        }

        $expectedTotalPages = 0
        foreach ($lp in $localPdfPaths) { $expectedTotalPages += _SSS-GetPdfPageCount -Path $lp -QpdfExe $qpdfExe }

        $needsFullRebuild = $forceRebuild
        $changedIndices = @()

        if (-not $needsFullRebuild -and (Test-Path -LiteralPath $manifestPath) -and -not $manifestSchemaStale) {
            try {
                $manifest2 = Read-StatusSetManifestLegacy -ManifestPath $manifestPath
                $manifestSources = @($manifest2.sources)
                if ($manifestSources.Count -ne $localPdfPaths.Count) { $needsFullRebuild = $true }
                else {
                    $pageEnd = 0
                    for ($i = 0; $i -lt $localPdfPaths.Count; $i++) {
                        $localPath = $localPdfPaths[$i]
                        $name = [System.IO.Path]::GetFileName($localPath)
                        $hash = _SSS-GetFileHashSha256 $localPath
                        $pageCount = _SSS-GetPdfPageCount -Path $localPath -QpdfExe $qpdfExe
                        $pageStart = $pageEnd + 1
                        $pageEnd = $pageStart + $pageCount - 1
                        $old = $manifestSources | Where-Object { $_.name -eq $name } | Select-Object -First 1
                        if (-not $old -or $old.hash -ne $hash) { $changedIndices += @{ Index=$i; Name=$name; PageStart=$pageStart; PageEnd=$pageEnd; LocalPath=$localPath } }
                    }
                }
            } catch { $needsFullRebuild = $true }
        } else { $needsFullRebuild = $true }

        if ($isDryRun) {
            return New-QCSuccessResult -Code 'STATUS_SET_DRYRUN' -Message 'Dry-run: would build StatusSet (legacy method, ProjectWise).' -Data @{
                jobId = [string]$Job.id
                manifestPath = $manifestPath
                cacheDir = $cacheDir
                docSearchPath = $docSearchPath
                outPdf = $outPdf
                needsFullRebuild = $needsFullRebuild
                changedCount = $changedIndices.Count
                wouldWriteBack = $writeBackToPW
            }
        }

        if ($needsFullRebuild) {
            _SSS-MergePdfs -PdfPaths $localPdfPaths -OutPath $outPdf -QpdfExe $qpdfExe
        } elseif ($changedIndices.Count -gt 0 -and (Test-Path -LiteralPath $outPdf)) {
            $actualPages = _SSS-GetPdfPageCount -Path $outPdf -QpdfExe $qpdfExe
            if ($actualPages -ne $expectedTotalPages) {
                _SSS-MergePdfs -PdfPaths $localPdfPaths -OutPath $outPdf -QpdfExe $qpdfExe
                $needsFullRebuild = $true
            } else {
                $sorted = $changedIndices | Sort-Object { $_.PageStart } -Descending
                $workPdf = $outPdf
                foreach ($ch in $sorted) {
                    $tmpOut = Join-Path $tempWorkDir ("status_set_replace_{0}.pdf" -f [guid]::NewGuid().ToString('N').Substring(0,8))
                    _SSS-ReplacePdfPages -CombinedPath $workPdf -NewPdfPath $ch.LocalPath -PageStart $ch.PageStart -PageEnd $ch.PageEnd -OutPath $tmpOut -QpdfExe $qpdfExe
                    if ($workPdf -ne $outPdf) { Remove-Item -LiteralPath $workPdf -Force -ErrorAction SilentlyContinue }
                    $workPdf = $tmpOut
                }
                if ($workPdf -ne $outPdf) { Move-Item -LiteralPath $workPdf -Destination $outPdf -Force }
            }
        } else {
            if (-not (Test-Path -LiteralPath $outPdf)) { _SSS-MergePdfs -PdfPaths $localPdfPaths -OutPath $outPdf -QpdfExe $qpdfExe }
        }

        # update manifest after rebuild/exchange
        if ($needsFullRebuild -or $changedIndices.Count -gt 0) {
            $sources = @()
            $pageEnd = 0
            foreach ($lp in $localPdfPaths) {
                $name = [System.IO.Path]::GetFileName($lp)
                $hash = _SSS-GetFileHashSha256 $lp
                $pageCount = _SSS-GetPdfPageCount -Path $lp -QpdfExe $qpdfExe
                $pageStart = $pageEnd + 1
                $pageEnd = $pageStart + $pageCount - 1
                $row = @{ name = $name; hash = $hash; pageStart = $pageStart; pageEnd = $pageEnd }
                $dlm = $sourcePwLastModByName[$name]
                if ($dlm) { $row.pwLastModified = ConvertTo-QCTimestamp $dlm }
                $sources += $row
            }
            $pinned = $null
            if ($statusSetLastModified) { $pinned = ConvertTo-QCTimestamp $statusSetLastModified }
            elseif ($manifest2 -and $manifest2.pinnedStatusSetLastModified) { $pinned = [string]$manifest2.pinnedStatusSetLastModified }
            $manifestObj = @{
                manifestSchemaVersion       = $script:StatusSetManifestSchemaVersion
                folderPath                  = $pwSheetsPath
                pinnedStatusSetLastModified = $pinned
                sources                     = $sources
            }
            Write-StatusSetManifestLegacy -ManifestPath $manifestPath -ManifestObj $manifestObj
        }

        # write-back if requested (same decision as legacy)
        $existingDoc = Get-PWDocumentsBySearch -FolderPath $docSearchPath -JustThisFolder -DocumentName $script:StatusSetOutputName -PopulatePath -ErrorAction SilentlyContinue
        $shouldWriteBack = $writeBackToPW -and (Test-Path -LiteralPath $outPdf) -and (($needsFullRebuild -or $changedIndices.Count -gt 0) -or -not $existingDoc)
        $pwUpload = $null
        $pwUploadError = $null
        if ($shouldWriteBack) {
            try {
                if ($existingDoc) {
                    _SSS-UpdatePWDocumentFileFromDisk -InputDocument $existingDoc -LocalFilePath $outPdf
                    $pwUpload = 'UPDATED'
                } else {
                    New-PWDocument -FolderPath $docSearchPath -FilePath $outPdf -DocumentName $script:StatusSetOutputName -ErrorAction Stop | Out-Null
                    $pwUpload = 'CREATED'
                }
                # refresh pinned date after write-back
                try {
                    $sd2 = _SSS-FindStatusSetDocLegacy -TryPaths $tryPaths -OutputName $script:StatusSetOutputName -DateCols $statusSetDateCols
                    if ($sd2) {
                        $dt2 = _SSS-GetDocLastModified $sd2
                        if ($dt2 -and (Test-Path -LiteralPath $manifestPath)) {
                            $m2 = Read-StatusSetManifestLegacy -ManifestPath $manifestPath
                            if ($m2) {
                                $m2.pinnedStatusSetLastModified = ConvertTo-QCTimestamp $dt2
                                $m2 | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
                            }
                        }
                    }
                } catch { }
            } catch {
                $pwUpload = 'FAILED'
                $pwUploadError = [string]$_.Exception.Message
            }
        }

        # If writeback was requested AND attempted AND threw, do NOT report success.
        # The local PDF was generated but the user's intent ("upload to PW") was not
        # met; treating this as success would silently strand the result and the job
        # would never be retried. Return a structured failure so the worker can move
        # the job to pending\ for another attempt (or to failed\ after maxAttempts).
        if ($pwUpload -eq 'FAILED') {
            return New-QCFailureResult -Code 'STATUS_SET_PW_UPLOAD_FAILED' -Message ('PW write-back failed: ' + $pwUploadError) -Data @{
                jobId            = [string]$Job.id
                outPdf           = $outPdf
                manifestPath     = $manifestPath
                cacheDir         = $cacheDir
                docSearchPath    = $docSearchPath
                needsFullRebuild = $needsFullRebuild
                changedCount     = $changedIndices.Count
                writeBackToPW    = $writeBackToPW
                pwUpload         = $pwUpload
                error            = $pwUploadError
            }
        }

        return New-QCSuccessResult -Code 'STATUS_SET_OK' -Message 'Status set generated (legacy method, ProjectWise).' -Data @{
            jobId = [string]$Job.id
            outPdf = $outPdf
            manifestPath = $manifestPath
            cacheDir = $cacheDir
            docSearchPath = $docSearchPath
            needsFullRebuild = $needsFullRebuild
            changedCount = $changedIndices.Count
            writeBackToPW = $writeBackToPW
            pwUpload = $pwUpload
        }
    } finally {
        try { Disconnect-PW | Out-Null } catch { }
    }
}

Export-ModuleMember -Function @(
    'Get-StatusSetManifestPathLegacy',
    'Get-StatusSetCacheDirLegacy',
    'Read-StatusSetManifestLegacy',
    'Write-StatusSetManifestLegacy',
    'Get-StatusSetLocalFolderState',
    'Get-StatusSetPWFolderState',
    'Invoke-StatusSetNativeJob'
)

function _SSS-BuildPWStatusSetState {
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath,
        [Parameter(Mandatory)]
        [bool]$OneLevelDeep,
        [Parameter(Mandatory)]
        [bool]$IncludeOrderedPdfDocuments
    )

    $paths = @([string]$FolderPath)
    if ($OneLevelDeep) {
        try {
            $kids = @(Get-PWImmediateChildFolders -FolderPath $FolderPath)
            foreach ($k in $kids) {
                $kp = _SSS-PWGetProp -Obj $k -Name 'FolderPath'
                if ($kp) { $paths += [string]$kp }
            }
        } catch { }
    }

    $docs = @()
    $dateCols = @('Name','DocumentID','DocumentGUID','FileUpdatedDate','FileUpdateDate','DocumentUpdateDate','VersionModifiedDate','Version Modified Date','FileSize','Size','StateName')
    foreach ($p in @($paths | Select-Object -Unique)) {
        # Targeted query: only the extensions that can affect pairing/fingerprint.
        $docs += @(_SSS-PWListDocsInFolderByExtensions -FolderPath $p -Extensions @('.pdf','.dgn','.dwg') -DateCols $dateCols)
    }
    $pdfs = @()
    $dgns = @()
    foreach ($d in @($docs)) {
        $name = _SSS-PWGetDocName -Doc $d
        if (-not $name) { continue }
        if ($name -match '(?i)-qc\.pdf$') { continue }
        if ($name -match '(?i)_statusset\.pdf$') { continue }
        if ($name -match '(?i)\.pdf$') { $pdfs += $d; continue }
        if ($name -match '(?i)\.dgn$') { $dgns += $d; continue }
    }

    $dgnByStem = @{}
    foreach ($d in $dgns) {
        $n = _SSS-PWGetDocName -Doc $d
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($n)
        if ($stem) { $dgnByStem[$stem.ToLowerInvariant()] = $d }
    }

    $lines = @()
    $pairRows = @()
    foreach ($p in $pdfs) {
        $pn = _SSS-PWGetDocName -Doc $p
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($pn)
        if (-not $stem) { continue }
        $stemKey = $stem.ToLowerInvariant()
        if (-not $dgnByStem.ContainsKey($stemKey)) { continue }
        $d = $dgnByStem[$stemKey]
        $dgnName = _SSS-PWGetDocName -Doc $d

        $pdfMod = _SSS-PWGetDocLastModifiedUtcIso -Doc $p
        $dgnMod = _SSS-PWGetDocLastModifiedUtcIso -Doc $d
        $pdfSize = _SSS-PWGetProp -Obj $p -Name 'FileSize'
        if (-not $pdfSize) { $pdfSize = _SSS-PWGetProp -Obj $p -Name 'Size' }
        $dgnSize = _SSS-PWGetProp -Obj $d -Name 'FileSize'
        if (-not $dgnSize) { $dgnSize = _SSS-PWGetProp -Obj $d -Name 'Size' }
        $pdfDocId = _SSS-PWGetProp -Obj $p -Name 'DocumentID'
        $dgnDocId = _SSS-PWGetProp -Obj $d -Name 'DocumentID'
        $pdfDocGuid = _SSS-PWGetProp -Obj $p -Name 'DocumentGUID'
        $dgnDocGuid = _SSS-PWGetProp -Obj $d -Name 'DocumentGUID'
        $pdfStateName = _SSS-PWGetProp -Obj $p -Name 'StateName'
        $dgnStateName = _SSS-PWGetProp -Obj $d -Name 'StateName'

        $dirKey = ([string]$FolderPath).ToLowerInvariant()
        $pairRows += [pscustomobject]@{
            Stem = $stemKey
            Dir = $dirKey
            PdfDoc = $p
            DgnDoc = $d
            PdfName = $pn
            PdfMod = $pdfMod
            PdfSize = $pdfSize
            DgnMod = $dgnMod
            DgnSize = $dgnSize
            PdfDocumentId = $pdfDocId
            DgnDocumentId = $dgnDocId
            PdfDocumentGuid = $pdfDocGuid
            DgnDocumentGuid = $dgnDocGuid
            PdfStateName = $pdfStateName
            DgnStateName = $dgnStateName
            DgnName = $dgnName
        }
        # Hash identity for dedupe + manifest:
        # - Include PW document IDs (stable identity) and file sizes (when available)
        # - Include FileUpdateDateUtc (via _SSS-PWGetDocLastModifiedUtcIso) so PW "reprints"/new versions
        #   trigger rebuilds even when the PDF size does not change.
        $pdfIdStr = if ($null -ne $pdfDocId -and -not [string]::IsNullOrWhiteSpace([string]$pdfDocId)) { [string]$pdfDocId } else { '' }
        $dgnIdStr = if ($null -ne $dgnDocId -and -not [string]::IsNullOrWhiteSpace([string]$dgnDocId)) { [string]$dgnDocId } else { '' }
        $lines += (@(
            $stemKey,
            ([string]$pdfSize),
            ([string]$dgnSize),
            $pdfIdStr,
            $dgnIdStr,
            ([string]$pdfMod),
            ([string]$dgnMod),
            $dirKey
        ) -join '|')
    }

    $lines = @($lines | Sort-Object)
    $stable = ($lines -join "`n")
    $hash = _SSS-Sha256TextHex -Text $stable

    $sortedRows = @($pairRows | Sort-Object -Property Dir, Stem)
    # Must be @(...) — a lone foreach result is one hashtable; piping that hashtable
    # iterates values (not rows) and breaks orderKey / watcher QC-PDF linking.
    $pairedSheets = @(foreach ($r in $sortedRows) {
        @{
            stem = $r.Stem
            dir = $r.Dir
            pdf = @{
                name = [string]$r.PdfName
                lastWriteTimeUtc = [string]$r.PdfMod
                length = $r.PdfSize
                documentId = $r.PdfDocumentId
                documentGuid = $r.PdfDocumentGuid
                stateName = $r.PdfStateName
                doc = $r.PdfDoc
            }
            dgn = @{
                name = [string]$r.DgnName
                lastWriteTimeUtc = [string]$r.DgnMod
                length = $r.DgnSize
                documentId = $r.DgnDocumentId
                documentGuid = $r.DgnDocumentGuid
                stateName = $r.DgnStateName
                doc = $r.DgnDoc
            }
        }
    })
    $orderKey = (@($pairedSheets | ForEach-Object { ($_.dir + '|' + $_.stem) }) -join "`n")
    $orderedPdfDocs = if ($IncludeOrderedPdfDocuments) { @($sortedRows | ForEach-Object { $_.PdfDoc }) } else { @() }

    # Collect QC PDFs (filtered out from pairing) so the caller can link them
    $qcPdfDocs = @()
    foreach ($d in @($docs)) {
        $dn = _SSS-PWGetDocName -Doc $d
        if ($dn -match '(?i)-qc\.pdf$') {
            $qcPdfDocs += @{
                name = [string]$dn
                documentId = _SSS-PWGetProp -Obj $d -Name 'DocumentID'
                documentGuid = _SSS-PWGetProp -Obj $d -Name 'DocumentGUID'
                stem = ([System.IO.Path]::GetFileNameWithoutExtension($dn) -replace '(?i)-qc$', '').ToLowerInvariant()
            }
        }
    }

    $listingMethod = 'unknown'
    try { if ($script:_SSS_LastDocListingMethod) { $listingMethod = [string]$script:_SSS_LastDocListingMethod } } catch { }

    return @{
        folderStateHash = $hash
        pairedCount = $sortedRows.Count
        pdfCount = $pdfs.Count
        dgnCount = $dgns.Count
        stableInput = $stable
        orderKey = $orderKey
        pairedSheets = $pairedSheets
        orderedPdfDocuments = $orderedPdfDocs
        qcPdfDocs = $qcPdfDocs
        docListingMethod = $listingMethod
    }
}

function Get-StatusSetPWFolderState {
    <#
    .SYNOPSIS
    Same pairing rules as local state, using PW document listings (optionally one level of subfolders).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath,
        [Parameter(Mandatory)]
        [bool]$OneLevelDeep
    )
    $discRes = _SSS-EnsurePWDiscoveryCmdlets
    if (-not $discRes.IsSuccess) {
        return @{
            folderStateHash = ''
            pairedCount = 0
            pdfCount = 0
            dgnCount = 0
            stableInput = ''
            orderKey = ''
            pairedSheets = @()
            orderedPdfDocuments = @()
            discoveryIncomplete = $true
            discoveryError = $discRes
        }
    }
    $state = _SSS-BuildPWStatusSetState -FolderPath $FolderPath -OneLevelDeep $OneLevelDeep -IncludeOrderedPdfDocuments:$false
    if ($OneLevelDeep -and [int]$state.pairedCount -le 0 -and [int]$state.pdfCount -le 0 -and [int]$state.dgnCount -le 0) {
        $flat = _SSS-BuildPWStatusSetState -FolderPath $FolderPath -OneLevelDeep:$false -IncludeOrderedPdfDocuments:$false
        if ([int]$flat.pairedCount -gt 0 -or [int]$flat.pdfCount -gt 0 -or [int]$flat.dgnCount -gt 0) {
            $flat['oneLevelDeepRetry'] = $true
            return $flat
        }
    }
    return $state
}

function Test-StatusSetSheetPdfTimestampMatch {
    <#
    .SYNOPSIS
    True when manifest sheet row and current paired row have the same PDF file size and PW lastWriteTimeUtc (UTC).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$ManifestSheetRow,
        [Parameter(Mandatory)]
        [object]$CurrentPairedRow
    )
    return _SSS-SheetPdfTimestampMatches -ManifestRow $ManifestSheetRow -CurrentRow $CurrentPairedRow
}

function Get-StatusSetPairedSheetSignature {
    <#
    .SYNOPSIS
    Fingerprint for one manifest or pairedSheets row (documentId+size when available).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Row
    )
    return _SSS-GetPairedSheetSignature -Row $Row
}

function Get-StatusSetWorkspaceDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LocalRoot,
        [Parameter(Mandatory)]
        [string]$SheetsFolderPath
    )

    if (_SSS-IsNullOrWhiteSpace $LocalRoot) {
        return New-QCFailureResult -Code 'STATUS_SET_LOCAL_ROOT_EMPTY' -Message 'LocalRoot is required.' -Data @{}
    }
    $h = _SSS-Sha256HexOfPath -Path $SheetsFolderPath
    if (-not $h) {
        return New-QCFailureResult -Code 'STATUS_SET_PATH_NORMALIZE_FAILED' -Message 'Failed to normalize SheetsFolderPath.' -Data @{ sheetsFolderPath = $SheetsFolderPath }
    }
    $key = $h.Substring(0, 16)
    $dir = Join-Path $LocalRoot $key
    return New-QCSuccessResult -Code 'STATUS_SET_WORKSPACE_OK' -Message 'Workspace directory resolved.' -Data @{
        workspaceDir = $dir
        folderKey = $key
        sheetsFolderNormalizedHash = $h
    }
}

function Get-StatusSetManifestPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$WorkspaceDir,
        [Parameter(Mandatory = $false)]
        [string]$ManifestFileName = '_statusset.manifest.json'
    )
    return Join-Path $WorkspaceDir $ManifestFileName
}

function Read-StatusSetManifestFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        return New-QCSuccessResult -Code 'STATUS_SET_MANIFEST_MISSING' -Message 'Manifest file not found.' -Data @{ manifest = $null; path = $Path }
    }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $h = @{}
        if ($obj.PSObject -and $obj.PSObject.Properties) {
            foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = $p.Value }
        }
        return New-QCSuccessResult -Code 'STATUS_SET_MANIFEST_READ' -Message 'Manifest loaded.' -Data @{ manifest = $h; path = $Path }
    } catch {
        return New-QCFailureResult -Code 'STATUS_SET_MANIFEST_READ_FAILED' -Message 'Failed to read manifest JSON.' -Data @{ path = $Path; errorMessage = $_.Exception.Message }
    }
}

function New-StatusSetManifestObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$State,
        [Parameter(Mandatory)]
        [string]$SheetsFolderDisplay
    )
    $sheets = @()
    foreach ($row in @([array]$State.pairedSheets)) {
        $key = ([string]$row.dir + '|' + [string]$row.stem).ToLowerInvariant()
        $dgnSide = $null
        if ($row -is [hashtable]) {
            if ($row.ContainsKey('dgn') -and $row['dgn']) { $dgnSide = $row['dgn'] }
            elseif ($row.ContainsKey('cad') -and $row['cad']) { $dgnSide = $row['cad'] }
        } elseif ($row -and $row.PSObject) {
            $pDgn = $row.PSObject.Properties['dgn']
            $pCad = $row.PSObject.Properties['cad']
            if ($pDgn -and $pDgn.Value) { $dgnSide = $pDgn.Value }
            elseif ($pCad -and $pCad.Value) { $dgnSide = $pCad.Value }
        }
        $sheets += @{
            key = $key
            stem = [string]$row.stem
            folderKey = [string]$row.dir
            pdf = $row.pdf
            dgn = $dgnSide
        }
    }
    $man = @{
        version = 1
        sheetsFolder = $SheetsFolderDisplay
        folderStateHash = [string]$State.folderStateHash
        generatedAtUtc = Get-QCTimestamp
        sheetCount = [int]$State.pairedCount
        sheets = $sheets
    }
    return $man
}

function Write-StatusSetManifestFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [hashtable]$Manifest
    )
    $tmp = $null
    try {
        $dir = Split-Path -Parent $Path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $json = $Manifest | ConvertTo-Json -Depth 12
        $tmpRoot = [System.IO.Path]::GetTempPath()
        $tmp = Join-Path $tmpRoot ('qc_statusset_man_' + [guid]::NewGuid().ToString('N') + '.json')
        $enc = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($tmp, $json, $enc)
        $pauseMs = [int]$script:_SSS_FsThrottleMs
        if ($pauseMs -lt 600) { $pauseMs = 600 }
        Start-Sleep -Milliseconds $pauseMs
        $lastEx = $null
        for ($a = 1; $a -le 22; $a++) {
            try {
                Move-Item -LiteralPath $tmp -Destination $Path -Force -ErrorAction Stop
                $tmp = $null
                return New-QCSuccessResult -Code 'STATUS_SET_MANIFEST_WRITTEN' -Message 'Manifest written.' -Data @{ path = $Path }
            } catch {
                $lastEx = $_
            }
            try {
                Copy-Item -LiteralPath $tmp -Destination $Path -Force -ErrorAction Stop
                Remove-Item -LiteralPath $tmp -Force -ErrorAction Stop
                $tmp = $null
                return New-QCSuccessResult -Code 'STATUS_SET_MANIFEST_WRITTEN' -Message 'Manifest written.' -Data @{ path = $Path }
            } catch {
                $lastEx = $_
            }
            if ($a -ge 22) { break }
            Start-Sleep -Milliseconds ([Math]::Min(3000, 200 + ($a * 150)))
        }
        throw $lastEx
    } catch {
        return New-QCFailureResult -Code 'STATUS_SET_MANIFEST_WRITE_FAILED' -Message 'Failed to write manifest.' -Data @{ path = $Path; errorMessage = $_.Exception.Message }
    } finally {
        if ($tmp -and (Test-Path -LiteralPath $tmp)) {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

function _SSS-GetPairedSheetSignature {
    <#
    .SYNOPSIS
    Stable fingerprint for one paired sheet row (manifest sheet or CurrentState.pairedSheets item).
    Aligns with Test-StatusSetRebuildNeeded: prefer PW documentId + file sizes; else timestamps.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Row
    )
    $pdf = $null
    $dgn = $null
    if ($Row -is [hashtable]) {
        if ($Row.ContainsKey('pdf')) { $pdf = $Row['pdf'] }
        if ($Row.ContainsKey('dgn')) { $dgn = $Row['dgn'] }
        elseif ($Row.ContainsKey('cad')) { $dgn = $Row['cad'] }
    } elseif ($Row -and $Row.PSObject) {
        try { $pdf = $Row.pdf } catch { }
        try {
            if ($Row.PSObject.Properties.Match('dgn').Count -gt 0) { $dgn = $Row.dgn }
            elseif ($Row.PSObject.Properties.Match('cad').Count -gt 0) { $dgn = $Row.cad }
        } catch { }
    }
    function _PartDoc([object]$Doc) {
        $pl = ''; $pm = ''; $docId = ''
        if ($Doc -is [hashtable]) {
            if ($Doc.ContainsKey('length')) { $pl = [string]$Doc['length'] }
            if ($Doc.ContainsKey('lastWriteTimeUtc')) { $pm = [string]$Doc['lastWriteTimeUtc'] }
            if ($Doc.ContainsKey('documentId')) { $docId = [string]$Doc['documentId'] }
        } elseif ($Doc -and $Doc.PSObject) {
            $pl = [string]($Doc | Select-Object -ExpandProperty length -ErrorAction SilentlyContinue)
            $pm = [string]($Doc | Select-Object -ExpandProperty lastWriteTimeUtc -ErrorAction SilentlyContinue)
            $docId = [string]($Doc | Select-Object -ExpandProperty documentId -ErrorAction SilentlyContinue)
        }
        return @{ pl = $pl; pm = $pm; docId = $docId }
    }
    $pp = _PartDoc -Doc $pdf
    $dp = _PartDoc -Doc $dgn
    if (-not [string]::IsNullOrWhiteSpace($pp.docId) -and -not [string]::IsNullOrWhiteSpace($dp.docId)) {
        return ($pp.docId + '|' + $pp.pl + '|' + $dp.docId + '|' + $dp.pl)
    }
    return ($pp.pl + '|' + $pp.pm + '|' + $dp.pl + '|' + $dp.pm)
}

function _SSS-SheetPdfTimestampMatches([object]$ManifestRow, [object]$CurrentRow) {
    $mp = $null
    $cp = $null
    if ($ManifestRow -is [hashtable] -and $ManifestRow.ContainsKey('pdf')) { $mp = $ManifestRow['pdf'] }
    elseif ($ManifestRow.PSObject) { try { $mp = $ManifestRow.pdf } catch { } }
    if ($CurrentRow -is [hashtable] -and $CurrentRow.ContainsKey('pdf')) { $cp = $CurrentRow['pdf'] }
    elseif ($CurrentRow.PSObject) { try { $cp = $CurrentRow.pdf } catch { } }
    if (-not $mp -or -not $cp) { return $false }
    function _Len([object]$Pdf) {
        if ($Pdf -is [hashtable] -and $Pdf.ContainsKey('length')) { return (_SSS-NormalizeFileSize $Pdf['length']) }
        if ($Pdf.PSObject) { return (_SSS-NormalizeFileSize ($Pdf | Select-Object -ExpandProperty length -ErrorAction SilentlyContinue)) }
        return ''
    }
    function _Mod([object]$Pdf) {
        if ($Pdf -is [hashtable] -and $Pdf.ContainsKey('lastWriteTimeUtc')) { return [string]$Pdf['lastWriteTimeUtc'] }
        if ($Pdf.PSObject) { return [string]($Pdf | Select-Object -ExpandProperty lastWriteTimeUtc -ErrorAction SilentlyContinue) }
        return ''
    }
    $l1 = _Len $mp
    $l2 = _Len $cp
    # Some PW APIs/configs do not reliably return file length, causing length-only drift and
    # forcing a full re-export every time. If either side lacks length, fall back to timestamp-only.
    if ([string]::IsNullOrWhiteSpace($l1) -or [string]::IsNullOrWhiteSpace($l2)) {
        $t1 = _SSS-ParseIsoDateTime -Value (_Mod $mp)
        $t2 = _SSS-ParseIsoDateTime -Value (_Mod $cp)
        if ($null -eq $t1 -or $null -eq $t2) { return $false }
        return ($t1 -eq $t2)
    }
    if ($l1 -ne $l2) { return $false }
    $t1 = _SSS-ParseIsoDateTime -Value (_Mod $mp)
    $t2 = _SSS-ParseIsoDateTime -Value (_Mod $cp)
    if ($null -eq $t1 -or $null -eq $t2) { return $false }
    return ($t1 -eq $t2)
}

function _SSS-StatusSetSheetCachePath {
    param(
        [Parameter(Mandatory)]
        [string]$WorkspaceDir,
        [Parameter(Mandatory)]
        [string]$SheetKey
    )
    $cacheRoot = Join-Path $WorkspaceDir '_sheet_cache'
    $h = _SSS-Sha256TextHex -Text (($SheetKey -as [string]).ToLowerInvariant())
    return (Join-Path $cacheRoot ($h + '.pdf'))
}

function _SSS-MaxPairedSheetPdfWriteUtc {
    param(
        [Parameter(Mandatory)]
        [hashtable]$CurrentState
    )
    $max = $null
    foreach ($row in @([array]$CurrentState['pairedSheets'])) {
        if (-not $row) { continue }
        $pdf = $null
        if ($row -is [hashtable] -and $row.ContainsKey('pdf')) { $pdf = $row['pdf'] }
        elseif ($row.PSObject) { try { $pdf = $row.pdf } catch { } }
        if (-not $pdf) { continue }
        $pm = $null
        if ($pdf -is [hashtable] -and $pdf.ContainsKey('lastWriteTimeUtc')) { $pm = $pdf['lastWriteTimeUtc'] }
        elseif ($pdf.PSObject) { $pm = $pdf | Select-Object -ExpandProperty lastWriteTimeUtc -ErrorAction SilentlyContinue }
        $dt = _SSS-ParseIsoDateTime -Value $pm
        if ($dt) {
            if ($null -eq $max -or $dt -gt $max) { $max = $dt }
        }
    }
    return $max
}

function Test-StatusSetRebuildNeeded {
    <#
    .SYNOPSIS
    Compares last-written manifest to current folder state (hash + per-sheet signatures).
    .DESCRIPTION
    When StatusSetPdfPath points to an existing workspace _statusset.pdf and paired sheet PDF
    timestamps are available, a rebuild is forced if that PDF is older than the newest paired
    sheet PDF. If structure/signatures match and the status PDF is not older, hash-only drift
    is ignored (skips rebuild) — mirrors "sheets unchanged since last merge" intent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Manifest,
        [Parameter(Mandatory)]
        [hashtable]$CurrentState,
        [Parameter(Mandatory = $false)]
        [switch]$ForceRebuild,
        [Parameter(Mandatory = $false)]
        [string]$StatusSetPdfPath
    )

    if ($ForceRebuild) {
        return New-QCSuccessResult -Code 'STATUS_SET_COMPARE' -Message 'Rebuild forced.' -Data @{
            needsRebuild = $true
            reasons = @('ForceRebuild')
            onlyInManifest = @()
            onlyInFolder = @()
            signatureMismatches = @()
            manifestHash = $null
            currentHash = [string]$CurrentState.folderStateHash
        }
    }

    if ($null -eq $Manifest) {
        return New-QCSuccessResult -Code 'STATUS_SET_COMPARE' -Message 'No manifest: rebuild needed.' -Data @{
            needsRebuild = $true
            reasons = @('NO_MANIFEST')
            onlyInManifest = @()
            onlyInFolder = @()
            signatureMismatches = @()
            manifestHash = $null
            currentHash = [string]$CurrentState.folderStateHash
        }
    }

    $m = $Manifest
    if ($m -isnot [hashtable]) {
        if ($m.PSObject -and $m.PSObject.Properties) {
            $mh = @{}
            foreach ($p in $m.PSObject.Properties) { $mh[$p.Name] = $p.Value }
            $m = $mh
        }
    }

    $mh = if ($m.ContainsKey('folderStateHash')) { [string]$m['folderStateHash'] } else { '' }
    $ch = [string]$CurrentState.folderStateHash
    if ($mh -and $ch -eq $mh) {
        if (-not [string]::IsNullOrWhiteSpace($StatusSetPdfPath) -and (Test-Path -LiteralPath $StatusSetPdfPath)) {
            $maxEarly = _SSS-MaxPairedSheetPdfWriteUtc -CurrentState $CurrentState
            if ($null -ne $maxEarly) {
                try {
                    $outEarly = (Get-Item -LiteralPath $StatusSetPdfPath -ErrorAction Stop).LastWriteTimeUtc
                    if ($outEarly -lt $maxEarly) {
                        return New-QCSuccessResult -Code 'STATUS_SET_COMPARE' -Message 'Manifest hash matches but status set PDF is older than a paired sheet PDF; rebuild.' -Data @{
                            needsRebuild = $true
                            reasons = @('STATUS_SET_PDF_OLDER_THAN_SHEET_PDF')
                            onlyInManifest = @()
                            onlyInFolder = @()
                            signatureMismatches = @()
                            manifestHash = $mh
                            currentHash = $ch
                            statusSetPdfPath = $StatusSetPdfPath
                            statusSetPdfUtc = $outEarly
                            newestPairedSheetPdfUtc = $maxEarly
                            usedStatusSetPdfMtimeRule = $true
                        }
                    }
                } catch { }
            }
        }
        return New-QCSuccessResult -Code 'STATUS_SET_COMPARE' -Message 'folderStateHash matches manifest; skip rebuild.' -Data @{
            needsRebuild = $false
            reasons = @()
            onlyInManifest = @()
            onlyInFolder = @()
            signatureMismatches = @()
            manifestHash = $mh
            currentHash = $ch
        }
    }

    $manifestKeys = @{}
    foreach ($s in @([array]$m['sheets'])) {
        $key = $null
        if ($s -is [hashtable] -and $s.ContainsKey('key')) { $key = [string]$s['key'] }
        elseif ($s.PSObject.Properties['key']) { $key = [string]$s.key }
        if ($key) {
            $manifestKeys[$key.ToLowerInvariant()] = $s
        }
    }

    $folderKeys = @{}
    foreach ($s in @([array]$CurrentState.pairedSheets)) {
        $k = ([string]$s.dir + '|' + [string]$s.stem).ToLowerInvariant()
        $folderKeys[$k] = $s
    }

    $onlyM = @()
    $onlyF = @()
    foreach ($k in $manifestKeys.Keys) {
        if (-not $folderKeys.ContainsKey($k)) { $onlyM += $k }
    }
    foreach ($k in $folderKeys.Keys) {
        if (-not $manifestKeys.ContainsKey($k)) { $onlyF += $k }
    }

    $sigMismatch = @()
    foreach ($k in $manifestKeys.Keys) {
        if (-not $folderKeys.ContainsKey($k)) { continue }
        $sm = _SSS-GetPairedSheetSignature -Row $manifestKeys[$k]
        $sf = _SSS-GetPairedSheetSignature -Row $folderKeys[$k]
        if ($sm -ne $sf) { $sigMismatch += $k }
    }

    $structuralDrift = ($onlyM.Count -gt 0) -or ($onlyF.Count -gt 0) -or ($sigMismatch.Count -gt 0)

    $maxSheetPdfUtc = _SSS-MaxPairedSheetPdfWriteUtc -CurrentState $CurrentState
    $statusSetPdfUtc = $null
    $staleVsSheetPdfs = $false
    $usedStatusSetPdfMtimeRule = $false
    if (-not [string]::IsNullOrWhiteSpace($StatusSetPdfPath) -and (Test-Path -LiteralPath $StatusSetPdfPath) -and ($null -ne $maxSheetPdfUtc)) {
        try {
            $statusSetPdfUtc = (Get-Item -LiteralPath $StatusSetPdfPath -ErrorAction Stop).LastWriteTimeUtc
            $usedStatusSetPdfMtimeRule = $true
            if ($statusSetPdfUtc -lt $maxSheetPdfUtc) { $staleVsSheetPdfs = $true }
        } catch {
            $usedStatusSetPdfMtimeRule = $false
            $statusSetPdfUtc = $null
        }
    }

    if ((-not $structuralDrift) -and $usedStatusSetPdfMtimeRule -and -not $staleVsSheetPdfs) {
        return New-QCSuccessResult -Code 'STATUS_SET_COMPARE' -Message 'Status set PDF is newer than all paired sheet PDFs; skip rebuild.' -Data @{
            needsRebuild = $false
            reasons = @()
            onlyInManifest = $onlyM
            onlyInFolder = $onlyF
            signatureMismatches = $sigMismatch
            manifestHash = $mh
            currentHash = $ch
            statusSetPdfPath = $StatusSetPdfPath
            statusSetPdfUtc = $statusSetPdfUtc
            newestPairedSheetPdfUtc = $maxSheetPdfUtc
            usedStatusSetPdfMtimeRule = $true
        }
    }

    $reasons = @()
    if ($onlyM.Count -gt 0) { $reasons += 'SHEETS_REMOVED_FROM_FOLDER' }
    if ($onlyF.Count -gt 0) { $reasons += 'SHEETS_ADDED_TO_FOLDER' }
    if ($sigMismatch.Count -gt 0) { $reasons += 'SHEET_METADATA_CHANGED' }
    if ($staleVsSheetPdfs) { $reasons += 'STATUS_SET_PDF_OLDER_THAN_SHEET_PDF' }
    if ($mh -ne $ch) { $reasons += 'FOLDER_STATE_HASH_MISMATCH' }

    $needs = $structuralDrift -or $staleVsSheetPdfs -or ($mh -ne $ch)

    return New-QCSuccessResult -Code 'STATUS_SET_COMPARE' -Message 'Comparison complete.' -Data @{
        needsRebuild = [bool]$needs
        reasons = $reasons
        onlyInManifest = $onlyM
        onlyInFolder = $onlyF
        signatureMismatches = $sigMismatch
        manifestHash = $mh
        currentHash = $ch
        statusSetPdfPath = $StatusSetPdfPath
        statusSetPdfUtc = $statusSetPdfUtc
        newestPairedSheetPdfUtc = $maxSheetPdfUtc
        usedStatusSetPdfMtimeRule = $usedStatusSetPdfMtimeRule
    }
}

function Merge-StatusSetPdfWithQpdf {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$OrderedInputPdfPaths,
        [Parameter(Mandatory)]
        [string]$OutputPdf,
        [Parameter(Mandatory)]
        [string]$QpdfExe
    )

    # Delegate to the internal _SSS-MergePdfs helper so there is exactly one place
    # that builds qpdf arguments. _SSS-MergePdfs mirrors legacy/combine_status_set.ps1:
    #   - one '--empty --pages <files...> -- <out>' invocation per batch
    #   - 1-file fast path uses Copy-Item
    #   - >100 files chunks recursively to avoid the Windows ~32K command-line limit
    # Returning rich Result objects is preserved for callers/diagnostics.

    $missing = @($OrderedInputPdfPaths | Where-Object { -not (Test-Path -LiteralPath $_) })
    if ($missing.Count -gt 0) {
        return New-QCFailureResult -Code 'STATUS_SET_MERGE_INPUT_MISSING' -Message 'One or more input PDFs are missing on disk.' -Data @{ missing = $missing }
    }
    if (-not (Test-Path -LiteralPath $QpdfExe)) {
        return New-QCFailureResult -Code 'STATUS_SET_QPDF_MISSING' -Message "qpdf not found: $QpdfExe" -Data @{ qpdfExe = $QpdfExe }
    }

    $outDir = Split-Path -Parent $OutputPdf
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    try {
        _SSS-MergePdfs -PdfPaths $OrderedInputPdfPaths -OutPath $OutputPdf -QpdfExe $QpdfExe
    } catch {
        $errMsg = [string]$_.Exception.Message
        return New-QCFailureResult -Code 'STATUS_SET_QPDF_FAILED' -Message 'qpdf merge failed.' -Data @{
            error      = $errMsg
            inputCount = @($OrderedInputPdfPaths).Count
            outputPdf  = $OutputPdf
            qpdfExe    = $QpdfExe
        }
    }

    if (-not (Test-Path -LiteralPath $OutputPdf)) {
        return New-QCFailureResult -Code 'STATUS_SET_QPDF_FAILED' -Message 'qpdf produced no output.' -Data @{
            inputCount = @($OrderedInputPdfPaths).Count
            outputPdf  = $OutputPdf
        }
    }

    return New-QCSuccessResult -Code 'STATUS_SET_MERGE_OK' -Message 'Merged PDF written.' -Data @{
        outputPdf  = $OutputPdf
        inputCount = @($OrderedInputPdfPaths).Count
    }
}

function Export-StatusSetPdfToFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $InputDocument,
        [Parameter(Mandatory)]
        [string]$TargetFolder
    )

    $cmd = Get-Command -Name Export-PWDocumentsSimple -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return New-QCFailureResult -Code 'STATUS_SET_EXPORT_CMD_MISSING' -Message 'Export-PWDocumentsSimple not available (run in ProjectWise PowerShell).' -Data @{}
    }
    if (-not (Test-Path -LiteralPath $TargetFolder)) {
        New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
    }
    try {
        Export-PWDocumentsSimple -InputDocuments $InputDocument -TargetFolder $TargetFolder -ErrorAction Stop | Out-Null
        # Match legacy combine_status_set.ps1 line ~828: a brief sleep after every PW
        # export gives Fortinet (and similar AV) time to scan the freshly-written
        # PDF before we issue the next export. Without this, back-to-back exports
        # can hit file-handle contention and produce partial/zero-byte PDFs.
        _SSS-PwExportThrottle
        $name = _SSS-PWGetDocName -Doc $InputDocument
        $direct = Join-Path $TargetFolder $name
        if (Test-Path -LiteralPath $direct) {
            return New-QCSuccessResult -Code 'STATUS_SET_EXPORT_OK' -Message 'Exported PDF.' -Data @{ localPath = $direct }
        }
        $cname = _SSS-PWGetProp -Obj $InputDocument -Name 'CopiedOutLocalFileName'
        if ($cname) {
            $alt = Join-Path $TargetFolder ([string]$cname)
            if (Test-Path -LiteralPath $alt) {
                return New-QCSuccessResult -Code 'STATUS_SET_EXPORT_OK' -Message 'Exported PDF (CopiedOutLocalFileName).' -Data @{ localPath = $alt }
            }
        }
        # Ignore numbered scratch files (NNN_name.pdf) from a prior partial run in the same export folder.
        $newest = Get-ChildItem -LiteralPath $TargetFolder -File -Filter '*.pdf' -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notmatch '^\d{3}_' } |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 1
        if ($newest) {
            return New-QCSuccessResult -Code 'STATUS_SET_EXPORT_OK' -Message 'Exported PDF (newest unnumbered pdf in folder).' -Data @{ localPath = [string]$newest.FullName }
        }
        return New-QCFailureResult -Code 'STATUS_SET_EXPORT_RESOLVE_FAILED' -Message 'Export ran but local PDF path could not be resolved.' -Data @{ targetFolder = $TargetFolder; documentName = $name }
    } catch {
        return New-QCFailureResult -Code 'STATUS_SET_EXPORT_FAILED' -Message 'Export-PWDocumentsSimple failed.' -Data @{ errorMessage = $_.Exception.Message }
    }
}

function Invoke-StatusSetNativeJob {
    <#
    .SYNOPSIS
    Manifest compare vs live PW inventory; export paired PDFs; qpdf merge; write manifest; optional PW write-back of _statusset.pdf.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    _SSS-ApplyFsThrottleConfig -Config $Config

    $ss = @{}
    if ($Config.ContainsKey('statusSet') -and $Config.statusSet) {
        $raw = $Config.statusSet
        if ($raw -is [hashtable]) { $ss = $raw } elseif ($raw.PSObject) {
            foreach ($p in $raw.PSObject.Properties) { $ss[$p.Name] = $p.Value }
        }
    }

    $localRoot = if ($ss.ContainsKey('localRoot') -and $ss.localRoot) { [string]$ss.localRoot } else { 'C:\PW_QC_LOCAL' }
    $manifestName = if ($ss.ContainsKey('manifestFileName') -and $ss.manifestFileName) { [string]$ss.manifestFileName } else { '_statusset.manifest.json' }
    $statusPdfName = if ($ss.ContainsKey('statusSetPdfName') -and $ss.statusSetPdfName) { [string]$ss.statusSetPdfName } else { '_statusset.pdf' }
    $qpdfExe = if ($ss.ContainsKey('qpdfExe') -and $ss.qpdfExe) { [string]$ss.qpdfExe } else { '' }
    if (_SSS-IsNullOrWhiteSpace $qpdfExe) {
        if ($Config.ContainsKey('qcPrepend') -and $Config.qcPrepend) {
            $qc = $Config.qcPrepend
            if ($qc -is [hashtable] -and $qc.ContainsKey('qpdfExePath')) { $qpdfExe = [string]$qc['qpdfExePath'] }
            elseif ($qc.PSObject.Properties['qpdfExePath']) { $qpdfExe = [string]$qc.qpdfExePath }
        }
    }
    $repoRoot = Split-Path -Parent $PSScriptRoot
    if (-not (_SSS-IsNullOrWhiteSpace $qpdfExe)) {
        try {
            if (-not [System.IO.Path]::IsPathRooted($qpdfExe)) {
                $qpdfExe = Join-Path $repoRoot $qpdfExe
            }
        } catch { }
    }
    if (_SSS-IsNullOrWhiteSpace $qpdfExe) { $qpdfExe = Join-Path $repoRoot 'tools\qpdf\bin\qpdf.exe' }

    $forceRebuild = $false
    if ($ss.ContainsKey('forceRebuild')) { try { $forceRebuild = [bool]$ss.forceRebuild } catch { $forceRebuild = $false } }
    $writeBackToPW = $false
    if ($ss.ContainsKey('writeBackToPW')) { try { $writeBackToPW = [bool]$ss.writeBackToPW } catch { $writeBackToPW = $false } }

    $pwSheetCacheEnabled = $true
    if ($ss.ContainsKey('pwSheetCache')) {
        try { $pwSheetCacheEnabled = [bool]$ss.pwSheetCache } catch { $pwSheetCacheEnabled = $true }
    }
    $pwSheetReusePolicy = 'Signature'
    if ($ss.ContainsKey('pwSheetReusePolicy') -and $ss.pwSheetReusePolicy) {
        $pwSheetReusePolicy = ([string]$ss.pwSheetReusePolicy).Trim()
    }

    $incrementalMode = _SSS-GetHashtableBool -Map $ss -Name 'incrementalMode' -Default $true
    if (-not $incrementalMode) { $forceRebuild = $true }
    $stagingRoot = if ($ss.ContainsKey('stagingRoot') -and $ss.stagingRoot) { [string]$ss.stagingRoot } else { '' }
    $retentionDays = _SSS-GetHashtableInt -Map $ss -Name 'retentionDays' -Default 14
    $cleanupEnabled = _SSS-GetHashtableBool -Map $ss -Name 'cleanupEnabled' -Default $false
    $cleanImmediateExportScratch = _SSS-GetHashtableBool -Map $ss -Name 'cleanImmediateExportScratch' -Default $false
    $atomicReplaceEnabled = _SSS-GetHashtableBool -Map $ss -Name 'atomicReplaceEnabled' -Default $true
    $dryRunOperationReport = _SSS-GetHashtableBool -Map $ss -Name 'dryRunOperationReport' -Default $true

    $sourceFolder = if ($Job.ContainsKey('sourceFolder') -and $Job.sourceFolder) { [string]$Job.sourceFolder } else { '' }
    if (_SSS-IsNullOrWhiteSpace $sourceFolder) {
        return New-QCFailureResult -Code 'STATUS_SET_MISSING_SOURCE_FOLDER' -Message 'Job.sourceFolder is required.' -Data @{}
    }

    $oneLevelDeep = $true
    if ($Job.ContainsKey('metadata') -and $Job.metadata -is [hashtable] -and $Job.metadata.ContainsKey('candidate')) {
        $cand = $Job.metadata.candidate
        if ($cand -is [hashtable] -and $cand.ContainsKey('oneLevelDeep')) {
            try { $oneLevelDeep = [bool]$cand.oneLevelDeep } catch { $oneLevelDeep = $true }
        }
    }

    $pwPath = $sourceFolder -replace '^Documents\\', ''

    $pathNormRes = Normalize-QCPath -Path $sourceFolder
    $normForDetect = if ($pathNormRes.IsSuccess) { [string]$pathNormRes.Data.path } else { $sourceFolder }
    $isPwLogicalPath = $normForDetect.StartsWith('documents\', [System.StringComparison]::OrdinalIgnoreCase) -or ($sourceFolder -match '(?i)^pw:\\')
    $useLocalFs = (-not $isPwLogicalPath) -and (Test-Path -LiteralPath $sourceFolder -PathType Container)

    $wsRes = Get-StatusSetWorkspaceDirectory -LocalRoot $localRoot -SheetsFolderPath $sourceFolder
    if (-not $wsRes.IsSuccess) { return $wsRes }
    $workspace = [string]$wsRes.Data.workspaceDir
    $manifestPath = Get-StatusSetManifestPath -WorkspaceDir $workspace -ManifestFileName $manifestName
    $outPdf = Join-Path $workspace $statusPdfName
    $stagingBase = if (-not (_SSS-IsNullOrWhiteSpace $stagingRoot)) { Join-Path $stagingRoot ([string]$wsRes.Data.folderKey) } else { $workspace }
    $operationReport = _SSS-NewStatusSetOperationReport -WorkspaceDir $workspace -OutputPdf $outPdf -ManifestPath $manifestPath -Options @{
        incrementalMode = $incrementalMode
        stagingRoot = $stagingBase
        retentionDays = $retentionDays
        cleanupEnabled = $cleanupEnabled
        cleanImmediateExportScratch = $cleanImmediateExportScratch
        atomicReplaceEnabled = $atomicReplaceEnabled
        dryRunOperationReport = $dryRunOperationReport
        pwSheetCache = $pwSheetCacheEnabled
        pwSheetReusePolicy = $pwSheetReusePolicy
    }

    $isDryRun = $false
    if ($Config.ContainsKey('dryRun')) { try { $isDryRun = [bool]$Config.dryRun } catch { $isDryRun = $false } }
    if ($cleanupEnabled) {
        _SSS-CleanupExpiredStatusSetStaging -WorkspaceDir $workspace -RetentionDays $retentionDays -OperationReport $operationReport -DryRun:$isDryRun
        if ($stagingBase -ne $workspace) {
            _SSS-CleanupExpiredStatusSetStaging -WorkspaceDir $stagingBase -RetentionDays $retentionDays -OperationReport $operationReport -DryRun:$isDryRun
        }
    }

    $pwCfg = @{}
    if ($Config.ContainsKey('projectWise')) {
        $p = $Config.projectWise
        if ($p -is [hashtable]) { $pwCfg = $p }
        elseif ($p.PSObject) { foreach ($x in $p.PSObject.Properties) { $pwCfg[$x.Name] = $x.Value } }
    }
    $credPath = if ($pwCfg.ContainsKey('credentialPath')) { [string]$pwCfg['credentialPath'] } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
    $ds = if ($Job.ContainsKey('metadata') -and $Job.metadata -is [hashtable] -and $Job.metadata.ContainsKey('candidate')) {
        $c = $Job.metadata.candidate
        if ($c -is [hashtable] -and $c.ContainsKey('datasourceName')) { [string]$c['datasourceName'] }
        else { '' }
    } else { '' }
    if (_SSS-IsNullOrWhiteSpace $ds) {
        $ds = if ($pwCfg.ContainsKey('datasourceName')) { [string]$pwCfg['datasourceName'] } else { 'typsa-us-pw.bentley.com:typsa-us-pw-03' }
    }

    $pwConnected = $false
    if (-not $useLocalFs) {
        Import-Module (Join-Path $PSScriptRoot 'PW.Connection.psm1') -Force
        $credRes = Get-PWCredentialFromFile -CredentialPath $credPath
        if (-not $credRes.IsSuccess) { return $credRes }
        $connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
        if (-not $connRes.IsSuccess) { return $connRes }
        $pwConnected = $true
    }

    try {
        if ($useLocalFs) {
            $ls = Get-StatusSetLocalFolderState -RootFolder $sourceFolder
            $fullState = @{
                folderStateHash = $ls.folderStateHash
                pairedCount = $ls.pairedCount
                pairedSheets = $ls.pairedSheets
                stableInput = $ls.stableInput
                orderKey = $ls.orderKey
                pdfCount = 0
                dgnCount = 0
                orderedPdfDocuments = @()
            }
        } else {
            $discRes = _SSS-EnsurePWDiscoveryCmdlets
            if (-not $discRes.IsSuccess) {
                return New-QCFailureResult -Code 'STATUS_SET_PW_DISCOVERY_INCOMPLETE' -Message $discRes.Message -Data @{ folder = $sourceFolder; pwPath = $pwPath; oneLevelDeep = $oneLevelDeep; useLocalFs = $useLocalFs; discovery = $discRes.Data }
            }
            $fullState = _SSS-BuildPWStatusSetState -FolderPath $pwPath -OneLevelDeep $oneLevelDeep -IncludeOrderedPdfDocuments:$true
        }

        if ([int]$fullState.pairedCount -le 0) {
            if (-not $useLocalFs -and (-not (_SSS-TestPWDiscoveryCmdlets))) {
                return New-QCFailureResult -Code 'STATUS_SET_PW_DISCOVERY_INCOMPLETE' -Message 'ProjectWise returned zero pairs while pwps_dab discovery cmdlets are missing; treating as incomplete discovery instead of no pairs.' -Data @{ folder = $sourceFolder; pwPath = $pwPath; oneLevelDeep = $oneLevelDeep; useLocalFs = $useLocalFs; pairedCount = [int]$fullState.pairedCount; pdfCount = [int]$fullState.pdfCount; dgnCount = [int]$fullState.dgnCount }
            }
            return New-QCFailureResult -Code 'STATUS_SET_NO_PAIRS' -Message 'No PDF/DGN pairs found for status set.' -Data @{ folder = $sourceFolder; oneLevelDeep = $oneLevelDeep; useLocalFs = $useLocalFs }
        }

        $readM = Read-StatusSetManifestFile -Path $manifestPath
        if (-not $readM.IsSuccess) { return $readM }
        $manifest = $readM.Data.manifest

        $cmp = Test-StatusSetRebuildNeeded -Manifest $manifest -CurrentState $fullState -ForceRebuild:$forceRebuild -StatusSetPdfPath $outPdf
        if (-not $cmp.IsSuccess) { return $cmp }

        if (-not [bool]$cmp.Data.needsRebuild) {
            _SSS-AddStatusSetOperation -Report $operationReport -Kind 'skips' -Operation @{ action='status-set-cycle'; reason='manifest-and-output-current'; outputPdf=$outPdf }
            return New-QCSuccessResult -Code 'STATUS_SET_UP_TO_DATE' -Message 'Status set manifest matches folder state; no rebuild.' -Data @{
                jobId = [string]$Job.id
                workspaceDir = $workspace
                manifestPath = $manifestPath
                folderStateHash = [string]$fullState.folderStateHash
                useLocalFs = $useLocalFs
                operationReport = if ($dryRunOperationReport) { $operationReport } else { $null }
            }
        }

        if ($isDryRun) {
            $renderPdf = Join-Path (Join-Path $stagingBase '_render') ($statusPdfName + '.next.' + ([string]$Job.id -replace '[^a-zA-Z0-9._-]', '_') + '.pdf')
            if ($useLocalFs) {
                foreach ($row in @($fullState.pairedSheets)) {
                    $pdfPath = $null
                    try { $pdfPath = [string]$row.pdf.fullName } catch { }
                    _SSS-AddStatusSetOperation -Report $operationReport -Kind 'skips' -Operation @{ action='download'; reason='local-filesystem-source'; source=$pdfPath }
                }
            } else {
                $manByKeyDry = @{}
                if ($pwSheetCacheEnabled -and (-not $forceRebuild) -and $manifest -and $manifest['sheets']) {
                    foreach ($s in @([array]$manifest['sheets'])) {
                        $sk = $null
                        if ($s -is [hashtable] -and $s['key']) { $sk = [string]$s['key'] }
                        elseif ($s.PSObject.Properties['key']) { $sk = [string]$s.key }
                        if ($sk) { $manByKeyDry[$sk.ToLowerInvariant()] = $s }
                    }
                }
                foreach ($pairRow in @([array]$fullState.pairedSheets)) {
                    $key = ([string]$pairRow.dir + '|' + [string]$pairRow.stem).ToLowerInvariant()
                    $cachePath = _SSS-StatusSetSheetCachePath -WorkspaceDir $workspace -SheetKey $key
                    $reuseOk = $false
                    if ($pwSheetCacheEnabled -and $manByKeyDry.ContainsKey($key)) {
                        $manRow = $manByKeyDry[$key]
                        if ($pwSheetReusePolicy.Equals('PdfTimestamp', [System.StringComparison]::OrdinalIgnoreCase)) { $reuseOk = _SSS-SheetPdfTimestampMatches -ManifestRow $manRow -CurrentRow $pairRow }
                        else { $reuseOk = ((_SSS-GetPairedSheetSignature -Row $manRow) -eq (_SSS-GetPairedSheetSignature -Row $pairRow)) }
                    }
                    if ($reuseOk -and (Test-Path -LiteralPath $cachePath)) {
                        _SSS-AddStatusSetOperation -Report $operationReport -Kind 'skips' -Operation @{ action='download'; reason='cache-current'; sheetKey=$key; cachePath=$cachePath }
                    } else {
                        _SSS-AddStatusSetOperation -Report $operationReport -Kind 'downloads' -Operation @{ action='projectwise-export'; sheetKey=$key; target='sheet-cache'; cachePath=$cachePath }
                    }
                }
            }
            _SSS-AddStatusSetOperation -Report $operationReport -Kind 'writes' -Operation @{ action='qpdf-merge'; output=$renderPdf; inputCount=[int]$fullState.pairedCount }
            _SSS-AddStatusSetOperation -Report $operationReport -Kind 'replaces' -Operation @{ action='install-output'; destination=$outPdf; atomicReplace=$atomicReplaceEnabled; preserveHistory=$true }
            _SSS-AddStatusSetOperation -Report $operationReport -Kind 'writes' -Operation @{ action='write-manifest'; path=$manifestPath }
            if ($writeBackToPW -and -not $useLocalFs) { _SSS-AddStatusSetOperation -Report $operationReport -Kind 'writes' -Operation @{ action='projectwise-upload'; folder=$pwPath; documentName=$statusPdfName } }
            return New-QCSuccessResult -Code 'STATUS_SET_DRYRUN_REBUILD' -Message 'Dry-run: would rebuild status set PDF (manifest drift detected).' -Data @{
                jobId = [string]$Job.id
                compare = $cmp.Data
                workspaceDir = $workspace
                qpdfExe = $qpdfExe
                useLocalFs = $useLocalFs
                operationReport = if ($dryRunOperationReport) { $operationReport } else { $null }
            }
        }

        $orderedPaths = @()
        $pwExportReuseCount = 0
        $pwExportFreshCount = 0
        if ($useLocalFs) {
            $sortedPairs = @($fullState.pairedSheets | Sort-Object @{ Expression = { $_.dir } }, @{ Expression = { $_.stem } })
            foreach ($row in $sortedPairs) {
                if ($row.pdf -is [hashtable] -and $row.pdf.ContainsKey('fullName')) {
                    $localInput = [string]$row.pdf['fullName']
                    $orderedPaths += $localInput
                    _SSS-AddStatusSetOperation -Report $operationReport -Kind 'skips' -Operation @{ action='download'; reason='local-filesystem-source'; source=$localInput }
                }
            }
        } else {
            _SSS-EnsureDir (Join-Path $workspace '_sheet_cache')
            _SSS-EnsureDir $stagingBase
            $exportJobTag = [string]$Job['id']
            if (_SSS-IsNullOrWhiteSpace $exportJobTag) { $exportJobTag = [guid]::NewGuid().ToString('N') }
            $exportJobTag = ($exportJobTag -replace '[^a-zA-Z0-9._-]', '_')
            if ($exportJobTag.Length -gt 72) { $exportJobTag = $exportJobTag.Substring(0, 72) }
            # One scratch folder per job: fewer directory trees for AV to flag than per-sheet folders.
            # Basename collisions (same PDF name from two subfolders) are avoided by renaming after each export.
            $exportWorkDir = Join-Path $stagingBase ('_export_' + $exportJobTag)
            _SSS-EnsureDir $exportWorkDir
            # Same job id reuses this folder on retry; clear stale NNN_*.pdf so rename does not fail.
            try { _SSS-RemoveExportDirContentsFileFirst -DirPath $exportWorkDir -RemoveEmptyDir $false } catch { }

            $manByKey = @{}
            if ($pwSheetCacheEnabled -and (-not $forceRebuild) -and $manifest) {
                $mht = $manifest
                if ($mht -isnot [hashtable] -and $mht.PSObject) {
                    $mht = @{}
                    foreach ($p in $manifest.PSObject.Properties) { $mht[$p.Name] = $p.Value }
                }
                if ($mht -is [hashtable] -and $mht['sheets']) {
                    foreach ($s in @([array]$mht['sheets'])) {
                        $sk = $null
                        if ($s -is [hashtable] -and $s['key']) { $sk = [string]$s['key'] }
                        elseif ($s.PSObject.Properties['key']) { $sk = [string]$s.key }
                        if ($sk) { $manByKey[$sk.ToLowerInvariant()] = $s }
                    }
                }
            }

            $pairList = @([array]$fullState.pairedSheets)
            $docsList = @([array]$fullState.orderedPdfDocuments)
            if ($pairList.Count -ne $docsList.Count) {
                return New-QCFailureResult -Code 'STATUS_SET_STATE_MISMATCH' -Message 'pairedSheets and orderedPdfDocuments counts differ.' -Data @{ paired = $pairList.Count; docs = $docsList.Count }
            }

            $policyNorm = $pwSheetReusePolicy
            if ([string]::IsNullOrWhiteSpace($policyNorm)) { $policyNorm = 'Signature' }

            for ($i = 0; $i -lt $docsList.Count; $i++) {
                $doc = $docsList[$i]
                $pairRow = $pairList[$i]
                $key = ([string]$pairRow.dir + '|' + [string]$pairRow.stem).ToLowerInvariant()
                $cachePath = _SSS-StatusSetSheetCachePath -WorkspaceDir $workspace -SheetKey $key

                $reuseOk = $false
                if ($pwSheetCacheEnabled -and $manByKey.ContainsKey($key)) {
                    $manRow = $manByKey[$key]
                    if ($policyNorm.Equals('PdfTimestamp', [System.StringComparison]::OrdinalIgnoreCase)) {
                        $reuseOk = _SSS-SheetPdfTimestampMatches -ManifestRow $manRow -CurrentRow $pairRow
                    } else {
                        $sigM = _SSS-GetPairedSheetSignature -Row $manRow
                        $sigC = _SSS-GetPairedSheetSignature -Row $pairRow
                        $reuseOk = ($sigM -eq $sigC)
                    }
                }

                if ($reuseOk -and (Test-Path -LiteralPath $cachePath)) {
                    try {
                        if ((Get-Item -LiteralPath $cachePath -ErrorAction Stop).Length -le 0) { $reuseOk = $false }
                    } catch { $reuseOk = $false }
                } else {
                    $reuseOk = $false
                }

                if ($reuseOk) {
                    $orderedPaths += $cachePath
                    $pwExportReuseCount++
                    _SSS-AddStatusSetOperation -Report $operationReport -Kind 'skips' -Operation @{ action='download'; reason='cache-current'; sheetKey=$key; cachePath=$cachePath }
                    continue
                }

                $idx = $i + 1
                _SSS-AddStatusSetOperation -Report $operationReport -Kind 'downloads' -Operation @{ action='projectwise-export'; sheetKey=$key; targetFolder=$exportWorkDir; cachePath=$cachePath }
                $ex = Export-StatusSetPdfToFolder -InputDocument $doc -TargetFolder $exportWorkDir
                if (-not $ex.IsSuccess) { return $ex }
                $localPath = [string]$ex.Data.localPath
                try {
                    $localPath = _SSS-MoveStatusSetExportToUniqueName -LocalPath $localPath -ExportWorkDir $exportWorkDir -SequenceIndex $idx
                } catch {
                    return New-QCFailureResult -Code 'STATUS_SET_EXPORT_RENAME_FAILED' -Message 'Could not move exported PDF to unique name in export folder.' -Data @{
                        from = [string]$ex.Data.localPath
                        to = (Join-Path $exportWorkDir (('{0:000}_{1}' -f $idx, (_SSS-NormalizeStatusSetExportLeaf -Leaf ([System.IO.Path]::GetFileName($localPath))))))
                        errorMessage = $_.Exception.Message
                    }
                }
                $orderedPaths += $localPath
                $pwExportFreshCount++
                try {
                    Copy-Item -LiteralPath $localPath -Destination $cachePath -Force -ErrorAction Stop
                } catch { }
            }
        }

        $renderDir = Join-Path $stagingBase '_render'
        _SSS-EnsureDir $renderDir
        $renderTag = [string]$Job['id']
        if (_SSS-IsNullOrWhiteSpace $renderTag) { $renderTag = [guid]::NewGuid().ToString('N') }
        $renderTag = ($renderTag -replace '[^a-zA-Z0-9._-]', '_')
        $renderStamp = (Get-QCWallClockNow).ToString('yyyyMMddHHmmssfff')
        $renderPdf = if ($atomicReplaceEnabled) { Join-Path $renderDir ($statusPdfName + '.next.' + $renderTag + '.' + $renderStamp + '.pdf') } else { $outPdf }
        _SSS-AddStatusSetOperation -Report $operationReport -Kind 'writes' -Operation @{ action='qpdf-merge'; output=$renderPdf; inputCount=$orderedPaths.Count }
        $merge = Merge-StatusSetPdfWithQpdf -OrderedInputPdfPaths $orderedPaths -OutputPdf $renderPdf -QpdfExe $qpdfExe
        if (-not $merge.IsSuccess) { return $merge }

        if ($atomicReplaceEnabled) {
            $install = _SSS-InstallStatusSetPdf -SourcePdf $renderPdf -OutputPdf $outPdf -AtomicReplace:$true -OperationReport $operationReport -HistoryDir (Join-Path $workspace '_history')
            if (-not $install.IsSuccess) { return $install }
        }

        if ($cleanImmediateExportScratch -and -not $useLocalFs) {
            try { _SSS-RemoveExportDirContentsFileFirst -DirPath $exportWorkDir -RemoveEmptyDir $true } catch { }
        } elseif ($cleanupEnabled) {
            _SSS-CleanupExpiredStatusSetStaging -WorkspaceDir $workspace -RetentionDays $retentionDays -OperationReport $operationReport
            if ($stagingBase -ne $workspace) { _SSS-CleanupExpiredStatusSetStaging -WorkspaceDir $stagingBase -RetentionDays $retentionDays -OperationReport $operationReport }
        } else {
            _SSS-AddStatusSetOperation -Report $operationReport -Kind 'skips' -Operation @{ action='cleanup'; reason='cleanup-disabled'; stagingRoot=$stagingBase }
        }

        # Pause before manifest replace: AV often scans the merged PDF in the same folder.
        _SSS-FsThrottle

        $manObj = New-StatusSetManifestObject -State $fullState -SheetsFolderDisplay $sourceFolder
        _SSS-AddStatusSetOperation -Report $operationReport -Kind 'writes' -Operation @{ action='write-manifest'; path=$manifestPath }
        $mw = Write-StatusSetManifestFile -Path $manifestPath -Manifest $manObj
        if (-not $mw.IsSuccess) { return $mw }

        $uploadNote = $null
        $uploadError = $null
        if ($writeBackToPW) {
            if ($useLocalFs) {
                _SSS-AddStatusSetOperation -Report $operationReport -Kind 'skips' -Operation @{ action='projectwise-upload'; reason='local-filesystem-source'; documentName=$statusPdfName }
                $uploadNote = 'SKIPPED_LOCAL_SOURCE_FOLDER'
            } elseif ($pwConnected) {
                $up = Get-Command -Name Update-PWDocumentFile -ErrorAction SilentlyContinue
                $find = Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue
                if ($up -and $find) {
                    try {
                        $existing = Get-PWDocumentsBySearch -FolderPath $pwPath -JustThisFolder -DocumentName $statusPdfName -PopulatePath -ErrorAction SilentlyContinue | Select-Object -First 1
                        if ($existing) {
                            _SSS-AddStatusSetOperation -Report $operationReport -Kind 'writes' -Operation @{ action='projectwise-upload'; mode='update'; folder=$pwPath; documentName=$statusPdfName }
                            # PW upload is a single-file write to the server, not a local
                            # delete/move; AV throttle does not apply here.
                            _SSS-UpdatePWDocumentFileFromDisk -InputDocument $existing -LocalFilePath $outPdf
                            $uploadNote = 'UPDATED'
                        } else {
                            $newCmd = Get-Command -Name New-PWDocument -ErrorAction SilentlyContinue
                            if ($newCmd) {
                                _SSS-AddStatusSetOperation -Report $operationReport -Kind 'writes' -Operation @{ action='projectwise-upload'; mode='create'; folder=$pwPath; documentName=$statusPdfName }
                                New-PWDocument -FolderPath $pwPath -FilePath $outPdf -DocumentName $statusPdfName -ErrorAction Stop | Out-Null
                                $uploadNote = 'CREATED'
                            } else {
                                $uploadNote = 'SKIPPED_NO_New-PWDocument'
                            }
                        }
                    } catch {
                        $uploadNote = 'FAILED'
                        $uploadError = [string]$_.Exception.Message
                    }
                } else {
                    $uploadNote = 'SKIPPED_MISSING_CMDLETS'
                }
            } else {
                $uploadNote = 'SKIPPED_PW_DISCONNECTED'
            }
        }

        # If writeBackToPW=true and the upload was attempted-and-threw, fail the
        # job. Returning success with pwUpload='FAILED' silently strands the
        # work: the local _StatusSet.pdf exists but PW never got it, and the
        # job moves to succeeded\ where the watcher would never re-enqueue it.
        # By failing, the worker re-enqueues to pending\ (or to failed\ at
        # maxAttempts) and the dashboard surfaces the actual cause.
        if ($uploadNote -eq 'FAILED') {
            return New-QCFailureResult -Code 'STATUS_SET_PW_UPLOAD_FAILED' -Message ('PW write-back failed: ' + $uploadError) -Data @{
                jobId        = [string]$Job.id
                workspaceDir = $workspace
                outputPdf    = $outPdf
                manifestPath = $manifestPath
                mergedInputs = $orderedPaths.Count
                writeBackToPW= $writeBackToPW
                pwUpload     = $uploadNote
                error        = $uploadError
                pwFolder     = $pwPath
                useLocalFs   = $useLocalFs
                operationReport = if ($dryRunOperationReport) { $operationReport } else { $null }
            }
        }

        return New-QCSuccessResult -Code 'STATUS_SET_OK' -Message 'Status set PDF regenerated.' -Data @{
            jobId = [string]$Job.id
            workspaceDir = $workspace
            outputPdf = $outPdf
            manifestPath = $manifestPath
            mergedInputs = $orderedPaths.Count
            writeBackToPW = $writeBackToPW
            pwUpload = $uploadNote
            compare = $cmp.Data
            useLocalFs = $useLocalFs
            pwExportReuseCount = $pwExportReuseCount
            pwExportFreshCount = $pwExportFreshCount
            pwSheetReusePolicy = $pwSheetReusePolicy
            operationReport = if ($dryRunOperationReport) { $operationReport } else { $null }
        }
    } finally {
        if ($pwConnected) {
            try { Disconnect-PW | Out-Null } catch { }
        }
    }
}

function Get-StatusSetWorkspaceManifests {
    <#
    .SYNOPSIS
    Walk a status-set localRoot and return one record per workspace that contains
    both a manifest and a local _StatusSet.pdf.
    .DESCRIPTION
    The active processor stores each per-folder workspace as a 16-char-hex
    directory under localRoot, containing the manifest JSON and the merged
    _StatusSet.pdf. The reconcile pass needs to enumerate these workspaces to
    decide which ones still need to be uploaded to ProjectWise.
    .PARAMETER LocalRoot
    The statusSet.localRoot from appsettings.json.
    .PARAMETER ManifestFileName
    Manifest filename inside each workspace. Defaults to '_statusset.manifest.json'.
    .PARAMETER StatusSetPdfName
    Local merged PDF filename. Defaults to '_StatusSet.pdf' to match appsettings.json.
    .OUTPUTS
    Array of hashtables: { workspaceDir; manifestPath; outputPdf; manifest;
                            sheetsFolder; pwPath; outputPdfLastWriteUtc }.
    Workspaces missing either the manifest or the PDF are skipped and reported
    in Data.skipped (with reason).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LocalRoot,
        [Parameter(Mandatory = $false)]
        [string]$ManifestFileName = '_statusset.manifest.json',
        [Parameter(Mandatory = $false)]
        [string]$StatusSetPdfName = '_StatusSet.pdf'
    )

    $records = @()
    $skipped = @()

    if (_SSS-IsNullOrWhiteSpace $LocalRoot) {
        return New-QCFailureResult -Code 'STATUS_SET_LOCAL_ROOT_EMPTY' -Message 'LocalRoot is required.' -Data @{}
    }
    if (-not (Test-Path -LiteralPath $LocalRoot -PathType Container)) {
        return New-QCSuccessResult -Code 'STATUS_SET_LOCAL_ROOT_MISSING' -Message 'LocalRoot does not exist.' -Data @{
            localRoot = $LocalRoot; records = @(); skipped = @()
        }
    }

    foreach ($d in @(Get-ChildItem -LiteralPath $LocalRoot -Directory -ErrorAction SilentlyContinue)) {
        $manifestPath = Join-Path $d.FullName $ManifestFileName
        $outPdf = Join-Path $d.FullName $StatusSetPdfName
        if (-not (Test-Path -LiteralPath $manifestPath)) {
            $skipped += @{ workspaceDir = $d.FullName; reason = 'NO_MANIFEST' }
            continue
        }
        if (-not (Test-Path -LiteralPath $outPdf)) {
            $skipped += @{ workspaceDir = $d.FullName; reason = 'NO_LOCAL_PDF' }
            continue
        }
        $r = Read-StatusSetManifestFile -Path $manifestPath
        if (-not $r.IsSuccess -or -not $r.Data.manifest) {
            $skipped += @{ workspaceDir = $d.FullName; reason = 'MANIFEST_UNREADABLE' }
            continue
        }
        $man = $r.Data.manifest
        $sheetsFolder = ''
        try { if ($man.ContainsKey('sheetsFolder')) { $sheetsFolder = [string]$man.sheetsFolder } } catch { }
        if (_SSS-IsNullOrWhiteSpace $sheetsFolder) {
            $skipped += @{ workspaceDir = $d.FullName; reason = 'MANIFEST_MISSING_SHEETS_FOLDER' }
            continue
        }
        $pwPath = $sheetsFolder -replace '^Documents\\', ''
        $pdfInfo = Get-Item -LiteralPath $outPdf -ErrorAction SilentlyContinue
        $localMtime = if ($pdfInfo) { $pdfInfo.LastWriteTimeUtc } else { $null }

        $records += @{
            workspaceDir          = $d.FullName
            manifestPath          = $manifestPath
            outputPdf             = $outPdf
            manifest              = $man
            sheetsFolder          = $sheetsFolder
            pwPath                = $pwPath
            outputPdfLastWriteUtc = $localMtime
            outputPdfSize         = if ($pdfInfo) { [long]$pdfInfo.Length } else { 0 }
        }
    }

    return New-QCSuccessResult -Code 'STATUS_SET_WORKSPACES_OK' -Message 'Walked status set workspaces.' -Data @{
        localRoot   = $LocalRoot
        records     = $records
        skipped     = $skipped
        recordCount = $records.Count
        skipCount   = $skipped.Count
    }
}

function Sync-StatusSetWorkspaceToPw {
    <#
    .SYNOPSIS
    Reconcile a single status-set workspace's _StatusSet.pdf against ProjectWise.
    .DESCRIPTION
    Compares the local _StatusSet.pdf against the PW copy in the manifest's
    folder. Decision matrix:
      - PW has no _StatusSet.pdf and local PDF exists  -> CREATED  (New-PWDocument)
      - PW has it and local PDF mtime > PW lastModified -> UPDATED (Update-PWDocumentFile)
      - PW has it and is the same/newer                 -> IN_SYNC (no-op)
    Caller is responsible for opening/closing the PW connection.
    .PARAMETER WorkspaceRecord
    A record produced by Get-StatusSetWorkspaceManifests.
    .PARAMETER StatusSetPdfName
    Defaults to '_StatusSet.pdf'.
    .PARAMETER MtimeSkewSeconds
    PW timestamps are second-precision; local file timestamps are sub-second.
    Differences smaller than this are treated as "same". Default 60.
    .OUTPUTS
    Result with Code in:
      STATUS_SET_RECONCILE_IN_SYNC, STATUS_SET_RECONCILE_UPDATED,
      STATUS_SET_RECONCILE_CREATED, STATUS_SET_RECONCILE_SKIPPED,
      STATUS_SET_RECONCILE_FAILED.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$WorkspaceRecord,
        [Parameter(Mandatory = $false)]
        [string]$StatusSetPdfName = '_StatusSet.pdf',
        [Parameter(Mandatory = $false)]
        [int]$MtimeSkewSeconds = 60
    )

    $pwPath = [string]$WorkspaceRecord.pwPath
    $outPdf = [string]$WorkspaceRecord.outputPdf
    $sheetsFolder = [string]$WorkspaceRecord.sheetsFolder

    if (-not (Test-Path -LiteralPath $outPdf)) {
        return New-QCFailureResult -Code 'STATUS_SET_RECONCILE_FAILED' -Message 'Local _StatusSet.pdf not found.' -Data @{
            workspaceDir = [string]$WorkspaceRecord.workspaceDir; outputPdf = $outPdf
        }
    }

    function _SSS-TryImportPWModules {
        # Reconcile needs pwps_dab cmdlets (Get-PWDocumentsBySearch/Update-PWDocumentFile/New-PWDocument).
        # A process can have Open-PWConnection (pwps) loaded but still be missing pwps_dab.
        if (Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue) { return $true }

        # Prefer pwps_dab first (it may auto-load pwps as a dependency).
        foreach ($name in @('pwps_dab', 'pwps')) {
            try { Import-Module $name -Force -ErrorAction Stop | Out-Null } catch { }
            if (Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue) { return $true }
        }

        # Common explicit path (Bentley install) if modules aren't in PSModulePath.
        $pwpsPath = 'C:\Program Files (x86)\Bentley\ProjectWise\bin\PowerShell\pwps\pwps.psd1'
        if (Test-Path -LiteralPath $pwpsPath) {
            try { Import-Module $pwpsPath -Force -ErrorAction Stop | Out-Null } catch { }
        }
        # After pwps.psd1 import, pwps_dab may still not be present; we still check the cmdlet.
        return [bool](Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue)
    }

    $null = _SSS-TryImportPWModules

    $find = Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue
    $up   = Get-Command -Name Update-PWDocumentFile  -ErrorAction SilentlyContinue
    $newC = Get-Command -Name New-PWDocument         -ErrorAction SilentlyContinue
    if (-not $find -or -not $up) {
        return New-QCFailureResult -Code 'STATUS_SET_RECONCILE_FAILED' -Message 'pwps_dab/pwps cmdlets not loaded in this PowerShell process.' -Data @{
            haveFind     = [bool]$find
            haveUpdate   = [bool]$up
            haveNew      = [bool]$newC
            psModulePath = [string]$env:PSModulePath
            is64Bit      = [Environment]::Is64BitProcess
        }
    }

    $existing = $null
    try {
        $existing = Get-PWDocumentsBySearch -FolderPath $pwPath -JustThisFolder -DocumentName $StatusSetPdfName -PopulatePath -ErrorAction SilentlyContinue | Select-Object -First 1
    } catch {
        return New-QCFailureResult -Code 'STATUS_SET_RECONCILE_FAILED' -Message ('PW search failed: ' + $_.Exception.Message) -Data @{
            workspaceDir = [string]$WorkspaceRecord.workspaceDir; pwFolder = $pwPath
        }
    }

    $pdfInfo = Get-Item -LiteralPath $outPdf -ErrorAction SilentlyContinue
    $localMtime = if ($pdfInfo) { $pdfInfo.LastWriteTimeUtc } else { $null }

    if (-not $existing) {
        if (-not $newC) {
            return New-QCFailureResult -Code 'STATUS_SET_RECONCILE_FAILED' -Message 'PW has no _StatusSet.pdf and New-PWDocument cmdlet is missing.' -Data @{
                workspaceDir = [string]$WorkspaceRecord.workspaceDir; pwFolder = $pwPath
            }
        }
        try {
            New-PWDocument -FolderPath $pwPath -FilePath $outPdf -DocumentName $StatusSetPdfName -ErrorAction Stop | Out-Null
        } catch {
            return New-QCFailureResult -Code 'STATUS_SET_RECONCILE_FAILED' -Message ('New-PWDocument failed: ' + $_.Exception.Message) -Data @{
                workspaceDir = [string]$WorkspaceRecord.workspaceDir; pwFolder = $pwPath; outputPdf = $outPdf
            }
        }
        return New-QCSuccessResult -Code 'STATUS_SET_RECONCILE_CREATED' -Message 'Uploaded missing _StatusSet.pdf to PW.' -Data @{
            workspaceDir = [string]$WorkspaceRecord.workspaceDir
            pwFolder = $pwPath; sheetsFolder = $sheetsFolder
            outputPdf = $outPdf; localMtime = if ($localMtime) { ConvertTo-QCTimestamp $localMtime } else { $null }
            action = 'CREATED'
        }
    }

    $pwMtime = _SSS-GetDocLastModified $existing
    $needsUpdate = $true
    if ($pwMtime -and $localMtime) {
        $delta = ($localMtime - $pwMtime).TotalSeconds
        if ($delta -le $MtimeSkewSeconds) { $needsUpdate = $false }
    }
    if (-not $needsUpdate) {
        return New-QCSuccessResult -Code 'STATUS_SET_RECONCILE_IN_SYNC' -Message 'PW _StatusSet.pdf is up-to-date.' -Data @{
            workspaceDir = [string]$WorkspaceRecord.workspaceDir
            pwFolder = $pwPath; sheetsFolder = $sheetsFolder
            outputPdf = $outPdf
            localMtime = if ($localMtime) { ConvertTo-QCTimestamp $localMtime } else { $null }
            pwMtime    = if ($pwMtime)    { ConvertTo-QCTimestamp $pwMtime    } else { $null }
            action = 'IN_SYNC'
        }
    }

    try {
        _SSS-UpdatePWDocumentFileFromDisk -InputDocument $existing -LocalFilePath $outPdf
    } catch {
        return New-QCFailureResult -Code 'STATUS_SET_RECONCILE_FAILED' -Message ('Update-PWDocumentFile failed: ' + $_.Exception.Message) -Data @{
            workspaceDir = [string]$WorkspaceRecord.workspaceDir; pwFolder = $pwPath; outputPdf = $outPdf
        }
    }
    return New-QCSuccessResult -Code 'STATUS_SET_RECONCILE_UPDATED' -Message 'PW _StatusSet.pdf updated to match local copy.' -Data @{
        workspaceDir = [string]$WorkspaceRecord.workspaceDir
        pwFolder = $pwPath; sheetsFolder = $sheetsFolder
        outputPdf = $outPdf
        localMtime = if ($localMtime) { ConvertTo-QCTimestamp $localMtime } else { $null }
        pwMtime    = if ($pwMtime)    { ConvertTo-QCTimestamp $pwMtime    } else { $null }
        action = 'UPDATED'
    }
}

function Invoke-StatusSetReconcile {
    <#
    .SYNOPSIS
    On startup, sync every locally-built _StatusSet.pdf back to ProjectWise.
    .DESCRIPTION
    Walks statusSet.localRoot for workspaces that have both a manifest and a
    local _StatusSet.pdf. For each, compares to the PW copy and uploads when
    out-of-sync (legacy parity: every restart re-checks every manifest). Caller
    is responsible for opening/closing the PW connection so this can be wired
    into Watch-QCTrigger or run as a standalone script.
    .PARAMETER Config
    Loaded appsettings hashtable.
    .PARAMETER LogCallback
    Optional ScriptBlock invoked once per workspace with a hashtable describing
    the action taken. Lets callers stream structured events to dashboards/logs.
    .OUTPUTS
    Result with aggregate counts: created, updated, inSync, failed, skipped.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory = $false)]
        [scriptblock]$LogCallback
    )

    $ss = @{}
    if ($Config.ContainsKey('statusSet') -and $Config.statusSet) {
        $raw = $Config.statusSet
        if ($raw -is [hashtable]) { $ss = $raw } elseif ($raw.PSObject) {
            foreach ($p in $raw.PSObject.Properties) { $ss[$p.Name] = $p.Value }
        }
    }
    $localRoot = if ($ss.ContainsKey('localRoot') -and $ss.localRoot) { [string]$ss.localRoot } else { 'C:\PW_QC_LOCAL' }
    $manifestName = if ($ss.ContainsKey('manifestFileName') -and $ss.manifestFileName) { [string]$ss.manifestFileName } else { '_statusset.manifest.json' }
    $statusPdfName = if ($ss.ContainsKey('statusSetPdfName') -and $ss.statusSetPdfName) { [string]$ss.statusSetPdfName } else { '_StatusSet.pdf' }

    $walk = Get-StatusSetWorkspaceManifests -LocalRoot $localRoot -ManifestFileName $manifestName -StatusSetPdfName $statusPdfName
    if (-not $walk.IsSuccess) { return $walk }

    $records = @($walk.Data.records)
    $counts = [ordered]@{
        considered = $records.Count
        created    = 0
        updated    = 0
        inSync     = 0
        failed     = 0
        skipped    = [int]$walk.Data.skipCount
    }
    $failures = @()
    foreach ($rec in $records) {
        $r = Sync-StatusSetWorkspaceToPw -WorkspaceRecord $rec -StatusSetPdfName $statusPdfName
        switch ($r.Code) {
            'STATUS_SET_RECONCILE_CREATED' { $counts.created++ }
            'STATUS_SET_RECONCILE_UPDATED' { $counts.updated++ }
            'STATUS_SET_RECONCILE_IN_SYNC' { $counts.inSync++ }
            default {
                $counts.failed++
                $failures += @{
                    workspaceDir = [string]$rec.workspaceDir
                    pwFolder     = [string]$rec.pwPath
                    code         = [string]$r.Code
                    message      = [string]$r.Message
                }
            }
        }
        if ($LogCallback) {
            try {
                $payload = @{
                    code         = [string]$r.Code
                    isSuccess    = [bool]$r.IsSuccess
                    message      = [string]$r.Message
                    workspaceDir = [string]$rec.workspaceDir
                    pwFolder     = [string]$rec.pwPath
                    sheetsFolder = [string]$rec.sheetsFolder
                    outputPdf    = [string]$rec.outputPdf
                    data         = $r.Data
                }
                & $LogCallback $payload
            } catch { }
        }
    }
    return New-QCSuccessResult -Code 'STATUS_SET_RECONCILE_DONE' -Message 'Status set reconciliation completed.' -Data @{
        localRoot = $localRoot
        counts    = $counts
        failures  = $failures
        skipped   = $walk.Data.skipped
    }
}

function Test-StatusSetWatcherShouldEnqueue {
    <#
    .SYNOPSIS
    Watch-QCTrigger gate: enqueue STATUS_SET_GEN only if work is needed vs local manifest + _statusset.pdf.
    .DESCRIPTION
    Uses the same Test-StatusSetRebuildNeeded rules as Invoke-StatusSetNativeJob (hash, sheet drift,
    _statusset.pdf vs newest paired sheet PDF). When shouldEnqueue is $false, the watcher should not
    call Add-QCQueueJob so the worker is not bothered with STATUS_SET_UP_TO_DATE jobs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$SourceFolder,
        [Parameter(Mandatory)]
        [hashtable]$FolderState
    )

    $ss = @{}
    if ($Config.ContainsKey('statusSet') -and $Config.statusSet) {
        $raw = $Config.statusSet
        if ($raw -is [hashtable]) { $ss = $raw } elseif ($raw.PSObject) {
            foreach ($p in $raw.PSObject.Properties) { $ss[$p.Name] = $p.Value }
        }
    }

    $localRoot = if ($ss.ContainsKey('localRoot') -and $ss.localRoot) { [string]$ss.localRoot } else { 'C:\PW_QC_LOCAL' }
    $manifestName = if ($ss.ContainsKey('manifestFileName') -and $ss.manifestFileName) { [string]$ss.manifestFileName } else { '_statusset.manifest.json' }
    $statusPdfName = if ($ss.ContainsKey('statusSetPdfName') -and $ss.statusSetPdfName) { [string]$ss.statusSetPdfName } else { '_statusset.pdf' }
    $forceRebuild = $false
    if ($ss.ContainsKey('forceRebuild')) { try { $forceRebuild = [bool]$ss.forceRebuild } catch { $forceRebuild = $false } }

    try {
        $pc = [int]$FolderState.pairedCount
    } catch {
        $pc = 0
    }
    if ($pc -le 0) {
        return New-QCSuccessResult -Code 'WATCH_STATUS_SET_GATE' -Message 'No paired sheets in folder state.' -Data @{
            shouldEnqueue = $false
            gateReason = 'NO_PAIRS'
        }
    }

    $wsRes = Get-StatusSetWorkspaceDirectory -LocalRoot $localRoot -SheetsFolderPath $SourceFolder
    if (-not $wsRes.IsSuccess) {
        return New-QCSuccessResult -Code 'WATCH_STATUS_SET_GATE' -Message 'Could not resolve workspace; allow enqueue.' -Data @{
            shouldEnqueue = $true
            gateReason = 'WORKSPACE_PATH_FAILED'
            workspaceError = [string]$wsRes.Message
        }
    }

    $workspace = [string]$wsRes.Data.workspaceDir
    $manifestPath = Get-StatusSetManifestPath -WorkspaceDir $workspace -ManifestFileName $manifestName
    $outPdf = Join-Path $workspace $statusPdfName

    $readM = Read-StatusSetManifestFile -Path $manifestPath
    if (-not $readM.IsSuccess) {
        return New-QCSuccessResult -Code 'WATCH_STATUS_SET_GATE' -Message 'Manifest read error; allow enqueue.' -Data @{
            shouldEnqueue = $true
            gateReason = 'MANIFEST_READ_FAILED'
            workspaceDir = $workspace
            manifestPath = $manifestPath
            readCode = [string]$readM.Code
        }
    }

    $manifest = $readM.Data.manifest

    $cmp = Test-StatusSetRebuildNeeded -Manifest $manifest -CurrentState $FolderState -ForceRebuild:$forceRebuild -StatusSetPdfPath $outPdf
    if (-not $cmp.IsSuccess) {
        return New-QCSuccessResult -Code 'WATCH_STATUS_SET_GATE' -Message 'Compare failed; allow enqueue.' -Data @{
            shouldEnqueue = $true
            gateReason = 'COMPARE_FAILED'
            compareError = [string]$cmp.Message
            workspaceDir = $workspace
        }
    }

    $needs = [bool]$cmp.Data.needsRebuild
    return New-QCSuccessResult -Code 'WATCH_STATUS_SET_GATE' -Message $(if ($needs) { 'Enqueue: rebuild needed.' } else { 'Skip enqueue: already current.' }) -Data @{
        shouldEnqueue = $needs
        gateReason = if ($needs) { 'WORK_NEEDED' } else { 'ALREADY_CURRENT' }
        workspaceDir = $workspace
        manifestPath = $manifestPath
        statusSetPdfPath = $outPdf
        compare = $cmp.Data
    }
}

Export-ModuleMember -Function @(
    'Get-StatusSetLocalFolderState',
    'Get-StatusSetPWFolderState',
    'Get-StatusSetWorkspaceDirectory',
    'Get-StatusSetManifestPath',
    'Read-StatusSetManifestFile',
    'Write-StatusSetManifestFile',
    'New-StatusSetManifestObject',
    'Test-StatusSetRebuildNeeded',
    'Test-StatusSetSheetPdfTimestampMatch',
    'Get-StatusSetPairedSheetSignature',
    'Test-StatusSetWatcherShouldEnqueue',
    'Merge-StatusSetPdfWithQpdf',
    'Export-StatusSetPdfToFolder',
    'Invoke-StatusSetNativeJob',
    'Get-StatusSetWorkspaceManifests',
    'Sync-StatusSetWorkspaceToPw',
    'Invoke-StatusSetReconcile'
)
