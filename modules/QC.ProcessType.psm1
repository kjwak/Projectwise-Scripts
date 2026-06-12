# QC.ProcessType.psm1
# Process type normalization, lane PDF resolution, and stamp configuration.

Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'Core.Logging.psm1') -Force -ErrorAction SilentlyContinue

function _QCPT-IsBlank([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCPT-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) { return $Value }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _QCPT-Log([string]$Code, [string]$Level, [string]$Message, [hashtable]$Data) {
    try {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level $Level -Code $Code -Message $Message -Data $Data | Out-Null
        }
    } catch { }
}

function _QCPT-DefaultProcessTypeSettings {
    return @{
        DefaultProcessType = 'production'
        ProcessTypes = @{
            production = @{
                PdfSuffix = 'prod'
                ResetToProductionAfterPrepend = $false
                SyncWithSiblingSheets = $true
                DefaultStamp = 'Production'
            }
            check = @{
                PdfSuffix = 'chk'
                ResetToProductionAfterPrepend = $true
                SyncWithSiblingSheets = $false
                DefaultStamp = 'Check'
            }
            review = @{
                PdfSuffix = 'rev'
                ResetToProductionAfterPrepend = $true
                SyncWithSiblingSheets = $false
                DefaultStamp = 'Review'
            }
        }
        StampProfiles = @{
            Default = @{
                production = 'Production'
                check = 'Check'
                review = 'Review'
            }
        }
        StampAssets = @{
            Production = 'stamps/Production_Stamp.pdf'
            Check = 'stamps/IC_Stamp.pdf'
            Review = 'stamps/Peer_Review_Stamp.pdf'
        }
        RootOverrides = @()
    }
}

function Get-QCProcessTypeSettings {
    <#
    .SYNOPSIS
    Loads and merges QCProcess configuration from appsettings.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Config
    )
    $defaults = _QCPT-DefaultProcessTypeSettings
    $raw = $null
    if ($Config) {
        if ($Config.ContainsKey('QCProcess')) { $raw = _QCPT-ToHashtable $Config.QCProcess }
        elseif ($Config.ContainsKey('qcProcess')) { $raw = _QCPT-ToHashtable $Config.qcProcess }
    }
    if (-not $raw) { return $defaults }

    $merged = @{}
    foreach ($k in $defaults.Keys) { $merged[$k] = $defaults[$k] }
    foreach ($k in $raw.Keys) {
        if ($null -eq $raw[$k]) { continue }
        if ($raw[$k] -is [System.Array]) {
            $merged[$k] = @($raw[$k])
            continue
        }
        if ($raw[$k] -is [hashtable] -or ($raw[$k].PSObject -and $raw[$k].PSObject.Properties)) {
            $sub = _QCPT-ToHashtable $merged[$k]
            if (-not $sub) { $sub = @{} }
            $src = _QCPT-ToHashtable $raw[$k]
            foreach ($sk in $src.Keys) { $sub[$sk] = $src[$sk] }
            $merged[$k] = $sub
        } else {
            $merged[$k] = $raw[$k]
        }
    }
    return $merged
}

function Normalize-QCProcessType {
    <#
    .SYNOPSIS
    Normalizes QC process type values to production, check, or review.
    Returns $null for unknown values (with telemetry).
    #>
    [CmdletBinding()]
    param(
        [string]$ProcessType = '',
        [string]$ReviewType = '',
        [hashtable]$Context = $null,
        [switch]$AllowNullOnEmpty
    )

    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not (_QCPT-IsBlank $ProcessType)) { [void]$candidates.Add([string]$ProcessType) }
    if (-not (_QCPT-IsBlank $ReviewType)) { [void]$candidates.Add([string]$ReviewType) }

    if ($Context) {
        foreach ($key in @('qc_process_type', 'qcProcessType', 'processType', 'qc_review_type', 'qcReviewType', 'reviewType')) {
            if ($Context.ContainsKey($key) -and -not (_QCPT-IsBlank $Context[$key])) {
                [void]$candidates.Add([string]$Context[$key])
            }
        }
    }

    foreach ($raw in $candidates) {
        $norm = ([string]$raw).Trim().ToLowerInvariant()
        switch -Regex ($norm) {
            '^production(\s+qc)?$|^production$|^qc$' { return 'production' }
            '^check$|^independent(\s+(check|review))?$|^independent_check$|^independent$|^ic$' { return 'check' }
            '^review$|^peer(\s+review)?$|^peer_review$|^peer$' { return 'review' }
            default { }
        }
    }

    if ($AllowNullOnEmpty -and $candidates.Count -eq 0) { return $null }

    $inputValue = if ($candidates.Count -gt 0) { $candidates[0] } else { '' }
    _QCPT-Log -Code 'QC_PROCESS_TYPE_UNKNOWN' -Level 'Warning' `
        -Message 'Unknown QC process type value; normalization failed.' `
        -Data @{ inputValue = $inputValue; candidateCount = $candidates.Count }
    return $null
}

function Get-QCProcessTypeDisplayLabel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ProcessType
    )
    switch (([string]$ProcessType).Trim().ToLowerInvariant()) {
        'production' { return 'Production' }
        'check' { return 'Check' }
        'review' { return 'Review' }
        default { return [string]$ProcessType }
    }
}

function Get-QCProcessTypePdfSuffix {
    <#
    .SYNOPSIS
    Maps normalized process type to PDF lane suffix (prod, chk, rev).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ProcessType,
        [hashtable]$Config = $null
    )
    $normalized = Normalize-QCProcessType -ProcessType $ProcessType
    if (-not $normalized) { return $null }

    $settings = Get-QCProcessTypeSettings -Config $Config
    $pt = _QCPT-ToHashtable $settings.ProcessTypes
    if ($pt -and $pt.ContainsKey($normalized)) {
        $entry = _QCPT-ToHashtable $pt[$normalized]
        if ($entry -and $entry.PdfSuffix) { return [string]$entry.PdfSuffix }
        if ($entry -and $entry.pdfSuffix) { return [string]$entry.pdfSuffix }
    }

    switch ($normalized) {
        'production' { return 'prod' }
        'check' { return 'chk' }
        'review' { return 'rev' }
        default { return $null }
    }
}

function Test-QCProcessTypeSyncsWithSiblingSheets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ProcessType,
        [hashtable]$Config = $null
    )
    $normalized = Normalize-QCProcessType -ProcessType $ProcessType
    if (-not $normalized) { return $true }
    $settings = Get-QCProcessTypeSettings -Config $Config
    $pt = _QCPT-ToHashtable $settings.ProcessTypes
    if ($pt -and $pt.ContainsKey($normalized)) {
        $entry = _QCPT-ToHashtable $pt[$normalized]
        if ($entry -and $entry.ContainsKey('SyncWithSiblingSheets')) {
            try { return [bool]$entry.SyncWithSiblingSheets } catch { }
        }
        if ($entry -and $entry.ContainsKey('syncWithSiblingSheets')) {
            try { return [bool]$entry.syncWithSiblingSheets } catch { }
        }
    }
    return ($normalized -eq 'production')
}

function Test-QCProcessTypeResetsAfterPrepend {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ProcessType,
        [hashtable]$Config = $null
    )
    $normalized = Normalize-QCProcessType -ProcessType $ProcessType
    if (-not $normalized) { return $false }
    $settings = Get-QCProcessTypeSettings -Config $Config
    $pt = _QCPT-ToHashtable $settings.ProcessTypes
    if ($pt -and $pt.ContainsKey($normalized)) {
        $entry = _QCPT-ToHashtable $pt[$normalized]
        if ($entry -and $entry.ContainsKey('ResetToProductionAfterPrepend')) {
            try { return [bool]$entry.ResetToProductionAfterPrepend } catch { }
        }
        if ($entry -and $entry.ContainsKey('resetToProductionAfterPrepend')) {
            try { return [bool]$entry.resetToProductionAfterPrepend } catch { }
        }
    }
    return ($normalized -in @('check', 'review'))
}

function Get-PWQcPdfLaneFromDocumentName {
    <#
    .SYNOPSIS
    Returns production, check, review, or $null for non-lane PDF names.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DocumentName
    )
    if ($DocumentName -notmatch '(?i)\.pdf$') { return $null }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
    if ([string]::IsNullOrWhiteSpace($stem)) { return $null }
    if ($stem -match '(?i)-prod$') { return 'production' }
    if ($stem -match '(?i)-chk$') { return 'check' }
    if ($stem -match '(?i)-rev$') { return 'review' }
    return $null
}

function Test-PWQcPdfLaneSuffix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DocumentName
    )
    return (-not (_QCPT-IsBlank (Get-PWQcPdfLaneFromDocumentName -DocumentName $DocumentName)))
}

function Get-QCLaneQcPdfExpectedName {
    <#
    .SYNOPSIS
    Builds expected lane QC PDF filename: {base}-{suffix}.pdf
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SheetBaseName,
        [Parameter(Mandatory)][string]$ProcessType,
        [hashtable]$Config = $null
    )
    $suffix = Get-QCProcessTypePdfSuffix -ProcessType $ProcessType -Config $Config
    if (-not $suffix) { return $null }
    $base = ([string]$SheetBaseName).Trim()
    if ([string]::IsNullOrWhiteSpace($base)) { return $null }
    return ($base + '-' + $suffix + '.pdf')
}

function Resolve-QCLaneQcPdf {
    <#
    .SYNOPSIS
    Resolves lane-specific QC PDF in a ProjectWise folder by process type.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SheetBaseName,
        [Parameter(Mandatory)][string]$ProcessType,
        [string]$WatchRoot = ''
    )

    $normalized = Normalize-QCProcessType -ProcessType $ProcessType
    if (-not $normalized) {
        return @{
            IsSuccess = $false
            Code = 'QC_PROCESS_TYPE_UNKNOWN'
            Message = 'Cannot resolve lane QC PDF without a valid process type.'
            Data = @{ processType = $ProcessType }
        }
    }

    $suffix = Get-QCProcessTypePdfSuffix -ProcessType $normalized -Config $Config
    $expectedName = Get-QCLaneQcPdfExpectedName -SheetBaseName $SheetBaseName -ProcessType $normalized -Config $Config
    if (-not $expectedName) {
        return @{
            IsSuccess = $false
            Code = 'QC_LANE_PDF_INVALID'
            Message = 'Could not build expected lane QC PDF name.'
            Data = @{ sheetBaseName = $SheetBaseName; processType = $normalized }
        }
    }

    $searchCmd = Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction SilentlyContinue
    if (-not $searchCmd) {
        return @{
            IsSuccess = $false
            Code = 'QC_LANE_PDF_SEARCH_UNAVAILABLE'
            Message = 'Get-PWDocumentsBySearch is not available.'
            Data = @{ expectedName = $expectedName }
        }
    }

    $apiPath = $FolderPath
    if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
        $converted = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
        if (-not (_QCPT-IsBlank $converted)) { $apiPath = $converted }
    }

    $matches = @()
    try {
        $params = @{
            FolderPath = $apiPath
            JustThisFolder = $true
            DocumentName = $expectedName
            ErrorAction = 'Stop'
        }
        if ($searchCmd.Parameters.ContainsKey('PopulatePath')) { $params['PopulatePath'] = $true }
        $matches = @(& $searchCmd @params)
    } catch {
        $matches = @()
    }

    if ($matches.Count -eq 0) {
        _QCPT-Log -Code 'QC_LANE_PDF_MISSING' -Level 'Warning' `
            -Message 'Expected lane QC PDF not found in folder.' `
            -Data @{
                qc_process_type = $normalized
                pdfSuffix = $suffix
                expectedName = $expectedName
                folderPath = $FolderPath
                watchRoot = $WatchRoot
            }
        return @{
            IsSuccess = $false
            Code = 'QC_LANE_PDF_MISSING'
            Message = "Lane QC PDF not found: $expectedName"
            Data = @{
                qc_process_type = $normalized
                pdfSuffix = $suffix
                expectedName = $expectedName
                folderPath = $FolderPath
            }
        }
    }

    if ($matches.Count -gt 1) {
        _QCPT-Log -Code 'QC_LANE_PDF_AMBIGUOUS' -Level 'Warning' `
            -Message 'Multiple lane QC PDFs matched expected name.' `
            -Data @{
                qc_process_type = $normalized
                pdfSuffix = $suffix
                expectedName = $expectedName
                matchCount = $matches.Count
            }
        return @{
            IsSuccess = $false
            Code = 'QC_LANE_PDF_AMBIGUOUS'
            Message = "Ambiguous lane QC PDF matches for: $expectedName"
            Data = @{
                qc_process_type = $normalized
                expectedName = $expectedName
                matchCount = $matches.Count
            }
        }
    }

    $doc = $matches[0]
    $guid = ''
    foreach ($prop in @('DocumentGUID', 'DocumentGuid', 'GUID', 'Guid')) {
        try {
            if ($doc.PSObject.Properties[$prop] -and $doc.$prop) {
                $guid = [string]$doc.$prop
                break
            }
        } catch { }
    }

    _QCPT-Log -Code 'QC_LANE_PDF_RESOLVED' -Level 'Information' `
        -Message 'Lane QC PDF resolved.' `
        -Data @{
            qc_process_type = $normalized
            pdfSuffix = $suffix
            expectedName = $expectedName
            resolvedGuid = $guid
            folderPath = $FolderPath
        }

    return @{
        IsSuccess = $true
        Code = 'QC_LANE_PDF_RESOLVED'
        Message = 'Lane QC PDF resolved.'
        Data = @{
            qc_process_type = $normalized
            pdfSuffix = $suffix
            expectedName = $expectedName
            documentName = $expectedName
            documentGuid = $guid
            document = $doc
        }
    }
}

function _QCPT-NormalizeRootPath([string]$Path) {
    if (_QCPT-IsBlank $Path) { return '' }
    $p = ([string]$Path).Trim().Replace('\', '/').TrimEnd('/')
    return $p.ToLowerInvariant()
}

function _QCPT-ResolveWatchlistRoot {
    param(
        [hashtable]$Config,
        [string]$FolderPath
    )
    if (_QCPT-IsBlank $FolderPath) { return '' }
    $pw = _QCPT-ToHashtable $Config.projectWise
    if (-not $pw) { return '' }
    $wl = _QCPT-ToHashtable $pw.watchList
    if (-not $wl) { return '' }
    $roots = @($wl.roots)
    $folderNorm = _QCPT-NormalizeRootPath $FolderPath
    $best = ''
    $bestLen = -1
    foreach ($root in $roots) {
        $r = _QCPT-ToHashtable $root
        if ((-not $r) -or (_QCPT-IsBlank $r.path)) { continue }
        $rootNorm = _QCPT-NormalizeRootPath ([string]$r.path)
        if ($folderNorm.StartsWith($rootNorm) -and $rootNorm.Length -gt $bestLen) {
            $best = [string]$r.path
            $bestLen = $rootNorm.Length
        }
    }
    return $best
}

function _QCPT-ResolveStampAssetPath {
    param(
        [hashtable]$Settings,
        [string]$StampName,
        [string]$RepoRoot
    )
    if (_QCPT-IsBlank $StampName) { return $null }
    $assets = _QCPT-ToHashtable $Settings.StampAssets
    if (-not $assets) { $assets = _QCPT-ToHashtable (_QCPT-DefaultProcessTypeSettings).StampAssets }
    $rel = $null
    if ($assets.ContainsKey($StampName)) { $rel = [string]$assets[$StampName] }
    if (_QCPT-IsBlank $rel) { return $null }
    if (-not [System.IO.Path]::IsPathRooted($rel)) {
        if (_QCPT-IsBlank $RepoRoot) { $RepoRoot = Split-Path -Parent $PSScriptRoot }
        $rel = Join-Path $RepoRoot $rel
    }
    return $rel
}

function Resolve-QCStampForProcess {
    <#
    .SYNOPSIS
    Resolves stamp profile and stamp asset path for a process type using root overrides.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$ProcessType,
        [string]$FolderPath = '',
        [string]$WatchRoot = '',
        [string]$RepoRoot = ''
    )

    $normalized = Normalize-QCProcessType -ProcessType $ProcessType
    if (-not $normalized) {
        _QCPT-Log -Code 'QC_STAMP_UNRESOLVED' -Level 'Warning' `
            -Message 'Stamp resolution failed: unknown process type.' `
            -Data @{ processType = $ProcessType }
        return @{
            IsSuccess = $false
            Code = 'QC_STAMP_UNRESOLVED'
            Message = 'Unknown process type for stamp resolution.'
            resolvedStampProfile = $null
            resolvedStampName = $null
            stampPath = $null
            usedFallback = $false
        }
    }

    $settings = Get-QCProcessTypeSettings -Config $Config
    if (_QCPT-IsBlank $WatchRoot) { $WatchRoot = _QCPT-ResolveWatchlistRoot -Config $Config -FolderPath $FolderPath }

    $candidates = @()
    if (-not (_QCPT-IsBlank $FolderPath)) { $candidates += @{ path = $FolderPath; kind = 'folder' } }
    if (-not (_QCPT-IsBlank $WatchRoot)) { $candidates += @{ path = $WatchRoot; kind = 'watchlist' } }

    $overrides = @($settings.RootOverrides)
    $matchedOverride = $null
    $bestLen = -1
    foreach ($cand in $candidates) {
        $candNorm = _QCPT-NormalizeRootPath $cand.path
        foreach ($ov in $overrides) {
            $o = _QCPT-ToHashtable $ov
            if ((-not $o) -or (_QCPT-IsBlank $o.RootPath)) { continue }
            $rootNorm = _QCPT-NormalizeRootPath ([string]$o.RootPath)
            if ($candNorm.StartsWith($rootNorm) -and $rootNorm.Length -gt $bestLen) {
                $matchedOverride = $o
                $bestLen = $rootNorm.Length
            }
        }
    }

    $profileName = 'Default'
    $usedFallback = $false
    if ($matchedOverride -and $matchedOverride.StampProfile) {
        $profileName = [string]$matchedOverride.StampProfile
    }

    $profiles = _QCPT-ToHashtable $settings.StampProfiles
    $profile = $null
    if ($profiles -and $profiles.ContainsKey($profileName)) {
        $profile = _QCPT-ToHashtable $profiles[$profileName]
    }
    if (-not $profile -and $profileName -ne 'Default') {
        $usedFallback = $true
        $profileName = 'Default'
        if ($profiles -and $profiles.ContainsKey('Default')) {
            $profile = _QCPT-ToHashtable $profiles['Default']
        }
    }

    $stampName = $null
    if ($profile -and $profile.ContainsKey($normalized)) { $stampName = [string]$profile[$normalized] }
    if (_QCPT-IsBlank $stampName) {
        $usedFallback = $true
        $pt = _QCPT-ToHashtable $settings.ProcessTypes
        if ($pt -and $pt.ContainsKey($normalized)) {
            $entry = _QCPT-ToHashtable $pt[$normalized]
            if ($entry -and $entry.DefaultStamp) { $stampName = [string]$entry.DefaultStamp }
            elseif ($entry -and $entry.defaultStamp) { $stampName = [string]$entry.defaultStamp }
        }
    }

    if (_QCPT-IsBlank $stampName) {
        _QCPT-Log -Code 'QC_STAMP_UNRESOLVED' -Level 'Warning' `
            -Message 'No stamp could be resolved for process type.' `
            -Data @{
                qc_process_type = $normalized
                stampProfile = $profileName
                folderPath = $FolderPath
                watchRoot = $WatchRoot
                matchedOverride = if ($matchedOverride) { $matchedOverride.Name } else { $null }
            }
        return @{
            IsSuccess = $false
            Code = 'QC_STAMP_UNRESOLVED'
            Message = "No stamp resolved for process type '$normalized'."
            qc_process_type = $normalized
            resolvedStampProfile = $profileName
            resolvedStampName = $null
            stampPath = $null
            usedFallback = $usedFallback
            matchedRootOverride = if ($matchedOverride) { $matchedOverride.Name } else { $null }
        }
    }

    if (_QCPT-IsBlank $RepoRoot) { $RepoRoot = Split-Path -Parent $PSScriptRoot }
    $stampPath = _QCPT-ResolveStampAssetPath -Settings $settings -StampName $stampName -RepoRoot $RepoRoot

    # Legacy fallback: qcPrepend.reviewStamps paths
    if ((_QCPT-IsBlank $stampPath) -or -not (Test-Path -LiteralPath $stampPath)) {
        $qc = _QCPT-ToHashtable $Config.qcPrepend
        $rs = if ($qc) { _QCPT-ToHashtable $qc.reviewStamps } else { $null }
        if ($rs) {
            $legacyKey = if ($normalized -eq 'check') { 'independentCheck' } elseif ($normalized -eq 'review') { 'peerReview' } else { $null }
            if ($legacyKey) {
                $leg = _QCPT-ToHashtable $rs[$legacyKey]
                if ($leg -and $leg.stampPath) {
                    $lp = [string]$leg.stampPath
                    if (-not [System.IO.Path]::IsPathRooted($lp)) { $lp = Join-Path $RepoRoot $lp }
                    if (Test-Path -LiteralPath $lp) {
                        $stampPath = $lp
                        $usedFallback = $true
                    }
                }
            }
        }
    }

    if (_QCPT-IsBlank $stampPath) {
        _QCPT-Log -Code 'QC_STAMP_UNRESOLVED' -Level 'Warning' `
            -Message 'Stamp name resolved but asset path is missing.' `
            -Data @{
                qc_process_type = $normalized
                resolvedStampProfile = $profileName
                resolvedStampName = $stampName
            }
        return @{
            IsSuccess = $false
            Code = 'QC_STAMP_UNRESOLVED'
            Message = "Stamp asset not found for '$stampName'."
            qc_process_type = $normalized
            resolvedStampProfile = $profileName
            resolvedStampName = $stampName
            stampPath = $null
            usedFallback = $usedFallback
        }
    }

    $suffix = Get-QCProcessTypePdfSuffix -ProcessType $normalized -Config $Config
    _QCPT-Log -Code 'QC_STAMP_RESOLVED' -Level 'Information' `
        -Message 'Stamp resolved for process type.' `
        -Data @{
            qc_process_type = $normalized
            pdfSuffix = $suffix
            resolvedStampProfile = $profileName
            resolvedStampName = $stampName
            stampPath = $stampPath
            matchedRootOverride = if ($matchedOverride) { $matchedOverride.Name } else { $null }
            usedFallback = $usedFallback
        }

    return @{
        IsSuccess = $true
        Code = 'QC_STAMP_RESOLVED'
        Message = 'Stamp resolved.'
        qc_process_type = $normalized
        pdfSuffix = $suffix
        resolvedStampProfile = $profileName
        resolvedStampName = $stampName
        stampPath = $stampPath
        matchedRootOverride = if ($matchedOverride) { $matchedOverride.Name } else { $null }
        usedFallback = $usedFallback
        profileKey = if ($normalized -eq 'check') { 'independentCheck' } elseif ($normalized -eq 'review') { 'peerReview' } else { 'production' }
    }
}

function Resolve-QCProcessTypeFromContext {
    <#
    .SYNOPSIS
  Reads process type from a context hashtable with compatibility keys.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Context,
        [string]$DefaultProcessType = 'production'
    )
    if (-not $Context) {
        return Normalize-QCProcessType -ProcessType $DefaultProcessType
    }
    $resolved = Normalize-QCProcessType -Context $Context -AllowNullOnEmpty
    if ($resolved) { return $resolved }
    if (-not (_QCPT-IsBlank $DefaultProcessType)) {
        return Normalize-QCProcessType -ProcessType $DefaultProcessType
    }
    return $null
}

Export-ModuleMember -Function `
    Get-QCProcessTypeSettings, `
    Normalize-QCProcessType, `
    Get-QCProcessTypeDisplayLabel, `
    Get-QCProcessTypePdfSuffix, `
    Test-QCProcessTypeSyncsWithSiblingSheets, `
    Test-QCProcessTypeResetsAfterPrepend, `
    Get-PWQcPdfLaneFromDocumentName, `
    Test-PWQcPdfLaneSuffix, `
    Get-QCLaneQcPdfExpectedName, `
    Resolve-QCLaneQcPdf, `
    Resolve-QCStampForProcess, `
    Resolve-QCProcessTypeFromContext
