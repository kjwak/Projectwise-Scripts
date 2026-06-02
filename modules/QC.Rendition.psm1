# QC.Rendition.psm1
# Responsibility: QC_RENDITION jobs, profile resolution, readiness gating for Ready for QC notifications.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'PW.Connection.psm1') -Force -ErrorAction SilentlyContinue
if (-not (Get-Command -Name 'Get-PWDocumentDescriptionForFolder' -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $PSScriptRoot 'PW.Discovery.psm1') -ErrorAction SilentlyContinue
    if (-not (Get-Command -Name 'Get-PWDocumentDescriptionForFolder' -ErrorAction SilentlyContinue)) {
        Import-Module (Join-Path $PSScriptRoot 'PW.Discovery.psm1') -Force -ErrorAction SilentlyContinue
    }
}
if (-not (Get-Command -Name 'Test-QCDuplicateJob' -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $PSScriptRoot 'QC.Queue.Json.psm1') -Force -ErrorAction SilentlyContinue
}

function _QCR-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCR-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($key in @($Value.Keys)) { $h[[string]$key] = $Value[$key] }
        return $h
    }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _QCR-GetRepoRoot() {
    $root = $PSScriptRoot
    if ($root -match '[\\/]modules$') { return Split-Path -Parent $root }
    return $root
}

function _QCR-Sha256Hex([string]$Text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return -join ($hash | ForEach-Object { $_.ToString('x2') })
}

function _QCR-NormalizeFolderKey([string]$Path) {
    if (_QCR-IsNullOrWhiteSpace $Path) { return '' }
    $p = $Path.Trim().TrimEnd('\', '/').Replace('/', '\')
    if ($p.StartsWith('Documents\', [System.StringComparison]::OrdinalIgnoreCase)) {
        $p = $p.Substring('Documents\'.Length).TrimStart('\')
    }
    return $p.ToLowerInvariant()
}

function Get-QCRenditionSettings {
    [CmdletBinding()]
    param([hashtable]$Config)

    $defaults = @{
        enabled = $false
        deferReadyForQcNotification = $true
        readinessStorePath = (Join-Path (_QCR-GetRepoRoot) 'rendition-readiness')
        defaultSourceDocumentPattern = '%.dgn'
        deriveSourceFromQcPdf = $true
        makeSinglePlotRequest = $false
        completion = @{
            mode = 'submitOnly'
            pollAttempts = 1
            pollIntervalSeconds = 0
            fileNamePattern = '{stem}*.pdf'
        }
        datasources = @{}
        folderOverrides = @()
    }

    $raw = @{}
    if ($Config -and $Config.ContainsKey('qcRendition') -and $Config.qcRendition) {
        $norm = _QCR-ToHashtable $Config.qcRendition
        if ($norm) { $raw = $norm }
    }

    $settings = @{}
    foreach ($k in $defaults.Keys) { $settings[$k] = $defaults[$k] }
    foreach ($k in $raw.Keys) {
        if ($k -eq 'completion') {
            $c = _QCR-ToHashtable $raw.completion
            if ($c) { foreach ($ck in $c.Keys) { $settings.completion[$ck] = $c[$ck] } }
            continue
        }
        $settings[$k] = $raw[$k]
    }

    if ($settings.datasources -isnot [hashtable]) {
        $ds = _QCR-ToHashtable $settings.datasources
        $settings.datasources = if ($ds) { $ds } else { @{} }
    }
    if (-not ($settings.folderOverrides -is [System.Collections.IEnumerable]) -or ($settings.folderOverrides -is [string])) {
        $settings.folderOverrides = @()
    }

    return $settings
}

function _QCR-MergeProfileConfig([hashtable]$Merged, [hashtable]$Source) {
    if (-not $Merged -or -not $Source) { return }
    foreach ($k in $Source.Keys) {
        if ($k -eq 'completion') {
            $bc = _QCR-ToHashtable $Source.completion
            if ($bc) { foreach ($ck in $bc.Keys) { $Merged.completion[$ck] = $bc[$ck] } }
        } elseif ($k -notin @('folderPathPrefix', 'folderPath', 'path')) {
            $Merged[$k] = $Source[$k]
        }
    }
}

function _QCR-FindWatchRootRenditionProfile([hashtable]$Config, [string]$FolderPath) {
    $folderKey = _QCR-NormalizeFolderKey $FolderPath
    if (_QCR-IsNullOrWhiteSpace $folderKey) { return $null }

    $roots = @()
    if ($Config -and $Config.ContainsKey('projectWise') -and $Config.projectWise) {
        $pw = _QCR-ToHashtable $Config.projectWise
        if ($pw -and $pw.watchList) {
            $wl = _QCR-ToHashtable $pw.watchList
            if ($wl -and $wl.roots) { $roots = @($wl.roots) }
        }
    }

    $bestLen = -1
    $bestProfile = $null
    foreach ($entry in $roots) {
        $root = _QCR-ToHashtable $entry
        if (-not $root -or -not $root.path) { continue }
        $rk = _QCR-NormalizeFolderKey ([string]$root.path)
        if (_QCR-IsNullOrWhiteSpace $rk) { continue }
        if (-not $folderKey.StartsWith($rk, [System.StringComparison]::Ordinal)) { continue }
        if ($rk.Length -le $bestLen) { continue }

        $rendition = $null
        if ($root.qcRendition) { $rendition = _QCR-ToHashtable $root.qcRendition }
        if (-not $rendition) { continue }

        $bestLen = $rk.Length
        $bestProfile = $rendition
    }
    return $bestProfile
}

function Resolve-QCRenditionProfile {
    <#
    .SYNOPSIS
    Resolves rendition profile and options from watch roots, optional datasource defaults, and folder overrides.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$DatasourceName = '',
        [string]$FolderPath = ''
    )

    $settings = Get-QCRenditionSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        return New-QCFailureResult -Code 'QC_RENDITION_DISABLED' -Message 'qcRendition.enabled is false.' -Data @{}
    }

    $dsName = $DatasourceName
    if (_QCR-IsNullOrWhiteSpace $dsName -and $Config.ContainsKey('projectWise') -and $Config.projectWise) {
        $pw = _QCR-ToHashtable $Config.projectWise
        if ($pw -and $pw.datasourceName) { $dsName = [string]$pw.datasourceName }
    }

    $merged = @{
        datasourceName = $dsName
        profileName = $null
        sourceDocumentPattern = [string]$settings.defaultSourceDocumentPattern
        deriveSourceFromQcPdf = [bool]$settings.deriveSourceFromQcPdf
        useFolderProfileWhenUnspecified = $true
        outputFolderPath = $null
        outputFolderRelative = $null
        makeSinglePlotRequest = [bool]$settings.makeSinglePlotRequest
        completion = $settings.completion
    }

    $dsProfiles = $settings.datasources
    if ($dsProfiles -is [hashtable] -and -not (_QCR-IsNullOrWhiteSpace $dsName) -and $dsProfiles.ContainsKey($dsName)) {
        _QCR-MergeProfileConfig -Merged $merged -Source (_QCR-ToHashtable $dsProfiles[$dsName])
    }

    $rootProfile = _QCR-FindWatchRootRenditionProfile -Config $Config -FolderPath $FolderPath
    if ($rootProfile) { _QCR-MergeProfileConfig -Merged $merged -Source $rootProfile }

    $folderKey = _QCR-NormalizeFolderKey $FolderPath
    $bestLen = -1
    $bestOverride = $null
    foreach ($entry in @($settings.folderOverrides)) {
        $ov = _QCR-ToHashtable $entry
        if (-not $ov) { continue }
        $prefix = if ($ov.folderPathPrefix) { [string]$ov.folderPathPrefix } elseif ($ov.folderPath) { [string]$ov.folderPath } else { '' }
        if (_QCR-IsNullOrWhiteSpace $prefix) { continue }
        $pk = _QCR-NormalizeFolderKey $prefix
        if (_QCR-IsNullOrWhiteSpace $folderKey) { continue }
        if ($folderKey.StartsWith($pk, [System.StringComparison]::Ordinal) -and $pk.Length -gt $bestLen) {
            $bestLen = $pk.Length
            $bestOverride = $ov
        }
    }
    if ($bestOverride) { _QCR-MergeProfileConfig -Merged $merged -Source $bestOverride }

    return New-QCSuccessResult -Code 'QC_RENDITION_PROFILE_RESOLVED' -Message 'Rendition profile resolved.' -Data @{ profile = $merged }
}

function Get-QCReadinessKey {
    param(
        [string]$DocumentGuid = '',
        [string]$FolderPath = '',
        [string]$QcPdfName = ''
    )
    if (-not (_QCR-IsNullOrWhiteSpace $DocumentGuid)) {
        return ('guid:' + $DocumentGuid.Trim().ToLowerInvariant())
    }
    $fk = _QCR-NormalizeFolderKey $FolderPath
    $stem = [System.IO.Path]::GetFileNameWithoutExtension([string]$QcPdfName)
    if ($stem -match '(?i)-qc$') { $stem = $stem.Substring(0, $stem.Length - 3) }
    return ('folder:' + $fk + '|' + $stem.ToLowerInvariant())
}

function _QCR-ReadinessFilePath([hashtable]$Settings, [string]$Key) {
    $root = if ($Settings.readinessStorePath) { [string]$Settings.readinessStorePath } else { (Join-Path (_QCR-GetRepoRoot) 'rendition-readiness') }
    if (-not (Test-Path -LiteralPath $root)) {
        New-Item -ItemType Directory -Path $root -Force | Out-Null
    }
    $safe = (_QCR-Sha256Hex -Text $Key).Substring(0, 32)
    return (Join-Path $root ($safe + '.json'))
}

function Get-QCReadinessState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$ReadinessKey
    )

    $settings = Get-QCRenditionSettings -Config $Config
    $path = _QCR-ReadinessFilePath -Settings $settings -Key $ReadinessKey
    if (-not (Test-Path -LiteralPath $path)) {
        return @{
            readinessKey = $ReadinessKey
            prependComplete = $false
            renditionComplete = $false
            readyNotificationSent = $false
        }
    }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
        $h = _QCR-ToHashtable $raw
        if (-not $h) { $h = @{} }
        return @{
            readinessKey = $ReadinessKey
            prependComplete = [bool]$h.prependComplete
            renditionComplete = [bool]$h.renditionComplete
            readyNotificationSent = [bool]$h.readyNotificationSent
            updatedAtUtc = if ($h.updatedAtUtc) { [string]$h.updatedAtUtc } else { $null }
            parentPrependJobId = if ($h.parentPrependJobId) { [string]$h.parentPrependJobId } else { $null }
        }
    } catch {
        return @{
            readinessKey = $ReadinessKey
            prependComplete = $false
            renditionComplete = $false
            readyNotificationSent = $false
        }
    }
}

function Set-QCReadinessFlag {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$ReadinessKey,
        [switch]$PrependComplete,
        [switch]$RenditionComplete,
        [switch]$ReadyNotificationSent,
        [string]$ParentPrependJobId = '',
        [string]$DocumentGuid = '',
        [string]$FolderPath = '',
        [string]$QcPdfName = ''
    )

    $settings = Get-QCRenditionSettings -Config $Config
    $path = _QCR-ReadinessFilePath -Settings $settings -Key $ReadinessKey
    $state = Get-QCReadinessState -Config $Config -ReadinessKey $ReadinessKey
    if ($PrependComplete) { $state.prependComplete = $true }
    if ($RenditionComplete) { $state.renditionComplete = $true }
    if ($ReadyNotificationSent) { $state.readyNotificationSent = $true }
    if (-not (_QCR-IsNullOrWhiteSpace $ParentPrependJobId)) { $state.parentPrependJobId = $ParentPrependJobId }
    if (-not (_QCR-IsNullOrWhiteSpace $DocumentGuid)) { $state.documentGuid = $DocumentGuid }
    if (-not (_QCR-IsNullOrWhiteSpace $FolderPath)) { $state.folderPath = $FolderPath }
    if (-not (_QCR-IsNullOrWhiteSpace $QcPdfName)) { $state.qcPdfName = $QcPdfName }
    $state.updatedAtUtc = Get-QCTimestamp
    $state.readinessKey = $ReadinessKey

    $json = ($state | ConvertTo-Json -Depth 6 -Compress)
    $tmp = $path + '.tmp'
    [System.IO.File]::WriteAllText($tmp, $json, [System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tmp -Destination $path -Force
    return $state
}

function _QCR-GetReadyForQcStateName([hashtable]$Config) {
    $readyName = 'Ready for QC'
    if ($Config -and $Config.ContainsKey('qcWorkflow') -and $Config.qcWorkflow) {
        $wf = _QCR-ToHashtable $Config.qcWorkflow
        if ($wf -and $wf.states) {
            $st = _QCR-ToHashtable $wf.states
            if ($st -and $st.readyForQc) { $readyName = [string]$st.readyForQc }
        }
        if ($wf -and $wf.receivedStateName) { $readyName = [string]$wf.receivedStateName }
    }
    return $readyName
}

function Test-QCShouldDeferReadyForQcNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$CurrentState
    )

    $rendition = Get-QCRenditionSettings -Config $Config
    if (-not [bool]$rendition.enabled) { return $false }
    if (-not [bool]$rendition.deferReadyForQcNotification) { return $false }

    $readyName = _QCR-GetReadyForQcStateName -Config $Config
    return (([string]$CurrentState).Trim().ToLowerInvariant() -eq $readyName.ToLowerInvariant())
}

function _QCR-GetDocumentGuidFromObject([object]$Document) {
    if (-not $Document) { return $null }
    foreach ($n in @('DocumentGUID', 'DocumentGuid', 'GUID', 'Id', 'DocumentID')) {
        try {
            if ($Document.PSObject.Properties[$n] -and -not (_QCR-IsNullOrWhiteSpace $Document.$n)) {
                return [string]$Document.$n
            }
        } catch { }
    }
    return $null
}

function _QCR-DeriveSourceDocumentName([string]$QcPdfName, [bool]$DeriveFromQcPdf, [string]$ExplicitName) {
    if (-not (_QCR-IsNullOrWhiteSpace $ExplicitName)) { return [string]$ExplicitName }
    $base = [System.IO.Path]::GetFileName([string]$QcPdfName)
    if (_QCR-IsNullOrWhiteSpace $base) { return $null }
    if ($DeriveFromQcPdf) {
        if ($base -match '(?i)^(.+)-qc\.pdf$') { return ($Matches[1] + '.dgn') }
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($base)
        return ($stem + '.dgn')
    }
    return $base
}

function Test-QCRenditionOutputComplete {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Profile,
        [Parameter(Mandatory)][string]$SourceDocumentName,
        [string]$OutputFolderPath = ''
    )

    $completion = _QCR-ToHashtable $Profile.completion
    if (-not $completion) { $completion = @{ mode = 'outputFolder' } }
    $mode = if ($completion.mode) { ([string]$completion.mode).Trim().ToLowerInvariant() } else { 'outputFolder' }

    if ($mode -eq 'immediate' -or $mode -eq 'submitOnly') {
        return New-QCSuccessResult -Code 'QC_RENDITION_COMPLETE_IMMEDIATE' -Message 'Completion mode does not require output polling.' -Data @{ complete = $true }
    }

    $outFolder = $OutputFolderPath
    if (_QCR-IsNullOrWhiteSpace $outFolder -and $Profile.outputFolderPath) { $outFolder = [string]$Profile.outputFolderPath }
    if (_QCR-IsNullOrWhiteSpace $outFolder) {
        return New-QCFailureResult -Code 'QC_RENDITION_OUTPUT_FOLDER_MISSING' -Message 'outputFolderPath is required for outputFolder completion mode.' -Data @{ complete = $false }
    }

    $pattern = if ($completion.fileNamePattern) { [string]$completion.fileNamePattern } else { '{stem}*.pdf' }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($SourceDocumentName)
    if ($stem -match '\.dgn$') { $stem = [System.IO.Path]::GetFileNameWithoutExtension($stem) }
    $glob = $pattern.Replace('{stem}', $stem)

    if (-not (Test-Path -LiteralPath $outFolder)) {
        return New-QCSuccessResult -Code 'QC_RENDITION_OUTPUT_PENDING' -Message 'Output folder not found yet.' -Data @{ complete = $false; outputFolder = $outFolder; pattern = $glob }
    }

    $hits = @(Get-ChildItem -LiteralPath $outFolder -Filter $glob -File -ErrorAction SilentlyContinue)
    if ($hits.Count -gt 0) {
        return New-QCSuccessResult -Code 'QC_RENDITION_OUTPUT_FOUND' -Message 'Rendition output file found.' -Data @{
            complete = $true
            outputFolder = $outFolder
            files = @($hits | ForEach-Object { $_.FullName })
        }
    }
    return New-QCSuccessResult -Code 'QC_RENDITION_OUTPUT_PENDING' -Message 'Rendition output not found yet.' -Data @{ complete = $false; outputFolder = $outFolder; pattern = $glob }
}

function Invoke-QCReadyForQcNotificationIfReady {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$ReadinessKey,
        [string]$PreviousState = '',
        [string]$CurrentState = '',
        [object]$Document = $null,
        [hashtable]$Job = $null,
        [string]$DocumentName = '',
        [string]$DocumentGuid = '',
        [string]$FolderPath = ''
    )

    $rendition = Get-QCRenditionSettings -Config $Config
    if (-not [bool]$rendition.deferReadyForQcNotification) {
        return New-QCSuccessResult -Code 'QC_READY_NOTIFICATION_NOT_DEFERRED' -Message 'Ready for QC notification is sent at prepend workflow state change; deferred rendition send skipped.' -Data @{ readinessKey = $ReadinessKey }
    }

    $state = Get-QCReadinessState -Config $Config -ReadinessKey $ReadinessKey
    if ($state.readyNotificationSent) {
        return New-QCSuccessResult -Code 'QC_READY_NOTIFICATION_ALREADY_SENT' -Message 'Ready for QC notification already sent for this key.' -Data @{ readinessKey = $ReadinessKey }
    }
    if (-not $state.prependComplete -or -not $state.renditionComplete) {
        return New-QCSuccessResult -Code 'QC_READY_NOTIFICATION_DEFERRED' -Message 'Prepend and rendition must both complete before notification.' -Data @{
            readinessKey = $ReadinessKey
            prependComplete = $state.prependComplete
            renditionComplete = $state.renditionComplete
        }
    }

    $curr = _QCR-GetReadyForQcStateName -Config $Config
    if (-not (_QCR-IsNullOrWhiteSpace $CurrentState)) { $curr = ([string]$CurrentState).Trim() }
    if (_QCR-IsNullOrWhiteSpace $curr) { $curr = 'Ready for QC' }
    if (-not (Get-Command -Name 'Invoke-QCNotificationForStateChange' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'QC_NOTIFICATION_UNAVAILABLE' -Message 'QC.Notifications module not loaded.' -Data @{}
    }

    $prev = if (_QCR-IsNullOrWhiteSpace $PreviousState) { 'In Production' } else { $PreviousState }
    $docPath = if ($FolderPath -and $DocumentName) { $FolderPath + '\' + $DocumentName } else { $null }
    $notifParams = @{
        Config = $Config
        PreviousState = $prev
        CurrentState = $curr
        Force = $true
    }
    if ($Document) { $notifParams.Document = $Document }
    if (-not (_QCR-IsNullOrWhiteSpace $DocumentName)) { $notifParams.DocumentName = $DocumentName }
    if (-not (_QCR-IsNullOrWhiteSpace $DocumentGuid)) { $notifParams.DocumentGuid = $DocumentGuid }
    if (-not (_QCR-IsNullOrWhiteSpace $docPath)) { $notifParams.DocumentPath = $docPath }
    if ($Job) { $notifParams.Job = $Job }
    $notif = Invoke-QCNotificationForStateChange @notifParams

    if ($notif -and $notif.IsSuccess) {
        Set-QCReadinessFlag -Config $Config -ReadinessKey $ReadinessKey -ReadyNotificationSent | Out-Null
    }
    return $notif
}

function New-QCRenditionQueueJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$ParentJob,
        [string]$DocumentGuid = '',
        [object]$Document = $null
    )

    $profileRes = Resolve-QCRenditionProfile -Config $Config -FolderPath ([string]$ParentJob.sourceFolder)
    if (-not $profileRes.IsSuccess) { return $profileRes }
    $profile = $profileRes.Data.profile

    $folder = if ($ParentJob.sourceFolder) { [string]$ParentJob.sourceFolder } else { '' }
    $qcName = if ($ParentJob.sourceName) { [string]$ParentJob.sourceName } else { [System.IO.Path]::GetFileName([string]$ParentJob.sourcePath) }
    if (-not $DocumentGuid) { $DocumentGuid = _QCR-GetDocumentGuidFromObject $Document }

    $sourceDoc = _QCR-DeriveSourceDocumentName -QcPdfName $qcName -DeriveFromQcPdf ([bool]$profile.deriveSourceFromQcPdf) -ExplicitName ''
    if (_QCR-IsNullOrWhiteSpace $sourceDoc) {
        return New-QCFailureResult -Code 'QC_RENDITION_SOURCE_NAME_MISSING' -Message 'Could not derive source document name from QC PDF.' -Data @{ qcPdfName = $qcName }
    }
    if ($sourceDoc -notmatch '(?i)\.dgn$') {
        return New-QCFailureResult -Code 'QC_RENDITION_SOURCE_NOT_DGN' -Message 'Derived rendition source must be a DGN file name.' -Data @{ qcPdfName = $qcName; sourceDocument = $sourceDoc }
    }

    $readinessKey = Get-QCRenditionSheetReadinessKey -FolderPath $folder -SourceDgnFileName $sourceDoc
    $dedupeKey = _QCR-GetRenditionDedupeKeyForSheet -FolderPath $folder -SourceDgnFileName $sourceDoc
    $jobId = _QCR-GetRenditionJobIdForSheet -FolderPath $folder -SourceDgnFileName $sourceDoc
    $sourcePath = if ($folder) { Join-Path $folder $sourceDoc } else { $sourceDoc }

    $job = @{
        id = $jobId
        type = 'QC_RENDITION'
        sourcePath = $sourcePath
        sourceName = $sourceDoc
        sourceFolder = $folder
        dedupeKey = $dedupeKey
        status = 'queued'
        createdAt = Get-QCTimestamp
        attempts = 0
        triggerRule = @{
            id = 'qc-rendition-after-prepend'
            jobType = 'QC_RENDITION'
            triggerType = 'processor'
        }
        metadata = @{
            rendition = @{
                parentPrependJobId = [string]$ParentJob.id
                readinessKey = $readinessKey
                sheetReadinessKey = $readinessKey
                documentGuid = $DocumentGuid
                qcPdfName = $qcName
                profileName = $profile.profileName
                sourceDocumentPattern = $profile.sourceDocumentPattern
            }
        }
    }

    return New-QCSuccessResult -Code 'QC_RENDITION_JOB_BUILT' -Message 'QC_RENDITION job payload built.' -Data @{ job = $job; readinessKey = $readinessKey; profile = $profile }
}

function _QCR-ResolveTargetStateFromWriteback([object]$Writeback, [hashtable]$WfSettings) {
    $targetState = $null
    if ($Writeback -and $Writeback.Data) {
        $wb = _QCR-ToHashtable $Writeback.Data
        if ($wb -and $wb.actions) {
            foreach ($a in @($wb.actions)) {
                if (-not $a) { continue }
                if ([string]$a.Code -notmatch '^QC_WORKFLOW_STATE') { continue }
                $ad = _QCR-ToHashtable $a.Data
                if ($ad -and $ad.stateName) { $targetState = [string]$ad.stateName; break }
            }
        }
    }
    if (_QCR-IsNullOrWhiteSpace $targetState -and $WfSettings) {
        if ($WfSettings.stateAfterSuccessfulPrepend) { $targetState = [string]$WfSettings.stateAfterSuccessfulPrepend }
        elseif ($WfSettings.defaultStateAfterPrepend) { $targetState = [string]$WfSettings.defaultStateAfterPrepend }
        else { $targetState = _QCR-GetReadyForQcStateName -Config @{ qcWorkflow = $WfSettings } }
    }
    return $targetState
}

function Add-QCRenditionJobAfterPrepend {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Job,
        [hashtable]$Writeback,
        [object]$Document = $null
    )

    $rendition = Get-QCRenditionSettings -Config $Config
    if (-not [bool]$rendition.enabled) { return $null }

    $wfRaw = if ($Config.qcWorkflow) { _QCR-ToHashtable $Config.qcWorkflow } else { @{} }
    $readyName = _QCR-GetReadyForQcStateName -Config $Config
    $targetState = _QCR-ResolveTargetStateFromWriteback -Writeback $Writeback -WfSettings $wfRaw
    $targetTrim = ([string]$targetState).Trim()
    if ($targetTrim.Length -eq 0 -or ($targetTrim.ToLowerInvariant() -ne $readyName.ToLowerInvariant())) { return $null }

    $docGuid = _QCR-GetDocumentGuidFromObject $Document
    $folder = if ($Job.sourceFolder) { [string]$Job.sourceFolder } else { '' }
    $qcName = if ($Job.sourceName) { [string]$Job.sourceName } else { '' }
    $built = New-QCRenditionQueueJob -Config $Config -ParentJob $Job -DocumentGuid $docGuid -Document $Document
    if (-not $built.IsSuccess) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Warning' -Code 'QC_RENDITION_ENQUEUE_SKIPPED' -Message $built.Message -Data @{
                jobId = [string]$Job.id; code = [string]$built.Code
            } | Out-Null
        }
        $sheetKeyFail = $null
        try {
            if ($built.Data -and $built.Data.job -and $built.Data.job.sourceName) {
                $sheetKeyFail = Get-QCRenditionSheetReadinessKey -FolderPath $folder -SourceDgnFileName ([string]$built.Data.job.sourceName)
            }
        } catch { }
        if (-not $sheetKeyFail) { $sheetKeyFail = Get-QCReadinessKey -DocumentGuid $docGuid -FolderPath $folder -QcPdfName $qcName }
        if ([bool]$rendition.deferReadyForQcNotification) {
            Set-QCReadinessFlag -Config $Config -ReadinessKey $sheetKeyFail -RenditionComplete | Out-Null
            Invoke-QCReadyForQcNotificationIfReady -Config $Config -ReadinessKey $sheetKeyFail `
                -Document $Document -Job $Job -DocumentName $qcName -DocumentGuid $docGuid -FolderPath $folder | Out-Null
        }
        return $built
    }

    $renditionJobPreview = $built.Data.job
    $sheetKey = Get-QCRenditionSheetReadinessKey -FolderPath $folder -SourceDgnFileName ([string]$renditionJobPreview.sourceName)
    Set-QCReadinessFlag -Config $Config -ReadinessKey $sheetKey -PrependComplete `
        -ParentPrependJobId ([string]$Job.id) -DocumentGuid $docGuid -FolderPath $folder -QcPdfName $qcName | Out-Null

    if (-not (Get-Command -Name 'Add-QCQueueJob' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'QC_RENDITION_QUEUE_UNAVAILABLE' -Message 'QC.Queue.Json not loaded.' -Data @{}
    }

    $renditionJob = $built.Data.job
    $sheetKey = Get-QCRenditionSheetReadinessKey -FolderPath $folder -SourceDgnFileName ([string]$renditionJob.sourceName)
    $block = _QCR-TestRenditionEnqueueBlocked -Config $Config -DedupeKey ([string]$renditionJob.dedupeKey) -SheetReadinessKey $sheetKey
    if ([bool]$block.blocked) {
        _QCR-WriteRenditionStateEnqueueLog -Level 'Information' -Code 'QC_RENDITION_SKIPPED_ALREADY_DONE' `
            -Message 'QC_RENDITION skipped after prepend; sheet rendition already queued or complete.' -Data @{
            parentJobId = [string]$Job.id
            renditionJobId = [string]$renditionJob.id
            dedupeKey = [string]$renditionJob.dedupeKey
            skipReason = [string]$block.reason
        }
        return New-QCSuccessResult -Code 'QC_RENDITION_SKIPPED_ALREADY_DONE' -Message 'Rendition already in flight or complete for this sheet.' -Data @{
            jobId = [string]$renditionJob.id
            skipReason = [string]$block.reason
        }
    }

    $enq = Add-QCQueueJob -Job $renditionJob -Config $Config
    if ($enq.IsSuccess) {
        _QCR-WriteRenditionStateEnqueueLog -Level 'Information' -Code 'QC_RENDITION_ENQUEUED' `
            -Message 'QC_RENDITION job enqueued after prepend.' -Data @{
            parentJobId = [string]$Job.id
            renditionJobId = [string]$renditionJob.id
            readinessKey = $sheetKey
            dedupeKey = [string]$renditionJob.dedupeKey
        }
        return $enq
    }
    if ([string]$enq.Code -eq 'QUEUE_JOB_ALREADY_EXISTS') {
        return New-QCSuccessResult -Code 'QC_RENDITION_SKIPPED_ALREADY_DONE' -Message $enq.Message -Data @{ jobId = [string]$renditionJob.id }
    }
    _QCR-WriteRenditionStateEnqueueLog -Level 'Warning' -Code 'QC_RENDITION_ENQUEUE_FAILED' -Message $enq.Message -Data @{
        parentJobId = [string]$Job.id
        renditionJobId = [string]$renditionJob.id
        enqueueCode = [string]$enq.Code
    }
    return $enq
}

function _QCR-DeriveSourceFromStateChange([string]$TriggerDocumentName) {
    $name = [System.IO.Path]::GetFileName([string]$TriggerDocumentName)
    if (_QCR-IsNullOrWhiteSpace $name) { return $null }
    if ($name -match '(?i)\.dgn$') { return $name }
    if ($name -match '(?i)^(.+)-qc\.pdf$') { return ($Matches[1] + '.dgn') }
    if ($name -match '(?i)^(.+)\.pdf$') { return ($Matches[1] + '.dgn') }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($name)
    if (_QCR-IsNullOrWhiteSpace $stem) { return $null }
    return ($stem + '.dgn')
}

function Get-QCRenditionSheetReadinessKey {
    <#
    .SYNOPSIS
    Stable readiness key for a sheet (folder + DGN stem), shared by DGN/PDF/QC-PDF siblings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDgnFileName
    )

    $fk = _QCR-NormalizeFolderKey $FolderPath
    $stem = [System.IO.Path]::GetFileNameWithoutExtension([string]$SourceDgnFileName)
    if ($stem -match '(?i)\.dgn$') { $stem = [System.IO.Path]::GetFileNameWithoutExtension($stem) }
    return ('sheet:' + $fk + '|' + $stem.ToLowerInvariant())
}

function _QCR-GetRenditionDedupeKeyForSheet {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDgnFileName
    )

    $stable = ('dedupeV3_rendition|folder={0}|source={1}' -f $FolderPath, $SourceDgnFileName)
    return 'dq_qcrendition_' + (_QCR-Sha256Hex -Text $stable).Substring(0, 24)
}

function _QCR-GetRenditionJobIdForSheet {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDgnFileName
    )

    return 'rendition-' + (_QCR-Sha256Hex -Text ('sheet|' + $FolderPath + '|' + $SourceDgnFileName)).Substring(0, 16)
}

function _QCR-TestRenditionEnqueueBlocked {
    <#
    Returns $true when a sheet-level rendition is already pending, running, succeeded, or marked complete.
    Failed jobs are not blocked so a retry can be enqueued.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DedupeKey,
        [Parameter(Mandatory)][string]$SheetReadinessKey
    )

    $ready = Get-QCReadinessState -Config $Config -ReadinessKey $SheetReadinessKey
    if ([bool]$ready.renditionComplete) {
        return @{
            blocked = $true
            reason = 'readiness_complete'
            matches = @()
        }
    }

    if (-not (Get-Command -Name 'Test-QCDuplicateJob' -ErrorAction SilentlyContinue)) { return @{ blocked = $false } }

    $dup = Test-QCDuplicateJob -DedupeKey $DedupeKey -Config $Config
    if (-not $dup.IsSuccess -or -not [bool]$dup.Data.isDuplicate) {
        return @{ blocked = $false }
    }

    $matches = @()
    try { if ($dup.Data.matches) { $matches = @($dup.Data.matches) } } catch { }
    $blockStates = @('pending', 'running', 'succeeded')
    $blocking = @($matches | Where-Object { $blockStates -contains [string]$_.state })
    if ($blocking.Count -eq 0) {
        return @{ blocked = $false }
    }

    $reason = 'queue_pending'
    if (@($blocking | Where-Object { $_.state -eq 'running' }).Count -gt 0) { $reason = 'queue_running' }
    elseif (@($blocking | Where-Object { $_.state -eq 'succeeded' }).Count -gt 0) { $reason = 'queue_succeeded' }

    return @{
        blocked = $true
        reason = $reason
        matches = $blocking
    }
}

function _QCR-WriteRenditionStateEnqueueLog {
    param(
        [string]$Level,
        [string]$Code,
        [string]$Message,
        [hashtable]$Data
    )

    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    Write-QCJsonLog -Level $Level -Code $Code -Message $Message -Data $Data | Out-Null
}

function Add-QCRenditionJobForReadyForQcStateChange {
    <#
    .SYNOPSIS
    Enqueues QC_RENDITION when a non-automation actor sets state to Ready for QC.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$TriggerDocumentGuid,
        [Parameter(Mandatory)][string]$TriggerDocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$CurrentStateName,
        [bool]$DryRun = $false,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )

    $rendition = Get-QCRenditionSettings -Config $Config
    if (-not [bool]$rendition.enabled) { return $null }

    $readyName = _QCR-GetReadyForQcStateName -Config $Config
    $curr = ([string]$CurrentStateName).Trim()
    if ($curr.Length -eq 0 -or $curr.ToLowerInvariant() -ne $readyName.ToLowerInvariant()) { return $null }

    if (Get-Command -Name 'Test-QCIsAutomationPwActor' -ErrorAction SilentlyContinue) {
        if (Test-QCIsAutomationPwActor -Config $Config -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername) {
            return $null
        }
    }

    $triggerFile = [System.IO.Path]::GetFileName([string]$TriggerDocumentName)
    if ($triggerFile -notmatch '(?i)\.dgn$') {
        return $null
    }

    if (-not (Get-Command -Name 'Add-QCQueueJob' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'QC_RENDITION_QUEUE_UNAVAILABLE' -Message 'QC.Queue.Json not loaded.' -Data @{}
    }

    $sourceDoc = _QCR-DeriveSourceFromStateChange -TriggerDocumentName $TriggerDocumentName
    if (_QCR-IsNullOrWhiteSpace $sourceDoc) {
        return New-QCFailureResult -Code 'QC_RENDITION_SOURCE_NAME_MISSING' -Message 'Could not derive source DGN name from state-change document.' -Data @{
            triggerDocumentName = $TriggerDocumentName
        }
    }

    $sheetReadinessKey = Get-QCRenditionSheetReadinessKey -FolderPath $FolderPath -SourceDgnFileName $sourceDoc
    $dedupeKey = _QCR-GetRenditionDedupeKeyForSheet -FolderPath $FolderPath -SourceDgnFileName $sourceDoc
    $jobId = _QCR-GetRenditionJobIdForSheet -FolderPath $FolderPath -SourceDgnFileName $sourceDoc

    $block = _QCR-TestRenditionEnqueueBlocked -Config $Config -DedupeKey $dedupeKey -SheetReadinessKey $sheetReadinessKey
    if ([bool]$block.blocked) {
        _QCR-WriteRenditionStateEnqueueLog -Level 'Information' -Code 'QC_RENDITION_SKIPPED_ALREADY_DONE' `
            -Message 'QC_RENDITION skipped; sheet rendition already queued, running, succeeded, or marked complete.' -Data @{
            renditionJobId = $jobId
            dedupeKey = $dedupeKey
            sheetReadinessKey = $sheetReadinessKey
            folderPath = $FolderPath
            sourceDocument = $sourceDoc
            triggerDocumentGuid = $TriggerDocumentGuid
            skipReason = [string]$block.reason
            queueMatches = @($block.matches)
        }
        return New-QCSuccessResult -Code 'QC_RENDITION_SKIPPED_ALREADY_DONE' -Message 'Rendition already in flight or complete for this sheet.' -Data @{
            jobId = $jobId
            dedupeKey = $dedupeKey
            sheetReadinessKey = $sheetReadinessKey
            skipReason = [string]$block.reason
        }
    }

    Set-QCReadinessFlag -Config $Config -ReadinessKey $sheetReadinessKey -PrependComplete `
        -DocumentGuid $TriggerDocumentGuid -FolderPath $FolderPath -QcPdfName $TriggerDocumentName | Out-Null

    $profileRes = Resolve-QCRenditionProfile -Config $Config -FolderPath $FolderPath
    if (-not $profileRes.IsSuccess) {
        if ([bool]$rendition.deferReadyForQcNotification) {
            Set-QCReadinessFlag -Config $Config -ReadinessKey $sheetReadinessKey -RenditionComplete | Out-Null
            Invoke-QCReadyForQcNotificationIfReady -Config $Config -ReadinessKey $sheetReadinessKey `
                -DocumentName $TriggerDocumentName -DocumentGuid $TriggerDocumentGuid -FolderPath $FolderPath | Out-Null
        }
        return $profileRes
    }
    $profile = $profileRes.Data.profile

    $job = @{
        id         = $jobId
        type       = 'QC_RENDITION'
        sourcePath = (Join-Path $FolderPath $sourceDoc)
        sourceName = $sourceDoc
        sourceFolder = $FolderPath
        dedupeKey  = $dedupeKey
        status     = 'queued'
        createdAt  = Get-QCTimestamp
        attempts   = 0
        triggerRule = @{
            id          = 'qc-rendition-readyforqc'
            jobType     = 'QC_RENDITION'
            triggerType = 'audit_state_change'
        }
        metadata = @{
            rendition = @{
                readinessKey         = $sheetReadinessKey
                sheetReadinessKey    = $sheetReadinessKey
                documentGuid         = $TriggerDocumentGuid
                triggerDocumentName  = $TriggerDocumentName
                triggerCurrentState  = $curr
                changedByUser        = $ChangedByUser
                changedByUsername    = $ChangedByUsername
                profileName          = $profile.profileName
                sourceDocumentPattern = $profile.sourceDocumentPattern
            }
        }
    }

    if ($DryRun) {
        return New-QCSuccessResult -Code 'QC_RENDITION_PLANNED' -Message 'Dry-run: QC_RENDITION job planned from Ready for QC state change.' -Data @{
            job = $job
            readinessKey = $sheetReadinessKey
            profile = $profile
        }
    }

    $enq = Add-QCQueueJob -Job $job -Config $Config
    if ($enq.IsSuccess) {
        _QCR-WriteRenditionStateEnqueueLog -Level 'Information' -Code 'QC_RENDITION_ENQUEUED' `
            -Message 'QC_RENDITION job enqueued from Ready for QC state change (DGN trigger).' -Data @{
            renditionJobId = [string]$job.id
            dedupeKey = $dedupeKey
            sheetReadinessKey = $sheetReadinessKey
            folderPath = $FolderPath
            sourceDocument = $sourceDoc
            triggerDocumentGuid = $TriggerDocumentGuid
            changedByUser = $ChangedByUser
        }
        return $enq
    }

    if ([string]$enq.Code -eq 'QUEUE_JOB_ALREADY_EXISTS') {
        _QCR-WriteRenditionStateEnqueueLog -Level 'Information' -Code 'QC_RENDITION_SKIPPED_ALREADY_DONE' `
            -Message 'QC_RENDITION skipped; job id already pending in queue.' -Data @{
            renditionJobId = [string]$job.id
            dedupeKey = $dedupeKey
            sheetReadinessKey = $sheetReadinessKey
            folderPath = $FolderPath
            sourceDocument = $sourceDoc
            skipReason = 'queue_pending_same_job_id'
        }
        return New-QCSuccessResult -Code 'QC_RENDITION_SKIPPED_ALREADY_DONE' -Message $enq.Message -Data @{ jobId = $jobId; dedupeKey = $dedupeKey }
    }

    _QCR-WriteRenditionStateEnqueueLog -Level 'Warning' -Code 'QC_RENDITION_ENQUEUE_FAILED' `
        -Message $enq.Message -Data @{
        renditionJobId = [string]$job.id
        dedupeKey = $dedupeKey
        folderPath = $FolderPath
        sourceDocument = $sourceDoc
        enqueueCode = [string]$enq.Code
    }
    return $enq
}

function _QCR-GetFolderRenditionProfileName([string]$ApiFolderPath) {
    $cmd = Get-Command -Name 'Get-PWFolderRenditionProfile' -ErrorAction SilentlyContinue
    if ((-not $cmd) -or (_QCR-IsNullOrWhiteSpace $ApiFolderPath)) { return $null }
    try {
        $result = Get-PWFolderRenditionProfile -FolderPath $ApiFolderPath -ErrorAction Stop
    } catch { return $null }
    if (-not $result) { return $null }
    if ($result -is [string]) { return $result.Trim() }
    foreach ($prop in @('ProfileName', 'Name', 'RenditionProfileName')) {
        if ($result.PSObject.Properties[$prop] -and -not (_QCR-IsNullOrWhiteSpace $result.$prop)) {
            return [string]$result.$prop
        }
    }
    return [string]$result
}

function _QCR-ResolveOutputFolderPath([hashtable]$Profile, [string]$ApiFolderPath) {
    if ($Profile.outputFolderPath) { return [string]$Profile.outputFolderPath }
    if (_QCR-IsNullOrWhiteSpace $ApiFolderPath) { return $null }
    $api = $ApiFolderPath.TrimEnd('\')

    $rel = $null
    if ($Profile -and $Profile.ContainsKey('outputFolderRelative') -and $null -ne $Profile.outputFolderRelative) {
        $rel = [string]$Profile.outputFolderRelative
    }
    if (_QCR-IsNullOrWhiteSpace $rel) { return $api }

    $trimmed = $rel.Trim().TrimEnd('\')
    if ($trimmed -eq '.' -or $trimmed -eq '') { return $api }
    if ([System.IO.Path]::IsPathRooted($rel)) { return $rel }
    return ($api + '\' + $trimmed.TrimStart('\'))
}

function Invoke-QCRenditionProcessor {
    <#
    .SYNOPSIS
    Processes QC_RENDITION queue jobs: submit PW rendition and poll for output.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config
    )

    $renditionMeta = $null
    if ($Job.metadata -and $Job.metadata.rendition) { $renditionMeta = _QCR-ToHashtable $Job.metadata.rendition }
    if (-not $renditionMeta) { $renditionMeta = @{} }

    $readinessKey = $null
    if ($renditionMeta.sheetReadinessKey) { $readinessKey = [string]$renditionMeta.sheetReadinessKey }
    elseif ($renditionMeta.readinessKey) { $readinessKey = [string]$renditionMeta.readinessKey }
    else {
        $src = if ($Job.sourceName) { [string]$Job.sourceName } else { '' }
        if ($src -and $Job.sourceFolder) {
            $readinessKey = Get-QCRenditionSheetReadinessKey -FolderPath ([string]$Job.sourceFolder) -SourceDgnFileName $src
        } else {
            $readinessKey = Get-QCReadinessKey -DocumentGuid ([string]$renditionMeta.documentGuid) -FolderPath ([string]$Job.sourceFolder) -QcPdfName ([string]$renditionMeta.qcPdfName)
        }
    }

    $folder = if ($Job.sourceFolder) { [string]$Job.sourceFolder } else { '' }
    $profileRes = Resolve-QCRenditionProfile -Config $Config -FolderPath $folder
    if (-not $profileRes.IsSuccess) { return $profileRes }
    $profile = $profileRes.Data.profile

    $apiFolder = $folder
    if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
        try { $apiFolder = ConvertTo-PWCmdletFolderPath -InternalFolderPath $folder } catch { }
    }
    if (_QCR-IsNullOrWhiteSpace $apiFolder) { $apiFolder = $folder }

    $profileName = if ($profile.profileName) { [string]$profile.profileName } else { $null }
    $useFolderProfile = $true
    if ($null -ne $profile.useFolderProfileWhenUnspecified) { $useFolderProfile = [bool]$profile.useFolderProfileWhenUnspecified }
    if (_QCR-IsNullOrWhiteSpace $profileName -and $useFolderProfile) {
        $profileName = _QCR-GetFolderRenditionProfileName -ApiFolderPath $apiFolder
    }
    if (_QCR-IsNullOrWhiteSpace $profileName) {
        return New-QCFailureResult -Code 'QC_RENDITION_PROFILE_MISSING' -Message 'No rendition profileName configured or assigned to folder.' -Data @{ folderPath = $folder }
    }

    $sourceDoc = if ($Job.sourceName) { [string]$Job.sourceName } else { [System.IO.Path]::GetFileName([string]$Job.sourcePath) }
    if ($sourceDoc -notmatch '(?i)\.dgn$') {
        return New-QCFailureResult -Code 'QC_RENDITION_SOURCE_NOT_DGN' -Message 'QC_RENDITION only submits ProjectWise rendition for DGN source documents, not PDFs.' -Data @{
            sourceDocument = $sourceDoc
            folderPath = $folder
            jobId = if ($Job.id) { [string]$Job.id } else { $null }
        }
    }
    $pattern = if ($profile.sourceDocumentPattern) { [string]$profile.sourceDocumentPattern } else { '%.dgn' }

    $processors = _QCR-ToHashtable $Config.processors
    $dryRun = $false
    if ($processors -and $processors.dryRun) {
        $dr = _QCR-ToHashtable $processors.dryRun
        if ($dr -and $dr.ContainsKey('invokeHandler')) { try { $dryRun = -not [bool]$dr.invokeHandler } catch { } }
    }
    if ($Config.ContainsKey('dryRun') -and [bool]$Config.dryRun) { $dryRun = $true }

    if ($dryRun) {
        Set-QCReadinessFlag -Config $Config -ReadinessKey $readinessKey -RenditionComplete | Out-Null
        Invoke-QCReadyForQcNotificationIfReady -Config $Config -ReadinessKey $readinessKey -Job $Job `
            -DocumentName ([string]$renditionMeta.qcPdfName) -DocumentGuid ([string]$renditionMeta.documentGuid) -FolderPath $folder | Out-Null
        return New-QCSuccessResult -Code 'QC_RENDITION_DRY_RUN' -Message 'Dry-run: rendition planned, readiness marked complete.' -Data @{
            profileName = $profileName
            sourceDocument = $sourceDoc
            folderPath = $folder
            readinessKey = $readinessKey
        }
    }

    if (-not (Get-Command -Name 'Invoke-PWAuthenticatedCommand' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'QC_RENDITION_PW_UNAVAILABLE' -Message 'PW.Connection not available.' -Data @{}
    }

    $pw = _QCR-ToHashtable $Config.projectWise
    $ds = if ($pw -and $pw.datasourceName) { [string]$pw.datasourceName } else { '' }
    $credPath = if ($pw -and $pw.credentialPath) { [string]$pw.credentialPath } else { '' }
    if (_QCR-IsNullOrWhiteSpace $ds -or _QCR-IsNullOrWhiteSpace $credPath) {
        return New-QCFailureResult -Code 'QC_RENDITION_PW_CONFIG_MISSING' -Message 'projectWise.datasourceName and credentialPath are required.' -Data @{}
    }

    $submitBlock = {
        $docs = @(Get-PWDocumentsBySearch -FolderPath $script:qcrApiFolder -JustThisFolder -DocumentName $script:qcrSourceDoc -PopulatePath -ErrorAction Stop)
        if ($docs.Count -eq 0 -and $script:qcrPattern -and $script:qcrPattern -ne $script:qcrSourceDoc) {
            $docs = @(Get-PWDocumentsBySearch -FolderPath $script:qcrApiFolder -JustThisFolder -DocumentName $script:qcrPattern -PopulatePath -ErrorAction Stop)
            $docs = @($docs | Where-Object {
                $n = if ($_.Name) { [string]$_.Name } else { '' }
                $n.ToLowerInvariant() -eq $script:qcrSourceDoc.ToLowerInvariant()
            })
        }
        if ($docs.Count -eq 0) {
            return @{ ok = $false; code = 'QC_RENDITION_SOURCE_NOT_FOUND'; message = "Source document '$($script:qcrSourceDoc)' not found in folder." }
        }
        $splat = @{
            InputDocuments = @($docs | Select-Object -First 1)
            ProfileName    = $script:qcrProfileName
        }
        if ($script:qcrMakeSinglePlot) { $splat.MakeSinglePlotRequest = $true }
        New-PWRenditionRequest @splat -ErrorAction Stop
        return @{ ok = $true; code = 'QC_RENDITION_SUBMITTED'; document = $docs[0]; count = 1 }
    }

    try {
        $script:qcrApiFolder = $apiFolder
        $script:qcrSourceDoc = $sourceDoc
        $script:qcrPattern = $pattern
        $script:qcrProfileName = $profileName
        $script:qcrMakeSinglePlot = [bool]$profile.makeSinglePlotRequest
        $submitResult = Invoke-PWAuthenticatedCommand -DatasourceName $ds -CredentialPath $credPath -ScriptBlock $submitBlock
    } catch {
        return New-QCFailureResult -Code 'QC_RENDITION_PW_SESSION_FAILED' -Message $_.Exception.Message -Data @{ readinessKey = $readinessKey }
    } finally {
        Remove-Variable -Name qcrApiFolder, qcrSourceDoc, qcrPattern, qcrProfileName, qcrMakeSinglePlot -Scope Script -ErrorAction SilentlyContinue
    }

    $inner = $submitResult
    if ($inner -is [hashtable] -and $inner.ok -eq $false) {
        return New-QCFailureResult -Code ([string]$inner.code) -Message ([string]$inner.message) -Data @{ readinessKey = $readinessKey; folderPath = $folder }
    }

    $outFolder = _QCR-ResolveOutputFolderPath -Profile $profile -ApiFolderPath $apiFolder
    $completion = _QCR-ToHashtable $profile.completion
    $mode = if ($completion -and $completion.mode) { ([string]$completion.mode).Trim().ToLowerInvariant() } else { 'outputFolder' }
    $attempts = 1
    $intervalSec = 0
    if ($completion) {
        try { if ($completion.pollAttempts) { $attempts = [int]$completion.pollAttempts } } catch { }
        try { if ($completion.pollIntervalSeconds) { $intervalSec = [int]$completion.pollIntervalSeconds } } catch { }
    }
    if ($attempts -lt 1) { $attempts = 1 }

    $complete = ($mode -eq 'immediate' -or $mode -eq 'submitOnly')
    $foundData = @{}
    if (-not $complete) {
        for ($i = 0; $i -lt $attempts; $i++) {
            $check = Test-QCRenditionOutputComplete -Profile $profile -SourceDocumentName $sourceDoc -OutputFolderPath $outFolder
            if ($check.IsSuccess -and $check.Data -and $check.Data.complete -eq $true) {
                $complete = $true
                $foundData = $check.Data
                break
            }
            if ($i -lt ($attempts - 1) -and $intervalSec -gt 0) {
                Start-Sleep -Seconds $intervalSec
            }
        }
    }

    if (-not $complete) {
        return New-QCFailureResult -Code 'QC_RENDITION_OUTPUT_TIMEOUT' -Message 'Rendition output was not detected within the configured poll window.' -Data @{
            readinessKey = $readinessKey
            profileName = $profileName
            outputFolder = $outFolder
            pollAttempts = $attempts
        }
    }

    Set-QCReadinessFlag -Config $Config -ReadinessKey $readinessKey -RenditionComplete | Out-Null
    $notif = Invoke-QCReadyForQcNotificationIfReady -Config $Config -ReadinessKey $readinessKey -Job $Job `
        -DocumentName ([string]$renditionMeta.qcPdfName) -DocumentGuid ([string]$renditionMeta.documentGuid) -FolderPath $folder

    $submitOnly = ($mode -eq 'immediate' -or $mode -eq 'submitOnly')
    $resultCode = if ($submitOnly) { 'QC_RENDITION_SUBMITTED' } else { 'QC_RENDITION_SUCCEEDED' }
    $resultMsg = if ($submitOnly) {
        'Rendition request submitted to ProjectWise; job complete (no output polling).'
    } else {
        'Rendition submitted and output verified.'
    }

    return New-QCSuccessResult -Code $resultCode -Message $resultMsg -Data @{
        readinessKey = $readinessKey
        profileName = $profileName
        completionMode = $mode
        output = $foundData
        notification = $notif
    }
}

Export-ModuleMember -Function *
