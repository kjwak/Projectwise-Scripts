# QC.Processors.psm1
# Responsibility: Processor readiness checks and job-type-based dispatch.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Runtime.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Workflow/QC.Workflow.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Reporting/QC.Reporting.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Processing/QC.CommentStatusProcessor.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Processing/QC.ReviewStamp.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Workflow/QC.ProcessType.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Processing/QC.Rendition.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Processing/QC.PrependOverlayPolicy.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Workflow/QC.AuditTriggers.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Workflow/QC.ProcessType.psm1') -Force -ErrorAction SilentlyContinue

# Per-process throttle (milliseconds) inserted between PDF/cache file ops
# (Move/Remove/Copy) so AV scanners (Fortinet, etc.) don't flag rapid temp
# churn. Default 2000 ms; override via top-level Config.fileOpThrottleMs.
# Set to 0 to disable.
$script:_QCP_FsThrottleMs = 2000

function _QCP-FsThrottle {
    if ($script:_QCP_FsThrottleMs -and $script:_QCP_FsThrottleMs -gt 0) {
        Start-Sleep -Milliseconds $script:_QCP_FsThrottleMs
    }
}

function _QCP-ApplyFsThrottleConfig([hashtable]$Config) {
    if (-not $Config) { return }
    try {
        if ($Config.ContainsKey('fileOpThrottleMs') -and $null -ne $Config['fileOpThrottleMs']) {
            $script:_QCP_FsThrottleMs = [int]$Config['fileOpThrottleMs']
        }
    } catch { }
}

function _QCP-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCP-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) { return $Value }
    if ($Value -is [string] -or $Value -is [System.ValueType]) { return @{ value = $Value } }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _QCP-TruncateTextForStorage {
    param(
        [string]$Text,
        [int]$MaxLen = 8192
    )
    if ([string]::IsNullOrEmpty($Text)) { return $null }
    if ($Text.Length -le $MaxLen) { return $Text }
    $omitted = $Text.Length - $MaxLen
    return ('...(truncated ' + $omitted + ' chars)...' + $Text.Substring($Text.Length - $MaxLen))
}

function _QCP-SanitizeProcessorDataForStorage([object]$Data) {
    $h = _QCP-ToHashtable $Data
    if (-not $h) { return $null }
    $out = @{}
    foreach ($k in $h.Keys) {
        $v = $h[$k]
        if ($k -in @('stdout', 'stderr', 'argLine')) {
            if ($null -ne $v) { $out[$k] = _QCP-TruncateTextForStorage -Text ([string]$v) }
            continue
        }
        if ($k -eq 'args' -and $v -is [array]) {
            $out[$k] = @($v | ForEach-Object { [string]$_ })
            continue
        }
        $out[$k] = $v
    }
    return $out
}


function _QCP-TestQcPrependProjectWiseMode {
    param([string]$Mode)
    $m = ([string]$Mode).Trim().ToLowerInvariant()
    return $m -in @('legacypw', 'projectwise', 'pw')
}

function Test-QCFastAuditEnqueueEnabled {
    <#
    .SYNOPSIS
    True when qcPrepend.fastAuditEnqueue is enabled (watcher enqueues before sibling sync).
    #>
    [CmdletBinding()]
    param([hashtable]$Config)
    if (-not $Config) { return $false }
    try {
        if ($Config.ContainsKey('qcPrepend') -and $Config.qcPrepend) {
            $qc = $Config.qcPrepend
            if ($qc -is [hashtable] -and $qc.ContainsKey('fastAuditEnqueue')) {
                return [bool]$qc.fastAuditEnqueue
            }
            if ($qc.PSObject -and $qc.PSObject.Properties['fastAuditEnqueue']) {
                return [bool]$qc.fastAuditEnqueue
            }
        }
    } catch { }
    return $false
}

function _QCP-AttachSheetIndexProcessTypeHint {
    param(
        [hashtable]$Job,
        [hashtable]$Config,
        [string]$FolderPath,
        [string]$SheetPdfName
    )
    if (-not $Job) { return }
    $hint = ''
    try {
        $hint = [string](_QCP-ResolveProcessTypeFromSheetIndex -Config $Config -FolderPath $FolderPath -SourceDocumentName $SheetPdfName)
    } catch { $hint = '' }
    if (_QCP-IsNullOrWhiteSpace $hint) { return }
    if (-not $Job.ContainsKey('metadata') -or -not $Job.metadata) { $Job['metadata'] = @{} }
    $md = _QCP-ToHashtable $Job.metadata
    if (-not $md) { $md = @{} }
    if (_QCP-IsNullOrWhiteSpace ([string]$md['qcProcessType'])) {
        $md['qcProcessType'] = $hint
        $Job['metadata'] = $md
    }
}

function Invoke-QCFastAuditPrependEnqueue {
    <#
    .SYNOPSIS
    Cheap DOCUMENT_STATE prepend enqueue after the watcher live-state read (no sibling sync).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$TriggerDocumentGuid,
        [Parameter(Mandatory)][string]$TriggerDocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$LiveStateName = '',
        [bool]$IsSheetsFolder = $false,
        [bool]$EnableQcPrepend = $true,
        [bool]$DryRun = $false,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [string]$LastAuditEventAt = '',
        [Nullable[long]]$AuditEventId = $null
    )
    $out = @{
        attempted = $false
        skipPathD = $false
        enqueued  = $false
        code      = ''
        reason    = 'not_eligible'
    }
    if (-not $EnableQcPrepend) {
        $out.reason = 'prepend_disabled'
        return $out
    }
    if (-not $IsSheetsFolder) {
        $out.reason = 'not_sheets_folder'
        return $out
    }
    if (Get-Command -Name 'Test-QCIsStatusSetOutputPdfName' -ErrorAction SilentlyContinue) {
        if (Test-QCIsStatusSetOutputPdfName -FileName $TriggerDocumentName) {
            $out.reason = 'status_set_output'
            return $out
        }
    }
    if (-not (Get-Command -Name 'Test-QCIsSheetPdfDocumentName' -ErrorAction SilentlyContinue) `
        -or -not (Test-QCIsSheetPdfDocumentName -DocumentName $TriggerDocumentName)) {
        $out.reason = 'not_stem_sheet_pdf'
        return $out
    }
    if (Get-Command -Name 'Test-QCIsAutomationPwActor' -ErrorAction SilentlyContinue) {
        if (Test-QCIsAutomationPwActor -Config $Config -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername) {
            $out.reason = 'automation_actor'
            return $out
        }
    }
    $state = ([string]$LiveStateName).Trim()
    if ([string]::IsNullOrWhiteSpace($state)) {
        $out.reason = 'empty_live_state'
        return $out
    }
    $isInitiated = $false
    $isFinalizing = $false
    if (Get-Command -Name 'Test-QCWorkflowStateIsQcInitiated' -ErrorAction SilentlyContinue) {
        try { $isInitiated = [bool](Test-QCWorkflowStateIsQcInitiated -StateName $state -Config $Config) } catch { $isInitiated = $false }
    }
    if (Get-Command -Name 'Test-QCWorkflowStateIsQcFinalizing' -ErrorAction SilentlyContinue) {
        try { $isFinalizing = [bool](Test-QCWorkflowStateIsQcFinalizing -StateName $state -Config $Config) } catch { $isFinalizing = $false }
    }
    if (-not $isInitiated -and -not $isFinalizing) {
        $out.reason = 'state_not_prepend'
        return $out
    }

    $out.attempted = $true
    $enq = $null
    if ($isInitiated -and (Get-Command -Name 'Add-QCPrependJobForQcInitiatedStateChange' -ErrorAction SilentlyContinue)) {
        $enq = Add-QCPrependJobForQcInitiatedStateChange -Config $Config -TriggerDocumentGuid $TriggerDocumentGuid `
            -TriggerDocumentName $TriggerDocumentName -FolderPath $FolderPath -CurrentStateName $state `
            -DryRun:$DryRun -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId
    } elseif ($isFinalizing -and (Get-Command -Name 'Add-QCPrependJobForQcFinalizingStateChange' -ErrorAction SilentlyContinue)) {
        $enq = Add-QCPrependJobForQcFinalizingStateChange -Config $Config -TriggerDocumentGuid $TriggerDocumentGuid `
            -TriggerDocumentName $TriggerDocumentName -FolderPath $FolderPath -CurrentStateName $state `
            -DryRun:$DryRun -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId
    } else {
        $out.reason = 'enqueue_cmd_missing'
        return $out
    }

    if ($null -eq $enq) {
        $out.reason = 'enqueue_returned_null'
        return $out
    }
    $code = ''
    try { $code = [string]$enq.Code } catch { $code = '' }
    $out.code = $code
    $ok = $false
    try { $ok = [bool]$enq.IsSuccess } catch { $ok = $false }
    # Skip Path D after a real enqueue attempt (including dedupe / planned / sheet-active).
    $out.skipPathD = $true
    if ($ok -and ($code -eq 'QC_PREPEND_ENQUEUED' -or $code -eq 'QUEUE_ENQUEUED' -or $code -eq 'QC_PREPEND_PLANNED')) {
        $out.enqueued = $true
        $out.reason = 'enqueued'
    } elseif ($ok) {
        $out.reason = $(if ($code) { $code } else { 'enqueue_skipped' })
    } else {
        $out.reason = $(if ($code) { $code } else { 'enqueue_failed' })
    }
    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_FAST_ENQUEUE' `
            -Message 'Fast audit enqueue evaluated DOCUMENT_STATE prepend without sibling sync.' -Data @{
            triggerDocumentName = $TriggerDocumentName
            triggerDocumentGuid = $TriggerDocumentGuid
            folderPath = $FolderPath
            liveState = $state
            enqueued = [bool]$out.enqueued
            skipPathD = [bool]$out.skipPathD
            resultCode = $out.code
            reason = $out.reason
        } | Out-Null
    }
    return $out
}

function Invoke-QCPrependPreflight {
    <#
    .SYNOPSIS
    Confirm live workflow state, DGN pair, and lane before overlay. Skip = success (no overlay).
    When PW cmdlets are missing, proceed so local/checkpoint tests still run.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$IncomingFolder = '',
        [string]$IncomingDocName = ''
    )
    $guid = ''
    try { $guid = [string](_QCP-GetJobMetadataValue -Job $Job -Keys @('triggerDocumentGuid', 'sourceDocumentGuid', 'documentGuid')) } catch { $guid = '' }
    $folder = $IncomingFolder
    if (_QCP-IsNullOrWhiteSpace $folder) {
        try { $folder = [string](_QCP-GetJobMetadataValue -Job $Job -Keys @('sourceFolder', 'folderPath', 'incomingFolderPath')) } catch { $folder = '' }
        if (_QCP-IsNullOrWhiteSpace $folder -and $Job.sourceFolder) { $folder = [string]$Job.sourceFolder }
    }
    $docName = $IncomingDocName
    if (_QCP-IsNullOrWhiteSpace $docName) {
        if ($Job.sourceName) { $docName = [string]$Job.sourceName }
        elseif ($Job.sourcePath) { $docName = [System.IO.Path]::GetFileName([string]$Job.sourcePath) }
    }
    if ($folder -match '^(?i)Documents\\') {
        $folder = ($folder -replace '^(?i)Documents\\', '')
    }

    $hasStateCmd = [bool](Get-Command -Name 'Get-PWDocumentWorkflowStateName' -ErrorAction SilentlyContinue)
    $hasPairCmd = [bool](Get-Command -Name 'Test-PWSheetPdfHasMatchingPair' -ErrorAction SilentlyContinue)
    if (-not $hasStateCmd -and -not $hasPairCmd) {
        return New-QCSuccessResult -Code 'QC_PREPEND_PREFLIGHT_BYPASS' -Message 'Preflight bypassed (PW discovery cmdlets not loaded).' -Data @{
            proceed = $true; skipped = $false; reason = 'cmdlets_missing'
        }
    }

    $trigger = _QCP-ResolvePrependTrigger -Job $Job
    if ($hasStateCmd -and -not (_QCP-IsNullOrWhiteSpace $folder) -and -not (_QCP-IsNullOrWhiteSpace $docName)) {
        $liveState = ''
        try {
            $liveState = [string](Get-PWDocumentWorkflowStateName -FolderPath $folder -DocumentName $docName -DocumentGuid $guid)
        } catch { $liveState = '' }
        $stateOk = $false
        if ($trigger -eq 'finalQcComplete') {
            if (Get-Command -Name 'Test-QCWorkflowStateIsQcFinalizing' -ErrorAction SilentlyContinue) {
                try { $stateOk = [bool](Test-QCWorkflowStateIsQcFinalizing -StateName $liveState -Config $Config) } catch { $stateOk = $false }
            }
            if (-not $stateOk -and (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue)) {
                try {
                    $wf = Get-QCWorkflowSettings -Config $Config
                    $want = [string](Get-QCWorkflowStateName -Settings $wf -StateKey 'qcFinalizing')
                    if (-not [string]::IsNullOrWhiteSpace($want) -and -not [string]::IsNullOrWhiteSpace($liveState)) {
                        $stateOk = ($liveState.Trim().ToLowerInvariant() -eq $want.Trim().ToLowerInvariant())
                    }
                } catch { }
            }
        } else {
            if (Get-Command -Name 'Test-QCWorkflowStateIsQcInitiated' -ErrorAction SilentlyContinue) {
                try { $stateOk = [bool](Test-QCWorkflowStateIsQcInitiated -StateName $liveState -Config $Config) } catch { $stateOk = $false }
            }
            if (-not $stateOk -and (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue)) {
                try {
                    $wf = Get-QCWorkflowSettings -Config $Config
                    $want = [string](Get-QCWorkflowStateName -Settings $wf -StateKey 'qcInitiated')
                    if (-not [string]::IsNullOrWhiteSpace($want) -and -not [string]::IsNullOrWhiteSpace($liveState)) {
                        $stateOk = ($liveState.Trim().ToLowerInvariant() -eq $want.Trim().ToLowerInvariant())
                    }
                } catch { }
            }
        }
        if (-not $stateOk) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_SKIPPED_NOT_QC_INITIATED' `
                    -Message 'QC_PREPEND skipped: live workflow state is not a prepend trigger state.' -Data @{
                    jobId = [string]$Job.id; liveState = $liveState; prependTrigger = $trigger
                    folderPath = $folder; sourceName = $docName
                } | Out-Null
            }
            return New-QCSuccessResult -Code 'QC_PREPEND_SKIPPED_NOT_QC_INITIATED' `
                -Message 'QC_PREPEND skipped because live workflow state is not QC Initiated / QC Finalizing.' -Data @{
                skipped = $true; proceed = $false; liveState = $liveState; prependTrigger = $trigger
            }
        }
    }

    if ($hasPairCmd -and -not (_QCP-IsNullOrWhiteSpace $folder) -and -not (_QCP-IsNullOrWhiteSpace $docName)) {
        $hasPair = $false
        try { $hasPair = [bool](Test-PWSheetPdfHasMatchingPair -FolderPath $folder -DocumentName $docName) } catch { $hasPair = $false }
        if (-not $hasPair) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_SKIPPED_NO_DGN_PAIR' `
                    -Message 'QC_PREPEND skipped: no matching DGN pair for sheet stem.' -Data @{
                    jobId = [string]$Job.id; folderPath = $folder; sourceName = $docName
                } | Out-Null
            }
            return New-QCSuccessResult -Code 'QC_PREPEND_SKIPPED_NO_DGN_PAIR' `
                -Message 'QC_PREPEND skipped because the sheet PDF has no matching DGN pair.' -Data @{
                skipped = $true; proceed = $false; folderPath = $folder; sourceName = $docName
            }
        }
    }

    $laneRes = _QCP-TryResolvePrependLaneContext -Job $Job -Config $Config
    if (-not $laneRes.IsSuccess) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_SKIPPED_NO_PROCESS_TYPE' `
                -Message 'QC_PREPEND skipped at preflight: qc_process_type could not be resolved.' -Data @{
                jobId = [string]$Job.id; folderPath = $folder; sourceName = $docName
                errorCode = [string]$laneRes.Code; errorMessage = [string]$laneRes.Message
            } | Out-Null
        }
        return New-QCSuccessResult -Code 'QC_PREPEND_SKIPPED_NO_PROCESS_TYPE' `
            -Message 'QC_PREPEND skipped because qc_process_type could not be resolved.' -Data @{
            skipped = $true; proceed = $false; errorCode = [string]$laneRes.Code
        }
    }
    return New-QCSuccessResult -Code 'QC_PREPEND_PREFLIGHT_OK' -Message 'QC_PREPEND preflight passed.' -Data @{
        proceed = $true; skipped = $false; lane = [hashtable]$laneRes.Data
    }
}

function _QCP-ResolveProjectWisePrependScriptPath {
    param(
        [Parameter(Mandatory)][hashtable]$QcPrependConfig,
        [Parameter(Mandatory)][string]$RepoRoot
    )
    if ($QcPrependConfig.ContainsKey('projectWiseScriptPath') -and $QcPrependConfig.projectWiseScriptPath) {
        return [string]$QcPrependConfig.projectWiseScriptPath
    }
    if ($QcPrependConfig.ContainsKey('legacyScriptPath') -and $QcPrependConfig.legacyScriptPath) {
        return [string]$QcPrependConfig.legacyScriptPath
    }
    return (Join-Path $RepoRoot 'scripts\processing\Invoke-QCPrependPw.ps1')
}

function _QCP-GetRepoRoot() {
    $parent = Split-Path -Parent $PSScriptRoot
    if ((Split-Path -Leaf $parent) -eq 'modules') {
        return (Split-Path -Parent $parent)
    }
    return $parent
}

function _QCP-GetReviewStampRoleFieldsFromJob {
    param(
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$FolderPath = '',
        [string]$SourceDocumentName = ''
    )

    $roles = @{
        designerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('designerEmail', 'qcDesignerEmail')
        reviewerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewerEmail', 'qcReviewerEmail')
        checkerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('checkerEmail', 'qcCheckerEmail')
        qcReviewType = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewType', 'qcReviewType')
    }

    if (-not (_QCP-IsNullOrWhiteSpace $FolderPath) -and -not (_QCP-IsNullOrWhiteSpace $SourceDocumentName)) {
        if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
            $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName -Config $Config
            if ($pw.found) {
                $roles.designerEmail = [string]$pw.designerEmail
                $roles.reviewerEmail = [string]$pw.reviewerEmail
                $roles.checkerEmail = [string]$pw.checkerEmail
            }
        }
        if (Get-Command -Name 'Get-PWQcPrependProcessIntentFromSourcePdf' -ErrorAction SilentlyContinue) {
            $intent = Get-PWQcPrependProcessIntentFromSourcePdf -FolderPath $FolderPath `
                -SourceDocumentName $SourceDocumentName -Config $Config
            if ($intent.found -and -not (_QCP-IsNullOrWhiteSpace $intent.qcProcessType)) {
                $roles.qcProcessType = [string]$intent.qcProcessType
            }
        }
    }

    return $roles
}

function _QCP-IsStemSheetDocumentName {
    param([string]$DocumentName)
    if (_QCP-IsNullOrWhiteSpace $DocumentName) { return $false }
    if ([string]$DocumentName -match '(?i)\.dgn$') { return $true }
    if ([string]$DocumentName -match '(?i)\.pdf$' -and [string]$DocumentName -notmatch '(?i)-(prod|chk|rev|qc)\.pdf$') { return $true }
    return $false
}

function _QCP-ResolveProcessTypeFromSheetIndex {
    param(
        [hashtable]$Config,
        [string]$FolderPath,
        [string]$SourceDocumentName
    )
    if (-not (Get-Command -Name 'Invoke-QCDatabaseQuery' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return '' }
    if ((_QCP-IsNullOrWhiteSpace $FolderPath) -or (_QCP-IsNullOrWhiteSpace $SourceDocumentName)) { return '' }

    $stem = [System.IO.Path]::GetFileNameWithoutExtension([string]$SourceDocumentName)
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        $resolvedStem = Get-PWSheetStemFromDocumentName -DocumentName $SourceDocumentName
        if (-not (_QCP-IsNullOrWhiteSpace $resolvedStem)) { $stem = [string]$resolvedStem }
    }
    if (_QCP-IsNullOrWhiteSpace $stem) { return '' }

    $pdfName = $stem + '.pdf'
    $dgnName = $stem + '.dgn'
    $folderCandidates = @([string]$FolderPath)
    if ($FolderPath -notmatch '^(?i)Documents\\') {
        $folderCandidates += ('Documents\' + $FolderPath.Trim().TrimEnd('\'))
    }

    foreach ($fp in @($folderCandidates)) {
        if (_QCP-IsNullOrWhiteSpace $fp) { continue }
        try {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 qc_process_type
FROM sheet_index
WHERE folder_path = @folderPath
  AND qc_process_type IS NOT NULL
  AND LTRIM(RTRIM(qc_process_type)) <> ''
  AND (LOWER(document_name) = LOWER(@pdfName) OR LOWER(document_name) = LOWER(@dgnName))
ORDER BY CASE WHEN LOWER(document_name) = LOWER(@pdfName) THEN 0 WHEN LOWER(document_name) = LOWER(@dgnName) THEN 1 ELSE 2 END
"@ -Parameters @{ folderPath = $fp; pdfName = $pdfName; dgnName = $dgnName }
            if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
                $raw = if ($res.Data.table.Rows[0].qc_process_type -is [DBNull]) { '' } else { [string]$res.Data.table.Rows[0].qc_process_type }
                if (-not (_QCP-IsNullOrWhiteSpace $raw)) {
                    $norm = _QCP-NormalizePrependProcessTypeValue -RawProcessType $raw
                    if ($norm) { return [string]$norm }
                    return $raw.Trim()
                }
            }
        } catch { }
    }
    return ''
}

function _QCP-NormalizePrependProcessTypeValue {
    param([string]$RawProcessType)
    if (_QCP-IsNullOrWhiteSpace $RawProcessType) { return $null }
    if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
        return Normalize-QCProcessType -ProcessType ([string]$RawProcessType) -AllowNullOnEmpty
    }
    return $null
}

function _QCP-ResolveProcessTypeFromJob {
    param(
        [hashtable]$Job,
        [hashtable]$Config
    )
    $laneRes = _QCP-TryResolvePrependLaneContext -Job $Job -Config $Config
    if ($laneRes.IsSuccess -and $laneRes.Data -and $laneRes.Data.qcProcessType) {
        return [string]$laneRes.Data.qcProcessType
    }
    return $null
}

function _QCP-EnsureJobMetadataHashtable {
    param([hashtable]$Job)
    if (-not $Job) { return @{} }
    if (-not $Job.ContainsKey('metadata') -or -not $Job.metadata) {
        $Job['metadata'] = @{}
    } elseif (-not ($Job.metadata -is [hashtable])) {
        $md = _QCP-ToHashtable $Job.metadata
        if (-not $md) { $md = @{} }
        $Job['metadata'] = $md
    }
    return [hashtable]$Job.metadata
}

function _QCP-LogPrependLaneResolved {
    param(
        [hashtable]$Job,
        [hashtable]$LaneData,
        [string]$Stage = 'before_execution'
    )
    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    if (-not $LaneData) { return }
    Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_LANE_RESOLVED' `
        -Message 'QC_PREPEND lane context resolved.' -Data @{
        jobId = if ($Job -and $Job.id) { [string]$Job.id } else { '' }
        stage = $Stage
        qc_process_type = [string]$LaneData.qcProcessType
        qcProcessType = [string]$LaneData.qcProcessType
        pdfSuffix = [string]$LaneData.pdfSuffix
        expectedLanePdfName = [string]$LaneData.expectedLanePdfName
        triggerDocumentName = [string]$LaneData.triggerDocumentName
        sourceDocumentGuid = [string]$LaneData.sourceDocumentGuid
        resolutionSource = [string]$LaneData.resolutionSource
    } | Out-Null
}

function _QCP-LogPrependProgress {
    param(
        [hashtable]$Job,
        [Parameter(Mandatory)][string]$Code,
        [Parameter(Mandatory)][string]$Message,
        [hashtable]$Config = $null,
        [hashtable]$Extra = $null
    )
    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    $data = @{
        jobId = if ($Job -and $Job.id) { [string]$Job.id } else { '' }
        jobType = 'QC_PREPEND'
    }
    if ($Job) {
        $src = _QCP-GetJobMetadataValue -Job $Job -Keys @('sourceName', 'triggerDocumentName', 'documentName', 'incomingDocName')
        if (-not (_QCP-IsNullOrWhiteSpace $src)) { $data['sourceName'] = [string]$src }
    }
    if ($Extra) {
        foreach ($k in $Extra.Keys) { $data[$k] = $Extra[$k] }
    }
    Write-QCJsonLog -Level 'Information' -Code $Code -Message $Message -Data $data | Out-Null
    if ($Config -and $Job -and $Job.id -and (Get-Command -Name 'Update-QCJobHeartbeat' -ErrorAction SilentlyContinue)) {
        try { Update-QCJobHeartbeat -JobId ([string]$Job.id) -Config $Config -Job $Job | Out-Null } catch { }
    }
}

$script:_QCP_CheckpointPrependComplete = 'prepend_complete'
$script:_QCP_CheckpointWritebackRunning = 'writeback_running'
$script:_QCP_CheckpointWritebackComplete = 'writeback_complete'

function _QCP-GetJobCheckpointName {
    param([hashtable]$Job)
    if (-not $Job) { return '' }
    if ($Job.ContainsKey('checkpoint') -and -not (_QCP-IsNullOrWhiteSpace $Job.checkpoint)) {
        return ([string]$Job.checkpoint).Trim()
    }
    return ''
}

function _QCP-TestPrependWritebackOnlyResume {
    param([hashtable]$Job)
    $cp = _QCP-GetJobCheckpointName -Job $Job
    return ($cp -eq $script:_QCP_CheckpointPrependComplete -or $cp -eq $script:_QCP_CheckpointWritebackRunning)
}

function _QCP-TestPrependAlreadyComplete {
    param([hashtable]$Job)
    return ((_QCP-GetJobCheckpointName -Job $Job) -eq $script:_QCP_CheckpointWritebackComplete)
}

function _QCP-SetPrependJobCheckpoint {
    param(
        [hashtable]$Job,
        [hashtable]$Config,
        [Parameter(Mandatory)][string]$Checkpoint,
        [hashtable]$CheckpointData = $null
    )
    if (-not $Job) { return }
    $dataText = ''
    if ($CheckpointData) {
        try { $dataText = ($CheckpointData | ConvertTo-Json -Compress -Depth 6) } catch { $dataText = '' }
    }
    if (Get-Command -Name 'Set-QCJobCheckpoint' -ErrorAction SilentlyContinue -and $Job.id -and $Config) {
        try {
            Set-QCJobCheckpoint -JobId ([string]$Job.id) -Config $Config -Job $Job -Checkpoint $Checkpoint -CheckpointData $dataText | Out-Null
            return
        } catch { }
    }
    $Job.checkpoint = [string]$Checkpoint
    if ($dataText) { $Job.checkpointData = $dataText }
}

function _QCP-ReadProcessRedirectFile {
    param(
        [string]$Path,
        [int]$Retries = 8,
        [int]$DelayMs = 150
    )
    if (_QCP-IsNullOrWhiteSpace $Path) { return '' }
    for ($i = 0; $i -lt $Retries; $i++) {
        try {
            if (-not (Test-Path -LiteralPath $Path)) { return '' }
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                $sr = New-Object System.IO.StreamReader($fs)
                try { return [string]$sr.ReadToEnd() } finally { $sr.Dispose() }
            } finally { $fs.Dispose() }
        } catch {
            Start-Sleep -Milliseconds $DelayMs
        }
    }
    return [string](Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue)
}

function _QCP-WaitForLaunchedProcess {
    <#
    .SYNOPSIS
    Wait for a specific Start-Process -PassThru process to exit. Does not wait for descendants
    or for redirected stdout/stderr handles held by Bentley/ProjectWise child processes.
    #>
    param(
        [Parameter(Mandatory)][System.Diagnostics.Process]$Process,
        [hashtable]$Job = $null,
        [hashtable]$Config = $null,
        [int]$PollMilliseconds = 500,
        [int]$TimeoutSeconds = 0,
        [int]$HeartbeatSeconds = 15
    )
    if ($PollMilliseconds -lt 50) { $PollMilliseconds = 50 }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $lastHb = [DateTime]::UtcNow
    $pidValue = 0
    try { $pidValue = [int]$Process.Id } catch { $pidValue = 0 }
    # Touching Handle keeps the native process handle alive so ExitCode is readable
    # after the PID exits, without WaitForExit() (which waits on redirected descendants).
    $keepHandle = $null
    try { $keepHandle = $Process.Handle } catch { $keepHandle = $null }
    while ($true) {
        $alive = $false
        if ($pidValue -gt 0) {
            $alive = [bool](Get-Process -Id $pidValue -ErrorAction SilentlyContinue)
        } else {
            try {
                $Process.Refresh()
                $alive = -not [bool]$Process.HasExited
            } catch { $alive = $false }
        }
        if (-not $alive) {
            $code = $null
            try {
                $Process.Refresh()
                $code = [int]$Process.ExitCode
            } catch { $code = $null }
            [GC]::KeepAlive($Process)
            [void]$keepHandle
            return @{
                exited = $true
                timedOut = $false
                exitCode = $code
                elapsedMs = [int]$sw.ElapsedMilliseconds
                processId = $pidValue
            }
        }
        if ($TimeoutSeconds -gt 0 -and $sw.Elapsed.TotalSeconds -ge $TimeoutSeconds) {
            [GC]::KeepAlive($Process)
            [void]$keepHandle
            return @{
                exited = $false
                timedOut = $true
                exitCode = $null
                elapsedMs = [int]$sw.ElapsedMilliseconds
                processId = $pidValue
            }
        }
        if ($Job -and $Config -and $HeartbeatSeconds -gt 0 -and (([DateTime]::UtcNow - $lastHb).TotalSeconds -ge $HeartbeatSeconds)) {
            _QCP-LogPrependProgress -Job $Job -Config $Config -Code 'QC_PREPEND_PW_CHILD_WAIT' `
                -Message 'Waiting for ProjectWise prepend child PowerShell process (descendants ignored).' -Extra @{
                childPid = $pidValue
                elapsedMs = [int]$sw.ElapsedMilliseconds
            }
            $lastHb = [DateTime]::UtcNow
        }
        Start-Sleep -Milliseconds $PollMilliseconds
    }
}

function _QCP-StartAndWaitLaunchedProcess {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$ArgumentList,
        [Parameter(Mandatory)][string]$StdoutPath,
        [Parameter(Mandatory)][string]$StderrPath,
        [hashtable]$Job = $null,
        [hashtable]$Config = $null,
        [int]$TimeoutSeconds = 0,
        [int]$PollMilliseconds = 500,
        [int]$HeartbeatSeconds = 15
    )
    $p = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -NoNewWindow -PassThru `
        -RedirectStandardOutput $StdoutPath -RedirectStandardError $StderrPath
    if (-not $p) {
        return @{
            process = $null
            processId = 0
            exited = $false
            timedOut = $false
            exitCode = $null
            elapsedMs = 0
            stdout = ''
            stderr = ''
            startFailed = $true
        }
    }
    $startedPid = 0
    try { $startedPid = [int]$p.Id } catch { $startedPid = 0 }
    try { [void]$p.Handle } catch { }
    _QCP-LogPrependProgress -Job $Job -Config $Config -Code 'QC_PREPEND_PW_CHILD_PID' `
        -Message 'ProjectWise prepend child PowerShell PID retained; waiting for that process only.' -Extra @{
        childPid = $startedPid
    }
    $wait = _QCP-WaitForLaunchedProcess -Process $p -Job $Job -Config $Config `
        -TimeoutSeconds $TimeoutSeconds -PollMilliseconds $PollMilliseconds -HeartbeatSeconds $HeartbeatSeconds
    $stdout = _QCP-ReadProcessRedirectFile -Path $StdoutPath
    $stderr = _QCP-ReadProcessRedirectFile -Path $StderrPath
    return @{
        process = $p
        processId = [int]$wait.processId
        exited = [bool]$wait.exited
        timedOut = [bool]$wait.timedOut
        exitCode = $wait.exitCode
        elapsedMs = [int]$wait.elapsedMs
        stdout = $stdout
        stderr = $stderr
        startFailed = $false
    }
}

function _QCP-GetPrependChildWaitTimeoutSeconds {
    param([hashtable]$QcPrependConfig)
    if (-not $QcPrependConfig) { return 0 }
    if ($QcPrependConfig.ContainsKey('childWaitTimeoutSeconds') -and $null -ne $QcPrependConfig.childWaitTimeoutSeconds) {
        try { return [int]$QcPrependConfig.childWaitTimeoutSeconds } catch { return 0 }
    }
    return 0
}

function _QCP-InvokePrependPostSuccessWriteback {
    param(
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][object]$SuccessResult,
        [string]$IncomingFolder,
        [string]$IncomingDocName,
        [string]$DatasourceName,
        [hashtable]$ProjectWiseCfg,
        [switch]$ClearTriggerTag
    )

    $tagCleared = $null
    $tagClearError = $null
    $writebackResult = $null
    $connected = $false
    try {
        _QCP-SetPrependJobCheckpoint -Job $Job -Config $Config -Checkpoint $script:_QCP_CheckpointWritebackRunning
        _QCP-LogPrependProgress -Job $Job -Config $Config -Code 'QC_PREPEND_PW_RECONNECT_START' `
            -Message 'Reconnecting to ProjectWise for post-prepend workflow writeback.' -Extra @{
            incomingDocName = $IncomingDocName
        }
        $repoRoot = _QCP-GetRepoRoot
        $pwConnMod = Join-Path $repoRoot 'modules\ProjectWise\PW.Connection.psm1'
        if (Test-Path -LiteralPath $pwConnMod) { Import-Module $pwConnMod -Force -ErrorAction SilentlyContinue | Out-Null }

        $pwCfg = if ($ProjectWiseCfg) { $ProjectWiseCfg } else { @{} }
        $credPath = if ($pwCfg.ContainsKey('credentialPath') -and $pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
        $credRes = Get-PWCredentialFromFile -CredentialPath $credPath
        if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }
        $connRes = Connect-PW -DatasourceName $DatasourceName -Credential ([pscredential]$credRes.Data.credential)
        if (-not $connRes.IsSuccess) { throw ($connRes.Code + ': ' + $connRes.Message) }
        $connected = $true

        $discRes = Ensure-PWDiscoveryCmdlets
        if (-not $discRes.IsSuccess -or -not (Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue)) {
            throw ('QC_PREPEND_PW_DISCOVERY_INCOMPLETE: ProjectWise document search cmdlet Get-PWDocumentsBySearch is missing after connect/re-import.')
        }

        $doc = Get-PWDocumentsBySearch -FolderPath $IncomingFolder -JustThisFolder -DocumentName $IncomingDocName -PopulatePath -ErrorAction Stop
        if ($ClearTriggerTag.IsPresent) {
            _QCP-LogPrependProgress -Job $Job -Config $Config -Code 'QC_PREPEND_TAG_CLEAR_START' `
                -Message 'Clearing QC_Archivist trigger tag.' -Extra @{ incomingDocName = $IncomingDocName }
            if ($doc) {
                $currentDesc = $doc.Description
                if ($null -eq $currentDesc -and $doc.PSObject.Properties['Description']) { $currentDesc = $doc.Description }
                if ($null -eq $currentDesc) { $currentDesc = '' }
                $newDesc = ([string]$currentDesc -replace [regex]::Escape('QC_Archivist'), '').Trim()
                $doc.Description = $newDesc
                Update-PWDocumentProperties $doc -ErrorAction Stop
                $tagCleared = $true
            } else {
                $tagCleared = $false
                $tagClearError = 'Could not re-find document to clear trigger tag.'
            }
            _QCP-LogPrependProgress -Job $Job -Config $Config -Code 'QC_PREPEND_TAG_CLEAR_DONE' `
                -Message $(if ($tagCleared) { 'QC_Archivist trigger tag cleared.' } else { 'QC_Archivist trigger tag clear skipped or failed.' }) -Extra @{
                triggerTagCleared = $tagCleared
                triggerTagClearError = $tagClearError
                incomingDocName = $IncomingDocName
            }
        }

        if ($doc) { $Job['document'] = $doc }
        if ($SuccessResult.Data) {
            $SuccessResult.Data.triggerTagCleared = $tagCleared
            $SuccessResult.Data.triggerTagClearError = $tagClearError
        }

        _QCP-LogPrependProgress -Job $Job -Config $Config -Code 'QC_PREPEND_WORKFLOW_WRITEBACK_START' `
            -Message 'Workflow writeback starting (ProjectWise connected).' -Extra @{ incomingDocName = $IncomingDocName }
        $writebackResult = _QCP-AppendWorkflowWriteback -Result $SuccessResult -Job $Job -Config $Config `
            -SourcePath $IncomingDocName -OutputPath $null -HistoryPath $null
        if ($writebackResult -and $writebackResult.IsSuccess) {
            _QCP-SetPrependJobCheckpoint -Job $Job -Config $Config -Checkpoint $script:_QCP_CheckpointWritebackComplete
        }
        return $writebackResult
    } catch {
        $tagClearError = [string]$_.Exception.Message
        try {
            if ($SuccessResult.Data) {
                $SuccessResult.Data.triggerTagCleared = $false
                $SuccessResult.Data.triggerTagClearError = $tagClearError
            }
        } catch { }
        return New-QCFailureResult -Code 'QC_PREPEND_POST_WRITEBACK_FAILED' `
            -Message 'Post-prepend ProjectWise workflow writeback failed.' -Data @{
            error = $tagClearError
            incomingDocName = $IncomingDocName
            incomingFolderPath = $IncomingFolder
        }
    } finally {
        if ($Job.ContainsKey('document')) { $Job.Remove('document') }
        if ($connected) {
            try { Disconnect-PW | Out-Null } catch { }
        }
    }
}

function _QCP-ResolvePrependProcessTypeFromSourceAttributes {
    <#
    .SYNOPSIS
    Resolves qc_process_type for prepend from live ProjectWise QC_Process_Type first, then sheet_index fallback.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Config,
        [string]$FolderPath,
        [string]$SourceDocumentName
    )

    if ((_QCP-IsNullOrWhiteSpace $FolderPath) -or (_QCP-IsNullOrWhiteSpace $SourceDocumentName)) {
        return @{ processType = $null; resolutionSource = '' }
    }

    if (Get-Command -Name 'Get-PWQcPrependProcessIntentFromSourcePdf' -ErrorAction SilentlyContinue) {
        try {
            $pw = Get-PWQcPrependProcessIntentFromSourcePdf -FolderPath ([string]$FolderPath) `
                -SourceDocumentName ([string]$SourceDocumentName) -Config $Config
            if ($pw -and [bool]$pw.found) {
                $norm = _QCP-NormalizePrependProcessTypeValue -RawProcessType ([string]$pw.qcProcessType)
                if ($norm) {
                    return @{ processType = [string]$norm; resolutionSource = 'projectwise_attributes' }
                }
            }
        } catch { }
    } elseif (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
        try {
            $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath ([string]$FolderPath) `
                -SourceDocumentName ([string]$SourceDocumentName) -Config $Config
            if ($pw -and [bool]$pw.found) {
                $norm = _QCP-NormalizePrependProcessTypeValue -RawProcessType ([string]$pw.qcProcessType)
                if ($norm) {
                    return @{ processType = [string]$norm; resolutionSource = 'projectwise_attributes' }
                }
            }
        } catch { }
    }

    $idxProcess = _QCP-ResolveProcessTypeFromSheetIndex -Config $Config -FolderPath ([string]$FolderPath) `
        -SourceDocumentName ([string]$SourceDocumentName)
    if (-not (_QCP-IsNullOrWhiteSpace $idxProcess)) {
        return @{ processType = [string]$idxProcess; resolutionSource = 'sheet_index' }
    }

    return @{ processType = $null; resolutionSource = '' }
}

function _QCP-TryResolvePrependLaneContext {
    <#
    .SYNOPSIS
    Resolves QC process type and lane PDF naming before prepend execution or workflow writeback.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config
    )

    $folder = _QCP-GetJobMetadataValue -Job $Job -Keys @('folderPath', 'sourceFolder', 'incomingFolderPath')
    $sheetPdfRaw = _QCP-GetJobMetadataValue -Job $Job -Keys @('sourceName', 'incomingDocName', 'sourceDocumentName')
    $sheetPdfName = _QCP-NormalizeJobDocumentFileName -DocumentName ([string]$sheetPdfRaw)
    if (_QCP-IsNullOrWhiteSpace $sheetPdfName) {
        $sheetPdfFallback = if ($Job.sourceName) { [string]$Job.sourceName } else { '' }
        $sheetPdfName = _QCP-NormalizeJobDocumentFileName -DocumentName $sheetPdfFallback
    }
    $laneTriggerRaw = _QCP-GetJobMetadataOnlyValue -Job $Job -Keys @('triggerDocumentName')
    $laneTriggerName = _QCP-NormalizeJobDocumentFileName -DocumentName ([string]$laneTriggerRaw)
    if (_QCP-IsNullOrWhiteSpace $laneTriggerName) {
        $laneTriggerFallback = if ($Job.sourceName) { [string]$Job.sourceName } else { '' }
        $laneTriggerName = _QCP-NormalizeJobDocumentFileName -DocumentName $laneTriggerFallback
    }
    $sourceDocumentGuid = _QCP-GetJobMetadataValue -Job $Job -Keys @('triggerDocumentGuid', 'documentGuid', 'sourceDocumentGuid')

    $processType = $null
    $resolutionSource = ''

    $rawProcess = _QCP-GetJobMetadataValue -Job $Job -Keys @('qcProcessType', 'processType', 'qc_process_type')
    $norm = _QCP-NormalizePrependProcessTypeValue -RawProcessType ([string]$rawProcess)
    if ($norm) {
        $processType = $norm
        $resolutionSource = 'job_metadata'
    }

    if (-not $processType -and (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue)) {
        foreach ($laneName in @($laneTriggerName, $sheetPdfName)) {
            if (_QCP-IsNullOrWhiteSpace $laneName) { continue }
            $laneFromName = Get-PWQcPdfLaneFromDocumentName -DocumentName ([string]$laneName)
            if ($laneFromName) {
                $processType = $laneFromName
                $resolutionSource = 'document_name_lane'
                break
            }
        }
    }

    $attributeLookupName = if (-not (_QCP-IsNullOrWhiteSpace $sheetPdfName)) { [string]$sheetPdfName } else { [string]$laneTriggerName }
    if (-not $processType -and -not (_QCP-IsNullOrWhiteSpace $folder) -and -not (_QCP-IsNullOrWhiteSpace $attributeLookupName)) {
        if (_QCP-IsStemSheetDocumentName ([string]$attributeLookupName)) {
            $attrResolved = _QCP-ResolvePrependProcessTypeFromSourceAttributes -Config $Config `
                -FolderPath ([string]$folder) -SourceDocumentName ([string]$attributeLookupName)
            if ($attrResolved.processType) {
                $processType = [string]$attrResolved.processType
                $resolutionSource = [string]$attrResolved.resolutionSource
            }
        } elseif (-not $processType) {
            $idxProcess = _QCP-ResolveProcessTypeFromSheetIndex -Config $Config -FolderPath ([string]$folder) -SourceDocumentName ([string]$attributeLookupName)
            if (-not (_QCP-IsNullOrWhiteSpace $idxProcess)) {
                $processType = [string]$idxProcess
                $resolutionSource = 'sheet_index'
            }
        }
    }

    if (-not $processType) {
        $prependTrigger = _QCP-ResolvePrependTrigger -Job $Job
        if ($prependTrigger -eq 'initialQcPdf' `
                -and (_QCP-IsStemSheetDocumentName ([string]$laneTriggerName)) `
                -and (_QCP-IsStemSheetDocumentName ([string]$sheetPdfName))) {
            # QC Initiated intake on the stem sheet PDF always creates the production lane PDF.
            $processType = 'production'
            $resolutionSource = 'initial_prepend_default'
        }
    }

    if (-not $processType) {
        return New-QCFailureResult -Code 'QC_PROCESS_TYPE_UNKNOWN' `
            -Message 'QC_PREPEND could not resolve qc_process_type; refusing to default to production.' -Data @{
            jobId = if ($Job.id) { [string]$Job.id } else { '' }
            folderPath = [string]$folder
            triggerDocumentName = [string]$laneTriggerName
            sheetPdfName = [string]$sheetPdfName
            sourceDocumentGuid = [string]$sourceDocumentGuid
            metadataQcProcessType = [string]$rawProcess
        }
    }

    $laneSuffix = ''
    if (Get-Command -Name 'Get-QCProcessTypePdfSuffix' -ErrorAction SilentlyContinue) {
        $laneSuffix = Get-QCProcessTypePdfSuffix -ProcessType $processType -Config $Config
    }
    if (_QCP-IsNullOrWhiteSpace $laneSuffix) {
        return New-QCFailureResult -Code 'QC_LANE_PDF_RESOLUTION_FAILED' `
            -Message 'QC_PREPEND could not resolve lane PDF suffix for process type.' -Data @{
            jobId = if ($Job.id) { [string]$Job.id } else { '' }
            qcProcessType = $processType
            triggerDocumentName = [string]$laneTriggerName
            sheetPdfName = [string]$sheetPdfName
        }
    }

    $stemSourceName = if (-not (_QCP-IsNullOrWhiteSpace $sheetPdfName)) { [string]$sheetPdfName } else { [string]$laneTriggerName }
    $sheetStem = [System.IO.Path]::GetFileNameWithoutExtension([string]$stemSourceName)
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        $resolvedStem = Get-PWSheetStemFromDocumentName -DocumentName ([string]$stemSourceName)
        if (-not (_QCP-IsNullOrWhiteSpace $resolvedStem)) { $sheetStem = [string]$resolvedStem }
    }
    $expectedLanePdfName = $null
    if (Get-Command -Name 'Get-QCLaneQcPdfExpectedName' -ErrorAction SilentlyContinue) {
        $expectedLanePdfName = Get-QCLaneQcPdfExpectedName -SheetBaseName $sheetStem -ProcessType $processType -Config $Config
    }
    if (_QCP-IsNullOrWhiteSpace $expectedLanePdfName) {
        $expectedLanePdfName = ($sheetStem + '-' + $laneSuffix + '.pdf')
    }

    $laneData = @{
        qcProcessType = $processType
        pdfSuffix = [string]$laneSuffix
        expectedLanePdfName = [string]$expectedLanePdfName
        sheetStem = [string]$sheetStem
        triggerDocumentName = [string]$laneTriggerName
        sheetPdfName = [string]$sheetPdfName
        sourceDocumentGuid = [string]$sourceDocumentGuid
        resolutionSource = $resolutionSource
        folderPath = [string]$folder
    }

    $md = _QCP-EnsureJobMetadataHashtable -Job $Job
    $md['qcProcessType'] = $processType
    $md['pdfSuffix'] = [string]$laneSuffix
    $md['expectedLanePdfName'] = [string]$expectedLanePdfName
    if (-not (_QCP-IsNullOrWhiteSpace $sourceDocumentGuid)) { $md['sourceDocumentGuid'] = [string]$sourceDocumentGuid }

    return New-QCSuccessResult -Code 'QC_PREPEND_LANE_RESOLVED' -Message 'QC_PREPEND lane context resolved.' -Data $laneData
}

function _QCP-ReviewStampRequiredForReviewType {
    param(
        [hashtable]$StampSettings,
        [string]$ReviewType,
        [string]$ProcessType = ''
    )

    if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
        $norm = Normalize-QCProcessType -ProcessType $ProcessType -ReviewType $ReviewType
        if ($norm -in @('check', 'review')) { return $true }
        if ($norm -eq 'production') { return $false }
    }
    if (-not $StampSettings -or (_QCP-IsNullOrWhiteSpace $ReviewType)) { return $false }
    foreach ($p in @($StampSettings.profiles)) {
        if ([string]$p.reviewType -eq [string]$ReviewType) { return $true }
    }
    return $false
}

function _QCP-TryApplyReviewStampFromJob {
    param(
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$FolderPath = '',
        [string]$SourceDocumentName = '',
        [string]$OverlayExe = ''
    )


    if (-not (Get-Command -Name 'Invoke-QCReviewStamp' -ErrorAction SilentlyContinue)) {
        return @{ applied = $false; reason = 'QC.ReviewStamp module not loaded' }
    }

    $roles = _QCP-GetReviewStampRoleFieldsFromJob -Job $Job -Config $Config -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName
    $processType = _QCP-ResolveProcessTypeFromJob -Job $Job -Config $Config
    if (_QCP-IsNullOrWhiteSpace $processType) {
        $src = _QCP-ResolvePrependProcessTypeFromSourceAttributes -Config $Config `
            -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName
        if ($src.processType) { $processType = [string]$src.processType }
    }
    if (_QCP-IsNullOrWhiteSpace $processType -and -not (_QCP-IsNullOrWhiteSpace $roles.qcProcessType)) {
        $norm = _QCP-NormalizePrependProcessTypeValue -RawProcessType ([string]$roles.qcProcessType)
        if ($norm) { $processType = [string]$norm }
    }
    if (_QCP-IsNullOrWhiteSpace $processType) {
        return @{
            applied = $false
            skipped = $true
            reason = 'qc_process_type not resolved from job lane or stem sheet PDF (QC_Review_Type is not used for stamps)'
        }
    }
    $roles.qcProcessType = $processType

    if (Get-Command -Name 'Test-QCPrependSkipReviewStamp' -ErrorAction SilentlyContinue) {
        $prependTrigger = _QCP-ResolvePrependTrigger -Job $Job
        if (Test-QCPrependSkipReviewStamp -PrependTrigger $prependTrigger -ProcessType $processType) {
            return @{
                applied = $false
                skipped = $true
                reason = 'Review stamp skipped: QC Finalizing prepend (production lane).'
                qcProcessType = $processType
            }
        }
    }

    if (_QCP-IsNullOrWhiteSpace $OverlayExe) {
        $qc = _QCP-ToHashtable $Config.qcPrepend
        if ($qc -and $qc.overlayExePath) { $OverlayExe = [string]$qc.overlayExePath }
    }

    if (Get-Command -Name 'Resolve-QCStampForProcess' -ErrorAction SilentlyContinue) {
        $stampResolved = Resolve-QCStampForProcess -Config $Config -ProcessType $processType -FolderPath $FolderPath
        if ($stampResolved.IsSuccess -and $stampResolved.stampPath -and (Get-Command -Name 'Invoke-QCReviewStamp' -ErrorAction SilentlyContinue)) {
            $stampCfg = Get-QCReviewStampSettings -Config $Config
            $layout = if (Get-Command -Name 'Get-QCStampProfileLayout' -ErrorAction SilentlyContinue) {
                Get-QCStampProfileLayout -Config $Config -StampProfile ([string]$stampResolved.resolvedStampProfile)
            } else {
                $stampCfg
            }
            $profileKey = if ($stampResolved.profileKey) { [string]$stampResolved.profileKey } else { 'check' }
            $roleValues = @{
                designerEmail = [string]$roles.designerEmail
                reviewerEmail = [string]$roles.reviewerEmail
                checkerEmail = [string]$roles.checkerEmail
            }
            if (Get-Command -Name '_QCRS-ResolveStampRoleValues' -ErrorAction SilentlyContinue) {
                $roleValues = _QCRS-ResolveStampRoleValues -RoleFields $roleValues -ProfileKey $profileKey
            }
            $stampParams = @{
                OverlayExe = if ($OverlayExe) { $OverlayExe } else { [string]$stampCfg.overlayExe }
                PdfPath = $PdfPath
                StampPath = [string]$stampResolved.stampPath
                StampHeightPt = if ($layout) { [double]$layout.stampHeightPt } else { 200 }
                MarginOutsidePt = if ($layout) { [double]$layout.marginOutsidePt } else { 12 }
                PopulateTextFields = if ($layout) { [bool]$layout.populateTextFields } else { $false }
            }
            if ($layout -and $null -ne $layout.stampXPt -and $null -ne $layout.stampYPt) {
                $stampParams['StampXPt'] = [double]$layout.stampXPt
                $stampParams['StampYPt'] = [double]$layout.stampYPt
            }
            $result = Invoke-QCReviewStamp @stampParams
            $result['reviewType'] = $processType
            $result['qcProcessType'] = $processType
            $result['stampProfile'] = $stampResolved.resolvedStampProfile
            $result['stampName'] = $stampResolved.resolvedStampName
            return $result
        }
        if ($processType -in @('check', 'review') -and -not $stampResolved.IsSuccess) {
            return @{ applied = $false; skipped = $false; reason = [string]$stampResolved.Message; qcProcessType = $processType }
        }
        if (-not $stampResolved.IsSuccess) {
            return @{
                applied = $false
                skipped = $true
                reason = [string]$stampResolved.Message
                qcProcessType = $processType
            }
        }
    }

    return @{
        applied = $false
        skipped = $true
        reason = 'Resolve-QCStampForProcess not available'
        qcProcessType = $processType
    }
}

function _QCP-TryApplyPeerReviewStampFromJob {
    param(
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$FolderPath = '',
        [string]$SourceDocumentName = '',
        [string]$OverlayExe = ''
    )

    return _QCP-TryApplyReviewStampFromJob -PdfPath $PdfPath -Job $Job -Config $Config `
        -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName -OverlayExe $OverlayExe
}

function _QCP-EnsureDir([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function _QCP-Sha256Hex([string]$Text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
    } finally {
        $sha.Dispose()
    }
    return -join ($hash | ForEach-Object { $_.ToString('x2') })
}

function _QCP-GetJobMetadataValue([hashtable]$Job, [string[]]$Keys) {
    foreach ($k in $Keys) {
        if ($Job.ContainsKey($k) -and $null -ne $Job[$k]) { return $Job[$k] }
    }
    return _QCP-GetJobMetadataOnlyValue -Job $Job -Keys $Keys
}

function _QCP-GetJobMetadataOnlyValue([hashtable]$Job, [string[]]$Keys) {
    try {
        if ($Job.ContainsKey('metadata') -and $Job.metadata) {
            $md = _QCP-ToHashtable $Job.metadata
            if ($md) {
                foreach ($k in $Keys) {
                    if ($md.ContainsKey($k) -and $null -ne $md[$k]) { return $md[$k] }
                }
                if ($md.ContainsKey('candidate') -and $md.candidate) {
                    $cand = _QCP-ToHashtable $md.candidate
                    if ($cand) {
                        foreach ($k in $Keys) {
                            if ($cand.ContainsKey($k) -and $null -ne $cand[$k]) { return $cand[$k] }
                        }
                    }
                }
            }
        }
    } catch { }
    return $null
}

function _QCP-NormalizeJobDocumentFileName([string]$DocumentName) {
    if (_QCP-IsNullOrWhiteSpace $DocumentName) { return '' }
    $name = [string]$DocumentName
    if ($name -match '\\') { $name = [System.IO.Path]::GetFileName($name) }
    return $name
}

function _QCP-ResolvePrependTrigger([hashtable]$Job) {
    $t = _QCP-GetJobMetadataValue -Job $Job -Keys @('prependTrigger','qcPrependTrigger','workflowPrependTrigger')
    if (-not (_QCP-IsNullOrWhiteSpace $t)) { return ([string]$t).Trim() }

    $final = _QCP-GetJobMetadataValue -Job $Job -Keys @('finalQcComplete','qcFinalizing','finalPrepend')
    if (-not (_QCP-IsNullOrWhiteSpace $final)) {
        $fv = ([string]$final).Trim().ToLowerInvariant()
        if ($fv -in @('true','1','yes','y','finalqccomplete')) { return 'finalQcComplete' }
    }

    try {
        if ($Job -and $Job.ContainsKey('metadata') -and $Job.metadata) {
            $md = _QCP-ToHashtable $Job.metadata
            if ($md -and $md.ContainsKey('candidate') -and $md.candidate) {
                $cand = _QCP-ToHashtable $md.candidate
                if ($cand) {
                    foreach ($dk in @('description','documentDescription','desc')) {
                        if ($cand.ContainsKey($dk) -and $cand[$dk] -and ([string]$cand[$dk]).IndexOf('QC_Archivist', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                            return 'initialQcPdf'
                        }
                    }
                }
            }
        }
    } catch { }

    return $null
}

function _QCP-IsFinalQcPrependJob([hashtable]$Job) {
    $trigger = _QCP-ResolvePrependTrigger -Job $Job
    return ($trigger -eq 'finalQcComplete')
}

function _QCP-NewWorkflowContext([hashtable]$Job, [hashtable]$Config, [string]$SourcePath, [string]$OutputPath, [string]$HistoryPath, [string]$ResultStatus, [string]$ErrorMessage, [object]$Document) {
    $now = Get-QCTimestamp
    $designerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('designerEmail','qcDesignerEmail')
    $reviewerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewerEmail','qcReviewerEmail','reviewer','qcReviewer')
    $checkerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('checkerEmail','qcCheckerEmail')
    $reviewType = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewType','qcReviewType')
    $pwFolder = _QCP-GetJobMetadataValue -Job $Job -Keys @('folderPath','sourceFolder','incomingFolderPath')
    $pwDoc = _QCP-GetJobMetadataValue -Job $Job -Keys @('sourceName','incomingDocName','sourceDocumentName')
    if (_QCP-IsNullOrWhiteSpace $pwDoc) { $pwDoc = _QCP-GetJobMetadataValue -Job $Job -Keys @('sourcePath') }
    if (-not (_QCP-IsNullOrWhiteSpace $pwDoc) -and $pwDoc -match '\\') { $pwDoc = [System.IO.Path]::GetFileName([string]$pwDoc) }
    if (-not (_QCP-IsNullOrWhiteSpace $pwFolder) -and -not (_QCP-IsNullOrWhiteSpace $pwDoc)) {
        $pwRoles = _QCP-GetReviewStampRoleFieldsFromJob -Job $Job -Config $Config -FolderPath ([string]$pwFolder) -SourceDocumentName ([string]$pwDoc)
        if (-not (_QCP-IsNullOrWhiteSpace $pwRoles.designerEmail)) { $designerEmail = [string]$pwRoles.designerEmail }
        if (-not (_QCP-IsNullOrWhiteSpace $pwRoles.reviewerEmail)) { $reviewerEmail = [string]$pwRoles.reviewerEmail }
        if (-not (_QCP-IsNullOrWhiteSpace $pwRoles.checkerEmail)) { $checkerEmail = [string]$pwRoles.checkerEmail }
        if ((_QCP-IsNullOrWhiteSpace $reviewType) -and -not (_QCP-IsNullOrWhiteSpace $pwRoles.qcReviewType)) {
            $reviewType = [string]$pwRoles.qcReviewType
        }
    }
    if (_QCP-IsNullOrWhiteSpace $reviewType) {
        $qcRtEnabled = $true
        if (-not (_QCP-IsNullOrWhiteSpace $pwFolder) -and (Get-Command -Name 'Test-PWQcReviewTypeAttributesEnabled' -ErrorAction SilentlyContinue)) {
            $qcRtEnabled = Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath ([string]$pwFolder)
        }
        if ($qcRtEnabled -and -not (_QCP-IsNullOrWhiteSpace $pwFolder) -and -not (_QCP-IsNullOrWhiteSpace $pwDoc)) {
            if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
                $pwRoles = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath ([string]$pwFolder) -SourceDocumentName ([string]$pwDoc) -Config $Config
                if ($pwRoles.found -and -not (_QCP-IsNullOrWhiteSpace $pwRoles.qcReviewType)) {
                    $reviewType = [string]$pwRoles.qcReviewType
                }
            }
        }
    }

    $lastActionBy = _QCP-GetJobMetadataValue -Job $Job -Keys @('lastActionBy','userName','triggeredBy')
    $cycleId = _QCP-GetJobMetadataValue -Job $Job -Keys @('cycleId','qcCycleId')
    $cycleNumber = _QCP-GetJobMetadataValue -Job $Job -Keys @('cycleNumber','qcCycleNumber')
    if (_QCP-IsNullOrWhiteSpace $cycleId -and $Job.ContainsKey('id')) { $cycleId = [string]$Job.id }

    $lifecycleState = $null
    if ($Document) {
        try {
            if ($Document.PSObject.Properties['StateName'] -and $Document.StateName) { $lifecycleState = [string]$Document.StateName }
            elseif ($Document.PSObject.Properties['WorkflowState'] -and $Document.WorkflowState) { $lifecycleState = [string]$Document.WorkflowState }
        } catch { }
    }
    if (_QCP-IsNullOrWhiteSpace $lifecycleState) {
        $lifecycleState = _QCP-GetJobMetadataValue -Job $Job -Keys @('pwStateName','stateName','workflowState','currentState')
    }

    $settings = $null
    $assignedTo = $null
    $statusValue = $lifecycleState
    try {
        $settings = Get-QCWorkflowSettings -Config $Config
        if (_QCP-IsNullOrWhiteSpace $reviewType) {
            $pwFolderForRt = _QCP-GetJobMetadataValue -Job $Job -Keys @('folderPath','sourceFolder','incomingFolderPath')
            $applyDefaultRt = $true
            if (-not (_QCP-IsNullOrWhiteSpace $pwFolderForRt) -and (Get-Command -Name 'Test-PWQcReviewTypeAttributesEnabled' -ErrorAction SilentlyContinue)) {
                $applyDefaultRt = Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath ([string]$pwFolderForRt)
            }
            if ($applyDefaultRt) { $reviewType = [string]$settings.defaultReviewType }
        }
        $assignedTo = Resolve-QCWorkflowAssignee -Settings $settings -StateName $lifecycleState -ReviewType $reviewType `
            -DesignerEmail $designerEmail -ReviewerEmail $reviewerEmail -CheckerEmail $checkerEmail
        if (_QCP-IsNullOrWhiteSpace $assignedTo) {
            $assignedTo = _QCP-GetJobMetadataValue -Job $Job -Keys @('assignedTo','qcAssignedTo')
        }
        if (_QCP-IsNullOrWhiteSpace $statusValue -and -not (_QCP-IsNullOrWhiteSpace $lifecycleState)) { $statusValue = $lifecycleState }
        elseif (_QCP-IsNullOrWhiteSpace $statusValue) { $statusValue = $ResultStatus }
    } catch {
        if (_QCP-IsNullOrWhiteSpace $statusValue) { $statusValue = $ResultStatus }
        $assignedTo = _QCP-GetJobMetadataValue -Job $Job -Keys @('assignedTo','qcAssignedTo')
    }

    $docPath = if (-not (_QCP-IsNullOrWhiteSpace $SourcePath)) { $SourcePath } else { [string](_QCP-GetJobMetadataValue -Job $Job -Keys @('sourceDocumentPath','sourcePath')) }
    return @{
        jobId = [string]$Job.id
        jobType = [string]$Job.type
        resultStatus = $ResultStatus
        lifecycleState = $lifecycleState
        documentPath = $docPath
        sourcePath = $SourcePath
        outputPath = $OutputPath
        historyPath = $HistoryPath
        timestamp = $now
        attributes = @{
            qcActive = $true
            reviewType = $reviewType
            cycleId = $cycleId
            cycleNumber = $cycleNumber
            designerEmail = $designerEmail
            reviewerEmail = $reviewerEmail
            checkerEmail = $checkerEmail
            assignedTo = $assignedTo
            lastActionBy = $lastActionBy
            lastActionDate = $now
            status = $statusValue
            historyPdfPath = $HistoryPath
            latestOverlayPdfPath = $OutputPath
            sourceDocumentPath = $docPath
            automationLastRun = $now
            automationResult = $ResultStatus
            automationError = $ErrorMessage
        }
    }
}

function _QCP-ResolvePwDocumentForJob([hashtable]$Job, [hashtable]$Config) {
    if (-not $Job) { return $null }
    if ($Job.ContainsKey('document') -and $Job.document) { return $Job.document }

    if (Get-Command -Name 'Get-QCCommentSyncPwDocument' -ErrorAction SilentlyContinue) {
        try {
            $res = Get-QCCommentSyncPwDocument -Job $Job -Config $Config
            if ($res.IsSuccess -and $res.Data -and $res.Data.document) {
                $Job['document'] = $res.Data.document
                return $res.Data.document
            }
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Warning' -Code 'QC_PREPEND_PW_DOC_RESOLVE_FAILED' `
                    -Message 'Could not resolve ProjectWise document for workflow writeback.' -Data @{
                    jobId = if ($Job.id) { [string]$Job.id } else { $null }
                    code = [string]$res.Code
                    message = [string]$res.Message
                    sourceFolder = if ($Job.sourceFolder) { [string]$Job.sourceFolder } else { $null }
                    sourceName = if ($Job.sourceName) { [string]$Job.sourceName } else { $null }
                } | Out-Null
            }
        } catch {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Warning' -Code 'QC_PREPEND_PW_DOC_RESOLVE_FAILED' `
                    -Message 'ProjectWise document resolution threw.' -Data @{
                    jobId = if ($Job.id) { [string]$Job.id } else { $null }
                    errorMessage = [string]$_.Exception.Message
                } | Out-Null
            }
        }
    }
    return $null
}

function _QCP-LogPrependWorkflowWriteback([hashtable]$Job, [hashtable]$Config, [object]$Writeback, [string]$PrependTrigger) {
    if (-not (Get-Command -Name Write-QCJsonLog -ErrorAction SilentlyContinue)) { return }

    $jobId = $null
    if ($Job) {
        if ($Job.ContainsKey('id') -and $Job.id) { $jobId = [string]$Job.id }
        elseif ($Job.ContainsKey('jobId') -and $Job.jobId) { $jobId = [string]$Job.jobId }
    }

    $settings = $null
    try { $settings = Get-QCWorkflowSettings -Config $Config } catch { }

    $targetState = $null
    $previousState = $null
    $stateCode = $null
    $planned = $false
    $changed = $false
    $dryRun = $false
    $warnings = @()

    $transitionDetail = $null
    if ($Writeback -and $Writeback.Data) {
        $wbData = _QCP-ToHashtable $Writeback.Data
        if ($wbData) {
            if ($wbData.ContainsKey('dryRun')) { try { $dryRun = [bool]$wbData.dryRun } catch { } }
            if ($wbData.warnings) { $warnings = @($wbData.warnings) }
            if ($wbData.actions) {
                foreach ($a in @($wbData.actions)) {
                    if (-not $a) { continue }
                    $ac = [string]$a.Code
                    if ($ac -notmatch '^QC_WORKFLOW_STATE') { continue }
                    $stateCode = $ac
                    $ad = _QCP-ToHashtable $a.Data
                    if ($ad) {
                        if ($ad.stateName) { $targetState = [string]$ad.stateName }
                        if ($ad.currentStateName) { $previousState = [string]$ad.currentStateName }
                        if ($ad.ContainsKey('planned')) { try { $planned = [bool]$ad.planned } catch { } }
                        if ($ad.ContainsKey('changed')) { try { $changed = [bool]$ad.changed } catch { } }
                        if ($ad.transition) {
                            $tr = _QCP-ToHashtable $ad.transition
                            if ($tr -and $tr.Data) { $transitionDetail = _QCP-ToHashtable $tr.Data }
                        }
                    }
                    break
                }
            }
        }
    }

    if ((_QCP-IsNullOrWhiteSpace $targetState) -and $settings) {
        try {
            $resolveCtx = @{ resultStatus = 'Succeeded' }
            if (-not (_QCP-IsNullOrWhiteSpace $PrependTrigger)) { $resolveCtx['prependTrigger'] = $PrependTrigger }
            $targetState = Resolve-QCWorkflowStateAfterPrepend -Settings $settings -Context $resolveCtx
        } catch { }
    }

    $level = 'Information'
    if ($Writeback -and -not $Writeback.IsSuccess) { $level = 'Warning' }
    elseif ($stateCode -match 'FAILED|INVALID|UNAVAILABLE|MISSING') { $level = 'Warning' }

    $msg = if ($changed) {
        "QC_PREPEND set workflow state to '$targetState'."
    } elseif ($planned -or $dryRun) {
        "QC_PREPEND planned workflow state '$targetState' (dry-run)."
    } else {
        'QC_PREPEND workflow state writeback finished (no ProjectWise state change).'
    }

    Write-QCJsonLog -Level $level -Code 'QC_PREPEND_WORKFLOW_STATE' -Message $msg -Data @{
        jobId             = $jobId
        prependTrigger    = $PrependTrigger
        targetState       = $targetState
        previousState     = $previousState
        qcReviewType      = $(if (Get-Command -Name 'Resolve-QCWorkflowEventQcReviewType' -ErrorAction SilentlyContinue) {
            $docGuid = _QCP-GetJobMetadataValue -Job $Job -Keys @('triggerDocumentGuid','documentGuid')
            $folder = _QCP-GetJobMetadataValue -Job $Job -Keys @('folderPath','sourceFolder','incomingFolderPath')
            $docName = _QCP-GetJobMetadataValue -Job $Job -Keys @('sourceName','incomingDocName','sourceDocumentName')
            Resolve-QCWorkflowEventQcReviewType -Config $Config -DocumentGuid ([string]$docGuid) -FolderPath ([string]$folder) `
                -DocumentName ([string]$docName) -Context @{ job = $Job }
        } else { '' })
        stateActionCode   = $stateCode
        planned           = $planned
        changed           = $changed
        dryRun            = $dryRun
        writebackCode     = if ($Writeback) { [string]$Writeback.Code } else { $null }
        writebackSuccess  = if ($Writeback) { [bool]$Writeback.IsSuccess } else { $false }
        autoSetState      = if ($settings) { [bool]$settings.autoSetState } else { $null }
        mode              = if ($settings) { [string]$settings.mode } else { $null }
        workflowName      = if ($settings) { [string]$settings.workflowName } else { $null }
        warnings          = @($warnings)
        targetStateExists = if ($transitionDetail -and $transitionDetail.ContainsKey('targetStateExists')) { $transitionDetail.targetStateExists } else { $null }
        transitionValid   = if ($transitionDetail -and $transitionDetail.ContainsKey('transitionValid')) { $transitionDetail.transitionValid } else { $null }
        workflowStates    = if ($transitionDetail -and $transitionDetail.states) { @($transitionDetail.states) } else { @() }
    } | Out-Null
}

function _QCP-AppendWorkflowWriteback([object]$Result, [hashtable]$Job, [hashtable]$Config, [string]$SourcePath, [string]$OutputPath, [string]$HistoryPath) {
    if ($null -eq $Result -or -not $Result.IsSuccess) { return $Result }
    $document = _QCP-ResolvePwDocumentForJob -Job $Job -Config $Config
    $ctx = _QCP-NewWorkflowContext -Job $Job -Config $Config -SourcePath $SourcePath -OutputPath $OutputPath -HistoryPath $HistoryPath -ResultStatus 'Succeeded' -ErrorMessage $null -Document $document
    $ctx['config'] = $Config
    $ctx['job'] = $Job
    if ($document) { $ctx['document'] = $document }
    $prependTrigger = _QCP-ResolvePrependTrigger -Job $Job
    if (-not (_QCP-IsNullOrWhiteSpace $prependTrigger)) { $ctx['prependTrigger'] = $prependTrigger }
    $laneRes = _QCP-TryResolvePrependLaneContext -Job $Job -Config $Config
    if (-not $laneRes.IsSuccess) { return $laneRes }
    $laneCtx = [hashtable]$laneRes.Data
    $processType = [string]$laneCtx.qcProcessType
    $ctx['qcProcessType'] = $processType
    $ctx['expectedLanePdfName'] = [string]$laneCtx.expectedLanePdfName
    $ctx['pdfSuffix'] = [string]$laneCtx.pdfSuffix
    if (-not $Job.ContainsKey('metadata') -or -not ($Job.metadata -is [hashtable])) {
        $Job['metadata'] = @{}
    }
    $Job.metadata['qcProcessType'] = $processType
    $Job.metadata['expectedLanePdfName'] = [string]$laneCtx.expectedLanePdfName
    $Job.metadata['pdfSuffix'] = [string]$laneCtx.pdfSuffix
    $folderPath = _QCP-GetJobMetadataValue -Job $Job -Keys @('folderPath', 'sourceFolder', 'incomingFolderPath')
    $sourceDocumentGuid = _QCP-GetJobMetadataValue -Job $Job -Keys @('triggerDocumentGuid', 'documentGuid', 'sourceDocumentGuid')
    $lanePdfName = [string]$laneCtx.expectedLanePdfName
    if (-not (_QCP-IsNullOrWhiteSpace $lanePdfName) -and -not (_QCP-IsNullOrWhiteSpace $folderPath) `
            -and (Get-Command -Name 'Sync-QCLaneQcPdfGuidFromProjectWise' -ErrorAction SilentlyContinue)) {
        try {
            $reg = Sync-QCLaneQcPdfGuidFromProjectWise -Config $Config -FolderPath $folderPath -QcPdfName $lanePdfName `
                -QcProcessType $processType -SourceDocumentGuid $sourceDocumentGuid -Required
            if ($reg.IsSuccess -and $reg.Data -and $reg.Data.documentGuid) {
                $liveLaneGuid = [string]$reg.Data.documentGuid
                $ctx['notificationLaneDocumentGuid'] = $liveLaneGuid
                $ctx['laneQcPdfDocumentGuid'] = $liveLaneGuid
                $Job.metadata['notificationLaneDocumentGuid'] = $liveLaneGuid
                $Job.metadata['laneQcPdfDocumentGuid'] = $liveLaneGuid
            }
        } catch {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Warning' -Code 'QC_LANE_PDF_GUID_REGISTER_FAILED' `
                    -Message 'Failed to register live lane QC PDF GUID before workflow writeback.' -Data @{
                    folderPath = $folderPath
                    qcPdfName = $lanePdfName
                    qcProcessType = $processType
                    error = [string]$_.Exception.Message
                } | Out-Null
            }
        }
    }
    if (Get-Command -Name 'Test-QCProcessTypeSyncsWithSiblingSheets' -ErrorAction SilentlyContinue) {
        $ctx['skipSiblingStateSync'] = -not (Test-QCProcessTypeSyncsWithSiblingSheets -ProcessType $processType -Config $Config)
    } else {
        $ctx['skipSiblingStateSync'] = $true
    }

    $previousState = _QCP-GetJobMetadataValue -Job $Job -Keys @('pwStateName','stateName','workflowState','currentState')
    if (-not (_QCP-IsNullOrWhiteSpace $previousState)) {
        $ctx['previousState'] = [string]$previousState
    } elseif (-not (_QCP-IsNullOrWhiteSpace $prependTrigger)) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            if ($prependTrigger -eq 'finalQcComplete') {
                $ctx['previousState'] = Get-QCWorkflowStateName -Settings $wf -StateKey 'qcFinalizing'
            } elseif ($prependTrigger -eq 'initialQcPdf') {
                $ctx['previousState'] = Get-QCWorkflowStateName -Settings $wf -StateKey 'qcInitiated'
            }
        } catch { }
    }
    if ($ctx.ContainsKey('previousState') -and $ctx.previousState -and -not $ctx.ContainsKey('lifecycleState')) {
        $ctx['lifecycleState'] = [string]$ctx.previousState
    }

    $triggerGuid = _QCP-GetJobMetadataValue -Job $Job -Keys @('triggerDocumentGuid','documentGuid')
    if (-not (_QCP-IsNullOrWhiteSpace $triggerGuid)) { $ctx['documentGuid'] = [string]$triggerGuid }
    $triggerName = _QCP-GetJobMetadataValue -Job $Job -Keys @('triggerDocumentName')
    if (-not (_QCP-IsNullOrWhiteSpace $triggerName)) { $ctx['documentName'] = [string]$triggerName }
    $stKey = _QCP-GetJobMetadataValue -Job $Job -Keys @('stateTransitionKey')
    if (-not (_QCP-IsNullOrWhiteSpace $stKey)) { $ctx['stateTransitionKey'] = [string]$stKey }

    $writeback = Invoke-QCWorkflowWriteback -Config $Config -Context $ctx
    _QCP-LogPrependWorkflowWriteback -Job $Job -Config $Config -Writeback $writeback -PrependTrigger $prependTrigger
    if ($Result.IsSuccess -and (Get-Command -Name 'Test-QCProcessTypeResetsAfterPrepend' -ErrorAction SilentlyContinue) `
        -and (Test-QCProcessTypeResetsAfterPrepend -ProcessType $processType -Config $Config)) {
        _QCP-TryResetProcessTypeAfterPrepend -Job $Job -Config $Config -ProcessType $processType | Out-Null
    } elseif ($Result.IsSuccess -and (Get-Command -Name 'Test-QCResetProcessTypeAfterLanePrepend' -ErrorAction SilentlyContinue) `
        -and -not (Test-QCResetProcessTypeAfterLanePrepend -Config $Config)) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PROCESS_TYPE_RESET_SKIPPED' `
                -Message 'Process type reset after lane prepend is disabled.' -Data @{
                qcProcessType = $processType
                prependTrigger = $prependTrigger
            } | Out-Null
        }
    }
    $strict = $false
    try {
        $settings = Get-QCWorkflowSettings -Config $Config
        $strict = [bool]$settings.strictMode
    } catch { }
    $data = @{}
    if ($Result.Data) {
        $rd = _QCP-ToHashtable $Result.Data
        if ($rd) { foreach ($k in $rd.Keys) { $data[$k] = $rd[$k] } }
    }
    $data['workflowWriteback'] = $writeback
    $writebackUnverified = ($writeback.Code -eq 'QC_WORKFLOW_STATE_WRITE_UNVERIFIED')
    if ($writebackUnverified) {
        $data['laneStateVerified'] = $false
    }
    if ((-not $writeback.IsSuccess -or $writebackUnverified) -and $strict) {
        return New-QCFailureResult -Code 'QC_PREPEND_WORKFLOW_WRITEBACK_FAILED' -Message 'QC_PREPEND succeeded but strict QC workflow writeback failed.' -Data $data
    }

    return New-QCSuccessResult -Code $Result.Code -Message $Result.Message -Data $data
}

function _QCP-TryResetProcessTypeAfterPrepend {
    param(
        [hashtable]$Job,
        [hashtable]$Config,
        [string]$ProcessType
    )
    if (-not (Get-Command -Name 'Test-QCResetProcessTypeAfterLanePrepend' -ErrorAction SilentlyContinue)) { return @{ reset = $false } }
    if (-not (Test-QCResetProcessTypeAfterLanePrepend -Config $Config)) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PROCESS_TYPE_RESET_SKIPPED' -Message 'Process type reset after lane prepend is disabled.' -Data @{
                qcProcessType = $ProcessType
            } | Out-Null
        }
        return @{ reset = $false; skipped = $true; reason = 'reset_disabled' }
    }
    if (-not (Get-Command -Name 'Test-QCProcessTypeResetsAfterPrepend' -ErrorAction SilentlyContinue)) { return @{ reset = $false } }
    if (-not (Test-QCProcessTypeResetsAfterPrepend -ProcessType $ProcessType -Config $Config)) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PROCESS_TYPE_RESET_SKIPPED' -Message 'Process type reset not configured for lane.' -Data @{
                qcProcessType = $ProcessType
            } | Out-Null
        }
        return @{ reset = $false; skipped = $true }
    }

    $folder = _QCP-GetJobMetadataValue -Job $Job -Keys @('folderPath', 'sourceFolder', 'incomingFolderPath')
    $docName = _QCP-GetJobMetadataValue -Job $Job -Keys @('sourceName', 'incomingDocName', 'sourceDocumentName')
    $docGuid = _QCP-GetJobMetadataValue -Job $Job -Keys @('triggerDocumentGuid', 'documentGuid')
    if (Get-Command -Name '_PWD-SyncReferenceSheetProcessTypeAttributes' -ErrorAction SilentlyContinue) {
        try {
            _PWD-SyncReferenceSheetProcessTypeAttributes -Config $Config -DocumentGuid ([string]$docGuid) `
                -DocumentName ([string]$docName) -FolderPath ([string]$folder) -CanonicalProcessType 'production' -ControlDocumentOnly | Out-Null
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_PROCESS_RESET' -Message 'Reset stem/DGN qc_process_type to Production after Initiate Origination prepend.' -Data @{
                    fromProcessType = $ProcessType
                    toProcessType = 'production'
                    folderPath = $folder
                    documentName = $docName
                } | Out-Null
            }
            return @{ reset = $true; toProcessType = 'production' }
        } catch {
            return @{ reset = $false; error = [string]$_.Exception.Message }
        }
    }
    return @{ reset = $false }
}

function _QCP-ResolveSourcePdf([hashtable]$Job) {
    if ($Job.ContainsKey('source') -and $Job.source -is [hashtable] -and $Job.source.ContainsKey('localPath') -and $Job.source.localPath) {
        return [string]$Job.source.localPath
    }
    if ($Job.ContainsKey('sourcePath') -and $Job.sourcePath) {
        return [string]$Job.sourcePath
    }
    return $null
}

function _QCP-GetQpdfExePath([hashtable]$Config) {
    if ($Config.ContainsKey('qcPrepend') -and $Config.qcPrepend) {
        $qc = _QCP-ToHashtable $Config.qcPrepend
        if ($qc -and $qc.ContainsKey('qpdfExePath') -and $qc.qpdfExePath) {
            $p = [string]$qc.qpdfExePath
            try {
                if (-not [System.IO.Path]::IsPathRooted($p)) {
                    return (Join-Path (_QCP-GetRepoRoot) $p)
                }
            } catch { }
            return $p
        }
    }
    $default = Join-Path (_QCP-GetRepoRoot) 'tools\qpdf\bin\qpdf.exe'
    return $default
}

function _QCP-ResolveHistoryPath([hashtable]$Job, [hashtable]$Config) {
    $qc = $null
    if ($Config.ContainsKey('qcPrepend') -and $Config.qcPrepend) { $qc = _QCP-ToHashtable $Config.qcPrepend }
    $historyRoot = if ($qc -and $qc.historyRoot) { [string]$qc.historyRoot } else { $null }
    if (_QCP-IsNullOrWhiteSpace $historyRoot) { return New-QCFailureResult -Code 'QC_PREPEND_CONFIG_MISSING_HISTORY_ROOT' -Message 'qcPrepend.historyRoot is required.' -Data @{} }

    $md = $null
    try {
        if ($Job.ContainsKey('metadata') -and $Job.metadata) { $md = _QCP-ToHashtable $Job.metadata }
    } catch { }
    if ($md -and $md.ContainsKey('historyPdfPath') -and $md.historyPdfPath) {
        return New-QCSuccessResult -Code 'QC_PREPEND_HISTORY_PATH' -Message 'History path resolved from job metadata.' -Data @{ historyPdf = [string]$md.historyPdfPath }
    }

    $laneRes = _QCP-TryResolvePrependLaneContext -Job $Job -Config $Config
    if (-not $laneRes.IsSuccess) { return $laneRes }
    $laneCtx = [hashtable]$laneRes.Data
    $historyFileName = [string]$laneCtx.expectedLanePdfName

    $sourceFolder = if ($Job.ContainsKey('sourceFolder') -and $Job.sourceFolder) { [string]$Job.sourceFolder } else { '' }
    $folderKey = if (_QCP-IsNullOrWhiteSpace $sourceFolder) { 'root' } else { (_QCP-Sha256Hex -Text $sourceFolder).Substring(0, 8) }

    $destDir = Join-Path $historyRoot $folderKey
    $historyPdf = Join-Path $destDir $historyFileName
    return New-QCSuccessResult -Code 'QC_PREPEND_HISTORY_PATH' -Message 'History path resolved from lane context.' -Data @{
        historyPdf = $historyPdf
        historyDir = $destDir
        folderKey = $folderKey
        qcProcessType = [string]$laneCtx.qcProcessType
        pdfSuffix = [string]$laneCtx.pdfSuffix
        expectedLanePdfName = $historyFileName
    }
}

function _QCP-RunQpdfPrepend([string]$QpdfExe, [string]$SourcePdf, [string]$HistoryPdf, [string]$OutPdf) {
    # Prepend by concatenating pages: source + history
    $args = @('--empty', '--pages', $SourcePdf, $HistoryPdf, '--', $OutPdf)
    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    try {
        $p = Start-Process -FilePath $QpdfExe -ArgumentList $args -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $stdout = Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue
        $stderr = Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue
        return @{ exitCode = $p.ExitCode; stdout = $stdout; stderr = $stderr; args = $args }
    } finally {
        Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function _QCP-RunOverlayExe([string]$ExePath, [string[]]$OverlayArgs, [string]$WorkingDirectory) {
    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    try {
        # Start-Process rejects null/empty items in -ArgumentList (and some callers can accidentally pass them).
        $cleanArgs = @()
        foreach ($a in @($OverlayArgs)) {
            if ($null -eq $a) { continue }
            $s = [string]$a
            if ($s.Trim().Length -eq 0) { continue }
            $cleanArgs += $s
        }

        # Windows PowerShell can be picky about string[] ArgumentList. Use a single arg line string.
        $argLine = (($cleanArgs | ForEach-Object {
            $t = [string]$_
            if ($t -match '[\\s\"]') { return ('"' + ($t -replace '"', '\\"') + '"') }
            return $t
        }) -join ' ')

        if ($null -ne $WorkingDirectory -and ([string]$WorkingDirectory).Trim().Length -gt 0) {
            $p = Start-Process -FilePath $ExePath -ArgumentList $argLine -WorkingDirectory $WorkingDirectory -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        } else {
            $p = Start-Process -FilePath $ExePath -ArgumentList $argLine -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        }
        $stdout = Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue
        $stderr = Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue
        return @{ exitCode = $p.ExitCode; stdout = $stdout; stderr = $stderr; args = $cleanArgs; argLine = $argLine }
    } finally {
        Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function _QCP-TokenizeArgs([string]$CommandLine) {
    # Minimal quoted-string tokenizer.
    $args = @()
    $cur = ''
    $inQuote = $false
    for ($i = 0; $i -lt $CommandLine.Length; $i++) {
        $ch = $CommandLine[$i]
        if ($ch -eq '"') { $inQuote = -not $inQuote; continue }
        if (-not $inQuote -and [char]::IsWhiteSpace($ch)) {
            if ($cur.Length -gt 0) { $args += $cur; $cur = '' }
            continue
        }
        $cur += $ch
    }
    if ($cur.Length -gt 0) { $args += $cur }
    return $args
}

function Test-QCJobReady {
    <#
    .SYNOPSIS
    Validates job readiness for processing.
    .DESCRIPTION
    Confirms required job fields, state, and preconditions before dispatch.
    .PARAMETER Job
    Job payload.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    if ((-not $Job.ContainsKey('sourcePath') -or (_QCP-IsNullOrWhiteSpace $Job.sourcePath)) `
        -and $Job.ContainsKey('sourceFolder') -and -not (_QCP-IsNullOrWhiteSpace $Job.sourceFolder) `
        -and $Job.ContainsKey('sourceName') -and -not (_QCP-IsNullOrWhiteSpace $Job.sourceName)) {
        $Job['sourcePath'] = Join-Path ([string]$Job.sourceFolder) ([string]$Job.sourceName)
    }

    $missing = @()
    foreach ($k in @('id', 'type', 'sourcePath', 'dedupeKey')) {
        if (-not $Job.ContainsKey($k) -or (_QCP-IsNullOrWhiteSpace $Job[$k])) { $missing += $k }
    }
    if ($missing.Count -gt 0) {
        return New-QCFailureResult -Code 'PROCESSOR_JOB_NOT_READY' -Message 'Job is missing required fields for processing.' -Data @{ missing = $missing; jobId = [string]$Job.id }
    }

    return New-QCSuccessResult -Code 'PROCESSOR_JOB_READY' -Message 'Job is ready for processing.' -Data @{ jobId = [string]$Job.id }
}

function Invoke-QCProcessorByType {
    <#
    .SYNOPSIS
    Dispatches a job to the mapped processor by job type.
    .DESCRIPTION
    Resolves processor mapping and invokes the appropriate processor entry point.
    .PARAMETER Job
    Job payload.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: may trigger downstream local processing workflows.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $ready = Test-QCJobReady -Job $Job -Config $Config
    if (-not $ready.IsSuccess) { return $ready }

    $jobType = [string]$Job.type
    if (_QCP-IsNullOrWhiteSpace $jobType) {
        return New-QCFailureResult -Code 'PROCESSOR_MISSING_JOB_TYPE' -Message 'Job.type is required for dispatch.' -Data @{ jobId = [string]$Job.id }
    }

    $map = $null
    if ($Config.ContainsKey('processors') -and $Config.processors -and $Config.processors.ContainsKey('processorMap')) {
        $map = $Config.processors.processorMap
    } elseif ($Config.ContainsKey('processorMap')) {
        $map = $Config.processorMap
    }

    $handlerName = $null
    if ($map -is [hashtable] -and $map.ContainsKey($jobType)) {
        $handlerName = [string]$map[$jobType]
    }
    if (_QCP-IsNullOrWhiteSpace $handlerName) {
        # Default mapping for initial rollout.
        if ($jobType -eq 'QC_PREPEND') { $handlerName = 'Invoke-QCPrependProcessor' }
        elseif ($jobType -eq 'STATUS_SET_GEN') { $handlerName = 'Invoke-StatusSetProcessor' }
        elseif ($jobType -eq 'QC_REPORTING_SCAN') { $handlerName = 'Invoke-QCReportingScanProcessor' }
        elseif ($jobType -eq 'QC_COMMENT_STATUS_SYNC') { $handlerName = 'Invoke-QCCommentStatusSyncProcessor' }
        elseif ($jobType -eq 'QC_RENDITION') { $handlerName = 'Invoke-QCRenditionProcessor' }
        elseif ($jobType -eq 'QC_NOTIFICATION') { $handlerName = 'Invoke-QCNotificationProcessor' }
    }

    if (_QCP-IsNullOrWhiteSpace $handlerName) {
        return New-QCFailureResult -Code 'PROCESSOR_NO_HANDLER' -Message "No processor handler mapped for job type: $jobType" -Data @{ jobId = [string]$Job.id; jobType = $jobType }
    }

    $cmd = Get-Command -Name $handlerName -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return New-QCFailureResult -Code 'PROCESSOR_HANDLER_NOT_FOUND' -Message "Mapped processor handler not found: $handlerName" -Data @{ jobId = [string]$Job.id; jobType = $jobType; handler = $handlerName }
    }

    try {
        $r = & $handlerName -Job $Job -Config $Config
        # Handlers must return a single QCResult. In Windows PowerShell, any stray pipeline output
        # turns the return into an object[]; recover by selecting the last QCResult-like item.
        if ($r -is [object[]]) {
            $qcLike = @($r | Where-Object { $_ -ne $null -and $_.PSObject -and ($_.PSObject.Properties.Name -contains 'IsSuccess') })
            if ($qcLike.Count -gt 0) {
                $r = $qcLike[-1]
            }
        }
        if ($null -eq $r -or -not ($r.PSObject.Properties.Name -contains 'IsSuccess')) {
            $types = @()
            try {
                if ($r -is [object[]]) { $types = @($r | ForEach-Object { if ($_ -eq $null) { '<null>' } else { $_.GetType().FullName } } | Select-Object -First 10) }
                elseif ($null -ne $r) { $types = @($r.GetType().FullName) }
            } catch { }
            return New-QCFailureResult -Code 'PROCESSOR_INVALID_RESULT' -Message "Handler did not return a QCResult: $handlerName" -Data @{ jobId = [string]$Job.id; jobType = $jobType; handler = $handlerName; returnTypes = $types }
        }
        return $r
    } catch {
        return New-QCFailureResult -Code 'PROCESSOR_HANDLER_THROW' -Message "Handler threw: $handlerName" -Data @{ jobId = [string]$Job.id; jobType = $jobType; handler = $handlerName; error = $_.Exception.Message }
    }
}

function Invoke-QCPrependProcessor {
    <#
    .SYNOPSIS
    Entry point for QC_PREPEND jobs.
    .DESCRIPTION
    Handles QC prepend workflow orchestration for eligible jobs.
    .PARAMETER Job
    Job payload.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: local processing actions only (no ProjectWise write operations in this stub).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    _QCP-ApplyFsThrottleConfig -Config $Config

    $qc = @{}
    if ($Config.ContainsKey('qcPrepend') -and $Config.qcPrepend) {
        $qcNorm = _QCP-ToHashtable $Config.qcPrepend
        if ($qcNorm) { $qc = $qcNorm }
    }

    $mode = 'local'
    if ($qc.ContainsKey('mode') -and $qc.mode) { $mode = ([string]$qc.mode).Trim().ToLowerInvariant() }

    $isDryRun = $false
    if ($Config.ContainsKey('dryRun')) { $isDryRun = [bool]$Config.dryRun }

    $pwMode = _QCP-TestQcPrependProjectWiseMode -Mode $mode
    $laneCtx = $null
    if (-not $pwMode) {
        $laneRes = _QCP-TryResolvePrependLaneContext -Job $Job -Config $Config
        if (-not $laneRes.IsSuccess) { return $laneRes }
        $laneCtx = [hashtable]$laneRes.Data
        _QCP-LogPrependLaneResolved -Job $Job -LaneData $laneCtx -Stage 'before_execution'
    }

    # ProjectWise-triggered QC_PREPEND: Invoke-QCPrependPw.ps1 (PW export + overlay + history).
    if ($pwMode) {
        $repoRoot = _QCP-GetRepoRoot
        $pwPrependScript = _QCP-ResolveProjectWisePrependScriptPath -QcPrependConfig $qc -RepoRoot $repoRoot
        if (-not (Test-Path -LiteralPath $pwPrependScript)) {
            return New-QCFailureResult -Code 'QC_PREPEND_LEGACY_SCRIPT_MISSING' -Message "ProjectWise prepend script not found: $pwPrependScript" -Data @{ script = $pwPrependScript }
        }

        $pwCfg = @{}
        if ($Config.ContainsKey('projectWise') -and $Config.projectWise) {
            $pwNorm = _QCP-ToHashtable $Config.projectWise
            if ($pwNorm) { $pwCfg = $pwNorm }
        }
        $ds = if ($pwCfg.ContainsKey('datasourceName') -and $pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { '' }
        if (_QCP-IsNullOrWhiteSpace $ds) { $ds = 'typsa-us-pw.bentley.com:typsa-us-pw-03' }

        $incomingFolder = ''
        # Prefer the original PW folder path from watcher metadata (preserves casing / Documents\ prefix).
        try {
            if ($Job.ContainsKey('metadata') -and $Job.metadata -and $Job.metadata.ContainsKey('candidate') -and $Job.metadata.candidate) {
                $cand = _QCP-ToHashtable $Job.metadata.candidate
                if ($cand -is [hashtable] -and $cand.ContainsKey('sourceFolder') -and $cand.sourceFolder) {
                    $incomingFolder = [string]$cand.sourceFolder
                }
            }
        } catch { }
        if (_QCP-IsNullOrWhiteSpace $incomingFolder) {
            if ($Job.ContainsKey('sourceFolder') -and $Job.sourceFolder) { $incomingFolder = [string]$Job.sourceFolder }
        }
        # pwps_dab cmdlets generally expect FolderPath without the leading "Documents\".
        if ($incomingFolder -match '^(?i)Documents\\') {
            $incomingFolder = ($incomingFolder -replace '^(?i)Documents\\', '')
        }
        $incomingDocName = if ($Job.ContainsKey('sourceName') -and $Job.sourceName) { [string]$Job.sourceName } else { '' }
        if (_QCP-IsNullOrWhiteSpace $incomingDocName) {
            $incomingDocName = [System.IO.Path]::GetFileName([string]$Job.sourcePath)
        }

        if ((_QCP-IsNullOrWhiteSpace $incomingFolder) -or (_QCP-IsNullOrWhiteSpace $incomingDocName)) {
            return New-QCFailureResult -Code 'QC_PREPEND_LEGACY_MISSING_INPUTS' -Message 'ProjectWise prepend requires Job.sourceFolder and Job.sourceName.' -Data @{ jobId = [string]$Job.id; sourceFolder = [string]$Job.sourceFolder; sourceName = [string]$Job.sourceName; sourcePath = [string]$Job.sourcePath }
        }

        $clearTag = $true
        if ($pwCfg.ContainsKey('clearTriggerTagOnSuccess')) { try { $clearTag = [bool]$pwCfg.clearTriggerTagOnSuccess } catch { $clearTag = $true } }

        if (_QCP-TestPrependAlreadyComplete -Job $Job) {
            _QCP-LogPrependProgress -Job $Job -Config $Config -Code 'QC_PREPEND_RESUME_ALREADY_COMPLETE' `
                -Message 'QC_PREPEND writeback checkpoint already complete; skipping prepend and writeback.' -Extra @{
                checkpoint = _QCP-GetJobCheckpointName -Job $Job
                incomingDocName = $incomingDocName
            }
            return New-QCSuccessResult -Code 'QC_PREPEND_OK' -Message 'QC_PREPEND already completed (writeback checkpoint).' -Data @{
                resumedFromCheckpoint = $true
                checkpoint = _QCP-GetJobCheckpointName -Job $Job
                incomingFolderPath = $incomingFolder
                incomingDocName = $incomingDocName
            }
        }

        if (_QCP-TestPrependWritebackOnlyResume -Job $Job) {
            _QCP-LogPrependProgress -Job $Job -Config $Config -Code 'QC_PREPEND_RESUME_WRITEBACK' `
                -Message 'Resuming QC_PREPEND at workflow writeback; prepend child will not run.' -Extra @{
                checkpoint = _QCP-GetJobCheckpointName -Job $Job
                incomingDocName = $incomingDocName
            }
            if ($isDryRun) {
                return New-QCSuccessResult -Code 'QC_PREPEND_DRYRUN' -Message 'Dry-run: would resume ProjectWise prepend at workflow writeback.' -Data @{
                    resumedFromCheckpoint = $true
                    checkpoint = _QCP-GetJobCheckpointName -Job $Job
                    incomingFolderPath = $incomingFolder
                    incomingDocName = $incomingDocName
                }
            }
            $resumeSuccess = New-QCSuccessResult -Code 'QC_PREPEND_OK' -Message 'QC_PREPEND resumed at workflow writeback (prepend already completed).' -Data @{
                resumedFromCheckpoint = $true
                checkpoint = _QCP-GetJobCheckpointName -Job $Job
                incomingFolderPath = $incomingFolder
                incomingDocName = $incomingDocName
            }
            $resumeParams = @{
                Job = $Job
                Config = $Config
                SuccessResult = $resumeSuccess
                IncomingFolder = $incomingFolder
                IncomingDocName = $incomingDocName
                DatasourceName = $ds
                ProjectWiseCfg = $pwCfg
            }
            if ($clearTag) { $resumeParams['ClearTriggerTag'] = $true }
            return (_QCP-InvokePrependPostSuccessWriteback @resumeParams)
        }

        $preflight = Invoke-QCPrependPreflight -Job $Job -Config $Config -IncomingFolder $incomingFolder -IncomingDocName $incomingDocName
        $preData = $null
        try {
            if ($preflight.Data -is [hashtable]) { $preData = $preflight.Data }
            elseif ($preflight.Data) {
                $preData = @{}
                foreach ($p in @($preflight.Data.PSObject.Properties)) { $preData[$p.Name] = $p.Value }
            }
        } catch { $preData = $null }
        if ($preflight.IsSuccess -and $preData -and $preData.ContainsKey('skipped') -and [bool]$preData.skipped) {
            return $preflight
        }
        if ($preflight.IsSuccess -and $preData -and $preData.ContainsKey('lane') -and $preData.lane) {
            $laneCtx = [hashtable]$preData.lane
        } else {
            $laneRes = _QCP-TryResolvePrependLaneContext -Job $Job -Config $Config
            if (-not $laneRes.IsSuccess) { return $laneRes }
            $laneCtx = [hashtable]$laneRes.Data
        }
        _QCP-LogPrependLaneResolved -Job $Job -LaneData $laneCtx -Stage 'before_execution'

        $localRoot = if ($qc.ContainsKey('localRoot') -and $qc.localRoot) { [string]$qc.localRoot } else { 'C:\PW_QC_LOCAL' }
        $logDir = if ($qc.ContainsKey('logDir') -and $qc.logDir) { [string]$qc.logDir } else { '' }

        $qpdfExe = $null
        if ($qc.ContainsKey('qpdfExePath') -and $qc.qpdfExePath) { $qpdfExe = [string]$qc.qpdfExePath }
        if (-not (_QCP-IsNullOrWhiteSpace $qpdfExe)) {
            try {
                if (-not [System.IO.Path]::IsPathRooted($qpdfExe)) {
                    $qpdfExe = Join-Path $repoRoot $qpdfExe
                }
            } catch { }
        }
        if (_QCP-IsNullOrWhiteSpace $qpdfExe) { $qpdfExe = Join-Path $repoRoot 'tools\qpdf\bin\qpdf.exe' }
        $overlayExe = $null
        if ($qc.ContainsKey('overlayExePath') -and $qc.overlayExePath) { $overlayExe = [string]$qc.overlayExePath }
        if (_QCP-IsNullOrWhiteSpace $overlayExe) { $overlayExe = Join-Path $repoRoot 'dist\qc_overlay_prepend\qc_overlay_prepend.exe' }

        $ovOldFromHistoryOnly = $true
        if ($qc.ContainsKey('overlayOldFromHistoryOnly')) { try { $ovOldFromHistoryOnly = [bool]$qc.overlayOldFromHistoryOnly } catch { $ovOldFromHistoryOnly = $true } }
        $ovSheetWorkDir = $true
        if ($qc.ContainsKey('overlaySheetWorkDir')) { try { $ovSheetWorkDir = [bool]$qc.overlaySheetWorkDir } catch { $ovSheetWorkDir = $true } }

        if ($isDryRun) {
            return New-QCSuccessResult -Code 'QC_PREPEND_DRYRUN' -Message 'Dry-run: would run ProjectWise prepend (Invoke-QCPrependPw.ps1) for ProjectWise job.' -Data @{
                jobId = [string]$Job.id
                datasourceName = $ds
                incomingFolderPath = $incomingFolder
                incomingDocName = $incomingDocName
                prependScript = $pwPrependScript
                localRoot = $localRoot
                qpdfExe = $qpdfExe
                overlayExe = $overlayExe
                qcProcessType = [string]$laneCtx.qcProcessType
                pdfSuffix = [string]$laneCtx.pdfSuffix
                expectedLanePdfName = [string]$laneCtx.expectedLanePdfName
            }
        }

        $incomingGuid = [string]$laneCtx.sourceDocumentGuid
        if (_QCP-IsNullOrWhiteSpace $incomingGuid) {
            try {
                $incomingGuid = [string](_QCP-GetJobMetadataValue -Job $Job -Keys @('sourceDocumentGuid', 'triggerDocumentGuid', 'documentGuid'))
            } catch { $incomingGuid = '' }
        }

        $historyDocName = [string]$laneCtx.expectedLanePdfName
        $args = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', $pwPrependScript,
            '-IncomingFolderPath', $incomingFolder,
            '-IncomingDocName', $incomingDocName,
            '-DatasourceName', $ds,
            '-LocalRoot', $localRoot
        )
        if (-not (_QCP-IsNullOrWhiteSpace $incomingGuid)) {
            $args += @('-IncomingDocumentGuid', $incomingGuid)
        }
        if (-not (_QCP-IsNullOrWhiteSpace $logDir)) { $args += @('-LogDir', $logDir) }
        if (Test-Path -LiteralPath $qpdfExe) { $args += @('-QpdfExe', $qpdfExe) }
        if (Test-Path -LiteralPath $overlayExe) { $args += @('-QcOverlayExe', $overlayExe) }
        $enableOverlayPw = $true
        if ($qc.ContainsKey('enableOverlay')) { try { $enableOverlayPw = [bool]$qc.enableOverlay } catch { $enableOverlayPw = $true } }
        if (-not $enableOverlayPw) { $args += '-NoOverlayLayers' }
        $args += @('-OverlayOldFromHistoryOnly', ([string]$ovOldFromHistoryOnly), '-OverlaySheetWorkDir', ([string]$ovSheetWorkDir))
        $appsettingsPath = Join-Path $repoRoot 'appsettings.json'
        if ($Config.ContainsKey('appsettingsPath') -and $Config.appsettingsPath) {
            $appsettingsPath = [string]$Config.appsettingsPath
        }
        if (Test-Path -LiteralPath $appsettingsPath) {
            $args += @('-AppsettingsPath', $appsettingsPath)
        }
        $prependTrigger = _QCP-ResolvePrependTrigger -Job $Job
        if (-not (_QCP-IsNullOrWhiteSpace $prependTrigger)) {
            $args += @('-PrependTrigger', $prependTrigger)
        }
        $args += @(
            '-QcProcessType', [string]$laneCtx.qcProcessType,
            '-QcPdfSuffix', [string]$laneCtx.pdfSuffix,
            '-HistoryDocumentName', $historyDocName,
            '-HistoryDocName', $historyDocName
        )

        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_PW_LANE_PARAMS' `
                -Message 'ProjectWise prepend invoked with resolved lane parameters.' -Data @{
                jobId = [string]$Job.id
                qc_process_type = [string]$laneCtx.qcProcessType
                qcProcessType = [string]$laneCtx.qcProcessType
                pdfSuffix = [string]$laneCtx.pdfSuffix
                expectedLanePdfName = $historyDocName
                historyDocumentName = $historyDocName
                triggerDocumentName = [string]$laneCtx.triggerDocumentName
                sourceDocumentGuid = [string]$laneCtx.sourceDocumentGuid
                prependScript = $pwPrependScript
            } | Out-Null
        }

        $stdoutPath = [System.IO.Path]::GetTempFileName()
        $stderrPath = [System.IO.Path]::GetTempFileName()
        try {
            # Windows PowerShell can mis-handle string[] -ArgumentList when paths contain spaces.
            # Use a single quoted arg line (same approach as _QCP-RunOverlayExe).
            $cleanArgs = @('-MTA') + @($args)
            $argLine = (($cleanArgs | ForEach-Object {
                $t = [string]$_
                if ($t -match '[\s"]') { return ('"' + ($t -replace '"', '\\"') + '"') }
                return $t
            }) -join ' ')

            _QCP-LogPrependProgress -Job $Job -Config $Config -Code 'QC_PREPEND_PW_CHILD_START' `
                -Message 'ProjectWise prepend child process starting.' -Extra @{
                prependScript = $pwPrependScript
                incomingDocName = $incomingDocName
            }
            $childWaitTimeout = _QCP-GetPrependChildWaitTimeoutSeconds -QcPrependConfig $qc
            $child = _QCP-StartAndWaitLaunchedProcess -FilePath 'powershell.exe' -ArgumentList $argLine `
                -StdoutPath $stdoutPath -StderrPath $stderrPath -Job $Job -Config $Config `
                -TimeoutSeconds $childWaitTimeout
            if ($child.startFailed) {
                return New-QCFailureResult -Code 'QC_PREPEND_PW_FAILED' -Message 'Failed to start ProjectWise prepend child process.' -Data @{
                    prependScript = $pwPrependScript
                    incomingDocName = $incomingDocName
                }
            }
            _QCP-LogPrependProgress -Job $Job -Config $Config -Code 'QC_PREPEND_PW_CHILD_DONE' `
                -Message 'ProjectWise prepend child PowerShell process finished.' -Extra @{
                exitCode = $child.exitCode
                childPid = $child.processId
                timedOut = [bool]$child.timedOut
                elapsedMs = $child.elapsedMs
                incomingDocName = $incomingDocName
            }
            $stdout = [string]$child.stdout
            $stderr = [string]$child.stderr
            if ($child.timedOut -or -not $child.exited) {
                $failData = _QCP-SanitizeProcessorDataForStorage @{
                    exitCode = $child.exitCode
                    timedOut = $true
                    childPid = $child.processId
                    elapsedMs = $child.elapsedMs
                    stdout = $stdout
                    stderr = $stderr
                    args = $args
                    argLine = $argLine
                    prependScript = $pwPrependScript
                }
                return New-QCFailureResult -Code 'QC_PREPEND_PW_CHILD_TIMEOUT' -Message 'ProjectWise prepend child PowerShell process timed out; prepend was not marked complete.' -Data $failData
            }
            if ($null -eq $child.exitCode -or [int]$child.exitCode -ne 0) {
                $failData = _QCP-SanitizeProcessorDataForStorage @{
                    exitCode = $child.exitCode
                    childPid = $child.processId
                    stdout = $stdout
                    stderr = $stderr
                    args = $args
                    argLine = $argLine
                    prependScript = $pwPrependScript
                }
                return New-QCFailureResult -Code 'QC_PREPEND_PW_FAILED' -Message 'ProjectWise prepend failed.' -Data $failData
            }

            _QCP-SetPrependJobCheckpoint -Job $Job -Config $Config -Checkpoint $script:_QCP_CheckpointPrependComplete -CheckpointData @{
                exitCode = [int]$child.exitCode
                childPid = $child.processId
                atUtc = (Get-Date).ToUniversalTime().ToString('o')
                incomingDocName = $incomingDocName
            }
            $tagCleared = $null
            $tagClearError = $null
            $success = New-QCSuccessResult -Code 'QC_PREPEND_OK' -Message 'QC_PREPEND completed via ProjectWise prepend.' -Data (_QCP-SanitizeProcessorDataForStorage @{
                exitCode = [int]$child.exitCode
                childPid = $child.processId
                stdout = $stdout
                stderr = $stderr
                prependScript = $pwPrependScript
                incomingFolderPath = $incomingFolder
                incomingDocName = $incomingDocName
                triggerTagCleared = $tagCleared
                triggerTagClearError = $tagClearError
            })
            $writebackParams = @{
                Job = $Job
                Config = $Config
                SuccessResult = $success
                IncomingFolder = $incomingFolder
                IncomingDocName = $incomingDocName
                DatasourceName = $ds
                ProjectWiseCfg = $pwCfg
            }
            if ($clearTag) { $writebackParams['ClearTriggerTag'] = $true }
            return (_QCP-InvokePrependPostSuccessWriteback @writebackParams)
        } finally {
            Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
        }
    }

    $sourcePdf = _QCP-ResolveSourcePdf -Job $Job
    if (_QCP-IsNullOrWhiteSpace $sourcePdf) {
        return New-QCFailureResult -Code 'QC_PREPEND_SOURCE_MISSING' -Message 'Job does not include a source PDF path.' -Data @{ jobId = [string]$Job.id }
    }
    if (-not (Test-Path -LiteralPath $sourcePdf)) {
        return New-QCFailureResult -Code 'QC_PREPEND_SOURCE_NOT_FOUND' -Message "Source PDF not found: $sourcePdf" -Data @{ jobId = [string]$Job.id; sourcePdf = $sourcePdf }
    }

    $processType = [string]$laneCtx.qcProcessType
    $laneSuffix = [string]$laneCtx.pdfSuffix
    $isProductionLane = ($processType -eq 'production')

    $histRes = _QCP-ResolveHistoryPath -Job $Job -Config $Config
    if (-not $histRes.IsSuccess) { return $histRes }
    $historyPdf = [string]$histRes.Data.historyPdf
    $historyDir = Split-Path -Parent $historyPdf

    $enableOverlay = $false
    if ($qc.ContainsKey('enableOverlay')) { $enableOverlay = [bool]$qc.enableOverlay }
    $overlayExePath = if ($qc.ContainsKey('overlayExePath') -and $qc.overlayExePath) { [string]$qc.overlayExePath } else { (Join-Path (_QCP-GetRepoRoot) 'dist\qc_overlay_prepend\qc_overlay_prepend.exe') }

    # Size gate (parity with Invoke-QCPrependPw): skip overlay when incoming exceeds threshold.
    $overlaySkipDecision = $null
    if ($enableOverlay -and (Get-Command -Name 'Test-QCOverlaySkipByIncomingSize' -ErrorAction SilentlyContinue)) {
        $incomingLen = [long]0
        try { $incomingLen = [long](Get-Item -LiteralPath $sourcePdf -ErrorAction Stop).Length } catch { $incomingLen = [long]0 }
        $overlaySkipDecision = Test-QCOverlaySkipByIncomingSize -QcPrepend $qc -IncomingBytes $incomingLen
        if ($overlaySkipDecision.Skip) {
            $enableOverlay = $false
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Information' -Code 'OVERLAY_SKIPPED_INCOMING_SIZE' `
                    -Message 'Skipping layered overlay; incoming PDF exceeds size threshold (qpdf prepend).' -Data @{
                    jobId = [string]$Job.id
                    incomingBytes = $overlaySkipDecision.IncomingBytes
                    thresholdMb = $overlaySkipDecision.ThresholdMb
                    thresholdBytes = $overlaySkipDecision.ThresholdBytes
                    reason = $overlaySkipDecision.Reason
                    sourcePdf = $sourcePdf
                } | Out-Null
            }
        }
    }

    # Output path for overlay results (optional)
    $outputRoot = if ($qc.ContainsKey('outputRoot') -and $qc.outputRoot) { [string]$qc.outputRoot } else { '' }
    $tempRoot = if ($qc.ContainsKey('tempRoot') -and $qc.tempRoot) { [string]$qc.tempRoot } else { (Join-Path ([System.IO.Path]::GetTempPath()) 'qc_prepend') }
    _QCP-EnsureDir -Path $tempRoot

    $overlayOutPdf = $null
    if (-not (_QCP-IsNullOrWhiteSpace $outputRoot)) {
        _QCP-EnsureDir -Path $outputRoot
        $sourceName = if ($Job.sourceName) { [string]$Job.sourceName } else { [System.IO.Path]::GetFileName($sourcePdf) }
        $base = $sourceName
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            $stem = Get-PWSheetStemFromDocumentName -DocumentName $sourceName
            if (-not (_QCP-IsNullOrWhiteSpace $stem)) { $base = $stem }
        } else {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($sourceName)
        }
        if (_QCP-IsNullOrWhiteSpace $base) { $base = [string]$Job.id }
        $folderKey = 'root'
        if ($histRes.Data -and ($histRes.Data -is [hashtable]) -and $histRes.Data.ContainsKey('folderKey') -and $histRes.Data.folderKey) {
            $folderKey = [string]$histRes.Data.folderKey
        }
        $outDir = Join-Path $outputRoot $folderKey
        _QCP-EnsureDir -Path $outDir
        $overlayOutPdf = Join-Path $outDir ([string]$laneCtx.expectedLanePdfName)
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_LANE_RESOLVED' -Message 'Prepend lane output resolved.' -Data @{
                qc_process_type = $processType
                qcProcessType = $processType
                pdfSuffix = $laneSuffix
                expectedLanePdfName = [string]$laneCtx.expectedLanePdfName
                qcOutputPdf = $overlayOutPdf
                sheetBaseName = $base
                triggerDocumentName = [string]$laneCtx.triggerDocumentName
                sourceDocumentGuid = [string]$laneCtx.sourceDocumentGuid
            } | Out-Null
        }
    }

    $overlayTemplate = if ($qc.ContainsKey('overlayArgumentsTemplate') -and $qc.overlayArgumentsTemplate) { [string]$qc.overlayArgumentsTemplate } else { '' }
    $overlayCmdWouldRun = $null
    if ($enableOverlay) {
        $overlayCmdWouldRun = @{
            exe = $overlayExePath
            argsTemplate = $overlayTemplate
            outputPdf = $overlayOutPdf
            historyPdf = $historyPdf
            sourcePdf = $sourcePdf
        }
    }

    if ($isDryRun) {
        return New-QCSuccessResult -Code 'QC_PREPEND_DRYRUN' -Message 'Dry-run: QC_PREPEND would update history and optionally run overlay.' -Data @{
            jobId = [string]$Job.id
            sourcePdf = $sourcePdf
            targetHistoryPdf = $historyPdf
            historyExists = (Test-Path -LiteralPath $historyPdf)
            overlayEnabled = $enableOverlay
            overlayCommandWouldRun = $overlayCmdWouldRun
            qcOutputPdf = $overlayOutPdf
        }
    }

    $qpdfExe = _QCP-GetQpdfExePath -Config $Config
    if (-not (Test-Path -LiteralPath $qpdfExe)) {
        return New-QCFailureResult -Code 'PDF_MERGE_TOOL_MISSING' -Message "qpdf.exe not found: $qpdfExe" -Data @{ qpdfExe = $qpdfExe }
    }

    _QCP-EnsureDir -Path $historyDir

    $tmpHistory = Join-Path $tempRoot ([string]$Job.id + '.history.tmp.' + ([guid]::NewGuid().ToString('N')) + '.pdf')
    $tmpQcOut = Join-Path $tempRoot ([string]$Job.id + '.qc.tmp.' + ([guid]::NewGuid().ToString('N')) + '.pdf')

    try {
        $historyExists = Test-Path -LiteralPath $historyPdf

        # 1) If overlay enabled, build QC output first:
        #    qc.pdf = overlay(source vs page1(history)) + history tail pages
        #    Use tools/overlay/qc_overlay_prepend's contract directly: incoming + qc_history -> merged QC output.
        if ($enableOverlay) {
            if (-not (Test-Path -LiteralPath $overlayExePath)) {
                return New-QCFailureResult -Code 'QC_OVERLAY_EXE_MISSING' -Message "Overlay exe missing: $overlayExePath" -Data @{ overlayExePath = $overlayExePath }
            }
            if (_QCP-IsNullOrWhiteSpace $overlayOutPdf) {
                return New-QCFailureResult -Code 'QC_OVERLAY_OUTPUT_MISSING' -Message 'qcPrepend.outputRoot is required to produce QC output when overlay is enabled.' -Data @{}
            }

            $overlayArgs = @(
                $sourcePdf,
                $historyPdf,
                '-o',
                $tmpQcOut
            )

            # Pass through selected overlay engine options if present in config.
            # Mirrors tools/overlay/qc_overlay_prepend.py CLI.
            if ($qc.ContainsKey('overlayAlpha') -and $qc.overlayAlpha -ne $null -and ([string]$qc.overlayAlpha).Trim().Length -gt 0) {
                $overlayArgs += @('--alpha', [string]$qc.overlayAlpha)
            }
            if ($qc.ContainsKey('overlayFit') -and [bool]$qc.overlayFit) {
                $overlayArgs += @('--fit')
            }
            if ($qc.ContainsKey('overlayVerbose') -and [bool]$qc.overlayVerbose) {
                $overlayArgs += @('--verbose')
            }
            if ($qc.ContainsKey('overlayKeepTemp') -and [bool]$qc.overlayKeepTemp) {
                $overlayArgs += @('--keep-temp')
            }
            if ($qc.ContainsKey('overlayNoFlattenSources') -and [bool]$qc.overlayNoFlattenSources) {
                $overlayArgs += @('--no-flatten-sources')
            }
            if ($qc.ContainsKey('overlayFlattenRaster') -and [bool]$qc.overlayFlattenRaster) {
                $overlayArgs += @('--flatten-raster')
            }
            if ($qc.ContainsKey('overlayFlattenDpi') -and $qc.overlayFlattenDpi -ne $null -and ([string]$qc.overlayFlattenDpi).Trim().Length -gt 0) {
                $overlayArgs += @('--flatten-dpi', [string]$qc.overlayFlattenDpi)
            }
            if ($qc.ContainsKey('overlayCurrentMasterPath') -and $qc.overlayCurrentMasterPath) {
                $overlayArgs += @('--current-master', [string]$qc.overlayCurrentMasterPath)
            }
            if ($qc.ContainsKey('overlaySheetWorkDir') -and $qc.overlaySheetWorkDir) {
                $overlayArgs += @('--sheet-work-dir', [string]$qc.overlaySheetWorkDir)
            }

            $overlayRes = $null
            try {
                $overlayRes = _QCP-RunOverlayExe -ExePath $overlayExePath -OverlayArgs $overlayArgs -WorkingDirectory (Split-Path -Parent $sourcePdf)
            } catch {
                return New-QCFailureResult -Code 'QC_OVERLAY_LAUNCH_FAILED' -Message 'Failed to launch overlay exe.' -Data @{ exe = $overlayExePath; errorMessage = $_.Exception.Message; args = $overlayArgs }
            }
            if ([int]$overlayRes.exitCode -ne 0) {
                return New-QCFailureResult -Code 'QC_OVERLAY_FAILED' -Message 'Overlay exe failed.' -Data @{ exitCode = $overlayRes.exitCode; stderr = $overlayRes.stderr; stdout = $overlayRes.stdout; args = $overlayRes.args; argLine = $overlayRes.argLine; exe = $overlayExePath }
            }

            $stampCfg = if (Get-Command -Name 'Get-QCReviewStampSettings' -ErrorAction SilentlyContinue) {
                Get-QCReviewStampSettings -Config $Config
            } else { $null }
            if ($stampCfg) {
                $stampRes = _QCP-TryApplyReviewStampFromJob -PdfPath $tmpQcOut -Job $Job -Config $Config `
                    -FolderPath ([string]$Job.sourceFolder) -SourceDocumentName ([string]$Job.sourceName) `
                    -OverlayExe $overlayExePath
                $jobRt = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewType', 'qcReviewType')
                if (_QCP-IsNullOrWhiteSpace $jobRt) { $jobRt = [string]$stampRes.reviewType }
                if (_QCP-ReviewStampRequiredForReviewType -StampSettings $stampCfg -ReviewType $jobRt -ProcessType $processType -and -not $stampRes.skipped -and -not $stampRes.applied) {
                    return New-QCFailureResult -Code 'QC_REVIEW_STAMP_FAILED' -Message 'Review stamp failed on QC PDF.' -Data $stampRes
                }
            }

            try {
                Move-Item -LiteralPath $tmpQcOut -Destination $overlayOutPdf -Force -ErrorAction Stop
                _QCP-FsThrottle
            } catch {
                return New-QCFailureResult -Code 'QC_PREPEND_QC_WRITE_FAILED' -Message 'Failed to write QC output PDF.' -Data @{ tmpQcOut = $tmpQcOut; qcOutputPdf = $overlayOutPdf; errorMessage = $_.Exception.Message }
            }

            if (-not $isProductionLane) {
                $laneSuccess = New-QCSuccessResult -Code 'QC_PREPEND_OK' -Message 'QC_PREPEND lane PDF updated.' -Data @{
                    jobId = [string]$Job.id
                    sourcePdf = $sourcePdf
                    qcOutputPdf = $overlayOutPdf
                    qcProcessType = $processType
                    overlayEnabled = $enableOverlay
                }
                $laneResult = _QCP-AppendWorkflowWriteback -Result $laneSuccess -Job $Job -Config $Config -SourcePath $sourcePdf -OutputPath $overlayOutPdf -HistoryPath $null
                return $laneResult
            }
        }

        # 2) Update history (production lane only):
        #    history.pdf = source + oldHistory  (or init from source)
        if (-not $isProductionLane) {
            # Size-skipped overlay: still produce lane PDF via stamped copy of incoming (no Old/New/Current layers).
            if ($overlaySkipDecision -and $overlaySkipDecision.Skip) {
                if (_QCP-IsNullOrWhiteSpace $overlayOutPdf) {
                    return New-QCFailureResult -Code 'QC_OVERLAY_OUTPUT_MISSING' -Message 'qcPrepend.outputRoot is required to produce QC output when overlay is size-skipped.' -Data @{}
                }
                try {
                    Copy-Item -LiteralPath $sourcePdf -Destination $tmpQcOut -Force -ErrorAction Stop
                    _QCP-FsThrottle
                } catch {
                    return New-QCFailureResult -Code 'QC_PREPEND_QC_WRITE_FAILED' -Message 'Failed to stage size-skipped lane PDF.' -Data @{ sourcePdf = $sourcePdf; tmpQcOut = $tmpQcOut; errorMessage = $_.Exception.Message }
                }
                $stampCfg = if (Get-Command -Name 'Get-QCReviewStampSettings' -ErrorAction SilentlyContinue) {
                    Get-QCReviewStampSettings -Config $Config
                } else { $null }
                if ($stampCfg) {
                    $stampRes = _QCP-TryApplyReviewStampFromJob -PdfPath $tmpQcOut -Job $Job -Config $Config `
                        -FolderPath ([string]$Job.sourceFolder) -SourceDocumentName ([string]$Job.sourceName) `
                        -OverlayExe $overlayExePath
                    $jobRt = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewType', 'qcReviewType')
                    if (_QCP-IsNullOrWhiteSpace $jobRt) { $jobRt = [string]$stampRes.reviewType }
                    if (_QCP-ReviewStampRequiredForReviewType -StampSettings $stampCfg -ReviewType $jobRt -ProcessType $processType -and -not $stampRes.skipped -and -not $stampRes.applied) {
                        return New-QCFailureResult -Code 'QC_REVIEW_STAMP_FAILED' -Message 'Review stamp failed on QC PDF.' -Data $stampRes
                    }
                }
                try {
                    Move-Item -LiteralPath $tmpQcOut -Destination $overlayOutPdf -Force -ErrorAction Stop
                    _QCP-FsThrottle
                } catch {
                    return New-QCFailureResult -Code 'QC_PREPEND_QC_WRITE_FAILED' -Message 'Failed to write size-skipped QC output PDF.' -Data @{ tmpQcOut = $tmpQcOut; qcOutputPdf = $overlayOutPdf; errorMessage = $_.Exception.Message }
                }
                $laneSuccess = New-QCSuccessResult -Code 'QC_PREPEND_OK' -Message 'QC_PREPEND lane PDF updated (overlay size-skipped).' -Data @{
                    jobId = [string]$Job.id
                    sourcePdf = $sourcePdf
                    qcOutputPdf = $overlayOutPdf
                    qcProcessType = $processType
                    overlayEnabled = $false
                    overlaySkipReason = $overlaySkipDecision.Reason
                }
                return (_QCP-AppendWorkflowWriteback -Result $laneSuccess -Job $Job -Config $Config -SourcePath $sourcePdf -OutputPath $overlayOutPdf -HistoryPath $null)
            }
            return New-QCFailureResult -Code 'QC_PREPEND_LANE_HISTORY_SKIPPED' -Message 'Non-production lane requires overlay output.' -Data @{ qcProcessType = $processType }
        }
        if (-not $historyExists) {
            try {
                Copy-Item -LiteralPath $sourcePdf -Destination $tmpHistory -Force -ErrorAction Stop
                _QCP-FsThrottle
            } catch {
                return New-QCFailureResult -Code 'QC_PREPEND_HISTORY_INIT_FAILED' -Message 'Failed to create initial history PDF (copy failed).' -Data @{ sourcePdf = $sourcePdf; tmpHistory = $tmpHistory; errorMessage = $_.Exception.Message }
            }
        } else {
            $merge = _QCP-RunQpdfPrepend -QpdfExe $qpdfExe -SourcePdf $sourcePdf -HistoryPdf $historyPdf -OutPdf $tmpHistory
            if ([int]$merge.exitCode -ne 0) {
                return New-QCFailureResult -Code 'QC_PREPEND_MERGE_FAILED' -Message 'qpdf prepend/merge failed.' -Data @{ exitCode = $merge.exitCode; stderr = $merge.stderr; stdout = $merge.stdout; args = $merge.args }
            }
        }

        if (-not $enableOverlay) {
            $stampCfg = if (Get-Command -Name 'Get-QCReviewStampSettings' -ErrorAction SilentlyContinue) {
                Get-QCReviewStampSettings -Config $Config
            } else { $null }
            if ($stampCfg) {
                $stampRes = _QCP-TryApplyReviewStampFromJob -PdfPath $tmpHistory -Job $Job -Config $Config `
                    -FolderPath ([string]$Job.sourceFolder) -SourceDocumentName ([string]$Job.sourceName) `
                    -OverlayExe $overlayExePath
                $rt = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewType', 'qcReviewType')
                if (_QCP-ReviewStampRequiredForReviewType -StampSettings $stampCfg -ReviewType $rt -and -not $stampRes.skipped -and -not $stampRes.applied) {
                    return New-QCFailureResult -Code 'QC_REVIEW_STAMP_FAILED' -Message 'Review stamp failed on history PDF.' -Data $stampRes
                }
            }
        }

        try {
            Move-Item -LiteralPath $tmpHistory -Destination $historyPdf -Force -ErrorAction Stop
            _QCP-FsThrottle
        } catch {
            return New-QCFailureResult -Code 'QC_PREPEND_HISTORY_WRITE_FAILED' -Message 'Failed to write updated history PDF (move/replace failed).' -Data @{ tmpHistory = $tmpHistory; targetHistoryPdf = $historyPdf; errorMessage = $_.Exception.Message }
        }

        $success = New-QCSuccessResult -Code 'QC_PREPEND_OK' -Message 'QC_PREPEND history updated.' -Data @{
            jobId = [string]$Job.id
            sourcePdf = $sourcePdf
            targetHistoryPdf = $historyPdf
            overlayEnabled = $enableOverlay
            overlayExe = if ($enableOverlay) { $overlayExePath } else { $null }
            qcOutputPdf = $overlayOutPdf
        }
        return (_QCP-AppendWorkflowWriteback -Result $success -Job $Job -Config $Config -SourcePath $sourcePdf -OutputPath $overlayOutPdf -HistoryPath $historyPdf)
    } finally {
        if (Test-Path -LiteralPath $tmpHistory) {
            Remove-Item -LiteralPath $tmpHistory -Force -ErrorAction SilentlyContinue
            _QCP-FsThrottle
        }
        if (Test-Path -LiteralPath $tmpQcOut) {
            Remove-Item -LiteralPath $tmpQcOut -Force -ErrorAction SilentlyContinue
            _QCP-FsThrottle
        }
    }
}

function Invoke-StatusSetProcessor {
    <#
    .SYNOPSIS
    Entry point for STATUS_SET_GEN jobs.
    .DESCRIPTION
    Handles status set generation workflow orchestration for eligible jobs.
    .PARAMETER Job
    Job payload.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: depends on statusSet.mode (stub, native: manifest + PW, legacy: external script).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $isDryRun = $false
    if ($Config.ContainsKey('dryRun')) { $isDryRun = [bool]$Config.dryRun }

    $sourceFolder = if ($Job.ContainsKey('sourceFolder') -and $Job.sourceFolder) { [string]$Job.sourceFolder } else { '' }
    if (_QCP-IsNullOrWhiteSpace $sourceFolder) {
        return New-QCFailureResult -Code 'STATUS_SET_MISSING_SOURCE_FOLDER' -Message 'STATUS_SET_GEN requires Job.sourceFolder (target folder / sheets folder).' -Data @{ jobId = [string]$Job.id }
    }

    $ss = @{}
    if ($Config.ContainsKey('statusSet') -and $Config.statusSet) {
        $ssNorm = _QCP-ToHashtable $Config.statusSet
        if ($ssNorm) { $ss = $ssNorm }
    }

    # Default to the supported native implementation (QC.StatusSet.psm1). The previous
    # default was 'stub', which returned SUCCESS without doing any work whenever the
    # config lacked a statusSet block - jobs appeared to succeed in <1s while PW was
    # never updated and no PDF was written. Misconfigured/unknown modes now fail loudly
    # instead of silently succeeding.
    $mode = 'native'
    if ($ss.ContainsKey('mode') -and $ss.mode) { $mode = ([string]$ss.mode).Trim().ToLowerInvariant() }

    if ($mode -eq 'native') {
        $ssMod = Join-Path $PSScriptRoot 'QC.StatusSet.psm1'
        if (-not (Test-Path -LiteralPath $ssMod)) {
            return New-QCFailureResult -Code 'STATUS_SET_MODULE_MISSING' -Message "QC.StatusSet.psm1 not found: $ssMod" -Data @{ path = $ssMod }
        }
        Import-Module $ssMod -Force
        return Invoke-StatusSetNativeJob -Job $Job -Config $Config
    }

    if ($mode -eq 'stub') {
        return New-QCSuccessResult -Code 'STATUS_SET_STUB_OK' -Message 'STATUS_SET_GEN explicitly configured as stub (statusSet.mode = "stub"); no work performed.' -Data @{
            jobId = [string]$Job.id
            jobType = [string]$Job.type
            sourceFolder = $sourceFolder
            mode = $mode
        }
    }

    if ($mode -ne 'legacy') {
        return New-QCFailureResult -Code 'STATUS_SET_UNKNOWN_MODE' -Message ("statusSet.mode '{0}' is not recognized. Valid values: native, legacy, stub." -f $mode) -Data @{
            jobId = [string]$Job.id
            jobType = [string]$Job.type
            sourceFolder = $sourceFolder
            mode = $mode
        }
    }

    if ($isDryRun) {
        return New-QCSuccessResult -Code 'STATUS_SET_DRYRUN' -Message 'Dry-run: would invoke legacy combine_status_set.ps1.' -Data @{
            jobId = [string]$Job.id
            jobType = [string]$Job.type
            sourceFolder = $sourceFolder
            mode = $mode
        }
    }

    $repoRoot = _QCP-GetRepoRoot
    $combineScript = if ($ss.ContainsKey('legacyScriptPath') -and $ss.legacyScriptPath) { [string]$ss.legacyScriptPath } else { (Join-Path $repoRoot 'legacy\combine_status_set.ps1') }
    if (-not (Test-Path -LiteralPath $combineScript)) {
        return New-QCFailureResult -Code 'STATUS_SET_LEGACY_SCRIPT_MISSING' -Message "Legacy combine_status_set.ps1 not found: $combineScript" -Data @{ script = $combineScript }
    }

    $datasource = if ($ss.ContainsKey('datasourceName') -and $ss.datasourceName) { [string]$ss.datasourceName } else { 'typsa-us-pw.bentley.com:typsa-us-pw-03' }
    $localRoot = if ($ss.ContainsKey('localRoot') -and $ss.localRoot) { [string]$ss.localRoot } else { 'C:\PW_QC_LOCAL' }
    $logDir = if ($ss.ContainsKey('logDir') -and $ss.logDir) { [string]$ss.logDir } else { '' }
    $qpdfExe = if ($ss.ContainsKey('qpdfExe') -and $ss.qpdfExe) { [string]$ss.qpdfExe } else { 'qpdf' }
    $forceRebuild = $false
    if ($ss.ContainsKey('forceRebuild')) { try { $forceRebuild = [bool]$ss.forceRebuild } catch { $forceRebuild = $false } }
    $promptForCredential = $false
    if ($ss.ContainsKey('promptForCredential')) { try { $promptForCredential = [bool]$ss.promptForCredential } catch { $promptForCredential = $false } }
    $writeBackToPW = $false
    if ($ss.ContainsKey('writeBackToPW')) { try { $writeBackToPW = [bool]$ss.writeBackToPW } catch { $writeBackToPW = $false } }

    $args = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $combineScript,
        '-SheetsFolderPath', $sourceFolder,
        '-DatasourceName', $datasource,
        '-LocalRoot', $localRoot,
        '-QpdfExe', $qpdfExe,
        '-PollIntervalSeconds', '0',
        '-RunOnce'
    )
    if (-not (_QCP-IsNullOrWhiteSpace $logDir)) { $args += @('-LogDir', $logDir) }
    if ($writeBackToPW) { $args += @('-WriteBackToPW') }
    if ($forceRebuild) { $args += @('-ForceRebuild') }
    if ($promptForCredential) { $args += @('-PromptForCredential') }

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    try {
        # combine_status_set.ps1 re-launches itself in MTA if needed; but call powershell.exe -MTA anyway.
        $p = Start-Process -FilePath 'powershell.exe' -ArgumentList (@('-MTA') + $args) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $stdout = Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue
        $stderr = Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue

        if ($p.ExitCode -ne 0) {
            return New-QCFailureResult -Code 'STATUS_SET_LEGACY_FAILED' -Message 'Legacy combine_status_set.ps1 failed.' -Data @{
                jobId = [string]$Job.id
                exitCode = [int]$p.ExitCode
                stdout = $stdout
                stderr = $stderr
                script = $combineScript
                args = $args
                writeBackToPW = $writeBackToPW
            }
        }

        return New-QCSuccessResult -Code 'STATUS_SET_OK' -Message 'STATUS_SET_GEN completed via legacy combine_status_set.ps1.' -Data @{
            jobId = [string]$Job.id
            exitCode = [int]$p.ExitCode
            stdout = $stdout
            stderr = $stderr
            script = $combineScript
            args = $args
            writeBackToPW = $writeBackToPW
        }
    } finally {
        Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function _QCP-EnsureQueueModulesLoaded {
    if (-not (Get-Command -Name 'Add-QCQueueJob' -ErrorAction SilentlyContinue)) {
        Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Queue/QC.Queue.Json.psm1') -Force -ErrorAction Stop
    }
    if (-not (Get-Command -Name 'New-QCJobObject' -ErrorAction SilentlyContinue)) {
        Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Queue/QC.JobFactory.psm1') -Force -ErrorAction Stop
    }
}

function _QCP-EnsureNotificationModulesLoaded {
    if (Get-Command -Name 'Invoke-QCNotificationForStateChange' -ErrorAction SilentlyContinue) { return }
    $notifPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Notifications/QC.Notifications.psm1'
    if (Test-Path -LiteralPath $notifPath) {
        Import-Module $notifPath -Force -WarningAction SilentlyContinue | Out-Null
    }
}

function _QCP-ResolveSheetPdfForPrependTrigger {
    param([string]$TriggerDocumentName)
    $name = [System.IO.Path]::GetFileName([string]$TriggerDocumentName)
    if ([string]::IsNullOrWhiteSpace($name)) { return $null }
    if (Test-QCLegacyQcPdfDocumentName -DocumentName $name) { return $null }
    if (Get-Command -Name 'Test-QCIsQcPdfDocumentName' -ErrorAction SilentlyContinue) {
        if (Test-QCIsQcPdfDocumentName -DocumentName $name) {
            $stem = ''
            if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
                $stem = Get-PWSheetStemFromDocumentName -DocumentName $name
            }
            if ([string]::IsNullOrWhiteSpace($stem)) {
                $stem = [System.IO.Path]::GetFileNameWithoutExtension($name) -replace '(?i)-(prod|chk|rev)$', ''
            }
            if (-not [string]::IsNullOrWhiteSpace($stem)) { return ($stem + '.pdf') }
        }
    } elseif ($name -match '(?i)-(prod|chk|rev)\.pdf$') {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($name) -replace '(?i)-(prod|chk|rev)$', ''
        if (-not [string]::IsNullOrWhiteSpace($stem)) { return ($stem + '.pdf') }
    }
    if ($name -match '(?i)\.dgn$') {
        return ([System.IO.Path]::GetFileNameWithoutExtension($name) + '.pdf')
    }
    if ($name -match '(?i)\.pdf$') { return $name }
    return $null
}

function _QCP-GetQcInitiatedStateTransitionKey {
    param(
        [Nullable[long]]$AuditEventId = $null,
        [string]$LastAuditEventAt = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$TriggerDocumentGuid = ''
    )
    if (Get-Command -Name 'Get-QCAuditStateTransitionKey' -ErrorAction SilentlyContinue) {
        return Get-QCAuditStateTransitionKey -AuditEventId $AuditEventId -LastAuditEventAt $LastAuditEventAt `
            -ChangedByUser $ChangedByUser -TriggerDocumentGuid $TriggerDocumentGuid
    }
    if ($null -ne $AuditEventId -and $AuditEventId -gt 0) {
        return ('audit:' + [string]$AuditEventId)
    }
    $at = ([string]$LastAuditEventAt).Trim()
    $guid = ([string]$TriggerDocumentGuid).Trim()
    if ($at.Length -eq 0) { return $null }
    $userPart = ''
    if ($null -ne $ChangedByUser -and $ChangedByUser -gt 0) { $userPart = ('user:' + [string]$ChangedByUser + '|') }
    return ($userPart + 'at:' + $at + '|doc:' + $guid)
}

function _QCP-GetSheetPrependStateTransitionKey {
    param(
        [string]$SheetStem = '',
        [string]$FromState = '',
        [string]$ToState = ''
    )
    $stem = ([string]$SheetStem).Trim().ToLowerInvariant()
    $from = ([string]$FromState).Trim().ToLowerInvariant()
    $to = ([string]$ToState).Trim().ToLowerInvariant()
    if ($stem.Length -eq 0 -or $to.Length -eq 0) { return $null }
    return ('sheet:' + $stem + '|from:' + $from + '|to:' + $to)
}

function _QCP-GetQueueRootFromConfig {
    param([hashtable]$Config)
    if ($Config.ContainsKey('queue') -and $Config.queue) {
        if ($Config.queue.ContainsKey('rootDir') -and $Config.queue.rootDir) { return [string]$Config.queue.rootDir }
        if ($Config.queue.ContainsKey('root') -and $Config.queue.root) { return [string]$Config.queue.root }
        if ($Config.queue.ContainsKey('path') -and $Config.queue.path) { return [string]$Config.queue.path }
    }
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    return (Join-Path $repoRoot 'queue')
}

function _QCP-NormalizePrependFolderPath {
    param([string]$FolderPath)
    $path = ([string]$FolderPath).Trim()
    if ($path.Length -eq 0) { return '' }
    if (Get-Command -Name 'Normalize-QCDocumentsFolderPath' -ErrorAction SilentlyContinue) {
        try {
            $norm = Normalize-QCDocumentsFolderPath -Path $path
            if ($norm.IsSuccess -and $norm.Data.path) { $path = [string]$norm.Data.path }
        } catch { }
    }
    return $path.TrimEnd('\', '/').ToLowerInvariant()
}

function _QCP-GetJobMetadataProcessType {
    param([object]$Job)
    if ($null -eq $Job) { return '' }
    $raw = ''
    try {
        if ($Job.metadata -and $Job.metadata.qcProcessType) { $raw = [string]$Job.metadata.qcProcessType }
    } catch { }
    if (_QCP-IsNullOrWhiteSpace $raw) { return '' }
    if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
        $norm = Normalize-QCProcessType -ProcessType $raw -AllowNullOnEmpty
        if ($norm) { return [string]$norm }
    }
    return $raw.Trim().ToLowerInvariant()
}

function _QCP-ResolveIntendedPrependProcessType {
    param(
        [hashtable]$Config,
        [string]$FolderPath,
        [string]$SheetPdfName,
        [string]$SheetPdfGuid = ''
    )
    $resolved = _QCP-ResolvePrependProcessTypeFromSourceAttributes -Config $Config `
        -FolderPath $FolderPath -SourceDocumentName $SheetPdfName
    if ($resolved.processType) { return [string]$resolved.processType }
    return ''
}

function _QCP-JobMatchesPrependSheet {
    param(
        [object]$Job,
        [string]$NormFolder,
        [string]$NormSheetPdfName,
        [string]$PrependTrigger = '',
        [string]$NormQcProcessType = ''
    )
    if ($null -eq $Job) { return $false }
    $jobType = ''
    try { $jobType = [string]$Job.type } catch { }
    if ($jobType -ne 'QC_PREPEND') { return $false }

    $triggerFilter = ([string]$PrependTrigger).Trim().ToLowerInvariant()
    if ($triggerFilter.Length -gt 0) {
        $jobTrigger = ''
        try {
            if ($Job.metadata -and $Job.metadata.prependTrigger) { $jobTrigger = [string]$Job.metadata.prependTrigger }
        } catch { }
        if ($jobTrigger.Length -gt 0 -and $jobTrigger.ToLowerInvariant() -ne $triggerFilter) { return $false }
    }

    $jobFolder = ''
    try {
        if ($Job.sourceFolder) { $jobFolder = [string]$Job.sourceFolder }
        elseif ($Job.sourcePath) { $jobFolder = [System.IO.Path]::GetDirectoryName([string]$Job.sourcePath) }
    } catch { }
    $jobFolder = _QCP-NormalizePrependFolderPath -FolderPath $jobFolder
    if ($jobFolder.Length -eq 0 -or $jobFolder -ne $NormFolder) { return $false }

    $jobName = ''
    try {
        if ($Job.sourceName) { $jobName = [string]$Job.sourceName }
        elseif ($Job.sourcePath) { $jobName = [System.IO.Path]::GetFileName([string]$Job.sourcePath) }
    } catch { }
    $jobName = ([string]$jobName).ToLowerInvariant()
    if ($jobName.Length -eq 0 -or $jobName -ne $NormSheetPdfName) { return $false }

    $laneFilter = ([string]$NormQcProcessType).Trim().ToLowerInvariant()
    if ($laneFilter.Length -eq 0) { return $true }
    $jobLane = _QCP-GetJobMetadataProcessType -Job $Job
    if ($jobLane.Length -eq 0) { return $false }
    return ($jobLane -eq $laneFilter)
}

function Test-QCPrependEnqueueBlockedForSheet {
    <#
    .SYNOPSIS
    Returns whether a QC_PREPEND for the same sheet PDF should not be enqueued.
    .DESCRIPTION
    Blocks when another initialQcPdf prepend is pending or running for the same folder + sheet PDF + lane,
    or when one succeeded very recently for the same lane (covers fallback enqueue racing a just-finished audit job).
    A different StateTransitionKey (for example audit:44209 vs audit:44147) bypasses the recent-succeeded guard
    so a deliberate new Initiate Origination cycle can overwrite an existing lane PDF.
  #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SheetPdfName,
        [string]$PrependTrigger = 'initialQcPdf',
        [string]$QcProcessType = '',
        [int]$SucceededWithinMinutes = 20,
        [string]$StateTransitionKey = ''
    )

    try { _QCP-EnsureQueueModulesLoaded } catch { return @{ blocked = $false } }

    $normFolder = _QCP-NormalizePrependFolderPath -FolderPath $FolderPath
    $normName = ([System.IO.Path]::GetFileName([string]$SheetPdfName)).ToLowerInvariant()
    if ($normFolder.Length -eq 0 -or $normName.Length -eq 0) { return @{ blocked = $false } }

    $normLane = ''
    if (-not (_QCP-IsNullOrWhiteSpace $QcProcessType) -and (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue)) {
        $normLane = Normalize-QCProcessType -ProcessType ([string]$QcProcessType) -AllowNullOnEmpty
    }
    if (_QCP-IsNullOrWhiteSpace $normLane) { $normLane = '' } else { $normLane = [string]$normLane }

    $root = _QCP-GetQueueRootFromConfig -Config $Config
    $states = @('pending', 'running', 'succeeded')
    $matches = @()
    $now = [DateTime]::UtcNow

    foreach ($state in $states) {
        $dir = Join-Path $root $state
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)
        foreach ($f in $files) {
            $job = $null
            try { $job = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction Stop | ConvertFrom-Json } catch { continue }
            if (-not (_QCP-JobMatchesPrependSheet -Job $job -NormFolder $normFolder -NormSheetPdfName $normName `
                    -PrependTrigger $PrependTrigger -NormQcProcessType $normLane)) {
                continue
            }
            $dedupeKey = ''
            $stKey = ''
            try { $dedupeKey = [string]$job.dedupeKey } catch { }
            try {
                if ($job.metadata -and $job.metadata.stateTransitionKey) {
                    $stKey = [string]$job.metadata.stateTransitionKey
                }
            } catch { }
            $entry = @{
                jobId = [string]$job.id
                state = $state
                dedupeKey = $dedupeKey
                stateTransitionKey = $stKey
            }
            if ($state -eq 'pending' -or $state -eq 'running') {
                return @{ blocked = $true; reason = ('queue_' + $state); matches = @($entry) }
            }
            if ($state -eq 'succeeded' -and $SucceededWithinMinutes -gt 0) {
                $updated = $null
                try {
                    if ($job.updatedAtUtc) { $updated = [DateTime]::Parse([string]$job.updatedAtUtc).ToUniversalTime() }
                    elseif ($job.completedAtUtc) { $updated = [DateTime]::Parse([string]$job.completedAtUtc).ToUniversalTime() }
                } catch { $updated = $null }
                if ($null -ne $updated) {
                    $age = $now - $updated
                    if ($age.TotalMinutes -le $SucceededWithinMinutes) {
                        $incomingKey = ([string]$StateTransitionKey).Trim()
                        $existingKey = ([string]$stKey).Trim()
                        if ($incomingKey.Length -gt 0 -and $existingKey.Length -gt 0 `
                                -and (-not $incomingKey.Equals($existingKey, [System.StringComparison]::OrdinalIgnoreCase))) {
                            $matches += $entry
                            continue
                        }
                        $entry['succeededMinutesAgo'] = [math]::Round($age.TotalMinutes, 2)
                        return @{ blocked = $true; reason = 'queue_succeeded_recent'; matches = @($entry) }
                    }
                }
            }
            $matches += $entry
        }
    }

    return @{ blocked = $false; matches = $matches }
}

function Add-QCPrependJobForQcInitiatedStateChange {
    <#
    .SYNOPSIS
    Enqueues QC_PREPEND when a non-automation actor sets workflow state to QC Initiated.
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
        [string]$ChangedByUsername = '',
        [string]$LastAuditEventAt = '',
        [Nullable[long]]$AuditEventId = $null,
        [string]$StateTransitionKey = ''
    )

    $initiatedName = 'QC Initiated'
    if (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        $wf = Get-QCWorkflowSettings -Config $Config
        $resolved = Get-QCWorkflowStateName -Settings $wf -StateKey 'qcInitiated'
        if (-not (_QCP-IsNullOrWhiteSpace $resolved)) { $initiatedName = [string]$resolved }
    }

    $curr = ([string]$CurrentStateName).Trim()
    if ($curr.Length -eq 0 -or $curr.ToLowerInvariant() -ne $initiatedName.ToLowerInvariant()) { return $null }

    if (Get-Command -Name 'Test-QCIsAutomationPwActor' -ErrorAction SilentlyContinue) {
        if (Test-QCIsAutomationPwActor -Config $Config -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername) {
            return $null
        }
    }

    $sheetPdf = _QCP-ResolveSheetPdfForPrependTrigger -TriggerDocumentName $TriggerDocumentName
    if (_QCP-IsNullOrWhiteSpace $sheetPdf) { return $null }

    if (-not (Get-Command -Name 'Test-QCPrependBlockedByMissingEmailAttributes' -ErrorAction SilentlyContinue)) {
        try { Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Notifications/QC.Notifications.psm1') -Force -ErrorAction SilentlyContinue } catch { }
    }
    if (Get-Command -Name 'Test-QCPrependBlockedByMissingEmailAttributes' -ErrorAction SilentlyContinue) {
        $emailGate = Test-QCPrependBlockedByMissingEmailAttributes -Config $Config -FolderPath $FolderPath `
            -SheetPdfName $sheetPdf -DocumentGuid $TriggerDocumentGuid
        if ($emailGate -and [bool]$emailGate.blocked) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Warning' -Code 'QC_PREPEND_SKIPPED_MISSING_EMAIL' `
                    -Message 'QC_PREPEND skipped: required notification email attributes are missing.' -Data @{
                    folderPath = $FolderPath; sheetPdf = $sheetPdf; triggerDocumentGuid = $TriggerDocumentGuid
                    missingFields = @($emailGate.missingFields); postPrependState = [string]$emailGate.postPrependState
                } | Out-Null
            }
            return New-QCSuccessResult -Code 'QC_PREPEND_SKIPPED_MISSING_EMAIL' `
                -Message 'QC_PREPEND skipped because required email attributes are missing for post-prepend notification.' -Data @{
                missingFields = @($emailGate.missingFields); sheetPdf = $sheetPdf; folderPath = $FolderPath
            }
        }
    }

    try { _QCP-EnsureQueueModulesLoaded } catch {
        return New-QCFailureResult -Code 'QC_PREPEND_QUEUE_UNAVAILABLE' -Message $_.Exception.Message -Data @{}
    }

    $sourcePath = Join-Path $FolderPath $sheetPdf
    $stateTransitionKey = $null
    if (-not (_QCP-IsNullOrWhiteSpace $StateTransitionKey)) {
        $stateTransitionKey = [string]$StateTransitionKey
    } elseif (Get-Command -Name 'Get-QCPrependStateTransitionDedupeKey' -ErrorAction SilentlyContinue) {
        $sheetStem = ''
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            $sheetStem = Get-PWSheetStemFromDocumentName -DocumentName $sheetPdf
        }
        $stateTransitionKey = Get-QCPrependStateTransitionDedupeKey -AuditEventId $AuditEventId `
            -LastAuditEventAt $LastAuditEventAt -ChangedByUser $ChangedByUser -TriggerDocumentGuid $TriggerDocumentGuid `
            -SheetStem $sheetStem -PreviousSheetState '' -TargetStateName $curr -PrependTrigger 'initialQcPdf'
    }
    if (_QCP-IsNullOrWhiteSpace $stateTransitionKey) {
        $stateTransitionKey = _QCP-GetQcInitiatedStateTransitionKey -AuditEventId $AuditEventId `
            -LastAuditEventAt $LastAuditEventAt -ChangedByUser $ChangedByUser -TriggerDocumentGuid $TriggerDocumentGuid
    }

    $intendedProcessType = ''
    if (Test-QCFastAuditEnqueueEnabled -Config $Config) {
        try { $intendedProcessType = [string](_QCP-ResolveProcessTypeFromSheetIndex -Config $Config -FolderPath $FolderPath -SourceDocumentName $sheetPdf) } catch { $intendedProcessType = '' }
    } else {
        $intendedProcessType = _QCP-ResolveIntendedPrependProcessType -Config $Config -FolderPath $FolderPath `
            -SheetPdfName $sheetPdf -SheetPdfGuid $TriggerDocumentGuid
    }

    $sheetBlock = Test-QCPrependEnqueueBlockedForSheet -Config $Config -FolderPath $FolderPath `
        -SheetPdfName $sheetPdf -PrependTrigger 'initialQcPdf' -QcProcessType $intendedProcessType `
        -StateTransitionKey ([string]$stateTransitionKey)
    if ($sheetBlock -and [bool]$sheetBlock.blocked) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_SKIPPED_SHEET_ACTIVE' `
                -Message 'QC_PREPEND skipped: another prepend for this sheet lane is pending, running, or recently succeeded.' -Data @{
                folderPath = $FolderPath; sheetPdf = $sheetPdf; qcProcessType = $intendedProcessType
                reason = [string]$sheetBlock.reason; matches = @($sheetBlock.matches); stateTransitionKey = $stateTransitionKey
            } | Out-Null
        }
        return New-QCSuccessResult -Code 'QC_PREPEND_SKIPPED_SHEET_ACTIVE' `
            -Message 'QC_PREPEND skipped because another prepend for this sheet is already active or recently completed.' -Data @{
            reason = [string]$sheetBlock.reason; matches = @($sheetBlock.matches); sheetPdf = $sheetPdf
            folderPath = $FolderPath; stateTransitionKey = $stateTransitionKey
        }
    }

    $candidate = @{
        path = $sourcePath
        fileName = $sheetPdf
        description = 'QC_Archivist'
        detectedAtUtc = (Get-QCTimestamp)
        sourceFolder = $FolderPath
        triggerSource = 'audit_state_change'
        auditActionName = 'DOCUMENT_STATE'
        file = @{
            fullName = $sourcePath
            length = 0
            lastWriteTimeUtc = (Get-QCTimestamp)
        }
    }
    if (-not (_QCP-IsNullOrWhiteSpace $stateTransitionKey)) {
        $candidate['stateTransitionKey'] = [string]$stateTransitionKey
    }

    $rule = @{
        id = 'qc-prepend-qc-initiated'
        jobType = 'QC_PREPEND'
        triggerType = 'audit_state_change'
        grouping = @{ enabled = $false; groupBy = 'file' }
    }

    $jobRes = New-QCJobObject -Candidate $candidate -Rule $rule -Config $Config
    if (-not $jobRes.IsSuccess) { return $jobRes }
    $job = [hashtable]$jobRes.Data.job
    if (-not $job.ContainsKey('metadata') -or -not $job.metadata) { $job['metadata'] = @{} }
    $md = _QCP-ToHashtable $job.metadata
    if (-not $md) { $md = @{} }
    $md['prependTrigger'] = 'initialQcPdf'
    $md['pwStateName'] = $curr
    $md['triggerDocumentGuid'] = $TriggerDocumentGuid
    $md['triggerDocumentName'] = $TriggerDocumentName
    if (-not (_QCP-IsNullOrWhiteSpace $stateTransitionKey)) { $md['stateTransitionKey'] = [string]$stateTransitionKey }
    if ($null -ne $ChangedByUser) { $md['changedByUser'] = $ChangedByUser }
    if (-not (_QCP-IsNullOrWhiteSpace $ChangedByUsername)) { $md['changedByUsername'] = [string]$ChangedByUsername }
    $job['metadata'] = $md
    if (Test-QCFastAuditEnqueueEnabled -Config $Config) {
        _QCP-AttachSheetIndexProcessTypeHint -Job $job -Config $Config -FolderPath $FolderPath -SheetPdfName $sheetPdf
    } else {
        $laneEnqueue = _QCP-TryResolvePrependLaneContext -Job $job -Config $Config
        if (-not $laneEnqueue.IsSuccess) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Warning' -Code 'QC_PREPEND_SKIPPED_NO_PROCESS_TYPE' `
                    -Message 'QC_PREPEND skipped: qc_process_type could not be resolved at enqueue time.' -Data @{
                    jobId = [string]$job.id; sourcePath = $sourcePath; folderPath = $FolderPath
                    triggerDocumentGuid = $TriggerDocumentGuid; currentState = $curr; prependTrigger = 'initialQcPdf'
                    errorCode = [string]$laneEnqueue.Code; errorMessage = [string]$laneEnqueue.Message
                } | Out-Null
            }
            return New-QCSuccessResult -Code 'QC_PREPEND_SKIPPED_NO_PROCESS_TYPE' `
                -Message 'QC_PREPEND skipped because qc_process_type could not be resolved.' -Data @{
                skipped = $true; sheetPdf = $sheetPdf; folderPath = $FolderPath; prependTrigger = 'initialQcPdf'
                errorCode = [string]$laneEnqueue.Code; errorMessage = [string]$laneEnqueue.Message
            }
        }
    }

    if ($DryRun) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_PLANNED' -Message 'Dry-run: QC_PREPEND job planned from QC Initiated state change.' -Data @{
                jobId = [string]$job.id; sourcePath = $sourcePath; folderPath = $FolderPath
                triggerDocumentGuid = $TriggerDocumentGuid; currentState = $curr
                stateTransitionKey = $stateTransitionKey
            } | Out-Null
        }
        return New-QCSuccessResult -Code 'QC_PREPEND_PLANNED' -Message 'Dry-run: QC_PREPEND job planned.' -Data @{ job = $job }
    }

    if (Get-Command -Name 'Test-QCDuplicateJob' -ErrorAction SilentlyContinue) {
        $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $Config
        if ($dupRes.IsSuccess -and [bool]$dupRes.Data.isDuplicate) {
            return New-QCSuccessResult -Code 'QC_PREPEND_SKIPPED_DUPLICATE' -Message 'QC_PREPEND already queued for this QC Initiated transition.' -Data @{
                dedupeKey = [string]$job['dedupeKey']; sourcePath = $sourcePath; stateTransitionKey = $stateTransitionKey
            }
        }
    }

    $enq = Add-QCQueueJob -Job $job -Config $Config
    if ($enq.IsSuccess) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_ENQUEUED' -Message 'QC_PREPEND job enqueued from QC Initiated state change.' -Data @{
                jobId = [string]$job.id; sourcePath = $sourcePath; folderPath = $FolderPath
                triggerDocumentGuid = $TriggerDocumentGuid; currentState = $curr
            } | Out-Null
        }
    }
    return $enq
}

function Add-QCPrependJobForQcFinalizingStateChange {
    <#
    .SYNOPSIS
    Enqueues QC_PREPEND when workflow state is Initiate Verification (correction-cycle prepend).
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
        [string]$ChangedByUsername = '',
        [string]$LastAuditEventAt = '',
        [Nullable[long]]$AuditEventId = $null,
        [string]$StateTransitionKey = ''
    )

    $finalizingName = 'QC Finalizing'
    if (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        $wf = Get-QCWorkflowSettings -Config $Config
        $resolved = Get-QCWorkflowStateName -Settings $wf -StateKey 'qcFinalizing'
        if (-not (_QCP-IsNullOrWhiteSpace $resolved)) { $finalizingName = [string]$resolved }
    }

    $curr = ([string]$CurrentStateName).Trim()
    if ($curr.Length -eq 0 -or $curr.ToLowerInvariant() -ne $finalizingName.ToLowerInvariant()) { return $null }

    $sheetPdf = _QCP-ResolveSheetPdfForPrependTrigger -TriggerDocumentName $TriggerDocumentName
    if (_QCP-IsNullOrWhiteSpace $sheetPdf) { return $null }

    try { _QCP-EnsureQueueModulesLoaded } catch {
        return New-QCFailureResult -Code 'QC_PREPEND_QUEUE_UNAVAILABLE' -Message $_.Exception.Message -Data @{}
    }

    $sourcePath = Join-Path $FolderPath $sheetPdf
    $stateTransitionKey = $null
    if (-not (_QCP-IsNullOrWhiteSpace $StateTransitionKey)) {
        $stateTransitionKey = [string]$StateTransitionKey
    } elseif (Get-Command -Name 'Get-QCPrependStateTransitionDedupeKey' -ErrorAction SilentlyContinue) {
        $sheetStem = ''
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            $sheetStem = Get-PWSheetStemFromDocumentName -DocumentName $sheetPdf
        }
        $stateTransitionKey = Get-QCPrependStateTransitionDedupeKey -AuditEventId $AuditEventId `
            -LastAuditEventAt $LastAuditEventAt -ChangedByUser $ChangedByUser -TriggerDocumentGuid $TriggerDocumentGuid `
            -SheetStem $sheetStem -PreviousSheetState '' -TargetStateName $curr
    }
    if (_QCP-IsNullOrWhiteSpace $stateTransitionKey) {
        $stateTransitionKey = _QCP-GetQcInitiatedStateTransitionKey -AuditEventId $AuditEventId `
            -LastAuditEventAt $LastAuditEventAt -ChangedByUser $ChangedByUser -TriggerDocumentGuid $TriggerDocumentGuid
    }
    $candidate = @{
        path = $sourcePath
        fileName = $sheetPdf
        description = 'QC_FinalPrepend'
        detectedAtUtc = (Get-QCTimestamp)
        sourceFolder = $FolderPath
        triggerSource = 'audit_state_change'
        auditActionName = 'DOCUMENT_STATE'
        file = @{
            fullName = $sourcePath
            length = 0
            lastWriteTimeUtc = (Get-QCTimestamp)
        }
    }
    if (-not (_QCP-IsNullOrWhiteSpace $stateTransitionKey)) {
        $candidate['stateTransitionKey'] = [string]$stateTransitionKey
    }

    $rule = @{
        id = 'qc-prepend-qc-finalizing'
        jobType = 'QC_PREPEND'
        triggerType = 'audit_state_change'
        grouping = @{ enabled = $false; groupBy = 'file' }
    }

    $jobRes = New-QCJobObject -Candidate $candidate -Rule $rule -Config $Config
    if (-not $jobRes.IsSuccess) { return $jobRes }
    $job = [hashtable]$jobRes.Data.job
    if (-not $job.ContainsKey('metadata') -or -not $job.metadata) { $job['metadata'] = @{} }
    $md = _QCP-ToHashtable $job.metadata
    if (-not $md) { $md = @{} }
    $md['prependTrigger'] = 'finalQcComplete'
    $md['finalPrepend'] = $true
    $md['pwStateName'] = $curr
    $md['triggerDocumentGuid'] = $TriggerDocumentGuid
    $md['triggerDocumentName'] = $TriggerDocumentName
    if (-not (_QCP-IsNullOrWhiteSpace $stateTransitionKey)) { $md['stateTransitionKey'] = [string]$stateTransitionKey }
    if ($null -ne $ChangedByUser) { $md['changedByUser'] = $ChangedByUser }
    if (-not (_QCP-IsNullOrWhiteSpace $ChangedByUsername)) { $md['changedByUsername'] = [string]$ChangedByUsername }
    $job['metadata'] = $md
    if (Test-QCFastAuditEnqueueEnabled -Config $Config) {
        _QCP-AttachSheetIndexProcessTypeHint -Job $job -Config $Config -FolderPath $FolderPath -SheetPdfName $sheetPdf
    } else {
        $laneEnqueue = _QCP-TryResolvePrependLaneContext -Job $job -Config $Config
        if (-not $laneEnqueue.IsSuccess) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Warning' -Code 'QC_PREPEND_SKIPPED_NO_PROCESS_TYPE' `
                    -Message 'QC_PREPEND skipped: qc_process_type could not be resolved at enqueue time.' -Data @{
                    jobId = [string]$job.id; sourcePath = $sourcePath; folderPath = $FolderPath
                    triggerDocumentGuid = $TriggerDocumentGuid; currentState = $curr; prependTrigger = 'finalQcComplete'
                    errorCode = [string]$laneEnqueue.Code; errorMessage = [string]$laneEnqueue.Message
                } | Out-Null
            }
            return New-QCSuccessResult -Code 'QC_PREPEND_SKIPPED_NO_PROCESS_TYPE' `
                -Message 'QC_PREPEND skipped because qc_process_type could not be resolved.' -Data @{
                skipped = $true; sheetPdf = $sheetPdf; folderPath = $FolderPath; prependTrigger = 'finalQcComplete'
                errorCode = [string]$laneEnqueue.Code; errorMessage = [string]$laneEnqueue.Message
            }
        }
    }

    if ($DryRun) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_PLANNED' -Message 'Dry-run: QC_PREPEND job planned from QC Finalizing state change.' -Data @{
                jobId = [string]$job.id; sourcePath = $sourcePath; folderPath = $FolderPath
                triggerDocumentGuid = $TriggerDocumentGuid; currentState = $curr; prependTrigger = 'finalQcComplete'
                stateTransitionKey = $stateTransitionKey
            } | Out-Null
        }
        return New-QCSuccessResult -Code 'QC_PREPEND_PLANNED' -Message 'Dry-run: QC_PREPEND job planned.' -Data @{ job = $job }
    }

    if (Get-Command -Name 'Test-QCDuplicateJob' -ErrorAction SilentlyContinue) {
        $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $Config
        if ($dupRes.IsSuccess -and [bool]$dupRes.Data.isDuplicate) {
            return New-QCSuccessResult -Code 'QC_PREPEND_SKIPPED_DUPLICATE' -Message 'QC_PREPEND already queued for this QC Finalizing transition.' -Data @{
                dedupeKey = [string]$job['dedupeKey']; sourcePath = $sourcePath; prependTrigger = 'finalQcComplete'
                stateTransitionKey = $stateTransitionKey
            }
        }
    }

    $enq = Add-QCQueueJob -Job $job -Config $Config
    if ($enq.IsSuccess) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_PREPEND_ENQUEUED' -Message 'QC_PREPEND job enqueued from QC Finalizing state change.' -Data @{
                jobId = [string]$job.id; sourcePath = $sourcePath; folderPath = $FolderPath
                triggerDocumentGuid = $TriggerDocumentGuid; currentState = $curr; prependTrigger = 'finalQcComplete'
            } | Out-Null
        }
    }
    return $enq
}

function Invoke-QCNotificationProcessor {
    <#
    .SYNOPSIS
    Sends a deferred QC notification job (decoupled from QC_PREPEND).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config
    )

    if (-not (Get-Command -Name 'Invoke-QCNotificationForStateChange' -ErrorAction SilentlyContinue)) {
        try { _QCP-EnsureNotificationModulesLoaded } catch { }
    }
    if (-not (Get-Command -Name 'Invoke-QCNotificationForStateChange' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'QC_NOTIFICATION_UNAVAILABLE' -Message 'QC.Notifications module not loaded.' -Data @{}
    }
    if (-not (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue)) {
        $discPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'ProjectWise/PW.Discovery.psm1'
        if (Test-Path -LiteralPath $discPath) {
            Import-Module $discPath -Force -WarningAction SilentlyContinue | Out-Null
        }
    }
    $meta = @{}
    if ($Job.ContainsKey('metadata') -and $Job.metadata -is [hashtable]) { $meta = $Job.metadata }
    $prev = if ($meta.ContainsKey('previousState')) { [string]$meta.previousState } else { '' }
    $curr = if ($meta.ContainsKey('currentState')) { [string]$meta.currentState } else { '' }
    $resolvedDocumentGuid = ''
    foreach ($guidKey in @('notificationLaneDocumentGuid', 'laneQcPdfDocumentGuid', 'documentGuid')) {
        if ($meta.ContainsKey($guidKey) -and -not (_QCP-IsNullOrWhiteSpace $meta[$guidKey])) {
            $resolvedDocumentGuid = [string]$meta[$guidKey]
            break
        }
    }
    $doc = $null
    if (-not (_QCP-IsNullOrWhiteSpace $resolvedDocumentGuid)) {
        $doc = @{ DocumentGUID = $resolvedDocumentGuid; Name = if ($Job.sourceName) { [string]$Job.sourceName } else { '' } }
    }
    if (Get-Command -Name 'Set-QCJobCheckpoint' -ErrorAction SilentlyContinue) {
        Set-QCJobCheckpoint -JobId ([string]$Job.id) -Config $Config -Job $Job -Checkpoint 'notification_send' | Out-Null
    }
    $stKey = $null
    if ($meta.ContainsKey('stateTransitionKey') -and $meta.stateTransitionKey) {
        $stKey = [string]$meta.stateTransitionKey
    } elseif ($Job.id) {
        $stKey = 'workflow:job:' + [string]$Job.id + '|state:' + $curr
    }
    $changedByUser = $null
    if ($meta.ContainsKey('changedByUser') -and $null -ne $meta.changedByUser) {
        try { $changedByUser = [int]$meta.changedByUser } catch { }
    }
    $changedByUsername = if ($meta.ContainsKey('changedByUsername')) { [string]$meta.changedByUsername } else { '' }
    $transitionId = $null
    if ($meta.ContainsKey('transitionId') -and $null -ne $meta.transitionId) {
        try { $transitionId = [int]$meta.transitionId } catch { }
    }
    $notifyParams = @{
        Config = $Config
        PreviousState = $prev
        CurrentState = $curr
        Document = $doc
        Job = $Job
        StateTransitionKey = $stKey
        ChangedByUser = $changedByUser
        ChangedByUsername = $changedByUsername
    }
    if (-not (_QCP-IsNullOrWhiteSpace $resolvedDocumentGuid)) {
        $notifyParams['DocumentGuid'] = $resolvedDocumentGuid
    }
    if (-not (_QCP-IsNullOrWhiteSpace $Job.sourceName)) {
        $notifyParams['DocumentName'] = [string]$Job.sourceName
    }
    if (-not (_QCP-IsNullOrWhiteSpace $Job.sourceFolder)) {
        $notifyParams['DocumentPath'] = Join-Path ([string]$Job.sourceFolder) $(if ($Job.sourceName) { [string]$Job.sourceName } else { '' })
    }
    if ($null -ne $transitionId -and $transitionId -gt 0) {
        $notifyParams['TransitionId'] = $transitionId
    }
    $res = $null
    try {
        $res = Invoke-QCNotificationForStateChange @notifyParams
    } catch {
        $err = [string]$_.Exception.Message
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_NOTIFICATION_PROCESSOR_THROW' -Message $err -Data @{
                jobId = if ($Job.id) { [string]$Job.id } else { '' }
                documentName = if ($notifyParams.DocumentName) { [string]$notifyParams.DocumentName } else { '' }
                currentState = $curr
                previousState = $prev
            } | Out-Null
        }
        return New-QCFailureResult -Code 'PROCESSOR_HANDLER_THROW' -Message "Handler threw: Invoke-QCNotificationProcessor" `
            -Data @{ jobId = [string]$Job.id; jobType = 'QC_NOTIFICATION'; handler = 'Invoke-QCNotificationProcessor'; error = $err }
    }
    if (Get-Command -Name 'Set-QCJobCheckpoint' -ErrorAction SilentlyContinue) {
        $cp = if ($res -and $res.IsSuccess) { 'notification_complete' } else { 'notification_failed' }
        Set-QCJobCheckpoint -JobId ([string]$Job.id) -Config $Config -Job $Job -Checkpoint $cp | Out-Null
    }
    if ($null -eq $res) {
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED' -Message 'Notification skipped or unavailable.' -Data @{}
    }
    if ($res.IsSuccess) {
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_JOB_OK' -Message 'Notification job completed.' -Data @{ notification = $res.Data }
    }
    if ($res.Code -eq 'QC_NOTIFICATION_SKIPPED_NO_RECIPIENTS') {
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_NO_RECIPIENTS' -Message ([string]$res.Message) -Data @{ notification = $res.Data; skipped = $true }
    }
    return New-QCFailureResult -Code 'QC_NOTIFICATION_JOB_FAILED' -Message ([string]$res.Message) -Data @{ notificationCode = [string]$res.Code }
}

function Invoke-QCReportingScanProcessor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $mod = Join-Path (Split-Path -Parent $PSScriptRoot) 'Reporting\QC.Reporting.psm1'
    if (-not (Test-Path -LiteralPath $mod)) {
        return New-QCFailureResult -Code 'QC_REPORTING_MODULE_MISSING' -Message "QC.Reporting.psm1 not found: $mod" -Data @{ path = $mod }
    }
    Import-Module $mod -Force
    return Invoke-QCReportingScan -Job $Job -Config $Config
}

Export-ModuleMember -Function *
