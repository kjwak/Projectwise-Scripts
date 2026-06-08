# QC.AuditTriggers.psm1
# Responsibility: Audit-driven QC workflow state/attribute triggers (telemetry + notifications).

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
if (-not (Get-Command -Name 'Write-QCDocumentStateHistoryRow' -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $PSScriptRoot 'Core.Database.psm1') -Force
}
# Do not import QC.Notifications here: it loads PW.Discovery which re-imports this module and can
# break exported cmdlets during circular load. Notifications are lazy-loaded in Invoke-QCAuditWorkflowStateChangeTriggers.

function _QCAT-ToHashtable([object]$Value) {
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

function _QCAT-NormalizeValue([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return ([string]$Value).Trim()
}

function Get-QCAuditStateTransitionKey {
    <#
    .SYNOPSIS
    Stable id for one workflow state transition (audit event, DB transition row, or user+time).
    Used by notification dedupe and QC_PREPEND enqueue so repeat cycles are not blocked by prior sends.
    #>
    [CmdletBinding()]
    param(
        [Nullable[long]]$AuditEventId = $null,
        [string]$LastAuditEventAt = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$TriggerDocumentGuid = '',
        [Nullable[int]]$TransitionId = $null
    )
    if ($null -ne $AuditEventId -and $AuditEventId -gt 0) {
        return ('audit:' + [string]$AuditEventId)
    }
    if ($null -ne $TransitionId -and $TransitionId -gt 0) {
        return ('transition:' + [string]$TransitionId)
    }
    $at = ([string]$LastAuditEventAt).Trim()
    $guid = ([string]$TriggerDocumentGuid).Trim()
    if ($at.Length -eq 0) { return $null }
    $userPart = ''
    if ($null -ne $ChangedByUser -and $ChangedByUser -gt 0) { $userPart = ('user:' + [string]$ChangedByUser + '|') }
    return ($userPart + 'at:' + $at + '|doc:' + $guid)
}

function Get-QCPrependStateTransitionDedupeKey {
    <#
    .SYNOPSIS
    Dedupe key fragment for QC_PREPEND jobs tied to a workflow state transition.
    .DESCRIPTION
    Prefers the audit event id so one human DOCUMENT_STATE produces one prepend job (sibling echo
    syncs reuse the same audit id in dedupe). A new audit event on a later QC cycle gets a new key
    even when sheet|from|to matches a prior succeeded job. Falls back to sheet|from|to when no audit id.
    #>
    [CmdletBinding()]
    param(
        [Nullable[long]]$AuditEventId = $null,
        [string]$LastAuditEventAt = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$TriggerDocumentGuid = '',
        [string]$SheetStem = '',
        [string]$PreviousSheetState = '',
        [string]$TargetStateName = '',
        [string]$PrependTrigger = ''
    )

    if ($null -ne $AuditEventId -and $AuditEventId -gt 0) {
        $auditKey = Get-QCAuditStateTransitionKey -AuditEventId $AuditEventId -LastAuditEventAt $LastAuditEventAt `
            -ChangedByUser $ChangedByUser -TriggerDocumentGuid $TriggerDocumentGuid
        if (-not [string]::IsNullOrWhiteSpace($auditKey)) { return [string]$auditKey }
    }

    $stem = ([string]$SheetStem).Trim().ToLowerInvariant()
    $from = ([string]$PreviousSheetState).Trim().ToLowerInvariant()
    $to = ([string]$TargetStateName).Trim().ToLowerInvariant()
    if ($stem.Length -eq 0 -or $to.Length -eq 0) { return $null }
    $triggerPart = ''
    $pt = ([string]$PrependTrigger).Trim().ToLowerInvariant()
    if ($pt.Length -gt 0) { $triggerPart = ('trigger:' + $pt + '|') }
    return ('sheet:' + $stem + '|' + $triggerPart + 'from:' + $from + '|to:' + $to)
}

function Get-QCSheetGroupTransitionKey {
    <#
    .SYNOPSIS
    Stable logical transition key for one sheet-group member (idempotency across siblings and retries).
    #>
    [CmdletBinding()]
    param(
        [string]$SheetStem = '',
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$TargetState,
        [string]$TransitionSource = '',
        [Nullable[long]]$AuditEventId = $null,
        [string]$JobId = ''
    )
    $stem = ([string]$SheetStem).Trim().ToLowerInvariant()
    $guid = ([string]$DocumentGuid).Trim().ToLowerInvariant()
    $to = ([string]$TargetState).Trim().ToLowerInvariant()
    $src = ([string]$TransitionSource).Trim().ToLowerInvariant()
    if ($guid.Length -eq 0 -or $to.Length -eq 0) { return $null }
    $anchor = ''
    if ($null -ne $AuditEventId -and $AuditEventId -gt 0) {
        $anchor = 'audit:' + [string]$AuditEventId
    } elseif (-not [string]::IsNullOrWhiteSpace($JobId)) {
        $anchor = 'job:' + ([string]$JobId).Trim().ToLowerInvariant()
    }
    if ($anchor.Length -eq 0) { return $null }
    if ($stem.Length -eq 0) { $stem = '_' }
    return ('sg|' + $stem + '|' + $guid + '|' + $to + '|' + $anchor + '|' + $src)
}

function _QCAT-ResolveMemberFileRole {
    param([string]$DocumentName)
    $dn = [string]$DocumentName
    if ($dn -match '(?i)-qc\.pdf$') { return 'qcPdf' }
    if ($dn -match '(?i)\.dgn$') { return 'dgn' }
    if ($dn -match '(?i)\.pdf$') { return 'pdf' }
    return 'other'
}

function _QCAT-ResolveSheetPackageNotificationMember {
    param([array]$Members)
    $qcPdf = $null
    $sheetPdf = $null
    foreach ($member in @($Members)) {
        $dn = [string]$member.documentName
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }
        if (Test-QCIsQcPdfDocumentName -DocumentName $dn) { $qcPdf = $member; break }
        if (($dn -match '(?i)\.pdf$') -and ($dn -notmatch '(?i)-qc\.pdf$') -and -not $sheetPdf) { $sheetPdf = $member }
    }
    if ($qcPdf) { return $qcPdf }
    if ($sheetPdf) { return $sheetPdf }
    if (@($Members).Count -gt 0) { return $Members[0] }
    return $null
}

function _QCAT-ResolveSheetPackageSheetPdfMember {
    param([array]$Members)
    foreach ($member in @($Members)) {
        $dn = [string]$member.documentName
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }
        if (($dn -match '(?i)\.pdf$') -and ($dn -notmatch '(?i)-qc\.pdf$')) { return $member }
    }
    return $null
}

function _QCAT-ResolveSheetPackageDgnMember {
    param([array]$Members)
    foreach ($member in @($Members)) {
        $dn = [string]$member.documentName
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }
        if ($dn -match '(?i)\.dgn$') { return $member }
    }
    return $null
}

function _QCAT-ResolveCanonicalSheetPackageIdentity {
    <#
    Resolves the canonical DGN identity used for qc_cycle_completions and sheet_index rollups.
    #>
    param(
        [hashtable]$Config = @{},
        [string]$DocumentGuid = '',
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$SheetStem = '',
        [array]$Members = @()
    )

    $dgnMember = _QCAT-ResolveSheetPackageDgnMember -Members $Members
    if ($dgnMember -and -not [string]::IsNullOrWhiteSpace([string]$dgnMember.documentGuid)) {
        return @{
            documentGuid = [string]$dgnMember.documentGuid
            documentName = [string]$dgnMember.documentName
            document = $dgnMember.document
            resolved = $true
            resolution = 'sheet_package_dgn_member'
        }
    }

    $dn = ([string]$DocumentName).Trim()
    $dg = ([string]$DocumentGuid).Trim()
    if (($dn -match '(?i)\.dgn$') -and -not [string]::IsNullOrWhiteSpace($dg)) {
        return @{ documentGuid = $dg; documentName = $dn; document = $null; resolved = $true; resolution = 'dgn_input' }
    }

    $stem = if (-not [string]::IsNullOrWhiteSpace($SheetStem)) { [string]$SheetStem.Trim() } else { _QCAT-ResolveSheetStemFromDocumentName -DocumentName $dn }
    $dgnName = if ($stem -match '(?i)\.dgn$') { $stem } else { $stem + '.dgn' }
    $folder = ([string]$FolderPath).Trim()

    if ($folder.Length -gt 0 -and $Config -and (Get-Command -Name 'Invoke-QCDatabaseQuery' -ErrorAction SilentlyContinue)) {
        try {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_guid, document_name
FROM sheet_index
WHERE folder_path = @folderPath
  AND LOWER(document_name) = LOWER(@dgnName)
"@ -Parameters @{ folderPath = $folder; dgnName = $dgnName }
            if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
                $row = $res.Data.table.Rows[0]
                return @{
                    documentGuid = [string]$row.document_guid
                    documentName = [string]$row.document_name
                    document = $null
                    resolved = $true
                    resolution = 'sheet_index_dgn_lookup'
                }
            }
        } catch { }
    }

    return @{
        documentGuid = $dg
        documentName = $dgnName
        document = $null
        resolved = $false
        resolution = 'canonical_dgn_not_found'
    }
}

function _QCAT-GetQcCompleteStateName {
    param([hashtable]$Config = @{})
    $completeState = 'QC Complete'
    if (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            if ($wf -and $wf.states -and $wf.states.complete) { $completeState = [string]$wf.states.complete }
        } catch { }
    }
    return _QCAT-NormalizeValue $completeState
}

function _QCAT-GetPackageQcCompleteTransitionFromMemberResults {
    <#
    Returns package-level QC Complete transition evidence from per-member telemetry.
    Counts a completion when any member transitioned from a non-QC Complete state into QC Complete.
    #>
    param(
        [array]$MemberResults = @(),
        [Parameter(Mandatory)][string]$TargetState,
        [hashtable]$Config = @{}
    )

    $complete = _QCAT-GetQcCompleteStateName -Config $Config
    $target = _QCAT-NormalizeValue $TargetState
    if ($target -ne $complete) { return $null }

    foreach ($mr in @($MemberResults)) {
        if ([bool]$mr.skipped) { continue }
        if (-not ([bool]$mr.transitionInserted -or [bool]$mr.transitionReused)) { continue }
        $mPrev = _QCAT-NormalizeValue ([string]$mr.previousState)
        $mFinal = _QCAT-NormalizeValue ([string]$mr.finalState)
        if ($mPrev -ne $complete -and $mFinal -eq $complete) {
            return @{
                previousState = $mPrev
                currentState = $mFinal
                memberDocumentGuid = [string]$mr.documentGuid
                memberRole = [string]$mr.role
                evidenceSource = 'member_results'
            }
        }
    }
    return $null
}

function _QCAT-GetPackageQcCompleteTransitionFromSheetStateSync {
    <#
    Fallback for prepend writeback: sibling sync records pre-write fromState even when sheet_index
    was already updated to QC Complete before per-member telemetry runs.
    #>
    param(
        [hashtable]$Context = $null,
        [Parameter(Mandatory)][string]$TargetState,
        [hashtable]$Config = @{}
    )

    if (-not $Context -or -not $Context.ContainsKey('sheetStateSync')) { return $null }

    $complete = _QCAT-GetQcCompleteStateName -Config $Config
    $target = _QCAT-NormalizeValue $TargetState
    if ($target -ne $complete) { return $null }

    $sync = _QCAT-ToHashtable $Context.sheetStateSync
    if (-not $sync -or -not $sync.ContainsKey('updates')) { return $null }

    foreach ($upd in @($sync.updates)) {
        $u = _QCAT-ToHashtable $upd
        if (-not $u) { continue }
        $from = _QCAT-NormalizeValue ([string]$u.fromState)
        $to = if ($u.ContainsKey('toState') -and $u.toState) { _QCAT-NormalizeValue ([string]$u.toState) } else { $target }
        if ($from -ne $complete -and $to -eq $complete) {
            return @{
                previousState = $from
                currentState = $to
                memberDocumentGuid = [string]$u.documentGuid
                memberRole = ''
                evidenceSource = 'sheet_state_sync'
            }
        }
    }
    return $null
}

function _QCAT-ResolveQCCycleIdForCompletion {
    param(
        [hashtable]$Config = @{},
        [hashtable]$Context = $null,
        [string]$DocumentGuid = '',
        [string]$FolderPath = '',
        [string]$SheetStem = ''
    )

    if (Get-Command -Name 'Get-QCSheetIndexCycle' -ErrorAction SilentlyContinue) {
        try {
            $cycle = Get-QCSheetIndexCycle -Config $Config -DocumentGuid $DocumentGuid -FolderPath $FolderPath -SheetStem $SheetStem
            if ($cycle -and -not [string]::IsNullOrWhiteSpace([string]$cycle.cycleId)) {
                $cycleNumber = $null
                if ($null -ne $cycle.cycleNumber -and -not [string]::IsNullOrWhiteSpace([string]$cycle.cycleNumber)) {
                    try { $cycleNumber = [int][string]$cycle.cycleNumber } catch { $cycleNumber = $null }
                }
                return @{
                    cycleId = [string]$cycle.cycleId.Trim()
                    cycleNumber = $cycleNumber
                    source = 'sheet_index'
                }
            }
        } catch { }
    }

    $sources = [System.Collections.Generic.List[object]]::new()
    if ($Context) {
        $sources.Add($Context) | Out-Null
        if ($Context.ContainsKey('attributes') -and $Context.attributes) {
            $sources.Add((_QCAT-ToHashtable $Context.attributes)) | Out-Null
        }
        if ($Context.ContainsKey('job') -and $Context.job) {
            $job = _QCAT-ToHashtable $Context.job
            if ($job) {
                $sources.Add($job) | Out-Null
                if ($job.ContainsKey('metadata') -and $job.metadata) {
                    $md = _QCAT-ToHashtable $job.metadata
                    if ($md) {
                        $sources.Add($md) | Out-Null
                        if ($md.ContainsKey('attributes') -and $md.attributes) {
                            $sources.Add((_QCAT-ToHashtable $md.attributes)) | Out-Null
                        }
                    }
                }
            }
        }
    }

    foreach ($source in @($sources)) {
        if (-not $source) { continue }
        $src = _QCAT-ToHashtable $source
        if (-not $src) { continue }
        foreach ($key in @('cycleId', 'qcCycleId', 'QC_Cycle_ID')) {
            if (-not $src.ContainsKey($key)) { continue }
            $value = [string]$src[$key]
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            if ($value -match '^(audit:|transition:)') { continue }
            $cycleNumber = $null
            foreach ($numKey in @('cycleNumber', 'qcCycleNumber', 'QC_Cycle_Number')) {
                if ($src.ContainsKey($numKey) -and -not [string]::IsNullOrWhiteSpace([string]$src[$numKey])) {
                    try { $cycleNumber = [int][string]$src[$numKey] } catch { $cycleNumber = $null }
                    break
                }
            }
            return @{
                cycleId = $value.Trim()
                cycleNumber = $cycleNumber
                source = 'context'
            }
        }
    }
    return $null
}

function _QCAT-TryRecordQCCycleCompletion {
    <#
    Records a durable QC cycle completion when workflow transitions into QC Complete.
    Idempotent via Ensure-QCCycleCompletion. Caller must pass a genuine non-complete -> QC Complete transition.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [AllowEmptyString()][string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState,
        [string]$TransitionSource = '',
        [Parameter(Mandatory)][string]$DocumentGuid,
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$SheetStem = '',
        [Nullable[long]]$AuditEventId = $null,
        [Nullable[int]]$TransitionEventId = $null,
        [string]$ChangedByUsername = '',
        [hashtable]$Context = $null,
        [hashtable]$PwAttributes = $null,
        [object]$Document = $null,
        [array]$Members = @(),
        [bool]$DryRun = $false
    )

    $completeState = 'QC Complete'
    if (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            if ($wf -and $wf.states -and $wf.states.complete) { $completeState = [string]$wf.states.complete }
        } catch { }
    }

    $prev = _QCAT-NormalizeValue $PreviousState
    $curr = _QCAT-NormalizeValue $CurrentState
    if ([string]::IsNullOrWhiteSpace($prev) -or $prev -eq $completeState -or $curr -ne $completeState) { return }

    $source = ([string]$TransitionSource).Trim()

    if (-not (Get-Command -Name 'Ensure-QCCycleCompletion' -ErrorAction SilentlyContinue)) { return }

    $canonical = _QCAT-ResolveCanonicalSheetPackageIdentity -Config $Config -DocumentGuid $DocumentGuid `
        -DocumentName $DocumentName -FolderPath $FolderPath -SheetStem $SheetStem -Members $Members
    if (-not [bool]$canonical.resolved) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Warning' -Code 'QC_CYCLE_COMPLETION_SKIPPED' `
                -Message 'Skipped QC cycle completion: canonical DGN identity unresolved.' -Data @{
                documentGuid = $DocumentGuid; documentName = $DocumentName; folderPath = $FolderPath
                sheetStem = $SheetStem; transitionSource = $source; reason = 'canonical_dgn_not_found'
                resolution = $canonical.resolution
            } | Out-Null
        }
        return
    }

    $docGuid = [string]$canonical.documentGuid
    $docName = [string]$canonical.documentName
    $docObject = if ($canonical.document) { $canonical.document } else { $Document }

    $sheetPackageId = $null
    if (Get-Command -Name 'Resolve-QCCycleCompletionSheetPackageId' -ErrorAction SilentlyContinue) {
        try {
            $sheetPackageId = Resolve-QCCycleCompletionSheetPackageId -Config $Config -DocumentGuid $docGuid
        } catch { }
    } elseif (Get-Command -Name 'Get-SheetPackageIdForDocument' -ErrorAction SilentlyContinue) {
        try { $sheetPackageId = Get-SheetPackageIdForDocument -Config $Config -DocumentGuid $docGuid } catch { }
    }
    if (-not $sheetPackageId) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Warning' -Code 'QC_CYCLE_COMPLETION_SKIPPED' `
                -Message 'Skipped QC cycle completion: sheet_package_id unresolved.' -Data @{
                documentGuid = $docGuid; documentName = $docName; folderPath = $FolderPath
                sheetStem = $SheetStem; transitionSource = $source; reason = 'sheet_package_not_found'
            } | Out-Null
        }
        return
    }

    $resolvedCycle = _QCAT-ResolveQCCycleIdForCompletion -Config $Config -Context $Context `
        -DocumentGuid $docGuid -FolderPath $FolderPath -SheetStem $SheetStem
    $cycleId = if ($resolvedCycle) { [string]$resolvedCycle.cycleId } else { '' }
    $cycleNumber = if ($resolvedCycle) { $resolvedCycle.cycleNumber } else { $null }
    if ([string]::IsNullOrWhiteSpace($cycleId)) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Warning' -Code 'QC_CYCLE_COMPLETION_SKIPPED' `
                -Message 'Skipped QC cycle completion: qc_cycle_id unavailable.' -Data @{
                documentGuid = $docGuid; documentName = $docName; transitionSource = $source
            } | Out-Null
        }
        return
    }

    $reviewType = ''
    $reviewSourceMember = _QCAT-ResolveSheetPackageSheetPdfMember -Members $Members
    $reviewDocGuid = if ($reviewSourceMember) { [string]$reviewSourceMember.documentGuid } else { $docGuid }
    $reviewDocName = if ($reviewSourceMember) { [string]$reviewSourceMember.documentName } else { $docName }
    $reviewDocObject = if ($reviewSourceMember -and $reviewSourceMember.document) { $reviewSourceMember.document } else { $docObject }
    if (Get-Command -Name 'Resolve-QCWorkflowEventQcReviewType' -ErrorAction SilentlyContinue) {
        try {
            $reviewType = Resolve-QCWorkflowEventQcReviewType -Config $Config -DocumentGuid $reviewDocGuid `
                -FolderPath $FolderPath -DocumentName $reviewDocName -Context $Context -PwAttributes $PwAttributes -Document $reviewDocObject
        } catch { }
    }

    $normalizedReviewType = $null
    if (-not [string]::IsNullOrWhiteSpace($reviewType) -and (Get-Command -Name 'Get-QCReviewTypeBucket' -ErrorAction SilentlyContinue)) {
        $normalizedReviewType = Get-QCReviewTypeBucket -ReviewType $reviewType
    }
    if (-not $normalizedReviewType) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Warning' -Code 'QC_CYCLE_COMPLETION_SKIPPED' `
                -Message 'Skipped QC cycle completion: unrecognized review type.' -Data @{
                documentGuid = $docGuid; reviewType = $reviewType; transitionSource = $source
            } | Out-Null
        }
        return
    }

    if ($DryRun) { return }

    $completedBy = if ($ChangedByUsername) { [string]$ChangedByUsername.Trim() } else { '' }
    $result = Ensure-QCCycleCompletion -Config $Config -DocumentGuid $docGuid -DocumentName $docName `
        -SheetPackageId $sheetPackageId -QcCycleId $cycleId -QcCycleNumber $cycleNumber -QcReviewType $normalizedReviewType `
        -TransitionEventId $TransitionEventId -AuditEventId $AuditEventId -CompletedBy $completedBy

    $inserted = $false
    $reused = $false
    $completedAt = $null
    if ($result.IsSuccess -and $result.Data) {
        try { $inserted = [bool]$result.Data.inserted } catch { }
        try { $reused = [bool]$result.Data.reused } catch { }
        try {
            if ($null -ne $result.Data.completedAt) { $completedAt = [datetime]$result.Data.completedAt }
        } catch { }
    }
    if ($inserted -and (Get-Command -Name 'Update-QCSheetCycleCompletionSummary' -ErrorAction SilentlyContinue)) {
        try { Update-QCSheetCycleCompletionSummary -Config $Config -SheetPackageId $sheetPackageId | Out-Null } catch { }
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        $logData = @{
            documentGuid = $docGuid
            sheetPackageId = $sheetPackageId.ToString()
            cycleId = $cycleId
            cycleNumber = $cycleNumber
            reviewType = $reviewType
            normalizedReviewType = $normalizedReviewType
            completedAt = if ($completedAt) { $completedAt.ToString('o') } else { [datetime]::UtcNow.ToString('o') }
            inserted = $inserted
            transitionSource = $source
            auditEventId = $AuditEventId
            transitionEventId = $TransitionEventId
        }
        if ($inserted) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_CYCLE_COMPLETION_RECORDED' `
                -Message 'QC cycle completion recorded.' -Data $logData | Out-Null
        } elseif ($reused) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_CYCLE_COMPLETION_DUPLICATE' `
                -Message 'QC cycle completion duplicate suppressed.' -Data $logData | Out-Null
        }
    }
}

function _QCAT-WriteSheetGroupMemberWorkflowEvents {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Settings,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [AllowEmptyString()][string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState,
        [string]$TransitionSource = '',
        [string]$AuditActionName = 'DOCUMENT_STATE',
        [Nullable[long]]$AuditEventId = $null,
        [string]$JobId = '',
        [string]$JobType = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [object]$Document = $null,
        [hashtable]$Context = $null,
        [bool]$DryRun = $false,
        [Nullable[guid]]$SheetPackageId = $null,
        [Nullable[guid]]$TransitionGroupId = $null
    )

    $prev = _QCAT-NormalizeValue $PreviousState
    $curr = _QCAT-NormalizeValue $CurrentState
    $result = @{
        documentGuid = $DocumentGuid
        documentName = $DocumentName
        previousState = $prev
        currentState = $curr
        historyInserted = $false
        transitionInserted = $false
        transitionReused = $false
        mirrorInserted = $false
        skipped = $false
        skipReason = ''
        transitionId = $null
    }

    if ($prev -eq $curr) {
        $result.skipped = $true
        $result.skipReason = 'no_state_change'
        return $result
    }

    if (Test-QCShouldSuppressBaselineSheetIndexStateTransition -Config $Config -PreviousState $prev -CurrentState $curr) {
        $result.skipped = $true
        $result.skipReason = 'baseline_seed_suppressed'
        return $result
    }

    if ([bool]$Settings.recordStateHistory) {
        $hist = Write-QCDocumentStateHistoryRow -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -EventType 'STATE_CHANGE' -OldValue $prev -NewValue $curr `
            -FieldName 'pw_state_name' -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -SourceAuditId $AuditEventId -SheetPackageId $SheetPackageId -TransitionGroupId $TransitionGroupId
        try {
            if ($hist.IsSuccess -and $hist.Data -and $hist.Data.written -eq $true) { $result.historyInserted = $true }
        } catch { }
    }

    if ([bool]$Settings.recordTransitions) {
        $jobTypeValue = if ($JobType) { $JobType } elseif ($TransitionSource -eq 'automation_prepend_completion') { 'QC_PREPEND' } else { 'audit_trigger' }
        $tr = Ensure-QCTransitionEvent -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -TransitionType 'STATE_CHANGE' -FromValue $prev -ToValue $curr `
            -JobId $JobId -JobType $jobTypeValue -TriggerAuditId $AuditEventId `
            -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -SheetPackageId $SheetPackageId -TransitionGroupId $TransitionGroupId
        if ($tr.IsSuccess -and $tr.Data) {
            try {
                if ($null -ne $tr.Data.transitionId) { $result.transitionId = [int]$tr.Data.transitionId }
            } catch { }
            try {
                if ($tr.Data.written -eq $true) { $result.transitionInserted = $true }
                if ($tr.Data.reused -eq $true) { $result.transitionReused = $true }
            } catch { }
        }
        if ($result.transitionInserted -or $result.transitionReused) {
            _QCAT-WriteWorkflowEventMirror -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
                -FolderPath $FolderPath -FromValue $prev -ToValue $curr -JobId $JobId -JobType $jobTypeValue `
                -EventType 'STATE_CHANGE' -AuditActionName $AuditActionName -Context $Context -Document $Document `
                -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
                -TransitionEventId $result.transitionId -SheetPackageId $SheetPackageId -TransitionGroupId $TransitionGroupId
            $result.mirrorInserted = $true
        }
    }

    return $result
}

function Invoke-QCSheetGroupWorkflowTransition {
    <#
    .SYNOPSIS
    Records workflow telemetry and optional notifications for every resolved sheet-group sibling.
    .DESCRIPTION
    Pairs workflow state synchronization with lifecycle event recording across DGN, sheet PDF, and QC PDF.
    The triggering document may be any sibling; notifications emit once per logical sheet-group transition.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$TriggerDocumentGuid,
        [Parameter(Mandatory)][string]$TriggerDocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$SourceState = '',
        [Parameter(Mandatory)][string]$TargetState,
        [Parameter(Mandatory)][string]$TransitionSource,
        [array]$Members = @(),
        [hashtable]$StateByGuid = $null,
        [hashtable]$PreviousStateByGuid = $null,
        [Nullable[long]]$AuditEventId = $null,
        [string]$JobId = '',
        [string]$JobType = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [string]$LastAuditEventAt = '',
        [bool]$DryRun = $false,
        [switch]$SuppressNotification,
        [hashtable]$Context = $null,
        [string]$AuditActionName = 'DOCUMENT_STATE'
    )

    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        return @{ skipped = $true; skipReason = 'workflow_triggers_disabled'; members = @() }
    }

    $target = _QCAT-NormalizeValue $TargetState
    if ([string]::IsNullOrWhiteSpace($target)) {
        return @{ skipped = $true; skipReason = 'empty_target_state'; members = @() }
    }

    $triggerRole = _QCAT-ResolveMemberFileRole -DocumentName $TriggerDocumentName
    $sheetStem = _QCAT-ResolveSheetStemFromDocumentName -DocumentName $TriggerDocumentName

    if ($Context -and $Context.ContainsKey('sheetStateSync')) {
        $syncForMembers = _QCAT-ToHashtable $Context.sheetStateSync
        if ($syncForMembers -and $syncForMembers.ContainsKey('updates') -and @($syncForMembers.updates).Count -gt 0) {
            $syncMembers = [System.Collections.Generic.List[object]]::new()
            foreach ($upd in @($syncForMembers.updates)) {
                $u = _QCAT-ToHashtable $upd
                if (-not $u) { continue }
                $ug = [string]$u.documentGuid
                $un = [string]$u.documentName
                if ([string]::IsNullOrWhiteSpace($ug) -and [string]::IsNullOrWhiteSpace($un)) { continue }
                [void]$syncMembers.Add(@{
                    documentGuid = $ug
                    documentName = $un
                    document = $null
                })
            }
            if ($syncMembers.Count -gt @($Members).Count) {
                $Members = @($syncMembers)
            }
        }
    }

    if (@($Members).Count -eq 0 -and (Get-Command -Name 'Get-PWAssociatedSheetMembers' -ErrorAction SilentlyContinue)) {
        try {
            $Members = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $FolderPath `
                -DocumentName $TriggerDocumentName -DocumentGuid $TriggerDocumentGuid)
        } catch { $Members = @() }
    }

    if (@($Members).Count -eq 0) {
        if (-not [string]::IsNullOrWhiteSpace($TriggerDocumentGuid)) {
            $Members = @(@{
                documentGuid = $TriggerDocumentGuid
                documentName = $TriggerDocumentName
                document = $null
            })
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Information' -Code 'SHEET_GROUP_TRANSITION_TRIGGER_ONLY' `
                    -Message 'No associated members resolved; recording transition for trigger document only.' -Data @{
                    triggerDocumentGuid = $TriggerDocumentGuid
                    triggerDocumentName = $TriggerDocumentName
                    triggerFileType = $triggerRole
                    folderPath = $FolderPath
                    targetState = $target
                    transitionSource = $TransitionSource
                } | Out-Null
            }
        } else {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Warning' -Code 'SHEET_GROUP_TRANSITION_NO_MEMBERS' `
                    -Message 'Sheet-group transition skipped: no associated members resolved.' -Data @{
                    triggerDocumentGuid = $TriggerDocumentGuid
                    triggerDocumentName = $TriggerDocumentName
                    triggerFileType = $triggerRole
                    folderPath = $FolderPath
                    targetState = $target
                    transitionSource = $TransitionSource
                    auditEventId = $AuditEventId
                    jobId = $JobId
                } | Out-Null
            }
            return @{ skipped = $true; skipReason = 'no_members'; members = @() }
        }
    }

    if (-not $StateByGuid) { $StateByGuid = @{} }
    if (-not $PreviousStateByGuid) { $PreviousStateByGuid = @{} }

    if ($Context -and $Context.ContainsKey('sheetStateSync')) {
        $sync = _QCAT-ToHashtable $Context.sheetStateSync
        if ($sync -and $sync.ContainsKey('updates')) {
            foreach ($upd in @($sync.updates)) {
                $u = _QCAT-ToHashtable $upd
                if (-not $u) { continue }
                $ug = [string]$u.documentGuid
                if ($ug) { $PreviousStateByGuid[$ug.ToLowerInvariant()] = [string]$u.fromState }
            }
        }
    }

    foreach ($member in @($Members)) {
        $dg = [string]$member.documentGuid
        if (-not $dg) { continue }
        $key = $dg.ToLowerInvariant()
        if (-not $PreviousStateByGuid.ContainsKey($key)) {
            $prev = ''
            if (Get-Command -Name '_PWD-GetSheetIndexPwStateName' -ErrorAction SilentlyContinue) {
                $prev = [string](_PWD-GetSheetIndexPwStateName -Config $Config -DocumentGuid $dg)
            }
            if ([string]::IsNullOrWhiteSpace($prev) -and $StateByGuid.ContainsKey($key)) {
                $prev = [string]$StateByGuid[$key]
            } elseif ([string]::IsNullOrWhiteSpace($prev) -and $member.document -and (Get-Command -Name '_PWD-GetWorkflowStateFromDocumentRow' -ErrorAction SilentlyContinue)) {
                $prev = [string](_PWD-GetWorkflowStateFromDocumentRow -DocRow $member.document)
            }
            $PreviousStateByGuid[$key] = $prev
        }
    }

    $memberResults = [System.Collections.Generic.List[object]]::new()
    $notifyMember = _QCAT-ResolveSheetPackageNotificationMember -Members $Members
    $notifyGuid = if ($notifyMember) { [string]$notifyMember.documentGuid } else { '' }
    $notifyName = if ($notifyMember) { [string]$notifyMember.documentName } else { '' }
    $packagePreviousState = _QCAT-NormalizeValue $SourceState
    if ([string]::IsNullOrWhiteSpace($packagePreviousState) -and $notifyGuid) {
        $notifyKey = $notifyGuid.ToLowerInvariant()
        if ($PreviousStateByGuid.ContainsKey($notifyKey)) {
            $packagePreviousState = _QCAT-NormalizeValue ([string]$PreviousStateByGuid[$notifyKey])
        }
    }

    $expectedRoles = @('dgn', 'pdf', 'qcPdf')
    $resolvedRoles = @{}
    foreach ($member in @($Members)) {
        $role = _QCAT-ResolveMemberFileRole -DocumentName ([string]$member.documentName)
        if ($expectedRoles -contains $role) { $resolvedRoles[$role] = $member }
    }
    foreach ($role in @($expectedRoles)) {
        if (-not $resolvedRoles.ContainsKey($role)) {
            [void]$memberResults.Add(@{
                role = $role
                skipped = $true
                skipReason = 'sibling_missing'
                documentGuid = ''
                documentName = ''
            })
        }
    }

    $notifyTransitionId = $null
    $transitionGroupId = [guid]::NewGuid()
    $sheetPackageId = $null
    if (Get-Command -Name 'Resolve-SheetPackageIdForSheetGroup' -ErrorAction SilentlyContinue) {
        $sheetPackageId = Resolve-SheetPackageIdForSheetGroup -Config $Config -FolderPath $FolderPath `
            -SheetStem $sheetStem -DocumentGuid $TriggerDocumentGuid -DocumentName $TriggerDocumentName
    }

    foreach ($member in @($Members)) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        if (-not $dg) {
            [void]$memberResults.Add(@{
                role = (_QCAT-ResolveMemberFileRole -DocumentName $dn)
                skipped = $true
                skipReason = 'missing_document_guid'
                documentGuid = ''
                documentName = $dn
            })
            continue
        }

        $role = _QCAT-ResolveMemberFileRole -DocumentName $dn
        $key = $dg.ToLowerInvariant()
        $prevState = if ($PreviousStateByGuid.ContainsKey($key)) { [string]$PreviousStateByGuid[$key] } else { '' }
        if ([string]::IsNullOrWhiteSpace($prevState) -and -not [string]::IsNullOrWhiteSpace($packagePreviousState)) {
            $prevState = $packagePreviousState
        }
        $liveState = if ($StateByGuid.ContainsKey($key)) { [string]$StateByGuid[$key] } else { '' }
        if ([string]::IsNullOrWhiteSpace($liveState) -and $member.document -and (Get-Command -Name '_PWD-GetWorkflowStateFromDocumentRow' -ErrorAction SilentlyContinue)) {
            $liveState = [string](_PWD-GetWorkflowStateFromDocumentRow -DocRow $member.document)
        }
        $finalState = $target
        if (-not [string]::IsNullOrWhiteSpace($liveState)) { $finalState = _QCAT-NormalizeValue $liveState }

        $logicalKey = Get-QCSheetGroupTransitionKey -SheetStem $sheetStem -DocumentGuid $dg -TargetState $target `
            -TransitionSource $TransitionSource -AuditEventId $AuditEventId -JobId $JobId

        $shouldRecord = (_QCAT-NormalizeValue $prevState) -ne $target

        $telemetry = @{
            role = $role
            documentGuid = $dg
            documentName = $dn
            previousState = _QCAT-NormalizeValue $prevState
            liveState = _QCAT-NormalizeValue $liveState
            finalState = $finalState
            logicalTransitionKey = $logicalKey
            skipped = $true
            skipReason = 'no_state_change'
            historyInserted = $false
            transitionInserted = $false
            transitionReused = $false
            mirrorInserted = $false
            transitionId = $null
        }

        if ($shouldRecord) {
            $recordSettings = $settings
            if ($TransitionSource -eq 'automation_prepend_completion') {
                if (-not [bool]$settings.recordFromProcessor) {
                    $telemetry.skipReason = 'processor_telemetry_disabled'
                    [void]$memberResults.Add($telemetry)
                    continue
                }
            } elseif (-not [bool]$settings.recordStateHistory -and -not [bool]$settings.recordTransitions) {
                $telemetry.skipReason = 'telemetry_disabled'
                [void]$memberResults.Add($telemetry)
                continue
            }

            $rec = _QCAT-WriteSheetGroupMemberWorkflowEvents -Config $Config -Settings $settings `
                -DocumentGuid $dg -DocumentName $dn -FolderPath $FolderPath `
                -PreviousState $prevState -CurrentState $target -TransitionSource $TransitionSource `
                -AuditActionName $AuditActionName -AuditEventId $AuditEventId -JobId $JobId -JobType $JobType `
                -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
                -Document $member.document -Context $Context -DryRun:$DryRun `
                -SheetPackageId $sheetPackageId -TransitionGroupId $transitionGroupId
            $telemetry.skipped = [bool]$rec.skipped
            $telemetry.skipReason = if ($rec.skipReason) { [string]$rec.skipReason } else { '' }
            $telemetry.historyInserted = [bool]$rec.historyInserted
            $telemetry.transitionInserted = [bool]$rec.transitionInserted
            $telemetry.transitionReused = [bool]$rec.transitionReused
            $telemetry.mirrorInserted = [bool]$rec.mirrorInserted
            $telemetry.transitionId = $rec.transitionId
            if ($dg -eq $notifyGuid -and $null -ne $rec.transitionId) { $notifyTransitionId = $rec.transitionId }
        }

        [void]$memberResults.Add($telemetry)
    }

    $notificationEmitted = $false
    $shouldNotify = -not $SuppressNotification
    if ($shouldNotify) {
        $shouldNotify = Test-QCShouldNotifyForSheetPackageMember -Config $Config -DocumentName $notifyName `
            -NotifyOnStateChange ([bool]$settings.notifyOnStateChange)
        if ($shouldNotify -and (Test-QCShouldSuppressAuditStateChangeNotification -Config $Config -ChangedByUser $ChangedByUser)) {
            $shouldNotify = $false
        }
        if ($shouldNotify -and (Test-QCShouldSuppressAuditReadyForQcBaselineNotification -Config $Config `
                -PreviousState $packagePreviousState -CurrentState $target)) {
            $shouldNotify = $false
        }
    }

    if ($shouldNotify -and -not [string]::IsNullOrWhiteSpace($notifyGuid) -and $packagePreviousState -ne $target) {
        if (-not (Get-Command -Name 'Invoke-QCWorkflowStateChangeNotification' -ErrorAction SilentlyContinue)) {
            try { Import-Module (Join-Path $PSScriptRoot 'QC.Workflow.psm1') -ErrorAction SilentlyContinue } catch { }
        }
        if (Get-Command -Name 'Invoke-QCWorkflowStateChangeNotification' -ErrorAction SilentlyContinue) {
            $stateTransitionKey = $null
            if (Get-Command -Name 'Get-QCAuditStateTransitionKey' -ErrorAction SilentlyContinue) {
                $stateTransitionKey = Get-QCAuditStateTransitionKey -AuditEventId $AuditEventId -LastAuditEventAt $LastAuditEventAt `
                    -ChangedByUser $ChangedByUser -TriggerDocumentGuid $TriggerDocumentGuid -TransitionId $notifyTransitionId
            }
            $sheetPdfName = $notifyName
            if ($notifyName -match '(?i)-qc\.pdf$') {
                $sheetPdfName = [System.IO.Path]::GetFileNameWithoutExtension($notifyName) + '.pdf'
            }
            $notifyContext = @{
                config = $Config
                folderPath = $FolderPath
                documentPath = ($FolderPath + '\' + $notifyName)
                sourceName = $sheetPdfName
                stateTransitionKey = $stateTransitionKey
                changedByUser = $ChangedByUser
                changedByUsername = $ChangedByUsername
                notificationStateSource = 'sheet_group_transition'
            }
            if ($null -ne $notifyTransitionId -and $notifyTransitionId -gt 0) {
                $notifyContext['transitionId'] = $notifyTransitionId
            }
            if ($Context -and $Context.ContainsKey('attributes')) {
                $notifyContext['attributes'] = $Context.attributes
            }
            if (-not $DryRun) {
                try {
                    $doc = _QCAT-BuildNotificationDocument -Config $Config -FolderPath $FolderPath `
                        -DocumentName $notifyName -DocumentGuid $notifyGuid -Attributes @{}
                    $notif = Invoke-QCWorkflowStateChangeNotification -Config $Config -Context $notifyContext `
                        -PreviousState $packagePreviousState -CurrentState $target -Document $doc
                    $notificationEmitted = $true
                    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                        Write-QCJsonLog -Level 'Information' -Code 'SHEET_GROUP_TRANSITION_NOTIFY' `
                            -Message 'Sheet-group workflow notification evaluated once per logical transition.' -Data @{
                            triggerDocumentGuid = $TriggerDocumentGuid
                            triggerDocumentName = $TriggerDocumentName
                            triggerFileType = $triggerRole
                            notifyDocumentGuid = $notifyGuid
                            notifyDocumentName = $notifyName
                            previousState = $packagePreviousState
                            currentState = $target
                            transitionSource = $TransitionSource
                            auditEventId = $AuditEventId
                            jobId = $JobId
                            notificationCode = if ($notif) { [string]$notif.Code } else { '' }
                        } | Out-Null
                    }
                } catch {
                    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                        Write-QCJsonLog -Flush -Level 'Warning' -Code 'SHEET_GROUP_TRANSITION_NOTIFY_FAILED' `
                            -Message $_.Exception.Message -Data @{
                            notifyDocumentGuid = $notifyGuid
                            notifyDocumentName = $notifyName
                            transitionSource = $TransitionSource
                        } | Out-Null
                    }
                }
            }
        }
    }

    $qcCompleteTransition = _QCAT-GetPackageQcCompleteTransitionFromMemberResults -MemberResults $memberResults `
        -TargetState $target -Config $Config
    if (-not $qcCompleteTransition) {
        $qcCompleteTransition = _QCAT-GetPackageQcCompleteTransitionFromSheetStateSync -Context $Context `
            -TargetState $target -Config $Config
    }
    if ($qcCompleteTransition) {
        _QCAT-TryRecordQCCycleCompletion -Config $Config `
            -PreviousState ([string]$qcCompleteTransition.previousState) -CurrentState ([string]$qcCompleteTransition.currentState) `
            -TransitionSource $TransitionSource -DocumentGuid $TriggerDocumentGuid `
            -DocumentName $TriggerDocumentName -FolderPath $FolderPath -SheetStem $sheetStem `
            -AuditEventId $AuditEventId -TransitionEventId $notifyTransitionId -ChangedByUsername $ChangedByUsername `
            -Context $Context -Members $Members -DryRun:$DryRun
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'SHEET_GROUP_TRANSITION_COMPLETE' `
            -Message 'Sheet-group workflow transition telemetry completed.' -Data @{
            triggerDocumentGuid = $TriggerDocumentGuid
            triggerDocumentName = $TriggerDocumentName
            triggerFileType = $triggerRole
            sheetStem = $sheetStem
            folderPath = $FolderPath
            sheetPackageId = if ($sheetPackageId) { $sheetPackageId.ToString() } else { '' }
            transitionGroupId = $transitionGroupId.ToString()
            targetState = $target
            sourceState = $packagePreviousState
            transitionSource = $TransitionSource
            auditEventId = $AuditEventId
            jobId = $JobId
            notificationEmitted = $notificationEmitted
            members = @($memberResults)
        } | Out-Null
    }

    return @{
        skipped = $false
        triggerDocumentGuid = $TriggerDocumentGuid
        triggerDocumentName = $TriggerDocumentName
        triggerFileType = $triggerRole
        sheetStem = $sheetStem
        sheetPackageId = $sheetPackageId
        transitionGroupId = $transitionGroupId
        targetState = $target
        transitionSource = $TransitionSource
        auditEventId = $AuditEventId
        jobId = $JobId
        notificationEmitted = $notificationEmitted
        members = @($memberResults)
    }
}

function Test-QCIsQcPdfDocumentName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DocumentName)
    return ([string]$DocumentName -match '(?i)-qc\.pdf$')
}

function Get-QCAuditWorkflowTriggerSettings {
    <#
    .SYNOPSIS
    Resolves auditPoller.workflowTriggers settings (defaults enable telemetry; notifications follow notifications.enabled).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $defaults = @{
        enabled                         = $true
        recordStateHistory              = $true
        recordAttributeHistory          = $true
        recordTransitions               = $true
        recordFromProcessor             = $true
        recordProcessingJobs            = $true
        notifyOnStateChange             = $true
        qcPdfNotificationsOnly          = $true
        ignoreStateChangeFromAutomation = $false
        suppressBaselineIndexStateTransition = $true
        baselineStateNames              = @()
        processingGoLiveUtc             = ''
        automationPwUsernames           = @('srv_typsa_archivist')
        automationPwUserNumbers         = @()
    }
    try {
        if ($Config.ContainsKey('auditPoller') -and $Config.auditPoller) {
            $ap = _QCAT-ToHashtable $Config.auditPoller
            if ($ap -and $ap.ContainsKey('workflowTriggers') -and $ap.workflowTriggers) {
                $wt = _QCAT-ToHashtable $ap.workflowTriggers
                if ($wt) {
                    foreach ($k in $wt.Keys) {
                        if (-not $defaults.ContainsKey($k)) { continue }
                        if ($k -eq 'processingGoLiveUtc') {
                            if (-not [string]::IsNullOrWhiteSpace([string]$wt[$k])) {
                                $defaults[$k] = [string]$wt[$k].Trim()
                            }
                            continue
                        }
                        if ($k -in @('automationPwUsernames', 'automationPwUserNumbers', 'baselineStateNames')) {
                            if ($wt[$k] -is [System.Collections.IEnumerable] -and -not ($wt[$k] -is [string])) {
                                $defaults[$k] = @($wt[$k] | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                            } elseif (-not [string]::IsNullOrWhiteSpace([string]$wt[$k])) {
                                $defaults[$k] = @([string]$wt[$k])
                            }
                            continue
                        }
                        try { $defaults[$k] = [bool]$wt[$k] } catch { }
                    }
                }
            }
        }
    } catch { }
    return $defaults
}

function Test-QCIsAutomationPwActor {
    <#
    .SYNOPSIS
    True when the audit actor matches configured automation service accounts (by PW user number or login name).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )

    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.ignoreStateChangeFromAutomation) { return $false }

    if ($null -ne $ChangedByUser) {
        try {
            $n = [int]$ChangedByUser
            if ($n -gt 0) {
                foreach ($configured in @($settings.automationPwUserNumbers)) {
                    try { if ([int]$configured -eq $n) { return $true } } catch { }
                }
            }
        } catch { }
    }

    $user = _QCAT-NormalizeValue $ChangedByUsername
    if ([string]::IsNullOrWhiteSpace($user)) { return $false }
    foreach ($configured in @($settings.automationPwUsernames)) {
        if ([string]::IsNullOrWhiteSpace([string]$configured)) { continue }
        if ($user.Equals([string]$configured, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}


function Get-QCRestartIntakeSourceStateNames {
    <#
    .SYNOPSIS
    Workflow state labels that may intentionally restart a QC cycle by moving back to QC Initiated.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $seen = @{}
    $names = [System.Collections.Generic.List[string]]::new()
    $addName = {
        param([string]$Value)
        $n = _QCAT-NormalizeValue $Value
        if ([string]::IsNullOrWhiteSpace($n)) { return }
        $key = $n.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { return }
        $seen[$key] = $true
        [void]$names.Add($n)
    }

    if (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            if ($wf -and (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue)) {
                foreach ($key in @('correctionsReceived','correctionsInProgress','complete','error')) {
                    $name = Get-QCWorkflowStateName -Settings $wf -StateKey $key
                    if (-not [string]::IsNullOrWhiteSpace($name)) { & $addName $name }
                }
            }
        } catch { }
    }
    foreach ($fallback in @('Corrections Received','Corrections In Progress','QC Complete','Error Needs Attention')) { & $addName $fallback }
    return @($names)
}

function Test-QCWorkflowStateIsRestartIntakeTransition {
    <#
    .SYNOPSIS
    True for deliberate QC cycle restart/intake transitions back to QC Initiated.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$PreviousState = '',
        [string]$CurrentState = ''
    )

    $prev = _QCAT-NormalizeValue $PreviousState
    $curr = _QCAT-NormalizeValue $CurrentState
    if ([string]::IsNullOrWhiteSpace($prev) -or [string]::IsNullOrWhiteSpace($curr)) { return $false }

    $initiated = 'QC Initiated'
    if (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            if ($wf -and (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue)) {
                $cfgInitiated = Get-QCWorkflowStateName -Settings $wf -StateKey 'qcInitiated'
                if (-not [string]::IsNullOrWhiteSpace($cfgInitiated)) { $initiated = [string]$cfgInitiated }
            }
        } catch { }
    }
    if (-not $curr.Equals($initiated, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }

    foreach ($source in @(Get-QCRestartIntakeSourceStateNames -Config $Config)) {
        if ($prev.Equals([string]$source, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Get-QCBaselineWorkflowStateNames {
    <#
    .SYNOPSIS
    Workflow state labels treated as sheet_index baseline (no telemetry when DB had no prior state).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $seen = @{}
    $names = [System.Collections.Generic.List[string]]::new()

    $addName = {
        param([string]$Value)
        $n = _QCAT-NormalizeValue $Value
        if ([string]::IsNullOrWhiteSpace($n)) { return }
        $key = $n.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { return }
        $seen[$key] = $true
        [void]$names.Add($n)
    }

    if (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            if ($wf -and (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue)) {
                $prod = Get-QCWorkflowStateName -Settings $wf -StateKey 'production'
                if (-not [string]::IsNullOrWhiteSpace($prod)) { & $addName $prod }
            }
        } catch { }
    }
    if ($names.Count -eq 0) { & $addName 'In Production' }

    $wt = Get-QCAuditWorkflowTriggerSettings -Config $Config
    foreach ($extra in @($wt.baselineStateNames)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$extra)) { & $addName ([string]$extra) }
    }

    return @($names)
}

function Test-QCShouldSuppressBaselineSheetIndexStateTransition {
    <#
    .SYNOPSIS
    True when sheet_index had no prior state and PW already shows a baseline lifecycle state (index seed only).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState
    )

    $wt = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$wt.suppressBaselineIndexStateTransition) { return $false }

    $prev = _QCAT-NormalizeValue $PreviousState
    if (-not [string]::IsNullOrWhiteSpace($prev)) { return $false }

    $curr = _QCAT-NormalizeValue $CurrentState
    if ([string]::IsNullOrWhiteSpace($curr)) { return $false }

    foreach ($baseline in @(Get-QCBaselineWorkflowStateNames -Config $Config)) {
        if ($curr.Equals($baseline, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Test-QCShouldSuppressAuditReadyForQcBaselineNotification {
    <#
    .SYNOPSIS
    True when audit telemetry shows empty prior state landing on Ready for QC (DOCUMENT_CREATE / index seed).
    The real QC Initiated -> Ready for QC notification is owned by prepend/worker with an explicit previousState.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState
    )

    $prev = _QCAT-NormalizeValue $PreviousState
    if (-not [string]::IsNullOrWhiteSpace($prev)) { return $false }

    $curr = _QCAT-NormalizeValue $CurrentState
    if ([string]::IsNullOrWhiteSpace($curr)) { return $false }

    $readyNames = [System.Collections.Generic.List[string]]::new()
    if (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            if ($wf -and (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue)) {
                foreach ($key in @('readyForQc', 'qcReceived')) {
                    $name = Get-QCWorkflowStateName -Settings $wf -StateKey $key
                    if (-not [string]::IsNullOrWhiteSpace($name)) { [void]$readyNames.Add([string]$name) }
                }
            }
        } catch { }
    }
    if ($readyNames.Count -eq 0) { [void]$readyNames.Add('Ready for QC') }

    foreach ($ready in @($readyNames)) {
        if ($curr.Equals([string]$ready, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function _QCAT-ParsePwActTimeUtc {
    param([string]$ActTime)
    if ([string]::IsNullOrWhiteSpace($ActTime)) { return $null }
    $s = [string]$ActTime.Trim()
    try {
        return [DateTimeOffset]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime()
    } catch { }
    try {
        $dt = [DateTime]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal)
        return [DateTimeOffset]::new($dt.ToUniversalTime())
    } catch { }
    return $null
}

function Test-QCShouldSkipAuditWorkflowProcessingForEvent {
    <#
    .SYNOPSIS
    True when audit event pw_acttime is before workflowTriggers.processingGoLiveUtc (expansion / backlog skip).
    .DESCRIPTION
    Skips DOCUMENT_STATE sibling sync and workflow telemetry for historical audit rows while still ingesting
    and marking them processed. Use when enabling QC on folders that already have sheet_index rows.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$ActTime = ''
    )

    $wt = Get-QCAuditWorkflowTriggerSettings -Config $Config
    $goLiveRaw = ''
    try { $goLiveRaw = [string]$wt.processingGoLiveUtc } catch { }
    if ([string]::IsNullOrWhiteSpace($goLiveRaw)) { return $false }

    $goLive = _QCAT-ParsePwActTimeUtc -ActTime $goLiveRaw
    $eventAt = _QCAT-ParsePwActTimeUtc -ActTime $ActTime
    if ($null -eq $goLive -or $null -eq $eventAt) { return $false }
    return ($eventAt -lt $goLive)
}

function Test-QCShouldSuppressAuditStateChangeNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )
    return (Test-QCIsAutomationPwActor -Config $Config -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername)
}

function Test-QCShouldSuppressAuditSheetStateSync {
    <#
    .SYNOPSIS
    Skips sibling sheet sync for automation-originated DOCUMENT_STATE echoes (e.g. DGN after QC PDF was already aligned).
    Allows sync when automation changed a *-qc.pdf (prepend set QC Received → propagate to siblings once).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentName,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )

    if (-not (Test-QCIsAutomationPwActor -Config $Config -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername)) {
        return $false
    }
    if (Test-QCIsQcPdfDocumentName -DocumentName $DocumentName) { return $false }
    return $true
}

function _QCAT-BuildNotificationDocument {
    param(
        [hashtable]$Config,
        [string]$FolderPath,
        [string]$DocumentName,
        [string]$DocumentGuid,
        [hashtable]$Attributes = @{}
    )
    $notifCfg = $null
    try {
        if ($Config.ContainsKey('notifications') -and $Config.notifications) {
            $notifCfg = _QCAT-ToHashtable $Config.notifications
        }
    } catch { }
    $attrCfg = @{}
    if ($notifCfg -and $notifCfg.ContainsKey('attributes')) {
        $attrCfg = _QCAT-ToHashtable $notifCfg.attributes
        if (-not $attrCfg) { $attrCfg = @{} }
    }
    $reviewerField = if ($attrCfg.reviewerEmailField) { [string]$attrCfg.reviewerEmailField } else { 'EM_Reviewer_Email' }
    $designerField = if ($attrCfg.designerEmailField) { [string]$attrCfg.designerEmailField } else { 'EM_Designer_Email' }
    $checkerField = if ($attrCfg.checkerEmailField) { [string]$attrCfg.checkerEmailField } else { 'EM_Checker_Email' }

    $doc = [ordered]@{
        Name         = $DocumentName
        DocumentName = $DocumentName
        DocumentGUID = $DocumentGuid
        FolderPath   = $FolderPath
    }
    if ($Attributes -and $Attributes.Count -gt 0) {
        foreach ($k in $Attributes.Keys) { $doc[$k] = $Attributes[$k] }
        if ($Attributes.ContainsKey($reviewerField)) { $doc[$reviewerField] = $Attributes[$reviewerField] }
        if ($Attributes.ContainsKey($designerField)) { $doc[$designerField] = $Attributes[$designerField] }
        if ($Attributes.ContainsKey($checkerField)) { $doc[$checkerField] = $Attributes[$checkerField] }
    }
    elseif ($Config -and -not [string]::IsNullOrWhiteSpace($FolderPath) -and -not [string]::IsNullOrWhiteSpace($DocumentName)) {
        $srcName = [string]$DocumentName
        if ($srcName -match '(?i)-qc\.pdf$') {
            $srcName = [System.IO.Path]::GetFileNameWithoutExtension($srcName) + '.pdf'
        }
        if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
            $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $srcName -Config $Config
            if ($pw.found) {
                if ($pw.designerEmail) { $doc[$designerField] = [string]$pw.designerEmail }
                if ($pw.reviewerEmail) { $doc[$reviewerField] = [string]$pw.reviewerEmail }
                if ($pw.checkerEmail) { $doc[$checkerField] = [string]$pw.checkerEmail }
            }
        }
    }
    return [pscustomobject]$doc
}

function Resolve-QCWorkflowEventQcReviewType {
    <#
    .SYNOPSIS
    Resolves QC_Review_Type for qc_workflow_events payload and related telemetry.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Config = @{},
        [string]$DocumentGuid = '',
        [string]$FolderPath = '',
        [string]$DocumentName = '',
        [hashtable]$Context = $null,
        [hashtable]$PwAttributes = $null,
        [object]$Document = $null
    )

    if ($Context) {
        if ($Context.ContainsKey('attributes') -and $Context.attributes) {
            $attrs = _QCAT-ToHashtable $Context.attributes
            if ($attrs) {
                foreach ($k in @('reviewType', 'qcReviewType', 'qc_review_type')) {
                    if ($attrs.ContainsKey($k) -and -not [string]::IsNullOrWhiteSpace([string]$attrs[$k])) {
                        return [string]$attrs[$k]
                    }
                }
            }
        }
        foreach ($k in @('reviewType', 'qcReviewType')) {
            if ($Context.ContainsKey($k) -and -not [string]::IsNullOrWhiteSpace([string]$Context[$k])) {
                return [string]$Context[$k]
            }
        }
    }

    $reviewCol = 'QC_Review_Type'
    if ($Config -and (Get-Command -Name 'Get-PWQcReviewTypeAttributeName' -ErrorAction SilentlyContinue)) {
        try { $reviewCol = Get-PWQcReviewTypeAttributeName -Config $Config } catch { }
    }

    if ($PwAttributes) {
        foreach ($k in @('qc_review_type', 'reviewType', 'qcReviewType', $reviewCol, 'QC_Review_Type')) {
            if ($PwAttributes.ContainsKey($k) -and -not [string]::IsNullOrWhiteSpace([string]$PwAttributes[$k])) {
                return [string]$PwAttributes[$k]
            }
        }
    }

    if ($Document) {
        foreach ($k in @($reviewCol, 'QC_Review_Type', 'qcReviewType', 'reviewType')) {
            try {
                if ($Document.PSObject.Properties[$k] -and -not [string]::IsNullOrWhiteSpace([string]$Document.$k)) {
                    return [string]$Document.$k
                }
            } catch { }
        }
    }

    $folder = [string]$FolderPath
    $source = [string]$DocumentName
    if ($Context) {
        if ([string]::IsNullOrWhiteSpace($folder) -and $Context.ContainsKey('folderPath') -and $Context.folderPath) {
            $folder = [string]$Context.folderPath
        }
        if ([string]::IsNullOrWhiteSpace($folder) -and $Context.ContainsKey('job') -and $Context.job) {
            $job = _QCAT-ToHashtable $Context.job
            if ($job -and $job.sourceFolder) { $folder = [string]$job.sourceFolder }
        }
        if ([string]::IsNullOrWhiteSpace($source) -and $Context.ContainsKey('documentName') -and $Context.documentName) {
            $source = [string]$Context.documentName
        }
        if ([string]::IsNullOrWhiteSpace($source) -and $Context.ContainsKey('job') -and $Context.job) {
            $job = _QCAT-ToHashtable $Context.job
            if ($job -and $job.sourceName) { $source = [string]$job.sourceName }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($source) -and $source -match '(?i)-qc\.pdf$') {
        $source = [string]([System.IO.Path]::GetFileNameWithoutExtension($source)) + '.pdf'
    }
    if ($Config -and -not [string]::IsNullOrWhiteSpace($folder) -and -not [string]::IsNullOrWhiteSpace($source)) {
        if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
            $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $folder -SourceDocumentName $source -Config $Config
            if ($pw.found -and -not [string]::IsNullOrWhiteSpace([string]$pw.qcReviewType)) {
                return [string]$pw.qcReviewType
            }
        }
    }

    if ($Config -and -not [string]::IsNullOrWhiteSpace($DocumentGuid)) {
        if (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue) {
            if (Test-QCDatabaseEnabled -Config $Config) {
                try {
                    $siRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT qc_review_type FROM sheet_index WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $DocumentGuid }
                    if ($siRes.IsSuccess -and $siRes.Data.table -and $siRes.Data.table.Rows.Count -gt 0) {
                        $r = $siRes.Data.table.Rows[0]
                        if (-not ($r.qc_review_type -is [DBNull]) -and -not [string]::IsNullOrWhiteSpace([string]$r.qc_review_type)) {
                            return [string]$r.qc_review_type
                        }
                    }
                } catch { }
            }
        }
    }

    return ''
}

function _QCAT-WriteWorkflowEventMirror {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$FromValue = '',
        [string]$ToValue = '',
        [string]$JobId = '',
        [string]$JobType = '',
        [string]$EventType = 'STATE_CHANGE',
        [string]$AuditActionName = '',
        [string]$QcReviewType = '',
        [hashtable]$Context = $null,
        [hashtable]$PwAttributes = $null,
        [object]$Document = $null,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [Nullable[int]]$TransitionEventId = $null,
        [Nullable[guid]]$SheetPackageId = $null,
        [Nullable[guid]]$TransitionGroupId = $null
    )

    if (-not (Get-Command -Name 'Write-QCWorkflowEventRow' -ErrorAction SilentlyContinue)) { return }
    if ([string]::IsNullOrWhiteSpace($QcReviewType)) {
        $QcReviewType = Resolve-QCWorkflowEventQcReviewType -Config $Config -DocumentGuid $DocumentGuid `
            -FolderPath $FolderPath -DocumentName $DocumentName -Context $Context -PwAttributes $PwAttributes -Document $Document
    }
    $payload = @{
        documentName = $DocumentName
        auditAction = $AuditActionName
    }
    if ($null -ne $ChangedByUser) { $payload['changedByUser'] = $ChangedByUser }
    if (-not [string]::IsNullOrWhiteSpace($ChangedByUsername)) { $payload['changedByUsername'] = [string]$ChangedByUsername }
    $payloadJson = ''
    try { $payloadJson = ($payload | ConvertTo-Json -Compress) } catch { }
    Write-QCWorkflowEventRow -Config $Config -DocumentId $DocumentGuid -JobId $JobId -EventType $EventType `
        -PreviousPwState $FromValue -TargetPwState $ToValue -DecisionCode $JobType -PayloadJson $payloadJson `
        -QcReviewType $QcReviewType -TransitionEventId $TransitionEventId `
        -SheetPackageId $SheetPackageId -TransitionGroupId $TransitionGroupId | Out-Null
}

function _QCAT-ResolveSheetStemFromDocumentName {
    param([string]$DocumentName)
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        try { return [string](Get-PWSheetStemFromDocumentName -DocumentName $DocumentName) } catch { }
    }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension([string]$DocumentName)
    if ($stem -match '(?i)-qc$') { $stem = $stem -replace '(?i)-qc$', '' }
    return $stem
}

function _QCAT-ResolveAssociatedSheetMemberNames {
    param([string]$SheetStem)
    if ([string]::IsNullOrWhiteSpace($SheetStem)) { return @() }
    if (Get-Command -Name 'Get-PWAssociatedSheetDocumentNames' -ErrorAction SilentlyContinue) {
        try { return @(Get-PWAssociatedSheetDocumentNames -SheetStem $SheetStem) } catch { }
    }
    return @(
        ($SheetStem + '.pdf')
        ($SheetStem + '.dgn')
        ($SheetStem + '-qc.pdf')
    )
}

function _QCAT-BuildAssociatedStateDiagnostics {
    param(
        [array]$Members = @(),
        [hashtable]$StateByGuid = @{},
        [hashtable]$Config = $null
    )
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($member in @($Members)) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        if (-not $dn) { continue }
        $role = 'other'
        if ($dn -match '(?i)-qc\.pdf$') { $role = 'qcPdf' }
        elseif ($dn -match '(?i)\.dgn$') { $role = 'dgn' }
        elseif ($dn -match '(?i)\.pdf$') { $role = 'pdf' }
        $liveState = ''
        $key = $dg.ToLowerInvariant()
        if ($dg -and $StateByGuid -and $StateByGuid.ContainsKey($key)) {
            $liveState = [string]$StateByGuid[$key]
        } elseif ($member.document -and (Get-Command -Name '_PWD-GetWorkflowStateFromDocumentRow' -ErrorAction SilentlyContinue)) {
            $liveState = [string](_PWD-GetWorkflowStateFromDocumentRow -DocRow $member.document)
        }
        $sheetIndexState = ''
        if ($dg -and $Config -and (Get-Command -Name '_PWD-GetSheetIndexPwStateName' -ErrorAction SilentlyContinue)) {
            $sheetIndexState = [string](_PWD-GetSheetIndexPwStateName -Config $Config -DocumentGuid $dg)
        }
        [void]$rows.Add(@{
            role = $role
            documentGuid = $dg
            documentName = $dn
            livePwState = $liveState
            sheetIndexState = $sheetIndexState
        })
    }
    return @($rows)
}

function _QCAT-TestSheetPdfWouldRegressFromCanonical {
    param(
        [hashtable]$Config,
        [array]$Members = @(),
        [hashtable]$StateByGuid = @{},
        [string]$CanonicalState = '',
        [string]$LastAuditEventAt = ''
    )
    $canonical = _QCAT-NormalizeValue $CanonicalState
    if ([string]::IsNullOrWhiteSpace($canonical)) { return $null }
    $currentAt = _QCAT-ParsePwActTimeUtc -ActTime $LastAuditEventAt
    foreach ($member in @($Members)) {
        $dn = [string]$member.documentName
        if (-not (($dn -match '(?i)\.pdf$') -and ($dn -notmatch '(?i)-qc\.pdf$'))) { continue }
        $dg = [string]$member.documentGuid
        if (-not $dg) { break }

        $indexState = ''
        $indexAtRaw = ''
        if (Get-Command -Name '_PWD-GetSheetIndexStateSnapshot' -ErrorAction SilentlyContinue) {
            $snap = _PWD-GetSheetIndexStateSnapshot -Config $Config -DocumentGuid $dg
            if ($snap) {
                $indexState = _QCAT-NormalizeValue ([string]$snap.pwStateName)
                if ($snap.lastAuditEventAt) { $indexAtRaw = [string]$snap.lastAuditEventAt }
            }
        }
        if ([string]::IsNullOrWhiteSpace($indexState)) { break }

        $liveState = ''
        $key = $dg.ToLowerInvariant()
        if ($StateByGuid -and $StateByGuid.ContainsKey($key)) {
            $liveState = _QCAT-NormalizeValue ([string]$StateByGuid[$key])
        }
        if ([string]::IsNullOrWhiteSpace($liveState)) { $liveState = $indexState }

        $indexAt = _QCAT-ParsePwActTimeUtc -ActTime $indexAtRaw
        if ($null -eq $indexAt) { break }
        if ($currentAt -and ($indexAt -le $currentAt)) { break }
        if ($indexState -eq $canonical) { break }
        if ($liveState -ne $indexState -and $liveState -eq $canonical) { break }

        return @{
            pdfDocumentGuid = $dg
            pdfDocumentName = $dn
            pdfLiveState = $liveState
            pdfIndexState = $indexState
            pdfLastAuditEventAt = $indexAtRaw
        }
    }
    return $null
}

function Test-QCDocumentStateAuditEventIsStale {
    <#
    .SYNOPSIS
    True when a DOCUMENT_STATE audit event is superseded by a newer sheet state event or would regress a newer manual PDF state.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = '',
        [Nullable[long]]$AuditEventId = $null,
        [string]$LastAuditEventAt = '',
        [string]$CanonicalState = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [array]$Members = @(),
        [hashtable]$StateByGuid = @{},
        [string]$SheetStem = ''
    )

    $sheetStem = ([string]$SheetStem).Trim()
    if ([string]::IsNullOrWhiteSpace($sheetStem)) {
        $sheetStem = _QCAT-ResolveSheetStemFromDocumentName -DocumentName $DocumentName
    }

    $memberNames = @()
    $memberGuids = @()
    foreach ($member in @($Members)) {
        if ($member.documentName) { $memberNames += [string]$member.documentName }
        if ($member.documentGuid) { $memberGuids += [string]$member.documentGuid }
    }
    if ($memberNames.Count -eq 0) {
        $memberNames = @(_QCAT-ResolveAssociatedSheetMemberNames -SheetStem $sheetStem)
    }
    if ($DocumentGuid -and ($memberGuids -notcontains $DocumentGuid)) {
        $memberGuids += [string]$DocumentGuid
    }

    $actorIsAutomation = $false
    if (Get-Command -Name 'Test-QCIsAutomationPwActor' -ErrorAction SilentlyContinue) {
        try { $actorIsAutomation = Test-QCIsAutomationPwActor -Config $Config -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername } catch { }
    }

    $decision = @{
        isStale = $false
        reason = ''
        decision = 'process'
        currentAuditEventId = $AuditEventId
        currentAuditTime = $LastAuditEventAt
        newerAuditEventId = $null
        newerAuditTime = ''
        actor = $ChangedByUsername
        actorIsAutomation = $actorIsAutomation
        triggerDocument = $DocumentName
        triggerDocumentGuid = $DocumentGuid
        sheetStem = $sheetStem
        liveSourceState = $CanonicalState
        associatedStates = @(_QCAT-BuildAssociatedStateDiagnostics -Members $Members -StateByGuid $StateByGuid -Config $Config)
    }

    if ($null -eq $AuditEventId -or $AuditEventId -le 0) {
        if ([string]::IsNullOrWhiteSpace($LastAuditEventAt)) { return $decision }
    }

    if (Get-Command -Name 'Get-QCNewerSheetDocumentStateAuditEvent' -ErrorAction SilentlyContinue) {
        try {
            $newerRes = Get-QCNewerSheetDocumentStateAuditEvent -Config $Config -FolderPath $FolderPath `
                -MemberDocumentNames $memberNames -MemberDocumentGuids $memberGuids `
                -CurrentAuditEventId $AuditEventId -CurrentAuditEventAt $LastAuditEventAt
            if ($newerRes.IsSuccess -and $newerRes.Data -and $newerRes.Data.found -eq $true) {
                $decision.isStale = $true
                $decision.reason = 'newer_audit_event'
                $decision.decision = 'skipped'
                try { $decision.newerAuditEventId = [long]$newerRes.Data.id } catch { }
                $decision.newerAuditTime = [string]$newerRes.Data.pwActtime
                $decision.newerAuditDocumentName = [string]$newerRes.Data.documentName
                $decision.newerAuditDocumentGuid = [string]$newerRes.Data.documentGuid
                $decision.newerAuditProcessed = [bool]$newerRes.Data.processed
                return $decision
            }
        } catch { }
    }

    $regression = _QCAT-TestSheetPdfWouldRegressFromCanonical -Config $Config -Members $Members -StateByGuid $StateByGuid `
        -CanonicalState $CanonicalState -LastAuditEventAt $LastAuditEventAt
    if ($regression) {
        $decision.isStale = $true
        $decision.reason = 'regressive_pdf_state'
        $decision.decision = 'skipped'
        $decision.pdfDocumentName = [string]$regression.pdfDocumentName
        $decision.pdfLiveState = [string]$regression.pdfLiveState
        $decision.pdfIndexState = [string]$regression.pdfIndexState
        $decision.pdfLastAuditEventAt = [string]$regression.pdfLastAuditEventAt
    }
    return $decision
}

function Test-QCShouldNotifyForSheetPackageMember {
    <#
    .SYNOPSIS
    True when a workflow state-change notification should be sent for this document (one notify per sheet package when qcPdfNotificationsOnly is enabled).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentName,
        [bool]$NotifyOnStateChange = $true
    )

    if (-not $NotifyOnStateChange) { return $false }
    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.notifyOnStateChange) { return $false }
    if ([bool]$settings.qcPdfNotificationsOnly -and -not (Test-QCIsQcPdfDocumentName -DocumentName $DocumentName)) {
        return $false
    }
    return $true
}

function Invoke-QCAuditWorkflowStateChangeTriggers {
    <#
    .SYNOPSIS
    Records workflow state transitions from audit events and optionally sends QC PDF notifications.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState,
        [object]$Document = $null,
        [hashtable]$PwAttributes = $null,
        [bool]$DryRun = $false,
        [string]$AuditActionName = 'DOCUMENT_STATE',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [string]$LastAuditEventAt = '',
        [Nullable[long]]$AuditEventId = $null,
        [string]$PreviousStateSource = '',
        [string]$CurrentStateSource = '',
        [array]$StaleCheckMembers = @(),
        [hashtable]$StaleCheckStateByGuid = @{},
        [string]$StaleCheckSheetStem = '',
        [string]$StaleCheckCanonicalState = '',
        [switch]$SuppressNotification
    )

    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.enabled) { return }

    $prev = _QCAT-NormalizeValue $PreviousState
    $curr = _QCAT-NormalizeValue $CurrentState
    if ($prev -eq $curr) { return }

    $isDocumentStateAudit = ([string]$AuditActionName).Trim() -eq 'DOCUMENT_STATE'
    $staleDecision = $null
    $skipStaleSideEffects = $false
    if ($isDocumentStateAudit) {
        $canonicalForStale = if (-not [string]::IsNullOrWhiteSpace($StaleCheckCanonicalState)) {
            [string]$StaleCheckCanonicalState
        } else {
            $curr
        }
        $staleDecision = Test-QCDocumentStateAuditEventIsStale -Config $Config -FolderPath $FolderPath `
            -DocumentName $DocumentName -DocumentGuid $DocumentGuid -AuditEventId $AuditEventId `
            -LastAuditEventAt $LastAuditEventAt -CanonicalState $canonicalForStale `
            -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -Members $StaleCheckMembers -StateByGuid $StaleCheckStateByGuid -SheetStem $StaleCheckSheetStem
        if ($staleDecision -and [bool]$staleDecision.isStale) { $skipStaleSideEffects = $true }
    }

    if (Test-QCShouldSuppressBaselineSheetIndexStateTransition -Config $Config -PreviousState $prev -CurrentState $curr) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_BASELINE_STATE_SUPPRESSED' `
                -Message 'Skipped workflow triggers for sheet_index baseline seed (empty prior state).' -Data @{
                documentGuid = $DocumentGuid; documentName = $DocumentName; currentState = $curr
            } | Out-Null
        }
        return
    }

    $shouldNotify = Test-QCShouldNotifyForSheetPackageMember -Config $Config -DocumentName $DocumentName `
        -NotifyOnStateChange ([bool]$settings.notifyOnStateChange)
    if ($SuppressNotification) { $shouldNotify = $false }
    if ($shouldNotify -and (Test-QCShouldSuppressAuditStateChangeNotification -Config $Config -ChangedByUser $ChangedByUser)) {
        $shouldNotify = $false
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_NOTIFY_SKIPPED_AUTOMATION' `
                -Message 'Skipped QC notification for automation-originated DOCUMENT_STATE.' -Data @{
                documentGuid = $DocumentGuid; documentName = $DocumentName; changedByUser = $ChangedByUser
                currentState = $curr; previousState = $prev
            } | Out-Null
        }
    }
    if ($shouldNotify -and (Test-QCShouldSuppressAuditReadyForQcBaselineNotification -Config $Config -PreviousState $prev -CurrentState $curr)) {
        $shouldNotify = $false
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_NOTIFY_SKIPPED_BASELINE_READY_FOR_QC' `
                -Message 'Skipped Ready for QC audit notification for empty prior state (DOCUMENT_CREATE / index baseline).' -Data @{
                auditEventId = $AuditEventId
                auditActionName = $AuditActionName
                documentGuid = $DocumentGuid
                documentName = $DocumentName
                folderPath = $FolderPath
                currentState = $curr
                previousState = $prev
            } | Out-Null
        }
    }

    if ([bool]$settings.recordStateHistory) {
        Write-QCDocumentStateHistoryRow -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -EventType 'STATE_CHANGE' -OldValue $prev -NewValue $curr `
            -FieldName 'pw_state_name' -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -SourceAuditId $AuditEventId | Out-Null
    }

    if ($skipStaleSideEffects) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            $staleLog = @{
                documentGuid = $DocumentGuid
                documentName = $DocumentName
                folderPath = $FolderPath
                previousState = $prev
                currentState = $curr
                decision = 'skipped'
                notify = 'skipped'
                sync = 'skipped'
            }
            foreach ($k in @($staleDecision.Keys)) { $staleLog[$k] = $staleDecision[$k] }
            Write-QCJsonLog -Flush -Level 'Information' -Code 'QC_NOTIFICATION_SKIPPED_STALE_STATE_EVENT' `
                -Message 'Skipped workflow notification for superseded DOCUMENT_STATE audit event.' -Data $staleLog | Out-Null
        }
        return
    }

    $transitionId = $null
    $transitionWritten = $false
    if ([bool]$settings.recordTransitions) {
        $tr = Ensure-QCTransitionEvent -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -TransitionType 'STATE_CHANGE' -FromValue $prev -ToValue $curr `
            -JobType 'audit_trigger' -TriggerAuditId $AuditEventId -ChangedByUser $ChangedByUser `
            -ChangedByUsername $ChangedByUsername
        if ($tr.IsSuccess -and $tr.Data -and $null -ne $tr.Data.transitionId) {
            try { $transitionId = [int]$tr.Data.transitionId } catch { $transitionId = $null }
        }
        $transitionReused = $false
        try {
            if ($tr.IsSuccess -and $tr.Data -and $tr.Data.written -eq $true) { $transitionWritten = $true }
            if ($tr.IsSuccess -and $tr.Data -and $tr.Data.reused -eq $true) { $transitionReused = $true }
        } catch { }
    }

    if ($transitionWritten -or $transitionReused) {
        $attrs = @{}
        if ($PwAttributes) { $attrs = $PwAttributes }
        _QCAT-WriteWorkflowEventMirror -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -FromValue $prev -ToValue $curr -JobType 'audit_trigger' -EventType 'STATE_CHANGE' `
            -AuditActionName $AuditActionName -PwAttributes $attrs -Document $Document `
            -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername -TransitionEventId $transitionId

        $sheetStem = _QCAT-ResolveSheetStemFromDocumentName -DocumentName $DocumentName
        _QCAT-TryRecordQCCycleCompletion -Config $Config -PreviousState $prev -CurrentState $curr `
            -TransitionSource 'audit_trigger' -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -SheetStem $sheetStem -AuditEventId $AuditEventId -TransitionEventId $transitionId `
            -ChangedByUsername $ChangedByUsername -PwAttributes $attrs -Document $Document `
            -Members $StaleCheckMembers -DryRun:$DryRun
    }

    if (-not $shouldNotify) { return }
    if (-not (Get-Command -Name 'Invoke-QCWorkflowStateChangeNotification' -ErrorAction SilentlyContinue)) {
        try { Import-Module (Join-Path $PSScriptRoot 'QC.Workflow.psm1') -ErrorAction SilentlyContinue } catch { }
    }
    if (-not (Get-Command -Name 'Invoke-QCWorkflowStateChangeNotification' -ErrorAction SilentlyContinue)) { return }

    $attrs = @{}
    if ($PwAttributes) { $attrs = $PwAttributes }
    $Document = _QCAT-BuildNotificationDocument -Config $Config -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid -Attributes $attrs

    if ($DryRun) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_NOTIFY_PLANNED' -Message 'QC notification planned (dry-run).' -Data @{
                documentGuid = $DocumentGuid; documentName = $DocumentName; previousState = $prev; currentState = $curr
            }
        }
        return
    }

    $stateTransitionKey = $null
    if (Get-Command -Name 'Get-QCAuditStateTransitionKey' -ErrorAction SilentlyContinue) {
        $stateTransitionKey = Get-QCAuditStateTransitionKey -AuditEventId $AuditEventId -LastAuditEventAt $LastAuditEventAt `
            -ChangedByUser $ChangedByUser -TriggerDocumentGuid $DocumentGuid -TransitionId $transitionId
    }

    if ($null -ne $ChangedByUser -and $ChangedByUser -gt 0 -and (Get-Command -Name 'Sync-PWUserDirectory' -ErrorAction SilentlyContinue)) {
        try { Sync-PWUserDirectory -Config $Config -UserNumbers @([int]$ChangedByUser) -MaxUsers 1 | Out-Null } catch { }
    }
    $notifyUsername = ''
    if (-not [string]::IsNullOrWhiteSpace($ChangedByUsername)) {
        $notifyUsername = [string]$ChangedByUsername.Trim()
    }

    $notificationStateSource = if (-not [string]::IsNullOrWhiteSpace($CurrentStateSource)) {
        [string]$CurrentStateSource
    } elseif (-not [string]::IsNullOrWhiteSpace($PreviousStateSource)) {
        [string]$PreviousStateSource
    } else {
        'auditTriggerParameter'
    }
    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_NOTIFY_STATE' `
            -Message 'Workflow notification state resolved for DOCUMENT_STATE trigger.' -Data @{
            auditEventId = $AuditEventId
            documentGuid = $DocumentGuid
            documentName = $DocumentName
            folderPath = $FolderPath
            previousState = $prev
            previousStateSource = if ($PreviousStateSource) { [string]$PreviousStateSource } else { 'auditTriggerParameter' }
            currentState = $curr
            currentStateSource = if ($CurrentStateSource) { [string]$CurrentStateSource } else { 'auditTriggerParameter' }
            notificationStateSource = $notificationStateSource
            notificationStateValue = $curr
            changedByUser = $ChangedByUser
            changedByUsername = $notifyUsername
        } | Out-Null
    }

    $sheetPdfName = $DocumentName
    if ($DocumentName -match '(?i)-qc\.pdf$') {
        $sheetPdfName = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName) + '.pdf'
    }

    $notifyContext = @{
        config = $Config
        folderPath = $FolderPath
        documentPath = ($FolderPath + '\' + $DocumentName)
        sourceName = $sheetPdfName
        stateTransitionKey = $stateTransitionKey
        changedByUser = $ChangedByUser
        changedByUsername = $notifyUsername
        notificationStateSource = $notificationStateSource
    }
    if ($null -ne $transitionId -and $transitionId -gt 0) {
        $notifyContext['transitionId'] = $transitionId
    }
    if ($attrs -and $attrs.Count -gt 0) {
        $notifyContext['attributes'] = $attrs
    }
    if ($Config -and (Get-Command -Name 'Get-QCSheetIndexCycle' -ErrorAction SilentlyContinue)) {
        try {
            $sheetStem = ''
            if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
                $sheetStem = [string](Get-PWSheetStemFromDocumentName -DocumentName $sheetPdfName)
            }
            $cycle = Get-QCSheetIndexCycle -Config $Config -DocumentGuid $DocumentGuid -FolderPath $FolderPath -SheetStem $sheetStem
            if ($cycle -and -not [string]::IsNullOrWhiteSpace($cycle.cycleId)) {
                if (-not $notifyContext.ContainsKey('attributes') -or -not $notifyContext.attributes) {
                    $notifyContext['attributes'] = @{}
                }
                $notifyContext.attributes['cycleId'] = [string]$cycle.cycleId
                if ($null -ne $cycle.cycleNumber) {
                    $notifyContext.attributes['cycleNumber'] = [string]$cycle.cycleNumber
                }
                $notifyContext['cycleId'] = [string]$cycle.cycleId
            }
        } catch { }
    }

    try {
        $notif = Invoke-QCWorkflowStateChangeNotification -Config $Config -Context $notifyContext `
            -PreviousState $prev -CurrentState $curr -Document $Document
        if ($notif -is [System.Array]) { $notif = $notif[-1] }
        if ($null -eq $notif) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_NOTIFY_UNAVAILABLE' `
                    -Message 'Workflow notification returned no result.' -Data @{
                    auditEventId = $AuditEventId
                    documentGuid = $DocumentGuid
                    documentName = $DocumentName
                    previousState = $prev
                    currentState = $curr
                } | Out-Null
            }
        }
        elseif ($notif.Code -eq 'QC_NOTIFICATION_MODULE_UNAVAILABLE') {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_NOTIFY_UNAVAILABLE' `
                    -Message ([string]$notif.Message) -Data @{
                    auditEventId = $AuditEventId
                    documentGuid = $DocumentGuid
                    documentName = $DocumentName
                    previousState = $prev
                    currentState = $curr
                    notificationCode = [string]$notif.Code
                } | Out-Null
            }
        }
        elseif ($notif.Code -in @('QC_NOTIFICATION_SKIPPED_DUPLICATE', 'QC_NOTIFICATION_ENQUEUE_SKIPPED_DUPLICATE')) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_NOTIFY_SKIPPED_DUPLICATE' `
                    -Message ([string]$notif.Message) -Data @{
                    auditEventId = $AuditEventId
                    documentGuid = $DocumentGuid
                    documentName = $DocumentName
                    previousState = $prev
                    currentState = $curr
                    dedupeKey = if ($notif.Data -and $notif.Data.dedupeKey) { [string]$notif.Data.dedupeKey } else { '' }
                } | Out-Null
            }
        }
        elseif (-not $notif.IsSuccess) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_NOTIFY_FAILED' `
                    -Message ([string]$notif.Message) -Data @{
                    auditEventId = $AuditEventId
                    documentGuid = $DocumentGuid
                    documentName = $DocumentName
                    previousState = $prev
                    currentState = $curr
                    notificationCode = [string]$notif.Code
                } | Out-Null
            }
        }
    } catch {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_NOTIFY_FAILED' -Message $_.Exception.Message -Data @{
                documentGuid = $DocumentGuid; documentName = $DocumentName; previousState = $prev; currentState = $curr
            }
        }
    }
}

function _QCAT-GetContextDocumentIdentity {
    param([hashtable]$Context)
    $document = $null
    if ($Context -and $Context.ContainsKey('document')) { $document = $Context.document }
    $docGuid = ''
    $docName = ''
    $folderPath = ''
    if ($document) {
        try {
            if ($document.PSObject.Properties['DocumentGUID'] -and $document.DocumentGUID) { $docGuid = [string]$document.DocumentGUID }
            elseif ($document.PSObject.Properties['DocumentGuid'] -and $document.DocumentGuid) { $docGuid = [string]$document.DocumentGuid }
        } catch { }
        try {
            if ($document.PSObject.Properties['Name'] -and $document.Name) { $docName = [string]$document.Name }
            elseif ($document.PSObject.Properties['DocumentName'] -and $document.DocumentName) { $docName = [string]$document.DocumentName }
        } catch { }
        try {
            if ($document.PSObject.Properties['FolderPath'] -and $document.FolderPath) { $folderPath = [string]$document.FolderPath }
        } catch { }
    }
    if ($Context) {
        if (-not $docGuid -and $Context.ContainsKey('documentGuid')) { $docGuid = [string]$Context.documentGuid }
        if (-not $docName -and $Context.ContainsKey('documentName')) { $docName = [string]$Context.documentName }
        if (-not $folderPath) {
            foreach ($k in @('folderPath', 'sourceFolder', 'incomingFolderPath')) {
                if ($Context.ContainsKey($k) -and $Context[$k]) { $folderPath = [string]$Context[$k]; break }
            }
        }
    }
    $job = $null
    if ($Context -and $Context.ContainsKey('job')) { $job = _QCAT-ToHashtable $Context.job }
    if ($job) {
        if (-not $docGuid -and $job.ContainsKey('documentGuid')) { $docGuid = [string]$job.documentGuid }
        if (-not $docName -and $job.ContainsKey('fileName')) { $docName = [string]$job.fileName }
        if (-not $folderPath -and $job.ContainsKey('sourceFolder')) { $folderPath = [string]$job.sourceFolder }
        if (-not $folderPath -and $job.ContainsKey('metadata')) {
            $meta = _QCAT-ToHashtable $job.metadata
            if ($meta) {
                if (-not $docGuid -and $meta.ContainsKey('documentId')) { $docGuid = [string]$meta.documentId }
                if (-not $docGuid -and $meta.ContainsKey('triggerDocumentGuid')) { $docGuid = [string]$meta.triggerDocumentGuid }
                if (-not $docName -and $meta.ContainsKey('triggerDocumentName')) { $docName = [string]$meta.triggerDocumentName }
                if (-not $folderPath -and $meta.ContainsKey('folderPath')) { $folderPath = [string]$meta.folderPath }
            }
        }
    }
    return @{ documentGuid = $docGuid; documentName = $docName; folderPath = $folderPath; job = $job }
}

function _QCAT-GetContextStateChangeActor {
    param([hashtable]$Context)
    $changedByUser = $null
    $changedByUsername = ''
    if ($Context) {
        if ($Context.ContainsKey('changedByUser') -and $null -ne $Context.changedByUser) {
            try {
                $n = [int]$Context.changedByUser
                if ($n -gt 0) { $changedByUser = $n }
            } catch { }
        }
        if ($Context.ContainsKey('changedByUsername') -and -not [string]::IsNullOrWhiteSpace($Context.changedByUsername)) {
            $changedByUsername = [string]$Context.changedByUsername.Trim()
        }
    }
    $id = _QCAT-GetContextDocumentIdentity -Context $Context
    if ($id.job -and $id.job.ContainsKey('metadata') -and $id.job.metadata) {
        $meta = _QCAT-ToHashtable $id.job.metadata
        if ($meta) {
            if ($null -eq $changedByUser -and $meta.ContainsKey('changedByUser') -and $null -ne $meta.changedByUser) {
                try {
                    $n = [int]$meta.changedByUser
                    if ($n -gt 0) { $changedByUser = $n }
                } catch { }
            }
            if ([string]::IsNullOrWhiteSpace($changedByUsername) -and $meta.ContainsKey('changedByUsername') -and -not [string]::IsNullOrWhiteSpace($meta.changedByUsername)) {
                $changedByUsername = [string]$meta.changedByUsername.Trim()
            }
        }
    }
    return @{ changedByUser = $changedByUser; changedByUsername = $changedByUsername }
}

function Invoke-QCProcessorWorkflowStateTelemetry {
    <#
    .SYNOPSIS
    Records processor-driven workflow state changes (QC_PREPEND / comment sync) without sending email.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Context,
        [string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState,
        [string]$JobType = 'processor'
    )

    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.enabled -or -not [bool]$settings.recordFromProcessor) { return }

    $prev = _QCAT-NormalizeValue $PreviousState
    $curr = _QCAT-NormalizeValue $CurrentState
    if ($prev -eq $curr) { return }

    $id = _QCAT-GetContextDocumentIdentity -Context $Context
    if ([string]::IsNullOrWhiteSpace($id.documentGuid)) { return }

    $jobId = ''
    if ($id.job -and $id.job.ContainsKey('id')) { $jobId = [string]$id.job.id }
    $actor = _QCAT-GetContextStateChangeActor -Context $Context

    $transitionSource = if ($JobType -eq 'QC_PREPEND') { 'automation_prepend_completion' } else { 'processor' }
    $sgResult = $null
    if (Get-Command -Name 'Invoke-QCSheetGroupWorkflowTransition' -ErrorAction SilentlyContinue) {
        $sgResult = Invoke-QCSheetGroupWorkflowTransition -Config $Config -TriggerDocumentGuid $id.documentGuid `
            -TriggerDocumentName $id.documentName -FolderPath $id.folderPath -SourceState $prev -TargetState $curr `
            -TransitionSource $transitionSource -JobId $jobId -JobType $JobType -Context $Context `
            -ChangedByUser $actor.changedByUser -ChangedByUsername $actor.changedByUsername -SuppressNotification
        if ($Context -and $sgResult -and $sgResult.members) {
            foreach ($m in @($sgResult.members)) {
                if ($null -ne $m.transitionId -and $m.documentGuid -eq $id.documentGuid) {
                    try { $Context['transitionId'] = [int]$m.transitionId } catch { }
                    break
                }
            }
        }
    } else {
        if ([bool]$settings.recordStateHistory) {
            Write-QCDocumentStateHistoryRow -Config $Config -DocumentGuid $id.documentGuid -DocumentName $id.documentName `
                -FolderPath $id.folderPath -EventType 'STATE_CHANGE' -OldValue $prev -NewValue $curr `
                -FieldName 'pw_state_name' -ChangedByUser $actor.changedByUser -ChangedByUsername $actor.changedByUsername | Out-Null
        }
        if ([bool]$settings.recordTransitions) {
            $tr = Ensure-QCTransitionEvent -Config $Config -DocumentGuid $id.documentGuid -DocumentName $id.documentName `
                -FolderPath $id.folderPath -TransitionType 'STATE_CHANGE' -FromValue $prev -ToValue $curr `
                -JobId $jobId -JobType $JobType -ChangedByUser $actor.changedByUser -ChangedByUsername $actor.changedByUsername
            if ($Context -and $tr.IsSuccess -and $tr.Data -and $null -ne $tr.Data.transitionId) {
                try { $Context['transitionId'] = [int]$tr.Data.transitionId } catch { }
            }
            $written = $false
            $reused = $false
            try {
                if ($tr.IsSuccess -and $tr.Data -and $tr.Data.written -eq $true) { $written = $true }
                if ($tr.IsSuccess -and $tr.Data -and $tr.Data.reused -eq $true) { $reused = $true }
            } catch { }
            if ($written -or $reused) {
                $mirrorTransitionId = $null
                if ($tr.IsSuccess -and $tr.Data -and $null -ne $tr.Data.transitionId) {
                    try { $mirrorTransitionId = [int]$tr.Data.transitionId } catch { }
                }
                _QCAT-WriteWorkflowEventMirror -Config $Config -DocumentGuid $id.documentGuid -DocumentName $id.documentName `
                    -FolderPath $id.folderPath -FromValue $prev -ToValue $curr -JobId $jobId -JobType $JobType `
                    -EventType 'STATE_CHANGE' -Context $Context -Document $(if ($Context -and $Context.document) { $Context.document } else { $null }) `
                    -TransitionEventId $mirrorTransitionId
            }
        }
    }

    if ([bool]$settings.recordProcessingJobs -and (Get-Command -Name 'Write-QCStateChangeJobTelemetry' -ErrorAction SilentlyContinue)) {
        if ($JobType -eq 'QC_COMMENT_STATUS_SYNC' -and $jobId) {
            return
        }
        $triggerSource = 'processor'
        if ($JobType -eq 'QC_PREPEND') { $triggerSource = 'qc_prepend' }
        $qcReviewType = Resolve-QCWorkflowEventQcReviewType -Config $Config -DocumentGuid $id.documentGuid `
            -FolderPath $id.folderPath -DocumentName $id.documentName -Context $Context `
            -Document $(if ($Context -and $Context.document) { $Context.document } else { $null })
        Write-QCStateChangeJobTelemetry -Config $Config -PreviousState $prev -CurrentState $curr `
            -ParentJobId $jobId -DocumentGuid $id.documentGuid -DocumentName $id.documentName `
            -SourceFolder $id.folderPath -TriggerSource $triggerSource -Operation $JobType `
            -QcReviewType $qcReviewType | Out-Null
    }
}

function Invoke-QCProcessorWorkflowAttributeTelemetry {
    <#
    .SYNOPSIS
    Records attribute writeback from QC_PREPEND processor (new values only; audit path has before/after).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)][hashtable]$MappedAttributes
    )

    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.enabled -or -not [bool]$settings.recordFromProcessor) { return }
    if (-not [bool]$settings.recordAttributeHistory) { return }
    if (-not $MappedAttributes -or $MappedAttributes.Keys.Count -eq 0) { return }

    $id = _QCAT-GetContextDocumentIdentity -Context $Context
    if ([string]::IsNullOrWhiteSpace($id.documentGuid)) { return }

    $jobId = ''
    if ($id.job -and $id.job.ContainsKey('id')) { $jobId = [string]$id.job.id }

    foreach ($field in @($MappedAttributes.Keys)) {
        $newVal = _QCAT-NormalizeValue ([string]$MappedAttributes[$field])
        if ([string]::IsNullOrWhiteSpace($newVal)) { continue }
        Write-QCDocumentStateHistoryRow -Config $Config -DocumentGuid $id.documentGuid -DocumentName $id.documentName `
            -FolderPath $id.folderPath -EventType 'ATTR_CHANGE' -OldValue '' -NewValue $newVal `
            -FieldName ([string]$field) | Out-Null
        if ([bool]$settings.recordTransitions) {
            Write-QCTransitionEvent -Config $Config -DocumentGuid $id.documentGuid -DocumentName $id.documentName `
                -FolderPath $id.folderPath -TransitionType 'ATTR_CHANGE' -FromValue '' -ToValue $newVal `
                -JobId $jobId -JobType 'QC_PREPEND' | Out-Null
        }
    }
}

function Invoke-QCAuditWorkflowAttributeChangeTriggers {
    <#
    .SYNOPSIS
    Records attribute field transitions detected after DOCUMENT_ATTR audit events (sheet_index diff).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][hashtable]$FieldChanges,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )

    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.enabled) { return }
    if (-not $FieldChanges -or $FieldChanges.Keys.Count -eq 0) { return }

    foreach ($field in @($FieldChanges.Keys)) {
        $change = _QCAT-ToHashtable $FieldChanges[$field]
        if (-not $change) { continue }
        $oldVal = _QCAT-NormalizeValue ([string]$change.oldValue)
        $newVal = _QCAT-NormalizeValue ([string]$change.newValue)
        if ($oldVal -eq $newVal) { continue }

        if ([bool]$settings.recordAttributeHistory) {
            Write-QCDocumentStateHistoryRow -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
                -FolderPath $FolderPath -EventType 'ATTR_CHANGE' -OldValue $oldVal -NewValue $newVal `
                -FieldName ([string]$field) -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername | Out-Null
        }

        if ([bool]$settings.recordTransitions) {
            Write-QCTransitionEvent -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
                -FolderPath $FolderPath -TransitionType 'ATTR_CHANGE' -FromValue $oldVal -ToValue $newVal `
                -JobType 'audit_trigger' | Out-Null
        }
    }
}

Export-ModuleMember -Function Get-QCAuditWorkflowTriggerSettings, Get-QCBaselineWorkflowStateNames, Get-QCRestartIntakeSourceStateNames, Test-QCWorkflowStateIsRestartIntakeTransition, Get-QCAuditStateTransitionKey, Get-QCPrependStateTransitionDedupeKey, Get-QCSheetGroupTransitionKey, Test-QCIsQcPdfDocumentName, `
    Test-QCIsAutomationPwActor, Test-QCShouldNotifyForSheetPackageMember, Test-QCDocumentStateAuditEventIsStale, Test-QCShouldSuppressBaselineSheetIndexStateTransition, Test-QCShouldSuppressAuditReadyForQcBaselineNotification, Test-QCShouldSkipAuditWorkflowProcessingForEvent, `
    Test-QCShouldSuppressAuditStateChangeNotification, Test-QCShouldSuppressAuditSheetStateSync, `
    Resolve-QCWorkflowEventQcReviewType, Invoke-QCSheetGroupWorkflowTransition, Invoke-QCAuditWorkflowStateChangeTriggers, Invoke-QCAuditWorkflowAttributeChangeTriggers, `
    Invoke-QCProcessorWorkflowStateTelemetry, Invoke-QCProcessorWorkflowAttributeTelemetry
