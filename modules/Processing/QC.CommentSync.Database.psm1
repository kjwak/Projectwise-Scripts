# QC.CommentSync.Database.psm1
# Responsibility: Persist comment sync runs, comments, status history, and workflow events.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core.Database.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'QC.PdfExport.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'QC.AuditTriggers.psm1') -Force -ErrorAction SilentlyContinue

function _QCDB-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCDB-ToHashtable([object]$Value) {
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

function Test-QCCommentSyncDatabaseWritesAllowed {
    param([hashtable]$Config)
    return (Test-QCDatabaseWritesAllowed -Config $Config)
}

function Write-QCWorkflowEvent {
    param(
        [hashtable]$Config,
        [long]$RunId = 0,
        [string]$JobId = '',
        [string]$DocumentId = '',
        [string]$EventType = '',
        [string]$PreviousPwState = '',
        [string]$TargetPwState = '',
        [string]$DecisionCode = '',
        [string]$ProcessorVersion = '',
        [string]$QcReviewType = '',
        [string]$PayloadJson = '',
        [Nullable[int]]$TransitionEventId = $null,
        [switch]$PlannedOnly
    )

    $runIdParam = $null
    if ($RunId -gt 0) { $runIdParam = $RunId }
    return Write-QCWorkflowEventRow -Config $Config -RunId $runIdParam -JobId $JobId -DocumentId $DocumentId `
        -EventType $EventType -PreviousPwState $PreviousPwState -TargetPwState $TargetPwState `
        -DecisionCode $DecisionCode -ProcessorVersion $ProcessorVersion -QcReviewType $QcReviewType `
        -PayloadJson $PayloadJson -TransitionEventId $TransitionEventId -PlannedOnly:$PlannedOnly
}

function Write-QCCommentSyncRun {
    param(
        [hashtable]$Config,
        [hashtable]$RunRecord,
        [switch]$PlannedOnly
    )

    if ($PlannedOnly -or -not (Test-QCCommentSyncDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'QC_COMMENT_RUN_PLANNED' -Message 'Run record not written (dry-run or DB disabled).' -Data @{ planned = $true; run = $RunRecord }
    }

    $sql = @"
INSERT INTO qc_comment_runs
    (job_id, document_id, project_id, pw_path, file_name, file_hash, source_modified_utc,
     previous_pw_state, target_pw_state, state_update_result, parser_status, processor_version)
OUTPUT INSERTED.run_id
VALUES
    (@jobId, @docId, @projectId, @pwPath, @fileName, @fileHash, @sourceMod,
     @prevState, @targetState, @stateResult, @parserStatus, @procVer)
"@
    $params = @{
        jobId = [string]$RunRecord.job_id
        docId = if ($RunRecord.document_id) { [string]$RunRecord.document_id } else { [DBNull]::Value }
        projectId = if ($RunRecord.project_id) { [string]$RunRecord.project_id } else { [DBNull]::Value }
        pwPath = if ($RunRecord.pw_path) { [string]$RunRecord.pw_path } else { [DBNull]::Value }
        fileName = if ($RunRecord.file_name) { [string]$RunRecord.file_name } else { [DBNull]::Value }
        fileHash = if ($RunRecord.file_hash) { [string]$RunRecord.file_hash } else { [DBNull]::Value }
        sourceMod = if ($RunRecord.source_modified_utc) { $RunRecord.source_modified_utc } else { [DBNull]::Value }
        prevState = if ($RunRecord.previous_pw_state) { [string]$RunRecord.previous_pw_state } else { [DBNull]::Value }
        targetState = if ($RunRecord.target_pw_state) { [string]$RunRecord.target_pw_state } else { [DBNull]::Value }
        stateResult = if ($RunRecord.state_update_result) { [string]$RunRecord.state_update_result } else { [DBNull]::Value }
        parserStatus = if ($RunRecord.parser_status) { [string]$RunRecord.parser_status } else { [DBNull]::Value }
        procVer = if ($RunRecord.processor_version) { [string]$RunRecord.processor_version } else { [DBNull]::Value }
    }
    try {
        $r = Invoke-QCDatabaseScalar -Config $Config -Sql $sql -Parameters $params
        if ($r.IsSuccess) {
            return New-QCSuccessResult -Code 'QC_COMMENT_RUN_WRITTEN' -Message 'Comment sync run inserted.' -Data @{ runId = [long]$r.Data.value }
        }
        return $r
    } catch {
        return New-QCFailureResult -Code 'QC_COMMENT_RUN_FAILED' -Message $_.Exception.Message -Data @{}
    }
}

function Write-QCCommentRows {
    param(
        [hashtable]$Config,
        [long]$RunId,
        [string]$DocumentId,
        [object[]]$Annotations,
        [switch]$PlannedOnly
    )

    if ($PlannedOnly -or -not (Test-QCCommentSyncDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'QC_COMMENTS_PLANNED' -Message 'Comments not written.' -Data @{ planned = $true; count = @($Annotations).Count }
    }

    $inserted = 0
    foreach ($a in @($Annotations)) {
        if (-not ($a -is [hashtable])) { continue }
        $rawJson = $null
        if ($a.raw) {
            try { $rawJson = ($a.raw | ConvertTo-Json -Depth 12 -Compress) } catch { $rawJson = [string]$a.raw }
        }
        $sql = @"
INSERT INTO qc_comments
    (run_id, document_id, annotation_id, page_number, author, subject, comment_text, color, status,
     status_author, status_timestamp_utc, created_utc, modified_utc, parent_annotation_id, raw_json)
VALUES
    (@runId, @docId, @annotId, @page, @author, @subject, @text, @color, @status,
     @statusAuthor, @statusTs, @created, @modified, @parent, @raw)
"@
        $params = @{
            runId = $RunId
            docId = if ($DocumentId) { $DocumentId } else { [DBNull]::Value }
            annotId = [string]$a.annotation_id
            page = if ($null -ne $a.page_number) { [int]$a.page_number } else { [DBNull]::Value }
            author = if ($a.author) { [string]$a.author } else { [DBNull]::Value }
            subject = if ($a.subject) { [string]$a.subject } else { [DBNull]::Value }
            text = if ($a.comment_text) { [string]$a.comment_text } else { [DBNull]::Value }
            color = if ($a.color) { [string]$a.color } else { [DBNull]::Value }
            status = if ($a.status) { [string]$a.status } else { [DBNull]::Value }
            statusAuthor = if ($a.status_author) { [string]$a.status_author } else { [DBNull]::Value }
            statusTs = if ($a.status_timestamp_utc) { $a.status_timestamp_utc } else { [DBNull]::Value }
            created = if ($a.created_utc) { $a.created_utc } else { [DBNull]::Value }
            modified = if ($a.modified_utc) { $a.modified_utc } else { [DBNull]::Value }
            parent = if ($a.parent_annotation_id) { [string]$a.parent_annotation_id } else { [DBNull]::Value }
            raw = if ($rawJson) { $rawJson } else { [DBNull]::Value }
        }
        try {
            $r = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
            if ($r.IsSuccess) { $inserted++ }
        } catch { }
    }
    return New-QCSuccessResult -Code 'QC_COMMENTS_WRITTEN' -Message "Inserted $inserted comment row(s)." -Data @{ inserted = $inserted }
}

function Write-QCCommentStatusHistoryRows {
    param(
        [hashtable]$Config,
        [long]$RunId,
        [string]$DocumentId,
        [object[]]$Annotations,
        [switch]$PlannedOnly
    )

    if ($PlannedOnly -or -not (Test-QCCommentSyncDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'QC_COMMENT_HISTORY_PLANNED' -Message 'Status history not written.' -Data @{ planned = $true }
    }

    $written = 0
    foreach ($a in @($Annotations)) {
        if (-not ($a -is [hashtable])) { continue }
        $annotId = [string]$a.annotation_id
        if (_QCDB-IsNullOrWhiteSpace $annotId) { continue }

        $prevStatus = $null
        $prevTs = $null
        try {
            $q = @"
SELECT TOP 1 current_status, current_status_timestamp_utc
FROM qc_comment_status_history
WHERE document_id = @docId AND annotation_id = @annotId
ORDER BY detected_utc DESC
"@
            $qr = Invoke-QCDatabaseQuery -Config $Config -Sql $q -Parameters @{ docId = $DocumentId; annotId = $annotId }
            if ($qr.IsSuccess -and $qr.Data.rowCount -gt 0) {
                $row = $qr.Data.table.Rows[0]
                $prevStatus = [string]$row['current_status']
                if ($row['current_status_timestamp_utc'] -isnot [DBNull]) {
                    $prevTs = $row['current_status_timestamp_utc']
                }
            }
        } catch { }

        $curStatus = [string]$a.status
        $curTs = if ($a.status_timestamp_utc) { $a.status_timestamp_utc } else { $null }
        if ($prevStatus -eq $curStatus) { continue }

        $sql = @"
INSERT INTO qc_comment_status_history
    (document_id, annotation_id, previous_status, current_status,
     previous_status_timestamp_utc, current_status_timestamp_utc, detected_run_id)
VALUES
    (@docId, @annotId, @prevStatus, @curStatus, @prevTs, @curTs, @runId)
"@
        $params = @{
            docId = if ($DocumentId) { $DocumentId } else { [DBNull]::Value }
            annotId = $annotId
            prevStatus = if ($prevStatus) { $prevStatus } else { [DBNull]::Value }
            curStatus = if ($curStatus) { $curStatus } else { [DBNull]::Value }
            prevTs = if ($prevTs) { $prevTs } else { [DBNull]::Value }
            curTs = if ($curTs) { $curTs } else { [DBNull]::Value }
            runId = $RunId
        }
        try {
            $r = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
            if ($r.IsSuccess) { $written++ }
        } catch { }
    }
    return New-QCSuccessResult -Code 'QC_COMMENT_HISTORY_WRITTEN' -Message "Wrote $written status history row(s)." -Data @{ written = $written }
}

function Invoke-QCCommentSyncPersist {
    param(
        [hashtable]$Config,
        [hashtable]$Job,
        [hashtable]$JobMetadata,
        [hashtable]$Decision,
        [object[]]$Annotations,
        [string]$ParserStatus,
        [string]$StateUpdateResult,
        [string]$ProcessorVersion,
        [string]$PreviousPwState,
        [switch]$DryRun
    )

    $planned = [bool]$DryRun -or -not (Test-QCCommentSyncDatabaseWritesAllowed -Config $Config)

    $runRecord = @{
        job_id = [string]$Job.id
        document_id = [string]$JobMetadata.documentId
        project_id = [string]$JobMetadata.projectId
        pw_path = [string]$JobMetadata.pwPath
        file_name = [string]$JobMetadata.fileName
        file_hash = [string]$JobMetadata.fileHash
        source_modified_utc = $JobMetadata.sourceModifiedUtc
        previous_pw_state = $PreviousPwState
        target_pw_state = [string]$Decision.targetState
        state_update_result = $StateUpdateResult
        parser_status = $ParserStatus
        processor_version = $ProcessorVersion
    }

    $payloadJson = $null
    $qcReviewType = ''
    try {
        $payloadObj = _QCDB-ToHashtable $Decision
        if (-not $payloadObj) { $payloadObj = @{} }
        if (Get-Command -Name 'Resolve-QCWorkflowEventQcReviewType' -ErrorAction SilentlyContinue) {
            $qcReviewType = Resolve-QCWorkflowEventQcReviewType -Config $Config `
                -DocumentGuid ([string]$JobMetadata.documentId) -FolderPath ([string]$JobMetadata.folderPath) `
                -DocumentName ([string]$JobMetadata.fileName) -Context @{ job = $Job }
        }
        $payloadJson = ($payloadObj | ConvertTo-Json -Depth 8 -Compress)
    } catch { }

    Write-QCWorkflowEvent -Config $Config -JobId ([string]$Job.id) -DocumentId ([string]$JobMetadata.documentId) `
        -EventType 'STATE_DECIDED' -PreviousPwState $PreviousPwState -TargetPwState ([string]$Decision.targetState) `
        -DecisionCode ([string]$Decision.decisionCode) -ProcessorVersion $ProcessorVersion -QcReviewType $qcReviewType `
        -PayloadJson $payloadJson -PlannedOnly:$planned | Out-Null

    $runRes = Write-QCCommentSyncRun -Config $Config -RunRecord $runRecord -PlannedOnly:$planned
    $runId = 0L
    if ($runRes.IsSuccess -and $runRes.Data.runId) { $runId = [long]$runRes.Data.runId }

    if ($runId -gt 0) {
        Write-QCCommentRows -Config $Config -RunId $runId -DocumentId ([string]$JobMetadata.documentId) -Annotations $Annotations -PlannedOnly:$false | Out-Null
        Write-QCCommentStatusHistoryRows -Config $Config -RunId $runId -DocumentId ([string]$JobMetadata.documentId) -Annotations $Annotations -PlannedOnly:$false | Out-Null
    }

    return @{
        runResult = $runRes
        runId = $runId
        planned = $planned
    }
}

function Update-QCCommentSyncRunStateResult {
    param(
        [hashtable]$Config,
        [long]$RunId,
        [string]$StateUpdateResult,
        [switch]$PlannedOnly
    )

    if ($RunId -le 0) { return New-QCSuccessResult -Code 'QC_COMMENT_RUN_SKIP_UPDATE' -Message 'No run_id to update.' -Data @{} }
    if ($PlannedOnly -or -not (Test-QCCommentSyncDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'QC_COMMENT_RUN_UPDATE_PLANNED' -Message 'Run state result update planned only.' -Data @{ planned = $true; runId = $RunId }
    }

    $sql = 'UPDATE qc_comment_runs SET state_update_result = @result WHERE run_id = @runId'
    try {
        return Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters @{ result = $StateUpdateResult; runId = $RunId }
    } catch {
        return New-QCFailureResult -Code 'QC_COMMENT_RUN_UPDATE_FAILED' -Message $_.Exception.Message -Data @{ runId = $RunId }
    }
}

Export-ModuleMember -Function Test-QCCommentSyncDatabaseWritesAllowed, Write-QCWorkflowEvent, Write-QCCommentSyncRun, Update-QCCommentSyncRunStateResult, Write-QCCommentRows, Write-QCCommentStatusHistoryRows, Invoke-QCCommentSyncPersist
