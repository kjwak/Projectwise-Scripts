# QC.NotificationThreads.psm1
# Responsibility: Durable QC notification email threading (sheet_package_id + review_type).

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue

function _QCNT-IsBlank([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function Get-QCNotificationThreadKey {
    <#
    .SYNOPSIS
    Builds the logical thread key from sheet_package_id and review_type.
    #>
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$SheetPackageId = '',
        [AllowEmptyString()][string]$ReviewType = ''
    )

    $pkg = ([string]$SheetPackageId).Trim().ToLowerInvariant()
    $rt = ([string]$ReviewType).Trim()
    if (_QCNT-IsBlank $pkg -or _QCNT-IsBlank $rt) { return '' }
    return "$pkg|$rt"
}

function _QCNT-ResolveOutputRoot([string]$Path) {
    if (_QCNT-IsBlank $Path) { return '' }
    $p = [string]$Path
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    $root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    return (Join-Path $root $p)
}

function _QCNT-ParseSheetPackageGuid([string]$SheetPackageId) {
    if (_QCNT-IsBlank $SheetPackageId) { return $null }
    try {
        return [guid]([string]$SheetPackageId.Trim())
    } catch {
        return $null
    }
}

function _QCNT-WriteThreadLog {
    param(
        [string]$Code,
        [string]$Level = 'Information',
        [string]$Message = '',
        [hashtable]$Data = @{}
    )
    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    if (_QCNT-IsBlank $Message) { $Message = $Code }
    Write-QCJsonLog -Level $Level -Code $Code -Message $Message -Data $Data | Out-Null
}

function _QCNT-GetMockThreadState([string]$ThreadKey, [string]$OutputRoot) {
    if (_QCNT-IsBlank $ThreadKey -or _QCNT-IsBlank $OutputRoot) { return $null }
    $dir = Join-Path $OutputRoot 'mock\threads'
    $safeKey = ($ThreadKey -replace '[\\/:*?"<>|]+', '_')
    $path = Join-Path $dir "$safeKey.json"
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        $state = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
        return @{
            id = if ($state.thread_id) { [int]$state.thread_id } else { 1 }
            sheetPackageId = ''
            reviewType = ''
            status = 'active'
            graphConversationId = if ($state.conversation_id) { [string]$state.conversation_id } else { '' }
            latestGraphMessageId = if ($state.latest_message_id) { [string]$state.latest_message_id } else { '' }
            latestGraphImmutableMessageId = if ($state.latest_message_id) { [string]$state.latest_message_id } else { '' }
            latestInternetMessageId = ''
        }
    } catch {
        return $null
    }
}

function Get-QCNotificationActiveThread {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$SheetPackageId,
        [Parameter(Mandatory)][string]$ReviewType
    )

    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return $null }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return $null }
    $pkgGuid = _QCNT-ParseSheetPackageGuid -SheetPackageId $SheetPackageId
    if ($null -eq $pkgGuid) { return $null }

    $sql = @"
SELECT TOP (1)
    id, sheet_package_id, review_type, status,
    graph_conversation_id, latest_graph_message_id,
    latest_graph_immutable_message_id, latest_internet_message_id,
    created_at, updated_at, superseded_at
FROM qc_notification_threads
WHERE sheet_package_id = @sheetPackageId
  AND review_type = @reviewType
  AND status = 'active'
ORDER BY id DESC
"@
    $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters @{
        sheetPackageId = $pkgGuid
        reviewType = ([string]$ReviewType).Trim()
    }
    if (-not $res.IsSuccess -or -not $res.Data.table -or $res.Data.table.Rows.Count -eq 0) { return $null }

    $row = $res.Data.table.Rows[0]
    return @{
        id = [int]$row.id
        sheetPackageId = [string]$row.sheet_package_id
        reviewType = [string]$row.review_type
        status = [string]$row.status
        graphConversationId = if ($row.graph_conversation_id -is [DBNull]) { '' } else { [string]$row.graph_conversation_id }
        latestGraphMessageId = if ($row.latest_graph_message_id -is [DBNull]) { '' } else { [string]$row.latest_graph_message_id }
        latestGraphImmutableMessageId = if ($row.latest_graph_immutable_message_id -is [DBNull]) { '' } else { [string]$row.latest_graph_immutable_message_id }
        latestInternetMessageId = if ($row.latest_internet_message_id -is [DBNull]) { '' } else { [string]$row.latest_internet_message_id }
    }
}

function New-QCNotificationThread {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$SheetPackageId,
        [Parameter(Mandatory)][string]$ReviewType,
        [string]$GraphConversationId = '',
        [string]$GraphMessageId = '',
        [string]$GraphImmutableMessageId = '',
        [string]$InternetMessageId = ''
    )

    if (-not (Get-Command -Name 'Invoke-QCDatabaseScalar' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'QC_NOTIFICATION_THREAD_DB_UNAVAILABLE' -Message 'Database module not available.' -Data @{}
    }
    $pkgGuid = _QCNT-ParseSheetPackageGuid -SheetPackageId $SheetPackageId
    if ($null -eq $pkgGuid) {
        return New-QCFailureResult -Code 'QC_NOTIFICATION_THREAD_INVALID_PACKAGE' -Message 'Invalid sheet_package_id for thread.' -Data @{}
    }

    $sql = @"
INSERT INTO qc_notification_threads
    (sheet_package_id, review_type, status, graph_conversation_id,
     latest_graph_message_id, latest_graph_immutable_message_id, latest_internet_message_id)
OUTPUT INSERTED.id
VALUES
    (@sheetPackageId, @reviewType, 'active', @conversationId,
     @messageId, @immutableMessageId, @internetMessageId)
"@
    $params = @{
        sheetPackageId = $pkgGuid
        reviewType = ([string]$ReviewType).Trim()
        conversationId = if (_QCNT-IsBlank $GraphConversationId) { $null } else { $GraphConversationId }
        messageId = if (_QCNT-IsBlank $GraphMessageId) { $null } else { $GraphMessageId }
        immutableMessageId = if (_QCNT-IsBlank $GraphImmutableMessageId) { $null } else { $GraphImmutableMessageId }
        internetMessageId = if (_QCNT-IsBlank $InternetMessageId) { $null } else { $InternetMessageId }
    }
    $res = Invoke-QCDatabaseScalar -Config $Config -Sql $sql -Parameters $params
    if ($res.IsSuccess -and $null -ne $res.Data.value) {
        try {
            $threadId = [int]$res.Data.value
            return New-QCSuccessResult -Code 'QC_NOTIFICATION_THREAD_CREATED' -Message 'Notification thread created.' -Data @{ threadId = $threadId }
        } catch { }
    }

    $existing = Get-QCNotificationActiveThread -Config $Config -SheetPackageId $SheetPackageId -ReviewType $ReviewType
    if ($existing) {
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_THREAD_EXISTING' -Message 'Active thread already exists (concurrent create).' -Data @{ threadId = $existing.id; concurrent = $true }
    }
    $msg = if ($res.Message) { [string]$res.Message } else { 'Failed to create notification thread.' }
    return New-QCFailureResult -Code 'QC_NOTIFICATION_THREAD_CREATE_FAILED' -Message $msg -Data @{}
}

function Update-QCNotificationThreadLatestMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][int]$ThreadId,
        [string]$GraphConversationId = '',
        [string]$GraphMessageId = '',
        [string]$GraphImmutableMessageId = '',
        [string]$InternetMessageId = '',
        [string]$Status = ''
    )

    if (-not (Get-Command -Name 'Invoke-QCDatabaseNonQuery' -ErrorAction SilentlyContinue)) { return }
    $sets = @(
        'updated_at = SYSDATETIMEOFFSET()',
        'latest_graph_message_id = @messageId',
        'latest_graph_immutable_message_id = @immutableMessageId',
        'latest_internet_message_id = @internetMessageId'
    )
    if (-not (_QCNT-IsBlank $GraphConversationId)) {
        $sets += 'graph_conversation_id = @conversationId'
    }
    if (-not (_QCNT-IsBlank $Status)) {
        $sets += 'status = @status'
        if ($Status -eq 'superseded') { $sets += 'superseded_at = SYSDATETIMEOFFSET()' }
    }
    $sql = "UPDATE qc_notification_threads SET $($sets -join ', ') WHERE id = @threadId"
    $params = @{
        threadId = $ThreadId
        conversationId = if (_QCNT-IsBlank $GraphConversationId) { $null } else { $GraphConversationId }
        messageId = if (_QCNT-IsBlank $GraphMessageId) { $null } else { $GraphMessageId }
        immutableMessageId = if (_QCNT-IsBlank $GraphImmutableMessageId) { $null } else { $GraphImmutableMessageId }
        internetMessageId = if (_QCNT-IsBlank $InternetMessageId) { $null } else { $InternetMessageId }
        status = if (_QCNT-IsBlank $Status) { $null } else { $Status }
    }
    Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params | Out-Null
}

function Write-QCNotificationThreadMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][int]$ThreadId,
        [Nullable[int]]$NotificationLogId = $null,
        [string]$WorkflowEvent = '',
        [string]$GraphMessageId = '',
        [string]$GraphImmutableMessageId = '',
        [string]$GraphConversationId = '',
        [string]$InternetMessageId = '',
        [Parameter(Mandatory)][string]$SendMode,
        [string]$DeliveryStatus = 'sent',
        [string]$ErrorMessage = ''
    )

    if (-not (Get-Command -Name 'Invoke-QCDatabaseScalar' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'QC_NOTIFICATION_THREAD_MESSAGE_DB_UNAVAILABLE' -Message 'Database module not available.' -Data @{}
    }

    $sql = @"
INSERT INTO qc_notification_messages
    (thread_id, notification_log_id, workflow_event, graph_message_id, graph_immutable_message_id,
     graph_conversation_id, internet_message_id, send_mode, delivery_status, sent_at, error_message)
OUTPUT INSERTED.id
VALUES
    (@threadId, @notificationLogId, @workflowEvent, @graphMessageId, @graphImmutableMessageId,
     @graphConversationId, @internetMessageId, @sendMode, @deliveryStatus,
     CASE WHEN @deliveryStatus = 'sent' THEN SYSDATETIMEOFFSET() ELSE NULL END, @errorMessage)
"@
    $params = @{
        threadId = $ThreadId
        notificationLogId = if ($null -ne $NotificationLogId) { $NotificationLogId } else { $null }
        workflowEvent = if (_QCNT-IsBlank $WorkflowEvent) { $null } else { $WorkflowEvent }
        graphMessageId = if (_QCNT-IsBlank $GraphMessageId) { $null } else { $GraphMessageId }
        graphImmutableMessageId = if (_QCNT-IsBlank $GraphImmutableMessageId) { $null } else { $GraphImmutableMessageId }
        graphConversationId = if (_QCNT-IsBlank $GraphConversationId) { $null } else { $GraphConversationId }
        internetMessageId = if (_QCNT-IsBlank $InternetMessageId) { $null } else { $InternetMessageId }
        sendMode = $SendMode
        deliveryStatus = $DeliveryStatus
        errorMessage = if (_QCNT-IsBlank $ErrorMessage) { $null } else { $ErrorMessage }
    }
    $res = Invoke-QCDatabaseScalar -Config $Config -Sql $sql -Parameters $params
    if (-not $res.IsSuccess) {
        return New-QCFailureResult -Code 'QC_NOTIFICATION_THREAD_MESSAGE_WRITE_FAILED' -Message $res.Message -Data @{}
    }
    $messageId = $null
    if ($null -ne $res.Data.value) {
        try { $messageId = [int]$res.Data.value } catch { }
    }
    return New-QCSuccessResult -Code 'QC_NOTIFICATION_THREAD_MESSAGE_WRITTEN' -Message 'Thread message recorded.' -Data @{ messageRecordId = $messageId }
}

function Invoke-QCNotificationThreadedSend {
    <#
    .SYNOPSIS
    Resolves thread context, dispatches via provider, and persists thread/message records.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Event,
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Settings,
        [Parameter(Mandatory)][hashtable]$Payload,
        [string]$Provider = 'Mock',
        [switch]$DryRun
    )

    $sheetPackageId = ''
    if ($Event.sheetPackageId) { $sheetPackageId = [string]$Event.sheetPackageId }
    if (_QCNT-IsBlank $sheetPackageId -and $Event.documentGuid -and (Get-Command -Name 'Get-SheetPackageIdForDocument' -ErrorAction SilentlyContinue)) {
        try {
            $resolvedPkg = Get-SheetPackageIdForDocument -Config $Config -DocumentGuid ([string]$Event.documentGuid)
            if ($null -ne $resolvedPkg) { $sheetPackageId = $resolvedPkg.ToString() }
        } catch { }
    }

    $reviewType = ''
    if ($Event.reviewType) { $reviewType = [string]$Event.reviewType }
    elseif ($Event.qcReviewType) { $reviewType = [string]$Event.qcReviewType }

    $threadKey = Get-QCNotificationThreadKey -SheetPackageId $sheetPackageId -ReviewType $reviewType
    $workflowEvent = if ($Event.eventType) { [string]$Event.eventType } else { '' }

    $threadContext = @{
        sheetPackageId = $sheetPackageId
        reviewType = $reviewType
        threadKey = $threadKey
        sendMode = 'unthreaded'
        parentGraphMessageId = ''
        notificationThreadId = $null
        threadStatus = ''
    }

    if (_QCNT-IsBlank $threadKey) {
        $reason = if (_QCNT-IsBlank $sheetPackageId) { 'missing_sheet_package_id' } else { 'missing_review_type' }
        _QCNT-WriteThreadLog -Code 'QC_NOTIFICATION_THREAD_SKIPPED' -Level 'Information' `
            -Message "Email threading skipped: $reason." -Data @{
            sheetPackageId = $sheetPackageId
            reviewType = $reviewType
            reason = $reason
            workflowEvent = $workflowEvent
        }
        $Payload['threadContext'] = $threadContext
        $Payload['threadSendMode'] = 'unthreaded'
        return @{ threadContext = $threadContext; requiresThreadPersistence = $false }
    }

    $activeThread = $null
    if (Get-Command -Name 'Get-QCNotificationActiveThread' -ErrorAction SilentlyContinue) {
        $activeThread = Get-QCNotificationActiveThread -Config $Config -SheetPackageId $sheetPackageId -ReviewType $reviewType
    }
    if (-not $activeThread) {
        $outputRoot = ''
        if ($Settings.outputRoot) { $outputRoot = _QCNT-ResolveOutputRoot -Path ([string]$Settings.outputRoot) }
        if (_QCNT-IsBlank $outputRoot) {
            $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
            $outputRoot = Join-Path $repoRoot 'notifications'
        }
        $activeThread = _QCNT-GetMockThreadState -ThreadKey $threadKey -OutputRoot $outputRoot
    }

    $sendMode = 'root'
    $parentMessageId = ''
    if ($activeThread) {
        $sendMode = 'reply'
        $parentMessageId = $activeThread.latestGraphImmutableMessageId
        if (_QCNT-IsBlank $parentMessageId) { $parentMessageId = $activeThread.latestGraphMessageId }
        $threadContext.notificationThreadId = $activeThread.id
        $threadContext.threadStatus = $activeThread.status
    }

    $threadContext.sendMode = $sendMode
    $threadContext.parentGraphMessageId = $parentMessageId
    $Payload['threadContext'] = $threadContext
    $Payload['threadSendMode'] = $sendMode
    $Payload['parentGraphMessageId'] = $parentMessageId
    $Payload['threadKey'] = $threadKey
    if ($activeThread) {
        $Payload['notificationThreadId'] = $activeThread.id
        $Payload['graphConversationId'] = $activeThread.graphConversationId
    }

    return @{
        threadContext = $threadContext
        activeThread = $activeThread
        requiresThreadPersistence = $true
        sheetPackageId = $sheetPackageId
        reviewType = $reviewType
        threadKey = $threadKey
        sendMode = $sendMode
        parentMessageId = $parentMessageId
        workflowEvent = $workflowEvent
    }
}

function Complete-QCNotificationThreadedSend {
    <#
    .SYNOPSIS
    Persists thread and message records after a successful (or replacement) provider dispatch.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$ThreadPlan,
        [Parameter(Mandatory)][hashtable]$SendResultData,
        [Nullable[int]]$NotificationLogId = $null,
        [string]$WorkflowEvent = '',
        [switch]$ParentInvalid
    )

    if (-not $ThreadPlan.requiresThreadPersistence) { return }

    $sheetPackageId = [string]$ThreadPlan.sheetPackageId
    $reviewType = [string]$ThreadPlan.reviewType
    $sendMode = [string]$SendResultData.sendMode
    if (_QCNT-IsBlank $sendMode) { $sendMode = [string]$ThreadPlan.sendMode }

    $graphMessageId = if ($SendResultData.graphMessageId) { [string]$SendResultData.graphMessageId } else { '' }
    $graphImmutableId = if ($SendResultData.graphImmutableMessageId) { [string]$SendResultData.graphImmutableMessageId } else { $graphMessageId }
    $conversationId = if ($SendResultData.graphConversationId) { [string]$SendResultData.graphConversationId } else { '' }
    $internetMessageId = if ($SendResultData.internetMessageId) { [string]$SendResultData.internetMessageId } else { '' }

    $threadId = $null
    if ($SendResultData.notificationThreadId) {
        try { $threadId = [int]$SendResultData.notificationThreadId } catch { }
    }
    if (-not $threadId -and $ThreadPlan.activeThread) { $threadId = $ThreadPlan.activeThread.id }

    if ($sendMode -in @('root', 'replacement_root') -or (-not $threadId)) {
        if ($ParentInvalid -and $ThreadPlan.activeThread) {
            Update-QCNotificationThreadLatestMessage -Config $Config -ThreadId $ThreadPlan.activeThread.id -Status 'invalid' | Out-Null
        }
        $createRes = New-QCNotificationThread -Config $Config -SheetPackageId $sheetPackageId -ReviewType $reviewType `
            -GraphConversationId $conversationId -GraphMessageId $graphMessageId `
            -GraphImmutableMessageId $graphImmutableId -InternetMessageId $internetMessageId
        if ($createRes.IsSuccess -and $createRes.Data.threadId) {
            $threadId = [int]$createRes.Data.threadId
        } elseif ($ThreadPlan.activeThread) {
            $threadId = $ThreadPlan.activeThread.id
            Update-QCNotificationThreadLatestMessage -Config $Config -ThreadId $threadId `
                -GraphConversationId $conversationId -GraphMessageId $graphMessageId `
                -GraphImmutableMessageId $graphImmutableId -InternetMessageId $internetMessageId | Out-Null
        }
    } elseif ($threadId) {
        Update-QCNotificationThreadLatestMessage -Config $Config -ThreadId $threadId `
            -GraphConversationId $conversationId -GraphMessageId $graphMessageId `
            -GraphImmutableMessageId $graphImmutableId -InternetMessageId $internetMessageId | Out-Null
    }

    if ($threadId) {
        $deliveryStatus = if ($SendResultData.success -eq $false) { 'failed' } else { 'sent' }
        $errMsg = if ($SendResultData.threadError) { [string]$SendResultData.threadError } else { '' }
        [void](Write-QCNotificationThreadMessage -Config $Config -ThreadId $threadId `
            -NotificationLogId $NotificationLogId -WorkflowEvent $WorkflowEvent `
            -GraphMessageId $graphMessageId -GraphImmutableMessageId $graphImmutableId `
            -GraphConversationId $conversationId -InternetMessageId $internetMessageId `
            -SendMode $sendMode -DeliveryStatus $deliveryStatus -ErrorMessage $errMsg)
    }

    _QCNT-WriteThreadLog -Code 'QC_NOTIFICATION_THREAD_COMPLETED' -Level 'Information' `
        -Message 'Notification thread dispatch completed.' -Data @{
        sheetPackageId = $sheetPackageId
        reviewType = $reviewType
        notificationThreadId = $threadId
        graphConversationId = $conversationId
        parentGraphMessageId = if ($ThreadPlan.parentMessageId) { [string]$ThreadPlan.parentMessageId } else { '' }
        sentGraphMessageId = $graphMessageId
        sendMode = $sendMode
        threadStatus = if ($ThreadPlan.activeThread) { $ThreadPlan.activeThread.status } else { 'active' }
        notificationLogId = $NotificationLogId
        workflowEvent = $WorkflowEvent
    }
}

Export-ModuleMember -Function Get-QCNotificationThreadKey, Get-QCNotificationActiveThread, New-QCNotificationThread, `
    Update-QCNotificationThreadLatestMessage, Write-QCNotificationThreadMessage, `
    Invoke-QCNotificationThreadedSend, Complete-QCNotificationThreadedSend
