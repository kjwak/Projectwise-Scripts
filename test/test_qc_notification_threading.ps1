# Tests for durable QC notification email threading (sheet_package_id + review_type).
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Core/Core.Runtime.psm1" -Force
Import-Module (Join-Path $repoRoot 'modules\Notifications\QC.NotificationThreads.psm1') -Force -Global
Import-Module (Join-Path $repoRoot 'modules\Notifications\QC.NotificationMock.psm1') -Force -Global
Import-Module (Join-Path $repoRoot 'modules\Notifications\QC.NotificationGraph.psm1') -Force -Global
Import-Module (Join-Path $repoRoot 'modules\Notifications\QC.Notifications.psm1') -Force -Global

$testRoot = Join-Path $env:TEMP ("qc-thread-test-" + [guid]::NewGuid().ToString('N'))
$mockRoot = Join-Path $testRoot 'notifications'
$dedupePath = Join-Path $testRoot 'dedupe\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $dedupePath) -Force | Out-Null

$pkgA = [guid]::NewGuid().ToString()
$pkgB = [guid]::NewGuid().ToString()
$reviewProduction = 'Production QC'
$reviewPeer = 'Peer Review'

function New-ThreadTestConfig() {
    return @{
        database = @{ enabled = $false }
        notifications = @{
            enabled = $true
            provider = 'Mock'
            dryRun = $false
            outputRoot = $mockRoot
            dedupe = @{
                enabled = $true
                storePath = $dedupePath
                keyFields = @('sheetStem', 'qcProcessType', 'previousState', 'currentState', 'transitionSource', 'logicalTransitionAnchor', 'recipientKey')
                sheetPackageKeyFields = @('sheetStem', 'currentState', 'cycleId')
            }
            email = @{ bodyFormat = 'Text' }
            events = @{
                'QC Received' = @{
                    enabled = $true
                    eventType = 'QC_RECEIVED'
                    to = @('reviewers')
                    cc = @('designers')
                }
                'Verified' = @{
                    enabled = $true
                    eventType = 'VERIFIED'
                    to = @('reviewers')
                }
            }
        }
    }
}

function New-ThreadEvent([string]$SheetPackageId, [string]$ReviewType, [string]$EventType = 'QC_RECEIVED', [string]$DedupeSuffix = '') {
    $dedupe = "dedupe|$SheetPackageId|$ReviewType|$EventType$DedupeSuffix"
    return @{
        eventType = $EventType
        documentName = 'sheet001-prod.pdf'
        documentGuid = 'guid-prod-001'
        previousState = 'In Production'
        currentState = 'QC Received'
        sheetPackageId = $SheetPackageId
        reviewType = $ReviewType
        qcReviewType = $ReviewType
        dedupeKey = $dedupe
        sheetStem = 'sheet001'
        recipientKey = 'recipients:test@company.com'
        logicalTransitionAnchor = "sheet:sheet001|from:In Production|to:QC Received|cycle:1|audit:$dedupe"
        transitionSource = 'test'
        qcProcessType = 'production'
    }
}

function Send-ThreadNotification([hashtable]$Event, [hashtable]$Config, [string]$DedupeSuffix = '') {
    if ($DedupeSuffix) {
        $Event['dedupeKey'] = [string]$Event.dedupeKey + $DedupeSuffix
        $Event['logicalTransitionAnchor'] = [string]$Event.logicalTransitionAnchor + $DedupeSuffix
    }
    return Send-QCNotification -Event $Event -Config $Config -Subject '[Test] notification' -Body 'Test body' `
        -To @('test@company.com') -Cc @()
}

# 1. Thread key is sheet_package_id + review_type only
$keyA = Get-QCNotificationThreadKey -SheetPackageId $pkgA -ReviewType $reviewProduction
$keyB = Get-QCNotificationThreadKey -SheetPackageId $pkgA -ReviewType $reviewPeer
Assert-Eq $keyA "$($pkgA.ToLower())|$reviewProduction" 'Thread key uses package + review type'
Assert-True ($keyA -ne $keyB) 'Different review types produce different thread keys'
$keyOtherPkg = Get-QCNotificationThreadKey -SheetPackageId $pkgB -ReviewType $reviewProduction
Assert-True ($keyA -ne $keyOtherPkg) 'Different packages produce different thread keys'

# 2. First notification creates root thread (mock)
$cfg = New-ThreadTestConfig
$ev1 = New-ThreadEvent -SheetPackageId $pkgA -ReviewType $reviewProduction -DedupeSuffix '|1'
$r1 = Send-ThreadNotification -Event $ev1 -Config $cfg
Assert-True $r1.IsSuccess 'First notification should succeed'
Assert-Eq $r1.Data.send_mode 'root' 'First notification should be root'
Assert-True (-not [string]::IsNullOrWhiteSpace($r1.Data.message_id)) 'Root should have message_id'
Assert-Eq $r1.Data.thread_key $keyA 'Mock thread_key should match logical key'

# 3. Second notification for same package + review type is reply
$ev2 = New-ThreadEvent -SheetPackageId $pkgA -ReviewType $reviewProduction -DedupeSuffix '|2'
$ev2.currentState = 'In Review'
$ev2.previousState = 'QC Received'
$ev2.eventType = 'IN_REVIEW'
$ev2.dedupeKey = "dedupe|$pkgA|$reviewProduction|IN_REVIEW|2"
$r2 = Send-ThreadNotification -Event $ev2 -Config $cfg
Assert-True $r2.IsSuccess 'Second notification should succeed'
Assert-Eq $r2.Data.send_mode 'reply' 'Second notification should be reply'
Assert-Eq $r2.Data.parent_message_id $r1.Data.message_id 'Reply parent should be prior message'

# 4. Different review type creates separate thread
$evPeer = New-ThreadEvent -SheetPackageId $pkgA -ReviewType $reviewPeer -DedupeSuffix '|peer1'
$rPeer = Send-ThreadNotification -Event $evPeer -Config $cfg
Assert-Eq $rPeer.Data.send_mode 'root' 'Different review type should start new root thread'

# 5. Different sheet package creates separate thread
$evPkgB = New-ThreadEvent -SheetPackageId $pkgB -ReviewType $reviewProduction -DedupeSuffix '|b1'
$rPkgB = Send-ThreadNotification -Event $evPkgB -Config $cfg
Assert-Eq $rPkgB.Data.send_mode 'root' 'Different package should start new root thread'

# 6. Later workflow cycle reuses same thread (verified state does not close thread)
$evCycle2 = New-ThreadEvent -SheetPackageId $pkgA -ReviewType $reviewProduction -EventType 'VERIFIED' -DedupeSuffix '|cycle2'
$evCycle2.currentState = 'Verified'
$evCycle2.previousState = 'In Review'
$evCycle2.dedupeKey = "dedupe|$pkgA|$reviewProduction|VERIFIED|cycle2"
$rCycle2 = Send-ThreadNotification -Event $evCycle2 -Config $cfg
Assert-Eq $rCycle2.Data.send_mode 'reply' 'Post-verification notification should continue thread'

# 7. Duplicate events do not produce duplicate sends (dedupe preserved)
$dupCfg = New-ThreadTestConfig
$dupSettings = Get-QCNotificationSettings -Config $dupCfg
$dupEv = New-ThreadEvent -SheetPackageId ([guid]::NewGuid().ToString()) -ReviewType $reviewProduction -DedupeSuffix '|dup'
$dupKey = Get-QCNotificationDedupeKey -Event $dupEv -Settings $dupSettings -Config $dupCfg
Assert-True (Register-QCNotificationDedupeClaim -DedupeKey $dupKey -Settings $dupSettings -Config $dupCfg) 'First dedupe claim should succeed'
$dup1 = Send-ThreadNotification -Event $dupEv -Config $dupCfg
Register-QCNotificationDedupe -DedupeKey $dupKey -Settings $dupSettings -ResultData @{ success = $true } | Out-Null
Assert-True (Test-QCNotificationDedupe -DedupeKey $dupKey -Settings $dupSettings -Config $dupCfg) 'Dedupe key should be registered after send'
Assert-True (-not (Register-QCNotificationDedupeClaim -DedupeKey $dupKey -Settings $dupSettings -Config $dupCfg)) 'Duplicate claim should be rejected'

# 8. Missing parent Graph message creates replacement root (mock)
$cfgReplace = New-ThreadTestConfig
$pkgReplace = [guid]::NewGuid().ToString()
$replaceKey = Get-QCNotificationThreadKey -SheetPackageId $pkgReplace -ReviewType $reviewProduction
$threadDir = Join-Path $mockRoot 'mock\threads'
$safeKey = ($replaceKey -replace '[\\/:*?"<>|]+', '_')
$stalePath = Join-Path $threadDir "$safeKey.json"
New-Item -ItemType Directory -Path $threadDir -Force | Out-Null
@{
    thread_id = 99
    thread_key = $replaceKey
    conversation_id = 'mock-conv-stale'
    latest_message_id = 'stale-parent-id'
    message_count = 1
} | ConvertTo-Json | Set-Content -LiteralPath $stalePath -Encoding UTF8
$evReplace = New-ThreadEvent -SheetPackageId $pkgReplace -ReviewType $reviewProduction -DedupeSuffix '|replace'
$rReplace = Send-ThreadNotification -Event $evReplace -Config $cfgReplace
Assert-Eq $rReplace.Data.send_mode 'replacement_root' 'Stale parent should trigger replacement_root'

# 9. Mock provider fields exposed
foreach ($field in @('thread_id', 'thread_key', 'parent_message_id', 'message_id', 'send_mode')) {
    Assert-True ($r2.Data.ContainsKey($field)) "Mock result should expose $field"
}

# 10. Graph HTTP mock: root then reply via createReply
Import-Module (Join-Path $repoRoot 'modules\Notifications\QC.NotificationGraph.psm1') -Force -Global
Clear-QCNotificationGraphTestHttpHandler
$global:QCTestGraphState = @{
    messages = @{}
    sentOrder = [System.Collections.Generic.List[string]]::new()
}
Set-QCNotificationGraphTestHttpHandler -Handler {
    param($Method, $Uri, $Headers, $Body)
    $graphState = $global:QCTestGraphState
    if ($Uri -match 'oauth2/v2.0/token') {
        return @{ access_token = 'test-token' }
    }
    if ($Method -eq 'GET' -and $Uri -match '/messages/([^/?]+)') {
        $id = [Uri]::UnescapeDataString($Matches[1])
        if (-not $graphState.messages.ContainsKey($id)) { throw 'ErrorItemNotFound' }
        return $graphState.messages[$id]
    }
    if ($Method -eq 'POST' -and $Uri -match '/messages$') {
        $id = 'graph-root-' + ($graphState.sentOrder.Count + 1)
        $msg = @{
            id = $id
            conversationId = 'conv-test-1'
            internetMessageId = "<$id@test.local>"
        }
        $graphState.messages[$id] = $msg
        return $msg
    }
    if ($Method -eq 'POST' -and $Uri -match '/createReply$') {
        $parentId = ($Uri -split '/messages/')[1] -replace '/createReply.*', ''
        $parentId = [Uri]::UnescapeDataString($parentId)
        if (-not $graphState.messages.ContainsKey($parentId)) {
            $graphState.messages[$parentId] = @{
                id = $parentId
                conversationId = 'conv-test-1'
                internetMessageId = "<$parentId@test.local>"
            }
        }
        $id = 'graph-reply-' + ($graphState.sentOrder.Count + 1)
        $msg = @{
            id = $id
            conversationId = $graphState.messages[$parentId].conversationId
            internetMessageId = "<$id@test.local>"
        }
        $graphState.messages[$id] = $msg
        return $msg
    }
    if ($Method -eq 'PATCH' -and $Uri -match '/messages/([^/?]+)') {
        $id = [Uri]::UnescapeDataString($Matches[1])
        return $graphState.messages[$id]
    }
    if ($Method -eq 'POST' -and $Uri -match '/messages/([^/]+)/send$') {
        $id = [Uri]::UnescapeDataString($Matches[1])
        [void]$graphState.sentOrder.Add($id)
        return @{}
    }
    throw "Unexpected Graph mock request: $Method $Uri"
}

$graphPayload = @{
    eventType = 'QC_RECEIVED'
    documentName = 'sheet-graph.pdf'
    subject = 'Graph root'
    body = 'Root body'
    htmlBody = '<p>Root body</p>'
    to = @('graph-test@company.com')
    cc = @()
    threadSendMode = 'root'
    threadKey = 'graph-pkg|Production QC'
    parentGraphMessageId = ''
}
$graphSettings = @{
    tenantId = 'tenant'
    clientId = 'client'
    clientSecret = 'secret'
    senderMailbox = 'qc@company.com'
}
$gRoot = Send-QCNotificationGraph -GraphSettings $graphSettings -Payload $graphPayload
Assert-True $gRoot.IsSuccess 'Graph root send should succeed'
Assert-Eq $gRoot.Data.sendMode 'root' 'Graph root sendMode'
Assert-True (-not [string]::IsNullOrWhiteSpace($gRoot.Data.graphMessageId)) 'Graph root should return message id'

$graphPayload.threadSendMode = 'reply'
$graphPayload.parentGraphMessageId = $gRoot.Data.graphImmutableMessageId
$gReply = Send-QCNotificationGraph -GraphSettings $graphSettings -Payload $graphPayload
if (-not $gReply.IsSuccess) { Write-Host "Graph reply failed: $($gReply.Message)" -ForegroundColor Red }
Assert-True $gReply.IsSuccess "Graph reply send should succeed ($($gReply.Message))"
Assert-Eq $gReply.Data.sendMode 'reply' 'Graph reply sendMode'
Assert-Eq $global:QCTestGraphState.sentOrder.Count 2 'Graph should have sent root and reply'
Clear-QCNotificationGraphTestHttpHandler

# 11. Graph replacement root when parent missing
$global:QCTestGraphState = @{
    messages = @{}
    sentOrder = [System.Collections.Generic.List[string]]::new()
}
Set-QCNotificationGraphTestHttpHandler -Handler {
    param($Method, $Uri, $Headers, $Body)
    $graphState = $global:QCTestGraphState
    if ($Uri -match 'oauth2/v2.0/token') { return @{ access_token = 'test-token' } }
    if ($Method -eq 'GET' -and $Uri -match '/messages/') { throw 'ErrorItemNotFound' }
    if ($Method -eq 'POST' -and $Uri -match '/messages$') {
        $id = 'graph-replacement-1'
        $msg = @{ id = $id; conversationId = 'conv-repl'; internetMessageId = "<$id@test.local>" }
        $graphState.messages[$id] = $msg
        return $msg
    }
    if ($Method -eq 'POST' -and $Uri -match '/createReply$') { throw 'ErrorItemNotFound' }
    if ($Method -eq 'POST' -and $Uri -match '/sendMail$') { return @{} }
    if ($Method -eq 'POST' -and $Uri -match '/send$') { return @{} }
    if ($Method -eq 'PATCH') { return @{} }
    throw "Unexpected: $Method $Uri"
}
$replPayload = @{
    eventType = 'QC_RECEIVED'
    documentName = 'sheet.pdf'
    subject = 'Replacement'
    body = 'Body'
    htmlBody = '<p>Body</p>'
    to = @('a@b.com')
    cc = @()
    threadSendMode = 'reply'
    parentGraphMessageId = 'missing-parent-id'
}
$gRepl = Send-QCNotificationGraph -GraphSettings $graphSettings -Payload $replPayload
Assert-True $gRepl.IsSuccess 'Replacement root should succeed'
Assert-Eq $gRepl.Data.sendMode 'replacement_root' 'Should recover with replacement_root'
Assert-Eq $gRepl.Data.threadRecoveryReason 'parent_message_not_found' 'Missing parent should use not_found reason'
Assert-Eq $gRepl.Data.parentLookupStatus 'not_found' 'Missing parent lookup status'
Clear-QCNotificationGraphTestHttpHandler

# 11b. Graph replacement root when parent lookup is forbidden (Mail.Read missing)
$global:QCTestGraphState = @{
    messages = @{}
    sentOrder = [System.Collections.Generic.List[string]]::new()
}
Set-QCNotificationGraphTestHttpHandler -Handler {
    param($Method, $Uri, $Headers, $Body)
    if ($Uri -match 'oauth2/v2.0/token') { return @{ access_token = 'test-token' } }
    if ($Method -eq 'GET' -and $Uri -match '/messages/') {
        throw 'The remote server returned an error: (403) Forbidden.'
    }
    if ($Method -eq 'POST' -and $Uri -match '/sendMail$') { return @{} }
    throw "Unexpected: $Method $Uri"
}
$forbiddenPayload = @{
    eventType = 'QC_RECEIVED'
    documentName = 'sheet.pdf'
    subject = 'Forbidden parent lookup'
    body = 'Body'
    htmlBody = '<p>Body</p>'
    to = @('a@b.com')
    cc = @()
    threadSendMode = 'reply'
    parentGraphMessageId = 'parent-without-read-access'
    threadKey = 'pkg|Review'
}
$gForbidden = Send-QCNotificationGraph -GraphSettings $graphSettings -Payload $forbiddenPayload
Assert-True $gForbidden.IsSuccess 'Forbidden parent lookup should still send replacement_root'
Assert-Eq $gForbidden.Data.sendMode 'replacement_root' 'Forbidden parent should recover with replacement_root'
Assert-Eq $gForbidden.Data.threadRecoveryReason 'parent_message_read_forbidden' 'Forbidden parent should use read_forbidden reason'
Assert-Eq $gForbidden.Data.parentLookupStatus 'access_denied' 'Forbidden parent lookup status'
Assert-Eq $gForbidden.Data.parentLookupHttpStatus 403 'Forbidden parent HTTP status'
Clear-QCNotificationGraphTestHttpHandler

# 12. HTML unthreaded uses sendMail (Mail.Send only)
$global:QCTestGraphState = @{
    sendMailCalls = 0
    createMessageCalls = 0
}
Set-QCNotificationGraphTestHttpHandler -Handler {
    param($Method, $Uri, $Headers, $Body)
    if ($Uri -match 'oauth2/v2.0/token') { return @{ access_token = 'test-token' } }
    if ($Method -eq 'POST' -and $Uri -match '/sendMail$') {
        $global:QCTestGraphState.sendMailCalls++
        return @{}
    }
    if ($Method -eq 'POST' -and $Uri -match '/messages$') {
        $global:QCTestGraphState.createMessageCalls++
        throw 'The remote server returned an error: (403) Forbidden.'
    }
    throw "Unexpected Graph mock request: $Method $Uri"
}
$htmlPayload = @{
    eventType = 'EMAIL_TEMPLATE_TEST'
    documentName = 'sheet.pdf'
    subject = 'HTML sendMail'
    body = 'plain'
    htmlBody = '<p>HTML body</p>'
    logoPath = 'email/typsalogo.png.webp'
    to = @('html-test@company.com')
    cc = @()
    threadSendMode = 'unthreaded'
    threadKey = ''
}
$gHtml = Send-QCNotificationGraph -GraphSettings $graphSettings -Payload $htmlPayload
Assert-True $gHtml.IsSuccess "HTML unthreaded send should succeed ($($gHtml.Message))"
Assert-Eq $global:QCTestGraphState.sendMailCalls 1 'Unthreaded HTML should use sendMail'
Assert-Eq $global:QCTestGraphState.createMessageCalls 0 'Unthreaded HTML should not create draft messages'
Assert-Eq $gHtml.Data.threadWarning 'html_sendmail_no_message_id' 'Unthreaded HTML should note missing thread id'
Clear-QCNotificationGraphTestHttpHandler

# 13. Threaded root falls back to sendMail when create message is denied
$global:QCTestGraphState = @{
    sendMailCalls = 0
    createMessageCalls = 0
}
Set-QCNotificationGraphTestHttpHandler -Handler {
    param($Method, $Uri, $Headers, $Body)
    if ($Uri -match 'oauth2/v2.0/token') { return @{ access_token = 'test-token' } }
    if ($Method -eq 'POST' -and $Uri -match '/sendMail$') {
        $global:QCTestGraphState.sendMailCalls++
        return @{}
    }
    if ($Method -eq 'POST' -and $Uri -match '/messages$') {
        $global:QCTestGraphState.createMessageCalls++
        throw 'The remote server returned an error: (403) Forbidden.'
    }
    throw "Unexpected Graph mock request: $Method $Uri"
}
$fallbackPayload = @{
    eventType = 'QC_RECEIVED'
    documentName = 'sheet.pdf'
    subject = 'Fallback root'
    body = 'Body'
    htmlBody = '<p>Body</p>'
    to = @('fallback@company.com')
    cc = @()
    threadSendMode = 'root'
    threadKey = 'graph-pkg|Production QC'
    parentGraphMessageId = ''
}
$gFallback = Send-QCNotificationGraph -GraphSettings $graphSettings -Payload $fallbackPayload
Assert-True $gFallback.IsSuccess "Threaded root should fall back to sendMail ($($gFallback.Message))"
Assert-Eq $global:QCTestGraphState.sendMailCalls 1 'Denied create should fall back to sendMail'
Assert-Eq $global:QCTestGraphState.createMessageCalls 1 'Threaded root should attempt create first'
Assert-Eq $gFallback.Data.threadWarning 'create_message_denied_fallback_sendmail' 'Fallback warning should be set'
Clear-QCNotificationGraphTestHttpHandler

Write-Host 'test_qc_notification_threading.ps1: PASS' -ForegroundColor Green

# Optional SQL integration: concurrent thread create + DB tables
$appSettings = Join-Path $repoRoot 'appsettings.json'
if (Test-Path -LiteralPath $appSettings) {
    Import-Module "$repoRoot/modules/Database/Core.Database.psm1" -Force
    try { $dbConfig = Get-QCAppSettingsConfig -Path $appSettings } catch { $dbConfig = $null }
    if ($dbConfig -and (Test-QCDatabaseEnabled -Config $dbConfig)) {
        $connProbe = Get-QCDatabaseConnection -Config $dbConfig
        if ($connProbe.IsSuccess) {
            try { $connProbe.Data.connection.Close(); $connProbe.Data.connection.Dispose() } catch { }
            $schemaRes = Initialize-QCDatabaseSchema -Config $dbConfig
            if ($schemaRes.IsSuccess) {
                $markerPkg = [guid]::NewGuid()
                $markerReview = 'Production QC'
                $t1 = New-QCNotificationThread -Config $dbConfig -SheetPackageId $markerPkg.ToString() -ReviewType $markerReview `
                    -GraphMessageId 'int-msg-1' -GraphImmutableMessageId 'int-msg-1' -GraphConversationId 'int-conv-1'
                $t2 = New-QCNotificationThread -Config $dbConfig -SheetPackageId $markerPkg.ToString() -ReviewType $markerReview `
                    -GraphMessageId 'int-msg-2' -GraphImmutableMessageId 'int-msg-2' -GraphConversationId 'int-conv-1'
                Assert-True $t1.IsSuccess 'Concurrent thread create first should succeed'
                Assert-True $t2.IsSuccess 'Concurrent thread create second should succeed'
                $active = Get-QCNotificationActiveThread -Config $dbConfig -SheetPackageId $markerPkg.ToString() -ReviewType $markerReview
                Assert-True ($null -ne $active) 'Active thread should exist'
                $cnt = Invoke-QCDatabaseScalar -Config $dbConfig -Sql @"
SELECT COUNT(*) FROM qc_notification_threads
WHERE sheet_package_id = @p AND review_type = @r AND status = 'active'
"@ -Parameters @{ p = $markerPkg; r = $markerReview }
                Assert-Eq ([int]$cnt.Data.value) 1 'At most one active thread per package+review type'
                $null = Invoke-QCDatabaseNonQuery -Config $dbConfig -Sql @"
DELETE FROM qc_notification_messages WHERE thread_id IN (
  SELECT id FROM qc_notification_threads WHERE sheet_package_id = @p AND review_type = @r
);
DELETE FROM qc_notification_threads WHERE sheet_package_id = @p AND review_type = @r;
"@ -Parameters @{ p = $markerPkg; r = $markerReview }
                Write-Host 'SQL integration: PASS' -ForegroundColor Green
            }
        }
    }
}

try { Remove-Item -Recurse -Force $testRoot -ErrorAction SilentlyContinue } catch { }
