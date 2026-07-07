$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Paths.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\QC.StatusSetBatching.psm1') -Force

function Assert-True([bool]$Cond, [string]$Msg) {
    if (-not $Cond) { throw $Msg }
}

function Assert-Eq($Expected, $Actual, [string]$Msg) {
    if ("$Expected" -ne "$Actual") { throw "$Msg (expected='$Expected' actual='$Actual')" }
}

function New-TestStoreConfig([string]$StorePath, [hashtable]$Overrides = @{}) {
    $batch = @{
        enabled = $true
        intervalMinutes = 60
        maxFoldersPerRun = 100
        quietPeriodSeconds = 0
        staleWarningHours = 24
        dirtyFolderStorePath = $StorePath
        processOnWatcherStart = $false
    }
    foreach ($k in @($Overrides.Keys)) { $batch[$k] = $Overrides[$k] }
    return @{ statusSetBatching = $batch }
}

function Write-AgedDirtyStore {
    param(
        [string]$StorePath,
        [string[]]$Folders,
        [datetime]$FirstSeenUtc,
        [datetime]$LastSeenUtc
    )
    $entries = @{}
    foreach ($fp in @($Folders)) {
        $key = (Normalize-QCDocumentsFolderPath -Path $fp).Data.path
        $entries[$key] = @{
            folderPath = $fp
            folderKey = $key
            folderGuid = ('guid-' + $key)
            firstSeenUtc = $FirstSeenUtc.ToString('o')
            lastSeenUtc = $LastSeenUtc.ToString('o')
            eventCount = 1
            lastProcessedUtc = $null
            failureCount = 0
            lastError = $null
            oneLevelDeep = $true
            datasourceName = 'ds'
            triggerSource = 'audit_trail'
        }
    }
    $payload = @{
        version = 1
        lastBatchRunUtc = (Get-Date).ToUniversalTime().AddHours(-2).ToString('o')
        folders = $entries
    }
    $payload | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $StorePath -Encoding UTF8
}

$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('qc_ssb_' + [Guid]::NewGuid().ToString('n'))
$storePath = Join-Path $tmpRoot 'dirty-folders.json'
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null

$folder = 'Documents\AZDOT\TestProject\CADD\Sheets'
$config = New-TestStoreConfig -StorePath $storePath -Overrides @{ quietPeriodSeconds = 120 }

$statusRule = @{
    id = 'status-set-folder'
    enabled = $true
    priority = 50
    jobType = 'STATUS_SET_GEN'
    triggerType = 'pw'
    grouping = @{ enabled = $true; groupBy = 'folder' }
}

$capturedLogs = [System.Collections.Generic.List[object]]::new()
$logCb = {
    param($Code, $Message, $Data, $Level)
    [void]$capturedLogs.Add(@{ code = $Code; message = $Message; data = $Data; level = $Level })
}

try {
    $settings = Get-QCStatusSetBatchingSettings -Config $config -RepoRoot $repoRoot
    Assert-True ([bool]$settings.enabled) 'enabled default true'
    Assert-Eq 15 (Get-QCStatusSetBatchingSettings -Config @{} -RepoRoot $repoRoot).intervalMinutes 'default intervalMinutes is 15'
    Assert-Eq 24 $settings.staleWarningHours 'staleWarningHours default 24'
    Assert-Eq $storePath $settings.dirtyFolderStorePath 'store path resolved absolute'

    $queueResolved = Get-QCStatusSetBatchingSettings -Config @{
        queue = @{ rootDir = 'C:\QC_E2E_RealRun\queue' }
        statusSetBatching = @{ dirtyFolderStorePath = '_watcher/statusset-dirty-folders.json' }
    } -RepoRoot $repoRoot
    Assert-Eq 'C:\QC_E2E_RealRun\queue\_watcher\statusset-dirty-folders.json' $queueResolved.dirtyFolderStorePath 'relative path resolves under queue.rootDir'

    $m1 = Mark-StatusSetDirtyFolder -Config $config -FolderPath $folder -DatasourceName 'ds' -FolderGuid 'folder-guid-1' -RepoRoot $repoRoot -LogCallback $logCb
    Assert-True $m1.IsSuccess 'first mark success'
    Assert-True ([bool]$m1.Data.isNew) 'first mark is new'
    Assert-Eq 1 $m1.Data.dirtyFolderCount 'one dirty folder'

    $m2 = Mark-StatusSetDirtyFolder -Config $config -FolderPath $folder -DatasourceName 'ds' -RepoRoot $repoRoot -LogCallback $logCb
    Assert-True $m2.IsSuccess 'second mark success'
    Assert-True (-not [bool]$m2.Data.isNew) 'second mark not new'
    Assert-Eq 1 $m2.Data.dirtyFolderCount 'still one dirty folder'
    Assert-Eq 2 ([int]$m2.Data.entry.eventCount) 'eventCount increments'

    $reloaded = Get-Content -LiteralPath $storePath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($reloaded.folders.PSObject.Properties.Count -eq 1) 'store has one folder after restart read'
    $entryKey = @($reloaded.folders.PSObject.Properties.Name)[0]
    $entry = $reloaded.folders.$entryKey
    Assert-Eq 0 $entry.failureCount 'initial failureCount is 0'
    if ($null -ne $entry.lastError -and -not [string]::IsNullOrWhiteSpace([string]$entry.lastError)) {
        throw 'initial lastError should be null'
    }

    $disabledConfig = @{
        statusSetBatching = @{
            enabled = $false
            dirtyFolderStorePath = $storePath
        }
    }
    $m3 = Mark-StatusSetDirtyFolder -Config $disabledConfig -FolderPath $folder -RepoRoot $repoRoot
    Assert-True $m3.IsSuccess 'disabled mark returns success'
    Assert-True (-not [bool]$m3.Data.marked) 'disabled does not mark'

    $firstBatch = Invoke-StatusSetDirtyFolderBatch -Config $config -StatusRule $statusRule -RepoRoot $repoRoot -LogCallback $logCb
    Assert-True $firstBatch.IsSuccess 'first batch success'
    Assert-Eq 1 $firstBatch.Data.foldersSkippedQuietPeriod 'folder in quiet period on first batch'
    Assert-Eq 0 $firstBatch.Data.foldersProcessed 'no folders processed during quiet period'
    Assert-True ($firstBatch.Data.elapsedMs -ge 0) 'interval done includes elapsedMs'

    $notDue = Invoke-StatusSetDirtyFolderBatch -Config $config -StatusRule $statusRule -RepoRoot $repoRoot -LogCallback $logCb
    Assert-True $notDue.IsSuccess 'second batch call success'
    Assert-Eq 'STATUSSET_BATCH_NOT_DUE' $notDue.Code 'interval not due immediately after batch run'

    $afterQuiet = Get-Content -LiteralPath $storePath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($afterQuiet.folders.PSObject.Properties.Count -eq 1) 'folder kept during quiet period'

    $key = @($afterQuiet.folders.PSObject.Properties.Name)[0]
    $entry = $afterQuiet.folders.$key
    $entry.lastSeenUtc = (Get-Date).ToUniversalTime().AddMinutes(-10).ToString('o')
    $payload = @{
        version = 1
        lastBatchRunUtc = (Get-Date).ToUniversalTime().AddHours(-2).ToString('o')
        folders = @{ $key = $entry }
    }
    $payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $storePath -Encoding UTF8

    $capturedLogs.Clear()
    $failBatch = Invoke-StatusSetDirtyFolderBatch -Config $config -StatusRule $statusRule -RepoRoot $repoRoot -Force -LogCallback $logCb
    Assert-True $failBatch.IsSuccess 'eval-fail batch wrapper success'
    Assert-Eq 1 $failBatch.Data.foldersProcessed 'folder processed outside quiet period'
    Assert-Eq 1 $failBatch.Data.foldersFailed 'evaluation fails without PW'
    Assert-Eq 1 $failBatch.Data.foldersConsidered 'one folder considered'
    $doneLog = @($capturedLogs | Where-Object { $_.code -eq 'STATUSSET_BATCH_INTERVAL_DONE' } | Select-Object -Last 1)
    Assert-True ($doneLog.Count -eq 1) 'interval done log emitted'
    Assert-Eq 1 $doneLog[0].data.foldersFailed 'interval done foldersFailed'
    Assert-Eq 0 $doneLog[0].data.jobsQueued 'interval done jobsQueued'
    $afterFail = Get-Content -LiteralPath $storePath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($afterFail.folders.PSObject.Properties.Count -eq 1) 'folder kept after failed evaluation'
    $failedEntry = $afterFail.folders.$key
    Assert-Eq 1 ([int]$failedEntry.failureCount) 'failureCount incremented'
    Assert-True (-not [string]::IsNullOrWhiteSpace([string]$failedEntry.lastError)) 'lastError stored'

    $successOverride = {
        param($Cfg, $Fp, $Rule, $Ds, $OneLevel, $Trigger, $Dry)
        return New-QCSuccessResult -Code 'STATUSSET_FOLDER_EVALUATED' -Message 'stub success' -Data @{
            folderPath = $Fp
            evaluated = $true
            enqueued = $true
            duplicates = $false
            accepted = $true
            skippedReason = $null
            gateReason = $null
            pairedCount = 2
            jobId = 'job-stub-1'
        }
    }
    $successBatch = Invoke-StatusSetDirtyFolderBatch -Config $config -StatusRule $statusRule -RepoRoot $repoRoot -Force `
        -TestFolderEvaluationOverride $successOverride -LogCallback $logCb
    Assert-True $successBatch.IsSuccess 'success override batch'
    Assert-Eq 1 $successBatch.Data.foldersSucceeded 'success override succeeded'
    Assert-Eq 1 $successBatch.Data.jobsQueued 'success override queued'
    $afterSuccess = Get-Content -LiteralPath $storePath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-Eq 0 @($afterSuccess.folders.PSObject.Properties).Count 'successful evaluation clears dirty folder'

    $maxStore = Join-Path $tmpRoot 'max-folders.json'
    $maxConfig = New-TestStoreConfig -StorePath $maxStore -Overrides @{ maxFoldersPerRun = 2; quietPeriodSeconds = 0 }
    $folders = @(
        'Documents\AZDOT\TestProject\CADD\Sheets\A',
        'Documents\AZDOT\TestProject\CADD\Sheets\B',
        'Documents\AZDOT\TestProject\CADD\Sheets\C'
    )
    $aged = (Get-Date).ToUniversalTime().AddMinutes(-30)
    Write-AgedDirtyStore -StorePath $maxStore -Folders $folders -FirstSeenUtc $aged -LastSeenUtc $aged

    $maxBatch = Invoke-StatusSetDirtyFolderBatch -Config $maxConfig -StatusRule $statusRule -RepoRoot $repoRoot -Force `
        -TestFolderEvaluationOverride $successOverride -LogCallback $logCb
    Assert-Eq 2 $maxBatch.Data.foldersProcessed 'maxFoldersPerRun limits processed folders'
    Assert-Eq 1 $maxBatch.Data.remaining 'remaining dirty folder deferred'
    $maxAfter = Get-Content -LiteralPath $maxStore -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-Eq 1 @($maxAfter.folders.PSObject.Properties).Count 'one folder remains dirty'

    $staleStore = Join-Path $tmpRoot 'stale.json'
    $staleConfig = New-TestStoreConfig -StorePath $staleStore -Overrides @{ quietPeriodSeconds = 0; staleWarningHours = 24 }
    $staleFirst = (Get-Date).ToUniversalTime().AddHours(-30)
    $staleLast = (Get-Date).ToUniversalTime().AddMinutes(-30)
    Write-AgedDirtyStore -StorePath $staleStore -Folders @('Documents\AZDOT\Stale\CADD\Sheets') -FirstSeenUtc $staleFirst -LastSeenUtc $staleLast

    $capturedLogs.Clear()
    $staleBatch = Invoke-StatusSetDirtyFolderBatch -Config $staleConfig -StatusRule $statusRule -RepoRoot $repoRoot -Force `
        -TestFolderEvaluationOverride $successOverride -LogCallback $logCb
    Assert-True $staleBatch.IsSuccess 'stale batch success'
    $staleLog = @($capturedLogs | Where-Object { $_.code -eq 'STATUSSET_DIRTY_FOLDER_STALE' })
    Assert-True ($staleLog.Count -ge 1) 'stale warning emitted'
    Assert-True ($staleLog[0].data.ageHours -ge 24) 'stale ageHours reported'

    'ok'
} finally {
    Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
}
