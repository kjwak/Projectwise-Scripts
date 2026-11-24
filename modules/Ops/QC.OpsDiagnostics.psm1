Import-Module (Join-Path $PSScriptRoot 'QC.OpsConsole.psm1') -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot '..\Core\Core.Runtime.psm1') -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot '..\Core\QC.WatcherOrchestration.psm1') -DisableNameChecking

function Get-QCOpsSnapshot {
    param([hashtable]$Config, [switch]$IncludeQueue, [switch]$IncludeLogs)
    $status = (Get-QCOpsPipelineStatus -Config $Config).Data
    $root = $status.queueRoot
    $snapshot = @{
        status = $status; takenAt = [datetime]::UtcNow
        oldest = $null; holds = 0; queue = @(); logs = @()
        includeQueue = $IncludeQueue.IsPresent; includeLogs = $IncludeLogs.IsPresent
        scan = (Get-QCFullScanProgress -Config $Config -QueueRoot $root)
        schedule = @(Get-QCFullScanScheduleTimesFromConfig -Config $Config)
    }
    $pendingDir = Join-Path $root 'pending'
    if (Test-Path -LiteralPath $pendingDir) {
        $snapshot.oldest = Get-ChildItem -LiteralPath $pendingDir -Filter '*.json' -File |
            Sort-Object CreationTimeUtc | Select-Object -First 1
    }
    $holdDir = Join-Path $root '_prepend_holds'
    $holdFiles = @()
    if (Test-Path -LiteralPath $holdDir) {
        $holdFiles = @(Get-ChildItem -LiteralPath $holdDir -Filter '*.json' -File)
        $snapshot.holds = $holdFiles.Count
    }
    if ($IncludeQueue) {
        $prefer = @(Get-QCOpsPreferJobTypes -Config $Config)
        $snapshot.queue = @(foreach ($state in @('pending','running','failed','succeeded')) {
            foreach ($row in @(Get-QCOpsRunningJobRows -QueueRoot $root -State $state -Limit 250 -NewestPending -PreferJobTypes $prefer -Config $Config -IncludeDeferredStatusSet:($state -eq 'pending'))) {
                $row | Add-Member -NotePropertyName key -NotePropertyValue ("job:" + $row.jobId)
                $row
            }
        }
        foreach ($file in @($holdFiles | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 250)) {
            $hold = $null
            try { $hold = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json } catch { }
            [pscustomobject]@{
                key = 'hold:' + $file.BaseName; state = 'held'; jobId = [string]$hold.jobId; type = 'QC_PREPEND'
                sourceFolder = [string]$hold.lane.folder; sourceName = [string]$hold.lane.document
                disposition = 'lane_held'; errorDetail = 'Resolve the lane hold before attempting another upload.'
                lastWriteTimeUtc = $file.LastWriteTimeUtc; lastWriteTime = (Format-QCDisplayTime -Value $file.LastWriteTimeUtc)
                path = $null; detail = $hold
            }
        }) | Sort-Object lastWriteTimeUtc -Descending
    }
    if ($IncludeLogs) {
        $events = (Get-QCOpsRecentLogEvents -QueueRoot $root -MaxLines 250 -MaxFiles 12 -Unfiltered).events
        $snapshot.logs = @(ConvertTo-QCOpsLogRows -Events $events)
    }
    return $snapshot
}

function ConvertTo-QCOpsLogRows {
    param([object[]]$Events, [switch]$Ascending)
    $seen = @{}
    $rows = foreach ($event in $Events) {
        $time = [datetimeoffset]::MinValue
        if (-not [datetimeoffset]::TryParse([string]$event.ts, [ref]$time)) { continue }
        $code = [string]$event.code
        $data = $event.data
        $job = @($event.jobId,$event.job_id,$data.jobId,$data.job_id) | Where-Object { $_ } | Select-Object -First 1
        $document = @($event.documentGuid,$event.document_guid,$data.documentGuid,$data.document_guid,$data.docGuid,$data.documentName,$data.sourceName) |
            Where-Object { $_ } | Select-Object -First 1
        $key = @($event.logFile,$event.ts,$code,$job,$document,$event.message) -join '|'
        $seen[$key] = 1 + [int]$seen[$key]
        [pscustomobject]@{
            key = $key + '|' + $seen[$key]; when = $time.UtcDateTime; time = (Format-QCDisplayTime -Value $time.UtcDateTime)
            severity = [string]$event.level; source = [string]$event.logSource; code = $code
            job = [string]$job; document = [string]$document
            message = [string]$event.message -replace '\r?\n', ' '
            detail = $event
        }
    }
    $rows | Sort-Object @{ Expression = 'when'; Descending = (-not $Ascending) }, key
}

Export-ModuleMember -Function Get-QCOpsSnapshot, ConvertTo-QCOpsLogRows
