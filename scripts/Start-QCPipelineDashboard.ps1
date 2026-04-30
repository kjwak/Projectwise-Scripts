<#
.SYNOPSIS
Unified pipeline entrypoint + live terminal dashboard.

.DESCRIPTION
Runs Watch-QCTrigger.ps1 (enqueue) and Run-QCProcessor.ps1 (dequeue/process) in a loop,
while rendering a constantly-updating dashboard with color coding.

Architecture:
- Single stable renderer prints the full layout every frame (no conditional partial renders).
- Watcher/worker stdout is ingested into shared state; render only reads state.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath,

    [Parameter(Mandatory = $false)]
    [int]$PollSeconds = 0,

    [Parameter(Mandatory = $false)]
    [int]$MaxJobsPerPoll = 50,

    [Parameter(Mandatory = $false)]
    [int]$RecentJobs = 12,

    [Parameter(Mandatory = $false)]
    [int]$RecentErrors = 8,

    [Parameter(Mandatory = $false)]
    [int]$Workers = 0,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force -WarningAction SilentlyContinue

function _Resolve-AppSettingsPath([string]$ProvidedPath) {
    if (-not [string]::IsNullOrWhiteSpace($ProvidedPath)) {
        if (Test-Path -LiteralPath $ProvidedPath) { return (Resolve-Path -LiteralPath $ProvidedPath).Path }
        throw "appsettings.json not found: $ProvidedPath"
    }

    $candidates = @(
        (Join-Path $PSScriptRoot 'appsettings.json'),
        (Join-Path (Split-Path -Parent $PSScriptRoot) 'appsettings.json'),
        (Join-Path $repoRoot 'appsettings.json'),
        (Join-Path (Get-Location).Path 'appsettings.json')
    ) | Select-Object -Unique

    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c)) { return (Resolve-Path -LiteralPath $c).Path }
    }

    throw "appsettings.json not found. Looked in: $($candidates -join ', ')"
}

function _ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.ValueType]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) { return $Value }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $out = @()
        foreach ($i in $Value) { $out += (_ToHashtable $i) }
        return $out
    }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = (_ToHashtable $p.Value) }
        return $h
    }
    return $Value
}

function _Read-AppSettings([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { throw "appsettings.json not found: $Path" }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $obj = $raw | ConvertFrom-Json -ErrorAction Stop
    $cfg = [hashtable](_ToHashtable $obj)
    if (-not $cfg.ContainsKey('dryRun')) { $cfg['dryRun'] = $false }
    if ($DryRun.IsPresent) { $cfg['dryRun'] = $true }
    return $cfg
}

function _ColorForState([string]$State) {
    switch -Regex ($State) {
        '(?i)^pending$' { return 'Yellow' }
        '(?i)^running$' { return 'Cyan' }
        '(?i)^succeeded$' { return 'Green' }
        '(?i)^failed$' { return 'Red' }
        default { return 'Gray' }
    }
}

function _ColorForLevel([string]$Level) {
    $l = [string]$Level
    switch -Regex ($l) {
        '(?i)^error$' { return 'Red' }
        '(?i)^warning$' { return 'Yellow' }
        '(?i)^information$' { return 'Gray' }
        default { return 'Gray' }
    }
}

function _Draw-Bar([string]$Label, [int]$Value, [string]$Color) {
    Write-Host ("{0,-10} {1,6}" -f $Label, $Value) -ForegroundColor $Color
}

function _Get-TerminalWidth {
    try { return [int]$Host.UI.RawUI.WindowSize.Width } catch { return 120 }
}

function _Trunc([string]$Text, [int]$Max) {
    if ($null -eq $Text) { return '' }
    $t = [string]$Text
    if ($Max -lt 4) { return $t }
    if ($t.Length -le $Max) { return $t }
    return ($t.Substring(0, $Max - 1) + '…')
}

function _Parse-UtcIso([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    try {
        # RoundtripKind preserves the embedded Z/+offset and avoids local-time assumptions.
        return [DateTime]::Parse(
            $Value,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind
        ).ToUniversalTime()
    } catch {
        return $null
    }
}

function _Format-UiTs([string]$IsoOrNull) {
    $dt = _Parse-UtcIso -Value $IsoOrNull
    if (-not $dt) { return '' }
    return $dt.ToString('yy-MM-dd HH:mm:ss')
}

function _SafeAscii([string]$Text) {
    if ($null -eq $Text) { return '' }
    # Replace non-ASCII chars so dashboard columns don't glitch with PW oddities.
    return ([regex]::Replace([string]$Text, '[^\u0020-\u007E]', '?'))
}

function _FormatJobLine([object]$Job, [int]$ColWidth) {
    if (-not $Job) { return '' }
    $ts = _Format-UiTs -IsoOrNull ([string]$Job.lastWriteTimeUtc)
    $state = [string]$Job.state
    $id = [string]$Job.id
    $name = ''
    $path = ''
    try {
        if ($Job.job) {
            $j = $Job.job
            if ($j.sourceName) { $name = [string]$j.sourceName }
            elseif ($j.sourcePath) { $name = [System.IO.Path]::GetFileName([string]$j.sourcePath) }
            if ($j.sourcePath) { $path = [string]$j.sourcePath }
            elseif ($j.sourceFolder -and $name) { $path = ([string]$j.sourceFolder).TrimEnd('\') + '\' + $name }
        }
    } catch { }
    $detail = ''
    if ($name -and $path) { $detail = "$name  $path" }
    elseif ($path) { $detail = $path }
    elseif ($name) { $detail = $name }
    else { $detail = $id }
    return _Trunc -Text ("{0}  {1,-9}  {2}" -f $ts, $state, (_SafeAscii $detail)) -Max $ColWidth
}

function _Write-TwoColumns([string]$Left, [string]$LeftColor, [string]$Right, [string]$RightColor, [int]$ColWidth) {
    $l = _Trunc -Text $Left -Max $ColWidth
    $r = _Trunc -Text $Right -Max $ColWidth
    $pad = ' ' * ([Math]::Max(1, ($ColWidth - $l.Length) + 3))
    if ($l) { Write-Host $l -NoNewline -ForegroundColor $LeftColor } else { Write-Host (' ' * $ColWidth) -NoNewline }
    Write-Host $pad -NoNewline
    if ($r) { Write-Host $r -ForegroundColor $RightColor } else { Write-Host '' }
}

$AppSettingsPath = _Resolve-AppSettingsPath -ProvidedPath $AppSettingsPath
$watcher = Join-Path $repoRoot 'scripts\\Watch-QCTrigger.ps1'
$worker = Join-Path $repoRoot 'scripts\\Run-QCProcessor.ps1'
if (-not (Test-Path -LiteralPath $watcher)) { throw "Watcher script not found: $watcher" }
if (-not (Test-Path -LiteralPath $worker)) { throw "Worker script not found: $worker" }

$state = @{
    phase = 'starting'
    lastError = $null
    lastHeartbeatUtc = $null
    hasSeenPwScan = $false
    passCount = 0
    passStartedAtUtc = $null
    lastPassDurationMs = $null

    pwFolderTotal = 0
    pwFolderIndex = 0
    currentScanStage = ''
    recentScanFolders = New-Object System.Collections.Generic.List[string]

    scanRoot = ''
    scanProject = ''
    scanPath = ''

    queueStats = $null
    recentJobs = @()
    lastQueueRefreshUtc = $null

    lastWorkerEvent = $null
    workers = @{}
    workerSlotMax = 0
    errors = New-Object System.Collections.Generic.List[object]
}

function _State-PushError([object]$LogObj) {
    if (-not $LogObj) { return }
    try {
        $lvl = [string]$LogObj.level
        if ([string]::IsNullOrWhiteSpace($lvl)) { return }
        $l = $lvl.ToLowerInvariant()
        if ($l -in @('error', 'warning')) {
            $state.errors.Add($LogObj) | Out-Null
            while ($state.errors.Count -gt 200) { $state.errors.RemoveAt(0) }
        }
    } catch { }
}

function _State-SetScanContext([string]$FolderPath) {
    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return }
    $state.scanPath = $FolderPath
    try {
        # NOTE: -split uses regex; '\\' matches a single backslash.
        $parts = @(([string]$FolderPath) -split '\\' | Where-Object { $_ -ne '' })
        if ($parts.Count -ge 2 -and $parts[0].ToLowerInvariant() -eq 'documents') {
            # For PW paths, show: Documents\<root> and the immediate project folder (if present).
            $state.scanRoot = ('Documents\' + $parts[1])
            $state.scanProject = if ($parts.Count -ge 3) { [string]$parts[2] } else { '' }
        } else {
            # Local path (or unexpected shape): keep root/project blank but still show full path line.
            $state.scanRoot = ''
            $state.scanProject = ''
        }
    } catch { }
}

function _MaybeRefreshQueue([hashtable]$Cfg, [int]$MinIntervalMs = 1500) {
    try {
        $now = Get-Date
        if ($state.lastQueueRefreshUtc) {
            $age = ($now.ToUniversalTime() - [DateTime]::Parse([string]$state.lastQueueRefreshUtc)).TotalMilliseconds
            if ($age -lt $MinIntervalMs) { return }
        }
        $stats = Get-QCQueueStats -Config $Cfg
        if ($stats.IsSuccess) { $state.queueStats = $stats.Data }
        $recent = Get-QCRecentJobs -Config $Cfg -Limit ([Math]::Max(5, $RecentJobs))
        if ($recent.IsSuccess) { $state.recentJobs = @($recent.Data.jobs) }
        $state.lastQueueRefreshUtc = (Get-Date).ToUniversalTime().ToString('o')
    } catch { }
}

function _Render-Workers() {
    $max = [int]$state.workerSlotMax
    if ($max -le 0) { $max = $state.workers.Count }
    Write-Host ("Workers (max {0})" -f $max) -ForegroundColor White
    if (-not $state.workers -or $state.workers.Count -eq 0) {
        Write-Host '  (no workers spawned yet)' -ForegroundColor DarkGray
        return
    }
    $sorted = @($state.workers.Values | Sort-Object -Property label)
    foreach ($w in $sorted) {
        $st = [string]$w.state
        $color = switch -Regex ($st) {
            'RUNNING' { 'Cyan'; break }
            'EXITING' { 'DarkGray'; break }
            default   { 'Yellow' }
        }
        $elapsed = '-'
        if ($w.startedAtUtc) {
            try {
                $s = _Parse-UtcIso -Value ([string]$w.startedAtUtc)
                if ($s) { $elapsed = ("{0}s" -f [int]((Get-Date).ToUniversalTime() - $s).TotalSeconds) }
            } catch { }
        }
        $jobId = if ($w.jobId) { [string]$w.jobId } else { '-' }
        $jobType = if ($w.jobType) { [string]$w.jobType } else { '-' }
        Write-Host (
            "  {0,-4}  pid={1,-6}  {2,-8}  {3,-15}  {4,-36}  {5}" -f `
                [string]$w.label, [int]$w.pid, $st, $jobType, (_Trunc -Text $jobId -Max 36), $elapsed
        ) -ForegroundColor $color
    }
}

function _Render-RecentJobsTwoCol([object[]]$Jobs, [int]$Limit) {
    $width = _Get-TerminalWidth
    $colWidth = [int][Math]::Max(44, [Math]::Floor(($width - 6) / 2))

    $jobs = @($Jobs | Select-Object -First ($Limit * 2))
    # Get-QCRecentJobs returns entries shaped like: @{ state; job; lastWriteTimeUtc } (job is nested).
    $qcJobs = @($jobs | Where-Object { $_ -and $_.job -and ([string]$_.job.type) -ne 'STATUS_SET_GEN' } | Select-Object -First $Limit)
    $stJobs = @($jobs | Where-Object { $_ -and $_.job -and ([string]$_.job.type) -eq 'STATUS_SET_GEN' } | Select-Object -First $Limit)

    _Write-TwoColumns -Left 'QC (QC_PREPEND)' -LeftColor 'White' -Right 'Status (STATUS_SET_GEN)' -RightColor 'White' -ColWidth $colWidth
    for ($i = 0; $i -lt [Math]::Max($qcJobs.Count, $stJobs.Count); $i++) {
        $lJob = if ($i -lt $qcJobs.Count) { $qcJobs[$i] } else { $null }
        $rJob = if ($i -lt $stJobs.Count) { $stJobs[$i] } else { $null }
        $lText = _FormatJobLine -Job $lJob -ColWidth $colWidth
        $rText = _FormatJobLine -Job $rJob -ColWidth $colWidth
        $lColor = if ($lJob) { _ColorForState -State ([string]$lJob.state) } else { 'Gray' }
        $rColor = if ($rJob) { _ColorForState -State ([string]$rJob.state) } else { 'Gray' }
        _Write-TwoColumns -Left $lText -LeftColor $lColor -Right $rText -RightColor $rColor -ColWidth $colWidth
    }
}

function _Render-Full([hashtable]$Cfg) {
    Clear-Host
    $now = Get-Date
    $dry = $false
    try { $dry = [bool]$Cfg.dryRun } catch { $dry = $false }

    Write-Host ("QC Pipeline Dashboard   {0}   DryRun={1}" -f $now.ToString('yyyy-MM-dd HH:mm:ss'), $dry) -ForegroundColor Cyan
    Write-Host ("Config: {0}" -f $AppSettingsPath) -ForegroundColor DarkGray
    Write-Host ("Status: {0}" -f [string]$state.phase) -ForegroundColor Yellow
    Write-Host ("Root:   {0}" -f $(if ($state.scanRoot) { $state.scanRoot } else { '-' })) -ForegroundColor DarkGray
    Write-Host ("Proj:   {0}" -f $(if ($state.scanProject) { $state.scanProject } else { '-' })) -ForegroundColor DarkGray
    Write-Host ("Path:   {0}" -f (_Trunc -Text ($(if ($state.scanPath) { $state.scanPath } else { '-' })) -Max 170)) -ForegroundColor DarkGray
    $lastPass = ''
    if ($state.lastPassDurationMs) { $lastPass = ("Last pass: {0:0.0}s" -f ([double]$state.lastPassDurationMs / 1000.0)) }
    if ($state.passStartedAtUtc) {
        $elapsed = 0
        try {
            $start = _Parse-UtcIso -Value ([string]$state.passStartedAtUtc)
            if ($start) { $elapsed = [int]((Get-Date).ToUniversalTime() - $start).TotalSeconds } else { $elapsed = 0 }
        } catch { $elapsed = 0 }
        $suffix = if ($lastPass) { "  $lastPass" } else { '' }
        Write-Host ("Pass:   #{0}  Elapsed: {1}s{2}" -f [int]$state.passCount, $elapsed, $suffix) -ForegroundColor DarkGray
    } elseif ($lastPass) {
        Write-Host ("Pass:   #{0}  {1}" -f [int]$state.passCount, $lastPass) -ForegroundColor DarkGray
    }
    if ([int]$state.pwFolderTotal -gt 0) {
        Write-Host ("PW:     {0}/{1} folders" -f [int]$state.pwFolderIndex, [int]$state.pwFolderTotal) -ForegroundColor DarkGray
    }
    if ($state.currentScanStage) {
        Write-Host ("Stage:  {0}" -f (_Trunc -Text ([string]$state.currentScanStage) -Max 170)) -ForegroundColor DarkGray
    }
    if ($state.recentScanFolders -and $state.recentScanFolders.Count -gt 0) {
        $tail = @($state.recentScanFolders | Select-Object -Last 3)
        Write-Host ("Recent: {0}" -f (_Trunc -Text ($tail -join '  |  ') -Max 170)) -ForegroundColor DarkGray
    }
    if ($state.lastHeartbeatUtc) { Write-Host ("Heartbeat: {0}" -f (_Format-UiTs -IsoOrNull ([string]$state.lastHeartbeatUtc))) -ForegroundColor DarkGray }
    Write-Host ""

    Write-Host "Queue" -ForegroundColor White
    if ($state.queueStats) {
        $st = $state.queueStats.states
        _Draw-Bar -Label 'Pending'   -Value ([int]$st.pending)   -Color 'Yellow'
        _Draw-Bar -Label 'Running'   -Value ([int]$st.running)   -Color 'Cyan'
        _Draw-Bar -Label 'Succeeded' -Value ([int]$st.succeeded) -Color 'Green'
        _Draw-Bar -Label 'Failed'    -Value ([int]$st.failed)    -Color 'Red'
        Write-Host ("Locks: {0}" -f [int]$state.queueStats.locks.count) -ForegroundColor DarkGray
    } else {
        Write-Host "  (queue stats unavailable yet)" -ForegroundColor DarkGray
    }
    Write-Host ""

    _Render-Workers
    Write-Host ""

    Write-Host ("Recent jobs (last {0} per column)" -f $RecentJobs) -ForegroundColor White
    _Render-RecentJobsTwoCol -Jobs @($state.recentJobs) -Limit $RecentJobs
    Write-Host ""

    Write-Host "Processor activity" -ForegroundColor White
    if ($state.lastWorkerEvent) {
        $evt = $state.lastWorkerEvent
        $cc = _ColorForLevel -Level ([string]$evt.level)
        Write-Host ("  {0} {1} - {2}" -f [string]$evt.code, [string]$evt.level, [string]$evt.message) -ForegroundColor $cc
    } else {
        Write-Host "  (no worker activity yet)" -ForegroundColor DarkGray
    }
    Write-Host ""

    Write-Host ("Recent warnings/errors (last {0})" -f $RecentErrors) -ForegroundColor White
    $tail = @($state.errors | Select-Object -Last $RecentErrors)
    if ($tail.Count -eq 0) {
        Write-Host "  (none)" -ForegroundColor DarkGray
    } else {
        foreach ($e in $tail) {
            $cc = _ColorForLevel -Level ([string]$e.level)
            $folder = ''
            $errMsg = ''
            try { if ($e.data -and $e.data.folder) { $folder = [string]$e.data.folder } } catch { }
            try { if ($e.data -and $e.data.errorMessage) { $errMsg = [string]$e.data.errorMessage } } catch { }
            $suffix = ''
            if (-not [string]::IsNullOrWhiteSpace($folder)) { $suffix += ('  [' + (_Trunc -Text $folder -Max 80) + ']') }
            if (-not [string]::IsNullOrWhiteSpace($errMsg)) { $suffix += ('  ' + (_Trunc -Text $errMsg -Max 110)) }
            Write-Host ("  {0}  {1,-8}  {2}  {3}{4}" -f [string]$e.ts, [string]$e.level, [string]$e.code, [string]$e.message, $suffix) -ForegroundColor $cc
        }
    }

    if ($state.lastError) {
        Write-Host ""
        Write-Host "Last fatal error (will retry):" -ForegroundColor Red
        Write-Host ("  {0}" -f [string]$state.lastError) -ForegroundColor Red
    }
}

function _Get-WorkersConfig([hashtable]$Cfg) {
    $w = @{
        maxParallel      = 2
        maxJobsPerWorker = 25
        leaseSeconds     = 600
        idleSleepMs      = 750
        spawnStaggerMs   = 250
    }
    try {
        if ($Cfg.ContainsKey('workers') -and $Cfg.workers) {
            $src = $Cfg.workers
            if ($src -is [hashtable]) {
                foreach ($k in @('maxParallel','maxJobsPerWorker','leaseSeconds','idleSleepMs','spawnStaggerMs')) {
                    if ($src.ContainsKey($k) -and $src[$k] -ne $null) { $w[$k] = [int]$src[$k] }
                }
            }
        }
    } catch { }
    if ($Workers -gt 0) { $w.maxParallel = [int]$Workers }
    if ($w.maxParallel -lt 1) { $w.maxParallel = 1 }
    return $w
}

function _Build-ArgLine([string]$ScriptPath, [string[]]$ScriptArgs) {
    $childArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath) + @($ScriptArgs)
    return (($childArgs | ForEach-Object {
        $t = [string]$_
        if ($t -match '[\s"]') { return ('"' + ($t -replace '"', '\\"') + '"') }
        return $t
    }) -join ' ')
}

function _Get-ChildLogDir() {
    # Persistent log dir under the queue root so logs survive child-process death.
    # Used to diagnose AV-induced kills where workers vanish without a failure record.
    $logDir = $null
    try {
        if ($script:_QCDashLockPath) {
            $logDir = Join-Path (Split-Path -Parent $script:_QCDashLockPath) '_logs'
        }
    } catch { }
    if (-not $logDir) { $logDir = Join-Path $env:TEMP 'QC_Pipeline_Logs' }
    if (-not (Test-Path -LiteralPath $logDir)) {
        try { New-Item -ItemType Directory -Path $logDir -Force | Out-Null } catch { }
    }
    return $logDir
}

function _Start-Child([string]$ScriptPath, [string[]]$ScriptArgs) {
    # Persist child stdout/stderr under queue\_logs\ so we can read what a killed
    # process did/said even after the dashboard reaps it. Files are named with a
    # timestamp + script tag + PID-placeholder; PID is rewritten after spawn.
    $logDir = _Get-ChildLogDir
    $tag = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
    $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss_fff')
    $stdoutPath = Join-Path $logDir ("${stamp}_${tag}.out.log")
    $stderrPath = Join-Path $logDir ("${stamp}_${tag}.err.log")
    $argLine = _Build-ArgLine -ScriptPath $ScriptPath -ScriptArgs $ScriptArgs
    $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -NoNewWindow -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    return @{
        process = $p
        stdoutPath = $stdoutPath
        stderrPath = $stderrPath
        lastStdoutLen = 0
        lastHeartbeatAt = (Get-Date)
        scriptPath = $ScriptPath
    }
}

function _Stop-Child([hashtable]$Child) {
    # Intentionally do NOT delete the log files: they're the only forensic trail
    # when a worker is killed by AV and never gets to write a failure result.
    # A separate housekeeping step can prune _logs\ on age if needed.
}

function _Poll-Child([hashtable]$Child, [hashtable]$Cfg, [string]$Kind) {
    $p = $Child.process
    try { $p.Refresh() } catch { }

    $cur = ''
    try { $cur = [string](Get-Content -LiteralPath $Child.stdoutPath -Raw -ErrorAction SilentlyContinue) } catch { $cur = '' }
    if ($cur.Length -gt $Child.lastStdoutLen) {
        $delta = $cur.Substring($Child.lastStdoutLen)
        $Child.lastStdoutLen = $cur.Length

        foreach ($line in ($delta -split "(`r`n|`n|`r)")) {
            $t = ($line -as [string]).Trim()
            if (-not $t) { continue }
            if ($t.StartsWith('{')) {
                try {
                    $o = ($t | ConvertFrom-Json -ErrorAction Stop)
                    _State-PushError -LogObj $o

                    if ($Kind -eq 'worker' -and $o.code -match '^WORKER_') {
                        $state.lastWorkerEvent = $o
                        try {
                            $wlbl = ''
                            $wpid = 0
                            if ($o.data) {
                                if ($o.data.workerLabel) { $wlbl = [string]$o.data.workerLabel }
                                if ($o.data.workerPid) { $wpid = [int]$o.data.workerPid }
                            }
                            if (-not $wlbl) { $wlbl = "W?$wpid" }
                            if (-not $state.workers.ContainsKey($wlbl)) {
                                $state.workers[$wlbl] = @{
                                    label = $wlbl
                                    pid = $wpid
                                    jobId = ''
                                    jobType = ''
                                    state = 'IDLE'
                                    lastCode = [string]$o.code
                                    lastMessage = [string]$o.message
                                    startedAtUtc = $null
                                    updatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                                }
                            }
                            $w = $state.workers[$wlbl]
                            if ($wpid -gt 0) { $w.pid = $wpid }
                            $w.lastCode = [string]$o.code
                            $w.lastMessage = [string]$o.message
                            $w.updatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                            switch ([string]$o.code) {
                                'WORKER_START'   { $w.state = 'IDLE' }
                                'WORKER_SELECTED' {
                                    $w.state = 'RUNNING'
                                    if ($o.data) {
                                        if ($o.data.jobId)   { $w.jobId   = [string]$o.data.jobId }
                                        if ($o.data.jobType) { $w.jobType = [string]$o.data.jobType }
                                    }
                                    $w.startedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
                                }
                                'WORKER_SUCCEEDED' { $w.state = 'IDLE'; $w.jobId = ''; $w.jobType = ''; $w.startedAtUtc = $null }
                                'WORKER_FAILED'    { $w.state = 'IDLE'; $w.jobId = ''; $w.jobType = ''; $w.startedAtUtc = $null }
                                'WORKER_NO_JOB'    { $w.state = 'IDLE'; $w.jobId = ''; $w.jobType = '' }
                                'WORKER_LOCK_RACE' { $w.state = 'IDLE' }
                                'WORKER_BUDGET'    { $w.state = 'EXITING' }
                                'WORKER_LEASE'     { $w.state = 'EXITING' }
                            }
                        } catch { }
                    }

                    # Scan context: ANY PW event with a folder should update Root/Proj/Path
                    # (including error events where scanning fails early).
                    try {
                        if ($o.data -and $o.data.folder) {
                            _State-SetScanContext -FolderPath ([string]$o.data.folder)
                        }
                    } catch { }

                    # Phase transitions (keep it simple + monotonic enough to avoid "stuck connecting")
                    if (($o.code -as [string]) -match '^WATCH_PW_') { $state.hasSeenPwScan = $true }
                    if ($o.code -eq 'WATCH_PW_CONNECT_START') {
                        $state.phase = 'Connecting to ProjectWise...'
                        $state.currentScanStage = 'connecting to ProjectWise'
                    }
                    if ($o.code -eq 'WATCH_PW_CONNECT_OK') {
                        $state.phase = 'Scanning folders...'
                        $state.currentScanStage = 'connected; preparing folders'
                    }
                    if ($o.code -eq 'WATCH_PW_SCAN_START' -or $o.code -eq 'WATCH_PW_FOLDER_ERROR') { $state.phase = 'Scanning folders...' }
                    if ($o.code -eq 'WATCH_PW_DOC_SCAN') { $state.phase = 'Searching' }
                    if ($o.code -eq 'WATCH_ACCEPTED') { $state.phase = 'Searching' }

                    # Pass tracking + progress
                    if ($o.code -eq 'WATCH_START') {
                        # Use the event timestamp (already UTC) for consistent timing.
                        try { $state.passStartedAtUtc = [string]$o.ts } catch { $state.passStartedAtUtc = (Get-Date).ToUniversalTime().ToString('o') }
                        # Every WATCH_START is a new pass.
                        if ([int]$state.passCount -le 0) { $state.passCount = 1 } else { $state.passCount = ([int]$state.passCount + 1) }
                        $state.pwFolderTotal = 0
                        $state.pwFolderIndex = 0
                        $state.currentScanStage = 'watcher starting'
                        try { $state.recentScanFolders.Clear() } catch { }
                        # Do not override phase here; watcher may still be connecting.
                    }
                    if ($o.code -eq 'WATCH_PW_FOLDERS') {
                        try { $state.pwFolderTotal = [int]$o.data.folderCount } catch { $state.pwFolderTotal = 0 }
                        $state.pwFolderIndex = 0
                        $state.currentScanStage = "prepared $($state.pwFolderTotal) PW folders"
                        try { $state.recentScanFolders.Clear() } catch { }
                        # If we have a sample list, prime Root/Proj/Path immediately instead of waiting
                        # for the first WATCH_PW_SCAN_START event.
                        try {
                            $sample0 = $null
                            if ($o.data -and $o.data.sample) { $sample0 = @($o.data.sample | Select-Object -First 1)[0] }
                            if ($sample0) { _State-SetScanContext -FolderPath ([string]$sample0) }
                        } catch { }
                    }
                    if ($o.code -eq 'WATCH_PW_SCAN_START') {
                        $state.currentScanStage = "starting folder: $([string]$o.data.folder)"
                        # Make progress match what user sees: 1-based index, update on start.
                        if ([int]$state.pwFolderTotal -gt 0) {
                            $state.pwFolderIndex = [Math]::Min([int]$state.pwFolderTotal, ([int]$state.pwFolderIndex + 1))
                        } else {
                            $state.pwFolderIndex = [int]$state.pwFolderIndex + 1
                        }
                        try {
                            if ($o.data -and $o.data.folder) {
                                $state.recentScanFolders.Add([string]$o.data.folder) | Out-Null
                                while ($state.recentScanFolders.Count -gt 20) { $state.recentScanFolders.RemoveAt(0) }
                            }
                        } catch { }
                    }
                    if ($o.code -eq 'WATCH_PW_STATUSSET_SCAN_START') { $state.currentScanStage = "querying status set: $([string]$o.data.folder)" }
                    if ($o.code -eq 'WATCH_PW_STATUSSET_SCAN_DONE') { $state.currentScanStage = "status set done: $([string]$o.data.folder) ($([int]$o.data.pairedCount) pairs)" }
                    if ($o.code -eq 'WATCH_PW_DOC_SCAN_START') { $state.currentScanStage = "querying documents: $([string]$o.data.folder)" }
                    if ($o.code -eq 'WATCH_PW_DOC_SCAN') { $state.currentScanStage = "documents done: $([string]$o.data.folder) ($([int]$o.data.pdfCount) PDFs, $([int]$o.data.qcArchivistCount) tagged)" }
                    if ($o.code -eq 'WATCH_PW_FOLDER_DONE' -or $o.code -eq 'WATCH_PW_FOLDER_ERROR') {
                        $state.currentScanStage = if ($o.code -eq 'WATCH_PW_FOLDER_ERROR') { "folder error: $([string]$o.data.folder)" } else { "folder done: $([string]$o.data.folder)" }
                    }
                    if ($o.code -eq 'WATCH_DONE') {
                        try {
                            if ($state.passStartedAtUtc) {
                                $endTs = _Parse-UtcIso -Value ([string]$o.ts)
                                if (-not $endTs) { $endTs = (Get-Date).ToUniversalTime() }
                                $startTs = _Parse-UtcIso -Value ([string]$state.passStartedAtUtc)
                                if (-not $startTs) { $startTs = $endTs }
                                $ms = [int]($endTs - $startTs).TotalMilliseconds
                                $state.lastPassDurationMs = $ms
                            }
                        } catch { }
                        # Mark the pass as completed; keep last-pass info visible until next WATCH_START.
                        $state.passStartedAtUtc = $null
                        $state.pwFolderIndex = 0
                        $state.currentScanStage = 'watch pass completed'
                    }
                } catch { }
            }
        }
    }

    $now = Get-Date
    if ((($now - $Child.lastHeartbeatAt).TotalMilliseconds -ge 800)) {
        $Child.lastHeartbeatAt = $now
        $state.lastHeartbeatUtc = $now.ToUniversalTime().ToString('o')
        _MaybeRefreshQueue -Cfg $Cfg -MinIntervalMs 1500
        _Render-Full -Cfg $Cfg
    }
}

# Singleton guard: refuse to start if another dashboard is already running. Multiple
# concurrent dashboards multiply the watcher + worker process count and overwhelm
# Fortinet (which then kills processes mid-job, leaving orphan running\ jobs and
# empty output dirs). Use scripts\Stop-QCPipeline.ps1 to clean up stale instances.
$bootCfg = _Read-AppSettings -Path $AppSettingsPath
$queueRoot = $null
try {
    if ($bootCfg.queue -and $bootCfg.queue.rootDir) { $queueRoot = [string]$bootCfg.queue.rootDir }
    elseif ($bootCfg.queue -and $bootCfg.queue.root) { $queueRoot = [string]$bootCfg.queue.root }
} catch { }
if (-not $queueRoot) { $queueRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'queue' }
if (-not (Test-Path -LiteralPath $queueRoot)) { New-Item -ItemType Directory -Path $queueRoot -Force | Out-Null }

$dashLock = Join-Path $queueRoot '_dashboard.lock'
$alreadyRunning = $false
if (Test-Path -LiteralPath $dashLock) {
    try {
        $pl = Get-Content -LiteralPath $dashLock -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($pl -and $pl.pid) {
            $existing = Get-Process -Id ([int]$pl.pid) -ErrorAction SilentlyContinue
            if ($existing) { $alreadyRunning = $true }
        }
    } catch { }
}
if ($alreadyRunning) {
    Write-Host "Another dashboard is already running (see $dashLock)." -ForegroundColor Red
    Write-Host "Stop it first with: .\scripts\Stop-QCPipeline.ps1" -ForegroundColor Yellow
    exit 2
}
@{ pid = $PID; startedAtUtc = ([DateTime]::UtcNow.ToString('o')); host = $env:COMPUTERNAME } |
    ConvertTo-Json -Compress |
    Set-Content -LiteralPath $dashLock -Encoding utf8 -Force

# Best-effort cleanup of the singleton lock when this dashboard exits.
$script:_QCDashLockPath = $dashLock
[System.AppDomain]::CurrentDomain.add_ProcessExit({
    try { if ($script:_QCDashLockPath -and (Test-Path -LiteralPath $script:_QCDashLockPath)) { Remove-Item -LiteralPath $script:_QCDashLockPath -Force -ErrorAction SilentlyContinue } } catch { }
})

# One-time stale-job recovery: requeue or fail-over any orphan running\ jobs from a
# prior crashed dashboard/worker session. Safe (per-job lock files used).
# Also clear any stale watcher-active flag left over from older builds that used
# the global gate; current build processes STATUS_SET_GEN as soon as enqueued.
try {
    $rec = Recover-QCStaleJobs -Config $bootCfg
    if ($rec -and $rec.IsSuccess -and $rec.Data) {
        $state.lastWorkerEvent = [pscustomobject]@{
            ts = (Get-Date).ToUniversalTime().ToString('o')
            level = 'Information'
            code = 'WORKER_RECOVERY'
            message = ("Stale recovery: requeued={0} failed={1}" -f [int]$rec.Data.recoveredToPending, [int]$rec.Data.recoveredToFailed)
            data = $rec.Data
        }
    }
    Clear-QCWatcherActive -Config $bootCfg | Out-Null
} catch { }

$workerSlots = @{}
$nextWorkerIndex = 1
$lastSpawnAt = [DateTime]::MinValue
$lastRecoveryAt = [DateTime]::UtcNow

function _Spawn-Worker([hashtable]$Cfg, [hashtable]$WC, [string]$Label) {
    $xArgs = @(
        '-AppSettingsPath', $AppSettingsPath,
        '-MaxJobs', [string]([int]$WC.maxJobsPerWorker),
        '-LeaseSeconds', [string]([int]$WC.leaseSeconds),
        '-IdleSleepMs', [string]([int]$WC.idleSleepMs),
        '-WorkerLabel', $Label
    )
    if ([bool]$Cfg.dryRun) { $xArgs += '-DryRun' }
    $newChild = _Start-Child -ScriptPath $worker -ScriptArgs $xArgs
    $newChild['label'] = $Label
    if (-not $state.workers.ContainsKey($Label)) {
        $state.workers[$Label] = @{
            label = $Label
            pid = [int]$newChild.process.Id
            jobId = ''
            jobType = ''
            state = 'STARTING'
            lastCode = 'WORKER_SPAWN'
            lastMessage = 'Worker process starting.'
            startedAtUtc = $null
            updatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        }
    } else {
        $state.workers[$Label].pid = [int]$newChild.process.Id
        $state.workers[$Label].state = 'STARTING'
    }
    return $newChild
}

while ($true) {
    try {
        $cfg = _Read-AppSettings -Path $AppSettingsPath
        $wc = _Get-WorkersConfig -Cfg $cfg
        $state.workerSlotMax = [int]$wc.maxParallel

        $hasPw = $false
        try { $hasPw = ($cfg.ContainsKey('projectWise') -and $cfg.projectWise -and $cfg.projectWise.ContainsKey('watchList') -and $cfg.projectWise.watchList) } catch { $hasPw = $false }

        if ($hasPw -and (-not $state.scanPath -or -not $state.scanRoot)) {
            try {
                $wl = $cfg.projectWise.watchList
                $seed = $null
                if ($wl -and $wl.roots) {
                    $r0 = @($wl.roots | Select-Object -First 1)[0]
                    if ($r0 -and $r0.path) { $seed = [string]$r0.path }
                }
                if (-not $seed -and $wl -and $wl.folders) {
                    $f0 = @($wl.folders | Select-Object -First 1)[0]
                    if ($f0 -and $f0.root -and $f0.path) { $seed = ('Documents\' + [string]$f0.root + '\' + [string]$f0.path) }
                }
                if ($seed) { _State-SetScanContext -FolderPath $seed }
            } catch { }
        }

        if ($hasPw -and -not [bool]$state.hasSeenPwScan -and -not $state.passStartedAtUtc) {
            $state.phase = 'Connecting to ProjectWise...'
        } elseif (-not $hasPw) {
            $state.phase = 'Searching'
        } elseif ($state.phase -match 'Connecting') {
            $state.phase = 'Searching'
        }
        $state.lastError = $null

        _MaybeRefreshQueue -Cfg $cfg -MinIntervalMs 0
        _Render-Full -Cfg $cfg

        $wArgs = @('-AppSettingsPath', $AppSettingsPath)
        if ([bool]$cfg.dryRun) { $wArgs += '-DryRun' }

        $watcherChild = _Start-Child -ScriptPath $watcher -ScriptArgs $wArgs

        # Unified poll loop: watcher and worker pool run concurrently.
        # STATUS_SET_GEN and QC_PREPEND are both processed as soon as they're
        # enqueued. Per-folder STATUS_SET_GEN jobs only land in the queue after the
        # watcher has finished analyzing that folder, so picking them up immediately
        # is safe. Continue until watcher exited AND pool empty AND queue drained.
        while ($true) {
            $watcherAlive = $false
            try { $watcherAlive = -not $watcherChild.process.HasExited } catch { $watcherAlive = $false }
            if ($watcherAlive) {
                _Poll-Child -Child $watcherChild -Cfg $cfg -Kind 'watcher'
            }

            $deadLabels = @()
            foreach ($lbl in @($workerSlots.Keys)) {
                $wch = $workerSlots[$lbl]
                $alive = $false
                try { $alive = -not $wch.process.HasExited } catch { $alive = $false }
                if ($alive) {
                    _Poll-Child -Child $wch -Cfg $cfg -Kind 'worker'
                } else {
                    try { _Poll-Child -Child $wch -Cfg $cfg -Kind 'worker' } catch { }
                    _Stop-Child -Child $wch
                    # Drop reaped worker from the panel; only currently-spawned workers
                    # are shown. Sequential labels (W1, W2, ...) still appear in logs for
                    # cross-reference.
                    if ($state.workers.ContainsKey($lbl)) {
                        $state.workers.Remove($lbl) | Out-Null
                    }
                    $deadLabels += $lbl
                }
            }
            foreach ($lbl in $deadLabels) { $workerSlots.Remove($lbl) | Out-Null }

            _MaybeRefreshQueue -Cfg $cfg -MinIntervalMs 1500

            # Periodic stale-job recovery: every 30s, look for orphaned running\ jobs whose
            # owner-PID is dead (or has no lock file) and immediately requeue them. This
            # unsticks the queue without requiring a dashboard restart.
            if (([DateTime]::UtcNow - $lastRecoveryAt).TotalSeconds -ge 30) {
                $lastRecoveryAt = [DateTime]::UtcNow
                try {
                    $rec = Recover-QCStaleJobs -Config $cfg
                    if ($rec -and $rec.IsSuccess -and $rec.Data) {
                        $orph = 0; $req = 0; $fai = 0
                        try { $orph = [int]$rec.Data.recoveredOrphan } catch { }
                        try { $req  = [int]$rec.Data.recoveredToPending } catch { }
                        try { $fai  = [int]$rec.Data.recoveredToFailed } catch { }
                        if (($orph + $req + $fai) -gt 0) {
                            $state.lastWorkerEvent = [pscustomobject]@{
                                ts = (Get-Date).ToUniversalTime().ToString('o')
                                level = 'Information'
                                code = 'WORKER_RECOVERY'
                                message = ("Stale recovery: orphans={0} requeued={1} failed={2}" -f $orph, $req, $fai)
                                data = $rec.Data
                            }
                        }
                    }
                } catch { }
            }

            $pending = 0
            $running = 0
            try { $pending = [int]$state.queueStats.states.pending } catch { $pending = 0 }
            try { $running = [int]$state.queueStats.states.running } catch { $running = 0 }

            if ($workerSlots.Count -lt [int]$wc.maxParallel -and ($pending -gt 0)) {
                $now = Get-Date
                if (($now - $lastSpawnAt).TotalMilliseconds -ge [int]$wc.spawnStaggerMs) {
                    $label = "W$nextWorkerIndex"
                    $nextWorkerIndex++
                    try {
                        $newWorker = _Spawn-Worker -Cfg $cfg -WC $wc -Label $label
                        $workerSlots[$label] = $newWorker
                        $lastSpawnAt = $now
                    } catch {
                        $state.lastError = "Failed to spawn worker '$label': $($_.Exception.Message)"
                    }
                }
            }

            $watcherDone = -not $watcherAlive
            if ($watcherDone -and $workerSlots.Count -eq 0 -and $pending -le 0 -and $running -le 0) {
                break
            }

            Start-Sleep -Milliseconds 200
        }

        $exit = 0
        try { $exit = [int]$watcherChild.process.ExitCode } catch { $exit = 0 }
        _Stop-Child -Child $watcherChild
        if ($exit -ne 0) { throw "Watcher failed with exit code $exit" }

        $state.phase = 'Searching'
        _MaybeRefreshQueue -Cfg $cfg -MinIntervalMs 0
        _Render-Full -Cfg $cfg
        if ($PollSeconds -gt 0) { Start-Sleep -Seconds $PollSeconds }
    } catch {
        $state.lastError = [string]$_.Exception.Message
        $state.phase = 'ERROR (will retry)'
        try {
            $cfg = _Read-AppSettings -Path $AppSettingsPath
        } catch {
            $cfg = @{ dryRun = $false }
        }
        _Render-Full -Cfg $cfg
        if ($PollSeconds -gt 0) { Start-Sleep -Seconds ([Math]::Max(2, $PollSeconds)) }
    }
}

