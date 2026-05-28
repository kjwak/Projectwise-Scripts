<#
.SYNOPSIS
Unified pipeline entrypoint + live terminal dashboard.

.DESCRIPTION
Runs Watch-QCTrigger.ps1 (enqueue) and Run-QCProcessor.ps1 (dequeue/process) in a loop,
while rendering a constantly-updating dashboard with color coding.

Architecture:
- Single stable renderer prints the full layout every frame (no conditional partial renders).
- Watcher/worker stdout is ingested into shared state; render only reads state.
- After each Watch-QCTrigger run exits, the dashboard respawns it (inter-pass delay: -PollSeconds or 500ms for PW)
  while worker processes keep dequeuing jobs; the next scan does not wait for an empty queue.
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
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [switch]$SkipReconcileStatusSetsFirst,

    [Parameter(Mandatory = $false)]
    [ValidateSet('FullRedraw', 'DiffAnsi')]
    [string]$RenderMode = 'DiffAnsi',

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 60)]
    [int]$MaxFps = 8,

    # Bypass the singleton lock (normally only one dashboard per queue root). Use when
    # the lock file is wrong/stale or you intentionally need a second instance.
    [Parameter(Mandatory = $false)]
    [switch]$IgnoreSingletonLock
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force -WarningAction SilentlyContinue

function _Pause-IfInteractiveConsole {
    # Double-click / powershell.exe -File closes the window as soon as the script exits.
    # Pause only for the real console host so automation is not blocked.
    if ($Host.Name -ne 'ConsoleHost') { return }
    try {
        Write-Host ''
        Write-Host 'Press Enter to close this window...' -ForegroundColor Yellow
        $null = Read-Host
    } catch { }
}

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

function _Get-TerminalWidth {
    try { return [int]$Host.UI.RawUI.WindowSize.Width } catch { return 120 }
}

function _Trunc([string]$Text, [int]$Max) {
    if ($null -eq $Text) { return '' }
    $t = [string]$Text
    if ($Max -lt 4) { return $t }
    if ($t.Length -le $Max) { return $t }
    if ($Max -le 3) { return $t.Substring(0, $Max) }
    return ($t.Substring(0, $Max - 3) + '...')
}

function _Repeat-Char([string]$Char, [int]$Count) {
    if ($Count -le 0) { return '' }
    return $Char * $Count
}

function _PadRightVisible([string]$Text, [int]$Width) {
    $t = if ($null -eq $Text) { '' } else { [string]$Text }
    if ($t.Length -gt $Width) { $t = _Trunc -Text $t -Max $Width }
    return $t + (_Repeat-Char -Char ' ' -Count ($Width - $t.Length))
}

function _BoxLines([string]$Title, [string[]]$Content, [int]$Width) {
    $w = [Math]::Max(40, $Width)
    $inner = $w - 4
    $safeTitle = _SafeAscii $Title
    $titleText = if ($safeTitle) { ' ' + $safeTitle + ' ' } else { '' }
    if ($titleText.Length -gt ($w - 4)) { $titleText = ' ' + (_Trunc -Text $safeTitle -Max ($w - 6)) + ' ' }
    $topFill = [Math]::Max(0, $w - 2 - $titleText.Length)
    $rows = New-Object System.Collections.Generic.List[string]
    $rows.Add('+' + $titleText + (_Repeat-Char -Char '-' -Count $topFill) + '+') | Out-Null
    foreach ($line in @($Content)) {
        $rows.Add('| ' + (_PadRightVisible -Text (_SafeAscii $line) -Width $inner) + ' |') | Out-Null
    }
    $rows.Add('+' + (_Repeat-Char -Char '-' -Count ($w - 2)) + '+') | Out-Null
    return @($rows)
}

function _Progress-Bar([int]$Current, [int]$Total, [int]$Width = 20) {
    if ($Width -lt 4) { $Width = 4 }
    if ($Total -le 0) { return '[' + (_Repeat-Char -Char '-' -Count $Width) + ']' }
    $pct = [Math]::Max(0, [Math]::Min(1, ([double]$Current / [double]$Total)))
    $filled = [int][Math]::Round($pct * $Width)
    return '[' + (_Repeat-Char -Char '#' -Count $filled) + (_Repeat-Char -Char '-' -Count ($Width - $filled)) + ']'
}

function _Get-WorkerStageText([hashtable]$Worker) {
    if (-not $Worker) { return '-' }
    try { if ($Worker.stage) { return [string]$Worker.stage } } catch { }
    $code = [string]$Worker.lastCode
    $msg = [string]$Worker.lastMessage
    switch ($code) {
        'WORKER_SPAWN'          { return 'spawning process' }
        'WORKER_START'          { return 'started; polling queue' }
        'WORKER_SELECTED'       { return 'processing job' }
        'WORKER_STAGE'          { if ($msg) { return $msg }; return 'processing stage' }
        'WORKER_DRYRUN'         { return 'dry-run dispatch check' }
        'WORKER_DRYRUN_HANDLER' { return 'dry-run handler check' }
        'WORKER_MOVE_FAILED'    { return 'moving job state failed' }
        'WORKER_SUCCEEDED'      { return 'completed job; polling queue' }
        'WORKER_FAILED'         { return 'job failed; polling queue' }
        'WORKER_NO_JOB'         { return 'idle; waiting for pending jobs' }
        'WORKER_LOCK_RACE'      { return 'lock race; trying next job' }
        'WORKER_BUDGET'         { return 'max-jobs budget reached' }
        'WORKER_LEASE'          { return 'lease budget reached' }
        default                 { if ($msg) { return $msg }; return '-' }
    }
}

function _Get-LineColor([string]$Line) {
    $l = [string]$Line
    if ($l -match '^\+') { return 'DarkGray' }
    if ($l -match '(?i)\b(ERROR|failed|Fail [1-9]|WORKER_FAILED|MOVE_FAILED)\b') { return 'Red' }
    if ($l -match '(?i)\b(warning|WARN)\b') { return 'Yellow' }
    if ($l -match '(?i)\b(succeeded|SESSION ACTIVE|OK [1-9]|completed job)\b') { return 'Green' }
    if ($l -match '(?i)\b(RUNNING|processing|querying|scanning|audit|reconcil|CONNECTING|Run [1-9])\b') { return 'Cyan' }
    if ($l -match '(?i)\b(PENDING|Pend [1-9]|waiting|idle)\b') { return 'Yellow' }
    if ($l -match '^\| (QC Pipeline Dashboard|Status:|Watcher|Workers|Recent|Processor|Warnings)') { return 'White' }
    return 'Gray'
}

function _AnsiForColor([string]$Color) {
    switch ([string]$Color) {
        'Black'     { return '30' }
        'DarkRed'   { return '31' }
        'DarkGreen' { return '32' }
        'DarkYellow'{ return '33' }
        'DarkBlue'  { return '34' }
        'DarkMagenta'{ return '35' }
        'DarkCyan'  { return '36' }
        'Gray'      { return '37' }
        'DarkGray'  { return '90' }
        'Red'       { return '91' }
        'Green'     { return '92' }
        'Yellow'    { return '93' }
        'Blue'      { return '94' }
        'Magenta'   { return '95' }
        'Cyan'      { return '96' }
        'White'     { return '97' }
        default     { return '37' }
    }
}

function _Colorize-Line([string]$Line) {
    $esc = $script:_DashEsc
    $color = _AnsiForColor -Color (_Get-LineColor -Line $Line)
    return ("{0}[{1}m{2}{0}[0m" -f $esc, $color, $Line)
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
    return Format-QCTimestamp -IsoString $IsoOrNull
}

function _SafeAscii([string]$Text) {
    if ($null -eq $Text) { return '' }
    # Replace non-ASCII chars so dashboard columns don't glitch with PW oddities.
    return ([regex]::Replace([string]$Text, '[^\u0020-\u007E]', '?'))
}

# Returns @(segment1, segment2) immediately after a "Documents" path segment (PW-style roots).
function _Get-DocumentsAreaTwoSegments([string]$AnyPath) {
    if ([string]::IsNullOrWhiteSpace($AnyPath)) { return $null }
    $norm = ([string]$AnyPath).Trim() -replace '/', '\'
    $parts = @($norm -split '\\' | Where-Object { $_ -ne '' })
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ($parts[$i].Equals('documents', [StringComparison]::OrdinalIgnoreCase)) {
            if ($i + 2 -lt $parts.Count) {
                return @($parts[$i + 1], $parts[$i + 2])
            }
            return $null
        }
    }
    return $null
}

function _TryGet-ProjectNameFromFolder([hashtable]$Cfg, [string]$FolderPath) {
    # Extract "project name" from paths like:
    #   Documents\<AzDot root>\<project...>\CADD\Sheets
    # using appsettings watchList.roots[*].path + .sheetsPathFromProject + .projectDepth.
    # Works for any PW path under a configured watch root (status set or QC prepend).
    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return $null }
    if (-not $Cfg) { return $null }

    $norm = ([string]$FolderPath).Trim() -replace '/', '\'
    $norm = $norm.TrimEnd('\')

    $roots = $null
    try {
        if ($Cfg.ContainsKey('projectWise') -and $Cfg.projectWise -and
            $Cfg.projectWise.ContainsKey('watchList') -and $Cfg.projectWise.watchList -and
            $Cfg.projectWise.watchList.ContainsKey('roots') -and $Cfg.projectWise.watchList.roots) {
            $roots = @($Cfg.projectWise.watchList.roots)
        }
    } catch { $roots = $null }
    if (-not $roots) { return $null }

    foreach ($r in $roots) {
        $rootPath = $null
        $sheetsRel = $null
        $depth = 1
        try { if ($r.path) { $rootPath = [string]$r.path } } catch { $rootPath = $null }
        try { if ($r.sheetsPathFromProject) { $sheetsRel = [string]$r.sheetsPathFromProject } } catch { $sheetsRel = $null }
        try { if ($r.projectDepth) { $depth = [int]$r.projectDepth } } catch { $depth = 1 }
        if (-not $rootPath) { continue }
        if ($depth -lt 1) { $depth = 1 }

        $rootNorm = ([string]$rootPath).Trim() -replace '/', '\'
        $rootNorm = $rootNorm.TrimEnd('\')
        $prefix = $rootNorm + '\'
        if (-not ($norm.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase))) { continue }

        $rel = $norm.Substring($prefix.Length)
        if ([string]::IsNullOrWhiteSpace($rel)) { continue }

        if ($sheetsRel) {
            $suf = ([string]$sheetsRel).Trim() -replace '/', '\'
            $suf = $suf.Trim('\')
            if ($suf) {
                $suffix = '\' + $suf
                if ($rel.EndsWith($suffix, [StringComparison]::OrdinalIgnoreCase)) {
                    $rel = $rel.Substring(0, $rel.Length - $suffix.Length)
                }
            }
        }

        $rel = $rel.Trim('\')
        if ([string]::IsNullOrWhiteSpace($rel)) { continue }

        $parts = @($rel -split '\\' | Where-Object { $_ -ne '' })
        if ($parts.Count -le 0) { continue }
        $take = [Math]::Min([int]$depth, $parts.Count)
        $proj = ($parts | Select-Object -First $take) -join '\'
        if (-not [string]::IsNullOrWhiteSpace($proj)) { return $proj }
    }

    return $null
}

function _Format-QCPrependJobLine([object]$Entry, [int]$ColWidth) {
    if (-not $Entry) { return '' }
    $ts = _Format-UiTs -IsoOrNull ([string]$Entry.lastWriteTimeUtc)
    $state = [string]$Entry.state
    $name = ''
    $path = ''
    $id = ''
    try {
        if ($Entry.job) {
            $j = $Entry.job
            $id = [string]$j.id
            if ($j.sourceName) { $name = [string]$j.sourceName }
            elseif ($j.sourcePath) { $name = [System.IO.Path]::GetFileName([string]$j.sourcePath) }
            if ($j.sourcePath) { $path = [string]$j.sourcePath }
            elseif ($j.sourceFolder -and $name) {
                $sf = ([string]$j.sourceFolder) -replace '[\\/]+$', ''
                $path = $sf + '\' + $name
            }
        }
    } catch { }
    $name = _SafeAscii $name
    $path = _SafeAscii $path
    $segs = _Get-DocumentsAreaTwoSegments $path
    $root = if ($segs) { '\' + $segs[0] + '\' + $segs[1] } else { '' }
    if (-not $root -and $path) {
        try { $root = _SafeAscii ([System.IO.Path]::GetDirectoryName($path)) } catch { $root = '' }
    }
    if (-not $name -and $path) {
        try { $name = _SafeAscii ([System.IO.Path]::GetFileName($path)) } catch { }
    }
    $tail = ("{0,-9} {1}" -f $state, $ts)
    if ($ColWidth -le ($tail.Length + 2)) { return _Trunc -Text $tail -Max $ColWidth }
    $headBudget = $ColWidth - $tail.Length - 1
    $headRaw = if ($root -and $name) { "$root $name" } elseif ($name) { $name } elseif ($root) { $root } elseif ($id) { $id } else { '-' }
    $head = _Trunc -Text $headRaw -Max $headBudget
    return _Trunc -Text ("$head $tail") -Max $ColWidth
}

function _Format-StatusSetJobLine([hashtable]$Cfg, [object]$Entry, [int]$ColWidth) {
    if (-not $Entry) { return '' }
    $ts = _Format-UiTs -IsoOrNull ([string]$Entry.lastWriteTimeUtc)
    $state = [string]$Entry.state
    $folder = ''
    $id = ''
    try {
        if ($Entry.job) {
            $j = $Entry.job
            $id = [string]$j.id
            if ($j.sourceFolder) { $folder = [string]$j.sourceFolder }
            elseif ($j.sourcePath) { $folder = _SafeAscii ([System.IO.Path]::GetDirectoryName([string]$j.sourcePath)) }
        }
    } catch { }
    $folder = _SafeAscii $folder
    $proj = _TryGet-ProjectNameFromFolder -Cfg $Cfg -FolderPath $folder
    $dir = if ($proj) { $proj } else { '' }
    if (-not $dir) {
        # Fallback: show Documents\<root>\<project> (or last two segments) if we can't match a watch root.
        $segs = _Get-DocumentsAreaTwoSegments $folder
        $dir = if ($segs) { ($segs[0] + '/' + $segs[1]) } else { '' }
        if (-not $dir -and $folder) {
            $norm = $folder.Trim() -replace '\\', '/'
            $parts = @($norm -split '/' | Where-Object { $_ -ne '' })
            if ($parts.Count -ge 2) { $dir = $parts[$parts.Count - 2] + '/' + $parts[$parts.Count - 1] }
            else { $dir = $norm }
        }
    }
    $prefix = ("{0,-9} {1} " -f $state, $ts)
    if ($ColWidth -le ($prefix.Length + 2)) { return _Trunc -Text $prefix.TrimEnd() -Max $ColWidth }
    $dirBudget = $ColWidth - $prefix.Length
    $dirDisp = _Trunc -Text $dir -Max $dirBudget
    if (-not $dirDisp -and $id) { $dirDisp = _Trunc -Text $id -Max $dirBudget }
    return _Trunc -Text ($prefix + $dirDisp) -Max $ColWidth
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
    # True after the dashboard has started its main poll loop (watcher respawns while this stays true).
    passPipelineActive = $false
    watcherPid = 0
    # True only in the brief window after a pass begins and before Watch-QCTrigger is spawned (avoids bogus "Idle").
    awaitingWatcherSpawn = $false
    hasSeenPwScan = $false
    pwConnectOkSeen = $false
    pwFoldersPreparedSeen = $false
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
    # Updated each poll tick for status text (worker child processes still alive).
    activeWorkerSlots = 0
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

function _Test-HasPwWatchList([hashtable]$Cfg) {
    if (-not $Cfg) { return $false }
    try {
        if (-not ($Cfg.ContainsKey('projectWise') -and $Cfg.projectWise)) { return $false }
        $pw = $Cfg.projectWise
        if (-not ($pw -is [hashtable])) { return $false }
        return [bool]($pw.ContainsKey('watchList') -and $pw.watchList)
    } catch {
        return $false
    }
}

function _Get-PwSessionIndicator([hashtable]$Cfg) {
    if (-not (_Test-HasPwWatchList -Cfg $Cfg)) { return $null }
    $pipe = $false
    try { $pipe = [bool]$state.passPipelineActive } catch { $pipe = $false }
    $alive = $false
    try { $alive = [bool]$state.watcherAlive } catch { $alive = $false }
    $ok = $false
    try { $ok = [bool]$state.pwConnectOkSeen } catch { $ok = $false }
    $pidTxt = ''
    try { if ([int]$state.watcherPid -gt 0) { $pidTxt = (' pid=' + [int]$state.watcherPid) } } catch { }

    if (-not $pipe -and -not $alive) {
        return @{ text = ('NOT RUNNING (between passes)'); color = 'DarkGray' }
    }
    if ($alive -and -not $ok) {
        return @{ text = ('CONNECTING' + $pidTxt); color = 'Yellow' }
    }
    if ($alive -and $ok) {
        return @{ text = ('SESSION ACTIVE' + $pidTxt); color = 'Green' }
    }
    # Pipeline still draining after watcher exit — not an active PW API session.
    return @{ text = ('INACTIVE (watcher ended' + $pidTxt + ')'); color = 'Red' }
}

function _Get-ActivityStatusText([hashtable]$Cfg) {
    $hasPw = _Test-HasPwWatchList -Cfg $Cfg
    $max = [Math]::Max(48, (_Get-TerminalWidth) - 4)

    if ($state.lastError) { return _Trunc -Text ('ERROR: ' + [string]$state.lastError) -Max $max }

    $awaitSpawn = $false
    try { $awaitSpawn = [bool]$state.awaitingWatcherSpawn } catch { $awaitSpawn = $false }
    if ($awaitSpawn) {
        if ($hasPw) { return _Trunc -Text 'Spawning watcher; next step is ProjectWise connect…' -Max $max }
        return _Trunc -Text 'Spawning file watcher…' -Max $max
    }

    $pipe = $false
    try { $pipe = [bool]$state.passPipelineActive } catch { $pipe = $false }
    $watcherAlive = $false
    try { $watcherAlive = [bool]$state.watcherAlive } catch { $watcherAlive = $false }

    $stage = ''
    try { $stage = ([string]$state.currentScanStage).Trim() } catch { $stage = '' }
    if ($stage -eq 'watch pass completed') { $stage = '' }

    $pend = 0; $run = 0
    try {
        if ($state.queueStats -and $state.queueStats.states) {
            $pend = [int]$state.queueStats.states.pending
            $run = [int]$state.queueStats.states.running
        }
    } catch { }

    $slots = 0
    try { $slots = [int]$state.activeWorkerSlots } catch { $slots = 0 }

    if ($pipe) {
        if ($hasPw -and -not $watcherAlive) {
            $busy = ($slots -gt 0) -or ($pend -gt 0) -or ($run -gt 0)
            if ($busy) {
                return _Trunc -Text ("Watcher pass finished; workers/queue active (workers={0} pend={1} run={2}); next watch pass starts automatically." -f $slots, $pend, $run) -Max $max
            }
            return _Trunc -Text ("Watcher pass finished; starting next watch pass (queue pend={0} run={1})…" -f $pend, $run) -Max $max
        }
        if ($stage -and $stage -ne 'watcher starting') {
            return _Trunc -Text $stage -Max $max
        }
        if ($hasPw) {
            if ($watcherAlive -and -not [bool]$state.pwConnectOkSeen) {
                return _Trunc -Text 'ProjectWise: opening session (see _logs if this stalls)…' -Max $max
            }
            if ($watcherAlive -and [bool]$state.pwConnectOkSeen -and -not [bool]$state.pwFoldersPreparedSeen) {
                return _Trunc -Text 'ProjectWise: building folder list from watchList…' -Max $max
            }
            if ($watcherAlive) {
                return _Trunc -Text 'ProjectWise: scanning watch folders…' -Max $max
            }
        }
        if ($watcherAlive) {
            return _Trunc -Text ($(if ($stage) { $stage } else { 'Watcher running…' })) -Max $max
        }
        if ($slots -gt 0 -or $pend -gt 0 -or $run -gt 0) {
            return _Trunc -Text ("Watcher pass finished; workers/queue active (workers={0} pend={1} run={2}); next watch pass starts automatically." -f $slots, $pend, $run) -Max $max
        }
        return _Trunc -Text ("Finishing pipeline (queue pend={0} run={1})…" -f $pend, $run) -Max $max
    }

    return _Trunc -Text 'Idle - ready for next pass' -Max $max
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
        # Fetch a larger window so each column can show the last N jobs of that type (queue API returns a single mixed list).
        $recentLimit = [Math]::Max(80, [int]$RecentJobs * 20)
        $recent = Get-QCRecentJobs -Config $Cfg -Limit $recentLimit
        if ($recent.IsSuccess) { $state.recentJobs = @($recent.Data.jobs) }
        $state.lastQueueRefreshUtc = Get-QCTimestamp
    } catch { }
}

function _Compute-PhaseText([hashtable]$Cfg) {
    return (_Get-ActivityStatusText -Cfg $Cfg)
}

function _Format-TwoColumns([string]$Left, [string]$Right, [int]$ColWidth) {
    $l = _Trunc -Text $Left -Max $ColWidth
    $r = _Trunc -Text $Right -Max $ColWidth
    $pad = ' ' * ([Math]::Max(1, ($ColWidth - $l.Length) + 3))
    $leftText = if ($l) { $l } else { ' ' * $ColWidth }
    return ($leftText + $pad + $r).TrimEnd()
}

function _Get-WorkersLines() {
    $lines = New-Object System.Collections.Generic.List[string]
    $max = [int]$state.workerSlotMax
    if ($max -le 0) { $max = $state.workers.Count }
    $lines.Add(("Workers (max {0})" -f $max)) | Out-Null
    if (-not $state.workers -or $state.workers.Count -eq 0) {
        $lines.Add('  (no workers spawned yet)') | Out-Null
        return @($lines)
    }

    $sorted = @($state.workers.Values | Sort-Object -Property label)
    foreach ($w in $sorted) {
        $st = [string]$w.state
        $elapsed = '-'
        if ($w.startedAtUtc) {
            try {
                $s = _Parse-UtcIso -Value ([string]$w.startedAtUtc)
                if ($s) { $elapsed = ("{0}s" -f [int]((Get-Date).ToUniversalTime() - $s).TotalSeconds) }
            } catch { }
        }
        $jobId = if ($w.jobId) { [string]$w.jobId } else { '-' }
        $jobType = if ($w.jobType) { [string]$w.jobType } else { '-' }
        $proj = if ($w.projectName) { [string]$w.projectName } else { '-' }
        $lines.Add((
            "  {0,-4}  pid={1,-6}  {2,-8}  {3,-15}  {4,-30}  {5,-24}  {6}" -f `
                [string]$w.label, [int]$w.pid, $st, $jobType, (_Trunc -Text $proj -Max 30), (_Trunc -Text $jobId -Max 24), $elapsed
        )) | Out-Null
    }
    return @($lines)
}

function _Get-RecentJobsTwoColLines([object[]]$Jobs, [int]$Limit) {
    $lines = New-Object System.Collections.Generic.List[string]
    $width = _Get-TerminalWidth
    $colWidth = [int][Math]::Max(44, [Math]::Floor(($width - 6) / 2))

    $jobs = @($Jobs)
    # Get-QCRecentJobs returns entries shaped like: @{ state; job; lastWriteTimeUtc } (job is nested).
    $qcJobs = @($jobs | Where-Object { $_ -and $_.job -and ([string]$_.job.type) -eq 'QC_PREPEND' } | Select-Object -First $Limit)
    $stJobs = @($jobs | Where-Object { $_ -and $_.job -and ([string]$_.job.type) -eq 'STATUS_SET_GEN' } | Select-Object -First $Limit)

    $lines.Add((_Format-TwoColumns -Left 'QC (QC_PREPEND)' -Right 'Status (STATUS_SET_GEN)' -ColWidth $colWidth)) | Out-Null
    for ($i = 0; $i -lt [Math]::Max($qcJobs.Count, $stJobs.Count); $i++) {
        $lJob = if ($i -lt $qcJobs.Count) { $qcJobs[$i] } else { $null }
        $rJob = if ($i -lt $stJobs.Count) { $stJobs[$i] } else { $null }
        $lText = _Format-QCPrependJobLine -Entry $lJob -ColWidth $colWidth
        $rText = _Format-StatusSetJobLine -Cfg $script:_DashCfg -Entry $rJob -ColWidth $colWidth
        $lines.Add((_Format-TwoColumns -Left $lText -Right $rText -ColWidth $colWidth)) | Out-Null
    }
    return @($lines)
}

function _Get-FrameLines([hashtable]$Cfg) {
    $lines = New-Object System.Collections.Generic.List[string]
    $width = _Get-TerminalWidth
    $wideMax = [Math]::Max(48, $width - 2)
    $now = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, [TimeZoneInfo]::FindSystemTimeZoneById('Mountain Standard Time'))
    $dry = $false
    try { $dry = [bool]$Cfg.dryRun } catch { $dry = $false }

    $lines.Add(("QC Pipeline Dashboard   {0} MST   DryRun={1}" -f $now.ToString('yyyy-MM-dd HH:mm:ss'), $dry)) | Out-Null
    $lines.Add(("Config: {0}" -f $AppSettingsPath)) | Out-Null

    $status = 'Status: ' + (_Compute-PhaseText -Cfg $Cfg)
    if ($state.queueStats) {
        $st = $state.queueStats.states
        $p = [int]$st.pending; $ru = [int]$st.running; $su = [int]$st.succeeded; $fa = [int]$st.failed
        $lk = [int]$state.queueStats.locks.count
        $status += ("     Pend {0}  Run {1}  OK {2}  Fail {3}  Locks {4}" -f $p, $ru, $su, $fa, $lk)
    }
    $lines.Add((_Trunc -Text $status -Max $wideMax)) | Out-Null

    $pwInd = _Get-PwSessionIndicator -Cfg $Cfg
    if ($pwInd) { $lines.Add(('ProjectWise: ' + [string]$pwInd.text)) | Out-Null }
    $lines.Add(("Root:   {0}" -f $(if ($state.scanRoot) { $state.scanRoot } else { '-' }))) | Out-Null
    $lines.Add(("Proj:   {0}" -f $(if ($state.scanProject) { $state.scanProject } else { '-' }))) | Out-Null
    $lines.Add(("Path:   {0}" -f (_Trunc -Text ($(if ($state.scanPath) { $state.scanPath } else { '-' })) -Max ($wideMax - 8)))) | Out-Null

    $lastPass = ''
    if ($state.lastPassDurationMs) { $lastPass = ("Last pass: {0:0.0}s" -f ([double]$state.lastPassDurationMs / 1000.0)) }
    if ($state.passStartedAtUtc) {
        $elapsed = 0
        try {
            $start = _Parse-UtcIso -Value ([string]$state.passStartedAtUtc)
            if ($start) { $elapsed = [int]((Get-Date).ToUniversalTime() - $start).TotalSeconds } else { $elapsed = 0 }
        } catch { $elapsed = 0 }
        $suffix = if ($lastPass) { "  $lastPass" } else { '' }
        $lines.Add(("Pass:   #{0}  Elapsed: {1}s{2}" -f [int]$state.passCount, $elapsed, $suffix)) | Out-Null
    } elseif ($lastPass) {
        $lines.Add(("Pass:   #{0}  {1}" -f [int]$state.passCount, $lastPass)) | Out-Null
    }
    if ([int]$state.pwFolderTotal -gt 0) {
        $lines.Add(("PW:     {0}/{1} folders" -f [int]$state.pwFolderIndex, [int]$state.pwFolderTotal)) | Out-Null
    }
    if ($state.currentScanStage) {
        $lines.Add(("Stage:  {0}" -f (_Trunc -Text ([string]$state.currentScanStage) -Max ($wideMax - 8)))) | Out-Null
    }
    if ($state.recentScanFolders -and $state.recentScanFolders.Count -gt 0) {
        $tail = @($state.recentScanFolders | Select-Object -Last 3)
        $lines.Add(("Recent: {0}" -f (_Trunc -Text ($tail -join '  |  ') -Max ($wideMax - 8)))) | Out-Null
    }
    if ($state.lastHeartbeatUtc) { $lines.Add(("Heartbeat: {0}" -f (_Format-UiTs -IsoOrNull ([string]$state.lastHeartbeatUtc)))) | Out-Null }
    $lines.Add('') | Out-Null

    foreach ($line in @(_Get-WorkersLines)) { $lines.Add($line) | Out-Null }
    $lines.Add('') | Out-Null

    $lines.Add(("Recent jobs (last {0} per column)" -f $RecentJobs)) | Out-Null
    foreach ($line in @(_Get-RecentJobsTwoColLines -Jobs @($state.recentJobs) -Limit $RecentJobs)) { $lines.Add($line) | Out-Null }
    $lines.Add('') | Out-Null

    $lines.Add('Processor activity') | Out-Null
    if ($state.lastWorkerEvent) {
        $evt = $state.lastWorkerEvent
        $lines.Add(("  {0} {1} - {2}" -f [string]$evt.code, [string]$evt.level, [string]$evt.message)) | Out-Null
    } else {
        $lines.Add('  (no worker activity yet)') | Out-Null
    }
    $lines.Add('') | Out-Null

    $lines.Add(("Recent warnings/errors (last {0})" -f $RecentErrors)) | Out-Null
    $tail = @($state.errors | Select-Object -Last $RecentErrors)
    if ($tail.Count -eq 0) {
        $lines.Add('  (none)') | Out-Null
    } else {
        foreach ($e in $tail) {
            $folder = ''
            $errMsg = ''
            try { if ($e.data -and $e.data.folder) { $folder = [string]$e.data.folder } } catch { }
            try { if ($e.data -and $e.data.errorMessage) { $errMsg = [string]$e.data.errorMessage } } catch { }
            $suffix = ''
            if (-not [string]::IsNullOrWhiteSpace($folder)) { $suffix += ('  [' + (_Trunc -Text $folder -Max 80) + ']') }
            if (-not [string]::IsNullOrWhiteSpace($errMsg)) { $suffix += ('  ' + (_Trunc -Text $errMsg -Max 110)) }
            $lines.Add(("  {0}  {1,-8}  {2}  {3}{4}" -f [string]$e.ts, [string]$e.level, [string]$e.code, [string]$e.message, $suffix)) | Out-Null
        }
    }
    if ($state.lastError) {
        $lines.Add('') | Out-Null
        $lines.Add('Last fatal error (will retry):') | Out-Null
        $lines.Add(("  {0}" -f [string]$state.lastError)) | Out-Null
    }

    return @($lines | ForEach-Object { _Trunc -Text ([string]$_) -Max $wideMax })
}

function _Test-VtSupported() {
    try {
        if ($env:TERM_PROGRAM -or $env:WT_SESSION -or $env:ConEmuANSI -eq 'ON' -or $env:ANSICON) { return $true }
        if ($env:TERM -and $env:TERM -notin @('', 'dumb')) { return $true }
        if ($Host.UI -and ($Host.UI.PSObject.Properties.Name -contains 'SupportsVirtualTerminal') -and [bool]$Host.UI.SupportsVirtualTerminal) { return $true }
    } catch { }
    return $false
}

function _Write-Ansi([string]$Text) {
    [Console]::Write($Text)
}

$script:_DashEsc = [char]27
$script:_DashPrevFrame = $null
$script:_DashLastRenderUtc = [DateTime]::MinValue
$script:_DashEffectiveRenderMode = if ($RenderMode -eq 'DiffAnsi' -and -not (_Test-VtSupported)) { 'FullRedraw' } else { $RenderMode }
$script:_DashCursorHidden = $false
try {
    Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
        try { [Console]::Write("$([char]27)[?25h") } catch { }
    } | Out-Null
} catch { }

function _Hide-Cursor() {
    if ($script:_DashEffectiveRenderMode -ne 'DiffAnsi' -or $script:_DashCursorHidden) { return }
    try {
        _Write-Ansi ($script:_DashEsc + '[?25l')
        $script:_DashCursorHidden = $true
    } catch { }
}

function _Show-Cursor() {
    if (-not $script:_DashCursorHidden) { return }
    try { _Write-Ansi ($script:_DashEsc + '[?25h') } catch { }
    $script:_DashCursorHidden = $false
}

function _Render-Full([hashtable]$Cfg) {
    Clear-Host
    foreach ($line in @(_Get-FrameLines -Cfg $Cfg)) {
        Write-Host $line
    }
    $script:_DashPrevFrame = $null
}

function _Render-DiffAnsi([hashtable]$Cfg) {
    _Hide-Cursor
    $esc = $script:_DashEsc
    $frame = @(_Get-FrameLines -Cfg $Cfg)
    if ($null -eq $script:_DashPrevFrame) {
        _Write-Ansi ($esc + '[2J' + $esc + '[H')
        if ($frame.Count -gt 0) { _Write-Ansi (($frame -join "`r`n") + "`r`n") }
        $script:_DashPrevFrame = @($frame)
        return
    }

    $prev = @($script:_DashPrevFrame)
    $maxRows = [Math]::Max($prev.Count, $frame.Count)
    for ($i = 0; $i -lt $maxRows; $i++) {
        $old = if ($i -lt $prev.Count) { [string]$prev[$i] } else { $null }
        $new = if ($i -lt $frame.Count) { [string]$frame[$i] } else { '' }
        if ($old -ne $new) {
            $row = $i + 1
            _Write-Ansi ("{0}[{1};1H{0}[2K{2}" -f $esc, $row, $new)
        }
    }
    _Write-Ansi ("{0}[{1};1H" -f $esc, ([Math]::Max(1, $frame.Count + 1)))
    $script:_DashPrevFrame = @($frame)
}

function _Render-Dashboard([hashtable]$Cfg, [switch]$Force) {
    $nowUtc = [DateTime]::UtcNow
    $intervalMs = [int](1000 / [Math]::Max(1, [int]$MaxFps))
    if (-not $Force.IsPresent -and $script:_DashLastRenderUtc -ne [DateTime]::MinValue) {
        if (($nowUtc - $script:_DashLastRenderUtc).TotalMilliseconds -lt $intervalMs) { return }
    }
    $script:_DashLastRenderUtc = $nowUtc

    if ($script:_DashEffectiveRenderMode -eq 'DiffAnsi') {
        _Render-DiffAnsi -Cfg $Cfg
    } else {
        _Render-Full -Cfg $Cfg
    }
}

function _Get-WorkersConfig([hashtable]$Cfg) {
    $w = @{
        maxParallel      = 2
        maxJobsPerWorker = 25
        leaseSeconds     = 600
        idleSleepMs      = 750
        watcherIdleSleepMs = 750
        spawnStaggerMs   = 250
    }
    try {
        if ($Cfg.ContainsKey('workers') -and $Cfg.workers) {
            $src = $Cfg.workers
            if ($src -is [hashtable]) {
                foreach ($k in @('maxParallel','maxJobsPerWorker','leaseSeconds','idleSleepMs','spawnStaggerMs')) {
                    if ($src.ContainsKey($k) -and $src[$k] -ne $null) { $w[$k] = [int]$src[$k] }
                }
                $w.watcherIdleSleepMs = [int]$w.idleSleepMs
            }
        }
        if ($Cfg.ContainsKey('watcher') -and $Cfg.watcher -and ($Cfg.watcher -is [hashtable])) {
            if ($Cfg.watcher.ContainsKey('idleSleepMs') -and $Cfg.watcher.idleSleepMs -ne $null) { $w.watcherIdleSleepMs = [int]$Cfg.watcher.idleSleepMs }
        }
    } catch { }
    if ($Workers -gt 0) { $w.maxParallel = [int]$Workers }
    if ($w.maxParallel -lt 1) { $w.maxParallel = 1 }
    if ($w.watcherIdleSleepMs -lt 100) { $w.watcherIdleSleepMs = 100 }
    return $w
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

function _CmdLineEscapeDoubleQuotes([string]$Value) {
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Replace('"', '""')
}

function _Append-CmdLineArg([System.Text.StringBuilder]$Sb, [string]$Value) {
    if ($null -eq $Value) { return }
    $s = [string]$Value
    if ($s -match '[\s"]') {
        [void]$Sb.Append('"')
        [void]$Sb.Append((_CmdLineEscapeDoubleQuotes $s))
        [void]$Sb.Append('"')
    } else {
        [void]$Sb.Append($s)
    }
}

function _Start-Child([string]$ScriptPath, [string[]]$ScriptArgs, [switch]$Mta) {
    # Persist child stdout/stderr under queue\_logs\ so we can read what a killed
    # process did/said even after the dashboard reaps it. Files are named with a
    # timestamp + script tag + PID-placeholder; PID is rewritten after spawn.
    $logDir = _Get-ChildLogDir
    $tag = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
    $stamp = Get-QCTimestampShort
    $stdoutPath = Join-Path $logDir ("${stamp}_${tag}.out.log")
    $stderrPath = Join-Path $logDir ("${stamp}_${tag}.err.log")
    # Windows: Start-Process -ArgumentList @(...) joins arguments in a way that breaks paths
    # containing spaces (e.g. OneDrive - TYPSA). Pass one ArgumentList string with cmd-style quoting.
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('-NoProfile -ExecutionPolicy Bypass ')
    if ($Mta.IsPresent) { [void]$sb.Append('-MTA ') }
    [void]$sb.Append('-File ')
    _Append-CmdLineArg -Sb $sb -Value $ScriptPath
    foreach ($a in @($ScriptArgs)) {
        [void]$sb.Append(' ')
        _Append-CmdLineArg -Sb $sb -Value ([string]$a)
    }
    $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $sb.ToString() -NoNewWindow -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    return @{
        process = $p
        stdoutPath = $stdoutPath
        stderrPath = $stderrPath
        lastStdoutLen = 0
        # Incomplete tail of stdout (JSON split across reads) — kept until a full line arrives.
        stdoutTail = ''
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
    if ($null -eq $Child.stdoutTail) { $Child.stdoutTail = '' }
    if ($cur.Length -lt $Child.lastStdoutLen) {
        $Child.lastStdoutLen = 0
        $Child.stdoutTail = ''
    }
    if ($cur.Length -gt $Child.lastStdoutLen) {
        $delta = $cur.Substring($Child.lastStdoutLen)
        $Child.lastStdoutLen = $cur.Length
        $buffer = ([string]$Child.stdoutTail) + $delta

        $lines = New-Object System.Collections.Generic.List[string]
        $segStart = 0
        for ($i = 0; $i -lt $buffer.Length; $i++) {
            $ch = $buffer[$i]
            if ($ch -eq "`n") {
                $ln = $buffer.Substring($segStart, $i - $segStart)
                if ($ln.Length -gt 0 -and $ln[$ln.Length - 1] -eq "`r") { $ln = $ln.Substring(0, $ln.Length - 1) }
                $lines.Add($ln)
                $segStart = $i + 1
            }
        }
        if ($segStart -lt $buffer.Length) {
            $Child.stdoutTail = $buffer.Substring($segStart)
        } else {
            $Child.stdoutTail = ''
        }

        foreach ($line in $lines) {
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
                                    sourceFolder = ''
                                    projectName = ''
                                    state = 'IDLE'
                                    lastCode = [string]$o.code
                                    lastMessage = [string]$o.message
                                    stage = ''
                                    startedAtUtc = $null
                                    updatedAtUtc = Get-QCTimestamp
                                }
                            }
                            $w = $state.workers[$wlbl]
                            if ($wpid -gt 0) { $w.pid = $wpid }
                            $w.lastCode = [string]$o.code
                            $w.lastMessage = [string]$o.message
                            try { if ($o.data -and $o.data.stage) { $w.stage = [string]$o.data.stage } } catch { }
                            $w.updatedAtUtc = Get-QCTimestamp
                            switch ([string]$o.code) {
                                'WORKER_START'   { $w.state = 'IDLE' }
                                'WORKER_SELECTED' {
                                    $w.state = 'RUNNING'
                                    if ($o.data) {
                                        if ($o.data.jobId)        { $w.jobId        = [string]$o.data.jobId }
                                        if ($o.data.jobType)      { $w.jobType      = [string]$o.data.jobType }
                                        if ($o.data.sourceFolder) { $w.sourceFolder = [string]$o.data.sourceFolder }
                                        $w.projectName = ''
                                        try {
                                            $pn = _TryGet-ProjectNameFromFolder -Cfg $script:_DashCfg -FolderPath ([string]$w.sourceFolder)
                                            if ($pn) { $w.projectName = [string]$pn }
                                        } catch { }
                                    }
                                    $w.stage = 'selected job; preparing processor'
                                    $w.startedAtUtc = Get-QCTimestamp
                                }
                                'WORKER_STAGE' {
                                    if ($o.data) {
                                        if ($o.data.jobId)        { $w.jobId        = [string]$o.data.jobId }
                                        if ($o.data.jobType)      { $w.jobType      = [string]$o.data.jobType }
                                        if ($o.data.sourceFolder) {
                                            $w.sourceFolder = [string]$o.data.sourceFolder
                                            try {
                                                $pn = _TryGet-ProjectNameFromFolder -Cfg $script:_DashCfg -FolderPath ([string]$w.sourceFolder)
                                                $w.projectName = if ($pn) { [string]$pn } else { '' }
                                            } catch { }
                                        }
                                    }
                                    $w.state = if ($w.jobId) { 'RUNNING' } else { 'IDLE' }
                                    if (-not $w.startedAtUtc -and $w.jobId) { $w.startedAtUtc = Get-QCTimestamp }
                                }
                                'WORKER_SUCCEEDED' { $w.state = 'IDLE'; $w.jobId = ''; $w.jobType = ''; $w.sourceFolder = ''; $w.projectName = ''; $w.stage = 'completed job; polling queue'; $w.startedAtUtc = $null }
                                'WORKER_FAILED'    { $w.state = 'IDLE'; $w.jobId = ''; $w.jobType = ''; $w.sourceFolder = ''; $w.projectName = ''; $w.stage = 'job failed; polling queue'; $w.startedAtUtc = $null }
                                'WORKER_NO_JOB'    { $w.state = 'IDLE'; $w.jobId = ''; $w.jobType = ''; $w.sourceFolder = ''; $w.projectName = ''; $w.stage = 'idle; waiting for pending jobs' }
                                'WORKER_LOCK_RACE' { $w.state = 'IDLE'; $w.stage = 'lock race; trying next job' }
                                'WORKER_BUDGET'    { $w.state = 'EXITING'; $w.stage = 'max-jobs budget reached' }
                                'WORKER_LEASE'     { $w.state = 'EXITING'; $w.stage = 'lease budget reached' }
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
                        if ([bool]$state.pwConnectOkSeen) {
                            $state.phase = 'Scanning folders...'
                            $state.currentScanStage = 'establishing PW session'
                        } else {
                            $state.phase = 'Connecting to ProjectWise...'
                            $state.currentScanStage = 'connecting to ProjectWise'
                        }
                    }
                    if ($o.code -eq 'WATCH_PW_CONNECT_OK') {
                        $state.currentScanStage = 'connected to ProjectWise'
                        $state.pwConnectOkSeen = $true
                    }
                    if ($o.code -eq 'WATCH_RECONCILE_START') {
                        $state.currentScanStage = 'reconciling local _StatusSet.pdf copies to ProjectWise'
                    }
                    if ($o.code -eq 'WATCH_RECONCILE_UPDATED') {
                        try {
                            $pf = ''
                            if ($o.data.pwFolder) { $pf = [string]$o.data.pwFolder }
                            elseif ($o.data.sheetsFolder) { $pf = [string]$o.data.sheetsFolder }
                            if ($pf) {
                                $state.currentScanStage = "reconciling status set: $pf"
                                _State-SetScanContext -FolderPath $pf
                            }
                        } catch { }
                    }
                    if ($o.code -eq 'WATCH_RECONCILE_DONE') {
                        $updated = 0
                        try { if ($o.data.counts) { $updated = [int]$o.data.counts.updated } } catch { }
                        $state.currentScanStage = if ($updated -gt 0) {
                            "status set reconcile done ($updated updated)"
                        } else {
                            'status set reconcile done'
                        }
                    }
                    if ($o.code -eq 'WATCH_RECONCILE_FAILED') {
                        $state.currentScanStage = 'status set reconcile failed (see watcher log)'
                    }
                    if ($o.code -eq 'WATCH_AUDIT_SCAN_START') {
                        $since = ''; $until = ''
                        try {
                            if ($o.data.since) { $since = [string]$o.data.since }
                            if ($o.data.until) { $until = [string]$o.data.until }
                        } catch { }
                        if ($since -and $until) {
                            $state.currentScanStage = "audit trail scan ($since to $until)"
                        } else {
                            $state.currentScanStage = 'audit trail scan starting'
                        }
                    }
                    if ($o.code -eq 'WATCH_AUDIT_SCAN_DONE') {
                        $rel = 0; $tot = 0; $cand = 0
                        try {
                            $rel = [int]$o.data.relevantEvents
                            $tot = [int]$o.data.totalEvents
                            $cand = [int]$o.data.candidates
                        } catch { }
                        $state.currentScanStage = "audit scan done ($rel relevant / $tot events, $cand candidates)"
                    }
                    if ($o.code -eq 'WATCH_AUDIT_SCAN_FAILED' -or $o.code -eq 'WATCH_AUDIT_SCAN_ERROR') {
                        $state.currentScanStage = 'audit trail scan failed (see watcher log)'
                    }
                    if ($o.code -eq 'WATCH_AUDIT_FALLBACK') {
                        $state.currentScanStage = 'audit scan failed; falling back to full folder scan'
                        $state.pwFoldersPreparedSeen = $false
                    }
                    if ($o.code -eq 'WATCH_RECONCILE_CYCLE') {
                        $cn = 0; $every = 20
                        try {
                            $cn = [int]$o.data.cycleNum
                            $every = [int]$o.data.reconcileEvery
                        } catch { }
                        $state.currentScanStage = "scheduled full folder scan (cycle $cn, every $every)"
                        $state.pwFoldersPreparedSeen = $false
                    }
                    if ($o.code -eq 'WATCH_PW_ERROR') {
                        $em = ''
                        try { if ($o.data -and $o.data.errorMessage) { $em = [string]$o.data.errorMessage } } catch { }
                        $state.currentScanStage = if ($em) { "ProjectWise watch error: $em" } else { 'ProjectWise watch error (see watcher log)' }
                    }
                    if ($o.code -eq 'WATCH_PW_FOLDERS') {
                        $state.pwFoldersPreparedSeen = $true
                    }
                    if ($o.code -eq 'WATCH_PW_SCAN_START' -or $o.code -eq 'WATCH_PW_FOLDER_ERROR') { $state.phase = 'Scanning folders...' }
                    if ($o.code -eq 'WATCH_PW_ONELEVEL_EXPAND_PROGRESS') {
                        $state.phase = 'Scanning folders...'
                        try {
                            $folder = if ($o.data.folder) { [string]$o.data.folder } else { '' }
                            $inProg = $false
                            try { $inProg = [bool]$o.data.inProgress } catch { $inProg = $false }
                            if ($inProg) {
                                $state.currentScanStage = if ($folder) { "listing discipline subfolders: $folder..." } else { 'listing discipline subfolders (Sheets)...' }
                            } elseif ($folder) {
                                $cn = 0
                                try { $cn = [int]$o.data.childCount } catch { $cn = 0 }
                                $state.currentScanStage = "listing discipline subfolders: $folder ($cn found)"
                            } else {
                                $state.currentScanStage = 'listing discipline subfolders (Sheets)'
                            }
                        } catch {
                            $state.currentScanStage = 'listing discipline subfolders (Sheets)'
                        }
                    }
                    # (phase is derived in renderer; keep these events for stage/context only)

                    # Pass tracking + progress
                    if ($o.code -eq 'WATCH_START') {
                        # Use the event timestamp (already UTC) for consistent timing.
                        try { $state.passStartedAtUtc = [string]$o.ts } catch { $state.passStartedAtUtc = Get-QCTimestamp }
                        # Every WATCH_START is a new pass.
                        if ([int]$state.passCount -le 0) { $state.passCount = 1 } else { $state.passCount = ([int]$state.passCount + 1) }
                        $state.pwFolderTotal = 0
                        $state.pwFolderIndex = 0
                        $state.currentScanStage = 'watcher starting'
                        # Fresh PW session indicators each pass so status line progresses Connect → Folders → Scan again.
                        $state.pwConnectOkSeen = $false
                        $state.pwFoldersPreparedSeen = $false
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
        $state.lastHeartbeatUtc = Get-QCTimestamp
        _MaybeRefreshQueue -Cfg $Cfg -MinIntervalMs 1500
        _Render-Dashboard -Cfg $Cfg
    }
}

# Singleton guard + boot config: any failure here used to kill the process before the
# main loop, which makes a double-clicked console window vanish in ~1 second.
try {
    $bootCfg = Get-QCAppSettingsConfig -Path $AppSettingsPath -DryRun:$DryRun.IsPresent
    $queueRoot = $null
    try {
        if ($bootCfg.queue -and $bootCfg.queue.rootDir) { $queueRoot = [string]$bootCfg.queue.rootDir }
        elseif ($bootCfg.queue -and $bootCfg.queue.root) { $queueRoot = [string]$bootCfg.queue.root }
    } catch { }
    if (-not $queueRoot) { $queueRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'queue' }
    if (-not (Test-Path -LiteralPath $queueRoot)) { New-Item -ItemType Directory -Path $queueRoot -Force | Out-Null }

    # Singleton guard: refuse to start if another dashboard is already running. Multiple
    # concurrent dashboards multiply the watcher + worker process count and overwhelm
    # Fortinet (which then kills processes mid-job, leaving orphan running\ jobs and
    # empty output dirs). Use scripts\Stop-QCPipeline.ps1 to clean up stale instances.
    $dashLock = Join-Path $queueRoot '_dashboard.lock'
    $alreadyRunning = $false
    if (-not $IgnoreSingletonLock.IsPresent) {
        if (Test-Path -LiteralPath $dashLock) {
            try {
                $pl = Get-Content -LiteralPath $dashLock -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($pl -and $pl.pid) {
                    # IMPORTANT: PID reuse is common on Windows. Validate that the PID is
                    # actually a PowerShell process running THIS dashboard script, not
                    # an unrelated process that happened to reuse the same PID.
                    $existing = Get-Process -Id ([int]$pl.pid) -ErrorAction SilentlyContinue
                    if (-not $existing) {
                        # Stale lock (process gone); remove so we can start.
                        try { Remove-Item -LiteralPath $dashLock -Force -ErrorAction SilentlyContinue } catch { }
                    } else {
                        $isPw = ($existing.ProcessName -in @('powershell', 'pwsh'))
                        $cmdOk = $false
                        if ($isPw) {
                            try {
                                $p2 = Get-CimInstance Win32_Process -Filter ("ProcessId={0}" -f ([int]$pl.pid)) -ErrorAction SilentlyContinue
                                if ($p2 -and $p2.CommandLine) {
                                    $cmd = [string]$p2.CommandLine
                                    if ($cmd -match '(?i)Start-QCPipelineDashboard\.ps1') { $cmdOk = $true }
                                }
                            } catch { }
                        }
                        if ($isPw -and $cmdOk) {
                            $alreadyRunning = $true
                        } elseif ($isPw) {
                            # Live PowerShell but we could not confirm the command line (WMI/CIM flakey): assume it is the
                            # dashboard instance rather than deleting the lock and risking two writers to the same queue.
                            $alreadyRunning = $true
                        } else {
                            # PID reuse: lock points at a non-PowerShell process; treat lock as stale.
                            try { Remove-Item -LiteralPath $dashLock -Force -ErrorAction SilentlyContinue } catch { }
                        }
                    }
                } else {
                    try { Remove-Item -LiteralPath $dashLock -Force -ErrorAction SilentlyContinue } catch { }
                }
            } catch {
                try { Remove-Item -LiteralPath $dashLock -Force -ErrorAction SilentlyContinue } catch { }
            }
        }
    } else {
        try { if (Test-Path -LiteralPath $dashLock) { Remove-Item -LiteralPath $dashLock -Force -ErrorAction SilentlyContinue } } catch { }
    }
    if ($alreadyRunning) {
        Write-Host "Another dashboard is already running (see $dashLock)." -ForegroundColor Red
        Write-Host "Stop it first with: .\scripts\Stop-QCPipeline.ps1" -ForegroundColor Yellow
        Write-Host "Or override with: -IgnoreSingletonLock (not recommended if a real instance is running)" -ForegroundColor Yellow
        _Pause-IfInteractiveConsole
        exit 2
    }
    @{
        pid          = $PID
        startedAtUtc = Get-QCTimestamp
        host         = $env:COMPUTERNAME
        scriptPath   = $MyInvocation.MyCommand.Path
        queueRoot    = $queueRoot
    } |
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
                ts = Get-QCTimestamp
                level = 'Information'
                code = 'WORKER_RECOVERY'
                message = ("Stale recovery: requeued={0} failed={1}" -f [int]$rec.Data.recoveredToPending, [int]$rec.Data.recoveredToFailed)
                data = $rec.Data
            }
        }
        Clear-QCWatcherActive -Config $bootCfg | Out-Null
    } catch { }
} catch {
    Write-Host ''
    Write-Host 'Dashboard startup failed (before main loop):' -ForegroundColor Red
    Write-Host ("  {0}" -f $_.Exception.Message) -ForegroundColor Red
    try { if ($_.ScriptStackTrace) { Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray } } catch { }
    _Pause-IfInteractiveConsole
    exit 1
}

$workerSlots = @{}
$nextWorkerIndex = 1
$lastSpawnAt = [DateTime]::MinValue
$lastRecoveryAt = [DateTime]::UtcNow
$watcherChild = $null
# Reconcile-on-first-watcher only (same as prior single outer-pass behavior).
$watcherReconcileNext = -not $SkipReconcileStatusSetsFirst.IsPresent

function _Spawn-Worker([hashtable]$Cfg, [hashtable]$WC, [string]$Label) {
    $xArgs = @(
        '-AppSettingsPath', $AppSettingsPath,
        '-MaxJobs', [string]([int]$WC.maxJobsPerWorker),
        '-LeaseSeconds', [string]([int]$WC.leaseSeconds),
        '-IdleSleepMs', [string]([int]$WC.idleSleepMs),
        '-WorkerLabel', $Label
    )
    if ([bool]$Cfg.dryRun) { $xArgs += '-DryRun' }
    $newChild = _Start-Child -ScriptPath $worker -ScriptArgs $xArgs -Mta
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
            stage = 'spawning process'
            startedAtUtc = $null
            updatedAtUtc = Get-QCTimestamp
        }
    } else {
        $state.workers[$Label].pid = [int]$newChild.process.Id
        $state.workers[$Label].state = 'STARTING'
    }
    return $newChild
}

while ($true) {
    try {
        $cfg = Get-QCAppSettingsConfig -Path $AppSettingsPath -DryRun:$DryRun.IsPresent
        # Make config available to render helpers without threading it through every call.
        $script:_DashCfg = $cfg
        $wc = _Get-WorkersConfig -Cfg $cfg
        $state.workerSlotMax = [int]$wc.maxParallel

        $hasPw = _Test-HasPwWatchList -Cfg $cfg

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

        # Phase is derived in _Compute-PhaseText to avoid flicker from competing writers.
        $state.lastError = $null

        # --- Watcher: when the previous run exits, wait (inter-pass throttle) and spawn another.
        # Workers keep running; we no longer wait for an empty queue before the next watch pass.
        if ($watcherChild) {
            $watcherProcExited = $false
            try {
                $null = $watcherChild.process.Refresh()
                $watcherProcExited = [bool]$watcherChild.process.HasExited
            } catch {
                $watcherProcExited = $true
            }
            if ($watcherProcExited) {
                try { _Poll-Child -Child $watcherChild -Cfg $cfg -Kind 'watcher' } catch { }
                $exit = 0
                try {
                    if ($watcherChild.process) { $exit = [int]$watcherChild.process.ExitCode }
                } catch { $exit = 1 }
                _Stop-Child -Child $watcherChild
                if ($exit -ne 0) {
                    $watcherChild = $null
                    throw "Watcher failed with exit code $exit"
                }
                $watcherChild = $null
                $state.watcherAlive = $false
                $state.watcherPid = 0
                if ($PollSeconds -gt 0) { Start-Sleep -Seconds $PollSeconds }
                else { Start-Sleep -Milliseconds ([int]$wc.watcherIdleSleepMs) }
            }
        }

        if (-not $watcherChild) {
            $state.awaitingWatcherSpawn = $true
            $state.watcherAlive = $false
            $state.watcherPid = 0
            $state.currentScanStage = ''
            $state.pwConnectOkSeen = $false
            $state.pwFoldersPreparedSeen = $false
            _MaybeRefreshQueue -Cfg $cfg -MinIntervalMs 0
            _Render-Dashboard -Cfg $cfg -Force

            $wArgs = @('-AppSettingsPath', $AppSettingsPath)
            if ([bool]$cfg.dryRun) { $wArgs += '-DryRun' }
            if ($watcherReconcileNext) {
                $wArgs += '-ReconcileStatusSetsFirst'
                $watcherReconcileNext = $false
            }

            $watcherChild = _Start-Child -ScriptPath $watcher -ScriptArgs $wArgs -Mta
            $state.awaitingWatcherSpawn = $false
            $state.watcherAlive = $true
            $state.passPipelineActive = $true
            try { $state.watcherPid = [int]$watcherChild.process.Id } catch { $state.watcherPid = 0 }
            _Render-Dashboard -Cfg $cfg -Force
        }

        $watcherAlive = $false
        try { $watcherAlive = -not $watcherChild.process.HasExited } catch { $watcherAlive = $false }
        $state.watcherAlive = $watcherAlive
        if ($watcherAlive) {
            _Poll-Child -Child $watcherChild -Cfg $cfg -Kind 'watcher'
        } else {
            try { _Poll-Child -Child $watcherChild -Cfg $cfg -Kind 'watcher' } catch { }
        }

        $state.activeWorkerSlots = $workerSlots.Count

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
                if ($state.workers.ContainsKey($lbl)) {
                    $state.workers.Remove($lbl) | Out-Null
                }
                $deadLabels += $lbl
            }
        }
        foreach ($lbl in $deadLabels) { $workerSlots.Remove($lbl) | Out-Null }

        _MaybeRefreshQueue -Cfg $cfg -MinIntervalMs 1500

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
                            ts = Get-QCTimestamp
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

        Start-Sleep -Milliseconds 200
    } catch {
        $state.passPipelineActive = $false
        $state.lastError = [string]$_.Exception.Message
        $state.phase = 'ERROR (will retry)'
        try {
            $cfg = Get-QCAppSettingsConfig -Path $AppSettingsPath -DryRun:$DryRun.IsPresent
        } catch {
            $cfg = @{ dryRun = $false }
        }
        _Render-Dashboard -Cfg $cfg -Force
        if ($PollSeconds -gt 0) {
            Start-Sleep -Seconds ([Math]::Max(2, $PollSeconds))
        } else {
            # Default PollSeconds=0 otherwise tight-spins on repeated failures (watcher exit, spawn errors, etc.).
            Start-Sleep -Milliseconds 1500
        }
    }
}

