<#
.SYNOPSIS
One-shot filesystem watcher tick: detect → classify → enqueue.

.DESCRIPTION
Loads appsettings.json, scans configured watchFolders once, applies:
  - Core.Paths (normalize)
  - QC.Filters (allow/deny)
  - QC.Triggers (match/no match)
  - QC.JobFactory (build job + dedupe key)
  - QC.Queue.Json (enqueue)

Constraints:
  - No ProjectWise writes
  - No processor execution
  - Run-once only (no continuous loop)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'appsettings.json'),

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [int]$MaxFiles = 0,

    # When set, the watcher first walks every workspace under statusSet.localRoot
    # and reconciles its local _StatusSet.pdf back to ProjectWise (legacy parity:
    # every restart re-checks every manifest). Pass this on the first invocation
    # after restart; omit it from subsequent watcher ticks to skip the re-walk.
    [Parameter(Mandatory = $false)]
    [switch]$ReconcileStatusSetsFirst
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Paths.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Filters.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Triggers.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.JobFactory.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force
$pwConnPath = (Join-Path $repoRoot 'modules\PW.Connection.psm1')
if (-not (Test-Path -LiteralPath $pwConnPath)) {
    throw "PW.Connection.psm1 not found at expected path: $pwConnPath"
}
Import-Module $pwConnPath -Force | Out-Null
Import-Module (Join-Path $repoRoot 'modules\QC.StatusSet.psm1') -Force

function _Sha256FileHex([string]$Path) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $hash = $sha.ComputeHash($fs)
        } finally {
            $fs.Dispose()
        }
    } finally {
        $sha.Dispose()
    }
    return -join ($hash | ForEach-Object { $_.ToString('x2') })
}

function _Sha256TextHex([string]$Text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
    } finally {
        $sha.Dispose()
    }
    return -join ($hash | ForEach-Object { $_.ToString('x2') })
}

function _PW-GetProp([object]$Obj, [string]$Name) {
    try {
        if ($null -eq $Obj) { return $null }
        if ($Obj.PSObject -and $Obj.PSObject.Properties[$Name]) { return $Obj.$Name }
    } catch { }
    return $null
}

function _PW-GetDocName([object]$Doc) {
    foreach ($n in @('Name', 'DocumentName')) {
        $v = _PW-GetProp -Obj $Doc -Name $n
        if ($v) { return [string]$v }
    }
    return ''
}

function _PW-GetDocDescription([object]$Doc) {
    foreach ($n in @('Description', 'DocumentDescription')) {
        $v = _PW-GetProp -Obj $Doc -Name $n
        if ($null -ne $v) { return [string]$v }
    }
    return ''
}

function _PW-GetDocLastModifiedUtcIso([object]$Doc) {
    foreach ($n in @('FileUpdatedDate','FileUpdateDate','DocumentUpdateDate','VersionModifiedDate','Version Modified Date')) {
        $v = _PW-GetProp -Obj $Doc -Name $n
        if ($v) {
            try {
                if ($v -is [DateTime]) { return $v.ToUniversalTime().ToString('o') }
                if ($v -is [DateTimeOffset]) { return $v.UtcDateTime.ToString('o') }
                $dt = [DateTime]::Parse([string]$v, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
                return $dt.ToUniversalTime().ToString('o')
            } catch {
                try { return ([string]$v).Trim() } catch { }
            }
        }
    }
    return ''
}

function _PW-DiscoverSheetsFoldersUnderRoot([string]$RootPath, [string]$SheetsSuffix, [string]$DatasourceName, [int]$ProjectDepth = 1) {
    # Discovers Sheets folders under a root by walking N levels of "project folders".
    # Depth=1 (default) matches legacy behavior: Root\<Project>\SheetsSuffix
    # Depth=2: Root\<Area>\<Project>\SheetsSuffix, etc.
    $rootPathRaw = ($RootPath -as [string]).Trim().TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($rootPathRaw)) { return @() }

    $suffix = ($SheetsSuffix -as [string]).Trim().TrimStart('\')
    if ([string]::IsNullOrWhiteSpace($suffix)) { $suffix = 'CADD\Sheets' }

    $depth = $ProjectDepth
    if ($depth -lt 1) { $depth = 1 }
    if ($depth -gt 5) { $depth = 5 } # hard cap to avoid runaway scans

    # Maintain internal paths as "Documents\..." so downstream filtering stays consistent.
    $rootDocs = $rootPathRaw
    if ($rootDocs -notmatch '^(?i)Documents\\') { $rootDocs = ('Documents\' + $rootDocs.TrimStart('\')) }

    function _PwFolderExists([string]$DocsFolderPath) {
        # Best-effort: returns $true if we can view the folder, otherwise $false.
        # Some PW cmdlet builds accept paths without "Documents\"; others require it.
        # IMPORTANT: do not build "Documents\<apiPath>" when <apiPath> was derived by stripping an
        # already-present Documents\ prefix — that yields Documents\Documents\... and false negatives.
        $internal = ($DocsFolderPath -as [string]).Trim().TrimEnd('\')
        if ([string]::IsNullOrWhiteSpace($internal)) { return $false }
        $apiPath = _PW-ToPwCmdletFolderPath -InternalFolderPath $internal
        if ([string]::IsNullOrWhiteSpace($apiPath)) { return $false }

        $candidates = New-Object System.Collections.Generic.List[string]
        if ($internal -match '^(?i)Documents\\') { [void]$candidates.Add($internal) }
        [void]$candidates.Add($apiPath)
        if ($internal -notmatch '^(?i)Documents\\') { [void]$candidates.Add(('Documents\' + $apiPath)) }

        foreach ($p in ($candidates | Select-Object -Unique)) {
            if ([string]::IsNullOrWhiteSpace($p)) { continue }
            try {
                $null = Get-PWFolderView -FolderPath $p -WarningAction SilentlyContinue -ErrorAction Stop
                return $true
            } catch {
                try {
                    $f = Get-PWFolders -FolderPath $p -JustOne -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                    if ($f) { return $true }
                } catch { }
            }
        }
        return $false
    }

    function _ListChildNames([string]$DocsFolderPath) {
        $internal = ($DocsFolderPath -as [string]).Trim().TrimEnd('\')
        if ([string]::IsNullOrWhiteSpace($internal)) { return @() }
        $apiPath = _PW-ToPwCmdletFolderPath -InternalFolderPath $internal
        if ([string]::IsNullOrWhiteSpace($apiPath)) { return @() }
        $names = @()

        $candidates = New-Object System.Collections.Generic.List[string]
        if ($internal -match '^(?i)Documents\\') { [void]$candidates.Add($internal) }
        [void]$candidates.Add($apiPath)
        if ($internal -notmatch '^(?i)Documents\\') { [void]$candidates.Add(('Documents\' + $apiPath)) }

        foreach ($p in ($candidates | Select-Object -Unique)) {
            if ([string]::IsNullOrWhiteSpace($p)) { continue }
            try {
                $view = Get-PWFolderView -FolderPath $p -WarningAction SilentlyContinue -ErrorAction Stop
                if ($view.Children) {
                    foreach ($c in $view.Children) {
                        $name = _PW-GetProp -Obj $c -Name 'Name'
                        if (-not $name) {
                            $fp = _PW-GetProp -Obj $c -Name 'FolderPath'
                            if ($fp) { $name = [System.IO.Path]::GetFileName(([string]$fp).TrimEnd('\')) }
                        }
                        if ($name) { $names += [string]$name }
                    }
                }
                if ($names.Count -eq 0 -and $view.Folders) {
                    foreach ($f in $view.Folders) {
                        $name = _PW-GetProp -Obj $f -Name 'Name'
                        if ($name) { $names += [string]$name }
                    }
                }
            } catch {
                try {
                    $children = Get-PWFoldersImmediateChildren -FolderPath $p -WarningAction SilentlyContinue -ErrorAction Stop
                    foreach ($c in @($children)) {
                        $name = _PW-GetProp -Obj $c -Name 'Name'
                        if (-not $name) {
                            $fp = _PW-GetProp -Obj $c -Name 'FolderPath'
                            if ($fp) { $name = [System.IO.Path]::GetFileName(([string]$fp).TrimEnd('\')) }
                        }
                        if ($name) { $names += [string]$name }
                    }
                } catch { }
            }

            if ($names.Count -gt 0) { break }
        }

        return @($names | Where-Object { $_ } | Select-Object -Unique)
    }

    # Two modes:
    #  1) If RootPath\SheetsSuffix exists, treat RootPath as the *project* and walk
    #     depth under the Sheets folder (e.g. CADD\Sheets\01-Title\...).
    #  2) Otherwise, treat RootPath as a *portfolio root* and discover project folders
    #     under it (legacy), then append SheetsSuffix.
    $rootSheets = (($rootDocs.TrimEnd('\') + '\' + $suffix).TrimEnd('\'))
    if (_PwFolderExists -DocsFolderPath $rootSheets) {
        # In "project has a Sheets folder" mode, interpret depth as the overall depth
        # from the project root:
        #   depth=1 => <project>\CADD\Sheets
        #   depth=2 => <project>\CADD\Sheets\<child>
        #   depth=3 => <project>\CADD\Sheets\<child>\<grandchild>
        # So the number of levels to walk under the Sheets folder is (depth - 1).
        $walkDepth = ($depth - 1)
        if ($walkDepth -lt 0) { $walkDepth = 0 }

        $all = @($rootSheets)
        $current = @($rootSheets)
        for ($i = 1; $i -le $walkDepth; $i++) {
            $next = @()
            foreach ($p in @($current)) {
                foreach ($n in @(_ListChildNames -DocsFolderPath $p)) {
                    $next += (($p.TrimEnd('\') + '\' + $n).TrimEnd('\'))
                }
            }
            $current = @($next | Select-Object -Unique)
            if ($current.Count -eq 0) { break }
            $all += $current
        }

        $list = @()
        foreach ($sPath in @($all | Select-Object -Unique)) {
            $list += @{ DatasourceName = $DatasourceName; FolderPath = $sPath; OneLevelDeep = $true }
        }
        return $list
    }

    # Legacy: Build all project folder paths at the requested depth (under RootPath),
    # then append SheetsSuffix.
    $levelProjects = @($rootDocs)
    for ($i = 1; $i -le $depth; $i++) {
        $next = @()
        foreach ($p in @($levelProjects)) {
            foreach ($n in @(_ListChildNames -DocsFolderPath $p)) {
                $next += (($p.TrimEnd('\') + '\' + $n).TrimEnd('\'))
            }
        }
        $levelProjects = @($next | Select-Object -Unique)
        if ($levelProjects.Count -eq 0) { break }
    }

    $list = @()
    foreach ($projPath in @($levelProjects)) {
        $folderPath = (($projPath.TrimEnd('\') + '\' + $suffix).TrimEnd('\'))
        $list += @{ DatasourceName = $DatasourceName; FolderPath = $folderPath; OneLevelDeep = $true }
    }
    return $list
}

# Watch list keeps "Documents\..." paths; PW cmdlets expect the same logical path WITHOUT a leading Documents\ segment.
# Also collapse accidental "Documents\Documents\..." (e.g. oneLevelDeep merge when PW already returns FolderPath with Documents).
function _PW-ToPwCmdletFolderPath([string]$InternalFolderPath) {
    $s = ($InternalFolderPath -as [string]).Trim().TrimEnd('\')
    while ($s -match '^(?i)Documents\\') { $s = $s -replace '^(?i)Documents\\', '' }
    return $s
}

function _PW-CanonicalDocumentsFolderPath([string]$FolderPathProperty) {
    $t = ($FolderPathProperty -as [string]).Trim().TrimStart('\').TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }
    if ($t -match '^(?i)documents\\') { return $t }
    return ('Documents\' + $t)
}

function _PW-ListDocsInFolder([string]$FolderPath) {
    # Read-only doc listing; mirror legacy prepend_qc_on_trigger.ps1 behavior so Description is populated.
    try {
        $allDocs = @()

        $view = $null
        try { $view = Get-PWFolderView -FolderPath $FolderPath -ErrorAction Stop } catch { }
        if ($view -and $view.Documents) {
            $allDocs = @($view.Documents)
        } elseif ($view -and $view.Children) {
            $allDocs = @($view.Children | Where-Object { $_.DocumentID -or $_.Name })
        }

        if ($allDocs.Count -eq 0) {
            try {
                $allDocs = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath -ErrorAction Stop)
            } catch {
                $cmd = Get-Command -Name Get-PWDocumentsBySearchWithReturnColumns -ErrorAction SilentlyContinue
                if ($cmd) {
                    $returnColsParam = if ($cmd.Parameters.ContainsKey('ReturnColumns')) { 'ReturnColumns' } elseif ($cmd.Parameters.ContainsKey('ColumnsToReturn')) { 'ColumnsToReturn' } else { $null }
                    if ($returnColsParam) {
                        $cols = @('Description', 'Name', 'DocumentID')
                        if ($returnColsParam -eq 'ReturnColumns') {
                            $withCols = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -ReturnColumns $cols -ErrorAction SilentlyContinue
                        } else {
                            $withCols = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -ColumnsToReturn $cols -ErrorAction SilentlyContinue
                        }
                        if ($withCols) { $allDocs = @($withCols) }
                    }
                }
            }
        }

        return @($allDocs)
    } catch {
        # Fallback: folder view (can be limited)
        try {
            $folder = Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction SilentlyContinue
            if ($folder) {
                $view = $folder | Get-PWFolderView -ErrorAction SilentlyContinue
                if ($view -and $view.Documents) { return @($view.Documents) }
            }
        } catch { }
    }
    return @()
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
        foreach ($p in $Value.PSObject.Properties) {
            $h[$p.Name] = (_ToHashtable $p.Value)
        }
        return $h
    }
    return $Value
}

function _Read-AppSettingsJson([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return New-QCFailureResult -Code 'CONFIG_MISSING_FILE' -Message "appsettings.json not found: $Path" -Data @{ path = $Path }
    }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $cfg = _ToHashtable $obj
        return New-QCSuccessResult -Code 'CONFIG_LOADED' -Message 'Config loaded.' -Data @{ config = $cfg; path = $Path }
    } catch {
        return New-QCFailureResult -Code 'CONFIG_PARSE_ERROR' -Message 'Failed to read/parse appsettings.json.' -Data @{ path = $Path; error = $_ }
    }
}

function _Log([string]$Level, [string]$Code, [string]$Message, [hashtable]$Data) {
    $ts = [DateTime]::UtcNow.ToString('o')
    $payload = @{
        ts = $ts
        level = $Level
        code = $Code
        message = $Message
        data = $Data
    } | ConvertTo-Json -Depth 20 -Compress
    # IMPORTANT: write + flush stdout so dashboard redirection can ingest progress live.
    [Console]::Out.WriteLine($payload)
    [Console]::Out.Flush()
}

$cfgRes = _Read-AppSettingsJson -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

if (-not $config.ContainsKey('dryRun')) { $config['dryRun'] = $false }
if ($DryRun.IsPresent) { $config['dryRun'] = $true }
$isDryRun = [bool]$config['dryRun']

$ignoreSampleEvery = 50
if ($config.ContainsKey('logging') -and $config.logging -and $config.logging.ContainsKey('ignoredSampleEvery') -and $config.logging.ignoredSampleEvery) {
    $ignoreSampleEvery = [int]$config.logging.ignoredSampleEvery
}
if ($ignoreSampleEvery -lt 1) { $ignoreSampleEvery = 1 }

if (-not $config.ContainsKey('watchFolders') -or -not $config.watchFolders) {
    $config['watchFolders'] = @()
}

$watchFolders = @($config.watchFolders | ForEach-Object { ($_ -as [string]).Trim() } | Where-Object { $_ })
$hasPwWatchList = ($config.ContainsKey('projectWise') -and $config.projectWise -and ($config.projectWise.ContainsKey('watchList') -and $config.projectWise.watchList))
if ($watchFolders.Count -eq 0 -and -not $hasPwWatchList) { throw "watchFolders is empty and projectWise.watchList not configured." }

_Log -Level 'Information' -Code 'WATCH_START' -Message 'Watch run started.' -Data @{
    appSettingsPath = $AppSettingsPath
    dryRun = $isDryRun
    watchFolderCount = $watchFolders.Count
    maxFiles = $MaxFiles
    ignoredSampleEvery = $ignoreSampleEvery
}

$accepted = 0
$ignored = 0
$filtered = 0
$matched = 0
$enqueued = 0
$duplicates = 0
$skippedStatusSetCurrent = 0
$errors = 0

$statusSetRules = @()
try {
    if ($config.ContainsKey('triggers') -and $config.triggers -and $config.triggers.ContainsKey('rules') -and $config.triggers.rules) {
        foreach ($r in @($config.triggers.rules)) {
            $rh = _ToHashtable $r
            if (-not ($rh -is [hashtable])) { continue }
            if (-not ($rh.ContainsKey('enabled'))) { $rh['enabled'] = $true }
            if (-not [bool]$rh.enabled) { continue }
            if (($rh.ContainsKey('jobType') -and ([string]$rh.jobType) -eq 'STATUS_SET_GEN') -and $rh.ContainsKey('grouping') -and $rh.grouping) {
                $g = _ToHashtable $rh.grouping
                if ($g -is [hashtable]) {
                    $gEnabled = $false
                    try { $gEnabled = [bool]$g.enabled } catch { $gEnabled = $false }
                    $gBy = $null
                    if ($g.ContainsKey('groupBy') -and $g.groupBy) { $gBy = ([string]$g.groupBy).Trim().ToLowerInvariant() }
                    if ($gEnabled -and $gBy -eq 'folder') {
                        $statusSetRules += $rh
                    }
                }
            }
        }
    }
} catch { }

# ProjectWise watchList processing (STATUS_SET_GEN and/or QC_PREPEND).
# This must run even when STATUS_SET_GEN rules are disabled, because QC_PREPEND can be PW-triggered too.
if ($statusSetRules.Count -ge 0) {
    # ProjectWise sources (watchList) — read-only.
    if ($hasPwWatchList) {
        try {
            $pwCfg = _ToHashtable $config.projectWise
            $ds = if ($pwCfg.ContainsKey('datasourceName') -and $pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { '' }
            $credPath = if ($pwCfg.ContainsKey('credentialPath') -and $pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
            $watchList = _ToHashtable $pwCfg.watchList

            # Re-import here to avoid any odd module/session state where exports are not visible.
            Import-Module $pwConnPath -Force | Out-Null
            $credRes = Get-PWCredentialFromFile -CredentialPath $credPath
            if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }
            _Log -Level 'Information' -Code 'WATCH_PW_CONNECT_START' -Message 'Connecting to ProjectWise.' -Data @{
                datasourceName = $ds
                credentialPath = $credPath
            }
            $connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
            if (-not $connRes.IsSuccess) { throw ($connRes.Code + ': ' + $connRes.Message) }
            _Log -Level 'Information' -Code 'WATCH_PW_CONNECT_OK' -Message 'Connected to ProjectWise.' -Data @{
                datasourceName = $ds
                userName = if ($credRes.Data -and $credRes.Data.userName) { [string]$credRes.Data.userName } else { '' }
            }

            # One-shot reconciliation: walk every locally-built _StatusSet.pdf
            # and push to PW when the local copy is newer / PW is missing it.
            # Gated by -ReconcileStatusSetsFirst so it runs once per restart,
            # not on every watcher tick (and never blocks normal triggering).
            if ($ReconcileStatusSetsFirst.IsPresent) {
                _Log -Level 'Information' -Code 'WATCH_RECONCILE_START' -Message 'Reconciling local status sets to ProjectWise.' -Data @{}
                try {
                    $cb = {
                        param($evt)
                        $level = if ([bool]$evt.isSuccess) { 'Information' } else { 'Warning' }
                        $code  = "WATCH_RECONCILE_$([string]$evt.code -replace '^STATUS_SET_RECONCILE_','')"
                        _Log -Level $level -Code $code -Message ([string]$evt.message) -Data @{
                            workspaceDir = [string]$evt.workspaceDir
                            pwFolder     = [string]$evt.pwFolder
                            sheetsFolder = [string]$evt.sheetsFolder
                            outputPdf    = [string]$evt.outputPdf
                            data         = $evt.data
                        }
                    }
                    $rec = Invoke-StatusSetReconcile -Config $config -LogCallback $cb
                    if ($rec.IsSuccess) {
                        _Log -Level 'Information' -Code 'WATCH_RECONCILE_DONE' -Message 'Reconciliation completed.' -Data @{
                            counts   = $rec.Data.counts
                            failures = $rec.Data.failures
                            skipped  = $rec.Data.skipped
                        }
                    } else {
                        _Log -Level 'Error' -Code 'WATCH_RECONCILE_FAILED' -Message ([string]$rec.Message) -Data @{ code = [string]$rec.Code }
                    }
                } catch {
                    _Log -Level 'Error' -Code 'WATCH_RECONCILE_FAILED' -Message ('Reconciliation threw: ' + $_.Exception.Message) -Data @{}
                }
            }

            $pwFolders = @()
            if ($watchList -and $watchList.ContainsKey('roots') -and $watchList.roots) {
                foreach ($r in @($watchList.roots)) {
                    $rh = _ToHashtable $r
                    if (-not ($rh -is [hashtable])) { continue }
                    $rootPath = [string]$rh.path
                    $suffix = if ($rh.ContainsKey('sheetsPathFromProject') -and $rh.sheetsPathFromProject) { [string]$rh.sheetsPathFromProject } else { 'CADD\Sheets' }
                    $projectDepth = 1
                    if ($rh.ContainsKey('projectDepth') -and $null -ne $rh.projectDepth) {
                        try { $projectDepth = [int]$rh.projectDepth } catch { $projectDepth = 1 }
                    }
                    $enableQcPrepend = $false
                    if ($rh.ContainsKey('enableQcPrepend')) { try { $enableQcPrepend = [bool]$rh.enableQcPrepend } catch { $enableQcPrepend = $false } }
                    $enableStatusSet = $false
                    if ($rh.ContainsKey('enableStatusSet')) { try { $enableStatusSet = [bool]$rh.enableStatusSet } catch { $enableStatusSet = $false } }
                    $discovered = @(_PW-DiscoverSheetsFoldersUnderRoot -RootPath $rootPath -SheetsSuffix $suffix -DatasourceName $ds -ProjectDepth $projectDepth)
                    foreach ($d in $discovered) {
                        $d['EnableQcPrepend'] = $enableQcPrepend
                        $d['EnableStatusSet'] = $enableStatusSet
                        $pwFolders += $d
                    }
                }
            }
            if ($watchList -and $watchList.ContainsKey('folders') -and $watchList.folders) {
                foreach ($f in @($watchList.folders)) {
                    $fh = _ToHashtable $f
                    if (-not ($fh -is [hashtable])) { continue }
                    $root = [string]$fh.root
                    $path = [string]$fh.path
                    $oneLevelDeep = $false
                    if ($fh.ContainsKey('oneLevelDeep')) { try { $oneLevelDeep = [bool]$fh.oneLevelDeep } catch { $oneLevelDeep = $false } }
                    $enableQcPrepend = $false
                    if ($fh.ContainsKey('enableQcPrepend')) { try { $enableQcPrepend = [bool]$fh.enableQcPrepend } catch { $enableQcPrepend = $false } }
                    $enableStatusSet = $false
                    if ($fh.ContainsKey('enableStatusSet')) { try { $enableStatusSet = [bool]$fh.enableStatusSet } catch { $enableStatusSet = $false } }
                    $full = ($root.TrimEnd('\') + '\' + $path.TrimStart('\')).Trim()
                    $pwFolders += @(@{
                        DatasourceName = $ds
                        FolderPath = $full
                        OneLevelDeep = $oneLevelDeep
                        EnableQcPrepend = $enableQcPrepend
                        EnableStatusSet = $enableStatusSet
                    })
                }
            }

            # Expand oneLevelDeep for explicit folders
            $expanded = @()
            foreach ($e in @($pwFolders)) {
                $expanded += $e
                try {
                    if ($e.OneLevelDeep) {
                        $fp = [string]$e.FolderPath
                        $apiPath = _PW-ToPwCmdletFolderPath -InternalFolderPath $fp
                        _Log -Level 'Information' -Code 'WATCH_PW_ONELEVEL_EXPAND_PROGRESS' -Message 'Querying ProjectWise for discipline subfolders under Sheets.' -Data @{
                            folder = $fp
                            inProgress = $true
                        }
                        $kids = @(Get-PWImmediateChildFolders -FolderPath $apiPath)
                        _Log -Level 'Information' -Code 'WATCH_PW_ONELEVEL_EXPAND_PROGRESS' -Message 'Discipline subfolder listing completed.' -Data @{
                            folder = $fp
                            inProgress = $false
                            childCount = [int]@($kids).Count
                        }
                        if (@($kids).Count -eq 0) {
                            _Log -Level 'Information' -Code 'WATCH_PW_ONELEVEL_NO_CHILDREN' -Message 'oneLevelDeep: no discipline subfolders under this Sheets path; only this folder will be scanned (normal for flat Sheets or empty areas).' -Data @{
                                folder = $fp
                                apiPath = $apiPath
                            }
                        }
                        foreach ($k in $kids) {
                            $kp = _PW-GetProp -Obj $k -Name 'FolderPath'
                            if ($kp) {
                                $canonical = _PW-CanonicalDocumentsFolderPath -FolderPathProperty ([string]$kp)
                                if (-not $canonical) { continue }
                                $expanded += @{
                                    DatasourceName = $ds
                                    FolderPath = $canonical
                                    OneLevelDeep = $false
                                    EnableQcPrepend = [bool]$e.EnableQcPrepend
                                    EnableStatusSet = [bool]$e.EnableStatusSet
                                }
                            }
                        }
                    }
                } catch {
                    _Log -Level 'Warning' -Code 'WATCH_PW_ONELEVEL_EXPAND_FAILED' -Message ('oneLevelDeep expansion failed: ' + $_.Exception.Message) -Data @{
                        folder = [string]$e.FolderPath
                    }
                }
            }
            $pwFolders = $expanded
            _Log -Level 'Information' -Code 'WATCH_PW_FOLDERS' -Message 'ProjectWise watch folders prepared.' -Data @{
                folderCount = [int]$pwFolders.Count
                sample = @($pwFolders | Select-Object -First 5 | ForEach-Object { [string]$_.FolderPath })
            }

            # Select-Object -First 1 already yields a single hashtable (or $null). Avoid @() which forces object[].
            $statusRuleObj = ($statusSetRules | Sort-Object -Property priority | Select-Object -First 1)
            foreach ($entry in @($pwFolders)) {
                try {
                    $fp = [string]$entry.FolderPath
                    if ([string]::IsNullOrWhiteSpace($fp)) { continue }

                    $oneLevelDeep = $false
                    $enableQcPrepend = $false
                    $enableStatusSet = $false
                    try { $oneLevelDeep = [bool]$entry.OneLevelDeep } catch { $oneLevelDeep = $false }
                    try { $enableQcPrepend = [bool]$entry.EnableQcPrepend } catch { $enableQcPrepend = $false }
                    try { $enableStatusSet = [bool]$entry.EnableStatusSet } catch { $enableStatusSet = $false }

                    # Emit a "scan start" event even if filters later skip the folder.
                    _Log -Level 'Information' -Code 'WATCH_PW_SCAN_START' -Message 'PW scanning folder.' -Data @{
                        folder = $fp
                        oneLevelDeep = $oneLevelDeep
                        enableQcPrepend = $enableQcPrepend
                        enableStatusSet = $enableStatusSet
                    }

                    # STATUS_SET_GEN (folder-level)
                    if ($enableStatusSet -and $statusRuleObj) {
                        $allowRes = Test-QCPathAllowed -CandidatePath $fp -Config $config
                        if (-not $allowRes.IsSuccess) { throw $allowRes.Message }
                        if (-not [bool]$allowRes.Data.allowed) {
                            $filtered++
                            _Log -Level 'Information' -Code 'WATCH_PW_FOLDER_DONE' -Message 'PW folder skipped by filters.' -Data @{
                                folder = $fp
                                reason = 'filtered'
                                enableQcPrepend = $enableQcPrepend
                                enableStatusSet = $enableStatusSet
                            }
                            continue
                        }

                        _Log -Level 'Information' -Code 'WATCH_PW_STATUSSET_SCAN_START' -Message 'PW status-set folder query started.' -Data @{
                            folder = $fp
                            oneLevelDeep = $oneLevelDeep
                        }
                        $state = Get-StatusSetPWFolderState -FolderPath (_PW-ToPwCmdletFolderPath -InternalFolderPath $fp) -OneLevelDeep:$oneLevelDeep
                        _Log -Level 'Information' -Code 'WATCH_PW_STATUSSET_SCAN_DONE' -Message 'PW status-set folder query completed.' -Data @{
                            folder = $fp
                            oneLevelDeep = $oneLevelDeep
                            pdfCount = [int]$state.pdfCount
                            dgnCount = [int]$state.dgnCount
                            pairedCount = [int]$state.pairedCount
                        }
                        if ([int]$state.pairedCount -gt 0) {
                            $gateRes = Test-StatusSetWatcherShouldEnqueue -Config $config -SourceFolder $fp -FolderState $state
                            $skipUpToDate = ($gateRes.IsSuccess -and -not [bool]$gateRes.Data.shouldEnqueue)
                            if ($skipUpToDate) {
                                $skippedStatusSetCurrent++
                                _Log -Level 'Information' -Code 'WATCH_PW_STATUSSET_SKIP_CURRENT' -Message 'PW folder status set already current; not enqueueing STATUS_SET_GEN.' -Data @{
                                    folder = $fp
                                    pairedCount = [int]$state.pairedCount
                                    gateReason = [string]$gateRes.Data.gateReason
                                    workspaceDir = [string]$gateRes.Data.workspaceDir
                                    manifestPath = [string]$gateRes.Data.manifestPath
                                    compareReasons = if ($gateRes.Data.compare -and $gateRes.Data.compare.reasons) { @($gateRes.Data.compare.reasons) } else { @() }
                                }
                            } else {
                                $candidate = @{
                                    path = $fp
                                    fileName = '_folder_'
                                    description = ''
                                    detectedAtUtc = ([DateTime]::UtcNow.ToString('o'))
                                    sourceFolder = $fp
                                    datasourceName = $ds
                                    groupKey = ('STATUS_SET_GEN|' + $fp).ToLowerInvariant()
                                    folderStateHash = [string]$state.folderStateHash
                                    oneLevelDeep = $oneLevelDeep
                                    statusSet = @{
                                        pairedCount = [int]$state.pairedCount
                                        orderKey = [string]$state.orderKey
                                        pairedSheets = @($state.pairedSheets)
                                    }
                                    file = @{
                                        fullName = $fp
                                        length = 0
                                        lastWriteTimeUtc = ([DateTime]::UtcNow.ToString('o'))
                                    }
                                }

                                $jobRes = New-QCJobObject -Candidate $candidate -Rule $statusRuleObj -Config $config
                                if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
                                $job = [hashtable]$jobRes.Data.job

                                $accepted++
                                $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
                                if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
                                $wouldDedupe = [bool]$dupRes.Data.isDuplicate
                                $wouldEnqueue = (-not $wouldDedupe)
                                $enqueueSkippedReason = $null
                                if ($wouldDedupe) { $enqueueSkippedReason = 'duplicate' }
                                elseif ($isDryRun) { $enqueueSkippedReason = 'dryRun' }

                                _Log -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'PW folder change candidate accepted (STATUS_SET_GEN).' -Data @{
                                    jobId = [string]$job['id']
                                    jobType = [string]$job['type']
                                    dedupeKey = [string]$job['dedupeKey']
                                    sourcePath = [string]$job['sourcePath']
                                    sourceFolder = [string]$job['sourceFolder']
                                    groupKey = [string]$job['groupKey']
                                    triggeringFile = $fp
                                    ruleId = [string]$job['triggerRule']['id']
                                    dryRun = $isDryRun
                                    wouldEnqueue = $wouldEnqueue
                                    wouldDedupe = $wouldDedupe
                                    enqueueSkippedReason = $enqueueSkippedReason
                                    folderStateHash = [string]$candidate.folderStateHash
                                    pairedCount = [int]$state.pairedCount
                                }

                                if (-not $isDryRun -and -not $wouldDedupe) {
                                    $enqRes = Add-QCQueueJob -Job $job -Config $config
                                    if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
                                    $enqueued++
                                } elseif ($wouldDedupe) { $duplicates++ }
                            }
                        } elseif ([int]$state.pdfCount -gt 0 -or [int]$state.dgnCount -gt 0) {
                            _Log -Level 'Information' -Code 'WATCH_PW_STATUSSET_NO_PAIRS' -Message 'PW folder scanned but no PDF/DGN pairs found.' -Data @{
                                folder = $fp
                                oneLevelDeep = $oneLevelDeep
                                pdfCount = [int]$state.pdfCount
                                dgnCount = [int]$state.dgnCount
                            }
                        }
                    }

                    # QC_PREPEND (description tag)
                    if ([bool]$entry.EnableQcPrepend) {
                        _Log -Level 'Information' -Code 'WATCH_PW_DOC_SCAN_START' -Message 'PW folder doc query started.' -Data @{
                            folder = $fp
                        }
                        $docs = _PW-ListDocsInFolder -FolderPath (_PW-ToPwCmdletFolderPath -InternalFolderPath $fp)
                        $pdfDocs = @()
                        $withDesc = @()
                        $tagged = @()
                        foreach ($d in @($docs)) {
                            $n = _PW-GetDocName -Doc $d
                            if (-not $n -or -not ($n -match '(?i)\.pdf$')) { continue }
                            $pdfDocs += $d
                            $dd = _PW-GetDocDescription -Doc $d
                            if (-not [string]::IsNullOrWhiteSpace($dd)) {
                                $withDesc += $d
                                if ($dd.IndexOf('QC_Archivist', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                                    $tagged += $d
                                }
                            }
                        }
                        $sample = @()
                        foreach ($d in @($withDesc | Select-Object -First 2)) {
                            $sn = _PW-GetDocName -Doc $d
                            $sd = _PW-GetDocDescription -Doc $d
                            if ($sd -and $sd.Length -gt 160) { $sd = $sd.Substring(0, 160) }
                            $sample += (@{ name = $sn; description = $sd })
                        }
                        _Log -Level 'Information' -Code 'WATCH_PW_DOC_SCAN' -Message 'PW folder doc scan completed.' -Data @{
                            folder = $fp
                            docCount = [int](@($docs).Count)
                            pdfCount = [int](@($pdfDocs).Count)
                            withDescriptionCount = [int](@($withDesc).Count)
                            qcArchivistCount = [int](@($tagged).Count)
                            descriptionSample = $sample
                            propertyNamesSample = if (@($docs).Count -gt 0) { @($docs[0].PSObject.Properties | Select-Object -First 30 | ForEach-Object { $_.Name }) } else { @() }
                        }
                        foreach ($doc in @($docs)) {
                            $docName = _PW-GetDocName -Doc $doc
                            if (-not $docName -or -not ($docName -match '(?i)\.pdf$')) { continue }
                            $desc = _PW-GetDocDescription -Doc $doc
                            if ([string]::IsNullOrWhiteSpace($desc)) { continue }
                            if ($desc.IndexOf('QC_Archivist', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }

                            _Log -Level 'Information' -Code 'WATCH_PW_TAGGED' -Message 'PW doc has QC_Archivist tag.' -Data @{
                                folder = $fp
                                fileName = $docName
                                description = $desc
                            }

                            $mod = _PW-GetDocLastModifiedUtcIso -Doc $doc
                            $sz = _PW-GetProp -Obj $doc -Name 'FileSize'
                            if (-not $sz) { $sz = _PW-GetProp -Obj $doc -Name 'Size' }
                            $pseudo = _Sha256TextHex -Text (([string]$docName) + '|' + ([string]$mod) + '|' + ([string]$sz) + '|' + ([string]$fp))

                            $candidate = @{
                                path = ($fp.TrimEnd('\') + '\' + $docName)
                                fileName = $docName
                                description = $desc
                                detectedAtUtc = ([DateTime]::UtcNow.ToString('o'))
                                sourceFolder = $fp
                                datasourceName = $ds
                                file = @{
                                    fullName = ($fp.TrimEnd('\') + '\' + $docName)
                                    length = if ($sz) { [int64]$sz } else { 0 }
                                    lastWriteTimeUtc = $mod
                                    sha256 = $pseudo
                                }
                            }

                            $allowRes = Test-QCPathAllowed -CandidatePath ([string]$candidate.path) -Config $config
                            if (-not $allowRes.IsSuccess) { throw $allowRes.Message }
                            if (-not [bool]$allowRes.Data.allowed) {
                                $filtered++
                                continue
                            }

                            $matchRes = Test-QCTriggerCandidate -Candidate $candidate -Config $config
                            if (-not $matchRes.IsSuccess) { throw $matchRes.Message }
                            if (-not [bool]$matchRes.Data.matched) {
                                _Log -Level 'Information' -Code 'WATCH_PW_NO_MATCH' -Message 'PW doc had QC_Archivist but did not match any trigger rule.' -Data @{
                                    path = [string]$candidate.path
                                    fileName = [string]$candidate.fileName
                                    ruleReason = if ($matchRes.Data.ContainsKey('reason')) { [string]$matchRes.Data.reason } else { '' }
                                    candidateDescription = $desc
                                }
                                continue
                            }
                            $ruleObj = $matchRes.Data.rule
                            if ([string]$ruleObj.jobType -ne 'QC_PREPEND') { continue }

                            $jobRes = New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config
                            if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
                            $job = [hashtable]$jobRes.Data.job

                            $accepted++
                            $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
                            if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
                            $wouldDedupe = [bool]$dupRes.Data.isDuplicate
                            $wouldEnqueue = (-not $wouldDedupe)
                            $enqueueSkippedReason = $null
                            if ($wouldDedupe) { $enqueueSkippedReason = 'duplicate' }
                            elseif ($isDryRun) { $enqueueSkippedReason = 'dryRun' }

                            _Log -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'PW doc accepted (QC_PREPEND via description tag).' -Data @{
                                jobId = [string]$job['id']
                                jobType = [string]$job['type']
                                dedupeKey = [string]$job['dedupeKey']
                                sourcePath = [string]$job['sourcePath']
                                sourceFolder = [string]$job['sourceFolder']
                                triggeringFile = $docName
                                ruleId = [string]$job['triggerRule']['id']
                                dryRun = $isDryRun
                                wouldEnqueue = $wouldEnqueue
                                wouldDedupe = $wouldDedupe
                                enqueueSkippedReason = $enqueueSkippedReason
                            }

                            if (-not $isDryRun -and -not $wouldDedupe) {
                                $enqRes = Add-QCQueueJob -Job $job -Config $config
                                if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
                                $enqueued++
                            } elseif ($wouldDedupe) { $duplicates++ }
                        }
                    }
                    _Log -Level 'Information' -Code 'WATCH_PW_FOLDER_DONE' -Message 'PW folder processing completed.' -Data @{
                        folder = $fp
                        enableQcPrepend = $enableQcPrepend
                        enableStatusSet = $enableStatusSet
                    }
                } catch {
                    $errors++
                    $ex = $_.Exception
                    _Log -Level 'Error' -Code 'WATCH_PW_FOLDER_ERROR' -Message 'Error processing PW folder for STATUS_SET_GEN.' -Data @{
                        folder = [string]$entry.FolderPath
                        errorMessage = [string]$_.Exception.Message
                        errorType = if ($ex) { [string]$ex.GetType().FullName } else { '' }
                        scriptStackTrace = [string]$_.ScriptStackTrace
                    }
                }
            }

            Disconnect-PW | Out-Null
        } catch {
            $errors++
            _Log -Level 'Error' -Code 'WATCH_PW_ERROR' -Message 'ProjectWise watchList processing failed.' -Data @{ errorMessage = [string]$_.Exception.Message; scriptStackTrace = [string]$_.ScriptStackTrace }
        }
    }

    foreach ($folder in $watchFolders) {
        try {
            if (-not (Test-Path -LiteralPath $folder)) { continue }
            $state = Get-StatusSetLocalFolderState -RootFolder $folder
            if ([int]$state.pairedCount -le 0) { continue }

            $normFolderRes = Normalize-QCPath -Path $folder
            if (-not $normFolderRes.IsSuccess) { throw $normFolderRes.Message }
            $normFolder = [string]$normFolderRes.Data.path

            # use the highest-priority rule (lowest priority number wins; our triggers use higher=more important, but
            # this keeps deterministic selection if multiple rules exist)
            # NOTE: Select-Object -First 1 already returns a single hashtable (or $null).
            # Avoid @() which forces object[] and can break downstream Rule property access.
            $ruleObj = ($statusSetRules | Sort-Object -Property priority | Select-Object -First 1)
            $jobType = 'STATUS_SET_GEN'

            if (-not $ruleObj) {
                _Log -Level 'Warning' -Code 'WATCH_FS_STATUSSET_RULE_MISSING' -Message 'STATUS_SET_GEN rule not found/enabled; skipping folder status-set enqueue.' -Data @{
                    folder = $folder
                    normFolder = $normFolder
                    pairedCount = [int]$state.pairedCount
                }
                continue
            }

            $gateRes = Test-StatusSetWatcherShouldEnqueue -Config $config -SourceFolder $normFolder -FolderState $state
            $skipUpToDate = ($gateRes.IsSuccess -and -not [bool]$gateRes.Data.shouldEnqueue)
            if ($skipUpToDate) {
                $skippedStatusSetCurrent++
                _Log -Level 'Information' -Code 'WATCH_FS_STATUSSET_SKIP_CURRENT' -Message 'Local folder status set already current; not enqueueing STATUS_SET_GEN.' -Data @{
                    folder = $folder
                    normFolder = $normFolder
                    pairedCount = [int]$state.pairedCount
                    gateReason = [string]$gateRes.Data.gateReason
                    workspaceDir = [string]$gateRes.Data.workspaceDir
                }
                continue
            }

            $candidate = @{
                path = $normFolder
                fileName = '_folder_'
                description = ''
                detectedAtUtc = ([DateTime]::UtcNow.ToString('o'))
                sourceFolder = $normFolder
                groupKey = ($jobType + '|' + $normFolder).ToLowerInvariant()
                folderStateHash = [string]$state.folderStateHash
                oneLevelDeep = $false
                statusSet = @{
                    pairedCount = [int]$state.pairedCount
                    orderKey = [string]$state.orderKey
                    pairedSheets = @($state.pairedSheets)
                }
                file = @{
                    fullName = $folder
                    length = 0
                    lastWriteTimeUtc = ([DateTime]::UtcNow.ToString('o'))
                }
            }

            $jobRes = New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config
            if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
            $job = [hashtable]$jobRes.Data.job

            $accepted++
            $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
            if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
            $wouldDedupe = [bool]$dupRes.Data.isDuplicate
            $wouldEnqueue = (-not $wouldDedupe)
            $enqueueSkippedReason = $null
            if ($wouldDedupe) { $enqueueSkippedReason = 'duplicate' }
            elseif ($isDryRun) { $enqueueSkippedReason = 'dryRun' }

            _Log -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Folder change candidate accepted (STATUS_SET_GEN).' -Data @{
                jobId = [string]$job['id']
                jobType = [string]$job['type']
                dedupeKey = [string]$job['dedupeKey']
                sourcePath = [string]$job['sourcePath']
                sourceFolder = [string]$job['sourceFolder']
                groupKey = [string]$job['groupKey']
                triggeringFile = $folder
                ruleId = [string]$job['triggerRule']['id']
                dryRun = $isDryRun
                wouldEnqueue = $wouldEnqueue
                wouldDedupe = $wouldDedupe
                enqueueSkippedReason = $enqueueSkippedReason
                folderStateHash = [string]$candidate.folderStateHash
                pairedCount = [int]$state.pairedCount
                orderKeyLines = (if ($state.orderKey) { ([string]$state.orderKey -split "`n").Count } else { 0 })
            }

            if (-not $isDryRun -and -not $wouldDedupe) {
                $enqRes = Add-QCQueueJob -Job $job -Config $config
                if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
                $enqueued++
            } elseif ($wouldDedupe) {
                $duplicates++
            }
        } catch {
            $errors++
            $ex = $_.Exception
            _Log -Level 'Error' -Code 'WATCH_FOLDER_ERROR' -Message 'Error processing folder for STATUS_SET_GEN.' -Data @{
                folder = $folder
                errorMessage = [string]$_.Exception.Message
                errorType = if ($ex) { [string]$ex.GetType().FullName } else { '' }
                scriptStackTrace = [string]$_.ScriptStackTrace
            }
        }
    }
}

$fileItems = @()
foreach ($folder in $watchFolders) {
    if (-not (Test-Path -LiteralPath $folder)) {
        _Log -Level 'Warning' -Code 'WATCH_FOLDER_MISSING' -Message 'Watch folder missing.' -Data @{ folder = $folder }
        continue
    }
    $fileItems += Get-ChildItem -LiteralPath $folder -File -Recurse -ErrorAction SilentlyContinue
}

if ($MaxFiles -gt 0) { $fileItems = @($fileItems | Select-Object -First $MaxFiles) }

foreach ($fi in $fileItems) {
    try {
        $pathRes = Normalize-QCPath -Path ([string]$fi.FullName)
        if (-not $pathRes.IsSuccess) { throw $pathRes.Message }
        $normPath = [string]$pathRes.Data.path

        $allowRes = Test-QCPathAllowed -CandidatePath $normPath -Config $config
        if (-not $allowRes.IsSuccess) { throw $allowRes.Message }
        if (-not [bool]$allowRes.Data.allowed) {
            $filtered++
            continue
        }

        $candidate = @{
            path = $normPath
            fileName = [string]$fi.Name
            description = '' # local filesystem has no PW description; triggers should use filename/path/extension.
            detectedAtUtc = ([DateTime]::UtcNow.ToString('o'))
            file = @{
                fullName = [string]$fi.FullName
                length = [int64]$fi.Length
                lastWriteTimeUtc = $fi.LastWriteTimeUtc.ToString('o')
            }
        }
        $sfRes = Normalize-QCPath -Path ([string]$fi.DirectoryName)
        if (-not $sfRes.IsSuccess) { throw $sfRes.Message }
        $candidate.sourceFolder = [string]$sfRes.Data.path

        $matchRes = Test-QCTriggerCandidate -Candidate $candidate -Config $config
        if (-not $matchRes.IsSuccess) { throw $matchRes.Message }
        if (-not [bool]$matchRes.Data.matched) {
            $ignored++
            if (($ignored % $ignoreSampleEvery) -eq 0) {
                _Log -Level 'Information' -Code 'WATCH_IGNORED_SAMPLE' -Message 'Ignored file (no trigger match).' -Data @{
                    path = $normPath
                    fileName = $candidate.fileName
                    ignoredCount = $ignored
                }
            }
            continue
        }

        $matched++

        $ruleObj = $matchRes.Data.rule
        $jobType = [string]$ruleObj.jobType

        $groupingEnabled = $false
        $groupBy = $null
        $grouping = $null
        if ($ruleObj -and $ruleObj.grouping) { $grouping = _ToHashtable $ruleObj.grouping }
        if ($grouping -is [hashtable]) {
            try { $groupingEnabled = [bool]$grouping.enabled } catch { $groupingEnabled = $false }
            if ($grouping.ContainsKey('groupBy') -and $grouping.groupBy) { $groupBy = ([string]$grouping.groupBy).Trim().ToLowerInvariant() }
        }

        if (-not $groupingEnabled -or $groupBy -ne 'folder' -or $jobType -ne 'STATUS_SET_GEN') {
            # file-level workflows: compute a stable file hash for dedupe (read-only).
            $candidate.file.sha256 = _Sha256FileHex -Path ([string]$fi.FullName)
        } else {
            # grouped folder workflow: establish groupKey = jobType + sourceFolder
            $candidate.groupKey = ($jobType + '|' + [string]$candidate.sourceFolder).ToLowerInvariant()
        }

        $jobRes = New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config
        if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
        $job = [hashtable]$jobRes.Data.job

        $accepted++
        $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
        if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
        $wouldDedupe = [bool]$dupRes.Data.isDuplicate
        $wouldEnqueue = (-not $wouldDedupe)
        $enqueueSkippedReason = $null
        if ($wouldDedupe) {
            $enqueueSkippedReason = 'duplicate'
        } elseif ($isDryRun) {
            $enqueueSkippedReason = 'dryRun'
        }

        _Log -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Trigger matched; job accepted.' -Data @{
            jobId = [string]$job['id']
            jobType = [string]$job['type']
            dedupeKey = [string]$job['dedupeKey']
            sourcePath = [string]$job['sourcePath']
            sourceFolder = [string]$job['sourceFolder']
            groupKey = [string]$job['groupKey']
            triggeringFile = [string]$fi.FullName
            ruleId = [string]$job['triggerRule']['id']
            dryRun = $isDryRun
            wouldEnqueue = $wouldEnqueue
            wouldDedupe = $wouldDedupe
            enqueueSkippedReason = $enqueueSkippedReason
        }

        if ($isDryRun) { continue }

        if ($wouldDedupe) {
            $duplicates++
            continue
        }

        $enqRes = Add-QCQueueJob -Job $job -Config $config
        if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
        $enqueued++
    } catch {
        $errors++
        $ex = $_.Exception
        _Log -Level 'Error' -Code 'WATCH_FILE_ERROR' -Message 'Error processing file.' -Data @{
            file = [string]$fi.FullName
            errorMessage = [string]$_.Exception.Message
            errorType = if ($ex) { [string]$ex.GetType().FullName } else { '' }
            scriptStackTrace = [string]$_.ScriptStackTrace
        }
    }
}

_Log -Level 'Information' -Code 'WATCH_DONE' -Message 'Watch run completed.' -Data @{
    dryRun = $isDryRun
    scanned = $fileItems.Count
    filtered = $filtered
    ignored = $ignored
    matched = $matched
    accepted = $accepted
    duplicates = $duplicates
    skippedStatusSetCurrent = $skippedStatusSetCurrent
    enqueued = $enqueued
    errors = $errors
}

exit 0

