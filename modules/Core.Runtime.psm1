# Core.Runtime.psm1
# Responsibility: Shared script runtime helpers for config conversion/loading, JSON log output, and timezone utilities.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

# ---------------------------------------------------------------------------
# Timezone utilities -- wall-clock display from appsettings runtime.displayTimeZoneId
# ---------------------------------------------------------------------------

$script:_QCTimezoneId = 'US Mountain Standard Time'
$script:_QCTimezone = [TimeZoneInfo]::FindSystemTimeZoneById($script:_QCTimezoneId)

function Get-QCDisplayTimeZoneIdFromConfig {
    <#
    .SYNOPSIS
    Resolves Windows time zone id: runtime.displayTimeZoneId, then auditPoller.displayTimeZoneId, then module default.
    #>
    [CmdletBinding()]
    param([hashtable]$Config = @{})

    foreach ($sectionKey in @('runtime', 'auditPoller')) {
        try {
            if (-not $Config -or -not $Config.ContainsKey($sectionKey)) { continue }
            $sec = $Config[$sectionKey]
            if ($null -eq $sec) { continue }
            $id = $null
            if ($sec -is [hashtable] -and $sec.ContainsKey('displayTimeZoneId') -and $sec.displayTimeZoneId) {
                $id = [string]$sec.displayTimeZoneId
            } elseif ($sec.PSObject -and $null -ne $sec.displayTimeZoneId -and [string]$sec.displayTimeZoneId) {
                $id = [string]$sec.displayTimeZoneId
            }
            if (-not [string]::IsNullOrWhiteSpace($id)) { return $id.Trim() }
        } catch { }
    }
    return $script:_QCTimezoneId
}

function Set-QCDisplayTimeZone {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$TimeZoneId)
    $script:_QCTimezoneId = $TimeZoneId.Trim()
    $script:_QCTimezone = [TimeZoneInfo]::FindSystemTimeZoneById($script:_QCTimezoneId)
}

function Set-QCDisplayTimeZoneFromConfig {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)
    Set-QCDisplayTimeZone -TimeZoneId (Get-QCDisplayTimeZoneIdFromConfig -Config $Config)
}

function Get-QCDisplayTimeZone {
    <#
    .SYNOPSIS
    Active display time zone (set via Set-QCDisplayTimeZoneFromConfig after loading appsettings).
    #>
    if (-not $script:_QCTimezone) {
        $script:_QCTimezone = [TimeZoneInfo]::FindSystemTimeZoneById($script:_QCTimezoneId)
    }
    return $script:_QCTimezone
}

function Get-QCWallClockNow {
    <#
    .SYNOPSIS
    Current UTC instant as DateTime in the configured display time zone (Unspecified Kind, wall clock).
    #>
    return [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, (Get-QCDisplayTimeZone))
}

function Get-QCTimestamp {
    <#
    .SYNOPSIS
    Returns the current time in the configured display zone as an ISO 8601 string with offset.
    #>
    $utcNow = [DateTime]::UtcNow
    $tz = Get-QCDisplayTimeZone
    $local = [TimeZoneInfo]::ConvertTimeFromUtc($utcNow, $tz)
    $offset = $tz.GetUtcOffset($local)
    return ([DateTimeOffset]::new($local, $offset)).ToString('o')
}

function ConvertTo-QCTimestamp {
    <#
    .SYNOPSIS
    Converts a DateTime to the configured display zone as an ISO 8601 string with offset.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][DateTime]$DateTime)
    $utc = $DateTime.ToUniversalTime()
    $tz = Get-QCDisplayTimeZone
    $local = [TimeZoneInfo]::ConvertTimeFromUtc($utc, $tz)
    $offset = $tz.GetUtcOffset($local)
    return ([DateTimeOffset]::new($local, $offset)).ToString('o')
}

function Format-QCTimestamp {
    <#
    .SYNOPSIS
    Parses an ISO 8601 string and formats it for display in the configured zone (yyyy-MM-dd HH:mm:ss).
    #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$IsoString)
    if ([string]::IsNullOrWhiteSpace($IsoString)) { return '' }
    try {
        $dto = [DateTimeOffset]::Parse($IsoString, [System.Globalization.CultureInfo]::InvariantCulture)
        $tz = Get-QCDisplayTimeZone
        $local = [TimeZoneInfo]::ConvertTime($dto, $tz)
        return $local.ToString('yyyy-MM-dd HH:mm:ss')
    } catch {
        try {
            $dt = [DateTime]::Parse($IsoString, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
            $local = [TimeZoneInfo]::ConvertTimeFromUtc($dt.ToUniversalTime(), (Get-QCDisplayTimeZone))
            return $local.ToString('yyyy-MM-dd HH:mm:ss')
        } catch { return $IsoString }
    }
}

function Get-QCTimestampShort {
    <#
    .SYNOPSIS
    Returns current display-zone time as a compact stamp for file naming (yyyyMMdd_HHmmss).
    #>
    return (Get-QCWallClockNow).ToString('yyyyMMdd_HHmmss')
}

function Get-QCLogHourStamp {
    <#
    .SYNOPSIS
    Returns current display-zone time as an hour bucket for log rotation (yyyy-MM-dd_HH).
    #>
    return (Get-QCWallClockNow).ToString('yyyy-MM-dd_HH')
}

function Get-QCJsonLogFilePath {
    <#
    .SYNOPSIS
    Resolves the hourly JSON log path when QC_JSON_LOG_DIR is set (QC_JSON_LOG_TAG optional).
    #>
    [CmdletBinding()]
    param(
        [string]$HourStamp = '',
        [string]$Tag = ''
    )
    $dir = $env:QC_JSON_LOG_DIR
    if ([string]::IsNullOrWhiteSpace($dir)) { return $null }
    if ([string]::IsNullOrWhiteSpace($HourStamp)) { $HourStamp = Get-QCLogHourStamp }
    $tag = $Tag
    if ([string]::IsNullOrWhiteSpace($tag)) {
        $tag = if ($env:QC_JSON_LOG_TAG) { [string]$env:QC_JSON_LOG_TAG } else { 'qc' }
    }
    return (Join-Path $dir ("${tag}_${HourStamp}.jsonl"))
}

$script:_QCJsonLogWriter = $null
$script:_QCJsonLogWriterHour = $null
$script:_QCJsonLogWriterPath = $null
$script:_QCJsonLogFileLock = [object]::new()

function Close-QCJsonLogWriter {
    if ($script:_QCJsonLogWriter) {
        try {
            $script:_QCJsonLogWriter.Flush()
            $script:_QCJsonLogWriter.Dispose()
        } catch { }
        $script:_QCJsonLogWriter = $null
        $script:_QCJsonLogWriterHour = $null
        $script:_QCJsonLogWriterPath = $null
    }
}

function Write-QCJsonLogFileLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Line
    )
    if ([string]::IsNullOrWhiteSpace($env:QC_JSON_LOG_DIR)) { return $false }
    $hour = Get-QCLogHourStamp
    $path = Get-QCJsonLogFilePath -HourStamp $hour
    if (-not $path) { return $false }
    [System.Threading.Monitor]::Enter($script:_QCJsonLogFileLock)
    try {
        if ($script:_QCJsonLogWriterHour -ne $hour) {
            Close-QCJsonLogWriter
            $parent = Split-Path -Parent $path
            if (-not (Test-Path -LiteralPath $parent)) {
                New-Item -ItemType Directory -Path $parent -Force | Out-Null
            }
            $script:_QCJsonLogWriter = New-Object System.IO.StreamWriter($path, $true, [System.Text.UTF8Encoding]::new($false))
            $script:_QCJsonLogWriter.AutoFlush = $true
            $script:_QCJsonLogWriterHour = $hour
            $script:_QCJsonLogWriterPath = $path
        }
        $script:_QCJsonLogWriter.WriteLine($Line)
    } catch {
        Close-QCJsonLogWriter
        return $false
    } finally {
        [System.Threading.Monitor]::Exit($script:_QCJsonLogFileLock)
    }
    return $true
}

function ConvertTo-HashtableDeep {
    <#
    .SYNOPSIS
    Recursively converts PSCustomObject values into hashtables.
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [AllowNull()]
        [object]$Value
    )

    process {
        if ($null -eq $Value) { return $null }
        if ($Value -is [string]) { return $Value }
        if ($Value -is [System.ValueType]) { return $Value }
        if ($Value -is [System.Collections.IDictionary]) {
            $h = @{}
            foreach ($key in $Value.Keys) { $h[$key] = ConvertTo-HashtableDeep -Value $Value[$key] }
            return $h
        }
        if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
            $out = @()
            foreach ($i in $Value) { $out += (ConvertTo-HashtableDeep -Value $i) }
            return $out
        }
        if ($Value.PSObject -and $Value.PSObject.Properties) {
            $h = @{}
            foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = ConvertTo-HashtableDeep -Value $p.Value }
            return $h
        }
        return $Value
    }
}

function Remove-QCJsonComments {
    <#
    .SYNOPSIS
    Strips // line and /* block */ comments so appsettings.json can be documented inline.
  .DESCRIPTION
    String literals are preserved. Used by Read-QCAppSettings before ConvertFrom-Json.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Text)

    if ([string]::IsNullOrEmpty($Text)) { return $Text }

    $sb = New-Object System.Text.StringBuilder
    $len = $Text.Length
    $inString = $false
    $escape = $false
    $i = 0
    while ($i -lt $len) {
        $c = $Text[$i]
        if ($inString) {
            [void]$sb.Append($c)
            if ($escape) { $escape = $false }
            elseif ($c -eq '\') { $escape = $true }
            elseif ($c -eq '"') { $inString = $false }
            $i++
            continue
        }
        if ($c -eq '"') {
            $inString = $true
            [void]$sb.Append($c)
            $i++
            continue
        }
        if ($c -eq '/' -and ($i + 1) -lt $len) {
            $n = $Text[$i + 1]
            if ($n -eq '/') {
                $i += 2
                while ($i -lt $len -and $Text[$i] -ne "`n" -and $Text[$i] -ne "`r") { $i++ }
                continue
            }
            if ($n -eq '*') {
                $i += 2
                while ($i + 1 -lt $len -and -not ($Text[$i] -eq '*' -and $Text[$i + 1] -eq '/')) { $i++ }
                $i += 2
                continue
            }
        }
        [void]$sb.Append($c)
        $i++
    }
    return $sb.ToString()
}

function Merge-QCHashtableDeep {
    <#
    .SYNOPSIS
    Deep-merges overlay into base. Nested hashtables merge; scalars and arrays are replaced by overlay values.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Base,
        [AllowNull()][object]$Overlay
    )

    if ($null -eq $Overlay) { return $Base }
    if ($null -eq $Base) { return $Overlay }
    if ($Base -isnot [hashtable] -or $Overlay -isnot [hashtable]) { return $Overlay }

    $merged = @{}
    foreach ($key in $Base.Keys) { $merged[$key] = $Base[$key] }
    foreach ($key in $Overlay.Keys) {
        if ($merged.ContainsKey($key) -and $merged[$key] -is [hashtable] -and $Overlay[$key] -is [hashtable]) {
            $merged[$key] = Merge-QCHashtableDeep -Base $merged[$key] -Overlay $Overlay[$key]
        } else {
            $merged[$key] = $Overlay[$key]
        }
    }
    return $merged
}

function Resolve-QCAppSettingsMergeChain {
    <#
    .SYNOPSIS
    Returns ordered config file paths to load for a given appsettings path.
  .DESCRIPTION
    appsettings.json + optional appsettings.local.json + optional appsettings.secrets.json;
    appsettings.{profile}.json merges appsettings.json then profile then appsettings.{profile}.local.json then secrets.
    Other filenames load only that file (no automatic merge).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $fullPath = $Path
    if (Test-Path -LiteralPath $Path) {
        $fullPath = (Resolve-Path -LiteralPath $Path).Path
    } else {
        $fullPath = [System.IO.Path]::GetFullPath($Path)
    }

    $dir = [System.IO.Path]::GetDirectoryName($fullPath)
    $name = [System.IO.Path]::GetFileName($fullPath)
    $chain = [System.Collections.Generic.List[string]]::new()

    $secretsFile = Join-Path $dir 'appsettings.secrets.json'
    $appendSecrets = {
        if (Test-Path -LiteralPath $secretsFile) {
            $chain.Add((Resolve-Path -LiteralPath $secretsFile).Path)
        }
    }

    if ($name -eq 'appsettings.json') {
        $chain.Add($fullPath)
        $local = Join-Path $dir 'appsettings.local.json'
        if (Test-Path -LiteralPath $local) { $chain.Add((Resolve-Path -LiteralPath $local).Path) }
        & $appendSecrets
        return [string[]]$chain
    }

    if ($name -match '^appsettings\.([^.]+)\.json$') {
        $profile = $Matches[1]
        if ($profile -in @('local', 'secrets')) {
            $chain.Add($fullPath)
            return [string[]]$chain
        }
        $base = Join-Path $dir 'appsettings.json'
        if (-not (Test-Path -LiteralPath $base)) {
            throw "Profile config '$name' requires appsettings.json in the same directory: $dir"
        }
        $chain.Add((Resolve-Path -LiteralPath $base).Path)
        $chain.Add($fullPath)
        $profileLocal = Join-Path $dir ("appsettings.{0}.local.json" -f $profile)
        if (Test-Path -LiteralPath $profileLocal) { $chain.Add((Resolve-Path -LiteralPath $profileLocal).Path) }
        & $appendSecrets
        return [string[]]$chain
    }

    $chain.Add($fullPath)
    return [string[]]$chain
}

function Read-QCAppSettingsFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $json = Remove-QCJsonComments -Text $raw
    $obj = $json | ConvertFrom-Json -ErrorAction Stop
    return ConvertTo-HashtableDeep -Value $obj
}

function Read-QCAppSettings {
    <#
    .SYNOPSIS
    Reads appsettings.json (and profile/local overlays) and returns a QC result with a hashtable config.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return New-QCFailureResult -Code 'CONFIG_MISSING_FILE' -Message "appsettings.json not found: $Path" -Data @{ path = $Path }
    }

    try {
        $mergeChain = @(Resolve-QCAppSettingsMergeChain -Path $Path)
    } catch {
        return New-QCFailureResult -Code 'CONFIG_MERGE_ERROR' -Message $_.Exception.Message -Data @{ path = $Path }
    }

    $cfg = $null
    $loadedPaths = [System.Collections.Generic.List[string]]::new()
    foreach ($filePath in $mergeChain) {
        if (-not (Test-Path -LiteralPath $filePath)) {
            return New-QCFailureResult -Code 'CONFIG_MISSING_FILE' -Message "Config file not found: $filePath" -Data @{ path = $filePath; mergeChain = @($mergeChain) }
        }
        try {
            $layer = Read-QCAppSettingsFile -Path $filePath
            $cfg = if ($null -eq $cfg) { $layer } else { Merge-QCHashtableDeep -Base $cfg -Overlay $layer }
            $loadedPaths.Add($filePath) | Out-Null
        } catch {
            return New-QCFailureResult -Code 'CONFIG_PARSE_ERROR' -Message 'Failed to read/parse appsettings.json.' -Data @{
                path = $filePath
                mergeChain = @($mergeChain)
                errorMessage = $_.Exception.Message
                error = $_
            }
        }
    }

    if ($null -eq $cfg) { $cfg = @{} }
    Set-QCDisplayTimeZoneFromConfig -Config $cfg
    return New-QCSuccessResult -Code 'CONFIG_LOADED' -Message 'Config loaded.' -Data @{
        config = $cfg
        path = $Path
        mergeChain = @($loadedPaths)
    }
}


function Get-QCAppSettingsConfig {
    <#
    .SYNOPSIS
    Reads appsettings.json and returns the hashtable config or throws on failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [switch]$DryRun
    )

    $res = Read-QCAppSettings -Path $Path
    if (-not $res.IsSuccess) { throw $res.Message }
    $cfg = [hashtable]$res.Data.config
    if ($DryRun.IsPresent) {
        if (-not $cfg.ContainsKey('dryRun')) { $cfg['dryRun'] = $false }
        $cfg['dryRun'] = $true
    }
    Set-QCDisplayTimeZoneFromConfig -Config $cfg
    return $cfg
}

function Write-QCJsonLog {
    <#
    .SYNOPSIS
    Writes one structured JSON log event to stdout, or to hourly files when QC_JSON_LOG_DIR is set.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Level,
        [Parameter(Mandatory)]
        [string]$Code,
        [Parameter(Mandatory)]
        [string]$Message,
        [hashtable]$Data,
        [string]$WorkerLabel = '',
        [switch]$IncludeWorkerPid,
        [switch]$Flush
    )

    if (-not $Data) { $Data = @{} }
    if (-not [string]::IsNullOrWhiteSpace($WorkerLabel) -and -not $Data.ContainsKey('workerLabel')) {
        $Data['workerLabel'] = $WorkerLabel
    }
    if ($IncludeWorkerPid.IsPresent -and -not $Data.ContainsKey('workerPid')) { $Data['workerPid'] = $PID }

    $payload = @{
        ts      = Get-QCTimestamp
        level   = $Level
        code    = $Code
        message = $Message
        data    = $Data
    } | ConvertTo-Json -Depth 20 -Compress

    $fileSink = [bool](Write-QCJsonLogFileLine -Line $payload)
    if ($fileSink) { return }

    if ($Flush.IsPresent) {
        [Console]::Out.WriteLine($payload)
        [Console]::Out.Flush()
    } else {
        Write-Host $payload
    }
}

Export-ModuleMember -Function *
