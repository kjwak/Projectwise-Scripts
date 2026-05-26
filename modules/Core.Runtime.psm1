# Core.Runtime.psm1
# Responsibility: Shared script runtime helpers for config conversion/loading, JSON log output, and timezone utilities.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

# ---------------------------------------------------------------------------
# Timezone utilities -- all timestamps standardized to Mountain Time (MST/MDT)
# ---------------------------------------------------------------------------

$script:_QCTimezone = [TimeZoneInfo]::FindSystemTimeZoneById('Mountain Standard Time')

function Get-QCTimestamp {
    <#
    .SYNOPSIS
    Returns the current time in Mountain Time as an ISO 8601 string with offset.
    #>
    $utcNow = [DateTime]::UtcNow
    $mt = [TimeZoneInfo]::ConvertTimeFromUtc($utcNow, $script:_QCTimezone)
    $offset = $script:_QCTimezone.GetUtcOffset($mt)
    return ([DateTimeOffset]::new($mt, $offset)).ToString('o')
}

function ConvertTo-QCTimestamp {
    <#
    .SYNOPSIS
    Converts a DateTime to Mountain Time as an ISO 8601 string with offset.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][DateTime]$DateTime)
    $utc = $DateTime.ToUniversalTime()
    $mt = [TimeZoneInfo]::ConvertTimeFromUtc($utc, $script:_QCTimezone)
    $offset = $script:_QCTimezone.GetUtcOffset($mt)
    return ([DateTimeOffset]::new($mt, $offset)).ToString('o')
}

function Format-QCTimestamp {
    <#
    .SYNOPSIS
    Parses an ISO 8601 string and formats it for display in Mountain Time (yyyy-MM-dd HH:mm:ss).
    #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$IsoString)
    if ([string]::IsNullOrWhiteSpace($IsoString)) { return '' }
    try {
        $dto = [DateTimeOffset]::Parse($IsoString, [System.Globalization.CultureInfo]::InvariantCulture)
        $mt = [TimeZoneInfo]::ConvertTime($dto, $script:_QCTimezone)
        return $mt.ToString('yyyy-MM-dd HH:mm:ss')
    } catch {
        try {
            $dt = [DateTime]::Parse($IsoString, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
            $mt = [TimeZoneInfo]::ConvertTimeFromUtc($dt.ToUniversalTime(), $script:_QCTimezone)
            return $mt.ToString('yyyy-MM-dd HH:mm:ss')
        } catch { return $IsoString }
    }
}

function Get-QCTimestampShort {
    <#
    .SYNOPSIS
    Returns current Mountain Time as a compact stamp for file naming (yyyyMMdd_HHmmss).
    #>
    $utcNow = [DateTime]::UtcNow
    $mt = [TimeZoneInfo]::ConvertTimeFromUtc($utcNow, $script:_QCTimezone)
    return $mt.ToString('yyyyMMdd_HHmmss')
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

function Read-QCAppSettings {
    <#
    .SYNOPSIS
    Reads appsettings.json and returns a QC result with a hashtable config.
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
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        return New-QCSuccessResult -Code 'CONFIG_LOADED' -Message 'Config loaded.' -Data @{ config = (ConvertTo-HashtableDeep -Value $obj); path = $Path }
    } catch {
        return New-QCFailureResult -Code 'CONFIG_PARSE_ERROR' -Message 'Failed to read/parse appsettings.json.' -Data @{ path = $Path; errorMessage = $_.Exception.Message; error = $_ }
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
    return $cfg
}

function Write-QCJsonLog {
    <#
    .SYNOPSIS
    Writes one structured JSON log event to stdout.
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

    if ($Flush.IsPresent) {
        [Console]::Out.WriteLine($payload)
        [Console]::Out.Flush()
    } else {
        Write-Host $payload
    }
}

Export-ModuleMember -Function *
