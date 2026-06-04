<#
.SYNOPSIS
Resolve ProjectWise folder GUIDs via Get-PWFoldersByGUIDs and SQL (dms_proj).

.DESCRIPTION
Read-only probe for parent-folder GUID resolution used by the audit poller.
Compares cmdlet results with Select-PWSQL on dms_proj.o_projguid.

.PARAMETER FolderGuids
Folder GUIDs to resolve (default: sample from pw_folder_cache failures).

.PARAMETER AppSettingsPath
Path to appsettings.json.
#>
[CmdletBinding()]
param(
    [string[]]$FolderGuids = @(
        '3d5b6db8-d1b9-4aad-a863-7224f5bd94a4',
        '6c7a981c-c889-430b-8dba-53b3957ed247',
        '9475dfe8-1a85-46de-8986-3e59744591ca',
        'abb63939-f013-4853-9819-9765d282b5ab',
        'b4af91f6-ba9f-453b-a8c6-a1f3c9deda71'
    ),
    [string]$AppSettingsPath = ''
)

$ErrorActionPreference = 'Continue'

$scriptDir = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force

function _SqlCastGuidList([string[]]$Guids) {
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($g in @($Guids)) {
        $k = ($g -as [string]).Trim().Trim('{}').ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($k)) { continue }
        $escaped = $k -replace '''', ''''''
        [void]$parts.Add("CAST('$escaped' AS UNIQUEIDENTIFIER)")
    }
    return ($parts -join ',')
}

function _Prop([object]$o, [string]$n) {
    if ($null -eq $o) { return $null }
    try {
        if ($o.PSObject.Properties[$n]) { return $o.$n }
    } catch { }
    return $null
}

function _PrepareFolder([object]$folder) {
    if ($null -eq $folder) { return $null }
    try {
        $m = $folder.GetType().GetMethod('GetFullPath')
        if ($m) { $null = $m.Invoke($folder, @()) }
    } catch { }
    return $folder
}

function _FolderPathFromObject([object]$folder) {
    if ($null -eq $folder) { return $null }
    $folder = _PrepareFolder $folder
    if ($folder -is [string]) {
        $s = ([string]$folder).Trim()
        if ($s -match '\\') { return $s }
        return $null
    }
    foreach ($p in @('FolderPath', 'Path', 'FullPath', 'folderPath', 'CanonicalPath', 'PWPath')) {
        $v = _Prop $folder $p
        if ($v -and "$v" -match '\\') { return [string]$v }
    }
    $names = @()
    try { $names = @($folder.PSObject.Properties.Name) } catch { }
    return @{ path = $null; propertyNames = $names; typeName = $folder.GetType().FullName }
}

function _RunPwSql([string]$Sql, [string]$Label) {
    $warnBefore = @($WarningPreference)
    $WarningPreference = 'Continue'
    $errs = $null
    $dt = Select-PWSQL -SQLSelectStatement $Sql -ErrorVariable errs -ErrorAction SilentlyContinue
    foreach ($w in @($errs)) {
        if ($w) { Write-Host "  SQL warning ($Label): $w" -ForegroundColor Yellow }
    }
    if (-not $dt) {
        Write-Host "  SQL ($Label): no result table" -ForegroundColor Yellow
        return $null
    }
    return $dt
}

$configRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $configRes.IsSuccess) {
    Write-Error $configRes.Message
    exit 1
}
$config = $configRes.Data.config
$ds = [string]$config.projectWise.datasourceName
$credPath = [string]$config.projectWise.credentialPath

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) {
    Write-Error $credRes.Message
    exit 1
}

$conn = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $conn.IsSuccess) {
    Write-Error "Connect-PW failed: $($conn.Message)"
    exit 1
}

try {
    Select-PWSQL -SQLSelectStatement 'SELECT 1 AS ok' -ErrorAction Stop | Out-Null
} catch {
    Write-Error "Connected but Select-PWSQL not available: $($_.Exception.Message)"
    exit 1
}

Write-Host "Connected: $ds" -ForegroundColor Green
Write-Host "Get-PWFoldersByGUIDs: $(Get-Command Get-PWFoldersByGUIDs -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source)"
Write-Host ''

$inCast = _SqlCastGuidList -Guids $FolderGuids
$sqlProj = "SELECT o_projguid, o_projectno, o_itemname, o_itemdesc FROM dms_proj WHERE o_projguid IN ($inCast)"
$sqlRows = _RunPwSql -Sql $sqlProj -Label 'dms_proj'
if ($sqlRows) {
    Write-Host "SQL dms_proj hits: $($sqlRows.Rows.Count) / $($FolderGuids.Count)" -ForegroundColor Cyan
    foreach ($row in $sqlRows.Rows) {
        Write-Host ("  proj {0} no={1} name={2}" -f $row.o_projguid, $row.o_projectno, $row.o_itemname)
    }
}

$sqlDoc = "SELECT o_docguid, o_projectno, o_itemname FROM dms_doc WHERE o_docguid IN ($inCast)"
$docRows = _RunPwSql -Sql $sqlDoc -Label 'dms_doc by docguid'
if ($docRows) {
    Write-Host "SQL dms_doc by o_docguid: $($docRows.Rows.Count)" -ForegroundColor Cyan
    foreach ($row in @($docRows.Rows | Select-Object -First 5)) {
        Write-Host ("  doc={0} projno={1} name={2}" -f $row.o_docguid, $row.o_projectno, $row.o_itemname)
    }
}

Write-Host ''
try {
    $folders = @(Get-PWFoldersByGUIDs -FolderGUIDs @($FolderGuids) -ErrorAction Stop) | Where-Object { $null -ne $_ }
    Write-Host "Get-PWFoldersByGUIDs (batch canonical): $($folders.Count) non-null object(s)" -ForegroundColor Cyan
    foreach ($f in $folders) {
        $fp = _FolderPathFromObject $f
        $fg = _Prop $f 'FolderGUID'
        if (-not $fg) { $fg = _Prop $f 'GUID' }
        if ($fp -is [hashtable]) {
            Write-Host "  guid=$fg type=$($fp.typeName) path=<none> props=$($fp.propertyNames -join ', ')" -ForegroundColor Yellow
        } else {
            Write-Host "  guid=$fg path=$fp"
        }
    }
} catch {
    Write-Host "Get-PWFoldersByGUIDs (batch) error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ''
Write-Host 'Per-GUID cmdlet + document fallback:' -ForegroundColor Cyan
foreach ($g in $FolderGuids) {
    $resolved = $false
    foreach ($v in @($g, ('{' + $g + '}'))) {
        try {
            $one = @(Get-PWFoldersByGUIDs -FolderGUIDs $v -ErrorAction Stop) | Where-Object { $null -ne $_ }
            if ($one.Count -gt 0) {
                $fp = _FolderPathFromObject $one[0]
                $pathStr = if ($fp -is [string]) { $fp } else { '<no path after GetFullPath>' }
                Write-Host "  folder $g variant '$v' -> $pathStr"
                $resolved = $true
                break
            }
        } catch { }
    }
    if (-not $resolved) {
        try {
            $doc = @(Get-PWDocumentsByGUIDs -DocumentGUIDs $g -ErrorAction Stop) | Where-Object { $null -ne $_ } | Select-Object -First 1
            if ($doc) {
                $fp = $null
                if ($doc.FolderPath) { $fp = [string]$doc.FolderPath }
                elseif ($doc.FullPath) { $fp = [System.IO.Path]::GetDirectoryName([string]$doc.FullPath) -replace '/', '\' }
                Write-Host "  document $g -> $(if ($fp) { $fp } else { '<no FolderPath>' })"
                $resolved = $true
            }
        } catch { }
    }
    if (-not $resolved) { Write-Host "  $g -> not found as folder or document" -ForegroundColor Yellow }
}

Write-Host ''
Write-Host 'SQL dms_proj + Get-PWFolders -FolderID:' -ForegroundColor Cyan
foreach ($g in $FolderGuids) {
    $cast = _SqlCastGuidList -Guids @($g)
    $sqlOne = "SELECT o_projguid, o_projectno, o_itemname FROM dms_proj WHERE o_projguid = $cast"
    $r = _RunPwSql -Sql $sqlOne -Label "proj $g"
    if (-not $r -or $r.Rows.Count -eq 0) {
        Write-Host "  $g -> no dms_proj row"
        continue
    }
    $row = $r.Rows[0]
    try {
        $projNo = [int]$row.o_projectno
        $f = Get-PWFolders -FolderID $projNo -JustOne -ErrorAction Stop
        $fp = _FolderPathFromObject $f
        $pathOut = if ($fp -is [string]) { $fp } else { '<no path on folder object>' }
        Write-Host "  $g -> projectno=$projNo path=$pathOut"
    } catch {
        Write-Host "  $g -> Get-PWFolders failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ''
Write-Host 'Done. Disconnecting.'
Disconnect-PW | Out-Null
