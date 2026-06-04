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

function _Prop([object]$o, [string]$n) {
    if ($null -eq $o) { return $null }
    try {
        if ($o.PSObject.Properties[$n]) { return $o.$n }
    } catch { }
    return $null
}

function _FolderPathFromObject([object]$folder) {
    if ($null -eq $folder) { return $null }
    if ($folder -is [string]) { return $folder }
    foreach ($p in @('FolderPath', 'Path', 'FullPath', 'folderPath', 'Name')) {
        $v = _Prop $folder $p
        if ($v -and "$v" -match '\\') { return [string]$v }
    }
    $names = @()
    try { $names = @($folder.PSObject.Properties.Name) } catch { }
    return @{ path = $null; propertyNames = $names }
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
Write-Host "Get-PWFoldersHashTableByGuid params: FolderPath, FolderID (NOT -FolderGUIDs)" -ForegroundColor Yellow
Write-Host ''

$results = [System.Collections.Generic.List[object]]::new()

# --- SQL: dms_proj by o_projguid ---
$guidList = ($FolderGuids | ForEach-Object { "'$($_ -replace '''', '''''')'" }) -join ','
$sqlProj = @"
SELECT o_projguid, o_projectno, o_itemname, o_itemdesc
FROM dms_proj
WHERE o_projguid IN ($guidList)
"@
try {
    $sqlRows = Select-PWSQL -SQLSelectStatement $sqlProj -ErrorAction Stop
    Write-Host "SQL dms_proj hits: $($sqlRows.Rows.Count) / $($FolderGuids.Count)" -ForegroundColor Cyan
    foreach ($row in $sqlRows.Rows) {
        Write-Host ("  proj {0} no={1} name={2}" -f $row.o_projguid, $row.o_projectno, $row.o_itemname)
    }
} catch {
    Write-Host "SQL dms_proj failed: $($_.Exception.Message)" -ForegroundColor Red
}

# --- SQL: dms_doc parent (in case GUID is document not folder) ---
$sqlDoc = @"
SELECT o_docguid, o_parentguid, o_itemname, o_projectno
FROM dms_doc
WHERE o_docguid IN ($guidList) OR o_parentguid IN ($guidList)
"@
try {
    $docRows = Select-PWSQL -SQLSelectStatement $sqlDoc -ErrorAction Stop
    Write-Host "SQL dms_doc rows (guid as doc or parent): $($docRows.Rows.Count)" -ForegroundColor Cyan
    foreach ($row in @($docRows.Rows | Select-Object -First 10)) {
        Write-Host ("  doc={0} parent={1} name={2} projno={3}" -f $row.o_docguid, $row.o_parentguid, $row.o_itemname, $row.o_projectno)
    }
} catch {
    Write-Host "SQL dms_doc failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ''

# --- Get-PWFoldersByGUIDs: single array (blog / help example) ---
try {
    $folders = @(Get-PWFoldersByGUIDs -FolderGUIDs @($FolderGuids) -ErrorAction Stop)
    Write-Host "Get-PWFoldersByGUIDs (batch): returned $($folders.Count) object(s)" -ForegroundColor Cyan
    foreach ($f in $folders) {
        $fp = _FolderPathFromObject $f
        $fg = _Prop $f 'FolderGUID'
        if (-not $fg) { $fg = _Prop $f 'GUID' }
        if ($fp -is [hashtable]) {
            Write-Host "  guid=$fg path=<none> props=$($fp.propertyNames -join ', ')" -ForegroundColor Yellow
        } else {
            Write-Host "  guid=$fg path=$fp"
        }
        [void]$results.Add(@{ guid = $fg; path = if ($fp -is [string]) { $fp } else { $null }; source = 'Get-PWFoldersByGUIDs-batch' })
    }
} catch {
    Write-Host "Get-PWFoldersByGUIDs (batch) error: $($_.Exception.Message)" -ForegroundColor Red
}

# --- Per-GUID with brace variants (do not pipe strings to Select-Object; it splits into chars) ---
foreach ($g in $FolderGuids) {
    $variants = @($g, $g.ToUpperInvariant(), ('{' + $g + '}'), ('{' + $g.ToUpperInvariant() + '}'))
    $seen = @{}
    foreach ($v in $variants) {
        if ($seen.ContainsKey($v)) { continue }
        $seen[$v] = $true
        try {
            $one = @(Get-PWFoldersByGUIDs -FolderGUIDs $v -ErrorAction Stop)
            if ($one.Count -gt 0) {
                $fp = _FolderPathFromObject $one[0]
                $pathStr = if ($fp -is [string]) { $fp } else { '<no path>' }
                Write-Host "  OK variant '$v' -> $pathStr"
                break
            }
        } catch {
            Write-Host "  FAIL variant '$v': $($_.Exception.Message)"
        }
    }
}

# --- SQL + Get-PWFolders -FolderID (poller v3 fallback) ---
Write-Host ''
Write-Host 'SQL dms_proj + Get-PWFolders -FolderID:' -ForegroundColor Cyan
foreach ($g in $FolderGuids) {
    $gk = $g.Trim().Trim('{}').ToLowerInvariant()
    $sqlOne = "SELECT o_projguid, o_projectno, o_itemname FROM dms_proj WHERE LOWER(RTRIM(o_projguid)) = '$gk'"
    try {
        $r = Select-PWSQL -SQLSelectStatement $sqlOne -ErrorAction Stop
        if ($r.Rows.Count -eq 0) {
            Write-Host "  $g -> no dms_proj row"
            continue
        }
        $row = $r.Rows[0]
        $projNo = [int]$row.o_projectno
        $f = Get-PWFolders -FolderID $projNo -JustOne -ErrorAction Stop
        try { $null = $f.GetFullPath() } catch { }
        $fp = _FolderPathFromObject $f
        $pathOut = if ($fp -is [string]) { $fp } else { '<no path on folder object>' }
        Write-Host "  $g -> projectno=$projNo path=$pathOut"
    } catch {
        Write-Host "  $g -> $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ''
Write-Host 'Done. Disconnecting.'
Disconnect-PW | Out-Null

$results | ConvertTo-Json -Depth 4
