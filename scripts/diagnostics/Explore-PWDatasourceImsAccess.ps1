<#
.SYNOPSIS
Read-only probe of ProjectWise datasource access via Bentley IMS (and optional PWPS_Cloud).

.DESCRIPTION
Logs in with IMS (prefers -NonAdminLogin), inventories available modules/cmdlets,
probes generic root discovery vs exact known-folder metadata reads, and optionally
Cloud Get-* cmdlets. Distinguishes report-only metadata success from export denials.

Does not create, update, delete, checkout, or change workflow/attributes.
Export is attempted only when -ProbeExport is set (expected deny under report-only).

.PARAMETER DatasourceName
ProjectWise datasource, e.g. azdot-pw.bentley.com:azdot-pw-02

.PARAMETER PreferCloudModule
If PWPS_Cloud is installed, import it alone (never with PWPS_DAB in the same session).

.PARAMETER KnownFolderPaths
Exact folder paths known to be visible in ProjectWise Explorer. These are probed with
Get-PWFolders -JustOne and Get-PWDocumentsBySearch -JustThisFolder. Generic root
enumeration often fails under report-only even when these succeed.

.PARAMETER ProbeExport
If set, attempt Export-PWDocuments on the first document found in a known folder
(expected to fail without PowerShell Users / admin).

.PARAMETER OutDir
Directory for JSON summary (default: %TEMP%\pw-ims-probe\<timestamp>)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$DatasourceName = 'azdot-pw.bentley.com:azdot-pw-02',

    [switch]$PreferCloudModule,

    [Parameter(Mandatory = $false)]
    [string[]]$KnownFolderPaths = @(
        '02_HELP\Digital_Delivery_Docs\Modeling Guidelines'
    ),

    [switch]$ProbeExport,

    [Parameter(Mandatory = $false)]
    [string]$OutDir = ''
)

# pwps_dab / PWPS_Cloud require MTA; Cursor/VS Code terminals are usually STA.
$scriptPath = $PSCommandPath
if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
if (-not $scriptPath) { $scriptPath = Join-Path $PSScriptRoot 'Explore-PWDatasourceImsAccess.ps1' }
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        throw "MTA relaunch: could not resolve script path. Tried: $scriptPath"
    }
    $staMtaHelper = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'legacy\StaMtaRelaunch.ps1'
    if (-not (Test-Path -LiteralPath $staMtaHelper)) {
        throw "MTA relaunch helper not found: $staMtaHelper"
    }
    . $staMtaHelper
    $exeArgs = Build-PowerShellExeFileArgs -ScriptPath $scriptPath -BoundParameters $PSBoundParameters
    & powershell.exe @exeArgs
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Continue'
$Progressive = [ordered]@{}

function Write-Section([string]$Title) {
    Write-Host ''
    Write-Host ('=' * 72) -ForegroundColor Cyan
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ('=' * 72) -ForegroundColor Cyan
}

function Test-Cmd([string]$Name) {
    return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

function Invoke-Probe {
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][scriptblock]$Script,
        [int]$MaxItems = 20
    )
    $row = [ordered]@{ label = $Label; ok = $false; error = $null; summary = $null; sample = @() }
    try {
        $result = & $Script
        $row.ok = $true
        if ($null -eq $result) {
            $row.summary = 'null'
        }
        elseif ($result -is [System.Array] -or $result -is [System.Collections.ICollection]) {
            $arr = @($result)
            $row.summary = "count=$($arr.Count)"
            $row.sample = @(
                $arr | Select-Object -First $MaxItems | ForEach-Object {
                    if ($_ -is [string]) { $_ }
                    else {
                        $n = $null
                        foreach ($p in @('Name', 'FullName', 'FolderPath', 'Path', 'DisplayName', 'Email', 'UserName', 'ProjectName')) {
                            try {
                                if ($null -ne $_.PSObject.Properties[$p] -and $_.$p) { $n = [string]$_.$p; break }
                            } catch { }
                        }
                        if ($n) { $n } else { $_.ToString() }
                    }
                }
            )
        }
        elseif ($result -is [bool] -or $result -is [string] -or $result -is [ValueType]) {
            $row.summary = [string]$result
        }
        else {
            $row.summary = $result.GetType().FullName
            $props = @{}
            foreach ($p in @('Name', 'FullName', 'UserName', 'Email', 'WorkingFolder', 'DatasourceName', 'IsAdmin')) {
                try {
                    if ($null -ne $result.PSObject.Properties[$p]) { $props[$p] = [string]$result.$p }
                } catch { }
            }
            if ($props.Count -gt 0) { $row.sample = @($props) } else { $row.sample = @($result.ToString()) }
        }
        Write-Host ("[OK]   {0} :: {1}" -f $Label, $row.summary) -ForegroundColor Green
        if ($row.sample -and $row.sample.Count -gt 0) {
            $row.sample | Select-Object -First 10 | ForEach-Object { Write-Host ("       - {0}" -f $_) -ForegroundColor DarkGray }
        }
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match 'Refresh Token|Decoded Payload|OIDCAuthToken') {
            $msg = 'OIDC token type/binding error (details redacted)'
        }
        $row.error = $msg
        Write-Host ("[FAIL] {0} :: {1}" -f $Label, $row.error) -ForegroundColor Yellow
    }
    return [pscustomobject]$row
}

function Clear-PWPSCloudModule {
    Get-Module PWPS_Cloud, pwps_cloud -ErrorAction SilentlyContinue |
        Remove-Module -Force -ErrorAction SilentlyContinue
}

if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $OutDir = Join-Path $env:TEMP ("pw-ims-probe\{0:yyyyMMdd_HHmmss}" -f (Get-Date))
}
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

Write-Section 'Host'
$Progressive.host = [ordered]@{
    computerName = $env:COMPUTERNAME
    userName     = $env:USERNAME
    whoami       = (whoami)
    psVersion    = $PSVersionTable.PSVersion.ToString()
    bitness      = [Environment]::Is64BitProcess
    outDir       = $OutDir
    datasource   = $DatasourceName
}
$Progressive.host.GetEnumerator() | ForEach-Object { Write-Host ("{0}={1}" -f $_.Key, $_.Value) }

Write-Section 'Modules (ListAvailable)'
$modNames = @('pwps', 'PWPS', 'pwps_dab', 'PWPS_DAB', 'pwps_cloud', 'PWPS_Cloud')
$available = @(Get-Module -ListAvailable | Where-Object {
        $n = $_.Name
        ($modNames -contains $n) -or ($n -match '(?i)pwps|projectwise|bentley.*cloud')
    } | Sort-Object Name, Version -Unique)
$Progressive.modulesAvailable = @(
    $available | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Version = $_.Version.ToString(); Path = $_.Path }
    }
)
$Progressive.modulesAvailable | Format-Table -AutoSize | Out-String | Write-Host

$hasCloud = [bool]($available | Where-Object { $_.Name -match '(?i)^PWPS_Cloud$' })
$hasDab = [bool]($available | Where-Object { $_.Name -match '(?i)^PWPS_DAB$' })
$hasPwps = [bool]($available | Where-Object { $_.Name -match '(?i)^pwps$' })

Write-Section 'Import strategy'
$loginCmd = $null
$undoCmd = $null
$imported = @()

try { Undo-PWLogin -ErrorAction SilentlyContinue | Out-Null } catch { }
try { Close-PWConnection -ErrorAction SilentlyContinue | Out-Null } catch { }

# Start clean: overlapping DAB/Cloud OIDC types and cmdlet resolution cause false denials.
Get-Module PWPS_Cloud, pwps_cloud, PWPS_DAB, pwps_dab -ErrorAction SilentlyContinue |
    Remove-Module -Force -ErrorAction SilentlyContinue

function Invoke-ImsNonAdminLogin {
    param([string]$Ds, [string]$ModuleQualifier = '')
    # PWPS_DAB uses -DatasourceName; PWPS_Cloud docs often show -Datasource.
    $cmdName = if ($ModuleQualifier) { "${ModuleQualifier}\New-PWLogin" } else { 'New-PWLogin' }
    $cmd = Get-Command $cmdName -ErrorAction Stop
    $params = @{ BentleyIMS = $true; NonAdminLogin = $true; Verbose = $true }
    if ($cmd.Parameters.ContainsKey('DatasourceName')) { $params.DatasourceName = $Ds }
    elseif ($cmd.Parameters.ContainsKey('Datasource')) { $params.Datasource = $Ds }
    else { throw "$cmdName has neither -DatasourceName nor -Datasource" }
    & $cmd @params
}

if ($PreferCloudModule -and $hasCloud) {
    Import-Module PWPS_Cloud -Force -ErrorAction Stop
    $imported += 'PWPS_Cloud'
    $loginCmd = { param($ds) Invoke-ImsNonAdminLogin -Ds $ds -ModuleQualifier 'PWPS_Cloud' }
    $undoCmd = { PWPS_Cloud\Undo-PWLogin -ErrorAction SilentlyContinue }
    Write-Host 'Using PWPS_Cloud (PreferCloudModule)'
}
elseif ($hasCloud -and -not $hasDab) {
    Import-Module PWPS_Cloud -Force -ErrorAction Stop
    $imported += 'PWPS_Cloud'
    $loginCmd = { param($ds) Invoke-ImsNonAdminLogin -Ds $ds }
    $undoCmd = { Undo-PWLogin -ErrorAction SilentlyContinue }
    Write-Host 'Using PWPS_Cloud (only cloud module present)'
}
else {
    if ($hasDab) {
        Import-Module PWPS_DAB -Force -ErrorAction Stop
        $imported += 'PWPS_DAB'
    }
    elseif ($hasCloud) {
        Import-Module PWPS_Cloud -Force -ErrorAction Stop
        $imported += 'PWPS_Cloud'
    }
    else {
        throw 'Neither PWPS_DAB nor PWPS_Cloud is installed.'
    }
    $loginCmd = { param($ds) Invoke-ImsNonAdminLogin -Ds $ds }
    $undoCmd = { Undo-PWLogin -ErrorAction SilentlyContinue }
    Write-Host ("Using module(s): {0}" -f ($imported -join ', '))
}

# Do not keep PWPS_DAB and PWPS_Cloud loaded together: OIDC types conflict and
# Bentley recommends one module per session. DAB import can auto-pull Cloud.
if ($imported -contains 'PWPS_DAB') {
    Get-Module PWPS_Cloud, pwps_cloud -ErrorAction SilentlyContinue |
        Remove-Module -Force -ErrorAction SilentlyContinue
    if ($hasCloud) {
        Write-Host 'PWPS_Cloud is installed; re-run with -PreferCloudModule for a Cloud-only session (not co-loaded).'
    }
}
if ($imported -contains 'PWPS_Cloud') {
    Get-Module PWPS_DAB, pwps_dab -ErrorAction SilentlyContinue |
        Remove-Module -Force -ErrorAction SilentlyContinue
}

$Progressive.modulesImported = $imported
$Progressive.loginCommandOwners = @(
    Get-Command New-PWLogin -All -ErrorAction SilentlyContinue |
        Select-Object Name, ModuleName, Version, Source
)

Write-Section 'IMS login (NonAdminLogin)'
$login = [ordered]@{ nonAdmin = $null; full = $null }
try {
    $ok = & $loginCmd $DatasourceName
    $login.nonAdmin = [ordered]@{ ok = [bool]$ok; raw = [string]$ok }
    Write-Host ("NonAdminLogin => {0}" -f $ok) -ForegroundColor $(if ($ok) { 'Green' } else { 'Yellow' })
}
catch {
    $login.nonAdmin = [ordered]@{ ok = $false; error = $_.Exception.Message }
    Write-Host ("NonAdminLogin exception: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
}

if (-not $login.nonAdmin.ok) {
    Write-Host 'Retrying full IMS login (no -NonAdminLogin)...' -ForegroundColor DarkYellow
    try { & $undoCmd | Out-Null } catch { }
    try {
        $cmd = Get-Command New-PWLogin -ErrorAction Stop
        $fp = @{ BentleyIMS = $true; Verbose = $true }
        if ($cmd.Parameters.ContainsKey('DatasourceName')) { $fp.DatasourceName = $DatasourceName }
        elseif ($cmd.Parameters.ContainsKey('Datasource')) { $fp.Datasource = $DatasourceName }
        $okFull = & $cmd @fp
        $login.full = [ordered]@{ ok = [bool]$okFull; raw = [string]$okFull }
        Write-Host ("Full IMS login => {0}" -f $okFull) -ForegroundColor $(if ($okFull) { 'Green' } else { 'Yellow' })
    }
    catch {
        $login.full = [ordered]@{ ok = $false; error = $_.Exception.Message }
        Write-Host ("Full IMS login exception: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}
$Progressive.login = $login

$loggedIn = ($login.nonAdmin.ok -eq $true) -or ($login.full.ok -eq $true)
if (-not $loggedIn) {
    Write-Host 'No session - skipping live probes.' -ForegroundColor Red
    $Progressive | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutDir 'summary.json') -Encoding UTF8
    Write-Host ("Wrote {0}" -f (Join-Path $OutDir 'summary.json'))
    return
}

Write-Section 'Session identity'
$sessionProbes = @()
$sessionProbes += Invoke-Probe 'Get-PWCurrentUser' { Get-PWCurrentUser }
$sessionProbes += Invoke-Probe 'Get-PWCurrentDatasource' { Get-PWCurrentDatasource }
$sessionProbes += Invoke-Probe 'Get-PWCurrentDSSession' { Get-PWCurrentDSSession }
$sessionProbes += Invoke-Probe 'Get-PWDatasource -Active' { Get-PWDatasource -Active }
$sessionProbes += Invoke-Probe 'Get-PWUserWorkingDirectory' { Get-PWUserWorkingDirectory }
$sessionProbes += Invoke-Probe 'Get-PWServerVersion' { Get-PWServerVersion }
$Progressive.session = $sessionProbes

Write-Section 'Cmdlet inventory (post-login)'
# IMPORTANT: Get-Command *-Cloud* auto-imports PWPS_Cloud and can break DAB
# report-only folder/document metadata reads. Only inventory Cloud when requested.
$usingCloudSession = ($imported -contains 'PWPS_Cloud')
$cloudCmds = @()
$pwpsCloudCmds = @()
if ($usingCloudSession) {
    $cloudCmds = @(Get-Command *-Cloud* -ErrorAction SilentlyContinue | Select-Object Name, ModuleName, CommandType)
    $pwpsCloudCmds = @(Get-Command -Module PWPS_Cloud -ErrorAction SilentlyContinue | Select-Object Name, CommandType)
    Write-Host ("*-Cloud* cmdlets: {0}" -f $cloudCmds.Count)
    $cloudCmds | Select-Object -First 30 | ForEach-Object { Write-Host ("  {0}\{1}" -f $_.ModuleName, $_.Name) }
    Write-Host ("PWPS_Cloud exports: {0}" -f $pwpsCloudCmds.Count)
}
else {
    Clear-PWPSCloudModule
    Write-Host 'Skipping Cloud cmdlet auto-discovery (DAB-only session). Use -PreferCloudModule separately.'
}
$Progressive.cmdlets = [ordered]@{
    cloudStar      = @($cloudCmds | ForEach-Object { "$($_.ModuleName)\$($_.Name)" })
    pwpsCloudCount = $pwpsCloudCmds.Count
    pwpsCloudSample = @($pwpsCloudCmds | Select-Object -First 40 -ExpandProperty Name)
    hasGetCloudProject = if ($usingCloudSession) { (Test-Cmd 'Get-CloudProject') } else { $false }
    hasGetWorkAreas = (Test-Cmd 'Get-PWWorkAreas') -or (Test-Cmd 'Get-PWRichProjects')
    dabOnlySession = (-not $usingCloudSession)
}

Write-Section 'Generic root discovery (often blocked under report-only)'
$readProbes = @()
$readProbes += Invoke-Probe 'Get-PWFolders (Documents)' { Get-PWFolders -FolderPath 'Documents' -ErrorAction Stop }
$readProbes += Invoke-Probe 'Get-PWFoldersImmediateChildren Documents' {
    if (Test-Cmd 'Get-PWFoldersImmediateChildren') {
        Get-PWFoldersImmediateChildren -FolderPath 'Documents' -ErrorAction Stop
    }
    else { throw 'cmdlet missing' }
}
$readProbes += Invoke-Probe 'Get-PWFolderView Documents' {
    if (Test-Cmd 'Get-PWFolderView') { Get-PWFolderView -FolderPath 'Documents' -ErrorAction Stop }
    else { throw 'cmdlet missing' }
}

foreach ($root in @('Documents\AZDOT', 'AZDOT', '02_HELP')) {
    $readProbes += Invoke-Probe ("Get-PWFolders '{0}' (no -JustOne)" -f $root) {
        Get-PWFolders -FolderPath $root -ErrorAction Stop
    }
}

if (Test-Cmd 'Get-PWWorkAreas') {
    $readProbes += Invoke-Probe 'Get-PWWorkAreas' { Get-PWWorkAreas -ErrorAction Stop }
}
elseif (Test-Cmd 'Get-PWRichProjects') {
    $readProbes += Invoke-Probe 'Get-PWRichProjects' { Get-PWRichProjects -ErrorAction Stop }
}
$Progressive.genericReads = $readProbes

Write-Section 'Known-folder metadata (report-only path that can succeed)'
# Ensure Cloud is not loaded before exact-path metadata probes.
if (-not $usingCloudSession) { Clear-PWPSCloudModule }
$knownProbes = @()
$firstKnownDoc = $null
foreach ($known in @($KnownFolderPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
    $knownProbes += Invoke-Probe ("Get-PWFolders -JustOne '{0}'" -f $known) {
        Get-PWFolders -FolderPath $known -JustOne -ErrorAction Stop
    }
    if (Test-Cmd 'Get-PWDocumentsBySearch') {
        $docProbe = Invoke-Probe ("Get-PWDocumentsBySearch JustThisFolder '{0}'" -f $known) {
            Get-PWDocumentsBySearch -FolderPath $known -JustThisFolder -PopulatePath -ErrorAction Stop |
                Select-Object -First 25
        }
        $knownProbes += $docProbe
        if ($docProbe.ok -and -not $firstKnownDoc) {
            try {
                $firstKnownDoc = @(
                    Get-PWDocumentsBySearch -FolderPath $known -JustThisFolder -PopulatePath -ErrorAction Stop
                ) | Select-Object -First 1
            }
            catch { }
        }
    }
}
$Progressive.knownFolderReads = $knownProbes

if ($ProbeExport) {
    Write-Section 'Export probe (expected deny without PowerShell Users)'
    $exportProbes = @()
    if ($null -eq $firstKnownDoc) {
        $exportProbes += [pscustomobject]@{
            label = 'Export-PWDocuments'; ok = $false; error = 'no known-folder document to export'; summary = $null; sample = @()
        }
        Write-Host '[SKIP] Export-PWDocuments :: no known-folder document' -ForegroundColor DarkGray
    }
    else {
        $exportDir = Join-Path $OutDir 'export-probe'
        New-Item -ItemType Directory -Force -Path $exportDir | Out-Null
        $exportProbes += Invoke-Probe 'Export-PWDocuments (known folder doc)' {
            Export-PWDocuments -InputDocuments @($firstKnownDoc) -OutputFolder $exportDir -ErrorAction Stop
        }
    }
    $Progressive.export = $exportProbes
}

Write-Section 'Cloud / PWPS_Cloud Get-* probes (read-only)'
$cloudProbes = @()
$usingCloud = $usingCloudSession
if (-not $usingCloud) {
    Write-Host 'Skipped (DAB-only session). Re-run with -PreferCloudModule for Cloud probes.' -ForegroundColor DarkGray
    $Progressive.cloud = @([pscustomobject]@{ label = 'CloudProbes'; ok = $false; error = 'skipped_dab_only_session'; summary = $null; sample = @() })
}
else {

# Never capture/print Get-OIDCToken return value (ToString can dump refresh token).
$token = $null
$tokenOk = $false
if (Test-Cmd 'Get-OIDCToken') {
    $oidcRow = [ordered]@{ label = 'Get-OIDCToken'; ok = $false; error = $null; summary = $null; sample = @() }
    try {
        $token = Get-OIDCToken -ErrorAction Stop
        $tokenOk = $true
        $oidcRow.ok = $true
        $oidcRow.summary = ('token_acquired type={0} (value redacted)' -f $token.GetType().FullName)
        Write-Host ('[OK]   Get-OIDCToken :: {0}' -f $oidcRow.summary) -ForegroundColor Green
    }
    catch {
        $oidcRow.error = $_.Exception.Message
        if ($oidcRow.error -match 'Refresh Token|Decoded Payload|OIDCAuthToken') {
            $oidcRow.error = 'OIDC error (details redacted)'
        }
        Write-Host ('[FAIL] Get-OIDCToken :: {0}' -f $oidcRow.error) -ForegroundColor Yellow
    }
    $cloudProbes += [pscustomobject]$oidcRow
}
else {
    $cloudProbes += [pscustomobject]@{ label = 'Get-OIDCToken'; ok = $false; error = 'cmdlet missing'; summary = $null; sample = @() }
    Write-Host '[MISS] Get-OIDCToken' -ForegroundColor DarkGray
}

$cloudGets = @(
    'Get-CloudProject',
    'Get-CloudTemplates',
    'Get-CloudProjectWorkAreaConnections',
    'Get-CloudProjectRoles',
    'Get-CloudProjectId',
    'Get-DataCenterLocations',
    'Get-BillingCountries'
)
foreach ($name in $cloudGets) {
    if (-not (Test-Cmd $name)) {
        $cloudProbes += [pscustomobject]@{ label = $name; ok = $false; error = 'cmdlet missing'; summary = $null; sample = @() }
        Write-Host ("[MISS] {0}" -f $name) -ForegroundColor DarkGray
        continue
    }
    if (-not $tokenOk) {
        # Still try no-arg form for cmdlets that might not need a token in-session.
        $cloudProbes += Invoke-Probe $name { & $name -ErrorAction Stop }
        continue
    }
    $cmd = Get-Command $name -ErrorAction Stop
    if ($cmd.Parameters.ContainsKey('OIDCToken')) {
        $cloudProbes += Invoke-Probe $name { & $name -OIDCToken $token -ErrorAction Stop }
    }
    elseif ($cmd.Parameters.ContainsKey('Token') -and $cmd.Parameters.ContainsKey('CloudProjectId')) {
        $cloudProbes += [pscustomobject]@{
            label = $name; ok = $false; error = 'needs CloudProjectId (skipped)'; summary = $null; sample = @()
        }
        Write-Host ("[SKIP] {0} :: needs CloudProjectId" -f $name) -ForegroundColor DarkGray
    }
    else {
        $cloudProbes += Invoke-Probe $name { & $name -ErrorAction Stop }
    }
}
$token = $null
$Progressive.cloud = $cloudProbes
} # end PreferCloudModule cloud probes

Write-Section 'Capability matrix'
if (-not $usingCloudSession) { Clear-PWPSCloudModule }
$matrixNames = @(
    'Get-PWFolders',
    'Get-PWFoldersImmediateChildren',
    'Get-PWDocumentsBySearch',
    'Get-PWDocumentEAttributes',
    'Export-PWDocumentsSimple',
    'CopyOut-PWDocuments',
    'New-PWDocument',
    'Set-PWDocumentState',
    'Update-PWDocumentFile',
    'Get-PWUser',
    'Get-PWGroup',
    'Invoke-PWCelExpression',
    'Get-PWWorkAreas',
    'Get-PWEnvironments'
)
if ($usingCloudSession) { $matrixNames += 'Get-CloudProject' }
$matrix = @()
foreach ($n in $matrixNames) {
    $exists = Test-Cmd $n
    $entry = [ordered]@{ name = $n; exists = $exists; invoke = 'skipped' }
    if ($exists -and $n -match '^(Get-|Invoke-PWCelExpression)') {
        # Only invoke a few Get-* that take no/minimal args; CEL with a trivial expression
        try {
            if ($n -eq 'Invoke-PWCelExpression') {
                $null = Invoke-PWCelExpression -Expression '1' -ErrorAction Stop
            }
            elseif ($n -eq 'Get-PWDocumentEAttributes') {
                $entry.invoke = 'exists-only (needs document)'
                $matrix += [pscustomobject]$entry
                Write-Host ("[EXISTS] {0} (not invoked - needs document)" -f $n)
                continue
            }
            elseif ($n -eq 'Get-PWFolders') {
                $null = Get-PWFolders -FolderPath 'Documents' -ErrorAction Stop
            }
            elseif ($n -eq 'Get-PWFoldersImmediateChildren') {
                $null = Get-PWFoldersImmediateChildren -FolderPath 'Documents' -ErrorAction Stop
            }
            elseif ($n -eq 'Get-PWDocumentsBySearch') {
                $entry.invoke = 'exists-only (needs folder)'
                $matrix += [pscustomobject]$entry
                Write-Host ("[EXISTS] {0} (not invoked - needs folder)" -f $n)
                continue
            }
            else {
                $null = & $n -ErrorAction Stop
            }
            $entry.invoke = 'ok'
            Write-Host ("[OK]   {0}" -f $n) -ForegroundColor Green
        }
        catch {
            $entry.invoke = 'fail'
            $entry.error = $_.Exception.Message
            Write-Host ("[FAIL] {0} :: {1}" -f $n, $_.Exception.Message) -ForegroundColor Yellow
        }
    }
    elseif ($exists) {
        $entry.invoke = 'exists-only (mutating/export - not run)'
        Write-Host ("[EXISTS] {0} (not invoked)" -f $n) -ForegroundColor DarkCyan
    }
    else {
        $entry.invoke = 'missing'
        Write-Host ("[MISS] {0}" -f $n) -ForegroundColor DarkGray
    }
    $matrix += [pscustomobject]$entry
}
$Progressive.matrix = $matrix

Write-Section 'Cleanup'
try { & $undoCmd | Out-Null } catch { }
try { Undo-PWLogin -ErrorAction SilentlyContinue | Out-Null } catch { }
Write-Host 'Logged out.'

$summaryPath = Join-Path $OutDir 'summary.json'
$Progressive | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
Write-Section 'Done'
Write-Host ("JSON summary: {0}" -f $summaryPath) -ForegroundColor Green
Write-Host 'Paste the console output (or summary.json) back into chat for analysis.'
