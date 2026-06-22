<#
.SYNOPSIS
    Smoke test for New-PWRenditionRequest using repo appsettings and a target folder.

.DESCRIPTION
    By default this is read-only discovery:
      - connects via appsettings.json projectWise settings
      - lists rendition profiles
      - resolves the folder's assigned profile (if any)
      - finds candidate source documents in the folder
      - prints the exact New-PWRenditionRequest command that would run

    Pass -Submit to actually queue a rendition request.
    Use -WhatIf with -Submit to preview without calling ProjectWise.

.PARAMETER FolderPath
    PW folder path without the Documents\ prefix (pwps_dab style).
    Default: JFlint Prepend Test folder used for rendition discovery probes.

.PARAMETER DocumentName
    Exact document name. When set, -FileName is ignored.

.PARAMETER FileName
    Search pattern for Get-PWDocumentsBySearch (default %.dgn).

.PARAMETER ProfileName
    Rendition profile to use. When omitted, uses Get-PWFolderRenditionProfile
    on the target folder, then prompts if still unset.

.PARAMETER Submit
    Actually call New-PWRenditionRequest. Without this switch the script only discovers.

.PARAMETER MakeSinglePlotRequest
    Pass -MakeSinglePlotRequest to New-PWRenditionRequest (one plot job for all docs).

.PARAMETER MaxDocuments
    Limit how many documents are included in one submission (default 1 for safety).

.EXAMPLE
    .\tools\discovery\Test-PWRenditionRequest.ps1
    Discovery only: list profiles, find docs, show planned command.

.EXAMPLE
    .\tools\discovery\Test-PWRenditionRequest.ps1 -FolderPath "AZDOT 2024\MyProject\CADD\Sheets" -ProfileName "PDF to Renditions Folder" -Submit -WhatIf

.EXAMPLE
    .\tools\discovery\Test-PWRenditionRequest.ps1 -DocumentName "sheet001.dgn" -ProfileName "For Construction" -Submit -Verbose
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath = '',

    [Parameter(Mandatory = $false)]
    [string]$FolderPath = '',

    [Parameter(Mandatory = $false)]
    [switch]$ListProfilesOnly,

    [Parameter(Mandatory = $false)]
    [string]$DocumentName = '',

    [Parameter(Mandatory = $false)]
    [string]$FileName = '%.dgn',

    [Parameter(Mandatory = $false)]
    [string]$ProfileName = '',

    [Parameter(Mandatory = $false)]
    [switch]$Submit,

    [Parameter(Mandatory = $false)]
    [switch]$MakeSinglePlotRequest,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$MaxDocuments = 1,

    # Get-PWRenditionJobsStatus queries the PW SQL Server directly (not via Connect-PW).
    [Parameter(Mandatory = $false)]
    [string]$RenditionDatabaseServer = '',

    [Parameter(Mandatory = $false)]
    [string]$RenditionDatabaseName = '',

    [Parameter(Mandatory = $false)]
    [string]$RenditionDatabaseUser = '',

    [Parameter(Mandatory = $false)]
    [switch]$RenditionUseWindowsAuth
)

$ErrorActionPreference = 'Stop'

function Write-Step([string]$Title) {
    Write-Host ''
    Write-Host "=== $Title ===" -ForegroundColor Cyan
}

function Write-Info([string]$Message) {
    Write-Host $Message -ForegroundColor Gray
}

function Write-Ok([string]$Message) {
    Write-Host $Message -ForegroundColor Green
}

function Write-Warn([string]$Message) {
    Write-Host $Message -ForegroundColor Yellow
}

function Get-RenditionProfileNames {
    $profiles = @(Get-PWRenditionProfiles -ErrorAction Stop)
    $names = @()
    foreach ($p in $profiles) {
        if ($p -is [string]) { $names += $p; continue }
        foreach ($prop in @('Name', 'ProfileName', 'RenditionProfileName')) {
            if ($p.PSObject.Properties[$prop] -and -not [string]::IsNullOrWhiteSpace([string]$p.$prop)) {
                $names += [string]$p.$prop
                break
            }
        }
    }
    return @($names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Get-FolderRenditionProfileName([string]$Path) {
    $cmd = Get-Command -Name 'Get-PWFolderRenditionProfile' -ErrorAction SilentlyContinue
    if (-not $cmd) { return $null }
    try {
        $result = Get-PWFolderRenditionProfile -FolderPath $Path -ErrorAction Stop
    } catch {
        Write-Warn "Get-PWFolderRenditionProfile failed: $($_.Exception.Message)"
        return $null
    }
    if (-not $result) { return $null }
    if ($result -is [string]) { return $result.Trim() }
    foreach ($prop in @('ProfileName', 'Name', 'RenditionProfileName')) {
        if ($result.PSObject.Properties[$prop] -and -not [string]::IsNullOrWhiteSpace([string]$result.$prop)) {
            return [string]$result.$prop
        }
    }
    return [string]$result
}

function Resolve-DefaultRenditionFolderPath([hashtable]$ProjectWiseConfig) {
    if (-not $ProjectWiseConfig -or -not $ProjectWiseConfig.watchList) { return $null }
    $watch = ConvertTo-HashtableDeep -Value $ProjectWiseConfig.watchList
    if (-not $watch -or -not $watch.roots) { return $null }

    foreach ($rootEntry in @($watch.roots)) {
        $root = ConvertTo-HashtableDeep -Value $rootEntry
        if (-not $root) { continue }
        $rootPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath ([string]$root.path)
        $sheetsSuffix = if ($root.sheetsPathFromProject) { [string]$root.sheetsPathFromProject } else { 'CADD\Sheets' }
        if ([string]::IsNullOrWhiteSpace($rootPath)) { continue }

        $children = @(Get-PWFoldersImmediateChildren -FolderPath $rootPath -ErrorAction SilentlyContinue)
        foreach ($child in $children) {
            $childName = $null
            if ($child -is [string]) { $childName = $child }
            elseif ($child.Name) { $childName = [string]$child.Name }
            if ([string]::IsNullOrWhiteSpace($childName)) { continue }
            if ($childName.StartsWith('_')) { continue }

            $candidate = "$rootPath\$childName\$sheetsSuffix".TrimEnd('\')
            $folder = Get-PWFolders -FolderPath $candidate -JustOne -ErrorAction SilentlyContinue
            if ($folder) { return $candidate }
        }
    }
    return $null
}

function Show-RenditionOptionCatalog {
    $helpers = @(
        @{ Cmd = 'Get-PWRenditionPresentationOptions'; Label = 'Presentation options' },
        @{ Cmd = 'Get-PWRenditionOptions'; Label = 'Rendition options' },
        @{ Cmd = 'Get-PWRenditionFileNamingOptions'; Label = 'Filename options' },
        @{ Cmd = 'Get-PWRenditionDestinations'; Label = 'Destination options' }
    )
    foreach ($h in $helpers) {
        $cmd = Get-Command -Name $h.Cmd -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        try {
            $items = @(& $h.Cmd -ErrorAction Stop)
            if ($items.Count -eq 0) { continue }
            Write-Info "$($h.Label):"
            foreach ($item in $items) {
                if ($item -is [string]) { Write-Info "  - $item"; continue }
                $label = $null
                foreach ($prop in @('Name', 'OptionName', 'DisplayName')) {
                    if ($item.PSObject.Properties[$prop] -and -not [string]::IsNullOrWhiteSpace([string]$item.$prop)) {
                        $label = [string]$item.$prop
                        break
                    }
                }
                if ($label) { Write-Info "  - $label" }
            }
        } catch {
            Write-Warn "$($h.Cmd) failed: $($_.Exception.Message)"
        }
    }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.Connection.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.Discovery.psm1') -Force

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

$pw = @{}
if ($config.ContainsKey('projectWise') -and $config.projectWise) {
    $pw = ConvertTo-HashtableDeep -Value $config.projectWise
}
$script:renditionProbeProjectWise = $pw

$ds = if ($pw.datasourceName) { [string]$pw.datasourceName } else { 'typsa-us-pw.bentley.com:typsa-us-pw-03' }
$credPath = if ($pw.credentialPath) { [string]$pw.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }

Write-Step 'Configuration'
Write-Info "Repo root:       $repoRoot"
Write-Info "App settings:    $AppSettingsPath"
Write-Info "Datasource:      $ds"
Write-Info "Credential path: $credPath"
Write-Info "Submit:          $Submit"
Write-Info "WhatIf:          $($WhatIfPreference.IsPresent)"

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }

$connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $connRes.IsSuccess) { throw ($connRes.Code + ': ' + $connRes.Message) }
Write-Ok "Connected as $($connRes.Data.userName)"

try {
    Write-Step 'Rendition profiles (Get-PWRenditionProfiles)'
    $profileNames = @(Get-RenditionProfileNames)
    if ($profileNames.Count -eq 0) {
        Write-Warn 'No rendition profiles returned.'
    } else {
        foreach ($name in $profileNames) { Write-Info "  - $name" }
    }

    if ($ListProfilesOnly) {
        Show-RenditionOptionCatalog
        Write-Warn 'ListProfilesOnly set; skipping folder/document discovery.'
        return
    }

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        $FolderPath = Resolve-DefaultRenditionFolderPath -ProjectWiseConfig $pw
        if ([string]::IsNullOrWhiteSpace($FolderPath)) {
            throw 'FolderPath not specified and no watchList sheets folder could be resolved. Pass -FolderPath explicitly.'
        }
        Write-Warn "Using watchList-derived folder: $FolderPath"
    }

    $apiFolderPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
    if ([string]::IsNullOrWhiteSpace($apiFolderPath)) { $apiFolderPath = $FolderPath.Trim().TrimEnd('\') }
    Write-Info "Folder path:     $apiFolderPath"

    Write-Step 'Folder rendition profile (Get-PWFolderRenditionProfile)'
    $folderProfile = Get-FolderRenditionProfileName -Path $apiFolderPath
    if ($folderProfile) {
        Write-Ok "Folder profile: $folderProfile"
    } else {
        Write-Warn 'No folder-level rendition profile found (or cmdlet unavailable).'
    }

    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        $ProfileName = $folderProfile
    }
    if ([string]::IsNullOrWhiteSpace($ProfileName) -and ($profileNames -contains 'Quality Control - PDF ANSI D')) {
        $ProfileName = 'Quality Control - PDF ANSI D'
        Write-Warn "Using QC default profile: $ProfileName"
    }
    if ([string]::IsNullOrWhiteSpace($ProfileName) -and $profileNames.Count -gt 0) {
        $ProfileName = [string]$profileNames[0]
        Write-Warn "Using first available profile: $ProfileName"
    }
    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        throw 'ProfileName is required. Pass -ProfileName or assign a profile to the folder in ProjectWise.'
    }
    if ($profileNames.Count -gt 0 -and ($ProfileName -notin $profileNames)) {
        Write-Warn "Profile '$ProfileName' was not in Get-PWRenditionProfiles output; submission may still work if the name is valid."
    }

    Show-RenditionOptionCatalog

    Write-Step 'Candidate documents'
    if (-not [string]::IsNullOrWhiteSpace($DocumentName)) {
        $docs = @(Get-PWDocumentsBySearch -FolderPath $apiFolderPath -JustThisFolder -DocumentName $DocumentName -PopulatePath -ErrorAction Stop)
        Write-Info "Search: exact name '$DocumentName'"
    } else {
        $docs = @(Get-PWDocumentsBySearch -FolderPath $apiFolderPath -JustThisFolder -DocumentName $FileName -PopulatePath -ErrorAction Stop)
        Write-Info "Search: pattern '$FileName'"
    }

    if ($docs.Count -eq 0) {
        Write-Warn "No documents found in '$apiFolderPath' matching '$(
            if ([string]::IsNullOrWhiteSpace($DocumentName)) { $FileName } else { $DocumentName }
        )'."
        Write-Warn 'Pass -FolderPath and -DocumentName for a known source file, or try -FileName %.dwg / %.pdf.'
        return
    }

    $docs = @($docs | Select-Object -First $MaxDocuments)
    Write-Ok "Selected $($docs.Count) document(s):"
    foreach ($d in $docs) {
        $path = if ($d.FullPath) { [string]$d.FullPath } else { [string]$d.Name }
        Write-Info "  - $path"
    }

    $splat = @{
        InputDocuments = $docs
        ProfileName    = $ProfileName
    }
    if ($MakeSinglePlotRequest) { $splat.MakeSinglePlotRequest = $true }

    Write-Step 'Planned submission'
    $cmdParts = @('New-PWRenditionRequest')
    $cmdParts += "-InputDocuments @(`$docs)"
    $cmdParts += "-ProfileName '$ProfileName'"
    if ($MakeSinglePlotRequest) { $cmdParts += '-MakeSinglePlotRequest' }
    Write-Info ($cmdParts -join ' ')

    if (-not $Submit) {
        Write-Warn 'Discovery complete. Re-run with -Submit to queue the rendition request.'
        return
    }

    $target = "$ProfileName -> $($docs.Count) doc(s) in $apiFolderPath"
    if ($PSCmdlet.ShouldProcess($target, 'Submit ProjectWise rendition request')) {
        Write-Step 'Submitting (New-PWRenditionRequest)'
        New-PWRenditionRequest @splat -Verbose
        Write-Ok 'Rendition request submitted.'

        $statusCmd = Get-Command -Name 'Get-PWRenditionJobsStatus' -ErrorAction SilentlyContinue
        if ($statusCmd) {
            Write-Step 'Recent rendition jobs (Get-PWRenditionJobsStatus)'
            $sqlServer = $RenditionDatabaseServer
            $sqlDb = $RenditionDatabaseName
            if ([string]::IsNullOrWhiteSpace($sqlServer) -or [string]::IsNullOrWhiteSpace($sqlDb)) {
                $pwCfg = $script:renditionProbeProjectWise
                if ($pwCfg -and $pwCfg.renditionJobsSql) {
                    $rj = ConvertTo-HashtableDeep -Value $pwCfg.renditionJobsSql
                    if ($rj) {
                        if ([string]::IsNullOrWhiteSpace($sqlServer) -and $rj.databaseServer) { $sqlServer = [string]$rj.databaseServer }
                        if ([string]::IsNullOrWhiteSpace($sqlDb) -and $rj.databaseName) { $sqlDb = [string]$rj.databaseName }
                    }
                }
            }
            if ([string]::IsNullOrWhiteSpace($sqlServer) -or [string]::IsNullOrWhiteSpace($sqlDb)) {
                Write-Warn 'Skipped: Get-PWRenditionJobsStatus requires -RenditionDatabaseServer and -RenditionDatabaseName (PW SQL backend, not the Connect-PW datasource).'
                Write-Info '  Example: -RenditionDatabaseServer "sqlhost" -RenditionDatabaseName "pwmdb" -RenditionUseWindowsAuth'
            } else {
                try {
                    $statusArgs = @{
                        DatabaseServer = $sqlServer
                        DatabaseName   = $sqlDb
                    }
                    if ($RenditionUseWindowsAuth.IsPresent) { $statusArgs['UseWindowsAuth'] = $true }
                    elseif (-not [string]::IsNullOrWhiteSpace($RenditionDatabaseUser)) {
                        $statusArgs['DatabaseUser'] = $RenditionDatabaseUser
                    }
                    $jobs = @(Get-PWRenditionJobsStatus @statusArgs -ErrorAction Stop)
                    $jobs | Select-Object -First 10 | ForEach-Object {
                        if ($_ -is [string]) { Write-Info "  $_"; return }
                        $line = ($_.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; '
                        Write-Info "  $line"
                    }
                } catch {
                    Write-Warn "Get-PWRenditionJobsStatus failed: $($_.Exception.Message)"
                }
            }
        }
    }
}
finally {
    $disc = Disconnect-PW
    if ($disc.IsSuccess) { Write-Ok 'Disconnected from ProjectWise.' }
}
