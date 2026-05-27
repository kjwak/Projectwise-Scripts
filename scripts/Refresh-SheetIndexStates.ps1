<#
.SYNOPSIS
Refreshes sheet_index.pw_state_name from ProjectWise workflow state for all indexed documents.

.DESCRIPTION
Reads GUIDs from dbo.sheet_index (document_guid, and optionally qc_pdf_guid), fetches the current
ProjectWise workflow state (WorkflowState/StateName) via Get-PWDocumentsByGUIDs, then updates
sheet_index.pw_state_name for rows whose document_guid matches.

This script is intentionally focused on database state refresh only. It does NOT change ProjectWise
state; it only reads PW and writes the database.

Requires:
- database.enabled = true (appsettings.json)
- pwps_dab available + ProjectWise connectivity (credentialPath/datasourceName)

.EXAMPLE
PS> .\scripts\Refresh-SheetIndexStates.ps1 -DryRun -FolderPathFilter 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'
PS> .\scripts\Refresh-SheetIndexStates.ps1 -ConfirmWrites -FolderPathFilter 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'

#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath = '',
    [switch]$ConfirmWrites,
    [switch]$DryRun,
    [string]$FolderPathFilter = '',
    [switch]$IncludeQcPdfGuids,
    [int]$Limit = 0,
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force

foreach ($moduleName in @('pwps', 'pwps_dab')) {
    if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending | Select-Object -First 1
        if ($available) { Import-Module $available.Name -ErrorAction SilentlyContinue | Out-Null }
    }
}

if ($DryRun.IsPresent -and $ConfirmWrites.IsPresent) {
    throw 'Use -DryRun (preview only) OR -ConfirmWrites (apply DB changes), not both.'
}

$doWrites = $ConfirmWrites.IsPresent
if (-not $DryRun.IsPresent -and -not $ConfirmWrites.IsPresent) {
    Write-Host 'Refusing DB writes: pass -ConfirmWrites to update sheet_index, or -DryRun to preview only.' -ForegroundColor Yellow
    $DryRun = $true
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

if (-not (Test-QCDatabaseEnabled -Config $config)) {
    throw 'database.enabled must be true.'
}

Initialize-QCDatabaseSchema -Config $config | Out-Null

$sql = @"
SELECT
  document_guid,
  document_name,
  folder_path,
  extension,
  source_type,
  pw_state_name,
  qc_pdf_guid,
  qc_pdf_name
FROM sheet_index
"@
$params = @{}
if (-not [string]::IsNullOrWhiteSpace($FolderPathFilter)) {
    $sql += " WHERE folder_path LIKE @fp"
    $params['fp'] = ('%' + $FolderPathFilter.Trim().Trim('\') + '%')
}
$sql += " ORDER BY folder_path, document_name"

$qRes = Invoke-QCDatabaseQuery -Config $config -Sql $sql -Parameters $params
if (-not $qRes.IsSuccess) { throw $qRes.Message }

$rows = @()
foreach ($r in @($qRes.Data.table.Rows)) {
    $rows += [pscustomobject]@{
        documentGuid = [string]$r.document_guid
        documentName = [string]$r.document_name
        folderPath   = [string]$r.folder_path
        extension    = if ($r.extension -is [DBNull]) { '' } else { [string]$r.extension }
        sourceType   = if ($r.source_type -is [DBNull]) { '' } else { [string]$r.source_type }
        pwStateName  = if ($r.pw_state_name -is [DBNull]) { '' } else { [string]$r.pw_state_name }
        qcPdfGuid    = if ($r.qc_pdf_guid -is [DBNull]) { '' } else { [string]$r.qc_pdf_guid }
        qcPdfName    = if ($r.qc_pdf_name -is [DBNull]) { '' } else { [string]$r.qc_pdf_name }
    }
}
if ($Limit -gt 0) { $rows = @($rows | Select-Object -First $Limit) }

$pwCfg = @{}
if ($config.ContainsKey('projectWise') -and $config.projectWise) {
    $raw = $config.projectWise
    if ($raw -is [hashtable]) { $pwCfg = $raw }
    elseif ($raw.PSObject) { foreach ($p in $raw.PSObject.Properties) { $pwCfg[$p.Name] = $p.Value } }
}
$credPath = if ($pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
$ds = if ($pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { '' }

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw $credRes.Message }
$connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $connRes.IsSuccess) { throw $connRes.Message }

$summary = [ordered]@{
    rowsScanned          = $rows.Count
    guidCountRequested   = 0
    guidCountValid       = 0
    invalidGuids         = @()
    statesRead           = 0
    dbUpdatesPlanned     = 0
    dbUpdatesApplied     = 0
    missingStateGuids    = 0
    errors               = @()
}

try {
    $guids = [System.Collections.Generic.List[string]]::new()
    foreach ($row in $rows) {
        if ($row.documentGuid) { $guids.Add([string]$row.documentGuid) | Out-Null }
        if ($IncludeQcPdfGuids.IsPresent -and $row.qcPdfGuid) { $guids.Add([string]$row.qcPdfGuid) | Out-Null }
    }
    $unique = @($guids | Select-Object -Unique)
    $summary.guidCountRequested = $unique.Count

    $valid = @()
    $invalid = [System.Collections.Generic.List[string]]::new()
    foreach ($g in $unique) {
        if (Test-PWValidDocumentGuid -DocumentGuid $g) { $valid += [string]$g }
        elseif (-not [string]::IsNullOrWhiteSpace($g)) { $invalid.Add([string]$g) | Out-Null }
    }
    $summary.guidCountValid = $valid.Count
    if ($invalid.Count -gt 0) { $summary.invalidGuids = @($invalid | Select-Object -Unique) }

    $stateByGuid = @{}
    if ($valid.Count -gt 0) {
        try { $stateByGuid = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $valid } catch { }
    }
    $summary.statesRead = $stateByGuid.Keys.Count

    foreach ($row in $rows) {
        $dg = [string]$row.documentGuid
        if (-not (Test-PWValidDocumentGuid -DocumentGuid $dg)) { continue }
        $key = $dg.ToLowerInvariant()
        $pwState = if ($stateByGuid.ContainsKey($key)) { [string]$stateByGuid[$key] } else { '' }
        if ([string]::IsNullOrWhiteSpace($pwState)) {
            $summary.missingStateGuids++
            continue
        }

        $differs = ([string]$row.pwStateName).Trim() -ne ([string]$pwState).Trim()
        if (-not $differs) { continue }

        $summary.dbUpdatesPlanned++
        if ($doWrites) {
            $target = "$($row.folderPath)\$($row.documentName)"
            if ($PSCmdlet.ShouldProcess($target, "Update sheet_index.pw_state_name to '$pwState'")) {
                try {
                    Write-QCSheetIndex -Config $config `
                        -DocumentGuid $dg `
                        -DocumentName $row.documentName `
                        -FolderPath $row.folderPath `
                        -Extension $row.extension `
                        -SourceType $row.sourceType `
                        -PwStateName $pwState `
                        -SetOwnershipFromProjectWise
                    $summary.dbUpdatesApplied++
                } catch {
                    $summary.errors += "DB update failed for $dg : $($_.Exception.Message)"
                }
            }
        }
    }
} finally {
    Disconnect-PW | Out-Null
}

if ($Pretty) {
    $summary | ConvertTo-Json -Depth 6
} else {
    Write-Host "rowsScanned=$($summary.rowsScanned) guidValid=$($summary.guidCountValid) statesRead=$($summary.statesRead) planned=$($summary.dbUpdatesPlanned) applied=$($summary.dbUpdatesApplied) missingStateGuids=$($summary.missingStateGuids) invalidGuids=$($summary.invalidGuids.Count) errors=$($summary.errors.Count)"
}

