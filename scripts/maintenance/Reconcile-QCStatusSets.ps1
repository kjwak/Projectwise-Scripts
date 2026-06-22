<#
.SYNOPSIS
Walks every locally-built _StatusSet.pdf manifest under statusSet.localRoot and
re-syncs it to ProjectWise. Use this on application restart to catch up any
status sets that finished locally but never made it to PW (legacy parity:
every startup re-checks every manifest against the PW copy).

.DESCRIPTION
For each workspace under localRoot that contains both _statusset.manifest.json
and _StatusSet.pdf:
  - PW has no _StatusSet.pdf      -> create   (New-PWDocument)
  - Local PDF newer than PW copy  -> update   (Update-PWDocumentFile)
  - Already in sync               -> no-op

Outputs structured one-line JSON per workspace plus a final summary.
This script is read-write to ProjectWise but does not touch the queue.

.EXAMPLE
PS> .\scripts\Reconcile-QCStatusSets.ps1
PS> .\scripts\Reconcile-QCStatusSets.ps1 -DryRun     # show what would happen, no PW writes
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath = (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'appsettings.json'),
    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Paths.psm1')   -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.StatusSet.psm1') -Force

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

Write-QCJsonLog 'Information' 'RECONCILE_START' 'Status set reconciliation started.' @{
    appSettingsPath = $AppSettingsPath
    dryRun          = [bool]$DryRun.IsPresent
}

$pwCfg = @{}
if ($config.ContainsKey('projectWise') -and $config.projectWise) {
    $raw = $config.projectWise
    if ($raw -is [hashtable]) { $pwCfg = $raw } elseif ($raw.PSObject) { foreach ($p in $raw.PSObject.Properties) { $pwCfg[$p.Name] = $p.Value } }
}
$credPath = if ($pwCfg.ContainsKey('credentialPath') -and $pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
$ds = if ($pwCfg.ContainsKey('datasourceName') -and $pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { '' }

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }

$conn = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $conn.IsSuccess) { throw ($conn.Code + ': ' + $conn.Message) }
Write-QCJsonLog 'Information' 'RECONCILE_PW_CONNECT_OK' 'Connected to ProjectWise.' @{ datasourceName = $ds }

try {
    if ($DryRun.IsPresent) {
        # Dry run: enumerate only; do not call Sync-StatusSetWorkspaceToPw.
        $ss = @{}
        if ($config.ContainsKey('statusSet') -and $config.statusSet) {
            $raw = $config.statusSet
            if ($raw -is [hashtable]) { $ss = $raw } elseif ($raw.PSObject) { foreach ($p in $raw.PSObject.Properties) { $ss[$p.Name] = $p.Value } }
        }
        $localRoot     = if ($ss.ContainsKey('localRoot') -and $ss.localRoot) { [string]$ss.localRoot } else { 'C:\PW_QC_LOCAL' }
        $manifestName  = if ($ss.ContainsKey('manifestFileName') -and $ss.manifestFileName) { [string]$ss.manifestFileName } else { '_statusset.manifest.json' }
        $statusPdfName = if ($ss.ContainsKey('statusSetPdfName') -and $ss.statusSetPdfName) { [string]$ss.statusSetPdfName } else { '_StatusSet.pdf' }

        $walk = Get-StatusSetWorkspaceManifests -LocalRoot $localRoot -ManifestFileName $manifestName -StatusSetPdfName $statusPdfName
        if (-not $walk.IsSuccess) { throw ($walk.Code + ': ' + $walk.Message) }
        foreach ($rec in @($walk.Data.records)) {
            Write-QCJsonLog 'Information' 'RECONCILE_DRYRUN_CANDIDATE' 'Would reconcile.' @{
                workspaceDir = [string]$rec.workspaceDir
                pwFolder     = [string]$rec.pwPath
                outputPdf    = [string]$rec.outputPdf
                localMtime   = if ($rec.outputPdfLastWriteUtc) { ConvertTo-QCTimestamp -DateTime ([datetime]$rec.outputPdfLastWriteUtc) } else { $null }
            }
        }
        Write-QCJsonLog 'Information' 'RECONCILE_DRYRUN_DONE' 'Dry-run reconciliation completed.' @{
            considered = [int]$walk.Data.recordCount
            skipped    = [int]$walk.Data.skipCount
        }
    } else {
        $cb = {
            param($evt)
            $level = if ([bool]$evt.isSuccess) { 'Information' } else { 'Warning' }
            $code = "RECONCILE_$([string]$evt.code -replace '^STATUS_SET_RECONCILE_','' )"
            Write-QCJsonLog $level $code ([string]$evt.message) @{
                workspaceDir = [string]$evt.workspaceDir
                pwFolder     = [string]$evt.pwFolder
                sheetsFolder = [string]$evt.sheetsFolder
                outputPdf    = [string]$evt.outputPdf
                data         = $evt.data
            }
        }
        $res = Invoke-StatusSetReconcile -Config $config -LogCallback $cb
        if (-not $res.IsSuccess) { throw ($res.Code + ': ' + $res.Message) }
        Write-QCJsonLog 'Information' 'RECONCILE_DONE' 'Status set reconciliation completed.' @{
            counts   = $res.Data.counts
            failures = $res.Data.failures
            skipped  = $res.Data.skipped
        }
    }
} finally {
    try { Disconnect-PW | Out-Null } catch { }
}
