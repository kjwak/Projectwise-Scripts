<#
.SYNOPSIS
  Build/refresh _StatusSet.pdf using legacy combine_status_set.ps1 method (native implementation).

.DESCRIPTION
  Thin wrapper over modules\QC.StatusSet.psm1.
  For ProjectWise folders, run from pwps (ProjectWise PowerShell) so pwps_dab cmdlets are available.

.PARAMETER SheetsFolderPath
  Logical PW folder (e.g. Documents\Proj\CADD\Sheets) or local disk path.

.PARAMETER LocalRoot
  Root for legacy-style manifest/cache:
    - status_set_manifest_<safe>.json
    - status_set_cache\<safe>\*.pdf
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $SheetsFolderPath,

    [Parameter(Mandatory = $false)]
    [string] $LocalRoot = 'C:\PW_QC_LOCAL',

    [Parameter(Mandatory = $false)]
    [bool] $OneLevelDeep = $true,

    [Parameter(Mandatory = $false)]
    [switch] $ForceRebuild,

    [Parameter(Mandatory = $false)]
    [switch] $DryRun
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.StatusSet.psm1') -Force

$job = @{
    id = 'manual_combine'
    type = 'STATUS_SET_GEN'
    sourceFolder = $SheetsFolderPath
    sourcePath = $SheetsFolderPath
    sourceName = '_folder_'
    metadata = @{
        candidate = @{
            oneLevelDeep = $OneLevelDeep
        }
    }
}

$cfg = @{
    dryRun = [bool]$DryRun
    statusSet = @{
        localRoot = $LocalRoot
        forceRebuild = [bool]$ForceRebuild
    }
    qcPrepend = @{
        qpdfExePath = (Join-Path $repoRoot 'tools\qpdf\bin\qpdf.exe')
    }
}

$r = Invoke-StatusSetNativeJob -Job $job -Config $cfg
if (-not $r.IsSuccess) {
    Write-Error $r.Message
    exit 1
}
$r | ConvertTo-Json -Depth 8
