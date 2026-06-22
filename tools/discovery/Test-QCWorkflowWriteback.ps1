<#
.SYNOPSIS
Controlled ProjectWise QC workflow writeback test harness.

.DESCRIPTION
This script can execute state and/or attribute writeback against one explicit test
ProjectWise document. It is intentionally hard to run accidentally: callers must
provide both -ConfirmWrites and -TestDocumentPath. It prints dry-run/planned
operations before any write, supports state-only, attribute-only, or combined
mode, and can attempt a best-effort rollback of state and attributes.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$ConfirmWrites,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$TestDocumentPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('StateOnly','AttributeOnly','Combined')]
    [string]$Mode = 'Combined',

    [Parameter(Mandatory = $false)]
    [string]$TargetState,

    [Parameter(Mandatory = $false)]
    [string]$ExpectedWorkflowName = 'QC Review Workflow',

    [Parameter(Mandatory = $false)]
    [string]$AttributesJson,

    [Parameter(Mandatory = $false)]
    [switch]$Rollback,

    [Parameter(Mandatory = $false)]
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'

if (-not $ConfirmWrites) {
    throw 'Refusing to run: -ConfirmWrites is required for this controlled writeback test.'
}
if ([string]::IsNullOrWhiteSpace($TestDocumentPath)) {
    throw 'Refusing to run: -TestDocumentPath is required.'
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.Workflow.psm1') -Force

foreach ($moduleName in @('pwps','pwps_dab')) {
    if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        if ($available) { Import-Module $available.Name -ErrorAction SilentlyContinue | Out-Null }
    }
}

function Split-TestDocumentPath {
    param([string]$Path)
    $clean = $Path.Trim().TrimEnd('\')
    if ($clean -match '^(?i)pw:\\[^\\]+\\Documents\\(.+)$') { $clean = $Matches[1] }
    elseif ($clean -match '^(?i)Documents\\(.+)$') { $clean = $Matches[1] }
    $folder = Split-Path -Path $clean -Parent
    $name = Split-Path -Path $clean -Leaf
    if ([string]::IsNullOrWhiteSpace($folder) -or [string]::IsNullOrWhiteSpace($name)) { throw "TestDocumentPath must include folder and document name: $Path" }
    return @{ folder = $folder; name = $name }
}

function Get-TestDocument {
    param([string]$Path)
    $parts = Split-TestDocumentPath -Path $Path
    $cmd = Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue
    if (-not $cmd) { throw 'Get-PWDocumentsBySearch is required to locate the test document.' }
    $doc = Get-PWDocumentsBySearch -FolderPath $parts.folder -JustThisFolder -DocumentName $parts.name -PopulatePath -ErrorAction Stop | Select-Object -First 1
    if (-not $doc) { throw "Test document not found: $Path" }
    return @{ document = $doc; folder = $parts.folder; name = $parts.name }
}

function Get-OriginalAttributeSnapshot {
    param([object]$Document, [string[]]$AttributeNames)
    $snapshot = @{}
    $cmd = Get-Command -Name Get-PWDocumentEAttributes -ErrorAction SilentlyContinue
    if (-not $cmd -or -not $Document.PSObject.Properties['DocumentID'] -or -not $Document.PSObject.Properties['ProjectID']) { return $snapshot }
    try {
        $attrs = @(Get-PWDocumentEAttributes -DocumentID $Document.DocumentID -ProjectID $Document.ProjectID -ErrorAction Stop)
        foreach ($name in @($AttributeNames)) {
            foreach ($row in @($attrs)) {
                $rowName = $null
                foreach ($prop in @('Name','ColumnName','AttributeName','DisplayName')) {
                    if ($row.PSObject.Properties[$prop] -and $row.$prop) { $rowName = [string]$row.$prop; break }
                }
                if ($rowName -eq $name) {
                    $snapshot[$name] = if ($row.PSObject.Properties['Value']) { $row.Value } elseif ($row.PSObject.Properties['AttributeValue']) { $row.AttributeValue } else { $null }
                    break
                }
            }
        }
    } catch { }
    return $snapshot
}

$located = Get-TestDocument -Path $TestDocumentPath
$doc = $located.document
$originalInfo = Get-PWDocumentWorkflowInfo -Document $doc -Context @{ documentPath = $TestDocumentPath; folderPath = $located.folder }
$target = if ([string]::IsNullOrWhiteSpace($TargetState)) { $originalInfo.Data.stateName } else { $TargetState }
if (($Mode -eq 'StateOnly' -or $Mode -eq 'Combined') -and [string]::IsNullOrWhiteSpace($target)) {
    throw 'TargetState is required when the current state cannot be read from the test document.'
}

$attributes = @{
    automationLastRun = Get-QCTimestamp
    automationResult = 'ControlledWritebackTest'
    automationError = $null
}
if (-not [string]::IsNullOrWhiteSpace($AttributesJson)) {
    $custom = $AttributesJson | ConvertFrom-Json
    foreach ($p in $custom.PSObject.Properties) { $attributes[$p.Name] = $p.Value }
}

$workflowMode = if ($Mode -eq 'AttributeOnly') { 'AttributesOnly' } else { 'StateAndAttributes' }

$config = @{
    dryRun = $false
    qcWorkflow = @{
        enabled = $true
        strictMode = $true
        dryRunWriteback = $true
        mode = $workflowMode
        workflowName = $ExpectedWorkflowName
        expectedWorkflowName = $ExpectedWorkflowName
        autoAssignWorkflow = $false
        autoSetState = ($Mode -eq 'StateOnly' -or $Mode -eq 'Combined')
        autoWriteAttributes = ($Mode -eq 'AttributeOnly' -or $Mode -eq 'Combined')
        stateAfterSuccessfulPrepend = $target
        attributeMap = @{
            automationLastRun = 'QC_Automation_Last_Run'
            automationResult = 'QC_Automation_Result'
            automationError = 'QC_Automation_Error'
        }
    }
}
$context = @{
    document = $doc
    documentPath = $TestDocumentPath
    folderPath = $located.folder
    targetState = $target
    resultStatus = 'Succeeded'
    attributes = $attributes
}

$planned = Invoke-QCWorkflowWriteback -Config $config -Context $context
Write-Host 'Planned operations (dry-run):'
$planned | ConvertTo-Json -Depth 50 | Write-Host

$config.qcWorkflow.dryRunWriteback = $false
$originalAttributeSnapshot = Get-OriginalAttributeSnapshot -Document $doc -AttributeNames @('QC_Automation_Last_Run','QC_Automation_Result','QC_Automation_Error')
$writeResult = Invoke-QCWorkflowWriteback -Config $config -Context $context
Write-Host 'Executed operations:'
$writeResult | ConvertTo-Json -Depth 50 | Write-Host

$rollbackResult = $null
if ($Rollback) {
    Write-Host 'Rollback requested; attempting best-effort rollback.'
    $rollbackActions = @()
    if (($Mode -eq 'StateOnly' -or $Mode -eq 'Combined') -and -not [string]::IsNullOrWhiteSpace($originalInfo.Data.stateName) -and $originalInfo.Data.stateName -ne $target) {
        $rollbackActions += Set-PWQCWorkflowState -Settings (Get-QCWorkflowSettings -Config $config) -Context $context -StateName $originalInfo.Data.stateName -DryRun:$false
    }
    if (($Mode -eq 'AttributeOnly' -or $Mode -eq 'Combined') -and $originalAttributeSnapshot.Keys.Count -gt 0) {
        $rollbackContext = @{
            document = $doc
            attributes = @{
                automationLastRun = $originalAttributeSnapshot['QC_Automation_Last_Run']
                automationResult = $originalAttributeSnapshot['QC_Automation_Result']
                automationError = $originalAttributeSnapshot['QC_Automation_Error']
            }
        }
        $rollbackActions += Set-PWQCAttributes -Settings (Get-QCWorkflowSettings -Config $config) -Context $rollbackContext -DryRun:$false
    }
    $rollbackResult = [pscustomobject]@{ attempted = $true; actions = @($rollbackActions) }
    $rollbackResult | ConvertTo-Json -Depth 50 | Write-Host
}

$result = [ordered]@{
    testDocumentPath = $TestDocumentPath
    mode = $Mode
    confirmWrites = [bool]$ConfirmWrites
    rollbackRequested = [bool]$Rollback
    originalWorkflowInfo = $originalInfo
    planned = $planned
    executed = $writeResult
    rollback = $rollbackResult
}

if ($Pretty) { $result | ConvertTo-Json -Depth 50 }
else { $result | ConvertTo-Json -Depth 50 -Compress }
