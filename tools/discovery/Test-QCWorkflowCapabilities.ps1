<#
.SYNOPSIS
Read-only discovery for ProjectWise QC workflow/state/attribute cmdlet capabilities.

.DESCRIPTION
Imports pwps/pwps_dab when available, reports module versions, candidate cmdlet
availability, parameter sets, and read/write capability hints needed by the QC
workflow framework. The script does not modify ProjectWise documents.

If a test document is supplied, the script attempts read-only inspection of the
document object's workflow/state-like properties and environment attributes.
It assumes the caller has already opened a ProjectWise connection when document
reads require one.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$FolderPath,

    [Parameter(Mandatory = $false)]
    [string]$DocumentName,

    [Parameter(Mandatory = $false)]
    [string]$DocumentGuid,

    [Parameter(Mandatory = $false)]
    [int]$DocumentID,

    [Parameter(Mandatory = $false)]
    [int]$ProjectID,

    [Parameter(Mandatory = $false)]
    [switch]$Pretty
)

$ErrorActionPreference = 'Continue'

function Add-WarningMessage {
    param(
        [System.Collections.Generic.List[string]]$Warnings,
        [string]$Message
    )
    if (-not [string]::IsNullOrWhiteSpace($Message)) { $Warnings.Add($Message) | Out-Null }
}

function Import-PWModuleIfAvailable {
    param([string]$Name)

    $result = [ordered]@{
        name = $Name
        loaded = $false
        imported = $false
        version = $null
        path = $null
        error = $null
    }

    $loaded = Get-Module -Name $Name -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($loaded) {
        $result.loaded = $true
        $result.version = [string]$loaded.Version
        $result.path = [string]$loaded.Path
        return [pscustomobject]$result
    }

    $available = Get-Module -ListAvailable -Name $Name -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $available) { return [pscustomobject]$result }

    try {
        Import-Module $available.Name -ErrorAction Stop | Out-Null
        $loaded = Get-Module -Name $Name -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        $result.loaded = [bool]$loaded
        $result.imported = [bool]$loaded
        if ($loaded) {
            $result.version = [string]$loaded.Version
            $result.path = [string]$loaded.Path
        }
    } catch {
        $result.error = $_.Exception.Message
    }
    return [pscustomobject]$result
}

function Get-CommandShape {
    param([System.Management.Automation.CommandInfo]$Command)

    if (-not $Command) { return $null }
    $parameters = @()
    foreach ($p in @($Command.Parameters.GetEnumerator() | Sort-Object Key)) {
        $parameters += [pscustomobject]@{
            name = [string]$p.Key
            parameterType = if ($p.Value.ParameterType) { [string]$p.Value.ParameterType.FullName } else { $null }
            aliases = @($p.Value.Aliases)
            isDynamic = [bool]$p.Value.IsDynamic
        }
    }

    $parameterSets = @()
    try {
        foreach ($set in @($Command.ParameterSets)) {
            $parameterSets += [pscustomobject]@{
                name = [string]$set.Name
                isDefault = [bool]$set.IsDefault
                parameters = @($set.Parameters | ForEach-Object {
                    [pscustomobject]@{
                        name = [string]$_.Name
                        isMandatory = [bool]$_.IsMandatory
                        position = [int]$_.Position
                        valueFromPipeline = [bool]$_.ValueFromPipeline
                        valueFromPipelineByPropertyName = [bool]$_.ValueFromPipelineByPropertyName
                    }
                })
            }
        }
    } catch { }

    return [pscustomobject]@{
        name = [string]$Command.Name
        moduleName = [string]$Command.ModuleName
        source = [string]$Command.Source
        commandType = [string]$Command.CommandType
        version = if ($Command.Version) { [string]$Command.Version } else { $null }
        parameters = $parameters
        parameterSets = $parameterSets
    }
}

function Test-CommandHasAnyParameter {
    param(
        [string]$CommandName,
        [string[]]$ParameterNames
    )
    $cmd = Get-Command -Name $CommandName -ErrorAction SilentlyContinue
    if (-not $cmd) { return $false }
    foreach ($name in $ParameterNames) {
        if ($cmd.Parameters.ContainsKey($name)) { return $true }
    }
    return $false
}

function Get-DocumentByReadOnlyInputs {
    param(
        [string]$FolderPath,
        [string]$DocumentName,
        [string]$DocumentGuid,
        [System.Collections.Generic.List[string]]$Warnings
    )

    if (-not [string]::IsNullOrWhiteSpace($FolderPath) -and -not [string]::IsNullOrWhiteSpace($DocumentName)) {
        $cmd = Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue
        if (-not $cmd) {
            Add-WarningMessage -Warnings $Warnings -Message 'Get-PWDocumentsBySearch is unavailable; cannot read test document by FolderPath/DocumentName.'
            return $null
        }
        try {
            return Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocumentName -PopulatePath -ErrorAction Stop | Select-Object -First 1
        } catch {
            Add-WarningMessage -Warnings $Warnings -Message ('Read-only document search failed: ' + $_.Exception.Message)
            return $null
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($DocumentGuid)) {
        foreach ($candidate in @('Get-PWDocumentsBySearchExtended','Get-PWDocumentsBySearch')) {
            $cmd = Get-Command -Name $candidate -ErrorAction SilentlyContinue
            if (-not $cmd) { continue }
            if (-not $cmd.Parameters.ContainsKey('DocumentGUID') -and -not $cmd.Parameters.ContainsKey('DocumentGuid')) { continue }
            try {
                $paramName = if ($cmd.Parameters.ContainsKey('DocumentGUID')) { 'DocumentGUID' } else { 'DocumentGuid' }
                $args = @{ ErrorAction = 'Stop' }
                $args[$paramName] = $DocumentGuid
                if ($cmd.Parameters.ContainsKey('PopulatePath')) { $args['PopulatePath'] = $true }
                return & $candidate @args | Select-Object -First 1
            } catch {
                Add-WarningMessage -Warnings $Warnings -Message ("Read-only document GUID lookup via $candidate failed: " + $_.Exception.Message)
            }
        }
        Add-WarningMessage -Warnings $Warnings -Message 'No available document search cmdlet exposed a DocumentGUID/DocumentGuid parameter.'
    }

    return $null
}

function Get-WorkflowStateProperties {
    param([object]$Document)

    if (-not $Document) { return $null }
    $names = @('Workflow','WorkflowName','WorkflowState','WorkflowStateName','State','StateName','DocumentState','CurrentState')
    $values = [ordered]@{}
    foreach ($name in $names) {
        try {
            if ($Document.PSObject.Properties[$name]) { $values[$name] = $Document.$name }
        } catch { }
    }
    return [pscustomobject]$values
}

function Get-ReadOnlyEAttributes {
    param(
        [object]$Document,
        [int]$DocumentID,
        [int]$ProjectID,
        [System.Collections.Generic.List[string]]$Warnings
    )

    $cmd = Get-Command -Name Get-PWDocumentEAttributes -ErrorAction SilentlyContinue
    if (-not $cmd) { return $null }

    $docId = $DocumentID
    $projId = $ProjectID
    try { if (-not $docId -and $Document -and $Document.PSObject.Properties['DocumentID']) { $docId = [int]$Document.DocumentID } } catch { }
    try { if (-not $projId -and $Document -and $Document.PSObject.Properties['ProjectID']) { $projId = [int]$Document.ProjectID } } catch { }

    if (-not $docId -or -not $projId) {
        Add-WarningMessage -Warnings $Warnings -Message 'Get-PWDocumentEAttributes is available, but DocumentID and ProjectID were not supplied or found on the test document.'
        return $null
    }

    try {
        return Get-PWDocumentEAttributes -DocumentID $docId -ProjectID $projId -ErrorAction Stop
    } catch {
        Add-WarningMessage -Warnings $Warnings -Message ('Read-only Get-PWDocumentEAttributes failed: ' + $_.Exception.Message)
        return $null
    }
}

$candidateCmdlets = @(
    'Get-PWWorkflows',
    'Get-PWWorkflowStateLinks',
    'Get-PWDocumentEAttributes',
    'Get-PWEnvironmentColumns',
    'Get-PWDocumentsBySearch',
    'Get-PWDocumentsBySearchExtended',
    'Get-PWDocumentsBySearchWithReturnColumns',
    'Update-PWDocumentAttributes',
    'Set-PWDocumentState',
    'Set-PWDocumentWorkflow'
)

$writeCandidateNames = @('Update-PWDocumentAttributes','Set-PWDocumentState','Set-PWDocumentWorkflow','Set-PWDocumentWorkflowState','Update-PWDocumentProperties')
$warnings = [System.Collections.Generic.List[string]]::new()

$pwps = Import-PWModuleIfAvailable -Name 'pwps'
$pwpsDab = Import-PWModuleIfAvailable -Name 'pwps_dab'
if (-not $pwps.loaded) { Add-WarningMessage -Warnings $warnings -Message 'pwps module is not loaded/available.' }
if (-not $pwpsDab.loaded) { Add-WarningMessage -Warnings $warnings -Message 'pwps_dab module is not loaded/available.' }
if ($pwps.error) { Add-WarningMessage -Warnings $warnings -Message ('pwps import error: ' + $pwps.error) }
if ($pwpsDab.error) { Add-WarningMessage -Warnings $warnings -Message ('pwps_dab import error: ' + $pwpsDab.error) }

$available = @()
$missing = @()
$commandDetails = @{}
foreach ($name in $candidateCmdlets) {
    $cmd = Get-Command -Name $name -ErrorAction SilentlyContinue
    if ($cmd) {
        $available += $name
        $commandDetails[$name] = Get-CommandShape -Command $cmd
    } else {
        $missing += $name
    }
}

$matchingCmdlets = @()
foreach ($pattern in @('*Workflow*','*State*','*Attribute*','*Environment*','*Document*')) {
    $matchingCmdlets += @(Get-Command -Name $pattern -ErrorAction SilentlyContinue | Where-Object { $_.ModuleName -match '^(pwps|pwps_dab)$' -or $_.Source -match '^(pwps|pwps_dab)$' })
}
$matchingCmdlets = @($matchingCmdlets | Sort-Object Name -Unique | ForEach-Object {
    [pscustomobject]@{
        name = [string]$_.Name
        moduleName = [string]$_.ModuleName
        commandType = [string]$_.CommandType
    }
})

$candidateWriteCmdlets = @()
foreach ($name in $writeCandidateNames) {
    $cmd = Get-Command -Name $name -ErrorAction SilentlyContinue
    if ($cmd) { $candidateWriteCmdlets += (Get-CommandShape -Command $cmd) }
}

$readCapabilities = [ordered]@{
    listWorkflows = [bool](Get-Command -Name Get-PWWorkflows -ErrorAction SilentlyContinue)
    listWorkflowStateLinks = [bool](Get-Command -Name Get-PWWorkflowStateLinks -ErrorAction SilentlyContinue)
    readDocumentEnvironmentAttributes = [bool](Get-Command -Name Get-PWDocumentEAttributes -ErrorAction SilentlyContinue)
    readEnvironmentColumns = [bool](Get-Command -Name Get-PWEnvironmentColumns -ErrorAction SilentlyContinue)
    searchDocumentsBasic = [bool](Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue)
    searchDocumentsExtended = [bool](Get-Command -Name Get-PWDocumentsBySearchExtended -ErrorAction SilentlyContinue)
    searchDocumentsWithReturnColumns = [bool](Get-Command -Name Get-PWDocumentsBySearchWithReturnColumns -ErrorAction SilentlyContinue)
}

$writeCapabilities = [ordered]@{
    assignWorkflow = [bool](Get-Command -Name Set-PWDocumentWorkflow -ErrorAction SilentlyContinue)
    setState = [bool](Get-Command -Name Set-PWDocumentState -ErrorAction SilentlyContinue)
    setWorkflowState = [bool](Get-Command -Name Set-PWDocumentWorkflowState -ErrorAction SilentlyContinue)
    updateEnvironmentAttributes = [bool](Get-Command -Name Update-PWDocumentAttributes -ErrorAction SilentlyContinue)
    updateDocumentProperties = [bool](Get-Command -Name Update-PWDocumentProperties -ErrorAction SilentlyContinue)
    updateAttributesHasInputDocuments = Test-CommandHasAnyParameter -CommandName 'Update-PWDocumentAttributes' -ParameterNames @('InputDocuments','InputDocument')
    updateAttributesHasAttributes = Test-CommandHasAnyParameter -CommandName 'Update-PWDocumentAttributes' -ParameterNames @('Attributes')
    setStateHasStateName = Test-CommandHasAnyParameter -CommandName 'Set-PWDocumentState' -ParameterNames @('StateName','State')
    setWorkflowHasWorkflowName = Test-CommandHasAnyParameter -CommandName 'Set-PWDocumentWorkflow' -ParameterNames @('WorkflowName','Workflow')
}

$documentRead = $null
if ($FolderPath -or $DocumentName -or $DocumentGuid -or ($DocumentID -and $ProjectID)) {
    $doc = Get-DocumentByReadOnlyInputs -FolderPath $FolderPath -DocumentName $DocumentName -DocumentGuid $DocumentGuid -Warnings $warnings
    $workflowStateProps = Get-WorkflowStateProperties -Document $doc
    $eattrs = Get-ReadOnlyEAttributes -Document $doc -DocumentID $DocumentID -ProjectID $ProjectID -Warnings $warnings
    $documentRead = [ordered]@{
        attempted = $true
        foundDocument = [bool]$doc
        workflowStateProperties = $workflowStateProps
        environmentAttributes = $eattrs
    }
} else {
    $documentRead = [ordered]@{
        attempted = $false
        foundDocument = $false
        workflowStateProperties = $null
        environmentAttributes = $null
    }
}

$result = [ordered]@{
    pwpsLoaded = [bool]$pwps.loaded
    pwpsDabLoaded = [bool]$pwpsDab.loaded
    modules = @($pwps, $pwpsDab)
    availableCmdlets = @($available | Sort-Object)
    missingCmdlets = @($missing | Sort-Object)
    commandDetails = $commandDetails
    matchingProjectWiseCmdlets = @($matchingCmdlets)
    candidateWriteCmdlets = @($candidateWriteCmdlets)
    readCapabilities = [pscustomobject]$readCapabilities
    writeCapabilities = [pscustomobject]$writeCapabilities
    documentRead = [pscustomobject]$documentRead
    warnings = @($warnings)
}

$depth = 50
if ($Pretty) { $result | ConvertTo-Json -Depth $depth }
else { $result | ConvertTo-Json -Depth $depth -Compress }
