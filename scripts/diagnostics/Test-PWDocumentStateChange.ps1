<#
.SYNOPSIS
Manually test whether PWPS_DAB can change a ProjectWise document workflow state and verify read-back.

.DESCRIPTION
Resolves a document without attribute bags for Set-PWDocumentState, probes the installed cmdlet shape,
and verifies read-back after each non-throwing attempt.

Exit codes:
  0 = read-back WorkflowState equals TargetState (verified)
  2 = at least one call did not throw but read-back did not match TargetState
  1 = all attempts threw hard errors

.EXAMPLE
.\scripts\Test-PWDocumentStateChange.ps1 -ResolveOnly

.EXAMPLE
.\scripts\Test-PWDocumentStateChange.ps1

.NOTES
pwps_dab requires MTA. This script auto re-launches with powershell.exe -MTA when started from an STA host.

Validated in this environment (080J082001ca001-chk.pdf, workflow TYPSA QC, In Development -> Originated):
  - Set-PWDocumentState -InputDocuments @($cleanDoc) -State 'Originated' returned without error but read-back stayed In Development (unverified).
  - The same call with -Force persisted and read-back verified as Originated.
  - Runtime lane PDF workflow writeback must therefore reload a clean document (no -GetAttributes) and use -Force with read-back verification.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$Datasource = 'typsa-us-pw.bentley.com:typsa-us-pw-03',
    [string]$FolderPath = 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1',
    [string]$DocumentName = '080J082001ca001-chk.pdf',
    [string]$TargetState = 'Originated',
    [switch]$UseStoredCredential,
    [switch]$WhatIfOnly,
    [switch]$ResolveOnly,
    [switch]$NoAttributesForStateChange
)

# Default -NoAttributesForStateChange to $true when the switch is not explicitly passed.
$useNoAttributesForStateChange = $true
if ($PSBoundParameters.ContainsKey('NoAttributesForStateChange') -and -not $NoAttributesForStateChange) {
    $useNoAttributesForStateChange = $false
}

# pwps_dab requires MTA; Cursor/VS Code terminals often use STA.
$scriptPath = $PSCommandPath
if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
if (-not $scriptPath) {
    $scriptPath = Join-Path $PSScriptRoot 'Test-PWDocumentStateChange.ps1'
}
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

$ErrorActionPreference = 'Stop'

function Write-Step([string]$Message) {
    Write-Host $Message -ForegroundColor Cyan
}

function Get-PWCredFromKeyValueFile([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { throw "Credential file not found: $Path" }
    $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    $uLine = $lines | Where-Object { $_ -match '^\s*username\s*=' } | Select-Object -First 1
    $pLine = $lines | Where-Object { $_ -match '^\s*password\s*=' } | Select-Object -First 1
    if (-not $uLine -or -not $pLine) { throw "Bad credential file format (expected username=/password=): $Path" }
    $user = ($uLine -split '=', 2)[1].Trim()
    $pass = ($pLine -split '=', 2)[1].Trim()
    if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) { throw "Credential file missing values: $Path" }
    $sec = ConvertTo-SecureString $pass -AsPlainText -Force
    return [pscredential]::new($user, $sec)
}

function Get-DocumentWorkflowStateName {
    param(
        [object]$Document,
        [string]$FolderPath = '',
        [string]$DocumentName = '',
        [string]$DocumentGuid = ''
    )
    if ($Document) {
        foreach ($name in @('WorkflowState', 'StateName', 'State', 'WorkflowStateName', 'CurrentState')) {
            try {
                if ($Document.PSObject.Properties[$name]) {
                    $v = [string]$Document.$name
                    if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }
                }
            } catch { }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($FolderPath) -and -not [string]::IsNullOrWhiteSpace($DocumentName) `
        -and (Get-Command -Name 'Get-PWDocumentWorkflowStateName' -ErrorAction SilentlyContinue)) {
        try {
            $pw = Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $DocumentName -DocumentGuid $DocumentGuid
            if (-not [string]::IsNullOrWhiteSpace($pw)) { return [string]$pw }
        } catch { }
    }
    return ''
}

function Get-PwSearchParams {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [bool]$IncludeAttributes = $false
    )
    $searchCmd = Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction Stop
    $params = @{
        FolderPath = $FolderPath
        JustThisFolder = $true
        DocumentName = $DocumentName
        PopulatePath = $true
    }
    if ($searchCmd.Parameters.ContainsKey('GetAttributes')) {
        $params['GetAttributes'] = [bool]$IncludeAttributes
    }
    return $params
}

function Get-CleanPwDocumentForStateChange {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = '',
        [bool]$NoAttributes = $true
    )

    $includeAttrs = -not $NoAttributes
    $searchParams = Get-PwSearchParams -FolderPath $FolderPath -DocumentName $DocumentName -IncludeAttributes:$includeAttrs
    $doc = Get-PWDocumentsBySearch @searchParams -ErrorAction Stop
    if ($doc) { return $doc }

    if (-not [string]::IsNullOrWhiteSpace($DocumentGuid) `
        -and (Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue)) {
        $guidCmd = Get-Command -Name 'Get-PWDocumentsByGUIDs'
        $guidParams = @{ DocumentGUIDs = @($DocumentGuid) }
        if ($guidCmd.Parameters.ContainsKey('GetAttributes')) {
            $guidParams['GetAttributes'] = [bool]$includeAttrs
        }
        $byGuid = Get-PWDocumentsByGUIDs @guidParams -ErrorAction Stop | Select-Object -First 1
        if ($byGuid) { return $byGuid }
    }

    throw "Document not found: $FolderPath\$DocumentName"
}

function Get-PwDocumentWithAttributesForDiagnostics {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName
    )
    $searchParams = Get-PwSearchParams -FolderPath $FolderPath -DocumentName $DocumentName -IncludeAttributes:$true
    return (Get-PWDocumentsBySearch @searchParams -ErrorAction SilentlyContinue)
}

function Get-AttributeBagIssues {
    param([object]$Bag, [string]$BagName = 'Attributes')

    $issues = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $Bag) {
        return @{ present = $false; bagName = $BagName; nullOrBlankKeys = @(); issues = @() }
    }

    $present = $true
    $nullKeys = [System.Collections.Generic.List[string]]::new()

    if ($Bag -is [System.Collections.IDictionary]) {
        foreach ($k in @($Bag.Keys)) {
            if ($null -eq $k -or ([string]::IsNullOrWhiteSpace([string]$k))) {
                [void]$nullKeys.Add([string]$k)
            }
        }
    } elseif ($Bag.PSObject -and $Bag.PSObject.Properties) {
        foreach ($p in @($Bag.PSObject.Properties)) {
            if ($null -eq $p.Name -or [string]::IsNullOrWhiteSpace([string]$p.Name)) {
                [void]$nullKeys.Add('<blank property name>')
            }
        }
    } else {
        [void]$issues.Add('bag is not dictionary-like')
    }

    if ($nullKeys.Count -gt 0) {
        [void]$issues.Add('null or blank keys present')
    }

    return @{
        present = $present
        bagName = $BagName
        nullOrBlankKeys = @($nullKeys)
        issues = @($issues)
    }
}

function Write-PwDocumentAttributeDiagnostics {
    param(
        [Parameter(Mandatory)][string]$Label,
        [object]$Document
    )

    Write-Host ("Document object diagnostics ({0}):" -f $Label) -ForegroundColor Green
    if (-not $Document) {
        Write-Host '  document: <null>'
        return
    }

    $hasAttributes = $false
    $hasCustomAttributes = $false
    $attributes = $null
    $customAttributes = $null
    try {
        if ($Document.PSObject.Properties['Attributes']) {
            $hasAttributes = $true
            $attributes = $Document.Attributes
        }
    } catch { }
    try {
        if ($Document.PSObject.Properties['CustomAttributes']) {
            $hasCustomAttributes = $true
            $customAttributes = $Document.CustomAttributes
        }
    } catch { }

    Write-Host ("  has Attributes        : {0}" -f $hasAttributes)
    Write-Host ("  has CustomAttributes  : {0}" -f $hasCustomAttributes)

    foreach ($bagInfo in @(
        (Get-AttributeBagIssues -Bag $attributes -BagName 'Attributes')
        (Get-AttributeBagIssues -Bag $customAttributes -BagName 'CustomAttributes')
    )) {
        if (-not $bagInfo.present) { continue }
        $keySummary = if (@($bagInfo.nullOrBlankKeys).Count -gt 0) {
            (@($bagInfo.nullOrBlankKeys) -join ', ')
        } else {
            'none'
        }
        Write-Host ("  {0} null/blank keys : {1}" -f $bagInfo.bagName, $keySummary)
        if (@($bagInfo.issues).Count -gt 0) {
            Write-Host ("  {0} issues          : {1}" -f $bagInfo.bagName, (@($bagInfo.issues) -join '; '))
        }
    }
    Write-Host ''
}

function Write-SetPwDocumentStateDiagnostics {
    $cmd = Get-Command -Name 'Set-PWDocumentState' -ErrorAction SilentlyContinue
    if (-not $cmd) {
        $related = @(Get-Command -CommandType Function, Cmdlet -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match 'PW.*(State|Workflow)' } |
            Select-Object -ExpandProperty Name)
        throw "Set-PWDocumentState is unavailable. Related cmdlets: $($related -join ', ')"
    }

    Write-Host ''
    Write-Host 'Set-PWDocumentState syntax:' -ForegroundColor Green
    try {
        Get-Command Set-PWDocumentState -Syntax | ForEach-Object { Write-Host $_ }
    } catch {
        Write-Host ("  (syntax unavailable: {0})" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host 'Set-PWDocumentState parameters:' -ForegroundColor Green
    foreach ($p in @($cmd.Parameters.GetEnumerator() | Sort-Object Key)) {
        $req = if ($p.Value.Attributes.Mandatory) { 'Mandatory' } else { 'Optional' }
        Write-Host ("  -{0} [{1}] ({2})" -f $p.Key, $p.Value.ParameterType.FullName, $req)
    }
    Write-Host ''
    return $cmd
}

function Test-IsPwpsDabObjectShapeFailure {
    param([string]$ErrorMessage)
    if ([string]::IsNullOrWhiteSpace($ErrorMessage)) { return $false }
    return ($ErrorMessage -match '(?i)null key is not allowed in a hash literal')
}

function New-SetPwDocumentStateAttemptList {
    param(
        [Parameter(Mandatory)][System.Management.Automation.CommandInfo]$Command,
        [Parameter(Mandatory)][string]$TargetState
    )

    $params = $Command.Parameters
    if (-not $params.ContainsKey('InputDocuments') -or -not $params.ContainsKey('State')) {
        throw 'Installed Set-PWDocumentState does not expose -InputDocuments and -State.'
    }

    $base = "Set-PWDocumentState -InputDocuments `@(`$cleanDoc) -State '$TargetState'"
    $variants = @(
        @{ Label = 'InputDocuments + State'; CommandText = $base; Extra = @{} }
    )
    if ($params.ContainsKey('Force')) {
        $variants += @{ Label = 'InputDocuments + State + Force'; CommandText = "$base -Force"; Extra = @{ Force = $true } }
    }
    if ($params.ContainsKey('IgnoreStatus')) {
        $variants += @{ Label = 'InputDocuments + State + IgnoreStatus'; CommandText = "$base -IgnoreStatus"; Extra = @{ IgnoreStatus = $true } }
    }
    if ($params.ContainsKey('SkipIntermediateStates')) {
        $variants += @{ Label = 'InputDocuments + State + SkipIntermediateStates'; CommandText = "$base -SkipIntermediateStates"; Extra = @{ SkipIntermediateStates = $true } }
    }

    $attempts = [System.Collections.Generic.List[object]]::new()
    foreach ($v in @($variants)) {
        [void]$attempts.Add([pscustomobject]@{
            Label = [string]$v.Label
            CommandText = [string]$v.CommandText
            Extra = $v.Extra
        })
    }
    return @($attempts)
}

function Invoke-AllSetPwDocumentStateAttempts {
    param(
        [Parameter(Mandatory)][System.Management.Automation.CommandInfo]$Command,
        [Parameter(Mandatory)][object]$CleanDocument,
        [Parameter(Mandatory)][string]$TargetState,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = ''
    )

    $attemptList = New-SetPwDocumentStateAttemptList -Command $Command -TargetState $TargetState
    $results = [System.Collections.Generic.List[object]]::new()
    $anyNonThrowing = $false
    $anyObjectShapeFailure = $false
    $attemptNum = 0

    foreach ($attempt in $attemptList) {
        $attemptNum++
        Write-Host ''
        Write-Host ("--- Attempt {0}/{1}: {2} ---" -f $attemptNum, $attemptList.Count, $attempt.Label) -ForegroundColor Cyan
        Write-Host ("Command: {0}" -f $attempt.CommandText)

        $row = [ordered]@{
            attempt = $attemptNum
            label = $attempt.Label
            commandText = $attempt.CommandText
            threw = $false
            error = ''
            failureClass = ''
            readBackState = ''
            verified = $false
        }

        try {
            $invokeParams = @{
                InputDocuments = @($CleanDocument)
                State = $TargetState
                ErrorAction = 'Stop'
            }
            foreach ($k in @($attempt.Extra.Keys)) {
                $invokeParams[$k] = $attempt.Extra[$k]
            }
            & $Command @invokeParams | Out-Null

            $row.threw = $false
            $anyNonThrowing = $true
            Write-Host 'Threw: no' -ForegroundColor Green

            Write-Step 'Waiting 2 seconds before read-back...'
            Start-Sleep -Seconds 2
            $readDoc = Get-CleanPwDocumentForStateChange -FolderPath $FolderPath -DocumentName $DocumentName `
                -DocumentGuid $DocumentGuid -NoAttributes:$true
            $readBack = Get-DocumentWorkflowStateName -Document $readDoc -FolderPath $FolderPath `
                -DocumentName $DocumentName -DocumentGuid $DocumentGuid
            $row.readBackState = $readBack
            Write-Host ("Read-back WorkflowState: {0}" -f $readBack)

            if ($readBack -eq $TargetState) {
                $row.verified = $true
                Write-Host 'Attempt result: state change verified.' -ForegroundColor Green
                [void]$results.Add([pscustomobject]$row)
                return @{
                    results = @($results)
                    verified = $true
                    anyNonThrowing = $true
                    anyObjectShapeFailure = $anyObjectShapeFailure
                    winningAttempt = $attempt
                }
            }
            Write-Host 'Attempt result: command ran but read-back does not match target.' -ForegroundColor Yellow
        } catch {
            $row.threw = $true
            $row.error = [string]$_.Exception.Message
            if (Test-IsPwpsDabObjectShapeFailure -ErrorMessage $row.error) {
                $row.failureClass = 'pwps_dab_object_shape_failure'
                $anyObjectShapeFailure = $true
                Write-Host 'Threw: yes - PWPS_DAB object-shape/runtime failure (null key in hash literal).' -ForegroundColor Red
                Write-Host ("  Detail: {0}" -f $row.error) -ForegroundColor Red
                Write-Host '  Note: this is not evidence of a ProjectWise workflow/rules/permission rejection.' -ForegroundColor Yellow
            } else {
                Write-Host ("Threw: yes - {0}" -f $row.error) -ForegroundColor Red
            }
        }

        [void]$results.Add([pscustomobject]$row)
    }

    return @{
        results = @($results)
        verified = $false
        anyNonThrowing = $anyNonThrowing
        anyObjectShapeFailure = $anyObjectShapeFailure
        winningAttempt = $null
    }
}

function Write-ResolvedDocumentInfo {
    param(
        [string]$DocumentName,
        [string]$DocumentGuid,
        [string]$DocumentId,
        [string]$WorkflowName,
        [string]$CurrentState,
        [string]$TargetState
    )
    Write-Host ''
    Write-Host 'Resolved document info:' -ForegroundColor Green
    Write-Host ("  name         : {0}" -f $DocumentName)
    Write-Host ("  documentGUID : {0}" -f $DocumentGuid)
    Write-Host ("  documentID   : {0}" -f $DocumentId)
    Write-Host ("  workflowName : {0}" -f $WorkflowName)
    Write-Host ("  currentState : {0}" -f $CurrentState)
    Write-Host ("  targetState  : {0}" -f $TargetState)
    Write-Host ''
}

Write-Step 'Importing PWPS_DAB...'
Import-Module pwps_dab -Force -ErrorAction Stop

if (-not (Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction SilentlyContinue)) {
    throw 'Get-PWDocumentsBySearch is unavailable after importing pwps_dab.'
}

$credPath = 'C:\PW_QC_LOCAL\pw_cred.txt'
$skipConnect = $WhatIfOnly.IsPresent

if (-not $skipConnect) {
    if ($UseStoredCredential -or (Test-Path -LiteralPath $credPath)) {
        $cred = Get-PWCredFromKeyValueFile -Path $credPath
        Write-Step ("Connecting to {0} as {1}" -f $Datasource, $cred.UserName)
        Open-PWConnection -DatasourceName $Datasource -UserName $cred.UserName -Password $cred.Password | Out-Null
    } else {
        Write-Step ("Connecting to {0} (interactive credential)" -f $Datasource)
        Open-PWConnection -DatasourceName $Datasource | Out-Null
    }
}

try {
    if ($WhatIfOnly) {
        Write-Host 'WhatIfOnly: would connect, resolve document, and probe Set-PWDocumentState.' -ForegroundColor Yellow
        exit 0
    }

    Write-Step ("Resolving document: {0}\{1}" -f $FolderPath, $DocumentName)
    Write-Host ("NoAttributesForStateChange: {0}" -f $useNoAttributesForStateChange)

    $cleanDoc = Get-CleanPwDocumentForStateChange -FolderPath $FolderPath -DocumentName $DocumentName `
        -NoAttributes:$useNoAttributesForStateChange
    $diagDoc = Get-PwDocumentWithAttributesForDiagnostics -FolderPath $FolderPath -DocumentName $DocumentName

    $docGuid = ''
    $docId = ''
    $workflowName = ''
    try { $docGuid = [string]$cleanDoc.DocumentGUID } catch { }
    try { if (-not $docGuid) { $docGuid = [string]$cleanDoc.DocumentGuid } } catch { }
    try { $docId = [string]$cleanDoc.DocumentID } catch { }
    try { if (-not $docId) { $docId = [string]$cleanDoc.DocumentId } } catch { }
    try { $workflowName = [string]$cleanDoc.WorkflowName } catch { }
    try { if (-not $workflowName) { $workflowName = [string]$cleanDoc.Workflow } } catch { }

    $currentState = Get-DocumentWorkflowStateName -Document $cleanDoc -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $docGuid

    Write-ResolvedDocumentInfo -DocumentName $DocumentName -DocumentGuid $docGuid -DocumentId $docId `
        -WorkflowName $workflowName -CurrentState $currentState -TargetState $TargetState

    $setStateCmd = Write-SetPwDocumentStateDiagnostics

    Write-PwDocumentAttributeDiagnostics -Label 'clean state-change object' -Document $cleanDoc
    if ($diagDoc) {
        Write-PwDocumentAttributeDiagnostics -Label 'attributes-loaded diagnostic object' -Document $diagDoc
    }

    if ($ResolveOnly) {
        Write-Host 'ResolveOnly: document resolved; Set-PWDocumentState syntax printed; no state change attempted.' -ForegroundColor Green
        exit 0
    }

    if ($currentState -eq $TargetState) {
        Write-Host 'Document is already at target state.' -ForegroundColor Green
        Write-Host 'CONCLUSION: state change verified (already at target).' -ForegroundColor Green
        exit 0
    }

    Write-Step 'Attempting Set-PWDocumentState with clean document object (no attribute bags)...'
    $run = Invoke-AllSetPwDocumentStateAttempts -Command $setStateCmd -CleanDocument $cleanDoc `
        -TargetState $TargetState -FolderPath $FolderPath -DocumentName $DocumentName -DocumentGuid $docGuid

    Write-Host ''
    Write-Host 'Attempt summary:' -ForegroundColor Cyan
    foreach ($r in @($run.results)) {
        $status = if ($r.verified) { 'verified' } elseif ($r.threw) { 'threw' } else { 'unverified' }
        Write-Host ("  [{0}] {1} - {2}" -f $r.attempt, $r.label, $status)
        if ($r.threw) {
            if ($r.failureClass) { Write-Host ("         class: {0}" -f $r.failureClass) }
            Write-Host ("         error: {0}" -f $r.error)
        } elseif (-not [string]::IsNullOrWhiteSpace($r.readBackState)) {
            Write-Host ("         read-back: {0}" -f $r.readBackState)
        }
    }

    if ($run.verified) {
        Write-Host ''
        Write-Host ("CONCLUSION: state change verified via {0}." -f $run.winningAttempt.Label) -ForegroundColor Green
        exit 0
    }

    if ($run.anyNonThrowing) {
        Write-Host ''
        Write-Host 'CONCLUSION: command ran but state did not persist (read-back did not match target).' -ForegroundColor Yellow
        Write-Host 'Possible causes: invalid workflow transition, insufficient permission, workflow rules rejection, or PW state write did not persist.' -ForegroundColor Yellow
        exit 2
    }

    Write-Host ''
    if ($run.anyObjectShapeFailure) {
        Write-Host 'CONCLUSION: PWPS_DAB object-shape/runtime failure - all attempts threw before workflow outcome could be determined.' -ForegroundColor Red
        Write-Host 'Use a clean document object without attribute bags (default -NoAttributesForStateChange). This is not a ProjectWise workflow rejection.' -ForegroundColor Yellow
    } else {
        Write-Host 'CONCLUSION: all Set-PWDocumentState attempts threw hard errors.' -ForegroundColor Red
    }
    exit 1
}
catch {
    Write-Host ("ERROR: {0}" -f $_.Exception.Message) -ForegroundColor Red
    if ($_.ScriptStackTrace) { Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed }
    exit 1
}
finally {
    if (-not $WhatIfOnly) {
        try { Close-PWConnection -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
}
