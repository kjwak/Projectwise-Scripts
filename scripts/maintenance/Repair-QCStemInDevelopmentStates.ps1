<#
.SYNOPSIS
Finds stem DGN/PDF documents in ProjectWise that are not in the production reference state and optionally resets them.

.DESCRIPTION
Scans one or more ProjectWise Sheets folders for stem documents (.dgn and non-lane .pdf files).
Lane PDFs (*-prod.pdf, *-chk.pdf, *-rev.pdf, legacy *-qc.pdf) and _StatusSet.pdf are ignored.
Documents with an empty workflow state are ignored.

Default target state is qcWorkflow.states.production (usually "In Development"). Default mode is
preview only; pass -ConfirmWrites to apply changes after you confirm the selection.

Folder paths accept pw:\datasource\Documents\..., Documents\..., or the pwps_dab form without
the Documents\ prefix (for example AZDOT 2024\...\CADD\Sheets).

.EXAMPLE
.\scripts\maintenance\Repair-QCStemInDevelopmentStates.ps1 `
  -FolderPath 'AZDOT 2024\AZFWY1704-FD02-SR202 - I-10 to SR101\CADD\Sheets' -Recursive

.EXAMPLE
.\scripts\maintenance\Repair-QCStemInDevelopmentStates.ps1 `
  -FolderPath 'AZDOT 2024\AZFWY1704-FD02-SR202 - I-10 to SR101\CADD\Sheets' `
  -Selection all -ConfirmWrites

.EXAMPLE
.\scripts\maintenance\Repair-QCStemInDevelopmentStates.ps1 -Interactive
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$FolderPath = '',
    [string]$TargetState = '',
    [string]$Selection = '',
    [string]$AppSettingsPath = '',
    [switch]$ConfirmWrites,
    [switch]$DryRun,
    [switch]$Recursive,
    [switch]$Interactive,
    [switch]$NoPrompt,
    [switch]$Pretty
)

# pwps_dab requires MTA; Cursor/VS Code terminals often use STA.
$scriptPath = $PSCommandPath
if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
if (-not $scriptPath) {
    $scriptPath = Join-Path $PSScriptRoot 'Repair-QCStemInDevelopmentStates.ps1'
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
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'Core\Core.Paths.psm1'
    'ProjectWise\PW.Connection.psm1'
    'ProjectWise\PW.Discovery.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'ConvertTo-PWCmdletFolderPath'
    'Get-PWDocumentsInFolder'
    'Set-PWDocumentWorkflowStateVerified'
) -Context 'Repair-QCStemInDevelopmentStates bootstrap'

foreach ($moduleName in @('pwps', 'pwps_dab')) {
    if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending | Select-Object -First 1
        if ($available) { Import-Module $available.Name -ErrorAction SilentlyContinue | Out-Null }
    }
}

function _RQSD-PauseIfInteractiveConsole {
    if ($Host.Name -ne 'ConsoleHost') { return }
    try {
        Write-Host ''
        Write-Host 'Press Enter to close this window...' -ForegroundColor Yellow
        $null = Read-Host
    } catch { }
}

function _RQSD-TestStemSheetDocumentName {
    param([string]$DocumentName)
    if ([string]::IsNullOrWhiteSpace($DocumentName)) { return $false }
    if ($DocumentName -match '(?i)^_StatusSet\.pdf$') { return $false }
    if ($DocumentName -match '(?i)\.dgn$') { return $true }
    if ($DocumentName -match '(?i)\.pdf$' -and $DocumentName -notmatch '(?i)-(prod|chk|rev|qc)\.pdf$') { return $true }
    return $false
}

function _RQSD-NormalizeStateLabel {
    param(
        [string]$StateName,
        [hashtable]$Config
    )
    $s = ([string]$StateName).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    if (Get-Command -Name 'Format-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try { return [string](Format-QCWorkflowStateName -StateName $s -Config $Config) } catch { }
    }
    return $s
}

function _RQSD-StatesEqual {
    param(
        [string]$Left,
        [string]$Right
    )
    return ([string]$Left).Trim().Equals(([string]$Right).Trim(), [StringComparison]::OrdinalIgnoreCase)
}

function _RQSD-ResolveInputFolderPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }

    $raw = $Path.Trim().TrimEnd('\')
    $raw = $raw -replace '/', '\'

    if ($raw -match '^(?i)pw:\\') {
        $idx = $raw.IndexOf('\Documents\', [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -ge 0) {
            $raw = $raw.Substring($idx + 1)
        }
    }

    $norm = Normalize-QCDocumentsFolderPath -Path $raw
    if (-not $norm.IsSuccess) { throw $norm.Message }
    return [string]$norm.Data.path
}

function _RQSD-GetPwApiFolderPath {
    param([string]$DocumentsFolderPath)
    $api = ConvertTo-PWCmdletFolderPath -InternalFolderPath $DocumentsFolderPath
    if ([string]::IsNullOrWhiteSpace($api)) { return $DocumentsFolderPath }
    return $api
}

function _RQSD-GetDocumentNameFromPwObject {
    param([object]$Document)
    foreach ($prop in @('Name', 'DocumentName', 'FileName')) {
        try {
            if ($Document.PSObject.Properties[$prop] -and $Document.$prop) {
                return [string]$Document.$prop
            }
        } catch { }
    }
    return ''
}

function _RQSD-GetDocumentGuidFromPwObject {
    param([object]$Document)
    foreach ($prop in @('DocumentGUID', 'DocumentGuid', 'GUID')) {
        try {
            if ($Document.PSObject.Properties[$prop] -and $Document.$prop) {
                return [string]$Document.$prop
            }
        } catch { }
    }
    return ''
}

function _RQSD-GetChildApiFolderPaths {
    param([string]$ParentApiPath)
    $children = @()
    try {
        $childFolders = @(Get-PWImmediateChildFolders -FolderPath $ParentApiPath)
        foreach ($child in $childFolders) {
            $childPath = $null
            foreach ($prop in @('FolderPath', 'Path')) {
                try {
                    if ($child.PSObject.Properties[$prop] -and $child.$prop) {
                        $childPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath ([string]$child.$prop)
                        break
                    }
                } catch { }
            }
            if ([string]::IsNullOrWhiteSpace($childPath)) {
                $name = _RQSD-GetDocumentNameFromPwObject -Document $child
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    $childPath = ($ParentApiPath.TrimEnd('\') + '\' + $name)
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($childPath)) {
                $children += $childPath
            }
        }
    } catch { }
    return @($children | Select-Object -Unique)
}

function _RQSD-GetFolderPathsToScan {
    param(
        [string]$RootApiPath,
        [bool]$IncludeSubfolders
    )
    $paths = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    function _AddPath([string]$p) {
        if ([string]::IsNullOrWhiteSpace($p)) { return }
        $norm = $p.Trim().TrimEnd('\')
        if ($seen.Add($norm)) { [void]$paths.Add($norm) }
    }

    _AddPath $RootApiPath
    if (-not $IncludeSubfolders) { return @($paths) }

    $queue = [System.Collections.Queue]::new()
    $queue.Enqueue($RootApiPath.Trim().TrimEnd('\'))
    while ($queue.Count -gt 0) {
        $current = [string]$queue.Dequeue()
        foreach ($child in @(_RQSD-GetChildApiFolderPaths -ParentApiPath $current)) {
            if ($seen.Add($child)) {
                [void]$paths.Add($child)
                $queue.Enqueue($child)
            }
        }
    }
    return @($paths)
}

function _RQSD-ParseSelectionInput {
    param(
        [string]$SelectionText,
        [object[]]$Candidates
    )
    $text = $SelectionText.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    if ($text -match '^(?i)(all|\*|a)$') { return @($Candidates) }

    $selected = [System.Collections.Generic.List[object]]::new()
    foreach ($part in ($text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if ($part -match '^(\d+)\s*-\s*(\d+)$') {
            $start = [int]$Matches[1]
            $end = [int]$Matches[2]
            if ($start -gt $end) { throw "Invalid range: $part" }
            for ($n = $start; $n -le $end; $n++) {
                if ($n -lt 1 -or $n -gt $Candidates.Count) {
                    throw "Selection out of range: $n (1-$($Candidates.Count))"
                }
                [void]$selected.Add($Candidates[$n - 1])
            }
            continue
        }
        if ($part -match '^\d+$') {
            $n = [int]$part
            if ($n -lt 1 -or $n -gt $Candidates.Count) {
                throw "Selection out of range: $n (1-$($Candidates.Count))"
            }
            [void]$selected.Add($Candidates[$n - 1])
            continue
        }
        throw "Unrecognized selection '$part'. Use all, numbers (1,3,5 or 1-3)."
    }

    return @($selected | Select-Object -Unique)
}

function _RQSD-PromptFolderPathFromDatabase {
    param([hashtable]$Config)
    if (-not (Get-Command -Name 'Invoke-QCDatabaseQuery' -ErrorAction SilentlyContinue)) {
        throw 'Database helpers are unavailable for interactive folder selection.'
    }
    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) {
        throw 'Database helpers are unavailable for interactive folder selection.'
    }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        throw 'database.enabled must be true for interactive folder selection.'
    }

    Initialize-QCDatabaseSchema -Config $Config | Out-Null
    $sql = @"
SELECT folder_path, COUNT_BIG(*) AS index_count
FROM sheet_index
WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> ''
GROUP BY folder_path
ORDER BY folder_path
"@
    $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql
    if (-not $res.IsSuccess) { throw $res.Message }
    $rows = @($res.Data.table.Rows)
    if ($rows.Count -eq 0) { throw 'No folder paths found in sheet_index.' }

    Write-Host ''
    Write-Host 'Available folder paths:' -ForegroundColor Cyan
    $paths = [System.Collections.Generic.List[string]]::new()
    $idx = 1
    foreach ($row in $rows) {
        $fp = [string]$row.folder_path
        $paths.Add($fp) | Out-Null
        Write-Host ("  {0,4}) {1}  (index rows: {2})" -f $idx, $fp, [long]$row.index_count)
        $idx++
    }
    Write-Host ''
    Write-Host 'Enter selection: number, numbers (1,3,5), range (1-3), or type a folder path.' -ForegroundColor DarkGray
    $raw = (Read-Host 'Folder selection').Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) { throw 'No folder selected.' }

    if ($raw -match '[\\/]' -or $raw -imatch '^documents') {
        return @(_RQSD-ResolveInputFolderPath -Path $raw)
    }
    if ($raw -match '^\d+$') {
        $n = [int]$raw
        if ($n -lt 1 -or $n -gt $paths.Count) { throw "Selection out of range: $n (1-$($paths.Count))" }
        return @([string]$paths[$n - 1])
    }
    throw "Unrecognized folder selection: $raw"
}

if ($DryRun.IsPresent -and $ConfirmWrites.IsPresent) {
    throw 'Use -DryRun (preview only) OR -ConfirmWrites (apply PW changes), not both.'
}

$doWrites = $ConfirmWrites.IsPresent
if (-not $DryRun.IsPresent -and -not $ConfirmWrites.IsPresent) {
    Write-Host 'Preview mode: pass -ConfirmWrites to apply changes after selection, or -DryRun to skip the apply prompt.' -ForegroundColor Yellow
    $DryRun = $true
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

$resolvedTarget = if (-not [string]::IsNullOrWhiteSpace($TargetState)) {
    [string]$TargetState
} elseif ($config.ContainsKey('qcWorkflow') -and $config.qcWorkflow -and $config.qcWorkflow.states -and $config.qcWorkflow.states.production) {
    [string]$config.qcWorkflow.states.production
} else {
    'In Development'
}
$resolvedTarget = _RQSD-NormalizeStateLabel -StateName $resolvedTarget -Config $config

$launchInteractive = $Interactive.IsPresent
if ([string]::IsNullOrWhiteSpace($FolderPath) -and -not $launchInteractive) {
    if ($Host.Name -eq 'ConsoleHost' -and -not $NoPrompt.IsPresent) {
        $launchInteractive = $true
    } else {
        throw 'FolderPath is required unless -Interactive is used.'
    }
}

$folderPaths = [System.Collections.Generic.List[string]]::new()
if ($launchInteractive -and [string]::IsNullOrWhiteSpace($FolderPath)) {
    foreach ($fp in @(_RQSD-PromptFolderPathFromDatabase -Config $config)) {
        [void]$folderPaths.Add($fp)
    }
} else {
    [void]$folderPaths.Add((_RQSD-ResolveInputFolderPath -Path $FolderPath))
}

$pwCfg = @{}
if ($config.ContainsKey('projectWise') -and $config.projectWise) {
    $rawPw = $config.projectWise
    if ($rawPw -is [hashtable]) { $pwCfg = $rawPw }
    elseif ($rawPw.PSObject) { foreach ($p in $rawPw.PSObject.Properties) { $pwCfg[$p.Name] = $p.Value } }
}
$credPath = if ($pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
$ds = if ($pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { '' }
if ([string]::IsNullOrWhiteSpace($ds)) { throw 'projectWise.datasourceName is required in appsettings.' }

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw $credRes.Message }
$connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $connRes.IsSuccess) { throw $connRes.Message }

$summary = [ordered]@{
    targetState          = $resolvedTarget
    folderPaths          = @($folderPaths)
    recursive            = [bool]$Recursive.IsPresent
    foldersScanned       = 0
    documentsScanned     = 0
    stemDocumentsScanned = 0
    candidatesFound      = 0
    selectedCount        = 0
    writesPlanned        = 0
    writesVerified       = 0
    writesFailed         = 0
    skippedEmptyState    = 0
    skippedAlreadyTarget = 0
    candidates           = @()
    results              = @()
    errors               = @()
}

try {
    $candidates = [System.Collections.Generic.List[object]]::new()

    foreach ($documentsFolderPath in @($folderPaths)) {
        $apiRoot = _RQSD-GetPwApiFolderPath -DocumentsFolderPath $documentsFolderPath
        if (-not (Test-PWFolderResolvable -FolderPath $documentsFolderPath)) {
            $summary.errors += "Folder not found in ProjectWise: $documentsFolderPath (api: $apiRoot)"
            continue
        }

        $apiPaths = @(_RQSD-GetFolderPathsToScan -RootApiPath $apiRoot -IncludeSubfolders:$Recursive.IsPresent)
        foreach ($apiPath in $apiPaths) {
            $summary.foldersScanned++

            $docsFolderPath = $documentsFolderPath
            if ($apiPath -ne $apiRoot) {
                $suffix = $apiPath.Substring($apiRoot.Length).TrimStart('\')
                if ($suffix) {
                    $docsFolderPath = ($documentsFolderPath.TrimEnd('\') + '\' + $suffix)
                }
            }

            $docs = @()
            try {
                $docs = @(Get-PWDocumentsInFolder -FolderPath $apiPath)
            } catch {
                $summary.errors += "Failed to list documents in $apiPath : $($_.Exception.Message)"
                continue
            }

            $summary.documentsScanned += $docs.Count
            foreach ($doc in $docs) {
                $name = _RQSD-GetDocumentNameFromPwObject -Document $doc
                if (-not (_RQSD-TestStemSheetDocumentName -DocumentName $name)) { continue }
                $summary.stemDocumentsScanned++

                $guid = _RQSD-GetDocumentGuidFromPwObject -Document $doc
                if ([string]::IsNullOrWhiteSpace($guid)) { continue }

                $candidates.Add([pscustomobject]@{
                    folderPathDocuments = $docsFolderPath
                    folderPathApi       = $apiPath
                    documentName        = $name
                    documentGuid        = $guid
                    currentState        = ''
                }) | Out-Null
            }
        }
    }

    if ($candidates.Count -eq 0) {
        Write-Host 'No stem DGN/PDF documents found in the scanned folders.' -ForegroundColor Yellow
        if ($Pretty) { $summary | ConvertTo-Json -Depth 8 }
        exit 0
    }

    $guids = @($candidates | ForEach-Object { [string]$_.documentGuid })
    $stateByGuid = @{}
    try {
        $stateByGuid = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $guids
    } catch {
        $summary.errors += "Workflow state batch read failed: $($_.Exception.Message)"
    }

    $filtered = [System.Collections.Generic.List[object]]::new()
    foreach ($row in @($candidates)) {
        $key = ([string]$row.documentGuid).ToLowerInvariant()
        $state = if ($stateByGuid.ContainsKey($key)) { [string]$stateByGuid[$key] } else { '' }
        $state = _RQSD-NormalizeStateLabel -StateName $state -Config $config
        $row.currentState = $state

        if ([string]::IsNullOrWhiteSpace($state)) {
            $summary.skippedEmptyState++
            continue
        }
        if (_RQSD-StatesEqual -Left $state -Right $resolvedTarget) {
            $summary.skippedAlreadyTarget++
            continue
        }

        [void]$filtered.Add($row)
    }

    $summary.candidatesFound = $filtered.Count
    $summary.candidates = @($filtered | ForEach-Object {
        [ordered]@{
            folderPath   = $_.folderPathDocuments
            documentName = $_.documentName
            documentGuid = $_.documentGuid
            currentState = $_.currentState
            targetState  = $resolvedTarget
        }
    })

    if ($filtered.Count -eq 0) {
        Write-Host "No stem documents need a state change to '$resolvedTarget' (empty states ignored)." -ForegroundColor Green
        if ($Pretty) { $summary | ConvertTo-Json -Depth 8 }
        exit 0
    }

    Write-Host ''
    Write-Host "Stem documents not in '$resolvedTarget':" -ForegroundColor Cyan
    $idx = 1
    foreach ($row in @($filtered)) {
        Write-Host ("  {0,4}) {1}\{2}  [{3}]" -f $idx, $row.folderPathDocuments, $row.documentName, $row.currentState)
        $idx++
    }
    Write-Host ''

    $selected = @()
    if (-not [string]::IsNullOrWhiteSpace($Selection)) {
        $selected = @(_RQSD-ParseSelectionInput -SelectionText $Selection -Candidates @($filtered))
    } elseif ($NoPrompt.IsPresent) {
        $selected = @($filtered)
    } else {
        Write-Host 'Enter selection to reset: all, numbers (1,3,5), range (1-3), or blank to cancel.' -ForegroundColor DarkGray
        $raw = (Read-Host 'Document selection').Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) {
            Write-Host 'Cancelled; no ProjectWise changes made.' -ForegroundColor Yellow
            if ($Pretty) { $summary | ConvertTo-Json -Depth 8 }
            exit 0
        }
        $selected = @(_RQSD-ParseSelectionInput -SelectionText $raw -Candidates @($filtered))
    }

    if ($selected.Count -eq 0) {
        Write-Host 'No documents selected.' -ForegroundColor Yellow
        if ($Pretty) { $summary | ConvertTo-Json -Depth 8 }
        exit 0
    }

    $summary.selectedCount = $selected.Count
    Write-Host ''
    Write-Host ("Selected $($selected.Count) document(s) for -> $resolvedTarget") -ForegroundColor Cyan

    if (-not $doWrites) {
        Write-Host 'Preview only. Re-run with -ConfirmWrites to apply the selected changes.' -ForegroundColor Yellow
        if ($Pretty) { $summary | ConvertTo-Json -Depth 8 }
        exit 0
    }

    if (-not $NoPrompt.IsPresent) {
        $confirm = (Read-Host 'Apply these ProjectWise state changes? [y/N]').Trim()
        if ($confirm -notmatch '^(?i)y|yes$') {
            Write-Host 'Cancelled; no ProjectWise changes made.' -ForegroundColor Yellow
            exit 0
        }
    }

    foreach ($row in @($selected)) {
        $target = "$($row.folderPathDocuments)\$($row.documentName)"
        $summary.writesPlanned++
        $resultRow = [ordered]@{
            folderPath   = $row.folderPathDocuments
            documentName = $row.documentName
            documentGuid = $row.documentGuid
            fromState    = $row.currentState
            targetState  = $resolvedTarget
            verified     = $false
            error        = ''
        }

        if ($PSCmdlet.ShouldProcess($target, "Set workflow state to '$resolvedTarget'")) {
            try {
                $write = Set-PWDocumentWorkflowStateVerified -Config $config `
                    -FolderPath $row.folderPathApi `
                    -DocumentName $row.documentName `
                    -DocumentGuid $row.documentGuid `
                    -TargetState $resolvedTarget `
                    -DryRun:$false `
                    -IsLaneAuthority:$false `
                    -WriteScope 'stem'
                $resultRow.verified = [bool]$write.verified
                $resultRow.readBackState = [string]$write.readBackState
                if ($write.verified) {
                    $summary.writesVerified++
                    Write-Host ("  OK  {0}  {1} -> {2}" -f $row.documentName, $row.currentState, $write.readBackState) -ForegroundColor Green
                } else {
                    $summary.writesFailed++
                    $resultRow.error = if ($write.error) { [string]$write.error } else { 'read_back_not_verified' }
                    Write-Host ("  FAIL {0}  {1} -> (read-back: {2}) {3}" -f $row.documentName, $row.currentState, $write.readBackState, $resultRow.error) -ForegroundColor Red
                }
            } catch {
                $summary.writesFailed++
                $resultRow.error = $_.Exception.Message
                $summary.errors += "Write failed for $($row.documentName): $($_.Exception.Message)"
                Write-Host ("  FAIL {0}  {1}" -f $row.documentName, $_.Exception.Message) -ForegroundColor Red
            }
        }

        $summary.results += ,([pscustomobject]$resultRow)
    }
} finally {
    Disconnect-PW | Out-Null
}

Write-Host ''
Write-Host ("Done. planned=$($summary.writesPlanned) verified=$($summary.writesVerified) failed=$($summary.writesFailed) skippedEmpty=$($summary.skippedEmptyState) alreadyTarget=$($summary.skippedAlreadyTarget)") -ForegroundColor Cyan

if ($Pretty) {
    $summary | ConvertTo-Json -Depth 8
}

_RQSD-PauseIfInteractiveConsole
if ($summary.writesFailed -gt 0) { exit 2 }
if ($summary.writesPlanned -gt 0 -and $summary.writesVerified -lt $summary.writesPlanned) { exit 2 }
exit 0
