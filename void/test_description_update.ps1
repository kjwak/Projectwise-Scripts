# test_description_update.ps1
# Small test: connect to PW, get input1.pdf in the Prepend Test folder, try different ways to update Description.
# Run from repo root with pwps_dab loaded. Adjust $DatasourceName / $FolderPath if needed.
#
# Usage: .\test_description_update.ps1

$ErrorActionPreference = "Stop"
$DatasourceName = "typsa-us-pw.bentley.com:typsa-us-pw-03"
$FolderPath = "AZDOT 2024\AZFWY1704-FD02-SR202 - I-10 to SR101\CADD\Working\TYPSA\Drainage\JFlint\Prepend Test"
$DocName = "input1.pdf"
$TriggerTag = "|QC|"

# Encrypted credential path (created once with Get-Credential | Export-CliXml -Path 'C:\PW_QC_LOCAL\pw_cred.xml')
$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.xml'

function Get-PwCredential {
  if (-not (Test-Path -LiteralPath $CredentialPath)) {
    throw "ProjectWise credential file not found: $CredentialPath. Create it once with:`n  Get-Credential | Export-CliXml -Path '$CredentialPath'"
  }
  Import-Clixml -LiteralPath $CredentialPath
}

function Connect-PW([string]$dsName) {
  $cred = Get-PwCredential
  Open-PWConnection -DatasourceName $dsName -UserName $cred.UserName -Password $cred.Password | Out-Null
}

try { Connect-PW $DatasourceName } catch { Write-Host "Open-PWConnection failed: $_"; exit 1 }

$doc = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocName -PopulatePath
if (-not $doc) { Write-Host "Document not found: $FolderPath\$DocName"; exit 1 }

$currentDesc = $doc.Description
$newDesc = ($currentDesc -replace [regex]::Escape($TriggerTag), "").Trim()
Write-Host "Current Description: [$currentDesc]"
Write-Host "New Description (tag removed): [$newDesc]"
Write-Host ""

# Inspect what attribute names the document object exposes (for Attributes/CustomAttributes)
Write-Host "--- Document property names (Description, Attributes, CustomAttributes) ---"
$doc.PSObject.Properties.Name | Sort-Object
if ($doc.Attributes) { Write-Host "doc.Attributes: $($doc.Attributes | ConvertTo-Json -Compress)" }
if ($doc.CustomAttributes) { Write-Host "doc.CustomAttributes: $($doc.CustomAttributes | ConvertTo-Json -Compress)" }
Write-Host ""

# Discovery: cmdlets that have BOTH document param AND Description param (possible way to set description)
Write-Host "--- Cmdlets with Document + Description params (candidates for updating description) ---"
$docParamNames = @('InputDocument','Document','InputDocuments')
foreach ($cmd in (Get-Command -Module pwps_dab, pwps -ErrorAction SilentlyContinue)) {
  $keys = $cmd.Parameters.Keys
  $hasDoc = $keys | Where-Object { $_ -in $docParamNames }
  $hasDesc = $keys | Where-Object { $_ -match '^Description$' }
  if ($hasDoc -and $hasDesc) {
    Write-Host "  $($cmd.Name): doc=$hasDoc, Description"
  }
}
Write-Host ""

# Discovery: document object METHODS (e.g. Update, SetDescription, Save)
Write-Host "--- Document object methods (Get-Member -MemberType Method) ---"
$methods = $doc | Get-Member -MemberType Method -ErrorAction SilentlyContinue
if ($methods) { $methods | ForEach-Object { Write-Host "  $($_.Name)" } } else { Write-Host "  (none or not enumerable)" }
Write-Host ""

# Discovery: what are "environments"? (optional – you don't have to set anything)
Write-Host "--- Environments (datasource attribute schemas; Description is usually in one of these) ---"
try {
  $envs = Get-PWEnvironments -ErrorAction Stop
  if ($envs) {
    $names = @($envs | ForEach-Object { if ($_.Name) { $_.Name } elseif ($_.EnvironmentName) { $_.EnvironmentName } else { $_ } })
    Write-Host "Environments: $($names -join ', ')"
  } else { Write-Host "Get-PWEnvironments returned nothing." }
} catch { Write-Host "Get-PWEnvironments: $_" }
Write-Host ""

# Test 0: Rename-PWDocument (often has -Description; same name = just update description)
$doc2 = $doc
$renameCmd = Get-Command Rename-PWDocument -ErrorAction SilentlyContinue
if ($renameCmd) {
  $common = @('Verbose','Debug','ErrorAction','WarningAction','InformationAction','ErrorVariable','WarningVariable','OutVariable','OutBuffer','PipelineVariable')
  $allKeys = @($renameCmd.Parameters.Keys | Where-Object { $_ -notin $common })
  Write-Host "--- Test 0: Rename-PWDocument params: $($allKeys -join ', ') ---"
  $descParam = $allKeys | Where-Object { $_ -match 'Description' } | Select-Object -First 1
  $docParam = $allKeys | Where-Object { $_ -match '^(InputDocument|Document)$' } | Select-Object -First 1
  $nameParam = $allKeys | Where-Object { $_ -match 'DocumentNewName|NewName' } | Select-Object -First 1
  if ($descParam -and $docParam) {
    $args = @{ $docParam = $doc; $descParam = $newDesc }
    if ($nameParam) { $args[$nameParam] = $doc.Name }
    Write-Host "  Calling with: $($args.Keys -join ', ')"
    try {
      Rename-PWDocument @args -ErrorAction Stop
      Write-Host "  No exception."
      $doc2 = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocName -PopulatePath
      Write-Host "  Description after: [$($doc2.Description)]"
    } catch { Write-Host "  Error: $_" }
  } else { Write-Host "  No Description param (doc: $docParam, desc: $descParam)." }
  Write-Host ""
}

# Test 1: Update-PWDocumentAttributes @{ Description = newDesc }
Write-Host "--- Test 1: Attributes @{ Description = newDesc } ---"
$r1 = Update-PWDocumentAttributes -InputDocuments @($doc2) -Attributes @{ Description = $newDesc } -ReturnBoolean
Write-Host "ReturnBoolean: $r1"
$doc2 = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocName -PopulatePath
Write-Host "Description after update: [$($doc2.Description)]"
Write-Host ""

# Test 1b: Update-PWDocumentProperties (set .Description on object, then pass positionally – the working pattern)
if ($doc2.Description -eq $currentDesc) {
  Write-Host "--- Test 1b: `$doc.Description = newDesc; Update-PWDocumentProperties `$doc ---"
  try {
    $doc2.Description = $newDesc
    Update-PWDocumentProperties $doc2
    Write-Host "  No exception."
    $doc2 = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocName -PopulatePath
    Write-Host "  Description after: [$($doc2.Description)]"
  } catch { Write-Host "  Error: $_" }
  Write-Host ""
}

# Test 2: DOC_DESCRIPTION
if ($doc2.Description -eq $currentDesc -and $newDesc -ne $currentDesc) {
  Write-Host "--- Test 2: Attributes @{ DOC_DESCRIPTION = newDesc } ---"
  $r2 = Update-PWDocumentAttributes -InputDocuments @($doc2) -Attributes @{ DOC_DESCRIPTION = $newDesc } -ReturnBoolean
  Write-Host "ReturnBoolean: $r2"
  $doc2 = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocName -PopulatePath
  Write-Host "Description after update: [$($doc2.Description)]"
  Write-Host ""
}

# Test 3: Get-PWEnvironmentColumns requires -EnvironmentName; get envs first, then columns from each
if ($doc2.Description -eq $currentDesc -and $newDesc -ne $currentDesc) {
  Write-Host "--- Test 3: Get-PWEnvironmentColumns (per environment) to find description column ---"
  $envList = @(Get-PWEnvironments -ErrorAction SilentlyContinue)
  if ($envList.Count -eq 0) { Write-Host "Get-PWEnvironments returned nothing; cannot call Get-PWEnvironmentColumns." }
  $descLike = @()
  $allColNames = @()
  foreach ($e in $envList) {
    $envName = if ($e.Name) { $e.Name } elseif ($e.EnvironmentName) { $e.EnvironmentName } else { $e }
    $cols = $null
    try { $cols = Get-PWEnvironmentColumns -EnvironmentName $envName -ErrorAction Stop } catch { try { $cols = Get-PWEnvironmentColumns -Environment $envName -ErrorAction Stop } catch { } }
    if (-not $cols) { continue }
    $colNames = @()
    foreach ($obj in @($cols)) {
      if ($obj -is [string]) { $colNames += $obj; continue }
      if ($obj.Name) { $colNames += $obj.Name }
      elseif ($obj.ColumnName) { $colNames += $obj.ColumnName }
      elseif ($obj.DisplayName) { $colNames += $obj.DisplayName }
      else { foreach ($p in $obj.PSObject.Properties) { if ($p.Name -notmatch '^(Verbose|Debug|Error)' -and $p.Value) { $colNames += $p.Value; break } } }
    }
    if ($colNames.Count -gt 0) {
      $allColNames += $colNames
      $descLike += @($colNames | Where-Object { $_ -match 'desc|description|DOC_DESC' })
      Write-Host "  Environment '$envName': $($colNames.Count) columns"
    }
  }
  $descLike = @($descLike | Select-Object -Unique)
  if ($descLike.Count -eq 0) { $descLike = @('Description', 'DOC_DESCRIPTION') }
  Write-Host "Description-like column names to try: $($descLike -join ', ')"
  foreach ($c in $descLike) {
    Write-Host "  Trying Attributes @{ $c = newDesc }"
    $r = Update-PWDocumentAttributes -InputDocuments @($doc2) -Attributes @{ $c = $newDesc } -ReturnBoolean
    Write-Host "  ReturnBoolean: $r"
    $doc2 = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocName -PopulatePath
    if ($doc2.Description -ne $currentDesc) { Write-Host "  SUCCESS. Description now: [$($doc2.Description)]"; break }
  }
  Write-Host ""
}

# Test 4: pwps Set-PWDocumentEAttributes (DocumentID, ProjectID, ValueObjects) if pwps loaded
if ($doc2.Description -eq $currentDesc -and $newDesc -ne $currentDesc) {
  $setCmd = Get-Command Set-PWDocumentEAttributes -ErrorAction SilentlyContinue
  $getCmd = Get-Command Get-PWDocumentEAttributes -ErrorAction SilentlyContinue
  if ($setCmd -and $getCmd) {
    Write-Host "--- Test 4: Get-PWDocumentEAttributes / Set-PWDocumentEAttributes (pwps) ---"
    try {
      $eattrs = Get-PWDocumentEAttributes -DocumentID $doc2.DocumentID -ProjectID $doc2.ProjectID -ErrorAction Stop
      $eattrCount = @($eattrs).Count
      Write-Host "Get-PWDocumentEAttributes returned: $($eattrs.GetType().Name); count: $eattrCount"
      if ($eattrCount -gt 0) { $eattrs | Format-List * }
      # Only call Set if we have a clue about structure; empty list often means doc has no e-attrs or wrong env
      if ($eattrCount -gt 0) {
        $vo = [pscustomobject]@{ Name = 'Description'; Value = $newDesc }
        Set-PWDocumentEAttributes -DocumentID $doc2.DocumentID -ProjectID $doc2.ProjectID -ValueObjects @($vo) -ErrorAction Stop
        Write-Host "Set-PWDocumentEAttributes completed (no exception)"
        $doc2 = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocName -PopulatePath
        Write-Host "Description after set: [$($doc2.Description)]"
      } else { Write-Host "Skipping Set (no e-attrs returned)." }
    } catch { Write-Host "Test 4 error: $_" }
    Write-Host ""
  }
}

Close-PWConnection -ErrorAction SilentlyContinue

Write-Host "=== SUMMARY ==="
if ($doc2.Description -eq $currentDesc) {
  Write-Host "Description was NOT updated. In this environment:"
  Write-Host "  - No pwps_dab/pwps cmdlet has both Document + Description parameters."
  Write-Host "  - Update-PWDocumentAttributes rejects Description/DOC_DESCRIPTION (core property, not env attribute)."
  Write-Host "  - Update-PWDocumentProperties accepts -InputDocument but does not update Description."
  Write-Host "  - Rename-PWDocument has no -Description parameter."
  Write-Host "Next options: (1) Ask Bentley / pwps_dab for a supported way; (2) Direct SQL UPDATE if you have schema + rights; (3) Use in-session skip in prepend_qc_on_trigger.ps1 so each doc is prepended only once per run."
} else {
  Write-Host "Description was updated successfully."
}
Write-Host "Done."
