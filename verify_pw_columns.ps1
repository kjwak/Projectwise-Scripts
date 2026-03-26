# Verify PW document properties - run against a real Sheets folder to see column names
# Usage: .\verify_pw_columns.ps1                    # discover first Sheets folder and test
#        .\verify_pw_columns.ps1 -FolderPath "..."  # test specific folder
param(
  [string]$FolderPath,
  [string]$WatchUnderRoot = "Documents\AZDOT 2024",
  [string]$SheetsPathFromProject = "CADD\Sheets",
  [string]$DatasourceName = "typsa-us-pw.bentley.com:typsa-us-pw-03"
)

#Requires -Modules PWPS_DAB
$ErrorActionPreference = "Stop"
$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt'

function Get-PwCredential {
  if (-not (Test-Path -LiteralPath $CredentialPath)) {
    throw "Credential file not found: $CredentialPath"
  }
  $lines = Get-Content -LiteralPath $CredentialPath -ErrorAction Stop
  $uLine = $lines | Where-Object { $_ -match '^\s*username\s*=' } | Select-Object -First 1
  $pLine = $lines | Where-Object { $_ -match '^\s*password\s*=' } | Select-Object -First 1
  if (-not $uLine -or -not $pLine) { throw "Invalid format in $CredentialPath" }
  $user = ($uLine -split '=', 2)[1].Trim()
  $pass = ($pLine -split '=', 2)[1].Trim()
  $sec = ConvertTo-SecureString $pass -AsPlainText -Force
  return [pscredential]::new($user, $sec)
}

function Connect-PW([string]$dsName) {
  $cred = Get-PwCredential
  try {
    Open-PWConnection -DatasourceName $dsName -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null
  } catch {
    if ($_.Exception.Message -match 'connection is already open') {
      Close-PWConnection -ErrorAction SilentlyContinue
      Open-PWConnection -DatasourceName $dsName -UserName $cred.UserName -Password $cred.Password | Out-Null
    } else { throw }
  }
}

function Get-SheetsFoldersUnderRoot([string]$rootPath, [string]$sheetsSuffix, [string]$ds) {
  $rootPath = $rootPath -replace '^Documents\\', ''
  $childNames = @()
  try {
    $view = Get-PWFolderView -FolderPath $rootPath -ErrorAction Stop
    if ($view.Children) {
      foreach ($c in $view.Children) {
        $name = $c.Name
        if (-not $name -and $c.FolderPath) { $name = [System.IO.Path]::GetFileName($c.FolderPath.TrimEnd('\')) }
        if ($name) { $childNames += $name }
      }
    }
    if ($childNames.Count -eq 0 -and $view.Folders) {
      foreach ($f in $view.Folders) { if ($f.Name) { $childNames += $f.Name } }
    }
  } catch {
    $children = Get-PWFoldersImmediateChildren -FolderPath $rootPath -ErrorAction Stop
    foreach ($c in @($children)) {
      $name = $c.Name
      if (-not $name -and $c.FolderPath) { $name = [System.IO.Path]::GetFileName($c.FolderPath.TrimEnd('\')) }
      if ($name) { $childNames += $name }
    }
  }
  $suffix = $sheetsSuffix.Trim().TrimStart('\')
  $list = @()
  foreach ($name in $childNames) {
    $folderPath = $rootPath.TrimEnd('\') + '\' + $name + '\' + $suffix
    $list += $folderPath
  }
  return $list
}

$dateCols = @("Name", "DocumentID", "VersionModifiedDate", "Version Modified Date", "FileUpdatedDate")

Write-Host "1. Parameter check: -ColumnsToReturn (not -ReturnColumns)"
$params = (Get-Command Get-PWDocumentsBySearchWithReturnColumns).Parameters.Keys
Write-Host "   ColumnsToReturn: $($params -contains 'ColumnsToReturn')"
Write-Host ""

# Discover folder if not provided
if (-not $FolderPath) {
  Write-Host "2. Discovering first Sheets folder from $WatchUnderRoot..."
  Connect-PW $DatasourceName
  $folders = Get-SheetsFoldersUnderRoot -rootPath $WatchUnderRoot -sheetsSuffix $SheetsPathFromProject -ds $DatasourceName
  if ($folders.Count -eq 0) {
    Write-Host "   No Sheets folders found. Try -FolderPath explicitly."
    exit 1
  }
  $FolderPath = $folders[0]
  Write-Host "   Using: $FolderPath"
}

# Try Documents\ prefix if path doesn't have it
$tryPaths = @($FolderPath)
if ($FolderPath -notmatch '^Documents\\') {
  $tryPaths = @("Documents\$FolderPath", $FolderPath)
}

Write-Host ""
Write-Host "3. Fetching one PDF with ColumnsToReturn..."
$docs = $null
foreach ($tryPath in $tryPaths) {
  try {
    $docs = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $tryPath -JustThisFolder -FileName "%.pdf" -ColumnsToReturn $dateCols -ErrorAction Stop | Select-Object -First 1
    if ($docs) { $FolderPath = $tryPath; break }
  } catch { }
}

if (-not $docs) {
  Write-Host "   No PDFs found. Tried: $($tryPaths -join ', ')"
  exit 1
}

$d = $docs[0]
Write-Host "   Document: $($d.Name)"
Write-Host ""
Write-Host "4. All properties on returned object:"
$d.PSObject.Properties | ForEach-Object { Write-Host "     $($_.Name) = $($_.Value)" }
Write-Host ""
Write-Host "5. Date column check (for Get-DocLastModified):"
$foundDateCol = $null
foreach ($p in @("VersionModifiedDate", "Version Modified Date", "FileUpdatedDate", "Date Last Saved", "File Updated")) {
  $v = $null
  if ($d.PSObject.Properties[$p]) { $v = $d.PSObject.Properties[$p].Value }
  if (-not $v) { $v = $d.$p }
  $status = if ($v) { "OK: $v" } else { "null" }
  Write-Host "   $p : $status"
  if ($v -and -not $foundDateCol) { $foundDateCol = $p }
}

# Check for any property that looks like a date
Write-Host ""
Write-Host "6. Other properties that might be dates:"
$d.PSObject.Properties | Where-Object {
  $_.Name -match 'date|time|modified|updated|saved' -and $_.Value -and $_.Value -is [DateTime]
} | ForEach-Object { Write-Host "     $($_.Name) = $($_.Value)" }

Write-Host ""
Write-Host "7. Null-date fallback: when statusSetLastModified OR docLastMod is null, script exports (safe)."
