# test_project_id_from_workarea.ps1
# Run with PW connected. Tests ways to extract Project ID from work area type.
# Usage: .\test_project_id_from_workarea.ps1 -FolderPath "AZDOT 2024\SomeProject\CADD\Sheets"
#
param(
  [Parameter(Mandatory = $true)]
  [string] $FolderPath,
  [string] $DatasourceName = "typsa-us-pw.bentley.com:typsa-us-pw-03"
)

$ErrorActionPreference = "Stop"
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
  & powershell.exe -MTA -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -FolderPath $FolderPath -DatasourceName $DatasourceName
  exit $LASTEXITCODE
}

Import-Module pwps_dab -ErrorAction SilentlyContinue
# Assume already connected; or: Open-PWConnection -DatasourceName $DatasourceName ...

$path = $FolderPath -replace '^Documents\\', ''

Write-Host "=== Testing Project ID extraction for: $path ===" -ForegroundColor Cyan

# --- Method 1: Get-PWRichProjectScalar (direct scalar from work area property) ---
Write-Host "`n--- Method 1: Get-PWRichProjectScalar ---" -ForegroundColor Yellow
foreach ($prop in @('ProjectID', 'PROJECT_WorkAreaNumber', 'WorkAreaNumber', 'Project_Number')) {
  try {
    $val = Get-PWRichProjectScalar -FolderPath $path -KeyProjectProperty $prop -ErrorAction SilentlyContinue
    if ($val) { Write-Host "  $prop = $val" -ForegroundColor Green }
  } catch { Write-Host "  $prop : $_" -ForegroundColor Red }
}

# --- Method 2: Get-PWFolders + Get-PWRichProjectForFolder + Get-PWFolderPathAndProperties ---
Write-Host "`n--- Method 2: Get-PWFolders -> Get-PWRichProjectForFolder -> Get-PWFolderPathAndProperties ---" -ForegroundColor Yellow
try {
  $folder = Get-PWFolders -FolderPath $path -JustOne -ErrorAction Stop
  if ($folder) {
    $richProject = $folder | Get-PWRichProjectForFolder -ErrorAction SilentlyContinue
    if ($richProject) {
      $withProps = $richProject | Get-PWFolderPathAndProperties -ErrorAction SilentlyContinue
      if ($withProps) {
        Write-Host "  Rich project path: $($withProps.FullPath)"
        $withProps.PSObject.Properties | Where-Object { $_.Name -match 'Project|WorkArea|Number' } | ForEach-Object {
          Write-Host "  $($_.Name) = $($_.Value)"
        }
      }
    }
  }
} catch { Write-Host "  Error: $_" -ForegroundColor Red }

# --- Method 3: Get-PWWorkAreaTypeForRichProject (returns type definition) ---
Write-Host "`n--- Method 3: Get-PWWorkAreaTypeForRichProject ---" -ForegroundColor Yellow
try {
  $folder = Get-PWFolders -FolderPath $path -JustOne -ErrorAction Stop
  if ($folder) {
    $richProject = $folder | Get-PWRichProjectForFolder -ErrorAction SilentlyContinue
    if ($richProject) {
      $workAreaType = Get-PWWorkAreaTypeForRichProject -InputFolder $richProject -ErrorAction SilentlyContinue
      Write-Host "  Work area type: $workAreaType"
      if ($workAreaType) { $workAreaType | Format-List * }
    }
  }
} catch { Write-Host "  Error: $_" -ForegroundColor Red }

# --- Method 4: Get-PWRichProjectProperties (schema - ProjectType) ---
Write-Host "`n--- Method 4: Get-PWRichProjectProperties (schema for type) ---" -ForegroundColor Yellow
try {
  $folder = Get-PWFolders -FolderPath $path -JustOne -ErrorAction Stop
  if ($folder) {
    $richProject = $folder | Get-PWRichProjectForFolder -ErrorAction SilentlyContinue
    if ($richProject) {
      $typeObj = Get-PWWorkAreaTypeForRichProject -InputFolder $richProject -ErrorAction SilentlyContinue
      $typeName = if ($typeObj) { $typeObj.Name } elseif ($typeObj -is [string]) { $typeObj } else { $null }
      if ($typeName) {
        $schema = Get-PWRichProjectProperties -ProjectType $typeName -ErrorAction SilentlyContinue
        Write-Host "  Schema properties: $($schema | Out-String)"
      }
    }
  }
} catch { Write-Host "  Error: $_" -ForegroundColor Red }

Write-Host "`n=== Done ===" -ForegroundColor Cyan
