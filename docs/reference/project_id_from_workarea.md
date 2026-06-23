# Extracting Project ID from Work Area Type

## Summary

Project ID is stored as a **work area property** on the Rich Project (work area) root. The work area type defines the schema; each work area instance holds the values.

## Methods (in order of preference)

### 1. Get-PWRichProjectScalar (direct read)

```powershell
Get-PWRichProjectScalar -FolderPath "AZDOT 2024\ProjectName\CADD\Sheets" -KeyProjectProperty "ProjectID"
# or: PROJECT_WorkAreaNumber, WorkAreaNumber, Project_Number
```

Returns the scalar value for that work area property. Works on any folder under the work area.

### 2. Get-PWFolderPathAndProperties (rich project root)

```powershell
$folder = Get-PWFolders -FolderPath $path -JustOne
$richProject = $folder | Get-PWRichProjectForFolder
$withProps = $richProject | Get-PWFolderPathAndProperties
# $withProps has properties: ProjectID, PROJECT_WorkAreaNumber, etc.
```

### 3. Get-PWWorkAreaTypeForRichProject (type definition only)

Returns the **work area type** (schema name/definition), not the property values. Use with `Get-PWRichProjectProperties -ProjectType $typeName` to list available columns. Values come from Method 1 or 2.

## Relevant pwps_dab cmdlets

| Cmdlet | Purpose |
|--------|---------|
| `Get-PWRichProjectScalar` | Read a single work area property value by name |
| `Get-PWFolderPathAndProperties` | Populate folder with path + work area property values |
| `Get-PWRichProjectForFolder` | Get containing rich project (work area root) for a folder |
| `Get-PWWorkAreaTypeForRichProject` | Get work area type for a rich project (schema, not values) |
| `Get-PWRichProjectProperties` | List property definitions for a project type (schema) |

## Test script

Run `.\test_project_id_from_workarea.ps1 -FolderPath "AZDOT 2024\YourProject\CADD\Sheets"` (with PW connected) to verify which method returns data in your environment.
