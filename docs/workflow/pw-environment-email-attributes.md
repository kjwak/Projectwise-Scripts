# ProjectWise environment email attributes

## Purpose

Document how to read **designer** and **reviewer** email values from ProjectWise sheet documents across **multiple PW environments** (Caltrans, ADOT, MCDOT, POLB, and others). These fields support QC routing, notifications, and reporting without hard-coding a single environment.

Verified against:

`pw:\typsa-us-pw.bentley.com:typsa-us-pw-03\Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1\`

(May 2026, read-only `pwps` / `pwps_dab` automation session.)

---

## Summary

| Topic | Finding |
| --- | --- |
| Folder workflow | **TYPSA QC** is assigned on `Seg_1` (e.g. `In Development`, `Originated`). |
| Folder environment | **Caltrans** (`EnvironmentID` 109 on the folder object from `Get-PWFolders`). |
| Designer email column | **`EM_Designer_Email`** (Caltrans environment definition). |
| Reviewer email column | **`EM_Reviewer_Email`** — not `EM_Reveiewer_Email` (typo). |
| Values on documents | Present on some PDFs only; many documents have empty attribute bags until populated in PW. |
| Reliable read API | **`Get-PWDocumentsBySearchWithReturnColumns`** + parse **`.Attributes`** sorted-list bags. |
| Unreliable read APIs | `Get-PWDocumentEAttributes` (empty here), top-level properties on the search row, CEL `EM_*` expressions without document context. |

---

## Environment catalog (datasource snapshot)

Column names were confirmed via `Get-PWEnvironments` and `Get-PWEnvironmentColumns` on `typsa-us-pw.bentley.com:typsa-us-pw-03`. Use this table as the initial map; re-run discovery after PW Administrator changes.

| Environment | Designer email column | Reviewer email column | Notes |
| --- | --- | --- | --- |
| **Caltrans** | `EM_Designer_Email` | `EM_Reviewer_Email` | Used by CAFWY2200 `Seg_1` sheets folder. |
| **ADOT** | `EM_Designer_Email` | `EM_Reviewer_Email` | Same logical names as Caltrans. |
| **MCDOT** | `EM_Designer_Email` | `EM_Reviewer_Email` | Same logical names. |
| **POLB** | `EM_Designer_Email` | `EM_Reviewer_Email` | Same logical names. |

Other environments exist (renditions, lookups, etc.) but do not define these EM email columns.

**Resolution rule:** The environment name on the **folder** (`Get-PWFolders -JustOne` → `.Environment`) selects which column map to use. Documents inherit the folder environment; do not assume `Documents\...` path alone implies Caltrans.

---

## How extraction works

### 1. Connect and resolve the folder

- Open PW with `Open-PWConnection` (see `modules/ProjectWise/PW.Connection.psm1`).
- Folder path for cmdlets: `Caltrans\...\CADD\Sheets\Seg_1` (often **without** leading `Documents\`; both forms were tried during verification).
- Read folder metadata:

```powershell
$folder = Get-PWFolders -FolderPath $folderPath -JustOne
# $folder.Environment  -> e.g. "Caltrans"
# $folder.Workflow     -> e.g. "TYPSA QC"
```

### 2. Request columns by name

Use the configured column names for that environment (see [Configuration](#configuration) below):

```powershell
$cols = @('EM_Designer_Email', 'EM_Reviewer_Email')
$row = Get-PWDocumentsBySearchWithReturnColumns `
    -FolderPath $folderPath `
    -JustThisFolder `
    -DocumentName $docName `
    -ColumnsToReturn $cols `
    | Select-Object -First 1
```

### 3. Read values from `.Attributes` bags

Returned rows usually **do not** expose `EM_Designer_Email` as a top-level property. Values appear under `.Attributes`: a list of `SortedList<string,string>` bags.

```powershell
function Get-PWDocumentAttributeMap {
    param([object]$DocRow)
    $map = @{}
    if (-not $DocRow -or -not $DocRow.Attributes) { return $map }
    foreach ($bag in @($DocRow.Attributes)) {
        if ($bag -is [System.Collections.IDictionary]) {
            foreach ($k in $bag.Keys) {
                $map[[string]$k] = [string]$bag[$k]
            }
        }
    }
    return $map
}

$attrs = Get-PWDocumentAttributeMap -DocRow $row
$designerEmail = $attrs['EM_Designer_Email']
$reviewerEmail  = $attrs['EM_Reviewer_Email']
```

Example values read from `Seg_1` (May 2026):

- `0818000063ea501.pdf` → `JFlint@aztec.us` / `TMahon@aztec.us`
- `080J082001ab001.pdf` → `JFlint@aztec.us` / `TMahon@aztec.us`

At verification time, **2 of 77** PDFs in that folder had both fields populated; the rest had empty bags for those keys.

### 4. Optional: workflow state on the same row

The same search row can expose workflow context when populated, for example:

- `WorkflowState` → e.g. `Originated` (TYPSA)
- `Workflow` / `WorkflowId`

Use `Get-PWFolderTreeDocumentStateCount` for folder-level rollups (counts per state under TYPSA QC).

---

## Configuration

Automation should **not** hard-code Caltrans column names. Add a small map under `projectWise` in `appsettings.json` (or a dedicated config file loaded by `Core.Config`).

### Recommended shape

```json
{
  "projectWise": {
    "datasourceName": "typsa-us-pw.bentley.com:typsa-us-pw-03",
    "credentialPath": "C:\\PW_QC_LOCAL\\pw_cred.txt",
    "environmentEmailAttributes": {
      "default": {
        "designerEmailColumn": "EM_Designer_Email",
        "reviewerEmailColumn": "EM_Reviewer_Email"
      },
      "byEnvironment": {
        "Caltrans": {
          "designerEmailColumn": "EM_Designer_Email",
          "reviewerEmailColumn": "EM_Reviewer_Email"
        },
        "ADOT": {
          "designerEmailColumn": "EM_Designer_Email",
          "reviewerEmailColumn": "EM_Reviewer_Email"
        },
        "MCDOT": {
          "designerEmailColumn": "EM_Designer_Email",
          "reviewerEmailColumn": "EM_Reviewer_Email"
        },
        "POLB": {
          "designerEmailColumn": "EM_Designer_Email",
          "reviewerEmailColumn": "EM_Reviewer_Email"
        }
      }
    }
  }
}
```

### Resolution algorithm (for modules / scripts)

1. Load `environmentEmailAttributes` from config.
2. For a target folder, `Get-PWFolders -JustOne` and read `$folder.Environment`.
3. If `byEnvironment[$folder.Environment]` exists, use it; else use `default`.
4. Build `-ColumnsToReturn` from `designerEmailColumn` and `reviewerEmailColumn`.
5. Parse `.Attributes` bags into a hashtable (see [How extraction works](#how-extraction-works)).
6. If both values are empty, treat as “not populated” (not as “column missing”) unless environment column discovery fails.

If a future environment uses **different** column names, add only that environment under `byEnvironment`; leave `default` for the common TYPSA naming.

### Validating config against ProjectWise

After PW Administrator adds or renames columns:

```powershell
$envName = 'Caltrans'
$cols = Get-PWEnvironmentColumns -EnvironmentName $envName
# Confirm designerEmailColumn / reviewerEmailColumn exist in the environment definition
```

---

## Verification script

Read-only folder scan (JSON to stdout):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\Test-PWEmailAttributes-Extract.ps1 `
  -FolderPath "pw:\\typsa-us-pw.bentley.com:typsa-us-pw-03\Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1\" `
  -Pretty
```

Output includes:

- `environment`, `workflow`
- `workflowStateCounts`
- `summary.EM_Designer_Email_populated` / `EM_Reviewer_Email_populated`
- `samples` with extracted emails

Run under the same PowerShell host / MTA profile used by QC workers (`pwps_dab`).

---

## APIs to avoid for this data

| API / pattern | Result in verification |
| --- | --- |
| `Get-PWDocumentEAttributes -DocumentID -ProjectID` | Returned an empty `List<EAttribute>` for tested PDFs. |
| `$row.EM_Designer_Email` (top-level) | Property not present on the search object. |
| `EM_Reveiewer_Email` | Not defined in any environment column list. |
| `Invoke-PWCelExpression` with bare `EM_Designer_Email` | Activation errors without proper document context. |
| `Get-PWEnvironmentColumns` on entire datasource without iterating environments | Misleading; must unwrap `Get-PWEnvironments` list and query per environment name. |

---

## Implementation notes for QC / reporting

- **Environment-first:** Always resolve column names from folder environment before reading documents.
- **PDF vs DGN:** Email samples were verified on **PDFs** in `Seg_1`; DGN files had empty EM email bags in spot checks. Prefer PDFs for QC notification unless business rules say otherwise.
- **Population vs schema:** Empty values mean the attribute exists in the environment but is not set on that document version—not that extraction failed.
- **Workflow vs attributes:** TYPSA QC workflow state and Caltrans environment attributes are independent layers; folder carries both.
- **Future module home:** Consider `Get-PWEnvironmentEmailAttributeSettings`, `Get-PWDocumentAttributeMap`, and `Get-PWDocumentEmailContacts` in `modules/ProjectWise/PW.Discovery.psm1` or a thin `PW.EnvironmentAttributes.psm1`, driven by `environmentEmailAttributes` config.

---

## Related documentation

- QC workflow states: `docs/workflow/qc-workflow-framework.md` (when present on branch)
- ProjectWise cmdlet conventions: `.cursor/rules/projectwise-powershell.mdc`
- Path normalization (`pw:\...` URIs): `modules/Core/Core.Paths.psm1`

---

## Integration in Pipeline

The email attribute extraction method documented above is used in:

### Legacy prepend email sync (`*-qc.pdf` bridge)

**`scripts/processing/Invoke-QCPrependPw.ps1`** (production path when `qcPrepend.mode: projectWise`) — After each successful QC prepend, `Sync-PWQcPdfEmailAttributesFromSourcePdf` copies `EM_Designer_Email` and `EM_Reviewer_Email` from the **source sheet PDF** to the **lane QC PDF** (or legacy `*-qc.pdf` when lane params are absent). The source `.pdf` is the source of truth; the QC PDF is updated when those attributes are missing or differ.

Lane PDFs (`*-prod/-rev/-chk.pdf`) also receive `QC_Process_Type` via `_PWD-EnsureLaneQcPdfProcessTypeAttribute` during prepend.
- **`Watch-QCTrigger.ps1`** — During both full reconciliation scans and audit-trail scans, `Get-PWDocumentsBySearchWithReturnColumns` is called with `EM_Designer_Email` and `EM_Reviewer_Email` as return columns. Extracted values are written to the `sheet_index` database table via `Write-QCSheetIndex`.
- **`Test-SheetIndexAndAuditPoller.ps1`** — Validation script uses the same extraction pattern to populate `sheet_index` during testing.
- **`QC.Notifications.psm1`** — Uses the `_QCN-GetAttributeValue` helper to read email attributes from document objects for notification routing.

The `sheet_index` table (`docs/data/database-telemetry.md`) stores extracted emails in `designer_email` and `reviewer_email` columns, and workflow state in `pw_state_name`.

### Module helpers (`PW.Discovery.psm1`)

| Function | Purpose |
| --- | --- |
| `Get-PWDocumentAttributeMap` | Parse `.Attributes` bags from a search row |
| `Get-PWDocumentEmailContacts` | Read designer/reviewer emails for one document |
| `Sync-PWQcPdfEmailAttributesFromSourcePdf` | Copy emails from source PDF to matching lane QC PDF (legacy `*-qc.pdf` bridge) |

### Diagnostic Script

`tools/discovery/Test-PWDocumentProperties.ps1` can be used to enumerate all properties, attributes, and environment columns available on documents in the current PW environment. Run this when email attributes are not populating as expected.

## Change log

| Date | Change |
| --- | --- |
| 2026-05-19 | Initial write-up from `Seg_1` verification; multi-environment column map and config schema. |
| 2026-05-26 | Added pipeline integration notes; documented `sheet_index` usage and diagnostic script. |
| 2026-05-26 | QC print sync: source PDF → `*-qc.pdf` via `Sync-PWQcPdfEmailAttributesFromSourcePdf`. |
