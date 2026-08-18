$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Core\Core.Paths.psm1') -Force
Import-Module (Join-Path $root 'modules\Processing\QC.StatusSet.psm1') -Force

$failures = 0
function _Assert([bool]$ok, [string]$msg) {
    if (-not $ok) { Write-Host "  FAIL: $msg" -ForegroundColor Red; $script:failures++ }
    else          { Write-Host "  ok:   $msg" -ForegroundColor Green }
}

# Build a synthetic localRoot mimicking the workspace layout the active processor
# emits: <localRoot>\<16hex>\_statusset.manifest.json + _StatusSet.pdf
$localRoot = Join-Path $env:TEMP ("ssreconcile_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $localRoot -Force | Out-Null

function _NewWorkspace([string]$key, [string]$sheetsFolder, [bool]$WithPdf, [bool]$WithManifest) {
    $ws = Join-Path $localRoot $key
    New-Item -ItemType Directory -Path $ws -Force | Out-Null
    if ($WithPdf) {
        $pdf = Join-Path $ws '_StatusSet.pdf'
        # Tiny content; reconcile only inspects mtime/size, not bytes.
        Set-Content -LiteralPath $pdf -Value 'pdfbody' -Encoding ASCII
    }
    if ($WithManifest) {
        $man = @{
            version = 1
            sheetsFolder = $sheetsFolder
            folderStateHash = '0' * 64
            generatedAtUtc = ([datetime]::UtcNow.ToString('o'))
            sheetCount = 1
            sheets = @()
        }
        $man | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $ws '_statusset.manifest.json') -Encoding UTF8
    }
    return $ws
}

try {
    Write-Host "Test: Get-StatusSetWorkspaceManifests filters out incomplete workspaces" -ForegroundColor Cyan

    $wsGood = _NewWorkspace 'aaaaaaaaaaaaaaaa' 'Documents\AZDOT\Prj1\CADD\Sheets' $true  $true
    $wsNoPdf = _NewWorkspace 'bbbbbbbbbbbbbbbb' 'Documents\AZDOT\Prj2\CADD\Sheets' $false $true
    $wsNoMan = _NewWorkspace 'cccccccccccccccc' 'Documents\AZDOT\Prj3\CADD\Sheets' $true  $false
    $wsMissingFolder = _NewWorkspace 'dddddddddddddddd' '' $true $true

    $walk = Get-StatusSetWorkspaceManifests -LocalRoot $localRoot
    _Assert ($walk.IsSuccess)                                      "walk succeeds"
    $records = @($walk.Data.records)
    _Assert ($records.Count -eq 1)                                 "exactly one complete workspace surfaced"
    _Assert ($records[0].workspaceDir -eq $wsGood)                 "the complete workspace is the one with both files"
    _Assert ($records[0].pwPath -eq 'AZDOT\Prj1\CADD\Sheets')      "pwPath strips leading 'Documents\\'"
    _Assert ($records[0].sheetsFolder -eq 'Documents\AZDOT\Prj1\CADD\Sheets') "sheetsFolder preserved"
    _Assert ($null -ne $records[0].outputPdfLastWriteUtc)          "local PDF mtime captured"

    $skipped = @($walk.Data.skipped)
    _Assert ($skipped.Count -eq 3)                                 "skipped count is 3 (no-pdf + no-manifest + missing-folder)"
    $skipReasons = @($skipped | ForEach-Object { [string]$_.reason } | Sort-Object -Unique)
    _Assert ($skipReasons -contains 'NO_LOCAL_PDF')                "NO_LOCAL_PDF skip reason emitted"
    _Assert ($skipReasons -contains 'NO_MANIFEST')                 "NO_MANIFEST skip reason emitted"
    _Assert ($skipReasons -contains 'MANIFEST_MISSING_SHEETS_FOLDER') "MANIFEST_MISSING_SHEETS_FOLDER skip reason emitted"

    Write-Host ""
    Write-Host "Test: Sync-StatusSetWorkspaceToPw decision matrix uses stub PW cmdlets" -ForegroundColor Cyan

    # Override the PW cmdlets in the parent script's scope so Sync-StatusSetWorkspaceToPw
    # can pick them up via Get-Command. We use script:GlobalState to drive each scenario.
    $script:_PwExisting = $null  # null => PW has no _StatusSet.pdf
    $script:_PwUpdateCalls = 0
    $script:_PwNewCalls = 0

    function Get-PWDocumentsBySearch {
        param($FolderPath, [switch]$JustThisFolder, $DocumentName, [switch]$PopulatePath)
        $script:_PwExisting
    }
    function Update-PWDocumentFile {
        param($InputDocuments, [string]$NewFilePathName)
        $script:_PwUpdateCalls++
        $null
    }
    function New-PWDocument {
        param($FolderPath, $FilePath, $DocumentName)
        $script:_PwNewCalls++
        $null
    }

    # Scenario A: PW missing the doc -> CREATE
    $script:_PwExisting = $null
    $rA = Sync-StatusSetWorkspaceToPw -WorkspaceRecord $records[0]
    _Assert ($rA.IsSuccess)                                        "scenario A succeeds"
    _Assert ($rA.Code -eq 'STATUS_SET_RECONCILE_CREATED')          "scenario A code is CREATED"
    _Assert ($script:_PwNewCalls -eq 1)                            "New-PWDocument called once"
    _Assert ($script:_PwUpdateCalls -eq 0)                         "Update-PWDocumentFile not called"

    # Scenario B: PW has it AND is newer -> IN_SYNC
    $script:_PwUpdateCalls = 0; $script:_PwNewCalls = 0
    $script:_PwExisting = [pscustomobject]@{
        DocumentUpdateDate = [datetime]::UtcNow.AddMinutes(60)  # PW is in the future relative to local
        Name = '_StatusSet.pdf'
    }
    $rB = Sync-StatusSetWorkspaceToPw -WorkspaceRecord $records[0]
    _Assert ($rB.IsSuccess)                                        "scenario B succeeds"
    _Assert ($rB.Code -eq 'STATUS_SET_RECONCILE_IN_SYNC')          "scenario B code is IN_SYNC"
    _Assert ($script:_PwUpdateCalls -eq 0)                         "Update not called when PW newer"
    _Assert ($script:_PwNewCalls -eq 0)                            "Create not called when PW newer"

    # Scenario C: PW has it BUT is older -> UPDATE
    $script:_PwUpdateCalls = 0; $script:_PwNewCalls = 0
    $script:_PwExisting = [pscustomobject]@{
        DocumentUpdateDate = [datetime]::UtcNow.AddDays(-5)
        Name = '_StatusSet.pdf'
    }
    $rC = Sync-StatusSetWorkspaceToPw -WorkspaceRecord $records[0]
    _Assert ($rC.IsSuccess)                                        "scenario C succeeds"
    _Assert ($rC.Code -eq 'STATUS_SET_RECONCILE_UPDATED')          "scenario C code is UPDATED"
    _Assert ($script:_PwUpdateCalls -eq 1)                         "Update-PWDocumentFile called once"
    _Assert ($script:_PwNewCalls -eq 0)                            "New-PWDocument not called"

    Write-Host ""
    Write-Host "Test: Invoke-StatusSetReconcile aggregates counts" -ForegroundColor Cyan

    # Two workspaces, force one to be IN_SYNC (PW newer) and one to be UPDATED (PW older)
    $wsB = _NewWorkspace 'eeeeeeeeeeeeeeee' 'Documents\AZDOT\PrjB\CADD\Sheets' $true $true
    $script:_PwUpdateCalls = 0; $script:_PwNewCalls = 0

    # We need a per-folder dispatch; Get-PWDocumentsBySearch is single-stub, so use a hashtable
    $script:_PwIndex = @{
        'AZDOT\Prj1\CADD\Sheets' = [pscustomobject]@{ DocumentUpdateDate = [datetime]::UtcNow.AddMinutes(60); Name = '_StatusSet.pdf' }  # IN_SYNC
        'AZDOT\PrjB\CADD\Sheets' = [pscustomobject]@{ DocumentUpdateDate = [datetime]::UtcNow.AddDays(-5);  Name = '_StatusSet.pdf' }     # UPDATED
    }
    function Get-PWDocumentsBySearch {
        param($FolderPath, [switch]$JustThisFolder, $DocumentName, [switch]$PopulatePath)
        if ($script:_PwIndex.ContainsKey($FolderPath)) { return $script:_PwIndex[$FolderPath] }
        return $null
    }
    $config = @{ statusSet = @{ localRoot = $localRoot } }
    $events = New-Object System.Collections.ArrayList
    $cb = { param($e) [void]$events.Add($e) }
    $rRec = Invoke-StatusSetReconcile -Config $config -LogCallback $cb
    _Assert ($rRec.IsSuccess)                                                "reconcile succeeds"
    _Assert ($null -ne $rRec.Data.historyRetention)                           "historyRetention attached"
    _Assert ([int]$rRec.Data.counts.considered -eq 2)                        "considered = 2 (good + new)"
    _Assert ([int]$rRec.Data.counts.inSync -eq 1)                            "inSync = 1"
    _Assert ([int]$rRec.Data.counts.updated -eq 1)                           "updated = 1"
    _Assert ([int]$rRec.Data.counts.created -eq 0)                           "created = 0"
    _Assert ([int]$rRec.Data.counts.failed -eq 0)                            "failed = 0"
    _Assert ($events.Count -eq 2)                                            "log callback invoked once per workspace"
    _Assert ([int]$rRec.Data.counts.skipped -ge 3)                           "skipped count carries over from walk"

    if ($failures -gt 0) { Write-Host "`nFAILED ($failures)" -ForegroundColor Red; exit 1 }
    Write-Host "`nPASSED" -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $localRoot -Recurse -Force -ErrorAction SilentlyContinue
}
