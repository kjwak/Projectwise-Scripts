# Orchestrator test with mocked export/extract (no PW).
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Processing/QC.PdfExport.psm1" -Force
Import-Module "$repoRoot/modules/Processing/QC.CommentStatusProcessor.psm1" -Force

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

$tmp = Join-Path $env:TEMP ("qc-sync-orch-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$pdf = Join-Path $tmp 'sample-qc.pdf'
Set-Content -LiteralPath $pdf -Value '%PDF-1.4 mock' -Encoding ASCII

$config = @{
    dryRun = $true
    qcCommentSync = @{
        enabled = $true
        processorVersion = '1.0.0-test'
        stagingRoot = $tmp
        retainTempFiles = $false
    }
    qcWorkflow = @{ dryRunWriteback = $true; autoSetState = $false }
    database = @{ enabled = $false }
    notifications = @{ dryRun = $true; adminRecipients = @('admin@test') }
}

$job = @{
    id = 'job-test-1'
    type = 'QC_COMMENT_STATUS_SYNC'
    sourcePath = 'Documents\Proj\Sheets\A101-qc.pdf'
    sourceName = 'A101-qc.pdf'
    sourceFolder = 'Documents\Proj\Sheets'
    triggerRule = @{ id = 'qc-comment-status-pw'; jobType = 'QC_COMMENT_STATUS_SYNC' }
    metadata = @{
        candidate = @{
            documentGuid = 'guid-test'
            file = @{ sha256 = 'abc' }
        }
    }
}

$mockAnnots = @(
    @{ annotation_id = '1'; author = 'QC Reviewer'; status = 'Open'; page_number = 1 }
)

$script:mockPdfPath = $pdf
$mockExport = {
    param($Doc, $Target)
    return @{
        IsSuccess = $true
        Code = 'OK'
        Message = 'mock'
        Data = @{ localPath = $script:mockPdfPath }
    }
}

$doc = [pscustomobject]@{ Name = 'A101-qc.pdf' }
$res = Invoke-QCCommentStatusSyncProcessor -Job $job -Config $config -InputDocument $doc `
    -MockExporter $mockExport -MockAnnotations $mockAnnots

Assert-True $res.IsSuccess 'Orchestrator should succeed'
Assert-True ($res.Data.decision.targetState -eq 'Redlines Received') 'Should decide redlines received'
Assert-True ([bool]$res.Data.dryRun) 'Dry run flag set'

Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
Write-Host 'OK test_qc_comment_sync_orchestrator.ps1'
