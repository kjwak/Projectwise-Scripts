# Parse-only: worker deployments need email/ assets at repo root; Publish-QCPipelineCode does not copy them.
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$publishScript = Join-Path $repoRoot 'scripts\Publish-QCPipelineCode.ps1'
$templateRel = 'email\templates\qc_notification.html'
$logoRel = 'email\typsalogo.png.webp'
$fail = 0

function Assert-True($Cond, $Msg) {
    if (-not $Cond) {
        Write-Host "FAIL: $Msg" -ForegroundColor Red
        $script:fail++
        return $false
    }
    Write-Host "OK:   $Msg" -ForegroundColor Green
    return $true
}

Assert-True (Test-Path -LiteralPath (Join-Path $repoRoot $templateRel)) 'repo email template exists locally'
Assert-True (Test-Path -LiteralPath (Join-Path $repoRoot $logoRel)) 'repo email logo exists locally'

$publishText = Get-Content -LiteralPath $publishScript -Raw
Assert-True ($publishText -notmatch [regex]::Escape($templateRel)) 'Publish copy plan does not include email template (documented gap)'
Assert-True ($publishText -notmatch 'email\\') 'Publish copy plan does not copy email\ tree'

# Simulate post-4E module folder layout: relative path must not resolve under modules/Notifications.
Import-Module (Join-Path $repoRoot 'modules\Notifications\QC.NotificationTemplates.psm1') -Force
$badUnderNotifications = Join-Path $repoRoot 'modules\Notifications\email\templates\qc_notification.html'
Assert-True (-not (Test-Path -LiteralPath $badUnderNotifications)) 'wrong path under modules/Notifications must not exist'

$html = ConvertTo-QCEmailHtml -TemplatePath 'email/templates/qc_notification.html' -Data @{
    NotificationTitle = 't'
    NotificationMessage = 'm'
    ProjectName = 'p'
    DocumentName = 'd'
    ReviewType = 'r'
    WorkflowState = 'w'
    AssignedTo = 'a'
    SubmittedBy = 's'
    SubmittedDate = '2026-01-01'
    QCPdfUrl = 'https://example.com/qc.pdf'
    GeneratedTimestamp = '2026-01-01'
}
Assert-True ($html -like '*Open Review Package*') 'relative template path resolves from repo root via QC.NotificationTemplates'

if ($fail -gt 0) {
    Write-Host "test_notification_template_paths.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host 'test_notification_template_paths.ps1 passed' -ForegroundColor Green
