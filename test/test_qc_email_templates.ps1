$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}
function Assert-Contains($Haystack, $Needle, $Message) {
    if ($Haystack -notlike "*$Needle*") {
        throw "ASSERT FAILED: $Message`nExpected substring: $Needle"
    }
}
function Assert-Throws($ScriptBlock, $Message) {
    $threw = $false
    try { & $ScriptBlock } catch { $threw = $true }
    if (-not $threw) { throw "ASSERT FAILED: $Message" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core/Core.Runtime.psm1" -Force
Import-Module "$repoRoot/modules/Core/Core.Config.psm1" -Force
Import-Module "$repoRoot/modules/Notifications/QC.NotificationGraph.psm1" -Force
Import-Module "$repoRoot/modules/Notifications/QC.Notifications.psm1" -Force
# Re-import after Notifications so exports remain in the global session (nested import unloads globals).
Import-Module "$repoRoot/modules/Notifications/QC.NotificationGraph.psm1" -Force
Import-Module "$repoRoot/modules/Notifications/QC.NotificationTemplates.psm1" -Force
Import-Module "$repoRoot/modules/Core/Core.Config.psm1" -Force

$pascalSubjectTokens = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
$pascalSubjectTokens['DocumentName'] = '0818000063ea501-qc.pdf'
$pascalSubjectTokens['documentName'] = '0818000063ea501-qc.pdf'
$pascalSubjectTokens['ProjectName'] = 'CAFWY2200-I-15_ELPSE'
$pascalSubjectTokens['project'] = 'CAFWY2200-I-15_ELPSE'
$pascalSubjectTokens['ReviewType'] = 'Production QC'
$pascalSubjectTokens['reviewType'] = 'Production QC'
$pascalSubjectTokens['WorkflowState'] = 'Corrections Received'
$pascalSubjectTokens['currentState'] = 'Corrections Received'
$pascalTmpl = '[{ReviewType}] {ProjectName} | {DocumentName} | {WorkflowState}'
$pascalSubject = Expand-QCNotificationTemplate -Template $pascalTmpl -Tokens $pascalSubjectTokens
Assert-Eq $pascalSubject '[Production QC] CAFWY2200-I-15_ELPSE | 0818000063ea501-qc.pdf | Corrections Received' `
    'PascalCase and camelCase subject placeholders must both expand'

$templatePath = Join-Path $repoRoot 'email\templates\qc_notification.html'
$logoPath = Join-Path $repoRoot 'email\typsalogo.png.webp'
$samplePath = Join-Path $repoRoot 'email\sample_notification_data.json'

Assert-True (Test-Path -LiteralPath $templatePath) 'qc_notification.html must exist'
Assert-True (Test-Path -LiteralPath $logoPath) 'typsalogo.png.webp must exist'
Assert-True (Test-Path -LiteralPath $samplePath) 'sample_notification_data.json must exist'

$sampleRaw = Get-Content -LiteralPath $samplePath -Raw -Encoding UTF8
$sample = $sampleRaw | ConvertFrom-Json
$sampleHt = @{}
foreach ($p in $sample.PSObject.Properties) {
    $sampleHt[$p.Name] = if ($null -eq $p.Value) { '' } else { [string]$p.Value }
}

$html = ConvertTo-QCEmailHtml -TemplatePath $templatePath -Data $sampleHt
Assert-Contains $html 'Open QC PDF' 'Rendered HTML should include CTA label'
Assert-Contains $html 'cid:typsa-logo' 'Rendered HTML should reference inline logo CID'
Assert-Contains $html '#bf1425' 'Rendered HTML should use TYPSA accent color'
Assert-Contains $html ([System.Net.WebUtility]::HtmlEncode([string]$sampleHt.ProjectName)) 'Project name should be HTML-encoded'

Assert-Throws {
    $bad = @{}
    foreach ($k in $sampleHt.Keys) { if ($k -ne 'QCPdfUrl') { $bad[$k] = $sampleHt[$k] } }
    ConvertTo-QCEmailHtml -TemplatePath $templatePath -Data $bad | Out-Null
} 'Missing QCPdfUrl should throw'

Assert-Throws {
    $bad = @{}
    foreach ($k in $sampleHt.Keys) { $bad[$k] = $sampleHt[$k] }
    $bad['QCPdfUrl'] = 'not-a-valid-url'
    ConvertTo-QCEmailHtml -TemplatePath $templatePath -Data $bad | Out-Null
} 'Invalid QCPdfUrl should throw'

$escaped = ConvertTo-QCEmailHtml -TemplatePath $templatePath -Data (@{
    NotificationTitle = 'Test'
    NotificationMessage = 'Test'
    ProjectName = '<script>alert(1)</script>'
    DocumentName = 'doc.pdf'
    ReviewType = 'Peer'
    WorkflowState = 'QC Received'
    AssignedTo = 'a@b.com'
    SubmittedBy = 'c@d.com'
    SubmittedDate = '2026-01-01'
    QCPdfUrl = 'https://example.com/doc.pdf'
    GeneratedTimestamp = '2026-01-01'
})
Assert-Contains $escaped '&lt;script&gt;' 'Script tags in project name must be encoded'

$graphMsg = New-QCGraphEmailMessage -ToRecipients @('test@example.com') -Subject 'Test' `
    -HtmlBody '<p>Hi</p>' -LogoPath $logoPath
Assert-True ($graphMsg.message.body.contentType -eq 'HTML') 'Graph body should be HTML'
Assert-True ($graphMsg.message.attachments.Count -eq 1) 'Graph message should include one attachment'
Assert-True ($graphMsg.message.attachments[0].contentId -eq 'typsa-logo') 'Inline logo contentId must be typsa-logo'
Assert-True ($graphMsg.message.attachments[0].isInline -eq $true) 'Logo attachment must be inline'
Assert-True ($graphMsg.saveToSentItems -eq $true) 'saveToSentItems must be true'

Assert-Throws {
    New-QCGraphEmailMessage -ToRecipients @() -Subject 'x' -HtmlBody 'y' -LogoPath $logoPath | Out-Null
} 'Empty To recipients should throw'

$event = New-QCNotificationEvent -EventType 'QC_RECEIVED' -DocumentName 'D-101-qc.pdf' `
    -CurrentState 'QC Received' -DocumentGuid 'guid-101' -QcPdfUrl 'https://example.com/qc.pdf'
$settings = Get-QCNotificationSettings -Config @{ notifications = @{ email = @{ bodyFormat = 'Text' } } }
$templateData = New-QCNotificationEmailTemplateData -Event $event -Settings $settings
Assert-Eq $templateData.QCPdfUrl 'https://example.com/qc.pdf' 'Template data should carry qcPdfUrl from event'

$configPath = Join-Path $repoRoot 'appsettings.json'
$configRes = Read-QCAppSettings -Path $configPath
if (-not $configRes.IsSuccess) { throw $configRes.Message }
$configHt = $configRes.Data.config
Import-Module "$repoRoot/modules/Core/Core.Config.psm1" -Force
$azdotFolder = 'Documents\AZDOT 2024\Sample Highway\CADD\Sheets'
$azdotProject = Get-QCProjectNameFromFolderPath -Config $configHt -FolderPath $azdotFolder
Assert-Eq $azdotProject 'Sample Highway' 'AZDOT 2024 project name should be folder segment before CADD\Sheets'

$caltransFolder = 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'
$caltransProject = Get-QCProjectNameFromFolderPath -Config $configHt -FolderPath $caltransFolder
Assert-Eq $caltransProject 'CAFWY2200-I-15_ELPSE' 'Caltrans project name should be between Documents\Caltrans and CADD\Sheets'

$caltransDocPath = 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1\0818000063ea501-qc.pdf'
$caltransFromDoc = Get-QCProjectNameFromFolderPath -Config $configHt -FolderPath ([System.IO.Path]::GetDirectoryName($caltransDocPath))
Assert-Eq $caltransFromDoc 'CAFWY2200-I-15_ELPSE' 'Caltrans project name should not include CADD\Sheets or Seg_1'

$pathEvent = @{
    eventType = 'READY_FOR_QC'
    documentName = 'D-101-qc.pdf'
    documentPath = 'pw:\ds\Documents\AZDOT 2024\Bridge Rehab\CADD\Sheets\D-101-qc.pdf'
    currentState = 'Ready for QC'
    qcPdfUrl = 'https://example.com/qc.pdf'
    reviewers = @('r@x.com')
    designers = @('d@x.com')
}
$pathData = New-QCNotificationEmailTemplateData -Event $pathEvent -Settings $settings -Config $configHt
Assert-Eq $pathData.ProjectName 'Bridge Rehab' 'Template data should derive project name from document path'

$builtUrl = Resolve-QCNotificationQcPdfUrl -Event $event -Settings $settings
Assert-Eq $builtUrl 'https://example.com/qc.pdf' 'Resolver should prefer event qcPdfUrl'

$guidEvent = New-QCNotificationEvent -EventType 'T' -DocumentName 'x-qc.pdf' -CurrentState 'S' `
    -DocumentGuid 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
$guidSettings = Get-QCNotificationSettings -Config @{
    notifications = @{
        email = @{
            pwLinkBaseUrl = 'https://connect.example.com/pwlink/'
            pwLinkApp = 'pwe'
        }
    }
}
$pwConfig = @{ projectWise = @{ datasourceName = 'Test:pw' } }
$link = Resolve-QCNotificationQcPdfUrl -Event $guidEvent -Settings $guidSettings -Config $pwConfig
Assert-Contains $link 'objectId=aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' 'pwlink builder should include document GUID'
Assert-Contains $link 'objectType=doc' 'pwlink builder should target document'

Write-Host 'All QC email template tests passed.' -ForegroundColor Green
