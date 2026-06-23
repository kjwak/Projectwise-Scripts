# Unit/integration tests for Invoke-QCPrependProcessor.
param(
    # Optional: provide real PDFs for the 3-step overlay progression test.
    # If all three are set, the test uses these files instead of generating PDFs.
    [Parameter(Mandatory = $false)]
    [string]$PdfV1,
    [Parameter(Mandatory = $false)]
    [string]$PdfV2,
    [Parameter(Mandatory = $false)]
    [string]$PdfV3
)

$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}
function Assert-Contains($Haystack, $Needle, $Message) {
    if ($Haystack -notlike "*$Needle*") { throw "ASSERT FAILED: $Message`nExpected substring: $Needle" }
}

function Assert-NonEmptyPath([object]$Value, [string]$Message) {
    $s = ''
    try { $s = [string]$Value } catch { $s = '' }
    if ([string]::IsNullOrWhiteSpace($s)) { throw "ASSERT FAILED: $Message" }
    return $s
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')

Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Processing/QC.Processors.psm1" -Force

function New-TempRoot() {
    $t = Join-Path $env:TEMP ("qc-prepend-proc-test-" + ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $t -Force | Out-Null
    return $t
}

function Write-ValidPdf([string]$QpdfExe, [string]$Path) {
    # Create a valid empty PDF using qpdf.
    $stdout = [System.IO.Path]::GetTempFileName()
    $stderr = [System.IO.Path]::GetTempFileName()
    try {
        $p = Start-Process -FilePath $QpdfExe -ArgumentList @('--empty','--', $Path) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        if ($p.ExitCode -ne 0) {
            $err = Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue
            throw "qpdf failed creating test PDF: $err"
        }
    } finally {
        Remove-Item -LiteralPath $stdout -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderr -Force -ErrorAction SilentlyContinue
    }
}

function Write-OnePagePdf([string]$Path) {
    # Minimal 1-page PDF (enough for qpdf/overlay tooling to treat as a valid single-page document).
    $pdf = @'
%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Resources << >> >>
endobj
4 0 obj
<< /Length 0 >>
stream
endstream
endobj
xref
0 5
0000000000 65535 f 
0000000010 00000 n 
0000000060 00000 n 
0000000121 00000 n 
0000000223 00000 n 
trailer
<< /Size 5 /Root 1 0 R >>
startxref
290
%%EOF
'@
    [System.IO.File]::WriteAllText($Path, $pdf, [System.Text.Encoding]::ASCII)
}


function Write-OnePagePdfForQpdf([string]$QpdfExe, [string]$Path) {
    $tmp = [System.IO.Path]::GetTempFileName() + '.pdf'
    try {
        Write-OnePagePdf -Path $tmp
        $stdout = [System.IO.Path]::GetTempFileName()
        $stderr = [System.IO.Path]::GetTempFileName()
        try {
            $p = Start-Process -FilePath $QpdfExe -ArgumentList @($tmp, $Path) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
            if (-not (Test-Path -LiteralPath $Path)) {
                $out = [string](Get-Content -LiteralPath $stdout -Raw -ErrorAction SilentlyContinue)
                $err = [string](Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue)
                throw "qpdf rewrite failed: exitCode=$($p.ExitCode); stdout=$out; stderr=$err"
            }
            $pageCount = Get-PdfPageCount -QpdfExe $QpdfExe -Path $Path
            if ($pageCount -lt 1) {
                throw "qpdf rewrite produced $pageCount pages for $Path"
            }
        } finally {
            Remove-Item -LiteralPath $stdout -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $stderr -Force -ErrorAction SilentlyContinue
        }
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

function Get-PdfPageCount([string]$QpdfExe, [string]$Path) {
    $stdout = [System.IO.Path]::GetTempFileName()
    $stderr = [System.IO.Path]::GetTempFileName()
    try {
        $p = Start-Process -FilePath $QpdfExe -ArgumentList @('--show-npages','--', $Path) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        $out = [string](Get-Content -LiteralPath $stdout -Raw -ErrorAction SilentlyContinue)
        $err = [string](Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue)
        $exitCode = [int]$p.ExitCode

        $pageCount = $null
        foreach ($line in ($out -split "(\r?\n)")) {
            $trimmed = ([string]$line).Trim()
            if ($trimmed -match '^\d+$') {
                $pageCount = [int]$trimmed
                break
            }
        }

        if ($null -ne $pageCount) {
            return $pageCount
        }

        $detail = "exitCode=$exitCode"
        if (-not [string]::IsNullOrWhiteSpace($out)) { $detail += "; stdout=$out" }
        if (-not [string]::IsNullOrWhiteSpace($err)) { $detail += "; stderr=$err" }
        throw "qpdf --show-npages failed: no parseable page count. $detail"
    } finally {
        Remove-Item -LiteralPath $stdout -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderr -Force -ErrorAction SilentlyContinue
    }
}

function New-Job([string]$Id, [string]$SourcePdf) {
    return @{
        id = $Id
        type = 'QC_PREPEND'
        sourcePath = $SourcePdf
        sourceName = [System.IO.Path]::GetFileName($SourcePdf)
        sourceFolder = [System.IO.Path]::GetDirectoryName($SourcePdf)
        dedupeKey = ('dq_' + $Id)
        triggerRule = @{ id = 'r1'; jobType = 'QC_PREPEND' }
        attempts = 0
        metadata = @{ qcProcessType = 'production' }
    }
}

$qpdfExe = Join-Path $repoRoot 'tools\qpdf\bin\qpdf.exe'

# source missing -> failure
$tmp1 = New-TempRoot
try {
    $cfg = @{
        dryRun = $false
        qcPrepend = @{ historyRoot = (Join-Path $tmp1 'history'); tempRoot = (Join-Path $tmp1 'temp'); outputRoot = (Join-Path $tmp1 'output'); enableOverlay = $false; qpdfExePath = $qpdfExe }
    }
    $job = New-Job -Id 'missing' -SourcePdf (Join-Path $tmp1 'nope.pdf')
    $r = Invoke-QCPrependProcessor -Job $job -Config $cfg
    Assert-True (-not $r.IsSuccess) 'Missing source should fail'
    Assert-Eq $r.Code 'QC_PREPEND_SOURCE_NOT_FOUND' 'Missing source code'
} finally { Remove-Item -LiteralPath $tmp1 -Recurse -Force -ErrorAction SilentlyContinue }

# initial history creation path
$tmp2 = New-TempRoot
try {
    $srcDir = Join-Path $tmp2 'src'
    New-Item -ItemType Directory -Path $srcDir -Force | Out-Null
    $src = Join-Path $srcDir 'a.pdf'
    Write-ValidPdf -QpdfExe $qpdfExe -Path $src

    $historyRoot = Join-Path $tmp2 'history'
    $cfg = @{
        dryRun = $false
        qcPrepend = @{ historyRoot = $historyRoot; tempRoot = (Join-Path $tmp2 'temp'); outputRoot = (Join-Path $tmp2 'output'); enableOverlay = $false; qpdfExePath = $qpdfExe }
    }
    $job = New-Job -Id 'init' -SourcePdf $src
    $r = Invoke-QCPrependProcessor -Job $job -Config $cfg
    Assert-True $r.IsSuccess 'Initial history creation should succeed'
    $histPath = Assert-NonEmptyPath $r.Data.targetHistoryPdf ("targetHistoryPdf missing for init run. Data=" + ($r.Data | ConvertTo-Json -Depth 8 -Compress))
    Assert-True (Test-Path -LiteralPath $histPath) 'History PDF should be created'
} finally { Remove-Item -LiteralPath $tmp2 -Recurse -Force -ErrorAction SilentlyContinue }

# prepend existing history path
$tmp3 = New-TempRoot
try {
    $srcDir = Join-Path $tmp3 'src'
    New-Item -ItemType Directory -Path $srcDir -Force | Out-Null
    $src1 = Join-Path $srcDir 'a.pdf'
    $src2 = Join-Path $srcDir 'a2.pdf'
    Write-ValidPdf -QpdfExe $qpdfExe -Path $src1
    Write-ValidPdf -QpdfExe $qpdfExe -Path $src2

    $cfg = @{
        dryRun = $false
        qcPrepend = @{ historyRoot = (Join-Path $tmp3 'history'); tempRoot = (Join-Path $tmp3 'temp'); outputRoot = (Join-Path $tmp3 'output'); enableOverlay = $false; qpdfExePath = $qpdfExe }
    }
    $job = New-Job -Id 'prep' -SourcePdf $src1
    $r1 = Invoke-QCPrependProcessor -Job $job -Config $cfg
    Assert-True $r1.IsSuccess 'First run should succeed'
    $hist = Assert-NonEmptyPath $r1.Data.targetHistoryPdf ("targetHistoryPdf missing for prepend run1. Data=" + ($r1.Data | ConvertTo-Json -Depth 8 -Compress))
    $len1 = (Get-Item -LiteralPath $hist).Length

    # second run with different source, but same target (override sourceName to match)
    $job2 = New-Job -Id 'prep' -SourcePdf $src2
    $job2.sourceName = $job.sourceName
    $r2 = Invoke-QCPrependProcessor -Job $job2 -Config $cfg
    Assert-True $r2.IsSuccess 'Prepend run should succeed'
    $len2 = (Get-Item -LiteralPath $hist).Length
    Assert-True ($len2 -ge $len1) 'History PDF should be updated (size non-decreasing)'
} finally { Remove-Item -LiteralPath $tmp3 -Recurse -Force -ErrorAction SilentlyContinue }

# overlay disabled (no exe call)
$tmp4 = New-TempRoot
try {
    $srcDir = Join-Path $tmp4 'src'
    New-Item -ItemType Directory -Path $srcDir -Force | Out-Null
    $src = Join-Path $srcDir 'a.pdf'
    Write-ValidPdf -QpdfExe $qpdfExe -Path $src
    $cfg = @{
        dryRun = $false
        qcPrepend = @{ historyRoot = (Join-Path $tmp4 'history'); tempRoot = (Join-Path $tmp4 'temp'); outputRoot = (Join-Path $tmp4 'output'); enableOverlay = $false; qpdfExePath = $qpdfExe }
    }
    $job = New-Job -Id 'nooverlay' -SourcePdf $src
    $r = Invoke-QCPrependProcessor -Job $job -Config $cfg
    Assert-True $r.IsSuccess 'Overlay disabled should succeed'
    Assert-Eq $r.Data.overlayEnabled $false 'overlayEnabled should be false'
} finally { Remove-Item -LiteralPath $tmp4 -Recurse -Force -ErrorAction SilentlyContinue }

# overlay enabled but exe missing -> failure
$tmp5 = New-TempRoot
try {
    $srcDir = Join-Path $tmp5 'src'
    New-Item -ItemType Directory -Path $srcDir -Force | Out-Null
    $src = Join-Path $srcDir 'a.pdf'
    Write-ValidPdf -QpdfExe $qpdfExe -Path $src
    $cfg = @{
        dryRun = $false
        qcPrepend = @{
            historyRoot = (Join-Path $tmp5 'history')
            tempRoot = (Join-Path $tmp5 'temp')
            outputRoot = (Join-Path $tmp5 'output')
            enableOverlay = $true
            overlayExePath = (Join-Path $tmp5 'missing.exe')
            overlayArgumentsTemplate = '"{sourcePdf}" "{historyPdf}" -o "{outputPdf}"'
            qpdfExePath = $qpdfExe
        }
    }
    $job = New-Job -Id 'overlayfail' -SourcePdf $src
    $r = Invoke-QCPrependProcessor -Job $job -Config $cfg
    Assert-True (-not $r.IsSuccess) 'Missing overlay exe should fail'
    Assert-Eq $r.Code 'QC_OVERLAY_EXE_MISSING' 'Overlay missing exe code'
} finally { Remove-Item -LiteralPath $tmp5 -Recurse -Force -ErrorAction SilentlyContinue }

# dry-run does not write files or call exe
$tmp6 = New-TempRoot
try {
    $srcDir = Join-Path $tmp6 'src'
    New-Item -ItemType Directory -Path $srcDir -Force | Out-Null
    $src = Join-Path $srcDir 'a.pdf'
    Write-ValidPdf -QpdfExe $qpdfExe -Path $src
    $historyRoot = Join-Path $tmp6 'history'
    $cfg = @{
        dryRun = $true
        qcPrepend = @{
            historyRoot = $historyRoot
            tempRoot = (Join-Path $tmp6 'temp')
            outputRoot = (Join-Path $tmp6 'output')
            enableOverlay = $true
            overlayExePath = (Join-Path $tmp6 'missing.exe')
            overlayArgumentsTemplate = '"{sourcePdf}" "{historyPdf}" -o "{outputPdf}"'
            qpdfExePath = $qpdfExe
        }
    }
    $job = New-Job -Id 'dryrun' -SourcePdf $src
    $r = Invoke-QCPrependProcessor -Job $job -Config $cfg
    Assert-True $r.IsSuccess 'Dry-run should succeed'
    Assert-True (-not ([string]::IsNullOrWhiteSpace([string]$historyRoot))) 'Test setup: historyRoot should not be empty'
    Assert-True (-not (Test-Path -LiteralPath $historyRoot)) 'Dry-run should not create history root'
    Assert-True ($null -ne $r.Data.overlayCommandWouldRun) 'Dry-run should report overlayCommandWouldRun'
} finally { Remove-Item -LiteralPath $tmp6 -Recurse -Force -ErrorAction SilentlyContinue }

# overlay enabled + qc output pages (1 -> 2 -> 3) and history pages (1 -> 2 -> 3)
$tmp7 = New-TempRoot
try {
    $overlayExe = Join-Path $repoRoot 'dist\qc_overlay_prepend\qc_overlay_prepend.exe'
    if (-not (Test-Path -LiteralPath $overlayExe)) {
        $overlayExe = Join-Path $repoRoot 'qc_overlay_prepend.exe'
    }
    if (-not (Test-Path -LiteralPath $overlayExe)) {
        Write-Host "Skipping overlay integration test: overlay exe missing at $overlayExe" -ForegroundColor Yellow
    } else {
        $srcDir = Join-Path $tmp7 'src'
        New-Item -ItemType Directory -Path $srcDir -Force | Out-Null
        $useProvided = (($PdfV1 -as [string]) -and ($PdfV2 -as [string]) -and ($PdfV3 -as [string]))
        $v1 = $null
        $v2 = $null
        $v3 = $null
        if ($useProvided) {
            $v1 = [string]$PdfV1
            $v2 = [string]$PdfV2
            $v3 = [string]$PdfV3
            Assert-True (Test-Path -LiteralPath $v1) "Provided V1 not found: $v1"
            Assert-True (Test-Path -LiteralPath $v2) "Provided V2 not found: $v2"
            Assert-True (Test-Path -LiteralPath $v3) "Provided V3 not found: $v3"
        } else {
            $v1 = Join-Path $srcDir 'a.pdf'
            $v2 = Join-Path $srcDir 'a_v2.pdf'
            $v3 = Join-Path $srcDir 'a_v3.pdf'
            Write-OnePagePdfForQpdf -QpdfExe $qpdfExe -Path $v1
            Write-OnePagePdfForQpdf -QpdfExe $qpdfExe -Path $v2
            Write-OnePagePdfForQpdf -QpdfExe $qpdfExe -Path $v3
        }

        $historyRoot = Join-Path $tmp7 'history'
        $outputRoot = Join-Path $tmp7 'output'
        $cfg = @{
            dryRun = $false
            qcPrepend = @{
                historyRoot = $historyRoot
                tempRoot = (Join-Path $tmp7 'temp')
                outputRoot = $outputRoot
                enableOverlay = $true
                overlayExePath = $overlayExe
                overlayArgumentsTemplate = '"{sourcePdf}" "{historyPdf}" -o "{outputPdf}"'
                qpdfExePath = $qpdfExe
            }
        }

        # Run 1 (v1): qc should be 1 page; history should be 1 page
        $job1 = New-Job -Id 'ov' -SourcePdf $v1
        $r1 = Invoke-QCPrependProcessor -Job $job1 -Config $cfg
        Assert-True $r1.IsSuccess ("Overlay run 1 should succeed. code=$($r1.Code) msg=$($r1.Message)")
        $qc1 = Assert-NonEmptyPath $r1.Data.qcOutputPdf ("qcOutputPdf missing for run1. Data=" + ($r1.Data | ConvertTo-Json -Depth 8 -Compress))
        $hist1 = Assert-NonEmptyPath $r1.Data.targetHistoryPdf ("targetHistoryPdf missing for run1. Data=" + ($r1.Data | ConvertTo-Json -Depth 8 -Compress))
        Assert-True (Test-Path -LiteralPath $qc1) 'QC output should be created'
        Assert-Eq (Get-PdfPageCount -QpdfExe $qpdfExe -Path $r1.Data.qcOutputPdf) 1 'QC output pages after v1'
        Assert-Eq (Get-PdfPageCount -QpdfExe $qpdfExe -Path $r1.Data.targetHistoryPdf) 1 'History pages after v1'

        # Run 2 (v2): qc should be overlay+oldHistory => 2 pages; history => 2 pages
        $job2 = New-Job -Id 'ov' -SourcePdf $v2
        $job2.sourceName = $job1.sourceName
        $r2 = Invoke-QCPrependProcessor -Job $job2 -Config $cfg
        Assert-True $r2.IsSuccess ("Overlay run 2 should succeed. code=$($r2.Code) msg=$($r2.Message)")
        $qc2 = Assert-NonEmptyPath $r2.Data.qcOutputPdf ("qcOutputPdf missing for run2. Data=" + ($r2.Data | ConvertTo-Json -Depth 8 -Compress))
        $hist2 = Assert-NonEmptyPath $r2.Data.targetHistoryPdf ("targetHistoryPdf missing for run2. Data=" + ($r2.Data | ConvertTo-Json -Depth 8 -Compress))
        Assert-Eq (Get-PdfPageCount -QpdfExe $qpdfExe -Path $r2.Data.qcOutputPdf) 2 'QC output pages after v2'
        Assert-Eq (Get-PdfPageCount -QpdfExe $qpdfExe -Path $r2.Data.targetHistoryPdf) 2 'History pages after v2'

        # Run 3 (v3): qc should be overlay+oldHistory(2) => 3 pages; history => 3 pages
        $job3 = New-Job -Id 'ov' -SourcePdf $v3
        $job3.sourceName = $job1.sourceName
        $r3 = Invoke-QCPrependProcessor -Job $job3 -Config $cfg
        Assert-True $r3.IsSuccess ("Overlay run 3 should succeed. code=$($r3.Code) msg=$($r3.Message)")
        $qc3 = Assert-NonEmptyPath $r3.Data.qcOutputPdf ("qcOutputPdf missing for run3. Data=" + ($r3.Data | ConvertTo-Json -Depth 8 -Compress))
        $hist3 = Assert-NonEmptyPath $r3.Data.targetHistoryPdf ("targetHistoryPdf missing for run3. Data=" + ($r3.Data | ConvertTo-Json -Depth 8 -Compress))
        Assert-Eq (Get-PdfPageCount -QpdfExe $qpdfExe -Path $r3.Data.qcOutputPdf) 3 'QC output pages after v3'
        Assert-Eq (Get-PdfPageCount -QpdfExe $qpdfExe -Path $r3.Data.targetHistoryPdf) 3 'History pages after v3'
    }
} finally { Remove-Item -LiteralPath $tmp7 -Recurse -Force -ErrorAction SilentlyContinue }

$procText = Get-Content -LiteralPath (Resolve-ModuleImplPath -ModuleName 'QC.Processors.psm1') -Raw -Encoding UTF8
Assert-Contains $procText '_QCP-IsFinalQcPrependJob' 'Processors should detect QC Finalizing prepend jobs'
Assert-Contains $procText 'Test-QCPrependSkipReviewStamp' 'Processors should gate finalizing stamps by process type'
Assert-Contains $procText '-PrependTrigger' 'ProjectWise prepend must receive prependTrigger from job metadata'

$pwPrependText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\processing\Invoke-QCPrependPw.ps1') -Raw -Encoding UTF8
Assert-Contains $pwPrependText 'Test-PrependQcSkipReviewStamp' 'ProjectWise prepend must gate review stamp on finalQcComplete by lane'
Assert-Contains $pwPrependText 'QC Finalizing prepend (production lane)' 'Legacy prepend skip message for production final prepend'

Write-Host 'All QC_PREPEND processor tests passed.' -ForegroundColor Green

