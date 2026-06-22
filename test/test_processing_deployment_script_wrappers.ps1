# Assert Phase 4 processing/deployment script wrappers forward to subfolder targets.
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptsDir = Join-Path $repoRoot 'scripts'

$expected = @(
    @{ Name = 'Combine-StatusSet.ps1'; Subfolder = 'processing' }
    @{ Name = 'Run-CombineStatusSet.ps1'; Subfolder = 'processing' }
    @{ Name = 'Promote-DevToMain.ps1'; Subfolder = 'deployment' }
    @{ Name = 'Sync-OverlayReviewStamp.ps1'; Subfolder = 'deployment' }
)

$fail = 0
foreach ($item in $expected) {
    $name = $item.Name
    $sub = $item.Subfolder
    $wrapperPath = Join-Path $scriptsDir $name
    $targetPath = Join-Path $scriptsDir (Join-Path $sub $name)
    $relTarget = "$sub\$name"

    if (-not (Test-Path -LiteralPath $wrapperPath)) {
        Write-Host "FAIL: missing wrapper $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        Write-Host "FAIL: missing target $targetPath" -ForegroundColor Red
        $fail++
        continue
    }

    $wrapperText = Get-Content -LiteralPath $wrapperPath -Raw
    if ($wrapperText -notmatch [regex]::Escape($relTarget)) {
        Write-Host "FAIL: wrapper does not reference $relTarget : $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }

    $resolved = Join-Path $scriptsDir $relTarget
    if (-not (Test-Path -LiteralPath $resolved)) {
        Write-Host "FAIL: wrapper target not resolvable: $resolved" -ForegroundColor Red
        $fail++
        continue
    }

    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($wrapperPath, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        Write-Host "FAIL: wrapper parse errors in $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }

    Write-Host "OK:   $name -> $relTarget" -ForegroundColor Green
}

if ($fail -gt 0) {
    Write-Host "test_processing_deployment_script_wrappers.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host 'All processing/deployment script wrapper checks passed (4 scripts).' -ForegroundColor Green
