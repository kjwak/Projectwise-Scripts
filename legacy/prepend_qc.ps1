# Deprecated compatibility shim - forwards to scripts/processing/Invoke-QCPrependPw.ps1 (Phase 4B).
Write-Warning 'legacy/prepend_qc.ps1 is deprecated; use scripts/processing/Invoke-QCPrependPw.ps1'

$target = Join-Path $PSScriptRoot '..\scripts\processing\Invoke-QCPrependPw.ps1'
if (-not (Test-Path -LiteralPath $target)) {
    throw "ProjectWise prepend script not found: $target"
}

$passArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-MTA', '-File', (Resolve-Path -LiteralPath $target).Path) + $args
& powershell.exe @passArgs
exit $LASTEXITCODE
