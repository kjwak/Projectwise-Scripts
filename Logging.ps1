# Logging.ps1
# Shared logging for prepend_qc scripts. Dot-source after setting $PrependQc_LogDir (optional).
# Usage:
#   $PrependQc_LogDir = "C:\PW_QC_LOCAL\logs"  # optional; default used if not set
#   . "$PSScriptRoot\Logging.ps1"
#   Write-Log "Activity message"
#   Write-Log "Something went wrong" -Severity WARNING
#   Write-Log "Fatal error" -Severity ERROR

if (-not (Get-Variable -Name PrependQc_LogDir -ErrorAction SilentlyContinue)) {
  $PrependQc_LogDir = "C:\PW_QC_LOCAL\logs"
}

function Write-Log {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string] $Msg,
    [ValidateSet('INFO', 'WARNING', 'ERROR')]
    [string] $Severity = 'INFO'
  )
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $line = "$ts  [$Severity]  $Msg"

  Write-Host $line

  $logDir = $PrependQc_LogDir
  if (-not $logDir) { return }

  try {
    if (-not (Test-Path -LiteralPath $logDir)) {
      New-Item -ItemType Directory -Force -Path $logDir -ErrorAction Stop | Out-Null
    }
    $dateStr = (Get-Date).ToString("yyyy-MM-dd")
    $activityLog = Join-Path $logDir "prepend_qc_$dateStr.log"
    $errorLog = Join-Path $logDir "prepend_qc_errors_$dateStr.log"

    Add-Content -LiteralPath $activityLog -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue

    if ($Severity -in 'WARNING', 'ERROR') {
      Add-Content -LiteralPath $errorLog -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    }
  } catch {
    Write-Host "Log write failed: $_"
  }
}
