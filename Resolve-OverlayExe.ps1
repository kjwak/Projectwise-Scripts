# Shared helpers: validate qc_overlay_prepend.exe and optionally copy to TEMP when the source path reads as zeros (cloud/AV quirks).
# Dot-source from prepend_qc.ps1 or test scripts: . "$PSScriptRoot\Resolve-OverlayExe.ps1"

# PyInstaller spec (overlay\build_overlay_exe.ps1) produces onedir: dist\qc_overlay_prepend\qc_overlay_prepend.exe (+ _internal). Optional onefile: dist\qc_overlay_prepend.exe
function Get-DefaultOverlayExeCandidates([string]$scriptRoot) {
  return @(
    (Join-Path $scriptRoot "dist\qc_overlay_prepend\qc_overlay_prepend.exe")
    (Join-Path $scriptRoot "dist\qc_overlay_prepend.exe")
  )
}

function Select-ExistingOverlayExePath([string]$scriptRoot) {
  foreach ($p in (Get-DefaultOverlayExeCandidates $scriptRoot)) {
    if (Test-Path -LiteralPath $p) { return $p }
  }
  return $null
}

function Get-HexPrefix([byte[]]$bytes, [int]$count) {
  $n = [Math]::Min($count, $bytes.Length)
  $parts = for ($i = 0; $i -lt $n; $i++) { '{0:X2}' -f $bytes[$i] }
  return ($parts -join ' ')
}

function Read-OverlayExeHeaderSample([string]$path) {
  $len = (Get-Item -LiteralPath $path).Length
  $fs = [System.IO.File]::OpenRead($path)
  try {
    $toRead = [Math]::Min(4096, [int]$len)
    if ($toRead -lt 2) {
      return @{ Ok = $false; Err = "Overlay exe too small ($len bytes): $path" }
    }
    $buf = New-Object byte[] $toRead
    $read = $fs.Read($buf, 0, $toRead)
    if ($read -lt 2) {
      return @{ Ok = $false; Err = "Could only read $read byte(s) from overlay exe (file locked or incomplete). Path: $path" }
    }
    return @{ Ok = $true; Buf = $buf; Read = $read; Len = $len }
  } finally {
    $fs.Dispose()
  }
}

function Test-OverlayHeaderBytes([byte[]]$buf, [int]$read, [long]$len, [string]$path) {
  $hex = Get-HexPrefix $buf 16
  if ($buf[0] -eq 0x4D -and $buf[1] -eq 0x5A) {
    return @{ Ok = $true }
  }
  $head = [System.Text.Encoding]::ASCII.GetString($buf[0..([Math]::Min(199, $read - 1))])
  if ($head -match 'git-lfs|oid sha256') {
    return @{ Ok = $false; Err = "Overlay exe is a Git LFS pointer, not the real binary. Run 'git lfs pull' or copy qc_overlay_prepend.exe from a build machine. Path: $path" }
  }
  if ($read -ge 4 -and $buf[0] -eq 0x7F -and $buf[1] -eq 0x45 -and $buf[2] -eq 0x4C -and $buf[3] -eq 0x46) {
    return @{ Ok = $false; Err = "Overlay exe is ELF (Linux), not Windows. Run overlay\build_overlay_exe.ps1 on a Windows machine and copy that qc_overlay_prepend.exe. First bytes: $hex Path: $path" }
  }
  if ($read -ge 4 -and $buf[0] -eq 0xCF -and $buf[1] -eq 0xFA -and $buf[2] -eq 0xED -and $buf[3] -eq 0xFE) {
    return @{ Ok = $false; Err = "Overlay exe is Mach-O (macOS), not Windows. Build with PyInstaller on Windows and deploy that .exe. First bytes: $hex Path: $path" }
  }
  if ($read -ge 4 -and $buf[0] -eq 0xFE -and $buf[1] -eq 0xED -and $buf[2] -eq 0xFA -and $buf[3] -eq 0xCE) {
    return @{ Ok = $false; Err = "Overlay exe is Mach-O (macOS), not Windows. Build with PyInstaller on Windows and deploy that .exe. First bytes: $hex Path: $path" }
  }
  $sampleEnd = [Math]::Min(512, $read) - 1
  $nonZero = 0
  for ($i = 0; $i -le $sampleEnd; $i++) { if ($buf[$i] -ne 0) { $nonZero++ } }
  if ($nonZero -eq 0) {
    return @{ Ok = $false; AllZeros = $true; Hex = $hex; Len = $len }
  }
  return @{ Ok = $false; Err = "Overlay exe does not look like a Windows PE (expected MZ at start). First bytes: $hex size=$len bytes Path: $path Rebuild with .\overlay\build_overlay_exe.ps1 on Windows, or copy a Windows-built qc_overlay_prepend.exe." }
}

function Write-OverlayResolveLog([string]$message) {
  if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
    Write-Log $message
  } else {
    Write-Host "[INFO] $message"
  }
}

# Returns path to invoke (original or TEMP copy when first read returned zeros; Copy-Item often yields real bytes).
function Resolve-OverlayExePath([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) {
    throw "Overlay exe not found: $path"
  }
  $s = Read-OverlayExeHeaderSample $path
  if (-not $s.Ok) { throw $s.Err }
  $v = Test-OverlayHeaderBytes $s.Buf $s.Read $s.Len $path
  if ($v.Ok) {
    return $path
  }
  if ($v.Err) {
    throw $v.Err
  }
  if ($v.AllZeros) {
    $tempDir = Join-Path $env:TEMP "PW_QC"
    if (-not (Test-Path -LiteralPath $tempDir)) {
      New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
    }
    $tempExe = Join-Path $tempDir "qc_overlay_prepend.exe"
    Copy-Item -LiteralPath $path -Destination $tempExe -Force
    $s2 = Read-OverlayExeHeaderSample $tempExe
    if (-not $s2.Ok) { throw $s2.Err }
    $v2 = Test-OverlayHeaderBytes $s2.Buf $s2.Read $s2.Len $tempExe
    if ($v2.Ok) {
      Write-OverlayResolveLog "Overlay exe first bytes were zero at source; using local copy: $tempExe"
      return $tempExe
    }
    if ($v2.Err) { throw $v2.Err }
    if ($v2.AllZeros) {
      $hz = $v2.Hex
      throw @"
Overlay exe reads as all zeros (first bytes: $hz) but reports size $($s.Len) bytes - file is not a valid PE (sparse placeholder, bad copy, or corrupt). Delete dist\qc_overlay_prepend.exe if it is wrong, then run .\overlay\build_overlay_exe.ps1 on Windows; it builds dist\qc_overlay_prepend\qc_overlay_prepend.exe (keep the whole dist\qc_overlay_prepend folder with _internal). Or pass -QcOverlayExe to a known-good exe under e.g. C:\Tools\.
Source: $path
"@
    }
    throw "Resolve-OverlayExePath: unexpected validation state for $tempExe"
  }
  throw "Resolve-OverlayExePath: unexpected state for $path"
}
