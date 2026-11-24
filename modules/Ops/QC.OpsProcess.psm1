# Windows PowerShell child execution for the operator console.
Import-Module (Join-Path $PSScriptRoot '..\Core\Core.Results.psm1') -DisableNameChecking

function ConvertTo-QCOpsWindowsArgument {
    param([AllowEmptyString()][string]$Value)
    '"' + (($Value -replace '(\\*)"', '$1$1\"') -replace '(\\+)$', '$1$1') + '"'
}

function Invoke-QCOpsChildProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [string[]]$ArgumentList = @(),
        [hashtable]$Parameters,
        [ValidateRange(1,2147483647)][int]$TimeoutMs = 300000,
        [System.Collections.Concurrent.ConcurrentQueue[string]]$OutputQueue
    )
    if (-not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) {
        return New-QCFailureResult -Code 'OPS_CHILD_MISSING' -Message "Script not found: $ScriptPath"
    }
    $argv = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-MTA')
    if ($null -ne $Parameters) {
        # Bind named arguments in PowerShell. Values never enter executable source.
        $payload = @{ path = $ScriptPath; parameters = $Parameters } | ConvertTo-Json -Depth 8 -Compress
        $data = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($payload))
        $command = @"
`$ErrorActionPreference = 'Stop'
`$ProgressPreference = 'SilentlyContinue'
`$payload = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$data')) | ConvertFrom-Json
`$parameters = @{}
foreach (`$property in `$payload.parameters.PSObject.Properties) { `$parameters[`$property.Name] = `$property.Value }
`$global:LASTEXITCODE = 0
& `$payload.path @parameters
if (-not `$?) { exit 1 }
exit `$global:LASTEXITCODE
"@
        $argv += @('-EncodedCommand', [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command)))
    } else {
        $argv += @('-File', $ScriptPath) + @($ArgumentList)
    }
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = Join-Path $PSHOME 'powershell.exe'
    if (-not (Test-Path -LiteralPath $psi.FileName)) {
        $psi.FileName = Join-Path ([Environment]::GetFolderPath('System')) 'WindowsPowerShell\v1.0\powershell.exe'
    }
    $psi.Arguments = ($argv | ForEach-Object { ConvertTo-QCOpsWindowsArgument $_ }) -join ' '
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $psi
    $lines = [System.Collections.Generic.Queue[string]]::new()
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $resultData = @{ script = $ScriptPath }
    try {
        if (-not $process.Start()) { throw 'Windows PowerShell did not start.' }
        $streams = @(
            @{ reader = $process.StandardOutput; prefix = ''; next = $process.StandardOutput.ReadLineAsync() }
            @{ reader = $process.StandardError; prefix = '[stderr] '; next = $process.StandardError.ReadLineAsync() }
        )
        do {
            $gotLine = $false
            foreach ($stream in $streams) {
                for ($n = 0; $n -lt 64 -and $stream.next -and $stream.next.IsCompleted; $n++) {
                    $line = $stream.next.GetAwaiter().GetResult()
                    if ($null -eq $line) { $stream.next = $null; break }
                    $stream.next = $stream.reader.ReadLineAsync()
                    if ($line -eq '#< CLIXML') { continue }
                    if ($line.StartsWith('<Objs')) {
                        try {
                            $records = @(([xml]$line).SelectNodes("//*[local-name()='S' and @S]"))
                            $line = ($records | ForEach-Object { [System.Xml.XmlConvert]::DecodeName($_.InnerText) }) -join ''
                        } catch { }
                    }
                    $gotLine = $true
                    $message = $stream.prefix + $line
                    $lines.Enqueue($message)
                    if ($lines.Count -gt 2000) { [void]$lines.Dequeue() }
                    if ($OutputQueue) {
                        $OutputQueue.Enqueue($message)
                        if ($OutputQueue.Count -gt 2000) {
                            $dropped = ''
                            [void]$OutputQueue.TryDequeue([ref]$dropped)
                        }
                    }
                }
            }
            if ($process.HasExited -and -not $streams[0].next -and -not $streams[1].next) { break }
            if (-not $gotLine) { Start-Sleep -Milliseconds 25 }
        } while ($clock.ElapsedMilliseconds -lt $TimeoutMs)

        if (-not $process.HasExited -or $streams[0].next -or $streams[1].next) {
            return New-QCFailureResult -Code 'OPS_CHILD_TIMEOUT' -Message "Timed out after $([math]::Round($TimeoutMs / 1000, 1)) seconds. Check completion before retrying." -Data $resultData
        }
        $resultData.exitCode = $process.ExitCode
        if ($process.ExitCode -ne 0) {
            return New-QCFailureResult -Code 'OPS_CHILD_FAILED' -Message "Script exited with code $($process.ExitCode)." -Data $resultData
        }
        return New-QCSuccessResult -Code 'OPS_CHILD_OK' -Message 'Script completed.' -Data $resultData
    } catch {
        return New-QCFailureResult -Code 'OPS_CHILD_ERROR' -Message $_.Exception.Message -Data $resultData
    } finally {
        $resultData.output = $lines.ToArray() -join [Environment]::NewLine
        $resultData.elapsedMs = $clock.ElapsedMilliseconds
        try {
            if (-not $process.HasExited) {
                $kill = Join-Path ([Environment]::GetFolderPath('System')) 'taskkill.exe'
                try { & $kill /PID $process.Id /T /F 2>&1 | Out-Null } catch { }
                if (-not $process.HasExited) {
                    $process.Kill()
                    [void]$process.WaitForExit(1000)
                }
            }
        } catch { }
        $process.Dispose()
    }
}

Export-ModuleMember -Function Invoke-QCOpsChildProcess, ConvertTo-QCOpsWindowsArgument
