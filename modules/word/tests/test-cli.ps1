param([Parameter(Mandatory)][string]$CliPath, [Parameter(Mandatory)][string]$OutputDirectory, [switch]$MutationCheck)
$ErrorActionPreference = 'Stop'
$CliPath = (Resolve-Path -LiteralPath $CliPath).Path
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Choose a fresh CLI evidence directory.' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$checks = [Collections.Generic.List[string]]::new()
function Check-Command([string]$Name, [string[]]$Arguments, [int]$ExitCode, [string]$Contains, [switch]$HelpText) {
    $start = [Diagnostics.ProcessStartInfo]::new($CliPath)
    $start.UseShellExecute = $false; $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
    $start.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
    foreach ($argument in $Arguments) { $start.ArgumentList.Add($argument) }
    $process = [Diagnostics.Process]::Start($start)
    $stdout = $process.StandardOutput.ReadToEndAsync(); $stderr = $process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit(20000)) { $process.Kill(); throw "CLI process timed out: $Name" }
    $text = $stdout.GetAwaiter().GetResult(); $errors = $stderr.GetAwaiter().GetResult()
    $code = $process.ExitCode; $process.Dispose()
    [IO.File]::WriteAllText((Join-Path $OutputDirectory ($Name+'.json')), $text, [Text.UTF8Encoding]::new($false))
    if ($code -ne $ExitCode -or -not $text.Contains($Contains) -or $errors) { throw "CLI case failed: $Name (exit $code; expected $ExitCode). $text $errors" }
    if (-not $HelpText) { $reply = $text | ConvertFrom-Json; if ($reply.ok -ne $false) { throw "Expected JSON error: $Name" } }
    $checks.Add($Name)
}
try {
    Check-Command 'help' @('help') 0 'wordplai <command>' -HelpText
    Check-Command 'unknown-command' @('not-a-command') 2 'Unknown command'
    Check-Command 'missing-target' @('state') 2 'explicit --document'
    Check-Command 'duplicate-option' @('state','--document','A','--document','B') 2 'Duplicate option'
    Check-Command 'missing-option-value' @('state','--document') 2 'missing value'
    Check-Command 'invalid-pid' @('documents','--word-pid','-1') 2 'positive process ID'
    foreach ($entry in @(@{name='array';text='[]'}, @{name='null';text='null'}, @{name='malformed';text='{'})) {
        $file = Join-Path $OutputDirectory ($entry.name+'-input.json')
        [IO.File]::WriteAllText($file, $entry.text, [Text.UTF8Encoding]::new($false))
        Check-Command $entry.name @('state','--document','A','--input',$file) 2 'JSON object'
    }
    $override = Join-Path $OutputDirectory 'override-input.json'
    [IO.File]::WriteAllText($override, '{"documentId":"B"}', [Text.UTF8Encoding]::new($false))
    Check-Command 'target-override' @('state','--document','A','--input',$override) 2 'cannot override'
    $large = Join-Path $OutputDirectory 'oversized-input.json'
    [IO.File]::WriteAllText($large, ('{"text":"'+('x'*131062)+'"}'), [Text.UTF8Encoding]::new($false))
    Check-Command 'oversized-input' @('state','--document','A','--input',$large,'--word-pid','2147483647') 2 'exceeds 131072'
    Check-Command 'missing-mcp-session' @('mcp','--method','tools/list') 2 'session-file'
    Check-Command 'no-pid-fallback' @('documents','--word-pid','2147483647') 3 'not running'
    if ($MutationCheck) {
        $mutant = Join-Path $OutputDirectory 'oversized-mutant'
        New-Item -ItemType Directory -Path $mutant | Out-Null
        $source = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '../cli/Program.cs'))
        $needle = 'if (used > 131072)'
        if (-not $source.Contains($needle)) { throw 'Input-limit mutation target is missing.' }
        $program = Join-Path $mutant 'Program.cs'
        [IO.File]::WriteAllText($program, $source.Replace($needle,'if (used > 131073)'), [Text.UTF8Encoding]::new($false))
        $binary = Join-Path $mutant 'wordplai.exe'
        $compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
        & $compiler /nologo /target:exe /langversion:5 "/out:$binary" /r:System.dll /r:System.Core.dll /r:System.Web.Extensions.dll /r:Microsoft.CSharp.dll $program (Join-Path $PSScriptRoot '../cli/WordConnection.cs')
        if ($LASTEXITCODE -ne 0) { throw 'CLI mutant compilation failed.' }
        $evidence = Join-Path $mutant 'checks'
        & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $PSCommandPath -CliPath $binary -OutputDirectory $evidence *> (Join-Path $mutant 'run.log')
        $negative = $LASTEXITCODE
        $result = Get-Content -Raw (Join-Path $evidence 'result.json') | ConvertFrom-Json
        if ($negative -eq 0 -or $result.status -ne 'FAIL' -or $result.error -notlike '*oversized-input*') { throw 'Oversized CLI input mutant did not fail the intended predicate.' }
        $checks.Add('effective-input-limit-negative-control')
    }
    [ordered]@{status='PASS';checks=$checks;cliSha256=(Get-FileHash -LiteralPath $CliPath).Hash;nativeWord='NOT RUN'} | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $OutputDirectory 'result.json')
    Write-Output "PASS: $($checks.Count) CLI command checks."
} catch {
    [ordered]@{status='FAIL';checks=$checks;error=$_.Exception.Message} | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $OutputDirectory 'result.json')
    throw
}
exit 0
