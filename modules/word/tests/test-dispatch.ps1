param([string]$OutputDirectory, [switch]$MutationCheck)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repo ('artifacts/word/dispatch-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Choose a new dispatch evidence directory.' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$sources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../src') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$testSource = Join-Path $PSScriptRoot 'WordDispatchTests.cs'
$contract = Join-Path $repo 'modules/review/selective-editor-contract.json'
$inputs = @(@($sources) + @($testSource,$contract,$PSCommandPath) | ForEach-Object {
    [ordered]@{path=[IO.Path]::GetRelativePath($repo,$_);sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()}
})
$compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$arguments = @('/nologo','/target:exe','/langversion:5','/r:System.Windows.Forms.dll','/r:System.Drawing.dll',
    '/r:System.Web.Extensions.dll','/r:System.Security.dll','/r:Microsoft.CSharp.dll',"/resource:$contract,SelectiveEditorContract")
$test = Join-Path $OutputDirectory 'WordDispatchTests.exe'
& $compiler @arguments "/out:$test" @sources $testSource *> (Join-Path $OutputDirectory 'compile.log')
if ($LASTEXITCODE -ne 0) { throw 'Dispatch harness compilation failed; inspect compile.log.' }
& $test (Join-Path $OutputDirectory 'actual') 2>&1 | Tee-Object -FilePath (Join-Path $OutputDirectory 'dispatch.log')
$actualCode = $LASTEXITCODE
$mutationStatus = 'NOT RUN'
if ($actualCode -eq 0 -and $MutationCheck) {
    $mutationDirectory = Join-Path $OutputDirectory 'mutation-cancellation-disabled'
    New-Item -ItemType Directory -Path $mutationDirectory | Out-Null
    $original = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '../src/ComAddin.cs'))
    $cancel = 'bool started = Interlocked.CompareExchange(ref pending, 2, 0) != 0;'
    $clear = 'if (!started) operation = null;'
    if ([regex]::Matches($original,[regex]::Escape($cancel)).Count -ne 1 -or [regex]::Matches($original,[regex]::Escape($clear)).Count -ne 1) {
        throw 'The copied cancellation mutation needs updating for the current dispatch implementation.'
    }
    $mutantSource = Join-Path $mutationDirectory 'ComAddin.cs'
    [IO.File]::WriteAllText($mutantSource,$original.Replace($cancel,'bool started = System.Threading.Volatile.Read(ref pending) != 0;').Replace($clear,'if (!started) { }'))
    $mutantSources = @($sources | Where-Object {[IO.Path]::GetFileName($_) -ne 'ComAddin.cs'}) + @($mutantSource)
    $mutantTest = Join-Path $mutationDirectory 'WordDispatchTests.exe'
    & $compiler @arguments "/out:$mutantTest" @mutantSources $testSource *> (Join-Path $mutationDirectory 'compile.log')
    if ($LASTEXITCODE -ne 0) { throw 'The copied dispatch mutation failed to compile.' }
    & $mutantTest (Join-Path $mutationDirectory 'actual') *> (Join-Path $mutationDirectory 'dispatch.log')
    $mutantCode = $LASTEXITCODE
    $mutationStatus = if ($mutantCode -eq 1 -and ([IO.File]::ReadAllText((Join-Path $mutationDirectory 'dispatch.log'))).Contains('A canceled queued dispatch never runs after the STA is unblocked.')) {'PASS'} else {'FAIL'}
}
$changed = @($inputs | Where-Object { (Get-FileHash -LiteralPath (Join-Path $repo $_.path)).Hash.ToLowerInvariant() -ne $_.sha256 })
$status = if ($actualCode -eq 0 -and $changed.Count -eq 0 -and $mutationStatus -ne 'FAIL') {'PASS'} else {'FAIL'}
[ordered]@{status=$status;exitCode=$actualCode;mutation=$mutationStatus;productionTimeoutSeconds=15;inputs=$inputs;changedInputs=$changed;harnessSha256=(Get-FileHash -LiteralPath $test).Hash.ToLowerInvariant();realWord='NOT STARTED';installedPackage='NOT RUN'} | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'manifest.json') -Encoding utf8
if ($status -ne 'PASS') { throw 'Dispatch recovery checks failed; inspect the native WinForms evidence.' }
Write-Output 'PASS: actual 15-second dispatch recovery checks and any requested copied-source mutation proof.'
