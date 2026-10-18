param([string]$OutputDirectory, [string]$WordExecutable, [switch]$ProtectedOnly, [switch]$DocumentChoiceOnly, [switch]$ConcurrentClientsOnly, [switch]$DirectCommentsOnly, [switch]$BackgroundSaveOnly, [switch]$MutationCheck)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repo ('artifacts/word/native-suggestions-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Choose a new suggestion evidence directory.' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
if (-not $WordExecutable) { $WordExecutable = Join-Path $env:ProgramFiles 'Microsoft Office/root/Office16/WINWORD.EXE' }
$sources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../src') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$testSources = @((Join-Path $PSScriptRoot 'NativeSuggestionWorkflowSmoke.cs'), (Join-Path $PSScriptRoot 'WordAutomation.cs'),
    (Join-Path $PSScriptRoot 'NativeMcpClient.cs'), (Join-Path $PSScriptRoot 'NativeConcurrentClients.cs'))
$contract = Join-Path $repo 'modules/review/selective-editor-contract.json'
$inputs = @(@($sources) + $testSources + @($contract, $PSCommandPath) | ForEach-Object { [ordered]@{path=[IO.Path]::GetRelativePath($repo,$_);sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()} })
$test = Join-Path $OutputDirectory 'NativeSuggestionWorkflowSmoke.exe'
& "$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe" /nologo /target:exe /langversion:5 "/out:$test" /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.Security.dll /r:Microsoft.CSharp.dll "/resource:$contract,SelectiveEditorContract" @sources @testSources
if ($LASTEXITCODE -ne 0) { throw 'Native suggestion harness compilation failed.' }
$nativeArguments = @($WordExecutable, (Join-Path $OutputDirectory 'native'))
if (@($ProtectedOnly, $DocumentChoiceOnly, $ConcurrentClientsOnly, $DirectCommentsOnly, $BackgroundSaveOnly).Where({ $_ }).Count -gt 1) { throw 'Choose only one focused native scenario.' }
if ($MutationCheck -and -not $BackgroundSaveOnly) { throw 'The timer mutation requires -BackgroundSaveOnly.' }
if ($ProtectedOnly) { $nativeArguments += 'protected-only' }
if ($DocumentChoiceOnly) { $nativeArguments += 'document-choice-only' }
if ($ConcurrentClientsOnly) { $nativeArguments += 'concurrency-only' }
if ($DirectCommentsOnly) { $nativeArguments += 'direct-comments-only' }
if ($BackgroundSaveOnly) { $nativeArguments += 'background-save-only' }
& $test @nativeArguments 2>&1 | Tee-Object -FilePath (Join-Path $OutputDirectory 'native.log')
$code = $LASTEXITCODE
$timerMutation = 'NOT RUN'
if ($MutationCheck -and $code -eq 0) {
    $mutationDirectory = Join-Path $OutputDirectory 'timer-disabled-mutation'
    New-Item -ItemType Directory -Path $mutationDirectory | Out-Null
    $original = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '../src/WordAutoSave.cs'))
    if ([regex]::Matches($original, [regex]::Escape('timer.Start();')).Count -ne 1) { throw 'Review the timer-start mutation for the current implementation.' }
    $mutantSource = Join-Path $mutationDirectory 'WordAutoSave.cs'
    [IO.File]::WriteAllText($mutantSource, $original.Replace('timer.Start();', 'timer.Stop();'))
    $mutantSources = @($sources | Where-Object { [IO.Path]::GetFileName($_) -ne 'WordAutoSave.cs' }) + @($mutantSource)
    $mutant = Join-Path $mutationDirectory 'NativeSuggestionWorkflowSmoke.exe'
    & "$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe" /nologo /target:exe /langversion:5 "/out:$mutant" /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.Security.dll /r:Microsoft.CSharp.dll "/resource:$contract,SelectiveEditorContract" @mutantSources @testSources
    if ($LASTEXITCODE -ne 0) { throw 'The copied disabled-timer native control did not compile.' }
    & $mutant $WordExecutable (Join-Path $mutationDirectory 'native') 'background-save-only' *> (Join-Path $mutationDirectory 'native.log')
    $mutantCode = $LASTEXITCODE
    $timerMutation = if ($mutantCode -eq 1 -and ([IO.File]::ReadAllText((Join-Path $mutationDirectory 'native.log'))).Contains('The production inactive-document coordinator completes a confirmed native save:')) { 'PASS' } else { 'FAIL' }
}
$changed = @($inputs | Where-Object { (Get-FileHash -LiteralPath (Join-Path $repo $_.path)).Hash.ToLowerInvariant() -ne $_.sha256 })
[ordered]@{status=$(if($code -eq 0 -and $changed.Count -eq 0 -and $timerMutation -ne 'FAIL'){'PASS'}else{'FAIL'});exitCode=$code;inputs=$inputs;changedInputs=$changed;timerMutation=$timerMutation;harnessSha256=(Get-FileHash -LiteralPath $test).Hash.ToLowerInvariant();wordSha256=(Get-FileHash -LiteralPath $WordExecutable).Hash.ToLowerInvariant();normalWord='NOT TOUCHED';installedPackage='NOT RUN';publicTunnel='NOT RUN'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'manifest.json') -Encoding utf8
if ($changed.Count -ne 0) { throw 'Source changed during native suggestion qualification.' }
if ($code -ne 0) { throw 'Isolated native suggestion qualification failed.' }
if ($timerMutation -eq 'FAIL') { throw 'The native background save check did not detect a disabled production timer.' }
