param([string]$OutputDirectory, [string]$WordExecutable, [switch]$ProtectedOnly, [switch]$DocumentChoiceOnly, [switch]$ConcurrentClientsOnly, [switch]$DirectCommentsOnly)
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
if (@($ProtectedOnly, $DocumentChoiceOnly, $ConcurrentClientsOnly, $DirectCommentsOnly).Where({ $_ }).Count -gt 1) { throw 'Choose only one focused native scenario.' }
if ($ProtectedOnly) { $nativeArguments += 'protected-only' }
if ($DocumentChoiceOnly) { $nativeArguments += 'document-choice-only' }
if ($ConcurrentClientsOnly) { $nativeArguments += 'concurrency-only' }
if ($DirectCommentsOnly) { $nativeArguments += 'direct-comments-only' }
& $test @nativeArguments 2>&1 | Tee-Object -FilePath (Join-Path $OutputDirectory 'native.log')
$code = $LASTEXITCODE
$changed = @($inputs | Where-Object { (Get-FileHash -LiteralPath (Join-Path $repo $_.path)).Hash.ToLowerInvariant() -ne $_.sha256 })
[ordered]@{status=$(if($code -eq 0 -and $changed.Count -eq 0){'PASS'}else{'FAIL'});exitCode=$code;inputs=$inputs;changedInputs=$changed;harnessSha256=(Get-FileHash -LiteralPath $test).Hash.ToLowerInvariant();wordSha256=(Get-FileHash -LiteralPath $WordExecutable).Hash.ToLowerInvariant();normalWord='NOT TOUCHED';installedPackage='NOT RUN';publicTunnel='NOT RUN'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'manifest.json') -Encoding utf8
if ($changed.Count -ne 0) { throw 'Source changed during native suggestion qualification.' }
if ($code -ne 0) { throw 'Isolated native suggestion qualification failed.' }
