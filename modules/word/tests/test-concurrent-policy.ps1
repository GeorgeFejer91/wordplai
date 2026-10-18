param([string]$OutputDirectory, [switch]$MutationCheck)
$ErrorActionPreference = 'Stop'
$wordRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $wordRoot ('artifacts/word/scoped-policy-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$wordOut = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $wordOut) { throw 'Choose fresh scoped-policy evidence.' }
New-Item -ItemType Directory -Path $wordOut | Out-Null
$wordCompiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$wordSources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../src') -Filter 'Policy*.cs' | Sort-Object Name | ForEach-Object FullName)
$wordTestSource = Join-Path $PSScriptRoot 'PolicyConcurrencyTests.cs'
$wordContract = Join-Path $wordRoot 'modules/review/selective-editor-contract.json'
$wordInputs = @(@($wordSources) + @($wordTestSource, $wordContract, $PSCommandPath) | ForEach-Object {
    [ordered]@{path=[IO.Path]::GetRelativePath($wordRoot,$_);sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()}
})
$wordExecutable = Join-Path $wordOut 'PolicyConcurrencyTests.exe'
& $wordCompiler /nologo /target:exe /langversion:5 "/out:$wordExecutable" /r:System.Web.Extensions.dll @wordSources $wordTestSource *> (Join-Path $wordOut 'compile.log')
if ($LASTEXITCODE -ne 0) { throw 'Scoped policy tests did not compile.' }
& $wordExecutable $wordContract 2>&1 | Tee-Object -FilePath (Join-Path $wordOut 'policy.log')
if ($LASTEXITCODE -ne 0) { throw 'Scoped policy tests failed.' }
if ($MutationCheck) {
    $wordMutation = Join-Path $wordOut 'cross-client-mutation'
    New-Item -ItemType Directory -Path $wordMutation | Out-Null
    $wordOriginalPolicy = Join-Path $wordRoot 'modules/word/src/Policy.cs'
    $wordSource = [IO.File]::ReadAllText($wordOriginalPolicy)
    $wordNeedle = 'if (!contextReads.TryGetValue(scope, out contextRead) || contextRead.Id != contextId || !contextRead.Matches(current))'
    if (-not $wordSource.Contains($wordNeedle)) { throw 'Cross-client witness mutation no longer matches production.' }
    $wordMutatedPolicy = Join-Path $wordMutation 'Policy.cs'
    [IO.File]::WriteAllText($wordMutatedPolicy, $wordSource.Replace($wordNeedle, 'if ((contextRead = contextReads.Values.FirstOrDefault(witness => witness.Id == contextId)) == null || !contextRead.Matches(current))'))
    $wordMutationSources = @($wordSources | Where-Object { [IO.Path]::GetFileName($_) -ne 'Policy.cs' }) + @($wordMutatedPolicy)
    $wordMutationExecutable = Join-Path $wordMutation 'PolicyConcurrencyTests.exe'
    & $wordCompiler /nologo /target:exe /langversion:5 "/out:$wordMutationExecutable" /r:System.Web.Extensions.dll @wordMutationSources $wordTestSource *> (Join-Path $wordMutation 'compile.log')
    if ($LASTEXITCODE -ne 0) { throw 'Cross-client witness mutation did not compile.' }
    & $wordMutationExecutable $wordContract *> (Join-Path $wordMutation 'policy.log')
    if ($LASTEXITCODE -eq 0 -or -not ([IO.File]::ReadAllText((Join-Path $wordMutation 'policy.log')).Contains('A complete-read witness cannot cross client scopes.'))) {
        throw 'The policy test did not reject the intended cross-client witness bypass.'
    }
    Write-Output 'PASS: scoped-policy test rejects cross-client witness authorization in an isolated copied build.'
    $wordContextMutation = Join-Path $wordOut 'request-context-mutation'
    New-Item -ItemType Directory -Path $wordContextMutation | Out-Null
    $wordContextNeedle = ' || existing.RequestContextHash != requestContextHash'
    if (-not $wordSource.Contains($wordContextNeedle)) { throw 'Request-context replay mutation no longer matches production.' }
    $wordContextPolicy = Join-Path $wordContextMutation 'Policy.cs'
    [IO.File]::WriteAllText($wordContextPolicy, $wordSource.Replace($wordContextNeedle, ''))
    $wordContextSources = @($wordSources | Where-Object { [IO.Path]::GetFileName($_) -ne 'Policy.cs' }) + @($wordContextPolicy)
    $wordContextExecutable = Join-Path $wordContextMutation 'PolicyConcurrencyTests.exe'
    & $wordCompiler /nologo /target:exe /langversion:5 "/out:$wordContextExecutable" /r:System.Web.Extensions.dll @wordContextSources $wordTestSource *> (Join-Path $wordContextMutation 'compile.log')
    if ($LASTEXITCODE -ne 0) { throw 'Request-context replay mutation did not compile.' }
    & $wordContextExecutable $wordContract *> (Join-Path $wordContextMutation 'policy.log')
    if ($LASTEXITCODE -eq 0 -or -not ([IO.File]::ReadAllText((Join-Path $wordContextMutation 'policy.log')).Contains('A consumed request cannot replay with a changed context identity.'))) {
        throw 'The policy test did not reject the intended changed-context replay bypass.'
    }
    Write-Output 'PASS: scoped-policy test rejects changed-context replay in an isolated copied build.'
}
$wordChanged = @($wordInputs | Where-Object { (Get-FileHash -LiteralPath (Join-Path $wordRoot $_.path)).Hash.ToLowerInvariant() -ne $_.sha256 })
[ordered]@{status=$(if($wordChanged.Count -eq 0){'PASS'}else{'FAIL'});inputs=$wordInputs;changedInputs=$wordChanged;assemblySha256=(Get-FileHash -LiteralPath $wordExecutable).Hash.ToLowerInvariant();nativeWord='NOT RUN';mutationCheck=[bool]$MutationCheck} |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $wordOut 'manifest.json') -Encoding utf8
if ($wordChanged.Count -ne 0) { throw 'Policy inputs changed during concurrency qualification.' }
Write-Output "Scoped policy evidence: $wordOut"
