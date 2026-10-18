param([string]$OutputDirectory, [switch]$MutationCheck)
$ErrorActionPreference = 'Stop'
$wordRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $wordRoot ('artifacts/word/uncertain-inspection-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$wordOutput = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $wordOutput) { throw 'Choose fresh uncertain-inspection evidence.' }
New-Item -ItemType Directory -Path $wordOutput | Out-Null
$wordSources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../src') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$wordFixture = Join-Path $PSScriptRoot 'UncertainInspectionTests.cs'
$wordContract = Join-Path $wordRoot 'modules/review/selective-editor-contract.json'
$wordInputs = @(@($wordSources) + @($wordFixture, $wordContract, $PSCommandPath) | ForEach-Object {
    [ordered]@{path=[IO.Path]::GetRelativePath($wordRoot,$_);sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()}
})
$wordCompiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$wordOriginal = [IO.File]::ReadAllText((Join-Path $wordRoot 'modules/word/src/WordRemoteReview.cs')).Replace("`r`n", "`n")
$wordBranch = @'
                // Retained uncertainty is receipt-only; both strict lookups must throw before snapshot access.
                if (state.Policy != null && HasUncertainReceipt(tool, state, args, session))
                {
                    if (tool == "suggest_text") return Suggest(state, null, args, session);
                    if (tool == "highlight_text") return HighlightText(state, null, args, session);
                }
'@
if (-not $wordOriginal.Contains($wordBranch)) { throw 'The retained-uncertainty route no longer matches; review the copied regression.' }
$wordSeams = @{
    'if (!connections.Status.LocalRunning)' = 'if (!global::UncertainInspectionTests.AccessRunning)'
    'WordDocumentReview state = ReviewDocument(tool, args, session);' = 'WordDocumentReview state = global::UncertainInspectionTests.Resolve(args, session);'
    'state = Current(DocumentReview(state.Document));' = 'state = global::UncertainInspectionTests.Inspect(state);'
}
foreach ($wordNeedle in $wordSeams.Keys) {
    if (-not $wordOriginal.Contains($wordNeedle)) { throw 'A copied native boundary no longer matches the uncertainty regression.' }
    $wordOriginal = $wordOriginal.Replace($wordNeedle, $wordSeams[$wordNeedle])
}
$wordResults = [Collections.Generic.List[object]]::new()
$wordFailure = $null
try {
foreach ($wordMode in $(if($MutationCheck){@('production','branch-removal-control')}else{@('production')})) {
    $wordCase = Join-Path $wordOutput $wordMode
    New-Item -ItemType Directory -Path $wordCase | Out-Null
    $wordSource = Join-Path $wordCase 'WordRemoteReview.cs'
    [IO.File]::WriteAllText($wordSource, $(if($wordMode -eq 'production'){$wordOriginal}else{$wordOriginal.Replace($wordBranch,'')}))
    $wordExecutable = Join-Path $wordCase 'UncertainInspectionTests.exe'
    $wordCaseSources = @($wordSources | Where-Object { [IO.Path]::GetFileName($_) -ne 'WordRemoteReview.cs' }) + @($wordSource, $wordFixture)
    & $wordCompiler /nologo /target:exe /langversion:5 "/out:$wordExecutable" /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.Security.dll /r:Microsoft.CSharp.dll "/resource:$wordContract,SelectiveEditorContract" @wordCaseSources *> (Join-Path $wordCase 'compile.log')
    if ($LASTEXITCODE -ne 0) { throw 'Uncertain-inspection tests did not compile.' }
    & $wordExecutable $wordContract *> (Join-Path $wordCase 'probe.log')
    $wordCode = $LASTEXITCODE
    $wordLog = [IO.File]::ReadAllText((Join-Path $wordCase 'probe.log'))
    if ($wordMode -eq 'production') {
        if ($wordCode -ne 0 -or -not $wordLog.Contains('CHECKS 12; FAILURES 0;')) { throw 'An owned exact uncertain receipt was blocked by native inspection or an envelope/scope guard failed.' }
    }
    elseif ($wordCode -ne 1 -or -not $wordLog.Contains('FAIL suggest_text: an exact same-scope uncertain retry remains unknown before blocked native color inspection.') -or
        -not $wordLog.Contains('FAIL highlight_text: an exact same-scope uncertain retry remains unknown before blocked native color inspection.')) {
        throw 'The copied branch-removal control did not expose both intended uncertainty downgrades.'
    }
    $wordResults.Add([ordered]@{mode=$wordMode;exitCode=$wordCode;status=$(if($wordMode -eq 'production'){'PASS'}else{'EFFECTIVE: EXPECTED FAILURE'});assemblySha256=(Get-FileHash -LiteralPath $wordExecutable).Hash.ToLowerInvariant()})
    Write-Output "$wordMode`: $(if($wordMode -eq 'production'){'PASS12'}else{'effective expected failure'})"
}
}
catch { $wordFailure=$_.Exception.Message; throw }
finally {
    $wordChanged = @($wordInputs | Where-Object { (Get-FileHash -LiteralPath (Join-Path $wordRoot $_.path)).Hash.ToLowerInvariant() -ne $_.sha256 })
    [ordered]@{status=$(if($wordFailure -or $wordChanged.Count){'FAIL'}else{'PASS'});error=$wordFailure;inputs=$wordInputs;changedInputs=$wordChanged;results=$wordResults;nativeWord='NOT RUN';nativeEffects='SYNTHETIC FIRST-EFFECT COUNTER; REAL POLICY RECEIPTS/PARSERS/SCOPE/REPLAY FLOW';mutationCheck=[bool]$MutationCheck} |
        ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $wordOutput 'manifest.json') -Encoding utf8
    if ($wordChanged.Count) { throw 'Uncertain-inspection production inputs changed during qualification.' }
}
