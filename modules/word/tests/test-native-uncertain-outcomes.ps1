param([string]$OutputDirectory, [string]$WordExecutable, [switch]$CompileOnly, [switch]$MutationCheck)
$ErrorActionPreference = 'Stop'
$wordRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $wordRoot ('artifacts/word/native-uncertain-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$wordOutput = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $wordOutput) { throw 'Choose fresh native uncertainty evidence.' }
New-Item -ItemType Directory -Path $wordOutput | Out-Null
$wordGateTimer = [Diagnostics.Stopwatch]::StartNew()
if (-not $WordExecutable) { $WordExecutable = Join-Path $env:ProgramFiles 'Microsoft Office/root/Office16/WINWORD.EXE' }
$wordSources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../src') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$wordTests = @('NativeSuggestionWorkflowSmoke.cs','WordAutomation.cs','NativeMcpClient.cs','NativeConcurrentClients.cs','NativeUncertainOutcomes.cs') |
    ForEach-Object { Join-Path $PSScriptRoot $_ }
$wordContract = Join-Path $wordRoot 'modules/review/selective-editor-contract.json'
$wordInputs = @(@($wordSources) + $wordTests + @($wordContract, $PSCommandPath) | ForEach-Object {
    [ordered]@{path=[IO.Path]::GetRelativePath($wordRoot,$_);sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()}
})
$wordCompiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$wordCommon = @('/nologo','/target:exe','/langversion:5','/r:System.Windows.Forms.dll','/r:System.Drawing.dll',
    '/r:System.Web.Extensions.dll','/r:System.Security.dll','/r:Microsoft.CSharp.dll',"/resource:$wordContract,SelectiveEditorContract")
$wordMain = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'NativeSuggestionWorkflowSmoke.cs'))
$wordEntry = 'string scenario = args.Length < 3 ? "all" : args[2];'
if (-not $wordMain.Contains($wordEntry)) { throw 'The copied native probe entry no longer matches.' }
$wordMain = $wordMain.Replace($wordEntry, $wordEntry + [Environment]::NewLine +
    'if (scenario.StartsWith("uncertain-probe:", StringComparison.Ordinal)) { NativeUncertainOutcome(scenario.Substring(16)); WriteResult("PASS", null); return 0; }')
$wordResults = [Collections.Generic.List[object]]::new()
$wordFailure = $null
try {
foreach ($wordScenario in @('suggestion-rollback','marking-rollback','marking-persistence','suggestion-continuation')) {
    $wordCase = Join-Path $wordOutput $wordScenario
    New-Item -ItemType Directory -Path $wordCase | Out-Null
    $wordResult = [ordered]@{scenario=$wordScenario;status='STARTED';nativeExitCode=$null;nativeMilliseconds=$null;assemblySha256=$null;negativeControl='NOT RUN'}
    $wordResults.Add($wordResult)
    $wordCopies = @{}
    $wordCopies['NativeSuggestionWorkflowSmoke.cs'] = $wordMain
    if ($wordScenario -eq 'suggestion-rollback') {
        $wordField = 'private static WordEditor editor;'
        if (-not $wordMain.Contains($wordField)) { throw 'The copied retained-uncertainty inspection flag no longer matches.' }
        $wordCopies['NativeSuggestionWorkflowSmoke.cs'] = $wordMain.Replace($wordField,
            'internal static bool BlockNativeInspection;' + [Environment]::NewLine + '    ' + $wordField)
        $wordEditor = [IO.File]::ReadAllText((Join-Path $wordRoot 'modules/word/src/WordEditor.cs'))
        $wordCurrent = '(private WordDocumentReview Current\(WordDocumentReview state\)\s*\{)'
        if ([regex]::Matches($wordEditor, $wordCurrent).Count -ne 1) { throw 'The copied native inspection boundary no longer matches.' }
        $wordCopies['WordEditor.cs'] = [regex]::Replace($wordEditor, $wordCurrent,
            '$1' + [Environment]::NewLine + '            if (global::NativeSuggestionWorkflowSmoke.BlockNativeInspection) throw new McpOperationException("Synthetic color inspection limit after retained native uncertainty.", "native_color_verification_limited");')
        $wordOutcome = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'NativeUncertainOutcomes.cs'))
        $wordRetry = 'var retry = client.Call(tool, request);'
        if (-not $wordOutcome.Contains($wordRetry)) { throw 'The copied native uncertainty retry boundary no longer matches.' }
        $wordBlocked = @'
BlockNativeInspection = true;
        var blocked = client.Call("read_manuscript", Values("manuscript_id", documentId));
        Check((bool)blocked["isError"] && Text(Map(Map(blocked["structuredContent"])["error"]), "code") == "native_color_verification_limited" && !state.Policy.IsShared,
            "After an actual retained native uncertainty, a fresh read encounters the copied native color-verification limit without restoring sharing.");
        var retry = client.Call(tool, request);
'@
        $wordCopies['NativeUncertainOutcomes.cs'] = $wordOutcome.Replace($wordRetry, $wordBlocked)
    }
    if ($wordScenario.StartsWith('suggestion')) {
        $wordFile = 'WordRemoteReview.cs'
        $wordSource = [IO.File]::ReadAllText((Join-Path $wordRoot ('modules/word/src/' + $wordFile)))
        if ($wordScenario -eq 'suggestion-continuation') {
            $wordNeedle = 'return SuggestResult(recorded, ContinuationContext(state, state.Snapshot, session), requestId);'
            $wordReplacement = 'throw new System.IO.IOException("Injected continuation failure after native commit.");'
        }
        else {
            $wordNeedle = 'before.Range(state.Document, proposal.From, proposal.To - retainedSuffix).Text = replacement.Substring(0, replacement.Length - retainedSuffix);'
            $wordReplacement = $wordNeedle + [Environment]::NewLine + 'throw new InvalidOperationException("Injected failure after native wording mutation.");'
            $wordRollback = '// A failed COM call may already have changed bookmarks or text.'
            if (-not $wordSource.Contains($wordRollback)) { throw 'The copied suggestion rollback probe no longer matches.' }
            $wordSource = $wordSource.Replace($wordRollback, 'throw new System.Runtime.InteropServices.COMException("Injected unavailable native rollback.");')
        }
        if (-not $wordSource.Contains($wordNeedle)) { throw 'The copied suggestion effect probe no longer matches.' }
        $wordCopies[$wordFile] = $wordSource.Replace($wordNeedle, $wordReplacement)
    }
    else {
        $wordFile = 'WordHighlighting.cs'
        $wordSource = [IO.File]::ReadAllText((Join-Path $wordRoot ('modules/word/src/' + $wordFile)))
        $wordNeedle = 'state.NativeEdit.AddGuard(state.Document);'
        if (-not $wordSource.Contains($wordNeedle) -or -not $wordSource.Contains('bool restored = !started;') -or -not $wordSource.Contains('if (started) try')) {
            throw 'The copied native marking rollback probes no longer match.'
        }
        $wordCopies[$wordFile] = $wordSource.Replace($wordNeedle, $wordNeedle + [Environment]::NewLine +
            'throw new InvalidOperationException("Injected failure after native marking.");').Replace('bool restored = !started;', 'bool restored = false;').Replace('if (started) try', 'if (false && started) try')
        if ($wordScenario -eq 'marking-persistence') {
            $wordHelper = [IO.File]::ReadAllText((Join-Path $wordRoot 'modules/word/src/WordNativeOutcome.cs'))
            if (-not $wordHelper.Contains('try { persist(); }')) { throw 'The copied uncertain-receipt persistence probe no longer matches.' }
            $wordCopies['WordNativeOutcome.cs'] = $wordHelper.Replace('try { persist(); }', 'try { throw new System.IO.IOException("Injected uncertain receipt write failure."); }')
        }
    }
    foreach ($wordName in $wordCopies.Keys) { [IO.File]::WriteAllText((Join-Path $wordCase $wordName), $wordCopies[$wordName]) }
    $wordCaseSources = @(@($wordSources) + $wordTests | Where-Object { -not $wordCopies.ContainsKey([IO.Path]::GetFileName($_)) }) + @($wordCopies.Keys | ForEach-Object { Join-Path $wordCase $_ })
    $wordProbeExecutable = Join-Path $wordCase 'NativeSuggestionWorkflowSmoke.exe'
    & $wordCompiler @wordCommon "/out:$wordProbeExecutable" @wordCaseSources *> (Join-Path $wordCase 'compile.log')
    if ($LASTEXITCODE -ne 0) { throw "The copied $wordScenario probe did not compile." }
    $wordResult.assemblySha256 = (Get-FileHash -LiteralPath $wordProbeExecutable).Hash.ToLowerInvariant()
    if (-not $CompileOnly) {
        $wordNativeTimer = [Diagnostics.Stopwatch]::StartNew()
        & $wordProbeExecutable $WordExecutable (Join-Path $wordCase 'native') ('uncertain-probe:' + $wordScenario) *> (Join-Path $wordCase 'native.log')
        $wordResult.nativeExitCode = $LASTEXITCODE
        $wordResult.nativeMilliseconds = $wordNativeTimer.ElapsedMilliseconds
        if ($LASTEXITCODE -ne 0) { throw "The copied $wordScenario native gate failed." }
    }
    $wordResult.status = $(if($CompileOnly){'COMPILE ONLY'}else{'PASS'})
    if ($MutationCheck -and $wordScenario -ne 'suggestion-continuation') {
        $wordMutant = Join-Path $wordCase 'negative-control'
        New-Item -ItemType Directory -Path $wordMutant | Out-Null
        $wordFaultSource = $wordCopies[$wordFile]
        if (-not $wordFaultSource.Contains('if (nativeEffects) WordNativeOutcome.Unknown(')) { throw 'The copied unknown-outcome negative control no longer matches.' }
        $wordMutantSource = Join-Path $wordMutant $wordFile
        [IO.File]::WriteAllText($wordMutantSource, $wordFaultSource.Replace('if (nativeEffects) WordNativeOutcome.Unknown(', 'if (false && nativeEffects) WordNativeOutcome.Unknown('))
        $wordMutantSources = @($wordCaseSources | Where-Object { [IO.Path]::GetFileName($_) -ne $wordFile }) + @($wordMutantSource)
        $wordMutantExecutable = Join-Path $wordMutant 'NativeSuggestionWorkflowSmoke.exe'
        & $wordCompiler @wordCommon "/out:$wordMutantExecutable" @wordMutantSources *> (Join-Path $wordMutant 'compile.log')
        if ($LASTEXITCODE -ne 0) { throw 'The copied native outcome negative control did not compile.' }
        if (-not $CompileOnly) {
            & $wordMutantExecutable $WordExecutable (Join-Path $wordMutant 'native') ('uncertain-probe:' + $wordScenario) *> (Join-Path $wordMutant 'native.log')
            if ($LASTEXITCODE -eq 0 -or -not ([IO.File]::ReadAllText((Join-Path $wordMutant 'native.log')).Contains('An unconfirmed or undelivered native effect remains outcome_unknown through rollback, persistence and finalizer failures.'))) {
                throw 'The copied native gate did not reject disabled unknown-outcome classification.'
            }
        }
        $wordResult.negativeControl = $(if($CompileOnly){'COMPILE ONLY'}else{'EFFECTIVE: EXPECTED FAILURE'})
    }
    Write-Output "Native uncertainty $wordScenario`: $(if($CompileOnly){'COMPILED; NOT RUN'}else{'PASS'})"
}
}
catch { $wordFailure = $_.Exception.Message; throw }
finally {
$wordChanged = @($wordInputs | Where-Object { (Get-FileHash -LiteralPath (Join-Path $wordRoot $_.path)).Hash.ToLowerInvariant() -ne $_.sha256 })
[ordered]@{status=$(if($wordFailure -or $wordChanged.Count -ne 0){'FAIL'}elseif($CompileOnly){'COMPILE ONLY'}else{'PASS'});error=$wordFailure;inputs=$wordInputs;changedInputs=$wordChanged;scenarios=$wordResults;elapsedMilliseconds=$wordGateTimer.ElapsedMilliseconds;mutationCheck=[bool]$MutationCheck;nativeWord=$(if($CompileOnly){'NOT RUN'}else{'OWNED DISPOSABLE PROCESS ONLY'});normalWord='NOT TOUCHED'} |
    ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $wordOutput 'manifest.json') -Encoding utf8
if ($wordChanged.Count -ne 0) { throw 'Native uncertainty inputs changed during qualification.' }
}
