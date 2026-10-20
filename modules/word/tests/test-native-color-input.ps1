param([string]$OutputDirectory, [switch]$RunNative)
$ErrorActionPreference = 'Stop'
$wordRepository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $wordRepository ('artifacts/word/color-input-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$wordOutput = [IO.Path]::GetFullPath($OutputDirectory)
if (-not $wordOutput.StartsWith([IO.Path]::GetFullPath((Join-Path $wordRepository 'artifacts/word')) + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Use fresh ignored Word artifact evidence.' }
if (Test-Path -LiteralPath $wordOutput) { throw 'Preserve existing evidence; choose a fresh output directory.' }
[IO.Directory]::CreateDirectory($wordOutput) | Out-Null
$wordInputDirectory = Join-Path $wordOutput 'inputs'
$wordOriginalSources = @(Get-ChildItem -LiteralPath (Join-Path $wordRepository 'modules/word/src') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$wordOriginalTests = @('NativeHighlightingSmoke.cs','WordAutomation.cs','NativeMcpClient.cs') | ForEach-Object { Join-Path $wordRepository ('modules/word/tests/' + $_) }
$wordOriginalTests += Join-Path $PSScriptRoot 'NativeColorInputSmoke.cs'
$wordOriginalContract = Join-Path $wordRepository 'modules/review/selective-editor-contract.json'
$wordInputs = @()
$wordResults = @()
$wordChanged = @()
$wordSnapshotComplete = $false
$wordCompiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$wordExecutable = Join-Path $env:ProgramFiles 'Microsoft Office/root/Office16/WINWORD.EXE'
function Replace-One([string]$Source, [string]$Needle, [string]$Replacement) {
    if ($Source.Split(@($Needle), [StringSplitOptions]::None).Length -ne 2) { throw "Named copied-source seam changed: $Needle" }
    return $Source.Replace($Needle, $Replacement)
}
try {
    foreach ($wordOriginal in @(@($wordOriginalSources) + @($wordOriginalTests) + @($wordOriginalContract, $PSCommandPath))) {
        $wordRelative = [IO.Path]::GetRelativePath($wordRepository, $wordOriginal).Replace('\','/')
        $wordSnapshot = Join-Path $wordInputDirectory $wordRelative
        [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($wordSnapshot)) | Out-Null
        [IO.File]::WriteAllBytes($wordSnapshot,[IO.File]::ReadAllBytes($wordOriginal))
        $wordDigest = (Get-FileHash -LiteralPath $wordSnapshot).Hash.ToLowerInvariant()
        $wordInputs += [ordered]@{path=$wordRelative;sha256=$wordDigest}
        if ((Get-FileHash -LiteralPath $wordOriginal).Hash.ToLowerInvariant() -ne $wordDigest) { throw 'An original input changed while its snapshot was being captured.' }
    }
    $wordInputs | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $wordOutput 'inputs.json') -Encoding utf8
    $wordSnapshotComplete = $true
    $wordHighlightRelative = 'modules/word/src/WordHighlighting.cs'
    $wordBase = [IO.File]::ReadAllText((Join-Path $wordInputDirectory $wordHighlightRelative))
    $wordBase = Replace-One $wordBase 'state.NativeEdit.AddGuard(state.Document);' 'if (global::NativeHighlightingSmoke.ColorInputForceRollback) throw new InvalidOperationException("Synthetic copied post-paint failure."); state.NativeEdit.AddGuard(state.Document);'
    $wordBase = Replace-One $wordBase 'WordSnapshot beforeUndo = WordSnapshot.Capture(state.Document);' 'global::NativeHighlightingSmoke.ColorInputAtRollbackBoundary((object)application, (object)state.Document); WordSnapshot beforeUndo = WordSnapshot.Capture(state.Document);'
    $wordMixed = Replace-One $wordBase 'if (!mixedRecord)' 'if (true)'
    $wordMixed = Replace-One $wordMixed 'if (coloringInput.CanRollback(colorFrame, admission, snapshot.FullText, afterFailure.FullText))' 'if (true)'
    $wordMixed = Replace-One $wordMixed 'if (coloringInput.CanRollback(colorFrame, admission, snapshot.FullText, beforeUndo.FullText))' 'if (true)'
    $wordBoundary = Replace-One $wordBase 'if (coloringInput.CanRollback(colorFrame, admission, snapshot.FullText, beforeUndo.FullText))' 'if (true)'
    $wordVariants = @(
        [pscustomobject]@{name='good';source=$wordBase;scenario='all';predicate=$null},
        [pscustomobject]@{name='unsafe-mixed-undo';source=$wordMixed;scenario='after-start-input';predicate='PRESERVE_USER_TEXT:'},
        [pscustomobject]@{name='unsafe-boundary-undo';source=$wordBoundary;scenario='rollback-boundary';predicate='PRESERVE_BOUNDARY_USER_TEXT:'}
    )
    $wordRuntime = @($wordOriginalSources | Where-Object { [IO.Path]::GetFileName($_) -ne 'WordHighlighting.cs' } | ForEach-Object { Join-Path $wordInputDirectory ([IO.Path]::GetRelativePath($wordRepository,$_)) })
    $wordTests = @($wordOriginalTests | ForEach-Object { Join-Path $wordInputDirectory ([IO.Path]::GetRelativePath($wordRepository,$_)) })
    $wordContract = Join-Path $wordInputDirectory ([IO.Path]::GetRelativePath($wordRepository,$wordOriginalContract))
    foreach ($wordVariant in $wordVariants) {
        $wordDirectory = Join-Path $wordOutput $wordVariant.name
        [IO.Directory]::CreateDirectory($wordDirectory) | Out-Null
        $wordHighlight = Join-Path $wordDirectory 'WordHighlighting.cs'
        [IO.File]::WriteAllText($wordHighlight,$wordVariant.source,[Text.UTF8Encoding]::new($false))
        $wordTest = Join-Path $wordDirectory 'NativeColorInputSmoke.exe'
        $wordCompileLog = Join-Path $wordDirectory 'compile.log'
        $wordResult = [ordered]@{variant=$wordVariant.name;compileExit=$null;compileLog=$wordCompileLog;native='NOT RUN';exitCode=$null;negativeControl='NOT RUN';sourcePath=$wordHighlight;sourceSha256=(Get-FileHash -LiteralPath $wordHighlight).Hash.ToLowerInvariant();harnessPath=$wordTest;harnessSha256=$null}
        $wordResults += $wordResult
        & $wordCompiler /nologo /target:exe /langversion:5 /main:NativeHighlightingSmoke "/out:$wordTest" /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.Security.dll /r:Microsoft.CSharp.dll "/resource:$wordContract,SelectiveEditorContract" @wordRuntime $wordHighlight @wordTests *> $wordCompileLog
        $wordResult.compileExit = $LASTEXITCODE
        if ($wordResult.compileExit -ne 0) { throw 'Native color-input compilation failed; see compile.log.' }
        $wordResult.harnessSha256 = (Get-FileHash -LiteralPath $wordTest).Hash.ToLowerInvariant()
        if ($RunNative) {
            $wordNative = Join-Path $wordDirectory 'native'
            & $wordTest $wordExecutable $wordNative 'colors-only' $wordVariant.scenario *> (Join-Path $wordDirectory 'native.log')
            $wordResult.exitCode = $LASTEXITCODE
            $wordText = [IO.File]::ReadAllText((Join-Path $wordDirectory 'native.log'))
            if ($wordVariant.name -eq 'good') {
                $wordResult.native = $(if ($wordResult.exitCode -eq 0) { 'PASS' } else { 'FAIL' })
                if ($wordResult.exitCode -ne 0) { throw 'Native color-input outcome gate failed; preserve native.log.' }
            } else {
                $wordEffective = $wordResult.exitCode -eq 1 -and $wordText.Contains($wordVariant.predicate) -and -not $wordText.Contains('FIXTURE_SEAM:')
                $wordResult.negativeControl = $(if ($wordEffective) { 'EFFECTIVE ASSERTION FAILURE' } else { 'FAIL' })
                if (-not $wordEffective) { throw 'Unsafe-Undo control did not fail the exact user-text preservation predicate.' }
            }
        }
    }
}
finally {
    foreach ($wordInput in $wordInputs) {
        foreach ($wordCheck in @(@{kind='original';root=$wordRepository},@{kind='snapshot';root=$wordInputDirectory})) {
            $wordPath = Join-Path $wordCheck.root $wordInput.path
            $wordActual = if (Test-Path -LiteralPath $wordPath) { (Get-FileHash -LiteralPath $wordPath).Hash.ToLowerInvariant() } else { 'missing' }
            if ($wordActual -ne $wordInput.sha256) { $wordChanged += [ordered]@{kind=$wordCheck.kind;path=$wordInput.path;expected=$wordInput.sha256;actual=$wordActual} }
        }
    }
    foreach ($wordResult in $wordResults) {
        foreach ($wordCheck in @(@{kind='injected source';path=$wordResult.sourcePath;sha256=$wordResult.sourceSha256},@{kind='compiled harness';path=$wordResult.harnessPath;sha256=$wordResult.harnessSha256})) {
            if ($null -eq $wordCheck.sha256) { continue }
            $wordActual = if (Test-Path -LiteralPath $wordCheck.path) { (Get-FileHash -LiteralPath $wordCheck.path).Hash.ToLowerInvariant() } else { 'missing' }
            if ($wordActual -ne $wordCheck.sha256) { $wordChanged += [ordered]@{kind=$wordCheck.kind;path=$wordCheck.path;expected=$wordCheck.sha256;actual=$wordActual} }
        }
    }
    $wordCompileGood = $wordSnapshotComplete -and $wordResults.Count -eq 3 -and @($wordResults | Where-Object { $null -eq $_.compileExit -or $_.compileExit -ne 0 }).Count -eq 0 -and $wordChanged.Count -eq 0
    $wordNativeGood = $wordCompileGood -and $wordResults[0].native -eq 'PASS' -and @($wordResults | Where-Object { $_.variant -ne 'good' -and $_.negativeControl -ne 'EFFECTIVE ASSERTION FAILURE' }).Count -eq 0
    [ordered]@{status=$(if (-not $wordCompileGood -or $RunNative -and -not $wordNativeGood) { 'FAIL' } elseif ($RunNative) { 'PASS' } else { 'COMPILE PASS' });nativeRequested=[bool]$RunNative;results=$wordResults;inputs=$wordInputs;changedInputs=$wordChanged;snapshotComplete=$wordSnapshotComplete;copiedFaultSeams=@('Failure after real paint before AddGuard when selected by this fixture','Separate ended user record after NativeIdentity before final rollback capture');normalWord='NOT TOUCHED';installation='NOT RUN';publicTunnel='NOT RUN';sampling='Injected samples do not qualify the real Win32/foreground/GUI-thread probe.'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $wordOutput 'result.json') -Encoding utf8
    if ($wordChanged.Count -ne 0) { throw 'Original, snapshot, injected source or compiled harness changed during qualification; see result.json.' }
}
Write-Output "Native color-input evidence: $wordOutput (native requested: $RunNative)"
exit 0
