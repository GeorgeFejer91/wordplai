param([string]$OutputDirectory, [switch]$NativeWord, [string]$WordExecutable)
$ErrorActionPreference = 'Stop'
$wordRepository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $wordRepository ('artifacts/word/autosave-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$wordOutput = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $wordOutput) { throw 'Choose a new evidence directory.' }
New-Item -ItemType Directory -Path $wordOutput | Out-Null
$wordCompiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$wordSources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../src') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$wordTests = @('WordAutoSaveTests.cs','NativeAutoSaveSmoke.cs','WordAutomation.cs') | ForEach-Object { Join-Path $PSScriptRoot $_ }
$wordContract = Join-Path $wordRepository 'modules/review/selective-editor-contract.json'
$wordInputs = @(@($wordSources) + $wordTests + @($wordContract, $PSCommandPath) | ForEach-Object {
    [ordered]@{path=[IO.Path]::GetRelativePath($wordRepository,$_);sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()}
})
$wordArguments = @('/nologo','/target:exe','/langversion:5','/r:System.Windows.Forms.dll','/r:System.Drawing.dll','/r:System.Web.Extensions.dll',
    '/r:System.Security.dll','/r:Microsoft.CSharp.dll',"/resource:$wordContract,SelectiveEditorContract")
$wordComponent = Join-Path $wordOutput 'WordAutoSaveTests.exe'
& $wordCompiler @wordArguments '/main:WordAutoSaveTests' "/out:$wordComponent" @wordSources $wordTests[0] *> (Join-Path $wordOutput 'component-compile.log')
if ($LASTEXITCODE -ne 0) { throw 'Automatic-save component compilation failed.' }
& $wordComponent 2>&1 | Tee-Object -FilePath (Join-Path $wordOutput 'component.log')
if ($LASTEXITCODE -ne 0) { throw 'Automatic-save component checks failed.' }
$wordMutationDirectory = Join-Path $wordOutput 'mutation'
New-Item -ItemType Directory -Path $wordMutationDirectory | Out-Null
$wordOriginal = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '../src/WordAutoSave.cs'))
if (-not $wordOriginal.Contains('document.Save();')) { throw 'The save mutation must be updated for the current implementation.' }
$wordMutationSource = Join-Path $wordMutationDirectory 'WordAutoSave.cs'
[IO.File]::WriteAllText($wordMutationSource, $wordOriginal.Replace('document.Save();', 'document.Saved = true;'))
$wordMutationSources = @($wordSources | Where-Object { [IO.Path]::GetFileName($_) -ne 'WordAutoSave.cs' }) + @($wordMutationSource)
$wordMutationTest = Join-Path $wordMutationDirectory 'WordAutoSaveTests.exe'
& $wordCompiler @wordArguments '/main:WordAutoSaveTests' "/out:$wordMutationTest" @wordMutationSources $wordTests[0] *> (Join-Path $wordMutationDirectory 'compile.log')
if ($LASTEXITCODE -ne 0) { throw 'Automatic-save mutation compilation failed.' }
& $wordMutationTest *> (Join-Path $wordMutationDirectory 'component.log')
if ($LASTEXITCODE -eq 0 -or -not ([IO.File]::ReadAllText((Join-Path $wordMutationDirectory 'component.log')).Contains('An idle existing named document refreshes and saves once.'))) {
    throw 'The automatic-save tests failed to detect a false saved flag without a native save.'
}
Write-Output 'PASS: automatic-save tests reject a false saved flag without persisting the document.'
$wordTargetMutations = @(
    @{name='target';pattern='return targetDocument == null ? (object)application.ActiveDocument : targetDocument();';replacement='return (object)application.ActiveDocument;';failure='Active and selected background documents each refresh and save without changing the active window.'},
    @{name='target-recheck';pattern='if (refreshedTarget == null || WordSnapshot.Identity(refreshedTarget) != identity) return;';replacement='if (targetDocument == null && (refreshedTarget == null || WordSnapshot.Identity(refreshedTarget) != identity)) return;';failure='Closing the selected background document during refresh saves no stale or unrelated document.'},
    @{name='fairness';pattern='worker.Pulse(now);';replacement='worker.Pulse(now); if (worker.LastError != null) break;';failure='All eligible background documents are processed before a save-error report can block the UI.'}
)
foreach ($wordTargetMutation in $wordTargetMutations) {
    if (-not $wordOriginal.Contains($wordTargetMutation.pattern)) { throw ('Update the automatic-save ' + $wordTargetMutation.name + ' mutation for the current implementation.') }
    $wordTargetDirectory = Join-Path $wordOutput ($wordTargetMutation.name + '-mutation')
    New-Item -ItemType Directory -Path $wordTargetDirectory | Out-Null
    $wordTargetSource = Join-Path $wordTargetDirectory 'WordAutoSave.cs'
    [IO.File]::WriteAllText($wordTargetSource, $wordOriginal.Replace($wordTargetMutation.pattern, $wordTargetMutation.replacement))
    $wordTargetSources = @($wordSources | Where-Object { [IO.Path]::GetFileName($_) -ne 'WordAutoSave.cs' }) + @($wordTargetSource)
    $wordTargetTest = Join-Path $wordTargetDirectory 'WordAutoSaveTests.exe'
    & $wordCompiler @wordArguments '/main:WordAutoSaveTests' "/out:$wordTargetTest" @wordTargetSources $wordTests[0] *> (Join-Path $wordTargetDirectory 'compile.log')
    if ($LASTEXITCODE -ne 0) { throw ('Automatic-save ' + $wordTargetMutation.name + ' mutation compilation failed.') }
    & $wordTargetTest *> (Join-Path $wordTargetDirectory 'component.log')
    if ($LASTEXITCODE -eq 0 -or -not ([IO.File]::ReadAllText((Join-Path $wordTargetDirectory 'component.log')).Contains($wordTargetMutation.failure))) {
        throw ('The automatic-save tests did not detect the ' + $wordTargetMutation.name + ' mutation.')
    }
    Write-Output ('PASS: automatic-save tests detect the ' + $wordTargetMutation.name + ' mutation.')
}
$wordNativeStatus = 'NOT RUN'
if ($NativeWord) {
    if (-not $WordExecutable) { $WordExecutable = Join-Path $env:ProgramFiles 'Microsoft Office/root/Office16/WINWORD.EXE' }
    if (-not (Test-Path -LiteralPath $WordExecutable)) { throw 'Specify the installed Word executable.' }
    $wordNative = Join-Path $wordOutput 'NativeAutoSaveSmoke.exe'
    & $wordCompiler @wordArguments '/main:NativeAutoSaveSmoke' "/out:$wordNative" @wordSources $wordTests[1] $wordTests[2] *> (Join-Path $wordOutput 'native-compile.log')
    if ($LASTEXITCODE -ne 0) { throw 'Native automatic-save compilation failed.' }
    & $wordNative $WordExecutable (Join-Path $wordOutput 'native') 2>&1 | Tee-Object -FilePath (Join-Path $wordOutput 'native.log')
    if ($LASTEXITCODE -ne 0) { throw 'Native automatic-save checks failed.' }
    $wordNativeStatus = 'PASS'
}
foreach ($wordInput in $wordInputs) {
    if ((Get-FileHash -LiteralPath (Join-Path $wordRepository $wordInput.path)).Hash.ToLowerInvariant() -ne $wordInput.sha256) {
        throw "Source changed during automatic-save verification: $($wordInput.path)"
    }
}
[ordered]@{status='PASS';inputs=$wordInputs;component='PASS';saveMutation='PASS';targetMutation='PASS';targetRecheckMutation='PASS';fairnessMutation='PASS';nativeWord=$wordNativeStatus;installedPackage='NOT RUN';publicClient='NOT RUN'} |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $wordOutput 'manifest.json') -Encoding utf8
Write-Output "Automatic-save evidence: $wordOutput"
