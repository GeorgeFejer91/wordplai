param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$wordRepository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $wordRepository ('artifacts/word/precommit-conflict-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$wordOutput = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $wordOutput) { throw 'Choose fresh conflict evidence.' }
New-Item -ItemType Directory -Path $wordOutput | Out-Null
$wordInputDirectory = Join-Path $wordOutput 'inputs'
$wordOriginalSources = @(Get-ChildItem -LiteralPath (Join-Path $wordRepository 'modules/word/src') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$wordOriginalTests = @('NativeHighlightingSmoke.cs', 'WordAutomation.cs', 'NativePrecommitConflict.cs', 'NativeMcpClient.cs') |
    ForEach-Object { Join-Path $wordRepository ('modules/word/tests/' + $_) }
$wordOriginalContract = Join-Path $wordRepository 'modules/review/selective-editor-contract.json'
$wordInputs = @(@($wordOriginalSources) + @($wordOriginalTests) + @($wordOriginalContract, $PSCommandPath) | ForEach-Object {
    $wordRelative = [IO.Path]::GetRelativePath($wordRepository, $_).Replace('\', '/')
    $wordSnapshot = Join-Path $wordInputDirectory $wordRelative
    New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($wordSnapshot)) -Force | Out-Null
    [IO.File]::WriteAllBytes($wordSnapshot, [IO.File]::ReadAllBytes($_))
    [ordered]@{path=$wordRelative;sha256=(Get-FileHash -LiteralPath $wordSnapshot).Hash.ToLowerInvariant()}
})
$wordInputs | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $wordOutput 'inputs.json') -Encoding utf8
$wordHighlightPath = 'modules/word/src/WordHighlighting.cs'
$wordSource = [IO.File]::ReadAllText((Join-Path $wordInputDirectory $wordHighlightPath))
$wordProductHash = ($wordInputs | Where-Object { $_.path -eq $wordHighlightPath }).sha256
$wordNeedle = 'WordSnapshot beforeCommit = WordSnapshot.Capture(state.Document);'
if ($wordSource.Split(@($wordNeedle), [StringSplitOptions]::None).Length -ne 2) { throw 'The one named precommit fault seam changed.' }
$wordInjection = 'state.Document.Range(0, 0).InsertBefore("User change. "); ' + $wordNeedle
$wordFaultSource = Join-Path $wordOutput 'WordHighlighting.cs'
[IO.File]::WriteAllText($wordFaultSource, $wordSource.Replace($wordNeedle, $wordInjection))
$wordFaultHash = (Get-FileHash -LiteralPath $wordFaultSource).Hash.ToLowerInvariant()
$wordSources = @($wordOriginalSources | Where-Object { [IO.Path]::GetFileName($_) -ne 'WordHighlighting.cs' } |
    ForEach-Object { Join-Path $wordInputDirectory ([IO.Path]::GetRelativePath($wordRepository, $_)) })
$wordSources += $wordFaultSource
$wordTests = @($wordOriginalTests | ForEach-Object { Join-Path $wordInputDirectory ([IO.Path]::GetRelativePath($wordRepository, $_)) })
$wordContract = Join-Path $wordInputDirectory ([IO.Path]::GetRelativePath($wordRepository, $wordOriginalContract))
$wordCompiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$wordExecutable = Join-Path $env:ProgramFiles 'Microsoft Office/root/Office16/WINWORD.EXE'
$wordTest = Join-Path $wordOutput 'NativePrecommitConflict.exe'
$wordCompileCode = $null
$wordNativeCode = $null
$wordChanged = @()
try {
    & $wordCompiler /nologo /target:exe /langversion:5 "/out:$wordTest" /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.Security.dll /r:Microsoft.CSharp.dll "/resource:$wordContract,SelectiveEditorContract" @wordSources @wordTests *> (Join-Path $wordOutput 'compile.log')
    $wordCompileCode = $LASTEXITCODE
    if ($wordCompileCode -ne 0) { throw 'Conflict probe compilation failed; see compile.log.' }
    & $wordTest $wordExecutable (Join-Path $wordOutput 'native') 'colors-only' 2>&1 | Tee-Object -FilePath (Join-Path $wordOutput 'native.log')
    $wordNativeCode = $LASTEXITCODE
}
finally {
    foreach ($wordInput in $wordInputs) {
        foreach ($wordCheck in @(@{kind='original';root=$wordRepository}, @{kind='snapshot';root=$wordInputDirectory})) {
            $wordPath = Join-Path $wordCheck.root $wordInput.path
            $wordActual = if (Test-Path -LiteralPath $wordPath) { (Get-FileHash -LiteralPath $wordPath).Hash.ToLowerInvariant() } else { 'missing' }
            if ($wordActual -ne $wordInput.sha256) { $wordChanged += [ordered]@{kind=$wordCheck.kind;path=$wordInput.path;expected=$wordInput.sha256;actual=$wordActual} }
        }
    }
    $wordActualFaultHash = if (Test-Path -LiteralPath $wordFaultSource) { (Get-FileHash -LiteralPath $wordFaultSource).Hash.ToLowerInvariant() } else { 'missing' }
    if ($wordActualFaultHash -ne $wordFaultHash) { $wordChanged += [ordered]@{kind='fault';path='WordHighlighting.cs';expected=$wordFaultHash;actual=$wordActualFaultHash} }
    [ordered]@{
        status=$(if ($wordCompileCode -eq 0 -and $wordNativeCode -eq 0 -and $wordChanged.Count -eq 0) { 'PASS' } else { 'FAIL' })
        compileExitCode=$wordCompileCode;exitCode=$wordNativeCode
        nativeWord=$(if ($null -eq $wordNativeCode) { 'NOT RUN' } elseif ($wordNativeCode -eq 0) { 'PASS' } else { 'FAIL' })
        fault='One named copied-source seam inserts real native user text immediately before commit validation. No product hook is added.'
        faultSeam=$wordNeedle;faultReplacement=$wordInjection;faultOccurrences=1
        faultStatus=$(if ($wordActualFaultHash -eq $wordFaultHash) { 'APPLIED_ONCE_UNCHANGED' } else { 'CHANGED_DURING_RUN' })
        productSourceSha256=$wordProductHash;injectedSourceSha256=$wordFaultHash;actualInjectedSourceSha256=$wordActualFaultHash
        inputs=$wordInputs;changedInputs=$wordChanged
        harnessSha256=$(if (Test-Path -LiteralPath $wordTest) { (Get-FileHash -LiteralPath $wordTest).Hash.ToLowerInvariant() } else { $null })
        normalWord='NOT TOUCHED';installedPackage='NOT RUN';publicTunnel='NOT RUN'
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $wordOutput 'result.json') -Encoding utf8
}
if ($wordChanged.Count -ne 0) { throw 'Inputs changed during native precommit-conflict qualification; see result.json.' }
if ($wordNativeCode -ne 0) { throw 'Native precommit-conflict qualification failed; see native.log and result.json.' }
Write-Output "Precommit conflict evidence: $wordOutput"
