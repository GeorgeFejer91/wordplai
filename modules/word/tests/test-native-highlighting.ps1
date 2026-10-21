param([string]$OutputDirectory, [string]$WordExecutable, [switch]$DirectMode, [switch]$ColorStress, [switch]$RecolorStress, [switch]$MutationCheck)
$ErrorActionPreference = 'Stop'
if ($MutationCheck -and -not $DirectMode) { throw 'The comment-author mutation check requires -DirectMode.' }
$repo = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repo ('artifacts/word/native-highlighting-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Choose a new evidence directory.' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
if (-not $WordExecutable) { $WordExecutable = Join-Path $env:ProgramFiles 'Microsoft Office/root/Office16/WINWORD.EXE' }
if (-not (Test-Path -LiteralPath $WordExecutable)) { throw 'Specify the installed Word executable.' }
$sources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../src') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$wordHarnessSource = if ($DirectMode) { 'NativeDirectHighlightingSmoke.cs' } else { 'NativeHighlightingSmoke.cs' }
$testSources = @((Join-Path $PSScriptRoot $wordHarnessSource), (Join-Path $PSScriptRoot 'WordAutomation.cs'), (Join-Path $PSScriptRoot 'NativeMcpClient.cs'))
if (-not $DirectMode) { $testSources += Join-Path $PSScriptRoot 'NativeRequestedColors.cs' }
if ($RecolorStress) { $testSources += Join-Path $PSScriptRoot 'NativeRecoloringSmoke.cs' }
if (@($DirectMode, $ColorStress, $RecolorStress | Where-Object { $_ }).Count -gt 1) { throw 'Choose one native highlighting test mode.' }
$defines = @()
if ($RecolorStress) { $defines += '/define:WORD_RECOLOR_TEST' }
$contract = Join-Path $repo 'modules/review/selective-editor-contract.json'
$snapshot = @(@($sources) + $testSources + @($contract, $PSCommandPath) | ForEach-Object { [ordered]@{path=[IO.Path]::GetRelativePath($repo,$_);sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()} })
$compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$test = Join-Path $OutputDirectory 'NativeHighlightingSmoke.exe'
& $compiler /nologo /target:exe /langversion:5 "/out:$test" /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.Security.dll /r:Microsoft.CSharp.dll "/resource:$contract,SelectiveEditorContract" @defines @sources @testSources
if ($LASTEXITCODE -ne 0) { throw 'Native highlighting harness compilation failed.' }
$nativeArguments = @($WordExecutable, (Join-Path $OutputDirectory 'native'))
if ($DirectMode) { $nativeArguments += 'direct' }
if ($ColorStress) { $nativeArguments += 'colors-only' }
if ($RecolorStress) { $nativeArguments += 'recolor-only' }
& $test @nativeArguments 2>&1 | Tee-Object -FilePath (Join-Path $OutputDirectory 'native.log')
$code = $LASTEXITCODE
$changed = @()
foreach ($entry in $snapshot) {
    $currentPath = Join-Path $repo $entry.path
    $actual = if (Test-Path -LiteralPath $currentPath) { (Get-FileHash -LiteralPath $currentPath).Hash.ToLowerInvariant() } else { 'missing' }
    if ($actual -ne $entry.sha256) { $changed += [ordered]@{path=$entry.path;expected=$entry.sha256;actual=$actual} }
}
[ordered]@{status=$(if($code -eq 0 -and $changed.Count -eq 0){'PASS'}else{'FAIL'});exitCode=$code;mode=$(if($DirectMode){'direct'}elseif($RecolorStress){'recolor-only'}elseif($ColorStress){'colors-only'}else{'native-track-changes'});defines=$defines;inputs=$snapshot;changedInputs=$changed;harnessSha256=(Get-FileHash -LiteralPath $test).Hash.ToLowerInvariant();wordSha256=(Get-FileHash -LiteralPath $WordExecutable).Hash.ToLowerInvariant();installedPackage='NOT RUN';publicTunnel='NOT RUN'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'manifest.json') -Encoding utf8
if ($changed.Count -ne 0) { throw 'Source changed during native qualification.' }
if ($code -ne 0) { throw 'Isolated native highlighting qualification failed.' }
if ($MutationCheck) {
    $copy = Join-Path $OutputDirectory 'comment-author-control'
    New-Item -ItemType Directory -Path $copy | Out-Null
    $original = Join-Path $PSScriptRoot '../src/WordRemoteReview.cs'
    $changedSource = [IO.File]::ReadAllText($original)
    $needle = 'comment.Author = connections.CommentAuthorName;'
    if (-not $changedSource.Contains($needle)) { throw 'Comment-author fault control seam is missing.' }
    $replacement = Join-Path $copy 'WordRemoteReview.cs'
    [IO.File]::WriteAllText($replacement, $changedSource.Replace($needle, 'comment.Author = "wordplai";'))
    $controlSources = @($sources | ForEach-Object { if ([IO.Path]::GetFileName($_) -eq 'WordRemoteReview.cs') { $replacement } else { $_ } })
    $control = Join-Path $copy 'NativeHighlightingSmoke.exe'
    & $compiler /nologo /target:exe /langversion:5 "/out:$control" /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.Security.dll /r:Microsoft.CSharp.dll "/resource:$contract,SelectiveEditorContract" @controlSources @testSources
    if ($LASTEXITCODE -ne 0) { throw 'Comment-author control compilation failed.' }
    $controlLog = & $control $WordExecutable (Join-Path $copy 'native') 'direct' 2>&1
    $controlCode = $LASTEXITCODE
    $controlLog | Set-Content -LiteralPath (Join-Path $copy 'native.log') -Encoding utf8
    $effective = $controlCode -ne 0 -and ($controlLog -join "`n").Contains('The direct native edit includes exactly one concise rationale with the saved comment author.')
    [ordered]@{status=$(if($effective){'PASS'}else{'FAIL'});fault='hardcoded historical comment author';exitCode=$controlCode;effective=$effective} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $copy 'result.json') -Encoding utf8
    if (-not $effective) { throw 'The hardcoded comment-author control did not fail the intended predicate.' }
    Write-Output 'PASS: a copied hardcoded-author implementation fails the actual native attribution check.'
}
