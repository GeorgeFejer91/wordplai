param([string]$OutputDirectory,[string]$WordExecutable,[switch]$NavigationOnly,[switch]$RealInput)
$ErrorActionPreference='Stop'
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
if(-not $OutputDirectory){$OutputDirectory=Join-Path $repo ('artifacts/word/native-double-click-'+(Get-Date -Format 'yyyyMMdd-HHmmss'))}
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Choose a new native event evidence directory.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
if(-not $WordExecutable){$WordExecutable=Join-Path $env:ProgramFiles 'Microsoft Office/root/Office16/WINWORD.EXE'}
$sources=@(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '../src') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$testSources=@((Join-Path $PSScriptRoot 'NativeDoubleClickSmoke.cs'),(Join-Path $PSScriptRoot 'WordAutomation.cs'))
$contract=Join-Path $repo 'modules/review/selective-editor-contract.json'
$icon=Join-Path $repo 'assets/wraiter/icon.ico'
$ribbon=Join-Path $repo 'assets/wraiter/icon16.png'
$inputs=@(@($sources)+$testSources+@($contract,$icon,$ribbon,$PSCommandPath)|ForEach-Object{[ordered]@{path=[IO.Path]::GetRelativePath($repo,$_);sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()}})
$test=Join-Path $OutputDirectory 'NativeDoubleClickSmoke.exe'
& "$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe" /nologo /target:exe /langversion:5 "/out:$test" /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.Security.dll /r:Microsoft.CSharp.dll "/resource:$contract,SelectiveEditorContract" "/resource:$icon,wrAIter.Icon.ico" "/resource:$ribbon,wrAIter.Ribbon.png" @sources @testSources
if($LASTEXITCODE-ne0){throw 'Native event harness compilation failed.'}
if($RealInput -and $NavigationOnly){throw 'Choose NavigationOnly or RealInput.'}
$mode=if($RealInput){'real-input'}else{'navigation-only'}
& $test $WordExecutable (Join-Path $OutputDirectory 'native') $mode 2>&1 | Tee-Object -FilePath (Join-Path $OutputDirectory 'native.log')
$code=$LASTEXITCODE
$changed=@($inputs|ForEach-Object{$actual=(Get-FileHash -LiteralPath (Join-Path $repo $_.path)).Hash.ToLowerInvariant();if($actual-ne$_.sha256){[ordered]@{path=$_.path;expected=$_.sha256;actual=$actual}}})
[ordered]@{status=$(if($code-eq0-and$changed.Count-eq0){'PASS'}else{'FAIL'});mode=$mode;exitCode=$code;inputs=$inputs;changedInputs=$changed;harnessSha256=(Get-FileHash -LiteralPath $test).Hash.ToLowerInvariant();wordSha256=(Get-FileHash -LiteralPath $WordExecutable).Hash.ToLowerInvariant();normalWord='NOT TOUCHED';installedPackage='NOT RUN';publicTunnel='NOT RUN';scope=$(if($RealInput){'Explicit operator physical input through the genuine Word event subscription; no synthetic mouse or foreground automation.'}else{'Native bookmark/color navigation and stale authorization denials only; no foreground input or event execution.'})}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $OutputDirectory 'manifest.json') -Encoding utf8
if($changed.Count-ne0){throw 'Source changed during native event qualification.'}
if($code-ne0){throw 'Isolated native event qualification failed.'}
