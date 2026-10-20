param([string]$OutputDirectory, [string]$ZrokExecutable, [string]$NsisExecutable, [switch]$Installer)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repo 'artifacts/word/build' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$assembly = Join-Path $OutputDirectory 'wrAIter.Word.dll'
$sources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$cliSources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'cli') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName)
$contract = Join-Path $repo 'modules/review/selective-editor-contract.json'
$branding = Join-Path $repo 'assets/wraiter'
$picker = Join-Path $PSScriptRoot 'assets/document-picker.html'
$widgetInputs = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'widget') -File | Where-Object { $_.Name -ne 'test.mjs' } | Sort-Object Name | ForEach-Object FullName)
$widgetInputs += $picker
$widgetInputs += Join-Path $repo 'package-lock.json'
$widgetManifest = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'widget/build-manifest.json') | ConvertFrom-Json
if ((Get-FileHash -LiteralPath $picker).Hash.ToLowerInvariant() -ne $widgetManifest.sha256) { throw 'Rebuild the Word document picker before packaging.' }
foreach ($entry in $widgetManifest.inputs) {
    if ($entry.path.StartsWith('node_modules/')) { continue }
    if ((Get-FileHash -LiteralPath (Join-Path $repo $entry.path)).Hash.ToLowerInvariant() -ne $entry.sha256) {
        throw "Rebuild the Word document picker after changing $($entry.path)."
    }
}
$inputs = @($sources) + $cliSources + $widgetInputs + @(Get-ChildItem -LiteralPath $branding -File | Sort-Object Name | ForEach-Object FullName) + @($contract, $PSCommandPath, (Join-Path $PSScriptRoot 'README.md'), (Join-Path $PSScriptRoot 'cli/README.md'),
    (Join-Path $PSScriptRoot 'THIRD-PARTY.md'), (Join-Path $repo 'package.json'), (Join-Path $PSScriptRoot 'packaging/installer.nsi'),
    (Join-Path $PSScriptRoot 'packaging/zrok-LICENSE'), (Join-Path $PSScriptRoot 'packaging/zrok.json'), (Join-Path $PSScriptRoot 'packaging/fetch-zrok.ps1'))
$sourceSnapshot = @($inputs | ForEach-Object { [ordered]@{path=[IO.Path]::GetRelativePath($repo,$_);sha256=(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash.ToLowerInvariant()} })
$packageVersion = (Get-Content -Raw -LiteralPath (Join-Path $repo 'package.json') | ConvertFrom-Json).version
function Assert-SourceSnapshot {
    foreach ($entry in $sourceSnapshot) {
        if ((Get-FileHash -LiteralPath (Join-Path $repo $entry.path) -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.sha256) {
            throw "Source changed during the Word build: $($entry.path). Rebuild from stable inputs."
        }
    }
}
$arguments = @('/nologo','/target:library','/platform:anycpu','/optimize+','/debug-','/langversion:5',"/out:$assembly",
    "/resource:$contract,SelectiveEditorContract", "/resource:$picker,wordplai.DocumentPicker", "/resource:$(Join-Path $branding 'icon16.png'),wrAIter.Ribbon.png", "/resource:$(Join-Path $branding 'icon.ico'),wrAIter.Icon.ico",'/r:System.dll','/r:System.Core.dll','/r:System.Drawing.dll',
    '/r:System.Windows.Forms.dll','/r:System.Web.Extensions.dll','/r:System.Xml.dll','/r:System.Security.dll','/r:Microsoft.CSharp.dll') + $sources
& $compiler @arguments
if ($LASTEXITCODE -ne 0) { throw 'Word add-in compilation failed.' }
& $compiler /nologo /target:exe /platform:anycpu /optimize+ /debug- /langversion:5 "/out:$(Join-Path $OutputDirectory 'wordplai.exe')" /r:System.dll /r:System.Core.dll /r:System.Web.Extensions.dll /r:Microsoft.CSharp.dll @cliSources
if ($LASTEXITCODE -ne 0) { throw 'Word CLI compilation failed.' }
$runtime = Join-Path $OutputDirectory 'runtime'
$zrokPin = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'packaging/zrok.json') | ConvertFrom-Json
if ($Installer -and -not $ZrokExecutable) {
    $ZrokExecutable = & (Join-Path $PSScriptRoot 'packaging/fetch-zrok.ps1') -OutputDirectory (Join-Path $OutputDirectory 'dependencies/zrok')
}
if ($ZrokExecutable) {
    if ((Get-FileHash -LiteralPath $ZrokExecutable -Algorithm SHA256).Hash.ToLowerInvariant() -ne $zrokPin.executableSha256) {
        throw 'The supplied zrok executable does not match the pinned official release.'
    }
    if ((Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'packaging/zrok-LICENSE') -Algorithm SHA256).Hash.ToLowerInvariant() -ne $zrokPin.licenseSha256) {
        throw 'The bundled zrok license does not match the pinned release.'
    }
    New-Item -ItemType Directory -Force $runtime | Out-Null
    Copy-Item -LiteralPath $ZrokExecutable -Destination (Join-Path $runtime 'zrok.exe') -Force
}
$bundledZrok = Join-Path $runtime 'zrok.exe'
if ((Test-Path -LiteralPath $bundledZrok) -and
    (Get-FileHash -LiteralPath $bundledZrok -Algorithm SHA256).Hash.ToLowerInvariant() -ne $zrokPin.executableSha256) {
    throw 'The packaged zrok executable does not match the pinned release.'
}
$packagedReadme = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'README.md')).Replace('../../assets/wraiter/icon.svg', 'wrAIter.ico')
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'README.md'), $packagedReadme)
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'THIRD-PARTY.md') -Destination $OutputDirectory -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'cli/README.md') -Destination (Join-Path $OutputDirectory 'CLI.md') -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'packaging/zrok-LICENSE') -Destination $OutputDirectory -Force
Copy-Item -LiteralPath (Join-Path $branding 'icon.ico') -Destination (Join-Path $OutputDirectory 'wrAIter.ico') -Force
$payloadFiles = @('wrAIter.Word.dll','wordplai.exe','CLI.md','README.md','THIRD-PARTY.md','zrok-LICENSE','wrAIter.ico')
if (Test-Path -LiteralPath (Join-Path $runtime 'zrok.exe')) { $payloadFiles += 'runtime/zrok.exe' }
Assert-SourceSnapshot
$manifest = [ordered]@{ product='wordplai';version=$packageVersion;platform='Windows desktop Word';
    inputs=$sourceSnapshot;
    tunnel=$(if(Test-Path -LiteralPath $bundledZrok){$zrokPin}else{$null});
    files=@($payloadFiles | ForEach-Object { $payload=Get-Item -LiteralPath (Join-Path $OutputDirectory $_); [ordered]@{path=$_;bytes=$payload.Length;sha256=(Get-FileHash -LiteralPath $payload.FullName -Algorithm SHA256).Hash.ToLowerInvariant()} }) }
$manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'build-manifest.json') -Encoding utf8
if ($Installer) {
    if (-not (Test-Path -LiteralPath (Join-Path $runtime 'zrok.exe'))) { throw 'A release installer must contain its tunnel executable.' }
    $nsis = $NsisExecutable
    if (-not $nsis) { $nsis = (Get-Command makensis.exe -CommandType Application -ErrorAction SilentlyContinue).Source }
    if (-not $nsis -or -not (Test-Path -LiteralPath $nsis -PathType Leaf)) {
        throw 'Install NSIS and put makensis.exe on PATH, or pass -NsisExecutable with its full path.'
    }
    $setup = Join-Path (Split-Path $OutputDirectory -Parent) 'wordplai-setup.exe'
    & $nsis '/V2' "/DPAYLOAD=$OutputDirectory" "/DOUTPUT=$setup" "/DVERSION=$packageVersion" (Join-Path $PSScriptRoot 'packaging/installer.nsi')
    if ($LASTEXITCODE -ne 0) { throw 'Word installer build failed.' }
    Assert-SourceSnapshot
    Get-Item -LiteralPath $setup | Select-Object FullName,Length
    Get-FileHash -LiteralPath $setup -Algorithm SHA256
}
Write-Output "Built $assembly"
