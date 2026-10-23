param(
  [Parameter(Mandatory)][string]$WordInstaller,
  [Parameter(Mandatory)][string]$LiteInstaller,
  [Parameter(Mandatory)][ValidatePattern('^[a-fA-F0-9]{64}$')][string]$LiteSha256,
  [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{40}$')][string]$LiteSourceRevision,
  [string]$OutputDirectory,
  [string]$NsisExecutable
)
$ErrorActionPreference='Stop'
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
if(-not $OutputDirectory){$OutputDirectory=Join-Path $repo 'artifacts/companion-release'}
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$word=(Resolve-Path -LiteralPath $WordInstaller).Path
$lite=(Resolve-Path -LiteralPath $LiteInstaller).Path
$release=Get-Content -LiteralPath (Join-Path $repo 'site/release.json') -Raw | ConvertFrom-Json
if($release.version -cne 'v0.1.0-preview.8' -or $release.sha256 -cne 'e0f8d9456798d8507021a893e48421ace9c8ae7ea67cc3104b152ab75a8de423'){throw 'Requalify a changed Word component before packaging it.'}
if((Get-FileHash -LiteralPath $word -Algorithm SHA256).Hash.ToLowerInvariant() -cne $release.sha256){throw 'Word installer does not match its published release.'}
if((Get-FileHash -LiteralPath $lite -Algorithm SHA256).Hash.ToLowerInvariant() -cne $LiteSha256.ToLowerInvariant()){throw 'Lite installer does not match the reviewed checksum.'}
if(-not $NsisExecutable){$NsisExecutable=Join-Path $env:LOCALAPPDATA 'tauri/NSIS/Bin/makensis.exe'}
if(-not(Test-Path -LiteralPath $NsisExecutable)){throw 'NSIS is required to build the combined installer.'}
$output=Join-Path $OutputDirectory 'wordplai-and-secret-tunnel-setup.exe'
& $NsisExecutable /V2 "/DOUTPUT=$output" "/DWORD_INSTALLER=$word" "/DLITE_INSTALLER=$lite" "/DICON=$(Join-Path $repo 'assets/wraiter/icon.ico')" (Join-Path $PSScriptRoot 'installer.nsi')
if($LASTEXITCODE -ne 0){throw 'Combined installer compilation failed.'}
$components=@(
  [ordered]@{product='wordplai';version=$release.version;asset='wordplai-setup.exe';sha256=$release.sha256;bytes=(Get-Item -LiteralPath $word).Length},
  [ordered]@{product='Secret Tunnel Lite';version='5.0.0';asset='secret-tunnel-lite-setup.exe';sha256=$LiteSha256.ToLowerInvariant();bytes=(Get-Item -LiteralPath $lite).Length;sourceRevision=$LiteSourceRevision}
)
$manifest=[ordered]@{product='Wordplai and Secret Tunnel Lite';version='0.1.0-preview.1';components=$components;asset='wordplai-and-secret-tunnel-setup.exe';sha256=(Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash.ToLowerInvariant();bytes=(Get-Item -LiteralPath $output).Length;inputs=@(Get-ChildItem -LiteralPath $PSScriptRoot -File | Sort-Object Name | ForEach-Object {[ordered]@{path=$_.Name;sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()}})}
$manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'bundle-manifest.json') -Encoding utf8
$manifest | Select-Object product,version,asset,sha256,bytes | ConvertTo-Json -Compress
