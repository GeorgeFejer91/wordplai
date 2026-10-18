param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repo 'artifacts/word/dependencies/zrok-2.0.4' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
$pin = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'zrok.json') | ConvertFrom-Json
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$archive = Join-Path $OutputDirectory 'zrok_2.0.4_windows_amd64.tar.gz'
if (-not (Test-Path -LiteralPath $archive)) {
    Invoke-WebRequest -Uri $pin.archiveUrl -OutFile $archive -TimeoutSec 120
}
if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant() -ne $pin.archiveSha256) {
    throw 'The zrok archive does not match the pinned upstream SHA-256. Use a fresh output directory.'
}
$tar = (Get-Command tar -CommandType Application -ErrorAction Stop).Source
& $tar -xzf $archive -C $OutputDirectory $pin.executable LICENSE
if ($LASTEXITCODE -ne 0) { throw 'The verified zrok archive could not be extracted.' }
$executable = Join-Path $OutputDirectory $pin.executable
if ((Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash.ToLowerInvariant() -ne $pin.executableSha256 -or
    (Get-FileHash -LiteralPath (Join-Path $OutputDirectory 'LICENSE') -Algorithm SHA256).Hash.ToLowerInvariant() -ne $pin.licenseSha256) {
    throw 'The extracted zrok executable or license does not match the pinned release.'
}
$executable
