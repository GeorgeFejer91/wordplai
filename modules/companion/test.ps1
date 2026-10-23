param([string]$NsisExecutable = (Join-Path $env:LOCALAPPDATA 'tauri/NSIS/Bin/makensis.exe'))
$ErrorActionPreference='Stop'
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$evidence=Join-Path $repo ('artifacts/companion-tests/'+[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Force -Path $evidence | Out-Null
$compiler=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$source=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'installer.nsi') -Raw
$checks=@()
foreach($case in @(@{name='success';word=0;lite=0;exit=0;calls="word`nlite`n"},@{name='word-failure';word=7;lite=0;exit=1;calls="word`n"},@{name='lite-failure';word=0;lite=9;exit=1;calls="word`nlite`n"})){
  $folder=Join-Path $evidence $case.name
  New-Item -ItemType Directory -Path $folder | Out-Null
  $marker=Join-Path $folder 'calls.txt'
  foreach($component in @('word','lite')){
    $code='using System.IO; class Fixture { static int Main() { File.AppendAllText(@"'+$marker.Replace('"','""')+'", "'+$component+'\n"); return '+$case[$component]+'; } }'
    $cs=Join-Path $folder ($component+'.cs');$code | Set-Content -LiteralPath $cs
    & $compiler /nologo /target:winexe "/out:$folder/$component.exe" $cs
    if($LASTEXITCODE -ne 0){throw 'Fixture compilation failed.'}
  }
  # Only the fixture bypasses the production open-Word guard.
  $fixtureSource=Join-Path $folder 'installer.nsi'
  $source.Replace('FindWindow $0 "OpusApp"','FindWindow $0 "WordplaiNonexistentFixtureWindow"') | Set-Content -LiteralPath $fixtureSource -Encoding utf8
  $setup=Join-Path $folder 'setup.exe'
  & $NsisExecutable /V1 "/DOUTPUT=$setup" "/DWORD_INSTALLER=$(Join-Path $folder 'word.exe')" "/DLITE_INSTALLER=$(Join-Path $folder 'lite.exe')" "/DICON=$(Join-Path $repo 'assets/wraiter/icon.ico')" $fixtureSource
  if($LASTEXITCODE -ne 0){throw 'Fixture installer compilation failed.'}
  $process=Start-Process -FilePath $setup -ArgumentList '/S' -WindowStyle Hidden -PassThru -Wait
  if($process.ExitCode -ne $case.exit){throw "$($case.name): wrong exit code $($process.ExitCode)."}
  if((Get-Content -LiteralPath $marker -Raw) -cne $case.calls){throw "$($case.name): component invocation order/count mismatch."}
  $checks+=@{name=$case.name;exitCode=$process.ExitCode;passed=$true}
}
if(Get-Process WINWORD -ErrorAction SilentlyContinue){
  $folder=Join-Path $evidence 'open-word';New-Item -ItemType Directory -Path $folder | Out-Null
  $setup=Join-Path $folder 'setup.exe'
  & $NsisExecutable /V1 "/DOUTPUT=$setup" "/DWORD_INSTALLER=$(Join-Path $evidence 'success/word.exe')" "/DLITE_INSTALLER=$(Join-Path $evidence 'success/lite.exe')" "/DICON=$(Join-Path $repo 'assets/wraiter/icon.ico')" (Join-Path $PSScriptRoot 'installer.nsi')
  if($LASTEXITCODE -ne 0){throw 'Preflight installer compilation failed.'}
  $before=Get-Content -LiteralPath (Join-Path $evidence 'success/calls.txt') -Raw
  $process=Start-Process -FilePath $setup -ArgumentList '/S' -WindowStyle Hidden -PassThru -Wait
  if($process.ExitCode -ne 2 -or (Get-Content -LiteralPath (Join-Path $evidence 'success/calls.txt') -Raw) -cne $before){throw 'Open-Word preflight failed.'}
  $checks+=@{name='production-open-word-preflight';exitCode=$process.ExitCode;passed=$true}
}
$result=@{result='PASS';checks=$checks}
$result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $evidence 'result.json')
$result | ConvertTo-Json -Depth 5
