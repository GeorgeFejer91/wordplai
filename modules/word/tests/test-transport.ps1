param(
    [string]$OutputDirectory,
    [switch]$MutationCheck,
    [switch]$LogicalOnly
)
$ErrorActionPreference = 'Stop'
$wordRepository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $wordRepository ('artifacts\word-transport-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
$wordOutput = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $wordOutput) { throw 'Use a new output directory to preserve earlier transport evidence.' }
New-Item -ItemType Directory -Path $wordOutput | Out-Null
$wordCompiler = Join-Path $env:SystemRoot 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $wordCompiler)) {
    $wordCompiler = Join-Path $env:SystemRoot 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
$wordCommon = @('/nologo', '/target:exe', '/langversion:5', '/r:System.Web.Extensions.dll', '/r:System.Core.dll', '/r:System.Security.dll',
    "/resource:$(Join-Path $wordRepository 'modules/word/assets/document-picker.html'),wordplai.DocumentPicker",
    "/resource:$(Join-Path $wordRepository 'modules/review/selective-editor-contract.json'),SelectiveEditorContract")
$wordSources = @('ConnectionService.cs', 'ConnectionProfile.cs', 'McpEndpoint.cs', 'McpCoordination.cs', 'WordDocumentPicker.cs', 'WordNativeOutcome.cs', 'TunnelProcess.cs', 'TunnelSupervisor.cs') |
    ForEach-Object { Join-Path $wordRepository "modules\word\src\$_" }
$wordFake = Join-Path $wordOutput 'FakeZrok.exe'
$wordTest = Join-Path $wordOutput 'TransportTests.exe'
& $wordCompiler @wordCommon "/out:$wordFake" (Join-Path $PSScriptRoot 'FakeZrok.cs') *> (Join-Path $wordOutput 'compile-fake.log')
if ($LASTEXITCODE -ne 0) { throw "Fake runtime compilation failed; see $wordOutput\compile-fake.log" }
$wordTestSources = @((Join-Path $PSScriptRoot 'TransportTests.cs'), (Join-Path $PSScriptRoot 'McpConcurrencyTests.cs'), (Join-Path $PSScriptRoot 'LogicalConversationTests.cs'))
$wordInputPaths = @($wordSources) + @($wordTestSources) + @((Join-Path $PSScriptRoot 'FakeZrok.cs'), $PSCommandPath,
    (Join-Path $wordRepository 'modules/word/assets/document-picker.html'), (Join-Path $wordRepository 'modules/review/selective-editor-contract.json'))
$wordInputs = @($wordInputPaths | ForEach-Object { [ordered]@{path=[IO.Path]::GetRelativePath($wordRepository,$_);sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()} })
& $wordCompiler @wordCommon "/out:$wordTest" @wordSources @wordTestSources *> (Join-Path $wordOutput 'compile-tests.log')
if ($LASTEXITCODE -ne 0) { throw "Transport compilation failed; see $wordOutput\compile-tests.log" }
$wordRunArguments = @((Join-Path $wordOutput 'profiles'), $wordFake)
if ($LogicalOnly) { $wordRunArguments += 'logical-only' }
& $wordTest @wordRunArguments 2>&1 | Tee-Object -FilePath (Join-Path $wordOutput 'transport.log')
if ($LASTEXITCODE -ne 0) { throw "Transport tests failed; see $wordOutput\transport.log" }

if ($MutationCheck) {
    $wordMutation = Join-Path $wordOutput 'mutation'
    New-Item -ItemType Directory -Path $wordMutation | Out-Null
    $wordMutatedSource = Join-Path $wordMutation 'McpEndpoint.cs'
    $wordOriginal = [IO.File]::ReadAllText((Join-Path $wordRepository 'modules\word\src\McpEndpoint.cs'))
    $wordNeedle = 'headers.TryGetValue("Origin", out origin) ||'
    if (-not $wordOriginal.Contains($wordNeedle)) { throw 'Origin validation mutation no longer matches production; review the mutation.' }
    [IO.File]::WriteAllText($wordMutatedSource, $wordOriginal.Replace($wordNeedle, '(headers.TryGetValue("Origin", out origin) && false) ||'))
    $wordMutationSources = $wordSources | Where-Object { [IO.Path]::GetFileName($_) -ne 'McpEndpoint.cs' }
    $wordMutationTest = Join-Path $wordMutation 'TransportTests.exe'
    & $wordCompiler @wordCommon "/out:$wordMutationTest" @wordMutationSources $wordMutatedSource @wordTestSources *> (Join-Path $wordMutation 'compile.log')
    if ($LASTEXITCODE -ne 0) { throw 'The transport mutation did not compile.' }
    & $wordMutationTest (Join-Path $wordMutation 'profiles') $wordFake *> (Join-Path $wordMutation 'transport.log')
    if ($LASTEXITCODE -eq 0 -or -not ([IO.File]::ReadAllText((Join-Path $wordMutation 'transport.log')).Contains('Browser Origin denied.'))) {
        throw 'The transport suite did not reject the intended Origin security mutation.'
    }
    Write-Output 'PASS: test suite rejects disabled Origin security control.'
    $wordScopeMutation = Join-Path $wordOutput 'session-isolation-mutation'
    New-Item -ItemType Directory -Path $wordScopeMutation | Out-Null
    $wordCoordinatorSource = Join-Path $wordRepository 'modules/word/src/McpCoordination.cs'
    $wordCoordinator = [IO.File]::ReadAllText($wordCoordinatorSource)
    $wordScopeNeedle = 'NativeScope = (conversation == null ? session.ScopePrefix : conversation.ScopePrefix) + "default";'
    if (-not $wordCoordinator.Contains($wordScopeNeedle)) { throw 'Session isolation mutation no longer matches production; review the mutation.' }
    $wordMutatedCoordinator = Join-Path $wordScopeMutation 'McpCoordination.cs'
    [IO.File]::WriteAllText($wordMutatedCoordinator, $wordCoordinator.Replace($wordScopeNeedle, 'NativeScope = "shared-default";'))
    $wordScopeSources = $wordSources | Where-Object { [IO.Path]::GetFileName($_) -ne 'McpCoordination.cs' }
    $wordScopeTest = Join-Path $wordScopeMutation 'TransportTests.exe'
    & $wordCompiler @wordCommon "/out:$wordScopeTest" @wordScopeSources $wordMutatedCoordinator @wordTestSources *> (Join-Path $wordScopeMutation 'compile.log')
    if ($LASTEXITCODE -ne 0) { throw 'The session isolation mutation did not compile.' }
    & $wordScopeTest (Join-Path $wordScopeMutation 'profiles') $wordFake 'concurrency-only' *> (Join-Path $wordScopeMutation 'transport.log')
    if ($LASTEXITCODE -eq 0 -or -not ([IO.File]::ReadAllText((Join-Path $wordScopeMutation 'transport.log')).Contains('ChatGPT, Claude and a second GPT with the same client name have isolated native identities.'))) {
        throw 'The concurrent-client suite did not reject the intended shared-default identity mutation.'
    }
    Write-Output 'PASS: concurrent-client tests reject shared client identity in an isolated copied build.'
    foreach ($wordCase in @(
        @{name='ticket-zero-mutation';file='McpEndpoint.cs';needle='if (ticket > 0) metadata.Add';replacement='if (ticket >= 0) metadata.Add';failure='A pre-admission validation refusal carries no fabricated coordination ticket.'},
        @{name='uncertain-persistence-mutation';file='WordNativeOutcome.cs';needle='try { persist(); }';replacement='persist(); try { }';failure='Unconfirmed native effects remain unknown when uncertain-receipt persistence fails.'},
        @{name='lost-native-reply-mutation';file='McpEndpoint.cs';needle='if (nativeReplyReceived && (name == "suggest_text" || name == "highlight_text"))';replacement='if (false && nativeReplyReceived && (name == "suggest_text" || name == "highlight_text"))';failure='Response construction failure after a native commit reports uncertain delivery while retaining its tracked receipt.'},
        @{name='native-first-failure-mutation';file='WordNativeOutcome.cs';needle='if (failed) return;';replacement='if (false) return;';failure='Native diagnostics retain the first stage, exception type and HRESULT across rollback, cleanup and callback failures.'},
        @{name='native-false-rollback-mutation';file='WordNativeOutcome.cs';needle='rollbackConfirmed = rollback();';replacement='rollbackConfirmed = rollback() || true;';failure='A false rollback result is never reported as confirmed.'}
    )) {
        $wordCaseDirectory = Join-Path $wordOutput $wordCase.name
        New-Item -ItemType Directory -Path $wordCaseDirectory | Out-Null
        $wordCaseOriginal = [IO.File]::ReadAllText((Join-Path $wordRepository ('modules/word/src/' + $wordCase.file)))
        if (-not $wordCaseOriginal.Contains($wordCase.needle)) { throw "The $($wordCase.name) control no longer matches production." }
        $wordCaseSource = Join-Path $wordCaseDirectory $wordCase.file
        [IO.File]::WriteAllText($wordCaseSource, $wordCaseOriginal.Replace($wordCase.needle, $wordCase.replacement))
        $wordCaseSources = $wordSources | Where-Object { [IO.Path]::GetFileName($_) -ne $wordCase.file }
        $wordCaseTest = Join-Path $wordCaseDirectory 'TransportTests.exe'
        & $wordCompiler @wordCommon "/out:$wordCaseTest" @wordCaseSources $wordCaseSource @wordTestSources *> (Join-Path $wordCaseDirectory 'compile.log')
        if ($LASTEXITCODE -ne 0) { throw "The $($wordCase.name) control did not compile." }
        & $wordCaseTest (Join-Path $wordCaseDirectory 'profiles') $wordFake 'concurrency-only' *> (Join-Path $wordCaseDirectory 'transport.log')
        if ($LASTEXITCODE -eq 0 -or -not ([IO.File]::ReadAllText((Join-Path $wordCaseDirectory 'transport.log')).Contains($wordCase.failure))) {
            throw "The transport test did not reject the intended $($wordCase.name) defect."
        }
        Write-Output "PASS: copied $($wordCase.name) control fails the intended regression."
    }
    foreach ($wordCase in @(
        @{name='logical-transport-owner-mutation';file='McpCoordination.cs';needle='NativeScope = (conversation == null ? session.ScopePrefix : conversation.ScopePrefix) + "default";';replacement='NativeScope = session.ScopePrefix + "default";';failure='A labelled conversation retains its native selection and complete-read ownership across fresh transports.'},
        @{name='logical-shared-hint-mutation';file='McpCoordination.cs';needle='string key = Digest(hint);';replacement='string key = "all-conversations";';failure='Two different provider conversations keep distinct random native owner scopes across their own transports.'},
        @{name='logical-physical-expiry-mutation';file='McpCoordination.cs';needle='RetireSession(session, false);';replacement='RetireSession(session, true);';failure="Physical idle expiry cancels its immutable call handle without invalidating a still-active logical owner's witness namespace."},
        @{name='logical-liveness-mutation';file='McpCoordination.cs';needle='return active && scope.Session.Active && (scope.Conversation == null || scope.Conversation.Active);';replacement='return active && scope.Session.Active;';failure='Logical idle expiry independently revokes an old conversation while its transport and another label stay live.'},
        @{name='logical-cleanup-mutation';file='McpCoordination.cs';needle='revoke(scope.Prefix);';replacement='try { revoke(scope.Prefix); } catch { }';failure='Failed native retirement remains queued for retry without reviving its logical handle.'},
        @{name='logical-native-admission-mutation';file='McpEndpoint.cs';needle='if (!active || !coordination.IsActive(scope) || Stopwatch.GetTimestamp() > deadline)';replacement='if (!active || !coordination.IsActive(session) || Stopwatch.GetTimestamp() > deadline)';failure='Native admission rechecks the original immutable logical handle and never dispatches a revoked generation or its fresh replacement.'},
        @{name='logical-final-delivery-mutation';file='McpEndpoint.cs';needle=' || callScope != null && !coordination.IsActive(callScope)';replacement='';failure='Logical revocation with a still-live physical transport withholds manuscript and current-context data.'}
    )) {
        $wordCaseDirectory = Join-Path $wordOutput $wordCase.name
        New-Item -ItemType Directory -Path $wordCaseDirectory | Out-Null
        $wordCaseOriginal = [IO.File]::ReadAllText((Join-Path $wordRepository ('modules/word/src/' + $wordCase.file)))
        if (-not $wordCaseOriginal.Contains($wordCase.needle)) { throw "The $($wordCase.name) control no longer matches production." }
        $wordCaseSource = Join-Path $wordCaseDirectory $wordCase.file
        [IO.File]::WriteAllText($wordCaseSource, $wordCaseOriginal.Replace($wordCase.needle, $wordCase.replacement))
        $wordCaseSources = $wordSources | Where-Object { [IO.Path]::GetFileName($_) -ne $wordCase.file }
        $wordCaseTest = Join-Path $wordCaseDirectory 'TransportTests.exe'
        & $wordCompiler @wordCommon "/out:$wordCaseTest" @wordCaseSources $wordCaseSource @wordTestSources *> (Join-Path $wordCaseDirectory 'compile.log')
        if ($LASTEXITCODE -ne 0) { throw "The $($wordCase.name) control did not compile." }
        & $wordCaseTest (Join-Path $wordCaseDirectory 'profiles') $wordFake 'logical-only' *> (Join-Path $wordCaseDirectory 'transport.log')
        if ($LASTEXITCODE -eq 0 -or -not ([IO.File]::ReadAllText((Join-Path $wordCaseDirectory 'transport.log')).Contains($wordCase.failure))) {
            throw "The logical-conversation suite did not reject the intended $($wordCase.name) defect."
        }
        Write-Output "PASS: copied $($wordCase.name) control fails the intended regression."
    }
}
$wordChanged = @($wordInputs | Where-Object { (Get-FileHash -LiteralPath (Join-Path $wordRepository $_.path)).Hash.ToLowerInvariant() -ne $_.sha256 })
[ordered]@{status=$(if($wordChanged.Count -eq 0){'PASS'}else{'FAIL'});inputs=$wordInputs;changedInputs=$wordChanged;assemblySha256=(Get-FileHash -LiteralPath $wordTest).Hash.ToLowerInvariant();nativeWord='NOT RUN';mutationCheck=[bool]$MutationCheck;logicalOnly=[bool]$LogicalOnly} |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $wordOutput 'manifest.json') -Encoding utf8
if ($wordChanged.Count -ne 0) { throw 'Word transport inputs changed during qualification.' }
Write-Output "Transport evidence: $wordOutput"
