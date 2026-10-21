param([Parameter(Mandatory)][string]$CliPath, [Parameter(Mandatory)][int]$WordProcessId, [Parameter(Mandatory)][string]$OutputDirectory, [switch]$IsolationCases, [switch]$AllowForegroundChanges, [string]$CommentAuthorName)
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -eq 'Core') {
    $child = @('-NoProfile','-STA','-File',$PSCommandPath,'-CliPath',$CliPath,'-WordProcessId',[string]$WordProcessId,'-OutputDirectory',$OutputDirectory)
    if ($IsolationCases) { $child += '-IsolationCases' }
    if ($AllowForegroundChanges) { $child += '-AllowForegroundChanges' }
    if ($CommentAuthorName) { $child += @('-CommentAuthorName',$CommentAuthorName) }
    & "$env:WINDIR/System32/WindowsPowerShell/v1.0/powershell.exe" @child
    if ($LASTEXITCODE -ne 0) { throw 'Installed CLI qualification failed.' }
    return
}
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding
$CliPath = (Resolve-Path -LiteralPath $CliPath).Path
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Choose a fresh installed CLI evidence directory.' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
Add-Type -ReferencedAssemblies 'Microsoft.CSharp','System.Core' -Path (Join-Path $PSScriptRoot 'WordAutomation.cs')
$filter = New-Object WordTestAutomation
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class CliNativeObservation {
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
}
'@
$checks = New-Object 'Collections.Generic.List[string]'
$unqualifiedForeground = New-Object 'Collections.Generic.List[string]'
$documents = @(); $initialHashes = @(); $sequence = 0; $started = $false
$originalCommentAuthor = $null
function Check($Condition, [string]$Message) { if (-not $Condition) { throw $Message }; $checks.Add($Message); Write-Output ('PASS: '+$Message) }
function Get-SavedHash([string]$Path) {
    $stream = [IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($hash.ComputeHash($stream)) }
    finally { $hash.Dispose(); $stream.Dispose() }
}
function Wait-Saved($Targets, [string]$Message) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while (@($Targets|Where-Object { -not $_.Saved }).Count -ne 0 -or $word.BackgroundSavingStatus -ne 0) {
        if ($timer.ElapsedMilliseconds -ge 15000) { throw ('Timeout: '+$Message) }
        Start-Sleep -Milliseconds 75
    }
    Check $true $Message
}
function Check-Foreground([IntPtr]$Expected, [string]$Message) {
    $same = [CliNativeObservation]::GetForegroundWindow() -eq $Expected
    if (-not $same -and $AllowForegroundChanges) {
        $unqualifiedForeground.Add($Message)
        Write-Output ('NOT QUALIFIED: '+$Message)
    } else { Check $same $Message }
}
function Trace-Position([string]$Stage) {
    if (-not $IsolationCases -or $null -eq $word -or -not $activePath) { return }
    [ordered]@{stage=$Stage;path=[string]$word.ActiveDocument.FullName;expectedPath=$activePath;start=[int]$word.Selection.Start;end=[int]$word.Selection.End;expectedStart=$selectionStart;expectedEnd=$selectionEnd;foreground=[CliNativeObservation]::GetForegroundWindow().ToInt64();expectedForeground=$foreground.ToInt64()}|ConvertTo-Json -Compress|Add-Content -Encoding UTF8 (Join-Path $OutputDirectory 'native-position-private.jsonl')
}
function Invoke-Cli([string[]]$Parameters, $Body = $null, [int]$ExpectedExit = 0) {
    $script:sequence++
    $prefix = Join-Path $OutputDirectory ('command-'+$script:sequence)
    $parameters = @($Parameters) + @('--word-pid',[string]$WordProcessId)
    if ($null -ne $Body) {
        $inputFile = $prefix+'-input-private.json'
        [IO.File]::WriteAllText($inputFile, ($Body|ConvertTo-Json -Depth 30 -Compress), (New-Object Text.UTF8Encoding($false)))
        $parameters += @('--input',$inputFile)
    }
    $stdout = & $CliPath @parameters 2> ($prefix+'-stderr.txt')
    $code = $LASTEXITCODE
    $text = $stdout -join "`n"
    [IO.File]::WriteAllText(($prefix+'-output-private.json'),$text,(New-Object Text.UTF8Encoding($false)))
    Trace-Position ('CLI-'+$script:sequence)
    if ($code -ne $ExpectedExit) { throw "CLI command $($script:sequence) exited $code; expected $ExpectedExit. Inspect its private receipt." }
    return ($text|ConvertFrom-Json)
}
function Native([string]$Command, [string]$Document, $Body = $null, [int]$ExpectedExit = 0) {
    return (Invoke-Cli @($Command,'--document',$Document) $Body $ExpectedExit).result
}
function Mcp([string]$Method, $Body = $null, [int]$ExpectedExit = 0, [string]$Session = 'session-private.json') {
    return (Invoke-Cli @('mcp','--method',$Method,'--session-file',(Join-Path $OutputDirectory $Session)) $Body $ExpectedExit).response
}
function Tool([string]$Name, $Body, [int]$ExpectedExit = 0, [string]$Session = 'session-private.json') {
    $reply = Mcp 'tools/call' @{name=$Name;arguments=$Body} $ExpectedExit $Session
    if ($ExpectedExit -ne 0) { return $reply }
    return ($reply.result.content[0].text|ConvertFrom-Json)
}
try {
    if ($WordProcessId -le 0) { throw 'Specify a positive Word process ID.' }
    $word = [WordTestAutomation]::AttachProcess($WordProcessId)
    $actual = [uint32]0
    [void][CliNativeObservation]::GetWindowThreadProcessId([IntPtr]$word.ActiveWindow.Hwnd,[ref]$actual)
    if ($actual -ne $WordProcessId) { throw 'Fixture creation refused a different Word process.' }
    $originalDocuments = @($word.Documents|ForEach-Object { @{path=[string]$_.FullName;text=[string]$_.Content.Text;saved=[bool]$_.Saved;comments=[int]$_.Comments.Count;revisions=[int]$_.Revisions.Count} })
    $originalDocuments | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 (Join-Path $OutputDirectory 'originals-before-private.json')
    if ($CommentAuthorName) {
        $authority = $word.COMAddIns.Item('UnderTree.WordAddin').Object
        $status = $authority.Execute('connection_status','{}') | ConvertFrom-Json
        if (-not $status.ok) { throw 'Cannot read the installed comment author preference.' }
        $originalCommentAuthor = [string]$status.result.commentAuthorName
        $wordUserName = [string]$word.UserName
        $saved = $authority.Execute('comment_author_save',(@{commentAuthorName=$CommentAuthorName}|ConvertTo-Json -Compress)) | ConvertFrom-Json
        Check ($saved.ok -and $saved.result.commentAuthorName -ceq $CommentAuthorName -and [string]$word.UserName -ceq $wordUserName) 'Installed comment author preference saves without changing the global Word user name'
    }
    $emoji = [char]::ConvertFromUtf32(0x1F600)
    $prompt = "An overly long $emoji sentence can be made shorter."
    foreach ($letter in @('A','B')) {
        $doc = $word.Documents.Add([Type]::Missing,$false,0,$false)
        $documents += $doc
        $doc.Content.Text = "Unmarked $letter introduction.`r$prompt`rProtected $letter ending.`r"
        $doc.Content.Font.Name = 'Segoe UI Emoji'
        $doc.Content.Font.NameBi = 'Segoe UI Emoji'
        $path = Join-Path $OutputDirectory ('CLI-fixture-'+$letter+'.docx')
        $doc.SaveAs2([ref][object]$path,[ref][object]16)
        Wait-Saved @($doc) 'The owned fixture completes its initial ordinary Save As'
        $initialHashes += Get-SavedHash $path
    }
    $foreground = [CliNativeObservation]::GetForegroundWindow()
    $activePath = [string]$word.ActiveDocument.FullName
    $selectionStart = [int]$word.Selection.Start; $selectionEnd = [int]$word.Selection.End
    Trace-Position 'baseline'
    $listing = Invoke-Cli @('documents')
    $targets = @($listing.result.documents|Where-Object title -eq 'CLI-fixture-A.docx')
    Check ($targets.Count -eq 1) 'Installed CLI discovers the exact owned fixture'
    $id = $targets[0].documentId
    $context = Native 'context' $id
    Check ($context.complete -and $context.text.Contains($emoji)) 'Installed CLI context preserves complete Unicode text'
    $state = Native 'rule-save' $id @{snapshot=$context.snapshot;id='cli-short';label='CLI short';color='#AACCEE';action='prompt';prompt='Shorten without changing meaning.'}
    $from = $context.text.IndexOf($prompt,[StringComparison]::Ordinal)
    $context = Native 'context' $id
    $state = Native 'mark' $id @{snapshot=$context.snapshot;ruleId='cli-short';from=$from;to=$from+$prompt.Length;quote=$prompt}
    Check (@($state.marks|Where-Object quote -ceq $prompt).Count -eq 1) 'Installed CLI marks an exact range without changing native selection'
    $refused = Invoke-Cli @('rule-save','--document',$id) @{snapshot='obsolete';id='unexpected';label='Wrong';color='#FF0000';action='protect'} 1
    Check (-not $refused.ok -and $refused.error.Contains('changed')) 'Installed CLI refuses a stale local snapshot'
    $state = Native 'state' $id
    Check (@($state.rules|Where-Object id -eq 'unexpected').Count -eq 0) 'Stale CLI refusal creates no rule'
    Check ([string]$word.ActiveDocument.FullName -ceq $activePath) 'Installed CLI restores the active document'
    Check ($word.Selection.Start -eq $selectionStart -and $word.Selection.End -eq $selectionEnd) 'Installed CLI preserves native selection'
    Check-Foreground $foreground 'Installed CLI preserves desktop foreground'
    Wait-Saved @($documents[0]) 'Local-only CLI changes to inactive A reach timer autosave before any MCP read'
    $localState = Native 'state' $id
    Check (-not $localState.document.shared -and (Get-SavedHash ([string]$documents[0].FullName)) -cne $initialHashes[0]) 'Local-only autosave changes saved bytes without sharing MCP authority'

    $status = (Invoke-Cli @('status')).result
    if (-not $status.running) { $null = Invoke-Cli @('connect-local'); $started = $true }
    $null = Mcp 'initialize'
    $catalogue = Mcp 'tools/list'
    Check (($catalogue.result.tools.name -join ',') -ceq 'list_manuscripts,read_manuscript,suggest_text,highlight_text') 'CLI MCP exposes exactly the existing four tools'
    $sessionHash = (Get-FileHash -LiteralPath (Join-Path $OutputDirectory 'session-private.json')).Hash
    $null = Mcp 'initialize' $null 2
    Check ((Get-FileHash -LiteralPath (Join-Path $OutputDirectory 'session-private.json')).Hash -ceq $sessionHash) 'CLI initialization never overwrites an existing client session'
    $read = Tool 'read_manuscript' @{manuscript_id=$id;offset=0}
    Check ($read.complete -and $read.text.Contains($emoji) -and $read.context_id) 'CLI reads actual MCP context and its complete-read witness'
    $request = @{manuscript_id=$id;request_id='cli-installed-direct';revision=$read.revision;context_id=$read.context_id;from=$from;to=$from+$prompt.Length;quote=$prompt;replacement="A shorter $emoji sentence.";rationale='Remove redundant wording.'}
    $foregroundPid = [uint32]0
    [void][CliNativeObservation]::GetWindowThreadProcessId([CliNativeObservation]::GetForegroundWindow(),[ref]$foregroundPid)
    if ($foregroundPid -ne $WordProcessId) {
        $applied = Tool 'suggest_text' $request
        Check ($applied.applied -and -not $applied.requiresApproval) 'CLI submits an authorized AI edit through the actual MCP safeguards'
        $replay = Tool 'suggest_text' $request
        Check ($replay.replayed -and $documents[0].Comments.Count -eq 1) 'CLI exact replay does not duplicate the native rationale'
        Check ($documents[0].Content.Text.Contains("A shorter $emoji sentence.") -and $documents[0].Revisions.Count -eq 0) 'CLI MCP applies exact native wording without tracked revisions'
    } else { Write-Output 'NOT RUN: positive direct write while Word holds foreground; guard preserved.' }
    Check ($documents[1].Content.Text.Contains($prompt) -and $documents[1].Comments.Count -eq 0) 'CLI writing leaves the other synthetic document unchanged'
    if ($IsolationCases -and $foregroundPid -ne $WordProcessId) {
        $second = @($listing.result.documents|Where-Object title -eq 'CLI-fixture-B.docx')
        Check ($second.Count -eq 1) 'Isolation fixture has a unique second document identity'
        $idB = $second[0].documentId
        $contextB = Native 'context' $idB
        $null = Native 'rule-save' $idB @{snapshot=$contextB.snapshot;id='cli-short';label='CLI short';color='#AACCEE';action='prompt';prompt='Shorten without changing meaning.'}
        $contextB = Native 'context' $idB
        $fromB = $contextB.text.IndexOf($prompt,[StringComparison]::Ordinal)
        $null = Native 'mark' $idB @{snapshot=$contextB.snapshot;ruleId='cli-short';from=$fromB;to=$fromB+$prompt.Length;quote=$prompt}
        Wait-Saved @($documents[1]) 'Local-only CLI changes to inactive B reach timer autosave before its MCP read'
        $localStateB = Native 'state' $idB
        Check (-not $localStateB.document.shared -and (Get-SavedHash ([string]$documents[1].FullName)) -cne $initialHashes[1]) 'Independent local-only B autosave changes saved bytes without MCP authority'
        $sessionB = 'session-b-private.json'
        $null = Mcp 'initialize' $null 0 $sessionB
        $choiceB = Tool 'read_manuscript' @{offset=0} 0 $sessionB
        Check ($choiceB.selection_required -and -not $choiceB.text -and -not $choiceB.context_id) 'A fresh unnamed client must choose among multiple documents before reading text'
        $readB = Tool 'read_manuscript' @{manuscript_id=$idB;offset=0} 0 $sessionB
        $retainedA = Tool 'read_manuscript' @{offset=0}
        $retainedB = Tool 'read_manuscript' @{offset=0} 0 $sessionB
        Check ($retainedA.manuscript_id -ceq $id -and $retainedA.text.Contains('Unmarked A introduction.')) 'Client A retains A after client B chooses B'
        Check ($retainedB.manuscript_id -ceq $idB -and $retainedB.text.Contains('Unmarked B introduction.')) 'Client B retains B across interleaved CLI processes'
        Check ($retainedA.context_id -and $retainedB.context_id -and $retainedA.context_id -cne $retainedB.context_id) 'Independent complete reads have distinct context witnesses'
        $beforeB = [string]$documents[1].Content.Text
        $writeB = @{manuscript_id=$idB;request_id='cli-isolation-b';revision=$retainedB.revision;context_id=$retainedA.context_id;from=$fromB;to=$fromB+$prompt.Length;quote=$prompt;replacement="Client B shortened $emoji wording.";rationale='Remove redundant wording in B.'}
        $denied = Tool 'suggest_text' $writeB 1 $sessionB
        Check ($denied.result.isError -and [string]$documents[1].Content.Text -ceq $beforeB -and $documents[1].Comments.Count -eq 0) 'Client A witness cannot authorize a write in client B'
        $freshB = Tool 'read_manuscript' @{manuscript_id=$idB;offset=0} 0 $sessionB
        $writeB.context_id = $freshB.context_id; $writeB.revision = $freshB.revision
        $appliedB = Tool 'suggest_text' $writeB 0 $sessionB
        Check ($appliedB.applied -and -not $appliedB.requiresApproval) 'Client B writes its authorized passage using its own witness and original request ID'
        $replayB = Tool 'suggest_text' $writeB 0 $sessionB
        Check ($replayB.replayed -and $documents[1].Comments.Count -eq 1 -and $documents[1].Revisions.Count -eq 0) 'Client B exact replay retains one rationale and no tracked revisions'
        Check ($documents[0].Content.Text.Contains("A shorter $emoji sentence.") -and -not $documents[0].Content.Text.Contains('Client B shortened') -and $documents[0].Comments.Count -eq 1) 'Interleaved client B writing leaves client A exact wording and rationale unchanged'
        $readA = Tool 'read_manuscript' @{offset=0}
        $intro = 'Unmarked A introduction.'
        $marking = @{requestId='cli-revoked-highlight';contextId=$readA.context_id;revision=$readA.revision;color='#445566';ranges=@(@{from=0;to=$intro.Length;quote=$intro})}
        $contextA = Native 'context' $id
        $beforeRevocation = Native 'state' $id
        $marksBefore = ConvertTo-Json -InputObject @($beforeRevocation.marks) -Depth 12 -Compress
        $rulesBefore = ConvertTo-Json -InputObject @($beforeRevocation.rules) -Depth 12 -Compress
        $null = Native 'revoke' $id @{snapshot=$contextA.snapshot}
        $denied = Tool 'highlight_text' @{documentId=$id;marking=$marking} 1
        $afterRevocation = Native 'state' $id
        Check ($denied.result.isError -and (ConvertTo-Json -InputObject @($afterRevocation.marks) -Depth 12 -Compress) -ceq $marksBefore -and (ConvertTo-Json -InputObject @($afterRevocation.rules) -Depth 12 -Compress) -ceq $rulesBefore) 'Local revocation invalidates the old witness without changing any marker or rule'
        $readA = Tool 'read_manuscript' @{offset=0}
        $marking.contextId = $readA.context_id; $marking.revision = $readA.revision
        $highlight = Tool 'highlight_text' @{documentId=$id;marking=$marking}
        $afterRecovery = Native 'state' $id
        Check (@($afterRecovery.marks|Where-Object quote -ceq $intro).Count -eq 1) 'Fresh complete context recovers the original refused highlight request identity'
        Wait-Saved $documents 'Actual timer autosave persists both named fixtures without a CLI Save command'
        Add-Type -AssemblyName System.IO.Compression
        foreach ($index in @(0,1)) {
            $stream = [IO.File]::Open([string]$documents[$index].FullName,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
            $zip = New-Object IO.Compression.ZipArchive($stream,[IO.Compression.ZipArchiveMode]::Read)
            try {
                $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
                try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
                $names = New-Object Xml.XmlNamespaceManager($xml.NameTable); $names.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                $savedText = -join @($xml.SelectNodes('//w:body//w:t',$names)|ForEach-Object InnerText)
                $expected = if($index -eq 0){"A shorter $emoji sentence."}else{"Client B shortened $emoji wording."}
                Check ($savedText.Contains($expected) -and -not $savedText.Contains($prompt) -and $xml.SelectNodes('//w:ins|//w:del',$names).Count -eq 0 -and $null -ne $zip.GetEntry('word/comments.xml')) 'Saved DOCX retains the authorized Unicode wording and rationale without revisions'
                if ($CommentAuthorName) {
                    $reader = New-Object IO.StreamReader($zip.GetEntry('word/comments.xml').Open())
                    try { [xml]$comments = $reader.ReadToEnd() } finally { $reader.Dispose() }
                    $commentNames = New-Object Xml.XmlNamespaceManager($comments.NameTable); $commentNames.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                    $authors = @($comments.SelectNodes('//w:comment/@w:author',$commentNames)|ForEach-Object Value)
                    Check ($authors.Count -eq 1 -and $authors[0] -ceq $CommentAuthorName -and [string]$documents[$index].Comments.Item(1).Author -ceq $CommentAuthorName) 'Both native and timer-saved MCP rationale comments use the configured author exactly'
                }
            } finally { $zip.Dispose(); $stream.Dispose() }
        }
        Check ([string]$word.ActiveDocument.FullName -ceq $activePath -and $word.Selection.Start -eq $selectionStart -and $word.Selection.End -eq $selectionEnd) 'Extended CLI operations preserve the original active document and native selection'
        Check-Foreground $foreground 'Extended CLI operations preserve desktop foreground'
        Trace-Position 'before-owned-close'
        $documents[1].Close(0); $documents[1] = $null
        Trace-Position 'after-owned-close'
        $closeTimer = [Diagnostics.Stopwatch]::StartNew()
        while (-not $AllowForegroundChanges -and [CliNativeObservation]::GetForegroundWindow() -ne $foreground) {
            if ($closeTimer.ElapsedMilliseconds -ge 15000) { throw 'Owned fixture closure did not restore desktop foreground; no input was injected.' }
            $null = [string]$word.ActiveDocument.FullName
            Start-Sleep -Milliseconds 75
        }
        Trace-Position 'owned-close-settled'
        $afterClosePath = [string]$word.ActiveDocument.FullName
        $afterCloseStart = [int]$word.Selection.Start; $afterCloseEnd = [int]$word.Selection.End
        $afterCloseForeground = [CliNativeObservation]::GetForegroundWindow()
        $closed = Tool 'read_manuscript' @{offset=0} 1 $sessionB
        $closedLocal = Invoke-Cli @('context','--document',$idB) $null 1
        Check ($closed.result.isError -and -not $closedLocal.ok -and $closedLocal.error.Contains('no longer open')) 'Closed client B target refuses through MCP and CLI without active-document fallback'
        $retainedA = Tool 'read_manuscript' @{offset=0}
        Check ($retainedA.manuscript_id -ceq $id -and $documents[0].Comments.Count -eq 1) 'Closing B preserves client A target and its single rationale'
        Check ([string]$word.ActiveDocument.FullName -ceq $afterClosePath -and $word.Selection.Start -eq $afterCloseStart -and $word.Selection.End -eq $afterCloseEnd) 'Closed-target CLI requests preserve the native document and selection left by Word fixture closure'
        Check-Foreground $afterCloseForeground 'Closed-target CLI requests preserve desktop foreground'
    }
    $originalsAfter = @($word.Documents|ForEach-Object { @{path=[string]$_.FullName;text=[string]$_.Content.Text;saved=[bool]$_.Saved;comments=[int]$_.Comments.Count;revisions=[int]$_.Revisions.Count} })
    $originalsAfter | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 (Join-Path $OutputDirectory 'originals-after-private.json')
    foreach ($item in $originalDocuments) {
        $current = @($word.Documents|Where-Object { [string]$_.FullName -ceq $item.path })
        if ($current.Count -ne 1 -or [string]$current[0].Content.Text -cne $item.text -or [bool]$current[0].Saved -ne $item.saved -or $current[0].Comments.Count -ne $item.comments -or $current[0].Revisions.Count -ne $item.revisions) { throw 'An original document changed during CLI qualification.' }
    }
    Check $true 'All pre-existing documents retain exact native text, saved state, comments and revisions'
} catch {
    [ordered]@{status='FAIL';checks=$checks;error=$_.Exception.Message}|ConvertTo-Json -Depth 6|Set-Content -Encoding UTF8 (Join-Path $OutputDirectory 'result.json')
    throw
} finally {
    if ($started) { $null = Invoke-Cli @('stop') }
    foreach ($doc in $documents) { if ($null -ne $doc) { $doc.Close(0) } }
    if ($null -ne $originalCommentAuthor) {
        $restored = $authority.Execute('comment_author_save',(@{commentAuthorName=$originalCommentAuthor}|ConvertTo-Json -Compress)) | ConvertFrom-Json
        if (-not $restored.ok -or $restored.result.commentAuthorName -cne $originalCommentAuthor) { throw 'The installed comment author preference could not be restored.' }
    }
    if ($filter) { $filter.Dispose() }
}
[ordered]@{status=$(if($foregroundPid -ne $WordProcessId -and $unqualifiedForeground.Count -eq 0){'PASS'}else{'PARTIAL'});checks=$checks;unqualifiedForegroundChecks=$unqualifiedForeground;cliSha256=(Get-FileHash -LiteralPath $CliPath).Hash;wordProcessId=$WordProcessId;isolationCasesRequested=[bool]$IsolationCases;positiveMcpWrite=$(if($foregroundPid -ne $WordProcessId){'PASS'}else{'NOT RUN'});userInterfaceInput='NONE';publicTunnel='NOT RUN'}|ConvertTo-Json -Depth 6|Set-Content -Encoding UTF8 (Join-Path $OutputDirectory 'result.json')
exit 0
