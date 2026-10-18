param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -eq 'Core') {
    & "$env:WINDIR/System32/WindowsPowerShell/v1.0/powershell.exe" -NoProfile -STA -ExecutionPolicy Bypass -File $PSCommandPath -OutputDirectory $OutputDirectory
    if ($LASTEXITCODE -ne 0) { throw 'Installed Word qualification did not pass.' }
    return
}
if (-not $OutputDirectory) { $OutputDirectory = Join-Path (Resolve-Path "$PSScriptRoot/../../..") ('artifacts/word/native-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Choose a new evidence directory to preserve previous Word checks.' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
Add-Type -ReferencedAssemblies 'Microsoft.CSharp','System.Core' -Path (Join-Path $PSScriptRoot 'WordAutomation.cs')
$automation = New-Object WordTestAutomation
$checks = New-Object Collections.Generic.List[string]
$fixture = $null
$api = $null
$address = $null
$rpcId = 0
function Check($condition, [string]$message) { if (-not $condition) { throw $message }; $checks.Add($message); Write-Host ("PASS: " + $message) }
function WaitFor([scriptblock]$ready, [string]$message) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    do {
        if (& $ready) { Check $true $message; return }
        if ($timer.ElapsedMilliseconds -ge 15000) { throw ("Timeout: " + $message) }
        Start-Sleep -Milliseconds 75
    } while ($true)
}
function WaitSaved([string]$message) {
    WaitFor { $fixture.Saved -and $word.BackgroundSavingStatus -eq 0 } $message
}
function NativeGreen([string]$quote) {
    $text = [WordTestAutomation]::Read($fixture)
    $from = $text.IndexOf($quote, [StringComparison]::Ordinal)
    if ($from -lt 0) { return $false }
    $range = $fixture.Range($from, $from + $quote.Length)
    return ($range.Font.Shading.BackgroundPatternColor -eq 10795380 -and $range.HighlightColorIndex -eq 0)
}
function CheckGreenProtected([string]$quote) {
    $current = Local 'state'
    $greenRules = @($current.rules | Where-Object { $_.action -eq 'protect' -and $_.color -eq '#74b9a4' } | ForEach-Object { $_.id })
    $marks = @($current.marks | Where-Object { $greenRules -contains $_.ruleId -and -not $_.stale -and $_.quote -ceq $quote })
    Check ($marks.Count -eq 1 -and (NativeGreen $quote)) 'The actual green passage has one current native Preserve text marker'
}
function DocumentHash([string]$path) {
    $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($stream)) }
    finally { $sha.Dispose(); $stream.Dispose() }
}
function Local([string]$operation, $arguments = @{}) {
    $reply = $api.Execute($operation, (ConvertTo-Json -InputObject $arguments -Depth 20 -Compress)) | ConvertFrom-Json
    if (-not $reply.ok) { throw $reply.error }
    return $reply.result
}
function Rpc([string]$method, $arguments) {
    $script:rpcId++
    $request = [Net.HttpWebRequest]::Create($address)
    $request.Method = 'POST'; $request.ContentType = 'application/json'; $request.Accept = 'application/json, text/event-stream'
    $request.Timeout = 20000; $request.ReadWriteTimeout = 20000; $request.KeepAlive = $false
    if ($method -ne 'initialize' -and $script:mcpSession) { $request.Headers['Mcp-Session-Id'] = $script:mcpSession }
    $body = @{jsonrpc='2.0';id=$script:rpcId;method=$method;params=$arguments} | ConvertTo-Json -Depth 20 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($body); $request.ContentLength = $bytes.Length
    $stream = $request.GetRequestStream(); $stream.Write($bytes,0,$bytes.Length); $stream.Dispose()
    $response = $request.GetResponse()
    if ($method -eq 'initialize') { $script:mcpSession = $response.Headers['Mcp-Session-Id'] }
    try { $reader = New-Object IO.StreamReader($response.GetResponseStream()); $json = $reader.ReadToEnd() | ConvertFrom-Json }
    finally { $response.Dispose() }
    if ($json.error) { throw $json.error.message }
    return $json.result
}
function Tool([string]$name, $arguments = @{}, [switch]$Denied) {
    $reply = Rpc 'tools/call' @{name=$name;arguments=$arguments}
    if ($Denied) { Check ($reply.isError -eq $true) "$name denied without a document mutation"; return }
    if ($reply.isError) { throw $reply.content[0].text }
    return ($reply.content[0].text | ConvertFrom-Json)
}
function Mark([string]$quote, [string]$rule) {
    $text = [WordTestAutomation]::Read($fixture); $from = $text.IndexOf($quote, [StringComparison]::Ordinal)
    if ($from -lt 0) { throw 'Qualification passage is missing.' }
    $fixture.Range($from, $from+$quote.Length).Select()
    $proof = Local 'selection_preview'
    Local 'mark_selection' @{ruleId=$rule;prompt='Preserve the meaning.';quote=$proof.quote;from=$proof.from;to=$proof.to;snapshot=$proof.snapshot} | Out-Null
}
function ReadContext {
    $chunks = New-Object Text.StringBuilder; $offset = 0
    do {
        $page = Tool 'read_manuscript' @{manuscript_id=$documentId;offset=$offset}
        [void]$chunks.Append($page.text); $offset=$page.to
    } while (-not $page.complete)
    $completeText = $chunks.ToString()
    return @{text=$completeText;main_text=$completeText.Substring(0,$page.main_body_length);revision=$page.revision;context_id=$page.context_id}
}
try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    $addin = $word.COMAddIns.Item('UnderTree.WordAddin')
    Check $addin.Connect 'Word loaded the installed native add-in'
    $api = $addin.Object; Check ($null -ne $api) 'Word exposed the local add-in interface'
    $fixture = $word.Documents.Add()
    $emoji = [char]::ConvertFromUtf32(0x1F600)
    $fixture.Content.Text = "Unmarked introduction stays unchanged.`rAlpha $emoji beta gamma.`rDirect passage wording.`rProtected passage stays unchanged.`rRevocation probe wording.`rDiscard this unnecessary idea.`rFootnote context anchor.`r"
    $fixture.Content.Font.Name = 'Segoe UI Emoji'
    $fixture.Content.Font.NameBi = 'Segoe UI Emoji'
    $fixture.Sections.Item(1).Headers.Item(1).Range.Text = 'Read-only header context.'
    [void]$fixture.Footnotes.Add($fixture.Range($fixture.Content.End-2,$fixture.Content.End-2),[Type]::Missing,'Read-only footnote context.')
    $body = [WordTestAutomation]::Read($fixture)
    $formatFrom = $body.IndexOf('Alpha',[StringComparison]::Ordinal)
    $fixture.Range($formatFrom,$formatFrom + ("Alpha $emoji beta gamma.").Length).Font.Italic = -1
    $path = Join-Path $OutputDirectory 'Word-qualification.docx'
    $fixture.SaveAs2([ref][object]$path,[ref][object]16)
    WaitSaved 'The named synthetic document completed its first ordinary Save As'
    $diskBefore = DocumentHash $path
    $legacyRule=@((Local 'state').rules | Where-Object {$_.id -eq 'snappier'})[0]
    Local 'save_rule' @{id=$legacyRule.id;label=$legacyRule.label;color=$legacyRule.color;action=$legacyRule.action;prompt=$legacyRule.prompt;accessMode='suggestions'} | Out-Null
    Mark "Alpha $emoji beta gamma." 'snappier'
    Mark 'Protected passage stays unchanged.' 'protect'
    $state = Local 'state'; $documentId = $state.document.id
    $connection = Local 'connection_status'; $startedConnection = -not $connection.running
    if ($startedConnection) { $connection = Local 'connection_local' }
    $address=$connection.localMcpUrl
    $initialized = Rpc 'initialize' @{protocolVersion='2025-03-26';capabilities=@{};clientInfo=@{name='Word qualification';version='1'}}
    Check ($initialized.protocolVersion -eq '2025-03-26') 'The real Word MCP endpoint initialized'
    $catalogue = Rpc 'tools/list' @{}
    Check (($catalogue.tools.name -join ',') -eq 'list_manuscripts,read_manuscript,suggest_text,highlight_text') 'Only the four native Review tools are exposed'
    $listed=Tool 'list_manuscripts'
    $listedFixture=@($listed.manuscripts | Where-Object { $_.manuscript_id -eq $documentId })
    Check ($listed.manuscripts.Count -eq $word.Documents.Count -and $listedFixture.Count -eq 1 -and $listedFixture[0].title -ceq $fixture.Name -and $listedFixture[0].active) 'Connected discovery lists every open document and the exact active synthetic title and identity'
    $before = [WordTestAutomation]::Read($fixture); $quote="Alpha $emoji beta gamma."; $replacement="Alpha $emoji clear beta gamma."
    $from=$before.IndexOf($quote,[StringComparison]::Ordinal)
    $request=@{manuscript_id=$documentId;revision=(Local 'state').document.revision;context_id='forged';from=$from;to=$from+$quote.Length;quote=$quote;replacement=$replacement;request_id='no_prior_read'}
    Tool 'suggest_text' $request -Denied
    Check ([WordTestAutomation]::Read($fixture) -eq $before) 'A proposal without complete context left Word text unchanged'
    $context = ReadContext
    Check ($context.text.Contains('Read-only header context.') -and $context.text.Contains('Read-only footnote context.')) 'Complete context included header and footnote stories'
    $storedMode=@((Local 'state').rules | Where-Object {$_.id -eq 'snappier'})[0].accessMode
    $request.revision=$context.revision; $request.context_id=$context.context_id; $request.request_id='first_direct_example'
    $pending = Tool 'suggest_text' $request
    Check ($pending.status -eq 'applied' -and $pending.applied -and -not $pending.requiresApproval -and $pending.reviewMode -eq 'direct' -and $fixture.Revisions.Count -eq 0) 'The installed prompt edit changed native Word wording immediately without tracked approval'
    Check (@((Local 'state').rules | Where-Object {$_.id -eq 'snappier' -and $_.accessMode -ceq $storedMode}).Count -eq 1) 'Effective direct editing preserved the stored local prompt rule mode'
    Check ($fixture.Comments.Count -eq 1 -and $fixture.Comments.Item(1).Author -eq 'wordplai' -and $fixture.Comments.Item(1).Range.Text.Trim().Length -gt 0 -and $fixture.Comments.Item(1).Range.Text.Trim().Length -le 240) 'One concise native wordplai rationale comment appeared with the direct edit in the same operation'
    Check ($fixture.ActiveWindow.View.ShowComments -and $fixture.ActiveWindow.View.ShowRevisionsAndComments) 'The installed direct edit immediately enabled native comments and review markup'
    Check ($pending.current_context.complete -and $pending.current_context.reviewMode -eq 'direct' -and $pending.current_context.manuscript_id -eq $documentId -and $pending.current_context.review.feedback.Count -eq 1) 'The first direct edit returned complete updated context for immediate continuation'
    CheckGreenProtected $replacement
    $proposedContext=ReadContext
    $expected=$before.Substring(0,$from)+$replacement+$before.Substring($from+$quote.Length)
    Check ($proposedContext.main_text -ceq $expected) 'Direct context exposes exact edited wording and unchanged surrounding text'
    $firstNative=[WordTestAutomation]::Read($fixture)
    $firstReplay=Tool 'suggest_text' $request
    Check ($firstReplay.status -eq 'applied' -and $firstReplay.replayed -and $fixture.Comments.Count -eq 1 -and $fixture.Revisions.Count -eq 0 -and [WordTestAutomation]::Read($fixture) -ceq $firstNative) 'Replaying the first installed direct edit did not duplicate native wording, revisions or rationale'
    $reviewerComment=$fixture.Comments.Add($fixture.Paragraphs.Item(2).Range, 'Keep gamma and explain the reason without inventing facts.')
    $feedbackContext=ReadContext
    Check ($feedbackContext.text.Contains('explain the reason without inventing facts')) 'Complete MCP context includes native comment feedback'
    Check ($feedbackContext.revision -ne $proposedContext.revision) 'Adding native feedback invalidated the earlier context revision'
    $staleFeedback=@{} + $request
    $staleFeedback.request_id='stale_feedback'
    $staleFeedback.revision=$proposedContext.revision
    $staleFeedback.context_id=$proposedContext.context_id
    $staleFeedback.quote=$replacement
    $staleFeedback.to=$from+$replacement.Length
    $staleFeedback.replacement="Alpha $emoji grounded beta gamma."
    $staleReply=Rpc 'tools/call' @{name='suggest_text';arguments=$staleFeedback}
    Check ($staleReply.isError -and $staleReply.structuredContent.error.code -eq 'manuscript_changed' -and [WordTestAutomation]::Read($fixture) -ceq $expected) 'Native feedback produced a precise stale-revision refusal without another edit'
    $reviewerComment.Delete()
    WaitSaved 'The installed idle timer automatically saved direct wording, green protection and feedback changes'
    Check ($fixture.Revisions.Count -eq 0 -and (DocumentHash $path) -ne $diskBefore) 'Direct wording and protection were persisted in the named document without pending revisions'
    Check ((Local 'state').proposals.Count -eq 0) 'Completed direct edits did not create a separate pending list'
    WaitFor { (NativeGreen $replacement) -and $fixture.Saved -and $word.BackgroundSavingStatus -eq 0 } 'Direct wording automatically became visibly green and saved without native acceptance'
    CheckGreenProtected $replacement
    Check ([WordTestAutomation]::Italic($fixture,$from,$from+$replacement.Length) -eq -1) 'Direct editing preserved italic formatting'
    $greenContext=ReadContext
    $greenFrom=$greenContext.main_text.IndexOf($replacement,[StringComparison]::Ordinal)
    Tool 'suggest_text' @{manuscript_id=$documentId;revision=$greenContext.revision;context_id=$greenContext.context_id;from=$greenFrom;to=$greenFrom+$replacement.Length;quote=$replacement;replacement='Changed accepted wording.';request_id='accepted_green_denied'} -Denied
    Check ([WordTestAutomation]::Read($fixture) -ceq $expected -and (NativeGreen $replacement)) 'Accepted green wording stayed unchanged and protected against a fresh AI edit'
    Local 'save_rule' @{id='direct';label='Direct wording';color='#AACCEE';action='prompt';prompt='Improve clarity.';accessMode='direct';documentId=$documentId} | Out-Null
    Mark 'Direct passage wording.' 'direct'
    $context = ReadContext; $before=[WordTestAutomation]::Read($fixture); $from=$before.IndexOf('Direct passage wording.',[StringComparison]::Ordinal)
    $request=@{manuscript_id=$documentId;revision=$context.revision;context_id=$context.context_id;from=$from;to=$from+23;quote='Direct passage wording.';replacement='Clear direct wording.';request_id='direct_example'}
    $directRequest=@{} + $request
    $direct = Tool 'suggest_text' $request
    Check ($direct.status -eq 'applied' -and -not $direct.requiresApproval -and $direct.applied -and $direct.reviewMode -eq 'direct' -and $fixture.Revisions.Count -eq 0) 'The locally direct-colored passage also commits immediately without tracked revisions'
    Check ($direct.current_context.complete -and $direct.current_context.review.feedback.Count -eq $fixture.Comments.Count) 'A committed direct edit returns complete current context including its native rationale comment'
    $after=$direct.current_context.text.Substring(0,$direct.current_context.main_body_length)
    Check ($after -ceq $before.Substring(0,$from)+'Clear direct wording.'+$before.Substring($from+23)) 'The native direct edit preserved all text outside the highlight'
    CheckGreenProtected 'Clear direct wording.'
    $commentCount=$fixture.Comments.Count
    $pendingNative=[WordTestAutomation]::Read($fixture)
    WaitSaved 'The native direct edit, green protection and rationale comment automatically saved'
    $replay = Tool 'suggest_text' $request
    Check ($replay.status -eq 'applied' -and $replay.replayed -and [WordTestAutomation]::Read($fixture) -ceq $pendingNative -and $fixture.Comments.Count -eq $commentCount -and $fixture.Revisions.Count -eq 0) 'Replaying a direct edit did not duplicate native wording, revisions or comments'
    Check ([bool]$fixture.Undo([ref][object]1)) 'Word accepted one native Undo for the direct edit, green protection and rationale'
    Check ([WordTestAutomation]::Read($fixture) -ceq $before -and $fixture.Comments.Count -eq $commentCount-1) 'One native Undo restored original wording and removed exactly its rationale comment'
    $directRequest.request_id='undo_stale_example'
    Tool 'suggest_text' $directRequest -Denied
    Check ([WordTestAutomation]::Read($fixture) -ceq $before) 'A stale request after Undo could not reapply the direct edit'
    Check ([bool]$fixture.Redo([ref][object]1)) 'Word accepted one native Redo for the direct edit, protection and rationale'
    Check ([WordTestAutomation]::Read($fixture) -ceq $pendingNative -and $fixture.Comments.Count -eq $commentCount -and $fixture.Revisions.Count -eq 0) 'One native Redo restored the same direct wording, green protection and rationale'
    Check ([WordTestAutomation]::Read($fixture) -ceq $after -and $fixture.Revisions.Count -eq 0) 'The direct-colored edit requires no local Word acceptance'
    WaitFor { (NativeGreen 'Clear direct wording.') -and $fixture.Saved -and $word.BackgroundSavingStatus -eq 0 } 'Direct wording became visibly green and automatically saved in the installed add-in'
    CheckGreenProtected 'Clear direct wording.'
    Mark 'Discard this unnecessary idea.' 'discard'
    $discardBefore=[WordTestAutomation]::Read($fixture)
    $discardContext=ReadContext
    $discardQuote='Discard this unnecessary idea.'
    $discardFrom=$discardContext.main_text.IndexOf($discardQuote,[StringComparison]::Ordinal)
    $discardRequest=@{manuscript_id=$documentId;revision=$discardContext.revision;context_id=$discardContext.context_id;from=$discardFrom;to=$discardFrom+$discardQuote.Length;quote=$discardQuote;replacement='';request_id='discard_reviewed_only'}
    $discardComments=$fixture.Comments.Count
    $discard=Tool 'suggest_text' $discardRequest
    Check ($discard.status -eq 'tracked' -and $discard.requiresApproval -and -not $discard.applied -and $fixture.Revisions.Count -eq 1 -and $fixture.Comments.Count -eq $discardComments+1) 'Discard remains one complete native deletion requiring Word review with one rationale'
    $fixture.Revisions.RejectAll()
    Check ([WordTestAutomation]::Read($fixture) -ceq $discardBefore -and $fixture.Revisions.Count -eq 0 -and $fixture.Comments.Count -eq $discardComments+1) 'Native Reject restores the complete discarded passage and retains its rationale without changing earlier direct edits'
    WaitSaved 'Discard rejection automatically refreshed and saved without a manual Save or Refresh'
    $after=$discardBefore
    $context=ReadContext; $request.request_id='outside';$request.revision=$context.revision;$request.context_id=$context.context_id;$request.from=0;$request.to=8;$request.quote='Unmarked';$request.replacement='Changed'
    Tool 'suggest_text' $request -Denied
    Check ([WordTestAutomation]::Read($fixture) -eq $after) 'Unmarked text stayed unchanged'
    $context=ReadContext; $request.revision=$context.revision; $request.context_id=$context.context_id
    $from=$after.IndexOf('Protected passage stays unchanged.',[StringComparison]::Ordinal)
    $request.request_id='protected';$request.from=$from;$request.quote='Protected passage stays unchanged.';$request.to=$from+$request.quote.Length;$request.replacement='Changed'
    Tool 'suggest_text' $request -Denied
    Check ([WordTestAutomation]::Read($fixture) -eq $after) 'Protected text and surrounding manuscript stayed unchanged'
    $context=ReadContext
    $greenFrom=$context.main_text.IndexOf('Clear direct wording.',[StringComparison]::Ordinal)
    Tool 'suggest_text' @{manuscript_id=$documentId;revision=$context.revision;context_id=$context.context_id;from=$greenFrom;to=$greenFrom+21;quote='Clear direct wording.';replacement='Changed direct wording.';request_id='direct_green_denied'} -Denied
    Check ([WordTestAutomation]::Read($fixture) -ceq $after -and (NativeGreen 'Clear direct wording.')) 'Direct green wording rejected a fresh edit without native text or color changes'
    Mark 'Revocation probe wording.' 'direct'
    WaitSaved 'The independent direct revocation probe was automatically saved'
    $context=ReadContext
    $probeState=Local 'state'
    $probeRule=@($probeState.rules | Where-Object {$_.id -eq 'direct' -and $_.action -eq 'prompt' -and $_.accessMode -eq 'direct'})
    Check ($probeRule.Count -eq 1) 'The independent revocation probe retains an explicit local direct-edit rule'
    $currentDirect=@($probeState.marks | Where-Object {$_.ruleId -eq 'direct' -and -not $_.stale -and $_.quote -ceq 'Revocation probe wording.'})
    Check ($currentDirect.Count -eq 1 -and $context.main_text.Substring($currentDirect[0].from,$currentDirect[0].to-$currentDirect[0].from) -ceq $currentDirect[0].quote) 'Revocation probe starts with a current exact authorized direct marker'
    $restrictiveRules=@($probeState.rules | Where-Object {$_.action -eq 'protect' -or $_.action -eq 'reject'} | ForEach-Object {$_.id})
    $overlappingProtection=@($probeState.marks | Where-Object {$restrictiveRules -contains $_.ruleId -and $_.from -lt $currentDirect[0].to -and $_.to -gt $currentDirect[0].from})
    Check ($overlappingProtection.Count -eq 0) 'The revocation probe has direct authority without overlapping protection'
    $revokedRequest=@{manuscript_id=$documentId;revision=$context.revision;context_id=$context.context_id;from=$currentDirect[0].from;to=$currentDirect[0].to;quote=$currentDirect[0].quote;replacement='Another clear wording.';request_id='revoked_witness'}
    Local 'revoke_context' @{documentId=$documentId} | Out-Null
    Tool 'suggest_text' $revokedRequest -Denied
    Check ([WordTestAutomation]::Read($fixture) -eq $after) 'Revocation invalidated the old edit witness without changing text'
    $listed=Tool 'list_manuscripts'
    Check ($listed.manuscripts.Count -eq $word.Documents.Count -and @($listed.manuscripts|Where-Object {$_.manuscript_id -eq $documentId}).Count -eq 1 -and -not (Local 'state').document.shared) 'Listing retains the chosen identity without restoring a revoked read witness'
    WaitSaved 'All authorized native changes were automatically saved before closing the synthetic document'
    Check ((DocumentHash $path) -ne $diskBefore) 'The on-disk DOCX contains the automatically persisted editing session'
    if ($startedConnection) { Local 'connection_stop' | Out-Null }
    $fixture.Close(0);$fixture=$null
    $fixture=$word.Documents.Open($path)
    $reopened=Local 'state'
    Check ([WordTestAutomation]::Read($fixture) -ceq $after) 'Reopening proved the autosaved manuscript retained exactly the authorized final wording'
    CheckGreenProtected $replacement
    CheckGreenProtected 'Clear direct wording.'
    Check (-not $reopened.document.shared) 'Reopening a saved document did not restore sharing'
    Check ($reopened.rules.Count -ge 5 -and $reopened.marks.Count -ge 3) 'Rules and passage markers survived DOCX save and reopen'
    $hiddenBefore=$fixture.Bookmarks.ShowHidden
    [void]$fixture.Bookmarks.Add('_WraiHiddenProbe',$fixture.Range(0,8))
    $fixture.Range(0,8).Select()
    $denial=''
    try { Local 'selection_preview' | Out-Null } catch { $denial=$_.Exception.Message }
    Check ($denial -like '*bookmark or cross-reference target*') 'A hidden native bookmark prevented passage marking'
    Check ($fixture.Bookmarks.ShowHidden -eq $hiddenBefore) 'Bookmark validation preserved the native visibility setting'
    $fixture.Bookmarks.Item('_WraiHiddenProbe').Delete()
    WaitSaved 'Removing the disposable hidden bookmark was automatically saved'
    [ordered]@{status='PASS';checks=$checks;document=$path;wordVersion=$word.Version;wordBuild=$word.Build;publicTunnel='NOT RUN'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'result.json') -Encoding utf8
    Write-Output "PASS: $($checks.Count) installed Word checks. Synthetic document left open: $path"
}
catch {
    [ordered]@{status='FAIL';checks=$checks;error=$_.Exception.Message;publicTunnel='NOT RUN'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'result.json') -Encoding utf8
    throw
}
finally {
    if ($startedConnection -and $api) {
        try { if ((Local 'connection_status').running) { Local 'connection_stop' | Out-Null } } catch { }
    }
    $automation.Dispose()
}


