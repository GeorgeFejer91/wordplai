using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using UnderTree.WordAddin;

internal static partial class NativeSuggestionWorkflowSmoke
{
    private static readonly object traceGate = new object();
    private static readonly Stopwatch traceClock = Stopwatch.StartNew();
    private static ComAddin nativeDispatcher;

    private static void ConfigureNativeDispatcher(int nativeThread)
    {
        nativeDispatcher = new ComAddin();
        typeof(ComAddin).GetField("dispatcher", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(nativeDispatcher, dispatcher);
        typeof(ComAddin).GetField("ownerThread", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(nativeDispatcher, nativeThread);
    }

    private static void TraceNative(string message)
    {
        lock (traceGate) File.AppendAllText(Path.Combine(output, "native-dispatch.log"),
            traceClock.ElapsedMilliseconds.ToString() + "ms thread=" + Thread.CurrentThread.ManagedThreadId + " " + message + Environment.NewLine);
    }

    private static object MarshalNative(Func<object> action)
    {
        TraceNative("marshal-queued");
        Func<object> recorded = delegate
        {
            TraceNative("native-start");
            try { return action(); }
            finally { TraceNative("native-end"); }
        };
        try { return typeof(ComAddin).GetMethod("MarshalToWord", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(nativeDispatcher, new object[] { recorded }); }
        catch (TargetInvocationException error) { throw error.InnerException; }
    }
    private static Dictionary<string, object> NativeTool(NativeMcpClient client, string name, object args)
    {
        var result = client.Call(name, args);
        if ((bool)result["isError"]) throw new Exception(Json.Serialize(result));
        return Map(result["structuredContent"]);
    }

    private static Dictionary<string, object> NativeCoordination(Dictionary<string, object> result)
    { return Map(Map(result["_meta"])["undertree/coordination"]); }

    private static McpEndpoint NativeEndpoint()
    {
        var connection = (ConnectionService)typeof(WordEditor).GetField("connections", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
        return (McpEndpoint)typeof(ConnectionService).GetField("endpoint", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(connection);
    }

    private static McpCoordination NativeCoordinator()
    { return (McpCoordination)typeof(McpEndpoint).GetField("coordination", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(NativeEndpoint()); }

    private static string OwnedNativeRequestId(NativeMcpClient client, string requestId)
    {
        if (client.Conversation == null) return client.NativeRequestId(requestId);
        McpCoordination coordinator = NativeCoordinator();
        return PolicyEngine.Hash(coordinator.Resolve(coordinator.Require(client.Session), client.Conversation).NativeScope + "\0" + requestId);
    }

    private static int NativeQueuedTurns()
    {
        McpCoordination coordinator = NativeCoordinator();
        object gate = typeof(McpCoordination).GetField("gate", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(coordinator);
        lock (gate) return ((Queue<long>)typeof(McpCoordination).GetField("turns", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(coordinator)).Count;
    }

    private static void WaitWithoutNativeDispatch(Func<bool> ready)
    {
        Stopwatch deadline = Stopwatch.StartNew();
        while (!ready() && deadline.ElapsedMilliseconds < 5000) Thread.Sleep(10);
        if (!ready()) throw new TimeoutException("The concurrent HTTP requests did not reach their native queue barrier.");
    }

    private static void ConcurrentMcpClients()
    {
        const string second = "A second rough paragraph repeats a different point.";
        const string secondRevised = "A second paragraph states its point clearly.";
        const string revocationSource = "A remaining paragraph is reserved for revocation checks.";
        const string protectedText = "Protected native wording stays exact.";
        object secondDocument = null, largeDocument = null;
        try
        {
            Local("connection_stop"); Local("connection_local");
            NewFixture("concurrent-clients", Original + "\r" + second + "\r" + revocationSource + "\r" + protectedText);
            object firstDocument = document; string firstId = documentId;
            Mark(Original, "snappier"); Mark(second, "snappier"); Mark(revocationSource, "snappier"); Mark(protectedText, "protect");
            var initialMarks = Records(Local("state")["marks"]);
            string protectedId = Text(initialMarks.Single(mark => Text(mark, "ruleId") == "protect"), "id");
            string revocationId = Text(initialMarks.Single(mark => Text(mark, "ruleId") == "snappier" && Text(mark, "quote") == revocationSource), "id");
            string protectedBefore = WordSnapshot.FormattingSignature(document.Bookmarks.Item(protectedId).Range);
            string beforeBody = WordSnapshot.Capture(document).MainText;
            string beforeTail = beforeBody.Substring(beforeBody.IndexOf(protectedText, StringComparison.Ordinal) + protectedText.Length);
            var other = AddChoiceFixture("concurrent-other.docx", "Another document has a rough paragraph.");
            secondDocument = other.Document;
            document = firstDocument; documentId = firstId;
            string address = Text(Local("connection_status"), "localMcpUrl");
            var a = new NativeMcpClient(address, "ChatGPT", "native-conversation-a");
            var b = new NativeMcpClient(address, "Claude", "native-conversation-b");
            var c = new NativeMcpClient(address, "ChatGPT", "native-conversation-c");
            NativeMcpClient originalA = a, originalB = b, originalC = c;
            var contextA = NativeTool(a, "read_manuscript", Values("manuscript_id", firstId));
            var contextB = NativeTool(b, "read_manuscript", Values("manuscript_id", firstId));
            NativeTool(c, "read_manuscript", Values("manuscript_id", other.Id));
            Check(a.Session != b.Session && a.Session != c.Session && Text(contextA, "revision") == Text(contextB, "revision") &&
                Text(contextA, "context_id") != Text(contextB, "context_id"), "ChatGPT, Claude and a second GPT initialize independent real native read sessions at one endpoint.");
            a = new NativeMcpClient(address, "ChatGPT", originalA.Conversation);
            b = new NativeMcpClient(address, "Claude", originalB.Conversation);
            c = new NativeMcpClient(address, "ChatGPT", originalC.Conversation);
            Check(a.Session != originalA.Session && b.Session != originalB.Session && c.Session != originalC.Session,
                "All three labelled native conversations continue through freshly initialized physical transports.");
            Check(Text(NativeTool(a, "read_manuscript", Values()), "manuscript_id") == firstId &&
                Text(NativeTool(c, "read_manuscript", Values()), "manuscript_id") == other.Id,
                "Independent client document choices survive a different active native window.");

            largeDocument = application.Documents.Add(Visible: false); ((dynamic)largeDocument).Content.Text = new string('x', 34000) + "\r";
            ((dynamic)largeDocument).Activate();
            document = largeDocument;
            string largeId = Text(Map(Local("state")["document"]), "id"); documentIds.Add(largeId);
            var pageA = NativeTool(a, "read_manuscript", Values("manuscript_id", largeId));
            var pageB = NativeTool(b, "read_manuscript", Values("manuscript_id", largeId));
            Check(!(bool)pageA["complete"] && !(bool)pageB["complete"] && pageA["context_id"] == null && pageB["context_id"] == null,
                "Interleaved first pages grant neither client a complete-read witness.");
            var lastA = NativeTool(a, "read_manuscript", Values("manuscript_id", largeId, "offset", pageA["to"]));
            var lastB = NativeTool(b, "read_manuscript", Values("manuscript_id", largeId, "offset", pageB["to"]));
            Check((bool)lastA["complete"] && (bool)lastB["complete"] && Text(lastA, "context_id") != Text(lastB, "context_id"),
                "Two interleaved native pagination sequences complete independently without resetting each other's offsets.");
            ((dynamic)largeDocument).Close(0); largeDocument = null;
            document = firstDocument; documentId = firstId;
            contextA = NativeTool(a, "read_manuscript", Values("manuscript_id", firstId));
            contextB = NativeTool(b, "read_manuscript", Values("manuscript_id", firstId));
            string unchanged = new WordNativeEditState().NativeIdentity(document);
            var omittedSelection = a.Call("suggest_text", ContextRequest("omitted-native-owner", contextA, Original, Revised), true);
            Check((bool)omittedSelection["isError"] && Text(Map(Map(omittedSelection["structuredContent"])["error"]), "code") == "document_selection_required" &&
                new WordNativeEditState().NativeIdentity(document) == unchanged,
                "Omitting provider metadata on the same physical session cannot inherit a labelled native document choice.");
            var defaultRead = a.Call("read_manuscript", Values("manuscript_id", firstId), true);
            var defaultTheft = a.Call("suggest_text", ContextRequest("omitted-native-proof", contextA, Original, Revised), true);
            Check(!(bool)defaultRead["isError"] && (bool)defaultTheft["isError"] &&
                Text(Map(Map(defaultTheft["structuredContent"])["error"]), "code") == "context_changed" &&
                new WordNativeEditState().NativeIdentity(document) == unchanged,
                "A separately read unlabelled native scope cannot steal a labelled complete-read witness on that same transport.");
            var forged = b.Call("suggest_text", ContextRequest("cross-client-witness", contextA, Original, Revised));
            Check((bool)forged["isError"] && Text(Map(Map(forged["structuredContent"])["error"]), "code") == "context_changed" &&
                new WordNativeEditState().NativeIdentity(document) == unchanged,
                "A Claude session cannot use another client's complete-read witness, even for its own selected document.");
            contextB = Map(Map(Map(forged["structuredContent"])["error"])["details"]);
            contextB = Map(contextB["current_context"]);
            var firstRequest = ContextRequest("shared-request-name", contextA, Original, Revised);
            var staleRequest = ContextRequest("claude-contender", contextB, second, secondRevised);
            Task<Dictionary<string, object>> first = a.CallAsync("suggest_text", firstRequest);
            WaitWithoutNativeDispatch(() => NativeQueuedTurns() == 1);
            TraceNative("first-turn-queued");
            Task<Dictionary<string, object>> contender = b.CallAsync("suggest_text", staleRequest);
            WaitWithoutNativeDispatch(() => NativeQueuedTurns() == 2);
            TraceNative("both-turns-queued");
            var winner = NativeMcpClient.Wait(first); var loser = NativeMcpClient.Wait(contender);
            Check(!(bool)winner["isError"] && Text(Map(winner["structuredContent"]), "status") == "applied" &&
                (bool)loser["isError"] && (int)((dynamic)document).Comments.Count == 1 && (int)((dynamic)document).Revisions.Count == 0 &&
                WordSnapshot.Capture(document).MainText.Contains(Revised + "\r" + second),
                "Two actually queued native HTTP edits from one revision produce one direct winner and one unchanged losing passage.");
            CheckGreen(Revised);
            var diagnostic = Map(Map(loser["structuredContent"])["error"]);
            Check(Text(diagnostic, "code") == "manuscript_changed" &&
                UInt64.Parse(Text(NativeCoordination(winner), "ticket")) < UInt64.Parse(Text(NativeCoordination(loser), "ticket")) &&
                Text(NativeCoordination(winner), "state") == "applied" && Text(NativeCoordination(loser), "state") == "rejected",
                "The native commit and stale refusal expose distinct increasing FIFO tickets and exact outcomes.");
            string afterFirst = new WordNativeEditState().NativeIdentity(document);
            var replay = a.Call("suggest_text", firstRequest);
            Check(!(bool)replay["isError"] && Text(NativeCoordination(replay), "state") == "replayed" &&
                new WordNativeEditState().NativeIdentity(document) == afterFirst && (int)((dynamic)document).Comments.Count == 1,
                "An exact network retry returns its scoped durable receipt without another native edit or rationale.");
            var changed = new Dictionary<string, object>(firstRequest); changed["replacement"] = "Changed network retry.";
            Check((bool)a.Call("suggest_text", changed)["isError"] && new WordNativeEditState().NativeIdentity(document) == afterFirst,
                "Changed arguments cannot reuse an already applied native receipt.");
            var changedContext = new Dictionary<string, object>(firstRequest); changedContext["context_id"] = "changed-original-context-only";
            Check((bool)a.Call("suggest_text", changedContext)["isError"] && new WordNativeEditState().NativeIdentity(document) == afterFirst && (int)((dynamic)document).Comments.Count == 1,
                "Changing only the original context identity cannot replay a committed native request or create another rationale.");
            var refreshedB = Map(Map(diagnostic["details"])["current_context"]);
            var resumed = NativeTool(b, "suggest_text", ContextRequest("claude-refreshed", refreshedB, second, secondRevised));
            Check(Text(resumed, "status") == "applied" && (bool)resumed["applied"] && !(bool)resumed["requiresApproval"] &&
                (int)((dynamic)document).Comments.Count == 2 && (int)((dynamic)document).Revisions.Count == 0 &&
                WordSnapshot.Capture(document).MainText.Contains(Revised + "\r" + secondRevised),
                "Claude resumes with a fresh computed proposal and new ID using its returned current context.");
            CheckGreen(secondRevised);
            string firstAfterTwo = new WordNativeEditState().NativeIdentity(document);
            var otherContext = NativeTool(c, "read_manuscript", Values());
            NativeTool(c, "suggest_text", ContextRequest("shared-request-name", otherContext, other.Source, "Another document states its point clearly."));
            Check(new WordNativeEditState().NativeIdentity(document) == firstAfterTwo && (int)((dynamic)secondDocument).Comments.Count == 1,
                "A second GPT can use the same request name independently on its chosen native document.");
            Check((bool)((dynamic)document).Bookmarks.Exists(protectedId), "Concurrent edits preserve the original protected native bookmark identity.");
            Check(WordSnapshot.FormattingSignature(((dynamic)document).Bookmarks.Item(protectedId).Range) == protectedBefore &&
                (string)((dynamic)document).Bookmarks.Item(protectedId).Range.Text == protectedText,
                "Concurrent clients preserve the protected native paragraph's exact wording and formatting.");
            string afterBody = WordSnapshot.Capture(document).MainText;
            Check(afterBody.Substring(afterBody.IndexOf(protectedText, StringComparison.Ordinal) + protectedText.Length) == beforeTail,
                "Every unmarked tail character and native terminal paragraph mark matches its captured pre-edit baseline.");

            var currentB = NativeTool(b, "read_manuscript", Values());
            var queuedRequest = ContextRequest("delete-before-native-start", currentB, revocationSource, "This queued edit must not appear.");
            var restrictiveRules = Records(currentB["rules"]).Where(rule => Text(rule, "action") == "protect" || Text(rule, "action") == "reject")
                .Select(rule => Text(rule, "id")).ToArray();
            var promptRules = Records(currentB["rules"]).Where(rule => Text(rule, "action") == "prompt").Select(rule => Text(rule, "id")).ToArray();
            Check(Records(currentB["highlights"]).Count(mark => !(bool)mark["stale"] &&
                Text(mark, "id") == revocationId && promptRules.Contains(Text(mark, "ruleId")) &&
                Convert.ToInt32(mark["from"]) == Convert.ToInt32(queuedRequest["from"]) && Convert.ToInt32(mark["to"]) == Convert.ToInt32(queuedRequest["to"])) == 1 &&
                !Records(currentB["highlights"]).Any(mark => restrictiveRules.Contains(Text(mark, "ruleId")) &&
                    Convert.ToInt32(mark["from"]) < Convert.ToInt32(queuedRequest["to"]) && Convert.ToInt32(queuedRequest["from"]) < Convert.ToInt32(mark["to"])),
                "The queued revocation probe uses one exact current prompt marker with no overlapping protected passage.");
            var noEffect = b.CallAsync("suggest_text", queuedRequest);
            WaitWithoutNativeDispatch(() => NativeQueuedTurns() == 1);
            Task<int> deletion = Task.Run(() => originalB.Delete());
            WaitWithoutNativeDispatch(delegate {
                try { NativeCoordinator().Require(originalB.Session); return false; }
                catch (McpHttpException error) { return error.Status == 404; }
            });
            bool revoked = false;
            try { NativeMcpClient.Wait(noEffect); } catch (Exception error) { revoked = error.Message.Contains("revoked"); }
            Check(revoked && NativeMcpClient.Wait(deletion) == 200 && new WordNativeEditState().NativeIdentity(document) == firstAfterTwo,
                "Deleting one MCP session cancels its already queued native edit without changing wording or metadata.");
            Check(Text(NativeTool(a, "read_manuscript", Values()), "manuscript_id") == firstId &&
                Text(NativeTool(c, "read_manuscript", Values()), "manuscript_id") == other.Id,
                "The surviving ChatGPT sessions retain their own document targets after Claude disconnects.");
            bool deletedDenied = false;
            try { originalB.Call("read_manuscript", Values()); }
            catch (WebException error) { using (var response = (HttpWebResponse)error.Response) deletedDenied = response != null && response.StatusCode == HttpStatusCode.NotFound; }
            Check(deletedDenied, "A retired client cannot resurrect its old native choice or witness by reusing its header.");
            var retiredProof = b.Call("suggest_text", ContextRequest("revoked-logical-generation", currentB, revocationSource, "This old logical proof must not apply."));
            Check((bool)retiredProof["isError"] && Text(Map(Map(retiredProof["structuredContent"])["error"]), "code") == "document_selection_required" &&
                new WordNativeEditState().NativeIdentity(document) == firstAfterTwo,
                "A still-live transport cannot inherit a logical selection or proof revoked through its earlier physical session.");
        }
        catch
        {
            if (document != null)
                try { File.WriteAllText(Path.Combine(output, "concurrency-failure-body.xml"), (string)((dynamic)document).Content.WordOpenXML); }
                catch (Exception diagnostic) { TraceNative("failure-capture " + diagnostic.GetType().Name); }
            throw;
        }
        finally
        {
            if (largeDocument != null) ((dynamic)largeDocument).Close(0);
            if (secondDocument != null) ((dynamic)secondDocument).Close(0);
            if (document != null) { ((dynamic)document).Close(0); document = null; }
            Local("connection_stop"); Local("connection_local");
        }
    }

    private static WordDocumentReview NativeDocumentState()
    {
        var reviews = (Dictionary<long, WordDocumentReview>)typeof(WordEditor).GetField("documents", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
        return reviews[WordSnapshot.Identity(document)];
    }

    private static void NativeFinalizerRecovery()
    {
        int openBefore = (int)application.Documents.Count;
        try
        {
        NewFixture("lost-native-finalizer", Original); Mark(Original, "snappier");
        string address = Text(Local("connection_status"), "localMcpUrl");
        const string conversation = "native-finalizer-recovery";
        var client = new NativeMcpClient(address, "finalizer-recovery", conversation);
        var context = NativeTool(client, "read_manuscript", Values("manuscript_id", documentId));
        var request = ContextRequest("committed-before-finalizer", context, Original, Revised);
        client = new NativeMcpClient(address, "finalizer-recovery", conversation);
        Action failFinalizer = delegate { throw new IOException("Synthetic response finalizer failure after native commit."); };
        Dictionary<string, object> failed;
        editor.ReviewChanged += failFinalizer;
        try { failed = client.Call("suggest_text", request); }
        finally { editor.ReviewChanged -= failFinalizer; }
        WordDocumentReview state = NativeDocumentState();
        string nativeId = OwnedNativeRequestId(client, "committed-before-finalizer");
        Check((bool)failed["isError"] && Text(NativeCoordination(failed), "state") == "outcome_unknown" &&
            state.Policy.Proposals.Single(receipt => receipt.Id == nativeId).Status == "applied" &&
            WordReviewStore.Load((object)document).Policy.Proposals.Single(receipt => receipt.Id == nativeId).Status == "applied" &&
            WordSnapshot.Capture(document).MainText.Contains(Revised) && (int)document.Comments.Count == 1 && (int)document.Revisions.Count == 0,
            "A throwing Review finalizer reports unknown delivery while retaining the confirmed native direct receipt and rationale.");
        CheckGreen(Revised);
        string afterSuggestion = new WordNativeEditState().NativeIdentity(document);
        string receiptPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wrAIterWord", "reviews", documentId + ".bin");
        string afterSuggestionReceipt = FileHash(receiptPath);
        int afterSuggestionComments = (int)document.Comments.Count, afterSuggestionRevisions = (int)document.Revisions.Count;
        client = new NativeMcpClient(address, "finalizer-recovery", conversation);
        Dictionary<string, object> failedReplay;
        editor.ReviewChanged += failFinalizer;
        try { failedReplay = client.Call("suggest_text", request); }
        finally { editor.ReviewChanged -= failFinalizer; }
        string afterFailedReplay = new WordNativeEditState().NativeIdentity(document), afterFailedReplayReceipt = FileHash(receiptPath);
        File.WriteAllText(Path.Combine(output, "finalizer-replay-default-evidence.json"), Json.Serialize(new { response = failedReplay,
            nativeBefore = afterSuggestion, nativeAfter = afterFailedReplay, receiptBefore = afterSuggestionReceipt, receiptAfter = afterFailedReplayReceipt,
            commentsBefore = afterSuggestionComments, commentsAfter = (int)document.Comments.Count, revisionsBefore = afterSuggestionRevisions, revisionsAfter = (int)document.Revisions.Count }));
        var replayDetails = Map(Map(Map(failedReplay["structuredContent"])["error"])["details"]);
        Check((bool)failedReplay["isError"] && Text(NativeCoordination(failedReplay), "state") == "outcome_unknown" &&
            replayDetails.Count == 2 && replayDetails.ContainsKey("operation_started") && replayDetails.ContainsKey("outcome_unknown") &&
            (bool)replayDetails["operation_started"] && (bool)replayDetails["outcome_unknown"] &&
            afterFailedReplay == afterSuggestion && afterFailedReplayReceipt == afterSuggestionReceipt &&
            (int)document.Comments.Count == afterSuggestionComments && (int)document.Revisions.Count == afterSuggestionRevisions,
            "A callback failure on exact applied replay retains legacy two-field diagnostics and leaves native state and durable receipt unchanged.");
        client = new NativeMcpClient(address, "finalizer-recovery", conversation);
        var replay = client.Call("suggest_text", request);
        Check(!(bool)replay["isError"] && Text(NativeCoordination(replay), "state") == "replayed" &&
            new WordNativeEditState().NativeIdentity(document) == afterSuggestion && (int)document.Comments.Count == 1,
            "An exact retry on a fresh transport after native finalizer failure retrieves the applied receipt without another edit or rationale.");

        context = NativeTool(client, "read_manuscript", Values("manuscript_id", documentId));
        int from = Text(context, "text").IndexOf(Outside, StringComparison.Ordinal);
        var marking = Values("documentId", documentId, "marking", Values("requestId", "marked-before-finalizer", "contextId", context["context_id"],
            "revision", context["revision"], "color", "#BACDEF", "ranges", new object[] { Values("from", from, "to", from + Outside.Length, "quote", Outside) }));
        client = new NativeMcpClient(address, "finalizer-recovery", conversation);
        editor.ReviewChanged += failFinalizer;
        try { failed = client.Call("highlight_text", marking); }
        finally { editor.ReviewChanged -= failFinalizer; }
        string markId = OwnedNativeRequestId(client, "marked-before-finalizer");
        Check((bool)failed["isError"] && Text(NativeCoordination(failed), "state") == "outcome_unknown" &&
            state.Policy.HighlightReceipts.Single(receipt => receipt.Request.RequestId == markId).Outcome == "applied" &&
            WordReviewStore.Load((object)document).Policy.HighlightReceipts.Single(receipt => receipt.Request.RequestId == markId).Outcome == "applied",
            "A throwing marking finalizer retains the confirmed applied receipt instead of downgrading native success.");
        File.WriteAllText(Path.Combine(output, "finalizer-highlight-default-evidence.json"), Json.Serialize(failed));
        var highlightDetails = Map(Map(Map(failed["structuredContent"])["error"])["details"]);
        Check(highlightDetails.Count == 2 && highlightDetails.ContainsKey("operation_started") && highlightDetails.ContainsKey("outcome_unknown") &&
            (bool)highlightDetails["operation_started"] && (bool)highlightDetails["outcome_unknown"],
            "A committed highlighting callback failure preserves the exact legacy two-field diagnostic details.");
        string afterMarking = new WordNativeEditState().NativeIdentity(document);
        client = new NativeMcpClient(address, "finalizer-recovery", conversation);
        replay = client.Call("highlight_text", marking);
        Check(!(bool)replay["isError"] && Text(NativeCoordination(replay), "state") == "replayed" &&
            new WordNativeEditState().NativeIdentity(document) == afterMarking,
            "An exact marking retry on a fresh transport after lost finalization preserves the native mark and adds no second effect.");
        }
        finally
        {
            if (document != null) { document.Close(0); document = null; }
        }
        Check((int)application.Documents.Count == openBefore, "Finalizer recovery closes its owned document and restores the captured native document count.");
    }
}
