using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using UnderTree.WordAddin;

internal static partial class TransportTests
{
    private static string NewSession(string address, string name)
    {
        int status; string issued;
        string reply = PostSession(address, "initialize", new { protocolVersion = "2025-06-18", clientInfo = new { name = name, version = "test" } }, out status, null, out issued);
        Assert(status == 200 && issued != null && issued.Length == 64, "Initialization returns a distinct standard MCP session header.");
        var contract = (IDictionary<string, object>)((IDictionary<string, object>)Result(reply)["_meta"])["undertree/concurrency"];
        Assert(Object.Equals(contract["protocol"], McpCoordination.Protocol) && Object.Equals(contract["sessionRequired"], true) &&
            Object.Equals(contract["maxSessions"], McpCoordination.MaxSessions) && Object.Equals(contract["maxConversations"], McpCoordination.MaxConversations) &&
            Object.Equals(contract["ticketPolicy"], "fifo"), "Initialization advertises the shared FIFO concurrency contract.");
        return issued;
    }

    private static IDictionary<string, object> SessionCall(string address, string session, string name, object args)
    {
        int status; string issued;
        string result = PostSession(address, "tools/call", new { name = name, arguments = args }, out status, session, out issued);
        Assert(status == 200 && issued == null, "Post-initialization calls retain their supplied session.");
        return Result(result);
    }

    private static IDictionary<string, object> Coordination(IDictionary<string, object> result)
    { return (IDictionary<string, object>)((IDictionary<string, object>)result["_meta"])["undertree/coordination"]; }

    private static int DeleteSession(string address, string session)
    {
        var request = (HttpWebRequest)WebRequest.Create(address);
        request.Proxy = null; request.Method = "DELETE"; request.Timeout = 10000; request.KeepAlive = false;
        if (session != null) request.Headers["Mcp-Session-Id"] = session;
        try { using (var reply = (HttpWebResponse)request.GetResponse()) return (int)reply.StatusCode; }
        catch (WebException error) { using (var reply = (HttpWebResponse)error.Response) { if (reply == null) throw; return (int)reply.StatusCode; } }
    }

    private static int PendingTurns(McpEndpoint endpoint)
    {
        var coordinator = (McpCoordination)typeof(McpEndpoint).GetField("coordination", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(endpoint);
        object gate = typeof(McpCoordination).GetField("gate", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(coordinator);
        lock (gate) return ((Queue<long>)typeof(McpCoordination).GetField("turns", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(coordinator)).Count;
    }

    private static void ConcurrentClients(string directory, string fake)
    {
        using (var entered = new ManualResetEvent(false))
        using (var release = new ManualResetEvent(false))
        {
            int changes = 0;
            bool hold = false;
            var receipts = new Dictionary<string, string>();
            var retired = new List<string>();
            using (var endpoint = new McpEndpoint(delegate(string name, IDictionary<string, object> values)
            {
                string scope = (string)values[WordDocumentPicker.SessionKey];
                if (name == "read_manuscript") return new { scope = scope, revision = changes.ToString() };
                string key = scope + ":" + (string)values["request_id"];
                string wording = (string)values["replacement"];
                if (receipts.ContainsKey(key))
                {
                    if (receipts[key] != wording) throw new McpOperationException("Changed replay denied.");
                    return new { replayed = true, revision = changes.ToString() };
                }
                if (hold) { entered.Set(); if (!release.WaitOne(5000)) throw new Exception("Concurrent native barrier was not released."); }
                if ((string)values["revision"] != changes.ToString()) throw new McpOperationException("Refresh current context.", "manuscript_changed", new { request_applied = false });
                changes++; receipts.Add(key, wording);
                return new { replayed = false, revision = changes.ToString() };
            }, Marshal, "concurrency-fixture", prefix => retired.Add(prefix)))
            {
                string address = endpoint.LocalAddress;
                int status; string issued;
                PostSession(address, "ping", new { }, out status, null, out issued);
                Assert(status == 400 && changes == 0, "An uninitialized request cannot inherit a shared default client.");
                PostSession(address, "ping", new { }, out status, "unknown", out issued);
                Assert(status == 404 && changes == 0, "Unknown sessions fail closed before native dispatch.");
                for (int count = 0; count < 70; count++)
                {
                    PostSession(address, "initialize", "invalid-parameters", out status, null, out issued);
                    Assert(status == 400 && issued == null, "Failed initialization grants no session header or live capacity.");
                }
                string a = NewSession(address, "ChatGPT"), b = NewSession(address, "Claude"), c = NewSession(address, "ChatGPT");
                var notAdmitted = SessionCall(address, a, "read_manuscript", new { forbidden_policy = true });
                Assert(Object.Equals(notAdmitted["isError"], true) && !((IDictionary<string, object>)notAdmitted["_meta"]).ContainsKey("undertree/coordination"),
                    "A pre-admission validation refusal carries no fabricated coordination ticket.");
                var readA = (IDictionary<string, object>)SessionCall(address, a, "read_manuscript", new { })["structuredContent"];
                var readB = (IDictionary<string, object>)SessionCall(address, b, "read_manuscript", new { })["structuredContent"];
                var readC = (IDictionary<string, object>)SessionCall(address, c, "read_manuscript", new { })["structuredContent"];
                Assert(a != b && b != c && a != c && !Object.Equals(readA["scope"], readB["scope"]) && !Object.Equals(readA["scope"], readC["scope"]),
                    "ChatGPT, Claude and a second GPT with the same client name have isolated native identities.");
                hold = true;
                object first = new { manuscript_id = "synthetic", request_id = "same-id", revision = "0", replacement = "First wording." };
                object second = new { manuscript_id = "synthetic", request_id = "same-id", revision = "0", replacement = "Second wording." };
                Task<IDictionary<string, object>> firstCall = Task.Run(() => SessionCall(address, a, "suggest_text", first));
                Assert(entered.WaitOne(5000), "The first HTTP client holds its native turn before the second arrives.");
                Task<IDictionary<string, object>> secondCall = Task.Run(() => SessionCall(address, b, "suggest_text", second));
                Wait(() => PendingTurns(endpoint) == 2, "second concurrent client queued");
                release.Set();
                Assert(Task.WaitAll(new Task[] { firstCall, secondCall }, 5000), "Both simultaneously admitted clients receive bounded results.");
                var winner = firstCall.Result; var loser = secondCall.Result;
                Assert(Object.Equals(winner["isError"], false) && Object.Equals(loser["isError"], true) && changes == 1 &&
                    Json.Serialize(loser).Contains("manuscript_changed"), "Only the first same-revision edit applies; the queued contender is explicitly stale.");
                Assert(UInt64.Parse((string)Coordination(winner)["ticket"]) < UInt64.Parse((string)Coordination(loser)["ticket"]) &&
                    Object.Equals(Coordination(winner)["state"], "applied") && Object.Equals(Coordination(loser)["state"], "rejected"),
                    "Monotonic server tickets show FIFO commit/refusal order without a client clock.");
                hold = false;
                var replay = SessionCall(address, a, "suggest_text", first);
                Assert(Object.Equals(Coordination(replay)["state"], "replayed") && changes == 1, "An exact same-session retry returns its receipt without applying again.");
                var changed = SessionCall(address, a, "suggest_text", new { manuscript_id = "synthetic", request_id = "same-id", revision = "0", replacement = "Altered replay." });
                Assert(Object.Equals(changed["isError"], true) && changes == 1, "Changed retry arguments cannot reuse the first client's receipt.");
                var independent = SessionCall(address, b, "suggest_text", new { manuscript_id = "synthetic", request_id = "same-id", revision = "1", replacement = "Second wording." });
                Assert(Object.Equals(independent["isError"], false) && changes == 2, "A second client can independently reuse a request name after refreshing its revision.");

                entered.Reset(); release.Reset(); hold = true;
                Task<IDictionary<string, object>> held = Task.Run(() => SessionCall(address, a, "suggest_text", new { manuscript_id = "synthetic", request_id = "held", revision = "2", replacement = "Held wording." }));
                Assert(entered.WaitOne(5000), "The first client owns the running turn during deletion of a different queued client.");
                Task<string> queued = Task.Run(delegate { int code; string sessionHeader; return PostSession(address, "tools/call", new { name = "suggest_text", arguments = new { manuscript_id = "synthetic", request_id = "revoked", revision = "2", replacement = "Never apply." } }, out code, c, out sessionHeader); });
                Wait(() => PendingTurns(endpoint) == 2, "revocable client queued");
                Task<int> deletion = Task.Run(() => DeleteSession(address, c));
                Assert(queued.Wait(5000) && queued.Result.Contains("revoked") && changes == 2, "Deleting a client cancels its queued request before any native change.");
                release.Set();
                Assert(held.Wait(5000) && deletion.Wait(5000) && deletion.Result == 200 && changes == 3 && retired.Count == 1,
                    "The running other-client edit finishes once and native revocation clears only the deleted client scope.");
                PostSession(address, "ping", new { }, out status, c, out issued);
                Assert(status == 404 && DeleteSession(address, null) == 400, "Deleted and missing session IDs cannot resume or revoke another client.");
                PostSession(address, "ping", new { }, out status, a, out issued);
                Assert(status == 200, "Deleting another session leaves the surviving client usable.");
                for (int count = 2; count < McpCoordination.MaxSessions; count++) NewSession(address, "capacity");
                PostSession(address, "initialize", new { protocolVersion = "2025-06-18" }, out status, null, out issued);
                Assert(status == 429 && issued == null, "Session overflow fails explicitly without evicting a live client.");
                PostSession(address, "ping", new { }, out status, a, out issued);
                Assert(status == 200, "A full session table retains its original client's identity.");
                endpoint.Dispose();
                Assert(!endpoint.Active, "Stopping the endpoint revokes every session and pending turn.");
            }
        }
    }

    private static void MonotonicSessionExpiry()
    {
        using (Stream stream = typeof(TransportTests).Assembly.GetManifestResourceStream("SelectiveEditorContract"))
        using (var reader = new StreamReader(stream))
        {
            var fixture = (IDictionary<string, object>)((IDictionary<string, object>)Json.DeserializeObject(reader.ReadToEnd()))["mcpConcurrency"];
            Assert(Object.Equals(fixture["protocol"], McpCoordination.Protocol) && Object.Equals(fixture["maxSessions"], McpCoordination.MaxSessions) &&
                Object.Equals(fixture["maxConversations"], McpCoordination.MaxConversations) &&
                Object.Equals(fixture["idleTimeoutSeconds"], McpCoordination.IdleTimeoutSeconds) && Object.Equals(fixture["maxQueuedEdits"], McpCoordination.MaxQueuedEdits),
                "Word concurrency constants match the three-host shared acceptance fixture.");
        }
        long now = 0;
        var coordinator = new McpCoordination(() => now, 100);
        McpClientSession original = coordinator.Create();
        now = 99;
        Assert(coordinator.Require(original.Id) == original, "A client remains usable before the monotonic idle deadline and refreshes activity.");
        now = 199;
        int status = 0;
        try { coordinator.Require(original.Id); } catch (McpHttpException error) { status = error.Status; }
        Assert(status == 404 && !coordinator.IsActive(original), "Monotonic idle expiry revokes the old session without a wall-clock comparison.");
        int retired = 0; coordinator.DrainRetired(prefix => retired++);
        McpClientSession fresh = coordinator.Create();
        Assert(retired == 1 && fresh.Id != original.Id && fresh.ScopePrefix != original.ScopePrefix, "Expired native context is retired and a fresh session cannot inherit it.");
        long[] tickets = Enumerable.Range(0, McpCoordination.MaxQueuedEdits).Select(index => coordinator.Enqueue()).ToArray();
        string code = null;
        try { coordinator.Enqueue(); } catch (McpOperationException error) { code = error.Code; }
        Assert(code == "word_busy" && tickets.Distinct().Count() == 8, "The edit queue rejects overflow while retaining eight distinct FIFO tickets.");
        coordinator.Stop();
        Assert(!coordinator.IsActive(fresh), "Stop invalidates the remaining session generation immediately.");
    }

    private static void UnknownRunningTurn()
    {
        using (var entered = new ManualResetEvent(false))
        using (var release = new ManualResetEvent(false))
        {
            int changes = 0, nativeCalls = 0, marshalCalls = 0;
            Func<Func<object>, object> uncertainMarshal = delegate(Func<object> callback)
            {
                if (Interlocked.Increment(ref marshalCalls) != 1) return Marshal(callback);
                Queue.Add(delegate { try { callback(); } catch (Exception) { } });
                if (!entered.WaitOne(5000)) throw new Exception("The uncertain native operation never started.");
                throw new McpOperationException("Word is still completing this operation. Inspect the result before retrying.", "word_busy",
                    new { operation_started = true, outcome_unknown = true });
            };
            using (var endpoint = new McpEndpoint(delegate(string name, IDictionary<string, object> values)
            {
                nativeCalls++;
                if ((string)values["request_id"] == "uncertain")
                {
                    if (changes == 1) return new { replayed = true };
                    entered.Set();
                    if (!release.WaitOne(5000)) throw new Exception("The uncertain running operation did not settle.");
                    changes++;
                    return new { replayed = false };
                }
                if (changes != 0) throw new McpOperationException("This revision is stale.", "manuscript_changed", new { request_applied = false });
                changes++;
                return new { replayed = false };
            }, uncertainMarshal, "unknown-turn-fixture", delegate { }))
            {
                try
                {
                    string address = endpoint.LocalAddress;
                    string a = NewSession(address, "ChatGPT"), b = NewSession(address, "Claude");
                    object first = new { manuscript_id = "synthetic", request_id = "uncertain", revision = "0" };
                    var uncertain = SessionCall(address, a, "suggest_text", first);
                    Assert(Object.Equals(uncertain["isError"], true) && Object.Equals(Coordination(uncertain)["state"], "outcome_unknown") &&
                        PendingTurns(endpoint) == 1 && changes == 0, "A running uncertain native result keeps its ticket until the actual callback completes.");
                    Task<IDictionary<string, object>> second = Task.Run(() => SessionCall(address, b, "suggest_text", new { manuscript_id = "synthetic", request_id = "second", revision = "0" }));
                    Wait(() => PendingTurns(endpoint) == 2, "second client waits behind uncertain native operation");
                    Assert(nativeCalls == 1 && changes == 0 && !second.IsCompleted, "No later client enters native authority while an uncertain turn is still running.");
                    var additional = Enumerable.Range(0, 6).Select(index => Task.Run(() => SessionCall(address, b, "suggest_text", new {
                        manuscript_id = "synthetic", request_id = "queued-" + index, revision = "0" }))).ToArray();
                    Wait(() => PendingTurns(endpoint) == 8, "full native queue behind uncertain operation");
                    var full = SessionCall(address, b, "suggest_text", new { manuscript_id = "synthetic", request_id = "overflow", revision = "0" });
                    Assert(Object.Equals(full["isError"], true) && Json.Serialize(full).Contains("word_busy") &&
                        !((IDictionary<string, object>)full["_meta"]).ContainsKey("undertree/coordination"), "A full native queue rejects without manufacturing an admission ticket.");
                    release.Set();
                    Assert(second.Wait(5000) && Object.Equals(second.Result["isError"], true) && changes == 1,
                        "The uncertain edit completes once; only then can a later client receive a stale-revision refusal.");
                    Assert(Task.WaitAll(additional.Cast<Task>().ToArray(), 5000) && additional.All(call => Object.Equals(call.Result["isError"], true)) && changes == 1,
                        "Every admitted queued contender remains stale after the uncertain native edit settles once.");
                    var replay = SessionCall(address, a, "suggest_text", first);
                    Assert(Object.Equals(Coordination(replay)["state"], "replayed") && changes == 1,
                        "An exact same-identity retry after native settlement cannot duplicate an uncertain edit.");
                }
                finally { release.Set(); }
            }
        }
    }

    private sealed class CyclicNativeReply { public object Result; }

    private static void NativeOutcomeFailures()
    {
        var legacy = (IDictionary<string, object>)Json.DeserializeObject(Json.Serialize(WordNativeOutcome.UnknownException().Details));
        Assert(legacy.Count == 2 && Object.Equals(legacy["operation_started"], true) && Object.Equals(legacy["outcome_unknown"], true),
            "Untouched unknown-outcome callers retain exactly their two legacy detail fields.");
        for (int faults = 0; faults < 8; faults++)
        {
            bool consumed = false, shared = true, persisted = false;
            int effects = 0;
            using (var endpoint = new McpEndpoint(delegate(string name, IDictionary<string, object> values)
            {
                effects++;
                WordNativeOutcome.Unknown(
                    delegate { consumed = true; if ((faults & 1) != 0) throw new IOException("synthetic recording failure"); },
                    delegate { shared = false; if ((faults & 2) != 0) throw new IOException("synthetic revoke failure"); },
                    delegate { if ((faults & 4) != 0) throw new IOException("synthetic receipt persistence failure"); persisted = true; });
                return null;
            }, Marshal, "uncertain-native-fixture", delegate { }))
            {
                string session = NewSession(endpoint.LocalAddress, "native-failure");
                var result = SessionCall(endpoint.LocalAddress, session, "suggest_text", new { request_id = "native-failure" });
                string message = (faults & 4) != 0 ? "Unconfirmed native effects remain unknown when uncertain-receipt persistence fails." :
                    "Unconfirmed native effects remain unknown across consumed-receipt and cleanup failures.";
                Assert(result.ContainsKey("structuredContent"), message);
                var error = (IDictionary<string, object>)((IDictionary<string, object>)result["structuredContent"])["error"];
                var details = (IDictionary<string, object>)error["details"];
                Assert(Object.Equals(result["isError"], true) && Object.Equals(error["code"], "outcome_unknown") &&
                    Object.Equals(Coordination(result)["state"], "outcome_unknown") && Object.Equals(details["operation_started"], true) &&
                    Object.Equals(details["outcome_unknown"], true) && details.Count == 2 && consumed && !shared && effects == 1 && persisted == ((faults & 4) == 0), message);
            }
        }
        NativeOutcomeDiagnosticFailures();
        int committed = 0;
        string receipt = null;
        using (var endpoint = new McpEndpoint(delegate(string name, IDictionary<string, object> values)
        {
            if (receipt != null) return new { status = receipt, replayed = true };
            committed++; receipt = "tracked";
            var reply = new CyclicNativeReply(); reply.Result = reply;
            return reply;
        }, Marshal, "lost-native-delivery", delegate { }))
        {
            string session = NewSession(endpoint.LocalAddress, "delivery-failure");
            var failed = SessionCall(endpoint.LocalAddress, session, "suggest_text", new { request_id = "committed" });
            Assert(Object.Equals(failed["isError"], true) && Object.Equals(Coordination(failed)["state"], "outcome_unknown") && committed == 1 && receipt == "tracked",
                "Response construction failure after a native commit reports uncertain delivery while retaining its tracked receipt.");
            var replay = SessionCall(endpoint.LocalAddress, session, "suggest_text", new { request_id = "committed" });
            Assert(Object.Equals(replay["isError"], false) && Object.Equals(Coordination(replay)["state"], "replayed") && committed == 1 && receipt == "tracked",
                "An exact retry after lost native delivery retrieves the committed receipt without another effect.");
        }
        foreach (string mode in new[] { "committed", "uncertain", "read" })
        {
            McpCoordination coordinator = null;
            McpClientSession client = null;
            int effects = 0;
            string retainedReceipt = null;
            using (var endpoint = new McpEndpoint(delegate(string name, IDictionary<string, object> values)
            {
                if (mode != "read") { effects++; retainedReceipt = mode == "committed" ? "tracked" : "outcome_uncertain"; }
                coordinator.Retire(client);
                if (mode == "uncertain") throw new McpOperationException("revoked-secret-manuscript", "outcome_unknown",
                    new { operation_started = true, outcome_unknown = true, current_context = "revoked-secret-manuscript" });
                return new { current_context = "revoked-secret-manuscript", replayed = false };
            }, Marshal, "revoked-native-delivery", delegate { }))
            {
                string session = NewSession(endpoint.LocalAddress, "revoked-delivery");
                coordinator = (McpCoordination)typeof(McpEndpoint).GetField("coordination", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(endpoint);
                client = coordinator.Require(session);
                int status; string issued;
                object arguments = mode == "read" ? (object)new { } : new { request_id = "revoked" };
                string reply = PostSession(endpoint.LocalAddress, "tools/call", new { name = mode == "read" ? "read_manuscript" : "suggest_text",
                    arguments = arguments }, out status, session, out issued);
                Assert(status == 200 && !reply.Contains("revoked-secret-manuscript"), "Revocation withholds native manuscript and current-context data.");
                if (mode == "read")
                    Assert(((IDictionary<string, object>)Json.DeserializeObject(reply)).ContainsKey("error") && effects == 0, "A revoked read remains withheld at the final response boundary.");
                else
                {
                    var result = Result(reply);
                    Assert(Object.Equals(result["isError"], true) && Object.Equals(Coordination(result)["state"], "outcome_unknown") &&
                        UInt64.Parse((string)Coordination(result)["ticket"]) > 0 && effects == 1 && retainedReceipt == (mode == "committed" ? "tracked" : "outcome_uncertain"),
                        "The final revocation boundary preserves an admitted mutation's unknown outcome and consumed receipt without leaking its content.");
                }
            }
        }
    }

    private static void NativeOutcomeDiagnosticFailures()
    {
        const string privateValue = "PRIVATE_MANUSCRIPT_CAPABILITY_REQUEST_ID https://private.invalid/t/private-cap/mcp C:\\private\\document.docx";
        Exception first = new System.Runtime.InteropServices.COMException(privateValue, unchecked((int)0x800A11FD));
        Exception rollbackFailure = new System.Runtime.InteropServices.COMException(privateValue, unchecked((int)0x80004005));
        Exception recordingFailure = new IOException(privateValue);
        Exception revokingFailure = new InvalidOperationException(privateValue);
        Exception persistenceFailure = new ArgumentException(privateValue);
        for (int faults = 0; faults < 8; faults++)
        foreach (bool primed in new[] { false, true })
        {
            var diagnostic = new WordNativeOutcome.Diagnostic();
            int calls = 0, callbackMask = 0, effects = 0, rollbackMode = primed ? faults % 4 : 0;
            bool? rollbackResult = null;
            Exception observedRollback = null;
            using (var endpoint = new McpEndpoint(delegate(string name, IDictionary<string, object> values)
            {
                effects++;
                if (primed)
                {
                    diagnostic.Capture(WordNativeOutcome.Stage.add_comment, first, false);
                    if (rollbackMode != 0)
                    {
                        try { rollbackResult = diagnostic.AttemptRollback(delegate { if (rollbackMode == 3) throw rollbackFailure; return rollbackMode == 2; }); }
                        catch (Exception error) { observedRollback = error; diagnostic.Capture(WordNativeOutcome.Stage.undo_cleanup, error, false); }
                    }
                    diagnostic.Capture(WordNativeOutcome.Stage.settings_cleanup, recordingFailure, false);
                    diagnostic.Capture(WordNativeOutcome.Stage.review_changed_callback, revokingFailure, true);
                    diagnostic.Capture(WordNativeOutcome.Stage.window_cleanup, persistenceFailure, false);
                }
                WordNativeOutcome.Unknown(
                    delegate { calls++; callbackMask |= 1; if ((faults & 1) != 0) throw recordingFailure; },
                    delegate { calls++; callbackMask |= 2; if ((faults & 2) != 0) throw revokingFailure; },
                    delegate { calls++; callbackMask |= 4; if ((faults & 4) != 0) throw persistenceFailure; }, diagnostic);
                return null;
            }, Marshal, "native-diagnostic-fixture", delegate { }))
            {
                string session = NewSession(endpoint.LocalAddress, "native-diagnostic");
                var result = SessionCall(endpoint.LocalAddress, session, "suggest_text", new { request_id = privateValue });
                Assert(result.ContainsKey("structuredContent"), "Native diagnostic failures retain their structured unknown-outcome payload.");
                var error = (IDictionary<string, object>)((IDictionary<string, object>)result["structuredContent"])["error"];
                var details = (IDictionary<string, object>)error["details"];
                var fields = (IDictionary<string, object>)details["diagnostic"];
                var rollback = (IDictionary<string, object>)fields["rollback"];
                Exception expected = primed ? first : (faults & 1) != 0 ? recordingFailure : (faults & 2) != 0 ? revokingFailure : (faults & 4) != 0 ? persistenceFailure : null;
                string phase = primed ? "add_comment" : (faults & 1) != 0 ? "uncertain_receipt" : (faults & 2) != 0 ? "revoke_authority" : (faults & 4) != 0 ? "persist_uncertainty" : null;
                Assert(Object.Equals(fields["stage"], phase) && Object.Equals(fields["exception_type"], expected == null ? null : expected.GetType().FullName) &&
                    Object.Equals(fields["hresult"], expected == null ? null : (object)expected.HResult),
                    "Native diagnostics retain the first stage, exception type and HRESULT across rollback, cleanup and callback failures.");
                Assert(Object.Equals(result["isError"], true) && Object.Equals(error["code"], "outcome_unknown") &&
                    Object.Equals(Coordination(result)["state"], "outcome_unknown") && calls == 3 && callbackMask == 7 && effects == 1 &&
                    Object.Equals(details["operation_started"], true) && Object.Equals(details["outcome_unknown"], true),
                    "Every record/revoke/persist failure combination remains unknown and attempts all callbacks exactly once.");
                Assert(Object.Equals(fields["committed"], primed) && diagnostic.Committed == primed,
                    "A later uncommitted failure cannot clear an already committed diagnostic.");
                if (rollbackMode == 1)
                    Assert(rollbackResult == false && Object.Equals(rollback["confirmed"], false) && rollback["error"] == null,
                        "A false rollback result is never reported as confirmed.");
                if (rollbackMode == 2)
                    Assert(rollbackResult == true && Object.Equals(rollback["confirmed"], true) && rollback["error"] == null,
                        "Only a successful rollback result confirms restoration.");
                if (rollbackMode == 3)
                {
                    var failure = (IDictionary<string, object>)rollback["error"];
                    Assert(Object.ReferenceEquals(observedRollback, rollbackFailure) && rollbackResult == null && failure.Count == 2 &&
                        Object.Equals(failure["exception_type"], rollbackFailure.GetType().FullName) && Object.Equals(failure["hresult"], rollbackFailure.HResult),
                        "Rollback errors retain only their own type and HRESULT without replacing the original failure.");
                }
                else Assert(rollback["error"] == null, "Unattempted or nonthrowing rollback has no fabricated error.");
                Assert(Object.Equals(rollback["attempted"], rollbackMode != 0) && Object.Equals(rollback["confirmed"], rollbackMode == 2),
                    "Rollback attempt and verified confirmation remain separate facts.");
                Assert(details.Count == 3 && fields.Count == 5 && fields.Keys.All(key => new[] { "stage", "exception_type", "hresult", "committed", "rollback" }.Contains(key)) &&
                    rollback.Count == 3 && rollback.Keys.All(key => new[] { "attempted", "confirmed", "error" }.Contains(key)),
                    "Native diagnostic details stay within their exact fixed safe-field schema.");
                string wire = Json.Serialize(result);
                Assert(!wire.Contains(privateValue) && !wire.Contains(session) && !wire.Contains(endpoint.LocalAddress) && !wire.Contains(endpoint.Path) &&
                    !wire.Contains("PRIVATE_MANUSCRIPT") && !wire.Contains("private-cap") && !wire.Contains("document.docx") && !wire.Contains("StackTrace"),
                    "Native diagnostics expose no exception message, manuscript text, capability, request ID or session ID.");
                var text = (IDictionary<string, object>)((object[])result["content"])[0];
                Assert(Json.Serialize(Json.DeserializeObject((string)text["text"])) == Json.Serialize(result["structuredContent"]),
                    "Actual HTTP text and structured errors carry identical safe native diagnostics.");
            }
        }
    }
}
