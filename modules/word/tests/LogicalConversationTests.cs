using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using UnderTree.WordAddin;

internal static partial class TransportTests
{
    private sealed class LogicalNativeFixture
    {
        internal readonly Dictionary<string, string> Choices = new Dictionary<string, string>();
        internal readonly Dictionary<string, string> Witnesses = new Dictionary<string, string>();
        internal readonly Dictionary<string, string> Receipts = new Dictionary<string, string>();
        internal int Effects, Calls;
        internal Action BeforeDispatch;

        internal object Dispatch(string name, IDictionary<string, object> args)
        {
            Calls++;
            if (BeforeDispatch != null) BeforeDispatch();
            string scope = (string)args[WordDocumentPicker.SessionKey];
            if (name == "read_manuscript")
            {
                string witness = Guid.NewGuid().ToString("N");
                Choices[scope] = "synthetic"; Witnesses[scope] = witness;
                return new { manuscript_id = "synthetic", documentId = "synthetic", revision = Effects.ToString(),
                    context_id = witness, scope = scope, complete = true };
            }
            if (!Choices.ContainsKey(scope)) throw new McpOperationException("Read this client's chosen document first.", "document_selection_required");
            var marking = (IDictionary<string, object>)args["marking"];
            string key = scope + "\0" + (string)marking["requestId"];
            string fingerprint = Json.Serialize(args);
            string recorded;
            if (Receipts.TryGetValue(key, out recorded))
            {
                if (recorded != fingerprint) throw new McpOperationException("Changed exact request denied.", "changed_request");
                return new { replayed = true, scope = scope };
            }
            string context;
            if (!Witnesses.TryGetValue(scope, out context) || context != (string)marking["contextId"])
                throw new McpOperationException("Another client's context grants no authority.", "context_changed");
            if ((string)marking["revision"] != Effects.ToString()) throw new McpOperationException("Read the new revision.", "manuscript_changed");
            Effects++; Witnesses.Clear(); Receipts.Add(key, fingerprint);
            return new { replayed = false, scope = scope };
        }

        internal void Retire(string prefix)
        {
            foreach (string key in Choices.Keys.Where(key => key.StartsWith(prefix, StringComparison.Ordinal)).ToArray()) Choices.Remove(key);
            foreach (string key in Witnesses.Keys.Where(key => key.StartsWith(prefix, StringComparison.Ordinal)).ToArray()) Witnesses.Remove(key);
        }
    }

    private static IDictionary<string, object> LogicalCall(string address, string transport, string name, object args, object metadata)
    { return Result(RawLogicalCall(address, transport, name, args, metadata)); }

    private static string RawLogicalCall(string address, string transport, string name, object args, object metadata)
    {
        int status; string issued;
        string reply = metadata == null
            ? PostSession(address, "tools/call", new { name = name, arguments = args }, out status, transport, out issued)
            : PostSession(address, "tools/call", new { name = name, arguments = args, _meta = metadata }, out status, transport, out issued);
        Assert(status == 200 && issued == null, "Logical calls retain a valid standard physical session.");
        return reply;
    }

    private static object Conversation(string hint)
    { return new Dictionary<string, object> { { "openai/session", hint } }; }

    private static IDictionary<string, object> LogicalRead(string address, string transport, string hint)
    {
        var reply = LogicalCall(address, transport, "read_manuscript", new { manuscript_id = "synthetic" }, hint == null ? null : Conversation(hint));
        Assert(Object.Equals(reply["isError"], false), "An independent logical owner can read its chosen document.");
        return (IDictionary<string, object>)reply["structuredContent"];
    }

    private static object LogicalMark(IDictionary<string, object> context, string requestId, string color = "#ffff00")
    {
        return new { documentId = "synthetic", marking = new { requestId = requestId, contextId = (string)context["context_id"],
            revision = (string)context["revision"], color = color, ranges = new[] { new { from = 0, to = 5, quote = "Alpha" } } } };
    }

    private static bool LogicalError(IDictionary<string, object> reply, string code)
    { return Object.Equals(reply["isError"], true) && Json.Serialize(reply).Contains("\"code\":\"" + code + "\""); }

    private static void LogicalConversationContinuity()
    {
        var native = new LogicalNativeFixture();
        using (var endpoint = new McpEndpoint(native.Dispatch, Marshal, "logical-authority", native.Retire))
        {
            string address = endpoint.LocalAddress;
            string a1 = NewSession(address, "ChatGPT"), a2 = NewSession(address, "ChatGPT"), a3 = NewSession(address, "ChatGPT");
            var contextA = LogicalRead(address, a1, "provider-chat-a");
            object request = LogicalMark(contextA, "same-logical-request");
            var applied = LogicalCall(address, a2, "highlight_text", request, Conversation("provider-chat-a"));
            Assert(Object.Equals(applied["isError"], false) && native.Effects == 1,
                "A labelled conversation retains its native selection and complete-read ownership across fresh transports.");
            var replay = LogicalCall(address, a3, "highlight_text", request, Conversation("provider-chat-a"));
            Assert(Object.Equals(replay["isError"], false) && Object.Equals(Coordination(replay)["state"], "replayed") && native.Effects == 1,
                "An exact labelled request replays across a third transport without another effect.");
            var changed = LogicalCall(address, a3, "highlight_text", LogicalMark(contextA, "same-logical-request", "#00ffff"), Conversation("provider-chat-a"));
            Assert(LogicalError(changed, "changed_request") && native.Effects == 1, "Logical continuity does not permit changed-envelope receipt reuse.");
            string b1 = NewSession(address, "ChatGPT"), b2 = NewSession(address, "ChatGPT");
            var contextB = LogicalRead(address, b1, "provider-chat-b");
            var continuedB = LogicalRead(address, b2, "provider-chat-b");
            Assert(Object.Equals(contextB["scope"], continuedB["scope"]) && !Object.Equals(contextA["scope"], contextB["scope"]),
                "Two different provider conversations keep distinct random native owner scopes across their own transports.");
            var theft = LogicalCall(address, b2, "highlight_text", request, Conversation("provider-chat-b"));
            Assert(LogicalError(theft, "context_changed") && native.Effects == 1 && native.Receipts.Count == 1,
                "A distinct labelled conversation cannot steal another owner's proof or replay receipt.");
            var defaultContext = LogicalRead(address, a3, null);
            Assert(!Object.Equals(defaultContext["scope"], contextA["scope"]), "Omitted provider metadata remains in the physical default namespace.");
            var omitted = LogicalCall(address, a3, "highlight_text", request, null);
            Assert(LogicalError(omitted, "context_changed") && native.Effects == 1, "Omitted metadata cannot inherit a labelled proof or receipt.");
            var unknownMetadata = LogicalCall(address, a3, "read_manuscript", new { }, new { progressToken = "unrelated" });
            Assert(Object.Equals(((IDictionary<string, object>)unknownMetadata["structuredContent"])["scope"], defaultContext["scope"]),
                "Valid metadata without a provider hint retains exactly the physical default namespace.");
            int before = native.Calls;
            foreach (object malformed in new object[] { "metadata", new object[] { "metadata" }, Conversation(""), Conversation(new string('x', 1025)), new Dictionary<string, object> { { "openai/session", 4 } } })
                Assert(Object.Equals(LogicalCall(address, a3, "read_manuscript", new { }, malformed)["isError"], true) && native.Calls == before,
                    "Malformed provider metadata is refused without entering native authority or adopting labelled state.");
            Assert(DeleteSession(address, a1) == 200, "Explicit DELETE retires its physical/default and used logical state.");
            var revoked = LogicalCall(address, a2, "highlight_text", request, Conversation("provider-chat-a"));
            Assert(LogicalError(revoked, "document_selection_required") && native.Effects == 1,
                "Deleting one used transport revokes the old logical selection across other physical sessions.");
            var freshA = LogicalRead(address, a2, "provider-chat-a");
            Assert(!Object.Equals(freshA["scope"], contextA["scope"]), "A reused provider hint receives a fresh owner generation after explicit revocation.");
            Assert(LogicalError(LogicalCall(address, a2, "highlight_text", request, Conversation("provider-chat-a")), "context_changed") && native.Effects == 1,
                "A fresh logical generation cannot resume the revoked generation's proof or request receipt.");
            var survivedB = LogicalRead(address, b2, "provider-chat-b");
            Assert(Object.Equals(survivedB["scope"], contextB["scope"]), "Deleting chat A preserves chat B's independent logical owner.");
            for (int index = 0; index < 70; index++)
            {
                bool chatA = index % 2 == 0;
                string transport = NewSession(address, "ChatGPT");
                var context = LogicalRead(address, transport, chatA ? "provider-chat-a" : "provider-chat-b");
                Assert(Object.Equals(context["scope"], chatA ? freshA["scope"] : contextB["scope"]),
                    "More than 64 fresh initialize/tool pairs retain two bounded logical owners rather than consuming one owner per transport.");
            }
            var registry = (McpCoordination)typeof(McpEndpoint).GetField("coordination", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(endpoint);
            var labelled = (Dictionary<string, McpLogicalConversation>)typeof(McpCoordination).GetField("conversations", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(registry);
            Assert(labelled.Count == 2 && native.Choices.Count == 3, "Repeated physical reconnects create two labelled scopes and one strict default, without registry growth.");
        }
    }

    private static void LogicalConversationLifetime()
    {
        long now = 0;
        var registry = new McpCoordination(() => now, 100);
        var first = registry.Create();
        McpCallScope original = registry.Resolve(first, "chat-a");
        var second = registry.Create();
        now = 90;
        McpCallScope continued = registry.Resolve(second, "chat-a");
        now = 100;
        registry.Require(second.Id);
        Assert(!registry.IsActive(original) && registry.IsActive(continued) && original.NativeScope == continued.NativeScope,
            "Physical idle expiry cancels its immutable call handle without invalidating a still-active logical owner's witness namespace.");
        var retired = new List<string>();
        registry.DrainRetired(prefix => retired.Add(prefix));
        Assert(retired.SequenceEqual(new[] { first.ScopePrefix }), "Physical idle cleanup retires only its default namespace, not its live labelled owner.");
        now = 150;
        McpCallScope other = registry.Resolve(second, "chat-b");
        now = 190;
        registry.Require(second.Id);
        Assert(!registry.IsActive(continued) && registry.IsActive(other) && registry.IsActive(second),
            "Logical idle expiry independently revokes an old conversation while its transport and another label stay live.");
        registry.DrainRetired(prefix => retired.Add(prefix));
        McpCallScope replacement = registry.Resolve(second, "chat-a");
        Assert(replacement.NativeScope != continued.NativeScope && !registry.IsActive(continued) && registry.IsActive(replacement),
            "A reused hint gets a fresh random generation; old immutable handles cannot attach to its replacement.");
        long ticket = registry.Enqueue();
        string code = null;
        try { registry.Await(ticket, continued, Stopwatch.GetTimestamp() + Stopwatch.Frequency); }
        catch (McpOperationException error) { code = error.Code; }
        finally { registry.Finish(ticket); }
        Assert(code == "session_revoked", "An expired logical handle cannot enter native authority even when its transport and replacement label are live.");
        var third = registry.Create();
        McpCallScope shared = registry.Resolve(third, "chat-a");
        var fourth = registry.Create();
        McpCallScope survivor = registry.Resolve(fourth, "chat-c");
        registry.Retire(second);
        Assert(!registry.IsActive(replacement) && !registry.IsActive(shared) && !registry.IsActive(other) && registry.IsActive(third) && registry.IsActive(survivor),
            "Explicit DELETE revokes every logical generation used by that transport across other sessions, preserving unrelated labels.");
        Assert(third.Conversations.Count == 0 && fourth.Conversations.Count == 1, "Retired logical references are pruned from every physical session.");
        bool cleanupFailed = false;
        try { registry.DrainRetired(delegate { throw new IOException("Synthetic native cleanup failure."); }); }
        catch (IOException) { cleanupFailed = true; }
        Assert(cleanupFailed && registry.HasRetired, "Failed native retirement remains queued for retry without reviving its logical handle.");
        registry.DrainRetired(prefix => retired.Add(prefix));
        Assert(!registry.HasRetired && retired.Contains(other.Conversation.ScopePrefix) && retired.Contains(replacement.Conversation.ScopePrefix),
            "Successful retry drains all retained logical and physical cleanup exactly into the native revocation path.");
        registry.Stop();
        Assert(!registry.IsActive(survivor) && !registry.IsActive(third) && fourth.Conversations.Count == 0,
            "Stop revokes every logical owner and physical handle and releases membership references.");
    }

    private static void LogicalConversationBounds()
    {
        long now = 0;
        var physical = new McpCoordination(() => now, 100);
        var sessions = Enumerable.Range(0, McpCoordination.MaxSessions).Select(index => physical.Create()).ToArray();
        Assert(PhysicalOverflow(physical) && physical.IsActive(sessions[0]), "The 512-session physical limit refuses overflow without evicting a live client.");
        physical.Retire(sessions[0]);
        Assert(PhysicalOverflow(physical), "Pending physical cleanup continues to consume bounded admission capacity.");
        try { physical.DrainRetired(delegate { throw new IOException("Keep physical retirement queued."); }); } catch (IOException) { }
        Assert(PhysicalOverflow(physical), "Failed native cleanup cannot free physical capacity prematurely.");
        physical.DrainRetired(delegate { });
        Assert(physical.Create() != null, "Successful physical cleanup releases exactly its retired admission slot.");
        physical.Stop();

        var logical = new McpCoordination(() => now, 100);
        McpClientSession transport = logical.Create();
        var owners = Enumerable.Range(0, McpCoordination.MaxConversations).Select(index => logical.Resolve(transport, "bounded-" + index)).ToArray();
        Assert(LogicalOverflow(logical, transport) && logical.IsActive(owners[0]), "The 64-owner logical limit refuses overflow without evicting an active conversation.");
        Assert(logical.Resolve(transport, null).Conversation == null, "A full logical registry still permits the transport's separate default namespace.");
        logical.Retire(transport);
        var reconnect = logical.Create();
        Assert(LogicalOverflow(logical, reconnect), "Pending logical cleanup continues to consume bounded owner capacity across fresh transports.");
        try { logical.DrainRetired(delegate { throw new IOException("Keep logical retirement queued."); }); } catch (IOException) { }
        Assert(LogicalOverflow(logical, reconnect), "Failed native cleanup cannot free logical capacity or resurrect prior owner scopes.");
        var cleaned = new List<string>();
        logical.DrainRetired(prefix => cleaned.Add(prefix));
        McpCallScope fresh = logical.Resolve(reconnect, "bounded-0");
        Assert(cleaned.Count == 65 && cleaned.Distinct().Count() == 65 && fresh.NativeScope != owners[0].NativeScope,
            "Cleanup releases the 64 logical entries and one physical entry without duplicating or inheriting their scopes.");
        for (int index = 0; index < 4; index++)
        {
            now += 99; logical.Require(reconnect.Id); now++;
            Assert(!logical.IsActive(fresh), "Logical expiry is independent from recently refreshed physical activity.");
            logical.DrainRetired(delegate { });
            fresh = logical.Resolve(reconnect, "bounded-0");
            Assert(reconnect.Conversations.Count == 1, "Repeated logical expiry and recreation cannot accumulate dead membership generations.");
        }
        logical.Stop();
    }

    private static bool PhysicalOverflow(McpCoordination registry)
    { try { registry.Create(); return false; } catch (McpHttpException error) { return error.Status == 429; } }

    private static bool LogicalOverflow(McpCoordination registry, McpClientSession session)
    {
        try { registry.Resolve(session, "overflow-owner"); return false; }
        catch (McpOperationException error)
        { return error.Code == "conversation_limit" && Json.Serialize(error.Details).Contains("\"operation_started\":false") && Json.Serialize(error.Details).Contains("\"outcome_unknown\":false"); }
    }

    private static void LogicalConversationQueuedRevocation()
    {
        using (var entered = new ManualResetEvent(false))
        using (var release = new ManualResetEvent(false))
        {
            var native = new LogicalNativeFixture();
            int hold = 0;
            native.BeforeDispatch = delegate
            {
                if (Interlocked.Exchange(ref hold, 0) != 1) return;
                entered.Set();
                if (!release.WaitOne(5000)) throw new TimeoutException("Logical queue barrier was not released.");
            };
            using (var endpoint = new McpEndpoint(native.Dispatch, Marshal, "logical-queued-revocation", native.Retire))
            {
                try
                {
                    string address = endpoint.LocalAddress;
                    string a1 = NewSession(address, "ChatGPT"), a2 = NewSession(address, "ChatGPT"), b = NewSession(address, "Claude");
                    var contextA = LogicalRead(address, a1, "queued-chat-a");
                    var contextB = LogicalRead(address, b, "queued-chat-b");
                    hold = 1;
                    Task<string> held = Task.Run(() => RawLogicalCall(address, b, "read_manuscript", new { }, Conversation("queued-chat-b")));
                    Assert(entered.WaitOne(5000), "An unrelated logical owner holds a real native queue turn.");
                    Task<string> queued = Task.Run(() => RawLogicalCall(address, a2, "highlight_text", LogicalMark(contextA, "deleted-logical-queue"), Conversation("queued-chat-a")));
                    Wait(() => PendingTurns(endpoint) == 2, "logical handle queued across a different physical session");
                    Task<int> deletion = Task.Run(() => DeleteSession(address, a1));
                    Assert(queued.Wait(5000) && queued.Result.Contains("revoked") && native.Effects == 0 && !held.IsCompleted,
                        "DELETE cancels an already queued same-label call on another physical session before native start.");
                    release.Set();
                    Assert(held.Wait(5000) && deletion.Wait(5000) && deletion.Result == 200 && !held.Result.Contains("\"error\""),
                        "Deleting chat A leaves chat B's running native turn and physical transport usable.");
                    var surviving = LogicalRead(address, b, "queued-chat-b");
                    Assert(Object.Equals(surviving["scope"], contextB["scope"]), "An unrelated logical choice and owner survive queued cancellation and native retirement cleanup.");
                    entered.Reset(); release.Reset(); hold = 1;
                    held = Task.Run(() => RawLogicalCall(address, b, "read_manuscript", new { }, Conversation("queued-chat-b")));
                    Assert(entered.WaitOne(5000), "The Stop case reaches a running native turn.");
                    string b2 = NewSession(address, "Claude");
                    queued = Task.Run(() => RawLogicalCall(address, b2, "highlight_text", LogicalMark(surviving, "stopped-logical-queue"), Conversation("queued-chat-b")));
                    Wait(() => PendingTurns(endpoint) == 2, "logical queue before Stop");
                    endpoint.Dispose();
                    Assert(queued.Wait(5000) && queued.Result.Contains("revoked") && native.Effects == 0,
                        "Stop cancels queued labelled calls on every physical session without an effect.");
                    release.Set();
                    Assert(held.Wait(5000) && held.Result.Contains("revoked"), "Stop withholds the running read's data at final delivery.");
                }
                finally { release.Set(); }
            }
        }
    }

    private sealed class LogicalRevokingReply
    {
        private readonly Action revoke;
        internal LogicalRevokingReply(Action revoke) { this.revoke = revoke; }
        public bool replayed { get { return false; } }
        public string current_context { get { revoke(); return "logical-secret-manuscript"; } }
    }

    private static void LogicalConversationNativeAdmission()
    {
        using (var admitted = new ManualResetEvent(false))
        using (var release = new ManualResetEvent(false))
        {
            int delay = 0;
            Func<Func<object>, object> delayedMarshal = callback =>
            {
                if (Interlocked.Exchange(ref delay, 0) == 1)
                {
                    admitted.Set();
                    if (!release.WaitOne(5000)) throw new TimeoutException("Native admission barrier was not released.");
                }
                return Marshal(callback);
            };
            var native = new LogicalNativeFixture();
            using (var endpoint = new McpEndpoint(native.Dispatch, delayedMarshal, "logical-native-admission", native.Retire))
            {
                try
                {
                    string address = endpoint.LocalAddress;
                    string first = NewSession(address, "ChatGPT"), second = NewSession(address, "ChatGPT");
                    var context = LogicalRead(address, first, "admitted-owner");
                    int before = native.Calls;
                    delay = 1;
                    Task<string> pending = Task.Run(() => RawLogicalCall(address, second, "highlight_text", LogicalMark(context, "admitted-but-not-native"), Conversation("admitted-owner")));
                    Assert(admitted.WaitOne(5000) && PendingTurns(endpoint) == 1, "A logical call passes FIFO admission before its native callback begins.");
                    Assert(DeleteSession(address, first) == 200, "Another used transport can revoke that admitted logical handle before native dispatch.");
                    var registry = (McpCoordination)typeof(McpEndpoint).GetField("coordination", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(endpoint);
                    registry.Resolve(registry.Require(second), "admitted-owner");
                    release.Set();
                    Assert(pending.Wait(5000) && pending.Result.Contains("revoked") && native.Calls == before && native.Effects == 0,
                        "Native admission rechecks the original immutable logical handle and never dispatches a revoked generation or its fresh replacement.");
                }
                finally { release.Set(); }
            }
        }
    }

    private static void LogicalConversationRevokedDelivery()
    {
        foreach (string mode in new[] { "committed", "legacy-tracked", "uncertain", "read", "after-response" })
        {
            McpCoordination registry = null;
            McpClientSession retiredTransport = null;
            int effects = 0;
            string receipt = null;
            using (var endpoint = new McpEndpoint(delegate(string name, IDictionary<string, object> args)
            {
                if (mode != "read") { effects++; receipt = mode == "uncertain" ? "outcome_uncertain" : mode == "legacy-tracked" ? "tracked" : "applied"; }
                if (mode == "after-response") return new LogicalRevokingReply(() => registry.Retire(retiredTransport));
                registry.Retire(retiredTransport);
                if (mode == "uncertain") throw new McpOperationException("logical-secret-manuscript", "outcome_unknown",
                    new { operation_started = true, outcome_unknown = true, current_context = "logical-secret-manuscript" });
                return new { current_context = "logical-secret-manuscript", replayed = false };
            }, Marshal, "logical-revoked-delivery", delegate { }))
            {
                string address = endpoint.LocalAddress;
                string a1 = NewSession(address, "ChatGPT"), a2 = NewSession(address, "ChatGPT");
                registry = (McpCoordination)typeof(McpEndpoint).GetField("coordination", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(endpoint);
                retiredTransport = registry.Require(a1);
                registry.Resolve(retiredTransport, "revoked-logical-owner");
                object arguments = mode == "read" ? (object)new { } : new { request_id = "logical-revoked" };
                string reply = RawLogicalCall(address, a2, mode == "read" ? "read_manuscript" : "suggest_text", arguments, Conversation("revoked-logical-owner"));
                Assert(registry.IsActive(registry.Require(a2)) && !reply.Contains("logical-secret-manuscript"),
                    "Logical revocation with a still-live physical transport withholds manuscript and current-context data.");
                if (mode == "read")
                    Assert(((IDictionary<string, object>)Json.DeserializeObject(reply)).ContainsKey("error") && effects == 0,
                        "A logically revoked read remains withheld at the final response boundary.");
                else
                {
                    var result = Result(reply);
                    Assert(Object.Equals(result["isError"], true) && Object.Equals(Coordination(result)["state"], "outcome_unknown") &&
                        UInt64.Parse((string)Coordination(result)["ticket"]) > 0 && effects == 1 && receipt == (mode == "uncertain" ? "outcome_uncertain" : mode == "legacy-tracked" ? "tracked" : "applied"),
                        "Logical revocation after a possible effect preserves unknown delivery, its ticket and consumed receipt without data leakage.");
                }
            }
        }
    }
}
