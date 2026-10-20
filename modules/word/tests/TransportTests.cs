using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using UnderTree.WordAddin;

internal static partial class TransportTests
{
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 16777216 };
    private static int checks;
    private static int dispatches;
    private static int revoked;
    private static int actorThread;
    private static string expectedTool;
    private static bool failDispatch;
    private static bool markerFailure;
    private static readonly BlockingCollection<Action> Queue = new BlockingCollection<Action>();
    private static readonly ConcurrentDictionary<string, string> Sessions = new ConcurrentDictionary<string, string>();

    public static int Main(string[] arguments)
    {
        string root = Path.GetFullPath(arguments[0]);
        string fake = Path.GetFullPath(arguments[1]);
        Environment.SetEnvironmentVariable("ZROK2_API_ENDPOINT", "https://foreign-enrollment.example");
        Environment.SetEnvironmentVariable("ZROK2_UNKNOWN_OVERRIDE", "foreign-profile");
        ThreadPool.SetMinThreads(32, 32);
        Thread actor = new Thread(delegate() { actorThread = Thread.CurrentThread.ManagedThreadId; foreach (Action action in Queue.GetConsumingEnumerable()) action(); });
        actor.SetApartmentState(ApartmentState.STA); actor.IsBackground = true; actor.Start();
        try
        {
            if (arguments.Length > 2 && arguments[2] == "logical-only")
            {
                LogicalConversationContinuity();
                LogicalConversationLifetime();
                LogicalConversationBounds();
                LogicalConversationQueuedRevocation();
                LogicalConversationNativeAdmission();
                LogicalConversationRevokedDelivery();
                Console.WriteLine("PASS: " + checks + " logical-conversation HTTP assertions; disposable fake native authority only.");
                return 0;
            }
            if (arguments.Length > 2 && arguments[2] == "concurrency-only")
            {
                ConcurrentClients(Path.Combine(root, "concurrent-clients"), fake);
                MonotonicSessionExpiry();
                UnknownRunningTurn();
                NativeOutcomeFailures();
                Console.WriteLine("PASS: " + checks + " concurrent-client HTTP assertions; disposable fake native authority only.");
                return 0;
            }
            VisualHighlights(Path.Combine(root, "visual-highlights"), fake);
            LocalEndpoint(Path.Combine(root, "local"), fake);
            TunnelLifecycle(Path.Combine(root, "tunnel"), fake);
            ForeignSharePreserved(Path.Combine(root, "foreign"), fake);
            ChildDeath(Path.Combine(root, "death"), fake);
            FailedCleanupRetry(Path.Combine(root, "cleanup"), fake);
            InvalidStatePreserved(Path.Combine(root, "invalid"), fake);
            QueuedRevocation(Path.Combine(root, "queued"), fake);
            FailedRevocation(Path.Combine(root, "revoke"), fake);
            StaShutdown(Path.Combine(root, "sta-shutdown"), fake);
            InterruptedCleanupRecovery(Path.Combine(root, "shutdown-recovery"), fake);
            PendingWorkerRevocation(Path.Combine(root, "worker-revoke"), fake);
            DisposeRevocationFailure(Path.Combine(root, "dispose-failure"), fake);
            ActiveDocumentFirstRead(Path.Combine(root, "active-first-read"), fake);
            ActiveDocumentRecovery(Path.Combine(root, "changed-document"), fake);
            ConversationRouting(Path.Combine(root, "conversation-routing"), fake);
            LogicalConversationContinuity();
            LogicalConversationLifetime();
            LogicalConversationBounds();
            LogicalConversationQueuedRevocation();
            LogicalConversationNativeAdmission();
            LogicalConversationRevokedDelivery();
            SuggestionCurrentContext(Path.Combine(root, "current-context"), fake);
            LegalEscapedPayloads(Path.Combine(root, "escaped-payloads"), fake);
            ConcurrentClients(Path.Combine(root, "concurrent-clients"), fake);
            MonotonicSessionExpiry();
            UnknownRunningTurn();
            NativeOutcomeFailures();
            Console.WriteLine("PASS: " + checks + " transport assertions; actual loopback HTTP, isolated filesystem and fake zrok subprocess only.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally { Queue.CompleteAdding(); actor.Join(1000); }
    }

    private static object Marshal(Func<object> callback)
    {
        object result = null; Exception failure = null;
        using (ManualResetEvent finished = new ManualResetEvent(false))
        {
            Queue.Add(delegate { try { result = callback(); } catch (Exception error) { failure = error; } finally { finished.Set(); } });
            if (!finished.WaitOne(10000)) throw new TimeoutException("Test Word actor stalled.");
        }
        if (failure != null) throw failure;
        return result;
    }

    private static ConnectionService Create(string directory, string fake)
    {
        return new ConnectionService(delegate(string name, IDictionary<string, object> values)
        {
            Assert(Thread.CurrentThread.ManagedThreadId == actorThread, "Review dispatch executes on Word STA.");
            dispatches++;
            expectedTool = name;
            if (failDispatch) throw new InvalidOperationException("SECRET C:\\private\\manuscript.docx");
            if (markerFailure) throw new McpOperationException("Word cannot verify the current editing permission for a marked passage.", "marker_permission_unresolved",
                new { markers = new[] { new { id = "Wrai_test", ruleId = "protect", reason = "native_bookmark_missing" } }, recovery = "Inspect the named passage in Word." });
            return new { manuscripts = new object[0], supplied = values.ContainsKey("quote") ? values["quote"] : null };
        }, delegate { Interlocked.Increment(ref revoked); }, Marshal, directory, fake);
    }

    private static void VisualHighlights(string directory, string fake)
    {
        int nativeCalls = 0;
        IDictionary<string, object> receivedMarking = null;
        using (var service = new ConnectionService(delegate(string name, IDictionary<string, object> values) {
            Assert(Thread.CurrentThread.ManagedThreadId == actorThread && name == "highlight_text", "Visual RGB marking reaches only the Word highlight authority on its STA.");
            nativeCalls++;
            receivedMarking = (IDictionary<string, object>)values["marking"];
            return new { colorId = receivedMarking.ContainsKey("colorId") ? receivedMarking["colorId"] : "native-resolved-green" };
        }, delegate { }, Marshal, directory, fake))
        {
            service.StartLocal();
            string instructions = (string)Result(Post(service.LocalAddress, "initialize", new { protocolVersion = "2025-06-18" }))["instructions"];
            Assert(instructions.Contains("including unmarked text") && instructions.Contains("without rule or wording-permission questions") &&
                instructions.Contains("Reuse an existing colorId with its unchanged local rule") && instructions.Contains("ambiguous RGB requires exact colorId"),
                "A color request immediately applies existing configured colors or requested RGB without changing their local rules.");
            Assert(instructions.Contains("Neutral coloration grants no wording, protect or discard authority") &&
                instructions.Contains("only the local user assigns color functions"), "Visual colors cannot grant themselves editing or protection functions.");
            Assert(instructions.Contains("#00FF00 to #FF0000") && instructions.Contains("red may remain neutral with an empty prompt") &&
                instructions.Contains("Do not ask for an editing prompt"), "Recolor discovery explicitly allows neutral red without an editing-prompt question.");
            object[] tools = (object[])Result(Post(service.LocalAddress, "tools/list", new { }))["tools"];
            var tool = (IDictionary<string, object>)tools[3];
            var schema = (IDictionary<string, object>)tool["inputSchema"];
            var marking = (IDictionary<string, object>)((IDictionary<string, object>)schema["properties"])["marking"];
            var properties = (IDictionary<string, object>)marking["properties"];
            var color = (IDictionary<string, object>)properties["color"];
            Assert(Object.Equals(color["type"], "string") && Object.Equals(color["pattern"], "^#[0-9A-Fa-f]{6}$") &&
                !((object[])marking["required"]).Contains("color") && !((object[])marking["required"]).Contains("colorId"),
                "RGB is an optional exact six-digit color; existing color identity and automatic-color calls remain compatible.");
            Assert(((string)color["description"]).Contains("visual color") && ((string)tool["description"]).Contains("previously unmarked text") &&
                ((string)tool["description"]).Contains("existing color/rule identity") &&
                ((string)((IDictionary<string, object>)properties["colorId"])["description"]).Contains("unchanged"),
                "Tool discovery permits unchanged configured-color reuse and distinguishes fresh neutral RGB from ambiguous identities.");
            var ranges = (IDictionary<string, object>)properties["ranges"];
            Assert(Object.Equals(marking["additionalProperties"], false) && Object.Equals(ranges["maxItems"], 64) &&
                ((string)ranges["description"]).Contains("ascending UTF-16 order"), "Visual marking retains the closed, bounded and sorted atomic-range contract.");
            object[] exactRanges = new object[] { new { from = 5, to = 19, quote = "(Author, 2026)" }, new { from = 31, to = 45, quote = "(Writer, 2025)" } };
            var reply = Result(Post(service.LocalAddress, "tools/call", new { name = "highlight_text", arguments = new { documentId = "chosen", marking = new {
                requestId = "citations-green", revision = "current-revision", contextId = "complete-context", color = "#00ff00", ranges = exactRanges } } }));
            Assert(Object.Equals(reply["isError"], false) && nativeCalls == 1 && Object.Equals(receivedMarking["color"], "#00ff00") &&
                !receivedMarking.ContainsKey("colorId") && ((object[])receivedMarking["ranges"]).Length == 2,
                "Requested citation RGB and exact ranges cross one loopback call unchanged for native identity resolution.");
            var configured = Result(Post(service.LocalAddress, "tools/call", new { name = "highlight_text", arguments = new { documentId = "chosen", marking = new {
                requestId = "extend-protect-green", revision = "current-revision", contextId = "complete-context", colorId = "protect-green", ranges = exactRanges } } }));
            Assert(Object.Equals(configured["isError"], false) && nativeCalls == 2 && Object.Equals(receivedMarking["colorId"], "protect-green") &&
                !receivedMarking.ContainsKey("color"), "An exact existing functional color identity can be extended without sending any rule definition.");
            var recolorSchema = (IDictionary<string, object>)properties["recolorMarkIds"];
            Assert(Object.Equals(recolorSchema["type"], "array") && Object.Equals(recolorSchema["maxItems"], 64) &&
                Object.Equals(recolorSchema["uniqueItems"], true) && !((object[])marking["required"]).Contains("recolorMarkIds"),
                "Recolor discovery exposes optional bounded ordered marker identities without changing ordinary highlighting.");
            var recolored = Result(Post(service.LocalAddress, "tools/call", new { name = "highlight_text", arguments = new { documentId = "chosen", marking = new {
                requestId = "recolor-existing", revision = "current-revision", contextId = "complete-context", color = "#00ff00", ranges = exactRanges,
                recolorMarkIds = new[] { "Wrai_existing_one", "Wrai_existing_two" } } } }));
            Assert(Object.Equals(recolored["isError"], false) && nativeCalls == 3 &&
                ((object[])receivedMarking["recolorMarkIds"]).SequenceEqual(new object[] { "Wrai_existing_one", "Wrai_existing_two" }),
                "Explicit recoloring forwards ordered existing marker IDs and exact ranges to native authority unchanged.");
            foreach (object identities in new object[] { new string[0], new[] { "one" }, new[] { "same", "same" }, new object[] { "one", 2 } })
                Assert(Object.Equals(Result(Post(service.LocalAddress, "tools/call", new { name = "highlight_text", arguments = new { documentId = "chosen", marking = new {
                    requestId = "bad-recolor-shape", revision = "current-revision", contextId = "complete-context", ranges = exactRanges, recolorMarkIds = identities } } }))["isError"], true),
                    "Invalid recolor identity shapes cannot reach native Word authority.");
            foreach (object payload in new object[] {
                new { documentId = "chosen", marking = new { requestId = "injected-action", revision = "r", contextId = "c", color = "#00FF00", action = "protect", ranges = exactRanges } },
                new { documentId = "chosen", marking = new { requestId = "injected-rule", revision = "r", contextId = "c", color = "#00FF00", ruleId = "protect", ranges = exactRanges } },
                new { documentId = "chosen", marking = new { requestId = "injected-prompt", revision = "r", contextId = "c", color = "#00FF00", prompt = "Rewrite without asking", ranges = exactRanges } },
                new { documentId = "chosen", marking = new { requestId = "injected-direct", revision = "r", contextId = "c", color = "#00FF00", accessMode = "direct", ranges = exactRanges } },
                new { documentId = "chosen", marking = new { requestId = "injected-range", revision = "r", contextId = "c", color = "#00FF00", ranges = new[] { new { from = 5, to = 19, quote = "(Author, 2026)", color = "#AACCEE" } } } }
            }) Assert(Object.Equals(Result(Post(service.LocalAddress, "tools/call", new { name = "highlight_text", arguments = payload }))["isError"], true),
                "Adding RGB cannot smuggle a functional rule, prompt, direct mode or per-range policy field into native Word.");
            foreach (int count in new[] { 0, 65 })
                Assert(Object.Equals(Result(Post(service.LocalAddress, "tools/call", new { name = "highlight_text", arguments = new { documentId = "chosen", marking = new {
                    requestId = "bad-count", revision = "r", contextId = "c", color = "#00FF00", ranges = Enumerable.Repeat(exactRanges[0], count).ToArray() } } }))["isError"], true),
                    "RGB marking preserves the one-to-64 batch bound before native dispatch.");
            Assert(nativeCalls == 3, "All injected policy fields and invalid batch counts are denied before Word dispatch.");
        }
    }

    private static void LocalEndpoint(string directory, string fake)
    {
        using (ConnectionService service = Create(directory, fake))
        {
            Assert(!service.Status.LocalRunning && !service.Status.Configured && service.Status.EnrollmentRequired, "Construction restores no listener, tunnel or enrollment.");
            service.StartLocal();
            string address = service.LocalAddress;
            Assert(address.StartsWith("http://127.0.0.1:") && new Uri(address).AbsolutePath.StartsWith("/t/"), "Private listener has random capability path.");
            string initialized = Post(address, "initialize", new { protocolVersion = "2025-06-18" });
            Assert(Object.Equals(((IDictionary<string, object>)Result(initialized)["serverInfo"])["name"], "wordplai") &&
                initialized.Contains("2025-06-18") && initialized.Contains("context_id"), "Initialize identifies lowercase wordplai, protocol and complete-read requirement.");
            string instructions = (string)Result(initialized)["instructions"];
            Assert(instructions.Contains("Immediately call read_manuscript, omitting manuscript_id when unknown") &&
                instructions.Contains("read_manuscript sequentially until complete=true") &&
                instructions.Contains("Use the returned identity for pagination and edits"), "A writing session reads the complete active document directly without a discovery round trip.");
            Assert(instructions.Contains("subject, purpose, structure and writing style") &&
                instructions.Contains("If there is no editing intent and no applicable writing task") && instructions.Contains("one brief question grounded in the actual document"),
                "The editor asks for direction only when both editing intent and an applicable task are absent.");
            Assert(instructions.Contains("'use wordplai MCP' alone, 'use MCP, edit text', 'edit', 'write' or 'improve this' are complete writing tasks") &&
                instructions.Contains("apply the saved color and passage prompts immediately") &&
                instructions.Contains("do not ask the user to repeat them or specify what to change"),
                "The wordplai-only trigger and minimal edit requests immediately execute marked instructions without a clarification loop.");
            Assert(instructions.StartsWith("You are wordplai, the writing assistant in Word. Word is the primary workspace.", StringComparison.Ordinal) &&
                instructions.Contains("short completion confirmation"), "wordplai starts in its writing role, with finished writing in Word and a short chat confirmation.");
            Assert(instructions.Contains("Immediately carry out an already-given writing request") &&
                instructions.Contains("Submit each paragraph as soon as ready") && instructions.Contains("do not wait to draft all paragraphs"),
                "A concrete task starts immediately and the first finished paragraph is submitted before drafting the rest.");
            Assert(instructions.Contains("No preliminary prose, capability lists, tool narration, permission talk or redundant questions"),
                "Writing proceeds without preliminary explanations or unnecessary questions.");
            Assert(instructions.Contains("State a genuine blocker once") && instructions.Contains("never retry it in a loop") &&
                instructions.Contains("continue independent authorized passages"), "Real native blockers remain explicit and safe without repetitive retries or abandoning unrelated writable passages.");
            Assert(instructions.Contains("No read confirmation is needed") && instructions.Contains("Never claim a read or edit succeeded without its result"),
                "Already-authorized reading needs no confirmation and the editor reports only confirmed outcomes.");
            Assert(instructions.Contains("current_context in a successful result") && instructions.Contains("Do not read again just to verify success") &&
                instructions.Contains("re-read only when current context is unavailable or Word reports a change"),
                "A successful native receipt supplies the next paragraph's context without a redundant verification read.");
            Assert(instructions.Contains("If selection_required=true") && instructions.Contains("wait for the user's choice") &&
                instructions.Contains("Keep that document regardless of the active Word window") && instructions.Contains("switch only when the user explicitly names another") &&
                instructions.Contains("For document_closed") && instructions.Contains("never retarget another window automatically"),
                "Ambiguous documents require one choice, retained across window changes and never replaced after closure.");
            Assert(instructions.Contains("applies eligible highlighted wording directly") && instructions.Contains("legacy prompt rules stored with suggestions access") &&
                instructions.Contains("Discard remains a complete-removal suggestion") && instructions.Contains("existing pending or tracked receipts never become new direct edits") &&
                !instructions.Contains("always creates native Word Track Changes"),
                "Word fixes eligible prompt wording to direct editing while retaining restrictive discard and legacy receipt semantics.");
            Assert(instructions.Contains("saves existing named writable documents") && instructions.Contains("first ordinary Save As") &&
                instructions.Contains("Completed direct edits become green and protected") && instructions.Contains("accepted earlier suggestions do too"),
                "Native autosave and local acceptance protection do not add a remote save or approval operation.");
            IDictionary<string, object> catalogue = Result(Post(address, "tools/list", new { }));
            object[] tools = (object[])catalogue["tools"];
            Assert(tools.Length == 4 && ((Dictionary<string, object>)tools[2])["name"].Equals("suggest_text") &&
                ((Dictionary<string, object>)tools[3])["name"].Equals("highlight_text"), "Catalogue has exactly four closed Review tools.");
            var pickerTool = (IDictionary<string, object>)tools[0];
            Assert(pickerTool.ContainsKey("outputSchema") && pickerTool["outputSchema"] is IDictionary<string, object>,
                "The picker-linked tool declares its structured result schema for host validation.");
            File.WriteAllText(Path.Combine(directory, "picker-output-schema.json"), Json.Serialize(pickerTool["outputSchema"]));
            Assert(((string)((IDictionary<string, object>)tools[0])["description"]).Contains("all open Word document names") &&
                ((string)((IDictionary<string, object>)tools[0])["description"]).Contains("without changing the chosen writing target") &&
                ((string)((IDictionary<string, object>)tools[1])["description"]).Contains("Omit manuscript_id on the first call"),
                "Discovery lists all open documents without selecting one; a direct read offers choice only when needed.");
            Assert(((string)((IDictionary<string, object>)tools[1])["description"]).Contains("Generic editing requests immediately use saved passage prompts") &&
                !((string)((IDictionary<string, object>)tools[1])["description"]).Contains("only if no task was given"),
                "Read-tool guidance recognizes generic editing intent without asking the user to restate the marked task.");
            object readMetadataValue;
            Assert(((IDictionary<string, object>)tools[1]).TryGetValue("_meta", out readMetadataValue) && readMetadataValue is IDictionary<string, object>,
                "The read helper advertises metadata for document choice through the legacy widget bridge.");
            var readMetadata = (IDictionary<string, object>)readMetadataValue;
            var readUi = (IDictionary<string, object>)readMetadata["ui"];
            Assert(Object.Equals(readMetadata["openai/widgetAccessible"], true) &&
                ((object[])readUi["visibility"]).SequenceEqual(new object[] { "model", "app" }),
                "The read helper explicitly supports model/app visibility and the legacy widget bridge.");
            Assert(!readUi.ContainsKey("resourceUri") && !readMetadata.ContainsKey("openai/outputTemplate") &&
                Object.Equals(((IDictionary<string, object>)((IDictionary<string, object>)((IDictionary<string, object>)tools[0])["_meta"])["ui"])["resourceUri"], WordDocumentPicker.Uri),
                "Only explicit document listing mounts a picker; an ordinary one-document read cannot mount selection UI.");
            foreach (IDictionary<string, object> listedTool in tools)
            {
                var annotations = (IDictionary<string, object>)listedTool["annotations"];
                Assert(annotations["readOnlyHint"] is bool && annotations["destructiveHint"] is bool &&
                    annotations.ContainsKey("openWorldHint") && Object.Equals(annotations["openWorldHint"], false),
                    "Every bounded Word tool explicitly advertises its read-only, destructive and closed-world meanings.");
                if (Object.Equals(listedTool["name"], "suggest_text") || Object.Equals(listedTool["name"], "highlight_text"))
                    Assert(!listedTool.ContainsKey("_meta"), "Writing tools receive no widget metadata or accessibility override.");
            }
            var readSchema = (IDictionary<string, object>)((IDictionary<string, object>)tools[1])["inputSchema"];
            Assert(!readSchema.ContainsKey("required") &&
                ((IDictionary<string, object>)readSchema["properties"]).ContainsKey("manuscript_id") &&
                Object.Equals(readSchema["additionalProperties"], false),
                "The closed read schema admits an empty first read and retains explicit manuscript identity for compatibility.");
            var suggestionTool = (IDictionary<string, object>)tools[2];
            var suggestionSchema = (IDictionary<string, object>)suggestionTool["inputSchema"];
            Assert(((object[])suggestionSchema["required"]).Contains("manuscript_id"),
                "Only reads may omit identity; suggestion discovery still requires the manuscript returned by current context.");
            var suggestionProperties = (IDictionary<string, object>)suggestionSchema["properties"];
            var rationale = (IDictionary<string, object>)suggestionProperties["rationale"];
            Assert(((string)suggestionTool["description"]).Contains("Applies eligible highlighted wording directly") &&
                ((string)suggestionTool["description"]).Contains("Discard permits only a complete deletion as native Word Track Changes") &&
                ((string)rationale["description"]).Contains("native Word comment"),
                "Tool discovery advertises direct wording, separately approved discard, and a concise native rationale comment.");
            var suggestionAnnotations = (IDictionary<string, object>)suggestionTool["annotations"];
            Assert(Object.Equals(suggestionAnnotations["readOnlyHint"], false) && Object.Equals(suggestionAnnotations["destructiveHint"], true) &&
                Object.Equals(suggestionAnnotations["idempotentHint"], true), "A direct document mutation remains explicitly writable, potentially destructive and replay-safe.");
            string reply = Post(address, "tools/call", new { name = "list_manuscripts", arguments = new { } });
            Assert(reply.Contains("\"manuscripts\":[]") && expectedTool == "list_manuscripts", "Real HTTP response carries native structured results.");
            int before = dispatches;
            string colorOnly = Post(address, "tools/call", new { name = "highlight_text", arguments = new { documentId = "a", marking = new {
                requestId = "mark", revision = "r", contextId = "c", ranges = new[] { new { from = 0, to = 4, quote = "word" } } } } });
            Assert(colorOnly.Contains("isError\":false") && expectedTool == "highlight_text" && dispatches == before + 1, "Only explicit color-only tool reaches native STA authority.");
            before = dispatches;
            foreach (object payload in new object[] {
                new { documentId = "a", action = "prompt", marking = new { requestId = "mark", revision = "r", contextId = "c", ranges = new[] { new { from = 0, to = 4, quote = "word" } } } },
                new { documentId = "a", marking = new { requestId = "mark", revision = "r", contextId = "c", accessMode = "direct", ranges = new[] { new { from = 0, to = 4, quote = "word" } } } },
                new { documentId = "a", marking = new { requestId = "mark", revision = "r", contextId = "c", ranges = new[] { new { from = 0, to = 4, quote = "word", prompt = "edit" } } } }
            }) Assert(Object.Equals(Result(Post(address, "tools/call", new { name = "highlight_text", arguments = payload }))["isError"], true), "Policy injection returns a tool denial at the closed HTTP boundary.");
            Assert(dispatches == before, "Policy fields cannot reach native Word dispatch.");
            string afterPolicyDenial = Post(address, "tools/call", new { name = "read_manuscript", arguments = new { manuscript_id = "a" } });
            Assert(Object.Equals(Result(afterPolicyDenial)["isError"], false) && expectedTool == "read_manuscript" && dispatches == before + 1,
                "A bad tool payload leaves the healthy connection usable for the next valid context read.");
            before = dispatches;
            Assert(Post(address, "tools/call", new { name = "save_document", arguments = new { } }).Contains("Unknown tool"), "Unknown tools denied.");
            Assert(dispatches == before, "Unknown tool cannot reach native authority.");
            int status;
            Post(address, "tools/call", new { name = "suggest_text", arguments = new { } }, out status, null, null, true);
            Assert(status == 202 && dispatches == before, "Notifications never invoke document tools.");
            Post(address, "ping", new { }, out status, "https://malicious.example", null, false);
            Assert(status == 403, "Browser Origin denied.");
            Assert(RawStatus(address, "Host: attacker.example\r\nContent-Type: application/json\r\nContent-Length: 2\r\n", "{}") == 403, "Forged proxy Host denied.");
            Assert(RawStatus(address, "Host: " + new Uri(address).Authority + "\r\nHost: evil.test\r\nContent-Type: application/json\r\nContent-Length: 2\r\n", "{}") == 400, "Ambiguous duplicate headers denied.");
            Assert(RawStatus(address, "Host: " + new Uri(address).Authority + "\r\nContent-Type: application/json\r\nContent-Length: 1048577\r\n", "") == 413, "Requests above the one-MiB encoded contract bound are rejected before reading body.");
            Assert(RawStatus(address, "Host: " + new Uri(address).Authority + "\r\nContent-Type: text/plain\r\nContent-Length: 2\r\n", "{}") == 415, "Non-JSON media type denied.");
            Assert(RawStatus(address, "Host: " + new Uri(address).Authority + "\r\nContent-Type: application/json\r\nTransfer-Encoding: chunked\r\nContent-Length: 2\r\n", "{}") == 415, "Ambiguous transfer encoding denied.");
            string ping = "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"ping\"}";
            Assert(RawStatus(address, "Host: " + new Uri(address).Authority + "\r\nAccept: application/json\r\nAccept: text/event-stream\r\nContent-Type: application/json\r\nContent-Length: " + ping.Length + "\r\n", ping) == 200, "Repeated Accept headers remain supported.");
            Post(address + "x", "ping", new { }, out status, null, null, false);
            Assert(status == 403, "Wrong capability denied.");
            Post(address, "ping", new { }, out status, null, "2099-01-01", false);
            Assert(status == 400, "Unsupported explicit MCP version denied.");
            failDispatch = true;
            string denied = Post(address, "tools/call", new { name = "read_manuscript", arguments = new { manuscript_id = "a" } });
            failDispatch = false;
            Assert(denied.Contains("isError\":true") && !denied.Contains("SECRET") && !denied.Contains("manuscript.docx"), "Unexpected native errors do not disclose private text or paths.");
            markerFailure = true;
            var markerDenied = Result(Post(address, "tools/call", new { name = "suggest_text", arguments = new { } }));
            markerFailure = false;
            var diagnostic = (Dictionary<string, object>)markerDenied["structuredContent"];
            Assert(Object.Equals(markerDenied["isError"], true) && Json.Serialize(diagnostic).Contains("marker_permission_unresolved") &&
                Json.Serialize(diagnostic).Contains("native_bookmark_missing") && Json.Serialize(diagnostic).Contains("Wrai_test"),
                "Marker refusals carry the exact local identity, reason and recovery action in structured diagnostics.");
            Assert((string)((Dictionary<string, object>)((object[])markerDenied["content"])[0])["text"] == Json.Serialize(diagnostic),
                "Text-only MCP clients receive the same complete actionable marker diagnostic.");
            bool locked = false;
            try { using (ConnectionService duplicate = Create(directory, fake)) { } }
            catch (InvalidOperationException) { locked = true; }
            Assert(locked, "One Word process exclusively owns its private profile.");
            int previousRevoked = revoked;
            service.Stop();
            Assert(!service.Status.LocalRunning && service.LocalAddress == "" && revoked == previousRevoked + 1, "Stop immediately clears listener and native grants.");
            service.StartLocal();
            Assert(new Uri(service.LocalAddress).AbsolutePath != new Uri(address).AbsolutePath, "Restart rotates bearer capability.");
            Assert(!service.Status.ExternallyVerified, "Local success never claims public verification.");
        }
        using (ConnectionService restart = Create(directory, fake))
            Assert(!restart.Status.LocalRunning && !restart.Status.TunnelRunning, "New Word session never restores transport or document sharing.");
        string config = File.ReadAllText(Path.Combine(directory, "connection.json"));
        Assert(!config.Contains("/t/") && !config.Contains("token"), "Persisted public setup contains no bearer capability or raw enrollment token.");
    }

    private static void TunnelLifecycle(string directory, string fake)
    {
        string home = Path.Combine(directory, "connection-home");
        using (ConnectionService service = Create(directory, fake))
        {
            service.Configure("fake-enrollment-token");
            Wait(delegate { return service.Status.Phase != "configuring" && !service.Status.Busy; }, "Fake enrollment completes.");
            Assert(service.Status.Configured && !service.Status.EnrollmentRequired, "Dedicated enrollment and assigned reservation confirmed.");
            string reservation = service.Status.Reservation;
            Assert(!File.ReadAllText(Path.Combine(directory, "connection.json")).Contains("fake-enrollment-token"), "Account enrollment token absent from application configuration.");
            service.Start();
            Wait(delegate { return service.Status.TunnelRunning || service.Status.Phase == "failed"; }, "Owned fake tunnel starts.");
            Assert(service.Status.TunnelRunning && service.PublicAddress.StartsWith("https://" + reservation + ".example.test/t/"), "Public address comes from exact controller frontend, retaining capability.");
            Assert(Post(service.LocalAddress, "ping", new { }).Contains("\"result\":{}"), "Private MCP remains available while tunnel runs.");
            service.Stop();
            Assert(!service.Status.LocalRunning && service.PublicAddress == "", "Stop conceals public address immediately.");
            Wait(delegate { return !service.Status.Busy && (service.Status.Phase == "stopped" || service.Status.Phase == "cleanup_required"); }, "Owned share cleanup finishes.");
            Assert(service.Status.Phase == "stopped" && File.ReadAllText(Path.Combine(home, "fake-shares.json")) == "[]", "Stop removes only exact owned share and independently verifies absence.");
            service.Start();
            Wait(delegate { return service.Status.TunnelRunning || service.Status.Phase == "failed"; }, "Reserved-name reconnect completes.");
            Assert(service.Status.TunnelRunning && service.Status.Reservation == reservation, "Reconnect reuses assigned reservation.");
            service.Stop();
            Wait(delegate { return service.Status.Phase == "stopped" && !service.Status.Busy; }, "Second share cleanup finishes.");
            Assert(File.ReadAllLines(Path.Combine(home, "commands.jsonl")).Count(line => line.StartsWith("[\"create\",\"name\"")) == 1, "Retries and reconnect never allocate another persistent name.");
            Assert(!File.ReadAllText(Path.Combine(home, "commands.jsonl")).Contains("fake-enrollment-token"), "Fake command evidence redacts enrollment credential.");
        }
    }

    private static void ForeignSharePreserved(string directory, string fake)
    {
        using (ConnectionService service = Create(directory, fake))
        {
            service.Configure("fake-enrollment-token");
            Wait(delegate { return service.Status.Configured && !service.Status.Busy; }, "Second independent fake enrollment completes.");
            string home = Path.Combine(directory, "connection-home");
            File.WriteAllText(Path.Combine(home, "scenario"), "foreign-owner");
            service.Start();
            Wait(delegate { return service.Status.Phase == "cleanup_required" && !service.Status.Busy; }, "Foreign share causes honest cleanup-required state.");
            Assert(!service.Status.LocalRunning && !service.Status.TunnelRunning && service.Status.CleanupPending, "Ownership mismatch revokes access and blocks reconnect.");
            Assert(File.ReadAllText(Path.Combine(home, "fake-shares.json")).Contains("foreign-environment"), "Foreign share preserved.");
            Assert(!File.ReadAllLines(Path.Combine(home, "commands.jsonl")).Any(line => line.StartsWith("[\"delete\"")), "Foreign identity never receives delete command.");
            bool rejected = false;
            try { service.Start(); } catch (InvalidOperationException) { rejected = true; }
            Assert(rejected, "Cleanup-required state refuses reconnect.");
            Assert(!service.Status.LocalRunning, "Blocked reconnect cannot leave a fresh local listener behind.");
        }
    }

    private static void ChildDeath(string directory, string fake)
    {
        using (ConnectionService service = Create(directory, fake))
        {
            service.Configure("fake-enrollment-token");
            Wait(delegate { return service.Status.Configured && !service.Status.Busy; }, "Death-case enrollment completes.");
            string home = Path.Combine(directory, "connection-home");
            File.WriteAllText(Path.Combine(home, "scenario"), "die");
            int previousRevoked = revoked;
            service.Start();
            Wait(delegate { return service.Status.Phase == "failed" && !service.Status.Busy; }, "Exited fake tunnel is observed and cleaned.");
            Assert(!service.Status.LocalRunning && !service.Status.TunnelRunning && revoked > previousRevoked, "Unexpected child death revokes both transport and native sharing.");
            Assert(File.ReadAllText(Path.Combine(home, "fake-shares.json")) == "[]", "Child-death rollback retires its exact controller share.");
        }
    }

    private static void FailedCleanupRetry(string directory, string fake)
    {
        using (ConnectionService service = Create(directory, fake))
        {
            service.Configure("fake-enrollment-token");
            Wait(delegate { return service.Status.Configured && !service.Status.Busy; }, "Cleanup-case enrollment completes.");
            service.Start();
            Wait(delegate { return service.Status.TunnelRunning; }, "Cleanup-case fake tunnel starts.");
            string home = Path.Combine(directory, "connection-home");
            File.WriteAllText(Path.Combine(home, "scenario"), "cleanup-fail");
            service.Stop();
            Wait(delegate { return service.Status.Phase == "cleanup_required" && !service.Status.Busy; }, "Unconfirmed delete is reported.");
            Assert(!service.Status.LocalRunning && service.Status.CleanupPending && service.PublicAddress == "", "Cleanup failure keeps ownership privately and exposes no capability.");
            File.Delete(Path.Combine(home, "scenario"));
            service.Stop();
            Wait(delegate { return service.Status.Phase == "stopped" && !service.Status.Busy; }, "Explicit Stop retries exact-share cleanup.");
            Assert(!service.Status.CleanupPending && File.ReadAllText(Path.Combine(home, "fake-shares.json")) == "[]", "Explicit cleanup retry clears pinned ownership only after verified absence.");
        }
    }

    private static void InvalidStatePreserved(string directory, string fake)
    {
        using (ConnectionService service = Create(directory, fake)) { }
        string path = Path.Combine(directory, "connection.json");
        const string invalid = "{\"version\":99,\"private-existing-value\":\"preserve\"}";
        File.WriteAllText(path, invalid);
        bool rejected = false;
        try { using (ConnectionService service = Create(directory, fake)) { } }
        catch (InvalidOperationException) { rejected = true; }
        Assert(rejected && File.ReadAllText(path) == invalid, "Corrupt or unknown persisted setup is refused and preserved byte-for-byte.");
    }

    private static void QueuedRevocation(string directory, string fake)
    {
        using (ManualResetEvent queued = new ManualResetEvent(false))
        using (ManualResetEvent release = new ManualResetEvent(false))
        {
            int executed = 0;
            using (ConnectionService service = new ConnectionService(delegate(string name, IDictionary<string, object> values) { executed++; return new { }; },
                delegate { }, delegate(Func<object> callback) { queued.Set(); if (!release.WaitOne(10000)) throw new TimeoutException(); return Marshal(callback); }, directory, fake))
            {
                service.StartLocal();
                string address = service.LocalAddress;
                Exception failure = null;
                Thread request = new Thread(delegate()
                {
                    try { Post(address, "tools/call", new { name = "suggest_text", arguments = new { } }); }
                    catch (Exception error) { failure = error; }
                });
                request.Start();
                Assert(queued.WaitOne(5000), "Remote request reaches pending Word marshal queue.");
                service.Stop();
                release.Set();
                Assert(request.Join(5000) && failure == null && executed == 0, "Stop cancels already queued remote work before native authority executes.");
            }
        }
    }

    private static void FailedRevocation(string directory, string fake)
    {
        bool fail = true;
        using (ConnectionService service = new ConnectionService(delegate(string name, IDictionary<string, object> values) { return new { }; },
            delegate { if (fail) throw new InvalidOperationException("Synthetic revocation failure"); }, Marshal, directory, fake))
        {
            service.StartLocal();
            bool failed = false;
            try { service.Stop(); } catch (InvalidOperationException) { failed = true; }
            Assert(failed && service.Status.Phase == "revocation_required" && !service.Status.LocalRunning, "Failed native revocation closes listener and reports unresolved authority.");
            bool blocked = false;
            try { service.StartLocal(); } catch (InvalidOperationException) { blocked = true; }
            Assert(blocked && !service.Status.LocalRunning, "Unresolved document revocation blocks fresh bearer capabilities.");
            fail = false;
            service.Stop(); service.StartLocal();
            Assert(service.Status.LocalRunning, "Explicit successful revoke permits a fresh local connection.");
        }
    }

    private static void StaShutdown(string directory, string fake)
    {
        ConnectionService service = Create(directory, fake);
        try
        {
            service.Configure("fake-enrollment-token");
            Wait(delegate { return service.Status.Configured && !service.Status.Busy; }, "STA-shutdown enrollment completes.");
            service.Start();
            Wait(delegate { return service.Status.TunnelRunning; }, "STA-shutdown tunnel starts.");
            Marshal(delegate { service.Dispose(); return null; });
            string home = Path.Combine(directory, "connection-home");
            Assert(File.ReadAllText(Path.Combine(home, "fake-shares.json")) == "[]", "Disposal on the Word STA completes exact cleanup without waiting for its own blocked marshal queue.");
            Assert(!File.Exists(Path.Combine(directory, "share-cleanup.bin")), "Successful STA shutdown leaves no false persistent cleanup barrier.");
        }
        finally { service.Dispose(); }
    }

    private static void InterruptedCleanupRecovery(string directory, string fake)
    {
        ConnectionService service = Create(directory, fake);
        string home = Path.Combine(directory, "connection-home");
        try
        {
            service.Configure("fake-enrollment-token");
            Wait(delegate { return service.Status.Configured && !service.Status.Busy; }, "Interrupted-cleanup enrollment completes.");
            service.Start();
            Wait(delegate { return service.Status.TunnelRunning; }, "Interrupted-cleanup tunnel starts.");
            string barrier = Path.Combine(directory, "share-cleanup.bin");
            Assert(File.Exists(barrier) && !Encoding.UTF8.GetString(File.ReadAllBytes(barrier)).Contains("fake-word-share"), "Exact ownership is durably protected with per-user encryption before a running share is exposed.");
            File.WriteAllText(Path.Combine(home, "scenario"), "slow-cleanup");
            Stopwatch shutdown = Stopwatch.StartNew();
            Marshal(delegate { service.Dispose(); return null; });
            Assert(shutdown.ElapsedMilliseconds < 6000 && File.Exists(barrier), "Slow controller cleanup cannot hang Word; encrypted ownership barrier survives bounded shutdown.");
            File.Delete(Path.Combine(home, "scenario"));
            using (ConnectionService restart = Create(directory, fake))
            {
                Assert(restart.Status.CleanupPending && restart.Status.Phase == "cleanup_required" && !restart.Status.LocalRunning, "New Word session restores cleanup barrier without exposing transport or shares.");
                bool blocked = false;
                try { restart.Start(); } catch (InvalidOperationException) { blocked = true; }
                Assert(blocked && !restart.Status.LocalRunning, "Persisted cleanup uncertainty blocks all fresh tunnel capabilities.");
                restart.Stop();
                Wait(delegate { return restart.Status.Phase == "stopped" && !restart.Status.Busy; }, "Restart Stop independently retires exactly pinned share.");
                Assert(!File.Exists(barrier) && File.ReadAllText(Path.Combine(home, "fake-shares.json")) == "[]", "Restart clears barrier only after exact token, environment, target and frontend revalidation plus verified absence.");
            }
        }
        finally { service.Dispose(); }
    }

    private static void PendingWorkerRevocation(string directory, string fake)
    {
        using (ManualResetEvent queued = new ManualResetEvent(false))
        using (ManualResetEvent release = new ManualResetEvent(false))
        using (ManualResetEvent finished = new ManualResetEvent(false))
        {
            int nativeRevokes = 0;
            using (ConnectionService service = new ConnectionService(delegate(string name, IDictionary<string, object> values) { return new { }; },
                delegate { nativeRevokes++; }, delegate(Func<object> callback)
                {
                    queued.Set();
                    try { if (!release.WaitOne(10000)) throw new TimeoutException(); return Marshal(callback); }
                    finally { finished.Set(); }
                }, directory, fake))
            {
                service.Configure("fake-enrollment-token");
                Wait(delegate { return service.Status.Configured && !service.Status.Busy; }, "Worker-revoke enrollment completes.");
                File.WriteAllText(Path.Combine(directory, "connection-home", "scenario"), "die");
                service.Start();
                Assert(queued.WaitOne(5000), "Background failure queues native revocation.");
                Wait(delegate { return !service.Status.Busy; }, "Controller cleanup proceeds while the Word STA is unavailable.");
                bool blocked = false;
                try { service.StartLocal(); } catch (InvalidOperationException) { blocked = true; }
                Assert(blocked && !service.Status.LocalRunning && service.Status.Phase == "revocation_required", "Pending STA revoke blocks fresh capability even after controller cleanup completed.");
                service.Stop();
                service.StartLocal();
                int settledRevokes = nativeRevokes;
                release.Set();
                Assert(finished.WaitOne(5000) && nativeRevokes == settledRevokes && service.Status.LocalRunning, "A stale queued revoke cannot revoke a fresh session after explicit Stop resolved it.");
            }
        }
    }

    private static void DisposeRevocationFailure(string directory, string fake)
    {
        ConnectionService service = new ConnectionService(delegate(string name, IDictionary<string, object> values) { return new { }; },
            delegate { throw new InvalidOperationException("Synthetic shutdown revoke failure"); }, Marshal, directory, fake);
        service.StartLocal();
        service.Dispose();
        Assert(service.LocalAddress == "" && service.Status.Phase == "shutdown_unverified" && service.Status.Error.Length != 0,
            "Native revoke failure cannot throw into Word shutdown or leave an exposed listener; unresolved shutdown remains visible in status.");
        service.Dispose();
    }

    private static void ActiveDocumentFirstRead(string directory, string fake)
    {
        int nativeCalls = 0;
        using (var service = new ConnectionService(delegate(string name, IDictionary<string, object> values) {
            Assert(Thread.CurrentThread.ManagedThreadId == actorThread && name == "read_manuscript" && values.Count == 1 && values.ContainsKey(WordDocumentPicker.SessionKey),
                "An empty first read reaches Word STA without discovery or a manufactured manuscript identity.");
            nativeCalls++;
            return new { manuscript_id = "active-document", text = "Current writing context.", complete = true,
                revision = "active-revision", context_id = "active-witness" };
        }, delegate { }, Marshal, directory, fake))
        {
            service.StartLocal();
            var reply = Result(Post(service.LocalAddress, "tools/call", new { name = "read_manuscript", arguments = new { } }));
            var current = (IDictionary<string, object>)reply["structuredContent"];
            Assert(nativeCalls == 1 && Object.Equals(reply["isError"], false) &&
                Object.Equals(current["manuscript_id"], "active-document") && Object.Equals(current["complete"], true) &&
                Object.Equals(current["revision"], "active-revision") && Object.Equals(current["context_id"], "active-witness"),
                "One first-read round trip carries the complete active context, identity and edit witness.");
        }
    }

    private static void ActiveDocumentRecovery(string directory, string fake)
    {
        int nativeCalls = 0;
        using (var service = new ConnectionService(delegate(string name, IDictionary<string, object> values) {
            nativeCalls++;
            if (Object.Equals(values["manuscript_id"], "old-document"))
                throw new McpOperationException("The chosen Word document closed.", "document_closed", new {
                    manuscript_id = "old-document", request_applied = false,
                    instruction = "Choose another document; do not follow the active window." });
            return new { manuscript_id = "new-document", text = "Current writing context.", complete = true, revision = "current-revision", context_id = "current-witness" };
        }, delegate { }, Marshal, directory, fake))
        {
            service.StartLocal();
            var denied = Result(Post(service.LocalAddress, "tools/call", new { name = "read_manuscript", arguments = new { manuscript_id = "old-document" } }));
            var diagnostic = (IDictionary<string, object>)((IDictionary<string, object>)denied["structuredContent"])["error"];
            var details = (IDictionary<string, object>)diagnostic["details"];
            Assert(Object.Equals(denied["isError"], true) && Object.Equals(diagnostic["code"], "document_closed") &&
                Object.Equals(details["manuscript_id"], "old-document") && Object.Equals(details["request_applied"], false) && !details.ContainsKey("active_document"),
                "A closed-target diagnostic retains its chosen identity and cannot instruct automatic active-window retargeting.");
            var current = Result(Post(service.LocalAddress, "tools/call", new { name = "read_manuscript", arguments = new { manuscript_id = "new-document" } }));
            Assert(nativeCalls == 2 && Object.Equals(current["isError"], false) &&
                Object.Equals(((IDictionary<string, object>)current["structuredContent"])["context_id"], "current-witness"),
                "An explicit new document read can continue after closure without a sharing operation.");
        }
    }

    private static void ConversationRouting(string directory, string fake)
    {
        var sessions = new List<string>();
        using (var service = new ConnectionService(delegate(string name, IDictionary<string, object> values) {
            object session; sessions.Add(values.TryGetValue(WordDocumentPicker.SessionKey, out session) ? (string)session : null);
            return new { selection_required = true, manuscripts = new[] { new { manuscript_id = "a", title = "Draft.docx" } } };
        }, delegate { }, Marshal, directory, fake))
        {
            service.StartLocal(); string address = service.LocalAddress;
            foreach (string conversation in new[] { "chat-a", "chat-b", "chat-a" })
            {
                var result = Result(Post(address, "tools/call", new { name = "read_manuscript", arguments = new { },
                    _meta = new Dictionary<string, object> { { "openai/session", conversation } } }));
                Assert(Object.Equals(result["isError"], false) && Object.Equals(((IDictionary<string, object>)result["structuredContent"])["selection_required"], true),
                    "A choice-only response carries no manuscript read or edit witness.");
            }
            Assert(sessions.Count == 3 && sessions[0] == sessions[2] && sessions[0] != sessions[1] && sessions.All(value => value.StartsWith("mcp_")),
                "ChatGPT conversation metadata routes separate opaque native choices within the initialized MCP session.");
            int before = sessions.Count;
            foreach (object malformed in new object[] { "conversation", new object[] { "conversation" } })
                Assert(Object.Equals(Result(Post(address, "tools/call", new { name = "read_manuscript", arguments = new { }, _meta = malformed }))["isError"], true) &&
                    sessions.Count == before, "Malformed non-object metadata cannot silently collapse into a session's default document choice.");
            Assert(Object.Equals(Result(Post(address, "tools/call", new { name = "read_manuscript", arguments = new { }, _meta = new { progressToken = "valid-unknown-context" } }))["isError"], false) &&
                sessions.Count == before + 1, "Unknown valid object metadata remains accepted without granting authority.");
            before = sessions.Count;
            Assert(Object.Equals(Result(Post(address, "tools/call", new { name = "read_manuscript", arguments = new Dictionary<string, object> {
                { WordDocumentPicker.SessionKey, "forged" } } }))["isError"], true) && sessions.Count == before,
                "Internal conversation routing cannot be injected as a tool argument.");
            var resources = Result(Post(address, "resources/list", new { }));
            var listed = (IDictionary<string, object>)((object[])resources["resources"])[0];
            Assert(Object.Equals(listed["uri"], WordDocumentPicker.Uri) && ((object[])resources["resources"]).Length == 1,
                "Only the fixed embedded document picker is listed as a resource.");
            var resource = Result(Post(address, "resources/read", new { uri = WordDocumentPicker.Uri }));
            var html = (IDictionary<string, object>)((object[])resource["contents"])[0];
            Assert(Object.Equals(html["mimeType"], "text/html;profile=mcp-app") && ((string)html["text"]).Contains("Choose a Word document") && sessions.Count == before,
                "The real endpoint serves the embedded picker without dispatching Word document content.");
            var metadataResource = Result(Post(address, "resources/read", new { uri = WordDocumentPicker.Uri,
                _meta = new Dictionary<string, object> { { "openai/session", "widget-chat" }, { "progressToken", "widget-resource" } } }));
            Assert(Object.Equals(((IDictionary<string, object>)((object[])metadataResource["contents"])[0])["text"], html["text"]) && sessions.Count == before,
                "Standard resource request metadata does not block the fixed picker or dispatch Word authority.");
            string unexpectedResourceArgument = Post(address, "resources/read", new { uri = WordDocumentPicker.Uri, path = "C:/private.docx" });
            Assert(unexpectedResourceArgument.Contains("Unknown Word UI resource") && sessions.Count == before,
                "Resource metadata compatibility does not admit arbitrary resource arguments.");
            string forbidden = Post(address, "resources/read", new { uri = "file:///C:/private.docx" });
            Assert(forbidden.Contains("Unknown Word UI resource") && sessions.Count == before, "Resource dispatch refuses arbitrary paths and cannot invoke document authority.");
        }
    }

    private static void SuggestionCurrentContext(string directory, string fake)
    {
        string nativeRationale = null;
        string originalRequest = null;
        int nativeCalls = 0, nativeEffects = 0;
        object context = new { manuscript_id = "writing-test", revision = "next-revision", text = "A clearer paragraph.\nProtected text.",
            from = 0, to = 36, total = 36, complete = true, next_offset = (int?)null, context_id = "next-context", main_body_length = 36,
            highlights = new[] { new { from = 21, to = 36, protectedId = "protected-original" } } };
        using (var service = new ConnectionService(delegate(string name, IDictionary<string, object> values) {
            Assert(Thread.CurrentThread.ManagedThreadId == actorThread && name == "suggest_text", "The paragraph edit alone reaches native STA authority.");
            nativeCalls++;
            nativeRationale = (string)values["rationale"];
            string fingerprint = Json.Serialize(values);
            bool replayed = originalRequest != null;
            if (replayed && originalRequest != fingerprint) throw new McpOperationException("Changed replay denied.");
            if (!replayed) { originalRequest = fingerprint; nativeEffects++; }
            return new { id = values["request_id"], status = "applied", applied = true, requiresApproval = false, reviewMode = "direct", replayed = replayed,
                current_context = replayed ? null : context };
        }, delegate { }, Marshal, directory, fake))
        {
            service.StartLocal();
            object request = new { name = "suggest_text", arguments = new {
                manuscript_id = "writing-test", revision = "previous-revision", context_id = "previous-context", from = 0, to = 21,
                quote = "An unclear paragraph.", replacement = "A clearer paragraph.", request_id = "first-paragraph", rationale = "Clarify the wording." } };
            var reply = Result(Post(service.LocalAddress, "tools/call", request));
            var receipt = (IDictionary<string, object>)reply["structuredContent"];
            var current = (IDictionary<string, object>)receipt["current_context"];
            Assert(nativeCalls == 1 && nativeEffects == 1 && Object.Equals(receipt["status"], "applied") && Object.Equals(receipt["applied"], true) &&
                Object.Equals(receipt["requiresApproval"], false) && Object.Equals(receipt["reviewMode"], "direct"),
                "A single direct-edit round trip preserves the applied native receipt without an approval step.");
            Assert(nativeRationale == "Clarify the wording." && Object.Equals(current["revision"], "next-revision") &&
                Object.Equals(current["context_id"], "next-context") && Object.Equals(current["complete"], true),
                "The concise native rationale and current witness cross the endpoint unchanged.");
            object[] content = (object[])reply["content"];
            var textReceipt = (IDictionary<string, object>)Json.DeserializeObject((string)((IDictionary<string, object>)content[0])["text"]);
            Assert(Json.Serialize(textReceipt["current_context"]) == Json.Serialize(current) &&
                Object.Equals(((IDictionary<string, object>)((object[])current["highlights"])[0])["protectedId"], "protected-original"),
                "Both MCP result formats preserve the complete successor context and protected identity without an extra read.");
            var replay = (IDictionary<string, object>)Result(Post(service.LocalAddress, "tools/call", request))["structuredContent"];
            Assert(nativeCalls == 2 && nativeEffects == 1 && Object.Equals(replay["status"], "applied") && Object.Equals(replay["replayed"], true) &&
                Object.Equals(replay["applied"], true) && Object.Equals(replay["requiresApproval"], false) && replay["current_context"] == null,
                "An exact direct replay preserves the native receipt without another effect, comment or redundant context read.");
        }
        foreach (string status in new[] { "pending", "tracked" })
            using (var service = new ConnectionService(delegate(string name, IDictionary<string, object> values) {
                return new { id = values["request_id"], status = status, applied = false, replayed = true, requiresApproval = true,
                    reviewMode = status == "tracked" ? "word-track-changes" : "direct", current_context = (object)null };
            }, delegate { }, Marshal, Path.Combine(directory, status), fake))
            {
                service.StartLocal();
                var legacy = (IDictionary<string, object>)Result(Post(service.LocalAddress, "tools/call", new { name = "suggest_text", arguments = new {
                    manuscript_id = "writing-test", revision = "previous-revision", context_id = "previous-context", from = 0, to = 21,
                    quote = "An unclear paragraph.", replacement = "A clearer paragraph.", request_id = "legacy-paragraph" } }))["structuredContent"];
                Assert(Object.Equals(legacy["status"], status) && Object.Equals(legacy["replayed"], true) &&
                    Object.Equals(legacy["applied"], false) && Object.Equals(legacy["requiresApproval"], true),
                    "Transport cannot upgrade an existing " + status + " receipt into a new direct edit.");
            }
    }

    private static void LegalEscapedPayloads(string directory, string fake)
    {
        string prompt = new string('\\', 4000);
        object[] highlights = Enumerable.Range(0, 500).Select(index => (object)new {
            id = "Wrai_" + index, ruleId = "shorten", from = 0, to = 1, prompt = prompt, stale = false }).ToArray();
        object context = new { manuscript_id = "payload-test", revision = new string('a', 64), text = "a", from = 0, to = 1,
            total = 1, complete = true, context_id = new string('b', 32), highlights = highlights };
        Assert(Encoding.UTF8.GetByteCount(Json.Serialize(context)) < 4194304, "Legal compact context metadata fits the admitted four-MiB persisted aggregate.");
        bool oversizedResponse = false;
        string nativeQuote = null, nativeReplacement = null;
        using (ConnectionService service = new ConnectionService(delegate(string name, IDictionary<string, object> values)
        {
            Assert(Thread.CurrentThread.ManagedThreadId == actorThread, "Large encoded requests retain Word STA dispatch.");
            if (name == "read_manuscript") return oversizedResponse ? (object)new { text = new string('x', 8388609) } : context;
            nativeQuote = (string)values["quote"]; nativeReplacement = (string)values["replacement"];
            return new { request_id = values["request_id"], status = "tracked", requiresApproval = true };
        }, delegate { }, Marshal, directory, fake))
        {
            service.StartLocal();
            Stopwatch responseTime = Stopwatch.StartNew();
            string response = Post(service.LocalAddress, "tools/call", new { name = "read_manuscript", arguments = new { manuscript_id = "payload-test" } });
            responseTime.Stop();
            int encoded = Encoding.UTF8.GetByteCount(response);
            Console.WriteLine("Synthetic escaped-context transport: result=" + Encoding.UTF8.GetByteCount(Json.Serialize(context)) +
                " bytes, MCP response=" + encoded + " bytes, loopback=" + responseTime.ElapsedMilliseconds + " ms; no Word or model latency included.");
            Assert(encoded > 4194304 && encoded <= 16777216, "Escaped duplicated MCP context exceeds the old cap and fits the explicit sixteen-MiB response bound.");
            IDictionary<string, object> result = Result(response);
            IDictionary<string, object> structured = (IDictionary<string, object>)result["structuredContent"];
            object[] returned = (object[])structured["highlights"];
            Assert(returned.Length == 500 && returned.All(mark => Object.Equals(((IDictionary<string, object>)mark)["prompt"], prompt)), "All 500 legal escaped prompts arrive intact in structured content.");
            object[] content = (object[])result["content"];
            IDictionary<string, object> textResult = (IDictionary<string, object>)Json.DeserializeObject((string)((IDictionary<string, object>)content[0])["text"]);
            Assert(((object[])textResult["highlights"]).Length == 500 && Object.Equals(textResult["context_id"], new string('b', 32)), "MCP text content contains the same complete metadata and context witness.");
            string quote = new string('"', 32768), replacement = new string('\\', 32768);
            object proposal = new { name = "suggest_text", arguments = new { manuscript_id = "payload-test", revision = new string('a', 64),
                context_id = new string('b', 32), from = 0, to = 32768, quote = quote, replacement = replacement, request_id = "escaped-pair" } };
            int requestBytes = Encoding.UTF8.GetByteCount(Json.Serialize(new { jsonrpc = "2.0", id = 1, method = "tools/call", @params = proposal }));
            Assert(requestBytes > 131072 && requestBytes <= 1048576, "Legal maximum UTF-16 quote and replacement require more than the old encoded request cap.");
            IDictionary<string, object> submitted = Result(Post(service.LocalAddress, "tools/call", proposal));
            IDictionary<string, object> recorded = (IDictionary<string, object>)submitted["structuredContent"];
            Assert(nativeQuote == quote && nativeReplacement == replacement && Object.Equals(recorded["status"], "tracked") &&
                Object.Equals(recorded["requiresApproval"], true), "Maximum escaped quote and replacement reach native authority unchanged and remain a tracked suggestion for local review.");
            oversizedResponse = true;
            int status;
            string rejected = Post(service.LocalAddress, "tools/call", new { name = "read_manuscript", arguments = new { manuscript_id = "payload-test" } }, out status, null, null, false);
            Assert(status == 413 && rejected.Length == 0, "Responses above the sixteen-MiB encoded contract bound remain rejected without a misleading partial result.");
        }
    }

    private static string Post(string url, string method, object parameters)
    { int status; string response = Post(url, method, parameters, out status, null, null, false); Assert(status == 200, "MCP request returns valid HTTP response."); return response; }

    private static string Post(string url, string method, object parameters, out int status, string origin, string version, bool notification)
    {
        string session = null;
        if (method != "initialize" && !Sessions.TryGetValue(url, out session))
        {
            int initializeStatus;
            Post(url, "initialize", new { protocolVersion = "2025-06-18" }, out initializeStatus, null, null, false);
            Sessions.TryGetValue(url, out session);
        }
        string issued;
        string result = PostSession(url, method, parameters, out status, session, out issued, origin, version, notification);
        if (issued != null) Sessions[url] = issued;
        return result;
    }

    private static string PostSession(string url, string method, object parameters, out int status, string session, out string issued,
        string origin = null, string version = null, bool notification = false)
    {
        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
        request.Proxy = null; request.Method = "POST"; request.ContentType = "application/json"; request.Timeout = 10000;
        request.ServicePoint.Expect100Continue = false;
        request.ServicePoint.ConnectionLimit = 16;
        request.KeepAlive = false;
        if (session != null) request.Headers["Mcp-Session-Id"] = session;
        if (origin != null) request.Headers["Origin"] = origin;
        if (version != null) request.Headers["MCP-Protocol-Version"] = version;
        object message = notification ? (object)new { jsonrpc = "2.0", method = method, @params = parameters } : new { jsonrpc = "2.0", id = 1, method = method, @params = parameters };
        byte[] body = Encoding.UTF8.GetBytes(Json.Serialize(message));
        request.ContentLength = body.Length;
        using (Stream stream = request.GetRequestStream()) stream.Write(body, 0, body.Length);
        HttpWebResponse response;
        try { response = (HttpWebResponse)request.GetResponse(); }
        catch (WebException error) { response = (HttpWebResponse)error.Response; if (response == null) throw; }
        using (response)
        using (StreamReader source = new StreamReader(response.GetResponseStream())) { status = (int)response.StatusCode; issued = response.Headers["Mcp-Session-Id"]; return source.ReadToEnd(); }
    }

    private static IDictionary<string, object> Result(string json)
    { return (IDictionary<string, object>)((IDictionary<string, object>)Json.DeserializeObject(json))["result"]; }

    private static int RawStatus(string url, string headers, string body)
    {
        string session;
        if (Sessions.TryGetValue(url, out session) && !headers.Contains("Mcp-Session-Id:")) headers += "Mcp-Session-Id: " + session + "\r\n";
        Uri parsed = new Uri(url);
        using (TcpClient socket = new TcpClient())
        {
            socket.Connect(parsed.Host, parsed.Port); socket.ReceiveTimeout = 10000;
            using (NetworkStream stream = socket.GetStream())
            {
                byte[] request = Encoding.ASCII.GetBytes("POST " + parsed.AbsolutePath + " HTTP/1.1\r\n" + headers + "\r\n" + body);
                stream.Write(request, 0, request.Length);
                using (StreamReader source = new StreamReader(stream)) return Int32.Parse(source.ReadLine().Split(' ')[1]);
            }
        }
    }

    private static void Wait(Func<bool> condition, string label)
    {
        Stopwatch timeout = Stopwatch.StartNew();
        using (ManualResetEvent tick = new ManualResetEvent(false))
            while (!condition()) { if (timeout.ElapsedMilliseconds > 15000) throw new Exception("Timeout: " + label); tick.WaitOne(30); }
    }

    private static void Assert(bool condition, string label)
    { if (!condition) throw new Exception("FAILED: " + label); Interlocked.Increment(ref checks); }
}
