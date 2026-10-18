using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using UnderTree.WordAddin;

internal static partial class NativeSuggestionWorkflowSmoke
{
    private delegate bool Enumerate(IntPtr window, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumWindows(Enumerate callback, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr window, Enumerate callback, IntPtr state);
    [DllImport("user32.dll")] private static extern int GetClassName(IntPtr window, StringBuilder name, int length);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object value);
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 16777216 };
    private static readonly List<string> checks = new List<string>();
    private static readonly List<string> documentIds = new List<string>();
    private static dynamic application, document;
    private static WordEditor editor;
    private static Control dispatcher;
    private static string output, documentId, savedPath, savedHash;
    private const string Original = "A rough paragraph repeats the same point with unnecessary wording.";
    private const string Revised = "A clear paragraph makes its point once.";
    private const string Outside = "Unmarked tail stays exactly as written.";
    private const string ReopenedChoiceSource = "A reopened document has another rough paragraph.";
    private static readonly List<object> timings = new List<object>();
    private static readonly List<object> selectionObservations = new List<object>();

    [STAThread]
    public static int Main(string[] args)
    {
        Process owned = null;
        WordTestAutomation filter = null;
        try
        {
            filter = new WordTestAutomation();
            output = Path.GetFullPath(args[1]); Directory.CreateDirectory(output);
            string bootstrap = Path.Combine(output, "suggestion-bootstrap.rtf");
            File.WriteAllText(bootstrap, @"{\rtf1\ansi Disposable suggestion bootstrap.\par}");
            owned = Process.Start(new ProcessStartInfo(args[0], "/a /x /q \"" + bootstrap + "\"") { UseShellExecute = false, WindowStyle = ProcessWindowStyle.Hidden });
            application = AttachOwned(owned); application.Visible = false; application.DisplayAlerts = 0;
            Check((int)application.Documents.Count == 1 && (string)application.ActiveDocument.FullName == bootstrap,
                "The owned isolated Word process opened only the disposable bootstrap.");
            application.ActiveDocument.Close(0);
            dispatcher = new Control(); IntPtr dispatchHandle = dispatcher.Handle;
            int nativeThread = Thread.CurrentThread.ManagedThreadId;
            ConfigureNativeDispatcher(nativeThread);
            NativeMcpClient.Trace = TraceNative;
            Func<Func<object>, object> marshal = MarshalNative;
            editor = new WordEditor((object)application, marshal, Path.Combine(output, "profile"), Path.Combine(output, "unused-zrok.exe"));
            Local("connection_local");
            string scenario = args.Length < 3 ? "all" : args[2];
            if (scenario == "background-save-only")
            {
                BackgroundTimerSave(); WriteResult("PASS", null);
                Console.WriteLine("PASS " + checks.Count + " native background-timer save assertions.");
                return 0;
            }
            if (scenario == "direct-comments-only") DirectCommentAnchors();
            if (scenario == "all" || scenario == "concurrency-only") { ConcurrentMcpClients(); NativeFinalizerRecovery(); }
            if (scenario != "protected-only" && scenario != "concurrency-only" && scenario != "direct-comments-only")
            {
                MultipleOpenDocumentsRetainConversationChoice();
                if (scenario != "document-choice-only")
                {
                    SuccessorContextAndDocumentRecovery();
                    StaleRevisionReturnsCurrentContext();
                    LargeContextRequiresSequentialReading();
                    PromptColorsEditDirectly();
                    DirectCommentAnchors();
                    LegacyPendingRemainsPending();
                    FollowUpRejectAndDeletion();
                }
            }
            if (scenario != "document-choice-only" && scenario != "concurrency-only" && scenario != "direct-comments-only") ProtectedUnitsAndSave();
            WriteResult("PASS", null);
            Console.WriteLine("PASS " + checks.Count + " native suggestion-workflow assertions.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            if (document != null)
            {
                try { File.WriteAllText(Path.Combine(output, "failure-body.xml"), (string)document.Content.WordOpenXML); }
                catch (Exception diagnostic) { Console.Error.WriteLine(diagnostic.Message); }
            }
            WriteResult("FAIL", error.ToString());
            return 1;
        }
        finally
        {
            if (editor != null) editor.Dispose();
            if (dispatcher != null) dispatcher.Dispose();
            if (application != null) application.Quit(0);
            else if (owned != null && !owned.HasExited) owned.Kill();
            foreach (string id in documentIds)
            {
                string receipt = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wrAIterWord", "reviews", id + ".bin");
                if (File.Exists(receipt)) File.Delete(receipt);
            }
            if (filter != null) filter.Dispose();
        }
    }

    private static Dictionary<string, object> TimedSuggestion(Dictionary<string, object> request, string session = null)
    {
        var watch = Stopwatch.StartNew();
        Dictionary<string, object> response = session == null ? Review("suggest_text", request) : SessionReview(session, "suggest_text", request);
        watch.Stop();
        timings.Add(new { requestId = request["request_id"], requestToReturnMs = watch.Elapsed.TotalMilliseconds,
            nativeRevisions = (int)document.Revisions.Count, nativeComments = (int)document.Comments.Count });
        Console.WriteLine("TIMING: " + request["request_id"] + " " + watch.Elapsed.TotalMilliseconds.ToString("F1") + " ms request-to-return.");
        return response;
    }

    private sealed class ChoiceFixture
    {
        internal object Document;
        internal string Id, Title, Path, Source;
        internal bool Closed;
    }

    private static ChoiceFixture AddChoiceFixture(string relativePath, string source, string alternativePath = null)
    {
        document = application.Documents.Add(Visible: false); document.Activate(); document.TrackRevisions = false;
        document.Content.Text = source + "\r" + Outside + "\r" + ReopenedChoiceSource + "\r";
        string path = Path.Combine(output, "document-choice", relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(path));
        try { document.SaveAs2(path, 16); }
        catch (COMException error)
        {
            if (alternativePath == null) throw;
            selectionObservations.Add(new { scenario = "same-basename", status = "HOST_REFUSED", requestedPath = path, nativeError = error.Message, hresult = error.ErrorCode });
            Console.WriteLine("OBSERVATION: Word refused simultaneous same-basename documents: " + error.Message);
            path = Path.Combine(output, "document-choice", alternativePath);
            document.SaveAs2(path, 16);
        }
        documentId = Text(Map(Local("state")["document"]), "id"); documentIds.Add(documentId);
        Mark(source, "snappier"); Mark(Outside, "snappier"); Mark(ReopenedChoiceSource, "snappier"); document.Save();
        return new ChoiceFixture { Document = (object)document, Id = documentId, Title = (string)document.Name, Path = path, Source = source };
    }

    private static void UseChoiceFixture(ChoiceFixture fixture)
    {
        document = fixture.Document; documentId = fixture.Id; savedPath = fixture.Path; savedHash = FileHash(savedPath);
    }

    private static string ChoiceNativeIdentity(ChoiceFixture fixture)
    { return new WordNativeEditState().NativeIdentity(fixture.Document) + ":" + WordReviewStore.Fingerprint(fixture.Document); }

    private static void CloseChoiceFixture(ChoiceFixture fixture)
    {
        if (fixture.Closed) return;
        ((dynamic)fixture.Document).Close(0); fixture.Closed = true;
    }

    private static Dictionary<string, object> SessionReview(string session, string tool, Dictionary<string, object> args)
    {
        var routed = new Dictionary<string, object>(args); routed[WordEditor.ReviewSessionKey] = session;
        return Review(tool, routed);
    }

    private static McpOperationException SessionRefusal(string session, string tool, Dictionary<string, object> args)
    {
        try { if (session == null) Review(tool, args); else SessionReview(session, tool, args); }
        catch (McpOperationException error) { return error; }
        return null;
    }

    private static McpOperationException ExpectSessionError(string session, string tool, Dictionary<string, object> args, string code)
    {
        McpOperationException refusal = SessionRefusal(session, tool, args);
        Check(refusal != null && refusal.Code == code, "The " + tool + " request returns the required " + code + " diagnostic.");
        return refusal;
    }

    private static void MultipleOpenDocumentsRetainConversationChoice()
    {
        object previous = document; string previousId = documentId, previousPath = savedPath, previousHash = savedHash;
        var fixtures = new List<ChoiceFixture>(); object unvisitedCopy = null;
        const string sessionA = "native-choice-A", sessionB = "native-choice-B";
        try
        {
            Local("connection_stop"); Local("connection_local");
            for (int index = 0; index < 8; index++)
            {
                string name = index == 0 ? Path.Combine("same-name-a", "Shared manuscript.docx") :
                    index == 1 ? Path.Combine("same-name-b", "Shared manuscript.docx") : "stress-" + index + ".docx";
                fixtures.Add(AddChoiceFixture(name, "Document " + (index + 1) + " has a rough paragraph with unnecessary wording.", index == 1 ? "same-name-alternative.docx" : null));
            }
            ChoiceFixture a = fixtures[0], b = fixtures[1];
            var unchanged = fixtures.ToDictionary(fixture => fixture.Id, ChoiceNativeIdentity);
            var saved = fixtures.ToDictionary(fixture => fixture.Id, fixture => FileHash(fixture.Path));
            ((dynamic)b.Document).Activate();
            var choice = SessionReview(sessionA, "read_manuscript", Values());
            Check(choice.ContainsKey("selection_required") && (bool)choice["selection_required"] &&
                !choice.ContainsKey("text") && !choice.ContainsKey("context_id") && !choice.ContainsKey("revision") && Text(choice, "instruction").Length > 0,
                "An unknown conversation with eight open documents receives a normal choice result without manuscript text or an editing witness.");
            var listed = Records(choice["manuscripts"]);
            Check(listed.Length == fixtures.Count && listed.Select(row => Text(row, "manuscript_id")).Distinct().Count() == fixtures.Count &&
                fixtures.All(fixture => listed.Any(row => Text(row, "manuscript_id") == fixture.Id && Text(row, "documentId") == fixture.Id &&
                    Text(row, "title") == fixture.Title && (bool)row["active"] == (fixture.Id == b.Id))),
                "Every native document is listed with its exact distinct identity, title and active-window status.");
            Check(fixtures.All(fixture => ChoiceNativeIdentity(fixture) == unchanged[fixture.Id] && FileHash(fixture.Path) == saved[fixture.Id]),
                "Requesting a choice leaves all eight native manuscripts, annotations, bookmarks and saved files unchanged.");
            if (a.Title == b.Title)
                Check(listed.Count(row => Text(row, "title") == a.Title) == 2 && a.Id != b.Id,
                    "Simultaneously open documents with the same displayed title retain different manuscript identities.");
            selectionObservations.Add(new { scenario = "same-basename", status = a.Title == b.Title ? "PASS" : "HOST_TITLES_DISTINCT", firstTitle = a.Title, secondTitle = b.Title, firstId = a.Id, secondId = b.Id });

            var contextA = SessionReview(sessionA, "read_manuscript", Values("manuscript_id", a.Id));
            var retainedA = SessionReview(sessionA, "read_manuscript", Values());
            Check(Text(contextA, "manuscript_id") == a.Id && Text(retainedA, "manuscript_id") == a.Id &&
                Text(retainedA, "text") == WordSnapshot.Capture(a.Document).FullText && (bool)retainedA["complete"],
                "An explicit choice and later omitted identity retain inactive document A while document B is active.");
            var selectedList = SessionReview(sessionA, "list_manuscripts", Values());
            Check(Text(selectedList, "selected_manuscript_id") == a.Id && Records(selectedList["manuscripts"]).Length == fixtures.Count &&
                fixtures.All(fixture => Records(selectedList["manuscripts"]).Any(row => Text(row, "manuscript_id") == fixture.Id && Text(row, "title") == fixture.Title)),
                "The explicit manuscript list enumerates every native title and identity while retaining this conversation's chosen document.");
            var independent = SessionReview(sessionB, "read_manuscript", Values());
            Check((bool)independent["selection_required"] && !independent.ContainsKey("text"),
                "Another conversation does not inherit the first conversation's document choice.");
            var contextB = SessionReview(sessionB, "read_manuscript", Values("manuscript_id", b.Id));
            ExpectSessionError(sessionA, "suggest_text", ContextRequest("choice-wrong-session-B", contextB, b.Source, "Document B is concise."), "document_selection_required");
            ExpectSessionError("native-choice-unselected", "suggest_text", ContextRequest("choice-unselected-A", retainedA, a.Source, "Document A is concise."), "document_selection_required");
            Check(fixtures.All(fixture => ChoiceNativeIdentity(fixture) == unchanged[fixture.Id]),
                "A valid witness from another document or conversation cannot bypass the required local conversation choice.");

            UseChoiceFixture(a); ((dynamic)b.Document).Activate();
            int activeWindowBefore = (int)application.ActiveWindow.Hwnd;
            bool screenUpdatingBefore = (bool)application.ScreenUpdating;
            IntPtr foregroundBefore = GetForegroundWindow();
            var editedA = TimedSuggestion(ContextRequest("choice-write-inactive-A", retainedA, a.Source, "Document A states its point clearly."), sessionA);
            Check(WordSnapshot.Identity((object)application.ActiveDocument) == WordSnapshot.Identity(b.Document) &&
                (int)application.ActiveWindow.Hwnd == activeWindowBefore && (bool)application.ScreenUpdating == screenUpdatingBefore,
                "An inactive-document edit restores the exact previously active B window and ScreenUpdating state.");
            selectionObservations.Add(new { scenario = "inactive-edit-window", foregroundBefore = foregroundBefore.ToInt64(),
                foregroundAfter = GetForegroundWindow().ToInt64(), activeWindow = activeWindowBefore, screenUpdating = screenUpdatingBefore });
            Check(editedA.ContainsKey("current_context") && editedA["current_context"] != null,
                "A committed inactive-document suggestion returns fresh context for its chosen manuscript.");
            var freshA = Map(editedA["current_context"]);
            Check(Text(editedA, "status") == "applied" && (bool)editedA["applied"] && !(bool)editedA["requiresApproval"] &&
                (int)((dynamic)a.Document).Revisions.Count == 0 && Text(freshA, "manuscript_id") == a.Id && Text(freshA, "title") == a.Title &&
                Text(freshA, "text") == WordSnapshot.Capture(a.Document).FullText && (int)((dynamic)a.Document).Comments.Count == 1 &&
                (string)((dynamic)a.Document).Comments.Item(1).Author == "wordplai" && ChoiceNativeIdentity(b) == unchanged[b.Id] &&
                fixtures.Skip(2).All(fixture => ChoiceNativeIdentity(fixture) == unchanged[fixture.Id]),
                "A direct edit, rationale and returned context target chosen inactive A while B and all other native documents remain exact.");

            contextB = SessionReview(sessionB, "read_manuscript", Values());
            Check(Text(contextB, "manuscript_id") == b.Id && Text(SessionReview(sessionA, "read_manuscript", Values()), "manuscript_id") == a.Id,
                "Separate conversations retain their own writing targets as native windows change.");
            string aBeforeB = ChoiceNativeIdentity(a);
            UseChoiceFixture(b); ((dynamic)a.Document).Activate();
            TimedSuggestion(ContextRequest("choice-write-inactive-B", contextB, b.Source, "Document B states its point clearly."), sessionB);
            Check(ChoiceNativeIdentity(a) == aBeforeB && (int)((dynamic)b.Document).Comments.Count == 1 &&
                WordSnapshot.Capture(b.Document).MainText.Contains("Document B states its point clearly."),
                "The second conversation edits only its chosen B even while A is the active native window.");
            SessionReview(sessionA, "read_manuscript", Values("manuscript_id", b.Id));
            Check(Text(SessionReview(sessionA, "read_manuscript", Values()), "manuscript_id") == b.Id,
                "An explicit full read switches only the requesting conversation to another open document.");
            ExpectSessionError(sessionA, "suggest_text", ContextRequest("choice-prior-A-denied", freshA, Outside, "An earlier choice must not be reused."), "document_selection_required");

            ChoiceFixture unrelated = fixtures[7];
            dynamic unrelatedParts = ((dynamic)unrelated.Document).CustomXMLParts.SelectByNamespace("urn:undertree:word:selective-editor:v1");
            unrelatedParts.Item(1).Delete();
            ((dynamic)unrelated.Document).CustomXMLParts.Add("<review xmlns=\"urn:undertree:word:selective-editor:v1\">{invalid review JSON</review>");
            string unrelatedBefore = ChoiceNativeIdentity(unrelated);
            var firstReader = SessionReview(sessionA, "read_manuscript", Values("manuscript_id", a.Id));
            SessionReview("native-choice-same-document", "read_manuscript", Values("manuscript_id", a.Id));
            string concurrentBefore = ChoiceNativeIdentity(a), otherBefore = ChoiceNativeIdentity(b);
            var concurrent = ExpectSessionError("native-choice-same-document", "suggest_text", ContextRequest("choice-cross-client-witness", firstReader, Outside, "A freshly reviewed tail keeps its meaning."), "context_changed");
            Check(ChoiceNativeIdentity(a) == concurrentBefore && ChoiceNativeIdentity(b) == otherBefore,
                "A different client's same-document witness is refused without changing either native document.");
            var concurrentContext = Map(Map(concurrent.Details)["current_context"]);
            Check((bool)concurrentContext["complete"] && Text(concurrentContext, "manuscript_id") == a.Id &&
                Text(concurrentContext, "text") == WordSnapshot.Capture(a.Document).FullText,
                "A cross-client witness refusal returns complete context for the requesting client's chosen document.");
            UseChoiceFixture(a); ((dynamic)b.Document).Activate();
            TimedSuggestion(ContextRequest("choice-independent-read", firstReader, Outside, "A freshly reviewed tail keeps its meaning."), sessionA);
            Check(ChoiceNativeIdentity(b) == otherBefore && (int)((dynamic)a.Document).Comments.Count == 2 &&
                ChoiceNativeIdentity(unrelated) == unrelatedBefore && FileHash(unrelated.Path) == saved[unrelated.Id],
                "An independent same-document read keeps the first client's witness usable until a real mutation, preserving unrelated corrupt and native documents.");
            selectionObservations.Add(new { scenario = "unrelated-corrupt-json", status = "PASS", chosenId = a.Id, unchangedUnrelatedId = unrelated.Id });

            ((dynamic)a.Document).Save();
            var beforeCopy = SessionReview(sessionA, "read_manuscript", Values("manuscript_id", a.Id));
            string copyPath = Path.Combine(output, "document-choice", "late-unvisited-copy.docx"); File.Copy(a.Path, copyPath);
            unvisitedCopy = (object)application.Documents.Open(copyPath, ReadOnly: false, Visible: false); ((dynamic)b.Document).Activate();
            string originalBeforeCopy = ChoiceNativeIdentity(a), bBeforeCopy = ChoiceNativeIdentity(b);
            var copyRead = SessionRefusal(sessionA, "read_manuscript", Values());
            var copyWrite = SessionRefusal(sessionA, "suggest_text", ContextRequest("choice-unvisited-copy-denied", beforeCopy, ReopenedChoiceSource, "Document A is revised again."));
            Check(copyRead != null && copyWrite != null &&
                (copyRead.Message.IndexOf("copy", StringComparison.OrdinalIgnoreCase) >= 0 || copyRead.Message.IndexOf("identity", StringComparison.OrdinalIgnoreCase) >= 0) &&
                ChoiceNativeIdentity(a) == originalBeforeCopy && ChoiceNativeIdentity(b) == bBeforeCopy,
                "Opening an unvisited native copy with the same persisted identity blocks both reads and writes without changing the originals.");
            ((dynamic)unvisitedCopy).Close(0); unvisitedCopy = null;
            Check((bool)SessionReview(sessionA, "read_manuscript", Values("manuscript_id", a.Id))["complete"],
                "Closing the conflicting unvisited copy permits a new complete read of the chosen original.");

            Local("connection_stop"); Local("connection_local");
            var stopped = SessionReview(sessionA, "read_manuscript", Values());
            var stoppedB = SessionReview(sessionB, "read_manuscript", Values());
            Check((bool)stopped["selection_required"] && !stopped.ContainsKey("text") && (bool)stoppedB["selection_required"] && !stoppedB.ContainsKey("text"),
                "Stopping access clears prior conversation choices and requires a fresh choice when several documents remain.");
            var closingContext = SessionReview(sessionA, "read_manuscript", Values("manuscript_id", a.Id));
            var closedRequest = ContextRequest("choice-closed-target-denied", closingContext, ReopenedChoiceSource, "An explicit reopened read is required.");
            ((dynamic)a.Document).Save(); CloseChoiceFixture(a);
            foreach (ChoiceFixture fixture in fixtures.Skip(2)) CloseChoiceFixture(fixture);
            ((dynamic)b.Document).Activate(); string onlyB = ChoiceNativeIdentity(b);
            ExpectSessionError(sessionA, "read_manuscript", Values(), "document_closed");
            ExpectSessionError(sessionA, "suggest_text", closedRequest, "document_closed");
            Check((int)application.Documents.Count == 1 && ChoiceNativeIdentity(b) == onlyB,
                "A closed retained choice never falls through to the single remaining document.");

            a.Document = (object)application.Documents.Open(a.Path, ReadOnly: false, Visible: false); a.Closed = false;
            ((dynamic)b.Document).Activate();
            ExpectSessionError(sessionA, "read_manuscript", Values(), "document_closed");
            ExpectSessionError(sessionA, "suggest_text", closedRequest, "document_closed");
            var reopened = SessionReview(sessionA, "read_manuscript", Values("manuscript_id", a.Id, "offset", 0));
            Check((bool)reopened["complete"] && Text(reopened, "manuscript_id") == a.Id && Text(reopened, "text") == WordSnapshot.Capture(a.Document).FullText,
                "Reopening preserves document identity but requires an explicit complete read to restore the conversation's target.");
            UseChoiceFixture(a); ((dynamic)b.Document).Activate();
            TimedSuggestion(ContextRequest("choice-reopened-resumed", reopened, ReopenedChoiceSource, "Document A remains clear after reopening."), sessionA);
            Check(ChoiceNativeIdentity(b) == onlyB,
                "A deliberately rebound reopened document edits its remaining authorized paragraph while the other document and earlier green wording remain unchanged.");

            Local("connection_stop"); Local("connection_local");
            string capacityBeforeA = ChoiceNativeIdentity(a), capacityBeforeB = ChoiceNativeIdentity(b);
            bool complete = true;
            for (int index = 0; index < 64; index++)
            {
                var context = SessionReview("native-capacity-" + index, "read_manuscript", Values("manuscript_id", a.Id));
                complete &= (bool)context["complete"] && Text(context, "manuscript_id") == a.Id;
            }
            Check(complete, "Sixty-four native conversations can independently select and completely read the same open document.");
            ExpectSessionError("native-capacity-overflow", "read_manuscript", Values("manuscript_id", a.Id), "conversation_limit");
            Check(ChoiceNativeIdentity(a) == capacityBeforeA && ChoiceNativeIdentity(b) == capacityBeforeB,
                "The conversation bound refuses overflow without native document mutation.");
            Local("connection_stop"); Local("connection_local");
            Check((bool)SessionReview("native-capacity-overflow", "read_manuscript", Values("manuscript_id", a.Id))["complete"],
                "Stopping and restarting access releases conversation capacity and allows a fresh explicit choice.");
            BackgroundDirectSave(a, b);
        }
        finally
        {
            if (unvisitedCopy != null) ((dynamic)unvisitedCopy).Close(0);
            foreach (ChoiceFixture fixture in fixtures) CloseChoiceFixture(fixture);
            document = previous; documentId = previousId; savedPath = previousPath; savedHash = previousHash;
            if (previous != null) ((dynamic)previous).Activate();
            Local("connection_stop"); Local("connection_local");
        }
    }

    private static void BackgroundTimerSave()
    {
        ChoiceFixture a = null, b = null;
        try
        {
            a = AddChoiceFixture("timer-a.docx", Original);
            b = AddChoiceFixture("timer-b.docx", "Another document stays unchanged.");
            UseChoiceFixture(a); ((dynamic)b.Document).Activate();
            int request = 0;
            foreach (var edit in new[] {
                Tuple.Create(Original, "Document A states its point clearly."),
                Tuple.Create(Outside, "A freshly reviewed tail keeps its meaning."),
                Tuple.Create(ReopenedChoiceSource, "Document A remains clear after reopening.") })
            {
                var context = SessionReview("native-background-A", "read_manuscript", Values("manuscript_id", a.Id));
                TimedSuggestion(ContextRequest("timer-direct-" + request++, context, edit.Item1, edit.Item2), "native-background-A");
            }
            BackgroundDirectSave(a, b);
        }
        finally
        {
            if (a != null) CloseChoiceFixture(a);
            if (b != null) CloseChoiceFixture(b);
            document = null; documentId = null;
        }
    }

    private static void BackgroundDirectSave(ChoiceFixture a, ChoiceFixture b)
    {
        SessionReview("native-background-B", "read_manuscript", Values("manuscript_id", b.Id));
        UseChoiceFixture(a); ((dynamic)b.Document).Activate();
        string beforeA = FileHash(a.Path), beforeB = FileHash(b.Path), nativeB = ChoiceNativeIdentity(b);
        string expected = WordSnapshot.Capture(a.Document).MainText;
        int windowBefore = (int)application.ActiveWindow.Hwnd;
        bool screenBefore = (bool)application.ScreenUpdating;
        var failures = new List<string>();
        Check(editor.BackgroundReviewDocuments().Any(target => WordSnapshot.Identity(target) == WordSnapshot.Identity(a.Document)) &&
            !editor.BackgroundReviewDocuments().Any(target => WordSnapshot.Identity(target) == WordSnapshot.Identity(b.Document)),
            "The production background provider includes shared inactive A and excludes active B.");
        string[] greenIds;
        using (var autosave = new WordAutoSave((object)application, () => true,
            target => editor.RefreshReviewDocument(target), failures.Add, () => editor.BackgroundReviewDocuments()))
        {
            WaitBackgroundSaved(a, autosave, failures);
            Check(FileHash(a.Path) != beforeA && FileHash(b.Path) == beforeB && ChoiceNativeIdentity(b) == nativeB &&
                WordSnapshot.Identity((object)application.ActiveDocument) == WordSnapshot.Identity(b.Document) &&
                (int)application.ActiveWindow.Hwnd == windowBefore && (bool)application.ScreenUpdating == screenBefore,
                "Production background saving changes inactive A's real file while preserving B's native state, saved bytes, window and ScreenUpdating.");
            Check((int)((dynamic)a.Document).Revisions.Count == 0 &&
                WordSnapshot.Identity((object)application.ActiveDocument) == WordSnapshot.Identity(b.Document),
                "Inactive direct wording has no revisions awaiting acceptance while B remains the active document.");
            greenIds = BackgroundGreenIds(a, expected);
            Check(FileHash(a.Path) != beforeA && FileHash(b.Path) == beforeB && ChoiceNativeIdentity(b) == nativeB &&
                WordSnapshot.Identity((object)application.ActiveDocument) == WordSnapshot.Identity(b.Document) &&
                (int)application.ActiveWindow.Hwnd == windowBefore && (bool)application.ScreenUpdating == screenBefore,
                "The background coordinator saves immediately protected direct wording without a manuscript read or changing B.");
        }
        CloseChoiceFixture(a);
        a.Document = (object)application.Documents.Open(a.Path, ReadOnly: false, Visible: false); a.Closed = false;
        ((dynamic)b.Document).Activate(); UseChoiceFixture(a);
        Check(BackgroundGreenIds(a, expected).OrderBy(id => id).SequenceEqual(greenIds.OrderBy(id => id)),
            "Reopening the automatically saved document retains the exact green protected identities and accepted native wording.");
        var context = SessionReview("native-background-A", "read_manuscript", Values("manuscript_id", a.Id));
        string acceptedBefore = ChoiceNativeIdentity(a);
        var refusal = SessionRefusal("native-background-A", "suggest_text", Request("background-accepted-denied", context,
            "Document A remains clear after reopening.", "An accepted passage cannot be rewritten."));
        Check(refusal != null && ChoiceNativeIdentity(a) == acceptedBefore && ChoiceNativeIdentity(b) == nativeB && FileHash(b.Path) == beforeB,
            "A subsequent remote rewrite of background-accepted wording is refused with both native documents unchanged.");
        selectionObservations.Add(new { scenario = "inactive-background-direct-save", status = "PASS", acceptedIds = greenIds,
            beforeFileHash = beforeA, directFileHash = FileHash(a.Path), unchangedOtherFileHash = beforeB });
    }

    private static void WaitBackgroundSaved(ChoiceFixture fixture, WordAutoSave autosave, List<string> failures)
    {
        var watch = Stopwatch.StartNew();
        while (watch.Elapsed < TimeSpan.FromSeconds(10) && failures.Count == 0 &&
            (!(bool)((dynamic)fixture.Document).Saved || (int)application.BackgroundSavingStatus != 0))
        { Application.DoEvents(); Thread.Sleep(20); }
        Check(failures.Count == 0 && autosave.LastError == null && (bool)((dynamic)fixture.Document).Saved &&
            (int)application.BackgroundSavingStatus == 0,
            "The production inactive-document coordinator completes a confirmed native save: " + String.Join("; ", failures));
        selectionObservations.Add(new { scenario = "production-background-save-timer", elapsedMilliseconds = watch.Elapsed.TotalMilliseconds,
            trigger = "Production timer; no manual Pulse or Save call." });
    }

    private static string[] BackgroundGreenIds(ChoiceFixture fixture, string expected)
    {
        WordSnapshot snapshot = WordSnapshot.Capture(fixture.Document);
        SavedWordReview persisted = WordReviewStore.Load(fixture.Document);
        var greenRules = persisted.Policy.Rules.Where(rule => rule.Action == "protect" &&
            String.Equals(rule.Color, "#74b9a4", StringComparison.OrdinalIgnoreCase)).Select(rule => rule.Id).ToArray();
        var marks = persisted.Policy.Marks.Where(mark => greenRules.Contains(mark.RuleId) && !mark.Stale).ToArray();
        Check(snapshot.MainText == expected && (int)((dynamic)fixture.Document).Revisions.Count == 0 &&
            marks.Any(mark => mark.Quote == "Document A remains clear after reopening.") &&
            marks.Any(mark => mark.Quote == "Document A states its point clearly.") &&
            marks.Any(mark => mark.Quote == "A freshly reviewed tail keeps its meaning."),
            "Without refreshing local state, native saved metadata proves the accepted paragraphs are protected.");
        Check(marks.All(mark => (bool)((dynamic)fixture.Document).Bookmarks.Exists(mark.Id) &&
            (string)((dynamic)fixture.Document).Bookmarks.Item(mark.Id).Range.Text == mark.Quote &&
            (int)((dynamic)fixture.Document).Bookmarks.Item(mark.Id).Range.Font.Shading.BackgroundPatternColor == 10795380 &&
            (int)((dynamic)fixture.Document).Bookmarks.Item(mark.Id).Range.HighlightColorIndex == 0),
            "Every automatically accepted marker has its exact native bookmark, text and visible green shading.");
        return marks.Select(mark => mark.Id).ToArray();
    }

    private static Dictionary<string, object> ContextRequest(string id, Dictionary<string, object> context, string quote, string replacement)
    {
        string text = Text(context, "text");
        var mark = Records(context["highlights"]).Single(candidate => !(bool)candidate["stale"] &&
            text.Substring(Convert.ToInt32(candidate["from"]), Convert.ToInt32(candidate["to"]) - Convert.ToInt32(candidate["from"])) == quote);
        return Values("manuscript_id", context["manuscript_id"], "request_id", id, "revision", context["revision"], "context_id", context["context_id"],
            "from", mark["from"], "to", mark["to"], "quote", quote, "replacement", replacement, "rationale", "Clarify this passage while preserving its meaning.");
    }

    private static void VerifyCurrentContext(Dictionary<string, object> context)
    {
        WordSnapshot fresh = WordSnapshot.Capture(document);
        Check((bool)context["complete"] && Convert.ToInt32(context["from"]) == 0 && Convert.ToInt32(context["to"]) == fresh.FullText.Length &&
            Convert.ToInt32(context["total"]) == fresh.FullText.Length && Text(context, "text") == fresh.FullText && Text(context, "revision") == fresh.Revision &&
            Text(context, "manuscript_id") == documentId && Text(context, "documentId") == documentId && Text(context, "context_id").Length > 0 &&
            Text(context, "reviewMode") == "direct" && context["next_offset"] == null,
            "A successful small edit returns complete exact current native text, document identity, revision and next-edit witness.");
        Check(Json.Serialize(context["review"]) == Json.Serialize(Map(fresh.Review)) &&
            Records(Map(context["review"])["feedback"]).Any(comment => Text(comment, "author") == "wordplai"),
            "The successor context contains the actual new native rationale and revision structures.");
        var local = Local("state");
        var marks = Records(local["marks"]);
        Check(Records(context["highlights"]).Length == marks.Length && Records(context["highlights"]).All(mark => marks.Any(current =>
            Text(current, "id") == Text(mark, "id") && Text(current, "ruleId") == Text(mark, "ruleId") &&
            Convert.ToInt32(current["from"]) == Convert.ToInt32(mark["from"]) && Convert.ToInt32(current["to"]) == Convert.ToInt32(mark["to"]) &&
            (bool)current["stale"] == (bool)mark["stale"])) && Json.Serialize(context["rules"]) == Json.Serialize(local["rules"]),
            "Returned marker offsets, identities, validity and unchanged local rules match the native editor after mutation.");
    }

    private static void SuccessorContextAndDocumentRecovery()
    {
        const string second = "A second rough paragraph repeats a different point.";
        const string secondRevision = "A second paragraph states its point clearly.";
        NewFixture("successor-context", Original + "\r" + second);
        Mark(Original, "snappier"); Mark(second, "snappier"); Mark(Outside, "snappier");
        Local("connection_stop"); Local("connection_local");
        var initial = Review("read_manuscript", Values());
        Check((bool)initial["complete"] && Text(initial, "manuscript_id") == documentId && Text(initial, "text") == WordSnapshot.Capture(document).FullText,
            "After connection reset an initial read with one open document starts immediately with complete current context.");
        var firstRequest = ContextRequest("successor-first", initial, Original, Revised);
        var first = TimedSuggestion(firstRequest);
        Check(first.ContainsKey("current_context") && first["current_context"] != null,
            "A newly committed small suggestion supplies a successor context without another manuscript read.");
        var current = Map(first["current_context"]); VerifyCurrentContext(current);
        var next = ContextRequest("successor-next-paragraph", current, second, secondRevision);
        Check(Convert.ToInt32(next["from"]) == Revised.Length + 1 && Convert.ToInt32(next["to"]) == Revised.Length + 1 + second.Length,
            "The returned successor offsets include the preceding edit's length change for the next disjoint paragraph.");
        var secondResult = TimedSuggestion(next);
        Check(WordSnapshot.Capture(document).MainText.Contains(Revised + "\r" + secondRevision) && (int)document.Comments.Count == 2,
            "The next disjoint paragraph commits directly from the returned context without an intervening read.");
        var secondContext = Map(secondResult["current_context"]); VerifyCurrentContext(secondContext);
        string native = new WordNativeEditState().NativeIdentity(document);
        var replay = Review("suggest_text", firstRequest);
        Check(replay["current_context"] == null && new WordNativeEditState().NativeIdentity(document) == native && (int)document.Comments.Count == 2,
            "Replaying an older successful receipt returns no cached successor context and adds no review annotations.");
        var pending = ContextRequest("successor-comment-stale", secondContext, Outside, "The tail is now explicitly revised.");
        WordSnapshot beforeComment = WordSnapshot.Capture(document); int tail = beforeComment.MainText.IndexOf(Outside, StringComparison.Ordinal);
        document.Comments.Add(beforeComment.Range(document, tail, tail + Outside.Length), "Review this tail before the next change.");
        Check(Text(secondContext, "revision") != WordSnapshot.Capture(document).Revision,
            "A later native comment invalidates the returned successor revision.");
        Denied(pending, "A successor witness cannot bypass a later native comment; the denied attempt changes no native state.");
        TimedSuggestion(ContextRequest("successor-after-comment-read", Read(), Outside, "The tail is now explicitly revised."));

        string oldId = documentId;
        NewFixture("active-document-recovery", Original); Mark(Original, "snappier");
        string unchanged = new WordNativeEditState().NativeIdentity(document); McpOperationException refusal = null;
        try { Review("suggest_text", pending); } catch (McpOperationException error) { refusal = error; }
        Check(refusal != null && refusal.Code == "document_closed" && new WordNativeEditState().NativeIdentity(document) == unchanged,
            "Closing the selected document returns document_closed without writing to the only remaining open document.");
        var details = Map(refusal.Details);
        Check(Text(details, "manuscript_id") == oldId && !(bool)details["request_applied"] && Text(details, "instruction").Length > 0,
            "The closed-target diagnostic preserves the chosen identity and confirms that no edit was applied.");
        ExpectSessionError(null, "read_manuscript", Values(), "document_closed");
        Check(new WordNativeEditState().NativeIdentity(document) == unchanged,
            "An omitted manuscript identity cannot silently retarget a closed choice to the remaining active document.");
        var recovered = Review("read_manuscript", Values("manuscript_id", documentId, "offset", 0));
        TimedSuggestion(ContextRequest("active-document-resumed", recovered, Original, Revised));
        Check(WordSnapshot.Capture(document).MainText.Contains(Revised) && (int)document.Comments.Count == 1,
            "Explicitly choosing the remaining document and reading it completely resumes the writing request safely.");
    }

    private static void StaleRevisionReturnsCurrentContext()
    {
        const string second = "A second rough paragraph repeats a different point.";
        const string secondRevision = "A second paragraph states its point clearly.";
        NewFixture("stale-revision-recovery", Original + "\r" + second);
        Mark(Original, "snappier"); Mark(second, "snappier");
        var initial = Read();
        string address = Text(Local("connection_status"), "localMcpUrl");
        var httpContext = Map(NativeMcpClient.ForAddress(address).Call("read_manuscript", Values("manuscript_id", documentId))["structuredContent"]);
        var stale = ContextRequest("stale-revision-second", httpContext, second, secondRevision);
        var first = TimedSuggestion(ContextRequest("stale-revision-first", initial, Original, Revised));
        Check(Text(first, "status") == "applied" && (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 1,
            "Stale-revision recovery starts after exactly one successful direct paragraph edit and rationale.");
        WordSnapshot beforeStale = WordSnapshot.Capture(document);
        string beforeNative = new WordNativeEditState().NativeIdentity(document);
        string beforePolicy = Json.Serialize(WordReviewStore.Load((object)document).Policy);
        string beforeLocal = Json.Serialize(Local("state"));
        int beforeRevisions = document.Revisions.Count, beforeComments = document.Comments.Count;
        var refused = SuggestionRpc(stale);
        Check(refused.ContainsKey("isError") && (bool)refused["isError"] && refused.ContainsKey("structuredContent"),
            "The original pre-edit context is rejected as an actual MCP tool error after a successful mutation.");
        var diagnostic = Map(Map(refused["structuredContent"])["error"]);
        var diagnosticText = Map(Json.DeserializeObject(Text(Records(refused["content"])[0], "text")));
        Check(Text(diagnostic, "code") == "manuscript_changed" && Json.Serialize(Map(refused["structuredContent"])) == Json.Serialize(diagnosticText),
            "The stale-edit response identifies manuscript_changed consistently in structured and text MCP payloads.");
        Check(new WordNativeEditState().NativeIdentity(document) == beforeNative && WordSnapshot.Capture(document).MainText == beforeStale.MainText &&
            (int)document.Revisions.Count == beforeRevisions && (int)document.Comments.Count == beforeComments,
            "Returning stale-edit recovery leaves exact native wording, revisions, comments and bookmarks unchanged.");
        Check(Json.Serialize(WordReviewStore.Load((object)document).Policy) == beforePolicy && Json.Serialize(Local("state")) == beforeLocal && FileHash(savedPath) == savedHash,
            "A stale rejected edit creates no rule, mark, proposal, receipt, sharing or saved-file mutation.");
        var recovery = Map(diagnostic["details"]);
        Check(recovery.ContainsKey("current_context") && recovery["current_context"] != null,
            "A small-document stale-edit error supplies complete fresh current context for immediate recovery.");
        var recovered = Map(recovery["current_context"]); VerifyCurrentContext(recovered);
        Check(Text(recovered, "manuscript_id") == documentId && Text(recovered, "revision") != Text(initial, "revision"),
            "The recovery witness belongs to the correct current document and its post-edit revision.");
        var resumed = Map(SuggestionRpc(ContextRequest("stale-revision-resumed", recovered, second, secondRevision))["structuredContent"]);
        Check(Text(resumed, "status") == "applied" && (bool)resumed["applied"] && !(bool)resumed["requiresApproval"] &&
            WordSnapshot.Capture(document).MainText.Contains(Revised + "\r" + secondRevision) && (int)document.Comments.Count == beforeComments + 1 &&
            (int)document.Revisions.Count == beforeRevisions && beforeRevisions == 0,
            "The remaining paragraph resumes from the error's fresh quote, range and witness as a direct edit with one comment and no extra read.");
        CheckGreen(Revised); CheckGreen(secondRevision);
    }

    private static void LargeContextRequiresSequentialReading()
    {
        const string second = "A second rough paragraph needs clearer wording.";
        const string secondRevision = "A second paragraph states its point clearly.";
        NewFixture("large-context", Original + "\r" + second); Mark(Original, "snappier"); Mark(second, "snappier");
        document.Sections.Item(1).Headers.Item(1).Range.Text = new string('H', 33000);
        var firstPage = Review("read_manuscript", Values("manuscript_id", documentId, "offset", 0));
        Check(!(bool)firstPage["complete"] && firstPage["context_id"] == null && Convert.ToInt32(firstPage["total"]) > 32768 &&
            Text(firstPage, "manuscript_id") == documentId,
            "A large read-only header is included in a paged active-document read and cannot grant a partial-read witness.");
        var unread = Request("large-incomplete-denied", firstPage, Original, Revised);
        Denied(unread, "Reading only the first large-context page cannot authorize a suggestion.");
        var lastPage = Review("read_manuscript", Values("manuscript_id", firstPage["manuscript_id"], "offset", firstPage["to"]));
        Check((bool)lastPage["complete"] && Text(lastPage, "context_id").Length > 0,
            "Retaining the returned manuscript identity completes sequential reading and establishes the current witness.");
        var large = TimedSuggestion(Request("large-context-suggestion", lastPage, Original, Revised));
        Check(large["current_context"] == null && WordSnapshot.Capture(document).FullText.Length > 32768 && (int)document.Comments.Count == 1,
            "A successful large-context edit returns no truncated or oversized successor context.");
        Denied(Request("large-old-context-denied", lastPage, second, secondRevision),
            "Without a complete successor context the pre-edit witness cannot authorize another large-document edit.");
        var partial = Review("read_manuscript", Values());
        Denied(Request("large-second-partial-denied", partial, second, secondRevision),
            "A new first page after the edit also remains insufficient for editing.");
        var complete = Review("read_manuscript", Values("manuscript_id", partial["manuscript_id"], "offset", partial["to"]));
        TimedSuggestion(Request("large-context-resumed", complete, second, secondRevision));
        Check((int)document.Comments.Count == 2 && (int)document.Revisions.Count == 0 &&
            WordSnapshot.Capture(document).MainText.Contains(Revised + "\r" + secondRevision),
            "A fresh complete sequential read resumes the remaining large-document paragraph while preserving the first direct edit and both comments.");
        CheckGreen(Revised); CheckGreen(secondRevision);
    }

    private static void PromptColorsEditDirectly()
    {
        foreach (string mode in new[] { "suggestions", "direct" }) PromptColorEditsDirectly(mode);
    }

    private static void PromptColorEditsDirectly(string storedMode)
    {
        NewFixture("direct-color-" + storedMode, Original);
        SaveRule("blue", "prompt", storedMode, "#AACCEE"); Mark(Original, "blue");
        Mark(Outside, "protect");
        string neighborId = Text(Records(Local("state")["marks"]).Single(mark => Text(mark, "quote") == Outside), "id");
        string neighborSignature = WordSnapshot.FormattingSignature(document.Bookmarks.Item(neighborId).Range);
        document.ActiveWindow.View.ShowComments = false;
        document.ActiveWindow.View.ShowRevisionsAndComments = false;
        string rules = Json.Serialize(Local("state")["rules"]);
        WordSnapshot before = WordSnapshot.Capture(document);
        var request = Request("direct-color-edit", Read(), Original, Revised);
        request["rationale"] = "Remove repetition and keep the original meaning.";
        Dictionary<string, object> response = TimedSuggestion(request);
        Check(Text(response, "status") == "applied" && (bool)response["applied"] && !(bool)response["requiresApproval"] && Text(response, "reviewMode") == "direct",
            "A prompt with stored " + storedMode + " mode commits directly without requiring local approval.");
        Check((int)document.Revisions.Count == 0 && !(bool)document.TrackRevisions &&
            (bool)document.ActiveWindow.View.ShowComments && (bool)document.ActiveWindow.View.ShowRevisionsAndComments &&
            WordSnapshot.Capture(document).MainText == before.MainText.Replace(Original, Revised) && !((string)document.Content.Text).Contains(Original),
            "The exact replacement and rationale are visible immediately with unchanged surrounding prose and no tracked revisions.");
        Check(Json.Serialize(Local("state")["rules"]) == rules && AcceptedMarks().Count(mark => Text(mark, "quote") == Revised) == 1 &&
            (string)document.Bookmarks.Item(neighborId).Range.Text == Outside &&
            WordSnapshot.FormattingSignature(document.Bookmarks.Item(neighborId).Range) == neighborSignature,
            "Direct admission leaves stored rules unchanged, protects the edited passage and preserves the original protected neighbor's identity and formatting.");
        CheckGreen(Revised);
        VerifyCurrentContext(Map(response["current_context"]));
        Check((int)document.Comments.Count == 1 && (string)document.Comments.Item(1).Author == "wordplai" &&
            ((string)document.Comments.Item(1).Range.Text).Contains((string)request["rationale"]),
            "Exactly one native wordplai comment contains the concise reason for the change.");
        var feedback = Records(Map(Read()["review"])["feedback"]).Single();
        Check(Convert.ToInt32(feedback["paragraphFrom"]) == 0 && Convert.ToInt32(feedback["paragraphTo"]) == Revised.Length &&
            Convert.ToInt32(feedback["from"]) == 0 && Convert.ToInt32(feedback["to"]) == Revised.Length,
            "The rationale anchor identifies the revised passage and its paragraph in complete MCP context.");
        string native = new WordNativeEditState().NativeIdentity(document);
        Dictionary<string, object> replay = Review("suggest_text", request);
        Check(Text(replay, "status") == "applied" && (bool)replay["applied"] && (bool)replay["replayed"] &&
            !(bool)replay["requiresApproval"] && new WordNativeEditState().NativeIdentity(document) == native && (int)document.Comments.Count == 1,
            "Replaying a recorded direct request adds neither edits, green markers nor a duplicate rationale comment.");
        var changed = new Dictionary<string, object>(request); changed["rationale"] = "A different reason.";
        Denied(changed, "A consumed request identifier cannot be reused to attach different commentary.");
        DenyChange(Revised, "The green passage cannot be rewritten again.");
        Check((bool)document.Undo(1), "The complete direct edit, green marker and rationale share one Undo record.");
        Local("state");
        Check(WordSnapshot.Capture(document).MainText == before.MainText && (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0,
            "One native Undo restores the original wording and removes its rationale comment together.");
        Denied(request, "Undo does not permit replaying a consumed remote request.");
    }

    private static void DirectCommentAnchors()
    {
        DirectCommentSharedEnding();
        DirectCommentDenied("no-shared-ending", Original, "A clear paragraph conveys its point!", Original.Length);
        DirectCommentDenied("narrow-anchor", Original, Revised, 7);
        const string unicodeSource = "A rough paragraph \ud83d\ude00";
        const string unicodeReplacement = "A clear paragraph \ud83e\ude00";
        Check(Char.IsSurrogatePair(unicodeSource, unicodeSource.Length - 2) &&
            Char.IsSurrogatePair(unicodeReplacement, unicodeReplacement.Length - 2) &&
            unicodeSource[unicodeSource.Length - 1] == unicodeReplacement[unicodeReplacement.Length - 1] &&
            unicodeSource[unicodeSource.Length - 2] != unicodeReplacement[unicodeReplacement.Length - 2],
            "The Unicode refusal fixture uses valid whole scalars sharing only their low-surrogate half.");
        DirectCommentDenied("half-surrogate", unicodeSource, unicodeReplacement, unicodeSource.Length);
    }

    private static void DirectCommentSharedEnding()
    {
        const string earlier = "An earlier rough paragraph repeats unnecessary wording.";
        const string earlierRevision = "An earlier paragraph states its point clearly.";
        const string rationale = "Clarify this paragraph while preserving the reviewer's thread.";
        NewFixture("direct-comments-shared-ending", earlier + "\r" + Original);
        Mark(earlier, "snappier"); Mark(Original, "snappier"); Mark(Outside, "protect");
        TimedSuggestion(ContextRequest("direct-comment-earlier", Read(), earlier, earlierRevision));
        object earlierComment = (object)document.Comments.Item(1);
        WordSnapshot setup = WordSnapshot.Capture(document);
        int from = setup.MainText.IndexOf(Original, StringComparison.Ordinal), neighborFrom = setup.MainText.IndexOf(Outside, StringComparison.Ordinal);
        dynamic user = document.Comments.Add(setup.Range(document, from, from + Original.Length), "Preserve this complete native reviewer thread.");
        user.Author = "Local reviewer";
        dynamic reply = user.Replies.Add(user.Scope, "Keep the meaning and do not invent claims."); reply.Author = "Second reviewer";
        dynamic neighbor = document.Comments.Add(setup.Range(document, neighborFrom, neighborFrom + Outside.Length), "The protected neighbor remains unchanged.");
        neighbor.Author = "Neighbor reviewer"; neighbor.Done = true;
        var request = ContextRequest("direct-comment-shared-ending", Read(), Original, Revised); request["rationale"] = rationale;
        var neighborMark = AcceptedMarks().Single(mark => Text(mark, "quote") == Outside);
        string neighborId = Text(neighborMark, "id"), neighborBodyIdentity = DirectNeighborBodyIdentity(Text(neighborMark, "id"), "before");
        WordSnapshot before = WordSnapshot.Capture(document);
        string discussion = Json.Serialize(before.Review.feedback.Select(DirectCommentContent).ToArray());
        var beforeEvidence = DirectCommentEvidence("shared-ending-before");
        object[] retainedComments = { earlierComment, (object)user, (object)reply, (object)neighbor };
        var liveBefore = DirectCommentLiveEvidence("before", retainedComments);
        var result = TimedSuggestion(request);
        WordSnapshot after = WordSnapshot.Capture(document);
        var afterEvidence = DirectCommentEvidence("shared-ending-after");
        var liveAfter = DirectCommentLiveEvidence("after", retainedComments);
        Check(Text(result, "status") == "applied" && (bool)result["applied"] && !(bool)result["requiresApproval"] &&
            (int)document.Revisions.Count == 0 && after.MainText == before.MainText.Substring(0, from) + Revised + before.MainText.Substring(from + Original.Length),
            "A full-anchor direct rewrite changes only its authorized passage without tracked revisions.");
        Check(Json.Serialize(after.Review.feedback.Where(comment => comment.text != rationale).Select(DirectCommentContent).ToArray()) == discussion &&
            liveBefore.All(row => Convert.ToInt32(row["collectionMatches"]) == 1) && liveAfter.All(row => Convert.ToInt32(row["collectionMatches"]) == 1) &&
            liveBefore.Select(row => Json.Serialize(Map(row["heldState"])["content"])).SequenceEqual(liveAfter.Select(row => Json.Serialize(Map(row["heldState"])["content"]))),
            "The earlier rationale, user root/reply and neighboring discussion preserve their held live objects, exact content/date/Done and unique current collection membership.");
        var currentUser = after.Review.feedback.Single(comment => comment.author == "Local reviewer" && comment.text == "Preserve this complete native reviewer thread.");
        var currentReply = currentUser.replies.Single();
        var freshUser = Map(liveAfter[1]["freshState"]); var freshReply = Map(liveAfter[2]["freshState"]); var freshNeighbor = Map(liveAfter[3]["freshState"]);
        Check((int)user.Replies.Count == 1 && (string)user.Range.Text == "Preserve this complete native reviewer thread." &&
            (string)reply.Range.Text == "Keep the meaning and do not invent claims." &&
            currentUser.from == from + Revised.Length - 1 && currentUser.to == from + Revised.Length &&
            currentReply.from == currentUser.from && currentReply.to == currentUser.to &&
            Convert.ToInt32(freshUser["scopeFrom"]) == currentUser.from && Convert.ToInt32(freshUser["scopeTo"]) == currentUser.to &&
            Convert.ToInt32(freshReply["scopeFrom"]) == currentReply.from && Convert.ToInt32(freshReply["scopeTo"]) == currentReply.to &&
            Convert.ToInt32(freshReply["ancestorIndex"]) == Convert.ToInt32(freshUser["index"]) &&
            (string)neighbor.Range.Text == "The protected neighbor remains unchanged." && (bool)neighbor.Done,
            "The original live thread remains usable; fresh collection and review data prove exact shared-ending root/reply anchors and parentage.");
        Check((int)document.Comments.Count == Convert.ToInt32(beforeEvidence["comments"]) + 1 &&
            after.Review.feedback.Count(comment => comment.author == "wordplai" && comment.text == rationale &&
                comment.from == from && comment.to == from + Revised.Length) == 1 &&
            (string)document.Bookmarks.Item(neighborId).Name == neighborId && (string)document.Bookmarks.Item(neighborId).Range.Text == Outside &&
            (int)document.Bookmarks.Item(neighborId).Range.Start == after.MainText.IndexOf(Outside, StringComparison.Ordinal) &&
            (int)document.Bookmarks.Item(neighborId).Range.End == after.MainText.IndexOf(Outside, StringComparison.Ordinal) + Outside.Length &&
            Convert.ToInt32(freshNeighbor["scopeFrom"]) == (int)document.Bookmarks.Item(neighborId).Range.Start &&
            Convert.ToInt32(freshNeighbor["scopeTo"]) == (int)document.Bookmarks.Item(neighborId).Range.End &&
            DirectNeighborBodyIdentity(neighborId, "after") == neighborBodyIdentity &&
            Text(beforeEvidence, "rules") == Text(afterEvidence, "rules") && FileHash(savedPath) == savedHash,
            "The edit adds exactly one complete-passage rationale, preserves protected native formatting and rules, and never saves through MCP.");
        CheckGreen(Revised); CheckGreen(earlierRevision); CheckGreen(Outside); VerifyCurrentContext(Map(result["current_context"]));
        string native = new WordNativeEditState().NativeIdentity(document), review = after.Review.Serialize();
        var replay = Review("suggest_text", request);
        Check(Text(replay, "status") == "applied" && (bool)replay["replayed"] && new WordNativeEditState().NativeIdentity(document) == native &&
            WordSnapshot.Capture(document).Review.Serialize() == review,
            "Exact replay of a commented direct edit duplicates no native thread, rationale, green marker or wording.");
        Check((bool)document.Undo(1), "One Undo reverses the commented direct edit, rationale and new green protection together.");
        Local("state");
        var undo = DirectCommentEvidence("shared-ending-undo");
        Check(Text(undo, "mainText") == before.MainText && Json.Serialize(undo["review"]) == Json.Serialize(beforeEvidence["review"]) &&
            AcceptedMarks().All(mark => Text(mark, "quote") != Revised) &&
            (string)document.Bookmarks.Item(neighborId).Name == neighborId && (string)document.Bookmarks.Item(neighborId).Range.Text == Outside &&
            (int)document.Bookmarks.Item(neighborId).Range.Start == before.MainText.IndexOf(Outside, StringComparison.Ordinal) &&
            (int)document.Bookmarks.Item(neighborId).Range.End == before.MainText.IndexOf(Outside, StringComparison.Ordinal) + Outside.Length &&
            DirectNeighborBodyIdentity(neighborId, "undo") == neighborBodyIdentity,
            "Undo restores exact source, original comment/reply content, structure and anchors, retaining earlier and neighboring native protection.");
        CheckGreen(earlierRevision); CheckGreen(Outside);
        Denied(request, "Undo cannot replay the consumed commented direct request while its sharing witness is revoked.");
    }

    private static void DirectCommentDenied(string name, string quote, string replacement, int anchorLength)
    {
        NewFixture("direct-comments-" + name, quote);
        document.Content.Font.Name = "Segoe UI Emoji"; document.Content.Font.NameBi = "Segoe UI Emoji";
        Mark(quote, "snappier"); Mark(Outside, "protect");
        WordSnapshot setup = WordSnapshot.Capture(document);
        int from = setup.MainText.IndexOf(quote, StringComparison.Ordinal);
        dynamic user = document.Comments.Add(setup.Range(document, from, from + anchorLength), "Preserve this unsafe-anchor reviewer thread.");
        user.Author = "Local reviewer";
        dynamic reply = user.Replies.Add(user.Scope, "Do not lose the original reply."); reply.Author = "Second reviewer";
        var request = ContextRequest("direct-comment-" + name, Read(), quote, replacement);
        var before = DirectCommentEvidence(name + "-before");
        McpOperationException refusal = null;
        try { TimedSuggestion(request); } catch (McpOperationException error) { refusal = error; }
        var after = DirectCommentEvidence(name + "-after");
        File.WriteAllText(Path.Combine(output, "direct-comment-" + name + "-diagnostic.json"), Json.Serialize(new {
            code = refusal == null ? null : refusal.Code, message = refusal == null ? null : refusal.Message }));
        Check(refusal != null && refusal.Message == "Word cannot preserve this native comment anchor for that replacement. Keep a shared ending within the commented text, or review the change locally.",
            "The " + name + " direct edit returns the precise preflight comment-anchor refusal.");
        Check(Json.Serialize(after) == Json.Serialize(before) && (int)user.Replies.Count == 1 &&
            (string)user.Range.Text == "Preserve this unsafe-anchor reviewer thread." && (string)reply.Range.Text == "Do not lose the original reply.",
            "The " + name + " refusal leaves exact native text, full discussion, marker/receipt/rule metadata, formatting and saved bytes unchanged.");
    }

    private static object DirectCommentContent(WordCommentContext comment)
    { return new { comment.storyType, comment.author, comment.text, comment.date, comment.resolved, replies = comment.replies.Select(DirectCommentContent).ToArray() }; }

    private static Dictionary<string, object> DirectCommentLiveState(object value)
    {
        dynamic comment = value, scope = comment.Scope, range = comment.Range;
        object ancestor = (object)comment.Ancestor;
        return Values("content", new { author = (string)comment.Author, text = (string)range.Text,
            date = ((DateTime)comment.Date).ToUniversalTime().ToString("o"), done = (bool)comment.Done, replies = (int)comment.Replies.Count },
            "index", (int)comment.Index, "textStory", (int)range.StoryType, "textFrom", (int)range.Start, "textTo", (int)range.End,
            "scopeStory", (int)scope.StoryType, "scopeFrom", (int)scope.Start, "scopeTo", (int)scope.End,
            "ancestorIndex", ancestor == null ? 0 : (int)((dynamic)ancestor).Index);
    }

    private static Dictionary<string, object>[] DirectCommentLiveEvidence(string phase, object[] held)
    {
        object[] collection = ((IEnumerable)document.Comments).Cast<object>().ToArray();
        var rows = held.Select(comment => {
            var state = DirectCommentLiveState(comment);
            object[] matches = collection.Where(current => DirectCommentMembershipKey(DirectCommentLiveState(current)) == DirectCommentMembershipKey(state)).ToArray();
            return Values("heldState", state, "collectionMatches", matches.Length, "freshState", matches.Length == 1 ? DirectCommentLiveState(matches[0]) : null,
                "heldIUnknownDiagnostic", WordSnapshot.Identity(comment),
                "freshIUnknownDiagnostic", matches.Length == 1 ? (object)WordSnapshot.Identity(matches[0]) : null,
                "secondFreshIUnknownDiagnostic", WordSnapshot.Identity((object)document.Comments.Item(Convert.ToInt32(state["index"]))));
        }).ToArray();
        File.WriteAllText(Path.Combine(output, "direct-comment-live-" + phase + ".json"), Json.Serialize(new {
            identityLimit = "IUnknown/export IDs and held range positions are diagnostic; Word exposes aliases and cached ranges. Membership uses exact content/index/story/ancestor, while fresh collection and review data independently verify current anchors.", comments = rows }));
        return rows;
    }

    private static string DirectCommentMembershipKey(Dictionary<string, object> state)
    { return Json.Serialize(new { content = state["content"], index = state["index"], textStory = state["textStory"], scopeStory = state["scopeStory"], ancestorIndex = state["ancestorIndex"] }); }

    private static string DirectNeighborBodyIdentity(string bookmarkId, string phase)
    {
        string raw = (string)document.Bookmarks.Item(bookmarkId).Range.WordOpenXML;
        File.WriteAllText(Path.Combine(output, "direct-comment-neighbor-" + phase + ".xml"), raw);
        var xml = new System.Xml.XmlDocument(); xml.LoadXml(raw);
        var namespaces = new System.Xml.XmlNamespaceManager(xml.NameTable);
        const string word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"; namespaces.AddNamespace("w", word);
        var references = xml.SelectNodes("//w:body//w:commentRangeStart|//w:body//w:commentRangeEnd|//w:body//w:commentReference", namespaces).Cast<System.Xml.XmlElement>().ToArray();
        foreach (var reference in references)
        {
            var definition = xml.SelectNodes("//w:comments/w:comment", namespaces).Cast<System.Xml.XmlElement>().Single(comment => comment.GetAttribute("id", word) == reference.GetAttribute("id", word));
            if (definition.GetAttribute("author", word) != "Neighbor reviewer" ||
                String.Concat(definition.SelectNodes(".//w:t", namespaces).Cast<System.Xml.XmlNode>().Select(text => text.InnerText)) != "The protected neighbor remains unchanged.")
                throw new InvalidOperationException("The neighboring native comment reference lost its exact original discussion.");
            reference.SetAttribute("id", word, "0");
        }
        if (!references.Any(reference => reference.LocalName == "commentReference")) throw new InvalidOperationException("The neighboring native comment reference is missing.");
        // Export numbers are volatile. Bookmark name/range/current green and live comment membership are verified independently.
        return WordSnapshot.NativeBodyIdentityFromXml(xml.OuterXml);
    }

    private static Dictionary<string, object> DirectCommentEvidence(string name)
    {
        WordSnapshot snapshot = WordSnapshot.Capture(document);
        SavedWordReview saved = WordReviewStore.Load((object)document);
        string xml = (string)document.Content.WordOpenXML;
        var source = new System.Xml.XmlDocument(); source.LoadXml(xml);
        var namespaces = new System.Xml.XmlNamespaceManager(source.NameTable);
        const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"; namespaces.AddNamespace("w", wordNamespace);
        object[] bindings = source.SelectNodes("//w:comments/w:comment", namespaces).Cast<System.Xml.XmlElement>().Select(comment => (object)new {
            id = comment.GetAttribute("id", wordNamespace), author = comment.GetAttribute("author", wordNamespace),
            text = String.Concat(comment.SelectNodes(".//w:t", namespaces).Cast<System.Xml.XmlNode>().Select(text => text.InnerText)) }).ToArray();
        var evidence = Values("mainText", snapshot.MainText, "review", Map(snapshot.Review), "nativeIdentity", new WordNativeEditState().NativeIdentity(document),
            "metadata", WordReviewStore.Fingerprint((object)document), "policy", Json.Serialize(saved.Policy), "rules", Json.Serialize(saved.Policy.Rules),
            "comments", (int)document.Comments.Count, "revisions", (int)document.Revisions.Count, "commentBindings", bindings, "savedFileHash", FileHash(savedPath));
        File.WriteAllText(Path.Combine(output, "direct-comment-" + name + ".json"), Json.Serialize(evidence));
        File.WriteAllText(Path.Combine(output, "direct-comment-" + name + ".xml"), xml);
        return evidence;
    }

    private static WordDocumentReview SeedLegacyPending(Dictionary<string, object> request)
    {
        // Only fixture metadata is migrated to pending. Bindings and tracked application use the real current authorities.
        WordDocumentReview state = NativeDocumentState();
        WordSnapshot snapshot = WordSnapshot.Capture(document);
        state.Snapshot = snapshot;
        var proposal = new Proposal { Id = Text(request, "request_id"), DocumentId = documentId, Revision = Text(request, "revision"),
            From = Convert.ToInt32(request["from"]), To = Convert.ToInt32(request["to"]), Quote = Text(request, "quote"),
            Replacement = Text(request, "replacement"), Rationale = request.ContainsKey("rationale") ? Text(request, "rationale") : "" };
        state.Policy.Submit(proposal, Text(request, "context_id"), snapshot.MainText, snapshot.Revision);
        state.Policy.Proposals.Single(candidate => candidate.Id == proposal.Id).Status = "pending";
        return state;
    }

    private static void ApplyLegacyTracked(Dictionary<string, object> request, bool expectDenied = false)
    {
        PolicyState before = NativeDocumentState().Policy.ExportState();
        string native = new WordNativeEditState().NativeIdentity(document);
        WordDocumentReview state = SeedLegacyPending(request);
        Exception refusal = null;
        try { typeof(WordEditor).GetMethod("Apply", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(editor,
            new object[] { state, Text(request, "request_id"), false, true }); }
        catch (TargetInvocationException error) { refusal = error.InnerException; }
        if (!expectDenied)
        {
            if (refusal != null) throw refusal;
            Check(state.Policy.Proposals.Single(proposal => proposal.Id == Text(request, "request_id")).Status == "tracked",
                "The legacy fixture uses actual pending approval and native tracked application rather than a new MCP direct submission.");
            return;
        }
        Check(refusal != null && refusal.Message.Contains("comment anchor") &&
            new WordNativeEditState().NativeIdentity(document) == native,
            "A legacy tracked follow-up that would lose a comment anchor is refused without any native mutation.");
        state.Policy = (PolicyEngine)typeof(WordEditor).GetMethod("CreatePolicy", BindingFlags.Static | BindingFlags.NonPublic)
            .Invoke(null, new object[] { before });
        state.Policy.Share(state.Id);
        state.Snapshot = null;
    }

    private static void LegacyPendingRemainsPending()
    {
        NewFixture("legacy-pending", Original); Mark(Original, "snappier");
        var request = Request("legacy-pending-replay", Read(), Original, Revised);
        WordDocumentReview state = SeedLegacyPending(request);
        typeof(WordEditor).GetMethod("Persist", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(editor, new object[] { state, true });
        document.Save(); string path = savedPath;
        document.Close(0); document = application.Documents.Open(path, ReadOnly: false, Visible: false); document.Activate();
        Read();
        string native = new WordNativeEditState().NativeIdentity(document), receipt = Json.Serialize(WordReviewStore.Load((object)document).Policy.Proposals);
        var replay = Review("suggest_text", request);
        Check(Text(replay, "status") == "pending" && !(bool)replay["applied"] && (bool)replay["requiresApproval"] && (bool)replay["replayed"] &&
            new WordNativeEditState().NativeIdentity(document) == native && (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0 &&
            Json.Serialize(WordReviewStore.Load((object)document).Policy.Proposals) == receipt && AcceptedMarks().Length == 0,
            "A reopened legacy pending receipt replays as pending without promoting, editing, commenting or protecting its source.");
    }

    private static void FollowUpRejectAndDeletion()
    {
        NewFixture("follow-up", Original); Mark(Original, "snappier");
        WordSnapshot initial = WordSnapshot.Capture(document);
        dynamic paragraph = document.Paragraphs.First.Range;
        dynamic user = document.Comments.Add(document.Range((int)paragraph.End - 1, (int)paragraph.End), "Keep this reviewer thread.");
        user.Author = "Local reviewer";
        dynamic reply = user.Replies.Add(user.Scope, "Do not add new claims.");
        reply.Author = "Second reviewer";
        var first = Request("follow-up-first", Read(), Original, Revised); first.Remove("rationale");
        ApplyLegacyTracked(first);
        Check((int)document.Comments.Count == 3 && ((string)user.Range.Text).Contains("reviewer thread") &&
            (int)user.Replies.Count == 1 && ((string)user.Replies.Item(1).Range.Text).Contains("new claims"),
            "Adding a default rationale preserves the existing native reviewer and reply identities and text.");
        var firstFeedback = Records(Map(Read()["review"])["feedback"]);
        Check(firstFeedback.Count(comment => Text(comment, "author") == "wordplai" && Text(comment, "text").Trim().Length > 0 && Text(comment, "text").Length <= 240) == 1,
            "An omitted rationale still creates one brief useful native comment.");
        string trackedNative = new WordNativeEditState().NativeIdentity(document);
        var trackedReplay = Review("suggest_text", first);
        Check(Text(trackedReplay, "status") == "tracked" && !(bool)trackedReplay["applied"] && (bool)trackedReplay["requiresApproval"] &&
            (bool)trackedReplay["replayed"] && new WordNativeEditState().NativeIdentity(document) == trackedNative && AcceptedMarks().Length == 0,
            "An existing tracked receipt replays unchanged rather than being promoted into a direct edit or green protection.");
        Denied(Request("no-remote-promotion", Read(), Revised, "A new prompt cannot accept an existing tracked change."),
            "New direct wording cannot replace a paragraph with unresolved native tracked changes.");
        ApplyLegacyTracked(Request("follow-up-without-shared-anchor", Read(), Revised, "Fresh wording conveys the same idea!"), true);
        const string followUp = "Fresh wording conveys the same idea.";
        ApplyLegacyTracked(Request("follow-up-second", Read(), Revised, followUp));
        Check((int)document.Comments.Count == 4 && WordSnapshot.Capture(document).MainText.Contains(followUp) &&
            (int)user.Replies.Count == 1 && ((string)user.Range.Text).Contains("reviewer thread"),
            "A follow-up preserves rationale and reviewer threads through their shared native ending.");
        document.Revisions.RejectAll(); Local("state");
        Check(WordSnapshot.Capture(document).MainText == initial.MainText && (int)document.Comments.Count == 2 && AcceptedMarks().Length == 0 &&
            ((string)user.Range.Text).Contains("reviewer thread") && (int)user.Replies.Count == 1 && ((string)user.Replies.Item(1).Range.Text).Contains("new claims"),
            "Native Reject restores original source and the reviewer thread; Word removes rationales attached to rejected inserted wording.");
        var unsafeRequest = Request("follow-up-after-reject-unsafe-anchor", Read(), Original, Revised);
        var beforeUnsafe = DirectCommentEvidence("after-reject-unsafe-before");
        McpOperationException refusal = null;
        try { TimedSuggestion(unsafeRequest); } catch (McpOperationException error) { refusal = error; }
        var afterUnsafe = DirectCommentEvidence("after-reject-unsafe-after");
        Check(refusal != null && refusal.Message == "Word cannot preserve this native comment anchor for that replacement. Keep a shared ending within the commented text, or review the change locally.",
            "After native Reject, an edit that would lose the restored word-only reviewer anchor receives the precise preflight refusal.");
        Check(Json.Serialize(afterUnsafe) == Json.Serialize(beforeUnsafe),
            "The unsupported post-Reject request changes no native text, full discussion, marker/receipt/rule metadata, formatting or saved bytes.");
        const string afterRejectReplacement = "A clear paragraph uses concise wording.";
        var resumed = TimedSuggestion(Request("follow-up-after-reject", Read(), Original, afterRejectReplacement));
        var resumedContext = Map(resumed["current_context"]);
        var restoredThread = Records(Map(resumedContext["review"])["feedback"]).Single(comment => Text(comment, "author") == "Local reviewer" && Text(comment, "text") == "Keep this reviewer thread.");
        var restoredReply = Records(restoredThread["replies"]).Single();
        int retainedWord = afterRejectReplacement.IndexOf("wording", StringComparison.Ordinal);
        Check(Text(resumed, "status") == "applied" && (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 3 &&
            WordSnapshot.Capture(document).MainText == initial.MainText.Replace(Original, afterRejectReplacement) &&
            Convert.ToInt32(restoredThread["from"]) == retainedWord && Convert.ToInt32(restoredThread["to"]) == retainedWord + "wording".Length &&
            Convert.ToInt32(restoredReply["from"]) == retainedWord && Convert.ToInt32(restoredReply["to"]) == retainedWord + "wording".Length &&
            Text(restoredReply, "author") == "Second reviewer" && Text(restoredReply, "text") == "Do not add new claims." &&
            ((string)user.Range.Text).Contains("reviewer thread") && (int)user.Replies.Count == 1,
            "A fresh complete read and new request preserve the actual restored wording anchor, apply concise direct text and add one rationale without repairing comments or markers.");
        CheckGreen(afterRejectReplacement); VerifyCurrentContext(resumedContext);

        NewFixture("removal", Original); Mark(Original, "discard");
        var removal = Request("remove-complete", Read(), Original, ""); removal["rationale"] = " ";
        TimedSuggestion(removal);
        Check((int)document.Revisions.Count == 1 && (int)document.Comments.Count == 1 &&
            ((string)document.Comments.Item(1).Range.Text).Contains("Remove"),
            "A complete tracked removal receives a useful rationale anchored outside the deleted wording.");
        document.Revisions.RejectAll(); Local("state");
        Check(WordSnapshot.Capture(document).MainText.Contains(Original) && (int)document.Comments.Count == 1,
            "Rejecting complete removal restores its permission and retains its native rationale.");
        TimedSuggestion(Request("remove-again", Read(), Original, ""));
        document.Revisions.AcceptAll(); Local("state");
        Check(!WordSnapshot.Capture(document).MainText.Contains(Original) && (int)document.Comments.Count == 0 && AcceptedMarks().Length == 0,
            "Word's native Accept removes fully deleted comment anchors without an empty green or stale editable marker.");
        Mark(Outside, "snappier"); TimedSuggestion(Request("after-removal", Read(), Outside, "The unmarked tail is now explicitly authorized."));
    }

    private static void ProtectedUnitsAndSave()
    {
        const string source = "A rough [FIXED] paragraph repeats its point.";
        const string replacement = "A clear paragraph retains [FIXED].";
        NewFixture("protected-save", source); SaveRule("blue", "prompt", "direct", "#AACCEE"); Mark(source, "blue");
        WordSnapshot snapshot = WordSnapshot.Capture(document); int start = snapshot.MainText.IndexOf("[FIXED]", StringComparison.Ordinal);
        document.Hyperlinks.Add(snapshot.Range(document, start, start + 7), "https://example.invalid/fixed");
        Mark("[FIXED]", "protect");
        var unit = Records(Local("state")["marks"]).Single(mark => Text(mark, "ruleId") == "protect");
        string id = Text(unit, "id"), signature = WordSnapshot.FormattingSignature(document.Bookmarks.Item(id).Range);
        DenyChange("[FIXED]", "Changed protected unit");
        DenyChange(Outside, "Unauthorized outside rewrite");
        var context = Read(); var stale = Request("stale-exact-quote", context, source, replacement); stale["quote"] = "A forged quote.";
        Denied(stale, "An incorrect exact quote adds no native change or rationale comment.");
        var request = Request("preserve-and-review", Read(), source, "");
        request["replacementParts"] = new object[] { Values("text", "A clear paragraph retains "), Values("protectedId", id), Values("text", ".") };
        request["rationale"] = String.Join(" ", Enumerable.Repeat("Preserve the protected phrase while reducing repetition.", 10));
        WordSnapshot beforeEdit = WordSnapshot.Capture(document);
        string rulesBefore = Json.Serialize(Local("state")["rules"]);
        var result = TimedSuggestion(request); document.Activate();
        Check((string)document.Bookmarks.Item(id).Range.Text == "[FIXED]" &&
            WordSnapshot.FormattingSignature(document.Bookmarks.Item(id).Range) == signature &&
            (int)document.Bookmarks.Item(id).Range.Hyperlinks.Count == 1 &&
            (string)document.Bookmarks.Item(id).Range.Hyperlinks.Item(1).Address == "https://example.invalid/fixed",
            "A direct edit and rationale preserve the exact protected native identity, hyperlink and formatting.");
        Check(Text(result, "status") == "applied" && (bool)result["applied"] && !(bool)result["requiresApproval"] &&
            (int)document.Revisions.Count == 0 && WordSnapshot.Capture(document).MainText == beforeEdit.MainText.Replace(source, replacement) &&
            Json.Serialize(Local("state")["rules"]) == rulesBefore && AcceptedMarks().Count(mark => Text(mark, "id") != id) == 2,
            "The direct enclosing edit preserves all surrounding prose and local rules and immediately protects only the two rewritten prose gaps.");
        CheckGreen("A clear paragraph retains "); CheckGreen(".");
        var comments = Records(Map(Read()["review"])["feedback"]);
        Check(comments.Length == 1 && Text(comments[0], "text").Length <= 240 && Text(comments[0], "text").EndsWith("...", StringComparison.Ordinal),
            "A long supplied rationale becomes one concise comment without changing the recorded request.");
        string unchangedFile = FileHash(savedPath);
        Check(unchangedFile == savedHash, "Remote editing without the installed autosave timer leaves the existing saved file untouched until a local save occurs.");
        document.Save(); string path = savedPath;
        document.Close(0); document = application.Documents.Open(path, ReadOnly: false, Visible: false); document.Activate();
        Local("connection_stop"); Local("connection_local");
        Check((int)document.Revisions.Count == 0 && (int)document.Comments.Count == 1 && WordSnapshot.Capture(document).MainText.Contains(replacement),
            "Native save and reopen retain direct wording, its rationale and protected content without pending revisions.");
        Check(AcceptedMarks().Count(mark => Text(mark, "id") != id) == 2 && (int)document.Comments.Count == 1 &&
            (string)document.Bookmarks.Item(id).Range.Text == "[FIXED]" && WordSnapshot.FormattingSignature(document.Bookmarks.Item(id).Range) == signature,
            "Reopening retains both immediate green prose gaps and the original protected unit and rationale.");
        DenyChange(replacement, "Unauthorized accepted rewrite.");
        document.Save(); document.Close(0); document = application.Documents.Open(path, ReadOnly: false, Visible: false); document.Activate();
        Check(AcceptedMarks().Count(mark => Text(mark, "id") != id) == 2 && (int)document.Comments.Count == 1 && (int)document.Revisions.Count == 0,
            "Saving accepted wording persists green protection and its rationale without pending revisions.");
    }

    private static void Denied(Dictionary<string, object> request, string message)
    {
        string before = new WordNativeEditState().NativeIdentity(document); bool denied = false;
        try { Review("suggest_text", request); } catch (McpOperationException) { denied = true; }
        Check(denied && new WordNativeEditState().NativeIdentity(document) == before, message);
    }
    private static void NewFixture(string name, string text)
    {
        if (document != null) document.Close(0);
        document = application.Documents.Add(Visible: false); document.Activate(); document.TrackRevisions = false;
        document.Content.Text = text + "\r" + Outside + "\r";
        savedPath = Path.Combine(output, name + ".docx"); document.SaveAs2(savedPath, 16); savedHash = FileHash(savedPath);
        documentId = Text(Map(Local("state")["document"]), "id"); documentIds.Add(documentId);
    }

    private static void SaveRule(string id, string action, string mode, string color)
    { Local("save_rule", Values("id", id, "label", id, "color", color, "action", action, "accessMode", mode, "prompt", action == "prompt" ? "Retain the meaning." : "")); }

    private static void Mark(string quote, string rule)
    {
        WordSnapshot snap = WordSnapshot.Capture(document); int from = snap.MainText.IndexOf(quote, StringComparison.Ordinal);
        snap.Range(document, from, from + quote.Length).Select(); var selection = Local("selection_preview");
        selection["ruleId"] = rule; selection["prompt"] = ""; Local("mark_selection", selection);
    }

    private static Dictionary<string, object> Read()
    {
        int offset = 0; Dictionary<string, object> page;
        do { page = Review("read_manuscript", Values("manuscript_id", documentId, "offset", offset)); offset = Convert.ToInt32(page["to"]); } while (!(bool)page["complete"]);
        return page;
    }

    private static Dictionary<string, object> Request(string id, Dictionary<string, object> context, string quote, string replacement)
    {
        int from = WordSnapshot.Capture(document).MainText.IndexOf(quote, StringComparison.Ordinal);
        return Values("manuscript_id", documentId, "request_id", id, "revision", context["revision"], "context_id", context["context_id"],
            "from", from, "to", from + quote.Length, "quote", quote, "replacement", replacement, "rationale", "Clarify this passage while preserving its meaning.");
    }

    private static void DenyChange(string quote, string replacement)
    {
        var request = Request("denied_" + Guid.NewGuid().ToString("N"), Read(), quote, replacement);
        string before = new WordNativeEditState().NativeIdentity(document); bool denied = false;
        try { Review("suggest_text", request); } catch (McpOperationException) { denied = true; }
        Check(denied && new WordNativeEditState().NativeIdentity(document) == before,
            "Unauthorized wording changes are denied without changing native text, review or comments.");
    }

    private static Dictionary<string, object>[] AcceptedMarks()
    {
        var state = Local("state"); var rules = Records(state["rules"]);
        var greenRules = rules.Where(rule => Text(rule, "action") == "protect" && Text(rule, "color").ToLowerInvariant() == "#74b9a4").Select(rule => Text(rule, "id")).ToArray();
        return Records(state["marks"]).Where(mark => greenRules.Contains(Text(mark, "ruleId")) && !(bool)mark["stale"]).ToArray();
    }

    private static void CheckGreen(string quote)
    {
        WordDocumentReview state = NativeDocumentState();
        var greenRules = state.Policy.Rules.Where(rule => rule.Action == "protect" &&
            String.Equals(rule.Color, "#74b9a4", StringComparison.OrdinalIgnoreCase)).Select(rule => rule.Id).ToArray();
        var marks = state.Policy.Marks.Where(mark => greenRules.Contains(mark.RuleId) && !mark.Stale && mark.Quote == quote).ToArray();
        Check(marks.Length == 1 && (bool)document.Bookmarks.Exists(marks[0].Id) &&
            (string)document.Bookmarks.Item(marks[0].Id).Range.Text == quote &&
            (int)document.Bookmarks.Item(marks[0].Id).Range.Font.Shading.BackgroundPatternColor == 10795380 &&
            (int)document.Bookmarks.Item(marks[0].Id).Range.HighlightColorIndex == 0,
            "The direct passage has exactly one current protected identity and exact native green paint: " + quote);
    }

    private static Dictionary<string, object> Local(string operation, Dictionary<string, object> args = null)
    { return Map(editor.Dispatch(operation, args ?? Values())); }
    private static Dictionary<string, object> Review(string tool, Dictionary<string, object> args)
    {
        try { return Map(typeof(WordEditor).GetMethod("Review", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(editor, new object[] { tool, args })); }
        catch (TargetInvocationException error) { throw error.InnerException; }
    }

    private static Dictionary<string, object> SuggestionRpc(Dictionary<string, object> request)
    {
        string address = Text(Local("connection_status"), "localMcpUrl");
        return NativeMcpClient.ForAddress(address).Call("suggest_text", request);
    }

    private static object AttachOwned(Process process)
    {
        IntPtr target = IntPtr.Zero; DateTime deadline = DateTime.UtcNow.AddSeconds(30);
        while (target == IntPtr.Zero && DateTime.UtcNow < deadline)
        {
            EnumWindows(delegate(IntPtr parent, IntPtr state) {
                uint pid; GetWindowThreadProcessId(parent, out pid); if (pid != process.Id) return true;
                EnumChildWindows(parent, delegate(IntPtr child, IntPtr ignored) { var name = new StringBuilder(80); GetClassName(child, name, name.Capacity); if (name.ToString() == "_WwG") target = child; return true; }, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
            if (target == IntPtr.Zero) { Application.DoEvents(); Thread.Sleep(50); }
        }
        Check(target != IntPtr.Zero, "The exact spawned Word PID exposes an owned native document window.");
        object native = null; Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
        while (native == null && DateTime.UtcNow < deadline)
        {
            if (AccessibleObjectFromWindow(target, 0xfffffff0, ref dispatch, out native) != 0) { native = null; Application.DoEvents(); Thread.Sleep(50); }
        }
        Check(native != null, "The owned Word automation object becomes ready.");
        dynamic app = ((dynamic)native).Application; uint actual; GetWindowThreadProcessId(new IntPtr((int)app.ActiveWindow.Hwnd), out actual);
        Check(actual == process.Id, "Automation stays bound to the separately spawned Word process."); return app;
    }

    private static void WriteResult(string status, string error)
    { if (output != null) File.WriteAllText(Path.Combine(output, "result.json"), Json.Serialize(new { status = status, checks = checks, timings = timings, documentSelection = selectionObservations, timingMetric = "Synchronous request-to-return upper bound on native suggestion and comment appearance; model/network/save excluded.", error = error, installedPackage = "NOT RUN", normalWord = "NOT TOUCHED", publicTunnel = "NOT RUN" })); }
    private static Dictionary<string, object> Values(params object[] pairs)
    { var result = new Dictionary<string, object>(); for (int index = 0; index < pairs.Length; index += 2) result.Add((string)pairs[index], pairs[index + 1]); return result; }
    private static Dictionary<string, object> Map(object value) { return (Dictionary<string, object>)Json.DeserializeObject(Json.Serialize(value)); }
    private static Dictionary<string, object>[] Records(object value) { return ((IEnumerable)value).Cast<object>().Select(Map).ToArray(); }
    private static string Text(Dictionary<string, object> value, string key) { return (string)value[key]; }
    private static string FileHash(string path) { using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (SHA256 sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(stream)); }
    private static void Check(bool condition, string message) { if (!condition) throw new Exception(message); checks.Add(message); Console.WriteLine("PASS: " + message); }
}
