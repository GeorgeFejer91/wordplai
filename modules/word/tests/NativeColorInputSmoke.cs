using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using UnderTree.WordAddin;

internal static partial class NativeHighlightingSmoke
{
    private const string InputColor = "#13A7D6";
    private const string UserText = "SYNTHETIC USER INPUT. ";
    private static readonly List<object> colorResults = new List<object>();
    private static string colorReceiptDirectory;
    private static bool colorBoundaryEnabled, colorBoundaryInserted, colorForceRollback;
    private static int colorSeams;

    private sealed class ColorCase : IDisposable
    {
        internal readonly dynamic App, Document;
        internal readonly WordEditor Editor;
        internal readonly WordDocumentReview Native;
        internal readonly string Id, Original, Sidecar;
        internal readonly Dictionary<string, object> Context;

        internal ColorCase(dynamic app, WordEditor editor, string address)
        {
            App = app; Editor = editor;
            Document = app.Documents.Add(Visible: true);
            app.Visible = false;
            Document.TrackRevisions = false;
            Document.Content.Text = "One [A]. Two [B].";
            Document.Activate();
            app.Visible = false;
            NativeHighlightingSmoke.Check((int)Document.Windows.Count > 0 &&
                (long)WordSnapshot.Identity(app.ActiveDocument) == (long)WordSnapshot.Identity(Document),
                "The owned target has a real native window and is active before state and sharing.");
            Id = UiState.Text(UiState.Map(NativeHighlightingSmoke.Local(editor, "state")["document"]), "id");
            Context = NativeHighlightingSmoke.Share(editor, address, Id);
            var states = (IDictionary<long, WordDocumentReview>)typeof(WordEditor).GetField("documents", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
            Native = states[WordSnapshot.Identity(Document)];
            Original = WordSnapshot.Capture(Document).FullText;
            Sidecar = Path.Combine(colorReceiptDirectory, Id + ".bin");
            NativeHighlightingSmoke.Check(Native.Policy.Marks.Count == 0 && Native.Policy.HighlightReceipts.Count == 0, "A disposable color case starts without markers or receipts.");
        }

        internal object Marking(string request, int from = 0, int to = 3)
        {
            return new { documentId = Id, marking = new { requestId = request, revision = Context["revision"], contextId = Context["context_id"],
                color = InputColor, ranges = new[] { new { from = from, to = to, quote = Original.Substring(from, to - from) } } } };
        }

        public void Dispose() { Document.Close(0); }
    }

    private sealed class ColorView
    {
        private readonly long document;
        private readonly int window, start, end;
        private readonly string text;
        private readonly bool updating;

        internal ColorView(dynamic app)
        {
            document = WordSnapshot.Identity(app.ActiveDocument);
            window = (int)app.ActiveWindow.Hwnd; start = (int)app.Selection.Start; end = (int)app.Selection.End;
            text = (string)app.Selection.Range.Text; updating = (bool)app.ScreenUpdating;
        }

        internal void Verify(dynamic app)
        {
            NativeHighlightingSmoke.Check((long)WordSnapshot.Identity(app.ActiveDocument) == document && (int)app.ActiveWindow.Hwnd == window,
                "Coloring restores the exact active document and window.");
            NativeHighlightingSmoke.Check((int)app.Selection.Start == start && (int)app.Selection.End == end && (string)app.Selection.Range.Text == text,
                "Coloring preserves both collapsed caret and selected native ranges.");
            NativeHighlightingSmoke.Check((bool)app.ScreenUpdating == updating, "Coloring restores the previous ScreenUpdating value.");
        }
    }

    private sealed class ColorObservation
    {
        private readonly string text, native, policy, sidecar;
        private readonly long generation;
        private readonly bool shared;

        internal ColorObservation(ColorCase fixture)
        {
            text = WordSnapshot.Capture(fixture.Document).FullText;
            native = fixture.Native.NativeEdit.NativeIdentity(fixture.Document);
            policy = McpEndpoint.Json().Serialize(fixture.Native.Policy.ExportState());
            sidecar = NativeHighlightingSmoke.ReceiptDigest(fixture.Sidecar);
            generation = fixture.Native.Policy.ContextGeneration; shared = fixture.Native.Policy.IsShared;
        }

        internal void Verify(ColorCase fixture, string label)
        {
            NativeHighlightingSmoke.Check((string)WordSnapshot.Capture(fixture.Document).FullText == text && (string)fixture.Native.NativeEdit.NativeIdentity(fixture.Document) == native,
                label + " preserves complete native text, bookmarks and formatting.");
            NativeHighlightingSmoke.Check(McpEndpoint.Json().Serialize(fixture.Native.Policy.ExportState()) == policy &&
                fixture.Native.Policy.ContextGeneration == generation && fixture.Native.Policy.IsShared == shared,
                label + " preserves policy, local authority and consumed identities.");
            NativeHighlightingSmoke.Check(NativeHighlightingSmoke.ReceiptDigest(fixture.Sidecar) == sidecar, label + " does not rewrite the private receipt journal.");
        }
    }

    private sealed class ColorSampler : IDisposable
    {
        private readonly WordEditor editor;
        private readonly object previous;
        private readonly FieldInfo field;
        private readonly ColorCase fixture;
        internal int Calls;
        internal int ChangeAt, InsertAt;
        internal bool ChangeStamp, Busy, Forbid, Inserted;
        internal bool ForegroundWord = true;

        internal ColorSampler(ColorCase fixture)
        {
            this.fixture = fixture; editor = fixture.Editor;
            field = typeof(WordEditor).GetField("coloringInput", BindingFlags.Instance | BindingFlags.NonPublic);
            if (field == null || field.FieldType != typeof(WordColorInput)) throw new Exception("Frozen coloringInput reflection contract changed.");
            previous = field.GetValue(editor);
            var replacement = new WordColorInput(Sample);
            field.SetValue(editor, replacement);
            NativeHighlightingSmoke.Check(Object.ReferenceEquals(field.GetValue(editor), replacement), "Only this disposable editor receives the injected sampler.");
        }

        private WordColorInput.Sample Sample(IntPtr frame)
        {
            Calls++;
            if (Forbid) throw new Exception("The exact receipt must bypass a new input admission.");
            uint process;
            uint thread = GetWindowThreadProcessId(frame, out process);
            NativeHighlightingSmoke.Seam(frame != IntPtr.Zero && process != 0 && thread != 0, "The injected sample addresses a real native Word frame.");
            var result = new WordColorInput.Sample { Readable = true, Enabled = true, ForegroundWord = ForegroundWord, NormalForeground = true,
                QueueReadable = true, Now = 2000, LastInput = 1000, WordProcess = process, WordThread = thread };
            if (Busy) result.PendingInput = 1;
            if (Calls == InsertAt)
            {
                NativeHighlightingSmoke.Seam((int)fixture.App.UndoRecord.CustomRecordLevel == 1, "Synthetic user text arrives inside the real coloring Undo record.");
                NativeHighlightingSmoke.Seam(fixture.Native.Policy.HighlightReceipts.Count == 1 && fixture.Native.Policy.HighlightReceipts[0].Outcome == "reserved",
                    "The native request was durably reserved before the injected edit.");
                if (InsertAt == 3) NativeHighlightingSmoke.Seam((int)fixture.Document.Bookmarks.Count == 0, "Sample3 follows Start before new native markers.");
                if (InsertAt == 4) NativeHighlightingSmoke.Seam((int)fixture.Document.Bookmarks.Count > 0 && NativeHighlightingSmoke.NativePainted((object)fixture.Document, 0, 3),
                    "Sample4 follows actual native bookmark creation and confirmed paint.");
                fixture.Document.Range(0, 0).InsertBefore(UserText); Inserted = true;
            }
            if (ChangeAt > 0 && Calls >= ChangeAt || ChangeStamp && Inserted) result.LastInput++;
            return result;
        }

        public void Dispose()
        {
            field.SetValue(editor, previous);
            if (!Object.ReferenceEquals(field.GetValue(editor), previous)) throw new Exception("The disposable sampler was not restored.");
        }
    }

    private static void RequestedColors(dynamic app, WordEditor editor, string address, string output)
    {
        FieldInfo directory = typeof(WordReviewStore).GetField("DirectoryPath", BindingFlags.Static | BindingFlags.NonPublic);
        if (directory == null || directory.FieldType != typeof(string)) throw new Exception("Receipt-directory reflection contract changed.");
        object previous = directory.GetValue(null);
        colorReceiptDirectory = Path.GetFullPath(Path.Combine(output, "private-receipts"));
        if (!colorReceiptDirectory.StartsWith(Path.GetFullPath(output) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            throw new Exception("The disposable receipt directory escaped its evidence root.");
        Directory.CreateDirectory(colorReceiptDirectory);
        directory.SetValue(null, colorReceiptDirectory);
        try
        {
            Check(Object.Equals(directory.GetValue(null), colorReceiptDirectory), "Review sidecars are redirected to this disposable process's evidence directory.");
            string[] command = Environment.GetCommandLineArgs();
            string selected = command.Length > 4 ? command[4] : "all";
            if (selected != "all" && !new[] { "quiet-active", "quiet-inactive", "busy-retry", "planning-input", "after-start-input", "after-paint-input", "after-start-text", "after-paint-text", "rollback-boundary", "background-input", "background-rollback" }.Contains(selected))
                throw new Exception("Choose one known native color-input case.");
            if (Selected(selected, "quiet-active")) QuietActive(app, editor, address);
            if (Selected(selected, "quiet-inactive")) QuietInactive(app, editor, address);
            if (Selected(selected, "busy-retry")) EarlyBusy(app, editor, address);
            if (Selected(selected, "planning-input")) PlanningInput(app, editor, address);
            foreach (int phase in new[] { 3, 4 }) foreach (bool stamp in new[] { true, false })
            {
                string name = (phase == 3 ? "after-start" : "after-paint") + (stamp ? "-input" : "-text");
                if (Selected(selected, name)) MixedRecord(app, editor, address, phase, stamp, name);
            }
            if (Selected(selected, "rollback-boundary")) RollbackBoundary(app, editor, address);
            if (Selected(selected, "background-input")) BackgroundInput(app, editor, address);
            if (Selected(selected, "background-rollback")) BackgroundRollback(app, editor, address);
            File.WriteAllText(Path.Combine(output, "color-input-result.json"), McpEndpoint.Json().Serialize(new {
                status = "PASS", cases = colorResults, outcomeAssertions = assertions, samplerSeamAssertions = colorSeams,
                sampling = "Test-only injected samples; real Win32/foreground/GUI-thread idle detection NOT QUALIFIED", normalWord = "NOT TOUCHED" }));
        }
        finally
        {
            directory.SetValue(null, previous);
            if (!Object.Equals(directory.GetValue(null), previous)) throw new Exception("The disposable receipt-directory override was not restored.");
        }
    }

    private static void QuietActive(dynamic app, WordEditor editor, string address)
    {
        using (var fixture = new ColorCase(app, editor, address))
        using (var sampler = new ColorSampler(fixture))
        {
            fixture.Document.Range(4, 4).Select(); var view = new ColorView(app);
            object marking = fixture.Marking("quiet_active");
            var result = Rpc(address, "highlight_text", marking);
            Applied(fixture, result); view.Verify(app);
            Seam(sampler.Calls == 5, "Quiet fresh coloring reaches the five frozen sample boundaries.");
            var beforeReplay = new ColorObservation(fixture); sampler.Forbid = true;
            var replay = Rpc(address, "highlight_text", marking);
            Check(!UiState.Flag(replay, "isError") && UiState.Flag(UiState.Map(replay["structuredContent"]), "replayed"), "An exact successful receipt replays while new input admission would refuse.");
            beforeReplay.Verify(fixture, "Exact successful replay"); view.Verify(app);
            Record("quiet-active", sampler.Calls);
        }
    }

    private static void QuietInactive(dynamic app, WordEditor editor, string address)
    {
        using (var fixture = new ColorCase(app, editor, address))
        using (var sampler = new ColorSampler(fixture))
        {
            sampler.ForegroundWord = false;
            dynamic other = app.Documents.Add(Visible: true);
            app.Visible = false;
            other.Content.Text = "INACTIVE TARGET CONTROL.\r"; other.Activate();
            app.Visible = false;
            Check((int)fixture.Document.Windows.Count > 0 && (int)other.Windows.Count > 0 &&
                (long)WordSnapshot.Identity(app.ActiveDocument) == (long)WordSnapshot.Identity(other),
                "Both owned documents have real native windows and the other document is active before the inactive read.");
            string otherText = (string)other.Content.Text;
            int otherBookmarkCount = (int)other.Bookmarks.Count;
            var otherNative = new WordNativeEditState();
            string otherNativeIdentity = otherNative.NativeIdentity((object)other);
            bool updating = (bool)app.ScreenUpdating;
            try
            {
                foreach (bool selected in new[] { true, false })
                {
                    other.Range(2, selected ? 7 : 2).Select(); app.ScreenUpdating = selected;
                    var view = new ColorView(app);
                    var context = NativeMcpClient.ForAddress(address).Call("read_manuscript", new { manuscript_id = fixture.Id });
                    Check(!UiState.Flag(context, "isError"), "A retained inactive target can refresh its own context.");
                    Check((string)other.Content.Text == otherText && (int)other.Bookmarks.Count == otherBookmarkCount &&
                        otherNative.NativeIdentity((object)other) == otherNativeIdentity,
                        "Reading the retained inactive target preserves the other document's exact native baseline.");
                    var read = UiState.Map(context["structuredContent"]);
                    int from = selected ? 0 : 9;
                    var result = Rpc(address, "highlight_text", new { documentId = fixture.Id, marking = new {
                        requestId = selected ? "inactive_selected" : "inactive_caret", revision = read["revision"], contextId = read["context_id"],
                        color = InputColor, ranges = new[] { new { from = from, to = from + 3, quote = fixture.Original.Substring(from, 3) } } } });
                    Check(!UiState.Flag(result, "isError") && NativePainted(fixture.Document, from, from + 3), "Quiet coloring paints the retained inactive document.");
                    view.Verify(app);
                    Check((string)other.Content.Text == otherText && (int)other.Bookmarks.Count == otherBookmarkCount &&
                        otherNative.NativeIdentity((object)other) == otherNativeIdentity, "The other active document remains untouched.");
                }
                Check(fixture.Native.Policy.Marks.Count == 2 && fixture.Native.Policy.HighlightReceipts.Count == 2 && WordSnapshot.Capture(fixture.Document).FullText == fixture.Original,
                    "Both inactive-target color transactions commit once without changing wording.");
                Record("quiet-inactive", sampler.Calls);
            }
            finally { app.ScreenUpdating = updating; other.Close(0); }
        }
    }

    private static void EarlyBusy(dynamic app, WordEditor editor, string address)
    {
        using (var fixture = new ColorCase(app, editor, address))
        using (var sampler = new ColorSampler(fixture))
        {
            var before = new ColorObservation(fixture); object marking = fixture.Marking("busy_then_quiet"); sampler.Busy = true;
            var result = Rpc(address, "highlight_text", marking); KnownBusy(result);
            before.Verify(fixture, "Early known-busy refusal");
            Check(fixture.Native.Policy.Marks.Count == 0 && fixture.Native.Policy.HighlightReceipts.Count == 0, "Early busy consumes no native color identity or marker.");
            sampler.Busy = false;
            Applied(fixture, Rpc(address, "highlight_text", marking));
            Record("busy-retry", sampler.Calls);
        }
    }

    private static void PlanningInput(dynamic app, WordEditor editor, string address)
    {
        using (var fixture = new ColorCase(app, editor, address))
        using (var sampler = new ColorSampler(fixture))
        {
            string native = fixture.Native.NativeEdit.NativeIdentity(fixture.Document), journal = ReceiptDigest(fixture.Sidecar);
            sampler.ChangeAt = 2; sampler.ForegroundWord = true;
            KnownBusy(Rpc(address, "highlight_text", fixture.Marking("planning_input")));
            Check(fixture.Native.Policy.Marks.Count == 0 && fixture.Native.Policy.HighlightReceipts.Count == 0 && fixture.Native.Policy.Rules.Count == 4,
                "Planning input changes are refused before staged policy, colors or receipts exist.");
            Check(fixture.Native.NativeEdit.NativeIdentity(fixture.Document) == native && ReceiptDigest(fixture.Sidecar) == journal &&
                WordSnapshot.Capture(fixture.Document).FullText == fixture.Original, "Planning refusal changes no native text, paint, bookmark or private receipt.");
            Check(!fixture.Native.Policy.IsShared, "The interrupted plan retires its old context.");
            Record("planning-input", sampler.Calls);
        }
    }

    private static void MixedRecord(dynamic app, WordEditor editor, string address, int phase, bool stamp, string name)
    {
        using (var fixture = new ColorCase(app, editor, address))
        using (var sampler = new ColorSampler(fixture))
        {
            sampler.InsertAt = phase; sampler.ChangeStamp = stamp;
            object marking = fixture.Marking(name.Replace('-', '_'));
            var result = Rpc(address, "highlight_text", marking);
            Seam(sampler.Inserted, "The synthetic native user edit actually reached the selected open-record seam.");
            Check(WordSnapshot.Capture(fixture.Document).FullText == UserText + fixture.Original,
                "PRESERVE_USER_TEXT: a potentially mixed coloring record never erases real native user text (" + name + ").");
            UnknownConsumed(fixture, sampler, result, address, marking, name.Replace('-', '_'));
            Record(name, sampler.Calls);
        }
    }

    private static void RollbackBoundary(dynamic app, WordEditor editor, string address)
    {
        using (var fixture = new ColorCase(app, editor, address))
        using (var sampler = new ColorSampler(fixture))
        {
            colorBoundaryEnabled = true; colorBoundaryInserted = false; colorForceRollback = true;
            try
            {
                object marking = fixture.Marking("rollback_boundary");
                var result = Rpc(address, "highlight_text", marking);
                Seam(colorBoundaryInserted, "The separate native user Undo record was inserted after identity inspection.");
                Check(WordSnapshot.Capture(fixture.Document).FullText == UserText + fixture.Original,
                    "PRESERVE_BOUNDARY_USER_TEXT: the final rollback guard never consumes a newer separately ended native user record.");
                UnknownConsumed(fixture, sampler, result, address, marking, "rollback_boundary");
                Record("rollback-boundary", sampler.Calls);
            }
            finally { colorBoundaryEnabled = false; colorForceRollback = false; }
        }
    }

    private static void BackgroundInput(dynamic app, WordEditor editor, string address)
    {
        using (var fixture = new ColorCase(app, editor, address))
        using (var sampler = new ColorSampler(fixture))
        {
            sampler.ChangeAt = 3; sampler.ForegroundWord = false;
            Applied(fixture, Rpc(address, "highlight_text", fixture.Marking("background_input")));
            Seam(sampler.Calls == 5, "A changed other-app stamp crosses every forward coloring check while Word stays background.");
            Record("background-input", sampler.Calls);
        }
    }

    private static void BackgroundRollback(dynamic app, WordEditor editor, string address)
    {
        using (var fixture = new ColorCase(app, editor, address))
        using (var sampler = new ColorSampler(fixture))
        {
            sampler.ChangeAt = 3; sampler.ForegroundWord = false; colorForceRollback = true;
            try
            {
                object marking = fixture.Marking("background_rollback");
                var result = Rpc(address, "highlight_text", marking);
                Check(WordSnapshot.Capture(fixture.Document).FullText == fixture.Original && NativePainted(fixture.Document, 0, 3) && (int)fixture.Document.Bookmarks.Count > 0,
                    "BACKGROUND_ROLLBACK_NO_UNDO: a changed global input stamp leaves existing native paint intact and never grants rollback permission.");
                UnknownConsumed(fixture, sampler, result, address, marking, "background_rollback");
                Record("background-rollback", sampler.Calls);
            }
            finally { colorForceRollback = false; }
        }
    }

    internal static bool ColorInputForceRollback { get { return colorForceRollback; } }

    internal static void ColorInputAtRollbackBoundary(object application, object document)
    {
        if (!colorBoundaryEnabled) return;
        dynamic app = application, doc = document;
        Seam((int)app.UndoRecord.CustomRecordLevel == 0, "The coloring record ended before the newer user record begins.");
        Seam(NativePainted(doc, 0, 3), "The copied failure follows real confirmed native paint.");
        app.UndoRecord.StartCustomRecord("Synthetic user record at rollback boundary");
        try { doc.Range(0, 0).InsertBefore(UserText); }
        finally { app.UndoRecord.EndCustomRecord(); }
        colorBoundaryInserted = true;
        Seam((int)app.UndoRecord.CustomRecordLevel == 0, "The new user Undo record is independently ended before the final guard.");
    }

    private static void UnknownConsumed(ColorCase fixture, ColorSampler sampler, IDictionary<string, object> result, string address, object marking, string publicId)
    {
        Check(UiState.Flag(result, "isError"), "Interleaved native input returns an error instead of claiming a commit.");
        var error = UiState.Map(UiState.Map(result["structuredContent"])["error"]);
        var details = UiState.Map(error["details"]);
        Check(UiState.Text(error, "code") == "outcome_unknown" && UiState.Flag(details, "operation_started") && UiState.Flag(details, "outcome_unknown"),
            "An unconfirmed native effect is explicitly started and uncertain.");
        Check((int)fixture.App.UndoRecord.CustomRecordLevel == 0, "Uncertain coloring closes only its own custom record.");
        string request = NativeMcpClient.ForAddress(address).NativeRequestId(publicId);
        Check(fixture.Native.Policy.HighlightReceipts.Single(receipt => receipt.Request.RequestId == request).Outcome == "outcome_uncertain" && !fixture.Native.Policy.IsShared,
            "The original native identity remains consumed and its old edit authority is revoked.");
        byte[] saved = ProtectedData.Unprotect(File.ReadAllBytes(fixture.Sidecar), null, DataProtectionScope.CurrentUser);
        var privateState = McpEndpoint.Json().Deserialize<SavedWordReview>(System.Text.Encoding.UTF8.GetString(saved));
        Check(privateState.Policy.HighlightReceipts.Single(receipt => receipt.Request.RequestId == request).Outcome == "outcome_uncertain",
            "Consumed uncertainty is retained in the isolated durable receipt journal.");
        var observed = new ColorObservation(fixture);
        int samples = sampler.Calls;
        for (int retry = 0; retry < 2; retry++)
        {
            var repeat = Rpc(address, "highlight_text", marking);
            var repeatedError = UiState.Map(UiState.Map(repeat["structuredContent"])["error"]);
            Check(UiState.Flag(repeat, "isError") && UiState.Text(repeatedError, "code") == "outcome_unknown", "Exact uncertainty lookup cannot redispatch the edit.");
            Check(sampler.Calls == samples, "Exact consumed uncertainty lookup bypasses every new input-sampling boundary.");
            observed.Verify(fixture, "Exact uncertain receipt lookup");
        }
    }

    private static void Applied(ColorCase fixture, IDictionary<string, object> result)
    {
        Check(!UiState.Flag(result, "isError") && fixture.Native.Policy.Marks.Count == 1 && fixture.Native.Policy.HighlightReceipts.Count == 1,
            "Quiet native coloring commits exactly one mark and one receipt.");
        Check(fixture.Native.Policy.HighlightReceipts[0].Outcome == "applied" && NativePainted(fixture.Document, 0, 3), "The successful receipt corresponds to actual Word paint.");
        Check(WordSnapshot.Capture(fixture.Document).FullText == fixture.Original, "Color-only success preserves complete wording.");
    }

    private static void KnownBusy(IDictionary<string, object> result)
    {
        Check(UiState.Flag(result, "isError"), "Known busy refuses coloring.");
        var error = UiState.Map(UiState.Map(result["structuredContent"])["error"]); var details = UiState.Map(error["details"]);
        Check(UiState.Text(error, "code") == "word_busy" && details.ContainsKey("operation_started") && !UiState.Flag(details, "operation_started") &&
            details.ContainsKey("outcome_unknown") && !UiState.Flag(details, "outcome_unknown"), "Known busy explicitly reports no started or unknown operation.");
    }

    private static bool NativePainted(dynamic document, int from, int to)
    {
        var color = System.Drawing.ColorTranslator.FromHtml(InputColor);
        int expected = color.R | color.G << 8 | color.B << 16;
        return (int)document.Range(from, to).Font.Shading.BackgroundPatternColor == expected;
    }

    private static string ReceiptDigest(string path) { return File.Exists(path) ? FileDigest(path) : "missing"; }
    private static bool Selected(string selected, string name) { return selected == "all" || selected == name; }
    private static void Record(string name, int samples) { colorResults.Add(new { name = name, status = "PASS", injectedSamples = samples }); }
    private static void Seam(bool condition, string label)
    {
        colorSeams++;
        if (!condition) throw new Exception("FIXTURE_SEAM: " + label);
        Console.WriteLine("SEAM: " + label);
    }
}
