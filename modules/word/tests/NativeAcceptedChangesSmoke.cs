using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using UnderTree.WordAddin;

internal static class NativeAcceptedChangesSmoke
{
    private delegate bool Enumerate(IntPtr window, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumWindows(Enumerate callback, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr window, Enumerate callback, IntPtr state);
    [DllImport("user32.dll")] private static extern int GetClassName(IntPtr window, StringBuilder name, int length);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object value);
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 16777216 };
    private static readonly List<string> checks = new List<string>();
    private static readonly List<string> documentIds = new List<string>();
    private static dynamic application, document;
    private static WordEditor editor;
    private static string output, documentId, savedPath, savedHash;
    private const string Original = "A rough paragraph repeats the same point with unnecessary wording.";
    private const string Revised = "A clear paragraph makes its point once.";
    private const string Outside = "Unmarked tail stays exactly as written.";
    private const string Rationale = "Native acceptance qualification.";
    private const int Green = 10795380;

    [STAThread]
    public static int Main(string[] args)
    {
        Process owned = null;
        WordTestAutomation filter = null;
        try
        {
            filter = new WordTestAutomation();
            output = Path.GetFullPath(args[1]); Directory.CreateDirectory(output);
            string bootstrap = Path.Combine(output, "accepted-bootstrap.rtf");
            File.WriteAllText(bootstrap, @"{\rtf1\ansi Disposable acceptance bootstrap.\par}");
            owned = Process.Start(new ProcessStartInfo(args[0], "/a /x /q \"" + bootstrap + "\"") { UseShellExecute = false, WindowStyle = ProcessWindowStyle.Hidden });
            application = AttachOwned(owned); application.Visible = false; application.DisplayAlerts = 0;
            Check((int)application.Documents.Count == 1 && (string)application.ActiveDocument.FullName == bootstrap,
                "The owned isolated Word process opened only the disposable bootstrap.");
            application.ActiveDocument.Close(0);
            editor = new WordEditor((object)application, action => action(), Path.Combine(output, "profile"), Path.Combine(output, "unused-zrok.exe"));
            Local("connection_local");
            string scenario = args.Length < 3 ? "all" : args[2];
            if (scenario == "all")
            {
                LegacyDirectColorUsesNativeAcceptance();
                TrackedAcceptanceAndRejection();
                ProtectedUnitsAndLocalRules();
            }
            if (scenario == "all" || scenario == "unresolved-only") UnresolvedProtectionBeforeAcceptance();
            if (scenario == "all" || scenario == "rejection-only") RejectedPromptRestoresNativeIdentity();
            WriteResult("PASS", null);
            Console.WriteLine("PASS " + checks.Count + " native accepted-wording assertions.");
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

    private static void LegacyDirectColorUsesNativeAcceptance()
    {
        NewFixture("direct", Original);
        SaveRule("blue", "prompt", "direct", "#AACCEE"); Mark(Original, "blue");
        string before = WordSnapshot.Capture(document).MainText;
        var request = Request("direct-green", Read(), Original, Revised);
        var response = LegacyTracked(request);
        Check(Text(response, "status") == "tracked" && !(bool)response["applied"] && (bool)response["requiresApproval"] &&
            (int)document.Revisions.Count == 2 && AcceptedMarks().Length == 0,
            "A historical tracked receipt over a direct-colored passage stays unprotected pending native approval.");
        Check(WordSnapshot.Capture(document).MainText == before.Replace(Original, Revised), "The suggestion changes only its authorized paragraph.");
        WordSnapshot pendingSnapshot = WordSnapshot.Capture(document);
        Check((int)document.Comments.Count == 1 && pendingSnapshot.Review.feedback.Single().author == "wordplai" &&
            pendingSnapshot.Review.feedback.Single().text == Rationale,
            "The direct-colored suggestion creates exactly one concise native rationale comment.");
        string pending = new WordNativeEditState().NativeIdentity(document);
        Check(Text(Review("suggest_text", request), "status") == "tracked" && new WordNativeEditState().NativeIdentity(document) == pending,
            "Replaying a pending suggestion duplicates neither revisions nor its rationale comment.");
        Check((bool)document.Undo(1), "One Undo reverses the complete tracked suggestion together with its rationale comment.");
        Local("state");
        Check(WordSnapshot.Capture(document).MainText == before && (int)document.Revisions.Count == 0 &&
            (int)document.Comments.Count == 0 && AcceptedMarks().Length == 0,
            "Suggestion Undo restores original text, removes the rationale and adds no green protection.");
        Check((bool)document.Redo(1), "One Redo restores the tracked suggestion and its rationale comment.");
        Local("state");
        Check(WordSnapshot.Capture(document).MainText == before.Replace(Original, Revised) && (int)document.Revisions.Count == 2 &&
            (int)document.Comments.Count == 1 && AcceptedMarks().Length == 0,
            "Suggestion Redo restores the complete reviewable pair without accepting it.");
        document.Revisions.AcceptAll(); Local("state");
        var green = AcceptedMarks();
        Check(green.Length == 1 && Text(green[0], "quote") == Revised && (int)document.Revisions.Count == 0,
            "Native approval turns the replacement into one green protected passage.");
        CheckGreen(Revised);
        DenyChange(Revised, "Changed accepted wording.");
        string after = new WordNativeEditState().NativeIdentity(document);
        Check(Text(Review("suggest_text", request), "status") == "applied" && new WordNativeEditState().NativeIdentity(document) == after,
            "Replaying the accepted request never repeats editing, rationale comments or protection markers.");
        Check(FileHash(savedPath) == savedHash && !(bool)document.Saved, "Source authority without the installed autosave timer leaves saved bytes unchanged.");
        Check((bool)document.Undo(1), "One Undo removes the separate automatic green annotation after native approval.");
        Local("state");
        Check(WordSnapshot.Capture(document).MainText == before.Replace(Original, Revised) && AcceptedMarks().Length == 0,
            "Undoing green protection retains the locally accepted wording and its original rationale.");
        Check((bool)document.Redo(1), "One Redo restores the accepted protection.");
        Local("state"); CheckGreen(Revised); DenyChange(Revised, "Another unauthorized change.");
        string[] ids = AcceptedMarks().Select(mark => Text(mark, "id")).ToArray();
        document.Save(); document.Close(0); document = application.Documents.Open(savedPath, Visible: false); document.Activate();
        Check(Text(Map(Local("state")["document"]), "id") == documentId && AcceptedMarks().Select(mark => Text(mark, "id")).SequenceEqual(ids),
            "An explicit native save and reopen retain the exact green protected identities.");
        CheckGreen(Revised); DenyChange(Revised, "A reopened unauthorized change.");
    }

    private static void TrackedAcceptanceAndRejection()
    {
        NewFixture("tracked-accept", Original); Mark(Original, "snappier");
        var response = LegacyTracked(Request("tracked-green", Read(), Original, Revised));
        Check(Text(response, "status") == "tracked" && (int)document.Revisions.Count == 2 && AcceptedMarks().Length == 0,
            "A pending tracked replacement remains editable review text and has no accepted green protection.");
        document.Revisions.Item(1).Accept(); Local("state");
        Check((int)document.Revisions.Count == 1 && AcceptedMarks().Length == 0, "Partial native acceptance adds no green permission or protection metadata.");
        Check((bool)document.Undo(1), "Undo restores the complete native review pair after partial acceptance."); Local("state");
        document.Revisions.AcceptAll(); Local("state");
        Check((int)document.Revisions.Count == 0 && AcceptedMarks().Length == 1, "Exact complete native acceptance automatically adds green protection.");
        CheckGreen(Revised); DenyChange(Revised, "Changed after native acceptance.");
        string protectedNative = new WordNativeEditState().NativeIdentity(document);
        Local("state"); Read();
        Check(new WordNativeEditState().NativeIdentity(document) == protectedNative, "Repeated reads and refreshes never create duplicate acceptance annotations.");
        Check((bool)document.Undo(1), "The acceptance annotation has a native Undo record."); Local("state");
        Check(WordSnapshot.Capture(document).MainText.Contains(Revised) && AcceptedMarks().Length == 0,
            "Undoing the green annotation preserves accepted text and does not immediately reapply it.");
        Check((bool)document.Redo(1), "Redo restores the accepted green annotation."); Local("state"); CheckGreen(Revised);

        NewFixture("tracked-reject", Original); Mark(Original, "snappier");
        LegacyTracked(Request("reject-green", Read(), Original, Revised));
        document.Revisions.RejectAll(); Local("state");
        Check(WordSnapshot.Capture(document).MainText.Contains(Original) && (int)document.Revisions.Count == 0 && AcceptedMarks().Length == 0,
            "Native Reject restores original wording without green protection.");

        NewFixture("tracked-mixed", Original); Mark(Original, "snappier");
        LegacyTracked(Request("mixed-green", Read(), Original, Revised));
        document.Revisions.Item(1).Accept(); document.Revisions.Item(1).Reject(); Local("state");
        Check((int)document.Revisions.Count == 0 && AcceptedMarks().Length == 0,
            "An ambiguous mixed Accept/Reject result is never recognized as a completed acceptance.");
    }

    private static void ProtectedUnitsAndLocalRules()
    {
        const string source = "One [A]. Two [B].";
        NewFixture("native-units", source); SaveRule("blue", "prompt", "direct", "#AACCEE"); Mark(source, "blue");
        WordSnapshot before = WordSnapshot.Capture(document);
        int first = before.MainText.IndexOf("[A]", StringComparison.Ordinal);
        document.Hyperlinks.Add(before.Range(document, first, first + 3), "https://example.invalid/citation-a");
        Mark("[A]", "protect"); Mark("[B]", "reject");
        var units = Records(Local("state")["marks"]).Where(mark => Text(mark, "ruleId") == "protect" || Text(mark, "ruleId") == "reject").ToArray();
        var signatures = units.ToDictionary(mark => Text(mark, "id"), mark => WordSnapshot.FormattingSignature(document.Bookmarks.Item(Text(mark, "id")).Range));
        string a = Text(units.Single(mark => Text(mark, "quote") == "[A]"), "id"), b = Text(units.Single(mark => Text(mark, "quote") == "[B]"), "id");
        var request = Request("direct-native-green", Read(), source, "");
        request["replacementParts"] = new object[] { Values("text", "Both "), Values("protectedId", b), Values("text", " then "), Values("protectedId", a), Values("text", ".") };
        var response = LegacyTracked(request);
        document.Activate();
        Check(Text(response, "status") == "tracked" && (int)document.Revisions.Count == 2 && AcceptedMarks().Length == 1,
            "A historical protected-unit rewrite remains pending with only the original protected unit green.");
        document.Revisions.AcceptAll(); Local("state");
        Check(WordSnapshot.Capture(document).MainText.Contains("Both [B] then [A].") && (int)document.Hyperlinks.Count == 1,
            "Native acceptance preserves rearranged protected wording and its native hyperlink.");
        foreach (var unit in units)
        {
            string id = Text(unit, "id");
            Check((string)document.Bookmarks.Item(id).Range.Text == Text(unit, "quote") && WordSnapshot.FormattingSignature(document.Bookmarks.Item(id).Range) == signatures[id],
                "A retained protected unit keeps its exact identity, text and native formatting.");
        }
        Check(AcceptedMarks().Length == 4 && Records(Local("state")["marks"]).Count(mark => Text(mark, "ruleId") == "reject") == 1,
            "Only surrounding accepted gaps become green; existing separate protected units never merge or overlap.");
        DenyChange("Both [B] then [A].", "An unauthorized rewrite.");

        NewFixture("custom-rule", Original);
        SaveRule("protect", "prompt", "direct", "#74b9a4"); SaveRule("blue", "prompt", "direct", "#AACCEE"); Mark(Original, "blue");
        LegacyTracked(Request("custom-rule-green", Read(), Original, Revised));
        Check(AcceptedMarks().Length == 0 && (int)document.Revisions.Count == 2,
            "A repurposed green prompt color never accepts a pending remote suggestion.");
        document.Revisions.AcceptAll(); Local("state");
        var rules = Records(Local("state")["rules"]);
        Check(Text(rules.Single(rule => Text(rule, "id") == "protect"), "action") == "prompt" &&
            rules.Any(rule => Text(rule, "id") != "protect" && Text(rule, "action") == "protect" && Text(rule, "color").ToLowerInvariant() == "#74b9a4"),
            "Automatic accepted protection adds a separate green rule instead of overriding the user's repurposed color rule.");
        CheckGreen(Revised); DenyChange(Revised, "The custom prompt cannot bypass protection.");
    }

    private static void UnresolvedProtectionBeforeAcceptance()
    {
        const string source = "One [A]. Two [B].";
        NewFixture("unresolved-protection", source);
        Mark(source, "snappier"); Mark("[A]", "protect"); Mark("[B]", "reject");
        var marks = Records(Local("state")["marks"]);
        string a = Text(marks.Single(mark => Text(mark, "quote") == "[A]"), "id");
        string b = Text(marks.Single(mark => Text(mark, "quote") == "[B]"), "id");
        var request = Request("unresolved-accepted-green", Read(), source, "");
        request["replacementParts"] = new object[] { Values("protectedId", a), Values("text", " before "), Values("protectedId", b), Values("text", ".") };
        LegacyTracked(request); document.Activate(); document.Revisions.AcceptAll();
        dynamic protectedRange = document.Bookmarks.Item(b).Range.Duplicate;
        string signature = WordSnapshot.FormattingSignature(protectedRange);
        int red = (int)protectedRange.Font.Shading.BackgroundPatternColor;
        document.Bookmarks.Item(b).Delete();
        string beforeRefresh = new WordNativeEditState().NativeIdentity(document);
        var state = Local("state");
        Check(Records(state["marks"]).Any(mark => Text(mark, "id") == b && (bool)mark["stale"]),
            "A removed protected native identity is reported unresolved after native acceptance.");
        Check(new WordNativeEditState().NativeIdentity(document) == beforeRefresh &&
            WordSnapshot.FormattingSignature(protectedRange) == signature && (int)protectedRange.Font.Shading.BackgroundPatternColor == red,
            "Automatic accepted protection never repaints or mutates an unresolved protected native unit.");
        Check(AcceptedMarks().Length == 1 && WordSnapshot.Capture(document).MainText.Contains("[A] before [B].") && (int)document.Revisions.Count == 0,
            "Acceptance remains unannotated when protection is unresolved while existing green and red units stay intact.");
    }

    private static void RejectedPromptRestoresNativeIdentity()
    {
        NewFixture("rejected-original-prompt", Original); Mark(Original, "snappier");
        string promptId = Text(Records(Local("state")["marks"]).Single(), "id");
        WordSnapshot before = WordSnapshot.Capture(document);
        int from = before.MainText.IndexOf(Original, StringComparison.Ordinal);
        document.Comments.Add(before.Range(document, from, from + Original.Length), "Keep the original meaning when revising.");
        string feedback = UserFeedbackIdentity();
        LegacyTracked(Request("plain-rejected-original", Read(), Original, Revised));
        document.Revisions.RejectAll();
        var context = Read();
        var restored = Records(context["highlights"]).Single(mark => Text(mark, "id") == promptId);
        Check(!(bool)restored["stale"] && Text(Records(Local("state")["marks"]).Single(), "quote") == Original,
            "Plain native Reject restores the same original prompt identity and current quote without reapplying its color.");
        Check(WordSnapshot.Capture(document).MainText == before.MainText && UserFeedbackIdentity() == feedback &&
            (int)document.Comments.Count == 1 && AcceptedMarks().Length == 0,
            "Plain native Reject preserves original source and user feedback, removes the rationale on rejected inserted text and adds no green protection.");
        var response = LegacyTracked(Request("plain-rejected-followup", context, Original, Revised));
        Check(Text(response, "status") == "tracked" && WordSnapshot.Capture(document).MainText.Contains(Revised) && (int)document.Revisions.Count > 0,
            "A fresh complete read permits another historical tracked application after Reject with no manual marker repair.");
        document.Revisions.RejectAll(); Read();
        Check(Records(Local("state")["marks"]).Any(mark => Text(mark, "id") == promptId && !(bool)mark["stale"]),
            "Repeated plain rejection retains the original prompt identity without accumulating stale markers.");

        NewFixture("rejected-native-deletion", Original); Mark(Original, "discard");
        promptId = Text(Records(Local("state")["marks"]).Single(), "id");
        var deletion = Review("suggest_text", Request("native-deletion-reject", Read(), Original, ""));
        Check(Text(deletion, "status") == "tracked" && (int)document.Revisions.Count == 1,
            "A complete removal retains a native reviewable deletion with original marker evidence.");
        document.Revisions.RejectAll(); context = Read();
        Check(Records(context["highlights"]).Any(mark => Text(mark, "id") == promptId && !(bool)mark["stale"]) && WordSnapshot.Capture(document).MainText.Contains(Original),
            "Rejecting complete removal restores the same current discard marker without manual reapplication.");
        Review("suggest_text", Request("native-deletion-accept", context, Original, ""));
        document.Revisions.AcceptAll(); context = Read();
        Check(!Records(context["highlights"]).Any(mark => Text(mark, "id") == promptId) &&
            !WordSnapshot.Capture(document).MainText.Contains(Original) && AcceptedMarks().Length == 0,
            "Accepting complete removal retires its deleted marker without an invalid empty mark or green passage.");
        Mark(Outside, "snappier");
        var next = Review("suggest_text", Request("after-native-deletion", Read(), Outside, "The remaining paragraph stays available."));
        Check(Text(next, "status") == "applied" && (bool)next["applied"] && !(bool)next["requiresApproval"] && (int)document.Revisions.Count == 0,
            "An accepted removal leaves no stale marker blocking a new authorized direct paragraph edit.");

        NewFixture("deletion-later-protect", Original); Mark(Original, "discard");
        promptId = Text(Records(Local("state")["marks"]).Single(), "id");
        Review("suggest_text", Request("deletion-before-protect", Read(), Original, ""));
        SaveRule("discard", "protect", "suggestions", "#bc8fc7");
        document.Revisions.AcceptAll(); context = Read();
        Check(Records(context["highlights"]).Any(mark => Text(mark, "id") == promptId && (bool)mark["stale"]) &&
            Text(Records(Local("state")["rules"]).Single(rule => Text(rule, "id") == "discard"), "action") == "protect",
            "Accepted deletion cannot retire or weaken a later local Protect rule; its unresolved original marker fails closed.");
        Mark(Outside, "snappier"); DenyChange(Outside, "A stale protection cannot be bypassed.");

        NewFixture("deletion-explicit-removal", Original); Mark(Original, "discard");
        promptId = Text(Records(Local("state")["marks"]).Single(), "id");
        Review("suggest_text", Request("deletion-before-local-remove", Read(), Original, ""));
        Local("remove_mark", Values("id", promptId)); document.Revisions.RejectAll(); context = Read();
        Check(!Records(context["highlights"]).Any(mark => Text(mark, "id") == promptId) && WordSnapshot.Capture(document).MainText.Contains(Original),
            "Reject restores deleted words while preserving an explicit later local marker removal.");
        DenyChange(Original, "A removed marker cannot regain edit authority.");
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
            "from", from, "to", from + quote.Length, "quote", quote, "replacement", replacement, "rationale", Rationale);
    }

    private static void DenyChange(string quote, string replacement)
    {
        var request = Request("denied_" + Guid.NewGuid().ToString("N"), Read(), quote, replacement);
        string before = new WordNativeEditState().NativeIdentity(document); bool denied = false;
        try { Review("suggest_text", request); } catch (McpOperationException) { denied = true; }
        Check(denied && new WordNativeEditState().NativeIdentity(document) == before,
            "Further AI wording changes to accepted green text are denied without native mutation.");
    }

    private static Dictionary<string, object>[] AcceptedMarks()
    {
        var state = Local("state"); var rules = Records(state["rules"]);
        var greenRules = rules.Where(rule => Text(rule, "action") == "protect" && Text(rule, "color").ToLowerInvariant() == "#74b9a4").Select(rule => Text(rule, "id")).ToArray();
        return Records(state["marks"]).Where(mark => greenRules.Contains(Text(mark, "ruleId")) && !(bool)mark["stale"]).ToArray();
    }

    private static void CheckGreen(string quote)
    {
        WordSnapshot snap = WordSnapshot.Capture(document); int from = snap.MainText.IndexOf(quote, StringComparison.Ordinal);
        dynamic range = snap.Range(document, from, from + quote.Length);
        Check((int)range.Font.Shading.BackgroundPatternColor == Green && (int)range.HighlightColorIndex == 0,
            "Accepted replacement text is visibly green in the actual Word range.");
    }

    private static string UserFeedbackIdentity()
    { WordSnapshot snapshot = WordSnapshot.Capture(document); return Json.Serialize(snapshot.Review.feedback.Where(comment => comment.author != "wordplai" || comment.text != Rationale).Select(CommentContent).ToArray()); }

    private static object CommentContent(WordCommentContext comment)
    { return new { comment.storyType, comment.author, comment.text, comment.date, comment.resolved, replies = comment.replies.Select(CommentContent).ToArray() }; }

    private static Dictionary<string, object> Local(string operation, Dictionary<string, object> args = null)
    { return Map(editor.Dispatch(operation, args ?? Values())); }
    private static Dictionary<string, object> LegacyTracked(Dictionary<string, object> request)
    {
        NativeMcpClient.ApplyLegacyTracked(editor, (object)document, request);
        return Review("suggest_text", request);
    }
    private static Dictionary<string, object> Review(string tool, Dictionary<string, object> args)
    {
        try { return Map(typeof(WordEditor).GetMethod("Review", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(editor, new object[] { tool, args })); }
        catch (TargetInvocationException error) { throw error.InnerException; }
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
    { if (output != null) File.WriteAllText(Path.Combine(output, "result.json"), Json.Serialize(new { status = status, checks = checks, error = error, installedPackage = "NOT RUN", normalWord = "NOT TOUCHED", publicTunnel = "NOT RUN" })); }
    private static Dictionary<string, object> Values(params object[] pairs)
    { var result = new Dictionary<string, object>(); for (int index = 0; index < pairs.Length; index += 2) result.Add((string)pairs[index], pairs[index + 1]); return result; }
    private static Dictionary<string, object> Map(object value) { return (Dictionary<string, object>)Json.DeserializeObject(Json.Serialize(value)); }
    private static Dictionary<string, object>[] Records(object value) { return ((object[])value).Select(Map).ToArray(); }
    private static string Text(Dictionary<string, object> value, string key) { return (string)value[key]; }
    private static string FileHash(string path) { using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (SHA256 sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(stream)); }
    private static void Check(bool condition, string message) { if (!condition) throw new Exception(message); checks.Add(message); Console.WriteLine("PASS: " + message); }
}
