using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using UnderTree.WordAddin;

internal static partial class NativeReviewSmoke
{
    [DllImport("user32.dll", EntryPoint = "GetForegroundWindow")]
    private static extern IntPtr CliForegroundWindow();

    private static Dictionary<string, object> Automation(string command, string id = null, Dictionary<string, object> args = null)
    {
        args = args ?? Values();
        args["command"] = command;
        if (id != null) args["documentId"] = id;
        return Local("automation", args);
    }

    private static void AutomationDenied(string command, string id, Dictionary<string, object> args, string reason)
    {
        bool denied = false;
        try { Automation(command, id, args); }
        catch (Exception error)
        {
            if (!(error is ArgumentException) && !(error is InvalidOperationException)) throw;
            if (!error.Message.Contains(reason)) throw;
            denied = true;
        }
        Check(denied, "CLI automation refuses " + reason + ".");
    }

    private static void CliAutomation()
    {
        string paragraph = "A " + Char.ConvertFromUtf32(0x1F600) + " rough paragraph.";
        NewFixture("cli-target", paragraph);
        dynamic target = document;
        string targetId = documentId;
        string targetText = (string)target.Content.Text;
        string targetPath = (string)target.FullName;
        string targetHash = FileHash(targetPath);
        document = null;
        NewFixture("cli-neighbor", "This other document must not change.");
        string neighborId = documentId;
        string neighborText = (string)document.Content.Text;
        document.Range(3, 7).Select();
        int selectionStart = (int)application.Selection.Start, selectionEnd = (int)application.Selection.End;
        IntPtr foreground = CliForegroundWindow();

        var diagnostics = Local("connection_diagnostics");
        Check(Object.Equals(diagnostics["active"], true) && Obj(diagnostics["counters"]).Count == 9,
            "Local diagnostics expose the current endpoint counters through the native dispatcher.");
        Check((string)target.Content.Text == targetText && (string)document.Content.Text == neighborText &&
            WordSnapshot.Identity((object)application.ActiveDocument) == WordSnapshot.Identity((object)document) &&
            (int)application.Selection.Start == selectionStart && (int)application.Selection.End == selectionEnd && CliForegroundWindow() == foreground,
            "Reading local diagnostics preserves document text, active selection and desktop foreground.");

        var listed = Rows(Automation("documents")["documents"]);
        Check(listed.Length == 2 && listed.Any(row => Text(row, "documentId") == targetId), "CLI discovery lists exact open document identities.");
        AutomationDenied("state", null, Values(), "Missing documentId");
        AutomationDenied("state", "missing-document", Values(), "no longer open");
        AutomationDenied("select_mark", targetId, Values(), "Unsupported automation command");
        var state = Automation("state", targetId);
        Check(Text(Obj(state["document"]), "id") == targetId, "CLI state addresses an inactive document explicitly.");
        Check(!editor.BackgroundReviewDocuments().Any(candidate => WordSnapshot.Identity(candidate) == WordSnapshot.Identity((object)target)),
            "CLI reads leave an unchosen background document outside autosave eligibility.");
        string snapshot = Text(Obj(state["document"]), "revision");
        var rule = Values("id", "cli_prompt", "label", "CLI prompt", "color", "#AACCEE", "action", "prompt", "prompt", "Shorten the wording.", "snapshot", snapshot);
        AutomationDenied("save_rule", targetId, Values("snapshot", "obsolete"), "changed");
        Check(!editor.BackgroundReviewDocuments().Any(candidate => WordSnapshot.Identity(candidate) == WordSnapshot.Identity((object)target)),
            "A refused CLI mutation does not enroll its background target for autosave.");
        state = Automation("save_rule", targetId, rule);
        Check(Rows(state["rules"]).Any(row => Text(row, "id") == "cli_prompt"), "CLI rule save uses the native local policy owner.");
        Check(!Rows(Automation("state", neighborId)["rules"]).Any(row => Text(row, "id") == "cli_prompt"), "CLI rule save leaves the active neighbor's rules unchanged.");
        Check(editor.BackgroundReviewDocuments().Any(candidate => WordSnapshot.Identity(candidate) == WordSnapshot.Identity((object)target)) &&
            !(bool)Obj(state["document"])["shared"], "A successful local CLI mutation admits its background target to autosave without sharing MCP authority.");

        var context = Automation("context_preview", targetId);
        int from = Text(context, "text").IndexOf(paragraph, StringComparison.Ordinal);
        var mark = Values("ruleId", "cli_prompt", "from", from, "to", from + paragraph.Length, "quote", paragraph, "snapshot", context["snapshot"]);
        var wrong = new Dictionary<string, object>(mark); wrong["quote"] = "wrong";
        AutomationDenied("mark_range", targetId, wrong, "wording changed");
        Check(Rows(Automation("state", targetId)["marks"]).Length == 0, "A wrong CLI quote adds no native marker.");
        var split = new Dictionary<string, object>(mark);
        split["from"] = from + 3; split["to"] = from + 4; split["quote"] = paragraph.Substring(3, 1);
        AutomationDenied("mark_range", targetId, split, "Unicode");
        state = Automation("mark_range", targetId, mark);
        var marks = Rows(state["marks"]);
        Check(marks.Length == 1 && Text(marks[0], "quote") == paragraph && !(bool)marks[0]["stale"], "CLI range marking retains exact Unicode wording and a current native anchor.");
        Check((string)target.Content.Text == targetText && (int)target.Revisions.Count == 0, "Local CLI rule/marker operations do not rewrite wording or add revisions.");
        Check((string)document.Content.Text == neighborText && (bool)document.Saved && FileHash(savedPath) == savedHash, "CLI background operations preserve the neighbor and its saved bytes.");
        Check(WordSnapshot.Identity((object)application.ActiveDocument) == WordSnapshot.Identity((object)document) &&
            (int)application.Selection.Start == selectionStart && (int)application.Selection.End == selectionEnd && CliForegroundWindow() == foreground,
            "CLI operations preserve the active document, native selection and desktop foreground.");

        using (var autosave = new WordAutoSave((object)application, () => true,
            candidate => editor.RefreshReviewDocument(candidate), message => { throw new InvalidOperationException(message); },
            () => editor.BackgroundReviewDocuments()))
        {
            DateTime now = DateTime.UtcNow;
            autosave.Pulse(now);
            autosave.Pulse(now.AddSeconds(2));
        }
        Check((bool)target.Saved && FileHash(targetPath) != targetHash && (string)target.Content.Text == targetText,
            "The native background autosave port persists local-only CLI changes without an MCP read or wording change.");
        Check((bool)document.Saved && FileHash(savedPath) == savedHash,
            "Local-only CLI autosave leaves the active neighbor's saved bytes unchanged.");

        AutomationDenied("remove_mark", targetId, Values("id", marks[0]["id"], "snapshot", "obsolete"), "changed");
        state = Automation("remove_mark", targetId, Values("id", marks[0]["id"], "snapshot", Obj(state["document"])["revision"]));
        Check(Rows(state["marks"]).Length == 0, "CLI removes only the explicitly addressed native marker.");
        AutomationDenied("mark_range", targetId, mark, "changed");
        Check(Text(Obj(Local("state")["document"]), "id") == neighborId, "CLI target scope is cleared before ordinary ribbon/local operations.");
        target.Protect(3);
        state = Automation("state", targetId);
        rule["snapshot"] = Obj(state["document"])["revision"];
        AutomationDenied("save_rule", targetId, rule, "read-only or protected");
        target.Unprotect();
        string originalPath = (string)target.FullName;
        target.Save();
        target.SaveAs2(System.IO.Path.Combine(output, "cli-duplicate.docx"), 16);
        dynamic duplicate = application.Documents.Open(originalPath);
        AutomationDenied("state", targetId, Values(), "Another open copy");
        duplicate.Close(0);
        target.Close(0);
        AutomationDenied("state", targetId, Values(), "no longer open");
        Check((string)document.Content.Text == neighborText, "A closed CLI target never falls back to the active document.");
    }
}
