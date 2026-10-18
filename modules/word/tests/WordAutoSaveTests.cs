using System;
using System.Collections.Generic;
using System.Linq;
using UnderTree.WordAddin;

public sealed class AutoSaveTestCount { public int Count { get; set; } }
public sealed class AutoSaveTestRange { public int End { get; set; } public string Text { get; set; } }
public sealed class AutoSaveTestUndo { public int CustomRecordLevel { get; set; } }
public sealed class AutoSaveTestDocument
{
    public bool Saved { get; set; }
    public bool ReadOnly { get; set; }
    public int ProtectionType { get; set; }
    public string Path { get; set; }
    public AutoSaveTestRange Content { get; set; }
    public AutoSaveTestCount Revisions { get; set; }
    public AutoSaveTestCount Comments { get; set; }
    public int Saves;
    public bool SaveFails, SaveUnconfirmed;
    public Action BeforeSave;
    public void Save()
    {
        Saves++;
        if (BeforeSave != null) BeforeSave();
        if (SaveFails) throw new InvalidOperationException("The file is unavailable.");
        if (!SaveUnconfirmed) Saved = true;
    }
}
public sealed class AutoSaveTestApplication
{
    public AutoSaveTestCount Documents { get; set; }
    public AutoSaveTestDocument ActiveDocument { get; set; }
    public AutoSaveTestUndo UndoRecord { get; set; }
    public int BackgroundSavingStatus { get; set; }
}

internal static class WordAutoSaveTests
{
    private static int checks;
    private static readonly DateTime Start = new DateTime(2026, 10, 17, 12, 0, 0, DateTimeKind.Utc);

    [STAThread]
    public static int Main()
    {
        try
        {
            IdleSaveAndCurrentRefresh();
            NativeChangesAndUnnamedDocument();
            BusyReadOnlyAndProtection();
            FailureAndRecovery();
            DocumentSwitchAndDisposal();
            RefreshChangesAreRechecked();
            ActiveAndBackgroundDocumentsSaveIndependently();
            BackgroundTargetClosureAndSwitch();
            BackgroundTargetIsRecheckedAfterRefresh();
            SeveralBackgroundDocumentsMakeProgress();
            BackgroundRemovalDuringRefresh();
            BackgroundDocumentLimit();
            Console.WriteLine("PASS " + checks + " automatic-save assertions; isolated Word object doubles only.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static void IdleSaveAndCurrentRefresh()
    {
        var application = Application();
        int refreshes = 0;
        var errors = new List<string>();
        using (var automatic = new WordAutoSave(application, () => true, () => refreshes++, errors.Add))
        {
            application.ActiveDocument.BeforeSave = () => Check(refreshes > 0, "Native permissions refresh before a document save.");
            automatic.Pulse(Start);
            automatic.Pulse(Start.AddMilliseconds(749));
            Check(application.ActiveDocument.Saves == 0 && refreshes == 0, "Rapid activity coalesces without saving or rebuilding context.");
            automatic.Changed(Start.AddMilliseconds(600));
            automatic.Pulse(Start.AddMilliseconds(1000));
            Check(application.ActiveDocument.Saves == 0, "Selection activity extends the idle debounce.");
            automatic.Pulse(Start.AddMilliseconds(1350));
            Check(application.ActiveDocument.Saves == 1 && application.ActiveDocument.Saved && refreshes == 1, "An idle existing named document refreshes and saves once.");
            automatic.Pulse(Start.AddSeconds(5));
            Check(application.ActiveDocument.Saves == 1 && refreshes == 1 && errors.Count == 0, "An unchanged clean document does no repeated native refresh or save work.");
        }
    }

    private static void NativeChangesAndUnnamedDocument()
    {
        var application = Application();
        application.ActiveDocument.Path = "";
        int refreshes = 0;
        string observedText = null;
        using (var automatic = new WordAutoSave(application, () => true, () => { refreshes++; observedText = application.ActiveDocument.Content.Text; }, message => { throw new Exception(message); }))
        {
            automatic.Pulse(Start);
            automatic.Pulse(Start.AddSeconds(1));
            automatic.Pulse(Start.AddSeconds(2));
            Check(refreshes == 1 && application.ActiveDocument.Saves == 0 && !application.ActiveDocument.Saved,
                "An unnamed document receives live context without creating a path or opening Save As.");
            application.ActiveDocument.Revisions.Count = 1;
            automatic.Pulse(Start.AddSeconds(3));
            automatic.Pulse(Start.AddSeconds(4));
            Check(refreshes == 2, "Native revision changes refresh even while the new document remains dirty.");
            application.ActiveDocument.Comments.Count = 1;
            automatic.Pulse(Start.AddSeconds(5));
            automatic.Pulse(Start.AddSeconds(6));
            Check(refreshes == 3, "Native comment changes refresh without manual Refresh.");
            application.ActiveDocument.Content.End++;
            automatic.Pulse(Start.AddSeconds(7));
            automatic.Pulse(Start.AddSeconds(8));
            Check(refreshes == 4 && application.ActiveDocument.Saves == 0, "Typing is detected by a cheap native size observation without an automatic first save.");
            application.ActiveDocument.Revisions.Count = 0;
            automatic.Pulse(Start.AddSeconds(9));
            automatic.Pulse(Start.AddSeconds(10));
            Check(refreshes == 5, "Accepting native revisions is detected when the revision count decreases while already dirty.");
            int unchangedEnd = application.ActiveDocument.Content.End;
            application.ActiveDocument.Content.Text = "Same-length update";
            automatic.Changed(Start.AddSeconds(11));
            automatic.Pulse(Start.AddSeconds(12));
            Check(refreshes == 6 && observedText == "Same-length update" && application.ActiveDocument.Content.End == unchangedEnd,
                "Selection activity refreshes current same-length text while the native document remains dirty.");
            application.ActiveDocument.Path = "C:\\synthetic";
            application.ActiveDocument.Saved = true;
            automatic.Pulse(Start.AddSeconds(13));
            automatic.Pulse(Start.AddSeconds(14));
            application.ActiveDocument.Saved = false;
            automatic.Pulse(Start.AddSeconds(15));
            automatic.Pulse(Start.AddSeconds(16));
            Check(application.ActiveDocument.Saves == 1 && application.ActiveDocument.Saved, "After the user's ordinary first Save As, subsequent edits save automatically.");
        }
    }

    private static void BusyReadOnlyAndProtection()
    {
        var application = Application();
        bool available = false;
        int refreshes = 0;
        using (var automatic = new WordAutoSave(application, () => available, () => refreshes++, message => { throw new Exception(message); }))
        {
            automatic.Pulse(Start);
            automatic.Pulse(Start.AddSeconds(1));
            Check(refreshes == 0 && application.ActiveDocument.Saves == 0, "A modal window or in-progress authority operation defers all automatic work.");
            available = true;
            application.UndoRecord.CustomRecordLevel = 1;
            automatic.Pulse(Start.AddSeconds(2));
            Check(refreshes == 0, "Custom Undo records are never entered by background refresh or save.");
            application.UndoRecord.CustomRecordLevel = 0;
            application.BackgroundSavingStatus = 1;
            automatic.Pulse(Start.AddSeconds(3));
            Check(refreshes == 0, "An existing native background save is allowed to finish first.");
            application.BackgroundSavingStatus = 0;
            application.ActiveDocument.ReadOnly = true;
            automatic.Pulse(Start.AddSeconds(4));
            Check(refreshes == 0 && application.ActiveDocument.Saves == 0, "Read-only documents remain untouched.");
            application.ActiveDocument.ReadOnly = false;
            application.ActiveDocument.ProtectionType = 3;
            automatic.Pulse(Start.AddSeconds(5));
            Check(refreshes == 0 && application.ActiveDocument.Saves == 0, "Native document protection is never removed or bypassed.");
            application.ActiveDocument.ProtectionType = -1;
            automatic.Pulse(Start.AddSeconds(6));
            automatic.Pulse(Start.AddSeconds(7));
            Check(refreshes == 1 && application.ActiveDocument.Saves == 1, "Automatic work resumes after native busy and protection barriers clear.");
        }
    }

    private static void FailureAndRecovery()
    {
        var application = Application();
        application.ActiveDocument.SaveFails = true;
        var errors = new List<string>();
        using (var automatic = new WordAutoSave(application, () => true, () => { }, errors.Add))
        {
            automatic.Pulse(Start);
            automatic.Pulse(Start.AddSeconds(1));
            Check(!application.ActiveDocument.Saved && automatic.LastError != null && errors.Count == 1 && errors[0].Contains("Word's Save command"),
                "A failed save keeps the document dirty and surfaces an actionable local error.");
            automatic.Pulse(Start.AddSeconds(2));
            Check(application.ActiveDocument.Saves == 1, "Save failures are not retried on every timer tick.");
            automatic.Pulse(Start.AddSeconds(11));
            Check(application.ActiveDocument.Saves == 2 && errors.Count == 1, "A bounded retry does not repeat the same error dialog.");
            application.ActiveDocument.SaveFails = false;
            automatic.Pulse(Start.AddSeconds(21));
            Check(application.ActiveDocument.Saved && automatic.LastError == null && application.ActiveDocument.Saves == 3, "A later confirmed save clears the failure state.");
            application.ActiveDocument.Saved = false;
            application.ActiveDocument.SaveUnconfirmed = true;
            automatic.Pulse(Start.AddSeconds(22));
            automatic.Pulse(Start.AddSeconds(23));
            Check(!application.ActiveDocument.Saved && automatic.LastError.Contains("did not confirm") && errors.Count == 2,
                "A Save call returning without a saved document is not reported as success.");
        }
    }

    private static void DocumentSwitchAndDisposal()
    {
        var application = Application();
        AutoSaveTestDocument first = application.ActiveDocument;
        int refreshes = 0;
        var automatic = new WordAutoSave(application, () => true, () => refreshes++, message => { throw new Exception(message); });
        automatic.Pulse(Start);
        application.ActiveDocument = Application().ActiveDocument;
        automatic.Pulse(Start.AddMilliseconds(800));
        Check(first.Saves == 0 && application.ActiveDocument.Saves == 0, "A document switch cannot save the previous document through a stale reference.");
        automatic.Pulse(Start.AddMilliseconds(1600));
        Check(application.ActiveDocument.Saves == 1 && first.Saves == 0, "Only the current active named document is automatically saved.");
        automatic.Dispose();
        application.ActiveDocument.Saved = false;
        automatic.Changed(Start.AddSeconds(2));
        automatic.Pulse(Start.AddSeconds(5));
        Check(application.ActiveDocument.Saves == 1 && refreshes == 1, "Disconnection disposes the timer and prevents later document work.");
    }

    private static void RefreshChangesAreRechecked()
    {
        var application = Application();
        application.ActiveDocument.Saved = true;
        using (var automatic = new WordAutoSave(application, () => true, () => application.ActiveDocument.Saved = false,
            message => { throw new Exception(message); }))
        {
            automatic.Pulse(Start);
            automatic.Pulse(Start.AddSeconds(1));
            Check(application.ActiveDocument.Saves == 1 && application.ActiveDocument.Saved,
                "Native reconciliation changes made during refresh are included in the automatic save.");
        }
        application = Application();
        AutoSaveTestDocument first = application.ActiveDocument;
        using (var automatic = new WordAutoSave(application, () => true, () => application.ActiveDocument = Application().ActiveDocument,
            message => { throw new Exception(message); }))
        {
            automatic.Pulse(Start);
            automatic.Pulse(Start.AddSeconds(1));
            Check(first.Saves == 0 && application.ActiveDocument.Saves == 0,
                "A document switch during refresh cannot save either document through stale authority.");
        }
    }

    private static void ActiveAndBackgroundDocumentsSaveIndependently()
    {
        var application = Application(); application.Documents.Count = 2;
        AutoSaveTestDocument active = application.ActiveDocument, background = Application().ActiveDocument;
        int activeRefreshes = 0, backgroundRefreshes = 0;
        using (var activeSave = new WordAutoSave(application, () => true, () => activeRefreshes++, message => { throw new Exception(message); }))
        using (var backgroundSave = new WordAutoSave(application, () => true, () => backgroundRefreshes++, message => { throw new Exception(message); }, () => background))
        {
            background.BeforeSave = () => Check(backgroundRefreshes > 0 && Object.ReferenceEquals(application.ActiveDocument, active),
                "Selected background permissions refresh before saving without activating its document.");
            activeSave.Pulse(Start); backgroundSave.Pulse(Start);
            activeSave.Pulse(Start.AddSeconds(1)); backgroundSave.Pulse(Start.AddSeconds(1));
            Check(active.Saves == 1 && background.Saves == 1 && activeRefreshes == 1 && backgroundRefreshes == 1 &&
                Object.ReferenceEquals(application.ActiveDocument, active),
                "Active and selected background documents each refresh and save without changing the active window.");
            active.Saved = false; background.Saved = false;
            activeSave.Pulse(Start.AddSeconds(2)); backgroundSave.Pulse(Start.AddSeconds(2));
            activeSave.Changed(Start.AddMilliseconds(2800));
            backgroundSave.Pulse(Start.AddSeconds(3)); activeSave.Pulse(Start.AddSeconds(3));
            Check(background.Saves == 2 && active.Saves == 1 && backgroundRefreshes == 2 && activeRefreshes == 1,
                "Unrelated active-document typing never delays the selected background document's queued save.");
        }
    }

    private static void BackgroundTargetClosureAndSwitch()
    {
        var application = Application(); application.Documents.Count = 3;
        AutoSaveTestDocument original = Application().ActiveDocument, replacement = Application().ActiveDocument;
        object selected = original; int refreshes = 0, targetReads = 0;
        var automatic = new WordAutoSave(application, () => true, () => refreshes++, message => { throw new Exception(message); },
            () => { targetReads++; return selected; });
        try
        {
            automatic.Pulse(Start);
            selected = null; application.Documents.Count = 2;
            automatic.Pulse(Start.AddSeconds(1));
            Check(refreshes == 0 && original.Saves == 0 && application.ActiveDocument.Saves == 0,
                "A closed or unselected background target never falls back to saving the active document.");
            selected = replacement;
            automatic.Pulse(Start.AddSeconds(2));
            Check(replacement.Saves == 0, "A newly selected background document receives its own idle debounce.");
            automatic.Pulse(Start.AddSeconds(3));
            Check(replacement.Saves == 1 && original.Saves == 0 && application.ActiveDocument.Saves == 0 && refreshes == 1,
                "Switching the selected background target saves only the new current target after its refresh.");
            automatic.Dispose(); int observedReads = targetReads;
            selected = original; automatic.Changed(Start.AddSeconds(4)); automatic.Pulse(Start.AddSeconds(5));
            Check(targetReads == observedReads && original.Saves == 0 && refreshes == 1,
                "A disposed background timer performs no target lookup, refresh or save.");
        }
        finally { automatic.Dispose(); }
    }

    private static void BackgroundTargetIsRecheckedAfterRefresh()
    {
        foreach (bool closeTarget in new[] { true, false })
        {
            var application = Application(); application.Documents.Count = 3;
            AutoSaveTestDocument original = Application().ActiveDocument, replacement = Application().ActiveDocument;
            object selected = original; int refreshes = 0;
            using (var automatic = new WordAutoSave(application, () => true,
                () => { refreshes++; selected = closeTarget ? null : replacement; }, message => { throw new Exception(message); }, () => selected))
            {
                automatic.Pulse(Start); automatic.Pulse(Start.AddSeconds(1));
                Check(original.Saves == 0 && replacement.Saves == 0 && application.ActiveDocument.Saves == 0 && refreshes == 1,
                    closeTarget ? "Closing the selected background document during refresh saves no stale or unrelated document." :
                    "Switching the selected background document during refresh saves no stale or newly selected document early.");
                automatic.Pulse(Start.AddSeconds(2)); automatic.Pulse(Start.AddSeconds(3));
                Check(closeTarget ? refreshes == 1 && original.Saves == 0 && replacement.Saves == 0 :
                    refreshes == 2 && replacement.Saves == 1 && original.Saves == 0 && application.ActiveDocument.Saves == 0,
                    closeTarget ? "A missing selected background document stays idle until a target exists." :
                    "After a mid-refresh switch the replacement receives a fresh debounce, refresh and confirmed save.");
            }
        }
    }

    private static void SeveralBackgroundDocumentsMakeProgress()
    {
        var application = Application(); application.Documents.Count = 5;
        AutoSaveTestDocument failed = Application().ActiveDocument, good = Application().ActiveDocument;
        AutoSaveTestDocument unnamed = Application().ActiveDocument, readOnly = Application().ActiveDocument;
        failed.SaveFails = true; unnamed.Path = ""; readOnly.ReadOnly = true;
        var targets = new List<object> { failed, unnamed, readOnly, good };
        var refreshes = new Dictionary<AutoSaveTestDocument, int>(); var errors = new List<string>();
        using (var automatic = new WordAutoSave(application, () => true, target => {
            var document = (AutoSaveTestDocument)target;
            refreshes[document] = refreshes.ContainsKey(document) ? refreshes[document] + 1 : 1;
        }, message => {
            Check(good.Saved, "All eligible background documents are processed before a save-error report can block the UI.");
            errors.Add(message);
        }, () => targets))
        {
            automatic.Pulse(Start); automatic.Pulse(Start.AddSeconds(1));
            Check(failed.Saves == 1 && !failed.Saved && good.Saves == 1 && good.Saved && application.ActiveDocument.Saves == 0,
                "A failing inactive document cannot starve another inactive document or save the active document.");
            Check(unnamed.Saves == 0 && refreshes[unnamed] == 1 && readOnly.Saves == 0 && !refreshes.ContainsKey(readOnly),
                "Unnamed and read-only background documents never block writable named targets or create a new file.");
            good.Saved = false; good.Comments.Count++;
            automatic.Changed(Start.AddSeconds(2)); automatic.Pulse(Start.AddSeconds(2));
            failed.Comments.Count++; automatic.Changed(Start.AddMilliseconds(2900)); automatic.Pulse(Start.AddSeconds(3));
            Check(failed.Saves == 1 && good.Saves == 2 && errors.Count == 1 && automatic.LastError != null,
                "Per-document backoff and repeated activity on a failed target never delay another queued background save.");
            failed.SaveFails = false; automatic.Pulse(Start.AddSeconds(12));
            Check(failed.Saves == 2 && failed.Saved && automatic.LastError == null && errors.Count == 1,
                "The failed background target recovers on its own retry without duplicate error reports.");
        }
    }

    private static void BackgroundRemovalDuringRefresh()
    {
        var application = Application(); application.Documents.Count = 3;
        AutoSaveTestDocument closed = Application().ActiveDocument, remaining = Application().ActiveDocument;
        var targets = new List<object> { closed, remaining }; int refreshes = 0;
        using (var automatic = new WordAutoSave(application, () => true, target => {
            refreshes++;
            if (Object.ReferenceEquals(target, closed)) targets.Remove(closed);
        }, message => { throw new Exception(message); }, () => targets))
        {
            automatic.Pulse(Start); automatic.Pulse(Start.AddSeconds(1));
            Check(closed.Saves == 0 && remaining.Saves == 1 && application.ActiveDocument.Saves == 0 && refreshes == 2,
                "Removing one background target during refresh cancels its save while the other target still completes.");
            automatic.Pulse(Start.AddSeconds(2)); automatic.Pulse(Start.AddSeconds(3));
            Check(closed.Saves == 0 && refreshes == 2,
                "Closed background worker state is pruned without repeated refresh or a stale save.");
            automatic.Dispose(); remaining.Saved = false;
            automatic.Changed(Start.AddSeconds(4)); automatic.Pulse(Start.AddSeconds(5));
            Check(remaining.Saves == 1 && refreshes == 2, "Disposing the coordinator stops all per-document workers.");
        }
    }

    private static void BackgroundDocumentLimit()
    {
        var application = Application(); application.Documents.Count = 34;
        var targets = Enumerable.Range(0, 33).Select(index => (object)Application().ActiveDocument).ToList();
        var errors = new List<string>(); int refreshes = 0;
        using (var automatic = new WordAutoSave(application, () => true, target => refreshes++, errors.Add, () => targets))
        {
            automatic.Pulse(Start);
            Check(refreshes == 0 && targets.Cast<AutoSaveTestDocument>().All(document => document.Saves == 0) &&
                errors.Count == 1 && automatic.LastError.Contains("At most 32"),
                "More than 32 background targets are refused before any partial refresh, save or worker addition.");
            targets.RemoveAt(32); application.Documents.Count = 33;
            automatic.Pulse(Start.AddSeconds(11)); automatic.Pulse(Start.AddSeconds(12));
            Check(refreshes == 32 && targets.Cast<AutoSaveTestDocument>().All(document => document.Saves == 1) &&
                application.ActiveDocument.Saves == 0 && automatic.LastError == null,
                "At the 32-document bound every permitted background target receives its own confirmed save.");
        }
    }

    private static AutoSaveTestApplication Application()
    {
        return new AutoSaveTestApplication { Documents = new AutoSaveTestCount { Count = 1 }, UndoRecord = new AutoSaveTestUndo(),
            ActiveDocument = new AutoSaveTestDocument { Path = "C:\\synthetic", ProtectionType = -1, Content = new AutoSaveTestRange { End = 20 },
                Revisions = new AutoSaveTestCount(), Comments = new AutoSaveTestCount() } };
    }

    private static void Check(bool condition, string description)
    { if (!condition) throw new Exception("FAILED: " + description); checks++; }
}
