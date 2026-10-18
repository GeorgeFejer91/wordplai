using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using UnderTree.WordAddin;

internal static class NativeAutoSaveSmoke
{
    private delegate bool Enumerate(IntPtr window, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumWindows(Enumerate callback, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr window, Enumerate callback, IntPtr state);
    [DllImport("user32.dll")] private static extern int GetClassName(IntPtr window, StringBuilder name, int length);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object value);
    private static readonly List<string> checks = new List<string>();
    private static readonly List<string> errors = new List<string>();
    private static dynamic application, document;
    private static WordEditor editor;
    private static WordAutoSave automatic;
    private static string output;
    private static int refreshes;

    [STAThread]
    public static int Main(string[] arguments)
    {
        Process owned = null;
        WordTestAutomation filter = null;
        try
        {
            filter = new WordTestAutomation();
            output = Path.GetFullPath(arguments[1]);
            Directory.CreateDirectory(output);
            string bootstrap = Path.Combine(output, "autosave-bootstrap.rtf");
            File.WriteAllText(bootstrap, @"{\rtf1\ansi Disposable automatic-save bootstrap.\par}");
            owned = Process.Start(new ProcessStartInfo(arguments[0], "/a /x /q \"" + bootstrap + "\"") {
                UseShellExecute = false, WindowStyle = ProcessWindowStyle.Hidden });
            application = AttachOwned(owned);
            application.Visible = false;
            Check((int)application.Documents.Count == 1 && (string)application.ActiveDocument.FullName == bootstrap,
                "Only the synthetic bootstrap is open in the isolated process.");
            application.ActiveDocument.Close(0);
            document = application.Documents.Add();
            document.Content.Text = "Original body with context.\r";
            editor = new WordEditor((object)application, action => action(), Path.Combine(output, "profile"), Path.Combine(output, "unused-zrok.exe"));
            automatic = new WordAutoSave((object)application, () => true, () => { editor.Refresh(); refreshes++; }, errors.Add);
            DateTime now = DateTime.UtcNow;
            automatic.Pulse(now);
            automatic.Pulse(now.AddSeconds(1));
            Check((string)document.Path == "" && !(bool)document.Saved && Directory.GetFiles(output, "*.docx").Length == 0,
                "New native documents stay unnamed until the user's first ordinary Save As.");
            string path = Path.Combine(output, "autosave-native.docx");
            document.SaveAs2(path, 16);
            Wait(() => (int)application.BackgroundSavingStatus == 0 && (bool)document.Saved, "Initial native Save As completes.");
            automatic.Pulse(now.AddMilliseconds(1500));
            string before = Digest(path);
            int unchangedEnd = (int)document.Content.End;
            document.Range(0, 8).Text = "Revised!";
            string expected = (string)document.Content.Text;
            automatic.Pulse(now.AddSeconds(2));
            Check(Digest(path) == before && !(bool)document.Saved && (int)document.Content.End == unchangedEnd,
                "A same-length native edit is detected without saving before the idle debounce.");
            automatic.Pulse(now.AddSeconds(3));
            Wait(() => (int)application.BackgroundSavingStatus == 0 && (bool)document.Saved, "Automatic native save completes.");
            Check(Digest(path) != before && errors.Count == 0 && refreshes >= 2,
                "An existing named native document refreshes and persists the actual changed bytes.");
            document.Close(0);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            document.Activate();
            Check((string)document.Content.Text == expected, "Reopening proves the automatic save stored the complete native text.");
            before = Digest(path);
            application.UndoRecord.StartCustomRecord("Disposable automatic-save busy guard");
            try
            {
                document.Range(0, 7).Text = "Current";
                automatic.Pulse(now.AddSeconds(4));
                automatic.Pulse(now.AddSeconds(5));
                Check(Digest(path) == before && !(bool)document.Saved,
                    "The native custom Undo record remains untouched by automatic saving.");
            }
            finally { application.UndoRecord.EndCustomRecord(); }
            automatic.Pulse(now.AddSeconds(6));
            automatic.Pulse(now.AddSeconds(7));
            Wait(() => (int)application.BackgroundSavingStatus == 0 && (bool)document.Saved, "Save resumes after native Undo completes.");
            before = Digest(path);
            document.Range(0, 7).Text = "Guarded";
            document.Protect(3, true, "qualification-only");
            automatic.Pulse(now.AddSeconds(8));
            automatic.Pulse(now.AddSeconds(9));
            Check((int)document.ProtectionType == 3 && Digest(path) == before,
                "A native protected document is never unprotected or saved by the idle owner.");
            document.Unprotect("qualification-only");
            automatic.Pulse(now.AddSeconds(10));
            automatic.Pulse(now.AddSeconds(11));
            Wait(() => (int)application.BackgroundSavingStatus == 0 && (bool)document.Saved, "Save resumes after the user removes native protection.");
            document.TrackRevisions = true;
            document.Range(0, 7).Text = "Tracked";
            automatic.Pulse(now.AddSeconds(12));
            automatic.Pulse(now.AddSeconds(13));
            Wait(() => (int)application.BackgroundSavingStatus == 0 && (bool)document.Saved, "A native tracked edit is automatically saved.");
            Check((int)document.Revisions.Count > 0, "Native revision records remain available for the user's ordinary Accept action.");
            int previousRefreshes = refreshes;
            document.Revisions.AcceptAll();
            document.TrackRevisions = false;
            automatic.Pulse(now.AddSeconds(14));
            automatic.Pulse(now.AddSeconds(15));
            Wait(() => (int)application.BackgroundSavingStatus == 0 && (bool)document.Saved, "Native acceptance refreshes and saves automatically.");
            Check((int)document.Revisions.Count == 0 && refreshes > previousRefreshes,
                "The native Accept count change refreshes current writing state without a manual reload.");
            Check(errors.Count == 0, "Native automatic refresh and save produced no hidden errors.");
            WriteResult("PASS", null);
            Console.WriteLine("PASS " + checks.Count + " native automatic-save assertions; isolated Word only, installed package and public client NOT RUN.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            if (output != null) WriteResult("FAIL", error.ToString());
            return 1;
        }
        finally
        {
            if (automatic != null) automatic.Dispose();
            if (editor != null) editor.Dispose();
            if (application != null) application.Quit(0);
            else if (owned != null && !owned.HasExited) owned.CloseMainWindow();
            if (owned != null) owned.Dispose();
            if (filter != null) filter.Dispose();
        }
    }

    private static object AttachOwned(Process process)
    {
        IntPtr target = IntPtr.Zero;
        Wait(delegate {
            EnumWindows(delegate(IntPtr parent, IntPtr state) {
                uint pid; GetWindowThreadProcessId(parent, out pid);
                if (pid != process.Id) return true;
                EnumChildWindows(parent, delegate(IntPtr child, IntPtr ignored) {
                    var name = new StringBuilder(80); GetClassName(child, name, name.Capacity);
                    if (name.ToString() == "_WwG") target = child;
                    return true;
                }, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
            return target != IntPtr.Zero;
        }, "The exact spawned process exposes a native document window.");
        object native = null;
        Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
        Wait(delegate { return AccessibleObjectFromWindow(target, 0xfffffff0, ref dispatch, out native) == 0 && native != null; },
            "The isolated native Word automation object becomes available.");
        dynamic app = ((dynamic)native).Application;
        uint owned; GetWindowThreadProcessId(new IntPtr((int)app.ActiveWindow.Hwnd), out owned);
        Check(owned == process.Id, "Automation belongs to the exact isolated Word process.");
        return app;
    }

    private static void Wait(Func<bool> ready, string description)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(30);
        using (var interval = new ManualResetEvent(false))
            while (!ready())
            {
                if (DateTime.UtcNow > deadline) throw new Exception("Timeout: " + description);
                Application.DoEvents(); interval.WaitOne(30);
            }
        Check(true, description);
    }

    private static string Digest(string path)
    { using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (var hash = SHA256.Create()) return Convert.ToBase64String(hash.ComputeHash(stream)); }
    private static void Check(bool condition, string description)
    { if (!condition) throw new Exception("FAILED: " + description); checks.Add(description); }
    private static void WriteResult(string status, string error)
    {
        File.WriteAllText(Path.Combine(output, "result.json"), new JavaScriptSerializer().Serialize(new {
            status = status, checks = checks, errors = errors, error = error, installedPackage = "NOT RUN", publicClient = "NOT RUN" }));
    }
}
