using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

[assembly: AssemblyTitle("wordplai")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyFileVersion("0.1.0.0")]
[assembly: ComVisible(false)]

namespace UnderTree.WordAddin
{
    [ComImport, ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IOfficeLifecycle
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int mode,
            [MarshalAs(UnmanagedType.IDispatch)] object addin, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(int mode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComImport, ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IOfficeRibbon
    {
        [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }

    [ComVisible(true), Guid("4F8A9DB3-56E6-473F-B1A5-D1D4661D2E74"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IWordEditorAutomation
    {
        [DispId(1)] string Execute(string operation, string arguments);
        [DispId(2)] void ShowEditor();
        [DispId(3)] void RibbonOpen([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(4)] void RibbonConnections([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(5)] void RibbonLoaded([MarshalAs(UnmanagedType.IDispatch)] object ribbon);
        [DispId(6)] [return: MarshalAs(UnmanagedType.IDispatch)] object RibbonColorImage([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(7)] string RibbonColorLabel([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(8)] bool RibbonColorEnabled([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(9)] void RibbonColor([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(10)] void RibbonNextColors([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(11)] void RibbonColorRules([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(12)] void RibbonSelectionRules([MarshalAs(UnmanagedType.IDispatch)] object control);
    }

    [ComVisible(true), Guid("12E46C89-0A1D-4E2E-AECB-2C51D78F1907"), ProgId("UnderTree.WordAddin"),
     ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IWordEditorAutomation))]
    public sealed class ComAddin : IOfficeLifecycle, IOfficeRibbon, IWordEditorAutomation
    {
        [ComImport, Guid("0002095E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        private interface IWordRange { }
        private Control dispatcher;
        private WordEditor editor;
        private ConnectionsDialog connections;
        private WordRibbon ribbon;
        private ColorRuleWindows ruleWindows;
        private object wordApplication;
        private Action documentChanged;
        private Action<object> selectionChanged;
        private WordAutoSave automaticSave;
        private WordAutoSave backgroundAutomaticSave;
        private delegate void BeforeDoubleClick(object selection, ref bool cancel);
        private BeforeDoubleClick documentDoubleClick;
        private bool highlightedRulesPending;
        private bool paletteRefreshPending;
        private bool paletteRefreshQueued;
        private int authorityDepth;
        private static readonly Guid WordEvents = new Guid("00020A01-0000-0000-C000-000000000046");
        private int ownerThread;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 8388608 };

        private sealed class WordDialogOwner : IWin32Window
        {
            private readonly uint process;
            public IntPtr Handle { get; private set; }
            internal WordDialogOwner(IntPtr handle) { Handle = handle; GetWindowThreadProcessId(handle, out process); }
            internal bool IsCurrent { get { uint current; return Handle != IntPtr.Zero && GetWindowThreadProcessId(Handle, out current) != 0 && current == process; } }
        }

        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern uint GetMessagePos();
        [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(System.Drawing.Point point);
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);

        private static System.Drawing.Point ClickPoint(IWin32Window owner)
        {
            uint process;
            // In Word's UI thread the dequeued mouse message is stable even if the cursor moves.
            // An out-of-process native event test has its own message queue instead.
            if (GetWindowThreadProcessId(owner.Handle, out process) != GetCurrentThreadId()) return Cursor.Position;
            uint point = GetMessagePos();
            return new System.Drawing.Point(unchecked((short)(point & 65535)), unchecked((short)(point >> 16)));
        }

        private WordDialogOwner DialogOwner()
        {
            return new WordDialogOwner(new IntPtr((int)((dynamic)wordApplication).ActiveWindow.Hwnd));
        }

        private ColorRuleWindows RuleWindows
        {
            get { return ruleWindows ?? (ruleWindows = new ColorRuleWindows(LocalDispatch, DialogOwner)); }
        }

        public void OnConnection(object application, int mode, object addin, ref Array custom)
        {
            ownerThread = Thread.CurrentThread.ManagedThreadId;
            dispatcher = new Control();
            IntPtr handle = dispatcher.Handle;
            editor = new WordEditor(application, MarshalToWord);
            ribbon = new WordRibbon(LocalDispatch, RuleWindows);
            wordApplication = application;
            Action<string> reportAutomaticSave = message => MessageBox.Show(message, AppBranding.Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            automaticSave = new WordAutoSave(application, AutomaticWorkAvailable, RefreshCurrentDocument, reportAutomaticSave);
            backgroundAutomaticSave = new WordAutoSave(application, AutomaticWorkAvailable, RefreshBackgroundReviewDocument,
                reportAutomaticSave, () => editor.BackgroundReviewDocuments());
            documentChanged = delegate { QueuePaletteRefresh(); if (automaticSave != null) automaticSave.Changed(DateTime.UtcNow); };
            ComEventsHelper.Combine(wordApplication, WordEvents, 3, documentChanged);
            selectionChanged = delegate(object selection) { if (authorityDepth == 0 && automaticSave != null) automaticSave.Changed(DateTime.UtcNow); };
            ComEventsHelper.Combine(wordApplication, WordEvents, 12, selectionChanged);
            documentDoubleClick = OpenHighlightedRules;
            ComEventsHelper.Combine(wordApplication, WordEvents, 14, documentDoubleClick);
            QueuePaletteRefresh();
            editor.ReviewChanged += delegate {
                if (ribbon != null) ribbon.Invalidate();
                if (automaticSave != null) automaticSave.Changed(DateTime.UtcNow);
                if (backgroundAutomaticSave != null) backgroundAutomaticSave.Changed(DateTime.UtcNow);
            };
            ((dynamic)addin).Object = this;
        }

        private object MarshalToWord(Func<object> operation)
        {
            if (dispatcher == null || dispatcher.IsDisposed) throw new InvalidOperationException("Word is closing.");
            if (Thread.CurrentThread.ManagedThreadId == ownerThread) return InvokeWord(operation);
            var completion = new System.Threading.Tasks.TaskCompletionSource<object>();
            var pending = 0;
            dispatcher.BeginInvoke((Action)(() => {
                if (Interlocked.CompareExchange(ref pending, 1, 0) != 0) return;
                try { completion.SetResult(InvokeWord(operation)); }
                catch (Exception error) { completion.SetException(error); }
            }));
            bool completed;
            try { completed = completion.Task.Wait(TimeSpan.FromSeconds(15)); }
            catch (AggregateException) { return completion.Task.GetAwaiter().GetResult(); }
            if (!completed)
            {
                bool started = Interlocked.CompareExchange(ref pending, 2, 0) != 0;
                if (!started) operation = null;
                else if (completion.Task.IsCompleted) return completion.Task.GetAwaiter().GetResult();
                if (started) completion.Task.ContinueWith(failed => { var observed = failed.Exception; },
                    System.Threading.Tasks.TaskContinuationOptions.OnlyOnFaulted | System.Threading.Tasks.TaskContinuationOptions.ExecuteSynchronously);
                throw new McpOperationException(started
                    ? "Word is still completing this request. Inspect its result in Word before retrying."
                    : "Word is busy. This request did not start; continue after the current Word action finishes.",
                    "word_busy", new { operation_started = started, outcome_unknown = started });
            }
            return completion.Task.GetAwaiter().GetResult();
        }

        public string Execute(string operation, string arguments)
        {
            try
            {
                if (arguments == null || arguments.Length > 131072) throw new ArgumentException("Invalid local request.");
                var values = json.Deserialize<Dictionary<string, object>>(arguments);
                return json.Serialize(new { ok = true, result = MarshalToWord(() => LocalDispatch(operation, values)) });
            }
            catch (Exception error)
            {
                while (error is TargetInvocationException || error is AggregateException) error = error.InnerException;
                var refusal = error as McpOperationException;
                return json.Serialize(new { ok = false, error = error.Message,
                    error_code = refusal == null ? null : refusal.Code, details = refusal == null ? null : refusal.Details });
            }
        }

        public void ShowEditor()
        {
            MarshalToWord(delegate { ShowConnectionsCore(); return null; });
        }

        private void ShowConnectionsCore()
        {
            if (editor == null) throw new InvalidOperationException("Open Word normally to use " + AppBranding.Name + ".");
            if (ribbon != null) ribbon.CancelPendingHighlight();
            if (connections == null || connections.IsDisposed) connections = new ConnectionsDialog(LocalDispatch);
            if (!connections.Visible) connections.Show();
            connections.Activate();
        }

        public void RibbonOpen(object control) { ShowEditor(); }
        public void RibbonConnections(object control)
        {
            MarshalToWord(delegate { ShowConnectionsCore(); return null; });
        }
        private object LocalDispatch(string operation, IDictionary<string, object> arguments)
        {
            return InvokeWord(delegate {
                object result = json.DeserializeObject(json.Serialize(editor.Dispatch(operation, arguments)));
                if (ribbon != null && operation != "automation") ribbon.Observe(result);
                if (operation == "automation" && WordEditor.AutomationMutates(arguments)) QueuePaletteRefresh();
                if (automaticSave != null && (operation == "save_rule" || operation == "mark_selection" ||
                    operation == "remove_mark" || operation == "accept" || operation == "reject" ||
                    (operation == "automation" && WordEditor.AutomationMutates(arguments)))) automaticSave.Changed(DateTime.UtcNow);
                return result;
            });
        }
        private bool AutomaticWorkAvailable()
        {
            if (authorityDepth != 0 || editor == null || dispatcher == null || dispatcher.IsDisposed || wordApplication == null) return false;
            return (int)((dynamic)wordApplication).Documents.Count == 0 || IsWindowEnabled(DialogOwner().Handle);
        }
        private void RefreshCurrentDocument()
        {
            InvokeWord(delegate {
                object result = json.DeserializeObject(json.Serialize(editor.Refresh()));
                if (ribbon != null) ribbon.Observe(result);
                return null;
            });
        }
        private void RefreshBackgroundReviewDocument(object target)
        {
            InvokeWord(delegate { editor.RefreshReviewDocument(target); return null; });
        }
        private object InvokeWord(Func<object> operation)
        {
            authorityDepth++;
            try { return operation(); }
            finally { authorityDepth--; if (authorityDepth == 0 && paletteRefreshPending) QueuePaletteRefresh(); }
        }
        private void QueuePaletteRefresh()
        {
            paletteRefreshPending = true;
            if (paletteRefreshQueued || dispatcher == null || dispatcher.IsDisposed) return;
            paletteRefreshQueued = true;
            dispatcher.BeginInvoke((Action)delegate {
                paletteRefreshQueued = false;
                if (authorityDepth != 0 || editor == null) return;
                paletteRefreshPending = false;
                try { LocalDispatch("state", new Dictionary<string, object>()); }
                catch (Exception) { if (ribbon != null) ribbon.Observe(new Dictionary<string, object> { { "rules", new object[0] } }); }
            });
        }
        public void RibbonLoaded(object ui) { if (ribbon != null) { ribbon.Load(ui); QueuePaletteRefresh(); } }
        public object RibbonColorImage(object control) { return ribbon == null ? null : ribbon.Image(control); }
        public string RibbonColorLabel(object control) { return ribbon == null ? "Open a Word document" : ribbon.Label(control); }
        public bool RibbonColorEnabled(object control) { return ribbon != null && ribbon.Enabled(control); }
        public void RibbonColor(object control) { if (ribbon != null) ribbon.Highlight(control); }
        public void RibbonNextColors(object control) { if (ribbon != null) ribbon.Configure(); }
        public void RibbonColorRules(object control) { if (ribbon != null) ribbon.Configure(); }
        public void RibbonSelectionRules(object control)
        {
            try { MarshalToWord(delegate {
                if (ribbon != null) ribbon.CancelPendingHighlight();
                IWin32Window owner = DialogOwner();
                if (!HighlightedRulesDialog.OpenAtSelection(LocalDispatch, owner, RuleWindows))
                    MessageBox.Show(owner, "Place the cursor in highlighted text to open its color rule.", AppBranding.Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }); }
            catch (Exception error) { MessageBox.Show(error.Message, AppBranding.Name, MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
        private void OpenHighlightedRules(object selection, ref bool cancel)
        {
            if (editor == null || authorityDepth != 0 || dispatcher == null || dispatcher.IsDisposed) return;
            WordDialogOwner owner = null;
            bool posting = false;
            try {
                owner = DialogOwner();
                dynamic received = ((dynamic)selection).Range;
                dynamic current = ((dynamic)wordApplication).Selection.Range;
                if ((int)received.StoryType != 1 || (int)current.StoryType != 1) return;
                var point = ClickPoint(owner);
                IntPtr pointed = WindowFromPoint(point);
                if (pointed != owner.Handle && !IsChild(owner.Handle, pointed)) return;
                dynamic clicked = ((dynamic)wordApplication).ActiveWindow.RangeFromPoint(point.X, point.Y);
                if (clicked == null || !((object)clicked is IWordRange) || (int)clicked.StoryType != 1) return;
                object document = (object)clicked.Document;
                int from = (int)clicked.Start;
                int initialFrom = (int)current.Start, initialTo = (int)current.End;
                dynamic word = clicked.Words.First;
                int wordFrom = (int)word.Start, wordTo = (int)word.End;
                if (WordSnapshot.Identity(document) != WordSnapshot.Identity((object)((dynamic)wordApplication).ActiveDocument) ||
                    WordSnapshot.Identity(document) != WordSnapshot.Identity((object)received.Document) ||
                    WordSnapshot.Identity(document) != WordSnapshot.Identity((object)current.Document)) return;
                if (ribbon != null) ribbon.CancelPendingHighlight();
                var at = new Dictionary<string, object> { { "nativeFrom", from }, { "nativeTo", from } };
                var state = UiState.Map(LocalDispatch("selection_rules", at));
                if (UiState.Records(state, "choices").Count == 0) return;
                cancel = true;
                if (highlightedRulesPending) return;
                string choices = json.Serialize(state["choices"]);
                highlightedRulesPending = true;
                // Word must finish its input callback before showing a rule window.
                posting = true;
                dispatcher.BeginInvoke((Action)delegate {
                    try {
                        if (editor == null || authorityDepth != 0 || wordApplication == null || !owner.IsCurrent ||
                            (int)((dynamic)wordApplication).Documents.Count == 0 || (int)((dynamic)wordApplication).ActiveWindow.Hwnd != owner.Handle.ToInt32()) return;
                        dynamic active = ((dynamic)wordApplication).Selection.Range;
                        if ((int)active.StoryType != 1 || WordSnapshot.Identity(document) != WordSnapshot.Identity((object)active.Document) ||
                            !((initialFrom == (int)active.Start && initialTo == (int)active.End) ||
                              (wordFrom <= (int)active.Start && (int)active.End <= wordTo && (int)active.Start <= from && from <= (int)active.End))) return;
                        var fresh = UiState.Map(LocalDispatch("selection_rules", at));
                        if (UiState.Text(state, "documentId") != UiState.Text(fresh, "documentId") ||
                            UiState.Text(state, "revision") != UiState.Text(fresh, "revision") || choices != json.Serialize(fresh["choices"])) return;
                        RuleWindows.OpenAtSelection(at);
                    }
                    catch (Exception error) { if (!Disconnected(error) && editor != null && wordApplication != null && owner.IsCurrent) MessageBox.Show(owner, error.Message, AppBranding.Name, MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    finally { highlightedRulesPending = false; }
                });
            }
            catch (Exception error) {
                highlightedRulesPending = false;
                if (dispatcher == null || dispatcher.IsDisposed || !dispatcher.IsHandleCreated || Disconnected(error) || owner == null || !owner.IsCurrent) return;
                if (posting) throw;
                dispatcher.BeginInvoke((Action)delegate { if (editor != null && wordApplication != null && owner.IsCurrent) MessageBox.Show(owner, error.Message, AppBranding.Name, MessageBoxButtons.OK, MessageBoxIcon.Information); });
            }
        }
        private static bool Disconnected(Exception error)
        {
            var native = error as COMException;
            return native != null && (native.ErrorCode == unchecked((int)0x80010108) || native.ErrorCode == unchecked((int)0x800401FD) ||
                native.ErrorCode == unchecked((int)0x800706BA) || native.ErrorCode == unchecked((int)0x800706BE));
        }
        public string GetCustomUI(string ribbonId)
        {
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"RibbonLoaded\"><ribbon><tabs>" +
                "<tab idMso=\"TabHome\">" + WordRibbon.GroupXml("Home") + "</tab>" +
                "<tab id=\"wrAIterTab\" label=\"" + AppBranding.Name + "\">" + WordRibbon.GroupXml("Tab") + "</tab></tabs></ribbon></customUI>";
        }

        private void Shutdown()
        {
            highlightedRulesPending = false;
            if (automaticSave != null) { automaticSave.Dispose(); automaticSave = null; }
            if (backgroundAutomaticSave != null) { backgroundAutomaticSave.Dispose(); backgroundAutomaticSave = null; }
            if (wordApplication != null && documentChanged != null)
            {
                try { ComEventsHelper.Remove(wordApplication, WordEvents, 3, documentChanged); }
                catch (COMException) { }
            }
            if (wordApplication != null && documentDoubleClick != null)
            {
                try { ComEventsHelper.Remove(wordApplication, WordEvents, 14, documentDoubleClick); }
                catch (COMException) { }
            }
            if (wordApplication != null && selectionChanged != null)
            {
                try { ComEventsHelper.Remove(wordApplication, WordEvents, 12, selectionChanged); }
                catch (COMException) { }
            }
            wordApplication = null;
            documentChanged = null;
            documentDoubleClick = null;
            selectionChanged = null;
            if (ruleWindows != null) { ruleWindows.Dispose(); ruleWindows = null; }
            if (ribbon != null) ribbon.Dispose();
            if (connections != null && !connections.IsDisposed) connections.Close();
            if (editor != null) { editor.Dispose(); editor = null; }
            ribbon = null;
            if (dispatcher != null) { dispatcher.Dispose(); dispatcher = null; }
        }
        public void OnDisconnection(int mode, ref Array custom) { Shutdown(); }
        public void OnBeginShutdown(ref Array custom) { Shutdown(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
    }
}
