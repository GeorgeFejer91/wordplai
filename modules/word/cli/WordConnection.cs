using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Wordplai.Cli
{
    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int type, IntPtr caller, int elapsed, IntPtr info);
        [PreserveSig] int RetryRejectedCall(IntPtr callee, int elapsed, int rejection);
        [PreserveSig] int MessagePending(IntPtr callee, int elapsed, int pending);
    }

    internal sealed class WordConnection : IMessageFilter, IDisposable
    {
        private delegate bool Enumerate(IntPtr window, IntPtr state);
        [DllImport("user32.dll")] private static extern bool EnumWindows(Enumerate callback, IntPtr state);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr window, Enumerate callback, IntPtr state);
        [DllImport("user32.dll")] private static extern int GetClassName(IntPtr window, StringBuilder name, int length);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object value);
        [DllImport("ole32.dll")] private static extern int CoRegisterMessageFilter(IMessageFilter filter, out IMessageFilter previous);
        private IMessageFilter previous;
        private dynamic api;
        private bool registered;
        internal int ProcessId { get; private set; }
        internal bool RequestSent { get; private set; }

        internal WordConnection(int requested)
        {
            var processes = Process.GetProcessesByName("WINWORD");
            try
            {
                if (requested == 0 && processes.Length != 1) throw new InvalidOperationException("Open Word, or specify --word-pid when multiple Word processes are running.");
                if (requested != 0 && !processes.Any(process => process.Id == requested)) throw new InvalidOperationException("The requested Word process is not running.");
                ProcessId = requested == 0 ? processes[0].Id : requested;
            }
            finally { foreach (Process process in processes) process.Dispose(); }
            Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(this, out previous)); registered = true;
            try
            {
                object native = null;
                EnumWindows(delegate(IntPtr parent, IntPtr ignored) {
                    uint pid; GetWindowThreadProcessId(parent, out pid);
                    if (pid != ProcessId) return true;
                    EnumChildWindows(parent, delegate(IntPtr child, IntPtr unused) {
                        var name = new StringBuilder(80); GetClassName(child, name, name.Capacity);
                        if (name.ToString() == "_WwG")
                        {
                            object candidate; Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
                            if (AccessibleObjectFromWindow(child, 0xfffffff0, ref dispatch, out candidate) == 0 && candidate != null)
                            { native = candidate; return false; }
                        }
                        return true;
                    }, IntPtr.Zero);
                    return native == null;
                }, IntPtr.Zero);
                if (native == null) throw new InvalidOperationException("The requested Word process has no accessible document window. Open a document normally first.");
                dynamic application = ((dynamic)native).Application;
                uint actual; GetWindowThreadProcessId(new IntPtr((int)application.ActiveWindow.Hwnd), out actual);
                if (actual != ProcessId) throw new InvalidOperationException("Word process identity changed; no request was sent.");
                dynamic addin = application.COMAddIns.Item("UnderTree.WordAddin");
                if (!(bool)addin.Connect) throw new InvalidOperationException("Wordplai is not connected in the requested Word process.");
                api = addin.Object;
                if (api == null) throw new InvalidOperationException("Wordplai's local interface is not available.");
            }
            catch { Dispose(); throw; }
        }

        internal string Execute(string operation, string arguments)
        { RequestSent = true; return (string)api.Execute(operation, arguments); }
        public int HandleInComingCall(int type, IntPtr caller, int elapsed, IntPtr info) { return 0; }
        public int RetryRejectedCall(IntPtr callee, int elapsed, int rejection)
        { return elapsed < 10000 && (rejection == 1 || rejection == 2) ? 100 : -1; }
        public int MessagePending(IntPtr callee, int elapsed, int pending) { return 2; }
        public void Dispose()
        {
            if (!registered) return;
            registered = false;
            IMessageFilter ignored; Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(previous, out ignored));
            api = null;
        }
    }
}
