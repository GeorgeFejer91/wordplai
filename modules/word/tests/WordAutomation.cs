using System;
using System.Runtime.InteropServices;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IWordTestMessageFilter
{
    [PreserveSig] int HandleInComingCall(int type, IntPtr caller, int elapsed, IntPtr info);
    [PreserveSig] int RetryRejectedCall(IntPtr callee, int elapsed, int rejection);
    [PreserveSig] int MessagePending(IntPtr callee, int elapsed, int pending);
}

public sealed class WordTestAutomation : IWordTestMessageFilter, IDisposable
{
    private delegate bool Enumerate(IntPtr window, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumWindows(Enumerate callback, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr window, Enumerate callback, IntPtr state);
    [DllImport("user32.dll")] private static extern int GetClassName(IntPtr window, System.Text.StringBuilder name, int length);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object value);
    private IWordTestMessageFilter previous;
    [DllImport("ole32.dll")] private static extern int CoRegisterMessageFilter(IWordTestMessageFilter filter, out IWordTestMessageFilter old);
    public WordTestAutomation() { Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(this, out previous)); }
    public int HandleInComingCall(int type, IntPtr caller, int elapsed, IntPtr info) { return 0; }
    public int RetryRejectedCall(IntPtr callee, int elapsed, int rejection)
    { return elapsed < 10000 && (rejection == 1 || rejection == 2) ? 100 : -1; }
    public int MessagePending(IntPtr callee, int elapsed, int pending) { return 2; }
    public static string Read(object value) { dynamic document = value; return (string)document.Content.Text; }
    public static int Italic(object value, int from, int to) { dynamic document = value; return (int)document.Range(from, to).Font.Italic; }
    public static object AttachProcess(int processId)
    {
        object native = null;
        EnumWindows(delegate(IntPtr parent, IntPtr ignored) {
            uint pid; GetWindowThreadProcessId(parent, out pid);
            if (pid != processId) return true;
            EnumChildWindows(parent, delegate(IntPtr child, IntPtr unused) {
                var name = new System.Text.StringBuilder(80); GetClassName(child, name, name.Capacity);
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
        if (native == null) throw new InvalidOperationException("The owned Word process is not ready.");
        dynamic application = ((dynamic)native).Application;
        uint actual; GetWindowThreadProcessId(new IntPtr((int)application.ActiveWindow.Hwnd), out actual);
        if (actual != processId) throw new InvalidOperationException("Word process identity does not match the owned fixture.");
        return application;
    }
    public void Dispose() { IWordTestMessageFilter ignored; Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(previous, out ignored)); }
}
