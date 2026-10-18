using System;
using System.Runtime.InteropServices;

namespace UnderTree.WordAddin
{
    internal sealed class WordReviewWindow : IDisposable
    {
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
        private readonly dynamic application;
        private dynamic previousWindow;
        private bool screenUpdating;
        private bool changed;
        private IntPtr foreground;
        private uint wordProcess;

        internal WordReviewWindow(dynamic application, dynamic document)
        {
            this.application = application;
            try
            {
                if (WordSnapshot.Identity((object)application.ActiveDocument) == WordSnapshot.Identity((object)document)) return;
                previousWindow = application.ActiveWindow;
            }
            catch (COMException error)
            {
                if (error.ErrorCode != unchecked((int)0x800A1098)) throw;
            }
            if ((int)application.UndoRecord.CustomRecordLevel != 0)
                throw new InvalidOperationException("Finish the current Word operation before reviewing another document.");
            foreground = GetForegroundWindow();
            GetWindowThreadProcessId(new IntPtr((int)document.Windows.Item(1).Hwnd), out wordProcess);
            screenUpdating = (bool)application.ScreenUpdating;
            changed = true;
            try { application.ScreenUpdating = false; document.Activate(); }
            catch { Dispose(); throw; }
        }

        public void Dispose()
        {
            if (!changed) return;
            dynamic restore = previousWindow;
            previousWindow = null;
            changed = false;
            try { if (restore != null) restore.Activate(); }
            finally
            {
                application.ScreenUpdating = screenUpdating;
                uint currentProcess;
                GetWindowThreadProcessId(GetForegroundWindow(), out currentProcess);
                if (foreground != IntPtr.Zero && currentProcess == wordProcess && GetForegroundWindow() != foreground)
                    SetForegroundWindow(foreground);
            }
        }
    }
}
