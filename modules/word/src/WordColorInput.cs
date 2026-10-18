using System;
using System.Runtime.InteropServices;

namespace UnderTree.WordAddin
{
    internal sealed class WordColorInput
    {
        internal const uint QuietMilliseconds = 750;
        private const uint BusyGuiFlags = 0x1e, InputQueueFlags = 0x1c87;
        private readonly Func<IntPtr, Sample> sample;

        internal sealed class Sample
        {
            internal bool Readable, Enabled, ForegroundWord, NormalForeground, QueueReadable, HeldInput;
            internal uint Now, LastInput, GuiFlags, PendingInput, WordProcess, WordThread;
            internal IntPtr Capture;
        }

        internal WordColorInput() : this(Read) { }
        internal WordColorInput(Func<IntPtr, Sample> sample) { this.sample = sample; }

        internal static bool Held(short keyState) { return keyState < 0; }
        internal static uint Queued(uint queueStatus) { return (queueStatus >> 16) & InputQueueFlags; }

        internal Sample RequireQuiet(IntPtr frame)
        {
            if (frame == IntPtr.Zero) throw Busy();
            Sample current;
            try { current = sample(frame); }
            catch (Exception) { throw Busy(); }
            if (current == null || !current.Readable || !current.Enabled || current.WordProcess == 0 || current.WordThread == 0 ||
                (current.GuiFlags & BusyGuiFlags) != 0 || current.Capture != IntPtr.Zero || current.PendingInput != 0)
                throw Busy();
            uint age = unchecked(current.Now - current.LastInput);
            if (age > Int32.MaxValue || current.ForegroundWord &&
                (!current.NormalForeground || !current.QueueReadable || current.HeldInput || age < QuietMilliseconds))
                throw Busy();
            return current;
        }

        internal Sample RequireStable(IntPtr frame, Sample admission)
        {
            Sample current = RequireQuiet(frame);
            if (admission == null || current.WordProcess != admission.WordProcess || current.WordThread != admission.WordThread ||
                current.ForegroundWord != admission.ForegroundWord || current.ForegroundWord && current.LastInput != admission.LastInput)
                throw Busy();
            return current;
        }

        internal bool CanRollback(IntPtr frame, Sample admission, string beforeText, string currentText)
        {
            if (beforeText == null || currentText == null || !String.Equals(beforeText, currentText, StringComparison.Ordinal)) return false;
            try { return RequireStable(frame, admission).LastInput == admission.LastInput; }
            catch (McpOperationException) { return false; }
        }

        internal bool IsStable(IntPtr frame, Sample admission)
        {
            try { RequireStable(frame, admission); return true; }
            catch (McpOperationException) { return false; }
        }

        private static McpOperationException Busy()
        {
            return new McpOperationException("Word input is active or could not be verified. Finish the current action and pause briefly before coloring.",
                "word_busy", new { operation_started = false, outcome_unknown = false });
        }

        private static Sample Read(IntPtr frame)
        {
            var result = new Sample();
            if (frame == IntPtr.Zero || !IsWindow(frame)) return result;
            uint process;
            result.WordThread = GetWindowThreadProcessId(frame, out process); result.WordProcess = process;
            IntPtr foreground = GetForegroundWindow();
            uint foregroundProcess;
            if (foreground == IntPtr.Zero || GetWindowThreadProcessId(foreground, out foregroundProcess) == 0) return result;
            result.ForegroundWord = foregroundProcess == process;
            result.Enabled = IsWindowEnabled(frame);
            var gui = new GuiThreadInfo { Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo)) };
            var input = new LastInputInfo { Size = (uint)Marshal.SizeOf(typeof(LastInputInfo)) };
            if (!GetGUIThreadInfo(result.WordThread, ref gui) || !GetLastInputInfo(ref input)) return result;
            result.Readable = true;
            result.Now = unchecked((uint)Environment.TickCount); result.LastInput = input.Time;
            result.GuiFlags = gui.Flags; result.Capture = gui.Capture;
            result.NormalForeground = foreground == frame && gui.Active == frame;
            result.QueueReadable = GetCurrentThreadId() == result.WordThread;
            if (result.QueueReadable) result.PendingInput = Queued(GetQueueStatus(InputQueueFlags));
            if (result.ForegroundWord)
                for (int key = 1; key < 255; key++)
                    if (Held(GetAsyncKeyState(key))) { result.HeldInput = true; break; }
            return result;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LastInputInfo { internal uint Size, Time; }

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect { internal int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct GuiThreadInfo
        {
            internal uint Size, Flags;
            internal IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
            internal Rect CaretRectangle;
        }

        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint thread, ref GuiThreadInfo information);
        [DllImport("user32.dll")] private static extern bool GetLastInputInfo(ref LastInputInfo information);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")] private static extern uint GetQueueStatus(uint flags);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    }

    internal sealed partial class WordEditor
    {
        private readonly WordColorInput coloringInput = new WordColorInput();
    }
}
