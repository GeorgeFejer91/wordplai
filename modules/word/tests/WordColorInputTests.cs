using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;
using UnderTree.WordAddin;

internal static class WordColorInputTests
{
    private static readonly IntPtr Frame = new IntPtr(100);
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
    private static int checks, failures;

    public static int Main()
    {
        Admission(); Boundaries(); PlatformBits(); Stability(); Rollback();
        Console.WriteLine("CHECKS " + checks + "; FAILURES " + failures + "; injected samples only; no Win32, Word, browser or profile calls.");
        return failures == 0 ? 0 : 1;
    }

    private static WordColorInput.Sample Ready(bool foreground = true)
    {
        return new WordColorInput.Sample { Readable = true, Enabled = true, ForegroundWord = foreground,
            NormalForeground = true, QueueReadable = true, Now = 2000, LastInput = 1000, WordProcess = 10, WordThread = 20 };
    }

    private static void Check(bool condition, string label)
    {
        checks++;
        if (!condition) failures++;
        Console.WriteLine((condition ? "PASS " : "FAIL ") + label);
    }

    private static bool Busy(Action action)
    {
        try { action(); return false; }
        catch (McpOperationException error)
        {
            var details = Json.Deserialize<Dictionary<string, object>>(Json.Serialize(error.Details));
            return error.Code == "word_busy" && Object.Equals(details["operation_started"], false) && Object.Equals(details["outcome_unknown"], false);
        }
    }

    private static bool Allowed(WordColorInput.Sample sample)
    {
        try { return new WordColorInput(_ => sample).RequireQuiet(Frame) == sample; }
        catch (McpOperationException) { return false; }
    }

    private static void Admission()
    {
        Check(Allowed(Ready()), "Idle foreground Word is admitted without moving focus.");
        Check(Busy(() => new WordColorInput(_ => Ready()).RequireQuiet(IntPtr.Zero)), "An unavailable native frame refuses admission before sampling.");
        var background = Ready(false); background.Now = 1001; background.HeldInput = true;
        Check(Allowed(background), "Background Word retains admission despite recent or held input in another foreground app.");
        background.QueueReadable = false;
        Check(Allowed(background), "A background out-of-process fixture does not pretend to read Word's queue.");
        var foreground = Ready(); foreground.QueueReadable = false;
        Check(!Allowed(foreground), "Foreground admission refuses an unverified GUI queue owner.");
        foreach (Action<WordColorInput.Sample> change in new Action<WordColorInput.Sample>[] {
            sample => sample.Readable = false, sample => sample.Enabled = false, sample => sample.NormalForeground = false,
            sample => sample.WordProcess = 0, sample => sample.WordThread = 0, sample => sample.Capture = new IntPtr(9) })
        {
            var sample = Ready(); change(sample);
            Check(!Allowed(sample), "Unreadable, disabled, modal, captured or invalid Word state refuses before effects.");
        }
        Check(Busy(() => new WordColorInput(_ => null).RequireQuiet(Frame)), "A missing sample is a definite no-effect busy refusal.");
        Check(Busy(() => new WordColorInput(delegate { throw new IOException("Probe unavailable."); }).RequireQuiet(Frame)),
            "A failed probe is a definite no-effect busy refusal.");
        foreach (uint flag in new uint[] { 2, 4, 8, 16 })
        {
            var sample = Ready(); sample.GuiFlags = flag;
            Check(!Allowed(sample), "Move/size, menu, system menu and popup loops refuse admission.");
        }
        var caret = Ready(); caret.GuiFlags = 1;
        Check(Allowed(caret), "An idle blinking caret is not treated as user input.");
        var held = Ready(); held.LastInput = 1; held.HeldInput = true;
        Check(!Allowed(held), "Held keyboard, modifier or mouse input refuses even with an old timestamp.");
        foreach (uint pending in new uint[] { 1, 2, 4, 0x80, 0x400, 0x800, 0x1000 })
        {
            var sample = Ready(); sample.PendingInput = pending;
            Check(!Allowed(sample), "Queued key, mouse, hotkey, raw, touch and pointer input refuses admission.");
        }
    }

    private static void Boundaries()
    {
        foreach (uint age in new uint[] { 0, 1, 749, 750, 751, Int32.MaxValue })
        {
            var sample = Ready(); sample.Now = unchecked(sample.LastInput + age);
            Check(Allowed(sample) == (age >= 750), "Foreground quiet-age boundaries use unsigned elapsed ticks.");
        }
        foreach (uint age in new uint[] { 0x80000000, UInt32.MaxValue })
        {
            var sample = Ready(); sample.Now = unchecked(sample.LastInput + age);
            Check(!Allowed(sample), "Future and half-range ambiguous timestamps fail closed.");
        }
        var recentWrap = Ready(); recentWrap.Now = 10; recentWrap.LastInput = UInt32.MaxValue - 100;
        Check(!Allowed(recentWrap), "Recent input remains recent across DWORD tick wrap.");
        var quietWrap = Ready(); quietWrap.Now = 1000; quietWrap.LastInput = UInt32.MaxValue - 100;
        Check(Allowed(quietWrap), "A genuinely quiet interval remains valid across DWORD tick wrap.");
    }

    private static void PlatformBits()
    {
        Check(!WordColorInput.Held(0) && !WordColorInput.Held(1), "The unreliable recently-pressed low bit grants no held-input evidence.");
        Check(WordColorInput.Held(unchecked((short)0x8000)) && WordColorInput.Held(unchecked((short)0x8001)), "The current high/sign bit detects held input.");
        Check(WordColorInput.Queued(1) == 0, "New-message low bits do not masquerade as currently queued input.");
        foreach (uint flag in new uint[] { 1, 2, 4, 0x80, 0x400, 0x800, 0x1000 })
            Check(WordColorInput.Queued(flag << 16) == flag, "The queue high word retains each current input class.");
        Check(WordColorInput.Queued(0x00200000) == 0, "Paint work alone is not queued user input.");
    }

    private static void Stability()
    {
        var admission = Ready();
        var stable = Ready(); stable.Now += 300;
        Check(new WordColorInput(_ => stable).RequireStable(Frame, admission) == stable, "Stable input permits pre-effect and post-start verification.");
        foreach (Action<WordColorInput.Sample> change in new Action<WordColorInput.Sample>[] {
            sample => sample.LastInput += 1, sample => sample.LastInput -= 1, sample => sample.ForegroundWord = false,
            sample => sample.WordProcess += 1, sample => sample.WordThread += 1, sample => sample.Readable = false,
            sample => sample.HeldInput = true, sample => sample.PendingInput = 1 })
        {
            var sample = Ready(); change(sample);
            Check(Busy(() => new WordColorInput(_ => sample).RequireStable(Frame, admission)),
                "Changed, backward, unreadable, held, queued or retargeted input cannot reuse an admission stamp.");
        }
        Check(Busy(() => new WordColorInput(_ => Ready()).RequireStable(Frame, null)), "A missing original stamp cannot authorize native effects.");
        var backgroundAdmission = Ready(false);
        var backgroundInput = Ready(false); backgroundInput.LastInput += 500;
        Check(new WordColorInput(_ => backgroundInput).IsStable(Frame, backgroundAdmission),
            "Input in another foreground app permits unchanged background Word to finish coloring.");
        var foregroundInput = Ready(); foregroundInput.LastInput++;
        Check(!new WordColorInput(_ => foregroundInput).IsStable(Frame, admission),
            "Foreground Word still refuses a changed input stamp during coloring.");
        var backgroundQueued = Ready(false); backgroundQueued.PendingInput = 1;
        Check(!new WordColorInput(_ => backgroundQueued).IsStable(Frame, backgroundAdmission),
            "Queued input in background Word still refuses forward coloring.");
        Check(!new WordColorInput(_ => Ready()).IsStable(Frame, backgroundAdmission),
            "Word becoming foreground invalidates background coloring admission.");
    }

    private static void Rollback()
    {
        var admission = Ready();
        var guard = new WordColorInput(_ => Ready());
        Check(guard.CanRollback(Frame, admission, "Original prose.", "Original prose."), "Verified unchanged text and input permit the existing exact native rollback check.");
        foreach (string after in new[] { "User text. Original prose.", "Original prose!", "Original pros.", "Original prose.\r", "Original prose.\u0301" })
            Check(!guard.CanRollback(Frame, admission, "Original prose.", after), "Interleaved user text forbids Undo of a potentially mixed record.");
        Check(!guard.CanRollback(Frame, admission, null, null), "Unavailable native text cannot prove safe rollback.");
        var changed = Ready(); changed.LastInput--;
        Check(!new WordColorInput(_ => changed).CanRollback(Frame, admission, "same", "same"), "Backward input stamps forbid Undo even when source text is unchanged.");
        var unreadable = Ready(); unreadable.Readable = false;
        Check(!new WordColorInput(_ => unreadable).CanRollback(Frame, admission, "same", "same"), "Unreadable post-start state forbids Undo and requires uncertainty.");
        Check(!new WordColorInput(delegate { throw new IOException("Probe failed after start."); }).CanRollback(Frame, admission, "same", "same"),
            "A post-start probe exception cannot become permission to Undo.");
        var backgroundAdmission = Ready(false);
        Check(new WordColorInput(_ => Ready(false)).CanRollback(Frame, backgroundAdmission, "same", "same"),
            "Unchanged background input and text remain eligible for exact native rollback verification.");
        var backgroundInput = Ready(false); backgroundInput.LastInput += 500;
        Check(!new WordColorInput(_ => backgroundInput).CanRollback(Frame, backgroundAdmission, "same", "same"),
            "Changed global input forbids Undo even when Word stays background and text is unchanged.");
    }
}
