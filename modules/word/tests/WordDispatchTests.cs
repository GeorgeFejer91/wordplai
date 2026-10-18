using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using UnderTree.WordAddin;

internal static class WordDispatchTests
{
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
    private static readonly List<string> Checks = new List<string>();
    private static readonly List<object> Timings = new List<object>();
    private static string output;

    public static int Main(string[] args)
    {
        output = Path.GetFullPath(args[0]);
        Directory.CreateDirectory(output);
        try
        {
            ImmediateExecutionAndErrors();
            QueuedTimeoutCancelsLateMutation();
            RunningTimeoutReportsUnknownOutcome();
            Result("PASS", null);
            Console.WriteLine("PASS " + Checks.Count + " real STA dispatch assertions; no Word process was started.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            Result("FAIL", error.ToString());
            return 1;
        }
    }

    private static void ImmediateExecutionAndErrors()
    {
        using (var fixture = new DispatchFixture())
        {
            int executedThread = 0; ApartmentState apartment = ApartmentState.Unknown;
            object result = fixture.Invoke(delegate {
                executedThread = Thread.CurrentThread.ManagedThreadId;
                apartment = Thread.CurrentThread.GetApartmentState();
                return "native STA result";
            });
            Check(Object.Equals(result, "native STA result") && executedThread == fixture.OwnerThread && apartment == ApartmentState.STA,
                "The real marshaler executes on its hidden WinForms STA and returns the actual result.");
            int depth = 0;
            fixture.Post(delegate { fixture.Invoke(delegate { depth = fixture.AuthorityDepth; return null; }); });
            fixture.Drain();
            Check(depth == 1 && fixture.AuthorityDepth == 0,
                "An owner-thread call uses the normal authority scope and restores its depth.");
            var sentinel = new InvalidOperationException("Original dispatch failure"); Exception observed = null;
            try { fixture.Invoke(delegate { throw sentinel; }); } catch (Exception error) { observed = error; }
            Check(Object.ReferenceEquals(observed, sentinel) && fixture.AuthorityDepth == 0,
                "An operation failure propagates unchanged and releases the authority scope.");
            Check(Object.Equals(fixture.Invoke(() => "recovered"), "recovered"),
                "A failed callback does not prevent a subsequent dispatch.");
        }
    }

    private static void QueuedTimeoutCancelsLateMutation()
    {
        using (var fixture = new DispatchFixture())
        using (var blockerStarted = new ManualResetEvent(false))
        using (var release = new ManualResetEvent(false))
        {
            TimedCall request = null; int mutations = 0;
            try
            {
                fixture.Post(delegate {
                    blockerStarted.Set();
                    if (!release.WaitOne(35000)) throw new TimeoutException("The queued-call fixture was not released.");
                });
                Check(blockerStarted.WaitOne(5000), "The actual STA message loop is blocked before the test request is queued.");
                request = new TimedCall(() => fixture.Invoke(delegate { Interlocked.Increment(ref mutations); return "late mutation"; }));
                Check(request.Join(20000), "A queued request returns within a bounded interval while the STA remains blocked.");
                CheckBusy(request, false, "queued");
                Check(mutations == 0, "The timed-out queued callback has not started.");
                release.Set();
                fixture.Drain();
                Check(mutations == 0, "A canceled queued dispatch never runs after the STA is unblocked.");
                Check(Object.Equals(fixture.Invoke(() => "fresh after queue"), "fresh after queue") && mutations == 0,
                    "Fresh dispatch succeeds after cancellation without reviving the canceled mutation.");
            }
            finally
            {
                release.Set();
                fixture.Drain();
                if (request != null) request.Dispose();
            }
        }
    }

    private static void RunningTimeoutReportsUnknownOutcome()
    {
        using (var fixture = new DispatchFixture())
        using (var started = new ManualResetEvent(false))
        using (var release = new ManualResetEvent(false))
        {
            TimedCall request = null; int starts = 0, completions = 0;
            try
            {
                request = new TimedCall(() => fixture.Invoke(delegate {
                    Interlocked.Increment(ref starts); started.Set();
                    if (!release.WaitOne(35000)) throw new TimeoutException("The running-call fixture was not released.");
                    Interlocked.Increment(ref completions);
                    return "late completion";
                }));
                Check(started.WaitOne(5000), "The callback really starts on the STA before the timeout is measured.");
                Check(request.Join(20000), "A running request returns within a bounded interval while its callback is still blocked.");
                CheckBusy(request, true, "running");
                Check(starts == 1 && completions == 0,
                    "A running timeout reports an uncertain result while one callback is still in progress.");
                release.Set(); fixture.Drain();
                Check(starts == 1 && completions == 1 && fixture.AuthorityDepth == 0,
                    "Releasing a timed-out running operation completes it exactly once and restores native authority depth.");
                Check(Object.Equals(fixture.Invoke(() => "fresh after running"), "fresh after running") && starts == 1 && completions == 1,
                    "Subsequent dispatch recovers without repeating the operation whose result was uncertain.");
            }
            finally
            {
                release.Set();
                fixture.Drain();
                if (request != null) request.Dispose();
            }
        }
    }

    private static void CheckBusy(TimedCall request, bool started, string scenario)
    {
        var refusal = request.Error as McpOperationException;
        Check(refusal != null && refusal.Code == "word_busy", "The " + scenario + " timeout returns the typed word_busy diagnostic.");
        var details = (Dictionary<string, object>)Json.DeserializeObject(Json.Serialize(refusal.Details));
        Check(Object.Equals(details["operation_started"], started) && Object.Equals(details["outcome_unknown"], started),
            "The " + scenario + " diagnostic truthfully distinguishes canceled work from an unknown running outcome.");
        Check(request.Milliseconds >= 14500 && request.Milliseconds < 20000,
            "The " + scenario + " case uses the production 15-second timeout without a test bypass.");
        if (started) Check(refusal.Message.IndexOf("did not start", StringComparison.OrdinalIgnoreCase) < 0 &&
            refusal.Message.IndexOf("no changes", StringComparison.OrdinalIgnoreCase) < 0,
            "A running timeout never claims that no changes occurred.");
        Timings.Add(new { scenario = scenario, milliseconds = request.Milliseconds, operationStarted = started, outcomeUnknown = started });
        Console.WriteLine("TIMING: " + scenario + " " + request.Milliseconds.ToString("F1") + " ms.");
    }

    private sealed class TimedCall : IDisposable
    {
        private readonly Thread thread;
        internal Exception Error;
        internal double Milliseconds;
        internal TimedCall(Func<object> operation)
        {
            thread = new Thread(delegate() {
                var watch = Stopwatch.StartNew();
                try { operation(); } catch (Exception error) { Error = error; }
                finally { Milliseconds = watch.Elapsed.TotalMilliseconds; }
            }) { IsBackground = true, Name = "Dispatch test request" };
            thread.Start();
        }
        internal bool Join(int milliseconds) { return thread.Join(milliseconds); }
        public void Dispose() { if (!thread.Join(5000)) throw new TimeoutException("The request thread did not terminate."); }
    }

    private sealed class DispatchFixture : IDisposable
    {
        private static readonly MethodInfo Marshal = typeof(ComAddin).GetMethod("MarshalToWord", BindingFlags.NonPublic | BindingFlags.Instance);
        private static readonly FieldInfo Dispatcher = typeof(ComAddin).GetField("dispatcher", BindingFlags.NonPublic | BindingFlags.Instance);
        private static readonly FieldInfo Owner = typeof(ComAddin).GetField("ownerThread", BindingFlags.NonPublic | BindingFlags.Instance);
        private static readonly FieldInfo Depth = typeof(ComAddin).GetField("authorityDepth", BindingFlags.NonPublic | BindingFlags.Instance);
        private readonly ManualResetEvent ready = new ManualResetEvent(false);
        private readonly Thread thread;
        private volatile Exception failure;
        private volatile bool stopping;
        private Control control;
        private ComAddin addin;
        internal int OwnerThread;
        internal int AuthorityDepth { get { return (int)Depth.GetValue(addin); } }

        internal DispatchFixture()
        {
            thread = new Thread(delegate() {
                try
                {
                    using (var dispatcher = new Control())
                    {
                        IntPtr handle = dispatcher.Handle;
                        control = dispatcher; OwnerThread = Thread.CurrentThread.ManagedThreadId;
                        addin = new ComAddin(); Dispatcher.SetValue(addin, dispatcher); Owner.SetValue(addin, OwnerThread);
                        ready.Set();
                        if (!stopping) Application.Run();
                    }
                }
                catch (Exception error) { failure = error; }
                finally { ready.Set(); }
            }) { IsBackground = true, Name = "Hidden Word dispatch STA fixture" };
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            try
            {
                if (!ready.WaitOne(5000)) throw new TimeoutException("The hidden STA control did not become ready.");
                Healthy();
            }
            catch { Dispose(); throw; }
        }

        internal object Invoke(Func<object> operation)
        {
            Healthy();
            try { return Marshal.Invoke(addin, new object[] { operation }); }
            catch (TargetInvocationException error) { throw error.InnerException; }
        }
        internal void Post(Action operation)
        {
            Healthy();
            control.BeginInvoke((Action)delegate { try { operation(); } catch (Exception error) { failure = error; } });
        }
        internal void Drain()
        {
            using (var drained = new ManualResetEvent(false))
            {
                Post(() => drained.Set());
                if (!drained.WaitOne(5000)) throw new TimeoutException("The STA dispatch queue did not drain after release.");
            }
            Healthy();
        }
        private void Healthy() { if (failure != null) throw new InvalidOperationException("The hidden STA fixture failed.", failure); }
        public void Dispose()
        {
            stopping = true;
            if (control != null && !control.IsDisposed) control.BeginInvoke((Action)(() => Application.ExitThread()));
            if (!thread.Join(5000)) throw new TimeoutException("The hidden STA thread did not terminate.");
            ready.Dispose(); Healthy();
        }
    }

    private static void Result(string status, string error)
    { File.WriteAllText(Path.Combine(output, "result.json"), Json.Serialize(new { status = status, checks = Checks, timings = Timings, error = error, realWord = "NOT STARTED", productionTimeoutSeconds = 15 })); }
    private static void Check(bool condition, string message)
    { if (!condition) throw new Exception(message); Checks.Add(message); Console.WriteLine("PASS: " + message); }
}
