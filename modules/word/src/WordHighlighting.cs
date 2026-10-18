using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace UnderTree.WordAddin
{
    internal sealed partial class WordEditor
    {
        private object HighlightText(WordDocumentReview state, WordSnapshot snapshot, IDictionary<string, object> args, string session)
        {
            OnlyKeys(args, "documentId", "marking");
            var marking = Map(args, "marking");
            OnlyKeys(marking, "requestId", "contextId", "revision", "colorId", "color", "ranges", "recolorMarkIds");
            object supplied;
            if (!marking.TryGetValue("ranges", out supplied) || !(supplied is IEnumerable) || supplied is string || supplied is IDictionary)
                throw new ArgumentException("ranges must be an array.");
            var ranges = new List<HighlightRange>();
            foreach (object entry in (IEnumerable)supplied)
            {
                var range = entry as IDictionary<string, object>;
                if (range == null) throw new ArgumentException("Invalid marking range.");
                OnlyKeys(range, "from", "to", "quote");
                ranges.Add(new HighlightRange { From = Number(range, "from"), To = Number(range, "to"), Quote = Text(range, "quote") });
            }
            List<string> recolorMarkIds = null;
            if (marking.TryGetValue("recolorMarkIds", out supplied))
            {
                if (!(supplied is IEnumerable) || supplied is string || supplied is IDictionary)
                    throw new ArgumentException("recolorMarkIds must be an array of marker IDs.");
                recolorMarkIds = new List<string>();
                foreach (object entry in (IEnumerable)supplied)
                {
                    if (!(entry is string)) throw new ArgumentException("recolorMarkIds must contain only marker IDs.");
                    recolorMarkIds.Add((string)entry);
                }
            }
            string requestId = Text(marking, "requestId");
            var request = new HighlightRequest { DocumentId = Text(args, "documentId"), RequestId = ClientRequestId(session, requestId),
                Revision = Text(marking, "revision"), ContextId = Text(marking, "contextId"),
                ColorId = marking.ContainsKey("colorId") ? Text(marking, "colorId") : null,
                Color = marking.ContainsKey("color") ? Text(marking, "color") : null, Ranges = ranges,
                RecolorMarkIds = recolorMarkIds };
            if (state.Policy.UncertainReplay(request)) throw WordNativeOutcome.UnknownException();
            bool replay = state.Policy.HighlightReceipts.Any(receipt => receipt.Request.RequestId == request.RequestId);
            if (!replay && request.Revision != snapshot.Revision) throw new InvalidOperationException("The manuscript changed; read complete context again.");
            string beforeNative = state.NativeEdit.NativeIdentity(state.Document, snapshot);
            ISet<string> ownedBookmarks = OwnedBookmarks(state);
            foreach (HighlightRange range in ranges) snapshot.RequireProse(state.Document, range.From, range.To, ownedBookmarks, null, true);
            if ((int)application.UndoRecord.CustomRecordLevel != 0)
                throw new InvalidOperationException("Finish the current Word operation before marking text.");
            IntPtr colorFrame = IntPtr.Zero;
            WordColorInput.Sample admission = null;
            if (!replay)
            {
                colorFrame = new IntPtr((int)application.ActiveWindow.Hwnd);
                admission = coloringInput.RequireQuiet(colorFrame);
            }
            bool nativeEffects = false, committed = false;
            try
            {
            using (var window = new WordReviewWindow(application, state.Document))
            {
            PolicyEngine original = state.Policy;
            PolicyState beforePolicy = original.ExportState();
            state.NativeEdit.Capture(state.Document, beforePolicy, snapshot);
            HighlightReceipt result;
            if (replay)
            {
                result = original.Highlight(request, snapshot.MainText, snapshot.Revision, session);
                committed = true; reviewCommitted = true; return HighlightResult(state, result, true, requestId);
            }
            bool ended = false, started = false, staged = false, mixedRecord = false;
            try
            {
                WordSnapshot beforeCommit = WordSnapshot.Capture(state.Document);
                if (beforeCommit.Revision != snapshot.Revision || !snapshot.SameProseInspection(beforeCommit) ||
                    state.NativeEdit.NativeIdentity(state.Document, beforeCommit) != beforeNative)
                    throw new InvalidOperationException("The native document changed before marking.");
                colorFrame = new IntPtr((int)application.ActiveWindow.Hwnd);
                admission = coloringInput.RequireStable(colorFrame, admission);
                result = original.Highlight(request, snapshot.MainText, snapshot.Revision, session); staged = true;
                WordReviewStore.ValidateReceipts(new SavedWordReview { DocumentId = state.Id, Policy = state.Policy.ExportState() });
                original.ReserveHighlight(request.RequestId);
                Persist(state, false); // Consume the identity durably before native effects.
                application.UndoRecord.StartCustomRecord(AppBranding.Name + " AI color only"); started = true; nativeEffects = true;
                if (!coloringInput.IsStable(colorFrame, admission))
                {
                    mixedRecord = true;
                    throw new InvalidOperationException("Input changed after Word began the coloring record.");
                }
                foreach (string id in result.ReplacedMarkIds ?? new List<string>())
                    state.Document.Bookmarks.Item(id).Delete();
                foreach (string id in result.MarkIds)
                {
                    PassageMark mark = state.Policy.Marks.First(candidate => candidate.Id == id);
                    state.Document.Bookmarks.Add(mark.Id, snapshot.Range(state.Document, mark.From, mark.To));
                }
                RepaintMarks(state, state.Policy.Marks.Where(mark => result.MarkIds.Contains(mark.Id)), snapshot,
                    result.ReplacedMarkIds != null);
                if (!coloringInput.IsStable(colorFrame, admission))
                {
                    mixedRecord = true;
                    throw new InvalidOperationException("Input changed during native coloring.");
                }
                state.NativeEdit.AddGuard(state.Document);
                Persist(state, true); // Native Undo includes staged metadata; the durable receipt is still reserved.
                if ((int)application.UndoRecord.CustomRecordLevel != 1)
                    throw new InvalidOperationException("Word ended the marking Undo record before the native transaction completed.");
                application.UndoRecord.EndCustomRecord(); ended = true;
                WordSnapshot confirmed = WordSnapshot.Capture(state.Document);
                if (confirmed.FullText != snapshot.FullText || !coloringInput.IsStable(colorFrame, admission))
                {
                    mixedRecord = true;
                    throw new InvalidOperationException("Native coloring could not exclude intervening input or manuscript changes.");
                }
                original.CommitHighlightSnapshot(request.RequestId, confirmed.Revision);
                committed = true; reviewCommitted = true;
                state.NativeEdit.Capture(state.Document, state.Policy.ExportState(), confirmed);
                Persist(state, false); // Only the private replay receipt changes after native verification.
                result = original.HighlightReceipts.First(receipt => receipt.Request.RequestId == request.RequestId);
                return HighlightResult(state, result, false, requestId);
            }
            catch
            {
                if (committed) throw;
                if (!staged)
                {
                    state.Policy = CreatePolicy(beforePolicy);
                    state.Policy.Revoke(); state.ReadRevision = null; state.NativeEdit.Clear();
                    throw;
                }
                if (started && !ended) application.UndoRecord.EndCustomRecord(); ended = true;
                original.FailHighlight(request.RequestId);
                bool restored = !started;
                if (started) try
                {
                    if (!mixedRecord)
                    {
                        WordSnapshot afterFailure = WordSnapshot.Capture(state.Document);
                        if (coloringInput.CanRollback(colorFrame, admission, snapshot.FullText, afterFailure.FullText))
                        {
                            restored = state.NativeEdit.NativeIdentity(state.Document, afterFailure) == beforeNative;
                            if (!restored)
                            {
                                WordSnapshot beforeUndo = WordSnapshot.Capture(state.Document);
                                if (coloringInput.CanRollback(colorFrame, admission, snapshot.FullText, beforeUndo.FullText))
                                    restored = (bool)state.Document.Undo(1) && state.NativeEdit.NativeIdentity(state.Document) == beforeNative;
                            }
                        }
                    }
                }
                catch { restored = false; }
                if (!restored)
                {
                    throw new InvalidOperationException("The marking outcome needs inspection in Word before continuing; its request identity remains consumed.");
                }
                nativeEffects = false;
                beforePolicy.HighlightReceipts = original.ExportState().HighlightReceipts;
                state.Policy = CreatePolicy(beforePolicy);
                state.Policy.Revoke(); state.ReadRevision = null;
                Persist(state, started);
                state.NativeEdit.Clear(); state.NativeEdit.Capture(state.Document, state.Policy.ExportState());
                throw;
            }
            finally { if (started && !ended) application.UndoRecord.EndCustomRecord(); }
            }
            }
            catch
            {
                if (committed) throw WordNativeOutcome.UnknownException();
                if (nativeEffects) WordNativeOutcome.Unknown(
                    delegate { state.Policy.FailHighlight(request.RequestId); }, state.Policy.Revoke, delegate { Persist(state, false); });
                throw;
            }
        }

        private static object HighlightResult(WordDocumentReview state, HighlightReceipt receipt, bool replay, string requestId)
        {
            ColorRule rule = state.Policy.Rules.First(candidate => candidate.Id == receipt.ColorId);
            return new { requestId = requestId, colorId = receipt.ColorId, color = rule.Color,
                markIds = receipt.MarkIds.ToArray(), replacedMarkIds = (receipt.ReplacedMarkIds ?? new List<string>()).ToArray(),
                action = rule.Action, replayed = replay };
        }

        private static List<ReplacementPart> ParseReplacementParts(object supplied)
        {
            if (!(supplied is IEnumerable) || supplied is string || supplied is IDictionary) throw new ArgumentException("replacementParts must be an array.");
            var parts = new List<ReplacementPart>();
            foreach (object entry in (IEnumerable)supplied)
            {
                var part = entry as IDictionary<string, object>;
                if (part == null || part.Count != 1) throw new ArgumentException("Each replacement part contains only text or protectedId.");
                OnlyKeys(part, "text", "protectedId");
                parts.Add(part.ContainsKey("text") ? new ReplacementPart { Text = Text(part, "text") } : new ReplacementPart { ProtectedId = Text(part, "protectedId") });
            }
            return parts;
        }

        private static IDictionary<string, object> Map(IDictionary<string, object> values, string key)
        {
            object value;
            if (!values.TryGetValue(key, out value) || !(value is IDictionary<string, object>)) throw new ArgumentException("Invalid " + key + ".");
            return (IDictionary<string, object>)value;
        }

        private static void OnlyKeys(IDictionary<string, object> values, params string[] allowed)
        {
            if (values.Keys.Any(key => !allowed.Contains(key))) throw new ArgumentException("Unknown or policy-changing review argument.");
        }

        private void RepaintMarks(WordDocumentReview state, IEnumerable<PassageMark> marks = null, WordSnapshot snapshot = null,
            bool preserveMatchingShading = false)
        {
            snapshot = snapshot ?? WordSnapshot.CaptureTextPositions(state.Document);
            var repaint = (marks ?? state.Policy.Marks).Where(mark => !mark.Stale).ToList();
            var unchanged = marks == null ? new List<PassageMark>() : state.Policy.Marks.Where(mark => !mark.Stale &&
                !repaint.Any(candidate => candidate.Id == mark.Id) && state.Policy.Rules.Any(rule => rule.Id == mark.RuleId &&
                    (rule.Action == "protect" || rule.Action == "reject"))).OrderBy(mark => mark.From).ToList();
            bool originalTrackFormatting = (bool)state.Document.TrackFormatting;
            try
            {
                state.Document.TrackFormatting = false;
                foreach (PassageMark mark in repaint.OrderBy(mark => {
                    string action = state.Policy.Rules.First(rule => rule.Id == mark.RuleId).Action;
                    return action == "protect" || action == "reject" ? 2 : action == "neutral" ? 0 : 1;
                }))
                {
                    string color = state.Policy.Rules.First(rule => rule.Id == mark.RuleId).Color;
                    int cursor = mark.From;
                    foreach (PassageMark unit in unchanged.Where(unit => unit.From < mark.To && mark.From < unit.To))
                    {
                        if (cursor < unit.From) foreach (int[] span in snapshot.NativeSegments(cursor, unit.From)) PaintHighlightRange(state.Document.Range(span[0], span[1]), color, preserveMatchingShading, marks != null);
                        cursor = Math.Max(cursor, unit.To);
                    }
                    if (cursor < mark.To) foreach (int[] span in snapshot.NativeSegments(cursor, mark.To)) PaintHighlightRange(state.Document.Range(span[0], span[1]), color, preserveMatchingShading, marks != null);
                }
            }
            finally { state.Document.TrackFormatting = originalTrackFormatting; }
        }

        private void PaintHighlightRange(dynamic range, string color, bool preserveMatchingShading, bool verifyNativeColor)
        {
            if (!verifyNativeColor) { Paint(range, color); return; }
            System.Drawing.Color value = System.Drawing.ColorTranslator.FromHtml(color);
            int expected = value.R | (value.G << 8) | (value.B << 16);
            if (preserveMatchingShading && NativeColorMatches(range, expected)) return;
            range.Font.Shading.BackgroundPatternColor = expected;
            if (!NativeColorMatches(range, expected))
                throw new McpOperationException("Word could not confirm the requested color across the complete marked range.",
                    "native_color_not_verified", new { color = color });
        }
    }
}
