using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;

namespace UnderTree.WordAddin
{
    internal sealed partial class WordEditor
    {
        private int nativeColorProbes;
        private long nativeColorProbeDeadline;

        private sealed class NativeMarker
        {
            internal PassageMark Mark;
            internal ColorRule Rule;
            internal int From, To;
            internal string Quote, Issue;
        }

        private void RefreshMarks(WordDocumentReview state, WordSnapshot snapshot)
        {
            var native = state.Policy.Marks.Select(mark => LocateMarker(state, snapshot, mark)).ToList();
            state.MarkerIssues.Clear();
            bool changed = false;
            foreach (NativeMarker current in native)
            {
                PassageMark mark = current.Mark;
                if (current.Quote != null && current.To > current.From && current.To <= PolicyEngine.MaxText)
                {
                    changed |= mark.From != current.From || mark.To != current.To;
                    mark.From = current.From; mark.To = current.To;
                }
                string issue = current.Issue;
                if (issue == null && !MarkerColorMatches(state, snapshot, current, native)) issue = "visible_color_changed";
                bool stale = issue != null;
                changed |= mark.Stale != stale;
                mark.Stale = stale;
                if (stale) { state.MarkerIssues[mark.Id] = issue; continue; }
                changed |= mark.From != current.From || mark.To != current.To || mark.Quote != current.Quote;
                mark.From = current.From; mark.To = current.To; mark.Quote = current.Quote;
            }
            if (changed) { state.Policy.InvalidateContext(); state.ReadRevision = null; }
        }

        private static NativeMarker LocateMarker(WordDocumentReview state, WordSnapshot snapshot, PassageMark mark)
        {
            var current = new NativeMarker { Mark = mark, Rule = state.Policy.Rules.First(rule => rule.Id == mark.RuleId) };
            string bookmark = MarkBookmark(mark);
            if (!(bool)state.Document.Bookmarks.Exists(bookmark)) { current.Issue = "native_bookmark_missing"; return current; }
            dynamic range = state.Document.Bookmarks.Item(bookmark).Range;
            if ((int)range.StoryType != 1 || WordSnapshot.Identity((object)range.Document) != WordSnapshot.Identity((object)state.Document))
            { current.Issue = "native_story_changed"; return current; }
            if ((int)range.End <= (int)range.Start) { current.Issue = "native_bookmark_empty"; return current; }
            try
            {
                current.From = snapshot.VisibleOffset((int)range.Start);
                current.To = snapshot.BookmarkEnd(state.Document, range, mark.Quote);
                current.Quote = PolicyEngine.Slice(snapshot.MainText, current.From, current.To);
            }
            catch (InvalidOperationException) { current.Issue = "native_boundary_unresolved"; return current; }
            if (current.Quote.Length == 0 || current.Quote.Length > PolicyEngine.MaxReplacement)
            { current.Issue = "native_range_out_of_bounds"; return current; }
            if (mark.RequiresRemarking) { current.Issue = "partial_rewrite_ambiguous"; return current; }
            if (current.Quote != mark.Quote)
            {
                // Only Word's witnessed native-review reconciliation may change a tracked original unit.
                if (mark.NativeBookmark != null || (int)range.Revisions.Count != 0 || state.Policy.Proposals.Any(proposal =>
                    proposal.Status == "tracked" && proposal.NativeOriginalMarks != null && proposal.NativeOriginalMarks.Any(witness => witness.Id == mark.Id)))
                    current.Issue = "native_review_unresolved";
                else if (mark.Quote.IndexOfAny(new[] { '\r', '\v' }) >= 0 || current.Quote.IndexOfAny(new[] { '\r', '\v' }) >= 0)
                    current.Issue = "paragraph_boundary_changed";
            }
            return current;
        }

        private bool MarkerColorMatches(WordDocumentReview state, WordSnapshot snapshot, NativeMarker target, List<NativeMarker> native)
        {
            var overlapping = native.Where(marker => marker.Issue == null && marker.From < target.To && target.From < marker.To)
                .OrderBy(marker => marker.Rule.Action == "protect" || marker.Rule.Action == "reject" ? 2 : marker.Rule.Action == "neutral" ? 0 : 1).ToList();
            var boundaries = new SortedSet<int> { target.From, target.To };
            foreach (NativeMarker marker in overlapping)
            { boundaries.Add(Math.Max(target.From, marker.From)); boundaries.Add(Math.Min(target.To, marker.To)); }
            for (int index = target.From; index < target.To; index++)
                if (snapshot.MainText[index] == '\r' || snapshot.MainText[index] == '\v') { boundaries.Add(index); boundaries.Add(index + 1); }
            int[] cuts = boundaries.ToArray();
            for (int index = 1; index < cuts.Length; index++)
            {
                int from = cuts[index - 1], to = cuts[index];
                if (to == from + 1 && (snapshot.MainText[from] == '\r' || snapshot.MainText[from] == '\v')) continue;
                NativeMarker visible = overlapping.Last(marker => marker.From <= from && marker.To >= to);
                Color color = ColorTranslator.FromHtml(visible.Rule.Color);
                int expected = color.R | (color.G << 8) | (color.B << 16);
                foreach (int[] span in snapshot.NativeSegments(from, to))
                {
                    dynamic range = state.Document.Range(span[0], span[1]);
                    if (!NativeColorMatches(range, expected)) return false;
                }
            }
            return true;
        }

        private bool NativeColorMatches(dynamic range, int expected)
        {
            int character = (int)range.Font.Shading.BackgroundPatternColor;
            int highlight = (int)range.HighlightColorIndex;
            int paragraph = character == -16777216 ? (int)range.ParagraphFormat.Shading.BackgroundPatternColor : -16777216;
            if (character != 9999999 && highlight == 0 && (character == expected ||
                character == -16777216 && paragraph != 9999999 && paragraph == expected)) return true;
            if (highlight != 0 && highlight != 9999999 || character != 9999999 && character != expected && character != -16777216 ||
                character == -16777216 && paragraph != 9999999 && paragraph != expected) return false;
            int from = (int)range.Start, to = (int)range.End;
            if (to <= from + 1) return false;
            long now = Stopwatch.GetTimestamp();
            if (nativeColorProbeDeadline == 0) nativeColorProbeDeadline = now + Stopwatch.Frequency;
            if (nativeColorProbes >= 256 || now > nativeColorProbeDeadline)
                throw NativeColorInspectionLimit();
            nativeColorProbes += 2;
            int middle = from + (to - from) / 2;
            if (!NativeColorMatches(range.Document.Range(from, middle), expected)) return false;
            if (Stopwatch.GetTimestamp() > nativeColorProbeDeadline) throw NativeColorInspectionLimit();
            bool matches = NativeColorMatches(range.Document.Range(middle, to), expected);
            if (matches && Stopwatch.GetTimestamp() > nativeColorProbeDeadline) throw NativeColorInspectionLimit();
            return matches;
        }

        private static McpOperationException NativeColorInspectionLimit()
        {
            return new McpOperationException("Word could not verify this complex highlight within its inspection limit. Use smaller or more uniformly formatted marked passages.",
                "native_color_verification_limited", new { additionalProbes = 256, inspectionMilliseconds = 1000 });
        }

        private static void RequireResolvedMarkers(WordDocumentReview state, Proposal proposal)
        {
            var unresolved = state.Policy.Marks.Where(mark => mark.Stale &&
                (state.Policy.Rules.Any(rule => rule.Id == mark.RuleId && (rule.Action == "protect" || rule.Action == "reject" || rule.Action == "discard")) ||
                mark.From < proposal.To && proposal.From < mark.To)).ToArray();
            if (unresolved.Length == 0) return;
            string message = unresolved.Any(mark => state.MarkerIssues.ContainsKey(mark.Id) && state.MarkerIssues[mark.Id] == "native_review_unresolved")
                ? "This paragraph contains changed or unrelated revisions. Review them in Word before submitting another AI revision."
                : "Word cannot verify the current editing permission for a marked passage.";
            throw new McpOperationException(message, "marker_permission_unresolved",
                new { markers = unresolved.Select(mark => new { id = mark.Id, ruleId = mark.RuleId,
                    reason = state.MarkerIssues.ContainsKey(mark.Id) ? state.MarkerIssues[mark.Id] : "native_anchor_unresolved" }).ToArray(),
                    recovery = "Inspect the named passage in Word. Resolve native review changes, restore its original marker/color, or explicitly mark the intended passage locally. Protection remains enforced; another paragraph may be blocked when a protected anchor is unresolved." });
        }
    }
}
