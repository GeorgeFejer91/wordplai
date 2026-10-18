using System;
using System.Collections.Generic;
using System.Linq;

namespace UnderTree.WordAddin
{
    internal sealed partial class WordEditor
    {
        private const string AcceptedGreen = "#74b9a4";

        private sealed class AcceptedProtectionPlan
        {
            internal ColorRule Rule;
            internal readonly List<PassageMark> Marks = new List<PassageMark>();
        }

        private static AcceptedProtectionPlan PlanAcceptedProtection(WordDocumentReview state, int from, string quote,
            IEnumerable<ProtectedPlacement> retained)
        {
            var plan = new AcceptedProtectionPlan();
            int end = from + quote.Length, cursor = from;
            foreach (ProtectedPlacement unit in retained.Where(unit => unit.From < end && from < unit.To).OrderBy(unit => unit.From))
            {
                if (cursor < unit.From) AddAcceptedGap(plan, quote, from, cursor, Math.Min(unit.From, end));
                cursor = Math.Max(cursor, Math.Min(unit.To, end));
            }
            if (cursor < end) AddAcceptedGap(plan, quote, from, cursor, end);
            if (plan.Marks.Count == 0) return plan;
            plan.Rule = state.Policy.Rules.FirstOrDefault(rule => rule.Action == "protect" &&
                String.Equals(rule.Color, AcceptedGreen, StringComparison.OrdinalIgnoreCase));
            if (plan.Rule == null)
            {
                if (state.Policy.Rules.Count >= 24)
                    throw new InvalidOperationException("Accepted wording cannot be protected until an unused color rule is removed locally.");
                plan.Rule = new ColorRule { Id = "accepted_" + Guid.NewGuid().ToString("N"), Label = "Accepted wording",
                    Color = AcceptedGreen, Action = "protect", AccessMode = "suggestions", Prompt = "" };
            }
            if (state.Policy.Marks.Count + plan.Marks.Count > 500)
                throw new InvalidOperationException("Accepted wording cannot be protected until unused passage markers are removed locally.");
            foreach (PassageMark mark in plan.Marks) mark.RuleId = plan.Rule.Id;
            return plan;
        }

        private static void AddAcceptedGap(AcceptedProtectionPlan plan, string quote, int origin, int from, int to)
        {
            if (to <= from) return;
            plan.Marks.Add(new PassageMark { Id = "Wrai_" + Guid.NewGuid().ToString("N"), From = from, To = to,
                Quote = quote.Substring(from - origin, to - from), Prompt = "" });
        }

        private static void ApplyAcceptedProtection(WordDocumentReview state, WordSnapshot snapshot,
            Proposal receipt, AcceptedProtectionPlan plan, string quote)
        {
            if (plan.Rule != null && !state.Policy.Rules.Any(rule => rule.Id == plan.Rule.Id)) state.Policy.AddOrUpdateRule(plan.Rule);
            foreach (PassageMark mark in plan.Marks)
            {
                state.Policy.AddMark(mark, snapshot.MainText);
                dynamic range = snapshot.Range(state.Document, mark.From, mark.To);
                state.Document.Bookmarks.Add(mark.Id, range);
                range.HighlightColorIndex = 0;
                range.Shading.Texture = 0;
                Paint(range, plan.Rule.Color);
            }
            receipt.NativeAcceptanceQuote = quote;
            receipt.NativeAcceptedMarkIds = plan.Marks.Select(mark => mark.Id).ToList();
            state.Policy.InvalidateContext();
            state.ReadRevision = null;
        }

        private static void RecordAcceptanceCandidate(WordDocumentReview state, WordSnapshot snapshot,
            Proposal receipt, string replacement, string originalParagraph)
        {
            string bookmark = "WraiA_" + Guid.NewGuid().ToString("N");
            dynamic range = snapshot.Range(state.Document, receipt.From, receipt.From + replacement.Length);
            state.Document.Bookmarks.Add(bookmark, range);
            receipt.NativeAcceptanceBookmark = bookmark;
            receipt.NativeAcceptanceQuote = replacement;
            receipt.NativeOriginalParagraphQuote = originalParagraph;
        }

        private bool ReconcileAcceptedChanges(WordDocumentReview state, WordSnapshot snapshot)
        {
            if (state.Policy.Marks.Any(mark => mark.Stale && state.Policy.Rules.Any(rule => rule.Id == mark.RuleId &&
                (rule.Action == "protect" || rule.Action == "reject" || rule.Action == "discard")))) return false;
            bool changed = false;
            foreach (Proposal receipt in state.Policy.Proposals.Where(proposal => proposal.NativeAcceptanceBookmark != null &&
                proposal.NativeAcceptedMarkIds == null && (proposal.Status == "tracked" || proposal.Status == "applied")))
            {
                if (!(bool)state.Document.Bookmarks.Exists(receipt.NativeReviewBookmark)) continue;
                dynamic review = state.Document.Bookmarks.Item(receipt.NativeReviewBookmark).Range;
                if ((int)review.StoryType != 1 || WordSnapshot.Identity((object)review.Document) != WordSnapshot.Identity((object)state.Document)) continue;
                dynamic paragraph = review.Paragraphs.First.Range;
                if ((int)review.Start != (int)paragraph.Start || ((int)review.End != (int)paragraph.End &&
                    (int)review.End != (int)paragraph.End - 1) || (int)paragraph.Revisions.Count != 0) continue;
                int paragraphFrom, paragraphTo;
                try { paragraphFrom = snapshot.VisibleOffset((int)paragraph.Start); paragraphTo = snapshot.VisibleOffset((int)paragraph.End); }
                catch (InvalidOperationException) { continue; }
                string paragraphQuote = PolicyEngine.Slice(snapshot.MainText, paragraphFrom, paragraphTo);
                if (paragraphQuote == receipt.NativeOriginalParagraphQuote && paragraphQuote != receipt.NativeReviewQuote)
                {
                    receipt.Status = "rejected";
                    receipt.NativeAcceptedMarkIds = new List<string>();
                    Persist(state, false);
                    continue;
                }
                if (paragraphQuote != receipt.NativeReviewQuote || !(bool)state.Document.Bookmarks.Exists(receipt.NativeAcceptanceBookmark)) continue;
                dynamic accepted = state.Document.Bookmarks.Item(receipt.NativeAcceptanceBookmark).Range;
                if ((int)accepted.StoryType != 1 || (int)accepted.Start < (int)paragraph.Start || (int)accepted.End > (int)paragraph.End ||
                    WordSnapshot.Identity((object)accepted.Document) != WordSnapshot.Identity((object)state.Document)) continue;
                int from, to;
                try { from = snapshot.VisibleOffset((int)accepted.Start); to = snapshot.VisibleOffset((int)accepted.End); }
                catch (InvalidOperationException) { continue; }
                if (PolicyEngine.Slice(snapshot.MainText, from, to) != receipt.NativeAcceptanceQuote) continue;
                var retained = state.Policy.Marks.Where(mark => !mark.Stale && state.Policy.Rules.Any(rule => rule.Id == mark.RuleId &&
                    (rule.Action == "protect" || rule.Action == "reject"))).Select(mark => new ProtectedPlacement { Id = mark.Id, From = mark.From, To = mark.To });
                AcceptedProtectionPlan plan = PlanAcceptedProtection(state, from, receipt.NativeAcceptanceQuote, retained);
                ProtectNativeAcceptance(state, snapshot, receipt, plan);
                changed = true;
            }
            return changed;
        }

        private void ProtectNativeAcceptance(WordDocumentReview state, WordSnapshot snapshot, Proposal receipt, AcceptedProtectionPlan plan)
        {
            if ((int)application.UndoRecord.CustomRecordLevel != 0)
                throw new InvalidOperationException("Finish the current Word operation before accepted wording is protected.");
            PolicyState beforePolicy = state.Policy.ExportState();
            state.NativeEdit.Capture(state.Document, beforePolicy);
            string beforeNative = state.NativeEdit.NativeIdentity(state.Document);
            bool tracking = (bool)state.Document.TrackRevisions, trackFormatting = (bool)state.Document.TrackFormatting;
            bool ended = false;
            application.UndoRecord.StartCustomRecord(AppBranding.Name + " protect accepted wording");
            try
            {
                state.Document.TrackFormatting = false;
                state.Document.TrackRevisions = false;
                ApplyAcceptedProtection(state, snapshot, receipt, plan, receipt.NativeAcceptanceQuote);
                receipt.Status = "applied";
                state.NativeEdit.AddGuard(state.Document);
                Persist(state, true);
                state.Document.TrackFormatting = trackFormatting;
                state.Document.TrackRevisions = tracking;
                if ((int)application.UndoRecord.CustomRecordLevel != 1)
                    throw new InvalidOperationException("Word ended the accepted-wording Undo record before protection completed.");
                application.UndoRecord.EndCustomRecord(); ended = true;
                WordSnapshot after = WordSnapshot.Capture(state.Document);
                if (after.MainText != snapshot.MainText || after.ContextStories != snapshot.ContextStories ||
                    after.Review.FeedbackIdentity() != snapshot.Review.FeedbackIdentity())
                    throw new InvalidOperationException("Protecting accepted wording did not preserve the manuscript and its feedback.");
                state.NativeEdit.Capture(state.Document, state.Policy.ExportState());
            }
            catch
            {
                if (!ended) { application.UndoRecord.EndCustomRecord(); ended = true; }
                bool restored = false;
                try { restored = state.NativeEdit.NativeIdentity(state.Document) == beforeNative ||
                    (bool)state.Document.Undo(1) && state.NativeEdit.NativeIdentity(state.Document) == beforeNative; }
                catch { restored = false; }
                if (!restored)
                {
                    state.Policy.RecordOutcomeUncertain(receipt.Id);
                    state.Policy.Revoke();
                    Persist(state, false);
                    throw new InvalidOperationException("Accepted-wording protection needs inspection in Word before continuing.");
                }
                state.Policy = CreatePolicy(beforePolicy);
                state.Policy.InvalidateContext(); state.ReadRevision = null;
                Persist(state, true);
                state.NativeEdit.Clear(); state.NativeEdit.Capture(state.Document, state.Policy.ExportState());
                throw;
            }
            finally
            {
                state.Document.TrackFormatting = trackFormatting;
                state.Document.TrackRevisions = tracking;
                if (!ended) application.UndoRecord.EndCustomRecord();
            }
        }
    }
}
