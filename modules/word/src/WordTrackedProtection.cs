using System;
using System.Collections.Generic;
using System.Linq;

namespace UnderTree.WordAddin
{
    // Native review reconciliation preserves current rules and never rewrites document content.
    internal sealed class WordTrackedProtection
    {
        private readonly List<NativeMarkWitness> witnesses = new List<NativeMarkWitness>();
        private readonly List<PassageMark> originals = new List<PassageMark>();
        private readonly HashSet<string> deletedMarks = new HashSet<string>(StringComparer.Ordinal);
        private readonly string paragraphQuote;

        internal WordTrackedProtection(WordDocumentReview state, WordSnapshot before, Proposal proposal)
        {
            dynamic paragraph = before.Range(state.Document, proposal.From, proposal.To).Paragraphs.First.Range;
            paragraphQuote = PolicyEngine.Slice(before.MainText, before.VisibleOffset((int)paragraph.Start), before.VisibleOffset((int)paragraph.End));
            foreach (PassageMark mark in state.Policy.Marks.Where(mark => !mark.Stale && mark.From < proposal.To && proposal.From < mark.To))
            {
                originals.Add(new PassageMark { Id = mark.Id, From = mark.From, To = mark.To });
                witnesses.Add(new NativeMarkWitness { Id = mark.Id, Bookmark = "WraiO_" + Guid.NewGuid().ToString("N"),
                    BeforeQuote = mark.Quote, BeforeSignature = state.Policy.Rules.Any(rule => rule.Id == mark.RuleId && (rule.Action == "protect" || rule.Action == "reject"))
                        ? WordSnapshot.FormattingSignature(before.Range(state.Document, mark.From, mark.To)) : null });
            }
        }

        internal void AddOriginalBookmarks(dynamic document, WordSnapshot before)
        {
            foreach (NativeMarkWitness witness in witnesses)
            {
                PassageMark mark = originals.First(candidate => candidate.Id == witness.Id);
                document.Bookmarks.Add(witness.Bookmark, before.Range(document, mark.From, mark.To));
            }
        }

        internal void RetainDeletedMark(PassageMark mark)
        {
            NativeMarkWitness witness = witnesses.First(candidate => candidate.Id == mark.Id);
            if (witness.BeforeSignature != null) throw new InvalidOperationException("A protected native unit cannot be removed.");
            deletedMarks.Add(mark.Id);
            mark.Stale = true;
        }

        internal void PrepareWitnesses(WordDocumentReview state, string proposalId)
        {
            foreach (NativeMarkWitness witness in witnesses)
            {
                PassageMark mark = state.Policy.Marks.First(candidate => candidate.Id == witness.Id);
                witness.AfterQuote = deletedMarks.Contains(mark.Id) ? "" : mark.Quote;
                witness.AfterSignature = witness.BeforeSignature; // Verification must later prove the native unit retained this exact signature.
            }
            Proposal receipt = state.Policy.Proposals.First(candidate => candidate.Id == proposalId);
            receipt.NativeOriginalParagraphQuote = paragraphQuote;
            receipt.NativeOriginalMarks = witnesses;
            foreach (PassageMark mark in state.Policy.Marks.Where(mark => deletedMarks.Contains(mark.Id)))
                mark.NativeBookmark = witnesses.First(witness => witness.Id == mark.Id).Bookmark;
        }
    }

    internal sealed partial class WordEditor
    {
        private void ReconcileProtectedNativeReview(WordDocumentReview state, WordSnapshot current)
        {
            var claimed = new HashSet<string>(StringComparer.Ordinal);
            foreach (Proposal receipt in state.Policy.Proposals.Where(proposal => proposal.NativeOriginalMarks != null &&
                proposal.NativeOriginalMarks.Count != 0 && new[] { "tracked", "applied", "rejected" }.Contains(proposal.Status)).Reverse())
            {
                string reviewBookmark = receipt.NativeReviewBookmark;
                if (!(bool)state.Document.Bookmarks.Exists(reviewBookmark))
                {
                    if (receipt.NativeAcceptanceQuote != "" || receipt.NativeAcceptanceBookmark == null ||
                        !(bool)state.Document.Bookmarks.Exists(receipt.NativeAcceptanceBookmark)) continue;
                    reviewBookmark = receipt.NativeAcceptanceBookmark;
                }
                dynamic reviewAnchor = state.Document.Bookmarks.Item(reviewBookmark).Range;
                if ((int)reviewAnchor.StoryType != 1) continue;
                dynamic paragraph = reviewAnchor.Paragraphs.First.Range;
                if ((int)reviewAnchor.Start != (int)paragraph.Start || ((int)reviewAnchor.End != (int)paragraph.End && (int)reviewAnchor.End != (int)paragraph.End - 1)) continue;
                int count = (int)paragraph.Revisions.Count;
                string quote = PolicyEngine.Slice(current.MainText, current.VisibleOffset((int)paragraph.Start), current.VisibleOffset((int)paragraph.End));
                bool revised = quote == receipt.NativeReviewQuote;
                bool original = count == 0 && quote == receipt.NativeOriginalParagraphQuote;
                if (count > 0 && (!revised || NativeRevisionHash(paragraph) != receipt.NativeReviewHash)) continue;
                if (!revised && !original) continue;
                bool restoring = original && !revised;
                string status = count > 0 ? "tracked" : restoring ? "rejected" : "applied";
                var updates = new List<Action>();
                bool changed = receipt.Status != status;
                foreach (NativeMarkWitness witness in receipt.NativeOriginalMarks)
                {
                    PassageMark mark = state.Policy.Marks.FirstOrDefault(candidate => candidate.Id == witness.Id);
                    if (mark == null || claimed.Contains(witness.Id)) continue; // Keep explicit removals and the newest valid owner.
                    bool deleted = !restoring && witness.AfterQuote.Length == 0;
                    string bookmark = restoring ? witness.Bookmark : deleted ? receipt.NativeAcceptanceBookmark : witness.Id;
                    if (bookmark == null) { updates.Clear(); break; }
                    if (!(bool)state.Document.Bookmarks.Exists(bookmark)) { updates.Clear(); break; }
                    dynamic range = state.Document.Bookmarks.Item(bookmark).Range;
                    if ((int)range.StoryType != 1 || (int)range.Revisions.Count != 0 && restoring) { updates.Clear(); break; }
                    string expectedQuote = restoring ? witness.BeforeQuote : witness.AfterQuote;
                    int from = current.VisibleOffset((int)range.Start), to = current.BookmarkEnd(state.Document, range, expectedQuote);
                    string expectedSignature = restoring ? witness.BeforeSignature : witness.AfterSignature;
                    if (PolicyEngine.Slice(current.MainText, from, to) != expectedQuote || expectedSignature != null && WordSnapshot.FormattingSignature(range) != expectedSignature)
                    { updates.Clear(); break; }
                    if (deleted)
                    {
                        if ((int)range.Start != (int)range.End) { updates.Clear(); break; }
                        bool restrictive = state.Policy.Rules.Any(rule => rule.Id == mark.RuleId && (rule.Action == "protect" || rule.Action == "reject"));
                        if (count == 0 && !restrictive)
                        { changed = true; updates.Add(delegate { state.Policy.RemoveMark(mark.Id); }); }
                        else
                        { changed |= !mark.Stale; updates.Add(delegate { mark.Stale = true; }); }
                        continue;
                    }
                    string anchor = restoring ? bookmark : null;
                    changed |= mark.From != from || mark.To != to || mark.Quote != expectedQuote || mark.NativeBookmark != anchor || mark.Stale;
                    updates.Add(delegate { mark.From = from; mark.To = to; mark.Quote = expectedQuote; mark.NativeBookmark = anchor; mark.Stale = false; mark.RequiresRemarking = false; });
                }
                int existing = receipt.NativeOriginalMarks.Count(witness => !claimed.Contains(witness.Id) && state.Policy.Marks.Any(mark => mark.Id == witness.Id));
                if (updates.Count != existing || existing == 0) continue;
                foreach (NativeMarkWitness witness in receipt.NativeOriginalMarks.Where(witness => state.Policy.Marks.Any(mark => mark.Id == witness.Id))) claimed.Add(witness.Id);
                if (!changed) continue;
                foreach (Action update in updates) update();
                receipt.Status = status;
                state.Policy.Revoke(); state.ReadRevision = null;
                Persist(state, true);
                state.NativeEdit.Clear(); state.NativeEdit.Capture(state.Document, state.Policy.ExportState(), current);
            }
        }
    }
}
