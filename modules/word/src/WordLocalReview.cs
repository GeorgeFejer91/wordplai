using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace UnderTree.WordAddin
{
    internal sealed partial class WordEditor
    {
        private object SelectionPreview()
        {
            WordDocumentReview state = Current();
            WordSnapshot snapshot = state.Snapshot;
            dynamic selection = application.Selection.Range;
            if ((int)selection.StoryType != 1) throw new InvalidOperationException("Select prose in the main document body.");
            int from = snapshot.VisibleOffset((int)selection.Start), to = snapshot.VisibleOffset((int)selection.End);
            snapshot.RequireProse(state.Document, from, to, OwnedBookmarks(state), null, true);
            selectionProof = snapshot.Revision;
            selectionDocument = state.Id;
            selectionFrom = from;
            selectionTo = to;
            return new { quote = snapshot.MainText.Substring(from, to - from), from = from, to = to, snapshot = snapshot.Revision };
        }

        private object SaveRule(IDictionary<string, object> args)
        {
            WordDocumentReview state = Current();
            string action = Text(args, "action");
            var rule = new ColorRule { Id = Text(args, "id", Guid.NewGuid().ToString("N")), Label = Text(args, "label"),
                Color = Text(args, "color"), Action = action, Prompt = Text(args, "prompt", ""), AccessMode = Text(args, "accessMode", action == "prompt" ? "direct" : "suggestions") };
            if (String.IsNullOrEmpty(rule.Id)) rule.Id = Guid.NewGuid().ToString("N");
            state.Policy.ValidateRule(rule);
            RequireWritable(state);
            PolicyEngine candidate = PrepareLocalPolicy(state, policy => policy.AddOrUpdateRule(rule));
            state.Policy = candidate;
            RepaintMarks(state);
            Persist(state, true);
            state.NativeEdit.Clear(); state.NativeEdit.Capture(state.Document, state.Policy.ExportState());
            return State();
        }

        private object MarkSelection(IDictionary<string, object> args)
        {
            WordDocumentReview state = Current();
            WordSnapshot snapshot = state.Snapshot;
            if ((int)application.Selection.Range.StoryType != 1) throw new InvalidOperationException("Select prose in the main document body.");
            int from = Number(args, "from"), to = Number(args, "to");
            if (state.Id != selectionDocument || snapshot.Revision != selectionProof || Text(args, "snapshot") != selectionProof ||
                from != selectionFrom || to != selectionTo || snapshot.VisibleOffset((int)application.Selection.Start) != from || snapshot.VisibleOffset((int)application.Selection.End) != to)
                throw new InvalidOperationException("The document or selection changed. Select and preview the passage again.");
            snapshot.RequireProse(state.Document, from, to, OwnedBookmarks(state), null, true);
            string quote = snapshot.MainText.Substring(from, to - from);
            if (Text(args, "quote") != quote) throw new InvalidOperationException("The selected wording changed.");
            string ruleId = Text(args, "ruleId");
            ColorRule rule = state.Policy.Rules.FirstOrDefault(r => r.Id == ruleId);
            if (rule == null) throw new InvalidOperationException("Choose a current color rule.");
            var mark = new PassageMark { Id = "Wrai_" + Guid.NewGuid().ToString("N"), RuleId = ruleId, From = from, To = to, Quote = quote, Prompt = Text(args, "prompt", "") };
            PolicyEngine candidate = PrepareLocalPolicy(state, policy => policy.AddMark(mark, snapshot.MainText));
            state.Document.Bookmarks.Add(mark.Id, snapshot.Range(state.Document, from, to));
            state.Policy = candidate;
            RepaintMarks(state);
            Persist(state, true);
            state.NativeEdit.Clear(); state.NativeEdit.Capture(state.Document, state.Policy.ExportState());
            selectionProof = null;
            return State();
        }

        private object RemoveMark(IDictionary<string, object> args)
        {
            WordDocumentReview state = Current();
            RequireWritable(state);
            string id = Text(args, "id");
            PassageMark mark = state.Policy.Marks.FirstOrDefault(m => m.Id == id);
            if (mark == null) throw new InvalidOperationException("The marker no longer exists.");
            if ((bool)state.Document.Bookmarks.Exists(MarkBookmark(mark))) state.Document.Bookmarks.Item(MarkBookmark(mark)).Delete();
            state.Policy.RemoveMark(id);
            RepaintMarks(state);
            Persist(state, true);
            state.NativeEdit.Clear(); state.NativeEdit.Capture(state.Document, state.Policy.ExportState());
            return State();
        }

        private object SelectMark(IDictionary<string, object> args)
        {
            WordDocumentReview state = Current();
            PassageMark mark = state.Policy.Marks.FirstOrDefault(m => m.Id == Text(args, "id"));
            if (mark == null || mark.Stale) throw new InvalidOperationException("This marker is stale. Mark the passage again locally.");
            state.Snapshot.Range(state.Document, mark.From, mark.To).Select();
            return State();
        }

        private object SelectionRules(IDictionary<string, object> args)
        {
            WordDocumentReview state = Current();
            dynamic range;
            if (args.ContainsKey("nativeFrom") || args.ContainsKey("nativeTo"))
            {
                int nativeFrom = Number(args, "nativeFrom"), nativeTo = Number(args, "nativeTo");
                if (nativeFrom < (int)state.Document.Content.Start || nativeTo < nativeFrom || nativeTo > (int)state.Document.Content.End)
                    throw new InvalidOperationException("The clicked text is outside the current document body.");
                range = state.Document.Range(nativeFrom, nativeTo);
            }
            else range = application.Selection.Range;
            if ((int)range.StoryType != 1 || WordSnapshot.Identity((object)range.Document) != WordSnapshot.Identity((object)state.Document))
                throw new InvalidOperationException("Select a highlighted passage in the main document body.");
            int from = (int)range.Start, to = (int)range.End;
            var choices = new List<object>();
            foreach (PassageMark mark in state.Policy.Marks)
            {
                string bookmark = MarkBookmark(mark);
                // Rule navigation can survive changed wording; editing still requires a current quote.
                if (bookmark != mark.Id && !state.Policy.Proposals.Any(proposal => proposal.NativeOriginalMarks != null &&
                    proposal.NativeOriginalMarks.Any(witness => witness.Id == mark.Id && witness.Bookmark == bookmark))) continue;
                if (!(bool)state.Document.Bookmarks.Exists(bookmark)) continue;
                dynamic anchor = state.Document.Bookmarks.Item(bookmark).Range;
                if ((int)anchor.StoryType != 1 || WordSnapshot.Identity((object)anchor.Document) != WordSnapshot.Identity((object)state.Document) ||
                    (int)anchor.Start < (int)state.Document.Content.Start || (int)anchor.End > (int)state.Document.Content.End ||
                    (int)anchor.End <= (int)anchor.Start) continue;
                int start = (int)anchor.Start, end = (int)anchor.End;
                if (!(to == from ? start <= from && from < end : start < to && from < end)) continue;
                ColorRule rule = state.Policy.Rules.FirstOrDefault(candidate => candidate.Id == mark.RuleId);
                if (rule == null) continue;
                choices.Add(new { id = mark.Id, ruleId = mark.RuleId, quote = mark.Quote, stale = mark.Stale,
                    label = rule.Label, action = rule.Action });
            }
            return new { documentId = state.Id, revision = state.Snapshot.Revision, choices = choices.ToArray(),
                rules = state.Policy.Rules.Select(rule => new { id = rule.Id, label = rule.Label, color = rule.Color,
                    action = rule.Action, prompt = rule.Prompt, accessMode = rule.AccessMode }).ToArray() };
        }

        private object SelectProposal(IDictionary<string, object> args)
        {
            WordDocumentReview state = Current();
            Proposal proposal = state.Policy.Proposals.FirstOrDefault(p => p.Id == Text(args, "id"));
            WordSnapshot snapshot = state.Snapshot;
            if (proposal == null || proposal.Revision != snapshot.Revision) throw new InvalidOperationException("This proposal's document snapshot is stale.");
            snapshot.Range(state.Document, proposal.From, proposal.To).Select();
            return State();
        }

        private object Reject(IDictionary<string, object> args)
        {
            WordDocumentReview state = Current();
            state.Policy.Reject(Text(args, "id"));
            Persist(state, false);
            return State();
        }

        private static void RequireWritable(WordDocumentReview state)
        {
            if ((bool)state.Document.ReadOnly || (int)state.Document.ProtectionType != -1)
                throw new InvalidOperationException("This document is read-only or protected.");
        }

        private static void Paint(dynamic range, string color)
        {
            Color value = ColorTranslator.FromHtml(color);
            range.Shading.BackgroundPatternColor = value.R | (value.G << 8) | (value.B << 16);
        }
    }
}

