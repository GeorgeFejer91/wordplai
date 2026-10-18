using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace UnderTree.WordAddin
{
    internal sealed partial class WordEditor
    {
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        private bool reviewCommitted;
        private WordNativeOutcome.Diagnostic reviewDiagnostic;
        private object Review(string tool, IDictionary<string, object> args)
        {
            if (operationActive) throw new McpOperationException("Word is finishing another review operation. Read the current context again when it completes.");
            operationActive = true;
            reviewCommitted = false;
            reviewDiagnostic = null;
            WordNativeOutcome.Stage phase = WordNativeOutcome.Stage.review_core;
            McpOperationException uncertain = null;
            try
            {
                try { return ReviewCore(tool, args); }
                catch (Exception error)
                {
                    if (reviewDiagnostic != null) reviewDiagnostic.Capture(reviewDiagnostic.Phase, error, reviewCommitted);
                    McpOperationException operationError = error as McpOperationException;
                    if (operationError != null && operationError.Code == "outcome_unknown") uncertain = operationError;
                    throw;
                }
                finally
                {
                    operationActive = false;
                    phase = WordNativeOutcome.Stage.review_snapshot_cleanup;
                    ClearSnapshots();
                    phase = WordNativeOutcome.Stage.review_changed_callback;
                    if ((tool == "suggest_text" || tool == "highlight_text") && ReviewChanged != null) ReviewChanged();
                }
            }
            catch (Exception error)
            {
                if (reviewCommitted)
                {
                    if (reviewDiagnostic == null) throw WordNativeOutcome.UnknownException();
                    reviewDiagnostic.Capture(phase, error, true);
                    throw WordNativeOutcome.UnknownException(reviewDiagnostic);
                }
                if (uncertain != null) throw uncertain;
                throw;
            }
            finally { reviewCommitted = false; reviewDiagnostic = null; }
        }

        private object ReviewCore(string tool, IDictionary<string, object> args)
        {
            try
            {
                args = new Dictionary<string, object>(args);
                string session = args.ContainsKey(ReviewSessionKey) ? Text(args, ReviewSessionKey) : "default";
                args.Remove(ReviewSessionKey);
                if (!connections.Status.LocalRunning) throw new InvalidOperationException("Start access before reviewing a document.");
                if (tool == "list_manuscripts")
                {
                    OnlyKeys(args);
                    WordDocumentReview chosen;
                    reviewSelections.TryGetValue(session, out chosen);
                    WordDocumentReview[] open = OpenDocumentReviews();
                    bool retained = IsOpen(chosen);
                    return new { selection_required = !retained && open.Length > 1, manuscripts = DocumentChoices(open),
                        selected_manuscript_id = retained ? chosen.Id : null,
                        title = retained ? (string)chosen.Document.Name : open.Length == 1 ? (string)open[0].Document.Name : null };
                }
                WordDocumentReview state = ReviewDocument(tool, args, session);
                if (state == null)
                    return new { selection_required = true, manuscripts = DocumentChoices(OpenDocumentReviews()),
                        instruction = "Several Word documents are open. If the user explicitly named exactly one matching returned document, read its returned manuscript_id immediately and continue. Otherwise call list_manuscripts once to display the inline document picker, then wait for the user to choose before reading or editing. Continue the original writing request using that manuscript_id, regardless of the active Word window." };
                if (tool == "read_manuscript" && !reviewSelections.ContainsKey(session) && reviewSelections.Count >= 64)
                    throw new McpOperationException("This connection has reached its conversation limit. Restart Word access before starting another writing conversation.", "conversation_limit");
                // Retained uncertainty is receipt-only; both strict lookups must throw before snapshot access.
                if (state.Policy != null && HasUncertainReceipt(tool, state, args, session))
                {
                    if (tool == "suggest_text") return Suggest(state, null, args, session);
                    if (tool == "highlight_text") return HighlightText(state, null, args, session);
                }
                state = Current(DocumentReview(state.Document));
                if (tool == "read_manuscript" && !state.Policy.IsShared)
                {
                    state.Policy.Share(state.Id);
                    state.ReadRevision = null;
                    state.ReadOffset = 0;
                }
                if (!state.Policy.IsShared && !HasUncertainReceipt(tool, state, args, session))
                    throw new InvalidOperationException("Read this chosen document again before continuing its review.");
                WordSnapshot snapshot = state.Snapshot;
                if (tool == "read_manuscript")
                {
                    object context = ReadContext(state, snapshot, args, session);
                    reviewSelections[session] = state;
                    state.AutomaticReview = true;
                    return context;
                }
                if (tool == "suggest_text")
                    return Suggest(state, snapshot, args, session);
                if (tool == "highlight_text") return HighlightText(state, snapshot, args, session);
                throw new InvalidOperationException("This tool is unavailable.");
            }
            catch (InvalidOperationException error)
            { if (reviewDiagnostic != null) reviewDiagnostic.Capture(reviewDiagnostic.Phase, error, reviewCommitted); throw new McpOperationException(error.Message); }
            catch (ArgumentException error)
            { if (reviewDiagnostic != null) reviewDiagnostic.Capture(reviewDiagnostic.Phase, error, reviewCommitted); throw new McpOperationException("Invalid review arguments."); }
        }

        private object ReadContext(WordDocumentReview state, WordSnapshot snapshot, IDictionary<string, object> args, string session)
        {
            OnlyKeys(args, "manuscript_id", "offset");
            int from = Number(args, "offset", 0);
            WordClientRead progress;
            if (from == 0)
            {
                if (!state.ClientReads.ContainsKey(session) && state.ClientReads.Count >= 64)
                    throw new McpOperationException("This document reached its active reading-client limit.", "conversation_limit");
                state.ClientReads[session] = new WordClientRead { Policy = state.Policy, Revision = snapshot.Revision, Generation = state.Policy.ContextGeneration };
            }
            if (!state.ClientReads.TryGetValue(session, out progress) || progress.Policy != state.Policy || progress.Generation != state.Policy.ContextGeneration ||
                progress.Revision != snapshot.Revision || from != progress.Offset || from < 0 || from > snapshot.FullText.Length)
                throw new InvalidOperationException("Read the complete current context sequentially from offset zero.");
            int to = Math.Min(from + 32768, snapshot.FullText.Length);
            if (to < snapshot.FullText.Length && Char.IsLowSurrogate(snapshot.FullText[to])) to--;
            progress.Offset = to;
            bool complete = to == snapshot.FullText.Length;
            string witness = complete ? state.Policy.ReadContext(state.Id, snapshot.Revision, snapshot.MainText, snapshot.Revision, session) : null;
            return new { manuscript_id = state.Id, documentId = state.Id, title = (string)state.Document.Name, revision = snapshot.Revision,
                text = snapshot.FullText.Substring(from, to - from), from = from, to = to, total = snapshot.FullText.Length,
                complete = complete, next_offset = complete ? (int?)null : to, context_id = witness,
                main_body_length = snapshot.MainText.Length, reviewMode = "direct",
                review = snapshot.Review,
                paragraph_instruction = "Replace one complete highlighted paragraph at a time, excluding its paragraph mark. A partial highlight permits only that highlighted fragment; never expand it into unmarked text. Read review.feedback and its replies together with the color prompts before revising. Submit document edits sequentially: each successful edit changes the revision, so the next passage uses its returned complete current_context and updated ranges. For manuscript_changed, use the complete error.details.current_context, or read the complete current document once when it is unavailable, then resume the remaining passage prompts.",
                highlights = state.Policy.Marks.Select(m => new { id = m.Id, ruleId = m.RuleId, from = m.From, to = m.To, prompt = m.Prompt, stale = m.Stale,
                    staleReason = state.MarkerIssues.ContainsKey(m.Id) ? state.MarkerIssues[m.Id] : null,
                    protectedId = state.Policy.Rules.Any(r => r.Id == m.RuleId && (r.Action == "protect" || r.Action == "reject")) ? m.Id : null }).ToArray(),
                rules = state.Policy.Rules.Select(r => new { id = r.Id, label = r.Label, color = r.Color, action = r.Action, prompt = r.Prompt, accessMode = r.AccessMode }).ToArray() };
        }

        private object Suggest(WordDocumentReview state, WordSnapshot snapshot, IDictionary<string, object> args, string session)
        {
            OnlyKeys(args, "manuscript_id", "request_id", "revision", "context_id", "from", "to", "quote", "replacement", "rationale", "replacementParts");
            string requestId = Text(args, "request_id");
            var proposal = new Proposal { Id = ClientRequestId(session, requestId), DocumentId = state.Id, Revision = Text(args, "revision"),
                From = Number(args, "from"), To = Number(args, "to"), Quote = Text(args, "quote"), Replacement = Text(args, "replacement"), Rationale = Text(args, "rationale", "") };
            if (args.ContainsKey("replacementParts")) proposal.ReplacementParts = ParseReplacementParts(args["replacementParts"]);
            if (state.Policy.UncertainReplay(proposal, Text(args, "context_id"))) throw WordNativeOutcome.UnknownException();
            if (state.Policy.Proposals.Any(p => p.Id == proposal.Id))
            {
                Authorization replay = state.Policy.Submit(proposal, Text(args, "context_id"), snapshot.MainText, snapshot.Revision, session);
                reviewCommitted = replay.Proposal.Status == "tracked" || replay.Proposal.Status == "applied";
                return SuggestResult(replay.Proposal, null, requestId, true);
            }
            if (proposal.Revision != snapshot.Revision)
                throw new McpOperationException("The manuscript changed. This request made no changes; continue the remaining task from the current context.",
                    "manuscript_changed", new {
                        manuscript_id = state.Id,
                        request_applied = false,
                        current_context = ContinuationContext(state, snapshot, session),
                        recovery = "Use complete error.details.current_context to update the remaining passage's quote, range, revision and context_id. If unavailable, read the current document completely once. Then resume the existing passage prompts, with document edits in sequence."
                    });
            RequireResolvedMarkers(state, proposal);
            RequireParagraphUnit(state, snapshot, proposal);
            snapshot.RequireProse(state.Document, proposal.From, proposal.To, OwnedBookmarks(state), state.Policy.ProtectedUnits(proposal, snapshot.MainText));
            if (!state.Policy.HasReadWitness(Text(args, "context_id"), session))
                throw new McpOperationException("The reading context was refreshed by another request. This request made no changes; continue from the current context.",
                    "context_changed", new { manuscript_id = state.Id, request_applied = false,
                        current_context = ContinuationContext(state, snapshot, session),
                        recovery = "Use the complete current_context, or read the chosen document completely once, then resume its remaining passage prompts sequentially." });
            PolicyState beforeRequest = state.Policy.ExportState();
            Authorization authorization;
            try
            {
                authorization = state.Policy.Submit(proposal, Text(args, "context_id"), snapshot.MainText, snapshot.Revision, session);
                Persist(state, false);
            }
            catch
            {
                state.Policy = CreatePolicy(beforeRequest);
                state.Policy.Share(state.Id);
                throw;
            }
            if (!authorization.AlreadyRecorded)
            {
                try { Apply(state, authorization.Proposal.Id, authorization.Mode == "direct", authorization.Mode != "direct"); }
                catch
                {
                    if (state.Policy.Proposals.Any(p => p.Id == proposal.Id && (p.Status == "authorized" || p.Status == "pending")))
                    {
                        state.Policy = CreatePolicy(beforeRequest);
                        state.Policy.Share(state.Id);
                        Persist(state, false);
                    }
                    throw;
                }
            }
            Proposal recorded = state.Policy.Proposals.First(p => p.Id == proposal.Id);
            if (reviewDiagnostic != null) reviewDiagnostic.Phase = WordNativeOutcome.Stage.continuation_context;
            return SuggestResult(recorded, ContinuationContext(state, state.Snapshot, session), requestId);
        }

        private static bool HasUncertainReceipt(string tool, WordDocumentReview state, IDictionary<string, object> args, string session)
        {
            if (tool == "suggest_text")
            {
                string id = ClientRequestId(session, Text(args, "request_id"));
                return state.Policy.Proposals.Any(receipt => receipt.Id == id && (receipt.Status == "applying" || receipt.Status == "outcome_uncertain"));
            }
            if (tool == "highlight_text")
            {
                string id = ClientRequestId(session, Text(Map(args, "marking"), "requestId"));
                return state.Policy.HighlightReceipts.Any(receipt => receipt.Request.RequestId == id && (receipt.Outcome == "reserved" || receipt.Outcome == "outcome_uncertain"));
            }
            return false;
        }

        private object ContinuationContext(WordDocumentReview state, WordSnapshot snapshot, string session)
        {
            return snapshot.FullText.Length <= 32768
                ? ReadContext(state, snapshot, new Dictionary<string, object> { { "manuscript_id", state.Id }, { "offset", 0 } }, session) : null;
        }

        private object Accept(IDictionary<string, object> args)
        {
            WordDocumentReview state = Current();
            Apply(state, Text(args, "id"), false);
            return State();
        }

        private static void AddSuggestionComment(WordDocumentReview state, WordSnapshot snapshot, Proposal proposal, int replacementLength)
        {
            string rationale = String.Join(" ", proposal.Rationale.Split((char[])null, StringSplitOptions.RemoveEmptyEntries));
            if (rationale.Length == 0) rationale = replacementLength == 0 ? "Remove this passage as requested." : "Clarify the wording while preserving the meaning.";
            if (rationale.Length > 240)
            {
                int length = Char.IsLowSurrogate(rationale[237]) ? 236 : 237;
                rationale = rationale.Substring(0, length).TrimEnd() + "...";
            }
            dynamic comment = state.Document.Comments.Add(snapshot.Range(state.Document, proposal.From, proposal.From + replacementLength), rationale);
            comment.Author = AppBranding.Name;
            comment.Initial = "AI";
        }

        private static int CommentAnchorSuffix(WordDocumentReview state, WordSnapshot before, Proposal proposal, string replacement, bool direct, bool tracked, List<PassageMark> protectedUnits)
        {
            if (!direct && (!tracked || (int)before.Range(state.Document, proposal.From, proposal.To).Paragraphs.First.Range.Revisions.Count == 0)) return 0;
            var feedback = new List<WordCommentContext>(before.Review.feedback);
            for (int index = 0; index < feedback.Count; index++) feedback.AddRange(feedback[index].replies);
            var affected = feedback.Where(comment => comment.storyType == 1 &&
                (comment.from == comment.to ? proposal.From <= comment.from && comment.from < proposal.To : comment.from < proposal.To && proposal.From < comment.to)).ToList();
            if (affected.Count == 0) return 0;
            // Word deletes a comment whose entire inserted anchor is replaced. Keep its original shared ending.
            int retained = 0;
            while (retained < proposal.Quote.Length && retained < replacement.Length &&
                proposal.Quote[proposal.Quote.Length - retained - 1] == replacement[replacement.Length - retained - 1]) retained++;
            foreach (PassageMark unit in protectedUnits)
                if (unit.To > proposal.To - retained) retained = Math.Min(retained, proposal.To - unit.To);
            if (retained > 0 && Char.IsLowSurrogate(proposal.Quote[proposal.Quote.Length - retained])) retained--;
            int endingFrom = proposal.To - retained;
            if (retained == 0 || affected.Any(comment => comment.from == comment.to ? comment.from < endingFrom : comment.to <= endingFrom))
                throw new InvalidOperationException("Word cannot preserve this native comment anchor for that replacement. Keep a shared ending within the commented text, or review the change locally.");
            return retained;
        }
        private void Apply(WordDocumentReview state, string id, bool direct, bool tracked = false)
        {
            bool nativeEffects = false, committed = false;
            var diagnostic = new WordNativeOutcome.Diagnostic { Phase = WordNativeOutcome.Stage.window_prepare };
            if (tracked) reviewDiagnostic = diagnostic;
            try
            {
            using (var window = new WordReviewWindow(application, state.Document))
            {
            diagnostic.Phase = WordNativeOutcome.Stage.native_prepare;
            WordSnapshot before = state.Snapshot ?? WordSnapshot.Capture(state.Document);
            Authorization authorization = direct
                ? state.Policy.ValidateDirect(id, state.Id, before.Revision, before.MainText, before.Revision)
                : state.Policy.ValidateAcceptance(id, state.Id, before.Revision, before.MainText, before.Revision);
            Proposal proposal = authorization.Proposal;
            string replacement = state.Policy.ResolveReplacement(proposal, before.MainText);
            List<PassageMark> protectedUnits = state.Policy.ProtectedUnits(proposal, before.MainText);
            List<ProtectedPlacement> placements = state.Policy.ProtectedPlacements(proposal, before.MainText);
            RequireParagraphUnit(state, before, proposal);
            before.RequireProse(state.Document, proposal.From, proposal.To, OwnedBookmarks(state), protectedUnits);
            RequireNativeReviewRange(state, before, proposal, tracked);
            if (direct && !tracked && GetForegroundWindow() == new IntPtr((int)application.ActiveWindow.Hwnd))
                throw new InvalidOperationException("Finish typing and move focus from the Word document before a direct AI edit.");
            if ((int)application.UndoRecord.CustomRecordLevel != 0)
                throw new InvalidOperationException("Finish the current Word editing operation before applying a suggestion.");
            dynamic range = before.Range(state.Document, proposal.From, proposal.To);
            if (protectedUnits.Count == 0) WordSnapshot.RequireUniformFormatting(range);
            dynamic originalFont = range.Font.Duplicate;
            if (protectedUnits.Count == 0 && ((string)range.Font.Name == "" || (int)range.Font.Bold == 9999999 || (int)range.Font.Italic == 9999999 ||
                (float)range.Font.Size == 9999999 || (int)range.Font.Underline == 9999999 || (int)range.Font.Color == 9999999 ||
                (int)range.Font.Subscript == 9999999 || (int)range.Font.Superscript == 9999999 || (int)range.Font.Hidden == 9999999 ||
                (int)range.Font.SmallCaps == 9999999 || (int)range.LanguageID == 9999999))
                throw new InvalidOperationException("Select a passage with consistent character formatting so Word can preserve it.");
            if (replacement.Any(c => c < 32 && c != '\t'))
                throw new InvalidOperationException("Replacement wording cannot introduce structural paragraph or field characters.");
            int retainedSuffix = CommentAnchorSuffix(state, before, proposal, replacement, direct, tracked, protectedUnits);
            AcceptedProtectionPlan acceptedProtection = direct && !tracked ? PlanAcceptedProtection(state, proposal.From, replacement, placements) : null;
            dynamic originalParagraph = before.Range(state.Document, proposal.From, proposal.To).Paragraphs.First.Range;
            string originalParagraphQuote = PolicyEngine.Slice(before.MainText, before.VisibleOffset((int)originalParagraph.Start), before.VisibleOffset((int)originalParagraph.End));
            WordTrackedProtection trackedProtection = tracked ? new WordTrackedProtection(state, before, proposal) : null;
            diagnostic.Phase = WordNativeOutcome.Stage.protected_prepare;
            using (var preserved = new WordProtectedContent(application, state.Document, proposal, protectedUnits, placements))
            {
            WordSnapshot confirmed = WordSnapshot.Capture(state.Document);
            if (confirmed.Revision != before.Revision)
                throw new InvalidOperationException("The native document changed before replacement.");
            PolicyState beforePolicy = state.Policy.ExportState();
            state.NativeEdit.Capture(state.Document, beforePolicy, confirmed);
            string beforeNative = state.NativeEdit.NativeIdentity(state.Document, confirmed);
            diagnostic.Phase = WordNativeOutcome.Stage.begin_application;
            state.Policy.BeginApplication(id);
            Persist(state, false);
            bool ended = false;
            bool originalTracking = (bool)state.Document.TrackRevisions;
            bool originalTrackFormatting = (bool)state.Document.TrackFormatting;
            bool settingsRestored = false;
            WordSnapshot expectedReview = before;
            diagnostic.Phase = WordNativeOutcome.Stage.begin_undo;
            application.UndoRecord.StartCustomRecord(AppBranding.Name + " selective edit");
            try
            {
                nativeEffects = true;
                diagnostic.Phase = WordNativeOutcome.Stage.tracking_settings;
                state.Document.TrackFormatting = false;
                state.Document.TrackRevisions = tracked;
                diagnostic.Phase = WordNativeOutcome.Stage.original_anchors;
                if (trackedProtection != null) trackedProtection.AddOriginalBookmarks(state.Document, before);
                diagnostic.Phase = WordNativeOutcome.Stage.replace_text;
                before.Range(state.Document, proposal.From, proposal.To - retainedSuffix).Text = replacement.Substring(0, replacement.Length - retainedSuffix);
                diagnostic.Phase = WordNativeOutcome.Stage.restore_formatting;
                if (protectedUnits.Count == 0 && replacement.Length > retainedSuffix)
                    foreach (int[] span in new WordTextMap(state.Document).NativeSegments(proposal.From, proposal.From + replacement.Length - retainedSuffix))
                        state.Document.Range(span[0], span[1]).Font = originalFont;
                diagnostic.Phase = WordNativeOutcome.Stage.restore_protected;
                preserved.Restore(state.Document, proposal.From, replacement.Length);
                WordSnapshot reconstructed = WordSnapshot.CaptureTextPositions(state.Document);
                diagnostic.Phase = WordNativeOutcome.Stage.remap_markers;
                int delta = replacement.Length - (proposal.To - proposal.From);
                foreach (PassageMark mark in state.Policy.Marks.ToArray())
                {
                    if (mark.Stale) continue;
                    int from = mark.From, to = mark.To;
                    if (to <= proposal.From) continue;
                    if (from >= proposal.To) { from += delta; to += delta; }
                    else
                    {
                        ProtectedPlacement placement = placements.FirstOrDefault(candidate => candidate.Id == mark.Id);
                        PassageMark owner = protectedUnits.FirstOrDefault(unit => mark.From >= unit.From && mark.To <= unit.To);
                        if (placement != null) { from = placement.From; to = placement.To; }
                        else if (owner != null)
                        {
                            ProtectedPlacement moved = placements.First(candidate => candidate.Id == owner.Id);
                            from = moved.From + mark.From - owner.From; to = moved.From + mark.To - owner.From;
                        }
                        else
                        {
                            bool enclosing = from <= proposal.From && to >= proposal.To;
                            from = Math.Min(from, proposal.From); to = to <= proposal.To ? proposal.From + replacement.Length : to + delta;
                            mark.Stale = !enclosing;
                            mark.RequiresRemarking = !enclosing;
                        }
                    }
                    if (to == from)
                    {
                        if (tracked) trackedProtection.RetainDeletedMark(mark);
                        else state.Policy.RemoveMark(mark.Id);
                        if ((bool)state.Document.Bookmarks.Exists(mark.Id)) state.Document.Bookmarks.Item(mark.Id).Delete();
                        continue;
                    }
                    mark.NativeBookmark = null;
                    mark.From = from; mark.To = to;
                    mark.Quote = PolicyEngine.Slice(reconstructed.MainText, from, to);
                    state.Document.Bookmarks.Add(mark.Id, reconstructed.Range(state.Document, from, to));
                }
                diagnostic.Phase = WordNativeOutcome.Stage.repaint_markers;
                RepaintMarks(state);
                diagnostic.Phase = WordNativeOutcome.Stage.record_review;
                if (tracked) RecordNativeReview(state, reconstructed, proposal, replacement.Length);
                if (trackedProtection != null) trackedProtection.PrepareWitnesses(state, proposal.Id);
                Proposal acceptanceReceipt = state.Policy.Proposals.First(candidate => candidate.Id == proposal.Id);
                diagnostic.Phase = WordNativeOutcome.Stage.record_acceptance;
                if (tracked) RecordAcceptanceCandidate(state, reconstructed, acceptanceReceipt, replacement, originalParagraphQuote);
                else if (direct) ApplyAcceptedProtection(state, reconstructed, acceptanceReceipt, acceptedProtection, replacement);
                if (tracked || direct)
                {
                    diagnostic.Phase = WordNativeOutcome.Stage.verify_discussion;
                    WordSnapshot unchangedReview = WordSnapshot.CaptureReviewState(state.Document);
                    if (unchangedReview.ContextStories != before.ContextStories || unchangedReview.Review.FeedbackIdentity() != before.Review.FeedbackIdentity())
                        throw new InvalidOperationException("Word's replacement did not preserve the native review discussion.");
                    diagnostic.Phase = WordNativeOutcome.Stage.add_comment;
                    AddSuggestionComment(state, reconstructed, proposal, replacement.Length);
                }
                diagnostic.Phase = WordNativeOutcome.Stage.add_undo_guard;
                state.NativeEdit.AddGuard(state.Document);
                diagnostic.Phase = WordNativeOutcome.Stage.persist_native;
                Persist(state, true);
                diagnostic.Phase = WordNativeOutcome.Stage.review_visibility;
                if (tracked) state.Document.ShowRevisions = true;
                if (tracked || direct)
                {
                    state.Document.ActiveWindow.View.ShowComments = true;
                    state.Document.ActiveWindow.View.ShowRevisionsAndComments = true;
                }
                diagnostic.Phase = WordNativeOutcome.Stage.restore_settings;
                state.Document.TrackFormatting = originalTrackFormatting;
                state.Document.TrackRevisions = tracked ? true : originalTracking;
                settingsRestored = true;
                diagnostic.Phase = WordNativeOutcome.Stage.capture_review;
                if (tracked || direct) expectedReview = WordSnapshot.CaptureReviewState(state.Document);
                if ((int)application.UndoRecord.CustomRecordLevel != 1)
                    throw new InvalidOperationException("Word ended the selective edit's Undo record before its native transaction completed.");
                // WordOpenXML export terminates native Undo recording; verify after all grouped mutations.
                diagnostic.Phase = WordNativeOutcome.Stage.end_undo;
                application.UndoRecord.EndCustomRecord();
                ended = true;
                diagnostic.Phase = WordNativeOutcome.Stage.verify_protected;
                preserved.Verify(state.Document);
                diagnostic.Phase = WordNativeOutcome.Stage.verify_snapshot;
                string expected = before.MainText.Substring(0, proposal.From) + replacement + before.MainText.Substring(proposal.To);
                WordSnapshot after = WordSnapshot.Capture(state.Document);
                if (after.MainText != expected)
                    throw new InvalidOperationException("Word's replacement did not preserve the surrounding manuscript.");
                if (after.ContextStories != expectedReview.ContextStories || after.Review.FeedbackIdentity() != expectedReview.Review.FeedbackIdentity())
                    throw new InvalidOperationException("Word's replacement did not preserve the native review discussion.");
                diagnostic.Phase = WordNativeOutcome.Stage.commit_receipt;
                if (tracked)
                {
                    Proposal verified = state.Policy.Proposals.First(candidate => candidate.Id == proposal.Id);
                    state.Policy.RecordTracked(id, verified.NativeReviewBookmark, verified.NativeReviewHash, verified.NativeReviewQuote);
                }
                else state.Policy.RecordApplied(id);
                committed = true;
                reviewCommitted = true;
                diagnostic.Phase = WordNativeOutcome.Stage.persist_commit;
                Persist(state, true);
                diagnostic.Phase = WordNativeOutcome.Stage.capture_commit;
                state.NativeEdit.Capture(state.Document, state.Policy.ExportState(), after);
                state.Snapshot = after;
            }
            catch (Exception error)
            {
                diagnostic.Capture(diagnostic.Phase, error, committed);
                if (committed) throw;
                bool restored = diagnostic.AttemptRollback(delegate {
                    if (!ended) { application.UndoRecord.EndCustomRecord(); ended = true; }
                    settingsRestored = false;
                    // A failed COM call may already have changed bookmarks or text.
                    bool same = state.NativeEdit.NativeIdentity(state.Document) == beforeNative;
                    if (!same) same = (bool)state.Document.Undo(1) && state.NativeEdit.NativeIdentity(state.Document) == beforeNative;
                    return same;
                });
                if (!restored)
                    throw new InvalidOperationException("The selective edit outcome is uncertain. Inspect Word and its Undo history before continuing.");
                nativeEffects = false;
                state.Policy.RecordOutcomeUncertain(id);
                beforePolicy.Proposals = state.Policy.ExportState().Proposals;
                beforePolicy.HighlightReceipts = state.Policy.ExportState().HighlightReceipts;
                state.Policy = CreatePolicy(beforePolicy);
                state.Policy.Revoke(); state.ReadRevision = null;
                Persist(state, true);
                state.NativeEdit.Clear(); state.NativeEdit.Capture(state.Document, state.Policy.ExportState());
                throw;
            }
            finally
            {
                try
                {
                    diagnostic.Phase = WordNativeOutcome.Stage.settings_cleanup;
                    if (!settingsRestored)
                    {
                        state.Document.TrackFormatting = originalTrackFormatting;
                        state.Document.TrackRevisions = committed && tracked ? true : originalTracking;
                    }
                    diagnostic.Phase = WordNativeOutcome.Stage.undo_cleanup;
                    if (!ended) application.UndoRecord.EndCustomRecord();
                }
                catch (Exception error) { diagnostic.Capture(diagnostic.Phase, error, committed); throw; }
            }
            diagnostic.Phase = WordNativeOutcome.Stage.protected_cleanup;
            }
            diagnostic.Phase = WordNativeOutcome.Stage.window_cleanup;
            }
            }
            catch (Exception error)
            {
                diagnostic.Capture(diagnostic.Phase, error, committed);
                if (committed) throw WordNativeOutcome.UnknownException(diagnostic);
                if (nativeEffects) WordNativeOutcome.Unknown(
                    delegate { if (state.Policy.Proposals.First(proposal => proposal.Id == id).Status != "outcome_uncertain") state.Policy.RecordOutcomeUncertain(id); },
                    state.Policy.Revoke, delegate { Persist(state, false); }, diagnostic);
                throw;
            }
        }
    }
}

