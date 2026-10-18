using System;
using System.Collections.Generic;
using System.IO;
using UnderTree.WordAddin;

internal static class PolicyConcurrencyTests
{
    private const string Document = "concurrent-document", Revision = "revision", Source = "Alpha beta.", Snapshot = "all-native-stories";
    private static string contract;
    private static int checks;

    public static int Main(string[] args)
    {
        try
        {
            contract = File.ReadAllText(args[0]);
            IndependentReadWitnesses();
            ReplayRequiresOriginalContext();
            UncertainReceiptLookup();
            RetiredSubscopesAndCapacity();
            ScopedHighlighting();
            Console.WriteLine("PASS " + checks + " scoped native-policy concurrency assertions.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static PolicyEngine Shared()
    {
        var policy = new PolicyEngine(contract);
        policy.AddMark(new PassageMark { Id = "Wrai_scope", RuleId = "snappier", From = 0, To = Source.Length, Quote = Source, Prompt = "", Stale = false }, Source);
        policy.Share(Document);
        return policy;
    }

    private static Proposal Request(string id)
    { return new Proposal { Id = id, DocumentId = Document, Revision = Revision, From = 0, To = Source.Length, Quote = Source, Replacement = "A clear point.", Rationale = "Clarify the wording." }; }

    private static void IndependentReadWitnesses()
    {
        PolicyEngine policy = Shared();
        string a = policy.ReadContext(Document, Revision, Source, Snapshot, "transport-a:chat"),
            b = policy.ReadContext(Document, Revision, Source, Snapshot, "transport-b:chat");
        Check(a != b && policy.HasReadWitness(a, "transport-a:chat") && policy.HasReadWitness(b, "transport-b:chat"),
            "Independent reads retain separate valid witnesses at one native revision.");
        Denied(() => policy.Submit(Request("cross-client"), a, Source, Snapshot, "transport-b:chat"),
            "A complete-read witness cannot cross client scopes.");
        Check(policy.Proposals.Count == 0, "Cross-client witness denial records no native proposal.");
        Authorization result = policy.Submit(Request("first"), a, Source, Snapshot, "transport-a:chat");
        Check(result.Proposal.Status == "authorized" && result.Mode == "direct", "A later independent reader does not invalidate the first client's authorized direct proposal.");
        policy.BeginApplication("first");
        policy.RecordApplied("first");
        Denied(() => policy.Submit(Request("second"), b, Source, Snapshot, "transport-b:chat"),
            "A committed native direct edit invalidates every earlier witness even when the supplied source revision stays unchanged.");
        Check(policy.Submit(Request("first"), a, Source, Snapshot, "transport-a:chat").AlreadyRecorded && policy.Proposals.Count == 1,
            "The original exact request still replays its durable receipt after all witnesses were invalidated.");
    }

    private static void RetiredSubscopesAndCapacity()
    {
        PolicyEngine policy = Shared();
        string a = policy.ReadContext(Document, Revision, Source, Snapshot, "transport-a:one"),
            a2 = policy.ReadContext(Document, Revision, Source, Snapshot, "transport-a:two"),
            b = policy.ReadContext(Document, Revision, Source, Snapshot, "transport-b:one");
        policy.RevokeClientContexts("transport-a:");
        Check(!policy.HasReadWitness(a, "transport-a:one") && !policy.HasReadWitness(a2, "transport-a:two") && policy.HasReadWitness(b, "transport-b:one"),
            "Retiring one transport session clears all its conversation subscopes and preserves another client's witness.");
        Denied(() => policy.Submit(Request("retired"), a, Source, Snapshot, "transport-a:one"), "A retired subscope cannot submit its old witness.");
        Authorization survivor = policy.Submit(Request("survivor"), b, Source, Snapshot, "transport-b:one");
        Check(survivor.Proposal.Status == "authorized" && survivor.Mode == "direct", "An independent client's context survives another session's retirement.");
        policy = Shared();
        string first = null;
        for (int index = 0; index < 64; index++)
        {
            string token = policy.ReadContext(Document, Revision, Source, Snapshot, "scope:" + index);
            if (index == 0) first = token;
        }
        Denied(() => policy.ReadContext(Document, Revision, Source, Snapshot, "overflow"), "Witness capacity rejects a new client without evicting a live one.");
        Check(policy.HasReadWitness(first, "scope:0"), "Witness capacity preserves the first active reader.");
        policy.RevokeClientContexts("scope:");
        string fresh = policy.ReadContext(Document, Revision, Source, Snapshot, "fresh");
        Check(policy.HasReadWitness(fresh, "fresh"), "Retiring old subscopes makes bounded witness capacity reusable.");
        long generation = policy.ContextGeneration;
        policy.InvalidateContext();
        Check(policy.ContextGeneration > generation && !policy.HasReadWitness(fresh, "fresh"), "Policy invalidation advances the generation used by paginated reads.");
    }

    private static void ReplayRequiresOriginalContext()
    {
        PolicyEngine policy = Shared();
        string token = policy.ReadContext(Document, Revision, Source, Snapshot, "replay-owner");
        Proposal request = Request("context-bound-replay");
        policy.Submit(request, token, Source, Snapshot, "replay-owner");
        // Preserve a genuine historical tracked receipt rather than treating new prompt work as suggestions.
        policy.Proposals[0].Status = "pending";
        policy.BeginApplication(request.Id);
        policy.RecordTracked(request.Id, "WraiReview_context", PolicyEngine.Hash("native-revision"), "A clear point.");
        Check(policy.Submit(request, token, Source, Snapshot, "replay-owner").AlreadyRecorded, "An exact original-context retry reuses its recorded native receipt.");
        Denied(() => policy.Submit(request, "changed-context-only", Source, Snapshot, "replay-owner"),
            "A consumed request cannot replay with a changed context identity.");
        PolicyEngine reopened = new PolicyEngine(contract, policy.ExportState());
        reopened.Share(Document);
        Check(reopened.Submit(request, token, Source, Snapshot, "replay-owner").AlreadyRecorded,
            "The original request context remains bound after receipt serialization and reopening.");
        Denied(() => reopened.Submit(request, "changed-after-reopen", Source, Snapshot, "replay-owner"),
            "A reopened receipt cannot silently accept a changed original context.");
        PolicyState legacy = policy.ExportState();
        legacy.Proposals[0].RequestContextHash = null;
        var preserved = new PolicyEngine(contract, legacy);
        preserved.Share(Document);
        Check(preserved.Proposals[0].Status == "tracked", "Legacy receipts remain preserved without inventing their original request context.");
        Denied(() => preserved.Submit(request, token, Source, Snapshot, "replay-owner"), "A legacy receipt without its original context hash fails closed for replay.");
        legacy.Proposals[0].RequestContextHash = "malformed";
        Denied(() => new PolicyEngine(contract, legacy), "A malformed persisted request-context hash is refused.");
    }

    private static void ScopedHighlighting()
    {
        PolicyEngine policy = Shared();
        string a = policy.ReadContext(Document, Revision, Source, Snapshot, "a"), b = policy.ReadContext(Document, Revision, Source, Snapshot, "b");
        var marking = new HighlightRequest { DocumentId = Document, RequestId = "neutral-scope", Revision = Revision, ContextId = a,
            ColorId = "snappier", Ranges = new List<HighlightRange> { new HighlightRange { From = 0, To = 5, Quote = "Alpha" } } };
        Denied(() => policy.Highlight(marking, Source, Snapshot, "b"), "A color-only call cannot use another client's reading witness.");
        Check(policy.HighlightReceipts.Count == 0 && policy.Marks.Count == 1, "Rejected cross-client coloring creates no mark or replay receipt.");
        HighlightReceipt receipt = policy.Highlight(marking, Source, Snapshot, "a");
        Check(receipt.MarkIds.Count == 1 && !policy.HasReadWitness(a, "a") && !policy.HasReadWitness(b, "b"), "Authorized native highlighting invalidates both clients' earlier context.");
    }

    private static void UncertainReceiptLookup()
    {
        PolicyEngine policy = Shared();
        string token = policy.ReadContext(Document, Revision, Source, Snapshot, "uncertain-owner");
        Proposal request = Request("uncertain-original");
        policy.Submit(request, token, Source, Snapshot, "uncertain-owner");
        policy.BeginApplication(request.Id); policy.RecordOutcomeUncertain(request.Id); policy.Revoke();
        Check(policy.UncertainReplay(request, token) && !policy.IsShared && policy.Proposals[0].Status == "outcome_uncertain",
            "An exact uncertain proposal receipt remains inspectable after revocation without granting edit authority.");
        Denied(() => policy.UncertainReplay(request, "changed-context"), "A changed context cannot inherit a proposal's uncertain receipt.");
        Proposal changed = Request(request.Id); changed.Replacement = "Different wording.";
        Denied(() => policy.UncertainReplay(changed, token), "Changed wording cannot inherit a proposal's uncertain receipt.");
        var reopened = new PolicyEngine(contract, policy.ExportState());
        Check(reopened.UncertainReplay(request, token) && !reopened.IsShared, "Uncertainty and the exact original request survive persisted policy reopening.");
        Denied(() => reopened.Submit(Request("new-after-unknown"), token, Source, Snapshot, "uncertain-owner"), "Uncertainty receipt lookup does not reshare the document or authorize a new identity.");

        policy = Shared(); token = policy.ReadContext(Document, Revision, Source, Snapshot, "uncertain-marker");
        var marking = new HighlightRequest { DocumentId = Document, RequestId = "uncertain-mark", Revision = Revision, ContextId = token,
            ColorId = "snappier", Ranges = new List<HighlightRange> { new HighlightRange { From = 0, To = 5, Quote = "Alpha" } } };
        policy.Highlight(marking, Source, Snapshot, "uncertain-marker"); policy.ReserveHighlight(marking.RequestId); policy.FailHighlight(marking.RequestId);
        Check(policy.UncertainReplay(marking) && !policy.IsShared, "An exact uncertain marking receipt remains inspectable without resuming sharing.");
        marking.ContextId = "changed-context";
        Denied(() => policy.UncertainReplay(marking), "A changed context cannot inherit a marking's uncertain receipt.");
        marking.ContextId = token; marking.ColorId = "protect";
        Denied(() => policy.UncertainReplay(marking), "A changed color cannot inherit a marking's uncertain receipt.");
    }

    private static void Denied(Action action, string message)
    {
        try { action(); }
        catch (InvalidOperationException) { checks++; return; }
        throw new Exception(message);
    }

    private static void Check(bool passed, string message)
    { if (!passed) throw new Exception(message); checks++; }
}
