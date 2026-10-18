using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using UnderTree.WordAddin;

internal static class PolicyRecolorTests
{
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 16777216 };
    private const string Text = "Alpha beta gamma delta.";
    private const string Document = "recolor_document", Revision = "recolor_revision", Snapshot = "native_recolor_snapshot";
    private static string contract;
    private static int assertions;

    public static int Main(string[] args)
    {
        try
        {
            contract = File.ReadAllText(args[0]);
            GreenToRedAndReplay();
            OverlappingCurrentSelections();
            NeutralTargetPreservesInstructions();
            MixedExistingActions();
            SameColorAndOrdinaryHighlighting();
            InvalidSelectionIsAtomic();
            UnselectedOverlapRemainsProtected();
            UnresolvedNativeReviewIsRefused();
            ReplacementAtMarkerLimit();
            InvalidReplacementReceiptsAreRefused();
            Console.WriteLine("Word recolor policy: " + assertions + " assertions passed.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static PolicyEngine Engine()
    {
        var engine = new PolicyEngine(contract); engine.Rules.Clear();
        foreach (var rule in new[] {
            Rule("green", "#74B9A4", "protect"), Rule("red", "#E57575", "reject"),
            Rule("blue", "#AACCEE", "prompt", "direct"), Rule("yellow", "#E5BE50", "prompt"),
            Rule("discard", "#F4A7BC", "discard"), Rule("visual", "#BBBBBB", "neutral") }) engine.AddOrUpdateRule(rule);
        engine.Share(Document); return engine;
    }

    private static ColorRule Rule(string id, string color, string action, string mode = "suggestions")
    { return new ColorRule { Id = id, Label = id, Color = color, Action = action, AccessMode = mode, Prompt = action == "prompt" ? "Retain the meaning." : "" }; }

    private static PassageMark Add(PolicyEngine engine, string id, string rule, string quote, string prompt = "")
    {
        int from = Text.IndexOf(quote, StringComparison.Ordinal);
        var mark = new PassageMark { Id = id, RuleId = rule, From = from, To = from + quote.Length, Quote = quote, Prompt = prompt };
        engine.AddMark(mark, Text); return mark;
    }

    private static HighlightRequest Recolor(PolicyEngine engine, string id, string colorId, params PassageMark[] selected)
    {
        return Json.Deserialize<HighlightRequest>(Json.Serialize(new { DocumentId = Document, Revision = Revision, RequestId = id,
            ContextId = engine.ReadContext(Document, Revision, Text, Snapshot), ColorId = colorId,
            RecolorMarkIds = selected.Select(mark => mark.Id).ToArray(),
            Ranges = selected.Select(mark => new { From = mark.From, To = mark.To, Quote = mark.Quote }).ToArray() }));
    }

    private static HighlightRequest Change(HighlightRequest source, string key, object value)
    {
        var values = Json.Deserialize<Dictionary<string, object>>(Json.Serialize(source)); values[key] = value;
        return Json.Deserialize<HighlightRequest>(Json.Serialize(values));
    }

    private static string[] Replaced(HighlightReceipt receipt)
    {
        var values = Json.Deserialize<Dictionary<string, object>>(Json.Serialize(receipt));
        Equal(true, values.ContainsKey("ReplacedMarkIds") && values["ReplacedMarkIds"] != null, "A recolor receipt identifies every retired marker.");
        return ((System.Collections.IEnumerable)values["ReplacedMarkIds"]).Cast<string>().ToArray();
    }

    private static void GreenToRedAndReplay()
    {
        var engine = Engine(); PassageMark source = Add(engine, "Wrai_original_beta", "green", "beta");
        PassageMark neighbor = Add(engine, "Wrai_neighbor_delta", "green", "delta");
        string definitions = Json.Serialize(engine.Rules), neighboring = Json.Serialize(neighbor);
        var request = Recolor(engine, "green_to_red", "red", source);
        HighlightReceipt receipt = engine.Highlight(request, Text, Snapshot);
        Equal(true, Replaced(receipt).SequenceEqual(new[] { source.Id }), "Green-to-red returns the selected old identity in order.");
        Equal("red", receipt.ColorId, "The explicitly selected existing red rule is retained.");
        Equal(2, engine.Marks.Count, "Recolor replaces one marker without increasing the marker count.");
        Equal(false, engine.Marks.Any(mark => mark.Id == source.Id), "The selected original green identity is retired.");
        PassageMark replacement = engine.Marks.Single(mark => mark.Id == receipt.MarkIds.Single());
        Equal(true, replacement.Id != source.Id && replacement.RuleId == "red" && replacement.From == source.From && replacement.To == source.To &&
            replacement.Quote == source.Quote && replacement.Prompt == source.Prompt, "The new red identity preserves exact wording, range and passage instructions.");
        Equal(definitions, Json.Serialize(engine.Rules), "Recoloring changes no existing local rule definition.");
        Equal(neighboring, Json.Serialize(engine.Marks.Single(mark => mark.Id == neighbor.Id)), "Neighboring protection remains exact.");
        Equal(0, engine.Proposals.Count, "A recolor request creates no wording change or acceptance proposal.");
        string unchanged = Json.Serialize(engine.ExportState());
        Equal(receipt.MarkIds.Single(), engine.Highlight(request, Text, Snapshot).MarkIds.Single(), "A lost recolor reply replays the same new identity.");
        Equal(unchanged, Json.Serialize(engine.ExportState()), "Recolor replay has no further effect.");
        Denied(engine, Change(request, "ColorId", "blue"), Text, "A consumed recolor request cannot change its destination color.");
        engine.Revoke(); engine.Share(Document);
        Denied(engine, request, Text, "Revocation prevents recolor replay or restoration of retired identities.");
        var reopened = new PolicyEngine(contract, Json.Deserialize<PolicyState>(Json.Serialize(engine.ExportState())));
        Equal("red", reopened.Marks.Single(mark => mark.Id == replacement.Id).RuleId, "Serialized recolor state retains the new color binding.");
        Equal(definitions, Json.Serialize(reopened.Rules), "Reopening retains the unchanged local definitions.");
        reopened.Share(Document);
        Denied(reopened, request, Text, "A reopened process cannot replay the prior native operation.");
    }

    private static void OverlappingCurrentSelections()
    {
        foreach (var scenario in new[] {
            new { Name = "identical", Quotes = new[] { "beta", "beta" } },
            new { Name = "partial", Quotes = new[] { "beta gamma", "gamma delta" } },
            new { Name = "nested", Quotes = new[] { "beta gamma", "beta" } } })
        {
            var engine = Engine();
            for (int index = 0; index < scenario.Quotes.Length; index++)
            {
                string quote = scenario.Quotes[index]; int from = Text.IndexOf(quote, StringComparison.Ordinal);
                engine.Highlight(new HighlightRequest { DocumentId = Document, Revision = Revision, ColorId = "green",
                    RequestId = scenario.Name + "_ordinary_" + index, ContextId = engine.ReadContext(Document, Revision, Text, Snapshot),
                    Ranges = new List<HighlightRange> { new HighlightRange { From = from, To = from + quote.Length, Quote = quote } } }, Text, Snapshot);
            }
            PassageMark[] selected = engine.Marks.OrderBy(mark => mark.From).ThenBy(mark => mark.To).ToArray();
            Equal(2, selected.Length, "Separate ordinary same-RGB requests create two legitimate current " + scenario.Name + " overlaps.");
            PassageMark neighbor = Add(engine, "Wrai_overlap_neighbor", "green", "Alpha");
            string definitions = Json.Serialize(engine.Rules), neighboring = Json.Serialize(neighbor);
            Denied(engine, Recolor(engine, scenario.Name + "_selected_subset", "red", selected[0]), Text,
                "A selected subset cannot recolor an unselected overlapping protected identity.");
            var request = Recolor(engine, scenario.Name + "_all_selected", "red", selected);
            Denied(engine, Change(Change(request, "RecolorMarkIds", null), "ColorId", "green"), Text,
                "An ordinary marking batch still requires disjoint ranges.", "nonoverlapping");
            Denied(engine, Recolor(engine, scenario.Name + "_unsorted", "red", selected.Concat(new[] { neighbor }).ToArray()), Text,
                "Even exact paired recolor IDs must retain nondecreasing range starts.", "ordered");
            request.ContextId = engine.ReadContext(Document, Revision, Text, Snapshot);
            Denied(engine, Change(request, "RecolorMarkIds", new[] { selected[0].Id, selected[0].Id }), Text,
                "Equal or overlapping geometry never permits repeated source IDs.", "unique");
            Denied(engine, Change(request, "Ranges", new[] { new HighlightRange { From = 0, To = 5, Quote = "Alpha" }, request.Ranges[1] }), Text,
                "Every overlapping selection still requires the original ID's exact range and quote.", "paired current marker");
            HighlightReceipt receipt = engine.Highlight(request, Text, Snapshot);
            Equal(true, Replaced(receipt).SequenceEqual(selected.Select(mark => mark.Id)), "All selected overlapping identities are retired in exact request order.");
            Equal(true, selected.All(old => !engine.Marks.Any(mark => mark.Id == old.Id)) && receipt.MarkIds.Distinct().Count() == 2 && engine.Marks.Count == 3,
                "An atomic overlapping recolor replaces only the selected identities with distinct new IDs.");
            for (int index = 0; index < selected.Length; index++)
            {
                PassageMark replacement = engine.Marks.Single(mark => mark.Id == receipt.MarkIds[index]); PassageMark source = selected[index];
                Equal(true, replacement.RuleId == "red" && replacement.From == source.From && replacement.To == source.To &&
                    replacement.Quote == source.Quote && replacement.Prompt == source.Prompt,
                    "Each new overlapping marker preserves its source's exact geometry, quote and passage prompt.");
            }
            Equal(definitions, Json.Serialize(engine.Rules), "Overlapping recoloring leaves all local definitions unchanged.");
            Equal(neighboring, Json.Serialize(engine.Marks.Single(mark => mark.Id == neighbor.Id)), "Overlapping recoloring leaves neighboring protection exact.");
            string unchanged = Json.Serialize(engine.ExportState());
            Equal(true, engine.Highlight(request, Text, Snapshot).MarkIds.SequenceEqual(receipt.MarkIds), "An overlapping recolor replay returns the exact prior new IDs.");
            Equal(unchanged, Json.Serialize(engine.ExportState()), "Overlapping replay changes no markers, definitions or receipts.");
            PolicyState saved = engine.ExportState();
            Equal(unchanged, Json.Serialize(new PolicyEngine(contract, saved).ExportState()), "Overlapping replacement requests and receipts survive metadata validation and reopening.");
            saved.HighlightReceipts.Last().ReplacedMarkIds.Reverse();
            Refused(() => new PolicyEngine(contract, saved), "Overlapping receipts cannot reorder the replaced source identities.");
        }
    }

    private static void NeutralTargetPreservesInstructions()
    {
        var engine = Engine(); PassageMark source = Add(engine, "Wrai_neutral_source", "green", "beta");
        string definitions = Json.Serialize(engine.Rules);
        var request = Change(Recolor(engine, "protect_to_fresh_visual", null, source), "Color", "#123ABC");
        HighlightReceipt receipt = engine.Highlight(request, Text, Snapshot);
        ColorRule visual = engine.Rules.Single(rule => rule.Id == receipt.ColorId);
        Equal("neutral", visual.Action, "An unused requested recolor RGB creates a neutral identity.");
        Equal(true, engine.Marks.Single().RuleId == visual.Id && engine.Marks.Single().Quote == "beta" && engine.Marks.Single().Prompt == "",
            "Explicitly recoloring an empty-prompt protected passage to a fresh color changes only its association.");
        Equal(definitions, Json.Serialize(engine.Rules.Where(rule => rule.Id != visual.Id).ToArray()), "Allocating a neutral destination never alters prior rule definitions.");
        string token = engine.ReadContext(Document, Revision, Text, Snapshot);
        Refused(() => engine.Submit(new Proposal { Id = "neutral_no_wording", DocumentId = Document, Revision = Revision, From = 6, To = 10,
            Quote = "beta", Replacement = "changed", Rationale = "" }, token, Text, Snapshot), "A neutral recolor creates no wording permission.");
        engine = Engine(); source = Add(engine, "Wrai_prompted_source", "blue", "beta", "Keep the specialist term.");
        request = Recolor(engine, "do_not_drop_prompt", "visual", source);
        Denied(engine, request, Text, "Recoloring cannot silently discard a nonempty passage prompt.", "passage prompt");
    }

    private static void MixedExistingActions()
    {
        var engine = Engine(); var selected = new[] { Add(engine, "Wrai_alpha", "green", "Alpha"), Add(engine, "Wrai_beta", "red", "beta"),
            Add(engine, "Wrai_gamma", "blue", "gamma", "Preserve this exact term."), Add(engine, "Wrai_delta", "discard", "delta") };
        PassageMark neighbor = Add(engine, "Wrai_period", "green", ".");
        string definitions = Json.Serialize(engine.Rules), retained = Json.Serialize(neighbor);
        HighlightReceipt receipt = engine.Highlight(Recolor(engine, "mixed_sources", "yellow", selected), Text, Snapshot);
        Equal(true, Replaced(receipt).SequenceEqual(selected.Select(mark => mark.Id)), "A mixed-action batch returns exactly the selected old IDs in range order.");
        Equal(5, engine.Marks.Count, "A mixed batch replaces only its four selected passages.");
        for (int index = 0; index < selected.Length; index++)
        {
            PassageMark fresh = engine.Marks.Single(mark => mark.Id == receipt.MarkIds[index]); PassageMark old = selected[index];
            Equal(true, fresh.Id != old.Id && fresh.RuleId == "yellow" && fresh.From == old.From && fresh.To == old.To && fresh.Quote == old.Quote && fresh.Prompt == old.Prompt,
                "Every mixed-source replacement retains its exact text, geometry and passage prompt.");
        }
        Equal(definitions, Json.Serialize(engine.Rules), "Mixed-action recoloring never redefines a function or access mode.");
        Equal(retained, Json.Serialize(engine.Marks.Single(mark => mark.Id == neighbor.Id)), "A neighboring indivisible protected unit is untouched.");
    }

    private static void InvalidSelectionIsAtomic()
    {
        var engine = Engine(); PassageMark beta = Add(engine, "Wrai_beta", "green", "beta"), delta = Add(engine, "Wrai_delta", "green", "delta");
        var request = Recolor(engine, "invalid_recolor", "red", beta, delta);
        Denied(engine, Change(request, "DocumentId", "another_document"), Text, "A recolor cannot use another document's marker authority.");
        Denied(engine, Change(request, "ContextId", "forged_context"), Text, "A recolor requires the current complete-read witness.");
        Denied(engine, Change(request, "Revision", "another_revision"), Text, "A recolor cannot reuse a different revision's witness.");
        foreach (object ids in new object[] { new string[0], new[] { beta.Id }, new[] { beta.Id, beta.Id }, new[] { delta.Id, beta.Id },
            new[] { beta.Id, "missing_id" }, new[] { beta.Id, "invalid#id" }, new[] { beta.Id, (string)null } })
            Denied(engine, Change(request, "RecolorMarkIds", ids), Text, "Invalid, duplicate, missing or mismatched recolor IDs leave the entire batch unchanged.");
        var wrongRange = Change(request, "Ranges", new[] { new { From = 0, To = 5, Quote = "Alpha" }, new { From = 17, To = 22, Quote = "delta" } });
        Denied(engine, wrongRange, Text, "A source identity cannot be moved to a different exact quote by recoloring.");
        var wrongQuote = Change(request, "Ranges", new[] { new { From = 6, To = 10, Quote = "beta" }, new { From = 17, To = 22, Quote = "wrong" } });
        Denied(engine, wrongQuote, Text, "An invalid final quote cannot retire an earlier valid target.");
        engine.Marks.Single(mark => mark.Id == delta.Id).Stale = true;
        request.ContextId = engine.ReadContext(Document, Revision, Text, Snapshot);
        Denied(engine, request, Text, "A stale final target cannot produce a partial recolor.");
        engine.Marks.Single(mark => mark.Id == delta.Id).Stale = false;
        engine.Marks.Single(mark => mark.Id == delta.Id).RequiresRemarking = true;
        request.ContextId = engine.ReadContext(Document, Revision, Text, Snapshot);
        Denied(engine, request, Text, "An unresolved partial rewrite cannot be recolored as a current marker.");
        engine.Marks.Single(mark => mark.Id == delta.Id).RequiresRemarking = false;
        request.ContextId = engine.ReadContext(Document, Revision, Text, Snapshot);
        engine.AddOrUpdateRule(Rule("red", "#E57575", "protect"));
        Denied(engine, request, Text, "A local rule change invalidates a previously read recolor request.");
    }

    private static void SameColorAndOrdinaryHighlighting()
    {
        var engine = Engine(); PassageMark source = Add(engine, "Wrai_same_color_source", "green", "beta");
        string rules = Json.Serialize(engine.Rules);
        HighlightReceipt receipt = engine.Highlight(Recolor(engine, "same_color_replacement", "green", source), Text, Snapshot);
        Equal(true, receipt.MarkIds.Single() != source.Id && Replaced(receipt).Single() == source.Id && engine.Marks.Count == 1,
            "Even same-color recoloring produces a fresh identity with one uniform replacement receipt.");
        PassageMark retained = engine.Marks.Single(); string exact = Json.Serialize(retained);
        var ordinary = Change(Recolor(engine, "ordinary_addition", "red", retained), "RecolorMarkIds", null);
        ordinary.Ranges = new List<HighlightRange> { new HighlightRange { From = 11, To = 16, Quote = "gamma" } };
        HighlightReceipt added = engine.Highlight(ordinary, Text, Snapshot);
        Equal(2, engine.Marks.Count, "An ordinary request continues to add a marker without retiring another passage.");
        Equal(exact, Json.Serialize(engine.Marks.Single(mark => mark.Id == retained.Id)), "Ordinary highlighting preserves existing source markers exactly.");
        var encoded = Json.Deserialize<Dictionary<string, object>>(Json.Serialize(added));
        Equal(true, encoded["ReplacedMarkIds"] == null, "Ordinary highlighting records no retired identities.");
        Equal(rules, Json.Serialize(engine.Rules), "Same-color replacement and ordinary extension keep all rule definitions exact.");
    }

    private static void UnselectedOverlapRemainsProtected()
    {
        var engine = Engine(); PassageMark selected = Add(engine, "Wrai_selected", "green", "beta");
        Add(engine, "Wrai_unselected_protection", "green", "beta");
        Denied(engine, Recolor(engine, "keep_other_protection", "red", selected), Text, "Recoloring only one overlapping mark cannot repaint unselected protection.");
        engine = Engine(); selected = Add(engine, "Wrai_selected_visual", "visual", "beta");
        Add(engine, "Wrai_unselected_prompt", "blue", "beta");
        Denied(engine, Change(Recolor(engine, "do_not_hide_visual", null, selected), "Color", "#123ABC"), Text,
            "An unselected dominant prompt still blocks an invisible neutral recolor.");
    }

    private static void UnresolvedNativeReviewIsRefused()
    {
        var engine = Engine(); PassageMark source = Add(engine, "Wrai_tracked_source", "blue", "beta");
        string context = engine.ReadContext(Document, Revision, Text, Snapshot);
        engine.Submit(new Proposal { Id = "tracked", DocumentId = Document, Revision = Revision, From = 6, To = 10,
            Quote = "beta", Replacement = "clear beta", Rationale = "" }, context, Text, Snapshot);
        engine.BeginApplication("tracked"); engine.RecordTracked("tracked", "WraiReview_test", PolicyEngine.Hash("native review"), "Alpha clear beta gamma delta.");
        PolicyState state = engine.ExportState(); Proposal proposal = state.Proposals.Single();
        proposal.NativeOriginalParagraphQuote = Text;
        proposal.NativeOriginalMarks = new List<NativeMarkWitness> { new NativeMarkWitness { Id = source.Id, Bookmark = "WraiO_original",
            BeforeQuote = source.Quote, AfterQuote = "clear beta" } };
        engine = new PolicyEngine(contract, state); engine.Share(Document);
        Denied(engine, Recolor(engine, "tracked_witness_denied", "red", engine.Marks.Single()), Text, "A pending native tracked witness prevents recoloring its source identity.");
        state = engine.ExportState(); state.Proposals.Single().Status = "applied"; state.Marks.Single().NativeBookmark = "WraiO_original";
        engine = new PolicyEngine(contract, state); engine.Share(Document);
        Denied(engine, Recolor(engine, "native_alias_denied", "red", engine.Marks.Single()), Text, "An aliased native original anchor cannot be recolored.");
    }

    private static void ReplacementAtMarkerLimit()
    {
        var engine = Engine(); string text = String.Join(" ", Enumerable.Repeat("x", 500));
        for (int index = 0; index < 500; index++) engine.Marks.Add(new PassageMark { Id = "Wrai_limit_" + index, RuleId = "green",
            From = index * 2, To = index * 2 + 1, Quote = "x", Prompt = "" });
        var request = Json.Deserialize<HighlightRequest>(Json.Serialize(new { DocumentId = Document, Revision = Revision, RequestId = "recolor_at_limit",
            ContextId = engine.ReadContext(Document, Revision, text, Snapshot), ColorId = "red", RecolorMarkIds = engine.Marks.Take(64).Select(mark => mark.Id).ToArray(),
            Ranges = engine.Marks.Take(64).Select(mark => new { From = mark.From, To = mark.To, Quote = mark.Quote }).ToArray() }));
        HighlightReceipt receipt = engine.Highlight(request, text, Snapshot);
        Equal(500, engine.Marks.Count, "Replacing sixty-four markers at the limit consumes no additional marker slots.");
        Equal(64, Replaced(receipt).Length, "The maximum replacement batch records every retired identity.");
        Equal(64, engine.Marks.Count(mark => mark.RuleId == "red"), "Every selected limit-bound marker is recolored exactly once.");
        Equal(436, engine.Marks.Count(mark => mark.RuleId == "green"), "All unselected limit-bound markers retain their protection.");
    }

    private static void InvalidReplacementReceiptsAreRefused()
    {
        var engine = Engine(); PassageMark source = Add(engine, "Wrai_receipt_source", "green", "beta");
        engine.Highlight(Recolor(engine, "persisted_recolor", "red", source), Text, Snapshot);
        PolicyState state = engine.ExportState();
        foreach (object replaced in new object[] { null, new string[0], new[] { "wrong_source" }, new[] { source.Id, source.Id } })
            Refused(() => new PolicyEngine(contract, ReceiptChanged(state, "ReplacedMarkIds", replaced)), "Persisted replaced IDs must match the exact paired request.");
        Refused(() => new PolicyEngine(contract, ReceiptChanged(state, "MarkIds", new[] { source.Id })),
            "A recolor receipt cannot reuse an old identity as its new identity.");
        Equal(1, new PolicyEngine(contract, state).Marks.Count, "A valid serialized replacement receipt reopens without resurrecting retired markers.");
    }

    private static PolicyState ReceiptChanged(PolicyState state, string key, object value)
    {
        var encoded = Json.Deserialize<Dictionary<string, object>>(Json.Serialize(state));
        var receipts = (System.Collections.IEnumerable)encoded["HighlightReceipts"];
        var receipt = (Dictionary<string, object>)receipts.Cast<object>().Single(); receipt[key] = value;
        return Json.Deserialize<PolicyState>(Json.Serialize(encoded));
    }

    private static void Denied(PolicyEngine engine, HighlightRequest request, string text, string message, string reason = null)
    {
        string before = Json.Serialize(engine.ExportState()); string actual = Refused(() => engine.Highlight(request, text, Snapshot), message);
        if (reason != null) Equal(true, actual.IndexOf(reason, StringComparison.OrdinalIgnoreCase) >= 0,
            "The refusal identifies the actual recolor constraint: expected " + reason + "; received " + actual);
        Equal(before, Json.Serialize(engine.ExportState()), "Refused recolor preserves all rule definitions, marker identities and replay receipts.");
    }

    private static string Refused(Action action, string message)
    {
        assertions++;
        try { action(); } catch (InvalidOperationException error) { return error.Message; }
        throw new Exception("Expected policy refusal: " + message);
    }

    private static void Equal<T>(T expected, T actual, string message)
    {
        assertions++;
        if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception(message + " Expected " + expected + ", got " + actual + ".");
    }
}
