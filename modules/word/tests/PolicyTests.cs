using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using UnderTree.WordAddin;

internal static class PolicyTests
{
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 16777216 };
    private static string contract;
    private static int assertions;
    private const string Text = "Alpha \ud83d\ude00 beta gamma.";
    private const string DocumentId = "word_test_document";
    private const string Revision = "revision_1";
    private const string Snapshot = "all-stories-native-identity-1";

    public static int Main(string[] arguments)
    {
        try
        {
            contract = File.ReadAllText(arguments.Length > 0 ? arguments[0] : "modules/review/selective-editor-contract.json");
            FixedWordEditing();
            SharedRangeCases();
            SharedPreservationCases();
            NeutralHighlighting();
            DistinctNeutralColor();
            RequestedExistingColorsExtendUnchangedRules();
            RequestedRgbReplayAndLocalAssignment();
            InvalidRequestedRgbIsAtomic();
            RgbIdentitySelectionAndConflicts();
            ColorBatchBoundaries();
            HighlightOverlapAtomicity();
            ExactUtf16Boundaries();
            ContextAndRevocation();
            OverlapAndStalePermissions();
            DiscardRequiresCompleteLocalApproval();
            PendingAcceptanceRechecksCurrentState();
            ReplayCannotApplyAgainOrUpgrade();
            PersistenceDoesNotRestoreSharing();
            LimitsAndCorruptState();
            Console.WriteLine("Word policy: " + assertions + " assertions passed.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
    }

    private static void SharedPreservationCases()
    {
        Dictionary<string, object> shared = Json.Deserialize<Dictionary<string, object>>(contract);
        foreach (Dictionary<string, object> scenario in (IEnumerable)shared["preservationCases"])
        {
            string text = (string)scenario["text"];
            PolicyEngine engine = new PolicyEngine(contract);
            engine.Rules.Clear();
            var identities = new Dictionary<string, string>();
            foreach (Dictionary<string, object> highlight in (IEnumerable)scenario["marks"])
            {
                string id = "Wrai_" + Guid.NewGuid().ToString("N");
                identities.Add((string)highlight["id"], id);
                engine.Rules.Add(Rule(id, (string)highlight["action"], (string)highlight["accessMode"]));
                engine.Marks.Add(Mark(id, id, (int)highlight["from"], (int)highlight["to"], text));
            }
            var change = (Dictionary<string, object>)scenario["proposal"];
            Proposal request = Request((string)scenario["id"], (int)change["from"], (int)change["to"], (string)change["quote"], (string)change["replacement"]);
            request.ReplacementParts = new List<ReplacementPart>();
            foreach (Dictionary<string, object> part in (IEnumerable)change["replacementParts"])
                request.ReplacementParts.Add(part.ContainsKey("text") ? new ReplacementPart { Text = (string)part["text"] } :
                    new ReplacementPart { ProtectedId = identities[(string)part["protectedId"]] });
            string token = Read(engine, text);
            string expected = (string)scenario["result"];
            if (expected == "denied")
            {
                Denied(() => engine.Submit(request, token, text, Snapshot), (string)scenario["id"]);
                Equal(0, engine.Proposals.Count, "Denied preservation creates no receipt.");
                continue;
            }
            Authorization result = engine.Submit(request, token, text, Snapshot);
            Equal("direct", result.Mode, "Word preservation uses fixed direct wording behavior, including legacy prompts: " + scenario["id"]);
            Equal((string)scenario["expectedText"], engine.ResolveReplacement(result.Proposal, text), "Protected units reconstruct the exact expected wording.");
            List<ProtectedPlacement> placements = engine.ProtectedPlacements(result.Proposal, text);
            Equal(request.ReplacementParts.Count(part => part.ProtectedId != null), placements.Count, "Every original identity has one reanchor destination.");
            foreach (ProtectedPlacement placement in placements)
            {
                string unit = engine.Marks.First(mark => mark.Id == placement.Id).Quote;
                Equal(unit, PolicyEngine.Slice((string)scenario["expectedText"], placement.From, placement.To), "Native identity maps to the whole immutable unit.");
            }
            if (result.Mode == "direct") engine.ValidateDirect(request.Id, DocumentId, Revision, text, Snapshot);
            else engine.ValidateAcceptance(request.Id, DocumentId, Revision, text, Snapshot);
            engine.Rules[0].Prompt += " changed";
            if (result.Mode == "direct") Denied(() => engine.ValidateDirect(request.Id, DocumentId, Revision, text, Snapshot), "Policy changed before direct preservation.");
            else Denied(() => engine.ValidateAcceptance(request.Id, DocumentId, Revision, text, Snapshot), "Policy changed before preservation acceptance.");
        }
    }

    private static void NeutralHighlighting()
    {
        PolicyEngine engine = new PolicyEngine(contract);
        engine.Rules.Clear();
        engine.Share(DocumentId);
        var request = new HighlightRequest { DocumentId = DocumentId, Revision = Revision, RequestId = "highlight", ContextId = "forged",
            Ranges = new List<HighlightRange> { new HighlightRange { From = 9, To = 13, Quote = "beta" }, new HighlightRange { From = 14, To = 19, Quote = "gamma" } } };
        Denied(() => engine.Highlight(request, Text, Snapshot), "Coloring needs complete context.");
        request.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        request.Ranges[1].Quote = "wrong";
        Denied(() => engine.Highlight(request, Text, Snapshot), "A bad batch range rejects the whole batch.");
        Equal(0, engine.Rules.Count, "A bad batch allocates no color.");
        Equal(0, engine.Marks.Count, "A bad batch leaves no partial marks.");
        request.Ranges[1].Quote = "gamma";
        HighlightReceipt receipt = engine.Highlight(request, Text, Snapshot);
        Equal(2, engine.Marks.Count, "Batch marks both exact occurrences.");
        Equal("neutral", engine.Rules[0].Action, "Fresh AI color has no function.");
        Equal("", engine.Rules[0].Prompt, "Fresh AI color has no prompt.");
        Equal(2, engine.Highlight(request, Text, Snapshot).MarkIds.Count, "Lost marking reply replays the original receipt.");
        Equal(2, engine.Marks.Count, "Replay adds no duplicate marks.");
        string context = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        Denied(() => engine.Submit(Request("neutral_edit", 9, 13, "beta", "changed"), context, Text, Snapshot), "Neutral is not suggestion or direct permission.");
        request.RequestId = "reuse_neutral"; request.ColorId = receipt.ColorId; request.ContextId = context;
        engine.Highlight(request, Text, Snapshot);
        Equal(1, engine.Rules.Count, "Still-neutral reuse retains its identity.");
        engine.AddOrUpdateRule(new ColorRule { Id = receipt.ColorId, Label = "Protect citations", Color = engine.Rules[0].Color,
            Action = "protect", Prompt = "", AccessMode = "suggestions" });
        request.RequestId = "rule_bearing"; request.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        string assignedRule = Json.Serialize(engine.Rules.Single(rule => rule.Id == receipt.ColorId));
        Equal(receipt.ColorId, engine.Highlight(request, Text, Snapshot).ColorId, "A fresh complete read permits extending a locally assigned color rule.");
        Equal(assignedRule, Json.Serialize(engine.Rules.Single(rule => rule.Id == receipt.ColorId)), "Extending a color leaves its local function definition unchanged.");
        request.RequestId = "highlight";
        Denied(() => engine.Highlight(request, Text, Snapshot), "Changed policy cannot revive a previous color request.");
        engine.Revoke(); engine.Share(DocumentId);
        Denied(() => engine.Highlight(request, Text, Snapshot), "Marking requests cannot survive a revoked sharing epoch.");
        PolicyEngine reopened = new PolicyEngine(contract, engine.ExportState());
        reopened.Share(DocumentId);
        Denied(() => reopened.Highlight(request, Text, Snapshot), "Reopen cannot revive a marking receipt.");
        ColorRule neutral = Rule("invalid_neutral", "neutral", "direct");
        Denied(() => engine.AddOrUpdateRule(neutral), "Neutral cannot assign direct access.");
        neutral.AccessMode = "suggestions"; neutral.Prompt = "hidden rule";
        Denied(() => engine.AddOrUpdateRule(neutral), "Neutral cannot contain a hidden prompt.");
    }

    private static void DistinctNeutralColor()
    {
        var policy = new PolicyEngine(contract);
        policy.Rules.Clear();
        policy.AddOrUpdateRule(new ColorRule { Id = "near_user_color", Label = "User's citation color", Color = "#F6BD61", Action = "protect", Prompt = "", AccessMode = "suggestions" });
        policy.Share(DocumentId);
        string context = policy.ReadContext(DocumentId, Revision, Text, Snapshot);
        HighlightReceipt receipt = policy.Highlight(new HighlightRequest { DocumentId = DocumentId, Revision = Revision, RequestId = "distinct_neutral",
            ContextId = context, Ranges = new List<HighlightRange> { new HighlightRange { From = 9, To = 13, Quote = "beta" } } }, Text, Snapshot);
        string chosen = policy.Rules.First(rule => rule.Id == receipt.ColorId).Color;
        double red = Convert.ToInt32(chosen.Substring(1, 2), 16) - 246;
        double green = Convert.ToInt32(chosen.Substring(3, 2), 16) - 189;
        double blue = Convert.ToInt32(chosen.Substring(5, 2), 16) - 97;
        Equal(true, Math.Sqrt(red * red + green * green + blue * blue) >= 40,
            "A fresh neutral color remains visually distinct from a nearly identical user-defined palette hue.");
        Equal("#F6BD61", policy.Rules.First(rule => rule.Id == "near_user_color").Color, "Allocation never changes the user's existing color.");
    }

    private static void RequestedExistingColorsExtendUnchangedRules()
    {
        foreach (string behavior in new[] { "protect", "reject", "discard", "prompt_suggestions", "prompt_direct" })
        {
            string action = behavior.StartsWith("prompt_", StringComparison.Ordinal) ? "prompt" : behavior;
            string mode = behavior == "prompt_direct" ? "direct" : "suggestions";
            var engine = new PolicyEngine(contract);
            engine.Rules.Clear();
            var functional = new ColorRule { Id = "existing_green", Label = "Existing local function", Color = "#74b9a4",
                Action = action, Prompt = action == "prompt" ? "Improve clarity." : "", AccessMode = mode };
            engine.AddOrUpdateRule(functional);
            engine.AddMark(Mark("existing_passage", functional.Id, 14, 19, Text), Text);
            string originalRule = Json.Serialize(engine.Rules[0]), originalMark = Json.Serialize(engine.Marks[0]);
            engine.Share(DocumentId);
            var request = RgbMarking("rgb_collision_" + behavior, "#74B9A4");
            request.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
            HighlightReceipt receipt = engine.Highlight(request, Text, Snapshot);
            ColorRule selected = engine.Rules.Single(rule => rule.Id == receipt.ColorId);
            Equal(true, String.Equals("#74b9a4", selected.Color, StringComparison.OrdinalIgnoreCase),
                "An explicit green request uses the requested RGB even when it matches a local " + action + " color.");
            Equal(functional.Id, selected.Id, "Selecting an existing RGB extends that existing local rule identity.");
            Equal(1, engine.Rules.Count, "Extending an existing color allocates no duplicate rule definition.");
            Equal(action, selected.Action, "New passages use the selected existing function.");
            Equal(mode, selected.AccessMode, "New passages retain the selected existing access mode.");
            Equal(originalRule, Json.Serialize(engine.Rules.Single(rule => rule.Id == functional.Id)), "The original same-RGB local rule remains exact.");
            Equal(originalMark, Json.Serialize(engine.Marks.Single(mark => mark.Id == "existing_passage")), "The existing functional passage retains its exact identity and rule.");
            Equal(true, receipt.MarkIds.All(id => engine.Marks.Single(mark => mark.Id == id).RuleId == functional.Id),
                "Every new passage refers to the explicitly selected existing rule.");
            string token = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
            Proposal proposal = Request("rgb_existing_edit_" + behavior, 9, 13, "beta", action == "discard" ? "" : "changed");
            if (action == "protect" || action == "reject")
                Denied(() => engine.Submit(proposal, token, Text, Snapshot), "Extending a restrictive color protects the new exact passage.");
            else
                Equal(action == "discard" ? "suggestions" : "direct", engine.Submit(proposal, token, Text, Snapshot).Mode,
                    "The unchanged rule retains its function under Word's fixed wording behavior.");
            request = new HighlightRequest { DocumentId = DocumentId, Revision = Revision, RequestId = "rgb_functional_id_" + behavior,
                ContextId = token, ColorId = functional.Id, Ranges = new List<HighlightRange> { new HighlightRange { From = 0, To = 5, Quote = "Alpha" } } };
            HighlightReceipt byIdentity = engine.Highlight(request, Text, Snapshot);
            Equal(functional.Id, byIdentity.ColorId, "An explicit functional colorId can extend the same unchanged local rule.");
            Equal(originalRule, Json.Serialize(engine.Rules.Single(rule => rule.Id == functional.Id)), "Neither RGB nor identity selection can change the existing rule definition.");
            var reopened = new PolicyEngine(contract, Json.Deserialize<PolicyState>(Json.Serialize(engine.ExportState())));
            ColorRule persisted = reopened.Rules.Single(rule => rule.Id == functional.Id);
            Equal(originalRule, Json.Serialize(persisted), "Reopening preserves the original local rule definition exactly.");
            Equal(true, String.Equals(selected.Color, persisted.Color, StringComparison.OrdinalIgnoreCase) &&
                reopened.Marks.Single(mark => mark.Id == receipt.MarkIds[0]).RuleId == persisted.Id,
                "Serialization preserves exact selected RGB and newly extended passage linkage.");
        }
    }

    private static void RequestedRgbReplayAndLocalAssignment()
    {
        var engine = new PolicyEngine(contract); engine.Rules.Clear(); engine.Share(DocumentId);
        var request = RgbMarking("rgb_replay", "#008000");
        request.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        HighlightReceipt receipt = engine.Highlight(request, Text, Snapshot);
        Equal(true, String.Equals("#008000", engine.Rules.Single(rule => rule.Id == receipt.ColorId).Color, StringComparison.OrdinalIgnoreCase),
            "A requested non-palette green is honored exactly.");
        ColorRule freshRule = engine.Rules.Single(rule => rule.Id == receipt.ColorId);
        Equal("neutral", freshRule.Action, "A previously unused RGB starts with no assigned function.");
        Equal("", freshRule.Prompt, "A previously unused RGB starts with no assigned prompt.");
        string token = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        Denied(() => engine.Submit(Request("rgb_fresh_no_edit", 9, 13, "beta", "changed"), token, Text, Snapshot),
            "A new RGB cannot authorize wording changes until the local user assigns it a function.");
        string unchanged = Json.Serialize(engine.ExportState());
        Equal(receipt.ColorId, engine.Highlight(request, Text, Snapshot).ColorId, "A lost RGB marking reply replays its exact neutral identity.");
        Equal(unchanged, Json.Serialize(engine.ExportState()), "RGB replay creates no extra rules, marks or receipts.");
        var changed = RgbMarking("rgb_replay", "#123ABC"); changed.ContextId = request.ContextId;
        Denied(() => engine.Highlight(changed, Text, Snapshot), "Reusing a request ID with another RGB is rejected.");
        Equal(unchanged, Json.Serialize(engine.ExportState()), "Changed RGB replay cannot recolor any existing mark.");
        var reuse = new HighlightRequest { DocumentId = DocumentId, Revision = Revision, RequestId = "rgb_neutral_reuse",
            ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot), ColorId = receipt.ColorId,
            Ranges = new List<HighlightRange> { new HighlightRange { From = 9, To = 13, Quote = "beta" } } };
        Equal(receipt.ColorId, engine.Highlight(reuse, Text, Snapshot).ColorId, "Still-neutral colorId reuse remains supported for arbitrary RGB.");
        reuse.RequestId = "rgb_assignment_race"; reuse.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        engine.AddOrUpdateRule(new ColorRule { Id = receipt.ColorId, Label = "Locally protected beta", Color = "#008000",
            Action = "protect", Prompt = "", AccessMode = "suggestions" });
        unchanged = Json.Serialize(engine.ExportState());
        Denied(() => engine.Highlight(reuse, Text, Snapshot), "A local assignment invalidates an outstanding neutral-color request.");
        Equal(unchanged, Json.Serialize(engine.ExportState()), "A stale pre-assignment witness cannot extend or alter the new rule.");
        reuse.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        string assigned = Json.Serialize(engine.Rules.Single(rule => rule.Id == receipt.ColorId));
        Equal(receipt.ColorId, engine.Highlight(reuse, Text, Snapshot).ColorId, "A fresh witness can extend the newly assigned local rule.");
        Equal(assigned, Json.Serialize(engine.Rules.Single(rule => rule.Id == receipt.ColorId)), "Extending after a local assignment cannot change that definition.");
        var fresh = RgbMarking("rgb_fresh_after_assignment", "#008000");
        fresh.Ranges = new List<HighlightRange> { new HighlightRange { From = 14, To = 19, Quote = "gamma" } };
        fresh.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        HighlightReceipt next = engine.Highlight(fresh, Text, Snapshot);
        Equal(receipt.ColorId, next.ColorId, "A later uniquely matching RGB request extends the selected existing local rule.");
        Equal("protect", engine.Rules.Single(rule => rule.Id == receipt.ColorId).Action, "Reusing RGB never repurposes the user's assigned rule.");
        Equal("protect", engine.Rules.Single(rule => rule.Id == next.ColorId).Action, "The explicitly reused existing green protects its new passage.");
        token = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        Denied(() => engine.Submit(Request("rgb_after_assignment_no_edit", 14, 19, "gamma", "changed"), token, Text, Snapshot),
            "Newly extended protection is enforced on the exact added passage.");
    }

    private static void InvalidRequestedRgbIsAtomic()
    {
        foreach (string color in new[] { "", "green", "#ABC", "#0000000", "#GG0000", "#000000\n" })
        {
            var engine = new PolicyEngine(contract); engine.Share(DocumentId);
            var request = RgbMarking("rgb_invalid", color); request.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
            string unchanged = Json.Serialize(engine.ExportState());
            Denied(() => engine.Highlight(request, Text, Snapshot), "Invalid requested RGB must be rejected: " + Json.Serialize(color));
            Equal(unchanged, Json.Serialize(engine.ExportState()), "Invalid RGB allocates no rule, mark or replay receipt.");
        }
    }

    private static HighlightRequest RgbMarking(string id, string color)
    {
        return Json.Deserialize<HighlightRequest>(Json.Serialize(new { DocumentId = DocumentId, Revision = Revision, RequestId = id, Color = color,
            Ranges = new[] { new { From = 9, To = 13, Quote = "beta" } } }));
    }

    private static void RgbIdentitySelectionAndConflicts()
    {
        var engine = new PolicyEngine(contract); engine.Rules.Clear();
        engine.AddOrUpdateRule(new ColorRule { Id = "same_rgb_protect", Label = "Keep unchanged", Color = "#74b9a4",
            Action = "protect", Prompt = "", AccessMode = "suggestions" });
        engine.AddOrUpdateRule(new ColorRule { Id = "same_rgb_neutral", Label = "Color only", Color = "#74B9A4",
            Action = "neutral", Prompt = "", AccessMode = "suggestions" });
        string originalRules = Json.Serialize(engine.Rules); engine.Share(DocumentId);
        var request = RgbMarking("rgb_ambiguous", "#74b9a4");
        request.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        DeniedHighlightUnchanged(engine, request, Text, "A duplicate RGB requires an exact identity instead of guessing a function.");
        request.RequestId = "rgb_exact_neutral"; request.Color = null; request.ColorId = "same_rgb_neutral";
        HighlightReceipt neutral = engine.Highlight(request, Text, Snapshot);
        Equal("same_rgb_neutral", neutral.ColorId, "Exact neutral identity selection disambiguates duplicate RGB.");
        Equal(true, neutral.MarkIds.All(id => engine.Marks.Single(mark => mark.Id == id).RuleId == "same_rgb_neutral"),
            "The chosen neutral identity is retained on every new passage.");
        request.RequestId = "rgb_exact_protect"; request.ColorId = "same_rgb_protect";
        request.Ranges = new List<HighlightRange> { new HighlightRange { From = 0, To = 5, Quote = "Alpha" } };
        request.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        HighlightReceipt protection = engine.Highlight(request, Text, Snapshot);
        Equal("same_rgb_protect", protection.ColorId, "Exact protective identity selection disambiguates the same RGB.");
        Equal("same_rgb_protect", engine.Marks.Single(mark => mark.Id == protection.MarkIds[0]).RuleId,
            "Selecting an existing protection identity extends that precise rule.");
        Equal(originalRules, Json.Serialize(engine.Rules), "Choosing between duplicate RGB identities changes neither definition.");
        foreach (string color in new[] { "#74B9A4", "#123ABC" })
        {
            request = RgbMarking("rgb_conflicting_selectors", color); request.ColorId = "same_rgb_neutral";
            request.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
            DeniedHighlightUnchanged(engine, request, Text, "Supplying color and colorId together is refused even when their RGB agrees.");
        }
        request = RgbMarking("rgb_missing_identity", null); request.ColorId = "missing_color";
        request.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        DeniedHighlightUnchanged(engine, request, Text, "An absent exact color identity cannot silently allocate or select another rule.");
    }

    private static void ColorBatchBoundaries()
    {
        string text = String.Join(" ", Enumerable.Repeat("x", 65));
        var engine = new PolicyEngine(contract); engine.Rules.Clear(); engine.Share(DocumentId);
        var request = RgbMarking("rgb_batch_64", "#00BFA5");
        request.Ranges = Enumerable.Range(0, 64).Select(index => new HighlightRange { From = index * 2, To = index * 2 + 1, Quote = "x" }).ToList();
        request.ContextId = engine.ReadContext(DocumentId, Revision, text, Snapshot);
        HighlightReceipt receipt = engine.Highlight(request, text, Snapshot);
        Equal(64, receipt.MarkIds.Distinct().Count(), "The maximum batch creates sixty-four distinct exact passage identities.");
        Equal(64, engine.Marks.Count, "A sixty-four-range batch creates every requested marker once.");
        Equal(true, engine.Marks.All(mark => mark.To == mark.From + 1 && mark.From % 2 == 0 && mark.From < 128 && mark.Quote == "x" && mark.RuleId == receipt.ColorId),
            "Every maximum-batch marker has the requested exact geometry and selected rule.");
        string unchanged = Json.Serialize(engine.ExportState());
        engine.Highlight(request, text, Snapshot);
        Equal(unchanged, Json.Serialize(engine.ExportState()), "Replaying the maximum batch creates no extra markers or receipt.");
        request.RequestId = "rgb_batch_65"; request.Ranges.Add(new HighlightRange { From = 128, To = 129, Quote = "x" });
        request.ContextId = engine.ReadContext(DocumentId, Revision, text, Snapshot);
        DeniedHighlightUnchanged(engine, request, text, "A sixty-five-range batch is refused atomically.");
        var invalid = new PolicyEngine(contract); invalid.Rules.Clear(); invalid.Share(DocumentId);
        request = RgbMarking("rgb_batch_bad_last", "#00BFA5");
        request.Ranges = Enumerable.Range(0, 64).Select(index => new HighlightRange { From = index * 2, To = index * 2 + 1, Quote = index == 63 ? "y" : "x" }).ToList();
        request.ContextId = invalid.ReadContext(DocumentId, Revision, text, Snapshot);
        DeniedHighlightUnchanged(invalid, request, text, "A wrong final quote in the maximum batch leaves no valid-prefix marks or new RGB rule.");
    }

    private static void HighlightOverlapAtomicity()
    {
        foreach (string existingAction in new[] { "protect", "reject" })
        foreach (string requestedAction in new[] { "neutral", "prompt", "protect", "reject" })
        {
            var engine = new PolicyEngine(contract); engine.Rules.Clear();
            engine.AddOrUpdateRule(new ColorRule { Id = "existing_restriction", Label = "Existing protected passage", Color = "#74b9a4",
                Action = existingAction, Prompt = "", AccessMode = "suggestions" });
            engine.AddMark(Mark("existing_native_unit", "existing_restriction", 9, 13, Text), Text);
            if (requestedAction != "neutral")
                engine.AddOrUpdateRule(new ColorRule { Id = "requested_function", Label = "Existing selected function", Color = "#123ABC",
                    Action = requestedAction, Prompt = requestedAction == "prompt" ? "Keep meaning." : "", AccessMode = "suggestions" });
            engine.Share(DocumentId);
            var request = RgbMarking("rgb_overlap_" + existingAction + "_" + requestedAction, requestedAction == "neutral" ? "#123ABC" : null);
            if (requestedAction != "neutral") request.ColorId = "requested_function";
            request.Ranges = new List<HighlightRange> { new HighlightRange { From = 0, To = 5, Quote = "Alpha" },
                new HighlightRange { From = 9, To = 19, Quote = "beta gamma" } };
            request.ContextId = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
            DeniedHighlightUnchanged(engine, request, Text,
                "Different-RGB " + requestedAction + " cannot repaint existing " + existingAction + ", including equal-priority overlap; no prefix is marked.");
            request.RequestId += "_same_rgb"; request.Color = null; request.ColorId = "existing_restriction";
            string originalRule = Json.Serialize(engine.Rules.Single(rule => rule.Id == "existing_restriction"));
            string originalMark = Json.Serialize(engine.Marks.Single(mark => mark.Id == "existing_native_unit"));
            HighlightReceipt same = engine.Highlight(request, Text, Snapshot);
            Equal("existing_restriction", same.ColorId, "A visible same-RGB reuse succeeds without changing the protected color.");
            Equal(originalRule, Json.Serialize(engine.Rules.Single(rule => rule.Id == "existing_restriction")), "Same-RGB overlap preserves the original local rule definition.");
            Equal(originalMark, Json.Serialize(engine.Marks.Single(mark => mark.Id == "existing_native_unit")), "Same-RGB overlap preserves the original protected passage identity and bounds.");
        }
        var prompt = new PolicyEngine(contract); prompt.Rules.Clear();
        prompt.AddOrUpdateRule(new ColorRule { Id = "visible_prompt", Label = "Existing prompt", Color = "#E5BE50",
            Action = "prompt", Prompt = "Keep meaning.", AccessMode = "suggestions" });
        prompt.AddMark(Mark("prompt_native_unit", "visible_prompt", 9, 13, Text), Text);
        prompt.Share(DocumentId);
        var hidden = RgbMarking("rgb_hidden_by_prompt", "#123ABC");
        hidden.Ranges = new List<HighlightRange> { new HighlightRange { From = 0, To = 5, Quote = "Alpha" },
            new HighlightRange { From = 9, To = 19, Quote = "beta gamma" } };
        hidden.ContextId = prompt.ReadContext(DocumentId, Revision, Text, Snapshot);
        DeniedHighlightUnchanged(prompt, hidden, Text, "A new neutral color hidden by existing prompt paint is refused with no invisible success or prefix mutation.");
        prompt.AddOrUpdateRule(new ColorRule { Id = "same_rgb_visual", Label = "Matching visual color", Color = "#e5be50",
            Action = "neutral", Prompt = "", AccessMode = "suggestions" });
        hidden.RequestId = "rgb_visible_same_prompt"; hidden.Color = null; hidden.ColorId = "same_rgb_visual";
        hidden.ContextId = prompt.ReadContext(DocumentId, Revision, Text, Snapshot);
        HighlightReceipt visual = prompt.Highlight(hidden, Text, Snapshot);
        Equal("same_rgb_visual", visual.ColorId, "An exact same-RGB neutral identity remains visibly truthful under prompt paint.");
        string token = prompt.ReadContext(DocumentId, Revision, Text, Snapshot);
        Denied(() => prompt.Submit(Request("rgb_neutral_gap_still_denied", 14, 19, "gamma", "changed"), token, Text, Snapshot),
            "Same-RGB visual overlap never extends a different prompt identity's editing coverage.");
    }

    private static void DeniedHighlightUnchanged(PolicyEngine engine, HighlightRequest request, string text, string message)
    {
        string before = Json.Serialize(engine.ExportState());
        Denied(() => engine.Highlight(request, text, Snapshot), message);
        Equal(before, Json.Serialize(engine.ExportState()), "Refusal preserves every rule, marker and replay receipt exactly.");
    }

    private static void SharedRangeCases()
    {
        Dictionary<string, object> shared = Json.Deserialize<Dictionary<string, object>>(contract);
        foreach (Dictionary<string, object> scenario in (IEnumerable)shared["rangeCases"])
        {
            string text = (string)scenario["text"];
            PolicyEngine engine = new PolicyEngine(contract);
            engine.Rules.Clear();
            int index = 0;
            foreach (Dictionary<string, object> highlight in (IEnumerable)scenario["marks"])
            {
                string id = "rule_" + index;
                engine.Rules.Add(Rule(id, (string)highlight["action"], (string)highlight["accessMode"]));
                engine.Marks.Add(Mark("mark_" + index, id, (int)highlight["from"], (int)highlight["to"], text));
                index++;
            }
            Dictionary<string, object> change = (Dictionary<string, object>)scenario["proposal"];
            Proposal request = Request((string)scenario["id"], (int)change["from"], (int)change["to"], (string)change["quote"], (string)change["replacement"]);
            string token = Read(engine, text);
            string expected = (string)scenario["result"];
            if (expected == "denied")
            {
                Denied(() => engine.Submit(request, token, text, Snapshot), (string)scenario["id"]);
                Equal(0, engine.Proposals.Count, "Denied proposals create no pending state.");
            }
            else
            {
                Authorization result = engine.Submit(request, token, text, Snapshot);
                bool discard = engine.Marks.Any(mark => mark.From < request.To && request.From < mark.To && engine.Rules.Any(rule => rule.Id == mark.RuleId && rule.Action == "discard"));
                Equal(discard ? "suggestions" : "direct", result.Mode, "Word adapts shared legacy prompt modes without changing discard: " + scenario["id"]);
                Equal(discard ? "pending" : "authorized", result.Proposal.Status, "Word fixed-edit range state: " + scenario["id"]);
                Equal(Text, text, "Policy never mutates source text.");
            }
        }
    }

    private static void ExactUtf16Boundaries()
    {
        PolicyEngine engine = DirectEngine();
        string token = Read(engine, Text);
        Equal("\ud83d\ude00 beta", PolicyEngine.Slice(Text, 6, 13), "UTF-16 astral quote.");
        Denied(() => PolicyEngine.Slice(Text, 7, 13), "Split surrogate start.");
        Denied(() => PolicyEngine.Slice(Text, 0, 7), "Split surrogate end.");
        Denied(() => PolicyEngine.Slice(Text, -1, 2), "Negative boundary.");
        Denied(() => PolicyEngine.Slice(Text, 0, Int32.MaxValue), "Range overflow.");
        Denied(() => engine.Submit(Request("split", 7, 13, Text.Substring(7, 6), "new"), token, Text, Snapshot), "Incomplete quoted Unicode.");
        Denied(() => engine.Submit(Request("unpaired", 9, 13, "beta", "\ud800"), token, Text, Snapshot), "Incomplete replacement Unicode.");
        Denied(() => engine.Submit(Request("null", 9, 13, "beta", "a\0b"), token, Text, Snapshot), "Null replacement.");
        Denied(() => engine.Submit(Request("unchanged", 9, 13, "beta", "beta"), token, Text, Snapshot), "No-op proposal.");
        Denied(() => engine.Submit(Request("empty", 9, 9, "", "new"), token, Text, Snapshot), "Insertion has no authorized source range.");
        Denied(() => engine.Submit(Request("bad_quote", 9, 13, "Beta", "new"), token, Text, Snapshot), "Quote exactness.");
        Denied(() => engine.Submit(Request("long", 9, 13, "beta", new string('x', PolicyEngine.MaxReplacement + 1)), token, Text, Snapshot), "Replacement bound.");
        Equal(0, engine.Proposals.Count, "Malformed text never creates a proposal.");
    }

    private static void ContextAndRevocation()
    {
        PolicyEngine engine = DirectEngine();
        Proposal request = Request("context", 9, 13, "beta", "clear beta");
        Denied(() => engine.ReadContext(DocumentId, Revision, Text, Snapshot), "Reading without sharing.");
        engine.Share(DocumentId);
        Denied(() => engine.Submit(request, "guessed", Text, Snapshot), "No native prior read.");
        string token = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        Denied(() => engine.Submit(request, "forged", Text, Snapshot), "Forged witness.");
        Denied(() => engine.Submit(request, token, Text + "!", Snapshot), "Source changed without a claimed revision update.");
        Denied(() => engine.Submit(request, token, Text, "another-story-changed"), "Ancillary manuscript story changed.");
        request.Revision = "another_revision";
        Denied(() => engine.Submit(request, token, Text, Snapshot), "Revision changed.");
        request.Revision = Revision;
        request.DocumentId = "another_document";
        Denied(() => engine.Submit(request, token, Text, Snapshot), "Cross-document proposal.");
        request.DocumentId = DocumentId;
        engine.Marks[0].Prompt = "Passage instructions changed locally.";
        Denied(() => engine.Submit(request, token, Text, Snapshot), "Direct mark field changes invalidate witness.");
        token = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        engine.Rules[0].AccessMode = "suggestions";
        Denied(() => engine.Submit(request, token, Text, Snapshot), "Direct rule field changes invalidate witness.");
        token = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        engine.InvalidateContext();
        Denied(() => engine.Submit(request, token, Text, Snapshot), "Explicit native invalidation.");
        token = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        engine.Revoke();
        Denied(() => engine.Submit(request, token, Text, Snapshot), "Revoked sharing.");
        engine.Share(DocumentId);
        Denied(() => engine.Submit(request, token, Text, Snapshot), "Re-sharing does not revive old witness.");
        token = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
        Equal("authorized", engine.Submit(request, token, Text, Snapshot).Proposal.Status, "Fresh complete read restores fixed direct wording permission.");
        PolicyEngine other = DirectEngine();
        other.Share(DocumentId);
        Denied(() => other.Submit(Request("other", 9, 13, "beta", "new"), token, Text, Snapshot), "Witness cannot be moved to another native engine.");
        Denied(() => other.ReadContext(DocumentId, Revision, Text, ""), "Missing native complete-context identity.");
    }

    private static void OverlapAndStalePermissions()
    {
        foreach (string action in new[] { "protect", "reject", "discard" })
        {
            PolicyEngine engine = DirectEngine();
            engine.Rules.Add(Rule("restriction", action, "suggestions"));
            engine.Marks.Add(Mark("restriction_mark", "restriction", 14, 19, Text));
            engine.Marks[1].Stale = true;
            string token = Read(engine, Text);
            Denied(() => engine.Submit(Request("stale_" + action, 9, 13, "beta", "new"), token, Text, Snapshot), "Stale restrictive marker fails closed globally: " + action);
            engine.Marks[1].Stale = false;
            engine.Marks[1].Quote = "wrong";
            token = engine.ReadContext(DocumentId, Revision, Text, Snapshot);
            Denied(() => engine.Submit(Request("wrong_" + action, 9, 13, "beta", "new"), token, Text, Snapshot), "Mismatching restrictive quote is stale even without its flag.");
        }
        PolicyEngine consumed = DirectEngine();
        consumed.Marks.Add(Mark("used", "prompt", 14, 19, Text));
        consumed.Marks[1].Stale = true;
        string read = Read(consumed, Text);
        Equal("direct", consumed.Submit(Request("separate", 9, 13, "beta", "new"), read, Text, Snapshot).Mode, "Separate stale prompt marker does not expand or block current permission.");
        consumed.Marks[0].Stale = true;
        read = consumed.ReadContext(DocumentId, Revision, Text, Snapshot);
        Denied(() => consumed.Submit(Request("no_current", 9, 13, "beta", "new"), read, Text, Snapshot), "Stale prompt supplies no coverage.");
    }

    private static void DiscardRequiresCompleteLocalApproval()
    {
        PolicyEngine engine = DirectEngine();
        engine.Rules.Add(Rule("delete", "discard", "suggestions"));
        engine.Marks.Add(Mark("delete_mark", "delete", 9, 13, Text));
        string token = Read(engine, Text);
        Denied(() => engine.Submit(Request("discard_reword", 9, 13, "beta", "new"), token, Text, Snapshot), "Discard cannot be bypassed by overlapping direct prompt.");
        Denied(() => engine.Submit(Request("discard_partial", 10, 13, "eta", ""), token, Text, Snapshot), "Partial discard deletion.");
        Authorization allowed = engine.Submit(Request("discard_full", 9, 13, "beta", ""), token, Text, Snapshot);
        Equal("suggestions", allowed.Mode, "Discard always awaits approval.");
        Denied(() => engine.ValidateDirect("discard_full", DocumentId, Revision, Text, Snapshot), "Discard is never direct.");
        Equal("pending", engine.ValidateAcceptance("discard_full", DocumentId, Revision, Text, Snapshot).Proposal.Status, "Local acceptance can validate full deletion.");
        engine.BeginApplication("discard_full");
        engine.RecordApplied("discard_full");
        Equal("applied", engine.Proposals[0].Status, "Native deletion receipt.");
    }

    private static void PendingAcceptanceRechecksCurrentState()
    {
        PolicyEngine engine = SuggestionEngine();
        string token = Read(engine, Text);
        engine.Submit(Request("approval", 9, 13, "beta", "new"), token, Text, Snapshot);
        engine.Proposals.Single().Status = "pending";
        Denied(() => engine.ValidateAcceptance("approval", DocumentId, "new_revision", Text, Snapshot), "Acceptance rechecks revision.");
        Denied(() => engine.ValidateAcceptance("approval", DocumentId, Revision, Text.Replace("beta", "Beta"), Snapshot), "Acceptance rechecks source and quote.");
        Denied(() => engine.ValidateAcceptance("approval", DocumentId, Revision, Text, "changed_header"), "Acceptance rechecks complete-context snapshot.");
        engine.Marks[0].From = 10;
        Denied(() => engine.ValidateAcceptance("approval", DocumentId, Revision, Text, Snapshot), "Acceptance rechecks current highlights.");
        engine.Marks[0].From = 0;
        engine.Rules[0].Action = "protect";
        Denied(() => engine.ValidateAcceptance("approval", DocumentId, Revision, Text, Snapshot), "Acceptance rechecks protection.");
        engine.Rules[0].Action = "prompt";
        engine.Revoke();
        Equal("pending", engine.ValidateAcceptance("approval", DocumentId, Revision, Text, Snapshot).Proposal.Status, "A local user can accept while offline and unshared.");
        engine.BeginApplication("approval");
        Denied(() => engine.ValidateAcceptance("approval", DocumentId, Revision, Text, Snapshot), "A native edit already in progress cannot be accepted again.");
        engine.RecordOutcomeUncertain("approval");
        Equal("outcome_uncertain", engine.Proposals[0].Status, "Failed native outcome is terminal.");
        Denied(() => engine.BeginApplication("approval"), "An uncertain outcome cannot be replayed.");
    }

    private static void ReplayCannotApplyAgainOrUpgrade()
    {
        PolicyEngine engine = DirectEngine();
        string token = Read(engine, Text);
        Proposal request = Request("replay", 9, 13, "beta", "new");
        request.Status = "applied";
        request.SourceHash = new string('a', 64);
        Authorization first = engine.Submit(request, token, Text, Snapshot);
        Equal("authorized", first.Proposal.Status, "A remote status assertion is ignored.");
        Equal(PolicyEngine.Hash(Text), first.Proposal.SourceHash, "Native source binding overrides supplied hash.");
        Equal(false, first.AlreadyRecorded, "First direct request is new.");
        Equal(true, engine.Submit(request, token, Text, Snapshot).AlreadyRecorded, "Repeated prepared direct request cannot be reapplied.");
        Equal("direct", engine.ValidateDirect("replay", DocumentId, Revision, Text, Snapshot).Mode, "Native direct recheck.");
        engine.BeginApplication("replay");
        Equal("applying", engine.Submit(request, token, Text, Snapshot).Proposal.Status, "Reentrant retry returns existing unfinished receipt.");
        engine.RecordApplied("replay");
        Authorization replay = engine.Submit(request, token, Text.Replace("beta", "new"), "new_snapshot");
        Equal(true, replay.AlreadyRecorded, "Lost acknowledgement replay is idempotent after source changed.");
        Equal("applied", replay.Proposal.Status, "Replay retains terminal receipt.");
        Denied(() => engine.ValidateDirect("replay", DocumentId, Revision, Text, Snapshot), "An applied edit cannot be applied again.");
        request.Replacement = "different";
        Denied(() => engine.Submit(request, token, Text, Snapshot), "A reused id cannot replace its original request.");
        engine.RecordOutcomeUncertain("replay");
        Equal("outcome_uncertain", engine.Proposals[0].Status, "A metadata failure after native replacement must not report confirmed success.");
        Denied(() => engine.BeginApplication("replay"), "A changed document with uncertain metadata cannot be replayed.");

        PolicyEngine pending = SuggestionEngine();
        token = Read(pending, Text);
        Proposal suggestion = Request("no_upgrade", 9, 13, "beta", "new");
        pending.Submit(suggestion, token, Text, Snapshot);
        pending.Proposals.Single().Status = "pending";
        pending.Rules[0].AccessMode = "direct";
        Authorization stillPending = pending.Submit(suggestion, token, Text, Snapshot);
        Equal(true, stillPending.AlreadyRecorded, "Existing suggestion is a replay.");
        Equal("pending", stillPending.Proposal.Status, "Replay cannot upgrade a suggestion after permissions change.");
        Equal("suggestions", stillPending.Mode, "Pending always requires local approval.");

        PolicyEngine revoked = DirectEngine();
        token = Read(revoked, Text);
        revoked.Submit(Request("revoke_direct", 9, 13, "beta", "new"), token, Text, Snapshot);
        revoked.Revoke();
        Denied(() => revoked.ValidateDirect("revoke_direct", DocumentId, Revision, Text, Snapshot), "Revocation before native direct application.");
        revoked.Share(DocumentId);
        Denied(() => revoked.ValidateDirect("revoke_direct", DocumentId, Revision, Text, Snapshot), "A new sharing epoch cannot revive a prepared direct edit.");
    }

    private static void PersistenceDoesNotRestoreSharing()
    {
        PolicyEngine original = SuggestionEngine();
        string token = Read(original, Text);
        Proposal request = Request("persisted", 9, 13, "beta", "new");
        original.Submit(request, token, Text, Snapshot);
        original.Proposals.Single().Status = "pending";
        PolicyState saved = Json.Deserialize<PolicyState>(Json.Serialize(original.ExportState()));
        PolicyEngine restored = new PolicyEngine(contract, saved);
        Equal(false, restored.IsShared, "Reopen never restores sharing.");
        Equal("pending", restored.Proposals[0].Status, "Pending survives reopen.");
        Equal("suggestions", restored.ValidateAcceptance("persisted", DocumentId, Revision, Text, Snapshot).Mode, "Current persisted suggestion can be approved locally.");
        Denied(() => restored.ReadContext(DocumentId, Revision, Text, Snapshot), "Reopen requires sharing consent.");
        restored.Share(DocumentId);
        Denied(() => restored.Submit(Request("old_read", 9, 13, "beta", "new"), token, Text, Snapshot), "Reopen never restores prior-read authority.");
        restored.Reject("persisted");
        Equal("rejected", restored.Proposals[0].Status, "Rejection is a persisted receipt.");
        Equal("rejected", restored.Submit(request, token, Text, Snapshot).Proposal.Status, "A rejected id cannot become pending again.");
        Equal("pending", original.Proposals[0].Status, "Export/reopen use independent state copies.");

        PolicyEngine direct = DirectEngine();
        token = Read(direct, Text);
        direct.Submit(Request("unfinished", 9, 13, "beta", "new"), token, Text, Snapshot);
        PolicyEngine uncertain = new PolicyEngine(contract, direct.ExportState());
        Equal("outcome_uncertain", uncertain.Proposals[0].Status, "Reopen cannot replay an unacknowledged direct edit.");
        direct.BeginApplication("unfinished");
        uncertain = new PolicyEngine(contract, direct.ExportState());
        Equal("outcome_uncertain", uncertain.Proposals[0].Status, "Reopen cannot replay an interrupted native transaction.");
    }

    private static void LimitsAndCorruptState()
    {
        PolicyEngine engine = new PolicyEngine(contract);
        Equal(4, engine.Rules.Count, "Shared default color count.");
        Equal("Keep unchanged", engine.Rules[0].Label, "Shared rule labels loaded from contract.");
        Equal(true, engine.Rules.All(candidate => candidate.AccessMode == (candidate.Action == "prompt" ? "direct" : "suggestions")), "Only Word wording defaults change; restrictive wire values remain compatible.");
        ColorRule rule = Rule("utf8", "prompt", "suggestions");
        rule.Label = new string('\u00e9', 50);
        rule.Prompt = String.Concat(Enumerable.Repeat("\ud83d\ude00", 1000));
        engine.AddOrUpdateRule(rule);
        Equal(5, engine.Rules.Count, "UTF-8 limits permit exact boundaries.");
        rule.Label += "x";
        Denied(() => engine.AddOrUpdateRule(rule), "UTF-8 name byte overrun.");
        rule.Label = "Valid";
        rule.Prompt += "x";
        Denied(() => engine.AddOrUpdateRule(rule), "UTF-8 prompt byte overrun.");
        foreach (string action in new[] { "protect", "reject", "discard" })
            Denied(() => engine.AddOrUpdateRule(Rule("forbidden_" + action, action, "direct")), "Restrictive colors cannot authorize direct: " + action);
        for (int index = engine.Rules.Count; index < 24; index++) engine.AddOrUpdateRule(Rule("rule_" + index, "prompt", "suggestions"));
        Denied(() => engine.AddOrUpdateRule(Rule("extra", "prompt", "suggestions")), "Rule count bound.");

        PolicyEngine marks = DirectEngine();
        for (int index = 1; index < 500; index++) marks.Marks.Add(Mark("mark_" + index, "prompt", 9, 13, Text));
        marks.ValidateState(Text);
        Denied(() => marks.AddMark(Mark("extra", "prompt", 9, 13, Text), Text), "Mark count bound.");
        Equal(true, marks.RemoveMark("mark_499"), "Local mark removal.");
        marks.AddMark(Mark("replacement", "prompt", 9, 13, Text), Text);
        Equal(500, marks.Marks.Count, "Capacity restored by explicit mark removal.");

        PolicyEngine pending = DirectEngine();
        pending.Rules[0].Action = "discard";
        pending.Rules[0].AccessMode = "suggestions";
        pending.Marks[0].From = 9; pending.Marks[0].Quote = "beta";
        string token = Read(pending, Text);
        Proposal template = pending.Submit(Request("template", 9, 13, "beta", ""), token, Text, Snapshot).Proposal;
        for (int index = 1; index < 500; index++)
        {
            Proposal proposal = Json.Deserialize<Proposal>(Json.Serialize(template));
            proposal.Id = "pending_" + index;
            pending.Proposals.Add(proposal);
        }
        Denied(() => pending.Submit(Request("too_many", 9, 13, "beta", ""), token, Text, Snapshot), "Pending discard count bound remains enforced.");
        Equal(500, pending.Proposals.Count, "Full pending queue remains unchanged.");
        PolicyState corrupted = pending.ExportState();
        corrupted.Proposals[0].PolicyHash = "forged";
        Denied(() => new PolicyEngine(contract, corrupted), "Corrupted binding is preserved by refusal.");
        corrupted = pending.ExportState();
        corrupted.Proposals[0].Status = "model_says_applied";
        Denied(() => new PolicyEngine(contract, corrupted), "Unknown persisted status.");
        corrupted = pending.ExportState();
        corrupted.Marks.Add(Json.Deserialize<PassageMark>(Json.Serialize(corrupted.Marks[0])));
        Denied(() => new PolicyEngine(contract, corrupted), "Duplicate persisted markers.");
        corrupted = pending.ExportState();
        corrupted.Version = 99;
        Denied(() => new PolicyEngine(contract, corrupted), "Future metadata version is never reset.");

        PolicyEngine receipts = DirectEngine();
        token = Read(receipts, Text);
        Proposal receipt = receipts.Submit(Request("receipt", 9, 13, "beta", "new"), token, Text, Snapshot).Proposal;
        receipts.Proposals[0].Status = "rejected";
        for (int index = 1; index < PolicyEngine.MaxReceipts; index++)
        {
            Proposal previous = Json.Deserialize<Proposal>(Json.Serialize(receipt));
            previous.Id = "receipt_" + index;
            previous.Status = "rejected";
            receipts.Proposals.Add(previous);
        }
        Denied(() => receipts.Submit(Request("extra_receipt", 9, 13, "beta", "new"), token, Text, Snapshot), "Receipt history is bounded without evicting replay protection.");
        Equal(PolicyEngine.MaxReceipts, receipts.Proposals.Count, "Full receipt history remains unchanged.");
    }

    private static PolicyEngine DirectEngine()
    {
        PolicyEngine engine = new PolicyEngine(contract);
        engine.Rules.Clear();
        engine.Rules.Add(Rule("prompt", "prompt", "direct"));
        engine.Marks.Add(Mark("editable", "prompt", 0, 13, Text));
        return engine;
    }

    private static void FixedWordEditing()
    {
        PolicyEngine legacy = SuggestionEngine();
        string rules = Json.Serialize(legacy.Rules);
        legacy = new PolicyEngine(contract, legacy.ExportState());
        Equal(rules, Json.Serialize(legacy.Rules), "Opening a legacy document does not migrate its stored rule definitions.");
        string token = Read(legacy, Text);
        Proposal request = Request("legacy_edit_only", 9, 13, "beta", "clear beta");
        Authorization authorized = legacy.Submit(request, token, Text, Snapshot);
        Equal("direct", authorized.Mode, "Legacy prompt rules use Word's fixed direct editing behavior.");
        Equal("authorized", authorized.Proposal.Status, "A legacy prompt prepares a direct edit without a new approval step.");
        Equal(rules, Json.Serialize(legacy.Rules), "Effective direct behavior does not rewrite stored color rules.");
        Equal("direct", legacy.ValidateDirect(request.Id, DocumentId, Revision, Text, Snapshot).Mode, "Legacy direct edits retain the native witness recheck.");
        legacy.BeginApplication(request.Id);
        legacy.RecordApplied(request.Id);
        Authorization replay = legacy.Submit(request, token, Text.Replace("beta", "clear beta"), "updated-snapshot");
        Equal(true, replay.AlreadyRecorded, "A completed edit-only request replays its receipt after the source changes.");
        Equal("applied", replay.Proposal.Status, "The replay cannot reapply a completed direct edit.");
        Equal(rules, Json.Serialize(legacy.Rules), "Direct application and replay retain the original local rule definitions.");

        PolicyEngine fresh = new PolicyEngine(contract);
        Equal(true, fresh.Rules.Any(rule => rule.Action == "prompt"), "Word retains an existing default wording rule.");
        Equal(true, fresh.Rules.Where(rule => rule.Action == "prompt").All(rule => rule.AccessMode == "direct"), "New Word wording rules default to direct editing.");
        Equal(true, fresh.Rules.Where(rule => rule.Action != "prompt").All(rule => rule.AccessMode == "suggestions"), "Nonwording functions retain their compatible restrictive wire value.");
    }

    private static PolicyEngine SuggestionEngine()
    {
        PolicyEngine engine = DirectEngine();
        engine.Rules[0].AccessMode = "suggestions";
        return engine;
    }

    private static ColorRule Rule(string id, string action, string access)
    {
        return new ColorRule { Id = id, Label = id, Color = "#e5be50", Action = action, Prompt = action == "neutral" ? "" : "Keep meaning.", AccessMode = access };
    }

    private static PassageMark Mark(string id, string rule, int from, int to, string text)
    {
        return new PassageMark { Id = id, RuleId = rule, From = from, To = to, Quote = PolicyEngine.Slice(text, from, to), Prompt = "", Stale = false };
    }

    private static Proposal Request(string id, int from, int to, string quote, string replacement)
    {
        return new Proposal { Id = id, DocumentId = DocumentId, Revision = Revision, From = from, To = to, Quote = quote, Replacement = replacement, Rationale = "Clearer wording." };
    }

    private static string Read(PolicyEngine engine, string text)
    {
        engine.Share(DocumentId);
        return engine.ReadContext(DocumentId, Revision, text, Snapshot);
    }

    private static void Equal<T>(T expected, T actual, string message)
    {
        assertions++;
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new Exception(message + " Expected " + expected + ", got " + actual + ".");
    }

    private static void Denied(Action action, string message)
    {
        assertions++;
        try { action(); }
        catch (InvalidOperationException) { return; }
        throw new Exception("Expected policy refusal: " + message);
    }
}
