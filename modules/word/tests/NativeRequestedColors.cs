using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using UnderTree.WordAddin;

internal static partial class NativeHighlightingSmoke
{
    private static void RequestedColors(dynamic application, WordEditor editor, string address, string output)
    {
        dynamic document = application.Documents.Add();
        string citations = String.Join(" ", Enumerable.Range(1, 9).Select(index => "Reference [" + index + "]."));
        string tokens = String.Join(" ", Enumerable.Range(0, 65).Select(index => "token" + index.ToString("D2")));
        document.Content.Text = citations + "\rProtect passage. Reject passage. Discard passage. Suggest passage. Direct passage.\r" +
            "Fresh café 😀 text. Local assignment. Ambiguous identity. Another target.\rAcross first\rAcross second\r" + tokens + "\rField target. Link target.\r";
        document.Activate();
        string path = Path.Combine(output, "requested-colors.docx");
        document.SaveAs2(path, 16);
        string id = UiState.Text(UiState.Map(Local(editor, "state")["document"]), "id");
        var configured = new[] {
            new { Id = "protect", Color = "#74B9A4", Action = "protect", Mode = "suggestions", Text = "Protect passage." },
            new { Id = "color_reject", Color = "#BC5266", Action = "reject", Mode = "suggestions", Text = "Reject passage." },
            new { Id = "color_discard", Color = "#DEAC63", Action = "discard", Mode = "suggestions", Text = "Discard passage." },
            new { Id = "color_suggest", Color = "#7FA7D9", Action = "prompt", Mode = "suggestions", Text = "Suggest passage." },
            new { Id = "color_direct", Color = "#B584CE", Action = "prompt", Mode = "direct", Text = "Direct passage." }
        };
        foreach (var rule in configured)
            Local(editor, "save_rule", new { documentId = id, id = rule.Id, label = rule.Id, color = rule.Color, action = rule.Action,
                prompt = rule.Action == "prompt" ? "Improve clarity." : "", accessMode = rule.Mode });
        string definitions = McpEndpoint.Json().Serialize(Local(editor, "state")["rules"]);
        string original = WordSnapshot.Capture(document).MainText;
        WordSnapshot inspected = WordSnapshot.Capture(document);
        inspected.RequireProse(document, 0, 3, new HashSet<string>(), null, true);
        document.Bookmarks.Add("_HiddenInspectionTest", document.Range(1, 2));
        Check(!inspected.SameProseInspection(WordSnapshot.Capture(document)), "A hidden bookmark change invalidates a cached batch inspection even when prose is unchanged.");
        document.Bookmarks.Item("_HiddenInspectionTest").Delete();
        string diskBefore = FileDigest(path);
        var times = new List<object>();
        object[] citationRanges = Enumerable.Range(1, 9).Select(index => ExactColorRange(original, "[" + index + "]")).ToArray();
        var context = Share(editor, address, id);
        object citationRequest = ColorRequest(id, "nine_citations_green", context, "#74b9a4", null, citationRanges);
        var timer = Stopwatch.StartNew();
        var receipt = Tool(address, "highlight_text", citationRequest);
        times.Add(new { operation = "nine existing-green citations", milliseconds = timer.ElapsedMilliseconds });
        Check(UiState.Text(receipt, "colorId") == "protect" && UiState.Text(receipt, "action") == "protect", "Nine citations reuse the existing green identity and truthfully report protection.");
        foreach (int citation in Enumerable.Range(1, 9))
        {
            string quote = "[" + citation + "]";
            int from = original.IndexOf(quote, StringComparison.Ordinal);
            Check(NativePaint(document, from, from + quote.Length) == ColorNumber("#74B9A4"), "Citation " + citation + " displays the requested green in Word.");
        }
        Check(UiState.Records(Local(editor, "state"), "marks").Count == 9, "A single native batch creates nine independent citation markers.");
        Tool(address, "highlight_text", citationRequest);
        Check(UiState.Records(Local(editor, "state"), "marks").Count == 9, "An exact native replay cannot duplicate citation markers.");
        foreach (var rule in configured)
        {
            context = Share(editor, address, id);
            receipt = Tool(address, "highlight_text", ColorRequest(id, "extend_" + rule.Id, context, null, rule.Id, new[] { ExactColorRange(original, rule.Text) }));
            int from = original.IndexOf(rule.Text, StringComparison.Ordinal);
            Check(UiState.Text(receipt, "action") == rule.Action && NativePaint(document, from, from + rule.Text.Length) == ColorNumber(rule.Color),
                "Existing " + rule.Action + "/" + rule.Mode + " paints its exact text and reports its unchanged function.");
        }
        Check(definitions == McpEndpoint.Json().Serialize(Local(editor, "state")["rules"]), "Extending every existing function leaves all rule definitions exact.");
        Check(WordSnapshot.Capture(document).MainText == original && (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0,
            "Coloring prompt/direct/discard passages never executes wording edits, deletions or review changes.");
        context = Share(editor, address, id);
        receipt = Tool(address, "highlight_text", ColorRequest(id, "fresh_unicode", context, "#123ABC", null, new[] { ExactColorRange(original, "Fresh café 😀 text.") }));
        string freshId = UiState.Text(receipt, "colorId");
        Check(UiState.Text(receipt, "color") == "#123ABC" && UiState.Text(receipt, "action") == "neutral", "A requested unused RGB paints Unicode text and remains functionless.");
        int unicodeFrom = original.IndexOf("Fresh", StringComparison.Ordinal);
        var choices = Local(editor, "selection_rules", new { documentId = id, nativeFrom = unicodeFrom + 7, nativeTo = unicodeFrom + 7 });
        Check(UiState.Records(choices, "choices").Single()["ruleId"].ToString() == freshId, "The double-click navigation route resolves the exact new neutral color identity.");
        context = Share(editor, address, id);
        var stale = ColorRequest(id, "local_assignment_race", context, null, freshId, new[] { ExactColorRange(original, "Local assignment.") });
        Local(editor, "save_rule", new { documentId = id, id = freshId, label = "Local prompt assignment", color = "#123ABC", action = "prompt", prompt = "Keep concise.", accessMode = "suggestions" });
        DeniedTool(address, "highlight_text", stale);
        context = Share(editor, address, id);
        Tool(address, "highlight_text", ColorRequest(id, "after_local_assignment", context, null, freshId, new[] { ExactColorRange(original, "Local assignment.") }));
        Check(UiState.Records(Local(editor, "state"), "marks").Count(mark => mark["ruleId"].ToString() == freshId) == 2,
            "A fresh read extends the locally assigned prompt through the exact existing color.");
        Local(editor, "save_rule", new { documentId = id, id = "same_rgb_distinct", label = "Distinct identity", color = "#123ABC", action = "protect", prompt = "", accessMode = "suggestions" });
        AssertColorRefusal(editor, address, document, id, "ambiguous_rgb", "#123ABC", null, new[] { ExactColorRange(original, "Ambiguous identity.") }, "Several color identities");
        context = Share(editor, address, id);
        Tool(address, "highlight_text", ColorRequest(id, "disambiguated_identity", context, null, "same_rgb_distinct", new[] { ExactColorRange(original, "Ambiguous identity.") }));
        AssertColorRefusal(editor, address, document, id, "hidden_neutral", "#FEDCBA", null, new[] { ExactColorRange(original, "Suggest passage.") }, "overlaps marker");
        AssertColorRefusal(editor, address, document, id, "protect_recolor", null, "color_reject", new[] { ExactColorRange(original, "Protect passage.") }, "overlaps marker");
        AssertColorRefusal(editor, address, document, id, "reject_recolor", "#FF0102", null, new[] { ExactColorRange(original, "Reject passage.") }, "overlaps marker");
        context = Share(editor, address, id);
        int crossFrom = original.IndexOf("Across first", StringComparison.Ordinal), crossTo = original.IndexOf("Across second", StringComparison.Ordinal) + "Across second".Length;
        receipt = Tool(address, "highlight_text", ColorRequest(id, "paragraph_crossing", context, "#116677", null,
            new object[] { new { from = crossFrom, to = crossTo, quote = original.Substring(crossFrom, crossTo - crossFrom) } }));
        Check(UiState.Text(receipt, "action") == "neutral" && UiState.Records(Local(editor, "state"), "marks").All(mark => !UiState.Flag(mark, "stale")),
            "A multiline visual mark remains current across native paragraph boundaries.");
        var batch = Enumerable.Range(0, 65).Select(index => ExactColorRange(original, "token" + index.ToString("D2"))).ToArray();
        AssertColorRefusal(editor, address, document, id, "over_bound", "#665544", null, batch, null);
        context = Share(editor, address, id);
        timer.Restart();
        Tool(address, "highlight_text", ColorRequest(id, "max_batch", context, "#665544", null, batch.Take(64).ToArray()));
        times.Add(new { operation = "64-range maximum batch", milliseconds = timer.ElapsedMilliseconds });
        Check(WordSnapshot.Capture(document).MainText == original && (int)document.Revisions.Count == 0 && FileDigest(path) == diskBefore,
            "The maximum batch changes no prose or revisions and does not invoke Save.");
        string selectedRules = McpEndpoint.Json().Serialize(Local(editor, "state")["rules"]);
        document.Save(); document.Close(0);
        document = application.Documents.Open(path, ReadOnly: false, Visible: false); document.Activate();
        var reopened = Local(editor, "state");
        Check(UiState.Records(reopened, "marks").Count == 82 && UiState.Records(reopened, "marks").All(mark => !UiState.Flag(mark, "stale")),
            "DOCX save/reopen retains all 82 exact color markers without stale geometry.");
        Check(McpEndpoint.Json().Serialize(reopened["rules"]) == selectedRules && WordSnapshot.Capture(document).MainText == original,
            "Save/reopen retains exact local definitions, selected identities and all original prose.");
        InheritedProtectedPaint(application, editor, address, output);
        document.Activate();
        context = Share(editor, address, id);
        document.Close(0);
        DeniedTool(address, "highlight_text", ColorRequest(id, "closed_document", context, "#112244", null, new[] { ExactColorRange(original, "Another target.") }));
        File.WriteAllText(Path.Combine(output, "color-timing.json"), McpEndpoint.Json().Serialize(times));
        File.WriteAllText(Path.Combine(output, "color-result.txt"), "PASS " + assertions + " native requested-color assertions.");
        Console.WriteLine("PASS " + assertions + " native requested-color assertions.");
    }

    private static void InheritedProtectedPaint(dynamic application, WordEditor editor, string address, string output)
    {
        dynamic document = application.Documents.Add();
        const string quote = "Inherited citation formatting.";
        document.Content.Text = quote + "\r";
        document.SaveAs2(Path.Combine(output, "inherited-protection.docx"), 16);
        string id = UiState.Text(UiState.Map(Local(editor, "state")["document"]), "id");
        Mark(editor, document, id, "protect", 0, quote.Length);
        document.Range(0, quote.Length).Font.Shading.BackgroundPatternColor = -16777216;
        document.Range(0, quote.Length).ParagraphFormat.Shading.BackgroundPatternColor = ColorNumber("#74B9A4");
        var state = Local(editor, "state");
        Check(!UiState.Flag(UiState.Records(state, "marks").Single(), "stale"), "Paragraph-inherited green remains a valid visible protected mark.");
        string formatting = WordSnapshot.FormattingSignature(document.Range(0, quote.Length));
        var context = Share(editor, address, id);
        Tool(address, "highlight_text", ColorRequest(id, "inherited_same_rgb", context, "#74B9A4", null, new[] { ExactColorRange(quote, quote) }));
        Check(WordSnapshot.FormattingSignature(document.Range(0, quote.Length)) == formatting &&
            (int)document.Range(0, quote.Length).Font.Shading.BackgroundPatternColor == -16777216,
            "Applying the same green over protection never replaces inherited shading with direct formatting.");
        var choices = Local(editor, "selection_rules", new { documentId = id, nativeFrom = 5, nativeTo = 5 });
        Check(UiState.Records(choices, "choices").Count == 2 && UiState.Records(choices, "choices").Select(choice => choice["id"].ToString()).Distinct().Count() == 2,
            "Overlapping same-color marks remain separate native identities in the local navigation route.");
        document.Close(0);
    }

    private static object ExactColorRange(string text, string quote)
    {
        int from = text.IndexOf(quote, StringComparison.Ordinal);
        if (from < 0) throw new Exception("Missing fixture phrase: " + quote);
        return new { from = from, to = from + quote.Length, quote = quote };
    }

    private static object ColorRequest(string id, string requestId, IDictionary<string, object> context, string color, string colorId, object[] ranges)
    {
        var marking = new Dictionary<string, object> { { "requestId", requestId }, { "revision", context["revision"] }, { "contextId", context["context_id"] }, { "ranges", ranges } };
        if (color != null) marking["color"] = color;
        if (colorId != null) marking["colorId"] = colorId;
        return new { documentId = id, marking = marking };
    }

    private static void AssertColorRefusal(WordEditor editor, string address, dynamic document, string id, string requestId, string color, string colorId, object[] ranges, string reason)
    {
        var context = Share(editor, address, id);
        var before = Local(editor, "state");
        string metadata = McpEndpoint.Json().Serialize(new { rules = before["rules"], marks = before["marks"] });
        string native = new WordNativeEditState().NativeIdentity(document);
        var response = Rpc(address, "highlight_text", ColorRequest(id, requestId, context, color, colorId, ranges));
        string message = UiState.Text(UiState.Map(((IEnumerable)response["content"]).Cast<object>().First()), "text");
        var after = Local(editor, "state");
        Check(UiState.Flag(response, "isError") && (reason == null || message.Contains(reason)), "Native " + requestId + " returns its explicit refusal diagnostic.");
        Check(metadata == McpEndpoint.Json().Serialize(new { rules = after["rules"], marks = after["marks"] }) &&
            native == new WordNativeEditState().NativeIdentity(document), "Native " + requestId + " leaves all metadata, bookmarks, text and formatting exact.");
    }

    private static int ColorNumber(string color)
    { return Convert.ToInt32(color.Substring(1, 2), 16) | Convert.ToInt32(color.Substring(3, 2), 16) << 8 | Convert.ToInt32(color.Substring(5, 2), 16) << 16; }
    private static int NativePaint(dynamic document, int from, int to)
    { return (int)document.Range(from, to).Font.Shading.BackgroundPatternColor; }
}
