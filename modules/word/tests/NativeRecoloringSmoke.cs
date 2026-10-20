using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using UnderTree.WordAddin;

internal static partial class NativeHighlightingSmoke
{
    private static void Recoloring(dynamic application, WordEditor editor, string address, string output)
    {
        int initialAssertions = assertions;
        var evidence = new List<object>();
        dynamic document = null;
        string path = Path.Combine(output, "four-citation-recoloring.docx");
        string id = null;
        try
        {
            ExactNeutralRedRecolor(application, editor, address, output, evidence);
            InheritedSameRgbRecolor(application, editor, address, output, evidence);
            ExposedColorRecolor(application, editor, address, output, evidence);
            OverlappingRecolor(application, editor, address, output, evidence);
            MixedInheritedSameRgbRecolor(application, editor, address, output, evidence);
            SingleMarkerMixedColor(application, editor, address, output, evidence);
            ColorVerificationBudget(application, editor, address, output, evidence);
            document = application.Documents.Add();
            document.Content.Text = "Evidence [1], [2], [3], [4]; neighbor [5].\rUnmarked café text.\r";
            document.Content.Font.Name = "Calibri"; document.Content.Font.Size = 11;
            document.Activate(); document.SaveAs2(path, 16);
            id = UiState.Text(UiState.Map(Local(editor, "state")["document"]), "id");
            Check(!File.Exists(RecolorSidecar(id)), "The recoloring fixture has no pre-existing private review receipt.");
            Local(editor, "save_rule", new { documentId = id, id = "protect", label = "Preserve text", color = "#74B9A4",
                action = "protect", prompt = "", accessMode = "suggestions" });
            Local(editor, "save_rule", new { documentId = id, id = "recolor_red", label = "Reject wording", color = "#E57575",
                action = "reject", prompt = "", accessMode = "suggestions" });
            string original = WordSnapshot.Capture(document).FullText;
            object[] ranges = Enumerable.Range(1, 4).Select(number => ExactColorRange(original, "[" + number + "]")).ToArray();
            for (int index = 0; index < ranges.Length; index++)
            {
                var range = Map(ranges[index]);
                dynamic font = document.Range((int)range["from"], (int)range["to"]).Font;
                font.Bold = index % 2 == 0 ? -1 : 0; font.Italic = index % 2 == 1 ? -1 : 0;
                font.Size = 10 + index; font.Color = index == 2 ? ColorNumber("#163D75") : 0;
            }
            var context = Share(editor, address, id);
            var seed = Tool(address, "highlight_text", ColorRequest(id, "seed_five_green", context, null, "protect",
                ranges.Concat(new[] { ExactColorRange(original, "[5]") }).ToArray()));
            string[] greenIds = RecolorIds(seed, "markIds").Take(4).ToArray();
            var before = Local(editor, "state");
            var neighbor = UiState.Records(before, "marks").Single(mark => UiState.Text(mark, "quote") == "[5]");
            string neighborId = UiState.Text(neighbor, "id"), neighborState = McpEndpoint.Json().Serialize(neighbor);
            int neighborFrom = (int)neighbor["from"], neighborTo = (int)neighbor["to"];
            string neighborFormatting = WordSnapshot.FormattingSignature(document.Range(neighborFrom, neighborTo));
            string fonts = RecolorFonts(document), rules = McpEndpoint.Json().Serialize(before["rules"]);
            document.Save();
            string disk = FileDigest(path);
            evidence.Add(new { operation = "green_baseline", fullText = original, fontSignature = fonts, rules = before["rules"],
                marks = before["marks"], neighborFormatting = neighborFormatting, savedDocumentSha256 = disk });
            context = Share(editor, address, id);
            object request = RecolorRequest(id, "four_green_to_red", context, null, "recolor_red", ranges, greenIds);
            var timer = Stopwatch.StartNew();
            var receipt = Tool(address, "highlight_text", request);
            evidence.Add(new { operation = "four_green_to_existing_red", milliseconds = timer.ElapsedMilliseconds, request = request, receipt = receipt });
            string[] redIds = RecolorIds(receipt, "markIds");
            Check(RecolorIds(receipt, "replacedMarkIds").SequenceEqual(greenIds), "The native red receipt lists the four retired green IDs in exact range order.");
            Check(UiState.Text(receipt, "colorId") == "recolor_red" && UiState.Text(receipt, "action") == "reject" &&
                redIds.Length == 4 && redIds.Distinct().Count() == 4 && !redIds.Intersect(greenIds).Any(),
                "Four green citations become four distinct markers bound to the existing local red Reject rule.");
            AssertRecolorBindings(editor, document, redIds, greenIds, ranges, "recolor_red", "#E57575", "Red recoloring");
            before = Local(editor, "state");
            Check(McpEndpoint.Json().Serialize(before["rules"]) == rules && WordSnapshot.Capture(document).FullText == original &&
                RecolorFonts(document) == fonts && (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0,
                "Recoloring preserves every prose character, font, local rule and native review annotation count.");
            Check(McpEndpoint.Json().Serialize(UiState.Records(before, "marks").Single(mark => UiState.Text(mark, "id") == neighborId)) == neighborState &&
                WordSnapshot.FormattingSignature(document.Range(neighborFrom, neighborTo)) == neighborFormatting,
                "The neighboring green citation retains its exact marker, formatting and protection.");
            Check(FileDigest(path) == disk && !(bool)document.Saved, "Remote recoloring leaves the saved DOCX untouched and the changed Word document unsaved.");
            string nativeRed = new WordNativeEditState().NativeIdentity(document), stateRed = RecolorMetadata(before), sidecarRed = FileDigest(RecolorSidecar(id));
            var replay = Tool(address, "highlight_text", request);
            Check(UiState.Flag(replay, "replayed") && RecolorIds(replay, "markIds").SequenceEqual(redIds) &&
                RecolorIds(replay, "replacedMarkIds").SequenceEqual(greenIds), "A lost native recolor response replays exactly the same replacement identities.");
            Check(nativeRed == new WordNativeEditState().NativeIdentity(document) && stateRed == RecolorMetadata(Local(editor, "state")) &&
                sidecarRed == FileDigest(RecolorSidecar(id)), "Recolor replay duplicates no bookmark, paint, policy mark or receipt.");
            Check((bool)document.Undo(1), "One native Undo reverses the complete four-citation recoloring transaction.");
            AssertRecolorBindings(editor, document, greenIds, redIds, ranges, "protect", "#74B9A4", "Single Undo");
            Check(!UiState.Flag(UiState.Map(Local(editor, "state")["document"]), "shared"), "Undo revokes previously read recoloring authority.");
            Check((bool)document.Redo(1), "One native Redo reapplies the complete four-citation recoloring transaction.");
            AssertRecolorBindings(editor, document, redIds, greenIds, ranges, "recolor_red", "#E57575", "Single Redo");
            Check(WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts &&
                McpEndpoint.Json().Serialize(Local(editor, "state")["rules"]) == rules, "Undo and Redo retain all prose, fonts and current local rule definitions.");
            document.Save(); document.Close(0); document = null;
            document = application.Documents.Open(path, ReadOnly: false, Visible: false); document.Activate();
            AssertRecolorBindings(editor, document, redIds, greenIds, ranges, "recolor_red", "#E57575", "DOCX save/reopen");
            Check(!UiState.Flag(UiState.Map(Local(editor, "state")["document"]), "shared") &&
                WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts,
                "Reopening preserves exact recoloring and fonts without restoring the old sharing grant.");

            RecolorRefusals(editor, address, document, id, ranges, redIds, greenIds, evidence);
            context = Share(editor, address, id);
            request = RecolorRequest(id, "red_to_unused_neutral", context, "#1267AB", null, ranges, redIds);
            timer.Restart(); receipt = Tool(address, "highlight_text", request);
            evidence.Add(new { operation = "four_red_to_unused_neutral", milliseconds = timer.ElapsedMilliseconds, request = request, receipt = receipt });
            string neutralId = UiState.Text(receipt, "colorId");
            string[] neutralIds = RecolorIds(receipt, "markIds");
            Check(UiState.Text(receipt, "action") == "neutral" && UiState.Text(receipt, "color") == "#1267AB" &&
                RecolorIds(receipt, "replacedMarkIds").SequenceEqual(redIds), "A fresh recolor RGB creates only a neutral rule and retires every selected red identity.");
            AssertRecolorBindings(editor, document, neutralIds, greenIds.Concat(redIds).ToArray(), ranges, neutralId, "#1267AB", "Fresh neutral recoloring");
            before = Local(editor, "state");
            Check(McpEndpoint.Json().Serialize(UiState.Records(before, "rules").Where(rule => UiState.Text(rule, "id") != neutralId).ToArray()) == rules &&
                UiState.Records(before, "rules").Single(rule => UiState.Text(rule, "id") == neutralId)["prompt"].ToString() == "",
                "The fresh neutral destination leaves every pre-existing local rule and prompt exact.");
            context = Share(editor, address, id);
            var neutralHighlights = UiState.Records(context, "highlights").Where(mark => neutralIds.Contains(UiState.Text(mark, "id"))).ToArray();
            Check(neutralHighlights.Length == 4 && neutralHighlights.All(mark => mark["protectedId"] == null) &&
                UiState.Records(context, "highlights").Single(mark => UiState.Text(mark, "id") == neighborId)["protectedId"].ToString() == neighborId,
                "The four neutral citations carry no hidden old protection, while the unselected green neighbor remains protected.");
            Check(WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts &&
                WordSnapshot.FormattingSignature(document.Range(neighborFrom, neighborTo)) == neighborFormatting,
                "Fresh neutral replacement preserves all prose, fonts and neighboring native formatting.");
            RecolorRuleUndo(editor, address, document, id, neutralId, neutralIds, ranges, context, evidence);
            document.Save(); document.Close(0); document = null;
            document = application.Documents.Open(path, ReadOnly: false, Visible: false); document.Activate();
            before = Local(editor, "state");
            Check(UiState.Records(before, "rules").Single(rule => UiState.Text(rule, "id") == neutralId)["action"].ToString() == "protect" &&
                WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts,
                "Final save/reopen retains the newer local protection rather than the historical neutral authority.");
            evidence.Add(new { operation = "final_state", state = before, nativeIdentity = new WordNativeEditState().NativeIdentity(document),
                documentSha256 = FileDigest(path), assertions = assertions - initialAssertions });
            File.WriteAllText(Path.Combine(output, "recoloring-result.json"), McpEndpoint.Json().Serialize(new {
                status = "PASS", assertions = assertions - initialAssertions, events = evidence, installedPackage = "NOT RUN", publicClient = "NOT RUN" }));
            Console.WriteLine("PASS " + (assertions - initialAssertions) + " native recoloring assertions.");
        }
        finally
        {
            File.WriteAllText(Path.Combine(output, "recoloring-events.json"), McpEndpoint.Json().Serialize(evidence));
            if (document != null) document.Close(0);
            if (id != null && File.Exists(RecolorSidecar(id))) File.Delete(RecolorSidecar(id));
        }
    }

    private static void ExactNeutralRedRecolor(dynamic application, WordEditor editor, string address, string output, List<object> evidence)
    {
        dynamic document = application.Documents.Add();
        string id = null, path = Path.Combine(output, "exact-neutral-red.docx");
        try
        {
            document.Content.Text = "Alpha beta gamma delta.\r";
            document.Activate(); document.SaveAs2(path, 16);
            id = UiState.Text(UiState.Map(Local(editor, "state")["document"]), "id");
            Local(editor, "save_rule", new { documentId = id, id = "protect", label = "Green protection", color = "#00FF00",
                action = "protect", prompt = "", accessMode = "suggestions" });
            Local(editor, "save_rule", new { documentId = id, id = "neutral_red", label = "Red color only", color = "#FF0000",
                action = "neutral", prompt = "", accessMode = "suggestions" });
            string original = WordSnapshot.Capture(document).FullText;
            object[] ranges = new[] { ExactColorRange(original, "beta"), ExactColorRange(original, "delta") };
            var context = Share(editor, address, id);
            var seed = Tool(address, "highlight_text", ColorRequest(id, "exact_green_seed", context, "#00FF00", null, ranges));
            string[] greenIds = RecolorIds(seed, "markIds");
            var state = Local(editor, "state");
            string rules = McpEndpoint.Json().Serialize(state["rules"]), fonts = RecolorFonts(document);
            string neighbor = McpEndpoint.Json().Serialize(UiState.Records(state, "marks").Single(mark => UiState.Text(mark, "id") == greenIds[1]));
            string neighborPaint = WordSnapshot.FormattingSignature(document.Range(17, 22));
            context = Share(editor, address, id);
            var request = RecolorRequest(id, "exact_green_to_neutral_red", context, "#FF0000", null, ranges.Take(1).ToArray(), greenIds.Take(1).ToArray());
            var receipt = Tool(address, "highlight_text", request);
            string[] redIds = RecolorIds(receipt, "markIds");
            Check(UiState.Text(receipt, "colorId") == "neutral_red" && UiState.Text(receipt, "action") == "neutral" &&
                UiState.Text(receipt, "color") == "#FF0000", "Exact green-to-red uses the existing neutral red without an editing prompt.");
            AssertRecolorBindings(editor, document, redIds, greenIds.Take(1).ToArray(), ranges.Take(1).ToArray(), "neutral_red", "#FF0000", "Exact neutral red", 2);
            var replay = Tool(address, "highlight_text", request);
            Check(UiState.Flag(replay, "replayed") && RecolorIds(replay, "markIds").SequenceEqual(redIds), "Neutral red replay creates no duplicate marker.");
            document.Save(); document.Close(0); document = null;
            document = application.Documents.Open(path, ReadOnly: false, Visible: false); document.Activate();
            state = Local(editor, "state");
            AssertRecolorBindings(editor, document, redIds, greenIds.Take(1).ToArray(), ranges.Take(1).ToArray(), "neutral_red", "#FF0000", "Reopened neutral red", 2);
            Check(McpEndpoint.Json().Serialize(state["rules"]) == rules && UiState.Text(UiState.Records(state, "rules").Single(rule => UiState.Text(rule, "id") == "neutral_red"), "prompt") == "" &&
                WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts &&
                McpEndpoint.Json().Serialize(UiState.Records(state, "marks").Single(mark => UiState.Text(mark, "id") == greenIds[1])) == neighbor &&
                WordSnapshot.FormattingSignature(document.Range(17, 22)) == neighborPaint && (int)document.Comments.Count == 0 && (int)document.Revisions.Count == 0,
                "Save/reopen preserves neutral red, all rules, exact text, fonts, other highlights and native annotations.");
            evidence.Add(new { operation = "exact_green_to_neutral_red_reopened", request = request, receipt = receipt, state = state, documentSha256 = FileDigest(path) });
        }
        finally
        {
            if (document != null) document.Close(0);
            if (id != null && File.Exists(RecolorSidecar(id))) File.Delete(RecolorSidecar(id));
        }
    }

    private static void InheritedSameRgbRecolor(dynamic application, WordEditor editor, string address, string output, List<object> evidence)
    {
        dynamic document = application.Documents.Add();
        string id = null;
        try
        {
            document.Content.Text = "Inherited [1] keeps exact formatting.\r";
            document.Content.Font.Name = "Calibri"; document.Content.Font.Size = 11;
            document.Activate();
            string path = Path.Combine(output, "inherited-same-rgb-recoloring.docx");
            document.SaveAs2(path, 16);
            id = UiState.Text(UiState.Map(Local(editor, "state")["document"]), "id");
            Check(!File.Exists(RecolorSidecar(id)), "The inherited recoloring fixture has no prior private receipt.");
            Local(editor, "save_rule", new { documentId = id, id = "protect", label = "Original green", color = "#74B9A4",
                action = "protect", prompt = "", accessMode = "suggestions" });
            Local(editor, "save_rule", new { documentId = id, id = "inherited_green_target", label = "Other green identity", color = "#74B9A4",
                action = "protect", prompt = "", accessMode = "suggestions" });
            string original = WordSnapshot.Capture(document).FullText;
            object[] ranges = new[] { ExactColorRange(original, "[1]") };
            var range = Map(ranges[0]);
            int from = (int)range["from"], to = (int)range["to"];
            document.Range(from, to).Font.Italic = -1;
            var context = Share(editor, address, id);
            var seed = Tool(address, "highlight_text", ColorRequest(id, "seed_inherited_green", context, null, "protect", ranges));
            string oldId = RecolorIds(seed, "markIds").Single();
            document.Range(from, to).Font.Shading.BackgroundPatternColor = -16777216;
            document.Range(from, to).ParagraphFormat.Shading.BackgroundPatternColor = ColorNumber("#74B9A4");
            var state = Local(editor, "state");
            Check(UiState.Records(state, "marks").Count == 1 && !UiState.Flag(UiState.Records(state, "marks").Single(), "stale") &&
                NativePaint(document, from, to) == -16777216 &&
                (int)document.Range(from, to).ParagraphFormat.Shading.BackgroundPatternColor == ColorNumber("#74B9A4"),
                "A green source inherited from paragraph shading is current with automatic character shading.");
            string formatting = WordSnapshot.FormattingSignature(document.Range(from, to));
            string fonts = RecolorFonts(document), rules = McpEndpoint.Json().Serialize(state["rules"]);
            document.Save(); string disk = FileDigest(path);
            context = Share(editor, address, id);
            object request = RecolorRequest(id, "inherited_green_same_rgb", context, null, "inherited_green_target", ranges, new[] { oldId });
            var timer = Stopwatch.StartNew();
            var receipt = Tool(address, "highlight_text", request);
            state = Local(editor, "state");
            string newId = RecolorIds(receipt, "markIds").Single();
            evidence.Add(new { operation = "inherited_same_rgb_recolor", milliseconds = timer.ElapsedMilliseconds, request = request, receipt = receipt,
                beforeFormatting = formatting, afterFormatting = WordSnapshot.FormattingSignature(document.Range(from, to)),
                characterShading = NativePaint(document, from, to), paragraphShading = (int)document.Range(from, to).ParagraphFormat.Shading.BackgroundPatternColor });
            Check(WordSnapshot.FormattingSignature(document.Range(from, to)) == formatting && NativePaint(document, from, to) == -16777216 &&
                (int)document.Range(from, to).ParagraphFormat.Shading.BackgroundPatternColor == ColorNumber("#74B9A4"),
                "Explicit same-RGB replacement preserves exact inherited native formatting and automatic character shading.");
            Check(newId != oldId && RecolorIds(receipt, "replacedMarkIds").SequenceEqual(new[] { oldId }) &&
                UiState.Records(state, "marks").Count == 1 && UiState.Text(UiState.Records(state, "marks").Single(), "ruleId") == "inherited_green_target" &&
                !UiState.Flag(UiState.Records(state, "marks").Single(), "stale") && !(bool)document.Bookmarks.Exists(oldId) &&
                (bool)document.Bookmarks.Exists(newId) && (int)document.Bookmarks.Item(newId).Range.Start == from &&
                (int)document.Bookmarks.Item(newId).Range.End == to,
                "Inherited same-color recoloring changes only the exact selected marker identity and local rule binding.");
            Check(WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts &&
                McpEndpoint.Json().Serialize(state["rules"]) == rules && (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0 &&
                FileDigest(path) == disk && !(bool)document.Saved,
                "Inherited recoloring preserves prose, fonts, rules, review annotations and the saved DOCX.");
        }
        finally
        {
            document.Close(0);
            if (id != null && File.Exists(RecolorSidecar(id))) File.Delete(RecolorSidecar(id));
        }
    }

    private static void ExposedColorRecolor(dynamic application, WordEditor editor, string address, string output, List<object> evidence)
    {
        dynamic document = application.Documents.Add();
        string id = null;
        try
        {
            document.Content.Text = "Overlapping [2] keeps its prose.\r";
            document.Content.Font.Name = "Calibri"; document.Content.Font.Size = 11;
            document.Activate(); document.SaveAs2(Path.Combine(output, "exposed-color-recoloring.docx"), 16);
            id = UiState.Text(UiState.Map(Local(editor, "state")["document"]), "id");
            Check(!File.Exists(RecolorSidecar(id)), "The exposed-color fixture has no prior private receipt.");
            Local(editor, "save_rule", new { documentId = id, id = "protect", label = "Protect green", color = "#74B9A4",
                action = "protect", prompt = "", accessMode = "suggestions" });
            Local(editor, "save_rule", new { documentId = id, id = "neutral_green", label = "Neutral same green", color = "#74B9A4",
                action = "neutral", prompt = "", accessMode = "suggestions" });
            Local(editor, "save_rule", new { documentId = id, id = "visible_red_prompt", label = "Red wording prompt", color = "#E57575",
                action = "prompt", prompt = "Keep the meaning.", accessMode = "suggestions" });
            string original = WordSnapshot.Capture(document).FullText;
            object[] ranges = new[] { ExactColorRange(original, "[2]") };
            var range = Map(ranges[0]); int from = (int)range["from"], to = (int)range["to"];
            document.Range(from, to).Font.Bold = -1;
            var context = Share(editor, address, id);
            var seed = Tool(address, "highlight_text", ColorRequest(id, "seed_hidden_neutral", context, null, "neutral_green", ranges));
            string oldId = RecolorIds(seed, "markIds").Single();
            context = Share(editor, address, id);
            var covering = Tool(address, "highlight_text", ColorRequest(id, "cover_neutral_with_prompt", context, null, "visible_red_prompt", ranges));
            string promptId = RecolorIds(covering, "markIds").Single();
            var state = Local(editor, "state");
            Check(UiState.Records(state, "marks").Count == 2 && UiState.Records(state, "marks").All(mark => !UiState.Flag(mark, "stale")) &&
                NativePaint(document, from, to) == ColorNumber("#E57575"),
                "A current neutral-green marker is visibly hidden by a lower-priority-than-Protect red prompt.");
            string promptState = McpEndpoint.Json().Serialize(UiState.Records(state, "marks").Single(mark => UiState.Text(mark, "id") == promptId));
            string rules = McpEndpoint.Json().Serialize(state["rules"]), fonts = RecolorFonts(document);
            context = Share(editor, address, id);
            object request = RecolorRequest(id, "hidden_neutral_to_protect", context, null, "protect", ranges, new[] { oldId });
            var timer = Stopwatch.StartNew();
            var receipt = Tool(address, "highlight_text", request);
            state = Local(editor, "state");
            string newId = RecolorIds(receipt, "markIds").Single();
            evidence.Add(new { operation = "same_old_rgb_requires_actual_green_paint", milliseconds = timer.ElapsedMilliseconds, request = request, receipt = receipt,
                beforeVisibleColor = "#E57575", afterCharacterShading = NativePaint(document, from, to), marks = state["marks"] });
            Check(NativePaint(document, from, to) == ColorNumber("#74B9A4") && UiState.Records(state, "marks").All(mark => !UiState.Flag(mark, "stale")),
                "Recoloring the hidden neutral-green source to Protect-green actually paints green over the visible red prompt.");
            Check(UiState.Text(receipt, "action") == "protect" && UiState.Text(receipt, "colorId") == "protect" && newId != oldId &&
                RecolorIds(receipt, "replacedMarkIds").SequenceEqual(new[] { oldId }) && UiState.Records(state, "marks").Count == 2 &&
                UiState.Text(UiState.Records(state, "marks").Single(mark => UiState.Text(mark, "id") == newId), "ruleId") == "protect" &&
                !(bool)document.Bookmarks.Exists(oldId) && (bool)document.Bookmarks.Exists(newId) &&
                (int)document.Bookmarks.Item(newId).Range.Start == from && (int)document.Bookmarks.Item(newId).Range.End == to,
                "The exposed-color replacement retires only the selected neutral identity and creates exact current protection.");
            Check(promptState == McpEndpoint.Json().Serialize(UiState.Records(state, "marks").Single(mark => UiState.Text(mark, "id") == promptId)) &&
                (bool)document.Bookmarks.Exists(promptId) && McpEndpoint.Json().Serialize(state["rules"]) == rules &&
                WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts &&
                (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0,
                "Painting visible green preserves the overlapping prompt identity, every local rule, fonts, prose and review annotations.");
        }
        finally
        {
            document.Close(0);
            if (id != null && File.Exists(RecolorSidecar(id))) File.Delete(RecolorSidecar(id));
        }
    }

    private static void OverlappingRecolor(dynamic application, WordEditor editor, string address, string output, List<object> evidence)
    {
        dynamic document = application.Documents.Add();
        string id = null;
        try
        {
            document.Content.Text = "Identical [I]. Partial ABCDE. Nested UVWXY. Neighbor [N].\rUnmarked café.\r";
            document.Content.Font.Name = "Calibri"; document.Content.Font.Size = 11;
            document.Activate();
            string path = Path.Combine(output, "overlapping-recoloring.docx");
            document.SaveAs2(path, 16);
            id = UiState.Text(UiState.Map(Local(editor, "state")["document"]), "id");
            Check(!File.Exists(RecolorSidecar(id)), "The overlapping recoloring fixture has no prior private receipt.");
            Local(editor, "save_rule", new { documentId = id, id = "protect", label = "Protect green", color = "#74B9A4",
                action = "protect", prompt = "", accessMode = "suggestions" });
            Local(editor, "save_rule", new { documentId = id, id = "overlap_red", label = "Existing red Reject", color = "#E57575",
                action = "reject", prompt = "", accessMode = "suggestions" });
            string original = WordSnapshot.Capture(document).FullText;
            var partialFont = Map(ExactColorRange(original, "ABCD"));
            document.Range((int)partialFont["from"], (int)partialFont["to"]).Font.Bold = -1;
            var nestedFont = Map(ExactColorRange(original, "VWX"));
            document.Range((int)nestedFont["from"], (int)nestedFont["to"]).Font.Italic = -1;
            document.Range((int)nestedFont["from"], (int)nestedFont["to"]).Font.Size = 13;
            object[] firstBatch = new[] { "[I]", "ABCD", "UVWXY", "[N]" }.Select(quote => ExactColorRange(original, quote)).ToArray();
            object[] secondBatch = new[] { "[I]", "CDE", "VWX" }.Select(quote => ExactColorRange(original, quote)).ToArray();
            var context = Share(editor, address, id);
            Tool(address, "highlight_text", ColorRequest(id, "overlap_seed_first_batch", context, null, "protect", firstBatch));
            context = Share(editor, address, id);
            Tool(address, "highlight_text", ColorRequest(id, "overlap_seed_second_batch", context, null, "protect", secondBatch));
            var before = Local(editor, "state");
            var marks = UiState.Records(before, "marks");
            Check(marks.Count == 7 && marks.All(mark => !UiState.Flag(mark, "stale") && UiState.Text(mark, "ruleId") == "protect"),
                "Two ordinary green batches create seven current markers with identical, partial and nested overlaps.");
            var selected = marks.Where(mark => UiState.Text(mark, "quote") != "[N]").OrderBy(mark => (int)mark["from"]).ToArray();
            Check(selected.Select(mark => UiState.Text(mark, "quote")).SequenceEqual(new[] { "[I]", "[I]", "ABCD", "CDE", "UVWXY", "VWX" }) &&
                selected.All(mark => (int)mark["from"] == original.IndexOf(UiState.Text(mark, "quote"), StringComparison.Ordinal) &&
                    (int)mark["to"] == (int)mark["from"] + UiState.Text(mark, "quote").Length),
                "The selected source IDs retain the exact identical, partial-overlap and nested range geometry.");
            object[] ranges = selected.Select(mark => (object)new { from = mark["from"], to = mark["to"], quote = mark["quote"] }).ToArray();
            string[] greenIds = selected.Select(mark => UiState.Text(mark, "id")).ToArray();
            var neighbor = marks.Single(mark => UiState.Text(mark, "quote") == "[N]");
            string neighborId = UiState.Text(neighbor, "id"), neighborState = McpEndpoint.Json().Serialize(neighbor);
            string neighborFormatting = WordSnapshot.FormattingSignature(document.Range((int)neighbor["from"], (int)neighbor["to"]));
            string rules = McpEndpoint.Json().Serialize(before["rules"]), fonts = RecolorFonts(document);
            int[] originalPaint = RecolorCharacterPaint(document);
            document.Save(); string disk = FileDigest(path);
            context = Share(editor, address, id);
            AssertRecolorRefusal(editor, address, document, id,
                ColorRequest(id, "ordinary_overlap_batch_stays_invalid", context, null, "protect", ranges),
                "ordinary overlapping ranges without exact source replacement IDs", evidence);
            foreach (string quote in new[] { "[I]", "ABCD", "UVWXY", "VWX" })
            {
                var single = selected.First(mark => UiState.Text(mark, "quote") == quote);
                context = Share(editor, address, id);
                Dictionary<string, object> refused = AssertRecolorRefusal(editor, address, document, id,
                    RecolorRequest(id, "overlap_single_" + Array.IndexOf(selected, single), context, null, "overlap_red",
                        new[] { ExactColorRange(original, quote) }, new[] { UiState.Text(single, "id") }),
                    "one selected source leaves unselected overlapping protection for " + quote, evidence);
                Check(UiState.Text(Map(((IEnumerable)refused["content"]).Cast<object>().First()), "text").Contains("overlaps marker"),
                    "The partial selection of " + quote + " identifies the unselected overlapping marker instead of weakening its protection.");
            }
            context = Share(editor, address, id);
            object request = RecolorRequest(id, "all_overlapping_green_to_red", context, null, "overlap_red", ranges, greenIds);
            object nativeBefore = RecolorNativeEvidence(document, ranges);
            var timer = Stopwatch.StartNew();
            var receipt = Tool(address, "highlight_text", request);
            string[] redIds = RecolorIds(receipt, "markIds");
            evidence.Add(new { operation = "all_identical_partial_nested_sources_to_red", milliseconds = timer.ElapsedMilliseconds,
                request = request, receipt = receipt, beforeNative = nativeBefore, afterNative = RecolorNativeEvidence(document, ranges), state = Local(editor, "state") });
            Check(redIds.Length == 6 && redIds.Distinct().Count() == 6 && !redIds.Intersect(greenIds).Any() &&
                RecolorIds(receipt, "replacedMarkIds").SequenceEqual(greenIds) && UiState.Text(receipt, "colorId") == "overlap_red" && UiState.Text(receipt, "action") == "reject",
                "Selecting all six overlapping sources returns distinct red identities with the exact ordered retired green IDs.");
            AssertRecolorBindings(editor, document, redIds, greenIds, ranges, "overlap_red", "#E57575", "Overlapping batch", 7);
            var exactRanges = ranges.Select(Map).ToArray();
            int[] expectedPaint = originalPaint.Select((color, offset) => exactRanges.Any(range => (int)range["from"] <= offset && offset < (int)range["to"])
                ? ColorNumber("#E57575") : color).ToArray();
            Check(((int[])RecolorCharacterPaint(document)).SequenceEqual(expectedPaint),
                "The identical, partial and nested range union is painted red with every character outside the selected union unchanged.");
            before = Local(editor, "state");
            Check(McpEndpoint.Json().Serialize(before["rules"]) == rules && WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts &&
                McpEndpoint.Json().Serialize(UiState.Records(before, "marks").Single(mark => UiState.Text(mark, "id") == neighborId)) == neighborState &&
                WordSnapshot.FormattingSignature(document.Range((int)neighbor["from"], (int)neighbor["to"])) == neighborFormatting &&
                (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0 && FileDigest(path) == disk && !(bool)document.Saved,
                "Overlapping recoloring preserves local rules, all prose/fonts, the neighboring protected marker, review annotations and saved bytes.");
            string nativeRed = new WordNativeEditState().NativeIdentity(document), stateRed = RecolorMetadata(before), sidecarRed = FileDigest(RecolorSidecar(id));
            var replay = Tool(address, "highlight_text", request);
            Check(UiState.Flag(replay, "replayed") && RecolorIds(replay, "markIds").SequenceEqual(redIds) &&
                RecolorIds(replay, "replacedMarkIds").SequenceEqual(greenIds) && nativeRed == new WordNativeEditState().NativeIdentity(document) &&
                stateRed == RecolorMetadata(Local(editor, "state")) && sidecarRed == FileDigest(RecolorSidecar(id)),
                "An exact overlapping batch replay retains the same six replacements without any further native or durable effect.");
            Check((bool)document.Undo(1), "One native Undo reverses all identical, partial and nested recolor replacements.");
            AssertRecolorBindings(editor, document, greenIds, redIds, ranges, "protect", "#74B9A4", "Overlapping Undo", 7);
            Check(((int[])RecolorCharacterPaint(document)).SequenceEqual(originalPaint), "Overlapping Undo restores the exact original character paint inside and outside every source.");
            Check((bool)document.Redo(1), "One native Redo reapplies all identical, partial and nested recolor replacements.");
            AssertRecolorBindings(editor, document, redIds, greenIds, ranges, "overlap_red", "#E57575", "Overlapping Redo", 7);
            Check(((int[])RecolorCharacterPaint(document)).SequenceEqual(expectedPaint), "Overlapping Redo restores exactly the red union without enlarging it.");
            document.Save(); document.Close(0); document = null;
            document = application.Documents.Open(path, ReadOnly: false, Visible: false); document.Activate();
            AssertRecolorBindings(editor, document, redIds, greenIds, ranges, "overlap_red", "#E57575", "Overlapping DOCX reopen", 7);
            before = Local(editor, "state");
            Check(!UiState.Flag(UiState.Map(before["document"]), "shared") && McpEndpoint.Json().Serialize(before["rules"]) == rules &&
                WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts && ((int[])RecolorCharacterPaint(document)).SequenceEqual(expectedPaint) &&
                McpEndpoint.Json().Serialize(UiState.Records(before, "marks").Single(mark => UiState.Text(mark, "id") == neighborId)) == neighborState &&
                WordSnapshot.FormattingSignature(document.Range((int)neighbor["from"], (int)neighbor["to"])) == neighborFormatting,
                "DOCX reopen retains the exact overlapping red bindings, union paint, rules, fonts, prose and neighbor without sharing authority.");
            evidence.Add(new { operation = "overlap_saved_reopened", state = before, native = RecolorNativeEvidence(document, ranges), documentSha256 = FileDigest(path) });
        }
        finally
        {
            if (document != null) document.Close(0);
            if (id != null && File.Exists(RecolorSidecar(id))) File.Delete(RecolorSidecar(id));
        }
    }

    private static void MixedInheritedSameRgbRecolor(dynamic application, WordEditor editor, string address, string output, List<object> evidence)
    {
        dynamic document = application.Documents.Add();
        string id = null;
        try
        {
            document.Content.Text = "Mixed ABCDE shading.\r";
            document.Content.Font.Name = "Calibri"; document.Content.Font.Size = 11;
            document.Activate();
            string path = Path.Combine(output, "mixed-inherited-same-rgb-recoloring.docx");
            document.SaveAs2(path, 16);
            id = UiState.Text(UiState.Map(Local(editor, "state")["document"]), "id");
            Check(!File.Exists(RecolorSidecar(id)), "The mixed inherited/explicit recoloring fixture has no prior private receipt.");
            Local(editor, "save_rule", new { documentId = id, id = "protect", label = "Original green", color = "#74B9A4",
                action = "protect", prompt = "", accessMode = "suggestions" });
            Local(editor, "save_rule", new { documentId = id, id = "mixed_green_target", label = "Second Protect green", color = "#74B9A4",
                action = "protect", prompt = "", accessMode = "suggestions" });
            Local(editor, "save_rule", new { documentId = id, id = "mixed_green_inner", label = "Neutral green inner", color = "#74B9A4",
                action = "neutral", prompt = "", accessMode = "suggestions" });
            string original = WordSnapshot.Capture(document).FullText;
            object[] outerRanges = new[] { ExactColorRange(original, "ABCDE") };
            object[] innerRanges = new[] { ExactColorRange(original, "BCD") };
            var outer = Map(outerRanges[0]); int from = (int)outer["from"], to = (int)outer["to"];
            var inner = Map(innerRanges[0]); int innerFrom = (int)inner["from"], innerTo = (int)inner["to"];
            document.Range(innerFrom, innerTo).Font.Italic = -1;
            var context = Share(editor, address, id);
            var seededOuter = Tool(address, "highlight_text", ColorRequest(id, "mixed_seed_outer", context, null, "protect", outerRanges));
            string oldId = RecolorIds(seededOuter, "markIds").Single();
            context = Share(editor, address, id);
            var seededInner = Tool(address, "highlight_text", ColorRequest(id, "mixed_seed_inner", context, null, "mixed_green_inner", innerRanges));
            string innerId = RecolorIds(seededInner, "markIds").Single();
            document.Range(from, to).Font.Shading.BackgroundPatternColor = -16777216;
            document.Range(from, to).ParagraphFormat.Shading.BackgroundPatternColor = ColorNumber("#74B9A4");
            document.Range(innerFrom, innerTo).Font.Shading.BackgroundPatternColor = ColorNumber("#74B9A4");
            var state = Local(editor, "state");
            int[] beforePaint = RecolorCharacterPaint(document);
            Check(UiState.Records(state, "marks").Count == 2 && UiState.Records(state, "marks").All(mark => !UiState.Flag(mark, "stale")) &&
                Enumerable.Range(from, to - from).All(offset => beforePaint[offset] == (innerFrom <= offset && offset < innerTo ? ColorNumber("#74B9A4") : -16777216)) &&
                (int)document.Range(from, to).ParagraphFormat.Shading.BackgroundPatternColor == ColorNumber("#74B9A4"),
                "Overlapping current green markers can expose uniform green with inherited outer characters and explicit same-green inner characters.");
            string formatting = WordSnapshot.FormattingSignature(document.Range(from, to));
            string fonts = RecolorFonts(document), rules = McpEndpoint.Json().Serialize(state["rules"]);
            string innerState = McpEndpoint.Json().Serialize(UiState.Records(state, "marks").Single(mark => UiState.Text(mark, "id") == innerId));
            object nativeBefore = RecolorNativeEvidence(document, outerRanges.Concat(innerRanges).ToArray());
            document.Save(); string disk = FileDigest(path);
            context = Share(editor, address, id);
            object request = RecolorRequest(id, "mixed_same_rgb_outer", context, null, "mixed_green_target", outerRanges, new[] { oldId });
            var timer = Stopwatch.StartNew();
            var receipt = Tool(address, "highlight_text", request);
            long milliseconds = timer.ElapsedMilliseconds;
            state = Local(editor, "state");
            string newId = RecolorIds(receipt, "markIds").Single();
            evidence.Add(new { operation = "mixed_inherited_explicit_same_rgb", milliseconds = milliseconds, request = request, receipt = receipt,
                beforeNative = nativeBefore, afterNative = RecolorNativeEvidence(document, outerRanges.Concat(innerRanges).ToArray()), state = state });
            Check(WordSnapshot.FormattingSignature(document.Range(from, to)) == formatting &&
                ((int[])RecolorCharacterPaint(document)).SequenceEqual(beforePaint) &&
                (int)document.Range(from, to).ParagraphFormat.Shading.BackgroundPatternColor == ColorNumber("#74B9A4"),
                "Same-RGB outer recoloring preserves exact mixed inherited/explicit native formatting without flattening automatic character shading.");
            Check(newId != oldId && RecolorIds(receipt, "replacedMarkIds").SequenceEqual(new[] { oldId }) && UiState.Text(receipt, "colorId") == "mixed_green_target" &&
                UiState.Records(state, "marks").Count == 2 && UiState.Records(state, "marks").All(mark => !UiState.Flag(mark, "stale")) &&
                !(bool)document.Bookmarks.Exists(oldId) && (bool)document.Bookmarks.Exists(newId) &&
                (int)document.Bookmarks.Item(newId).Range.Start == from && (int)document.Bookmarks.Item(newId).Range.End == to &&
                innerState == McpEndpoint.Json().Serialize(UiState.Records(state, "marks").Single(mark => UiState.Text(mark, "id") == innerId)) &&
                (bool)document.Bookmarks.Exists(innerId) && (int)document.Bookmarks.Item(innerId).Range.Start == innerFrom &&
                (int)document.Bookmarks.Item(innerId).Range.End == innerTo,
                "The mixed-shading recolor replaces only the exact outer identity and leaves the inner neutral marker current and unchanged.");
            Check(McpEndpoint.Json().Serialize(state["rules"]) == rules && WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts &&
                (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0 && FileDigest(path) == disk && !(bool)document.Saved,
                "Mixed inherited/explicit recoloring preserves all local rules, prose, fonts, review annotations and saved bytes.");
        }
        finally
        {
            document.Close(0);
            if (id != null && File.Exists(RecolorSidecar(id))) File.Delete(RecolorSidecar(id));
        }
    }

    private static void SingleMarkerMixedColor(dynamic application, WordEditor editor, string address, string output, List<object> evidence)
    {
        dynamic document = application.Documents.Add();
        string id = null;
        try
        {
            document.Content.Text = "Single ABCDEFGHI marker. Unmarked prose.\r";
            document.Content.Font.Name = "Calibri"; document.Content.Font.Size = 11;
            document.Activate(); document.SaveAs2(Path.Combine(output, "single-marker-mixed-color.docx"), 16);
            id = UiState.Text(UiState.Map(Local(editor, "state")["document"]), "id");
            Check(!File.Exists(RecolorSidecar(id)), "The single-marker color fixture has no prior private receipt.");
            Local(editor, "save_rule", new { documentId = id, id = "protect", label = "Protect green", color = "#74B9A4",
                action = "protect", prompt = "", accessMode = "suggestions" });
            Local(editor, "save_rule", new { documentId = id, id = "single_green_target", label = "Second Protect green", color = "#74B9A4",
                action = "protect", prompt = "", accessMode = "suggestions" });
            string original = WordSnapshot.Capture(document).FullText;
            object[] ranges = new[] { ExactColorRange(original, "ABCDEFGHI") };
            var exact = Map(ranges[0]); int from = (int)exact["from"], to = (int)exact["to"];
            int innerFrom = from + 2, innerTo = to - 2;
            document.Range(innerFrom, innerTo).Font.Italic = -1;
            var context = Share(editor, address, id);
            var seed = Tool(address, "highlight_text", ColorRequest(id, "single_mixed_seed", context, null, "protect", ranges));
            string oldId = RecolorIds(seed, "markIds").Single();
            document.Range(from, to).Font.Shading.BackgroundPatternColor = -16777216;
            document.Range(from, to).ParagraphFormat.Shading.BackgroundPatternColor = ColorNumber("#74B9A4");
            document.Range(innerFrom, innerTo).Font.Shading.BackgroundPatternColor = ColorNumber("#74B9A4");
            string formatting = WordSnapshot.FormattingSignature(document.Range(from, to)), fonts = RecolorFonts(document);
            int[] originalPaint = RecolorCharacterPaint(document);
            var state = Local(editor, "state");
            string rules = McpEndpoint.Json().Serialize(state["rules"]);
            context = Share(editor, address, id);
            var highlights = UiState.Records(context, "highlights");
            Check(UiState.Flag(context, "complete") && UiState.Text(context, "manuscript_id") == id && highlights.Count == 1 &&
                UiState.Text(highlights.Single(), "id") == oldId && !UiState.Flag(highlights.Single(), "stale") && highlights.Single()["staleReason"] == null &&
                Enumerable.Range(from, to - from).All(offset => originalPaint[offset] == (innerFrom <= offset && offset < innerTo ? ColorNumber("#74B9A4") : -16777216)),
                "A complete MCP read validates one mixed inherited/direct green marker without inner marker cuts.");
            object nativeBefore = RecolorNativeEvidence(document, ranges);
            object request = RecolorRequest(id, "single_mixed_same_rgb", context, null, "single_green_target", ranges, new[] { oldId });
            var timer = Stopwatch.StartNew();
            var receipt = Tool(address, "highlight_text", request);
            long milliseconds = timer.ElapsedMilliseconds;
            string markId = RecolorIds(receipt, "markIds").Single();
            state = Local(editor, "state");
            evidence.Add(new { operation = "single_marker_mixed_same_rgb", milliseconds = milliseconds, request = request, receipt = receipt,
                beforeNative = nativeBefore, afterNative = RecolorNativeEvidence(document, ranges), state = state });
            Check(markId != oldId && RecolorIds(receipt, "replacedMarkIds").SequenceEqual(new[] { oldId }) &&
                UiState.Records(state, "marks").Count == 1 && UiState.Text(UiState.Records(state, "marks").Single(), "ruleId") == "single_green_target" &&
                !UiState.Flag(UiState.Records(state, "marks").Single(), "stale") && !(bool)document.Bookmarks.Exists(oldId) &&
                (bool)document.Bookmarks.Exists(markId) && (int)document.Bookmarks.Item(markId).Range.Start == from && (int)document.Bookmarks.Item(markId).Range.End == to,
                "The single mixed marker receives one exact replacement identity while remaining current.");
            Check(WordSnapshot.FormattingSignature(document.Range(from, to)) == formatting && ((int[])RecolorCharacterPaint(document)).SequenceEqual(originalPaint) &&
                WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts && McpEndpoint.Json().Serialize(state["rules"]) == rules,
                "Same-RGB recoloring preserves exact mixed native formatting, all prose/fonts and every local rule without flattening inheritance.");
            foreach (string mismatch in new[] { "one_wrong_character_rgb", "one_native_highlight" })
            {
                if (mismatch == "one_wrong_character_rgb") document.Range(innerFrom, innerFrom + 1).Font.Shading.BackgroundPatternColor = ColorNumber("#314B96");
                else document.Range(innerFrom, innerFrom + 1).HighlightColorIndex = 7;
                context = Share(editor, address, id);
                var highlight = UiState.Records(context, "highlights").Single();
                evidence.Add(new { operation = mismatch, context = context, state = Local(editor, "state"), native = RecolorNativeEvidence(document, ranges) });
                Check(UiState.Flag(context, "complete") && UiState.Text(highlight, "id") == markId && UiState.Flag(highlight, "stale") &&
                    UiState.Text(highlight, "staleReason") == "visible_color_changed" && UiState.Text(highlight, "protectedId") == markId,
                    "A complete native read marks " + mismatch + " stale with the exact current protected identity.");
                Dictionary<string, object> refused = AssertRecolorToolRefusal(editor, address, document, id, "suggest_text",
                    new { manuscript_id = id, request_id = mismatch + "_edit", revision = context["revision"], context_id = context["context_id"],
                        from = from, to = to, quote = exact["quote"], replacement = "changed" }, mismatch + " wording", evidence);
                var diagnostic = Map(Map(refused["structuredContent"])["error"]);
                var unresolved = UiState.Records(Map(diagnostic["details"]), "markers");
                Check(UiState.Text(diagnostic, "code") == "marker_permission_unresolved" && unresolved.Count == 1 &&
                    UiState.Text(unresolved.Single(), "id") == markId && UiState.Text(unresolved.Single(), "reason") == "visible_color_changed",
                    "The " + mismatch + " edit refusal identifies the actual unsafe character-color marker.");
                refused = AssertRecolorRefusal(editor, address, document, id,
                    RecolorRequest(id, mismatch + "_recolor", context, null, "protect", ranges, new[] { markId }), mismatch + " recoloring", evidence);
                Check(UiState.Text(Map(((IEnumerable)refused["content"]).Cast<object>().First()), "text").Contains("stale or unresolved") &&
                    WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts &&
                    McpEndpoint.Json().Serialize(Local(editor, "state")["rules"]) == rules && (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0,
                    "Refusing " + mismatch + " preserves body, fonts, rules and review annotations instead of treating a near-match as permission.");
                document.Range(innerFrom, innerFrom + 1).Font.Shading.BackgroundPatternColor = ColorNumber("#74B9A4");
                document.Range(innerFrom, innerFrom + 1).HighlightColorIndex = 0;
                context = Tool(address, "read_manuscript", new { manuscript_id = id });
                Check(UiState.Flag(context, "complete") && !UiState.Flag(UiState.Records(context, "highlights").Single(), "stale"),
                    "Restoring the exact native color after " + mismatch + " makes the next complete read current.");
            }
        }
        finally
        {
            document.Close(0);
            if (id != null && File.Exists(RecolorSidecar(id))) File.Delete(RecolorSidecar(id));
        }
    }

    private static void ColorVerificationBudget(dynamic application, WordEditor editor, string address, string output, List<object> evidence)
    {
        dynamic document = application.Documents.Add();
        string id = null;
        try
        {
            string quote = new string('x', 512);
            document.Content.Text = quote + "\r";
            document.Content.Font.Name = "Calibri"; document.Content.Font.Size = 11;
            document.Activate(); document.SaveAs2(Path.Combine(output, "color-verification-budget.docx"), 16);
            id = UiState.Text(UiState.Map(Local(editor, "state")["document"]), "id");
            Check(!File.Exists(RecolorSidecar(id)), "The color-budget fixture has no prior private receipt.");
            Local(editor, "save_rule", new { documentId = id, id = "protect", label = "Protect green", color = "#74B9A4",
                action = "protect", prompt = "", accessMode = "suggestions" });
            object[] ranges = new[] { ExactColorRange(quote, quote) };
            var context = Share(editor, address, id);
            var seed = Tool(address, "highlight_text", ColorRequest(id, "budget_seed", context, null, "protect", ranges));
            string markId = RecolorIds(seed, "markIds").Single();
            document.Range(0, quote.Length).Font.Shading.BackgroundPatternColor = -16777216;
            document.Range(0, quote.Length).ParagraphFormat.Shading.BackgroundPatternColor = ColorNumber("#74B9A4");
            for (int offset = 1; offset < quote.Length; offset += 2)
                document.Range(offset, offset + 1).Font.Shading.BackgroundPatternColor = ColorNumber("#74B9A4");
            int[] alternating = RecolorCharacterPaint(document);
            Check(Enumerable.Range(0, quote.Length).All(offset => alternating[offset] == (offset % 2 == 0 ? -16777216 : ColorNumber("#74B9A4"))),
                "The long marked range contains genuinely alternating inherited/direct same-green character shading.");
            string original = WordSnapshot.Capture(document).FullText, fonts = RecolorFonts(document);
            string native = new WordNativeEditState().NativeIdentity(document), sidecar = FileDigest(RecolorSidecar(id));
            string policy = McpEndpoint.Json().Serialize(WordReviewStore.Load((object)document).Policy);
            object nativeBefore = RecolorNativeEvidence(document, ranges);
            var timer = Stopwatch.StartNew();
            var refused = Rpc(address, "read_manuscript", new { manuscript_id = id });
            long milliseconds = timer.ElapsedMilliseconds;
            evidence.Add(new { operation = "alternating_color_budget", milliseconds = milliseconds, response = refused,
                beforeNative = nativeBefore, afterNative = RecolorNativeEvidence(document, ranges) });
            Check(UiState.Flag(refused, "isError") && refused.ContainsKey("structuredContent"),
                "Pathological mixed shading returns a structured inspection-limit refusal instead of hanging or sampling the range.");
            var diagnostic = Map(Map(refused["structuredContent"])["error"]);
            var details = Map(diagnostic["details"]);
            Check(UiState.Text(diagnostic, "code") == "native_color_verification_limited" && (int)details["additionalProbes"] == 256 &&
                (int)details["inspectionMilliseconds"] == 1000 && milliseconds < 20000,
                "The native refusal reports the exact 256-probe/1000ms verification limit within the existing 20-second MCP native timeout.");
            Check(native == new WordNativeEditState().NativeIdentity(document) && WordSnapshot.Capture(document).FullText == original &&
                RecolorFonts(document) == fonts && sidecar == FileDigest(RecolorSidecar(id)) &&
                McpEndpoint.Json().Serialize(WordReviewStore.Load((object)document).Policy) == policy &&
                (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0,
                "Exhausting the color-verification budget changes no native text, shading, bookmarks, fonts, rule or durable receipt.");
            document.Range(0, quote.Length).Font.Shading.BackgroundPatternColor = ColorNumber("#74B9A4");
            context = Tool(address, "read_manuscript", new { manuscript_id = id });
            Check(UiState.Flag(context, "complete") && UiState.Records(context, "highlights").Count == 1 &&
                UiState.Text(UiState.Records(context, "highlights").Single(), "id") == markId && !UiState.Flag(UiState.Records(context, "highlights").Single(), "stale"),
                "After a uniform native shading reset, the immediately following dispatch succeeds without a wedged review operation.");
            document.Range(0, quote.Length / 2).Font.Shading.BackgroundPatternColor = -16777216;
            context = Tool(address, "read_manuscript", new { manuscript_id = id });
            evidence.Add(new { operation = "color_budget_recovered", context = context, state = Local(editor, "state"), native = RecolorNativeEvidence(document, ranges) });
            Check(UiState.Flag(context, "complete") && !UiState.Flag(UiState.Records(context, "highlights").Single(), "stale") &&
                WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts &&
                McpEndpoint.Json().Serialize(WordReviewStore.Load((object)document).Policy) == policy,
                "A small mixed inherited/direct range verifies on a later dispatch, proving exhausted probe/deadline state was reset.");
        }
        finally
        {
            document.Close(0);
            if (id != null && File.Exists(RecolorSidecar(id))) File.Delete(RecolorSidecar(id));
        }
    }

    private static void RecolorRefusals(WordEditor editor, string address, dynamic document, string id, object[] ranges,
        string[] currentIds, string[] retiredIds, List<object> evidence)
    {
        var invalid = new[] {
            new { Name = "not_array", Ids = (object)currentIds[0] },
            new { Name = "dictionary_ids", Ids = (object)new Dictionary<string, object> { { "id", currentIds[0] } } },
            new { Name = "null_ids", Ids = (object)null },
            new { Name = "not_strings", Ids = (object)new object[] { currentIds[0], currentIds[1], currentIds[2], 7 } },
            new { Name = "null_id", Ids = (object)currentIds.Take(3).Concat(new string[] { null }).ToArray() },
            new { Name = "empty_ids", Ids = (object)new string[0] },
            new { Name = "missing_pair", Ids = (object)currentIds.Take(3).ToArray() },
            new { Name = "duplicate_id", Ids = (object)new[] { currentIds[0], currentIds[1], currentIds[2], currentIds[0] } },
            new { Name = "invalid_id", Ids = (object)currentIds.Take(3).Concat(new[] { "invalid#id" }).ToArray() },
            new { Name = "unknown_id", Ids = (object)currentIds.Take(3).Concat(new[] { "Wrai_missing_identity" }).ToArray() },
            new { Name = "retired_id", Ids = (object)currentIds.Take(3).Concat(new[] { retiredIds[3] }).ToArray() },
            new { Name = "wrong_order", Ids = (object)new[] { currentIds[1], currentIds[0], currentIds[2], currentIds[3] } }
        };
        foreach (var test in invalid)
        {
            var context = Share(editor, address, id);
            AssertRecolorRefusal(editor, address, document, id,
                RecolorRequest(id, "invalid_" + test.Name, context, "#1267AB", null, ranges, test.Ids), test.Name, evidence);
        }
        var partial = ranges.ToArray();
        var last = Map(partial[3]);
        partial[3] = new { from = (int)last["from"], to = (int)last["to"] - 1, quote = "[4" };
        var fresh = Share(editor, address, id);
        AssertRecolorRefusal(editor, address, document, id, RecolorRequest(id, "partial_final_range", fresh, "#1267AB", null, partial, currentIds),
            "partial final range after three valid pairs", evidence);
        var wrongQuote = ranges.ToArray(); wrongQuote[3] = new { from = last["from"], to = last["to"], quote = "[9]" };
        fresh = Share(editor, address, id);
        AssertRecolorRefusal(editor, address, document, id, RecolorRequest(id, "wrong_final_quote", fresh, "#1267AB", null, wrongQuote, currentIds),
            "wrong final quote after three valid pairs", evidence);
        fresh = Share(editor, address, id);
        var forged = Map(RecolorRequest(id, "forged_context", fresh, "#1267AB", null, ranges, currentIds));
        ((Dictionary<string, object>)forged["marking"])["contextId"] = "forged_context";
        AssertRecolorRefusal(editor, address, document, id, forged, "forged reading context", evidence);
        fresh = Share(editor, address, id);
        object revoked = RecolorRequest(id, "revoked_recolor", fresh, "#1267AB", null, ranges, currentIds);
        Local(editor, "revoke_context");
        AssertRecolorRefusal(editor, address, document, id, revoked, "revoked document grant", evidence);
        fresh = Share(editor, address, id);
        AssertRecolorRefusal(editor, address, document, id,
            RecolorRequest("different_document", "wrong_document", fresh, "#1267AB", null, ranges, currentIds), "wrong document identity", evidence);
        document.Bookmarks.Item(currentIds[3]).Delete();
        var staleMarks = UiState.Records(Local(editor, "state"), "marks");
        Check(staleMarks.Count(mark => UiState.Flag(mark, "stale")) == 1 &&
            UiState.Flag(staleMarks.Single(mark => UiState.Text(mark, "id") == currentIds[3]), "stale"),
            "Deleting one source bookmark makes exactly that native recolor target stale.");
        fresh = Share(editor, address, id);
        AssertRecolorRefusal(editor, address, document, id,
            RecolorRequest(id, "stale_final_marker", fresh, "#1267AB", null, ranges, currentIds), "stale final marker after three valid pairs", evidence);
        document.Bookmarks.Add(currentIds[3], document.Range((int)last["from"], (int)last["to"]));
        Check(UiState.Records(Local(editor, "state"), "marks").All(mark => !UiState.Flag(mark, "stale")),
            "Restoring the synthetic native bookmark makes the original red bindings current again.");
    }

    private static void RecolorRuleUndo(WordEditor editor, string address, dynamic document, string id, string colorId,
        string[] ids, object[] ranges, IDictionary<string, object> oldContext, List<object> evidence)
    {
        var originalState = Local(editor, "state");
        var neighbor = UiState.Records(originalState, "marks").Single(mark => !ids.Contains(UiState.Text(mark, "id")));
        string neighborId = UiState.Text(neighbor, "id"), neighborState = McpEndpoint.Json().Serialize(neighbor);
        string neighborFormatting = WordSnapshot.FormattingSignature(document.Range((int)neighbor["from"], (int)neighbor["to"]));
        string original = WordSnapshot.Capture(document).FullText, fonts = RecolorFonts(document);
        object stale = RecolorRequest(id, "old_neutral_authority", oldContext, null, "recolor_red", ranges, ids);
        Local(editor, "save_rule", new { documentId = id, id = colorId, label = "New local protection", color = "#1267AB",
            action = "protect", prompt = "", accessMode = "suggestions" });
        var beforeUndo = Local(editor, "state");
        string currentRules = McpEndpoint.Json().Serialize(beforeUndo["rules"]);
        AssertRecolorRefusal(editor, address, document, id, stale, "request read before a local rule change", evidence);
        object nativeBefore = RecolorNativeEvidence(document, ranges);
        bool undone = (bool)document.Undo(1);
        object nativeAfterUndo = RecolorNativeEvidence(document, ranges);
        var state = Local(editor, "state");
        evidence.Add(new { operation = "local_rule_change_and_undo", originalState = originalState, beforeState = beforeUndo, afterState = state,
            beforeNative = nativeBefore, afterUndoBeforeValidation = nativeAfterUndo, afterValidation = RecolorNativeEvidence(document, ranges) });
        Check(undone, "Native Undo can reach the metadata written before the latest local function assignment.");
        Check(McpEndpoint.Json().Serialize(state["rules"]) == currentRules &&
            UiState.Records(state, "rules").Single(rule => UiState.Text(rule, "id") == colorId)["action"].ToString() == "protect",
            "Native metadata Undo cannot restore the earlier neutral function over the newer local Protect rule.");
        var marks = UiState.Records(state, "marks");
        Check(marks.Count == 5 && marks.Select(mark => UiState.Text(mark, "id")).OrderBy(value => value).SequenceEqual(
            UiState.Records(beforeUndo, "marks").Select(mark => UiState.Text(mark, "id")).OrderBy(value => value)),
            "Undo retains the latest policy marker identities rather than resurrecting historical recolor authority.");
        for (int index = 0; index < ids.Length; index++)
        {
            string selected = ids[index]; var range = Map(ranges[index]);
            var mark = marks.Single(candidate => UiState.Text(candidate, "id") == selected);
            Check(UiState.Text(mark, "ruleId") == colorId && (int)mark["from"] == (int)range["from"] && (int)mark["to"] == (int)range["to"] &&
                UiState.Text(mark, "quote") == (string)range["quote"] && UiState.Text(mark, "prompt") == "",
                "Undo retains citation " + (index + 1) + " under the latest exact Protect binding even when its native witness is unresolved.");
        }
        Check(WordSnapshot.Capture(document).FullText == original && RecolorFonts(document) == fonts &&
            McpEndpoint.Json().Serialize(marks.Single(mark => UiState.Text(mark, "id") == neighborId)) == neighborState &&
            WordSnapshot.FormattingSignature(document.Range((int)neighbor["from"], (int)neighbor["to"])) == neighborFormatting &&
            (int)document.Revisions.Count == 0 && (int)document.Comments.Count == 0,
            "Local function change and Undo preserve all prose, fonts, the unselected green neighbor and native review annotations.");
        string[] unresolvedIds = marks.Where(mark => ids.Contains(UiState.Text(mark, "id")) && UiState.Flag(mark, "stale"))
            .Select(mark => UiState.Text(mark, "id")).ToArray();
        Check(unresolvedIds.Length > 0, "Undo beyond the retained native witness marks its affected protected identities unresolved.");
        AssertRecolorRefusal(editor, address, document, id, stale, "historical authority after native Undo", evidence);
        var first = Map(ranges[0]);
        object oldEdit = new { manuscript_id = id, request_id = "unsafe_old_edit_after_undo", revision = oldContext["revision"],
            context_id = oldContext["context_id"], from = first["from"], to = first["to"], quote = first["quote"], replacement = "[9]" };
        AssertRecolorToolRefusal(editor, address, document, id, "suggest_text", oldEdit, "historical wording authority after native Undo", evidence);
        var fresh = Share(editor, address, id);
        evidence.Add(new { operation = "fresh_fail_closed_undo_context", context = fresh, state = Local(editor, "state"), native = RecolorNativeEvidence(document, ranges) });
        var unresolved = UiState.Records(fresh, "highlights").Where(mark => unresolvedIds.Contains(UiState.Text(mark, "id"))).ToArray();
        Check(UiState.Flag(fresh, "complete") && UiState.Text(fresh, "manuscript_id") == id && unresolved.Length == unresolvedIds.Length &&
            unresolved.All(mark => UiState.Flag(mark, "stale") && UiState.Text(mark, "ruleId") == colorId &&
                UiState.Text(mark, "protectedId") == UiState.Text(mark, "id") && !String.IsNullOrWhiteSpace(UiState.Text(mark, "staleReason"))) &&
            McpEndpoint.Json().Serialize(fresh["rules"]) == currentRules,
            "A fresh complete read reports each unresolved current Protect identity with an actionable stale reason and protectedId.");
        object freshEdit = new { manuscript_id = id, request_id = "unsafe_fresh_edit_after_undo", revision = fresh["revision"],
            context_id = fresh["context_id"], from = first["from"], to = first["to"], quote = first["quote"], replacement = "[9]" };
        Dictionary<string, object> refusal = AssertRecolorToolRefusal(editor, address, document, id, "suggest_text", freshEdit, "fresh wording request over unresolved protection", evidence);
        var diagnostic = Map(Map(refusal["structuredContent"])["error"]);
        var diagnostics = UiState.Records(Map(diagnostic["details"]), "markers");
        Check(UiState.Text(diagnostic, "code") == "marker_permission_unresolved" && diagnostics.Count == unresolved.Length &&
            diagnostics.All(mark => unresolved.Any(current => UiState.Text(current, "id") == UiState.Text(mark, "id") &&
                UiState.Text(current, "ruleId") == UiState.Text(mark, "ruleId") && UiState.Text(current, "staleReason") == UiState.Text(mark, "reason"))),
            "The fresh unsafe edit fails closed with the exact unresolved marker identities and reasons from the complete read.");
        AssertRecolorRefusal(editor, address, document, id,
            RecolorRequest(id, "fresh_recolor_after_undo", fresh, null, "recolor_red", ranges, ids), "fresh recoloring of unresolved protected identities", evidence);
        Check(McpEndpoint.Json().Serialize(Local(editor, "state")["rules"]) == currentRules,
            "Rejected historical and fresh requests never restore the earlier neutral function or weaken current local protection.");
    }

    private static object RecolorNativeEvidence(dynamic document, object[] ranges)
    {
        var bookmarks = new List<object>();
        foreach (object entry in (IEnumerable)document.Bookmarks)
        {
            dynamic bookmark = entry;
            bookmarks.Add(new { name = (string)bookmark.Name, from = (int)bookmark.Range.Start, to = (int)bookmark.Range.End,
                text = (string)bookmark.Range.Text });
        }
        var paint = new List<object>();
        foreach (object selected in ranges)
        {
            var range = Map(selected); int from = (int)range["from"], to = (int)range["to"];
            var characters = new List<object>();
            for (int offset = from; offset < to; offset++)
                characters.Add(new { from = offset, to = offset + 1, text = (string)document.Range(offset, offset + 1).Text,
                    character = NativePaint(document, offset, offset + 1),
                    paragraph = (int)document.Range(offset, offset + 1).ParagraphFormat.Shading.BackgroundPatternColor,
                    highlight = (int)document.Range(offset, offset + 1).HighlightColorIndex });
            paint.Add(new { from = from, to = to, text = (string)document.Range(from, to).Text,
                character = NativePaint(document, from, to), paragraph = (int)document.Range(from, to).ParagraphFormat.Shading.BackgroundPatternColor,
                highlight = (int)document.Range(from, to).HighlightColorIndex, characters = characters,
                formattingSignature = WordSnapshot.FormattingSignature(document.Range(from, to)) });
        }
        return new { nativeIdentity = new WordNativeEditState().NativeIdentity(document), fullText = WordSnapshot.Capture(document).FullText,
            fontSignature = RecolorFonts(document), bookmarks = bookmarks, paint = paint, nativeBodyXml = (string)document.Content.WordOpenXML };
    }

    private static object RecolorRequest(string id, string requestId, IDictionary<string, object> context, string color,
        string colorId, object[] ranges, object ids)
    {
        var request = Map(ColorRequest(id, requestId, context, color, colorId, ranges));
        ((Dictionary<string, object>)request["marking"])["recolorMarkIds"] = ids;
        return request;
    }

    private static string[] RecolorIds(IDictionary<string, object> values, string key)
    { return ((IEnumerable)values[key]).Cast<object>().Select(value => (string)value).ToArray(); }

    private static void AssertRecolorBindings(WordEditor editor, dynamic document, string[] expectedIds, string[] retiredIds,
        object[] ranges, string ruleId, string color, string operation, int totalMarks = 5)
    {
        var state = Local(editor, "state");
        var marks = UiState.Records(state, "marks");
        Check(marks.Count == totalMarks && marks.All(mark => !UiState.Flag(mark, "stale")) && !marks.Any(mark => retiredIds.Contains(UiState.Text(mark, "id"))),
            operation + " retains " + totalMarks + " current markers without retired source identities.");
        Check(retiredIds.All(id => !(bool)document.Bookmarks.Exists(id)), operation + " removes every retired native bookmark.");
        for (int index = 0; index < expectedIds.Length; index++)
        {
            string expected = expectedIds[index];
            var range = Map(ranges[index]);
            int from = (int)range["from"], to = (int)range["to"];
            var mark = marks.Single(candidate => UiState.Text(candidate, "id") == expected);
            Check(UiState.Text(mark, "ruleId") == ruleId && (int)mark["from"] == from && (int)mark["to"] == to &&
                UiState.Text(mark, "quote") == (string)range["quote"] && UiState.Text(mark, "prompt") == "" &&
                (bool)document.Bookmarks.Exists(expected) && (int)document.Bookmarks.Item(expected).Range.Start == from &&
                (int)document.Bookmarks.Item(expected).Range.End == to && NativePaint(document, from, to) == ColorNumber(color),
                operation + " preserves citation " + (index + 1) + " with its exact new binding, native range and visible paint.");
        }
    }

    private static Dictionary<string, object> AssertRecolorRefusal(WordEditor editor, string address, dynamic document, string id, object request,
        string label, List<object> evidence)
    { return AssertRecolorToolRefusal(editor, address, document, id, "highlight_text", request, label, evidence); }

    private static Dictionary<string, object> AssertRecolorToolRefusal(WordEditor editor, string address, dynamic document, string id, string tool, object request,
        string label, List<object> evidence)
    {
        string before = RecolorMetadata(Local(editor, "state"));
        string native = new WordNativeEditState().NativeIdentity(document), sidecar = FileDigest(RecolorSidecar(id));
        var timer = Stopwatch.StartNew();
        var response = Rpc(address, tool, request);
        evidence.Add(new { operation = "refused_" + label, tool = tool, milliseconds = timer.ElapsedMilliseconds, request = request, response = response,
            beforeNative = native, afterNative = new WordNativeEditState().NativeIdentity(document) });
        Check(UiState.Flag(response, "isError"), "Native " + tool + " refuses " + label + ".");
        Check(before == RecolorMetadata(Local(editor, "state")) && native == new WordNativeEditState().NativeIdentity(document) &&
            sidecar == FileDigest(RecolorSidecar(id)), "Refusing " + label + " atomically preserves every marker, rule, bookmark, font, prose and durable receipt.");
        return response;
    }

    private static string RecolorMetadata(IDictionary<string, object> state)
    { return McpEndpoint.Json().Serialize(new { rules = state["rules"], marks = state["marks"], proposals = state["proposals"] }); }

    private static string RecolorSidecar(string id)
    { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wrAIterWord", "reviews", id + ".bin"); }

    private static int[] RecolorCharacterPaint(dynamic document)
    {
        int[] colors = new int[(int)document.Content.End];
        for (int offset = 0; offset < colors.Length; offset++) colors[offset] = NativePaint(document, offset, offset + 1);
        return colors;
    }

    private static string RecolorFonts(dynamic document)
    {
        var fonts = new List<object>();
        int end = (int)document.Content.End;
        for (int offset = 0; offset < end; offset++)
        {
            dynamic font = document.Range(offset, offset + 1).Font;
            fonts.Add(new { name = (string)font.Name, size = (double)font.Size, bold = (int)font.Bold, italic = (int)font.Italic,
                underline = (int)font.Underline, color = (int)font.Color, superscript = (int)font.Superscript, subscript = (int)font.Subscript });
        }
        return PolicyEngine.Hash(McpEndpoint.Json().Serialize(fonts));
    }
}
