using System;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace UnderTree.WordAddin
{
    public sealed class UiRibbonControl
    {
        public string Tag { get; set; }
    }

    internal static class UiLayoutSmoke
    {
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr argument, IntPtr value);
        [DllImport("user32.dll")] private static extern uint GetGuiResources(IntPtr process, uint flags);
        [DllImport("gdi32.dll")] private static extern uint GetObjectType(IntPtr handle);
        [DllImport("dwmapi.dll")] private static extern int DwmGetWindowAttribute(IntPtr window, int attribute, out int value, int size);
        private static readonly List<string> checks = new List<string>();
        private static readonly List<string> captures = new List<string>();
        private static readonly List<Tuple<string, IDictionary<string, object>>> calls = new List<Tuple<string, IDictionary<string, object>>>();
        private static readonly List<IDictionary<string, object>> rules = new List<IDictionary<string, object>>();
        private static readonly Dictionary<string, object> document = new Dictionary<string, object>
        {
            { "id", "synthetic-ui-document" }, { "title", "Synthetic manuscript · no user document.docx" },
            { "shared", false }, { "revision", "revision-1" }
        };
        private static string output;
        private static string assetDirectory;
        private static bool selectionUnavailable;
        private static string commentAuthorName = "wordplai";

        [STAThread]
        private static int Main(string[] arguments)
        {
            output = Path.GetFullPath(arguments.Length == 0 ? "artifacts/word-ui" : arguments[0]);
            assetDirectory = Path.GetFullPath(arguments.Length > 1 ? arguments[1] : Path.Combine("assets", "wraiter"));
            Directory.CreateDirectory(output);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            AddRules();
            try
            {
                CheckRules();
                CheckRibbon();
                CheckOverlappingRules();
                CheckPassageRuleWindows();
                CheckContext();
                CheckConnections();
                CheckSmallPaging();
                CheckDetails();
                CheckBranding();
                WriteReport(true, null);
                Console.WriteLine("PASS: " + checks.Count + " native UI checks; " + captures.Count + " control renders.");
                return 0;
            }
            catch (Exception error)
            {
                WriteReport(false, error.ToString());
                Console.Error.WriteLine(error);
                return 1;
            }
        }

        private static void AddRules()
        {
            string[] names = { "Keep unchanged", "Make snappier", "Reject changes", "Discard idea", "Shorten" };
            string[] colors = { "#74b9a4", "#e5be50", "#d86e6e", "#bc8fc7", "#668ddd" };
            string[] actions = { "protect", "prompt", "reject", "discard", "prompt" };
            for (int index = 0; index < names.Length; index++)
                rules.Add(new Dictionary<string, object>
                {
                    { "id", "rule-" + index }, { "label", names[index] }, { "color", colors[index] },
                    { "action", actions[index] }, { "prompt", index == 1 ? "Keep the meaning; improve the wording." : "" },
                    { "accessMode", "suggestions" }
                });
        }

        private static object Dispatch(string operation, IDictionary<string, object> arguments)
        {
            calls.Add(Tuple.Create(operation, arguments));
            if (operation == "state") return new Dictionary<string, object>
            {
                { "document", document }, { "rules", rules.ToArray() },
                { "marks", Enumerable.Range(0, 27).Select(index => new Dictionary<string, object>
                    { { "id", "mark-" + index }, { "ruleId", "rule-1" }, { "quote", "Marked passage " + index + " with complete original text, 😀 and e\u0301." }, { "prompt", "Passage instruction " + index }, { "stale", index == 3 } }).ToArray() },
                { "proposals", Enumerable.Range(0, 19).Select(index => new Dictionary<string, object>
                    { { "id", "proposal-" + index }, { "quote", "Original passage " + index }, { "replacement", "Revised passage " + index }, { "status", "pending" } }).ToArray() },
                { "connection", Connection() }
            };
            if (operation == "connection_status") return Connection();
            if (operation == "comment_author_save") { commentAuthorName = UiState.Text(arguments, "commentAuthorName").Trim(); return Connection(); }
            if (operation == "selection_preview" && selectionUnavailable) throw new InvalidOperationException("Select prose first.");
            if (operation == "selection_preview") return new Dictionary<string, object>
            {
                { "quote", "Selected passage 😀 with a composed e\u0301." }, { "from", 10 }, { "to", 48 }, { "snapshot", "selection-proof" }
            };
            if (operation == "context_preview") return Preview();
            if (operation == "share_context") document["shared"] = true;
            if (operation == "revoke_context" || operation == "connection_stop") document["shared"] = false;
            return new Dictionary<string, object>();
        }

        private static Dictionary<string, object> Connection()
        {
            return new Dictionary<string, object>
            {
                { "running", false }, { "status", "Access is stopped." }, { "lastError", "" }, { "configured", true }, { "commentAuthorName", commentAuthorName },
                { "publicMcpUrl", "https://example.invalid/synthetic-mcp-capability-do-not-use" }
            };
        }

        private static Dictionary<string, object> Preview()
        {
            string text = String.Join("\r\n\r\n", Enumerable.Repeat("Synthetic complete manuscript context. Unmarked text remains read-only. Highlighted passages may have prompt instructions, protection or proposal-only approval.", 12));
            return new Dictionary<string, object>
            {
                { "documentId", document["id"] }, { "text", text }, { "title", document["title"] }, { "characters", text.Length }, { "complete", true }, { "snapshot", "context-proof" }
            };
        }

        private static void CheckRules()
        {
            using (RuleDialog form = new RuleDialog(Dispatch, rules[1]))
            {
                Open(form);
                foreach (Size size in new[] { form.MinimumSize, new Size(760, 755), new Size(700, 950), new Size(1050, 720), new Size(1050, 950), form.MinimumSize })
                {
                    form.Size = size;
                    Settle(form);
                    CheckRuleGeometry(form, 4);
                }
                Capture(form, "color-rule-minimum");
                TextBox outlineInput = FindText(form, "Color");
                string originalColor = outlineInput.Text;
                CheckRuleOutline(form, originalColor);
                outlineInput.Text = "#d24b67";
                Settle(form);
                CheckRuleOutline(form, "#d24b67");
                outlineInput.Text = "#d2";
                Settle(form);
                CheckRuleOutline(form, "#d24b67");
                outlineInput.Text = originalColor;
                Settle(form);
                CheckRuleOutline(form, originalColor);
                Require(!Descendants(form).Any(control => control.AccessibleName == "Edit text directly"
                    || control.Text == "Suggestions · Track Changes" || control is GroupBox && control.Text == "Mode"),
                    "Color rules expose no editing-mode toggle or suggestion-mode choice.");
                CheckBox preserve = Descendants(form).OfType<CheckBox>().First(box => box.Text.StartsWith("Preserve text"));
                preserve.Checked = true;
                Require(!FindText(form, "Custom prompt").Enabled
                    && Descendants(form).OfType<CheckBox>().Where(box => box != preserve).All(box => !box.Enabled),
                    "Preserve text disables all wording presets and the custom prompt.");
                Capture(form, "color-rule-preserve");
                preserve.Checked = false;
                ComboBox function = Descendants(form).OfType<ComboBox>().Single();
                function.SelectedIndex = 1;
                Require(!FindText(form, "Custom prompt").Enabled, "Neutral colors expose no wording prompt.");
                function.SelectedIndex = 0;
                RadioButton shorten = Descendants(form).OfType<RadioButton>().First(box => box.Text == "Shorten");
                RadioButton detail = Descendants(form).OfType<RadioButton>().First(box => box.Text == "More detail");
                shorten.Checked = true;
                detail.Checked = true;
                Require(!shorten.Checked && detail.Checked, "Shorten and More detail are mutually exclusive.");
                shorten.Checked = true;
                Descendants(form).OfType<CheckBox>().First(box => box.Text == "Make snappier").Checked = true;
                Require(Descendants(form).OfType<CheckBox>().Any(box => box.Text == "Remove")
                    && Descendants(form).OfType<CheckBox>().Any(box => box.Text == "Get to the point"), "Requested remove and get-to-the-point presets are available.");
                FindText(form, "Custom prompt").Text = "Preserve measured results.";
                FindButton(form, "Save color rule").PerformClick();
                IDictionary<string, object> saved = calls.Last(call => call.Item1 == "save_rule").Item2;
                Require(UiState.Text(saved, "accessMode") == "direct" && UiState.Text(saved, "prompt").Contains("Shorten this passage")
                    && UiState.Text(saved, "prompt").Contains("Preserve measured results.") && UiState.Text(saved, "prompt").Contains("make the wording snappier")
                    && !UiState.Text(saved, "prompt").Contains("Add useful detail"), "Saving a legacy suggestion rule combines instructions with fixed direct editing.");
                Require(UiState.Text(saved, "documentId") == "synthetic-ui-document", "Color rule changes carry the displayed document identity.");
            }
            using (RuleDialog form = new RuleDialog(Dispatch, null))
            {
                Open(form);
                FindText(form, "Custom prompt").Text = "Keep the meaning.";
                FindButton(form, "Save color rule").PerformClick();
                var saved = calls.Last(call => call.Item1 == "save_rule").Item2;
                Require(UiState.Text(saved, "action") == "prompt" && UiState.Text(saved, "accessMode") == "direct"
                    && UiState.Text(saved, "prompt") == "Keep the meaning.",
                    "A new wording rule saves in direct edit mode without a mode choice.");
            }
            using (RuleDialog form = new RuleDialog(Dispatch, rules[1]))
            {
                Open(form);
                FindText(form, "Custom prompt").Text = "This cannot override protection.";
                Descendants(form).OfType<CheckBox>().First(box => box.Text.StartsWith("Preserve text")).Checked = true;
                FindButton(form, "Save color rule").PerformClick();
                var saved = calls.Last(call => call.Item1 == "save_rule").Item2;
                Require(UiState.Text(saved, "action") == "protect" && UiState.Text(saved, "prompt") == "" && UiState.Text(saved, "accessMode") == "suggestions",
                    "Saving Preserve text produces protection with no effective prompt or direct permission.");
            }
            foreach (string action in new[] { "neutral", "reject", "discard" })
            {
                var rule = new Dictionary<string, object>(rules[1]);
                rule["action"] = action;
                using (RuleDialog form = new RuleDialog(Dispatch, rule))
                {
                    Open(form);
                    FindButton(form, "Save color rule").PerformClick();
                    Require(UiState.Text(calls.Last(call => call.Item1 == "save_rule").Item2, "action") == action,
                        "Opening and saving retains the existing " + action + " function.");
                }
            }
            using (RuleDialog form = new RuleDialog(Dispatch, rules[1]))
            {
                Open(form);
                form.Font = new Font(form.Font.FontFamily, form.Font.SizeInPoints * 2);
                form.Size = new Size(1600, 884);
                Settle(form);
                CheckRuleGeometry(form, 3);
                Capture(form, "color-rule-enlarged-text");
            }
            using (RuleDialog form = new RuleDialog(Dispatch, rules[1], "synthetic-ui-document", "mark-0"))
            {
                Open(form);
                form.Size = form.MinimumSize;
                Settle(form);
                CheckRuleGeometry(form, 4);
                Capture(form, "color-rule-context-marker");
                form.Font = new Font(form.Font.FontFamily, form.Font.SizeInPoints * 2);
                form.Size = new Size(1600, 884);
                Settle(form);
                CheckRuleGeometry(form, 3);
                Capture(form, "color-rule-context-marker-enlarged");
                form.Font = SystemFonts.MessageBoxFont;
                form.Size = new Size(760, 755);
                Settle(form);
                CheckRuleGeometry(form, 4);
                FindButton(form, "Remove this marker").PerformClick();
                var removed = calls.Last(call => call.Item1 == "remove_mark").Item2;
                Require(UiState.Text(removed, "id") == "mark-0" && UiState.Text(removed, "documentId") == "synthetic-ui-document",
                    "In-text rule dialog removes only the exact selected document marker.");
            }
            using (Font enlarged = new Font(SystemFonts.MessageBoxFont.FontFamily, 11f))
            using (RuleDialog form = new RuleDialog(Dispatch, rules[0]))
            {
                Open(form);
                form.Font = enlarged;
                form.Size = new Size(760, 755);
                Settle(form);
                CheckRuleGeometry(form, 4);
                Capture(form, "color-rule-reported-size");
            }
            using (RuleDialog form = new RuleDialog(Dispatch, rules[1]))
            {
                Open(form);
                TextBox draft = FindText(form, "Custom prompt");
                draft.Text = "Keep this unsaved draft exactly.";
                int saves = calls.Count(call => call.Item1 == "save_rule");
                form.MaximumSize = new Size(760, 720);
                form.Font = new Font(form.Font.FontFamily, form.Font.SizeInPoints * 2);
                Settle(form);
                CheckRuleGeometry(form, 3);
                Require(Descendants(form).Any(control => control.Visible && control.AccessibleName == "More window space needed"),
                    "A bounded window shows a readable no-fit state instead of overlapping editing controls.");
                Require(draft.Text == "Keep this unsaved draft exactly." && calls.Count(call => call.Item1 == "save_rule") == saves,
                    "The no-fit state retains the exact unsaved draft without assigning a rule.");
                Capture(form, "color-rule-no-fit");
                form.MaximumSize = Size.Empty;
                form.Font = SystemFonts.MessageBoxFont;
                form.Size = new Size(760, 755);
                Settle(form);
                CheckRuleGeometry(form, 4);
                Require(draft.Visible && draft.Text == "Keep this unsaved draft exactly."
                    && !Descendants(form).Any(control => control.Visible && control.AccessibleName == "More window space needed"),
                    "Removing the space constraint restores the editor and preserves the unsaved draft.");
                Require(calls.Count(call => call.Item1 == "save_rule") == saves,
                    "No-fit recovery does not save or change the rule.");
            }
        }

        private static void CheckRuleGeometry(Form form, int promptLines)
        {
            CheckBounds(form);
            CheckChoiceText(form);
            if (Descendants(form).Any(control => control.Visible && control.AccessibleName == "More window space needed"))
            {
                Require(!FindText(form, "Custom prompt").Visible && !FindButton(form, "Save color rule").Visible
                    && form.AcceptButton == null && form.CancelButton != null,
                    "The no-fit state hides editing controls, disables Enter-to-save and keeps a close action.");
                return;
            }
            CheckPromptSpace(form, promptLines);
        }

        private static void CheckRibbon()
        {
            int before = calls.Count(call => call.Item1 == "mark_selection");
            using (WordRibbon ribbon = new WordRibbon(Dispatch))
            {
                ribbon.Observe(new Dictionary<string, object> { { "rules", Enumerable.Range(0, 24).Select(index => new Dictionary<string, object> {
                    { "id", "slot-" + index }, { "label", "Color " + (index + 1) }, { "color", "#ffffff" }
                }).ToArray() } });
                Require(ribbon.Enabled(new UiRibbonControl { Tag = "23" }) && ribbon.Label(new UiRibbonControl { Tag = "23" }).Contains("Color 24"),
                    "The twenty-fourth color is directly available without paging.");
                ribbon.Observe(Dispatch("state", new Dictionary<string, object>()));
                ribbon.Highlight(new UiRibbonControl { Tag = "1" });
                Require(calls.Count(call => call.Item1 == "mark_selection") == before, "A first swatch click waits for the double-click decision.");
                typeof(WordRibbon).GetMethod("CommitPendingHighlight", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(ribbon, null);
                var added = calls.Last(call => call.Item1 == "mark_selection").Item2;
                Require(calls.Count(call => call.Item1 == "mark_selection") == before + 1
                    && UiState.Text(added, "ruleId") == "rule-1" && UiState.Text(added, "snapshot") == "selection-proof"
                    && UiState.Text(added, "documentId") == "synthetic-ui-document" && UiState.Text(added, "prompt") == "",
                    "A single swatch click marks the captured selection directly without a confirmation window.");
                foreach (bool unavailable in new[] { false, true })
                {
                    bool opened = false;
                    selectionUnavailable = unavailable;
                    try
                    {
                        ribbon.Highlight(new UiRibbonControl { Tag = "1" });
                        ribbon.Highlight(new UiRibbonControl { Tag = "1" });
                        var form = Application.OpenForms.OfType<RuleDialog>().SingleOrDefault();
                        opened = form != null && form.Visible && !form.Modal;
                        if (form != null) form.Close();
                    }
                    finally { selectionUnavailable = false; }
                    Require(opened && calls.Count(call => call.Item1 == "mark_selection") == before + 1,
                        "Double-click opens only that color rule without marking, with " + (unavailable ? "no selected text." : "a selected passage."));
                }
            }
        }

        private static void CheckOverlappingRules()
        {
            var state = new Dictionary<string, object> {
                { "documentId", "synthetic-ui-document" }, { "rules", rules.ToArray() },
                { "choices", new[] {
                    new Dictionary<string, object> { { "id", "overlap-1" }, { "ruleId", "rule-0" }, { "label", "Keep unchanged" }, { "quote", "Protected words" } },
                    new Dictionary<string, object> { { "id", "overlap-2" }, { "ruleId", "rule-1" }, { "label", "Make snappier" }, { "quote", "Complete paragraph containing protected words" } }
                } }
            };
            using (HighlightedRulesDialog form = new HighlightedRulesDialog(Dispatch, state))
            {
                Open(form);
                form.Size = form.MinimumSize;
                Settle(form);
                CheckBounds(form);
                CheckPagedLists(form);
                Require(Descendants(form).OfType<ListBox>().Single().Items.Count == 2, "Overlapping colors remain an explicit choice.");
                Capture(form, "overlapping-color-choice");
                CheckEnlarged(form, "overlapping-color-enlarged");
            }
        }

        private static void CheckPassageRuleWindows()
        {
            string documentId = "synthetic-ui-document";
            var markers = new[] {
                new Dictionary<string, object> { { "id", "green" }, { "ruleId", "rule-0" }, { "label", "Keep unchanged" }, { "quote", "Green passage" } },
                new Dictionary<string, object> { { "id", "yellow" }, { "ruleId", "rule-1" }, { "label", "Make snappier" }, { "quote", "Changed yellow passage" }, { "stale", true } },
                new Dictionary<string, object> { { "id", "yellow-two" }, { "ruleId", "rule-1" }, { "label", "Make snappier" }, { "quote", "Another yellow passage" } }
            };
            IDictionary<string, object>[] selected = { markers[0] };
            var effects = new List<Tuple<string, IDictionary<string, object>>>();
            Func<string, IDictionary<string, object>, object> dispatch = (operation, args) => {
                if (args.ContainsKey("documentId") && UiState.Text(args, "documentId") != documentId)
                    throw new InvalidOperationException("The active document changed.");
                if (operation == "selection_rules") return new Dictionary<string, object> {
                    { "documentId", documentId }, { "choices", selected }, { "rules", rules.ToArray() }
                };
                if (operation != "state") effects.Add(Tuple.Create(operation, args));
                return new Dictionary<string, object> { { "document", document }, { "rules", rules.ToArray() }, { "marks", markers } };
            };
            using (var owner = new Form())
            using (var windows = new ColorRuleWindows(dispatch, () => owner))
            {
                Open(owner);
                Require(windows.OpenAtSelection(), "A green native marker resolves to its color-rule window.");
                RuleDialog green = Application.OpenForms.OfType<RuleDialog>().Single();
                selected = new IDictionary<string, object>[] { markers[1] };
                Require(windows.OpenAtSelection(), "A changed yellow native marker remains available for local rule navigation.");
                RuleDialog yellow = Application.OpenForms.OfType<RuleDialog>().Single(form => form != green);
                Settle(green); Settle(yellow);
                Require(owner.Enabled && !green.Modal && !yellow.Modal && green.Visible && yellow.Visible,
                    "Two different passage rules stay modeless so the document remains available for another double-click.");
                Require(green.Owner == owner && yellow.Owner == owner, "Both passage-rule windows belong to the native document owner.");
                Require(FindText(green, "Color").Text == "#74b9a4" && FindText(yellow, "Color").Text == "#e5be50"
                    && green.Text.StartsWith("Keep unchanged") && yellow.Text.StartsWith("Make snappier"),
                    "Each selected passage opens its exact rule name and color rather than the previously selected rule.");
                CheckRuleOutline(green, "#74b9a4"); CheckRuleOutline(yellow, "#e5be50");
                Capture(green, "passage-rule-green"); Capture(yellow, "passage-rule-yellow");
                FindText(yellow, "Custom prompt").Text = "Keep this unsaved yellow draft.";
                windows.Open(rules[0], documentId, "green");
                windows.Open(rules[1], documentId, "yellow-two");
                Require(Application.OpenForms.OfType<RuleDialog>().Count() == 2
                    && Application.OpenForms.OfType<RuleDialog>().Contains(yellow)
                    && FindText(yellow, "Custom prompt").Text == "Keep this unsaved yellow draft.",
                    "Returning to an existing color reuses its window and preserves its unsaved draft.");
                Require(effects.Count == 0, "Opening, switching and updating marker context do not save rules or mutate markers.");
                FindButton(yellow, "Remove this marker").PerformClick();
                Require(effects.Count == 1 && effects[0].Item1 == "remove_mark" && UiState.Text(effects[0].Item2, "id") == "yellow-two",
                    "Reusing a color updates removal to the most recently selected exact native marker.");
                selected = new IDictionary<string, object>[] { markers[0], markers[1] };
                windows.OpenAtSelection();
                HighlightedRulesDialog chooser = Application.OpenForms.OfType<HighlightedRulesDialog>().Single();
                Require(!chooser.Modal && owner.Enabled && Application.OpenForms.OfType<RuleDialog>().Count() == 1,
                    "Overlapping colors require an explicit modeless choice without automatically opening another rule.");
                Descendants(chooser).OfType<ListBox>().Single().SelectedIndex = 1;
                FindButton(chooser, "Open selected color rule").PerformClick();
                yellow = Application.OpenForms.OfType<RuleDialog>().Single(form => form != green);
                Require(FindText(yellow, "Color").Text == "#e5be50", "Explicit overlap selection opens the chosen yellow rule.");
                FindText(yellow, "Custom prompt").Text = "An explicitly saved yellow instruction.";
                FindButton(yellow, "Save color rule").PerformClick();
                var saved = effects.Single(effect => effect.Item1 == "save_rule").Item2;
                Require(UiState.Text(saved, "id") == "rule-1" && UiState.Text(saved, "color") == "#e5be50"
                    && UiState.Text(saved, "accessMode") == "direct" && UiState.Text(saved, "prompt").Contains("explicitly saved yellow"),
                    "Explicit Save applies only the selected color's exact prompt and direct mode.");
                FindButton(green, "Cancel").PerformClick();
                Require(effects.Count == 2, "Canceling the separate green rule adds no rule or marker mutation.");
                windows.Open(rules[0], documentId, "green");
                Require(Application.OpenForms.OfType<RuleDialog>().Count() == 1,
                    "A live rule window remains for the presenter shutdown cleanup check.");
                documentId = "another-document";
                bool denied = false;
                try { windows.Open(rules[1], "synthetic-ui-document", "yellow"); } catch (InvalidOperationException) { denied = true; }
                Require(denied, "An old document's rule cannot open against a newly active document.");
            }
            Require(!Application.OpenForms.OfType<RuleDialog>().Any() && !Application.OpenForms.OfType<HighlightedRulesDialog>().Any(),
                "Disposing the rule presenter closes all of its owned rule and chooser windows.");
        }

        private static void CheckContext()
        {
            int before = calls.Count(call => call.Item1 == "share_context");
            using (ContextConsentDialog form = new ContextConsentDialog(Dispatch, Preview()))
            {
                Open(form);
                form.Size = form.MinimumSize;
                Settle(form);
                CheckBounds(form);
                Capture(form, "context-consent");
                CheckEnlarged(form, "context-consent-enlarged");
                Require(calls.Count(call => call.Item1 == "share_context") == before, "Previewing complete context does not share it.");
                FindButton(form, "Share context").PerformClick();
                Require(UiState.Text(calls.Last(call => call.Item1 == "share_context").Item2, "snapshot") == "context-proof",
                    "Context consent returns the exact preview snapshot for native revalidation.");
                Require(UiState.Text(calls.Last(call => call.Item1 == "share_context").Item2, "documentId") == "synthetic-ui-document",
                    "Context sharing is bound to the document shown in the preview.");
            }
            IDictionary<string, object> incomplete = Preview();
            incomplete["complete"] = false;
            bool denied = false;
            try { using (ContextConsentDialog form = new ContextConsentDialog(Dispatch, incomplete)) { } }
            catch (InvalidOperationException) { denied = true; }
            Require(denied, "Incomplete context cannot expose a Share action.");
        }

        private static void CheckRuleOutline(RuleDialog form, string color)
        {
            TableLayoutPanel shell = form.Controls.OfType<TableLayoutPanel>().Single(panel =>
                panel.Controls.OfType<GroupBox>().Any(group => group.Text == "Custom prompt"));
            using (Bitmap bitmap = new Bitmap(shell.Width, shell.Height))
            {
                shell.DrawToBitmap(bitmap, shell.ClientRectangle);
                int expected = ColorTranslator.FromHtml(color).ToArgb();
                Require(bitmap.GetPixel(0, shell.Height / 2).ToArgb() == expected
                    && bitmap.GetPixel(15, shell.Height / 2).ToArgb() == expected
                    && bitmap.GetPixel(shell.Width - 1, shell.Height / 2).ToArgb() == expected
                    && bitmap.GetPixel(shell.Width / 2, 0).ToArgb() == expected
                    && bitmap.GetPixel(shell.Width / 2, shell.Height - 1).ToArgb() == expected,
                    "The complete rule-dialog outline matches the selected color; incomplete input retains its last valid color.");
            }
            Color selected = ColorTranslator.FromHtml(color);
            Color tint = Color.FromArgb((selected.R + 251) / 2, (selected.G + 248) / 2, (selected.B + 239) / 2);
            Require(shell.BackColor == tint && form.BackColor == tint && form.Opacity == 1,
                "The rule body uses a 50-percent color blend while the window and its text stay opaque.");
            int caption;
            if (DwmGetWindowAttribute(form.Handle, 35, out caption, sizeof(int)) == 0)
                Require(caption == (selected.R | selected.G << 8 | selected.B << 16),
                    "Windows stores the selected color on the native title bar.");
        }

        private static void CheckConnections()
        {
            var openedLinks = new List<string>();
            bool sharedBefore = UiState.Flag(document, "shared");
            int effectsBefore = calls.Count(call => call.Item1 != "state" && call.Item1 != "connection_status");
            using (ConnectionsDialog form = new ConnectionsDialog(Dispatch, openedLinks.Add))
            {
                Open(form);
                form.Size = form.MinimumSize;
                TabControl tabs = Descendants(form).OfType<TabControl>().First();
                for (int index = 0; index < tabs.TabPages.Count; index++)
                {
                    tabs.SelectedIndex = index;
                    foreach (Size size in new[] { form.MinimumSize, new Size(640, 950), new Size(1050, 700), new Size(1050, 950), form.MinimumSize })
                    {
                        form.Size = size;
                        Settle(form);
                        CheckBounds(form);
                        CheckConnectionGuidance(form);
                    }
                    Capture(form, "connections-" + index);
                }
                Require(tabs.TabPages.Count == 2 && tabs.TabPages[0].Text == "Overview" && tabs.TabPages[1].Text == "Setup", "Connections needs only Overview and Setup; connecting grants document reading.");
                tabs.SelectedIndex = 1;
                Settle(form);
                LinkLabel[] links = Descendants(tabs.SelectedTab).OfType<LinkLabel>().ToArray();
                Require(links.Length == 2 && links.All(link => link.TabStop && link.CanSelect && link.Links.Count == 1),
                    "Two native setup links are keyboard-focusable and expose one complete target each.");
                Require((string)links[0].Links[0].LinkData == "https://myzrok.io"
                    && (string)links[1].Links[0].LinkData == "https://netfoundry.io/docs/zrok/get-started/get-token/",
                    "Setup links target only the fixed public zrok account and official token instructions.");
                TextBox token = FindText(form, "zrok enrollment token");
                token.Text = "synthetic-account-token-not-a-credential";
                links[0].Focus();
                bool activated = (bool)typeof(LinkLabel).GetMethod("ProcessDialogKey", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(links[0], new object[] { Keys.Enter });
                Require(activated && openedLinks.SequenceEqual(new[] { "https://myzrok.io" }),
                    "Enter on the focused account link opens exactly its fixed public address.");
                form.SelectNextControl(links[0], true, true, true, false);
                Require(links[1].Focused, "Tab reaches the token instructions after the account link.");
                typeof(LinkLabel).GetMethod("OnLinkClicked", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(links[1], new object[] { new LinkLabelLinkClickedEventArgs(links[1].Links[0]) });
                Require(openedLinks.SequenceEqual(new[] { "https://myzrok.io", "https://netfoundry.io/docs/zrok/get-started/get-token/" }),
                    "Clicking token help opens its fixed official instructions without a token or document value.");
                Require(token.Text == "synthetic-account-token-not-a-credential", "Opening setup instructions preserves the unsubmitted account-token input.");
                form.SelectNextControl(links[1], true, true, true, false);
                Require(FindText(form, "zrok enrollment token").Focused, "Keyboard order continues from documentation to the masked token input.");
                Require(UiState.Flag(document, "shared") == sharedBefore
                    && calls.Count(call => call.Item1 != "state" && call.Item1 != "connection_status") == effectsBefore,
                    "Opening Connections and its setup links neither starts a connection nor grants document access.");
                TextBox author = FindText(form, "Edit comment author name");
                Require(author.Text == "wordplai" && author.MaxLength == 80, "Setup loads the saved comment name with a bounded input.");
                author.Text = "  Élodie – Editor 😀  ";
                FindButton(form, "Refresh").PerformClick();
                Require(author.Text == "  Élodie – Editor 😀  " && commentAuthorName == "wordplai", "Refresh preserves an unsaved comment-name draft without committing it.");
                FindButton(form, "Save comment name").PerformClick();
                Require(author.Text == "Élodie – Editor 😀" && commentAuthorName == author.Text
                    && calls.Last(call => call.Item1 == "comment_author_save").Item2.Count == 1,
                    "Explicit Save submits only the comment name and shows its normalized saved value.");
                author.Focus();
                form.SelectNextControl(author, true, true, true, false);
                Require(FindButton(form, "Save comment name").Focused, "Native keyboard traversal reaches Save immediately after the comment name.");
                author.Text = "Unsubmitted author draft";
                form.Font = new Font(form.Font.FontFamily, form.Font.SizeInPoints * 2);
                form.Size = new Size(1600, 884);
                for (int index = 0; index < tabs.TabPages.Count; index++)
                {
                    tabs.SelectedIndex = index;
                    foreach (Size size in new[] { new Size(1044, 788), new Size(1600, 884), new Size(1044, 788) })
                    {
                        form.Size = size;
                        Settle(form);
                        CheckBounds(form);
                        CheckConnectionGuidance(form);
                        CheckAppearance(form, SystemInformation.HighContrast);
                    }
                    Capture(form, "connections-enlarged-" + index);
                }
                Require(FindText(form, "zrok enrollment token").UseSystemPasswordChar, "Enrollment token is masked by default.");

            }
            using (ConnectionsDialog reopened = new ConnectionsDialog(Dispatch, openedLinks.Add))
            {
                Open(reopened);
                Require(FindText(reopened, "Edit comment author name").Text == "Élodie – Editor 😀",
                    "Closing and reopening Setup discards the unsubmitted author draft and restores the saved name.");
            }
        }

        private static void CheckConnectionGuidance(ConnectionsDialog form)
        {
            TabPage page = Descendants(form).OfType<TabControl>().Single().SelectedTab;
            foreach (Label label in Descendants(page).OfType<Label>().Where(label => label.Visible))
            {
                Size required = TextRenderer.MeasureText(label.Text, label.Font,
                    new Size(Math.Max(1, label.ClientSize.Width), Int32.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
                Require(required.Height <= label.ClientSize.Height + 1,
                    "Connection guidance fits its native text allocation: " + label.Text);
            }
            Require(!page.AutoScroll && Descendants(page).OfType<ScrollableControl>().All(control => !control.AutoScroll),
                "Connection instructions fit without an internal panel scrollbar.");
            if (page.Text == "Overview")
            {
                string text = String.Join(" ", Descendants(page).OfType<Label>().Select(label => label.Text));
                Require(text.Contains("MCP address in your AI client") && text.Contains("Choose a document")
                    && text.Contains("unsaved text") && text.Contains("Color rules") && text.Contains("apply directly")
                    && text.Contains("rationale comment") && text.Contains("green protection") && text.Contains("Stop / revoke ends access")
                    && !text.Contains("tracked suggestions with comments") && !text.Contains("whole active document"),
                    "Overview explains explicit document choice, direct edits with rationale and protection, and stopping access.");
            }
            if (page.Text == "Setup")
            {
                Label[] steps = Descendants(page).OfType<Label>().Where(label => label.Text.Length > 2
                    && label.Text[0] >= '1' && label.Text[0] <= '4' && label.Text[1] == '.').ToArray();
                Require(steps.Length == 4 && steps.Select(label => label.Text[0]).SequenceEqual(new[] { '1', '2', '3', '4' })
                    && steps.Zip(steps.Skip(1), (first, next) => first.PointToScreen(Point.Empty).Y < next.PointToScreen(Point.Empty).Y).All(value => value),
                    "Four setup steps retain their reading order from account creation through the AI connector.");
                Require(steps[1].Text.Contains("API console") && steps[1].Text.Contains("Getting started") && steps[1].Text.Contains("account token")
                    && steps[2].Text.Contains("Configure connection") && steps[3].Text.Contains("Start access") && steps[3].Text.Contains("MCP address"),
                    "Setup explains where to obtain the account token and how to configure and start the bundled connection.");
                Require(Descendants(page).OfType<LinkLabel>().All(link => link.LinkBehavior == LinkBehavior.AlwaysUnderline && link.LinkColor == link.ForeColor
                    && link.ActiveLinkColor == link.ForeColor && link.VisitedLinkColor == link.ForeColor),
                    "Setup links retain visible underlines and the active native surface text palette.");
            }

        }

        private static void CheckDetails()
        {
            using (TextDetailDialog form = new TextDetailDialog("Complete native status", "A complete read-only native detail value. Long content retains its own editor scroll behavior."))
            {
                Open(form);
                form.Size = form.MinimumSize;
                Settle(form);
                CheckBounds(form);
                Capture(form, "detail-minimum");
                CheckEnlarged(form, "detail-enlarged");
            }
        }

        private static void CheckBranding()
        {
            foreach (string resource in new[] { "wrAIter.Icon.ico", "wrAIter.Ribbon.png" })
            {
                string file = resource.EndsWith("ico") ? "icon.ico" : "icon16.png";
                using (var stream = typeof(AppBranding).Assembly.GetManifestResourceStream(resource))
                using (var bytes = new MemoryStream())
                {
                    stream.CopyTo(bytes);
                    Require(bytes.ToArray().SequenceEqual(File.ReadAllBytes(Path.Combine(assetDirectory, file))),
                        "The embedded " + resource + " bytes exactly match the canonical export.");
                }
            }
            using (Bitmap image = AppBranding.DialogImage(32))
                Require(image.Size == new Size(32, 32) && image.GetPixel(0, 0).A == 0, "The dialog uses a square transparent 32px frame from the canonical ICO.");
            using (Button swatch = new Button { BackColor = Color.FromArgb(116, 185, 164) })
            {
                AppBranding.ApplyPalette(swatch, false);
                Require(swatch.BackColor == Color.FromArgb(116, 185, 164), "Brand styling preserves a functional highlight swatch color.");
                AppBranding.ApplyPalette(swatch, true);
                Require(swatch.BackColor == Color.FromArgb(116, 185, 164), "System contrast fallback preserves a functional highlight swatch color.");
            }
            using (TextDetailDialog form = new TextDetailDialog("System palette fallback", "Synthetic style check; Windows high contrast is not changed."))
            {
                Open(form);
                form.ApplyAppearance(true);
                Settle(form);
                CheckAppearance(form, true);
                CheckBounds(form);
                Capture(form, "system-palette-simulated");
                form.ApplyAppearance(false);
                Settle(form);
                CheckAppearance(form, false);
                Capture(form, "brand-palette-restored");
            }
            for (int index = 0; index < 3; index++) CloseBrandedDetail();
            CollectNativeWrappers();
            uint gdiBefore = GetGuiResources(Process.GetCurrentProcess().Handle, 0);
            uint userBefore = GetGuiResources(Process.GetCurrentProcess().Handle, 1);
            for (int index = 0; index < 20; index++) CloseBrandedDetail();
            CollectNativeWrappers();
            checks.Add("Native resource counts: GDI " + gdiBefore + " → " + GetGuiResources(Process.GetCurrentProcess().Handle, 0)
                + "; USER " + userBefore + " → " + GetGuiResources(Process.GetCurrentProcess().Handle, 1));
            Require(GetGuiResources(Process.GetCurrentProcess().Handle, 0) <= gdiBefore + 2,
                "Twenty dialog open/close cycles retain no GDI handle growth after framework font wrappers are finalized.");
            Require(GetGuiResources(Process.GetCurrentProcess().Handle, 1) <= userBefore + 2,
                "Twenty dialog open/close cycles do not accumulate native window/icon handles.");
        }

        private static void CollectNativeWrappers()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Application.DoEvents();
        }

        private static void CloseBrandedDetail()
        {
            Image ownedImage;
            IntPtr ownedFont;
            using (TextDetailDialog form = new TextDetailDialog("Logo lifetime", "Synthetic disposal check."))
            {
                Open(form);
                ownedImage = Descendants(form).OfType<PictureBox>().Single(box => box.AccessibleName == "wordplai").Image;
                ownedFont = (IntPtr)typeof(Control).GetProperty("FontHandle", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form, null);
                form.Close();
            }
            bool disposed = false;
            try { int width = ownedImage.Width; }
            catch (ArgumentException) { disposed = true; }
            catch (ObjectDisposedException) { disposed = true; }
            Require(disposed, "Closing a dialog disposes its owned logo bitmap.");
            Require(GetObjectType(ownedFont) == 0, "Closing a dialog releases its allocated native font handle.");
        }

        private static void CheckAppearance(NativeDialog form, bool highContrast)
        {
            Color surface = highContrast ? SystemColors.Control : Color.FromArgb(251, 248, 239);
            Color text = highContrast ? SystemColors.ControlText : Color.FromArgb(27, 53, 48);
            if (form is RuleDialog && !highContrast)
            {
                Color selected = ColorTranslator.FromHtml(FindText(form, "Color").Text);
                surface = Color.FromArgb((selected.R + 251) / 2, (selected.G + 248) / 2, (selected.B + 239) / 2);
                text = Color.Black;
            }
            Require(form.BackColor == surface && form.ForeColor == text, "The shared native window uses the " + (highContrast ? "system contrast fallback." : "approved cream and deep-green palette."));
            Require(Descendants(form).OfType<Label>().All(label => label.BackColor == surface && label.ForeColor == text),
                "Native body labels follow the selected surface/text palette.");
            PictureBox[] marks = Descendants(form).OfType<PictureBox>().Where(box => box.AccessibleName == "wordplai").ToArray();
            Require(marks.Length == 1 && marks[0].Visible && marks[0].Width == marks[0].Height && marks[0].Width >= 24
                && !marks[0].TabStop && marks[0].Image != null, "One visible square logo stays outside the keyboard order.");
            if (form is RuleDialog)
            {
                TableLayoutPanel header = (TableLayoutPanel)marks[0].Parent;
                TableLayoutPanel shell = (TableLayoutPanel)header.Parent;
                Require(marks[0].Width >= 64 && shell.GetRow(header) == 0 && header.GetColumn(marks[0]) == 0,
                    "The prominent color-rule logo occupies the top-left header slot rather than the footer.");
                TableLayoutPanel presets = Descendants(form).OfType<GroupBox>().Single(group => group.Text.StartsWith("Wording instructions"))
                    .Controls.OfType<TableLayoutPanel>().Single();
                Require(presets.ColumnCount == 3 && presets.RowCount == 4, "All ten prompt suggestions use exactly three equal columns.");
                Require(!Descendants(form).OfType<GroupBox>().Any(group => group.Text == "Mode"),
                    "The color-rule layout reserves no group for an editing-mode choice.");
            }
            Require(SendMessage(form.Handle, 0x007f, new IntPtr(1), IntPtr.Zero) != IntPtr.Zero
                && SendMessage(form.Handle, 0x007f, IntPtr.Zero, IntPtr.Zero) != IntPtr.Zero, "The dialog retains its own large and small native icon handles.");
            Require(!form.AutoScroll && Descendants(form).OfType<ScrollableControl>().All(control => !control.AutoScroll),
                "Brand styling adds no internal panel scrollbar.");
            Require(Descendants(form).OfType<Button>().All(button => button.UseVisualStyleBackColor), "Actions retain native Windows rendering and focus behavior.");
        }

        private static void CheckEnlarged(NativeDialog form, string capture)
        {
            form.Font = new Font(form.Font.FontFamily, form.Font.SizeInPoints * 2);
            form.Size = new Size(1600, 884);
            Settle(form);
            CheckBounds(form);
            CheckChoiceText(form);
            CheckAppearance(form, SystemInformation.HighContrast);
            Capture(form, capture);
        }

        private static void CheckSmallPaging()
        {
            using (Form form = new Form { ClientSize = new Size(350, 95) })
            using (PagedRecordList records = new PagedRecordList("Small list", record => UiState.Text(record, "quote")))
            {
                form.Controls.Add(records);
                Open(form);
                records.UpdateRecords(Enumerable.Range(0, 40).Select(index => (IDictionary<string, object>)new Dictionary<string, object>
                    { { "id", "small-" + index }, { "quote", "Passage " + index } }).ToList());
                Settle(form);
                CheckPagedLists(form);
                string first = UiState.Text(records.Selected, "id");
                FindButton(form, "Next").PerformClick();
                Require(UiState.Text(records.Selected, "id") != first, "Paging reaches records beyond the bounded visible list.");
                string selected = UiState.Text(records.Selected, "id");
                form.ClientSize = new Size(430, 180);
                Settle(form);
                CheckPagedLists(form);
                Require(UiState.Text(records.Selected, "id") == selected, "Resizing the native page preserves the exact selected record.");
                Font original = form.Font;
                using (Font larger = new Font(original.FontFamily, original.SizeInPoints + 3, original.Style))
                {
                    form.Font = larger;
                    Settle(form);
                    CheckPagedLists(form);
                    Require(UiState.Text(records.Selected, "id") == selected, "Changing native list font preserves the exact selected record without duplicate rows.");
                    form.Font = original;
                    Settle(form);
                    CheckPagedLists(form);
                }
            }
        }

        private static void Open(Form form)
        {
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(-5000, -5000);
            form.ShowInTaskbar = false;
            form.Show();
            Settle(form);
            CheckPagedLists(form);
            NativeDialog branded = form as NativeDialog;
            if (branded != null)
            {
                CheckAppearance(branded, SystemInformation.HighContrast);
                branded.ApplyAppearance(true);
                CheckAppearance(branded, true);
                branded.ApplyAppearance(SystemInformation.HighContrast);
            }
        }

        private static void Settle(Form form)
        {
            form.PerformLayout();
            Application.DoEvents();
            form.PerformLayout();
            Application.DoEvents();
        }

        private static IEnumerable<Control> Descendants(Control root)
        {
            foreach (Control child in root.Controls)
            {
                yield return child;
                foreach (Control descendant in Descendants(child)) yield return descendant;
            }
        }

        private static Button FindButton(Control form, string label)
        {
            return Descendants(form).OfType<Button>().First(button => button.Text == label);
        }

        private static TextBox FindText(Control form, string name)
        {
            return Descendants(form).OfType<TextBox>().First(box => box.AccessibleName == name);
        }

        private static void CheckBounds(Form form)
        {
            foreach (Control control in Descendants(form).Where(control => control.Visible && !(control is TabPage)))
            {
                Rectangle bounds = control.Bounds;
                if (bounds.Left < 0 || bounds.Top < 0 || bounds.Right > control.Parent.ClientSize.Width + 1 || bounds.Bottom > control.Parent.ClientSize.Height + 1)
                {
                    Capture(form, "layout-failure");
                    throw new InvalidOperationException(form.Text + " at " + form.Size + " minimum " + form.MinimumSize + ": " + control.GetType().Name + " " + control.Text + " is outside its parent. " + bounds + " in " + control.Parent.ClientSize
                        + "; parent preferred " + control.Parent.PreferredSize + " minimum " + control.Parent.MinimumSize
                        + "; siblings " + String.Join("; ", control.Parent.Controls.Cast<Control>().Select(sibling => sibling.GetType().Name + " " + sibling.Bounds + " preferred " + sibling.PreferredSize + " margin " + sibling.Margin))
                        + (control.Parent is TableLayoutPanel ? "; rows " + String.Join(",", ((TableLayoutPanel)control.Parent).GetRowHeights()) : ""));
                }
                if (control is Button && (bounds.Width < 20 || bounds.Height < 20))
                    throw new InvalidOperationException(form.Text + ": action control is too small: " + control.Text);
                var table = control.Parent as TableLayoutPanel;
                if (table != null)
                {
                    TableLayoutPanelCellPosition cell = table.GetCellPosition(control);
                    int[] rowHeights = table.GetRowHeights();
                    int[] columnWidths = table.GetColumnWidths();
                    if (cell.Row >= 0 && cell.Column >= 0)
                    {
                        int top = table.Padding.Top + rowHeights.Take(cell.Row).Sum();
                        int left = table.Padding.Left + columnWidths.Take(cell.Column).Sum();
                        bool contained = bounds.Left >= left && bounds.Top >= top && bounds.Right <= left + columnWidths[cell.Column] + 1
                            && bounds.Bottom <= top + rowHeights[cell.Row] + 1;
                        if (!contained) Capture(form, "cell-layout-failure");
                        Require(contained,
                            "Native table cell contains " + control.GetType().Name + " " + control.Text + " " + bounds
                            + " in cell " + cell + " " + new Rectangle(left, top, columnWidths[cell.Column], rowHeights[cell.Row])
                            + (cell.Row < table.RowStyles.Count ? " row style " + table.RowStyles[cell.Row].SizeType + " " + table.RowStyles[cell.Row].Height : " implicit row style")
                            + (contained ? "" : "; geometry " + FailureGeometry(form, control, table)));
                    }
                }
            }
            checks.Add(form.Text + " native control bounds at " + form.Size + " / " + form.Font.SizeInPoints + "pt");
        }

        private static void CheckPromptSpace(Form form, int lines)
        {
            TextBox prompt = FindText(form, "Custom prompt");
            int minimumHeight = lines * prompt.Font.Height + 8;
            if (prompt.ClientSize.Height < minimumHeight) Capture(form, "custom-prompt-space-failure");
            Require(prompt.ClientSize.Height >= minimumHeight,
                "Custom prompt has room for " + lines + " readable lines: needs " + minimumHeight + " in " + prompt.ClientSize);
        }

        private static void CheckChoiceText(Form form)
        {
            foreach (ButtonBase choice in Descendants(form).OfType<ButtonBase>().Where(choice => choice.Visible && (choice is RadioButton || choice is CheckBox)))
            {
                Size required = TextRenderer.MeasureText(choice.Text, choice.Font, new Size(Math.Max(1, choice.ClientSize.Width - 24), Int32.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
                if (required.Height > choice.ClientSize.Height) Capture(form, "choice-text-failure");
                Require(required.Height <= choice.ClientSize.Height, "Choice text fits its measured native allocation: " + choice.Text + " needs " + required.Height + " in " + choice.ClientSize);
            }
        }

        private static void CheckPagedLists(Control form)
        {
            foreach (ListBox list in Descendants(form).OfType<ListBox>().Where(list => list.Visible))
            {
                int nativeCount = SendMessage(list.Handle, 0x018B, IntPtr.Zero, IntPtr.Zero).ToInt32();
                if (nativeCount != list.Items.Count) Capture((Form)form, "native-list-row-count-failure");
                Require(nativeCount == list.Items.Count, "Native list rows match the managed page: " + list.AccessibleName + " native=" + nativeCount + " managed=" + list.Items.Count);
                Require(nativeCount * list.ItemHeight <= list.ClientSize.Height,
                    "Paged list has no internal scrollbar: " + list.AccessibleName);
            }
        }

        private static void Capture(Form form, string name)
        {
            using (Bitmap bitmap = new Bitmap(form.Width, form.Height))
            {
                form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                string path = Path.Combine(output, name + ".png");
                bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                captures.Add(path);
            }
        }

        private static void Require(bool value, string message)
        {
            if (!value) throw new InvalidOperationException(message);
            checks.Add(message);
        }

        private static string FailureGeometry(Form form, Control control, TableLayoutPanel table)
        {
            return new JavaScriptSerializer().Serialize(new {
                window = new { form.Size, form.ClientSize, form.MinimumSize, form.MaximumSize,
                    form.DeviceDpi, font = form.Font.ToString(), form.AutoScaleDimensions, form.CurrentAutoScaleDimensions,
                    workingArea = Screen.FromControl(form).WorkingArea },
                parent = new { table.ClientSize, table.Padding, rows = table.GetRowHeights(),
                    styles = table.RowStyles.Cast<RowStyle>().Select(row => new { row.SizeType, row.Height }).ToArray() },
                children = table.Controls.Cast<Control>().Select(child => new {
                    type = child.GetType().Name, child.Text, child.Bounds, child.MinimumSize, child.PreferredSize,
                    child.Margin, font = child.Font.ToString(), cell = table.GetCellPosition(child) }).ToArray()
            });
        }

        private static void WriteReport(bool passed, string error)
        {
            File.WriteAllText(Path.Combine(output, "ui-report.json"), new JavaScriptSerializer().Serialize(
                new { passed = passed, checks = checks, captures = captures, error = error, scope = "Native WinForms component geometry and UI dispatch contract with synthetic document data; no installed Word or public connection qualification. System palette fallback is simulated without changing Windows high contrast settings." }));
        }
    }
}
