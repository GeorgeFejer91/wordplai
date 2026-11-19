using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace UnderTree.WordAddin
{
    internal sealed class WordRibbon : IDisposable
    {
        private readonly Func<string, IDictionary<string, object>, object> dispatch;
        private readonly ColorRuleWindows ruleWindows;
        private readonly Timer clickTimer = new Timer { Interval = SystemInformation.DoubleClickTime };
        private object ribbonUi;
        private string pendingRule;
        private string pendingDocument;
        private IDictionary<string, object> pendingSelection;
        private Exception pendingSelectionError;
        private List<IDictionary<string, object>> rules = new List<IDictionary<string, object>>();

        internal WordRibbon(Func<string, IDictionary<string, object>, object> dispatch) : this(dispatch, null) { }

        internal WordRibbon(Func<string, IDictionary<string, object>, object> dispatch, ColorRuleWindows ruleWindows)
        {
            this.dispatch = dispatch;
            this.ruleWindows = ruleWindows ?? new ColorRuleWindows(dispatch, () => null);
            clickTimer.Tick += delegate { Attempt(CommitPendingHighlight); };
        }

        internal void Load(object ui) { ribbonUi = ui; Invalidate(); }
        internal void Observe(object result)
        {
            var state = UiState.Map(result);
            if (!state.ContainsKey("rules")) return;
            rules = UiState.Records(state, "rules");
            Invalidate();
        }

        internal void Invalidate()
        {
            if (ribbonUi == null) return;
            try { ((dynamic)ribbonUi).Invalidate(); }
            catch (System.Runtime.InteropServices.COMException) { ribbonUi = null; }
        }

        internal static string GroupXml(string suffix)
        {
            return "<group id=\"wrAIterColors" + suffix + "\" label=\"wordplai\">" +
                "<button id=\"wrAIterNativeHighlights" + suffix + "\" size=\"normal\" label=\"Highlight rules\" imageMso=\"TextHighlightColorPicker\" screentip=\"Assign rules to Word text highlight colors\" keytip=\"WH\" onAction=\"RibbonOpen\"/>" +
                "<button id=\"wrAIterSelectionRules" + suffix + "\" size=\"normal\" imageMso=\"ReviewProtectDocument\" label=\"Legacy rules at selection\" screentip=\"Open a Wordplai rule from an earlier version at the selection\" keytip=\"WT\" onAction=\"RibbonSelectionRules\"/>" +
                "<button id=\"wrAIterSettings" + suffix + "\" tag=\"brand\" size=\"normal\" getImage=\"RibbonColorImage\" label=\"Connections\" screentip=\"Connections and document context sharing\" keytip=\"WE\" onAction=\"RibbonConnections\"/>" +
                "</group>";
        }

        private IDictionary<string, object> Rule(object control)
        {
            int slot = Int32.Parse(Convert.ToString(((dynamic)control).Tag, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            return slot >= 0 && slot < rules.Count && slot < 24 ? rules[slot] : null;
        }

        internal string Label(object control)
        {
            try { var rule = Rule(control); return rule == null ? "No color rule in this slot" : "Highlight selection: " + UiState.Text(rule, "label") + ". Double-click to configure this color."; }
            catch (Exception error) { return AppBranding.Name + ": " + error.Message; }
        }

        internal bool Enabled(object control)
        {
            try { return Rule(control) != null; }
            catch (Exception) { return false; }
        }

        internal object Image(object control)
        {
            if (Convert.ToString(((dynamic)control).Tag, CultureInfo.InvariantCulture) == "brand")
                using (var image = AppBranding.RibbonImage()) return RibbonPicture.Convert(image);
            Color color = SystemColors.Control;
            try { var rule = Rule(control); if (rule != null) color = ColorTranslator.FromHtml(UiState.Text(rule, "color")); }
            catch (Exception) { color = SystemColors.Control; }
            using (var image = new Bitmap(16, 16))
            using (Graphics graphics = Graphics.FromImage(image))
            using (var fill = new SolidBrush(color))
            {
                graphics.Clear(Color.Transparent);
                graphics.FillRectangle(fill, 1, 1, 13, 13);
                graphics.DrawRectangle(SystemPens.ControlDarkDark, 1, 1, 13, 13);
                return RibbonPicture.Convert(image);
            }
        }

        internal void Highlight(object control)
        {
            Attempt(delegate {
                var state = UiState.Map(dispatch("state", new Dictionary<string, object>()));
                Observe(state);
                var rule = Rule(control);
                if (rule == null) throw new InvalidOperationException("Choose an available color rule.");
                string id = UiState.Text(rule, "id");
                string documentId = UiState.Text(UiState.Map(UiState.Get(state, "document")), "id");
                if (clickTimer.Enabled && pendingRule == id && pendingDocument == documentId)
                {
                    CancelPendingHighlight();
                    ruleWindows.Open(rule, documentId, null);
                }
                else
                {
                    if (clickTimer.Enabled) CommitPendingHighlight();
                    pendingRule = id;
                    pendingDocument = documentId;
                    try { pendingSelection = UiState.Map(dispatch("selection_preview", UiState.Args("documentId", documentId))); }
                    catch (Exception error) { pendingSelectionError = error; }
                    clickTimer.Start();
                }
            });
        }

        private void CommitPendingHighlight()
        {
            var selection = pendingSelection;
            var error = pendingSelectionError;
            string ruleId = pendingRule;
            string documentId = pendingDocument;
            CancelPendingHighlight();
            if (error != null) throw error;
            if (selection == null || UiState.Text(selection, "quote").Length == 0)
                throw new InvalidOperationException("Select a passage to highlight, or double-click the color to configure it.");
            var arguments = new Dictionary<string, object>(selection);
            arguments["ruleId"] = ruleId;
            arguments["documentId"] = documentId;
            arguments["prompt"] = "";
            Observe(dispatch("mark_selection", arguments));
        }

        internal void CancelPendingHighlight()
        {
            clickTimer.Stop();
            pendingRule = null;
            pendingDocument = null;
            pendingSelection = null;
            pendingSelectionError = null;
        }

        internal void Configure()
        {
            CancelPendingHighlight();
            Attempt(delegate {
                var state = UiState.Map(dispatch("state", new Dictionary<string, object>()));
                Observe(state);
                if (rules.Count >= 24) throw new InvalidOperationException("The document already has 24 color rules. Double-click a color to edit it.");
                ruleWindows.Open(null, UiState.Text(UiState.Map(UiState.Get(state, "document")), "id"), null);
            });
        }

        private void Attempt(Action action)
        {
            try { action(); }
            catch (Exception error) { MessageBox.Show(error.Message, AppBranding.Name, MessageBoxButtons.OK, MessageBoxIcon.Information); }
            finally { Invalidate(); }
        }

        public void Dispose() { CancelPendingHighlight(); clickTimer.Dispose(); ruleWindows.Dispose(); }

        private sealed class RibbonPicture : AxHost
        {
            private RibbonPicture() : base("") { }
            internal static object Convert(Image image) { return GetIPictureDispFromPicture(image); }
        }
    }
}
