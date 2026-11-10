using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace UnderTree.WordAddin
{
    internal sealed class NativeHighlightsDialog : NativeDialog
    {
        private readonly Func<string, IDictionary<string, object>, object> dispatch;
        private readonly Label status;
        private readonly Timer refresh;
        private readonly Dictionary<int, Button> swatches = new Dictionary<int, Button>();
        private string documentId = "";
        private List<IDictionary<string, object>> colors = new List<IDictionary<string, object>>();

        internal NativeHighlightsDialog(Func<string, IDictionary<string, object>, object> dispatch)
            : base("Wordplai highlight rules", new Size(580, 460))
        {
            this.dispatch = dispatch;
            var palette = new FlowLayoutPanel { Dock = DockStyle.Fill, AccessibleName = "Colors in the active document" };
            for (int index = 1; index <= 16; index++)
            {
                int colorIndex = index;
                string name = WordNativeHighlights.Name(index);
                var button = NativeLayout.Button(name, delegate { NativeLayout.Attempt(this, () => Edit(colorIndex)); });
                button.Image = Swatch(WordNativeHighlights.Rgb(index));
                button.TextImageRelation = TextImageRelation.ImageBeforeText;
                button.Visible = false;
                swatches.Add(index, button);
                palette.Controls.Add(button);
            }
            status = NativeLayout.Label("");
            Button close = NativeLayout.Button("Close", delegate { Close(); });
            CancelButton = close;
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.Controls.Add(NativeLayout.Label("Mark text with Word's Home → Text Highlight Color. Choose a color here to assign its rule."), 0, 0);
            Shell.Controls.Add(status, 0, 1);
            Shell.Controls.Add(palette, 0, 2);
            Shell.Controls.Add(NativeLayout.Actions(close), 0, 3);
            refresh = new Timer { Interval = 1000 };
            refresh.Tick += delegate { if (Visible && Enabled) RefreshColors(); };
            Shown += delegate { RefreshColors(); refresh.Start(); };
        }

        internal void RefreshColors()
        {
            try
            {
                var current = UiState.Map(dispatch("native_highlight_colors", new Dictionary<string, object>()));
                documentId = UiState.Text(current, "documentId");
                colors = UiState.Records(current, "colors");
                status.Text = documentId.Length == 0 ? "Open a Word document first."
                    : colors.Count == 0 ? "This document has no text highlights." : "";
            }
            catch (InvalidOperationException error)
            {
                documentId = "";
                colors.Clear();
                status.Text = error.Message;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                documentId = "";
                colors.Clear();
                status.Text = "Word's highlight colors are temporarily unavailable.";
            }
            foreach (var swatch in swatches)
            {
                var color = FindColor(swatch.Key);
                swatch.Value.Visible = color != null;
                if (color != null)
                    swatch.Value.AccessibleDescription = RuleSummary(UiState.Text(color, "action"));
            }
        }

        private static string RuleSummary(string action)
        {
            if (action == "protect") return "Preserve text";
            if (action == "prompt") return "Editing instruction assigned";
            return "No rule assigned";
        }

        private IDictionary<string, object> FindColor(int index)
        {
            string key = index.ToString(CultureInfo.InvariantCulture);
            return colors.FirstOrDefault(color => UiState.Text(color, "index") == key);
        }

        private void Edit(int index)
        {
            string requestedDocument = documentId;
            RefreshColors();
            var color = FindColor(index);
            if (color == null || requestedDocument != documentId)
                throw new InvalidOperationException("The active document or its highlights changed. Choose a current color.");
            using (var rule = new NativeHighlightRuleDialog(color, (action, prompt) => dispatch("save_native_highlight_rule",
                new Dictionary<string, object> { { "documentId", requestedDocument }, { "index", index },
                    { "action", action }, { "prompt", prompt } })))
                rule.ShowDialog(this);
            RefreshColors();
        }

        private static Bitmap Swatch(string color)
        {
            var image = new Bitmap(16, 16);
            using (Graphics graphics = Graphics.FromImage(image))
            using (var brush = new SolidBrush(NativeLayout.ColorValue(color)))
            {
                graphics.FillRectangle(brush, 1, 1, 13, 13);
                graphics.DrawRectangle(SystemPens.ControlDarkDark, 1, 1, 13, 13);
            }
            return image;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (refresh != null) refresh.Dispose();
                foreach (Button swatch in swatches.Values)
                {
                    Image image = swatch.Image;
                    swatch.Image = null;
                    if (image != null) image.Dispose();
                }
            }
            base.Dispose(disposing);
        }
    }

    internal sealed class NativeHighlightRuleDialog : NativeDialog
    {
        private static readonly string[] Actions = { "neutral", "prompt", "protect" };

        internal NativeHighlightRuleDialog(IDictionary<string, object> color, Action<string, string> save)
            : base(UiState.Text(color, "name") + " highlight rule", new Size(420, 340))
        {
            var functions = new[] { "No rule", "Editing instruction", "Preserve text" }.Select(text =>
                new RadioButton { Text = text, AccessibleName = text, AutoSize = true, Dock = DockStyle.Fill }).ToArray();
            functions[Math.Max(0, Array.IndexOf(Actions, UiState.Text(color, "action")))].Checked = true;
            TextBox prompt = NativeLayout.TextBox("Editing instruction", true, false);
            prompt.MaxLength = 4000;
            prompt.Text = UiState.Text(color, "prompt");
            prompt.Enabled = functions[1].Checked;
            functions[1].CheckedChanged += delegate { prompt.Enabled = functions[1].Checked; };
            Button confirm = NativeLayout.Button("Save rule", delegate
            {
                NativeLayout.Attempt(this, delegate
                {
                    save(Actions[Array.FindIndex(functions, function => function.Checked)], prompt.Enabled ? prompt.Text.Trim() : "");
                    DialogResult = DialogResult.OK;
                    Close();
                });
            });
            Button cancel = NativeLayout.Button("Cancel", delegate { DialogResult = DialogResult.Cancel; Close(); });
            CancelButton = cancel;
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.Controls.Add(NativeLayout.Label("This rule applies to this color in the current document."), 0, 0);
            Shell.Controls.Add(NativeLayout.Stack(functions), 0, 1);
            Shell.Controls.Add(prompt, 0, 2);
            Shell.Controls.Add(NativeLayout.Actions(confirm, cancel), 0, 3);
        }
    }
}
