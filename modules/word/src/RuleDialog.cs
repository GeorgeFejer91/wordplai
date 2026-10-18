using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace UnderTree.WordAddin
{
    internal sealed class RuleDialog : NativeDialog
    {
        private readonly Func<string, IDictionary<string, object>, object> dispatch;
        private readonly string ruleId;
        private readonly string documentId;
        private string markerId;
        private readonly Button removeMarker;
        private readonly TextBox nameInput;
        private readonly TextBox colorInput;
        private readonly TextBox promptInput;
        private readonly CheckBox preserve;
        private readonly ComboBox function;
        private readonly TableLayoutPanel identity;
        private readonly GroupBox presetGroup;
        private readonly GroupBox lengthGroup;
        private readonly GroupBox promptGroup;
        private readonly TableLayoutPanel presetChoices;
        private readonly TableLayoutPanel functionRow;
        private readonly Dictionary<CheckBox, string> presets = new Dictionary<CheckBox, string>();
        private readonly Dictionary<RadioButton, string> lengths = new Dictionary<RadioButton, string>();
        private bool arrangingRows;
        private bool geometryQueued;
        private bool reflowingPresets;
        private int presetMinimumWidth;
        private Color outlineColor = ColorTranslator.FromHtml("#6c8fb5");
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);
        protected override int DialogLogoSize { get { return 64; } }
        protected override TableLayoutPanel DialogLogoPanel { get { return identity; } }
        protected override int DialogFrameInset { get { return 16; } }
        protected override Color DialogText { get { return Color.Black; } }
        protected override Color DialogSurface {
            get { return Color.FromArgb((outlineColor.R + 251) / 2, (outlineColor.G + 248) / 2, (outlineColor.B + 239) / 2); }
        }
        private static readonly string[] Actions = { "prompt", "neutral", "reject", "discard" };
        private static readonly Size RuleMinimum = new Size(700, 720);

        internal RuleDialog(Func<string, IDictionary<string, object>, object> dispatch, IDictionary<string, object> rule, string documentId = null, string markId = null)
            : base(rule == null ? "New color rule" : UiState.Text(rule, "label") + " · color rule", RuleMinimum)
        {
            this.dispatch = dispatch;
            if (documentId == null)
            {
                var state = UiState.Map(dispatch("state", new Dictionary<string, object>()));
                documentId = UiState.Text(UiState.Map(UiState.Get(state, "document")), "id");
            }
            this.documentId = documentId;
            rule = rule ?? new Dictionary<string, object>
            {
                { "label", "New color rule" }, { "color", "#6c8fb5" }, { "action", "prompt" },
                { "prompt", "" }, { "accessMode", "direct" }
            };
            ruleId = UiState.Text(rule, "id");
            if (ruleId.Length == 0) ruleId = Guid.NewGuid().ToString("N");
            nameInput = NativeLayout.TextBox("Rule name", false, false);
            nameInput.Text = UiState.Text(rule, "label");
            nameInput.MaxLength = 100;
            colorInput = NativeLayout.TextBox("Color", false, false);
            colorInput.Text = UiState.Text(rule, "color");
            colorInput.MaxLength = 7;
            colorInput.TextChanged += delegate { UpdateOutlineColor(); };
            Shell.Paint += delegate(object sender, PaintEventArgs args) {
                int width = (int)Math.Round(DialogFrameInset * args.Graphics.DpiX / 96);
                Color border = SystemInformation.HighContrast ? SystemColors.Highlight : outlineColor;
                ControlPaint.DrawBorder(args.Graphics, Shell.ClientRectangle,
                    border, width, ButtonBorderStyle.Solid, border, width, ButtonBorderStyle.Solid,
                    border, width, ButtonBorderStyle.Solid, border, width, ButtonBorderStyle.Solid);
            };
            Shell.SizeChanged += delegate { Shell.Invalidate(); };
            UpdateOutlineColor();
            Button choose = NativeLayout.Button("Choose color…", ChooseColor);
            TableLayoutPanel color = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2 };
            color.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            color.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            color.Controls.Add(colorInput, 0, 0);
            color.Controls.Add(choose, 1, 0);
            foreach (TextBox input in new[] { nameInput, colorInput })
            {
                input.Dock = DockStyle.None;
                input.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                input.Margin = Padding.Empty;
            }
            choose.Anchor = AnchorStyles.Left;
            choose.Margin = new Padding(8, 0, 0, 0);
            TableLayoutPanel nameRow = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2 };
            nameRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            nameRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Label nameLabel = NativeLayout.Label("Rule name");
            Label colorLabel = NativeLayout.Label("Color");
            nameLabel.Margin = colorLabel.Margin = new Padding(0, 0, 8, 0);
            nameRow.Controls.Add(nameLabel, 0, 0);
            nameRow.Controls.Add(nameInput, 1, 0);
            TableLayoutPanel colorRow = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2 };
            colorRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            colorRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            colorRow.Controls.Add(colorLabel, 0, 0);
            colorRow.Controls.Add(color, 1, 0);
            identity = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 3 };
            identity.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            identity.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            identity.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            identity.Controls.Add(nameRow, 1, 0);
            identity.Controls.Add(colorRow, 2, 0);

            preserve = Choice("Preserve text · AI cannot change this text");
            preserve.Checked = UiState.Text(rule, "action") == "protect";
            function = new ComboBox { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, DropDownStyle = ComboBoxStyle.DropDownList, AccessibleName = "Rule function" };
            function.MinimumSize = new Size(0, function.PreferredHeight);
            function.FontChanged += delegate { function.MinimumSize = new Size(0, function.PreferredHeight); };
            function.Items.AddRange(new object[] { "Follow wording instructions", "Color only · no AI function", "Reject AI changes", "Suggest deleting the complete passage" });
            int actionIndex = Array.IndexOf(Actions, UiState.Text(rule, "action"));
            function.SelectedIndex = Math.Max(0, actionIndex);
            functionRow = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2 };
            functionRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            functionRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            functionRow.Controls.Add(NativeLayout.Label("Function"), 0, 0);
            functionRow.Controls.Add(function, 1, 0);

            List<string> lines = UiState.Text(rule, "prompt").Replace("\r\n", "\n").Split('\n').ToList();
            var presetDefinitions = ReadPresets();
            presetChoices = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
            presetChoices.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            presetChoices.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            foreach (var definition in presetDefinitions.Where(entry => UiState.Text(entry, "id") != "shorten" && UiState.Text(entry, "id") != "detail"))
            {
                string prompt = UiState.Text(definition, "prompt");
                CheckBox choice = Choice(UiState.Text(definition, "label"));
                choice.Checked = lines.Remove(prompt);
                presets.Add(choice, prompt);
                int index = presetChoices.Controls.Count;
                if (index % 2 == 0) presetChoices.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                presetChoices.Controls.Add(choice, index % 2, index / 2);
            }
            presetGroup = NativeLayout.Group("Wording instructions · choose any", presetChoices);
            var unchanged = AccessChoice("Keep current length");
            unchanged.Checked = true;
            lengths.Add(unchanged, "");
            foreach (var definition in presetDefinitions.Where(entry => UiState.Text(entry, "id") == "shorten" || UiState.Text(entry, "id") == "detail"))
            {
                string prompt = UiState.Text(definition, "prompt");
                RadioButton choice = AccessChoice(UiState.Text(definition, "label"));
                lengths.Add(choice, prompt);
                if (lines.Remove(prompt)) choice.Checked = true;
            }
            lengthGroup = NativeLayout.Group("Length", ChoiceRow(lengths.Keys.ToArray()));
            promptInput = NativeLayout.TextBox("Custom prompt", true, false);
            promptInput.MinimumSize = new Size(0, 84);
            promptInput.MaxLength = 4000;
            promptInput.Text = String.Join(Environment.NewLine, lines);
            Label help = NativeLayout.Label("Save applies to every passage with this color. Wording changes edit the text directly.");
            Button save = NativeLayout.Button("Save color rule", SaveRule);
            Button cancel = NativeLayout.Button("Cancel", delegate { DialogResult = DialogResult.Cancel; Close(); });
            List<Control> actions = new List<Control> { save, cancel };
            removeMarker = NativeLayout.Button("Remove this marker", delegate {
                    NativeLayout.Attempt(this, delegate {
                        var state = UiState.Map(dispatch("state", UiState.Args("documentId", this.documentId)));
                        if (markerId == null || !UiState.Records(state, "marks").Any(mark => UiState.Text(mark, "id") == markerId && UiState.Text(mark, "ruleId") == ruleId))
                            throw new InvalidOperationException("This passage no longer has this color rule. Open its current rule again.");
                        dispatch("remove_mark", new Dictionary<string, object> { { "id", markerId }, { "documentId", this.documentId } });
                        DialogResult = DialogResult.OK;
                        Close();
                    });
                });
            removeMarker.AccessibleDescription = "Remove only the most recently opened passage's AI color marker. Keep its text and formatting.";
            SetMarkerContext(markId);
            actions.Insert(0, removeMarker);
            AcceptButton = save;
            CancelButton = cancel;
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.Controls.Add(identity, 0, 0);
            Shell.Controls.Add(preserve, 0, 1);
            Shell.Controls.Add(functionRow, 0, 2);
            Shell.Controls.Add(presetGroup, 0, 4);
            Shell.Controls.Add(lengthGroup, 0, 5);
            promptGroup = NativeLayout.Group("Custom prompt", promptInput);
            promptGroup.AutoSize = false;
            Shell.Controls.Add(promptGroup, 0, 6);
            Shell.Controls.Add(help, 0, 7);
            Shell.Controls.Add(NativeLayout.Actions(actions.ToArray()), 0, 8);
            foreach (GroupBox group in Shell.Controls.OfType<GroupBox>())
            {
                group.Padding = new Padding(10, 6, 10, 6);
                group.Margin = new Padding(0, 3, 0, 3);
            }
            promptGroup.Margin = Padding.Empty;
            presetChoices.SizeChanged += delegate { ReflowPresets(); };
            presetChoices.FontChanged += delegate { ReflowPresets(); };
            functionRow.SizeChanged += delegate { QueueGeometry(); };
            functionRow.FontChanged += delegate { QueueGeometry(); };
            SizeChanged += delegate { QueueGeometry(); };
            FontChanged += delegate { QueueGeometry(); };
            Shown += delegate { QueueGeometry(); };
            preserve.CheckedChanged += BehaviorChanged;
            function.SelectedIndexChanged += BehaviorChanged;
            BehaviorChanged(this, EventArgs.Empty);
        }

        private static List<IDictionary<string, object>> ReadPresets()
        {
            List<IDictionary<string, object>> definitions;
            using (var stream = typeof(RuleDialog).Assembly.GetManifestResourceStream("SelectiveEditorContract"))
            using (var reader = new StreamReader(stream))
                definitions = UiState.Records(UiState.Map(new JavaScriptSerializer().DeserializeObject(reader.ReadToEnd())), "presets");
            definitions.Add(new Dictionary<string, object> { { "id", "remove" }, { "label", "Remove" }, { "prompt", "Remove this passage." } });
            definitions.Add(new Dictionary<string, object> { { "id", "point" }, { "label", "Get to the point" }, { "prompt", "Get to the point; state the main claim directly and remove unnecessary lead-ins." } });
            return definitions;
        }

        private void ReflowPresets()
        {
            if (reflowingPresets) { QueueGeometry(); return; }
            reflowingPresets = true;
            try
            {
                int preferredWidth = presets.Keys.Max(choice => choice.PreferredSize.Width + choice.Margin.Horizontal);
                int columns = 3;
                presetMinimumWidth = Width - ClientSize.Width + columns * preferredWidth + presetChoices.Margin.Horizontal + presetGroup.Padding.Horizontal
                    + presetGroup.Margin.Horizontal + Shell.Padding.Horizontal;
                int rows = (presets.Count + columns - 1) / columns;
                int minimumHeight = rows * presets.Keys.Max(choice => choice.PreferredSize.Height + choice.Margin.Vertical);
                presetChoices.MinimumSize = new Size(0, minimumHeight);
                presetGroup.MinimumSize = new Size(0, minimumHeight + presetGroup.Font.Height + presetGroup.Padding.Vertical);
                AllocateRows();
                if (columns == presetChoices.ColumnCount) return;
                presetChoices.SuspendLayout();
                try
                {
                    presetChoices.Controls.Clear();
                    presetChoices.ColumnStyles.Clear();
                    presetChoices.RowStyles.Clear();
                    presetChoices.ColumnCount = columns;
                    presetChoices.RowCount = rows;
                    for (int index = 0; index < columns; index++) presetChoices.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / columns));
                    for (int index = 0; index < presetChoices.RowCount; index++) presetChoices.RowStyles.Add(new RowStyle(SizeType.Percent, 100f / presetChoices.RowCount));
                    int position = 0;
                    foreach (CheckBox choice in presets.Keys)
                    {
                        presetChoices.Controls.Add(choice, position % columns, position / columns);
                        position++;
                    }
                }
                finally { presetChoices.ResumeLayout(); }
            }
            finally { reflowingPresets = false; }
        }

        private void AllocateRows()
        {
            if (arrangingRows || Shell.RowStyles.Count < 9) return;
            arrangingRows = true;
            try
            {
                Shell.RowStyles[2].SizeType = SizeType.Absolute;
                functionRow.MinimumSize = new Size(0, functionRow.Controls.Cast<Control>()
                    .Max(control => control.GetPreferredSize(new Size(control.Width, 0)).Height + control.Margin.Vertical));
                Shell.RowStyles[2].Height = Math.Max(functionRow.MinimumSize.Height, functionRow.PreferredSize.Height) + functionRow.Margin.Vertical;
                Shell.RowStyles[4].SizeType = SizeType.Absolute;
                Shell.RowStyles[4].Height = presetGroup.MinimumSize.Height + presetGroup.Margin.Vertical;
                int promptHeight = Math.Max(84, promptInput.Font.Height * 3 + 12);
                promptInput.MinimumSize = new Size(0, promptHeight);
                promptGroup.MinimumSize = new Size(0, promptHeight + promptGroup.Font.Height + promptGroup.Padding.Vertical + promptInput.Margin.Vertical);
                Shell.RowStyles[6].SizeType = SizeType.Percent;
                Shell.RowStyles[6].Height = 100;
            }
            finally { arrangingRows = false; }
        }

        private void QueueGeometry()
        {
            if (geometryQueued || !IsHandleCreated || IsDisposed) return;
            geometryQueued = true;
            BeginInvoke((Action)delegate {
                geometryQueued = false;
                if (IsDisposed) return;
                ReflowPresets();
                AllocateRows();
                Shell.PerformLayout();
                int[] heights = Shell.GetRowHeights();
                int requiredHeight = Height - ClientSize.Height + Shell.Padding.Vertical;
                for (int row = 0; row < heights.Length; row++)
                    requiredHeight += row == 6 ? promptGroup.MinimumSize.Height + promptGroup.Margin.Vertical : heights[row];
                MinimumSize = new Size(Math.Max(RuleMinimum.Width, presetMinimumWidth), Math.Max(RuleMinimum.Height, requiredHeight));
            });
        }

        private static CheckBox Choice(string label)
        {
            return new CheckBox { Text = label, AccessibleName = label, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 1, 6, 1) };
        }

        private static RadioButton AccessChoice(string label)
        {
            return new RadioButton { Text = label, AccessibleName = label, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 1, 6, 1) };
        }

        private static TableLayoutPanel ChoiceRow(params RadioButton[] choices)
        {
            var row = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = choices.Length };
            RadioButton selected = choices.LastOrDefault(choice => choice.Checked);
            for (int index = 0; index < choices.Length; index++)
            {
                row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / choices.Length));
                row.Controls.Add(choices[index], index, 0);
            }
            if (selected != null) selected.Checked = true;
            return row;
        }

        private void BehaviorChanged(object sender, EventArgs args)
        {
            bool wording = !preserve.Checked && Actions[function.SelectedIndex] == "prompt";
            function.Enabled = !preserve.Checked;
            presetGroup.Enabled = wording;
            lengthGroup.Enabled = wording;
            promptInput.Enabled = !preserve.Checked && Actions[function.SelectedIndex] != "neutral" && Actions[function.SelectedIndex] != "reject";
        }

        private void ChooseColor(object sender, EventArgs args)
        {
            using (ColorDialog picker = new ColorDialog { Color = NativeLayout.ColorValue(colorInput.Text), FullOpen = true })
                if (picker.ShowDialog(this) == DialogResult.OK)
                    colorInput.Text = "#" + picker.Color.R.ToString("x2") + picker.Color.G.ToString("x2") + picker.Color.B.ToString("x2");
        }

        private void UpdateOutlineColor()
        {
            if (!Regex.IsMatch(colorInput.Text, "^#[0-9a-fA-F]{6}$")) return;
            outlineColor = ColorTranslator.FromHtml(colorInput.Text);
            if (IsHandleCreated) ApplyAppearance(SystemInformation.HighContrast);
            Shell.Invalidate();
        }

        internal void SetMarkerContext(string markId)
        {
            markerId = markId;
            removeMarker.Visible = markId != null;
        }

        protected override void OnHandleCreated(EventArgs args)
        {
            base.OnHandleCreated(args);
            OnAppearanceApplied(SystemInformation.HighContrast);
        }

        protected override void OnAppearanceApplied(bool highContrast)
        {
            if (!IsHandleCreated) return;
            int color = highContrast ? -1 : outlineColor.R | outlineColor.G << 8 | outlineColor.B << 16;
            Color ink = CaptionTextColor(outlineColor);
            int text = highContrast ? -1 : ink.R | ink.G << 8 | ink.B << 16;
            // Windows versions before Windows 11 retain their native caption colors.
            DwmSetWindowAttribute(Handle, 34, ref color, sizeof(int));
            DwmSetWindowAttribute(Handle, 35, ref color, sizeof(int));
            DwmSetWindowAttribute(Handle, 36, ref text, sizeof(int));
        }

        private static Color CaptionTextColor(Color color)
        {
            double[] channels = { color.R / 255.0, color.G / 255.0, color.B / 255.0 };
            for (int i = 0; i < channels.Length; i++)
                channels[i] = channels[i] <= 0.04045 ? channels[i] / 12.92 : Math.Pow((channels[i] + 0.055) / 1.055, 2.4);
            return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2] > 0.179 ? Color.Black : Color.White;
        }

        private void SaveRule(object sender, EventArgs args)
        {
            NativeLayout.Attempt(this, delegate
            {
                string label = nameInput.Text.Trim();
                if (label.Length == 0 || Encoding.UTF8.GetByteCount(label) > 100)
                    throw new InvalidOperationException("Enter a rule name using at most 100 UTF-8 bytes.");
                if (!Regex.IsMatch(colorInput.Text, "^#[0-9a-fA-F]{6}$"))
                    throw new InvalidOperationException("Enter a color as # followed by six hexadecimal digits, or use Choose color.");
                string action = preserve.Checked ? "protect" : Actions[function.SelectedIndex];
                List<string> instructions = new List<string>();
                if (action == "prompt")
                {
                    instructions.AddRange(presets.Where(pair => pair.Key.Checked).Select(pair => pair.Value));
                    instructions.AddRange(lengths.Where(pair => pair.Key.Checked && pair.Value.Length > 0).Select(pair => pair.Value));
                }
                if (promptInput.Enabled && promptInput.Text.Trim().Length > 0) instructions.Add(promptInput.Text.Trim());
                string prompt = String.Join("\n", instructions);
                if (Encoding.UTF8.GetByteCount(prompt) > 4000)
                    throw new InvalidOperationException("Wording instructions must use at most 4,000 UTF-8 bytes, including presets.");
                dispatch("save_rule", new Dictionary<string, object>
                {
                    { "id", ruleId }, { "label", label }, { "color", colorInput.Text.ToLowerInvariant() },
                    { "action", action }, { "prompt", prompt },
                    { "accessMode", action == "prompt" ? "direct" : "suggestions" },
                    { "documentId", documentId }
                });
                DialogResult = DialogResult.OK;
                Close();
            });
        }
    }
}
