using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace UnderTree.WordAddin
{
    internal sealed class HighlightedRulesDialog : NativeDialog
    {
        internal HighlightedRulesDialog(Func<string, IDictionary<string, object>, object> dispatch, IDictionary<string, object> state,
            Action<IDictionary<string, object>, string, string> openRule = null)
            : base("Choose highlighted color", new Size(570, 410))
        {
            var choices = new PagedRecordList("Overlapping highlight identities", mark => UiState.Text(mark, "label") + " · " +
                UiState.Text(mark, "action") + " · " + UiState.Text(mark, "id") + " · " + UiState.Text(mark, "quote"));
            choices.UpdateRecords(UiState.Records(state, "choices"));
            Action open = delegate {
                if (choices.Selected == null) return;
                string ruleId = UiState.Text(choices.Selected, "ruleId");
                var rule = UiState.Records(state, "rules").First(candidate => UiState.Text(candidate, "id") == ruleId);
                if (openRule == null) openRule = new ColorRuleWindows(dispatch, () => this).Open;
                openRule(rule, UiState.Text(state, "documentId"), UiState.Text(choices.Selected, "id"));
            };
            choices.RecordOpened += delegate { NativeLayout.Attempt(this, open); };
            Button choose = NativeLayout.Button("Open selected color rule", delegate { NativeLayout.Attempt(this, open); });
            Button close = NativeLayout.Button("Close", delegate { Close(); });
            AcceptButton = choose;
            CancelButton = close;
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Shell.Controls.Add(NativeLayout.Label("Choose the exact highlighted color. Opening its rule changes no permissions; only Save color rule assigns a function."), 0, 0);
            Shell.Controls.Add(choices, 0, 1);
            Shell.Controls.Add(NativeLayout.Actions(choose, close), 0, 2);
        }

        internal static bool OpenAtSelection(Func<string, IDictionary<string, object>, object> dispatch, IWin32Window owner = null,
            ColorRuleWindows windows = null)
        { return (windows ?? new ColorRuleWindows(dispatch, () => owner)).OpenAtSelection(); }
    }

    internal sealed class ColorRuleWindows : IDisposable
    {
        private readonly Func<string, IDictionary<string, object>, object> dispatch;
        private readonly Func<IWin32Window> owner;
        private readonly Dictionary<string, RuleDialog> rules = new Dictionary<string, RuleDialog>();
        private readonly Dictionary<string, HighlightedRulesDialog> choosers = new Dictionary<string, HighlightedRulesDialog>();

        internal ColorRuleWindows(Func<string, IDictionary<string, object>, object> dispatch, Func<IWin32Window> owner)
        { this.dispatch = dispatch; this.owner = owner; }

        internal void Open(IDictionary<string, object> requested, string documentId, string markId)
        {
            var state = UiState.Map(dispatch("state", UiState.Args("documentId", documentId)));
            if (UiState.Text(UiState.Map(UiState.Get(state, "document")), "id") != documentId)
                throw new InvalidOperationException("The active document changed. Open its color rule again.");
            string ruleId = requested == null ? "" : UiState.Text(requested, "id");
            IDictionary<string, object> rule = requested == null ? null : UiState.Records(state, "rules").FirstOrDefault(candidate => UiState.Text(candidate, "id") == ruleId);
            if (requested != null && rule == null) throw new InvalidOperationException("This color rule no longer exists.");
            if (markId != null && !UiState.Records(state, "marks").Any(mark => UiState.Text(mark, "id") == markId && UiState.Text(mark, "ruleId") == ruleId))
                throw new InvalidOperationException("This passage no longer has the selected color rule.");
            string key = documentId + "\0" + (rule == null ? Guid.NewGuid().ToString("N") : ruleId);
            RuleDialog dialog;
            if (rules.TryGetValue(key, out dialog))
            {
                dialog.SetMarkerContext(markId);
                dialog.Activate();
                return;
            }
            if (rules.Count >= 24) throw new InvalidOperationException("Close a color rule window before opening another.");
            IWin32Window nativeOwner = owner();
            dialog = new RuleDialog(dispatch, rule, documentId, markId);
            rules.Add(key, dialog);
            dialog.FormClosed += delegate { rules.Remove(key); dialog.Dispose(); };
            try { dialog.Show(nativeOwner); }
            catch { rules.Remove(key); dialog.Dispose(); throw; }
        }

        internal bool OpenAtSelection(IDictionary<string, object> args = null)
        {
            var state = UiState.Map(dispatch("selection_rules", args ?? new Dictionary<string, object>()));
            var marks = UiState.Records(state, "choices");
            if (marks.Count == 0) return false;
            if (marks.Count == 1)
            {
                string ruleId = UiState.Text(marks[0], "ruleId");
                var rule = UiState.Records(state, "rules").First(candidate => UiState.Text(candidate, "id") == ruleId);
                Open(rule, UiState.Text(state, "documentId"), UiState.Text(marks[0], "id"));
            }
            else
            {
                string documentId = UiState.Text(state, "documentId");
                HighlightedRulesDialog previous;
                if (choosers.TryGetValue(documentId, out previous)) previous.Close();
                if (choosers.Count >= 24) throw new InvalidOperationException("Close a highlighted color chooser before opening another.");
                IWin32Window nativeOwner = owner();
                var dialog = new HighlightedRulesDialog(dispatch, state, Open);
                choosers.Add(documentId, dialog);
                dialog.FormClosed += delegate { choosers.Remove(documentId); dialog.Dispose(); };
                try { dialog.Show(nativeOwner); }
                catch { choosers.Remove(documentId); dialog.Dispose(); throw; }
            }
            return true;
        }

        public void Dispose()
        {
            foreach (var dialog in choosers.Values.ToArray()) dialog.Close();
            foreach (var dialog in rules.Values.ToArray()) dialog.Close();
            choosers.Clear(); rules.Clear();
        }
    }
}
