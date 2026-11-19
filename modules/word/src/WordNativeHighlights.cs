using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace UnderTree.WordAddin
{
    internal sealed class WordNativeHighlights
    {
        private static readonly string[] Names = {
            "", "Black", "Blue", "Turquoise", "Bright Green", "Pink", "Red", "Yellow", "White",
            "Dark Blue", "Teal", "Green", "Violet", "Dark Red", "Dark Yellow", "Gray 50%", "Gray 25%"
        };
        private static readonly string[] Colors = {
            "", "#000000", "#0000FF", "#00FFFF", "#00FF00", "#FF00FF", "#FF0000", "#FFFF00", "#FFFFFF",
            "#000080", "#008080", "#008000", "#800080", "#800000", "#808000", "#808080", "#C0C0C0"
        };
        private static readonly string[] XmlColors = {
            "none", "black", "blue", "cyan", "green", "magenta", "red", "yellow", "white",
            "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray"
        };
        private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const int MaxNativeXml = 67108864;
        private const string RunHighlights = "w:r[w:t or w:tab or w:br or w:cr or w:sym]/w:rPr/w:highlight";
        private readonly long deadline = Stopwatch.GetTimestamp() + Stopwatch.Frequency;
        private int probes;

        internal static string Name(int index) { return Names[index]; }
        internal static string Rgb(int index) { return Colors[index]; }

        internal static int[] Palette(dynamic document)
        {
            var colors = new SortedSet<int>();
            int count = 0;
            foreach (object root in (IEnumerable)document.StoryRanges)
            {
                dynamic story = root;
                while (story != null)
                {
                    if (++count > 256) throw new InvalidOperationException("Word contains too many stories to read its highlight colors.");
                    colors.UnionWith(XmlPalette((string)story.WordOpenXML));
                    story = story.NextStoryRange;
                }
            }
            return colors.ToArray();
        }

        private static IEnumerable<int> XmlPalette(string source, bool bodyOnly = false)
        {
            XmlDocument xml = WordSnapshot.ReadNativeXml(source, MaxNativeXml);
            XmlNamespaceManager namespaces = WordSnapshot.WordNamespaces(xml);
            foreach (XmlElement highlight in xml.SelectNodes((bodyOnly ? "//w:body//" : "//") + RunHighlights, namespaces))
            {
                int index = Array.IndexOf(XmlColors, highlight.GetAttribute("val", WordNamespace));
                if (index > 0) yield return index;
            }
        }

        internal static bool ContainsHighlight(dynamic range)
        {
            return (int)range.HighlightColorIndex != 0 || XmlPalette((string)range.WordOpenXML, true).Any();
        }

        internal static bool TryVisitBody(WordSnapshot snapshot, Action<int, int, int> visit)
        {
            XmlDocument xml = WordSnapshot.ReadNativeXml(snapshot.NativeBodySource, MaxNativeXml);
            XmlNamespaceManager namespaces = WordSnapshot.WordNamespaces(xml);
            var spans = new List<int[]>();
            int offset = 0;
            foreach (XmlNode paragraph in xml.SelectNodes("//w:body//w:p[not(ancestor::w:txbxContent or ancestor::w:del or ancestor::w:moveFrom)]", namespaces))
            {
                foreach (XmlNode run in paragraph.SelectNodes(".//w:r[not(ancestor::w:txbxContent or ancestor::w:del or ancestor::w:moveFrom)]", namespaces))
                {
                    var highlight = run.SelectSingleNode("w:rPr/w:highlight", namespaces) as XmlElement;
                    int index = highlight == null ? 0 : Array.IndexOf(XmlColors, highlight.GetAttribute("val", WordNamespace));
                    if (index < 0) return false;
                    foreach (XmlNode part in run.ChildNodes)
                    {
                        if (part.NodeType != XmlNodeType.Element || part.NamespaceURI != WordNamespace) continue;
                        string text;
                        switch (part.LocalName)
                        {
                            case "rPr": case "instrText": case "fldChar": case "commentReference": case "lastRenderedPageBreak": continue;
                            case "t": text = part.InnerText; break;
                            case "tab": text = "\t"; break;
                            case "cr": case "br": text = "\v"; break;
                            default: return false;
                        }
                        if (!Matches(snapshot.MainText, offset, text)) return false;
                        if (index > 0 && text.Length > 0) spans.Add(new[] { offset, offset + text.Length, index });
                        offset += text.Length;
                    }
                }
                if (!Matches(snapshot.MainText, offset, "\r")) return false;
                offset++;
            }
            if (offset != snapshot.MainText.Length) return false;
            foreach (int[] span in spans) visit(span[0], span[1], span[2]);
            return true;
        }

        private static bool Matches(string body, int offset, string text)
        { return text.Length <= body.Length - offset && String.CompareOrdinal(body, offset, text, 0, text.Length) == 0; }

        internal void VisitStory(dynamic story, Action<int, int, int> visit)
        { Visit(story.Duplicate, (int)story.Start, (int)story.End, visit); }

        private void Probe()
        {
            if (++probes > 4096 || Stopwatch.GetTimestamp() > deadline)
                throw new InvalidOperationException("Word could not finish reading the document's text highlight colors.");
        }

        private int ColorAt(dynamic range, int from, int to)
        {
            Probe();
            range.SetRange(from, to);
            return (int)range.HighlightColorIndex;
        }

        private void Visit(dynamic range, int from, int to, Action<int, int, int> visit)
        {
            if (to <= from) return;
            int index = ColorAt(range, from, to);
            if (index < 0 || index >= Names.Length && index != 9999999)
                throw new InvalidOperationException("Word reported an unresolved text highlight color.");
            if (to - from == 1)
            {
                if (index == 9999999) throw new InvalidOperationException("Word reported an unresolved text highlight color.");
                if (index != 0) visit(from, to, index);
                return;
            }
            if (index == 0 && !XmlPalette((string)range.WordOpenXML, true).Any()) return;
            if (index != 9999999)
            {
                string text = (string)range.Text;
                if (text != null && text.Length == to - from)
                {
                    // Word reports the surrounding letters' color for some differently highlighted spaces.
                    int start = from;
                    for (int offset = 0; offset < text.Length; offset++)
                    {
                        if (!Char.IsWhiteSpace(text[offset])) continue;
                        int position = from + offset;
                        int space = ColorAt(range, position, position + 1);
                        if (space == index) continue;
                        if (index != 0 && position > start) visit(start, position, index);
                        if (space > 0 && space < Names.Length) visit(position, position + 1, space);
                        else if (space != 0) throw new InvalidOperationException("Word reported an unresolved text highlight color.");
                        start = position + 1;
                    }
                    if (index != 0 && to > start) visit(start, to, index);
                    return;
                }
            }
            int middle = from + (to - from) / 2;
            Visit(range, from, middle, visit);
            Visit(range, middle, to, visit);
        }
    }

    internal sealed partial class WordEditor
    {
        private sealed class HighlightSpan
        {
            internal int From, To;
            internal ColorRule Rule;
        }

        private void RefreshNativeHighlightMarks(WordDocumentReview state, WordSnapshot snapshot)
        {
            var rules = state.Policy.Rules.Where(rule => rule.NativeHighlightIndex != 0 && rule.Action != "neutral")
                .ToDictionary(rule => rule.NativeHighlightIndex);
            var marks = new List<PassageMark>();
            if (rules.Count > 0)
            {
                var spans = new List<HighlightSpan>();
                Action<int, int, int> add = (from, to, index) =>
                {
                    ColorRule rule;
                    if (!rules.TryGetValue(index, out rule)) return;
                    if (to <= from) return;
                    HighlightSpan last = spans.LastOrDefault();
                    if (last != null && last.Rule == rule && last.To == from) last.To = to;
                    else spans.Add(new HighlightSpan { From = from, To = to, Rule = rule });
                };
                if (!WordNativeHighlights.TryVisitBody(snapshot, add))
                    new WordNativeHighlights().VisitStory(state.Document.Content, (Action<int, int, int>)((from, to, index) =>
                        add(snapshot.VisibleOffset(from), snapshot.VisibleOffset(to), index)));
                foreach (HighlightSpan span in spans)
                {
                    int start = span.From;
                    for (int offset = start; offset < span.To; offset++)
                    {
                        char current = snapshot.MainText[offset];
                        if (current != '\r' && current != '\v' && current != '\a') continue;
                        AddNativeHighlight(marks, snapshot, span.Rule, start, offset);
                        start = offset + 1;
                    }
                    AddNativeHighlight(marks, snapshot, span.Rule, start, span.To);
                }
            }
            state.Policy.SetNativeHighlights(marks);
        }

        private static void AddNativeHighlight(List<PassageMark> marks, WordSnapshot snapshot, ColorRule rule, int from, int to)
        {
            if (to <= from) return;
            if (marks.Count >= 500 || to - from > PolicyEngine.MaxReplacement)
                throw new InvalidOperationException("The document has too many assigned highlight passages or a highlight exceeds the review limit.");
            string quote = PolicyEngine.Slice(snapshot.MainText, from, to);
            string key = rule.Id + "\0" + from.ToString(CultureInfo.InvariantCulture) + "\0" + to.ToString(CultureInfo.InvariantCulture) + "\0" + quote;
            marks.Add(new PassageMark { Id = "WraiN_" + PolicyEngine.Hash(key).Substring(0, 32), RuleId = rule.Id,
                From = from, To = to, Quote = quote, Prompt = "", NativeHighlight = true });
        }

        private object NativeHighlightColors()
        {
            if ((int)application.Documents.Count == 0) return new { documentId = "", colors = new object[0] };
            if ((int)application.UndoRecord.CustomRecordLevel != 0)
                throw new InvalidOperationException("Wait for Word to finish the current operation.");
            WordDocumentReview state = ActiveDocumentReview();
            return NativeHighlightColors(state, WordNativeHighlights.Palette(state.Document));
        }

        private static object NativeHighlightColors(WordDocumentReview state, int[] colors)
        {
            return new { documentId = state.Id, colors = colors.Select(index =>
            {
                ColorRule rule = state.Policy.Rules.FirstOrDefault(candidate => candidate.NativeHighlightIndex == index);
                return new { index = index, name = WordNativeHighlights.Name(index), color = WordNativeHighlights.Rgb(index),
                    action = rule == null ? "neutral" : rule.Action, prompt = rule == null ? "" : rule.Prompt };
            }).ToArray() };
        }

        private object SaveNativeHighlightRule(IDictionary<string, object> args)
        {
            OnlyKeys(args, "documentId", "index", "action", "prompt");
            WordDocumentReview state = Current();
            if (Text(args, "documentId") != state.Id) throw new InvalidOperationException("The active document changed. Open its highlight rule again.");
            RequireWritable(state);
            int index = Number(args, "index");
            int[] colors = WordNativeHighlights.Palette(state.Document);
            if (!colors.Contains(index)) throw new InvalidOperationException("This highlight color is no longer present in the document.");
            string action = Text(args, "action");
            string prompt = Text(args, "prompt", "").Trim();
            if (action == "prompt" && prompt.Length == 0) throw new InvalidOperationException("Enter an editing instruction.");
            ColorRule existing = state.Policy.Rules.FirstOrDefault(rule => rule.NativeHighlightIndex == index);
            var assignment = new ColorRule {
                Id = existing == null ? "word_" + Guid.NewGuid().ToString("N") : existing.Id,
                Label = WordNativeHighlights.Name(index), Color = WordNativeHighlights.Rgb(index), NativeHighlightIndex = index,
                Action = action, Prompt = action == "prompt" ? prompt : "", AccessMode = action == "prompt" ? "direct" : "suggestions"
            };
            PolicyEngine previous = state.Policy;
            state.Policy = PrepareLocalPolicy(state, candidate => candidate.AddOrUpdateRule(assignment));
            try { Persist(state, true); }
            catch { state.Policy = previous; throw; }
            state.NativeEdit.Capture(state.Document, state.Policy.ExportState());
            return NativeHighlightColors(state, colors);
        }
    }
}
