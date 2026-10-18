using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using UnderTree.WordAddin;

public sealed class MapTestRetrieval { public bool IncludeHiddenText { get; set; } }
public sealed class MapTestRange
{
    public int Start { get; set; }
    public int End { get; set; }
    public int StoryType { get { return 1; } }
    public string Text { get; set; }
    public List<MapTestField> Fields { get; set; }
    public MapTestRetrieval TextRetrievalMode { get { return new MapTestRetrieval(); } }
}
public sealed class MapTestRevision
{
    public int Type { get { return 2; } }
    public MapTestRange Range { get; set; }
}
public sealed class MapTestField
{
    public MapTestRange Code { get; set; }
    public MapTestRange Result { get; set; }
}
public sealed class MapTestDocument
{
    public string NativeText;
    public List<MapTestRevision> Revisions { get; set; }
    public List<MapTestField> Fields { get; set; }
    public MapTestRange Content { get { return Range(0, NativeText.Length); } }
    public MapTestRange Range(int from, int to)
    { return new MapTestRange { Start = from, End = to, Text = NativeText.Substring(from, to - from), Fields = new List<MapTestField>() }; }
}

public sealed class MapTableBounds { public int Start, End; public int StoryType = 1; }
public sealed class MapTestTable
{
    public MapTableBounds Initial, Later;
    private int reads;
    public object Range { get { return ++reads > 1 && Later != null ? Later : Initial; } }
}
public sealed class ExpandedMapDocument
{
    public string[] Units;
    public List<MapTestField> Fields = new List<MapTestField>();
    public List<MapTestRevision> Revisions = new List<MapTestRevision>();
    public List<MapTestTable> Tables = new List<MapTestTable>();
    public int NativeReads;
    public bool WrongRangeBounds;
    public Func<int, int, string, string> OverrideText;
    public Action<int, int> BeforeRead;
    public MapTestRange Content { get { return new MapTestRange { Start = 0, End = Units.Length }; } }
    public MapTestRange Range(int from, int to)
    {
        NativeReads++; if (BeforeRead != null) BeforeRead(from, to);
        string text = String.Join("", Units.Skip(from).Take(to - from));
        if (OverrideText != null) text = OverrideText(from, to, text);
        return new MapTestRange { Start = from + (WrongRangeBounds ? 1 : 0), End = to,
            Text = text, Fields = new List<MapTestField>() };
    }
}

internal static class WordTextMapTests
{
    private static int assertions;
    public static int Main()
    {
        try
        {
            var document = Document("abcdeX", 0, 3, 3, 5);
            var map = new WordTextMap(document);
            Equal("X", map.Text, "Adjacent deleted revisions leave only final text.");
            for (int native = 0; native <= 5; native++) Equal(0, map.VisibleOffset(native), "Every adjacent deleted position maps to the same final boundary.");
            Equal(1, map.VisibleOffset(6), "Visible suffix preserves its final boundary.");
            Equal(5, map.NativeOffset(0), "The writable final range starts after both deletion revisions.");
            Equal(6, map.NativeOffset(1), "The visible suffix maps to its exact native endpoint.");
            Denied(() => map.VisibleOffset(-1), "Out-of-bounds native positions stay invalid.");
            Denied(() => map.VisibleOffset(7), "Native positions cannot overrun the document.");

            map = new WordTextMap(Document("ab\ud83d\ude00cdX", 0, 2, 2, 4, 4, 6));
            Equal("X", map.Text, "Three adjacent deletions preserve the Unicode suffix.");
            for (int native = 0; native <= 6; native++) Equal(0, map.VisibleOffset(native), "A transitive deletion chain maps starts, interiors and endpoints.");
            Equal(6, map.NativeOffset(0), "A transitive deletion chain never exposes deleted source as editable text.");

            document = Document("###AB#cdX", 6, 8);
            document.Fields.Add(new MapTestField { Code = document.Range(1, 2), Result = document.Range(3, 5) });
            map = new WordTextMap(document);
            Equal("ABX", map.Text, "Field codes and deleted revisions are omitted from final text.");
            Equal(0, map.VisibleOffset(3), "Field result starts map through the native opening boundary.");
            Equal(2, map.VisibleOffset(5), "A field closing boundary resolves through an adjacent deletion.");
            Equal(2, map.VisibleOffset(7), "An interior deletion after a field stays at the same final boundary.");
            Equal(8, map.NativeOffset(2), "Native editing begins after the hidden field closing and deletion.");
            Equal(9, map.NativeOffset(3), "Final text still maps to the exact document end.");
            Denied(() => new WordTextMap(Document("abcdeX", 0, 3, 2, 5)), "Conflicting native deletion ranges fail closed.");
            document = Document("AxxByyCzz", 1, 3, 4, 6, 7, 9);
            map = new WordTextMap(document);
            Equal("ABC", map.Text, "Multiple removed native spans leave the exact logical current body.");
            Equal("0..1|3..4|6..7", String.Join("|", map.NativeSegments(0, 3).Select(span => span[0] + ".." + span[1])),
                "Formatting traverses only current native segments and never touches original deleted text.");
            MapTestRange currentUnit = map.Range(document, 2, 3);
            Equal(6, currentUnit.Start, "A protected current unit begins at its own native object.");
            Equal(7, currentUnit.End, "A trailing deletion cannot expand a protected current unit's native range.");
            Equal("C", currentUnit.Text, "Tight native unit ranges exclude original deleted content.");
            document = Document("###AB#cdX", 6, 8);
            document.Fields.Add(new MapTestField { Code = document.Range(1, 2), Result = document.Range(3, 5) });
            map = new WordTextMap(document);
            currentUnit = map.Range(document, 0, 2);
            Equal(0, currentUnit.Start, "Whole current fields retain their native opening structure.");
            Equal(6, currentUnit.End, "Whole current fields retain their closing structure while excluding adjacent deleted text.");
            TableMapping();
            Console.WriteLine("PASS " + assertions + " native final-text mapping assertions.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static MapTestDocument Document(string text, params int[] deleted)
    {
        var document = new MapTestDocument { NativeText = text, Fields = new List<MapTestField>(), Revisions = new List<MapTestRevision>() };
        for (int index = 0; index < deleted.Length; index += 2)
            document.Revisions.Add(new MapTestRevision { Range = document.Range(deleted[index], deleted[index + 1]) });
        return document;
    }

    private static ExpandedMapDocument Expanded(params string[] units)
    {
        var document = new ExpandedMapDocument { Units = units };
        document.Tables.Add(new MapTestTable { Initial = new MapTableBounds { Start = 1, End = units.Length - 1 } });
        return document;
    }

    private static void TableMapping()
    {
        var document = Expanded("A", "\r\a", "\r\a", "B");
        var map = new WordTextMap(document);
        Equal("A\r\a\r\aB", map.Text, "Verified native table expansion retains every returned UTF-16 unit.");
        Equal(1, map.VisibleOffset(1), "Expanded native start canonicalizes to the CR boundary.");
        Equal(3, map.VisibleOffset(2), "Adjacent expanded native start canonicalizes to the first equal boundary.");
        Equal(5, map.VisibleOffset(3), "After-table native position preserves its exact logical offset.");
        Equal(3, map.NativeOffset(5), "After-table logical position maps back to the real native position.");
        Equal("B", (string)map.Range(document, 5, 6).Text, "Outside prose resolves to its exact native range after expansion.");
        Denied(() => map.NativeOffset(2), "The first virtual half-terminator has no native edit boundary.");
        Denied(() => map.NativeOffset(4), "An adjacent virtual half-terminator stays inaccessible.");
        Denied(() => map.Range(document, 0, 2), "A range starting in outside prose cannot end inside a phantom delimiter half.");
        Denied(() => map.Range(document, 2, 2), "A collapsed range cannot turn a phantom delimiter half into native authority.");
        document.Tables.Clear();
        Denied(() => new WordTextMap(document), "Non-table CR/BEL expansion never receives table ownership.");
        Denied(() => new WordTextMap(Expanded("A", "XX", "B")), "Unknown same-size native expansion fails closed.");
        Denied(() => new WordTextMap(Expanded("A", "", "B")), "Unknown native contraction fails closed.");
        document = Expanded("A", "\r\a", "B"); document.WrongRangeBounds = true;
        Denied(() => new WordTextMap(document), "Unexpected actual native range bounds fail closed.");
        document = Expanded("A", "\r\a", "\r\a", "B");
        document.Tables.Add(new MapTestTable { Initial = new MapTableBounds { Start = 2, End = 3 } });
        Denied(() => new WordTextMap(document), "Overlapping table ownership cannot infer an expanded boundary.");
        document = Expanded("A", "\r\a", "B"); document.Tables[0].Initial.End = 10;
        Denied(() => new WordTextMap(document), "Out-of-document native table bounds fail closed.");
        document = Expanded("A", "\r\a", "B"); document.Tables[0].Initial.StoryType = 2;
        Denied(() => new WordTextMap(document), "A table in another story cannot own a body expansion.");
        document = Expanded("A", "\r\a", "B");
        document.Tables[0].Later = new MapTableBounds { Start = 1, End = 3 };
        Denied(() => new WordTextMap(document), "Native table bounds changed during mapping fail closed.");
        document = Expanded("A", "\r\a", "B");
        document.OverrideText = (from, to, value) => from == 2 && to == 3 ? "X" : value;
        Denied(() => new WordTextMap(document), "Parent and child native reads must expose identical text.");

        document = Expanded("#", "#", "#", "A", "B", "#", "c", "d", "\r\a", "\r\a", "X");
        document.Tables[0].Initial = new MapTableBounds { Start = 8, End = 10 };
        document.Fields.Add(new MapTestField { Code = document.Range(1, 2), Result = document.Range(3, 5) });
        document.Revisions.Add(new MapTestRevision { Range = document.Range(6, 8) });
        map = new WordTextMap(document);
        Equal("AB\r\a\r\aX", map.Text, "Table expansion preserves adjacent hidden field and deletion behavior.");
        Equal(0, map.VisibleOffset(3), "Field opening aliases remain canonical beside a table.");
        Equal(2, map.VisibleOffset(7), "Deletion interior resolves to the first canonical expanded boundary.");
        Equal(8, map.NativeOffset(2), "A table after a hidden field and deletion starts at its verified native unit.");
        Denied(() => map.NativeOffset(3), "Deletion alias resolution cannot authorize a virtual table interior.");
        Equal(10, map.NativeOffset(6), "The visible suffix remains accurately positioned after all hidden and expanded units.");

        document = Expanded("\ud83d", "\ude00", "\r\a", "e", "\u0301");
        document.Tables[0].Initial = new MapTableBounds { Start = 2, End = 3 };
        map = new WordTextMap(document);
        Equal("\ud83d\ude00\r\ae\u0301", map.Text, "Expanded table mapping preserves adjacent surrogate and combining UTF-16 units exactly.");
        Equal(3, map.NativeOffset(4), "Combining text after an expanded delimiter keeps its native start.");

        document = Expanded("A", "\r\a", "x", "y", "B");
        document.Tables[0].Initial = new MapTableBounds { Start = 1, End = 2 };
        document.Revisions.Add(new MapTestRevision { Range = document.Range(2, 3) });
        document.Revisions.Add(new MapTestRevision { Range = document.Range(3, 4) });
        map = new WordTextMap(document);
        Equal("A\r\aB", map.Text, "A table terminator followed immediately by adjacent deletions retains exact text.");
        Equal(3, map.VisibleOffset(2), "A post-terminator deletion start resolves to the canonical following logical boundary.");
        Equal(3, map.VisibleOffset(3), "Adjacent deletion fallback remains canonical after an expanded delimiter.");
        Equal(4, map.NativeOffset(3), "The end of the expanded pair includes native deletion aliases without inventing a half-boundary.");

        document = Expanded(Enumerable.Repeat("\r\a", 512).ToArray());
        document.Tables[0].Initial = new MapTableBounds { Start = 0, End = 512 };
        DeniedWith(() => new WordTextMap(document), "inspection limit", "Complex mapping stops at its bounded extra-probe budget.");
        Equal(true, document.NativeReads <= 513, "Probe exhaustion stops before walking every native table unit.");
        document = Expanded("A", "\r\a", "B"); bool rightRead = false;
        document.BeforeRead = (from, to) => { if (from == 0 && to == 1) Thread.Sleep(1100); if (from == 1 && to == 3) rightRead = true; };
        DeniedWith(() => new WordTextMap(document), "inspection limit", "A late conclusive child cannot turn an expired mapping into success.");
        Equal(false, rightRead, "An expired left child is detected before querying the right child.");
        document = Expanded("A", "\r\a", "B");
        document.BeforeRead = (from, to) => { if (from == 2 && to == 3) Thread.Sleep(1100); };
        DeniedWith(() => new WordTextMap(document), "inspection limit", "Late right success remains outside the inspection deadline.");
        Denied(() => new WordTextMap(Document(new String('x', PolicyEngine.MaxText + 1))), "The complete-text maximum remains enforced on the fast path.");
    }

    private static void DeniedWith(Action action, string reason, string message)
    {
        assertions++;
        try { action(); }
        catch (InvalidOperationException error)
        {
            if (error.Message.IndexOf(reason, StringComparison.Ordinal) >= 0) return;
            throw new Exception(message + " Wrong refusal: " + error.Message);
        }
        throw new Exception(message);
    }

    private static void Equal<T>(T expected, T actual, string message)
    { assertions++; if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception(message); }
    private static void Denied(Action action, string message)
    {
        assertions++;
        try { action(); }
        catch (InvalidOperationException) { return; }
        throw new Exception(message);
    }
}
