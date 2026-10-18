using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace UnderTree.WordAddin
{
    internal sealed class WordTextMap
    {
        private sealed class TableBounds { internal object Table; internal int From, To; }
        private sealed class NativeInspection
        {
            internal int Start, End;
            private int probes;
            private long deadline;
            private List<TableBounds> tables;

            internal void Begin() { if (deadline == 0) deadline = Stopwatch.GetTimestamp() + Stopwatch.Frequency; }
            internal void CheckTime()
            { if (deadline != 0 && Stopwatch.GetTimestamp() > deadline) throw Limit(); }
            internal void Probe()
            { CheckTime(); if (++probes > 512) throw Limit(); }
            private static InvalidOperationException Limit()
            { return new InvalidOperationException("Word's table positions exceed the safe inspection limit."); }

            private TableBounds Bounds(object value)
            {
                Probe(); dynamic range = ((dynamic)value).Range;
                int from = (int)range.Start, to = (int)range.End, story = (int)range.StoryType;
                CheckTime();
                if (story != 1 || from < Start || to > End || to <= from) throw MappingError();
                return new TableBounds { Table = value, From = from, To = to };
            }

            internal bool OwnsTableUnit(dynamic document, int from, int to)
            {
                if (tables == null)
                {
                    tables = new List<TableBounds>();
                    foreach (object table in (IEnumerable)document.Tables) tables.Add(Bounds(table));
                    tables = tables.OrderBy(table => table.From).ToList();
                    for (int index = 1; index < tables.Count; index++)
                        if (tables[index].From < tables[index - 1].To) throw MappingError();
                }
                foreach (TableBounds table in tables)
                    if (table.From <= from && to <= table.To)
                    {
                        TableBounds current = Bounds(table.Table);
                        if (current.From != table.From || current.To != table.To) throw MappingError();
                        return true;
                    }
                CheckTime(); return false;
            }

            internal void Complete(dynamic document)
            {
                CheckTime(); if (tables == null) return;
                Probe(); int count = (int)document.Tables.Count; CheckTime();
                if (count != tables.Count) throw MappingError();
                foreach (TableBounds table in tables)
                {
                    TableBounds current = Bounds(table.Table);
                    if (current.From != table.From || current.To != table.To) throw MappingError();
                }
            }
        }
        private sealed class HiddenFieldRange { internal int From, To; internal bool Closing; internal bool Deleted; }
        private readonly int[] offsets;
        private readonly Dictionary<int, int> aliases;
        private readonly List<HiddenFieldRange> deletions = new List<HiddenFieldRange>();
        internal readonly string Text;

        internal WordTextMap(dynamic document)
        {
            var hidden = new List<HiddenFieldRange>();
            foreach (object entry in (IEnumerable)document.Revisions)
            {
                dynamic revision = entry;
                int type = (int)revision.Type;
                if ((int)revision.Range.StoryType == 1 && (type == 2 || type == 14 || type == 21))
                    deletions.Add(new HiddenFieldRange { From = (int)revision.Range.Start, To = (int)revision.Range.End, Closing = true, Deleted = true });
                if (deletions.Count > 10000) throw new InvalidOperationException("Too many deleted revision ranges to map safely.");
            }
            hidden.AddRange(deletions);
            CollectFields(document.Fields, hidden, 0);
            int cursor = (int)document.Content.Start;
            int end = (int)document.Content.End;
            var inspection = new NativeInspection { Start = cursor, End = end };
            var boundaries = new List<int> { cursor };
            var text = new StringBuilder();
            aliases = new Dictionary<int, int>();
            foreach (HiddenFieldRange interval in hidden.OrderBy(part => part.From).ThenByDescending(part => part.To))
            {
                if (interval.To <= cursor) continue;
                if (interval.From < cursor) throw new InvalidOperationException("Conflicting native revision or field positions require local resolution.");
                Append(document, cursor, interval.From, text, boundaries, inspection);
                if (interval.Closing)
                {
                    boundaries[boundaries.Count - 1] = interval.To;
                    aliases[interval.From] = interval.To;
                }
                else aliases[interval.To] = interval.From;
                cursor = interval.To;
            }
            Append(document, cursor, end, text, boundaries, inspection);
            inspection.Complete(document);
            offsets = boundaries.ToArray();
            Text = text.ToString();
        }

        private void CollectFields(dynamic fields, List<HiddenFieldRange> ranges, int depth)
        {
            if (depth > 16) throw new InvalidOperationException("Nested fields exceed the native mapping bound.");
            foreach (object entry in (IEnumerable)fields)
            {
                if (ranges.Count >= 20000) throw new InvalidOperationException("Too many fields to map safely.");
                dynamic field = entry;
                if (deletions.Any(deleted => deleted.From <= (int)field.Code.Start - 1 && (int)field.Result.End + 1 <= deleted.To)) continue;
                if ((int)field.Result.Start >= (int)field.Result.End)
                    throw new InvalidOperationException("An empty field result cannot be mapped to a nonempty protected passage.");
                ranges.Add(new HiddenFieldRange { From = (int)field.Code.Start - 1, To = (int)field.Result.Start, Closing = false });
                ranges.Add(new HiddenFieldRange { From = (int)field.Result.End, To = (int)field.Result.End + 1, Closing = true });
                CollectFields(field.Result.Fields, ranges, depth + 1);
            }
        }

        private static InvalidOperationException MappingError()
        { return new InvalidOperationException("Word's native positions cannot be mapped safely in this document."); }

        private static string Read(dynamic document, int from, int to, NativeInspection inspection, bool extra)
        {
            if (extra) inspection.Probe();
            dynamic range = document.Range(from, to);
            range.TextRetrievalMode.IncludeHiddenText = true;
            string value = (string)range.Text ?? "";
            if ((int)range.Start != from || (int)range.End != to) throw MappingError();
            inspection.CheckTime(); return value;
        }

        private static void Append(dynamic document, int from, int to, StringBuilder text, List<int> boundaries, NativeInspection inspection)
        {
            if (to <= from) return;
            string value = Read(document, from, to, inspection, false);
            if (value.Length != to - from) inspection.Begin();
            AppendRead(document, from, to, value, text, boundaries, inspection);
        }

        private static void AppendRead(dynamic document, int from, int to, string value, StringBuilder text, List<int> boundaries, NativeInspection inspection)
        {
            inspection.CheckTime();
            if (text.Length + value.Length > PolicyEngine.MaxText) throw MappingError();
            if (value.Length == to - from)
            {
                text.Append(value);
                for (int index = from + 1; index <= to; index++) boundaries.Add(index);
            }
            else if (to == from + 1)
            {
                if (value != "\r\a" || !inspection.OwnsTableUnit(document, from, to)) throw MappingError();
                // The CR/BEL interior has no independently addressable native boundary.
                text.Append(value); boundaries.Add(boundaries[boundaries.Count - 1]); boundaries.Add(to);
            }
            else
            {
                int start = text.Length, middle = from + (to - from) / 2;
                AppendRead(document, from, middle, Read(document, from, middle, inspection, true), text, boundaries, inspection);
                inspection.CheckTime();
                AppendRead(document, middle, to, Read(document, middle, to, inspection, true), text, boundaries, inspection);
                if (text.Length - start != value.Length) throw MappingError();
                for (int index = 0; index < value.Length; index++)
                    if (text[start + index] != value[index]) throw MappingError();
            }
            inspection.CheckTime();
        }

        internal bool IsTableMarkerInterior(int visible)
        {
            if (visible < 0 || visible >= offsets.Length) throw new InvalidOperationException("Invalid visible UTF-16 range.");
            return visible > 0 && offsets[visible] == offsets[visible - 1];
        }

        internal int NativeOffset(int visible)
        {
            if (IsTableMarkerInterior(visible))
                throw new InvalidOperationException("This boundary cuts a native table marker.");
            return offsets[visible];
        }

        internal IEnumerable<int[]> NativeSegments(int from, int to)
        {
            int cursor = NativeOffset(from), end = NativeOffset(to);
            foreach (HiddenFieldRange deleted in deletions.OrderBy(interval => interval.From))
            {
                if (deleted.To <= cursor || deleted.From >= end) continue;
                if (cursor < deleted.From) yield return new[] { cursor, deleted.From };
                cursor = Math.Max(cursor, deleted.To);
            }
            if (cursor < end) yield return new[] { cursor, end };
        }

        internal dynamic Range(dynamic document, int from, int to)
        {
            if (from == to) return document.Range(NativeOffset(from), NativeOffset(to));
            var segments = NativeSegments(from, to).ToList();
            if (segments.Count == 0) throw new InvalidOperationException("The requested current text has no native range.");
            return document.Range(segments[0][0], segments[segments.Count - 1][1]);
        }

        internal int VisibleOffset(int native)
        {
            native = ResolveAlias(native);
            int result = Array.BinarySearch(offsets, native);
            while (result > 0 && offsets[result - 1] == native) result--;
            if (result < 0)
            {
                HiddenFieldRange deleted = deletions.FirstOrDefault(interval => interval.From <= native && native <= interval.To);
                if (deleted != null)
                {
                    int boundary = ResolveAlias(deleted.To);
                    result = Array.BinarySearch(offsets, boundary);
                    while (result > 0 && offsets[result - 1] == boundary) result--;
                }
            }
            if (result < 0) throw new InvalidOperationException("This native boundary cuts a field's hidden structure.");
            return result;
        }

        private int ResolveAlias(int native)
        {
            int mapped, hops = 0;
            while (aliases.TryGetValue(native, out mapped))
            {
                if (++hops > aliases.Count) throw new InvalidOperationException("Native range aliases contain a cycle.");
                native = mapped;
            }
            return native;
        }
    }
}
