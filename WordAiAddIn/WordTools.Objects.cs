using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public static partial class WordTools
    {
        // Every kind read_object understands. Each kind has its own 0-based index space, in the order
        // CollectObjects produces (document order for tables/notes/content controls/TOCs, inline-then-floating
        // for charts/SmartArt - the same order edit_chart/edit_smartart address). See WordTools.Objects.cs.md.
        private static readonly string[] ObjectKinds =
            { "table", "chart", "smartart", "shape", "content_control", "header_footer", "footnote", "endnote", "toc" };

        // read_object returns long text in windows of this many characters; offset continues (see TextWindow).
        private const int ObjectTextCap = TextWindow.DefaultCap;

        // Above this many floating shapes the counts line stops itemising them: classifying a shape costs ~10 ms of COM
        // calls (measured: 300 shapes ~3 s), so get_document_context would stall on shape-heavy documents.
        private const int ShapeScanCap = 100;

        private sealed class WordObject
        {
            public string Kind;
            public int Index;
            public string Name;          // shape / floating chart / floating SmartArt name, "" when it has none
            public int Start = -1;       // character offset in the main story, -1 when not anchored there
            public string Label;         // short tag shown next to a block in read_blocks
            public string Summary;       // one-line description used by read_object's list mode
        }

        private sealed class ShapeInfo
        {
            public dynamic Shape;
            public int Index;
            public int Type;
            public string TypeName;
            public string Name;
            public bool HasText;
            public int Start = -1;
        }

        private static string ShapeTypeName(int msoShapeType)
        {
            switch (msoShapeType)
            {
                case 1: return "shape";
                case 5: return "freeform";
                case 6: return "group";
                case 9: return "line";
                case 13: return "picture";
                case 17: return "textbox";
                case 20: return "canvas";
                default: return "type" + msoShapeType;
            }
        }

        private static int SafeShapeType(dynamic shp)
        {
            try { return (int)shp.Type; } catch { return -1; }
        }

        // A group/canvas has text when any child does; depth-capped because canvases can nest groups.
        private static bool ShapeHasText(dynamic shp, int type, int depth)
        {
            try
            {
                if (type == 6 || type == 20)
                {
                    if (depth >= 3) return false;
                    dynamic children = type == 6 ? shp.GroupItems : shp.CanvasItems;
                    foreach (dynamic child in children)
                    {
                        if (ShapeHasText(child, SafeShapeType(child), depth + 1)) return true;
                    }
                    return false;
                }
                // TextFrame throws on shapes without one (pictures, lines); that just means "no text".
                return (int)shp.TextFrame.HasText != 0;
            }
            catch { return false; }
        }

        // Character offset of a floating shape's anchor in the MAIN story (-1 for header/footer shapes etc.).
        private static int ShapeAnchorStart(dynamic shp)
        {
            try
            {
                dynamic anchor = shp.Anchor;
                if ((int)anchor.StoryType == 1 /* wdMainTextStory */) return (int)anchor.Start;
            }
            catch { }
            return -1;
        }

        // Charts and SmartArt are excluded - they have their own kinds and index spaces. Uses HasChart/HasSmartArt
        // (not just Type) so the exclusion matches ListChartShapes/ListSmartArtShapes exactly.
        private static List<ShapeInfo> ListShapes(dynamic doc)
        {
            var list = new List<ShapeInfo>();
            foreach (dynamic shp in doc.Shapes)
            {
                int type = SafeShapeType(shp);
                if (type == 3 || type == 24) continue;
                try { if ((int)shp.HasChart == -1 || (int)shp.HasSmartArt == -1) continue; } catch { }

                var info = new ShapeInfo { Shape = shp, Index = list.Count, Type = type, TypeName = ShapeTypeName(type) };
                try { info.Name = (string)shp.Name; } catch { info.Name = ""; }
                info.HasText = ShapeHasText(shp, type, 0);
                info.Start = ShapeAnchorStart(shp);
                list.Add(info);
            }
            return list;
        }

        // Chart/SmartArt entries come from ListChartShapes/ListSmartArtShapes and mix InlineShape (has Range) with
        // Shape (has Anchor).
        private static int InlineOrFloatingStart(dynamic shp)
        {
            try { return (int)shp.Range.Start; } catch { }
            return ShapeAnchorStart(shp);
        }

        // Name of a floating shape; InlineShape has none, so "" for inline charts/SmartArt.
        private static string FloatingName(dynamic shp)
        {
            try { return (string)shp.Name ?? ""; } catch { return ""; }
        }

        private static string Truncate(string s, int cap)
        {
            if (s == null) return "";
            s = s.Trim('\r', '\a', '\n', ' ');
            return s.Length > cap ? s.Substring(0, cap) + "..." : s;
        }

        private static string OneLine(string s, int cap)
        {
            return Truncate((s ?? "").Replace('\r', ' ').Replace('\n', ' ').Replace('\a', ' '), cap);
        }

        private static string ContentControlTypeName(int t)
        {
            switch (t)
            {
                case 0: return "rich text";
                case 1: return "plain text";
                case 2: return "picture";
                case 3: return "combo box";
                case 4: return "dropdown";
                case 5: return "building block";
                case 6: return "date";
                case 7: return "group";
                case 8: return "checkbox";
                case 9: return "repeating section";
                default: return "type " + t;
            }
        }

        private sealed class HeaderFooterInfo
        {
            public int Section;          // 1-based
            public string Which;         // header | footer
            public string Type;          // primary | first page | even pages
            public string Text;
        }

        // Only headers/footers that apply (first page / even pages only when the section enables them), are not
        // inherited from the previous section, and hold text.
        private static List<HeaderFooterInfo> ListHeaderFooters(dynamic doc)
        {
            var list = new List<HeaderFooterInfo>();
            int sectionIndex = 0;
            foreach (dynamic sec in doc.Sections)
            {
                sectionIndex++;
                bool firstPage = false, evenPages = false;
                try { firstPage = (int)sec.PageSetup.DifferentFirstPageHeaderFooter != 0; } catch { }
                try { evenPages = (int)sec.PageSetup.OddAndEvenPagesHeaderFooter != 0; } catch { }

                var types = new List<KeyValuePair<int, string>> { new KeyValuePair<int, string>(1, "primary") };
                if (firstPage) types.Add(new KeyValuePair<int, string>(2, "first page"));
                if (evenPages) types.Add(new KeyValuePair<int, string>(3, "even pages"));

                foreach (string which in new[] { "header", "footer" })
                {
                    dynamic parts = which == "header" ? sec.Headers : sec.Footers;
                    foreach (var t in types)
                    {
                        try
                        {
                            dynamic hf = parts.Item(t.Key);
                            if (sectionIndex > 1 && (bool)hf.LinkToPrevious) continue;
                            // Field codes on, so a page number etc. reads back as {PAGE} rather than its current value.
                            dynamic rng = hf.Range;
                            rng.TextRetrievalMode.IncludeFieldCodes = true;
                            string text = (FieldTokens.FromFieldCodes((string)rng.Text) ?? "").Trim('\r', '\a', '\n', ' ');
                            if (text.Length == 0) continue;
                            list.Add(new HeaderFooterInfo { Section = sectionIndex, Which = which, Type = t.Value, Text = text });
                        }
                        catch { }
                    }
                }
            }
            return list;
        }

        // Builds every object of the requested kind (null = all kinds). Cheap enough to call per read_blocks:
        // each kind starts with a .Count, so documents without that kind pay nothing for it.
        private static List<WordObject> CollectObjects(Word.Document doc, string onlyKind)
        {
            var all = new List<WordObject>();
            Func<string, bool> want = k => onlyKind == null || onlyKind == k;

            if (want("table") && doc.Tables.Count > 0)
            {
                int i = 0;
                foreach (Word.Table table in doc.Tables)
                {
                    string first = "";
                    try { first = OneLine(table.Cell(1, 1).Range.Text, 40); } catch { }
                    all.Add(new WordObject
                    {
                        Kind = "table", Index = i, Start = table.Range.Start, Label = "table#" + i,
                        Summary = table.Rows.Count + " rows x " + table.Columns.Count + " cols, first cell \"" + first + "\"",
                    });
                    i++;
                }
            }

            if (want("chart"))
            {
                var charts = ListChartShapes(doc);
                for (int i = 0; i < charts.Count; i++)
                {
                    string title = null;
                    try { dynamic c = charts[i].Chart; if ((bool)c.HasTitle) title = (string)c.ChartTitle.Text; } catch { }
                    all.Add(new WordObject
                    {
                        Kind = "chart", Index = i, Name = FloatingName(charts[i]), Start = InlineOrFloatingStart(charts[i]), Label = "chart#" + i,
                        Summary = title != null ? "title \"" + OneLine(title, 60) + "\"" : "no title",
                    });
                }
            }

            if (want("smartart"))
            {
                var arts = ListSmartArtShapes(doc);
                for (int i = 0; i < arts.Count; i++)
                {
                    int nodes = 0;
                    try { nodes = (int)arts[i].SmartArt.Nodes.Count; } catch { }
                    all.Add(new WordObject
                    {
                        Kind = "smartart", Index = i, Name = FloatingName(arts[i]), Start = InlineOrFloatingStart(arts[i]), Label = "smartart#" + i,
                        Summary = nodes + " node(s)",
                    });
                }
            }

            if (want("shape"))
            {
                foreach (ShapeInfo s in ListShapes(doc))
                {
                    all.Add(new WordObject
                    {
                        Kind = "shape", Index = s.Index, Name = s.Name, Start = s.Start,
                        Label = "shape#" + s.Index + " \"" + s.Name + "\" " + s.TypeName + (s.HasText ? " hasText" : ""),
                        Summary = "\"" + s.Name + "\" " + s.TypeName + (s.HasText ? " hasText" : ""),
                    });
                }
            }

            if (want("content_control") && doc.ContentControls.Count > 0)
            {
                int i = 0;
                foreach (dynamic cc in doc.ContentControls)
                {
                    string title = "", tag = "", text = "";
                    int type = -1, start = -1;
                    try { title = (string)cc.Title ?? ""; } catch { }
                    try { tag = (string)cc.Tag ?? ""; } catch { }
                    try { type = (int)cc.Type; } catch { }
                    try { start = (int)cc.Range.Start; } catch { }
                    try { text = OneLine((string)cc.Range.Text, 40); } catch { }
                    all.Add(new WordObject
                    {
                        Kind = "content_control", Index = i, Start = start,
                        Label = "content_control#" + i + (title.Length > 0 ? " \"" + title + "\"" : ""),
                        Summary = ContentControlTypeName(type) + (title.Length > 0 ? " \"" + title + "\"" : "") + (tag.Length > 0 ? " tag=" + tag : "") + ", text \"" + text + "\"",
                    });
                    i++;
                }
            }

            if (want("header_footer"))
            {
                var hfs = ListHeaderFooters(doc);
                for (int i = 0; i < hfs.Count; i++)
                {
                    all.Add(new WordObject
                    {
                        Kind = "header_footer", Index = i,
                        Summary = "section " + hfs[i].Section + " " + hfs[i].Which + " (" + hfs[i].Type + "): \"" + OneLine(hfs[i].Text, 60) + "\"",
                    });
                }
            }

            foreach (string noteKind in new[] { "footnote", "endnote" })
            {
                if (!want(noteKind)) continue;
                dynamic notes = noteKind == "footnote" ? (dynamic)doc.Footnotes : doc.Endnotes;
                if ((int)notes.Count == 0) continue;
                int i = 0;
                foreach (dynamic note in notes)
                {
                    int start = -1;
                    string text = "";
                    try { start = (int)note.Reference.Start; } catch { }
                    try { text = OneLine((string)note.Range.Text, 60); } catch { }
                    all.Add(new WordObject { Kind = noteKind, Index = i, Start = start, Label = noteKind + "#" + i, Summary = "\"" + text + "\"" });
                    i++;
                }
            }

            if (want("toc") && doc.TablesOfContents.Count > 0)
            {
                int i = 0;
                foreach (dynamic toc in doc.TablesOfContents)
                {
                    int start = -1, upper = 0, lower = 0;
                    try { start = (int)toc.Range.Start; } catch { }
                    try { upper = (int)toc.UpperHeadingLevel; lower = (int)toc.LowerHeadingLevel; } catch { }
                    all.Add(new WordObject
                    {
                        Kind = "toc", Index = i, Start = start, Label = "toc#" + i,
                        Summary = "heading levels " + upper + "-" + lower,
                    });
                    i++;
                }
            }

            return all;
        }

        // One counts-only line for get_document_context, e.g. "Objects: tables 2, shapes 6 (4 with text: textbox 3,
        // group 3)". Empty when the document has none, so plain documents' output is unchanged.
        private static string ObjectCountsLine(Word.Document doc)
        {
            var parts = new List<string>();
            Action<string, int> add = (name, n) => { if (n > 0) parts.Add(name + " " + n); };

            add("tables", doc.Tables.Count);

            int shapeTotal = doc.Shapes.Count; // a plain property: cheap however many shapes there are
            if (shapeTotal > ShapeScanCap)
            {
                // Too many to classify one by one (see ShapeScanCap). Charts and SmartArt are among them, uncounted.
                parts.Add("shapes " + shapeTotal + " (not itemised, may include charts/SmartArt - list them with read_object kind \"shape\")");
            }
            else
            {
                add("charts", ListChartShapes(doc).Count);
                add("smartart", ListSmartArtShapes(doc).Count);
                List<ShapeInfo> shapes = ListShapes(doc);
                if (shapes.Count > 0)
                {
                    string breakdown = string.Join(", ", shapes.GroupBy(s => s.TypeName).Select(g => g.Key + " " + g.Count()));
                    parts.Add("shapes " + shapes.Count + " (" + shapes.Count(s => s.HasText) + " with text: " + breakdown + ")");
                }
            }

            int pictures = 0;
            foreach (dynamic inl in doc.InlineShapes)
            {
                try { int t = (int)inl.Type; if (t == 3 || t == 4) pictures++; } catch { }
            }
            add("inline pictures", pictures);
            add("content controls", doc.ContentControls.Count);
            add("footnotes", doc.Footnotes.Count);
            add("endnotes", doc.Endnotes.Count);
            add("header/footer parts with text", ListHeaderFooters(doc).Count);
            add("TOC", doc.TablesOfContents.Count);

            return parts.Count == 0 ? "" : "Objects: " + string.Join(", ", parts);
        }

        // Anchors for read_blocks: the objects anchored inside the characters [spanStart, spanEnd) of the main story,
        // sorted by offset. Scoped to the requested range on purpose: asking Word for every object in the document cost
        // ~3.7 s per call on 300 text boxes (each shape fetch is ~3.6 ms), while the range-scoped collections take a few
        // milliseconds. Tables, content controls and notes keep their document-wide index (table#0, footnote#2 ...);
        // floating shapes/charts/SmartArt are tagged by NAME, because a shape's index would mean touching every shape
        // before it. See WordTools.Objects.cs.md.
        private static List<ObjectAnchor> CollectObjectAnchorsInRange(Word.Document doc, int spanStart, int spanEnd)
        {
            var anchors = new List<ObjectAnchor>();
            Word.Range span = doc.Range(spanStart, spanEnd);

            if (span.Tables.Count > 0)
            {
                foreach (Word.Table table in span.Tables)
                {
                    int start = table.Range.Start;
                    anchors.Add(new ObjectAnchor(start, "table#" + doc.Range(0, start).Tables.Count));
                }
            }

            if (span.ContentControls.Count > 0)
            {
                foreach (dynamic cc in span.ContentControls)
                {
                    int start = -1;
                    string title = "";
                    try { start = (int)cc.Range.Start; } catch { }
                    if (start < 0) continue;
                    try { title = (string)cc.Title ?? ""; } catch { }
                    // Range(0, start) also overlaps this control's own opening tag, so it counts itself: minus 1 (measured
                    // in Word, nested controls included). Tables need no such correction.
                    anchors.Add(new ObjectAnchor(start, "content_control#" + (doc.Range(0, start).ContentControls.Count - 1)
                        + (title.Length > 0 ? " \"" + title + "\"" : "")));
                }
            }

            if (span.Footnotes.Count > 0)
            {
                foreach (Word.Footnote note in span.Footnotes)
                    anchors.Add(new ObjectAnchor(note.Reference.Start, "footnote#" + (note.Index - 1)));
            }
            if (span.Endnotes.Count > 0)
            {
                foreach (Word.Endnote note in span.Endnotes)
                    anchors.Add(new ObjectAnchor(note.Reference.Start, "endnote#" + (note.Index - 1)));
            }

            if (doc.TablesOfContents.Count > 0) // a handful at most, so index them document-wide
            {
                int i = 0;
                foreach (dynamic toc in doc.TablesOfContents)
                {
                    try
                    {
                        int start = (int)toc.Range.Start;
                        if (start >= spanStart && start < spanEnd) anchors.Add(new ObjectAnchor(start, "toc#" + i));
                    }
                    catch { }
                    i++;
                }
            }

            dynamic floating = span.ShapeRange;
            int floatingCount = (int)floating.Count;
            for (int i = 1; i <= floatingCount; i++)
            {
                dynamic shp = floating.Item(i);
                int start = ShapeAnchorStart(shp);
                if (start < 0) continue; // anchored outside the main story
                string name = FloatingName(shp);
                bool isChart = false, isArt = false;
                try { isChart = (int)shp.HasChart == -1; } catch { }
                try { isArt = (int)shp.HasSmartArt == -1; } catch { }
                int type = SafeShapeType(shp);
                string label = isChart ? "chart \"" + name + "\""
                    : isArt ? "smartart \"" + name + "\""
                    : "shape \"" + name + "\" " + ShapeTypeName(type) + (ShapeHasText(shp, type, 0) ? " hasText" : "");
                anchors.Add(new ObjectAnchor(start, label));
            }

            if (span.InlineShapes.Count > 0) // inline pictures are only counted; inline charts/SmartArt are tagged
            {
                foreach (dynamic inl in span.InlineShapes)
                {
                    bool isChart = false, isArt = false;
                    try { isChart = (int)inl.HasChart == -1; } catch { }
                    try { isArt = !isChart && (int)inl.HasSmartArt == -1; } catch { }
                    if (!isChart && !isArt) continue;
                    int start = -1;
                    try { start = (int)inl.Range.Start; } catch { }
                    if (start < 0) continue;
                    anchors.Add(new ObjectAnchor(start, (isChart ? "chart#" : "smartart#") + InlineKindIndex(doc, isChart, start)));
                }
            }

            return anchors.OrderBy(a => a.Start).ToList();
        }

        // Index of the inline chart (or SmartArt) starting at `start` among the document's inline charts (SmartArt):
        // the same inline-first order ListChartShapes/ListSmartArtShapes use. Only reached when a range holds one.
        private static int InlineKindIndex(Word.Document doc, bool chart, int start)
        {
            int index = 0;
            foreach (dynamic inl in doc.InlineShapes)
            {
                bool match = false;
                try { match = (int)(chart ? inl.HasChart : inl.HasSmartArt) == -1; } catch { }
                if (!match) continue;
                try { if ((int)inl.Range.Start == start) return index; } catch { }
                index++;
            }
            return -1;
        }

        private static ToolResult ReadObject(JsonElement input)
        {
            string kind = input.TryGetProperty("kind", out var kEl) && kEl.ValueKind == JsonValueKind.String ? kEl.GetString() : null;
            if (kind == null || Array.IndexOf(ObjectKinds, kind) < 0)
                throw new ArgumentException("read_object: kind is required. Valid: " + string.Join(", ", ObjectKinds) + ".");
            bool hasIndex = input.TryGetProperty("index", out var iEl) && iEl.ValueKind == JsonValueKind.Number;
            int index = hasIndex ? iEl.GetInt32() : -1;
            // Long text comes back in windows; offset (characters, default 0) asks for a later one.
            int offset = input.TryGetProperty("offset", out var oEl) && oEl.ValueKind == JsonValueKind.Number ? (int)oEl.GetDouble() : 0;

            Word.Document doc = ActiveDoc;
            List<WordObject> objects = CollectObjects(doc, kind);
            string summary = "read_object " + kind;

            // A floating shape/chart/SmartArt is tagged by name in read_blocks, so accept the name in place of an index.
            string nameArg = input.TryGetProperty("name", out var nEl) && nEl.ValueKind == JsonValueKind.String ? nEl.GetString() : null;
            if (!hasIndex && !string.IsNullOrEmpty(nameArg))
            {
                if (kind != "shape" && kind != "chart" && kind != "smartart")
                    throw new ArgumentException("read_object: name only works with kind shape, chart or smartart; use index for " + kind + ".");
                WordObject named = objects.FirstOrDefault(o => string.Equals(o.Name, nameArg, StringComparison.OrdinalIgnoreCase));
                if (named == null)
                    throw new ArgumentException("read_object: no " + kind + " named \"" + nameArg + "\" (inline charts/SmartArt have no name - use index; omit index and name to list them).");
                hasIndex = true;
                index = named.Index;
            }

            if (!hasIndex)
            {
                if (objects.Count == 0)
                    return new ToolResult { Output = "No " + kind.Replace('_', ' ') + " objects in this document.", Summary = summary };

                // Paragraph numbers let the model jump from an object to read_blocks around it.
                var anchored = objects.Where(o => o.Start >= 0).OrderBy(o => o.Start).ToList();
                var paragraphOf = new Dictionary<WordObject, int>();
                if (anchored.Count > 0)
                {
                    var resolver = new ParagraphIndexResolver(doc);
                    foreach (WordObject o in anchored) paragraphOf[o] = resolver.IndexAt(o.Start);
                }

                var sb = new StringBuilder();
                sb.AppendLine(objects.Count + " " + kind + " object(s) - pass index to read one in full:");
                foreach (WordObject o in objects)
                {
                    int para;
                    sb.AppendLine("[" + o.Index + "] " + o.Summary + (paragraphOf.TryGetValue(o, out para) ? " (paragraph " + para + ")" : ""));
                }
                return new ToolResult { Output = sb.ToString().TrimEnd(), Summary = summary };
            }

            if (objects.Count == 0)
                return new ToolResult { Output = "No " + kind.Replace('_', ' ') + " objects in this document.", Summary = summary };
            if (index < 0 || index >= objects.Count)
                throw new ArgumentOutOfRangeException("index",
                    "index must be between 0 and " + (objects.Count - 1) + " (" + objects.Count + " " + kind + " object(s) in the document).");

            string output;
            switch (kind)
            {
                case "table": output = ReadTable(index); break;
                case "chart": output = ReadChart(index); break;
                case "smartart": output = ReadOneSmartArt(ListSmartArtShapes(doc)[index], index, objects.Count); break;
                case "shape": output = ReadShape(ListShapes(doc)[index], objects.Count, offset); break;
                case "content_control": output = ReadContentControl(doc, index, objects.Count, offset); break;
                case "header_footer":
                {
                    HeaderFooterInfo h = ListHeaderFooters(doc)[index];
                    output = "Header/footer " + index + " of " + objects.Count + " - section " + h.Section + " " + h.Which + " (" + h.Type + "):\n" + TextWindow.Slice(h.Text, offset, ObjectTextCap);
                    break;
                }
                case "footnote":
                case "endnote": output = ReadNote(doc, kind, index, objects.Count, offset); break;
                default: output = ReadToc(doc, index, objects.Count, offset); break;
            }

            WordObject picked = objects[index];
            if (picked.Start >= 0)
                output += "\n(anchored in paragraph " + new ParagraphIndexResolver(doc).IndexAt(picked.Start) + ")";
            return new ToolResult { Output = output, Summary = summary };
        }

        private static string ReadShape(ShapeInfo s, int total, int offset)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Shape " + s.Index + " of " + total + ": \"" + s.Name + "\" " + s.TypeName + (s.HasText ? " hasText" : ""));
            string alt = "";
            try { alt = (string)s.Shape.AlternativeText ?? ""; } catch { }
            if (alt.Length > 0) sb.AppendLine("Alt text: " + OneLine(alt, 200));
            var body = new StringBuilder();
            AppendShapeText(body, s.Shape, s.Type, 0, "");
            sb.Append(TextWindow.Slice(body.ToString().TrimEnd(), offset, ObjectTextCap));
            return sb.ToString().TrimEnd();
        }

        // Text boxes/autoshapes print their text; groups and canvases recurse into their children (depth-capped).
        private static void AppendShapeText(StringBuilder sb, dynamic shp, int type, int depth, string indent)
        {
            if (type == 6 || type == 20)
            {
                if (depth >= 3) return;
                int n = 0;
                try
                {
                    dynamic children = type == 6 ? shp.GroupItems : shp.CanvasItems;
                    foreach (dynamic child in children)
                    {
                        int ct = SafeShapeType(child);
                        string name = "";
                        try { name = (string)child.Name; } catch { }
                        sb.AppendLine(indent + "- [" + n + "] \"" + name + "\" " + ShapeTypeName(ct));
                        AppendShapeText(sb, child, ct, depth + 1, indent + "  ");
                        n++;
                    }
                }
                catch { }
                return;
            }
            try
            {
                if ((int)shp.TextFrame.HasText == 0) return;
                sb.AppendLine(indent + ((string)shp.TextFrame.TextRange.Text ?? "").Trim('\r', '\a', '\n', ' '));
            }
            catch { }
        }

        private static string ReadContentControl(Word.Document doc, int index, int total, int offset)
        {
            int i = 0;
            foreach (dynamic cc in doc.ContentControls)
            {
                if (i++ != index) continue;
                string title = "", tag = "", text = "";
                int type = -1;
                bool placeholder = false;
                try { title = (string)cc.Title ?? ""; } catch { }
                try { tag = (string)cc.Tag ?? ""; } catch { }
                try { type = (int)cc.Type; } catch { }
                try { placeholder = (bool)cc.ShowingPlaceholderText; } catch { }
                try { text = TextWindow.Slice(((string)cc.Range.Text ?? "").Trim('\r', '\a', '\n', ' '), offset, ObjectTextCap); } catch { }
                var sb = new StringBuilder();
                sb.AppendLine("Content control " + index + " of " + total + ": " + ContentControlTypeName(type) +
                    (title.Length > 0 ? ", title \"" + title + "\"" : "") + (tag.Length > 0 ? ", tag \"" + tag + "\"" : "") +
                    (placeholder ? " (showing placeholder text)" : ""));
                sb.Append(text);
                return sb.ToString();
            }
            throw new ArgumentOutOfRangeException("index");
        }

        private static string ReadNote(Word.Document doc, string kind, int index, int total, int offset)
        {
            dynamic notes = kind == "footnote" ? (dynamic)doc.Footnotes : doc.Endnotes;
            int i = 0;
            foreach (dynamic note in notes)
            {
                if (i++ != index) continue;
                string text = "";
                try { text = TextWindow.Slice(((string)note.Range.Text ?? "").Trim('\r', '\a', '\n', ' '), offset, ObjectTextCap); } catch { }
                return char.ToUpper(kind[0]) + kind.Substring(1) + " " + index + " of " + total + ":\n" + text;
            }
            throw new ArgumentOutOfRangeException("index");
        }

        private static string ReadToc(Word.Document doc, int index, int total, int offset)
        {
            int i = 0;
            foreach (dynamic toc in doc.TablesOfContents)
            {
                if (i++ != index) continue;
                int upper = 0, lower = 0;
                string text = "";
                try { upper = (int)toc.UpperHeadingLevel; lower = (int)toc.LowerHeadingLevel; } catch { }
                try { text = TextWindow.Slice(((string)toc.Range.Text ?? "").Trim('\r', '\a', '\n', ' '), offset, ObjectTextCap); } catch { }
                return "TOC " + index + " of " + total + " (heading levels " + upper + "-" + lower + "):\n" + text;
            }
            throw new ArgumentOutOfRangeException("index");
        }
    }
}
