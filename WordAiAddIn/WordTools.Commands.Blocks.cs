using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    // Spike 3: real COM tool execution against the live Word document, called
    // from the WebView2-hosted AgentLoop via the JSON WebMessage bridge.
    public static partial class WordTools
    {
        private static void DeleteBlocksCmd(JsonElement cmd)
        {
            var matches = ResolveTargetParagraphs(cmd.GetProperty("target"));
            if (matches.Count == 0)
            {
                throw new InvalidOperationException("deleteBlocks: no paragraphs matched target.");
            }
            if (matches.Count >= ActiveDoc.Paragraphs.Count)
            {
                // Deleting every paragraph would leave zero - clear content instead,
                // leaving one empty paragraph (mirrors genoffice's own guard).
                ActiveDoc.Content.Text = "";
                return;
            }
            // Delete in descending order so an earlier Paragraph object's Range
            // is never asked to reason about content after it disappearing.
            matches.Sort((a, b) => b.Index.CompareTo(a.Index));
            foreach (var (_, p) in matches)
            {
                p.Range.Delete();
            }
        }

        private static void MoveBlocksCmd(JsonElement cmd)
        {
            var blockIndexes = new List<int>();
            foreach (JsonElement e in cmd.GetProperty("blockIndexes").EnumerateArray()) blockIndexes.Add(e.GetInt32());
            if (blockIndexes.Count == 0)
            {
                throw new ArgumentException("moveBlocks: blockIndexes must contain at least one index.");
            }
            int afterBlockIndex = cmd.GetProperty("afterBlockIndex").GetInt32();

            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            int count = paragraphs.Count;
            if (blockIndexes.Any(i => i < 0 || i >= count) || afterBlockIndex < -1 || afterBlockIndex >= count)
            {
                throw new ArgumentException("moveBlocks: index out of range.");
            }
            if (blockIndexes.Contains(afterBlockIndex))
            {
                throw new ArgumentException("moveBlocks: afterBlockIndex cannot be one of the moved blocks.");
            }

            blockIndexes.Sort();
            // Capture each moved paragraph's content as an OOXML string - a true,
            // detached snapshot (plain text, not a live COM Range reference) -
            // before any deletion shifts indices, so formatting survives the
            // move without depending on FormattedText's live-vs-copy semantics.
            var captured = blockIndexes.Select(i => paragraphs[i + 1].Range.WordOpenXML).ToList();

            // Delete moved paragraphs in descending order.
            var deleteOrder = new List<int>(blockIndexes);
            deleteOrder.Reverse();
            foreach (int i in deleteOrder)
            {
                paragraphs[i + 1].Range.Delete();
            }

            // Recompute the insertion point: afterBlockIndex shifts down by however
            // many moved blocks were originally BEFORE it.
            int shift = blockIndexes.Count(i => i < afterBlockIndex);
            int adjustedAfter = afterBlockIndex - shift;

            Word.Range insertionPoint = adjustedAfter == -1
                ? ActiveDoc.Range(0, 0)
                : ActiveDoc.Paragraphs[adjustedAfter + 1].Range;
            insertionPoint.Collapse(adjustedAfter == -1 ? Word.WdCollapseDirection.wdCollapseStart : Word.WdCollapseDirection.wdCollapseEnd);

            foreach (string xml in captured)
            {
                insertionPoint.InsertXML(xml);
                insertionPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
        }

        private static void UpdateImageProperties(JsonElement cmd)
        {
            int imageIndex = cmd.GetProperty("imageIndex").GetInt32(); // 0-based index into doc.InlineShapes
            Word.InlineShapes shapes = ActiveDoc.InlineShapes;
            if (imageIndex < 0 || imageIndex >= shapes.Count)
            {
                throw new ArgumentException("updateImageProperties: imageIndex out of range.");
            }
            Word.InlineShape shape = shapes[imageIndex + 1];
            JsonElement properties = cmd.GetProperty("properties");
            HashSet<string> fields = new HashSet<string>();
            foreach (JsonElement f in cmd.GetProperty("fields").EnumerateArray()) fields.Add(f.GetString());

            const float pxToPoints = 0.75f; // 96dpi px -> points, matches genoffice's own pixel model
            float? newWidth = null, newHeight = null;
            if (fields.Contains("widthPx") && properties.TryGetProperty("widthPx", out var w) && w.ValueKind == JsonValueKind.Number)
                newWidth = (float)w.GetDouble() * pxToPoints;
            if (fields.Contains("heightPx") && properties.TryGetProperty("heightPx", out var h) && h.ValueKind == JsonValueKind.Number)
                newHeight = (float)h.GetDouble() * pxToPoints;

            if (newWidth.HasValue && !newHeight.HasValue)
            {
                newHeight = shape.Height * (newWidth.Value / shape.Width); // proportional scale from current size
            }
            else if (newHeight.HasValue && !newWidth.HasValue)
            {
                newWidth = shape.Width * (newHeight.Value / shape.Height);
            }
            if (newWidth.HasValue) shape.Width = newWidth.Value;
            if (newHeight.HasValue) shape.Height = newHeight.Value;

            if (fields.Contains("align") && properties.TryGetProperty("align", out var align) && align.ValueKind == JsonValueKind.String)
            {
                switch (align.GetString())
                {
                    case "left": shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;
                    case "center": shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;
                    case "right": shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; break;
                }
            }
        }

        private static void InsertTocCmd(JsonElement cmd)
        {
            int afterBlockIndex = cmd.GetProperty("afterBlockIndex").GetInt32();
            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            bool hasHeadings = false;
            foreach (Word.Paragraph p in paragraphs)
            {
                if (p.Range.get_Style().NameLocal.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) { hasHeadings = true; break; }
            }
            if (!hasHeadings)
            {
                throw new InvalidOperationException("insertToc: document has no heading-styled paragraphs to build a table of contents from.");
            }
            if (afterBlockIndex < -1 || afterBlockIndex >= paragraphs.Count)
            {
                throw new ArgumentException("insertToc: afterBlockIndex out of range.");
            }

            Word.Range insertionPoint = afterBlockIndex == -1
                ? ActiveDoc.Range(0, 0)
                : paragraphs[afterBlockIndex + 1].Range;
            insertionPoint.Collapse(afterBlockIndex == -1 ? Word.WdCollapseDirection.wdCollapseStart : Word.WdCollapseDirection.wdCollapseEnd);
            insertionPoint.InsertParagraphAfter();
            insertionPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            // Word's own native TOC field - auto-scans heading-styled paragraphs and
            // produces real, page-numbered entries directly. This is a more direct,
            // simpler native equivalent than genoffice's own hand-built TOC field-XML
            // workaround (real Word already paginates; genoffice's web renderer doesn't).
            ActiveDoc.TablesOfContents.Add(insertionPoint, UseHeadingStyles: true);
        }
    }
}

