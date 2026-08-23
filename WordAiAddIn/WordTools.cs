using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    // Task 11: server-side editing-mode gate. Defaults to FullAutonomy so
    // existing spike behavior (Tasks 8-10) is unchanged until a user
    // explicitly picks a more restrictive mode from the chat-ui mode menu.
    public enum EditingMode { ReadOnly, CommentOnly, TrackChanges, FullAutonomy }

    // Spike 3: real COM tool execution against the live Word document, called
    // from the WebView2-hosted AgentLoop via the JSON WebMessage bridge.
    public static class WordTools
    {
        public static EditingMode Mode = EditingMode.FullAutonomy;

        // Tools that are always safe to run regardless of editing mode - they
        // never touch document content. Everything else is gated below.
        private static readonly HashSet<string> AlwaysAllowedTools = new HashSet<string>
        {
            "get_document_context", "read_blocks",
        };

        public static ToolResult Execute(string name, JsonElement input)
        {
            try
            {
                bool isAlwaysAllowed = AlwaysAllowedTools.Contains(name);
                bool isAddComment = name == "add_comment";
                // "Mutating" here means "changes document content/structure" -
                // used only to decide whether TrackRevisions should be toggled.
                // add_comment does not mutate document content (it adds a
                // comment annotation), so it's excluded from this set even
                // though it is gated like a mutating tool below.
                bool isContentMutating = !isAlwaysAllowed && !isAddComment;

                if (Mode == EditingMode.ReadOnly && !isAlwaysAllowed)
                {
                    return new ToolResult { Output = "Blocked: editing mode is Read Only.", IsError = true, Summary = name };
                }
                if (Mode == EditingMode.CommentOnly && !isAlwaysAllowed && !isAddComment)
                {
                    return new ToolResult { Output = "Blocked: editing mode is Comment Only - use add_comment instead of editing content directly.", IsError = true, Summary = name };
                }

                if (isContentMutating)
                {
                    ActiveDoc.TrackRevisions = (Mode == EditingMode.TrackChanges);
                }

                switch (name)
                {
                    case "get_document_context":
                        return GetDocumentContext();
                    case "insert_content":
                        return InsertContent(input);
                    case "edit_chart":
                        return EditChart(input);
                    case "read_blocks":
                        return ReadBlocks(input);
                    case "replace_blocks":
                        return ReplaceBlocks(input);
                    case "apply_commands":
                        return ApplyCommands(input);
                    case "add_comment":
                        return AddComment(input);
                    default:
                        return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        private static ToolResult AddComment(JsonElement input)
        {
            string anchorText = input.GetProperty("anchorText").GetString();
            string commentText = input.GetProperty("commentText").GetString();
            Word.Document doc = ActiveDoc;
            Word.Range range = doc.Content;
            range.Find.ClearFormatting();
            range.Find.Text = anchorText;
            bool found = range.Find.Execute();
            if (!found)
            {
                return new ToolResult { Output = $"Could not find text to anchor comment: '{anchorText}'", IsError = true, Summary = "add_comment" };
            }
            doc.Comments.Add(range, commentText);
            return new ToolResult { Output = "Comment added.", Mutated = true, Summary = "add_comment" };
        }

        private static Word.Document ActiveDoc => Globals.ThisAddIn.Application.ActiveDocument;

        private static ToolResult GetDocumentContext()
        {
            Word.Document doc = ActiveDoc;
            int paraCount = doc.Paragraphs.Count;
            int wordCount = doc.Words.Count;
            string preview = doc.Content.Text;
            if (preview.Length > 300) preview = preview.Substring(0, 300) + "...";
            string output = $"Paragraphs: {paraCount}, Words: {wordCount}\nPreview: {preview}";
            return new ToolResult { Output = output, Summary = "read document context" };
        }

        private static ToolResult InsertContent(JsonElement input)
        {
            string text = input.TryGetProperty("text", out var t) ? t.GetString() : "(no text provided)";
            Word.Document doc = ActiveDoc;
            Word.Range range = doc.Content;
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.Text = text;
            return new ToolResult
            {
                Output = "Inserted text at end of document: " + text,
                Mutated = true,
                Summary = "insert_content",
            };
        }

        // dynamic: Word's chart object model (Shapes.AddChart2 / Chart / SeriesCollection) mirrors
        // Excel/PowerPoint's shared chart engine; using dynamic avoids pinning down the exact
        // Interop type names for this spike and lets any signature mismatch surface immediately at
        // runtime instead of guessing overloads at compile time.
        private static ToolResult EditChart(JsonElement input)
        {
            string title = input.TryGetProperty("title", out var t) ? t.GetString() : "Chart";
            double[] values = input.TryGetProperty("values", out var v) && v.ValueKind == JsonValueKind.Array
                ? v.EnumerateArray().Select(x => x.GetDouble()).ToArray()
                : new double[] { 1, 2, 3 };

            dynamic doc = ActiveDoc;
            dynamic chartShape = null;
            foreach (dynamic shp in doc.Shapes)
            {
                if ((int)shp.HasChart == -1 /* msoTrue */)
                {
                    chartShape = shp;
                    break;
                }
            }
            bool created = false;
            if (chartShape == null)
            {
                // 51 = xlColumnClustered
                chartShape = doc.Shapes.AddChart2(-1, 51, 0, 0, 300, 200);
                created = true;
            }

            dynamic chart = chartShape.Chart;
            chart.HasTitle = true;
            chart.ChartTitle.Text = title;
            dynamic series = chart.SeriesCollection(1);
            series.Values = values;

            return new ToolResult
            {
                Output = $"Chart {(created ? "created" : "updated")}: title='{title}', values=[{string.Join(", ", values)}]",
                Mutated = true,
                Summary = "edit_chart",
            };
        }

        private static ToolResult ReadBlocks(JsonElement input)
        {
            int startIndex = input.GetProperty("startIndex").GetInt32();
            int endIndex = input.GetProperty("endIndex").GetInt32();
            Word.Document doc = ActiveDoc;
            Word.Paragraphs paragraphs = doc.Paragraphs;
            int count = paragraphs.Count;
            endIndex = Math.Min(endIndex, count - 1);
            if (startIndex < 0 || startIndex > endIndex)
            {
                return new ToolResult { Output = "Invalid range.", IsError = true, Summary = "read_blocks" };
            }
            var sb = new System.Text.StringBuilder();
            for (int i = startIndex; i <= endIndex; i++)
            {
                Word.Paragraph p = paragraphs[i + 1];
                string text = p.Range.Text.TrimEnd('\r', '\a', '\n');
                sb.AppendLine($"[{i}] {text}");
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_blocks" };
        }

        private static ToolResult ReplaceBlocks(JsonElement input)
        {
            int startIndex = input.GetProperty("startIndex").GetInt32();
            int endIndex = input.GetProperty("endIndex").GetInt32();
            string text = input.GetProperty("text").GetString() ?? "";
            Word.Document doc = ActiveDoc;
            Word.Paragraphs paragraphs = doc.Paragraphs;
            endIndex = Math.Min(endIndex, paragraphs.Count - 1);
            if (startIndex < 0 || startIndex > endIndex)
            {
                return new ToolResult { Output = "Invalid range.", IsError = true, Summary = "replace_blocks" };
            }
            Word.Range range = doc.Range(paragraphs[startIndex + 1].Range.Start, paragraphs[endIndex + 1].Range.End);
            range.Text = text;
            return new ToolResult { Output = $"Replaced paragraphs {startIndex}-{endIndex} with: {text}", Mutated = true, Summary = "replace_blocks" };
        }

        private static ToolResult ApplyCommands(JsonElement input)
        {
            var lines = new System.Text.StringBuilder();
            bool anyMutated = false;
            bool anyError = false;
            foreach (JsonElement cmd in input.GetProperty("commands").EnumerateArray())
            {
                string kind = cmd.GetProperty("kind").GetString();
                try
                {
                    switch (kind)
                    {
                        case "set_bold":
                            SetRunProperty(cmd, (range, value) => range.Bold = value ? 1 : 0);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_italic":
                            SetRunProperty(cmd, (range, value) => range.Italic = value ? 1 : 0);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_heading":
                            SetHeading(cmd);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "find_replace":
                            int replacements = FindReplace(cmd);
                            lines.AppendLine($"{kind}: {replacements} replacement(s)");
                            if (replacements > 0) anyMutated = true;
                            break;
                        case "updateTextStyle":
                            UpdateTextStyle(cmd);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "updateParagraphStyle":
                            UpdateParagraphStyle(cmd);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "deleteBlocks":
                            DeleteBlocksCmd(cmd);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "moveBlocks":
                            MoveBlocksCmd(cmd);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "createParagraphBullets":
                            CreateParagraphBullets(cmd);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "deleteParagraphBullets":
                            DeleteParagraphBullets(cmd);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "updateImageProperties":
                            UpdateImageProperties(cmd);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "insertToc":
                            InsertTocCmd(cmd);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        default:
                            lines.AppendLine(kind + ": unknown command kind"); anyError = true; break;
                    }
                }
                catch (Exception ex)
                {
                    lines.AppendLine(kind + ": ERROR - " + ex.Message); anyError = true;
                }
            }
            return new ToolResult { Output = lines.ToString(), Mutated = anyMutated, IsError = anyError, Summary = "apply_commands" };
        }

        private static void SetRunProperty(JsonElement cmd, Action<Word.Range, bool> apply)
        {
            int startIndex = cmd.GetProperty("startIndex").GetInt32();
            int endIndex = cmd.GetProperty("endIndex").GetInt32();
            bool value = cmd.GetProperty("value").GetBoolean();
            Word.Document doc = ActiveDoc;
            Word.Paragraphs paragraphs = doc.Paragraphs;
            endIndex = Math.Min(endIndex, paragraphs.Count - 1);
            Word.Range range = doc.Range(paragraphs[startIndex + 1].Range.Start, paragraphs[endIndex + 1].Range.End);
            apply(range, value);
        }

        private static void SetHeading(JsonElement cmd)
        {
            int index = cmd.GetProperty("index").GetInt32();
            int level = cmd.GetProperty("level").GetInt32();
            Word.Paragraph p = ActiveDoc.Paragraphs[index + 1];
            p.Range.set_Style(level == 0 ? "Normal" : "Heading " + level);
        }

        private static int FindReplace(JsonElement cmd)
        {
            string find = cmd.GetProperty("find").GetString();
            string replace = cmd.GetProperty("replace").GetString();
            bool matchCase = cmd.TryGetProperty("matchCase", out var mc) && mc.GetBoolean();
            Word.Find findObj = ActiveDoc.Content.Find;
            findObj.ClearFormatting();
            findObj.Text = find;
            findObj.Replacement.ClearFormatting();
            findObj.Replacement.Text = replace;
            findObj.MatchCase = matchCase;
            bool found = findObj.Execute(Replace: Word.WdReplace.wdReplaceAll);
            return found ? 1 : 0;
        }

        private static List<int> ResolveTargetParagraphs(JsonElement target)
        {
            string nodeType = target.TryGetProperty("nodeType", out var nt) && nt.ValueKind == JsonValueKind.String ? nt.GetString() : null;
            int? headingLevel = target.TryGetProperty("headingLevel", out var hl) && hl.ValueKind == JsonValueKind.Number ? hl.GetInt32() : (int?)null;
            string containsText = target.TryGetProperty("containsText", out var ct) && ct.ValueKind == JsonValueKind.String ? ct.GetString() : null;
            bool matchCase = target.TryGetProperty("matchCase", out var mc) && mc.ValueKind == JsonValueKind.True;
            HashSet<int> blockIndexes = null;
            if (target.TryGetProperty("blockIndexes", out var bi) && bi.ValueKind == JsonValueKind.Array)
            {
                blockIndexes = new HashSet<int>();
                foreach (JsonElement e in bi.EnumerateArray()) blockIndexes.Add(e.GetInt32());
            }
            string scope = target.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String ? sc.GetString() : "document";

            if (nodeType == null && containsText == null && blockIndexes == null)
            {
                throw new ArgumentException("Target must specify at least one of nodeType, containsText, or blockIndexes.");
            }

            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            int selStart = -1, selEnd = -1;
            if (scope == "selection")
            {
                Word.Selection sel = Globals.ThisAddIn.Application.Selection;
                if (sel.Type != Word.WdSelectionType.wdNoSelection)
                {
                    selStart = sel.Range.Start;
                    selEnd = sel.Range.End;
                }
            }

            var result = new List<int>();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                Word.Paragraph p = paragraphs[i + 1];

                if (scope == "selection")
                {
                    if (selStart == -1) continue;
                    if (p.Range.Start > selEnd || p.Range.End < selStart) continue;
                }

                if (blockIndexes != null && !blockIndexes.Contains(i)) continue;

                string styleName = p.Range.get_Style().NameLocal;
                bool isHeading = styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
                bool isListItem = p.Range.ListFormat.ListType != Word.WdListType.wdListNoNumbering;

                if (nodeType == "heading" && !isHeading) continue;
                if (nodeType == "paragraph" && (isHeading || isListItem)) continue;
                if (nodeType == "listItem" && !isListItem) continue;

                if (nodeType == "heading" && headingLevel.HasValue)
                {
                    string levelDigits = new string(styleName.Where(char.IsDigit).ToArray());
                    if (!int.TryParse(levelDigits, out int actualLevel) || actualLevel != headingLevel.Value) continue;
                }

                if (containsText != null)
                {
                    string text = p.Range.Text ?? "";
                    bool found = matchCase
                        ? text.Contains(containsText)
                        : text.IndexOf(containsText, StringComparison.OrdinalIgnoreCase) >= 0;
                    if (!found) continue;
                }

                result.Add(i);
            }
            return result;
        }

        private static void UpdateTextStyle(JsonElement cmd)
        {
            List<int> indexes = ResolveTargetParagraphs(cmd.GetProperty("target"));
            JsonElement style = cmd.GetProperty("style");
            HashSet<string> fields = new HashSet<string>();
            foreach (JsonElement f in cmd.GetProperty("fields").EnumerateArray()) fields.Add(f.GetString());

            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            foreach (int i in indexes)
            {
                Word.Range range = paragraphs[i + 1].Range;
                if (fields.Contains("bold") && style.TryGetProperty("bold", out var bold))
                    range.Font.Bold = bold.ValueKind == JsonValueKind.True ? 1 : 0;
                if (fields.Contains("italic") && style.TryGetProperty("italic", out var italic))
                    range.Font.Italic = italic.ValueKind == JsonValueKind.True ? 1 : 0;
                if (fields.Contains("underline") && style.TryGetProperty("underline", out var underline))
                    range.Font.Underline = underline.ValueKind == JsonValueKind.True ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone;
                if (fields.Contains("strike") && style.TryGetProperty("strike", out var strike))
                    range.Font.StrikeThrough = strike.ValueKind == JsonValueKind.True ? 1 : 0;
                if (fields.Contains("sizeHalfPoints") && style.TryGetProperty("sizeHalfPoints", out var size) && size.ValueKind == JsonValueKind.Number)
                    range.Font.Size = (float)(size.GetDouble() / 2.0);
                if (fields.Contains("font") && style.TryGetProperty("font", out var font) && font.ValueKind == JsonValueKind.String)
                    range.Font.Name = font.GetString();
                if (fields.Contains("color") && style.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
                    range.Font.Color = HexToWdColor(color.GetString());
                if (fields.Contains("baselineOffset") && style.TryGetProperty("baselineOffset", out var baseline) && baseline.ValueKind == JsonValueKind.String)
                {
                    string b = baseline.GetString();
                    range.Font.Superscript = b == "SUPERSCRIPT" ? 1 : 0;
                    range.Font.Subscript = b == "SUBSCRIPT" ? 1 : 0;
                }
                if (fields.Contains("link") && style.TryGetProperty("link", out var link) && link.ValueKind == JsonValueKind.Object)
                {
                    string url = link.GetProperty("url").GetString();
                    ActiveDoc.Hyperlinks.Add(range, url);
                }
            }
        }

        private static Word.WdColor HexToWdColor(string hex)
        {
            hex = hex.TrimStart('#');
            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = Convert.ToInt32(hex.Substring(4, 2), 16);
            return (Word.WdColor)System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
        }

        private static void UpdateParagraphStyle(JsonElement cmd)
        {
            List<int> indexes = ResolveTargetParagraphs(cmd.GetProperty("target"));
            JsonElement style = cmd.GetProperty("style");
            HashSet<string> fields = new HashSet<string>();
            foreach (JsonElement f in cmd.GetProperty("fields").EnumerateArray()) fields.Add(f.GetString());

            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            foreach (int i in indexes)
            {
                Word.ParagraphFormat fmt = paragraphs[i + 1].Format;
                if (fields.Contains("align") && style.TryGetProperty("align", out var align) && align.ValueKind == JsonValueKind.String)
                {
                    switch (align.GetString())
                    {
                        case "left": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;
                        case "center": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;
                        case "right": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; break;
                        case "justify": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify; break;
                    }
                }
                if (fields.Contains("lineSpacing") && style.TryGetProperty("lineSpacing", out var ls) && ls.ValueKind == JsonValueKind.Number)
                    fmt.LineSpacing = (float)ls.GetDouble();
                if (fields.Contains("indentLeft") && style.TryGetProperty("indentLeft", out var il) && il.ValueKind == JsonValueKind.Number)
                    fmt.LeftIndent = (float)il.GetDouble();
                if (fields.Contains("indentRight") && style.TryGetProperty("indentRight", out var ir) && ir.ValueKind == JsonValueKind.Number)
                    fmt.RightIndent = (float)ir.GetDouble();
                if (fields.Contains("indentFirstLine") && style.TryGetProperty("indentFirstLine", out var ifl) && ifl.ValueKind == JsonValueKind.Number)
                    fmt.FirstLineIndent = (float)ifl.GetDouble();
                if (fields.Contains("spaceBefore") && style.TryGetProperty("spaceBefore", out var sb) && sb.ValueKind == JsonValueKind.Number)
                    fmt.SpaceBefore = (float)sb.GetDouble();
                if (fields.Contains("spaceAfter") && style.TryGetProperty("spaceAfter", out var sa) && sa.ValueKind == JsonValueKind.Number)
                    fmt.SpaceAfter = (float)sa.GetDouble();
                if (fields.Contains("pageBreakBefore") && style.TryGetProperty("pageBreakBefore", out var pbb))
                    fmt.PageBreakBefore = pbb.ValueKind == JsonValueKind.True ? 1 : 0;
                if (fields.Contains("shadingFill") && style.TryGetProperty("shadingFill", out var shading) && shading.ValueKind == JsonValueKind.String)
                    paragraphs[i + 1].Shading.BackgroundPatternColor = HexToWdColor(shading.GetString());
                if (fields.Contains("borders") && style.TryGetProperty("borders", out var borders))
                {
                    bool on = borders.ValueKind == JsonValueKind.True;
                    foreach (Word.Border border in paragraphs[i + 1].Borders)
                    {
                        border.LineStyle = on ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleNone;
                    }
                }
            }
        }

        private static void DeleteBlocksCmd(JsonElement cmd)
        {
            List<int> indexes = ResolveTargetParagraphs(cmd.GetProperty("target"));
            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            if (indexes.Count >= paragraphs.Count)
            {
                // Deleting every paragraph would leave zero - clear content instead,
                // leaving one empty paragraph (mirrors genoffice's own guard).
                ActiveDoc.Content.Text = "";
                return;
            }
            // Delete in descending order so earlier indices don't shift as later ones are removed.
            indexes.Sort();
            indexes.Reverse();
            foreach (int i in indexes)
            {
                paragraphs[i + 1].Range.Delete();
            }
        }

        private static void MoveBlocksCmd(JsonElement cmd)
        {
            var blockIndexes = new List<int>();
            foreach (JsonElement e in cmd.GetProperty("blockIndexes").EnumerateArray()) blockIndexes.Add(e.GetInt32());
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

        private static void CreateParagraphBullets(JsonElement cmd)
        {
            List<int> indexes = ResolveTargetParagraphs(cmd.GetProperty("target"));
            string preset = cmd.TryGetProperty("bulletPreset", out var bp) && bp.ValueKind == JsonValueKind.String ? bp.GetString() : "BULLET";
            bool numbered = preset.StartsWith("NUMBERED", StringComparison.OrdinalIgnoreCase);

            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            foreach (int i in indexes)
            {
                Word.Range range = paragraphs[i + 1].Range;
                string styleName = range.get_Style().NameLocal;
                if (styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) continue; // headings are matched but left unchanged, mirrors genoffice
                if (numbered) range.ListFormat.ApplyNumberDefault();
                else range.ListFormat.ApplyBulletDefault();
            }
        }

        private static void DeleteParagraphBullets(JsonElement cmd)
        {
            List<int> indexes = ResolveTargetParagraphs(cmd.GetProperty("target"));
            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            foreach (int i in indexes)
            {
                Word.Range range = paragraphs[i + 1].Range;
                if (range.ListFormat.ListType == Word.WdListType.wdListNoNumbering) continue; // non-list-item matches silently skipped, mirrors genoffice
                range.ListFormat.RemoveNumbers();
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
