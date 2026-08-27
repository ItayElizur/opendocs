using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // PP-23 Task 8 (post-hoc fix): previously only checked HasTextFrame,
        // so read_slide/get_deck_context reported table and SmartArt shapes
        // with no content at all - the model could add_table/add_smartart
        // and then have no way to see what it just created. Table reading
        // reuses the same statically-typed Table.Cell(r,c).Shape.TextFrame
        // pattern AddTable/EditTableCell already use in this file; SmartArt
        // is dynamic (HasSmartArt/.SmartArt aren't on the statically-typed
        // Shape interface), matching AddSmartArt's own existing pattern.
        private static string ShapeText(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                return shape.TextFrame.TextRange.Text;
            }
            if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                PowerPoint.Table table = shape.Table;
                var rowsOut = new List<string>();
                for (int r = 1; r <= table.Rows.Count; r++)
                {
                    var cellsOut = new List<string>();
                    for (int c = 1; c <= table.Columns.Count; c++)
                    {
                        cellsOut.Add(table.Cell(r, c).Shape.TextFrame.TextRange.Text.Replace("\r", " ").Trim());
                    }
                    rowsOut.Add(string.Join(" | ", cellsOut));
                }
                return "[table " + table.Rows.Count + "x" + table.Columns.Count + ": " + string.Join(" / ", rowsOut) + "]";
            }
            dynamic dshape = shape;
            bool hasSmartArt = false;
            // Post-hoc fix (2026-08-24, PP-23): HasSmartArt is an MsoTriState,
            // not a real bool - see WordTools.cs's ListSmartArtShapes for the
            // identical fix and the confirmed symptom this caused there.
            try { hasSmartArt = (int)dshape.HasSmartArt == -1 /* msoTrue */; } catch { }
            if (hasSmartArt)
            {
                dynamic nodes = dshape.SmartArt.Nodes;
                int count = (int)nodes.Count;
                var nodeTexts = new List<string>();
                for (int i = 1; i <= count; i++)
                {
                    try { nodeTexts.Add(((string)nodes.Item(i).TextFrame2.TextRange.Text).Replace("\r", " ").Trim()); }
                    catch { }
                }
                return "[SmartArt " + count + " node(s): " + string.Join(", ", nodeTexts) + "]";
            }
            return "";
        }

        private static ToolResult GetDeckContext()
        {
            var sb = new StringBuilder();
            int i = 0;
            foreach (PowerPoint.Slide slide in ActivePresentation.Slides)
            {
                var texts = new System.Collections.Generic.List<string>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    string t = ShapeText(shape).Replace("\r", " ").Trim();
                    if (t.Length > 0) texts.Add(t);
                }
                string preview = string.Join(" | ", texts);
                if (preview.Length > 120) preview = preview.Substring(0, 120) + "...";
                sb.AppendLine($"[{i}] {preview}");
                i++;
            }
            return new ToolResult { Output = sb.ToString(), Summary = "get_deck_context" };
        }

        private static ToolResult ReadSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (slideIndex < 0 || slideIndex >= slides.Count)
            {
                return new ToolResult { Output = "Invalid slide index.", IsError = true, Summary = "read_slide" };
            }
            PowerPoint.Slide slide = slides[slideIndex + 1];
            var sb = new StringBuilder();

            // PP-24: surfaces layout/transition/animation-count so the model
            // isn't blind to state set_slide_layout/set_slide_transition/
            // add_animation just created - same "the model can add
            // something but then can't see it" gap PP-23's read_chart/
            // read_table/read_smartart all exist to close.
            string layoutName = null;
            foreach (var kv in SlideLayoutMap) { if (kv.Value == slide.Layout) { layoutName = kv.Key; break; } }
            if (layoutName != null) sb.AppendLine("Layout: " + layoutName);
            else if (slide.Layout == PowerPoint.PpSlideLayout.ppLayoutCustom) sb.AppendLine("Layout: custom ('" + slide.CustomLayout.Name + "')");
            else sb.AppendLine("Layout: " + slide.Layout);

            string transitionName = null;
            foreach (var kv in TransitionEffectMap) { if (kv.Value == slide.SlideShowTransition.EntryEffect) { transitionName = kv.Key; break; } }
            sb.AppendLine("Transition: " + (transitionName ?? (slide.SlideShowTransition.EntryEffect == PowerPoint.PpEntryEffect.ppEffectNone ? "none" : slide.SlideShowTransition.EntryEffect.ToString())));

            int animCount = slide.TimeLine.MainSequence.Count;
            if (animCount > 0) sb.AppendLine(animCount + " animation(s) - call read_animations to see them.");

            // Post-hoc addition (2026-08-24, user-requested: "see the order
            // between objects"): slide.Shapes is already ordered back-to-
            // front by z-order (confirmed via reflection: Shape.ZOrderPosition
            // is a get-only int matching this same collection order) - the
            // shapeIndex below was already exactly this order, just never
            // stated explicitly. No new read tool needed; making the
            // existing order's meaning explicit is enough.
            if (slide.Shapes.Count > 1) sb.AppendLine("Shapes below are listed back-to-front (z-order) - index 0 is furthest back, the last index is drawn on top. Use set_element_order to change this.");

            int shapeIndex = 0;
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                sb.AppendLine($"[{shapeIndex}] {shape.Name}: {ShapeText(shape)}");
                shapeIndex++;
            }

            string notesText = GetSlideNotesText(slide);
            if (!string.IsNullOrWhiteSpace(notesText)) sb.AppendLine("Notes: " + notesText.Replace("\r", " ").Trim());

            return new ToolResult { Output = sb.ToString(), Summary = "read_slide" };
        }

        // Read-only search across every slide's shape text (via the existing
        // ShapeText helper, so text boxes/placeholders/tables/SmartArt are all
        // covered the same as get_deck_context/read_slide) plus speaker notes.
        // There was previously no way to locate text in a deck without
        // reading every slide one at a time via read_slide and scanning
        // yourself.
        private static ToolResult FindTextPpt(JsonElement input)
        {
            string query = input.GetProperty("query").GetString();
            bool useRegex = input.TryGetProperty("regex", out var rx) && rx.ValueKind == JsonValueKind.True;
            bool matchCase = input.TryGetProperty("matchCase", out var mc) && mc.ValueKind == JsonValueKind.True;
            int maxResults = input.GetProperty("max_results").GetInt32();

            System.Text.RegularExpressions.Regex regex = useRegex
                ? new System.Text.RegularExpressions.Regex(query, matchCase
                    ? System.Text.RegularExpressions.RegexOptions.None
                    : System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                : null;
            StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            bool IsMatch(string text) => regex != null ? regex.IsMatch(text) : text.IndexOf(query, comparison) >= 0;

            var sb = new StringBuilder();
            int found = 0;
            int slideIndex = 0;
            foreach (PowerPoint.Slide slide in ActivePresentation.Slides)
            {
                if (found >= maxResults) break;
                int shapeIndex = 0;
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (found >= maxResults) break;
                    string text = ShapeText(shape).Replace("\r", " ").Trim();
                    if (text.Length > 0 && IsMatch(text))
                    {
                        sb.AppendLine($"[slide {slideIndex}, shape {shapeIndex}] {text}");
                        found++;
                    }
                    shapeIndex++;
                }
                if (found < maxResults)
                {
                    string notes = GetSlideNotesText(slide).Replace("\r", " ").Trim();
                    if (notes.Length > 0 && IsMatch(notes))
                    {
                        sb.AppendLine($"[slide {slideIndex}, notes] {notes}");
                        found++;
                    }
                }
                slideIndex++;
            }
            return new ToolResult { Output = found > 0 ? sb.ToString() : "No matches.", Summary = "find_text" };
        }

        // Scoped to simple text-frame shapes (title/body placeholders, text
        // boxes) and speaker notes - NOT table cells or SmartArt node text,
        // which have their own dedicated edit tools (edit_table_cell,
        // edit_smartart) and aren't safely addressable by a single
        // TextRange.Text assignment the way a plain text frame is.
        private static ToolResult ReplaceTextPpt(JsonElement input)
        {
            string find = input.GetProperty("find").GetString();
            string replace = input.GetProperty("replace").GetString();
            bool useRegex = input.TryGetProperty("regex", out var rx) && rx.ValueKind == JsonValueKind.True;
            bool matchCase = input.TryGetProperty("matchCase", out var mc) && mc.ValueKind == JsonValueKind.True;
            bool includeNotes = !input.TryGetProperty("includeNotes", out var inEl) || inEl.ValueKind != JsonValueKind.False;

            System.Text.RegularExpressions.Regex regex = useRegex
                ? new System.Text.RegularExpressions.Regex(find, matchCase
                    ? System.Text.RegularExpressions.RegexOptions.None
                    : System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                : null;
            StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            string ApplyReplace(string text, out int count)
            {
                if (regex != null)
                {
                    count = regex.Matches(text).Count;
                    return count > 0 ? regex.Replace(text, replace) : text;
                }
                count = TextUtil.CountOccurrences(text, find, comparison);
                return count > 0 ? TextUtil.ReplaceAllOccurrences(text, find, replace, comparison) : text;
            }

            int replaced = 0;
            foreach (PowerPoint.Slide slide in ActivePresentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) continue;
                    if (shape.TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoTrue) continue;
                    PowerPoint.TextRange range = shape.TextFrame.TextRange;
                    string newText = ApplyReplace(range.Text, out int count);
                    if (count == 0) continue;
                    range.Text = newText;
                    ApplyAutoDirection(range, newText);
                    replaced += count;
                }
                if (includeNotes)
                {
                    PowerPoint.Shape notesBody = ResolveNotesBodyPlaceholder(slide);
                    if (notesBody != null && notesBody.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue &&
                        notesBody.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        PowerPoint.TextRange notesRange = notesBody.TextFrame.TextRange;
                        string newNotes = ApplyReplace(notesRange.Text, out int notesCount);
                        if (notesCount > 0)
                        {
                            notesRange.Text = newNotes;
                            replaced += notesCount;
                        }
                    }
                }
            }
            return new ToolResult { Output = replaced + " replacement(s).", Mutated = replaced > 0, Summary = "replace_text" };
        }

    }
}

