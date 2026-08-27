using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static class PowerPointTools
    {
        // Editing-mode gating (mirrors the Word/Excel Task 11/16 pattern this plan establishes
        // elsewhere): the tool list offered to the model is filtered client-side per mode
        // (web-src/entry.ts, first line of defense - smaller prompts, fewer wasted turns), but
        // Execute() independently re-checks mode here as defense-in-depth, since nothing stops a
        // misbehaving or malicious model response from calling a tool that wasn't offered.
        //
        // Per-document since PP-1 - see WordTools.cs's identical pattern for the rationale.
        private static readonly Dictionary<string, EditingMode> ModeByDoc = new Dictionary<string, EditingMode>();

        public static void SetMode(string docKey, EditingMode mode)
        {
            ModeByDoc[docKey] = mode;
        }

        private static EditingMode ModeFor(string docKey)
        {
            EditingMode m;
            return ModeByDoc.TryGetValue(docKey, out m) ? m : EditingMode.FullAutonomy;
        }

        // Tools always allowed regardless of editing mode (read-only, no document mutation).
        private static readonly System.Collections.Generic.HashSet<string> AlwaysAllowedTools =
            new System.Collections.Generic.HashSet<string> { "get_deck_context", "read_slide", "read_animations", "find_text" };

        public static ToolResult Execute(string docKey, string name, JsonElement input)
        {
            try
            {
                EditingMode mode = ModeFor(docKey);
                if (!AlwaysAllowedTools.Contains(name) && !IsMutationAllowed(mode))
                {
                    return new ToolResult
                    {
                        Output = "Blocked: editing mode is " + ModeLabel(mode) + ".",
                        IsError = true,
                        Summary = name,
                    };
                }

                switch (name)
                {
                    case "get_deck_context": return GetDeckContext();
                    case "read_slide": return ReadSlide(input);
                    case "find_text": return FindTextPpt(input);
                    case "replace_text": return ReplaceTextPpt(input);
                    case "set_element_text": return SetElementText(input);
                    case "set_slide_notes": return SetSlideNotes(input);
                    case "set_element_style": return SetElementStyle(input);
                    case "set_element_transform": return SetElementTransform(input);
                    case "set_element_order": return SetElementOrder(input);
                    case "add_text_box": return AddTextBox(input);
                    case "add_shape": return AddShape(input);
                    case "delete_element": return DeleteElement(input);
                    case "add_slide": return AddSlide(input);
                    case "set_element_fill": return SetElementFill(input);
                    case "set_element_stroke": return SetElementStroke(input);
                    case "set_slide_background": return SetSlideBackground(input);
                    case "ungroup_element": return UngroupElement(input);
                    case "add_table": return AddTable(input);
                    case "edit_table_cell": return EditTableCell(input);
                    case "edit_table_structure": return EditTableStructure(input);
                    case "edit_table_style": return EditTableStyle(input);
                    case "add_chart": return AddChartPpt(input);
                    case "edit_chart": return EditChartPpt(input);
                    case "add_smartart": return AddSmartArt(input);
                    case "crop_image": return CropImage(input);
                    case "replace_image": return ReplaceImagePpt(input);
                    case "set_picture_opacity": return SetPictureOpacity(input);
                    case "delete_slide": return DeleteSlide(input);
                    case "move_slide": return MoveSlide(input);
                    case "duplicate_slide": return DuplicateSlide(input);
                    case "set_slide_layout": return SetSlideLayout(input);
                    case "set_slide_transition": return SetSlideTransition(input);
                    case "add_animation": return AddAnimation(input);
                    case "read_animations": return ReadAnimations(input);
                    case "edit_animation": return EditAnimation(input);
                    default: return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        // Read Only and Comment Only modes block all mutating tools (PowerPoint has no
        // comment-equivalent tool in this pass - see plan backlog - so Comment Only currently
        // behaves identically to Read Only: no mutating tools available). Track Changes is scoped
        // to simple allow/block gating for now (same as Excel's Task 16 scoping note) rather than a
        // native PowerPoint revision-tracking UI. Full Autonomy allows everything.
        private static bool IsMutationAllowed(EditingMode mode)
        {
            return mode == EditingMode.TrackChanges || mode == EditingMode.FullAutonomy;
        }

        private static string ModeLabel(EditingMode mode)
        {
            switch (mode)
            {
                case EditingMode.ReadOnly: return "Read Only";
                case EditingMode.CommentOnly: return "Comment Only";
                case EditingMode.TrackChanges: return "Track Changes";
                default: return "Full Autonomy";
            }
        }

        // Known limitation (PP-1 Task 5 Step 5): resolves whichever presentation
        // is ACTIVE right now, not necessarily the one whose pane initiated
        // this tool call - see WordTools.cs's ActiveDoc for the identical
        // rationale and the same out-of-scope decision.
        private static PowerPoint.Presentation ActivePresentation => Globals.ThisAddIn.Application.ActivePresentation;

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

        private static PowerPoint.Shape ResolveShape(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int shapeIndex = input.GetProperty("shapeIndex").GetInt32();
            return ActivePresentation.Slides[slideIndex + 1].Shapes[shapeIndex + 1];
        }

        // The notes body is a placeholder on the slide's NotesPage (a separate
        // page object from the slide itself), found by placeholder TYPE rather
        // than a hardcoded index - the notes master can be customized, so
        // "index 2" isn't guaranteed to be the body across every deck.
        private static PowerPoint.Shape ResolveNotesBodyPlaceholder(PowerPoint.Slide slide)
        {
            foreach (PowerPoint.Shape shape in slide.NotesPage.Shapes.Placeholders)
            {
                if (shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody) return shape;
            }
            return null;
        }

        private static string GetSlideNotesText(PowerPoint.Slide slide)
        {
            PowerPoint.Shape body = ResolveNotesBodyPlaceholder(slide);
            if (body == null) return "";
            if (body.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return "";
            if (body.TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoTrue) return "";
            return body.TextFrame.TextRange.Text;
        }

        // Post-hoc fix (2026-08-26, user-reported): PowerPoint's TextRange
        // never auto-flips paragraph direction/alignment based on typed
        // content the way Word's editor does with "detect language
        // automatically" - every write here always came out left-to-right/
        // left-aligned, even for Hebrew. Same bidi mismatch class as
        // chat-ui.ts's dir="auto" fix, but PowerPoint's COM object model has
        // no built-in "auto" direction - it has to be decided per write from
        // the text's own script mix. IsRtlMajority itself now lives in
        // TextUtil (Phase 0) since it's free of COM types.
        private static void ApplyAutoDirection(PowerPoint.TextRange range, string text)
        {
            if (!TextUtil.IsRtlMajority(text)) return;
            range.ParagraphFormat.TextDirection = PowerPoint.PpDirection.ppDirectionRightToLeft;
            range.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignRight;
        }

        // User-reported bug: a placeholder from a layout like "Title and
        // Content" already renders its own native bullet per paragraph, so a
        // model that ALSO types a literal bullet character ("•"/"-"/"*") at
        // the start of each line ends up with two bullets per line. This
        // gives the model a real on/off switch instead, so it never needs to
        // embed a literal bullet character in the text. bulleted omitted
        // (the JSON property absent, not merely false) leaves the shape's
        // existing bullet setting untouched - matches prior behavior for
        // every existing caller that doesn't pass it.
        private static void ApplyBulletSetting(PowerPoint.TextRange range, JsonElement input)
        {
            if (!input.TryGetProperty("bulleted", out var el)) return;
            if (el.ValueKind != JsonValueKind.True && el.ValueKind != JsonValueKind.False) return;
            bool bulleted = el.ValueKind == JsonValueKind.True;
            range.ParagraphFormat.Bullet.Visible = bulleted
                ? Microsoft.Office.Core.MsoTriState.msoTrue
                : Microsoft.Office.Core.MsoTriState.msoFalse;
            if (bulleted) range.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
        }

        private static ToolResult SetElementText(JsonElement input)
        {
            string text = input.GetProperty("text").GetString();
            PowerPoint.TextRange range = ResolveShape(input).TextFrame.TextRange;
            range.Text = text;
            ApplyAutoDirection(range, text);
            ApplyBulletSetting(range, input);
            return new ToolResult { Output = "Text updated.", Mutated = true, Summary = "set_element_text" };
        }

        private static ToolResult SetSlideNotes(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            string text = input.GetProperty("text").GetString();
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            PowerPoint.Shape body = ResolveNotesBodyPlaceholder(slide);
            if (body == null)
            {
                return new ToolResult { Output = "Slide has no notes body placeholder.", IsError = true, Summary = "set_slide_notes" };
            }
            PowerPoint.TextRange range = body.TextFrame.TextRange;
            range.Text = text;
            ApplyAutoDirection(range, text);
            return new ToolResult { Output = "Notes updated.", Mutated = true, Summary = "set_slide_notes" };
        }

        // PP-20: left|center|right|justify -> PpParagraphAlignment, mirroring
        // this file's PptChartTypeMap/SmartArtLayoutNames dictionary pattern.
        private static readonly Dictionary<string, PowerPoint.PpParagraphAlignment> AlignmentMap =
            new Dictionary<string, PowerPoint.PpParagraphAlignment>
        {
            ["left"] = PowerPoint.PpParagraphAlignment.ppAlignLeft,
            ["center"] = PowerPoint.PpParagraphAlignment.ppAlignCenter,
            ["right"] = PowerPoint.PpParagraphAlignment.ppAlignRight,
            ["justify"] = PowerPoint.PpParagraphAlignment.ppAlignJustify,
        };

        private static ToolResult SetElementStyle(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            PowerPoint.TextRange range = shape.TextFrame.TextRange;
            var applied = new List<string>();

            if (input.TryGetProperty("bold", out var bold))
            {
                range.Font.Bold = bold.GetBoolean() ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
                applied.Add("bold");
            }
            if (input.TryGetProperty("italic", out var italic))
            {
                range.Font.Italic = italic.GetBoolean() ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
                applied.Add("italic");
            }
            if (input.TryGetProperty("fontSize", out var fontSize))
            {
                range.Font.Size = (float)fontSize.GetDouble();
                applied.Add("fontSize");
            }
            if (input.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
            {
                range.Font.Color.RGB = ColorUtil.HexToOle(color.GetString());
                applied.Add("color");
            }
            if (input.TryGetProperty("fontName", out var fontName) && fontName.ValueKind == JsonValueKind.String)
            {
                range.Font.Name = fontName.GetString();
                applied.Add("fontName");
            }
            if (input.TryGetProperty("underline", out var underline))
            {
                range.Font.Underline = underline.ValueKind == JsonValueKind.True ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
                applied.Add("underline");
            }
            if (input.TryGetProperty("shadow", out var shadow))
            {
                range.Font.Shadow = shadow.ValueKind == JsonValueKind.True ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
                applied.Add("shadow");
            }
            if (input.TryGetProperty("alignment", out var align) && align.ValueKind == JsonValueKind.String)
            {
                PowerPoint.PpParagraphAlignment a;
                if (!AlignmentMap.TryGetValue(align.GetString(), out a))
                    throw new ArgumentException("set_element_style: unknown alignment '" + align.GetString() +
                                                "'. Valid: " + string.Join(", ", AlignmentMap.Keys) + ".");
                range.ParagraphFormat.Alignment = a;
                applied.Add("alignment");
            }
            if (input.TryGetProperty("baselineOffset", out var baseline) && baseline.ValueKind == JsonValueKind.String)
            {
                string b = baseline.GetString();
                if (b != "SUPERSCRIPT" && b != "SUBSCRIPT" && b != "NONE")
                    throw new ArgumentException("set_element_style: unknown baselineOffset '" + b +
                                                "'. Valid: SUPERSCRIPT, SUBSCRIPT, NONE.");
                range.Font.Superscript = b == "SUPERSCRIPT" ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
                range.Font.Subscript = b == "SUBSCRIPT" ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
                applied.Add("baselineOffset");
            }
            // PP-20 Task 1 Step 2: strikethrough deliberately NOT implemented.
            // TextFrame (the older text model this file uses everywhere else)
            // has no Strikethrough member on this PIA. Excel's interop PIA
            // exposes TextFrame2/TextRange2 (the newer "DrawingML" text model,
            // used by ExcelTools.cs's chart formatting) with a Strikethrough
            // property, but Microsoft.Office.Interop.PowerPoint has no
            // TextFrame2 type or Shape.TextFrame2 member at all (confirmed:
            // absent from this PIA's own XML docs, and CS0234 - "TextRange2
            // does not exist in the namespace" - on a direct attempt). Per
            // this plan's own instruction: omit rather than ship a schema
            // field the handler can't back, or a `dynamic` call this
            // environment cannot runtime-verify.

            return new ToolResult
            {
                Output = applied.Count > 0
                    ? "Style updated: " + string.Join(", ", applied) + "."
                    : "No recognized style properties were provided - nothing changed.",
                Mutated = applied.Count > 0,
                Summary = "set_element_style",
            };
        }

        private static ToolResult SetElementTransform(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            if (input.TryGetProperty("left", out var left)) shape.Left = (float)left.GetDouble();
            if (input.TryGetProperty("top", out var top)) shape.Top = (float)top.GetDouble();
            if (input.TryGetProperty("width", out var width)) shape.Width = (float)width.GetDouble();
            if (input.TryGetProperty("height", out var height)) shape.Height = (float)height.GetDouble();
            if (input.TryGetProperty("rotation", out var rotation)) shape.Rotation = (float)rotation.GetDouble();
            return new ToolResult { Output = "Transform updated.", Mutated = true, Summary = "set_element_transform" };
        }

        // Post-hoc addition (2026-08-24, user-requested: "change the order
        // of an element" - stacking/z-order, distinct from set_element_
        // transform's position/size). Confirmed via reflection: Shape.ZOrder
        // (MsoZOrderCmd) is the real relative-move method; MsoZOrderCmd has
        // 6 values total, of which the first 4 apply to slide shapes
        // (msoBringInFrontOfText/msoSendBehindText are for a shape's
        // position relative to body text, not meaningful for a slide's flat
        // z-order stack) - only those 4 are exposed.
        private static readonly Dictionary<string, Microsoft.Office.Core.MsoZOrderCmd> ZOrderMap = new Dictionary<string, Microsoft.Office.Core.MsoZOrderCmd>
        {
            ["bringToFront"] = Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront,
            ["sendToBack"] = Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack,
            ["bringForward"] = Microsoft.Office.Core.MsoZOrderCmd.msoBringForward,
            ["sendBackward"] = Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward,
        };

        private static ToolResult SetElementOrder(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            string kind = input.GetProperty("kind").GetString();
            Microsoft.Office.Core.MsoZOrderCmd cmd;
            if (!ZOrderMap.TryGetValue(kind, out cmd))
                throw new ArgumentException("set_element_order: unknown kind '" + kind + "'. Valid: " + string.Join(", ", ZOrderMap.Keys) + ".");
            shape.ZOrder(cmd);
            // ZOrderPosition is 1-based in COM; reported 0-based to match
            // read_slide's shapeIndex convention. Structural edit - every
            // other shape's shapeIndex on this slide may have shifted too,
            // same caveat as delete_element/ungroup_element elsewhere in
            // this file.
            int newShapeIndex = shape.ZOrderPosition - 1;
            return new ToolResult { Output = "Shape order changed (" + kind + ") - now at shapeIndex " + newShapeIndex + ". Other shapes on this slide may have shifted index - re-read the slide (read_slide) before addressing another shape by index in the same run.", Mutated = true, Summary = "set_element_order" };
        }

        private static ToolResult AddTextBox(JsonElement input)
        {
            PowerPoint.Slide slide = ActivePresentation.Slides[input.GetProperty("slideIndex").GetInt32() + 1];
            float left = (float)input.GetProperty("left").GetDouble();
            float top = (float)input.GetProperty("top").GetDouble();
            float width = (float)input.GetProperty("width").GetDouble();
            float height = (float)input.GetProperty("height").GetDouble();
            PowerPoint.Shape shape = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            string text = input.GetProperty("text").GetString();
            PowerPoint.TextRange range = shape.TextFrame.TextRange;
            range.Text = text;
            ApplyAutoDirection(range, text);
            ApplyBulletSetting(range, input);
            return new ToolResult { Output = "Text box added.", Mutated = true, Summary = "add_text_box" };
        }

        // PP-20 Task 2: ported from ExcelAiAddIn/ExcelTools.cs's ShapeTypeMap -
        // keep the two in sync. Same PIA-omission note applies: msoShapePlus/
        // msoShapeMathPlus do not exist in this project's referenced
        // Microsoft.Office.Core PIA (confirmed via CS0117 on the Excel side) -
        // omitted here too, not re-attempted.
        // rect/ellipse are Excel's (canonical) spellings; rectangle/oval are
        // PowerPoint's pre-existing spellings, kept as aliases so every
        // existing call/prompt keeps working - both map to the same
        // MsoAutoShapeType. OrdinalIgnoreCase so a near-miss case still
        // resolves instead of erroring unnecessarily.
        private static readonly Dictionary<string, Microsoft.Office.Core.MsoAutoShapeType> ShapeTypeMap =
            new Dictionary<string, Microsoft.Office.Core.MsoAutoShapeType>(StringComparer.OrdinalIgnoreCase)
        {
            ["rect"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
            ["rectangle"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
            ["roundRect"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle,
            ["ellipse"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
            ["oval"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
            ["triangle"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle,
            ["rtTriangle"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightTriangle,
            ["parallelogram"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeParallelogram,
            ["trapezoid"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeTrapezoid,
            ["diamond"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond,
            ["pentagon"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapePentagon,
            ["hexagon"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeHexagon,
            ["octagon"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeOctagon,
            ["pie"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapePie,
            ["chord"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeChord,
            ["donut"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeDonut,
            ["foldedCorner"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeFoldedCorner,
            ["heart"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeHeart,
            ["lightningBolt"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeLightningBolt,
            ["sun"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeSun,
            ["moon"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeMoon,
            ["cloud"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeCloud,
            ["arc"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeArc,
            ["star5"] = Microsoft.Office.Core.MsoAutoShapeType.msoShape5pointStar,
            ["rightArrow"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightArrow,
            ["leftArrow"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeLeftArrow,
            ["upArrow"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeUpArrow,
            ["downArrow"] = Microsoft.Office.Core.MsoAutoShapeType.msoShapeDownArrow,
        };

        private static ToolResult AddShape(JsonElement input)
        {
            PowerPoint.Slide slide = ActivePresentation.Slides[input.GetProperty("slideIndex").GetInt32() + 1];
            string shapeType = input.GetProperty("shapeType").GetString();
            Microsoft.Office.Core.MsoAutoShapeType autoShapeType;
            if (!ShapeTypeMap.TryGetValue(shapeType, out autoShapeType))
                throw new ArgumentException("add_shape: unknown shapeType '" + shapeType + "'. Valid: " +
                                            string.Join(", ", ShapeTypeMap.Keys) + ".");
            float left = (float)input.GetProperty("left").GetDouble();
            float top = (float)input.GetProperty("top").GetDouble();
            float width = (float)input.GetProperty("width").GetDouble();
            float height = (float)input.GetProperty("height").GetDouble();
            PowerPoint.Shape shape = slide.Shapes.AddShape(autoShapeType, left, top, width, height);
            if (input.TryGetProperty("text", out var text)) shape.TextFrame.TextRange.Text = text.GetString();
            return new ToolResult { Output = "Shape added.", Mutated = true, Summary = "add_shape" };
        }

        private static ToolResult DeleteElement(JsonElement input)
        {
            ResolveShape(input).Delete();
            return new ToolResult { Output = "Shape deleted.", Mutated = true, Summary = "delete_element" };
        }

        private static ToolResult AddSlide(JsonElement input)
        {
            int sourceIndex = input.GetProperty("sourceIndex").GetInt32();
            bool clearText = !input.TryGetProperty("clearText", out var ct) || ct.ValueKind != JsonValueKind.False;
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (sourceIndex < 0 || sourceIndex >= slides.Count)
            {
                return new ToolResult { Output = "Invalid sourceIndex.", IsError = true, Summary = "add_slide" };
            }
            dynamic source = slides[sourceIndex + 1];
            dynamic dupRange = source.Duplicate(); // returns a SlideRange containing exactly the new slide
            dynamic newSlide = dupRange[1];
            newSlide.MoveTo(sourceIndex + 2);
            if (clearText)
            {
                foreach (PowerPoint.Shape shape in newSlide.Shapes)
                {
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        shape.TextFrame.TextRange.Text = "";
                    }
                }
            }
            return new ToolResult { Output = "Slide added after index " + sourceIndex + ".", Mutated = true, Summary = "add_slide" };
        }

        // PP-19 Task 1: delete_slide/move_slide/duplicate_slide. Deliberately no
        // slideIndexes:number[] batch form - deleting slide 2 shifts every later
        // slide's index down by one, so a batch would need to either resolve all
        // targets up front or delete in strict descending order. One slide per
        // call (with the index-shift warning in the output/description) is the
        // safer answer; it makes the model re-read the deck between deletes
        // instead of silently deleting the wrong slides.
        private static ToolResult DeleteSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Presentation pres = ActivePresentation;
            if (slideIndex < 0 || slideIndex >= pres.Slides.Count)
                throw new ArgumentOutOfRangeException("slideIndex",
                    "slideIndex must be between 0 and " + (pres.Slides.Count - 1) + ".");
            if (pres.Slides.Count == 1)
                throw new InvalidOperationException(
                    "delete_slide: cannot delete the only slide in the presentation.");
            pres.Slides[slideIndex + 1].Delete();
            return new ToolResult
            {
                Output = "Deleted slide " + slideIndex + ". " + pres.Slides.Count + " slide(s) remain; " +
                         "slides after it have shifted down by one index.",
                Mutated = true,
                Summary = "delete_slide",
            };
        }

        private static ToolResult MoveSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int toIndex = input.GetProperty("toIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (slideIndex < 0 || slideIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("slideIndex",
                    "slideIndex must be between 0 and " + (slides.Count - 1) + ".");
            if (toIndex < 0 || toIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("toIndex",
                    "toIndex must be between 0 and " + (slides.Count - 1) + ".");
            slides[slideIndex + 1].MoveTo(toIndex + 1);
            return new ToolResult
            {
                Output = "Moved slide " + slideIndex + " to position " + toIndex + ".",
                Mutated = true,
                Summary = "move_slide",
            };
        }

        private static ToolResult DuplicateSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (slideIndex < 0 || slideIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("slideIndex",
                    "slideIndex must be between 0 and " + (slides.Count - 1) + ".");
            slides[slideIndex + 1].Duplicate();
            return new ToolResult
            {
                Output = "Duplicated slide " + slideIndex + " (content included) - the copy is now at index " + (slideIndex + 1) + ".",
                Mutated = true,
                Summary = "duplicate_slide",
            };
        }

        // PP-24: curated subset of the 37 PpSlideLayout values (confirmed via
        // reflection against the real referenced PIA, not recalled) - the
        // pre-2007 leftovers (ppLayoutOrgchart, ppLayoutMediaClipAndText,
        // etc.) are omitted as unlikely to be what a model means by a
        // layout request. Same curated-map-with-throw-on-unknown pattern as
        // PptChartTypeMap/AlignmentMap elsewhere in this file.
        private static readonly Dictionary<string, PowerPoint.PpSlideLayout> SlideLayoutMap = new Dictionary<string, PowerPoint.PpSlideLayout>
        {
            ["title"] = PowerPoint.PpSlideLayout.ppLayoutTitle,
            ["titleOnly"] = PowerPoint.PpSlideLayout.ppLayoutTitleOnly,
            ["blank"] = PowerPoint.PpSlideLayout.ppLayoutBlank,
            ["text"] = PowerPoint.PpSlideLayout.ppLayoutText,
            ["twoColumnText"] = PowerPoint.PpSlideLayout.ppLayoutTwoColumnText,
            ["object"] = PowerPoint.PpSlideLayout.ppLayoutObject,
            ["objectAndText"] = PowerPoint.PpSlideLayout.ppLayoutObjectAndText,
            ["textAndObject"] = PowerPoint.PpSlideLayout.ppLayoutTextAndObject,
            ["twoObjects"] = PowerPoint.PpSlideLayout.ppLayoutTwoObjects,
            ["twoObjectsAndText"] = PowerPoint.PpSlideLayout.ppLayoutTwoObjectsAndText,
            ["fourObjects"] = PowerPoint.PpSlideLayout.ppLayoutFourObjects,
            ["table"] = PowerPoint.PpSlideLayout.ppLayoutTable,
            ["chart"] = PowerPoint.PpSlideLayout.ppLayoutChart,
            ["sectionHeader"] = PowerPoint.PpSlideLayout.ppLayoutSectionHeader,
            ["comparison"] = PowerPoint.PpSlideLayout.ppLayoutComparison,
            ["contentWithCaption"] = PowerPoint.PpSlideLayout.ppLayoutContentWithCaption,
            ["pictureWithCaption"] = PowerPoint.PpSlideLayout.ppLayoutPictureWithCaption,
        };

        // Mirrors WordTools.cs's ResolveSmartArtGalleryItem (PP-23): resolves
        // by case-insensitive substring match against this deck's own live
        // theme layouts, since custom layout names are not a fixed enum -
        // a miss lists the real available names so the caller can retry
        // correctly instead of guessing blind.
        private static PowerPoint.CustomLayout ResolveCustomLayout(PowerPoint.Slide slide, string query)
        {
            PowerPoint.CustomLayout firstMatch = null;
            var namesSeen = new List<string>();
            foreach (PowerPoint.CustomLayout layout in slide.Design.SlideMaster.CustomLayouts)
            {
                namesSeen.Add(layout.Name);
                if (firstMatch == null && layout.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) firstMatch = layout;
            }
            if (firstMatch != null) return firstMatch;
            throw new ArgumentException("set_slide_layout: no custom layout matching '" + query + "' found in this slide's theme. Available: " + string.Join(", ", namesSeen) + ".");
        }

        private static ToolResult SetSlideLayout(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            string kind = input.TryGetProperty("kind", out var kindEl) && kindEl.ValueKind == JsonValueKind.String ? kindEl.GetString() : "classic";

            if (kind == "custom")
            {
                string layoutName = input.GetProperty("layoutName").GetString();
                slide.CustomLayout = ResolveCustomLayout(slide, layoutName);
                return new ToolResult { Output = "Slide " + slideIndex + " layout set to custom layout '" + layoutName + "'.", Mutated = true, Summary = "set_slide_layout" };
            }
            if (kind != "classic")
                throw new ArgumentException("set_slide_layout: unknown kind '" + kind + "'. Valid: classic, custom.");

            string layoutKey = input.GetProperty("layout").GetString();
            PowerPoint.PpSlideLayout layoutValue;
            if (!SlideLayoutMap.TryGetValue(layoutKey, out layoutValue))
                throw new ArgumentException("set_slide_layout: unknown layout '" + layoutKey + "'. Valid: " + string.Join(", ", SlideLayoutMap.Keys) + ".");
            slide.Layout = layoutValue;
            return new ToolResult { Output = "Slide " + slideIndex + " layout set to '" + layoutKey + "'.", Mutated = true, Summary = "set_slide_layout" };
        }

        // PP-24: curated subset of PpEntryEffect's 189 values, all confirmed
        // present via reflection against the real referenced PIA.
        private static readonly Dictionary<string, PowerPoint.PpEntryEffect> TransitionEffectMap = new Dictionary<string, PowerPoint.PpEntryEffect>
        {
            ["none"] = PowerPoint.PpEntryEffect.ppEffectNone,
            ["cut"] = PowerPoint.PpEntryEffect.ppEffectCut,
            ["fade"] = PowerPoint.PpEntryEffect.ppEffectFade,
            ["dissolve"] = PowerPoint.PpEntryEffect.ppEffectDissolve,
            ["random"] = PowerPoint.PpEntryEffect.ppEffectRandom,
            ["wipeLeft"] = PowerPoint.PpEntryEffect.ppEffectWipeLeft,
            ["wipeRight"] = PowerPoint.PpEntryEffect.ppEffectWipeRight,
            ["wipeUp"] = PowerPoint.PpEntryEffect.ppEffectWipeUp,
            ["wipeDown"] = PowerPoint.PpEntryEffect.ppEffectWipeDown,
            ["pushLeft"] = PowerPoint.PpEntryEffect.ppEffectPushLeft,
            ["pushRight"] = PowerPoint.PpEntryEffect.ppEffectPushRight,
            ["pushUp"] = PowerPoint.PpEntryEffect.ppEffectPushUp,
            ["pushDown"] = PowerPoint.PpEntryEffect.ppEffectPushDown,
            ["coverLeft"] = PowerPoint.PpEntryEffect.ppEffectCoverLeft,
            ["coverRight"] = PowerPoint.PpEntryEffect.ppEffectCoverRight,
            ["coverUp"] = PowerPoint.PpEntryEffect.ppEffectCoverUp,
            ["coverDown"] = PowerPoint.PpEntryEffect.ppEffectCoverDown,
            ["uncoverLeft"] = PowerPoint.PpEntryEffect.ppEffectUncoverLeft,
            ["uncoverRight"] = PowerPoint.PpEntryEffect.ppEffectUncoverRight,
            ["uncoverUp"] = PowerPoint.PpEntryEffect.ppEffectUncoverUp,
            ["uncoverDown"] = PowerPoint.PpEntryEffect.ppEffectUncoverDown,
            ["zoomIn"] = PowerPoint.PpEntryEffect.ppEffectZoomIn,
            ["zoomOut"] = PowerPoint.PpEntryEffect.ppEffectZoomOut,
            ["zoomCenter"] = PowerPoint.PpEntryEffect.ppEffectZoomCenter,
            ["circle"] = PowerPoint.PpEntryEffect.ppEffectCircleOut,
            ["diamond"] = PowerPoint.PpEntryEffect.ppEffectDiamondOut,
            ["splitHorizontal"] = PowerPoint.PpEntryEffect.ppEffectSplitHorizontalOut,
            ["splitVertical"] = PowerPoint.PpEntryEffect.ppEffectSplitVerticalOut,
            ["wheel"] = PowerPoint.PpEntryEffect.ppEffectWheel1Spoke,
            ["blindsHorizontal"] = PowerPoint.PpEntryEffect.ppEffectBlindsHorizontal,
            ["blindsVertical"] = PowerPoint.PpEntryEffect.ppEffectBlindsVertical,
            ["checkerboard"] = PowerPoint.PpEntryEffect.ppEffectCheckerboardAcross,
        };

        private static ToolResult SetSlideTransition(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            PowerPoint.SlideShowTransition transition = slide.SlideShowTransition;

            string effectKey = input.GetProperty("effect").GetString();
            PowerPoint.PpEntryEffect effectValue;
            if (!TransitionEffectMap.TryGetValue(effectKey, out effectValue))
                throw new ArgumentException("set_slide_transition: unknown effect '" + effectKey + "'. Valid: " + string.Join(", ", TransitionEffectMap.Keys) + ".");
            transition.EntryEffect = effectValue;

            if (input.TryGetProperty("durationSeconds", out var durEl) && durEl.ValueKind == JsonValueKind.Number)
                transition.Duration = (float)durEl.GetDouble();
            if (input.TryGetProperty("advanceOnClick", out var clickEl))
                transition.AdvanceOnClick = clickEl.ValueKind == JsonValueKind.True ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            if (input.TryGetProperty("advanceAfterSeconds", out var advEl) && advEl.ValueKind == JsonValueKind.Number)
            {
                transition.AdvanceOnTime = Microsoft.Office.Core.MsoTriState.msoTrue;
                transition.AdvanceTime = (float)advEl.GetDouble();
            }

            return new ToolResult { Output = "Slide " + slideIndex + " transition set to '" + effectKey + "'.", Mutated = true, Summary = "set_slide_transition" };
        }

        // PP-24: curated subset of MsoAnimEffect's 151 values, all confirmed
        // present via reflection against the real referenced PIA.
        private static readonly Dictionary<string, PowerPoint.MsoAnimEffect> AnimationEffectMap = new Dictionary<string, PowerPoint.MsoAnimEffect>
        {
            ["appear"] = PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
            ["fade"] = PowerPoint.MsoAnimEffect.msoAnimEffectFade,
            ["fly"] = PowerPoint.MsoAnimEffect.msoAnimEffectFly,
            ["flashOnce"] = PowerPoint.MsoAnimEffect.msoAnimEffectFlashOnce,
            ["wipe"] = PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
            ["zoom"] = PowerPoint.MsoAnimEffect.msoAnimEffectZoom,
            ["dissolve"] = PowerPoint.MsoAnimEffect.msoAnimEffectDissolve,
            ["bounce"] = PowerPoint.MsoAnimEffect.msoAnimEffectBounce,
            ["spiral"] = PowerPoint.MsoAnimEffect.msoAnimEffectSpiral,
            ["swivel"] = PowerPoint.MsoAnimEffect.msoAnimEffectSwivel,
            ["wheel"] = PowerPoint.MsoAnimEffect.msoAnimEffectWheel,
            ["split"] = PowerPoint.MsoAnimEffect.msoAnimEffectSplit,
            ["box"] = PowerPoint.MsoAnimEffect.msoAnimEffectBox,
            ["circle"] = PowerPoint.MsoAnimEffect.msoAnimEffectCircle,
            ["diamond"] = PowerPoint.MsoAnimEffect.msoAnimEffectDiamond,
            ["plus"] = PowerPoint.MsoAnimEffect.msoAnimEffectPlus,
            ["checkerboard"] = PowerPoint.MsoAnimEffect.msoAnimEffectCheckerboard,
            ["randomBars"] = PowerPoint.MsoAnimEffect.msoAnimEffectRandomBars,
            ["growAndTurn"] = PowerPoint.MsoAnimEffect.msoAnimEffectGrowAndTurn,
            ["riseUp"] = PowerPoint.MsoAnimEffect.msoAnimEffectRiseUp,
        };

        // All 7 real MsoAnimTriggerType values are not exposed - msoAnimTriggerNone
        // and msoAnimTriggerOnMediaBookmark/msoAnimTriggerMixed are not
        // meaningful choices for add_animation/edit_animation's caller.
        private static readonly Dictionary<string, PowerPoint.MsoAnimTriggerType> AnimationTriggerMap = new Dictionary<string, PowerPoint.MsoAnimTriggerType>
        {
            ["onClick"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick,
            ["withPrevious"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious,
            ["afterPrevious"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious,
        };

        private static ToolResult AddAnimation(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            PowerPoint.Slide slide = (PowerPoint.Slide)shape.Parent;
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;

            string effectKey = input.GetProperty("effect").GetString();
            PowerPoint.MsoAnimEffect effectValue;
            if (!AnimationEffectMap.TryGetValue(effectKey, out effectValue))
                throw new ArgumentException("add_animation: unknown effect '" + effectKey + "'. Valid: " + string.Join(", ", AnimationEffectMap.Keys) + ".");

            string triggerKey = input.TryGetProperty("trigger", out var trEl) && trEl.ValueKind == JsonValueKind.String ? trEl.GetString() : "onClick";
            PowerPoint.MsoAnimTriggerType triggerValue;
            if (!AnimationTriggerMap.TryGetValue(triggerKey, out triggerValue))
                throw new ArgumentException("add_animation: unknown trigger '" + triggerKey + "'. Valid: " + string.Join(", ", AnimationTriggerMap.Keys) + ".");

            bool isExit = input.TryGetProperty("exit", out var exitEl) && exitEl.ValueKind == JsonValueKind.True;

            // UNVERIFIED (plan's own risk note, PP-24): -1 is documented
            // (not independently reflection-confirmed - reflection gives
            // signatures, not runtime semantics) to mean "append at the end
            // of the sequence." Test this specifically before trusting
            // multi-animation ordering.
            PowerPoint.Effect effect = sequence.AddEffect(shape, effectValue, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, triggerValue, -1);

            if (isExit)
                effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
            if (input.TryGetProperty("durationSeconds", out var durEl) && durEl.ValueKind == JsonValueKind.Number)
                effect.Timing.Duration = (float)durEl.GetDouble();
            if (input.TryGetProperty("delaySeconds", out var delEl) && delEl.ValueKind == JsonValueKind.Number)
                effect.Timing.TriggerDelayTime = (float)delEl.GetDouble();

            int newIndex = sequence.Count - 1; // AddEffect(-1) appends; stable immediately after the call
            return new ToolResult
            {
                Output = "Animation added at animationIndex " + newIndex + " ('" + effectKey + "', " + (isExit ? "exit" : "entrance") + ").",
                Mutated = true,
                Summary = "add_animation",
            };
        }

        private static ToolResult ReadAnimations(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
            int count = sequence.Count;
            if (count == 0)
                return new ToolResult { Output = "No animations on slide " + slideIndex + ".", Summary = "read_animations" };

            var sb = new StringBuilder();
            sb.AppendLine("Slide " + slideIndex + " has " + count + " animation(s):");
            for (int i = 1; i <= count; i++)
            {
                PowerPoint.Effect effect = sequence[i];
                string effectName = null;
                foreach (var kv in AnimationEffectMap) { if (kv.Value == effect.EffectType) { effectName = kv.Key; break; } }
                string shapeName;
                try { shapeName = effect.Shape.Name; } catch { shapeName = "(unknown shape)"; }
                string triggerName = null;
                foreach (var kv in AnimationTriggerMap) { if (kv.Value == effect.Timing.TriggerType) { triggerName = kv.Key; break; } }
                sb.AppendLine("[" + (i - 1) + "] shape=\"" + shapeName + "\" effect=" + (effectName ?? ("unrecognized (" + effect.EffectType + ")")) +
                    " kind=" + (effect.Exit == Microsoft.Office.Core.MsoTriState.msoTrue ? "exit" : "entrance") +
                    " trigger=" + (triggerName ?? effect.Timing.TriggerType.ToString()) +
                    " duration=" + effect.Timing.Duration + "s delay=" + effect.Timing.TriggerDelayTime + "s");
            }
            return new ToolResult { Output = sb.ToString().TrimEnd(), Summary = "read_animations" };
        }

        private static ToolResult EditAnimation(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int animationIndex = input.GetProperty("animationIndex").GetInt32();
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
            if (animationIndex < 0 || animationIndex >= sequence.Count)
                throw new ArgumentOutOfRangeException("animationIndex", "animationIndex must be between 0 and " + (sequence.Count - 1) + " (" + sequence.Count + " animation(s) on this slide).");
            PowerPoint.Effect effect = sequence[animationIndex + 1];

            string kind = input.GetProperty("kind").GetString();
            switch (kind)
            {
                case "delete":
                    effect.Delete();
                    return new ToolResult { Output = "Animation " + animationIndex + " deleted. Later animation indices have shifted - re-read (read_animations) before another edit in the same run.", Mutated = true, Summary = "edit_animation" };
                case "set_timing":
                {
                    bool changed = false;
                    if (input.TryGetProperty("durationSeconds", out var durEl) && durEl.ValueKind == JsonValueKind.Number) { effect.Timing.Duration = (float)durEl.GetDouble(); changed = true; }
                    if (input.TryGetProperty("delaySeconds", out var delEl) && delEl.ValueKind == JsonValueKind.Number) { effect.Timing.TriggerDelayTime = (float)delEl.GetDouble(); changed = true; }
                    if (input.TryGetProperty("trigger", out var trEl) && trEl.ValueKind == JsonValueKind.String)
                    {
                        PowerPoint.MsoAnimTriggerType triggerValue;
                        if (!AnimationTriggerMap.TryGetValue(trEl.GetString(), out triggerValue))
                            throw new ArgumentException("edit_animation: unknown trigger '" + trEl.GetString() + "'. Valid: " + string.Join(", ", AnimationTriggerMap.Keys) + ".");
                        effect.Timing.TriggerType = triggerValue;
                        changed = true;
                    }
                    if (!changed)
                        throw new ArgumentException("edit_animation: set_timing requires at least one of durationSeconds, delaySeconds, trigger.");
                    return new ToolResult { Output = "Animation " + animationIndex + " timing updated.", Mutated = true, Summary = "edit_animation" };
                }
                case "reorder":
                {
                    int toIndex = input.GetProperty("toIndex").GetInt32();
                    if (toIndex < 0 || toIndex >= sequence.Count)
                        throw new ArgumentOutOfRangeException("toIndex", "toIndex must be between 0 and " + (sequence.Count - 1) + ".");
                    // UNVERIFIED (plan's own risk note, PP-24): MoveTo's exact
                    // 0-based-vs-1-based indexing convention was confirmed to
                    // EXIST via reflection but not independently confirmed
                    // for its runtime semantics. If this lands one position
                    // off, MoveBefore/MoveAfter (relative-position moves
                    // against another Effect reference) are the fallback.
                    effect.MoveTo(toIndex + 1);
                    return new ToolResult { Output = "Animation moved from " + animationIndex + " to " + toIndex + ". Indices have shifted - re-read (read_animations) before another edit in the same run.", Mutated = true, Summary = "edit_animation" };
                }
                default:
                    throw new ArgumentException("edit_animation: unknown kind '" + kind + "'. Valid: delete, set_timing, reorder.");
            }
        }

        private static ToolResult SetElementFill(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            string fill = input.GetProperty("fill").GetString();
            if (fill == "none")
            {
                shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            else
            {
                shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                shape.Fill.ForeColor.RGB = ColorUtil.HexToOle(fill);
            }
            return new ToolResult { Output = "Fill updated.", Mutated = true, Summary = "set_element_fill" };
        }

        private static ToolResult SetElementStroke(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            bool remove = input.TryGetProperty("remove", out var r) && r.ValueKind == JsonValueKind.True;
            if (remove)
            {
                shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            else
            {
                shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                if (input.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
                {
                    shape.Line.ForeColor.RGB = ColorUtil.HexToOle(color.GetString());
                }
                shape.Line.Weight = input.TryGetProperty("widthPt", out var width) && width.ValueKind == JsonValueKind.Number ? (float)width.GetDouble() : 1f;
            }
            return new ToolResult { Output = "Stroke updated.", Mutated = true, Summary = "set_element_stroke" };
        }

        private static ToolResult SetSlideBackground(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int oleColor = ColorUtil.HexToOle(input.GetProperty("color").GetString());
            PowerPoint.Slides slides = ActivePresentation.Slides;

            void Apply(PowerPoint.Slide s)
            {
                s.Background.Fill.ForeColor.RGB = oleColor;
                s.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
            }

            if (slideIndex == -1)
            {
                foreach (PowerPoint.Slide s in slides) Apply(s);
            }
            else
            {
                Apply(slides[slideIndex + 1]);
            }
            return new ToolResult { Output = "Background updated.", Mutated = true, Summary = "set_slide_background" };
        }

        private static ToolResult UngroupElement(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            shape.Ungroup();
            return new ToolResult { Output = "Shape ungrouped - re-read the slide (read_slide) to get updated shape indices before addressing the promoted children.", Mutated = true, Summary = "ungroup_element" };
        }

        private static ToolResult AddTable(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int rows = input.GetProperty("rows").GetInt32();
            int cols = input.GetProperty("cols").GetInt32();
            float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
            float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
            float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
            float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 200f;

            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            PowerPoint.Shape tableShape = slide.Shapes.AddTable(rows, cols, left, top, width, height);
            if (input.TryGetProperty("cells", out var cells) && cells.ValueKind == JsonValueKind.Array)
            {
                int r = 0;
                foreach (JsonElement rowEl in cells.EnumerateArray())
                {
                    int c = 0;
                    foreach (JsonElement cellEl in rowEl.EnumerateArray())
                    {
                        string cellText = cellEl.GetString();
                        PowerPoint.TextRange cellRange = tableShape.Table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange;
                        cellRange.Text = cellText;
                        ApplyAutoDirection(cellRange, cellText);
                        c++;
                    }
                    r++;
                }
            }
            return new ToolResult { Output = "Table added.", Mutated = true, Summary = "add_table" };
        }

        private static PowerPoint.Table ResolveTable(JsonElement input)
        {
            return ResolveShape(input).Table;
        }

        private static ToolResult EditTableCell(JsonElement input)
        {
            PowerPoint.Table table = ResolveTable(input);
            int row = input.GetProperty("row").GetInt32();
            int col = input.GetProperty("col").GetInt32();
            string text = input.GetProperty("paragraphs").GetString();
            PowerPoint.TextRange range = table.Cell(row + 1, col + 1).Shape.TextFrame.TextRange;
            range.Text = text;
            ApplyAutoDirection(range, text);
            return new ToolResult { Output = "Cell updated.", Mutated = true, Summary = "edit_table_cell" };
        }

        private static ToolResult EditTableStructure(JsonElement input)
        {
            PowerPoint.Table table = ResolveTable(input);
            string kind = input.GetProperty("kind").GetString();
            int index = input.GetProperty("index").GetInt32();
            bool before = input.TryGetProperty("before", out var b) && b.ValueKind == JsonValueKind.True;
            // index always addresses an EXISTING row/column (0-based); before/
            // after decides which side of it the new one goes on - so the valid
            // range is the same for insert and delete. Un-validated, an
            // out-of-range index threw a raw, unhelpful COM error.
            switch (kind)
            {
                case "insert-row":
                    if (index < 0 || index >= table.Rows.Count)
                        throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Rows.Count - 1) + " for insert-row.");
                    table.Rows.Add(before ? index + 1 : index + 2);
                    break;
                case "delete-row":
                    if (index < 0 || index >= table.Rows.Count)
                        throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Rows.Count - 1) + " for delete-row.");
                    table.Rows[index + 1].Delete();
                    break;
                case "insert-col":
                    if (index < 0 || index >= table.Columns.Count)
                        throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Columns.Count - 1) + " for insert-col.");
                    table.Columns.Add(before ? index + 1 : index + 2);
                    break;
                case "delete-col":
                    if (index < 0 || index >= table.Columns.Count)
                        throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Columns.Count - 1) + " for delete-col.");
                    table.Columns[index + 1].Delete();
                    break;
                default:
                    return new ToolResult { Output = "Unknown structure kind: " + kind, IsError = true, Summary = "edit_table_structure" };
            }
            // Deleting/inserting a row or column shifts every later row/column's
            // index - the same trap PP-19's delete_slide has. Callers doing more
            // than one structural edit in a run should re-read the table between
            // calls; documented in the schema description too.
            return new ToolResult { Output = "Table structure updated.", Mutated = true, Summary = "edit_table_structure" };
        }

        private static ToolResult EditTableStyle(JsonElement input)
        {
            PowerPoint.Table table = ResolveTable(input);
            if (input.TryGetProperty("firstRow", out var firstRow))
            {
                table.FirstRow = firstRow.ValueKind == JsonValueKind.True;
            }
            if (input.TryGetProperty("bandRow", out var bandRow))
            {
                table.HorizBanding = bandRow.ValueKind == JsonValueKind.True;
            }
            if (input.TryGetProperty("shadingColor", out var shading) && shading.ValueKind == JsonValueKind.String)
            {
                int color = ColorUtil.HexToOle(shading.GetString());
                foreach (PowerPoint.Row row in table.Rows)
                {
                    foreach (PowerPoint.Cell cell in row.Cells)
                    {
                        cell.Shape.Fill.ForeColor.RGB = color;
                    }
                }
            }
            if (input.TryGetProperty("borderColor", out _) || input.TryGetProperty("borderWidthPt", out _) || input.TryGetProperty("borderPreset", out _))
            {
                string preset = input.TryGetProperty("borderPreset", out var bp) && bp.ValueKind == JsonValueKind.String ? bp.GetString() : "all";
                if (preset != "all" && preset != "none" && preset != "outline")
                    throw new ArgumentException("edit_table_style: unknown borderPreset '" + preset + "'. Valid: all, none, outline.");
                bool visible = preset != "none";
                float weight = input.TryGetProperty("borderWidthPt", out var bw) && bw.ValueKind == JsonValueKind.Number ? (float)bw.GetDouble() : 1f;
                int color = input.TryGetProperty("borderColor", out var bc) && bc.ValueKind == JsonValueKind.String ? ColorUtil.HexToOle(bc.GetString()) : ColorUtil.HexToOle("#000000");
                PowerPoint.PpBorderType[] sides = { PowerPoint.PpBorderType.ppBorderTop, PowerPoint.PpBorderType.ppBorderBottom, PowerPoint.PpBorderType.ppBorderLeft, PowerPoint.PpBorderType.ppBorderRight };
                int rowCount = table.Rows.Count;
                int colCount = table.Columns.Count;
                int rIdx = 0;
                foreach (PowerPoint.Row row in table.Rows)
                {
                    int cIdx = 0;
                    foreach (PowerPoint.Cell cell in row.Cells)
                    {
                        foreach (PowerPoint.PpBorderType side in sides)
                        {
                            // "outline" only draws the table's outer perimeter -
                            // suppress every interior edge per-cell rather than
                            // per-table, reusing the existing per-cell loop.
                            bool sideVisible = visible;
                            if (visible && preset == "outline")
                            {
                                sideVisible = (side == PowerPoint.PpBorderType.ppBorderTop && rIdx == 0)
                                           || (side == PowerPoint.PpBorderType.ppBorderBottom && rIdx == rowCount - 1)
                                           || (side == PowerPoint.PpBorderType.ppBorderLeft && cIdx == 0)
                                           || (side == PowerPoint.PpBorderType.ppBorderRight && cIdx == colCount - 1);
                            }
                            PowerPoint.LineFormat border = cell.Borders[side];
                            border.Visible = sideVisible ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
                            if (sideVisible)
                            {
                                border.Weight = weight;
                                border.ForeColor.RGB = color;
                            }
                        }
                        cIdx++;
                    }
                    rIdx++;
                }
            }
            return new ToolResult { Output = "Table style updated.", Mutated = true, Summary = "edit_table_style" };
        }

        // PP-21: one chart vocabulary across the repo, matching ExcelChartTypeMap
        // (ExcelAiAddIn/ExcelTools.cs) exactly. "bar" was previously mapped to 51
        // (xlColumnClustered, Excel's "column" code) instead of 57
        // (xlBarClustered) - a silent wrong result: even a *successful*
        // chartType:'bar' produced a column chart. Fixed here; "barStacked" had
        // the identical bug (52 = xlColumnStacked, not 58 = xlBarStacked).
        private static readonly Dictionary<string, int> PptChartTypeMap = new Dictionary<string, int>
        {
            ["column"] = 51,       // xlColumnClustered
            ["columnStacked"] = 52,// xlColumnStacked
            ["bar"] = 57,          // xlBarClustered
            ["barStacked"] = 58,   // xlBarStacked
            ["line"] = 4,          // xlLine
            ["area"] = 1,          // xlArea
            ["pie"] = 5,           // xlPie
            ["doughnut"] = -4120,  // xlDoughnut
        };

        // Post-hoc fix (2026-08-24, user-reported, same root cause ported
        // from WordTools.cs's identical fix): the embedded chart-data
        // workbook's OLE server occasionally still throws "The remote
        // procedure call failed" (HRESULT 0x800706BE) even after the
        // Clear()+batched-write fix below - a known, documented transient
        // failure mode for rapid COM calls against Office's embedded chart
        // Excel object. Only the specific known transient RPC HRESULTs are
        // retried, so a genuine logic error still fails immediately.
        private static readonly int[] TransientComHResults =
        {
            unchecked((int)0x800706BE), // RPC_S_CALL_FAILED
            unchecked((int)0x8001010A), // RPC_E_SERVERCALL_RETRYLATER
            unchecked((int)0x800706BA), // RPC_S_SERVER_UNAVAILABLE
        };

        private static void RetryTransientCom(Action action)
        {
            const int maxAttempts = 3;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try { action(); return; }
                // Post-hoc fix (2026-08-24, ported from Word's identical fix): widened
                // from COMException to the base Exception type, still filtered by
                // HResult - a COM error surfaced through dynamic late-binding is not
                // guaranteed to arrive as a raw COMException.
                catch (Exception ex) when (attempt < maxAttempts && Array.IndexOf(TransientComHResults, ex.HResult) >= 0)
                {
                    System.Threading.Thread.Sleep(200 * attempt);
                }
            }
        }

        private static ToolResult AddChartPpt(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            string kindStr = input.GetProperty("kind").GetString();
            int typeCode;
            if (!PptChartTypeMap.TryGetValue(kindStr, out typeCode))
                throw new ArgumentException("add_chart: unknown kind '" + kindStr + "'. Valid: " +
                                            string.Join(", ", PptChartTypeMap.Keys) + ".");
            var categories = new List<string>();
            foreach (JsonElement c in input.GetProperty("categories").EnumerateArray()) categories.Add(c.GetString());

            // PP-22: a series/categories length mismatch would otherwise write a
            // ragged grid and produce a silently wrong chart - reject it up front.
            JsonElement seriesForValidation = input.GetProperty("series");
            foreach (JsonElement s in seriesForValidation.EnumerateArray())
            {
                int valueCount = 0;
                foreach (JsonElement v in s.GetProperty("values").EnumerateArray()) valueCount++;
                if (valueCount != categories.Count)
                {
                    string seriesName = s.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String ? nameEl.GetString() : "(unnamed)";
                    throw new ArgumentException("add_chart: series '" + seriesName + "' has " + valueCount +
                                                " value(s) but there are " + categories.Count +
                                                " categories - every series must have exactly one value per category.");
                }
            }

            float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
            float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
            float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
            float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 300f;

            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            dynamic chartShape = slide.Shapes.AddChart2(-1, (Microsoft.Office.Core.XlChartType)typeCode, left, top, width, height);
            // 0-based index into slide.Shapes - AddChart2 appends the new shape at
            // the end of the collection, so this is stable immediately after the
            // call. Returned below so the model can edit_chart it without a
            // separate read_slide round trip (PP-22 Task 1 Step 5).
            int newShapeIndex = slide.Shapes.Count - 1;
            dynamic chart = chartShape.Chart;

            // Chart data lives in an embedded Excel workbook - open, write the grid,
            // close, and RELEASE explicitly so no hidden Excel host process leaks.
            // Post-hoc fix (2026-08-24, ported from Word's identical fix,
            // found via a real repro's DebugLog): moved inside the retry
            // lambda below so a flaky OPEN, not just a flaky subsequent
            // call, also gets retried.
            dynamic dataWorkbook = null;
            try
            {
                var seriesList = input.GetProperty("series").EnumerateArray().ToList();

                // Build the whole grid in memory up front (pure C#, no COM) -
                // only the write itself needs to go through RetryTransientCom.
                int rowCount = categories.Count + 1; // +1 header row
                int colCount = seriesList.Count + 1; // +1 category column
                object[,] grid = new object[rowCount, colCount];
                grid[0, 0] = "";
                int colIdx = 0;
                foreach (JsonElement s in seriesList)
                {
                    grid[0, colIdx + 1] = s.GetProperty("name").GetString();
                    colIdx++;
                }
                for (int r = 0; r < categories.Count; r++)
                {
                    grid[r + 1, 0] = categories[r];
                }
                colIdx = 0;
                foreach (JsonElement s in seriesList)
                {
                    int r = 0;
                    foreach (JsonElement v in s.GetProperty("values").EnumerateArray())
                    {
                        grid[r + 1, colIdx + 1] = v.GetDouble();
                        r++;
                    }
                    colIdx++;
                }

                RetryTransientCom(() =>
                {
                    // Post-hoc fix (2026-08-24, user-reported the RPC failure
                    // recurring even after the first fix - ported from
                    // Word's identical fix): a brief settle delay right after
                    // the embedded OLE workbook opens, before the first COM
                    // call against it - the automation surface is not always
                    // fully live the instant ChartData.Workbook returns.
                    System.Threading.Thread.Sleep(120);

                    dataWorkbook = chart.ChartData.Workbook;
                    dynamic sheet = dataWorkbook.Worksheets[1];

                    // Confirmed repro (Word's identical port of this same pattern,
                    // PP-9): a brand-new chart's embedded workbook comes pre-seeded
                    // by Office with placeholder sample data (a default chart
                    // template, commonly 4 categories x 3 series). Without
                    // clearing it first, only the cells the NEW data actually
                    // occupies get overwritten - any leftover placeholder cells
                    // stay in the sheet and get plotted alongside the real data.
                    sheet.Cells.Clear();

                    dynamic topLeft = sheet.Cells[1, 1];
                    dynamic writeRange = topLeft.Resize[rowCount, colCount];
                    writeRange.Value2 = grid;

                    // ACTUAL ROOT CAUSE (2026-08-24, confirmed via .NET
                    // reflection against the real referenced
                    // Microsoft.Office.Interop.PowerPoint.dll, not a guess -
                    // same finding as Word's identical code):
                    // PowerPoint.Chart.SetSourceData's real signature is
                    // SetSourceData(String Source, Object PlotBy) - the
                    // first parameter is a STRING, not a Range. Every prior
                    // attempt (a reused writeRange, then a Range built from
                    // an A1 string) was passing a Range COM object where the
                    // method actually expects a plain "SheetName!A1:B4"
                    // reference string - no Range object needed at all.
                    string a1Range = "A1:" + TextUtil.ColumnLetter(colCount) + rowCount;
                    string sourceRef = (string)sheet.Name + "!" + a1Range;
                    chart.SetSourceData(sourceRef);
                });
            }
            finally
            {
                // ROOT CAUSE FOUND (2026-08-24): this cleanup had no catch of
                // its own - when SetSourceData failed above, the chart/
                // embedded-workbook was left in a state where
                // dataWorkbook.Close() ALSO threw a second, unrelated
                // exception, which (per C# finally semantics) REPLACED the
                // real SetSourceData exception before it ever reached the
                // caller. Cleanup failures are now caught and swallowed
                // (not re-thrown) so they can never mask a real exception
                // already in flight.
                if (dataWorkbook != null)
                {
                    try
                    {
                        dataWorkbook.Close(SaveChanges: true);
                    }
                    catch { /* secondary failure - do not mask the real exception */ }
                    finally
                    {
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(dataWorkbook); } catch { }
                    }
                }
            }

            if (input.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
            }
            return new ToolResult { Output = "Chart added at shapeIndex " + newShapeIndex + ".", Mutated = true, Summary = "add_chart" };
        }

        // PP-21: legendPos's natural names (a model will say "right", not the
        // genoffice-ism "r") plus the original short aliases for back-compat.
        // xlLegendPositionCorner was considered and dropped - its code could not
        // be verified against a live Office install (no interactive GUI access
        // in this environment), and guessing it wrong would just replace one
        // silent-wrong-result bug with another.
        private static readonly Dictionary<string, int> PptLegendPositions =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["right"] = -4152, ["r"] = -4152,   // xlLegendPositionRight
            ["top"] = -4160, ["t"] = -4160,     // xlLegendPositionTop
            ["left"] = -4131, ["l"] = -4131,    // xlLegendPositionLeft
            ["bottom"] = -4107, ["b"] = -4107,  // xlLegendPositionBottom
        };

        private static ToolResult EditChartPpt(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            if (shape.HasChart != Microsoft.Office.Core.MsoTriState.msoTrue)
                throw new InvalidOperationException("edit_chart: shape " + input.GetProperty("shapeIndex").GetInt32() +
                                                    " on slide " + input.GetProperty("slideIndex").GetInt32() + " is not a chart.");
            dynamic chart = shape.Chart;
            var applied = new List<string>();

            if (input.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String)
            {
                int typeCode;
                if (!PptChartTypeMap.TryGetValue(ct.GetString(), out typeCode))
                    throw new ArgumentException("edit_chart: unknown chartType '" + ct.GetString() + "'. Valid: " +
                                                string.Join(", ", PptChartTypeMap.Keys) + ".");
                chart.ChartType = typeCode;
                applied.Add("chartType=" + ct.GetString());
            }
            if (input.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
                applied.Add("title");
            }
            if (input.TryGetProperty("legendPos", out var legendPos) && legendPos.ValueKind == JsonValueKind.String)
            {
                string pos = legendPos.GetString();
                if (string.Equals(pos, "none", StringComparison.OrdinalIgnoreCase))
                {
                    chart.HasLegend = false;
                }
                else
                {
                    int posCode;
                    if (!PptLegendPositions.TryGetValue(pos, out posCode))
                        throw new ArgumentException("edit_chart: unknown legendPos '" + pos + "'. Valid: none, " +
                                                    string.Join(", ", PptLegendPositions.Keys) + ".");
                    chart.HasLegend = true;
                    chart.Legend.Position = posCode;
                }
                applied.Add("legendPos=" + pos);
            }
            if (input.TryGetProperty("dataLabels", out var dl) && dl.ValueKind == JsonValueKind.String)
            {
                // Matches Excel's edit_chart vocabulary (none|value|percent) -
                // previously a boolean, which silently turned labels OFF for any
                // non-true value, including the string "value".
                string mode = dl.GetString();
                if (mode != "none" && mode != "value" && mode != "percent")
                    throw new ArgumentException("edit_chart: unknown dataLabels '" + mode + "'. Valid: none, value, percent.");
                bool show = mode != "none";
                foreach (dynamic series in chart.SeriesCollection())
                {
                    series.HasDataLabels = show;
                    if (show && mode == "percent") series.DataLabels().ShowPercentage = true;
                }
                applied.Add("dataLabels=" + mode);
            }
            if (input.TryGetProperty("gridlines", out var gl) && (gl.ValueKind == JsonValueKind.True || gl.ValueKind == JsonValueKind.False))
            {
                bool show = gl.ValueKind == JsonValueKind.True;
                try
                {
                    chart.Axes(2 /* xlValue */).HasMajorGridlines = show;
                    applied.Add("gridlines=" + show);
                }
                catch (Exception)
                {
                    throw new InvalidOperationException(
                        "edit_chart: this chart's type has no value axis (e.g. pie/doughnut) - gridlines do not apply to it.");
                }
            }
            return new ToolResult
            {
                Output = applied.Count > 0
                    ? "Chart updated: " + string.Join(", ", applied) + "."
                    : "No recognized chart properties were provided - nothing changed.",
                Mutated = applied.Count > 0,
                Summary = "edit_chart",
            };
        }

        // Verified against the standard English display names for PowerPoint's built-in SmartArt
        // layout gallery. Live cross-check against Application.SmartArtLayouts on this machine's
        // Office install (plan Task 6 Step 1) requires interactive Office GUI access that was not
        // available in this environment - remains a manual follow-up for a human with GUI access.
        private static readonly Dictionary<string, string> SmartArtLayoutNames = new Dictionary<string, string>
        {
            ["list"] = "Basic Block List",
            ["process"] = "Basic Process",
            ["cycle"] = "Basic Cycle",
            ["hierarchy"] = "Organization Chart",
            ["pyramid"] = "Basic Pyramid",
            ["matrix"] = "Basic Matrix",
            ["venn"] = "Basic Venn",
        };

        // PP-22 Task 2 Step 4: index-based lookup (SmartArtLayouts is
        // index-addressable, and the built-in gallery order is stable across
        // installs) was considered as a locale-independent alternative to
        // name-matching. Not switched - index stability across Office versions
        // is an assumption no better founded than the display-name assumption,
        // and names at least fail loudly with a diagnostic message below,
        // whereas a wrong index would silently insert the wrong diagram.
        private static dynamic ResolveSmartArtLayout(string layoutKey)
        {
            string targetName;
            if (!SmartArtLayoutNames.TryGetValue(layoutKey, out targetName))
                throw new ArgumentException("add_smartart: unknown layout '" + layoutKey + "'. Valid: " +
                                            string.Join(", ", SmartArtLayoutNames.Keys) + ".");
            dynamic layouts = Globals.ThisAddIn.Application.SmartArtLayouts;
            foreach (dynamic layout in layouts)
            {
                if (string.Equals((string)layout.Name, targetName, StringComparison.OrdinalIgnoreCase))
                {
                    return layout;
                }
            }
            // Distinct from the unknown-key case above: the key was valid, but
            // this Office install's gallery has no layout under that display
            // name - the one diagnosis that actually points at a non-English
            // install, which a silent fallback could never surface.
            throw new InvalidOperationException("add_smartart: no SmartArt layout named '" + targetName +
                                                "' was found in this Office install's gallery - this install may be " +
                                                "non-English, where the built-in gallery's display names differ from " +
                                                "the standard English ones this tool assumes (see plan Task 6 Step 1).");
        }

        private static ToolResult AddSmartArt(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            string layoutKey = input.GetProperty("layout").GetString();
            float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
            float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
            float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
            float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 300f;

            dynamic layout = ResolveSmartArtLayout(layoutKey);
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            dynamic shape = slide.Shapes.AddSmartArt(layout, left, top, width, height);
            dynamic smartArt = shape.SmartArt;

            // Post-hoc fix (2026-08-24, found via Word's identical port of
            // this code, PP-23): AddSmartArt seeds the diagram with the
            // layout's own default "[Text]" placeholder nodes. Without
            // clearing them first, the requested items were appended after
            // the placeholders instead of replacing them.
            dynamic existingNodes = smartArt.Nodes;
            for (int i = (int)existingNodes.Count; i >= 1; i--)
            {
                existingNodes.Item(i).Delete();
            }

            // genoffice's own version only ever produces a flat item list - maps
            // 1:1 to sequential top-level nodes, no nested tree-building needed.
            foreach (JsonElement item in input.GetProperty("items").EnumerateArray())
            {
                dynamic node = smartArt.Nodes.Add();
                node.TextFrame2.TextRange.Text = item.GetString();
            }
            return new ToolResult { Output = "SmartArt added.", Mutated = true, Summary = "add_smartart" };
        }

        private static ToolResult CropImage(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            float l = (float)input.GetProperty("l").GetDouble();
            float t = (float)input.GetProperty("t").GetDouble();
            float r = (float)input.GetProperty("r").GetDouble();
            float b = (float)input.GetProperty("b").GetDouble();
            // Approximation, documented deliberately: fractions are applied against
            // the shape's CURRENT on-slide size, not the original uncropped source
            // image - classic Interop has no reliable "natural size" property once a
            // picture has already been resized/cropped on the slide. Correct for a
            // freshly-inserted, never-before-cropped picture; imprecise under
            // repeated crop calls on the same shape.
            shape.PictureFormat.CropLeft = l * shape.Width;
            shape.PictureFormat.CropTop = t * shape.Height;
            shape.PictureFormat.CropRight = r * shape.Width;
            shape.PictureFormat.CropBottom = b * shape.Height;
            return new ToolResult { Output = "Image cropped.", Mutated = true, Summary = "crop_image" };
        }

        private static ToolResult ReplaceImagePpt(JsonElement input)
        {
            string localPath = input.GetProperty("localPath").GetString();
            if (localPath.StartsWith("http://") || localPath.StartsWith("https://"))
            {
                return new ToolResult { Output = "replace_image: remote URLs are not supported in this air-gapped deployment - use a local file path.", IsError = true, Summary = "replace_image" };
            }
            if (!System.IO.File.Exists(localPath))
            {
                return new ToolResult { Output = "replace_image: file not found: " + localPath, IsError = true, Summary = "replace_image" };
            }
            PowerPoint.Shape oldShape = ResolveShape(input);
            bool keepCrop = input.TryGetProperty("keepCrop", out var kc) && kc.ValueKind == JsonValueKind.True;

            float left = oldShape.Left, top = oldShape.Top, width = oldShape.Width, height = oldShape.Height, rotation = oldShape.Rotation;
            int zPos = oldShape.ZOrderPosition;
            float cropLeft = 0, cropTop = 0, cropRight = 0, cropBottom = 0;
            if (keepCrop)
            {
                cropLeft = oldShape.PictureFormat.CropLeft;
                cropTop = oldShape.PictureFormat.CropTop;
                cropRight = oldShape.PictureFormat.CropRight;
                cropBottom = oldShape.PictureFormat.CropBottom;
            }
            PowerPoint.Slide slide = (PowerPoint.Slide)oldShape.Parent;
            oldShape.Delete();

            PowerPoint.Shape newShape = slide.Shapes.AddPicture(localPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, left, top, width, height);
            newShape.Rotation = rotation;
            if (keepCrop)
            {
                newShape.PictureFormat.CropLeft = cropLeft;
                newShape.PictureFormat.CropTop = cropTop;
                newShape.PictureFormat.CropRight = cropRight;
                newShape.PictureFormat.CropBottom = cropBottom;
            }
            // Restore approximate z-order: send to back, then bring forward to the
            // original stack position.
            newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
            for (int i = 1; i < zPos; i++)
            {
                newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringForward);
            }
            return new ToolResult { Output = "Image replaced.", Mutated = true, Summary = "replace_image" };
        }

        private static ToolResult SetPictureOpacity(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            float opacity = (float)input.GetProperty("opacity").GetDouble();
            dynamic dShape = shape;
            dShape.Fill.Transparency = 1f - opacity;
            return new ToolResult { Output = "Opacity updated.", Mutated = true, Summary = "set_picture_opacity" };
        }
    }
}
