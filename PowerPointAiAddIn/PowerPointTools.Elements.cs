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
        // A shape reference resolved from a tool's slideIndex + shapeIndex.
        // shapeIndex is normally a 0-based number (a top-level shape on the
        // slide). It may also be a dotted-path STRING like "3.1.0" - shape 3 on
        // the slide, its child 1 (a group), that group's child 0 - to address a
        // shape nested inside one or more groups. read_group prints these paths.
        private struct ShapeRef
        {
            public PowerPoint.Shape Shape;
            public bool IsNested;   // true when the path descended into a group
            public string Path;     // normalized dotted path, for messages
            public int TopIndex;    // 0-based index of the top-level ancestor
        }

        private static int[] ParseShapePath(JsonElement shapeIndexEl)
        {
            if (shapeIndexEl.ValueKind == JsonValueKind.Number)
                return new[] { shapeIndexEl.GetInt32() };
            if (shapeIndexEl.ValueKind == JsonValueKind.String)
            {
                string raw = shapeIndexEl.GetString() ?? "";
                string[] parts = raw.Split('.');
                var path = new int[parts.Length];
                for (int i = 0; i < parts.Length; i++)
                {
                    if (!int.TryParse(parts[i].Trim(), out path[i]) || path[i] < 0)
                        throw new ArgumentException("shapeIndex '" + raw + "' is not valid - use a 0-based number, or a dotted path like \"3.1.0\" for a shape inside a group.");
                }
                return path;
            }
            throw new ArgumentException("shapeIndex must be a 0-based number, or a dotted path string like \"3.1.0\" for a shape inside a group.");
        }

        private static ShapeRef ResolveShapeRef(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int[] path = ParseShapePath(input.GetProperty("shapeIndex"));
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];

            PowerPoint.Shape shape = slide.Shapes[path[0] + 1];
            for (int i = 1; i < path.Length; i++)
            {
                if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoGroup)
                    throw new ArgumentException("shapeIndex path '" + string.Join(".", path) + "' - the shape at ." +
                                                string.Join(".", path.Take(i)) + " is not a group, so it has no children.");
                shape = shape.GroupItems[path[i] + 1];
            }
            return new ShapeRef { Shape = shape, IsNested = path.Length > 1, Path = string.Join(".", path), TopIndex = path[0] };
        }

        private static PowerPoint.Shape ResolveShape(JsonElement input)
        {
            return ResolveShapeRef(input).Shape;
        }

        // For tools whose behavior is undefined or confusing on a shape nested
        // inside a group (positional/structural edits): resolve, but refuse a
        // path target with a message that tells the model how to proceed.
        private static PowerPoint.Shape ResolveTopLevelShape(JsonElement input, string toolName)
        {
            ShapeRef r = ResolveShapeRef(input);
            if (r.IsNested)
                throw new ArgumentException(toolName + ": shape " + r.Path + " is inside a group. Call ungroup_element on shapeIndex " +
                                            r.TopIndex + " first, then address the promoted shape by its new top-level index.");
            return r.Shape;
        }

        private static PowerPoint.Slide ShapeSlide(PowerPoint.Shape shape)
        {
            try { return shape.Parent as PowerPoint.Slide; } catch { return null; }
        }

        // Shared by ApplyOptionalName below and DuplicateElement/CopyOrMoveElement
        // (PowerPointTools.CrossSlide.cs) - disambiguates `desired` against every
        // OTHER shape's current Name with a numeric suffix ("Rectangle 5" ->
        // "Rectangle 5 2"), the same convention read_slide/read_group key their
        // output on. siblingShapes defaults to resolving from the shape's own
        // slide; pass it explicitly for a Slide Master/layout shape, whose parent
        // is never a PowerPoint.Slide. See PowerPointTools.Elements.cs.md for why
        // this is one method with a defaultable parameter, not two overloads.
        private static string MakeUniqueNameOnSlide(PowerPoint.Shape shape, string desired, PowerPoint.Shapes siblingShapes = null)
        {
            if (siblingShapes == null)
            {
                PowerPoint.Slide slide = ShapeSlide(shape);
                siblingShapes = slide != null ? slide.Shapes : null;
            }

            string unique = desired;
            if (siblingShapes != null)
            {
                var taken = new HashSet<string>();
                foreach (PowerPoint.Shape s in siblingShapes)
                    if (s.Id != shape.Id) taken.Add(s.Name);
                int suffix = 2;
                while (taken.Contains(unique)) unique = desired + " " + suffix++;
            }
            return unique;
        }

        // Optional model-chosen shape name. PowerPoint permits duplicate names,
        // but read_slide/read_group key their output on the name, so a collision
        // among the sibling shapes is disambiguated with a numeric suffix.
        // Returns the name actually applied, or null when none was requested.
        private static string ApplyOptionalName(PowerPoint.Shape shape, JsonElement input, PowerPoint.Shapes siblingShapes = null)
        {
            if (!input.TryGetProperty("name", out var nameEl) || nameEl.ValueKind != JsonValueKind.String)
                return null;
            string desired = (nameEl.GetString() ?? "").Trim();
            if (desired.Length == 0) return null;
            if (desired.Length > 120) desired = desired.Substring(0, 120);

            string unique = MakeUniqueNameOnSlide(shape, desired, siblingShapes);
            shape.Name = unique;
            return unique;
        }

        private static string ShapeKindLabel(PowerPoint.Shape shape)
        {
            try
            {
                switch (shape.Type)
                {
                    case Microsoft.Office.Core.MsoShapeType.msoGroup: return "group";
                    case Microsoft.Office.Core.MsoShapeType.msoTextBox: return "text box";
                    case Microsoft.Office.Core.MsoShapeType.msoPicture: return "picture";
                    case Microsoft.Office.Core.MsoShapeType.msoLinkedPicture: return "picture";
                    case Microsoft.Office.Core.MsoShapeType.msoLine: return "line";
                    case Microsoft.Office.Core.MsoShapeType.msoFreeform: return "freeform";
                    case Microsoft.Office.Core.MsoShapeType.msoAutoShape: return "shape";
                    case Microsoft.Office.Core.MsoShapeType.msoPlaceholder: return "placeholder";
                    case Microsoft.Office.Core.MsoShapeType.msoChart: return "chart";
                    case Microsoft.Office.Core.MsoShapeType.msoTable: return "table";
                    case Microsoft.Office.Core.MsoShapeType.msoSmartArt: return "SmartArt";
                    case Microsoft.Office.Core.MsoShapeType.msoDiagram: return "SmartArt";
                    case Microsoft.Office.Core.MsoShapeType.msoMedia: return "media";
                    default: return shape.Type.ToString();
                }
            }
            catch { return "shape"; }
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

        // PowerPoint's TextRange never auto-flips direction/alignment for RTL text
        // the way Word's editor does - decided per write from the text's own script
        // mix via TextUtil.IsRtlMajority. See PowerPointTools.Elements.cs.md.
        private static void ApplyAutoDirection(PowerPoint.TextRange range, string text)
        {
            if (!TextUtil.IsRtlMajority(text)) return;
            range.ParagraphFormat.TextDirection = PowerPoint.PpDirection.ppDirectionRightToLeft;
            range.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignRight;
        }

        // Gives the model an explicit bulleted on/off switch instead of typing a
        // literal bullet character (which double-bullets against a placeholder's
        // own native bullet). Omitted `bulleted` leaves the existing setting
        // untouched. See PowerPointTools.Elements.cs.md for the full bug report.
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
        // this file's ChartTypes.ByName/SmartArtLayouts.ByName dictionary pattern.
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
            // Strikethrough deliberately NOT implemented: this PIA's TextFrame has no
            // Strikethrough member, and Microsoft.Office.Interop.PowerPoint has no
            // TextFrame2 type at all (confirmed absent). See .md for the full story.

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
            PowerPoint.Shape shape = ResolveTopLevelShape(input, "set_element_transform");
            if (input.TryGetProperty("left", out var left)) shape.Left = (float)left.GetDouble();
            if (input.TryGetProperty("top", out var top)) shape.Top = (float)top.GetDouble();
            if (input.TryGetProperty("width", out var width)) shape.Width = (float)width.GetDouble();
            if (input.TryGetProperty("height", out var height)) shape.Height = (float)height.GetDouble();
            if (input.TryGetProperty("rotation", out var rotation)) shape.Rotation = (float)rotation.GetDouble();
            return new ToolResult { Output = "Transform updated.", Mutated = true, Summary = "set_element_transform" };
        }

        // Stacking/z-order control, distinct from set_element_transform's
        // position/size. Only 4 of MsoZOrderCmd's 6 values are exposed - the other
        // two are relative-to-body-text, not meaningful for a slide's z-order stack.
        private static readonly Dictionary<string, Microsoft.Office.Core.MsoZOrderCmd> ZOrderMap = new Dictionary<string, Microsoft.Office.Core.MsoZOrderCmd>
        {
            ["bringToFront"] = Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront,
            ["sendToBack"] = Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack,
            ["bringForward"] = Microsoft.Office.Core.MsoZOrderCmd.msoBringForward,
            ["sendBackward"] = Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward,
        };

        private static ToolResult SetElementOrder(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveTopLevelShape(input, "set_element_order");
            string kind = input.GetProperty("kind").GetString();
            Microsoft.Office.Core.MsoZOrderCmd cmd;
            if (!ZOrderMap.TryGetValue(kind, out cmd))
                throw new ArgumentException("set_element_order: unknown kind '" + kind + "'. Valid: " + string.Join(", ", ZOrderMap.Keys) + ".");
            shape.ZOrder(cmd);
            // ZOrderPosition is 1-based in COM; reported 0-based to match
            // read_slide's shapeIndex convention.
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
            string named = ApplyOptionalName(shape, input);
            return new ToolResult { Output = "Text box added" + (named != null ? " (\"" + named + "\")" : "") + ".", Mutated = true, Summary = "add_text_box" };
        }

        // Shape-name lookup now lives in OfficeAi.Shared.ShapeTypes (Phase 0) -
        // union of this map and Excel's near-identical copy, including this
        // app's rectangle/oval aliases for rect/ellipse.
        private static ToolResult AddShape(JsonElement input)
        {
            PowerPoint.Slide slide = ActivePresentation.Slides[input.GetProperty("slideIndex").GetInt32() + 1];
            string shapeType = input.GetProperty("shapeType").GetString();
            int autoShapeTypeInt;
            if (!ShapeTypes.ByName.TryGetValue(shapeType, out autoShapeTypeInt))
                throw new ArgumentException("add_shape: unknown shapeType '" + shapeType + "'. Valid: " +
                                            string.Join(", ", ShapeTypes.ByName.Keys) + ".");
            Microsoft.Office.Core.MsoAutoShapeType autoShapeType = (Microsoft.Office.Core.MsoAutoShapeType)autoShapeTypeInt;
            float left = (float)input.GetProperty("left").GetDouble();
            float top = (float)input.GetProperty("top").GetDouble();
            float width = (float)input.GetProperty("width").GetDouble();
            float height = (float)input.GetProperty("height").GetDouble();
            PowerPoint.Shape shape = slide.Shapes.AddShape(autoShapeType, left, top, width, height);
            if (input.TryGetProperty("text", out var text)) shape.TextFrame.TextRange.Text = text.GetString();
            string named = ApplyOptionalName(shape, input);
            return new ToolResult { Output = "Shape added" + (named != null ? " (\"" + named + "\")" : "") + ".", Mutated = true, Summary = "add_shape" };
        }

        private static ToolResult DeleteElement(JsonElement input)
        {
            ResolveTopLevelShape(input, "delete_element").Delete();
            return new ToolResult { Output = "Shape deleted.", Mutated = true, Summary = "delete_element" };
        }

        // Shape.Duplicate() is a native in-place COM clone, uniform across every
        // shape kind (unlike copy_element/move_element's dispatch). The new position
        // is always computed from the ORIGINAL shape's Left/Top, since Duplicate()'s
        // own placement of the copy is unverified. See PowerPointTools.Elements.cs.md.
        private static ToolResult DuplicateElement(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveTopLevelShape(input, "duplicate_element");
            PowerPoint.ShapeRange range = shape.Duplicate();
            PowerPoint.Shape dup = range[1];

            // Each axis is independent: an explicit left/top is an exact coordinate;
            // an omitted one falls back to the default offset on that axis, not to
            // the original's exact coordinate. See .md for the bug this fixed.
            float offsetX = input.TryGetProperty("offsetX", out var ox) ? (float)ox.GetDouble() : 12f;
            float offsetY = input.TryGetProperty("offsetY", out var oy) ? (float)oy.GetDouble() : 12f;
            dup.Left = input.TryGetProperty("left", out var l) ? (float)l.GetDouble() : shape.Left + offsetX;
            dup.Top = input.TryGetProperty("top", out var t) ? (float)t.GetDouble() : shape.Top + offsetY;

            string named = ApplyOptionalName(dup, input);
            if (named == null)
            {
                // Shape.Duplicate() keeps the EXACT source Name (unlike a UI
                // Ctrl+D/paste, which auto-renames) - dedupe it the same way an
                // explicit name would be. See .md.
                string unique = MakeUniqueNameOnSlide(dup, dup.Name);
                if (unique != dup.Name)
                {
                    dup.Name = unique;
                    named = unique;
                }
            }
            // ZOrderPosition, not slide.Shapes.Count - Duplicate()'s exact
            // insertion point (end of collection vs. adjacent to the source)
            // is unverified; ZOrderPosition is correct either way, same
            // reasoning SetElementOrder/GroupElement already rely on.
            int newShapeIndex = dup.ZOrderPosition - 1;
            return new ToolResult
            {
                Output = "Shape duplicated" + (named != null ? " (\"" + named + "\")" : "") + " - new shape at shapeIndex " + newShapeIndex +
                         ". Other shapes' indices on this slide may have shifted - re-read the slide (read_slide) before addressing another shape by index in the same run.",
                Mutated = true,
                Summary = "duplicate_element",
            };
        }
    }
}

