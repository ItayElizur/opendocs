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

        // Optional model-chosen shape name. PowerPoint permits duplicate names,
        // but read_slide/read_group key their output on the name, so a collision
        // on the same slide is disambiguated with a numeric suffix. Returns the
        // name actually applied, or null when none was requested.
        private static string ApplyOptionalName(PowerPoint.Shape shape, JsonElement input)
        {
            if (!input.TryGetProperty("name", out var nameEl) || nameEl.ValueKind != JsonValueKind.String)
                return null;
            string desired = (nameEl.GetString() ?? "").Trim();
            if (desired.Length == 0) return null;
            if (desired.Length > 120) desired = desired.Substring(0, 120);

            PowerPoint.Slide slide = ShapeSlide(shape);
            string unique = desired;
            if (slide != null)
            {
                var taken = new HashSet<string>();
                foreach (PowerPoint.Shape s in slide.Shapes)
                    if (s.Id != shape.Id) taken.Add(s.Name);
                int suffix = 2;
                while (taken.Contains(unique)) unique = desired + " " + suffix++;
            }
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
            PowerPoint.Shape shape = ResolveTopLevelShape(input, "set_element_transform");
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
            PowerPoint.Shape shape = ResolveTopLevelShape(input, "set_element_order");
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

    }
}

