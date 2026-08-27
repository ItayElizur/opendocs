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
            return new ToolResult { Output = "Shape added.", Mutated = true, Summary = "add_shape" };
        }

        private static ToolResult DeleteElement(JsonElement input)
        {
            ResolveShape(input).Delete();
            return new ToolResult { Output = "Shape deleted.", Mutated = true, Summary = "delete_element" };
        }

    }
}

