using System;
using System.Collections.Generic;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // Plain Shape-to-Shape native-value copies (no JSON parsing, no hex
        // round trip) - shared by copy_element_style below and by
        // PowerPointTools.CrossSlide.cs's text-box/autoshape reconstruction
        // paths, which call these for fill/stroke/text fidelity after
        // building the destination shape.

        // Mixed-formatting design decision: samples the FIRST character/
        // paragraph only, not the whole TextRange - guarantees a single,
        // well-defined value regardless of whether the source's text is
        // uniformly formatted, matching real Format Painter's "the format
        // where you picked up from" semantics. Sidesteps needing to know
        // whether this PIA has a Word-style "mixed value" sentinel for a
        // whole-range read (unverified - see the plan's Risks).
        private static void CopyTextFormatting(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            if (source.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return;
            if (source.TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoTrue) return;
            if (target.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return;

            PowerPoint.TextRange srcText = source.TextFrame.TextRange;
            PowerPoint.Font srcFont = srcText.Characters(1, 1).Font;
            PowerPoint.ParagraphFormat srcPara = srcText.Paragraphs(1, 1).ParagraphFormat;

            PowerPoint.TextRange targetRange = target.TextFrame.TextRange;
            targetRange.Font.Bold = srcFont.Bold;
            targetRange.Font.Italic = srcFont.Italic;
            targetRange.Font.Size = srcFont.Size;
            targetRange.Font.Color.RGB = srcFont.Color.RGB;
            targetRange.Font.Name = srcFont.Name;
            targetRange.Font.Underline = srcFont.Underline;
            targetRange.Font.Shadow = srcFont.Shadow;
            targetRange.Font.Superscript = srcFont.Superscript;
            targetRange.Font.Subscript = srcFont.Subscript;
            targetRange.ParagraphFormat.Alignment = srcPara.Alignment;
            // Strikethrough deliberately not copied - same PIA gap
            // set_element_style already documents (no TextFrame2/TextRange2
            // on this PIA's Shape).
        }

        private static void CopyFillFormatting(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            target.Fill.Visible = source.Fill.Visible;
            if (source.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                target.Fill.ForeColor.RGB = source.Fill.ForeColor.RGB;
        }

        private static void CopyStrokeFormatting(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            target.Line.Visible = source.Line.Visible;
            if (source.Line.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                target.Line.ForeColor.RGB = source.Line.ForeColor.RGB;
                target.Line.Weight = source.Line.Weight;
            }
        }

        // Format painter. Same-slide only (v1 scope decision - no cross-slide
        // target matcher exists or is being built here). Source is addressed
        // via the standard slideIndex+shapeIndex (ResolveTopLevelShape, same
        // as every other single-shape PowerPoint tool) rather than a
        // "sourceShapeIndex" field, so it gets the existing nested-shape
        // rejection for free. Never touches position/size.
        private static ToolResult CopyElementStyle(JsonElement input)
        {
            PowerPoint.Shape source = ResolveTopLevelShape(input, "copy_element_style");

            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];

            if (!input.TryGetProperty("targetShapeIndexes", out var idxArr) || idxArr.ValueKind != JsonValueKind.Array)
                throw new ArgumentException("copy_element_style: targetShapeIndexes must be an array of at least one 0-based top-level shape index.");

            var targets = new List<PowerPoint.Shape>();
            var seen = new HashSet<int>();
            foreach (JsonElement el in idxArr.EnumerateArray())
            {
                int i = el.GetInt32();
                if (i < 0 || i >= slide.Shapes.Count)
                    throw new ArgumentException("copy_element_style: shapeIndex " + i + " is out of range (slide has " + slide.Shapes.Count + " shapes).");
                if (seen.Add(i)) targets.Add(slide.Shapes[i + 1]); // permissive dedup, matching group_element
            }
            if (targets.Count == 0)
                throw new ArgumentException("copy_element_style: targetShapeIndexes must contain at least one index.");

            bool hasSourceText = source.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                              && source.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue;

            foreach (PowerPoint.Shape target in targets)
            {
                if (hasSourceText) CopyTextFormatting(source, target);
                CopyFillFormatting(source, target);
                CopyStrokeFormatting(source, target);
            }

            string what = hasSourceText ? "text, fill, stroke" : "fill, stroke (source has no text)";
            return new ToolResult
            {
                Output = "Style copied to " + targets.Count + " shape(s): " + what + ".",
                Mutated = true,
                Summary = "copy_element_style",
            };
        }
    }
}
