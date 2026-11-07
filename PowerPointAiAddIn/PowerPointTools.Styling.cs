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
        private static readonly string[] TextStyleKeys = { "bold", "italic", "underline", "shadow", "fontSize", "fontName", "color", "alignment", "baselineOffset" };

        // One kind of formatting per call: text style, fill, stroke, or opacity.
        private static ToolResult FormatElement(JsonElement input)
        {
            switch (RequireOneGroup("format_element", input,
                Group("text style", TextStyleKeys),
                Group("fill", "fill"),
                Group("stroke", "strokeColor", "strokeWidthPt", "removeStroke"),
                Group("opacity", "opacity")))
            {
                case "text style": return SetElementStyle(input);
                case "fill": return SetElementFill(input);
                case "stroke": return SetElementStroke(input);
                default: return SetPictureOpacity(input);
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
            return new ToolResult { Output = "Fill updated.", Mutated = true };
        }

        private static ToolResult SetElementStroke(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            bool remove = ToolArgs.OptionalFlag(input, "removeStroke");
            if (remove)
            {
                shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            else
            {
                shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                if (input.TryGetProperty("strokeColor", out var color) && color.ValueKind == JsonValueKind.String)
                {
                    shape.Line.ForeColor.RGB = ColorUtil.HexToOle(color.GetString());
                }
                shape.Line.Weight = ToolArgs.OptionalFloat(input, "strokeWidthPt") ?? 1f;
            }
            return new ToolResult { Output = "Stroke updated.", Mutated = true };
        }

        // slideIndex -1 paints every slide.
        private static string ApplySlideBackground(int slideIndex, JsonElement input)
        {
            int oleColor = ColorUtil.HexToOle(input.GetProperty("background").GetString());

            void Apply(PowerPoint.Slide s)
            {
                s.Background.Fill.ForeColor.RGB = oleColor;
                s.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
            }

            if (slideIndex == -1)
            {
                foreach (PowerPoint.Slide s in ActivePresentation.Slides) Apply(s);
                return "Background updated on every slide.";
            }
            Apply(ResolveSlide(input));
            return "Slide " + slideIndex + " background updated.";
        }

        private static ToolResult UngroupElement(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            shape.Ungroup();
            return new ToolResult { Output = "Shape ungrouped - re-read the slide (read_slide) to get updated shape indices before addressing the promoted children.", Mutated = true };
        }

        // action "group" (shapeIndexes) or "ungroup" (shapeIndex).
        private static ToolResult GroupElement(JsonElement input)
        {
            string action = input.GetProperty("action").GetString();
            switch (action)
            {
                case "group": return GroupShapes(input);
                case "ungroup": return UngroupElement(input);
                default: throw new ArgumentException("group_element: unknown action '" + action + "'. Valid: group, ungroup.");
            }
        }

        private static ToolResult GroupShapes(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slide slide = ResolveSlide(input);

            if (!input.TryGetProperty("shapeIndexes", out var idxArr) || idxArr.ValueKind != JsonValueKind.Array)
                return new ToolResult { Output = "group_element needs shapeIndexes: an array of at least two 0-based top-level shape indices.", IsError = true };

            var seen = new HashSet<int>();
            var comIndexes = new List<object>();
            foreach (JsonElement el in idxArr.EnumerateArray())
            {
                int i = el.GetInt32();
                if (i < 0 || i >= slide.Shapes.Count)
                    return new ToolResult { Output = "group_element: shapeIndex " + i + " is out of range (slide has " + slide.Shapes.Count + " shapes).", IsError = true };
                if (seen.Add(i)) comIndexes.Add(i + 1); // ShapeRange is 1-based
            }
            if (comIndexes.Count < 2)
                return new ToolResult { Output = "group_element needs at least two distinct shapeIndexes.", IsError = true };

            PowerPoint.ShapeRange range = slide.Shapes.Range(comIndexes.ToArray());
            PowerPoint.Shape group = range.Group();

            string named = ApplyOptionalName(group, input);
            int newIndex = group.ZOrderPosition - 1;
            return new ToolResult
            {
                Output = "Grouped " + comIndexes.Count + " shapes into " + (named ?? group.Name) + " at shapeIndex " + newIndex +
                         ". Other shapes' indices on this slide shifted - call read_slide before addressing another shape by index in the same run. " +
                         "Address the grouped children with a dotted path (e.g. \"" + newIndex + ".0\") via read_slide (view=group).",
                Mutated = true,
            };
        }

    }
}

