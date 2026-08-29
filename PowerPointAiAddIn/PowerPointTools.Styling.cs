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

        private static ToolResult GroupElement(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];

            if (!input.TryGetProperty("shapeIndexes", out var idxArr) || idxArr.ValueKind != JsonValueKind.Array)
                return new ToolResult { Output = "group_element needs shapeIndexes: an array of at least two 0-based top-level shape indices.", IsError = true, Summary = "group_element" };

            var seen = new HashSet<int>();
            var comIndexes = new List<object>();
            foreach (JsonElement el in idxArr.EnumerateArray())
            {
                int i = el.GetInt32();
                if (i < 0 || i >= slide.Shapes.Count)
                    return new ToolResult { Output = "group_element: shapeIndex " + i + " is out of range (slide has " + slide.Shapes.Count + " shapes).", IsError = true, Summary = "group_element" };
                if (seen.Add(i)) comIndexes.Add(i + 1); // ShapeRange is 1-based
            }
            if (comIndexes.Count < 2)
                return new ToolResult { Output = "group_element needs at least two distinct shapeIndexes.", IsError = true, Summary = "group_element" };

            PowerPoint.ShapeRange range = slide.Shapes.Range(comIndexes.ToArray());
            PowerPoint.Shape group = range.Group();

            string named = ApplyOptionalName(group, input);
            int newIndex = group.ZOrderPosition - 1;
            return new ToolResult
            {
                Output = "Grouped " + comIndexes.Count + " shapes into " + (named ?? group.Name) + " at shapeIndex " + newIndex +
                         ". Other shapes' indices on this slide shifted - call read_slide before addressing another shape by index in the same run. " +
                         "Address the grouped children with a dotted path (e.g. \"" + newIndex + ".0\") via read_group.",
                Mutated = true,
                Summary = "group_element",
            };
        }

    }
}

