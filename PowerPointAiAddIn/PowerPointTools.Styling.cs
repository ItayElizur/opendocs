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

    }
}

