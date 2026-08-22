using System;
using System.Text;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static class PowerPointTools
    {
        public static ToolResult Execute(string name, JsonElement input)
        {
            try
            {
                switch (name)
                {
                    case "get_deck_context": return GetDeckContext();
                    case "read_slide": return ReadSlide(input);
                    default: return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        private static PowerPoint.Presentation ActivePresentation => Globals.ThisAddIn.Application.ActivePresentation;

        private static string ShapeText(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                return shape.TextFrame.TextRange.Text;
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
            int shapeIndex = 0;
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                sb.AppendLine($"[{shapeIndex}] {shape.Name}: {ShapeText(shape)}");
                shapeIndex++;
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_slide" };
        }
    }
}
