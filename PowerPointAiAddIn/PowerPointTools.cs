using System;
using System.Text;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    // Mirrors the Word/Excel Task 11/16 editing-mode enum pattern.
    public enum EditingMode { ReadOnly, CommentOnly, TrackChanges, FullAutonomy }

    public static class PowerPointTools
    {
        // Editing-mode gating (mirrors the Word/Excel Task 11/16 pattern this plan establishes
        // elsewhere): the tool list offered to the model is filtered client-side per mode
        // (web-src/entry.ts, first line of defense - smaller prompts, fewer wasted turns), but
        // Execute() independently re-checks mode here as defense-in-depth, since nothing stops a
        // misbehaving or malicious model response from calling a tool that wasn't offered.
        public static EditingMode Mode = EditingMode.FullAutonomy;

        // Tools always allowed regardless of editing mode (read-only, no document mutation).
        private static readonly System.Collections.Generic.HashSet<string> AlwaysAllowedTools =
            new System.Collections.Generic.HashSet<string> { "get_deck_context", "read_slide" };

        public static ToolResult Execute(string name, JsonElement input)
        {
            try
            {
                if (!AlwaysAllowedTools.Contains(name) && !IsMutationAllowed(Mode))
                {
                    return new ToolResult
                    {
                        Output = "Blocked: editing mode is " + ModeLabel(Mode) + ".",
                        IsError = true,
                        Summary = name,
                    };
                }

                switch (name)
                {
                    case "get_deck_context": return GetDeckContext();
                    case "read_slide": return ReadSlide(input);
                    case "set_element_text": return SetElementText(input);
                    case "set_element_style": return SetElementStyle(input);
                    case "set_element_transform": return SetElementTransform(input);
                    case "add_text_box": return AddTextBox(input);
                    case "add_shape": return AddShape(input);
                    case "delete_element": return DeleteElement(input);
                    case "add_slide": return AddSlide(input);
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

        private static PowerPoint.Shape ResolveShape(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int shapeIndex = input.GetProperty("shapeIndex").GetInt32();
            return ActivePresentation.Slides[slideIndex + 1].Shapes[shapeIndex + 1];
        }

        private static ToolResult SetElementText(JsonElement input)
        {
            ResolveShape(input).TextFrame.TextRange.Text = input.GetProperty("text").GetString();
            return new ToolResult { Output = "Text updated.", Mutated = true, Summary = "set_element_text" };
        }

        private static ToolResult SetElementStyle(JsonElement input)
        {
            PowerPoint.TextRange range = ResolveShape(input).TextFrame.TextRange;
            if (input.TryGetProperty("bold", out var bold)) range.Font.Bold = bold.GetBoolean() ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            if (input.TryGetProperty("italic", out var italic)) range.Font.Italic = italic.GetBoolean() ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            if (input.TryGetProperty("fontSize", out var fontSize)) range.Font.Size = (float)fontSize.GetDouble();
            if (input.TryGetProperty("color", out var color))
            {
                string hex = color.GetString().TrimStart('#');
                int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                range.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
            }
            return new ToolResult { Output = "Style updated.", Mutated = true, Summary = "set_element_style" };
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

        private static ToolResult AddTextBox(JsonElement input)
        {
            PowerPoint.Slide slide = ActivePresentation.Slides[input.GetProperty("slideIndex").GetInt32() + 1];
            float left = (float)input.GetProperty("left").GetDouble();
            float top = (float)input.GetProperty("top").GetDouble();
            float width = (float)input.GetProperty("width").GetDouble();
            float height = (float)input.GetProperty("height").GetDouble();
            PowerPoint.Shape shape = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            shape.TextFrame.TextRange.Text = input.GetProperty("text").GetString();
            return new ToolResult { Output = "Text box added.", Mutated = true, Summary = "add_text_box" };
        }

        private static ToolResult AddShape(JsonElement input)
        {
            PowerPoint.Slide slide = ActivePresentation.Slides[input.GetProperty("slideIndex").GetInt32() + 1];
            string shapeType = input.GetProperty("shapeType").GetString();
            Microsoft.Office.Core.MsoAutoShapeType autoShapeType =
                shapeType == "oval" ? Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval :
                shapeType == "roundRect" ? Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle :
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle;
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
    }
}
