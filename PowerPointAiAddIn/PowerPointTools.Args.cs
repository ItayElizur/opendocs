using System;
using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // Resolves a 0-based slide index argument to the slide; throws (Execute turns it into an IsError result).
        private static PowerPoint.Slide ResolveSlide(JsonElement input, string key = "slideIndex")
        {
            int index = input.GetProperty(key).GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (index < 0 || index >= slides.Count)
                throw new ArgumentOutOfRangeException(key, key + " must be between 0 and " + (slides.Count - 1) + ".");
            return slides[index + 1];
        }

        // Required left/top/width/height (add_text_box, add_shape).
        private static void ReadRect(JsonElement input, out float left, out float top, out float width, out float height)
        {
            left = (float)input.GetProperty("left").GetDouble();
            top = (float)input.GetProperty("top").GetDouble();
            width = (float)input.GetProperty("width").GetDouble();
            height = (float)input.GetProperty("height").GetDouble();
        }

        // Optional x/y/w/h with the 100,100,400,defaultHeight fallback (add_table/chart/smartart).
        private static void ReadXywh(JsonElement input, float defaultHeight, out float left, out float top, out float width, out float height)
        {
            left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
            top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
            width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
            height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : defaultHeight;
        }

        // Applies whichever of left/top/width/height/rotation are present; returns the names applied.
        private static List<string> ApplyTransform(PowerPoint.Shape shape, JsonElement input)
        {
            var applied = new List<string>();
            if (input.TryGetProperty("left", out var left)) { shape.Left = (float)left.GetDouble(); applied.Add("left"); }
            if (input.TryGetProperty("top", out var top)) { shape.Top = (float)top.GetDouble(); applied.Add("top"); }
            if (input.TryGetProperty("width", out var width)) { shape.Width = (float)width.GetDouble(); applied.Add("width"); }
            if (input.TryGetProperty("height", out var height)) { shape.Height = (float)height.GetDouble(); applied.Add("height"); }
            if (input.TryGetProperty("rotation", out var rotation)) { shape.Rotation = (float)rotation.GetDouble(); applied.Add("rotation"); }
            return applied;
        }

        // Writes text and applies the RTL auto-direction in one step.
        private static void SetText(PowerPoint.TextRange range, string text)
        {
            range.Text = text;
            ApplyAutoDirection(range, text);
        }

        private static MsoTriState Tri(bool value) => value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
    }
}
