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
        // add and duplicate both clone the slide right after itself; add then clears every text frame,
        // leaving the layout and non-text content.
        // delete is deliberately one slide per call, no batch form - see PowerPointTools.Slides.cs.md for why.
        private static ToolResult ManageSlide(JsonElement input)
        {
            string action = input.GetProperty("action").GetString();
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slide slide = ResolveSlide(input);
            PowerPoint.Slides slides = ActivePresentation.Slides;

            switch (action)
            {
                case "add":
                case "duplicate":
                    dynamic newSlide = slide.Duplicate()[1]; // SlideRange holding exactly the new slide
                    if (action == "add")
                    {
                        foreach (PowerPoint.Shape shape in newSlide.Shapes)
                        {
                            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                                shape.TextFrame.TextRange.Text = "";
                        }
                        return new ToolResult { Output = "Slide added after index " + slideIndex + ".", Mutated = true };
                    }
                    return new ToolResult
                    {
                        Output = "Duplicated slide " + slideIndex + " (content included) - the copy is now at index " + (slideIndex + 1) + ".",
                        Mutated = true,
                    };

                case "move":
                    if (!input.TryGetProperty("toIndex", out var toEl))
                        throw new ArgumentException("manage_slide: action \"move\" needs toIndex.");
                    int toIndex = toEl.GetInt32();
                    if (toIndex < 0 || toIndex >= slides.Count)
                        throw new ArgumentOutOfRangeException("toIndex", "toIndex must be between 0 and " + (slides.Count - 1) + ".");
                    slide.MoveTo(toIndex + 1);
                    return new ToolResult { Output = "Moved slide " + slideIndex + " to position " + toIndex + ".", Mutated = true };

                case "delete":
                    if (slides.Count == 1)
                        throw new InvalidOperationException("manage_slide: cannot delete the only slide in the presentation.");
                    slide.Delete();
                    return new ToolResult
                    {
                        Output = "Deleted slide " + slideIndex + ". " + slides.Count + " slide(s) remain; " +
                                 "slides after it have shifted down by one index.",
                        Mutated = true,
                    };

                default:
                    throw new ArgumentException("manage_slide: unknown action '" + action + "'. Valid: add, duplicate, move, delete.");
            }
        }

        // Applies whichever of layout / background / transition / notes are present, in that order.
        // slideIndex -1 (whole deck) is valid only for background.
        private static ToolResult SetSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            bool hasLayout = input.TryGetProperty("layout", out _) || input.TryGetProperty("layoutName", out _);
            bool hasBackground = input.TryGetProperty("background", out _);
            bool hasTransition = SlideTransitionKeys.Any(k => input.TryGetProperty(k, out _));
            bool hasNotes = input.TryGetProperty("notes", out _);

            if (!(hasLayout || hasBackground || hasTransition || hasNotes))
                throw new ArgumentException("set_slide: nothing to set - give layout/layoutName, background, notes, or effect/durationSeconds/advanceOnClick/advanceAfterSeconds.");
            if (slideIndex == -1 && (hasLayout || hasTransition || hasNotes))
                throw new ArgumentException("set_slide: slideIndex -1 (every slide) is only valid with background.");

            PowerPoint.Slide slide = slideIndex == -1 ? null : ResolveSlide(input);
            var done = new List<string>();
            try
            {
                if (hasLayout) done.Add(ApplySlideLayout(slide, slideIndex, input));
                if (hasBackground) done.Add(ApplySlideBackground(slideIndex, input));
                if (hasTransition) done.Add(ApplySlideTransition(slide, slideIndex, input));
                if (hasNotes) done.Add(ApplySlideNotes(slide, slideIndex, input));
            }
            catch (Exception ex) when (done.Count > 0)
            {
                throw new InvalidOperationException(string.Join(" ", done) + " Then failed: " + ex.Message);
            }
            return new ToolResult { Output = string.Join(" ", done), Mutated = true };
        }

    }
}

