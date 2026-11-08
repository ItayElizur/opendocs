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
                        throw new ArgumentException("toIndex must be between 0 and " + (slides.Count - 1) + ".");
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

        // One kind of change per call (layout, background, transition, or notes), so a call is one edit
        // and one undo step, and a failure can never leave a half-applied slide.
        // The transition fields (effect, duration, advance options) count as one change.
        // slideIndex -1 (whole deck) is valid only for background.
        private static ToolResult SetSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            var groups = new List<string>();
            if (input.TryGetProperty("layout", out _) || input.TryGetProperty("layoutName", out _)) groups.Add("layout");
            if (input.TryGetProperty("background", out _)) groups.Add("background");
            if (SlideTransitionKeys.Any(k => input.TryGetProperty(k, out _))) groups.Add("transition");
            if (input.TryGetProperty("notes", out _)) groups.Add("notes");

            if (groups.Count == 0)
                throw new ArgumentException("set_slide: nothing to set - give layout/layoutName, background, notes, or the transition fields (effect/durationSeconds/advanceOnClick/advanceAfterSeconds).");
            if (groups.Count > 1)
                throw new ArgumentException("set_slide: change one thing per call - got " + string.Join(", ", groups) + ". Make a separate set_slide call for each.");
            if (slideIndex == -1 && groups[0] != "background")
                throw new ArgumentException("set_slide: slideIndex -1 (every slide) is only valid with background.");

            PowerPoint.Slide slide = slideIndex == -1 ? null : ResolveSlide(input);
            string message;
            switch (groups[0])
            {
                case "layout": message = ApplySlideLayout(slide, slideIndex, input); break;
                case "background": message = ApplySlideBackground(slideIndex, input); break;
                case "transition": message = ApplySlideTransition(slide, slideIndex, input); break;
                default: message = ApplySlideNotes(slide, slideIndex, input); break;
            }
            return new ToolResult { Output = message, Mutated = true };
        }

    }
}

