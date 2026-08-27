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

        // PP-19 Task 1: delete_slide/move_slide/duplicate_slide. Deliberately no
        // slideIndexes:number[] batch form - deleting slide 2 shifts every later
        // slide's index down by one, so a batch would need to either resolve all
        // targets up front or delete in strict descending order. One slide per
        // call (with the index-shift warning in the output/description) is the
        // safer answer; it makes the model re-read the deck between deletes
        // instead of silently deleting the wrong slides.
        private static ToolResult DeleteSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Presentation pres = ActivePresentation;
            if (slideIndex < 0 || slideIndex >= pres.Slides.Count)
                throw new ArgumentOutOfRangeException("slideIndex",
                    "slideIndex must be between 0 and " + (pres.Slides.Count - 1) + ".");
            if (pres.Slides.Count == 1)
                throw new InvalidOperationException(
                    "delete_slide: cannot delete the only slide in the presentation.");
            pres.Slides[slideIndex + 1].Delete();
            return new ToolResult
            {
                Output = "Deleted slide " + slideIndex + ". " + pres.Slides.Count + " slide(s) remain; " +
                         "slides after it have shifted down by one index.",
                Mutated = true,
                Summary = "delete_slide",
            };
        }

        private static ToolResult MoveSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int toIndex = input.GetProperty("toIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (slideIndex < 0 || slideIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("slideIndex",
                    "slideIndex must be between 0 and " + (slides.Count - 1) + ".");
            if (toIndex < 0 || toIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("toIndex",
                    "toIndex must be between 0 and " + (slides.Count - 1) + ".");
            slides[slideIndex + 1].MoveTo(toIndex + 1);
            return new ToolResult
            {
                Output = "Moved slide " + slideIndex + " to position " + toIndex + ".",
                Mutated = true,
                Summary = "move_slide",
            };
        }

        private static ToolResult DuplicateSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (slideIndex < 0 || slideIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("slideIndex",
                    "slideIndex must be between 0 and " + (slides.Count - 1) + ".");
            slides[slideIndex + 1].Duplicate();
            return new ToolResult
            {
                Output = "Duplicated slide " + slideIndex + " (content included) - the copy is now at index " + (slideIndex + 1) + ".",
                Mutated = true,
                Summary = "duplicate_slide",
            };
        }

    }
}

