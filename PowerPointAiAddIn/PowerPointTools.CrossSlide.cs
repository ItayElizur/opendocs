using System;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // copy_element (move:true): cross-slide relocation via PowerPoint's own
        // native Copy()/Paste() (Shape.Copy() + Slide.Shapes.Paste()), not property
        // reconstruction - COM-surface gaps (SmartArt, grouping, table styling,
        // gradients) made reconstruction unreliable. This is the only tool in this
        // codebase that touches the real Windows clipboard; SaveClipboard/
        // RestoreClipboard below mitigate (not eliminate) the clobbering risk.
        // Full rationale and history: PowerPointTools.CrossSlide.cs.md.

        // Best-effort clipboard save: GetDataObject() can itself throw (the
        // clipboard is a shared OS resource), caught so a failed save doesn't
        // block the copy/move - it only widens the restore gap.
        private static object SaveClipboard()
        {
            try { return System.Windows.Forms.Clipboard.GetDataObject(); }
            catch { return null; }
        }

        // Restores the saved clipboard payload. Deliberately does NOT clear the
        // clipboard when nothing was saved - leaving our own shape's data there is
        // less destructive than guessing it's safe to wipe. Passes "copy: false" to
        // SetDataObject (not true) - eager flush of every format was measured to
        // dominate runtime for multi-format payloads. See PowerPointTools.CrossSlide.cs.md.
        private static void RestoreClipboard(object saved)
        {
            var dataObj = saved as System.Windows.Forms.IDataObject;
            if (dataObj == null) return;
            try { System.Windows.Forms.Clipboard.SetDataObject(dataObj, false); }
            catch { /* best-effort - see SaveClipboard's own comment */ }
        }

        // A completely empty shape (no text/fill/content) makes PowerPoint's Copy()
        // put no pasteable payload on the clipboard, raising a cryptic raw COM error -
        // caught and re-thrown with the likely cause named plainly. See .md for the
        // exact native error text this was confirmed against.
        private static PowerPoint.Shape CopyPasteShape(PowerPoint.Shape source, PowerPoint.Slide destSlide)
        {
            PowerPoint.ShapeRange pasted;
            try
            {
                source.Copy();
                pasted = destSlide.Shapes.Paste();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "Copy/paste failed for this shape - the most common cause is a completely empty shape " +
                    "(no text, no fill, no other content), which PowerPoint's own Copy() doesn't put a pasteable payload on " +
                    "the clipboard for. PowerPoint's own error: " + ex.Message, ex);
            }
            if (pasted.Count != 1)
            {
                // Whatever Paste() actually put on the destination slide is a
                // real orphan at this point - the caller's own dest-cleanup
                // try/catch never runs, since dest is never assigned when this
                // throws. Clean it up here instead of leaving it behind while
                // reporting a failure.
                try { pasted.Delete(); } catch { }
                throw new InvalidOperationException("Paste produced " + pasted.Count + " shape(s) instead of exactly 1.");
            }
            return pasted[1];
        }

        private static ToolResult CopyOrMoveElement(JsonElement input, bool cut)
        {
            const string toolName = "copy_element";
            PowerPoint.Shape source = ResolveTopLevelShape(input, toolName);

            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int targetSlideIndex = input.GetProperty("targetSlideIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (targetSlideIndex < 0 || targetSlideIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("targetSlideIndex", toolName + ": targetSlideIndex must be between 0 and " + (slides.Count - 1) + ".");

            PowerPoint.Slide destSlide = slides[targetSlideIndex + 1];

            object savedClipboard = SaveClipboard();
            try
            {
                PowerPoint.Shape dest = CopyPasteShape(source, destSlide);

                // The paste already succeeded by this point - delete the orphaned
                // dest shape before rethrowing so a failure here doesn't silently
                // leave it behind. See PowerPointTools.CrossSlide.cs.md.
                try
                {
                    // A native paste lands at PowerPoint's own default position;
                    // honor an explicit left/top override the same way
                    // copy_element (same slide) does, otherwise leave it alone.
                    if (input.TryGetProperty("left", out var l)) dest.Left = (float)l.GetDouble();
                    if (input.TryGetProperty("top", out var t)) dest.Top = (float)t.GetDouble();

                    string named = ApplyOptionalName(dest, input);
                    if (named == null)
                    {
                        // A pasted shape can keep the source's exact Name, colliding
                        // with it if the source and destination slides share names.
                        string unique = MakeUniqueNameOnSlide(dest, source.Name);
                        if (unique != dest.Name) dest.Name = unique;
                    }
                }
                catch
                {
                    try { dest.Delete(); } catch { }
                    throw;
                }

                int newShapeIndex = dest.ZOrderPosition - 1;
                if (cut)
                {
                    // Mirrors the positioning/naming guarantee above: if deleting
                    // the source fails after a successful paste, delete dest too
                    // rather than leave two copies behind.
                    try { source.Delete(); }
                    catch
                    {
                        try { dest.Delete(); } catch { }
                        throw;
                    }
                }

                string action = cut ? "moved" : "copied";
                string extra = cut ? " The shape has been removed from its original slide - other shapes' indices there may have shifted too." : "";
                return new ToolResult
                {
                    Output = "Shape " + action + " to slide " + targetSlideIndex + " - new shape at shapeIndex " + newShapeIndex +
                             ". Other shapes' indices on the destination slide may have shifted." + extra +
                             " Re-read the affected slide(s) (read_slide) before addressing another shape by index in the same run.",
                    Mutated = true,
                };
            }
            finally
            {
                RestoreClipboard(savedClipboard);
            }
        }

        // copy_element: a target on another slide goes through the clipboard (move:true also removes the
        // original); no target, or the same slide, is a native same-slide Duplicate.
        private static ToolResult CopyElement(JsonElement input)
        {
            bool move = ToolArgs.OptionalFlag(input, "move");
            bool otherSlide = input.TryGetProperty("targetSlideIndex", out var target) &&
                              target.GetInt32() != input.GetProperty("slideIndex").GetInt32();
            if (otherSlide) return CopyOrMoveElement(input, move);
            if (move)
                throw new ArgumentException("copy_element: move needs a targetSlideIndex on a different slide - to reposition within a slide use set_element_transform.");
            return DuplicateElement(input);
        }
    }
}
