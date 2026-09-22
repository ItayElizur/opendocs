using System;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // copy_element/move_element: cross-slide relocation via PowerPoint's
        // own native Copy/Paste (Shape.Copy() + Slide.Shapes.Paste()).
        //
        // History (2026-09-23): an earlier version of this file avoided the
        // OS clipboard entirely, instead reading a source shape's properties
        // and manually reconstructing an equivalent on the destination slide
        // (matching this codebase's documented rule for Word - never the
        // clipboard, since it clobbers the user's and races with anything
        // else in the session). After several rounds of live testing, that
        // approach kept hitting real PowerPoint object-model gaps that no
        // amount of extra property-copying could close: a SmartArt graphic
        // has no COM API to read its layout's own placeholder text/hierarchy
        // (had to be rebuilt node-by-node, losing layout-specific structure),
        // PowerPoint flatly refuses to Group() a SmartArt with any other
        // shape (so a reconstructed nested group could never truly nest one),
        // table cell shading/borders/merged-cell structure aren't fully
        // readable through this PIA, and gradient fills don't round-trip
        // cleanly through GradientStops. A native Copy()/Paste() sidesteps
        // ALL of this at once, because it reproduces the exact underlying
        // OOXML the same way Ctrl+C/Ctrl+V in the UI does - there is no
        // "kind" this can fail to support.
        //
        // The tradeoff is real and was made deliberately, on the user's
        // explicit instruction (2026-09-23: "we'll deal with the clipboard
        // issue somehow") after the reconstruction approach was judged not
        // worth its remaining gaps: this DOES touch the real Windows
        // clipboard, unlike every other tool in this codebase. The
        // clobbering/racing risk is mitigated, not eliminated, by saving and
        // restoring the clipboard's prior contents around the operation
        // (SaveClipboard/RestoreClipboard below) - a narrow window remains
        // where something else changes the clipboard mid-call, the same risk
        // any programmatic clipboard user (including a PowerPoint macro)
        // accepts. duplicate_element (PowerPointTools.Elements.cs) and
        // copy_element_style (PowerPointTools.FormatPainter.cs) are
        // unaffected - Shape.Duplicate() and Shape.PickUp()/Apply() are both
        // native, non-clipboard mechanisms and still avoid this tradeoff
        // entirely.

        // Best-effort save/restore of whatever was on the clipboard before
        // this tool touched it. GetDataObject() can itself throw (the
        // clipboard is a shared OS resource another process can be holding
        // momentarily) - caught, since failing to SAVE shouldn't block the
        // copy/move itself, only widen (not create) the restore gap.
        private static object SaveClipboard()
        {
            try { return System.Windows.Forms.Clipboard.GetDataObject(); }
            catch { return null; }
        }

        // Restores a previously saved clipboard payload. Deliberately does
        // NOT clear the clipboard when nothing was saved (e.g. SaveClipboard
        // itself failed, or the clipboard was already empty in a way
        // GetDataObject couldn't wrap) - leaving our own shape's data there
        // is less destructive than guessing it's safe to wipe.
        //
        // Real-user-confirmed (2026-09-23): this whole operation was much
        // slower than a manual Ctrl+C/Ctrl+V. Root cause: SetDataObject's
        // second argument ("copy") - originally passed as true - tells
        // Windows to eagerly render and flush EVERY format the data object
        // exposes right now, so the clipboard survives even after the owning
        // app exits. That's the right call for something like "user hit
        // Ctrl+C, now leaving the app" but not for restoring a save taken a
        // few milliseconds ago in the same still-running session - a
        // multi-format payload (e.g. an image someone had copied, which
        // carries bitmap/PNG/DIB/etc. simultaneously) can make that eager
        // flush the dominant cost of the whole tool call. Passing false
        // leaves the data lazily/delay-rendered (satisfied on demand by this
        // same live IDataObject reference) - correct for restoring within an
        // active session, and avoids the flush entirely.
        private static void RestoreClipboard(object saved)
        {
            var dataObj = saved as System.Windows.Forms.IDataObject;
            if (dataObj == null) return;
            try { System.Windows.Forms.Clipboard.SetDataObject(dataObj, false); }
            catch { /* best-effort - see SaveClipboard's own comment */ }
        }

        // Real-user-confirmed (2026-09-23): a completely empty default text
        // box (no text, no fill, no other content) raised PowerPoint's own
        // raw COM error - "Shapes (unknown member): Invalid request.
        // Clipboard is empty or contains data which may not be pasted here."
        // - because Copy() on a shape with nothing to render doesn't put a
        // pasteable payload on the clipboard at all. Caught and re-thrown
        // with the likely cause named plainly, instead of surfacing that
        // cryptic native message as-is.
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
                throw new InvalidOperationException("Paste produced " + pasted.Count + " shape(s) instead of exactly 1.");
            return pasted[1];
        }

        private static ToolResult CopyOrMoveElement(JsonElement input, bool cut)
        {
            string toolName = cut ? "move_element" : "copy_element";
            PowerPoint.Shape source = ResolveTopLevelShape(input, toolName);

            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int targetSlideIndex = input.GetProperty("targetSlideIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (targetSlideIndex < 0 || targetSlideIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("targetSlideIndex", toolName + ": targetSlideIndex must be between 0 and " + (slides.Count - 1) + ".");
            if (targetSlideIndex == slideIndex)
                throw new ArgumentException(toolName + ": source and target are the same slide - use duplicate_element for a same-slide copy, or set_element_transform to reposition within a slide.");

            PowerPoint.Slide destSlide = slides[targetSlideIndex + 1];

            object savedClipboard = SaveClipboard();
            try
            {
                PowerPoint.Shape dest = CopyPasteShape(source, destSlide);

                // A native paste lands wherever PowerPoint's own paste logic
                // puts it (typically the same Left/Top as the source, on top
                // of whatever's already on the destination slide) - honor an
                // explicit left/top override the same way duplicate_element
                // does, otherwise leave PowerPoint's own placement alone.
                if (input.TryGetProperty("left", out var l)) dest.Left = (float)l.GetDouble();
                if (input.TryGetProperty("top", out var t)) dest.Top = (float)t.GetDouble();

                string named = ApplyOptionalName(dest, input);
                if (named == null)
                {
                    // Real-user-confirmed (2026-09-22, live testing, same
                    // root cause as DuplicateElement's own fix): a pasted
                    // shape can keep the exact source Name, colliding with
                    // it on the destination slide if the source's own slide
                    // happens to share names with the destination's.
                    string unique = MakeUniqueNameOnSlide(dest, source.Name);
                    if (unique != dest.Name) dest.Name = unique;
                }

                int newShapeIndex = dest.ZOrderPosition - 1;
                if (cut) source.Delete();

                string action = cut ? "moved" : "copied";
                string extra = cut ? " The shape has been removed from its original slide - other shapes' indices there may have shifted too." : "";
                return new ToolResult
                {
                    Output = "Shape " + action + " to slide " + targetSlideIndex + " - new shape at shapeIndex " + newShapeIndex +
                             ". Other shapes' indices on the destination slide may have shifted." + extra +
                             " Re-read the affected slide(s) (read_slide) before addressing another shape by index in the same run.",
                    Mutated = true,
                    Summary = toolName,
                };
            }
            finally
            {
                RestoreClipboard(savedClipboard);
            }
        }

        private static ToolResult CopyElement(JsonElement input) { return CopyOrMoveElement(input, false); }
        private static ToolResult MoveElement(JsonElement input) { return CopyOrMoveElement(input, true); }
    }
}
