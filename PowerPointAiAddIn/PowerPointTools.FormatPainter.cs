using System;
using System.Collections.Generic;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // Plain Shape-to-Shape native-value copies (no JSON parsing, no hex round
        // trip) powering copy_element_style below. See PowerPointTools.FormatPainter.cs.md
        // for why copy_element/move_element no longer share this code.

        // Shape.PickUp()/Shape.Apply() is PowerPoint's own native format painter (an
        // Application-level "last picked up formatting" slot, not the clipboard).
        // Called FIRST as a broad generic baseline, before the targeted Copy*Formatting
        // calls below refine/guarantee specifics (especially per-run text formatting,
        // which a whole-shape format painter pass can't reproduce). See .md.
        private static void CopyViaPickUpApply(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            try { source.PickUp(); target.Apply(); } catch { }
        }

        // Shared by shape-level bevel (Shape.ThreeD) and text-level bevel
        // (TextFrame2.ThreeD - the same ThreeDFormat type). Contour/extrusion color
        // are wrapped separately since they can throw when unset on the source,
        // which shouldn't block copying the rest.
        private static void CopyThreeD(PowerPoint.ThreeDFormat src, PowerPoint.ThreeDFormat dest)
        {
            try
            {
                if (src.Visible != Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    dest.Visible = src.Visible;
                    return;
                }
                dest.BevelTopType = src.BevelTopType;
                dest.BevelTopDepth = src.BevelTopDepth;
                dest.BevelTopInset = src.BevelTopInset;
                dest.BevelBottomType = src.BevelBottomType;
                dest.BevelBottomDepth = src.BevelBottomDepth;
                dest.BevelBottomInset = src.BevelBottomInset;
                dest.Depth = src.Depth;
                dest.PresetMaterial = src.PresetMaterial;
                dest.PresetLighting = src.PresetLighting;
                dest.PresetLightingDirection = src.PresetLightingDirection;
                dest.PresetLightingSoftness = src.PresetLightingSoftness;
                dest.RotationX = src.RotationX;
                dest.RotationY = src.RotationY;
                dest.RotationZ = src.RotationZ;
                dest.Perspective = src.Perspective;
                dest.FieldOfView = src.FieldOfView;
                dest.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            }
            catch { }
            try { dest.ContourWidth = src.ContourWidth; dest.ContourColor.RGB = src.ContourColor.RGB; } catch { }
            try { dest.ExtrusionColorType = src.ExtrusionColorType; dest.ExtrusionColor.RGB = src.ExtrusionColor.RGB; } catch { }
        }

        private struct FontSnapshot
        {
            public Microsoft.Office.Core.MsoTriState Bold, Italic, Underline, Shadow, Superscript, Subscript;
            public float Size;
            public int Color;
            public string Name;

            public bool Matches(FontSnapshot other)
            {
                return Bold == other.Bold && Italic == other.Italic && Underline == other.Underline &&
                       Shadow == other.Shadow && Superscript == other.Superscript && Subscript == other.Subscript &&
                       Size == other.Size && Color == other.Color && Name == other.Name;
            }
        }

        private static FontSnapshot ReadFontSnapshot(PowerPoint.Font f)
        {
            return new FontSnapshot
            {
                Bold = f.Bold, Italic = f.Italic, Underline = f.Underline, Shadow = f.Shadow,
                Superscript = f.Superscript, Subscript = f.Subscript,
                Size = f.Size, Color = f.Color.RGB, Name = f.Name,
            };
        }

        private static void ApplyFontSnapshot(PowerPoint.Font f, FontSnapshot snap)
        {
            f.Bold = snap.Bold; f.Italic = snap.Italic; f.Underline = snap.Underline; f.Shadow = snap.Shadow;
            f.Superscript = snap.Superscript; f.Subscript = snap.Subscript;
            f.Size = snap.Size; f.Color.RGB = snap.Color; f.Name = snap.Name;
        }

        // Per-run text formatting fidelity: since there is no bulk "list the
        // formatting runs" API, run boundaries are found by scanning each
        // character's Font and starting a new run on any tracked-property
        // difference. Paragraph alignment is copied per-paragraph (boundaries found
        // by scanning for '\r'). Best-effort against a target of different length;
        // O(text length) COM round trips, not capped. See PowerPointTools.FormatPainter.cs.md.
        private static void CopyTextFormatting(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            // A table's outer graphic-frame Shape doesn't carry text the way a text
            // box/autoshape does, and querying HasTextFrame/TextFrame on it can throw
            // instead of reporting false - this gate must be guarded. See .md.
            try
            {
                if (source.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return;
                if (source.TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoTrue) return;
                if (target.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return;
            }
            catch { return; }

            // Text-frame-level layout (vertical anchor, etc.) - plain TextFrame
            // properties, no TextFrame2 needed. See .md for the gap this closed.
            try { target.TextFrame.VerticalAnchor = source.TextFrame.VerticalAnchor; } catch { }
            try { target.TextFrame.HorizontalAnchor = source.TextFrame.HorizontalAnchor; } catch { }
            try { target.TextFrame.WordWrap = source.TextFrame.WordWrap; } catch { }
            try
            {
                target.TextFrame.MarginLeft = source.TextFrame.MarginLeft;
                target.TextFrame.MarginRight = source.TextFrame.MarginRight;
                target.TextFrame.MarginTop = source.TextFrame.MarginTop;
                target.TextFrame.MarginBottom = source.TextFrame.MarginBottom;
            }
            catch { }

            PowerPoint.TextRange srcText = source.TextFrame.TextRange;
            PowerPoint.TextRange targetText = target.TextFrame.TextRange;
            string text = srcText.Text;
            int len = text.Length;
            if (len == 0) return;

            // Paragraph alignment, per paragraph (0-based start/length pairs,
            // each including its own trailing '\r' if present).
            var paragraphs = new List<KeyValuePair<int, int>>();
            int segStart = 0;
            for (int i = 0; i < len; i++)
            {
                if (text[i] == '\r')
                {
                    paragraphs.Add(new KeyValuePair<int, int>(segStart, i - segStart + 1));
                    segStart = i + 1;
                }
            }
            if (segStart < len) paragraphs.Add(new KeyValuePair<int, int>(segStart, len - segStart));
            foreach (var para in paragraphs)
            {
                try
                {
                    var align = srcText.Characters(para.Key + 1, para.Value).ParagraphFormat.Alignment;
                    targetText.Characters(para.Key + 1, para.Value).ParagraphFormat.Alignment = align;
                }
                catch { }
            }

            FontSnapshot? current = null;
            int runStart = 0;
            for (int i = 0; i < len; i++)
            {
                FontSnapshot snap;
                try { snap = ReadFontSnapshot(srcText.Characters(i + 1, 1).Font); }
                catch { continue; }

                if (current.HasValue && !current.Value.Matches(snap))
                {
                    ApplyFontRun(targetText, runStart, i - runStart, current.Value);
                    runStart = i;
                }
                current = snap;
            }
            if (current.HasValue) ApplyFontRun(targetText, runStart, len - runStart, current.Value);
        }

        private static void ApplyFontRun(PowerPoint.TextRange targetText, int start0, int length, FontSnapshot snap)
        {
            if (length <= 0) return;
            try { ApplyFontSnapshot(targetText.Characters(start0 + 1, length).Font, snap); }
            catch { }
        }

        // Text outline/glow/reflection/shadow/strikethrough/bevel - not accessible via
        // the classic Font object, only via Shape.TextFrame2 -> Core.TextRange2 ->
        // Core.Font2 (this PIA's Shape DOES have TextFrame2, contrary to an earlier,
        // wrong claim elsewhere in this codebase - see .md). Sampled from the FIRST
        // character only, not per-run like CopyTextFormatting, since these effects
        // are rarely mixed within one text box and a full per-run scan here would
        // double an already-significant COM cost. See PowerPointTools.FormatPainter.cs.md.
        private static void CopyTextEffects(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            try
            {
                if (source.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return;
                if (source.TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoTrue) return;
                if (target.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return;

                // TextFrame2.ThreeD is a whole-text-frame property, not per-character,
                // so it's set here rather than per-run below.
                try { CopyThreeD(source.TextFrame2.ThreeD, target.TextFrame2.ThreeD); } catch { }

                Microsoft.Office.Core.TextRange2 srcRange2 = source.TextFrame2.TextRange;
                Microsoft.Office.Core.TextRange2 destRange2 = target.TextFrame2.TextRange;
                if (srcRange2.Length == 0) return;

                Microsoft.Office.Core.Font2 srcFont2 = srcRange2.Characters[1, 1].Font;
                Microsoft.Office.Core.Font2 destFont2 = destRange2.Font;

                try { destFont2.StrikeThrough = srcFont2.StrikeThrough; } catch { }
                try { destFont2.DoubleStrikeThrough = srcFont2.DoubleStrikeThrough; } catch { }

                try
                {
                    if (srcFont2.Line.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        destFont2.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                        destFont2.Line.ForeColor.RGB = srcFont2.Line.ForeColor.RGB;
                        destFont2.Line.Weight = srcFont2.Line.Weight;
                    }
                    else
                    {
                        destFont2.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                    }
                }
                catch { }

                // GlowFormat has no Visible gate - a zero Radius IS the "no glow" state,
                // so only touch destFont2.Glow when the source has a real (>0) radius.
                try
                {
                    if (srcFont2.Glow.Radius > 0)
                    {
                        destFont2.Glow.Radius = srcFont2.Glow.Radius;
                        destFont2.Glow.Transparency = srcFont2.Glow.Transparency;
                        destFont2.Glow.Color.RGB = srcFont2.Glow.Color.RGB;
                    }
                }
                catch { }

                // Like Glow above, ReflectionFormat has no Visible gate -
                // msoReflectionTypeNone is the only off switch, so only write
                // Type/Size/Transparency/Blur when the source's Type isn't None
                // (writing unconditionally used to materialize a reflection that
                // shouldn't exist). Offset is deliberately left alone. See .md.
                try
                {
                    Microsoft.Office.Core.MsoReflectionType srcReflType = srcFont2.Reflection.Type;
                    if (srcReflType != Microsoft.Office.Core.MsoReflectionType.msoReflectionTypeNone)
                    {
                        destFont2.Reflection.Type = srcReflType;
                        destFont2.Reflection.Size = srcFont2.Reflection.Size;
                        destFont2.Reflection.Transparency = srcFont2.Reflection.Transparency;
                        destFont2.Reflection.Blur = srcFont2.Reflection.Blur;
                    }
                }
                catch { }

                // Font2.Shadow is a full ShadowFormat object (Visible/Type/Style/
                // ForeColor/Transparency/Blur/OffsetX/OffsetY/Size), NOT the classic
                // bool Font.Shadow. Font2.SoftEdgeFormat below is a plain MsoSoftEdgeType
                // enum despite the name (not an object), so a direct value copy suffices.
                try
                {
                    if (srcFont2.Shadow.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        destFont2.Shadow.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                        destFont2.Shadow.Type = srcFont2.Shadow.Type;
                        destFont2.Shadow.Style = srcFont2.Shadow.Style;
                        destFont2.Shadow.ForeColor.RGB = srcFont2.Shadow.ForeColor.RGB;
                        destFont2.Shadow.Transparency = srcFont2.Shadow.Transparency;
                        destFont2.Shadow.Blur = srcFont2.Shadow.Blur;
                        destFont2.Shadow.OffsetX = srcFont2.Shadow.OffsetX;
                        destFont2.Shadow.OffsetY = srcFont2.Shadow.OffsetY;
                        destFont2.Shadow.Size = srcFont2.Shadow.Size;
                    }
                    else
                    {
                        destFont2.Shadow.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                    }
                }
                catch { }
                try { destFont2.SoftEdgeFormat = srcFont2.SoftEdgeFormat; } catch { }
            }
            catch { /* best-effort - TextFrame2 access itself is the least-tested part of this change */ }
        }

        // Solid (as before) plus gradient/patterned - branches on Fill.Type.
        // Picture/textured fills only get Visible/Transparency: the actual
        // image/texture content would need the same export step picture
        // SHAPES do (not implemented) - not silently pretended to be handled.
        private static void CopyFillFormatting(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            // Same defensive gate as CopyTextFormatting's, same reason (table's outer
            // graphic-frame Shape has no simple uniform Fill) - see that method's .md entry.
            try
            {
                target.Fill.Visible = source.Fill.Visible;
                if (source.Fill.Visible != Microsoft.Office.Core.MsoTriState.msoTrue) return;
            }
            catch { return; }

            try { target.Fill.Transparency = source.Fill.Transparency; } catch { }

            Microsoft.Office.Core.MsoFillType fillType;
            try { fillType = source.Fill.Type; }
            catch { fillType = Microsoft.Office.Core.MsoFillType.msoFillSolid; }

            switch (fillType)
            {
                case Microsoft.Office.Core.MsoFillType.msoFillGradient:
                    CopyGradientFill(source.Fill, target.Fill);
                    break;
                case Microsoft.Office.Core.MsoFillType.msoFillPatterned:
                    try
                    {
                        target.Fill.Patterned(source.Fill.Pattern);
                        target.Fill.ForeColor.RGB = source.Fill.ForeColor.RGB;
                        target.Fill.BackColor.RGB = source.Fill.BackColor.RGB;
                    }
                    catch { }
                    break;
                case Microsoft.Office.Core.MsoFillType.msoFillPicture:
                case Microsoft.Office.Core.MsoFillType.msoFillTextured:
                    break;
                default:
                    try { target.Fill.ForeColor.RGB = source.Fill.ForeColor.RGB; } catch { }
                    break;
            }
        }

        // Always bootstraps gradient mode with a fixed, safe style/variant (not the
        // source's own GradientStyle/GradientVariant, which aren't safe to round-trip
        // for every fill) just to get a real GradientStops collection to write into -
        // the visible result comes entirely from the explicit per-stop copy below.
        // GradientAngle is copied separately for the linear styles. See
        // PowerPointTools.FormatPainter.cs.md for the bug this fixed.
        private static void CopyGradientFill(PowerPoint.FillFormat srcFill, PowerPoint.FillFormat destFill)
        {
            Microsoft.Office.Core.GradientStops srcStops;
            int srcCount;
            try
            {
                srcStops = srcFill.GradientStops;
                srcCount = srcStops.Count;
            }
            catch { return; }
            if (srcCount == 0) return;

            try { destFill.TwoColorGradient(Microsoft.Office.Core.MsoGradientStyle.msoGradientHorizontal, 1); }
            catch { return; }

            try
            {
                Microsoft.Office.Core.MsoGradientStyle srcStyle = srcFill.GradientStyle;
                if (srcStyle == Microsoft.Office.Core.MsoGradientStyle.msoGradientHorizontal ||
                    srcStyle == Microsoft.Office.Core.MsoGradientStyle.msoGradientVertical ||
                    srcStyle == Microsoft.Office.Core.MsoGradientStyle.msoGradientDiagonalUp ||
                    srcStyle == Microsoft.Office.Core.MsoGradientStyle.msoGradientDiagonalDown)
                {
                    destFill.GradientAngle = srcFill.GradientAngle;
                }
            }
            catch { }

            // The bootstrap above leaves exactly 2 default stops, but a real-world
            // gradient (especially PowerPoint's own Shape Style gallery gradients)
            // often has 3+. First make the dest stop COUNT match the source's
            // (re-reading indices/counts at every step), then do one final explicit
            // pass setting every stop's Color/Position/Transparency by index so
            // insert-time values are corrected regardless. See .md.
            Microsoft.Office.Core.GradientStops destStops = destFill.GradientStops;
            int destCount;
            try { destCount = destStops.Count; } catch { destCount = 0; }

            while (destCount > srcCount && destCount > 1)
            {
                try { destFill.GradientStops.Delete(destCount); destCount--; }
                catch { break; }
            }
            while (destCount < srcCount)
            {
                try
                {
                    Microsoft.Office.Core.GradientStop s = srcStops[destCount + 1];
                    destFill.GradientStops.Insert(s.Color.RGB, s.Position, s.Transparency, destCount + 1);
                    destCount++;
                }
                catch { break; }
            }

            int finalCount;
            try { finalCount = destFill.GradientStops.Count; } catch { finalCount = 0; }
            int shared = Math.Min(srcCount, finalCount);
            for (int i = 1; i <= shared; i++)
            {
                try
                {
                    Microsoft.Office.Core.GradientStop s = srcStops[i];
                    Microsoft.Office.Core.GradientStop d = destFill.GradientStops[i];
                    d.Color.RGB = s.Color.RGB;
                    d.Position = s.Position;
                    d.Transparency = s.Transparency;
                }
                catch { }
            }
        }

        // Existing (visible/color/weight) plus the rest of LineFormat - dash
        // style, line style, transparency, both arrowheads. Meaningful mainly
        // for lines but harmless to attempt on any shape's outline.
        private static void CopyStrokeFormatting(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            // Same defensive gate, same reason, as CopyTextFormatting/CopyFillFormatting above.
            try
            {
                target.Line.Visible = source.Line.Visible;
                if (source.Line.Visible != Microsoft.Office.Core.MsoTriState.msoTrue) return;
            }
            catch { return; }

            try
            {
                target.Line.ForeColor.RGB = source.Line.ForeColor.RGB;
                target.Line.Weight = source.Line.Weight;
            }
            catch { return; }
            try { target.Line.BackColor.RGB = source.Line.BackColor.RGB; } catch { }
            try { target.Line.Transparency = source.Line.Transparency; } catch { }
            try { target.Line.DashStyle = source.Line.DashStyle; } catch { }
            try { target.Line.Style = source.Line.Style; } catch { }
            try { target.Line.Pattern = source.Line.Pattern; } catch { }
            try
            {
                target.Line.BeginArrowheadStyle = source.Line.BeginArrowheadStyle;
                target.Line.BeginArrowheadLength = source.Line.BeginArrowheadLength;
                target.Line.BeginArrowheadWidth = source.Line.BeginArrowheadWidth;
                target.Line.EndArrowheadStyle = source.Line.EndArrowheadStyle;
                target.Line.EndArrowheadLength = source.Line.EndArrowheadLength;
                target.Line.EndArrowheadWidth = source.Line.EndArrowheadWidth;
            }
            catch { }
        }

        // Rotation is a plain property; flip is NOT - Shape.HorizontalFlip/VerticalFlip
        // are get-only, so flipping requires the Flip(MsoFlipCmd) TOGGLE method. This
        // compares current state to the source's and calls Flip() only when they differ.
        private static void CopyRotationAndFlip(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            try { target.Rotation = source.Rotation; } catch { }
            try
            {
                if (source.HorizontalFlip != target.HorizontalFlip) target.Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipHorizontal);
                if (source.VerticalFlip != target.VerticalFlip) target.Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipVertical);
            }
            catch { }
        }

        // AutoShape "adjustment handles" (e.g. a rounded rectangle's corner
        // radius). Adjustments.Count is 0 for shape kinds without any -
        // harmless no-op for table/chart/SmartArt/line/etc.
        private static void CopyShapeAdjustments(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            try
            {
                int n = Math.Min(source.Adjustments.Count, target.Adjustments.Count);
                for (int i = 1; i <= n; i++)
                    target.Adjustments[i] = source.Adjustments[i];
            }
            catch { }
        }

        // Format painter: copies style from one source shape to one or more targets
        // (each its own {slideIndex, shapeIndex}, so targets can be cross-slide).
        // Never touches position/size. See PowerPointTools.FormatPainter.cs.md.
        private static ToolResult CopyElementStyle(JsonElement input)
        {
            PowerPoint.Shape source = ResolveTopLevelShape(input, "copy_element_style");
            PowerPoint.Slides slides = ActivePresentation.Slides;

            if (!input.TryGetProperty("targets", out var targetsArr) || targetsArr.ValueKind != JsonValueKind.Array)
                throw new ArgumentException("copy_element_style: targets must be an array of at least one {slideIndex, shapeIndex}.");

            var targets = new List<PowerPoint.Shape>();
            var seen = new HashSet<string>();
            foreach (JsonElement el in targetsArr.EnumerateArray())
            {
                if (!el.TryGetProperty("slideIndex", out var siEl) || !el.TryGetProperty("shapeIndex", out var shEl))
                    throw new ArgumentException("copy_element_style: each target needs slideIndex and shapeIndex.");
                int si = siEl.GetInt32();
                int sh = shEl.GetInt32();
                if (si < 0 || si >= slides.Count)
                    throw new ArgumentException("copy_element_style: target slideIndex " + si + " is out of range.");
                PowerPoint.Slide tSlide = slides[si + 1];
                if (sh < 0 || sh >= tSlide.Shapes.Count)
                    throw new ArgumentException("copy_element_style: target shapeIndex " + sh + " is out of range on slide " + si + " (has " + tSlide.Shapes.Count + " shapes).");
                if (seen.Add(si + ":" + sh)) targets.Add(tSlide.Shapes[sh + 1]); // permissive dedup, matching group_element
            }
            if (targets.Count == 0)
                throw new ArgumentException("copy_element_style: targets must contain at least one entry.");

            bool hasSourceText = source.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                              && source.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue;

            foreach (PowerPoint.Shape target in targets)
            {
                // Native format painter first (broad generic pass), THEN the targeted
                // calls for per-run text fidelity and other specifics - see
                // CopyViaPickUpApply's own comment for why both are used.
                CopyViaPickUpApply(source, target);
                if (hasSourceText)
                {
                    CopyTextFormatting(source, target);
                    CopyTextEffects(source, target);
                }
                CopyFillFormatting(source, target);
                CopyStrokeFormatting(source, target);
                CopyRotationAndFlip(source, target);
                CopyShapeAdjustments(source, target);
            }

            string what = hasSourceText ? "text (per-run), text outline/strikethrough/glow/reflection/bevel, text anchor/margins, fill, stroke, rotation, and any other native shape formatting (via format painter)" : "fill, stroke, rotation, and any other native shape formatting (via format painter) - source has no text";
            return new ToolResult
            {
                Output = "Style copied to " + targets.Count + " shape(s): " + what + ".",
                Mutated = true,
                Summary = "copy_element_style",
            };
        }
    }
}
