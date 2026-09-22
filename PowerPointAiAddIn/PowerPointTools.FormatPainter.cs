using System;
using System.Collections.Generic;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // Plain Shape-to-Shape native-value copies (no JSON parsing, no hex
        // round trip) - shared by copy_element_style below and by
        // PowerPointTools.CrossSlide.cs's reconstruction paths, which call
        // these for full fidelity after building the destination shape.

        // User-asked (2026-09-23): "isn't there a way to just copy all
        // formatting generically instead of hardcoding every property?" -
        // yes: Shape.PickUp()/Shape.Apply() is PowerPoint's own native format
        // painter (confirmed via reflection to exist on both Shape and
        // ShapeRange, no arguments) - it's an internal Application-level
        // "last picked up formatting" slot, NOT the Windows clipboard (a
        // different mechanism from Shape.Copy()/Shapes.Paste(), which this
        // codebase avoids for the documented clobbering/racing reason). This
        // is called FIRST, as a broad baseline pass, before any of the
        // targeted Copy*Formatting calls below - it should already cover
        // fill/line/shadow/3-D/adjustments generically (including things we
        // haven't individually hardcoded, like shape bevel), and the
        // targeted calls that follow then refine/guarantee the specific
        // properties this tool documents (most importantly per-RUN text
        // formatting, which a cursor-position-based format painter cannot
        // reproduce - PickUp/Apply on a whole shape samples one formatting
        // state, not each run individually). Kept as a supplement, not a
        // replacement, for exactly that reason.
        private static void CopyViaPickUpApply(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            try { source.PickUp(); target.Apply(); } catch { }
        }

        // Shared by shape-level bevel (Shape.ThreeD) and text-level bevel
        // (TextFrame2.ThreeD, confirmed via reflection to be the SAME
        // ThreeDFormat type) - both expose an identical member set. Contour/
        // extrusion color are wrapped separately since reading/writing them
        // can throw when no contour/extrusion is actually set on the source,
        // which shouldn't block copying the rest (bevel type/depth/inset,
        // rotation, lighting, material).
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

        // User-requested (2026-09-22) upgrade from the original "sample the
        // first character only" design to real per-run fidelity. There is no
        // bulk "list the formatting runs" API - confirmed via reflection,
        // TextRange.Runs(start,length) is an indexed accessor exactly like
        // Characters/Paragraphs, not an enumerable collection - so run
        // boundaries are found by scanning: read every character's Font,
        // compare to the previous one, and a difference in ANY tracked
        // property starts a new run. Paragraph Alignment is copied
        // per-paragraph (paragraph boundaries found by scanning the string
        // itself for '\r', PowerPoint's own paragraph separator when reading
        // TextRange.Text - avoids relying on Paragraphs(start,length)'s own
        // indexing semantics, which were never independently verified).
        // Replaying onto the SAME character offsets on the destination is
        // valid because the destination was built from this exact, unmodified
        // text string (see ReconstructTextBox/ReconstructAutoShape).
        //
        // Real cost, not hidden: this is O(text length) COM round trips (one
        // Font read per character, up to 9 properties each) to detect
        // boundaries, plus O(run count) writes. Not capped - a hard cap would
        // silently lose fidelity on long text, contradicting the point of
        // this change - but a very long text body will be proportionally
        // slower. Strikethrough still isn't copied (unchanged pre-existing
        // gap - no TextFrame2/TextRange2 on this PIA's Shape, same as
        // set_element_style's own documented limitation).
        private static void CopyTextFormatting(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            // Real-user-confirmed (2026-09-23): copying a TABLE crashed the
            // whole reconstruction with "The specified value is out of
            // range" - traced to THIS gate being unguarded. A table's outer
            // graphic-frame Shape doesn't carry text the way a text box/
            // autoshape does (text lives in the individual cells' own
            // shapes), and querying HasTextFrame/TextFrame on it can throw
            // instead of just reporting false - every other Copy* method in
            // this file already treats its own entry gate as best-effort;
            // this one didn't.
            try
            {
                if (source.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return;
                if (source.TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoTrue) return;
                if (target.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return;
            }
            catch { return; }

            // Text-frame-level layout (real-user-confirmed gap, 2026-09-23:
            // vertical anchor - top/middle/bottom - wasn't copied at all).
            // Confirmed via reflection: PowerPoint.TextFrame has plain
            // read/write VerticalAnchor/HorizontalAnchor/margins/WordWrap -
            // no TextFrame2 needed for these.
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

            // Run-level font formatting.
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

        // User-requested (2026-09-23): text outline and "text effects" -
        // NOT accessible via the classic Font object CopyTextFormatting
        // above uses (confirmed via reflection: no Line/Glow/Reflection/
        // Strikethrough members on PowerPoint.Font at all). They ARE
        // accessible via Shape.TextFrame2 -> Core.TextRange2 -> Core.Font2 -
        // confirmed via reflection this session, CORRECTING an earlier,
        // wrong claim elsewhere in this codebase (set_element_style's own
        // comment) that TextFrame2 doesn't exist on this PIA's Shape at all;
        // it does, just needs Microsoft.Office.Core's own TextRange2/Font2
        // types, not a PowerPoint-namespace TextRange2 (which is what a
        // previous, unqualified attempt actually failed to find).
        //
        // Deliberately sampled from the FIRST character only (uniform
        // across the whole shape), NOT per-run like CopyTextFormatting -
        // text outline/glow are rarely mixed within one text box in
        // practice, and giving this its own full per-run scan would double
        // the already-significant COM-call cost of CopyTextFormatting for a
        // much rarer case. Covers Line (outline), StrikeThrough/
        // DoubleStrikeThrough, Glow, Reflection, Shadow, SoftEdge, and Bevel/
        // 3-D (TextFrame2.ThreeD) - every text-level DrawingML effect
        // confirmed via reflection on Font2/TextFrame2. This is still a
        // manually-enumerated list, not a generic "copy everything" - Font2
        // has no bulk clone method - but CopyViaPickUpApply (see above)
        // provides the generic broad pass at the whole-SHAPE level; this
        // method's job is specifically the effects a shape-level format
        // painter pass can't guarantee at true per-run granularity.
        private static void CopyTextEffects(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            try
            {
                if (source.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return;
                if (source.TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoTrue) return;
                if (target.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue) return;

                // Text "bevel"/3-D (real-user-confirmed gap, 2026-09-23):
                // TextFrame2.ThreeD (confirmed via reflection to be the same
                // ThreeDFormat type as Shape.ThreeD) - a whole-text-frame
                // property, not per-character, so it's set here rather than
                // per-run below.
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

                try
                {
                    destFont2.Glow.Radius = srcFont2.Glow.Radius;
                    destFont2.Glow.Transparency = srcFont2.Glow.Transparency;
                    destFont2.Glow.Color.RGB = srcFont2.Glow.Color.RGB;
                }
                catch { }

                // Real-user-confirmed gap (2026-09-23): text reflection wasn't
                // copied at all - Font2.Reflection (Type/Size/Transparency/
                // Blur) confirmed via reflection. Offset is deliberately left
                // alone (its exact type wasn't independently verified).
                try
                {
                    destFont2.Reflection.Type = srcFont2.Reflection.Type;
                    destFont2.Reflection.Size = srcFont2.Reflection.Size;
                    destFont2.Reflection.Transparency = srcFont2.Reflection.Transparency;
                    destFont2.Reflection.Blur = srcFont2.Reflection.Blur;
                }
                catch { }

                // Text shadow (Font2.Shadow -> ShadowFormat, confirmed via
                // reflection - a full object with Visible/Type/Style/
                // ForeColor/Transparency/Blur/OffsetX/OffsetY/Size, NOT the
                // classic bool Font.Shadow) and soft edge (Font2.
                // SoftEdgeFormat, confirmed via reflection to be a plain
                // MsoSoftEdgeType ENUM property here - despite the name, not
                // an object with its own Radius/Type - so a direct value
                // copy is correct and complete, not a partial gap).
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
            // Same defensive gate as CopyTextFormatting's, and for the same
            // real-user-confirmed reason (2026-09-23) - a table's outer
            // graphic-frame Shape doesn't have a simple uniform Fill the way
            // a text box/autoshape does, and this was unguarded.
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

        // Real-user-confirmed (2026-09-23, live testing): the original version
        // of this method - which bootstrapped gradient mode by calling
        // destFill.TwoColorGradient(srcFill.GradientStyle, srcFill.
        // GradientVariant), i.e. replaying the SOURCE's own style/variant -
        // produced a visibly wrong gradient. Root cause: GradientStyle/
        // GradientVariant aren't safe to round-trip for every source fill -
        // a fill set up via PresetGradient, a theme gradient, or anything
        // other than an explicit TwoColorGradient call can report a
        // style/variant combination TwoColorGradient itself rejects or
        // reinterprets differently (msoGradientMixed being the most obvious
        // case, but not the only one). This version instead ALWAYS bootstraps
        // with a fixed, safe style/variant just to enter gradient mode and
        // get a real GradientStops collection to write into - the actual
        // visible result comes entirely from the explicit per-stop
        // color/position/transparency copy below, which doesn't depend on
        // the bootstrap style at all. GradientAngle is copied separately
        // afterward for the linear styles (horizontal/vertical/diagonal),
        // since that's the one visually-significant property the bootstrap
        // style choice would otherwise have baked in.
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

            // Real-user-confirmed (2026-09-23, live testing, SECOND round):
            // still wrong after the bootstrap-style fix above. Root cause:
            // the bootstrap leaves exactly 2 default stops, but a real-world
            // gradient fill - especially PowerPoint's own Shape Style
            // gallery gradients - very often has 3+ stops (a subtle sheen
            // effect), not 2. The previous version kept the 2 defaults and
            // only INSERT()-ed extras beyond index 2, on an unverified
            // assumption about Insert()'s index semantics. This version
            // instead first makes the dest stop COUNT match the source's
            // (deleting extras or inserting - confirmed via reflection that
            // GradientStops.Delete(index) exists alongside Insert) using
            // fresh, re-read indices/counts at every step rather than
            // precomputed ones, then does one final explicit pass setting
            // every stop's Color/Position/Transparency by matching index -
            // so even if an insert's initial values didn't stick exactly
            // right, the final pass corrects them.
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
            // Same defensive gate, same reason (2026-09-23 table crash) as
            // CopyTextFormatting/CopyFillFormatting above.
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

        // Rotation is a plain property; flip is NOT - Shape.HorizontalFlip/
        // VerticalFlip are get-only state (confirmed via reflection: no
        // set_HorizontalFlip/set_VerticalFlip exist on this PIA's Shape) -
        // actually flipping requires the Flip(MsoFlipCmd) TOGGLE method, so
        // this compares current state to the source's and calls Flip() only
        // when they differ, rather than (wrongly) assigning the property
        // directly.
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

        // Format painter. Source is addressed via the standard
        // slideIndex+shapeIndex (ResolveTopLevelShape, same as every other
        // single-shape PowerPoint tool) rather than a "sourceShapeIndex"
        // field, so it gets the existing nested-shape rejection for free.
        // Never touches position/size. Cross-slide (user-requested,
        // 2026-09-22): each target now carries its own slideIndex, not just a
        // shapeIndex resolved against the source's slide - a breaking schema
        // change from targetShapeIndexes:number[] to targets:{slideIndex,
        // shapeIndex}[], safe since this tool has never shipped on main.
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
                // Native format painter first (broad, generic pass covering
                // bevel/3-D and anything else not individually enumerated
                // below), THEN the targeted calls to guarantee per-run text
                // fidelity and the other specifics this tool documents - see
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
