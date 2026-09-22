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
        // Confirmed via .NET reflection against the real referenced PIA
        // (Microsoft.Office.Interop.PowerPoint.dll, GAC 15.0.0.0), same
        // discipline as PP-24's enum maps: _Presentation.SlideMaster returns
        // Master directly - no fallback needed. Known limitation, same class
        // as ActivePresentation itself (PP-1): a deck combining more than one
        // theme has more than one Slide Master, and only this default
        // (first) one is targeted.
        private static PowerPoint.Master ResolveSlideMaster()
        {
            return ActivePresentation.SlideMaster;
        }

        // User-requested (2026-09-22, live testing): add_master_element/
        // read_master_elements/remove_master_element/set_master_element_
        // transform previously only ever targeted the single top-level
        // Slide Master - there was no way to address one specific layout
        // (Title Slide, Title and Content, etc.), even though each layout
        // has its own independent shape collection (the exact thing that
        // caused the footer-text-not-inheriting bug). optional layoutName
        // resolves by case-insensitive substring against this presentation's
        // own layout names - same resolve-by-substring pattern
        // ResolveCustomLayout (PowerPointTools.LayoutAnim.cs) already uses
        // for set_slide_layout - and returns that layout's own Shapes
        // collection instead of the master's when given.
        private struct MasterTarget
        {
            public PowerPoint.Shapes Shapes;
            public string Label;
        }

        private static PowerPoint.CustomLayout ResolveLayoutByName(string query, string toolName)
        {
            PowerPoint.Master master = ResolveSlideMaster();
            PowerPoint.CustomLayout firstMatch = null;
            var namesSeen = new List<string>();
            foreach (PowerPoint.CustomLayout layout in master.CustomLayouts)
            {
                namesSeen.Add(layout.Name);
                if (firstMatch == null && layout.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) firstMatch = layout;
            }
            if (firstMatch != null) return firstMatch;
            throw new ArgumentException(toolName + ": no layout matching '" + query + "' found in this presentation's theme. Available: " + string.Join(", ", namesSeen) + ".");
        }

        // User-requested (2026-09-22): before this, the only way to
        // discover a deck's real layout names was to deliberately pass a
        // bad layoutName and read the error message's "Available: ..."
        // list. Read-only, lists every layout in this presentation's
        // default Slide Master's theme, plus how many slides currently use
        // each one (compared by CustomLayout.Index, not Name, in case two
        // layouts ever share a name).
        private static ToolResult ListLayouts(JsonElement input)
        {
            PowerPoint.Master master = ResolveSlideMaster();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            var sb = new StringBuilder();
            int i = 0;
            foreach (PowerPoint.CustomLayout layout in master.CustomLayouts)
            {
                int usedByCount = 0;
                foreach (PowerPoint.Slide s in slides)
                {
                    try { if (s.CustomLayout.Index == layout.Index) usedByCount++; } catch { }
                }
                sb.AppendLine("[" + i + "] \"" + layout.Name + "\"" + (usedByCount > 0 ? " - used by " + usedByCount + " slide(s)" : ""));
                i++;
            }
            if (i == 0)
                return new ToolResult { Output = "No layouts found on this presentation's default Slide Master.", Summary = "list_layouts" };
            return new ToolResult { Output = "This presentation's Slide Master has " + i + " layout(s):\n" + sb.ToString().TrimEnd(), Summary = "list_layouts" };
        }

        private static string CapitalizeFirst(string s)
        {
            return string.IsNullOrEmpty(s) ? s : char.ToUpperInvariant(s[0]) + s.Substring(1);
        }

        private static MasterTarget ResolveMasterTarget(JsonElement input, string toolName)
        {
            if (input.TryGetProperty("layoutName", out var layoutEl) && layoutEl.ValueKind == JsonValueKind.String)
            {
                string query = layoutEl.GetString();
                if (!string.IsNullOrEmpty(query))
                {
                    PowerPoint.CustomLayout layout = ResolveLayoutByName(query, toolName);
                    return new MasterTarget { Shapes = layout.Shapes, Label = "layout \"" + layout.Name + "\"" };
                }
            }
            return new MasterTarget { Shapes = ResolveSlideMaster().Shapes, Label = "the Slide Master" };
        }

        // Slide.HeadersFooters.DateAndTime/.Footer/.SlideNumber are all the
        // same HeaderFooter type. DisplayOnTitleSlide is NOT among them here -
        // real-Office-confirmed (2026-09-21, live error): setting it via a
        // SLIDE's HeadersFooters throws "HeadersFooters (unknown member) :
        // Invalid request. This property must be set using slide or title
        // master." even though reflection shows the member on the shared
        // HeadersFooters type with no hint of that restriction - a runtime
        // COM restriction reflection can't see, only member existence/type.
        // See ApplySkipTitleSlide below, which sets it on the MASTER instead.
        //
        // Takes a HeadersFooters directly (not a Slide) so the SAME logic
        // applies to a slide's OWN HeadersFooters and to the Slide/Title
        // Master's (Master.HeadersFooters is the identical COM type) -
        // real-user-confirmed (2026-09-22): a slide added AFTER
        // set_headers_footers ran did not inherit the deck-wide toggle,
        // because only existing slides were ever touched. Now
        // SetHeadersFooters also applies to the master(s) on a deck-wide
        // call, on the hypothesis that a brand-new slide's initial
        // HeadersFooters state is seeded from its master at creation time -
        // unverified without a live retest, but harmless either way (the
        // master's own header/footer placeholders, if the theme has them,
        // are otherwise never addressed by this tool at all).
        private static void ApplyHeadersFooters(
            PowerPoint.HeadersFooters hf, bool? slideNumberVisible,
            bool? footerVisible, string footerText,
            bool? dateVisible, string dateMode, string dateText)
        {
            if (slideNumberVisible.HasValue)
                hf.SlideNumber.Visible = slideNumberVisible.Value ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            if (footerVisible.HasValue)
                hf.Footer.Visible = footerVisible.Value ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            if (footerText != null)
                hf.Footer.Text = footerText;
            if (dateVisible.HasValue)
                hf.DateAndTime.Visible = dateVisible.Value ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            if (dateMode == "auto")
            {
                hf.DateAndTime.UseFormat = Microsoft.Office.Core.MsoTriState.msoTrue;
                hf.DateAndTime.Format = PowerPoint.PpDateTimeFormat.ppDateTimeMdyy;
            }
            else if (dateMode == "fixed")
            {
                hf.DateAndTime.UseFormat = Microsoft.Office.Core.MsoTriState.msoFalse;
                hf.DateAndTime.Text = dateText;
            }
        }

        // DisplayOnTitleSlide (the "Don't show on title slide" checkbox) is
        // deck-wide by nature, same as PageSetup.FirstSlideNumber - it lives
        // on the Slide Master's (and, if the deck has one, the Title
        // Master's - a deck can carry a second, separate master used only by
        // title-layout slides) HeadersFooters, not any individual slide's.
        private static void ApplySkipTitleSlide(bool skipTitleSlide)
        {
            PowerPoint.Presentation pres = ActivePresentation;
            Microsoft.Office.Core.MsoTriState value = skipTitleSlide ? Microsoft.Office.Core.MsoTriState.msoFalse : Microsoft.Office.Core.MsoTriState.msoTrue;
            pres.SlideMaster.HeadersFooters.DisplayOnTitleSlide = value;
            // Real-user-confirmed (2026-09-22): even READING pres.TitleMaster
            // (not just writing to it) can throw "Master (unknown member) :
            // Invalid request" in some deck states despite HasTitleMaster
            // reporting true first - best-effort, never let this abort the
            // rest of set_headers_footers.
            try
            {
                if (pres.HasTitleMaster == Microsoft.Office.Core.MsoTriState.msoTrue)
                    pres.TitleMaster.HeadersFooters.DisplayOnTitleSlide = value;
            }
            catch (Exception ex) { DebugLog.WriteException("ApplySkipTitleSlide TitleMaster", ex); }
        }

        // Shared by ReadSlide (PowerPointTools.Read.cs) so the model can see
        // what set_headers_footers already did, same "don't leave the model
        // blind to state it just set" precedent as PP-24's layout/transition/
        // animation-count line.
        private static string DescribeHeadersFooters(PowerPoint.Slide slide)
        {
            PowerPoint.HeadersFooters hf = slide.HeadersFooters;
            var parts = new List<string>();
            parts.Add("slide number=" + (hf.SlideNumber.Visible == Microsoft.Office.Core.MsoTriState.msoTrue ? "on" : "off"));
            parts.Add("footer=" + (hf.Footer.Visible == Microsoft.Office.Core.MsoTriState.msoTrue ? "on (\"" + hf.Footer.Text + "\")" : "off"));
            bool dateOn = hf.DateAndTime.Visible == Microsoft.Office.Core.MsoTriState.msoTrue;
            parts.Add("date=" + (dateOn ? (hf.DateAndTime.UseFormat == Microsoft.Office.Core.MsoTriState.msoTrue ? "on (auto)" : "on (\"" + hf.DateAndTime.Text + "\")") : "off"));
            return "Headers/footers: " + string.Join(", ", parts) + ".";
        }

        private static ToolResult SetHeadersFooters(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Presentation pres = ActivePresentation;
            if (slideIndex != -1 && (slideIndex < 0 || slideIndex >= pres.Slides.Count))
                throw new ArgumentOutOfRangeException("slideIndex", "slideIndex must be -1 (whole deck) or between 0 and " + (pres.Slides.Count - 1) + ".");

            bool? slideNumberVisible = input.TryGetProperty("slideNumberVisible", out var snEl) && (snEl.ValueKind == JsonValueKind.True || snEl.ValueKind == JsonValueKind.False)
                ? (bool?)(snEl.ValueKind == JsonValueKind.True) : null;

            bool? footerVisible = input.TryGetProperty("footerVisible", out var fvEl) && (fvEl.ValueKind == JsonValueKind.True || fvEl.ValueKind == JsonValueKind.False)
                ? (bool?)(fvEl.ValueKind == JsonValueKind.True) : null;
            string footerText = input.TryGetProperty("footerText", out var ftEl) && ftEl.ValueKind == JsonValueKind.String ? ftEl.GetString() : null;
            if (footerText != null && footerVisible == null) footerVisible = true;

            bool? dateVisible = input.TryGetProperty("dateVisible", out var dvEl) && (dvEl.ValueKind == JsonValueKind.True || dvEl.ValueKind == JsonValueKind.False)
                ? (bool?)(dvEl.ValueKind == JsonValueKind.True) : null;
            string dateMode = input.TryGetProperty("dateMode", out var dmEl) && dmEl.ValueKind == JsonValueKind.String ? dmEl.GetString() : null;
            string dateText = input.TryGetProperty("dateText", out var dtEl) && dtEl.ValueKind == JsonValueKind.String ? dtEl.GetString() : null;
            if (dateMode != null && dateMode != "auto" && dateMode != "fixed")
                throw new ArgumentException("set_headers_footers: unknown dateMode '" + dateMode + "'. Valid: auto, fixed.");
            if (dateMode == "fixed" && dateText == null)
                throw new ArgumentException("set_headers_footers: dateMode 'fixed' requires dateText.");
            if (dateMode != null && dateVisible == null) dateVisible = true;

            bool? skipTitleSlide = input.TryGetProperty("skipTitleSlide", out var stEl) && (stEl.ValueKind == JsonValueKind.True || stEl.ValueKind == JsonValueKind.False)
                ? (bool?)(stEl.ValueKind == JsonValueKind.True) : null;

            int? startNumber = input.TryGetProperty("startNumber", out var snoEl) && snoEl.ValueKind == JsonValueKind.Number ? (int?)snoEl.GetInt32() : null;

            var applied = new List<string>();
            if (slideNumberVisible.HasValue) applied.Add("slideNumberVisible");
            if (footerVisible.HasValue) applied.Add("footerVisible");
            if (footerText != null) applied.Add("footerText");
            if (dateVisible.HasValue) applied.Add("dateVisible");
            if (dateMode != null) applied.Add("dateMode");
            if (skipTitleSlide.HasValue) applied.Add("skipTitleSlide");
            if (startNumber.HasValue) applied.Add("startNumber");

            if (applied.Count == 0)
                return new ToolResult { Output = "set_headers_footers: no recognized fields were provided - nothing changed.", IsError = true, Summary = "set_headers_footers" };

            bool anyPerSlideField = slideNumberVisible.HasValue || footerVisible.HasValue || footerText != null || dateVisible.HasValue || dateMode != null;
            int layoutFailures = 0;
            if (anyPerSlideField)
            {
                if (slideIndex == -1)
                {
                    // Real-user-confirmed (2026-09-22, fourth round): this
                    // per-SLIDE loop was the one place left unguarded - a
                    // dateVisible:false + dateMode:"auto" combination (never
                    // exercised before; every prior test used dateVisible:
                    // true) threw "HeaderFooter (unknown member)" here,
                    // aborting the whole call before even reaching the
                    // master/layout seeding below, so none of THEIR failure-
                    // counting ever got a chance to run. Same per-item
                    // best-effort treatment as the master/layout loops now.
                    int slideFailures = 0;
                    foreach (PowerPoint.Slide s in pres.Slides)
                    {
                        try { ApplyHeadersFooters(s.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText); }
                        catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters slide " + (s.SlideIndex - 1), ex); slideFailures++; }
                    }
                    layoutFailures += slideFailures;

                    // Also seed the master(s) AND every custom layout, so a
                    // slide added AFTER this call starts with the same
                    // state, instead of only ever touching slides that
                    // existed at call time. Real-user-confirmed (2026-09-22):
                    // seeding only the Master was NOT enough - slideNumber/
                    // date still showed on a new slide (this theme's layouts
                    // already defaulted those on), but footerText did not,
                    // because CustomLayout.HeadersFooters is confirmed (via
                    // reflection) to be a SEPARATE object from the Master's,
                    // not something a layout reads through to the master for -
                    // a new slide inherits from whichever layout it's based
                    // on, not from the top-level Master directly.
                    //
                    // Real-user-confirmed (2026-09-22, second round): SOME
                    // layouts throw "HeaderFooter (unknown member) : Invalid
                    // request" on at least one of these property writes -
                    // exact cause unconfirmed (a "Blank"-style layout with no
                    // footer/date/number placeholder at all is the leading
                    // guess, mirroring DisplayOnTitleSlide's own confirmed
                    // must-be-set-on-master-or-slide restriction), but the
                    // whole deck-wide call must not abort partway through
                    // over one uncooperative layout. Best-effort per layout;
                    // failures are counted, not silently hidden, and don't
                    // stop the rest.
                    // Every one of the calls below is wrapped individually -
                    // real-user-confirmed (2026-09-22, third round): even
                    // pres.SlideMaster.HeadersFooters / pres.TitleMaster
                    // itself (not just a per-layout write) can throw
                    // "Master (unknown member) : Invalid request" - exact
                    // trigger still unconfirmed, so nothing here is trusted
                    // to succeed unguarded any more.
                    try { ApplyHeadersFooters(pres.SlideMaster.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText); }
                    catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters SlideMaster", ex); layoutFailures++; }

                    try
                    {
                        foreach (PowerPoint.CustomLayout layout in pres.SlideMaster.CustomLayouts)
                        {
                            try { ApplyHeadersFooters(layout.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText); }
                            catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters layout '" + layout.Name + "'", ex); layoutFailures++; }
                        }
                    }
                    catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters SlideMaster.CustomLayouts enumeration", ex); layoutFailures++; }

                    try
                    {
                        if (pres.HasTitleMaster == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            PowerPoint.Master titleMaster = pres.TitleMaster;
                            try { ApplyHeadersFooters(titleMaster.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText); }
                            catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters TitleMaster", ex); layoutFailures++; }
                            foreach (PowerPoint.CustomLayout layout in titleMaster.CustomLayouts)
                            {
                                try { ApplyHeadersFooters(layout.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText); }
                                catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters title-master layout '" + layout.Name + "'", ex); layoutFailures++; }
                            }
                        }
                    }
                    catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters TitleMaster access", ex); layoutFailures++; }
                }
                else
                {
                    try { ApplyHeadersFooters(pres.Slides[slideIndex + 1].HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText); }
                    catch (Exception ex)
                    {
                        DebugLog.WriteException("SetHeadersFooters slide " + slideIndex, ex);
                        return new ToolResult { Output = "set_headers_footers: slide " + slideIndex + " rejected this combination of fields (" + ex.Message + ").", IsError = true, Summary = "set_headers_footers" };
                    }
                }
            }

            // Both deck-wide regardless of slideIndex, same as
            // set_slide_background's slideIndex:-1 convention but stronger:
            // FirstSlideNumber (PageSetup) and DisplayOnTitleSlide (the
            // Slide/Title Master's HeadersFooters) are simply not per-slide
            // properties in PowerPoint's object model at all.
            if (startNumber.HasValue)
                pres.PageSetup.FirstSlideNumber = startNumber.Value;
            if (skipTitleSlide.HasValue)
                ApplySkipTitleSlide(skipTitleSlide.Value);

            return new ToolResult
            {
                Output = "Headers/footers updated (" + string.Join(", ", applied) + ") on " + (slideIndex == -1 ? "every slide, plus every layout and the Slide/Title Master so new slides inherit it too" : "slide " + slideIndex) +
                         (layoutFailures > 0 ? " (" + layoutFailures + " slide(s)/layout(s) rejected this combination of fields and were skipped - the rest still updated)." : ".") +
                         " (skipTitleSlide/startNumber, if given, always apply deck-wide regardless of slideIndex)." +
                         " If slide numbers/footer/date still don't appear, this slide's layout may not include that placeholder - " +
                         "check read_master_elements or try a different layout (set_slide_layout).",
                Mutated = true,
                Summary = "set_headers_footers",
            };
        }

        private static readonly string[] MasterCorners = { "topLeft", "topRight", "bottomLeft", "bottomRight" };

        private static ToolResult AddMasterElement(JsonElement input)
        {
            string kind = input.GetProperty("kind").GetString();
            if (kind != "icon" && kind != "text")
                throw new ArgumentException("add_master_element: unknown kind '" + kind + "'. Valid: icon, text.");

            string corner = input.TryGetProperty("corner", out var cEl) && cEl.ValueKind == JsonValueKind.String ? cEl.GetString() : null;
            if (corner != null && Array.IndexOf(MasterCorners, corner) < 0)
                throw new ArgumentException("add_master_element: unknown corner '" + corner + "'. Valid: " + string.Join(", ", MasterCorners) + ".");
            float? left = input.TryGetProperty("left", out var lEl) && lEl.ValueKind == JsonValueKind.Number ? (float?)lEl.GetDouble() : null;
            float? top = input.TryGetProperty("top", out var tEl) && tEl.ValueKind == JsonValueKind.Number ? (float?)tEl.GetDouble() : null;
            bool hasExplicitPosition = left.HasValue && top.HasValue;
            if (!hasExplicitPosition && corner == null)
                throw new ArgumentException("add_master_element: provide either corner, or both left and top.");

            float? width = input.TryGetProperty("width", out var wEl) && wEl.ValueKind == JsonValueKind.Number ? (float?)wEl.GetDouble() : null;
            float? height = input.TryGetProperty("height", out var hEl) && hEl.ValueKind == JsonValueKind.Number ? (float?)hEl.GetDouble() : null;
            float marginPt = input.TryGetProperty("marginPt", out var mEl) && mEl.ValueKind == JsonValueKind.Number ? (float)mEl.GetDouble() : 12f;

            MasterTarget target = ResolveMasterTarget(input, "add_master_element");
            PowerPoint.Shape shape;

            if (kind == "icon")
            {
                string localPath = input.GetProperty("localPath").GetString();
                if (localPath.StartsWith("http://") || localPath.StartsWith("https://"))
                    return new ToolResult { Output = "add_master_element: remote URLs are not supported in this air-gapped deployment - use a local file path.", IsError = true, Summary = "add_master_element" };
                if (!System.IO.File.Exists(localPath))
                    return new ToolResult { Output = "add_master_element: file not found: " + localPath, IsError = true, Summary = "add_master_element" };

                // Same pattern as WordTools.Images.cs's floating-picture path:
                // insert at natural size first (-1,-1), then let
                // GeometryUtil.ResolveImageSize scale a single missing
                // dimension proportionally instead of guessing/distorting.
                shape = target.Shapes.AddPicture(localPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, -1, -1);
                float naturalW = shape.Width, naturalH = shape.Height;
                float finalW, finalH;
                GeometryUtil.ResolveImageSize(naturalW, naturalH, width, height, out finalW, out finalH);
                shape.Width = finalW;
                shape.Height = finalH;
                width = finalW;
                height = finalH;
            }
            else
            {
                string text = input.GetProperty("text").GetString();
                float w = width ?? 200f;
                float h = height ?? 24f;
                shape = target.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, w, h);
                PowerPoint.TextRange range = shape.TextFrame.TextRange;
                range.Text = text;
                range.Font.Size = input.TryGetProperty("fontSize", out var fsEl) && fsEl.ValueKind == JsonValueKind.Number ? (float)fsEl.GetDouble() : 10f;
                range.Font.Color.RGB = input.TryGetProperty("color", out var colorEl) && colorEl.ValueKind == JsonValueKind.String
                    ? ColorUtil.HexToOle(colorEl.GetString())
                    : ColorUtil.HexToOle("#808080");
                width = w;
                height = h;
            }

            float finalLeft, finalTop;
            if (hasExplicitPosition)
            {
                finalLeft = left.Value;
                finalTop = top.Value;
            }
            else
            {
                PowerPoint.PageSetup pageSetup = ActivePresentation.PageSetup;
                GeometryUtil.ResolveCornerPosition(pageSetup.SlideWidth, pageSetup.SlideHeight, width.Value, height.Value, marginPt, corner, out finalLeft, out finalTop);
            }
            shape.Left = finalLeft;
            shape.Top = finalTop;

            string named = ApplyOptionalName(shape, input);
            int newIndex = shape.ZOrderPosition - 1;
            return new ToolResult
            {
                Output = "Master element added" + (named != null ? " (\"" + named + "\")" : "") + " to " + target.Label + " at masterShapeIndex " + newIndex +
                         (target.Label == "the Slide Master"
                             ? ". It will appear on every slide whose layout doesn't hide master shapes (\"Hide Background Graphics\")."
                             : ". It will appear only on slides using " + target.Label + " (not other layouts, and not through the Slide Master).") +
                         " Only this presentation's default Slide Master/theme is affected - a deck combining more than one theme has more than one.",
                Mutated = true,
                Summary = "add_master_element",
            };
        }

        // User-reported gap (2026-09-22, live testing): add_master_element
        // takes a starting position/size, but there was no way to adjust an
        // EXISTING master element afterward short of remove + re-add.
        // Mirrors PowerPointTools.Elements.cs's SetElementTransform exactly,
        // just addressed by masterShapeIndex instead of slideIndex+shapeIndex.
        private static ToolResult SetMasterElementTransform(JsonElement input)
        {
            int masterShapeIndex = input.GetProperty("masterShapeIndex").GetInt32();
            MasterTarget target = ResolveMasterTarget(input, "set_master_element_transform");
            if (masterShapeIndex < 0 || masterShapeIndex >= target.Shapes.Count)
                throw new ArgumentOutOfRangeException("masterShapeIndex", "masterShapeIndex must be between 0 and " + (target.Shapes.Count - 1) + " (in " + target.Label + ").");
            PowerPoint.Shape shape = target.Shapes[masterShapeIndex + 1];

            var applied = new List<string>();
            if (input.TryGetProperty("left", out var left)) { shape.Left = (float)left.GetDouble(); applied.Add("left"); }
            if (input.TryGetProperty("top", out var top)) { shape.Top = (float)top.GetDouble(); applied.Add("top"); }
            if (input.TryGetProperty("width", out var width)) { shape.Width = (float)width.GetDouble(); applied.Add("width"); }
            if (input.TryGetProperty("height", out var height)) { shape.Height = (float)height.GetDouble(); applied.Add("height"); }
            if (input.TryGetProperty("rotation", out var rotation)) { shape.Rotation = (float)rotation.GetDouble(); applied.Add("rotation"); }

            if (applied.Count == 0)
                return new ToolResult { Output = "set_master_element_transform: no recognized fields were provided - nothing changed.", IsError = true, Summary = "set_master_element_transform" };

            return new ToolResult
            {
                Output = "Master element " + masterShapeIndex + " in " + target.Label + " updated (" + string.Join(", ", applied) + ").",
                Mutated = true,
                Summary = "set_master_element_transform",
            };
        }

        // Real-user-confirmed (2026-09-22, live testing): the Slide Master's
        // Title/Text/Footer/DateAndTime/SlideNumber theme placeholders live
        // in master.Shapes right alongside anything add_master_element adds
        // - masterShapeIndex 0 is the Title Placeholder on a typical theme,
        // NOT whatever was added last. Deleting one of these did NOT remove
        // its appearance from actual slides even after save/close/reopen -
        // each slide's own CustomLayout carries an independent placeholder
        // object (inheriting the master's formatting, not sharing the same
        // shape reference), so master.Shapes only ever reflects the Slide
        // Master's OWN copy. Net effect of the naive index-0-by-default
        // testing: two theme placeholders (Title, Text) were deleted from
        // the Slide Master with no visible effect and no way to undo via
        // this tool surface. Refusing to delete a placeholder here closes
        // that hole - a custom element add_master_element actually created
        // (msoTextBox/msoPicture, never msoPlaceholder) is unaffected.
        private static ToolResult RemoveMasterElement(JsonElement input)
        {
            int masterShapeIndex = input.GetProperty("masterShapeIndex").GetInt32();
            MasterTarget target = ResolveMasterTarget(input, "remove_master_element");
            if (masterShapeIndex < 0 || masterShapeIndex >= target.Shapes.Count)
                throw new ArgumentOutOfRangeException("masterShapeIndex", "masterShapeIndex must be between 0 and " + (target.Shapes.Count - 1) + " (in " + target.Label + ").");
            PowerPoint.Shape shape = target.Shapes[masterShapeIndex + 1];
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                throw new ArgumentException("remove_master_element: masterShapeIndex " + masterShapeIndex + " (\"" + shape.Name + "\") in " + target.Label + " is one of this theme's own placeholders " +
                    "(Title/Text/Footer/Date/SlideNumber style), not something add_master_element created - refusing to delete it. " +
                    "To turn off a slide number/footer/date placeholder's visibility (not delete the shape), use set_headers_footers with the matching *Visible:false field instead. " +
                    "Deleting a Slide Master placeholder would not even remove its appearance from slides using a layout with its own independent copy - " +
                    "use PowerPoint's own Slide Master view (Master Layout dialog) if you really need to remove a theme placeholder.");
            shape.Delete();
            return new ToolResult
            {
                Output = "Master element " + masterShapeIndex + " deleted from " + target.Label + ". Other master shapes' indices may have shifted - call read_master_elements before another edit in the same run.",
                Mutated = true,
                Summary = "remove_master_element",
            };
        }

        private static ToolResult ReadMasterElements(JsonElement input)
        {
            MasterTarget target = ResolveMasterTarget(input, "read_master_elements");
            int count = target.Shapes.Count;
            if (count == 0)
                return new ToolResult { Output = "No shapes in " + target.Label + ".", Summary = "read_master_elements" };

            var sb = new StringBuilder();
            sb.AppendLine(CapitalizeFirst(target.Label) + " has " + count + " shape(s):");
            for (int i = 1; i <= count; i++)
            {
                PowerPoint.Shape shape = target.Shapes[i];
                string text = "";
                try
                {
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        text = shape.TextFrame.TextRange.Text.Replace("\r", " ").Trim();
                }
                catch { }
                sb.AppendLine("[" + (i - 1) + "] \"" + shape.Name + "\" (" + ShapeKindLabel(shape) + ") left=" + shape.Left + " top=" + shape.Top +
                              " width=" + shape.Width + " height=" + shape.Height + (text.Length > 0 ? " text=\"" + text + "\"" : ""));
            }
            return new ToolResult { Output = sb.ToString().TrimEnd(), Summary = "read_master_elements" };
        }
    }
}
