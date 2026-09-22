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

        // Slide.HeadersFooters.DateAndTime/.Footer/.SlideNumber are all the
        // same HeaderFooter type. DisplayOnTitleSlide is NOT among them here -
        // real-Office-confirmed (2026-09-21, live error): setting it via a
        // SLIDE's HeadersFooters throws "HeadersFooters (unknown member) :
        // Invalid request. This property must be set using slide or title
        // master." even though reflection shows the member on the shared
        // HeadersFooters type with no hint of that restriction - a runtime
        // COM restriction reflection can't see, only member existence/type.
        // See ApplySkipTitleSlide below, which sets it on the MASTER instead.
        private static void ApplyHeadersFooters(
            PowerPoint.Slide slide, bool? slideNumberVisible,
            bool? footerVisible, string footerText,
            bool? dateVisible, string dateMode, string dateText)
        {
            PowerPoint.HeadersFooters hf = slide.HeadersFooters;
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
            if (pres.HasTitleMaster == Microsoft.Office.Core.MsoTriState.msoTrue)
                pres.TitleMaster.HeadersFooters.DisplayOnTitleSlide = value;
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

            if (slideNumberVisible.HasValue || footerVisible.HasValue || footerText != null || dateVisible.HasValue || dateMode != null)
            {
                if (slideIndex == -1)
                {
                    foreach (PowerPoint.Slide s in pres.Slides)
                        ApplyHeadersFooters(s, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText);
                }
                else
                {
                    ApplyHeadersFooters(pres.Slides[slideIndex + 1], slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText);
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
                Output = "Headers/footers updated (" + string.Join(", ", applied) + ") on " + (slideIndex == -1 ? "every slide" : "slide " + slideIndex) +
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

            PowerPoint.Master master = ResolveSlideMaster();
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
                shape = master.Shapes.AddPicture(localPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, -1, -1);
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
                shape = master.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, w, h);
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
                Output = "Master element added" + (named != null ? " (\"" + named + "\")" : "") + " at masterShapeIndex " + newIndex +
                         ". It will appear on every slide whose layout doesn't hide master shapes (\"Hide Background Graphics\"). " +
                         "Only this presentation's default Slide Master is affected - a deck combining more than one theme has more than one master.",
                Mutated = true,
                Summary = "add_master_element",
            };
        }

        private static ToolResult RemoveMasterElement(JsonElement input)
        {
            int masterShapeIndex = input.GetProperty("masterShapeIndex").GetInt32();
            PowerPoint.Master master = ResolveSlideMaster();
            if (masterShapeIndex < 0 || masterShapeIndex >= master.Shapes.Count)
                throw new ArgumentOutOfRangeException("masterShapeIndex", "masterShapeIndex must be between 0 and " + (master.Shapes.Count - 1) + ".");
            master.Shapes[masterShapeIndex + 1].Delete();
            return new ToolResult
            {
                Output = "Master element " + masterShapeIndex + " deleted. Other master shapes' indices may have shifted - call read_master_elements before another edit in the same run.",
                Mutated = true,
                Summary = "remove_master_element",
            };
        }

        private static ToolResult ReadMasterElements(JsonElement input)
        {
            PowerPoint.Master master = ResolveSlideMaster();
            int count = master.Shapes.Count;
            if (count == 0)
                return new ToolResult { Output = "No shapes on the Slide Master.", Summary = "read_master_elements" };

            var sb = new StringBuilder();
            sb.AppendLine("Slide Master has " + count + " shape(s):");
            for (int i = 1; i <= count; i++)
            {
                PowerPoint.Shape shape = master.Shapes[i];
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
