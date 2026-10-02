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
        // Known limitation, same class as ActivePresentation itself (PP-1): a deck
        // combining more than one theme has more than one Slide Master, and only
        // this default (first) one is targeted. See PowerPointTools.Master.cs.md.
        private static PowerPoint.Master ResolveSlideMaster()
        {
            return ActivePresentation.SlideMaster;
        }

        // Lets add_master_element/read_master_elements/remove_master_element/
        // set_master_element_transform target one specific layout (optional
        // layoutName, resolved by substring like ResolveCustomLayout) instead of
        // only the top-level Slide Master. See PowerPointTools.Master.cs.md.
        private struct MasterTarget
        {
            public PowerPoint.Shapes Shapes;
            public string Label;
        }

        // Shared by ResolveLayoutByName (below) and ResolveCustomLayout
        // (PowerPointTools.LayoutAnim.cs). DesignLabel is non-null only when
        // searching more than one design at once.
        private struct LayoutCandidate
        {
            public PowerPoint.CustomLayout Layout;
            public string DesignLabel;
        }

        // Exact (case-insensitive) name matches are checked as their own tier, ahead
        // of substring matches, and more than one surviving candidate is a thrown
        // ambiguity error rather than a silent first-match pick. See
        // PowerPointTools.Master.cs.md for the bug this fixed.
        private static LayoutCandidate ResolveLayoutByQuery(List<LayoutCandidate> candidates, string query, string toolName, string scopeDescription)
        {
            var exactMatches = new List<LayoutCandidate>();
            var substringMatches = new List<LayoutCandidate>();
            var namesSeen = new List<string>();
            foreach (var c in candidates)
            {
                namesSeen.Add(c.DesignLabel != null ? c.Layout.Name + " (design \"" + c.DesignLabel + "\")" : c.Layout.Name);
                if (string.Equals(c.Layout.Name, query, StringComparison.OrdinalIgnoreCase)) exactMatches.Add(c);
                else if (c.Layout.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) substringMatches.Add(c);
            }

            List<LayoutCandidate> winners = exactMatches.Count > 0 ? exactMatches : substringMatches;
            if (winners.Count == 1) return winners[0];
            if (winners.Count > 1)
            {
                var winnerNames = new List<string>();
                foreach (var w in winners)
                    winnerNames.Add(w.DesignLabel != null ? "\"" + w.Layout.Name + "\" (design \"" + w.DesignLabel + "\")" : "\"" + w.Layout.Name + "\"");
                throw new ArgumentException(toolName + ": '" + query + "' matches more than one layout in " + scopeDescription + ": " +
                    string.Join(", ", winnerNames) + ". Use a more specific or exact name.");
            }
            throw new ArgumentException(toolName + ": no layout matching '" + query + "' found in " + scopeDescription + ". Available: " + string.Join(", ", namesSeen) + ".");
        }

        // Searches every design in the presentation (Presentation.Designs), not just
        // the default one - a deck combining more than one theme/design could
        // otherwise never reach a non-default design's layout by name. See .md.
        private static LayoutCandidate ResolveLayoutByName(string query, string toolName)
        {
            var candidates = new List<LayoutCandidate>();
            bool multipleDesigns = ActivePresentation.Designs.Count > 1;
            foreach (PowerPoint.Design design in ActivePresentation.Designs)
                foreach (PowerPoint.CustomLayout layout in design.SlideMaster.CustomLayouts)
                    candidates.Add(new LayoutCandidate { Layout = layout, DesignLabel = multipleDesigns ? design.Name : null });
            return ResolveLayoutByQuery(candidates, query, toolName, "this presentation's theme(s)");
        }

        // Read-only: lists every layout in every design/theme in this presentation
        // (not just the default one), plus how many slides currently use each one.
        // See PowerPointTools.Master.cs.md.
        private static ToolResult ListLayouts(JsonElement input)
        {
            // CustomLayout.Index is only unique WITHIN its own design's SlideMaster -
            // key on the (design, layout) pair, not Index alone, or two designs'
            // same-Index layouts would be silently conflated.
            var usedByCountByKey = new Dictionary<string, int>();
            foreach (PowerPoint.Slide s in ActivePresentation.Slides)
            {
                try
                {
                    PowerPoint.CustomLayout layout = s.CustomLayout;
                    string key = layout.Design.Index + "|" + layout.Index;
                    usedByCountByKey[key] = usedByCountByKey.TryGetValue(key, out var c) ? c + 1 : 1;
                }
                catch { }
            }

            var sb = new StringBuilder();
            int i = 0;
            bool multipleDesigns = ActivePresentation.Designs.Count > 1;
            foreach (PowerPoint.Design design in ActivePresentation.Designs)
            {
                if (multipleDesigns) sb.AppendLine("Design \"" + design.Name + "\":");
                foreach (PowerPoint.CustomLayout layout in design.SlideMaster.CustomLayouts)
                {
                    string key = design.Index + "|" + layout.Index;
                    int usedByCount = usedByCountByKey.TryGetValue(key, out var count) ? count : 0;
                    sb.AppendLine("[" + i + "] \"" + layout.Name + "\"" + (usedByCount > 0 ? " - used by " + usedByCount + " slide(s)" : ""));
                    i++;
                }
            }
            if (i == 0)
                return new ToolResult { Output = "No layouts found in this presentation's theme(s).", Summary = "list_layouts" };
            return new ToolResult { Output = "This presentation has " + i + " layout(s)" + (multipleDesigns ? " across " + ActivePresentation.Designs.Count + " design(s)" : "") + ":\n" + sb.ToString().TrimEnd(), Summary = "list_layouts" };
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
                    LayoutCandidate match = ResolveLayoutByName(query, toolName);
                    string label = "layout \"" + match.Layout.Name + "\"" + (match.DesignLabel != null ? " (design \"" + match.DesignLabel + "\")" : "");
                    return new MasterTarget { Shapes = match.Layout.Shapes, Label = label };
                }
            }
            return new MasterTarget { Shapes = ResolveSlideMaster().Shapes, Label = "the Slide Master" };
        }

        // Curated subset of PpDateTimeFormat for set_headers_footers's dateFormat -
        // excludes time-only, combined date+time, locale-only, and sentinel/mixed
        // members that don't fit a settable "date format" field. See .md.
        private static readonly Dictionary<string, PowerPoint.PpDateTimeFormat> DateAutoFormats = new Dictionary<string, PowerPoint.PpDateTimeFormat>
        {
            { "M/d/yy", PowerPoint.PpDateTimeFormat.ppDateTimeMdyy },
            { "dddd, MMMM dd, yyyy", PowerPoint.PpDateTimeFormat.ppDateTimeddddMMMMddyyyy },
            { "d MMMM, yyyy", PowerPoint.PpDateTimeFormat.ppDateTimedMMMMyyyy },
            { "MMMM d, yyyy", PowerPoint.PpDateTimeFormat.ppDateTimeMMMMdyyyy },
            { "d-MMM-yy", PowerPoint.PpDateTimeFormat.ppDateTimedMMMyy },
            { "MMMM yy", PowerPoint.PpDateTimeFormat.ppDateTimeMMMMyy },
            { "MM/yy", PowerPoint.PpDateTimeFormat.ppDateTimeMMyy },
        };

        // DisplayOnTitleSlide is deliberately NOT handled here - setting it via a
        // SLIDE's HeadersFooters throws a COM restriction reflection can't see (see
        // ApplySkipTitleSlide below, which sets it on the MASTER instead). Takes a
        // HeadersFooters directly (not a Slide) so the same logic applies to a
        // slide's own HeadersFooters and to a Master's identical COM type - this is
        // also applied to the master(s) on a deck-wide call so new slides inherit
        // it too. See PowerPointTools.Master.cs.md.
        private static void ApplyHeadersFooters(
            PowerPoint.HeadersFooters hf, bool? slideNumberVisible,
            bool? footerVisible, string footerText,
            bool? dateVisible, string dateMode, string dateText, PowerPoint.PpDateTimeFormat dateFormat)
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
                hf.DateAndTime.Format = dateFormat;
            }
            else if (dateMode == "fixed")
            {
                hf.DateAndTime.UseFormat = Microsoft.Office.Core.MsoTriState.msoFalse;
                hf.DateAndTime.Text = dateText;
            }
        }

        // Shared wrapper replacing 6 near-identical hand-repeated try/catch sites in
        // SetHeadersFooters below. Returns null on success, or the caught
        // exception's message on failure (never throws). See .md.
        private static string TryApplyHeadersFooters(
            PowerPoint.HeadersFooters hf, bool? slideNumberVisible,
            bool? footerVisible, string footerText,
            bool? dateVisible, string dateMode, string dateText, PowerPoint.PpDateTimeFormat dateFormat,
            string debugLabel)
        {
            try
            {
                ApplyHeadersFooters(hf, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat);
                return null;
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("SetHeadersFooters " + debugLabel, ex);
                return ex.Message;
            }
        }

        // DisplayOnTitleSlide (the "Don't show on title slide" checkbox) is deck-wide
        // by nature - it lives on the Slide Master's (and, if present, the Title
        // Master's) HeadersFooters, not any individual slide's. Both writes are
        // independently guarded so one throwing doesn't block the other. See .md.
        private static string ApplySkipTitleSlide(bool skipTitleSlide)
        {
            PowerPoint.Presentation pres = ActivePresentation;
            Microsoft.Office.Core.MsoTriState value = skipTitleSlide ? Microsoft.Office.Core.MsoTriState.msoFalse : Microsoft.Office.Core.MsoTriState.msoTrue;
            string slideMasterError = null;
            try { pres.SlideMaster.HeadersFooters.DisplayOnTitleSlide = value; }
            catch (Exception ex) { DebugLog.WriteException("ApplySkipTitleSlide SlideMaster", ex); slideMasterError = ex.Message; }

            // Even READING pres.TitleMaster (not just writing to it) can throw in some
            // deck states despite HasTitleMaster reporting true first - best-effort,
            // never let this abort the rest of set_headers_footers. See .md.
            try
            {
                if (pres.HasTitleMaster == Microsoft.Office.Core.MsoTriState.msoTrue)
                    pres.TitleMaster.HeadersFooters.DisplayOnTitleSlide = value;
            }
            catch (Exception ex) { DebugLog.WriteException("ApplySkipTitleSlide TitleMaster", ex); }
            return slideMasterError;
        }

        // Shared by ReadSlide (PowerPointTools.Read.cs) so the model can see what
        // set_headers_footers already did.
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
            string dateFormatKey = input.TryGetProperty("dateFormat", out var dfEl) && dfEl.ValueKind == JsonValueKind.String ? dfEl.GetString() : null;
            // Infer dateMode "fixed" from dateText alone (mirrors footerText inferring
            // footerVisible:true above); "auto" is opt-in only, never inferred. See .md.
            if (dateMode == null && dateText != null) dateMode = "fixed";
            if (dateMode != null && dateMode != "auto" && dateMode != "fixed")
                throw new ArgumentException("set_headers_footers: unknown dateMode '" + dateMode + "'. Valid: auto, fixed.");
            if (dateMode == "fixed" && dateText == null)
                throw new ArgumentException("set_headers_footers: dateMode 'fixed' requires dateText.");
            if (dateFormatKey != null && dateMode != "auto")
                throw new ArgumentException("set_headers_footers: dateFormat only applies when dateMode is 'auto' - pass dateMode:\"auto\" explicitly to use it (dateMode defaults to \"fixed\" when not given).");
            PowerPoint.PpDateTimeFormat dateFormat = PowerPoint.PpDateTimeFormat.ppDateTimeMdyy;
            if (dateFormatKey != null && !DateAutoFormats.TryGetValue(dateFormatKey, out dateFormat))
                throw new ArgumentException("set_headers_footers: unknown dateFormat '" + dateFormatKey + "'. Valid: " + string.Join(", ", DateAutoFormats.Keys) + ".");
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
            string singleSlideError = null;
            if (anyPerSlideField)
            {
                if (slideIndex == -1)
                {
                    // Per-item best-effort treatment, same as the master/layout loops
                    // below - a rejected combination on one slide must not abort the
                    // rest of the deck-wide call. See PowerPointTools.Master.cs.md.
                    int slideFailures = 0;
                    foreach (PowerPoint.Slide s in pres.Slides)
                    {
                        if (TryApplyHeadersFooters(s.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat, "slide " + (s.SlideIndex - 1)) != null)
                            slideFailures++;
                    }
                    layoutFailures += slideFailures;

                    // Also seed the master(s) AND every custom layout, so a slide added
                    // AFTER this call starts with the same state - seeding only the
                    // Master is NOT enough, since CustomLayout.HeadersFooters is a
                    // separate object a new slide actually inherits from. Every call
                    // below is individually guarded, since some layouts/masters throw
                    // on at least one of these writes. Best-effort per item; failures
                    // are counted, not hidden, and don't stop the rest. See .md.
                    if (TryApplyHeadersFooters(pres.SlideMaster.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat, "SlideMaster") != null)
                        layoutFailures++;

                    try
                    {
                        foreach (PowerPoint.CustomLayout layout in pres.SlideMaster.CustomLayouts)
                        {
                            if (TryApplyHeadersFooters(layout.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat, "layout '" + layout.Name + "'") != null)
                                layoutFailures++;
                        }
                    }
                    catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters SlideMaster.CustomLayouts enumeration", ex); layoutFailures++; }

                    try
                    {
                        if (pres.HasTitleMaster == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            PowerPoint.Master titleMaster = pres.TitleMaster;
                            if (TryApplyHeadersFooters(titleMaster.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat, "TitleMaster") != null)
                                layoutFailures++;
                            foreach (PowerPoint.CustomLayout layout in titleMaster.CustomLayouts)
                            {
                                if (TryApplyHeadersFooters(layout.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat, "title-master layout '" + layout.Name + "'") != null)
                                    layoutFailures++;
                            }
                        }
                    }
                    catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters TitleMaster access", ex); layoutFailures++; }
                }
                else
                {
                    // Record the failure and keep going, rather than returning
                    // immediately - startNumber/skipTitleSlide below are always
                    // deck-wide and must still be attempted. See .md.
                    singleSlideError = TryApplyHeadersFooters(pres.Slides[slideIndex + 1].HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat, "slide " + slideIndex);
                }
            }

            // FirstSlideNumber (PageSetup) and DisplayOnTitleSlide are simply not
            // per-slide properties in PowerPoint's object model - both deck-wide
            // regardless of slideIndex, and applied even when the per-slide write
            // above failed. Each is guarded independently so one failing doesn't
            // mask the other succeeding or the per-slide work above. See .md.
            bool startNumberApplied = false, skipTitleSlideApplied = false;
            var deckWideFailures = new List<string>();
            if (startNumber.HasValue)
            {
                try { pres.PageSetup.FirstSlideNumber = startNumber.Value; startNumberApplied = true; }
                catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters startNumber", ex); deckWideFailures.Add("startNumber (" + ex.Message + ")"); }
            }
            if (skipTitleSlide.HasValue)
            {
                string skipTitleSlideError = ApplySkipTitleSlide(skipTitleSlide.Value);
                if (skipTitleSlideError == null) skipTitleSlideApplied = true;
                else deckWideFailures.Add("skipTitleSlide (" + skipTitleSlideError + ")");
            }

            if (singleSlideError != null)
            {
                bool deckWideStillApplied = startNumberApplied || skipTitleSlideApplied;
                return new ToolResult
                {
                    Output = "set_headers_footers: slide " + slideIndex + " rejected this combination of fields (" + singleSlideError + ")." +
                             (deckWideStillApplied ? " skipTitleSlide/startNumber, since given, were still applied deck-wide" +
                                                      (deckWideFailures.Count > 0 ? " except " + string.Join(", ", deckWideFailures) : "") + "."
                                                    : deckWideFailures.Count > 0 ? " skipTitleSlide/startNumber were also rejected: " + string.Join(", ", deckWideFailures) + "." : ""),
                    IsError = true,
                    Mutated = deckWideStillApplied,
                    Summary = "set_headers_footers",
                };
            }

            // Describes only what actually ran - a deck-wide call with no
            // per-slide fields (e.g. only skipTitleSlide/startNumber) never touches
            // the per-slide/layout/master loop above, so must not claim it did.
            string scopeDescription;
            if (!anyPerSlideField)
                scopeDescription = "no per-slide fields were given";
            else if (slideIndex == -1)
                scopeDescription = "every slide, plus every layout and the Slide/Title Master so new slides inherit it too" +
                                    (layoutFailures > 0 ? " (" + layoutFailures + " slide(s)/layout(s) rejected this combination of fields and were skipped - the rest still updated)" : "");
            else
                scopeDescription = "slide " + slideIndex;

            return new ToolResult
            {
                Output = "Headers/footers updated (" + string.Join(", ", applied) + ") - " + scopeDescription + "." +
                         (deckWideFailures.Count > 0 ? " (" + string.Join(", ", deckWideFailures) + " rejected and were skipped.)" : "") +
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
            float? left = input.TryGetProperty("left", out var lEl) && lEl.ValueKind == JsonValueKind.Number ? (float?)lEl.GetDouble() : null;
            float? top = input.TryGetProperty("top", out var tEl) && tEl.ValueKind == JsonValueKind.Number ? (float?)tEl.GetDouble() : null;
            // left/top must be given together, or a single coordinate is silently
            // dropped in favor of corner-based placement. See PowerPointTools.Master.cs.md.
            if (left.HasValue != top.HasValue)
                throw new ArgumentException("add_master_element: left and top must be given together (a single coordinate alone is not supported) - provide both, or use corner instead.");
            bool hasExplicitPosition = left.HasValue && top.HasValue;
            // Only require/validate corner when it will actually be used - checking it
            // before hasExplicitPosition is known would reject valid left+top alongside
            // a stray corner value that's never consulted.
            if (!hasExplicitPosition)
            {
                if (corner == null)
                    throw new ArgumentException("add_master_element: provide either corner, or both left and top.");
                if (Array.IndexOf(MasterCorners, corner) < 0)
                    throw new ArgumentException("add_master_element: unknown corner '" + corner + "'. Valid: " + string.Join(", ", MasterCorners) + ".");
            }

            float? width = input.TryGetProperty("width", out var wEl) && wEl.ValueKind == JsonValueKind.Number ? (float?)wEl.GetDouble() : null;
            float? height = input.TryGetProperty("height", out var hEl) && hEl.ValueKind == JsonValueKind.Number ? (float?)hEl.GetDouble() : null;
            float marginPt = input.TryGetProperty("marginPt", out var mEl) && mEl.ValueKind == JsonValueKind.Number ? (float)mEl.GetDouble() : 12f;

            // Resolve/validate color and font size up front, before creating
            // anything - doing it after shape creation orphaned an unpositioned,
            // unnamed shape on an invalid hex string. See .md.
            string text = null;
            float textFontSize = 10f;
            int textColor = 0;
            if (kind == "text")
            {
                text = input.GetProperty("text").GetString();
                textFontSize = input.TryGetProperty("fontSize", out var fsEl) && fsEl.ValueKind == JsonValueKind.Number ? (float)fsEl.GetDouble() : 10f;
                textColor = input.TryGetProperty("color", out var colorEl) && colorEl.ValueKind == JsonValueKind.String
                    ? ColorUtil.HexToOle(colorEl.GetString())
                    : ColorUtil.HexToOle("#808080");
            }

            MasterTarget target = ResolveMasterTarget(input, "add_master_element");
            PowerPoint.Shape shape;

            if (kind == "icon")
            {
                string localPath = input.GetProperty("localPath").GetString();
                if (localPath.StartsWith("http://") || localPath.StartsWith("https://"))
                    return new ToolResult { Output = "add_master_element: remote URLs are not supported in this air-gapped deployment - use a local file path.", IsError = true, Summary = "add_master_element" };
                if (!System.IO.File.Exists(localPath))
                    return new ToolResult { Output = "add_master_element: file not found: " + localPath, IsError = true, Summary = "add_master_element" };

                shape = target.Shapes.AddPicture(localPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, -1, -1);
            }
            else
            {
                float w = width ?? 200f;
                float h = height ?? 24f;
                shape = target.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, w, h);
                width = w;
                height = h;
            }

            // Point of no return: the shape now really exists. If anything from here
            // on throws, clean up the orphan rather than leave it unpositioned/unnamed
            // (same discipline as CrossSlide's CopyPasteShape/CopyOrMoveElement).
            string named = null;
            try
            {
                if (kind == "icon")
                {
                    // Same pattern as WordTools.Images.cs: insert at natural size
                    // first (-1,-1), then scale a missing dimension proportionally.
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
                    PowerPoint.TextRange range = shape.TextFrame.TextRange;
                    range.Text = text;
                    range.Font.Size = textFontSize;
                    range.Font.Color.RGB = textColor;
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

                // Pass target.Shapes explicitly - ApplyOptionalName's dedup resolves
                // the parent via `shape.Parent as PowerPoint.Slide`, always null for a
                // master/layout shape, so dedup would otherwise silently no-op here.
                named = ApplyOptionalName(shape, input, target.Shapes);
            }
            catch
            {
                try { shape.Delete(); } catch { }
                throw;
            }

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

        // Adjusts an existing master element's position/size (previously only
        // possible via remove + re-add). Mirrors PowerPointTools.Elements.cs's
        // SetElementTransform, addressed by masterShapeIndex instead.
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

        // Incident (2026-09-22, live-tested): the Slide Master's own theme
        // placeholders (Title/Text/Footer/Date/SlideNumber) live in master.Shapes
        // alongside anything add_master_element adds, and deleting one does NOT
        // remove its appearance from slides - each slide's own CustomLayout carries
        // an independent, non-shared placeholder copy. Two theme placeholders were
        // lost this way with no undo path, which is why this refuses to delete any
        // msoPlaceholder outright. Full incident writeup: PowerPointTools.Master.cs.md.
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
