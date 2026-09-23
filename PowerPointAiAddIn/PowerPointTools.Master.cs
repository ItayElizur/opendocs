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

        // Shared by ResolveLayoutByName (below, searches every design) and
        // ResolveCustomLayout (PowerPointTools.LayoutAnim.cs, searches one
        // slide's own design) - DesignLabel is only non-null when the caller
        // is searching more than one design at once, so a single-theme deck's
        // messages are unaffected.
        private struct LayoutCandidate
        {
            public PowerPoint.CustomLayout Layout;
            public string DesignLabel;
        }

        // Review finding: neither of this struct's two callers used to
        // detect an ambiguous match - both took the FIRST layout whose name
        // merely CONTAINED the query substring, in whatever order
        // CustomLayouts happened to enumerate, with no signal to the caller
        // that a second (or third) candidate existed at all. Two layouts can
        // collide on name (nothing in the object model enforces uniqueness -
        // ListLayouts's own count-by-index logic already works around this
        // same risk) or simply both happen to contain the same substring
        // (e.g. query "Title" against both "Title Slide" and "Title and
        // Content"). Exact (case-insensitive) name matches are checked as
        // their own tier, ahead of substring matches, so typing the literal
        // full name is never "ambiguous" just because some other layout's
        // name happens to contain it.
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

        // Review finding: this used to only ever search
        // ActivePresentation.SlideMaster (the presentation's default/first
        // design), because there was no slide to derive scope from - unlike
        // ResolveCustomLayout (PowerPointTools.LayoutAnim.cs), which is
        // handed a specific slide and correctly searches THAT slide's own
        // Design. A deck combining more than one theme/design (e.g. slides
        // pasted in with "Preserve Source Formatting") could never reach a
        // non-default design's layout by name through add_master_element/
        // read_master_elements/remove_master_element/set_master_element_
        // transform's layoutName. Now searches every design in the
        // presentation (confirmed via reflection: Presentation.Designs is a
        // real, enumerable collection, each with its own SlideMaster).
        private static LayoutCandidate ResolveLayoutByName(string query, string toolName)
        {
            var candidates = new List<LayoutCandidate>();
            bool multipleDesigns = ActivePresentation.Designs.Count > 1;
            foreach (PowerPoint.Design design in ActivePresentation.Designs)
                foreach (PowerPoint.CustomLayout layout in design.SlideMaster.CustomLayouts)
                    candidates.Add(new LayoutCandidate { Layout = layout, DesignLabel = multipleDesigns ? design.Name : null });
            return ResolveLayoutByQuery(candidates, query, toolName, "this presentation's theme(s)");
        }

        // User-requested (2026-09-22): before this, the only way to
        // discover a deck's real layout names was to deliberately pass a
        // bad layoutName and read the error message's "Available: ..."
        // list. Read-only, lists every layout in every design/theme in this
        // presentation (not just the default one - same widening as
        // ResolveLayoutByName above, for the same reason), plus how many
        // slides currently use each one.
        private static ToolResult ListLayouts(JsonElement input)
        {
            // Review finding: CustomLayout.Index is only unique WITHIN its
            // own design's SlideMaster - once more than one design is in
            // play, a bare Dictionary<int,int> keyed on Index alone would
            // silently conflate two different designs' layouts that happen
            // to share an Index (e.g. both designs' first layout is Index 1).
            // Key on the (design, layout) pair instead. CustomLayout.Design
            // and Design.Index are both confirmed via reflection.
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

        // Curated subset of the real PpDateTimeFormat enum (confirmed via
        // .NET reflection against the referenced PIA, Microsoft.Office.
        // Interop.PowerPoint.dll GAC 15.0.0.0) for set_headers_footers's
        // dateFormat. Deliberately excludes: the four time-only members
        // (Hmm/Hmmss/hmmAMPM/hmmssAMPM - a "date format" picker showing no
        // date at all doesn't fit this field); the two combined date+time
        // members (out of scope here - dateMode:"fixed" + dateText already
        // covers a time-bearing footer if ever needed); the seven locale-only
        // UAQ1-7 members (undocumented/regional); ppDateTimeFigureOut (a
        // sentinel PowerPoint uses internally, not a settable format); and
        // ppDateTimeFormatMixed (what PowerPoint reports when a range's
        // formats differ, never a value you set).
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
            string dateFormatKey = input.TryGetProperty("dateFormat", out var dfEl) && dfEl.ValueKind == JsonValueKind.String ? dfEl.GetString() : null;
            // Review finding: dateText alone (no dateMode) used to be
            // silently dropped - never applied, never counted in `applied`,
            // and the call could even return the "no recognized fields" error
            // despite dateText being a valid, given field. Infer dateMode the
            // same way footerText infers footerVisible:true just above.
            // User-directed: "auto" (and therefore dateFormat, which only
            // applies to "auto") is opt-in only and never inferred - dateMode
            // defaults to "fixed" and stays "fixed" unless the caller states
            // dateMode:"auto" explicitly, regardless of which dateFormat is
            // given.
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
                        try { ApplyHeadersFooters(s.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat); }
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
                    try { ApplyHeadersFooters(pres.SlideMaster.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat); }
                    catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters SlideMaster", ex); layoutFailures++; }

                    try
                    {
                        foreach (PowerPoint.CustomLayout layout in pres.SlideMaster.CustomLayouts)
                        {
                            try { ApplyHeadersFooters(layout.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat); }
                            catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters layout '" + layout.Name + "'", ex); layoutFailures++; }
                        }
                    }
                    catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters SlideMaster.CustomLayouts enumeration", ex); layoutFailures++; }

                    try
                    {
                        if (pres.HasTitleMaster == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            PowerPoint.Master titleMaster = pres.TitleMaster;
                            try { ApplyHeadersFooters(titleMaster.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat); }
                            catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters TitleMaster", ex); layoutFailures++; }
                            foreach (PowerPoint.CustomLayout layout in titleMaster.CustomLayouts)
                            {
                                try { ApplyHeadersFooters(layout.HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat); }
                                catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters title-master layout '" + layout.Name + "'", ex); layoutFailures++; }
                            }
                        }
                    }
                    catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters TitleMaster access", ex); layoutFailures++; }
                }
                else
                {
                    // Review finding: this branch used to return immediately
                    // on failure, before ever reaching the startNumber/
                    // skipTitleSlide handling below - since those two fields
                    // are documented as ALWAYS deck-wide regardless of
                    // slideIndex, a rejected per-slide field combination on
                    // one slide had the side effect of silently dropping an
                    // unrelated deck-wide field given in the same call.
                    // Record the failure and keep going instead.
                    try { ApplyHeadersFooters(pres.Slides[slideIndex + 1].HeadersFooters, slideNumberVisible, footerVisible, footerText, dateVisible, dateMode, dateText, dateFormat); }
                    catch (Exception ex)
                    {
                        DebugLog.WriteException("SetHeadersFooters slide " + slideIndex, ex);
                        singleSlideError = ex.Message;
                    }
                }
            }

            // Both deck-wide regardless of slideIndex, same as
            // set_slide_background's slideIndex:-1 convention but stronger:
            // FirstSlideNumber (PageSetup) and DisplayOnTitleSlide (the
            // Slide/Title Master's HeadersFooters) are simply not per-slide
            // properties in PowerPoint's object model at all. Applied even
            // when the per-slide write above failed (see singleSlideError).
            //
            // Review finding: these two writes used to be unguarded, unlike
            // every other write in this function - if a deck-wide call had
            // already mutated every slide's footer/date/number and THEN
            // PageSetup.FirstSlideNumber or ApplySkipTitleSlide's unguarded
            // DisplayOnTitleSlide write threw (the same class of COM
            // restriction this file documents extensively elsewhere), the
            // exception would escape to Execute's outer catch and report a
            // bare error with no Mutated:true, hiding the changes that had
            // already landed. Guard each independently so one failing
            // doesn't mask the other succeeding or the per-slide work above.
            bool startNumberApplied = false, skipTitleSlideApplied = false;
            var deckWideFailures = new List<string>();
            if (startNumber.HasValue)
            {
                try { pres.PageSetup.FirstSlideNumber = startNumber.Value; startNumberApplied = true; }
                catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters startNumber", ex); deckWideFailures.Add("startNumber (" + ex.Message + ")"); }
            }
            if (skipTitleSlide.HasValue)
            {
                try { ApplySkipTitleSlide(skipTitleSlide.Value); skipTitleSlideApplied = true; }
                catch (Exception ex) { DebugLog.WriteException("SetHeadersFooters skipTitleSlide", ex); deckWideFailures.Add("skipTitleSlide (" + ex.Message + ")"); }
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

            // Review finding: this message used to unconditionally describe
            // "every slide, plus every layout and the Slide/Title Master" for
            // a deck-wide call even when anyPerSlideField was false (e.g. only
            // skipTitleSlide/startNumber given) - overstating what actually
            // changed, since the per-slide/layout/master loop above never ran
            // in that case.
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
            // Review finding: left/top used to only count as an explicit
            // position when BOTH were given; giving just one alongside/without
            // corner silently fell through to corner-based placement (or the
            // margin default) with no error - the supplied coordinate was
            // dropped without a trace.
            if (left.HasValue != top.HasValue)
                throw new ArgumentException("add_master_element: left and top must be given together (a single coordinate alone is not supported) - provide both, or use corner instead.");
            bool hasExplicitPosition = left.HasValue && top.HasValue;
            // Review finding: corner's validity used to be checked before
            // hasExplicitPosition was known, so valid left+top alongside a
            // stray/garbage corner value threw even though corner would
            // never actually be used once an explicit position is given.
            // Only require/validate corner when it will actually be used.
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

            // Review finding: for kind:"text", ColorUtil.HexToOle used to run
            // AFTER the textbox was already created on the master/layout - an
            // invalid hex string (e.g. "gray") threw with the shape already
            // inserted but never positioned or named, orphaning it at (0,0)
            // with no shapeIndex returned for cleanup. Resolve/validate
            // color and font size up front, before creating anything.
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

            // Point of no return: the shape now really exists on the
            // master/layout. Everything from here on is best-effort setup -
            // if any of it throws, clean up the orphan instead of leaving a
            // stray, unpositioned/unnamed shape behind (same orphan-cleanup
            // discipline as CrossSlide's CopyPasteElement/MoveElement).
            string named = null;
            try
            {
                if (kind == "icon")
                {
                    // Same pattern as WordTools.Images.cs's floating-picture
                    // path: insert at natural size first (-1,-1), then let
                    // GeometryUtil.ResolveImageSize scale a single missing
                    // dimension proportionally instead of guessing/distorting.
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

                // Review finding: ApplyOptionalName's dedup-against-siblings
                // used to resolve the shape's parent via `shape.Parent as
                // PowerPoint.Slide`, which is always null for a master/layout
                // shape - two add_master_element calls with the same name
                // both got the literal name, unlike slide shapes which
                // auto-suffix. Pass target.Shapes explicitly so dedup works
                // here too.
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
