using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    // Spike 3: real COM tool execution against the live Word document, called
    // from the WebView2-hosted AgentLoop via the JSON WebMessage bridge.
    public static partial class WordTools
    {
        // PP-12 Task 1: Word highlighting is a fixed 16-entry palette
        // (WdColorIndex), NOT arbitrary RGB - unlike Font.Color above, which
        // "color" uses. Accept only these names; anything else is an error
        // rather than a silent nearest-match.
        private static readonly Dictionary<string, Word.WdColorIndex> HighlightColors =
            new Dictionary<string, Word.WdColorIndex>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = Word.WdColorIndex.wdNoHighlight,
            ["yellow"] = Word.WdColorIndex.wdYellow,
            ["brightGreen"] = Word.WdColorIndex.wdBrightGreen,
            ["turquoise"] = Word.WdColorIndex.wdTurquoise,
            ["pink"] = Word.WdColorIndex.wdPink,
            ["blue"] = Word.WdColorIndex.wdBlue,
            ["red"] = Word.WdColorIndex.wdRed,
            ["darkBlue"] = Word.WdColorIndex.wdDarkBlue,
            ["teal"] = Word.WdColorIndex.wdTeal,
            ["green"] = Word.WdColorIndex.wdGreen,
            ["violet"] = Word.WdColorIndex.wdViolet,
            ["darkRed"] = Word.WdColorIndex.wdDarkRed,
            ["darkYellow"] = Word.WdColorIndex.wdDarkYellow,
            ["gray50"] = Word.WdColorIndex.wdGray50,
            ["gray25"] = Word.WdColorIndex.wdGray25,
            ["black"] = Word.WdColorIndex.wdBlack,
            ["white"] = Word.WdColorIndex.wdWhite,
        };

        // PP-12 Task 1 Step 3: the general false-success hole - any
        // misspelled/unimplemented field name in `fields` previously matched
        // no `if` and silently applied nothing while still reporting "ok".
        private static readonly HashSet<string> KnownTextStyleFields = new HashSet<string>
        { "bold", "italic", "underline", "strike", "sizeHalfPoints", "font", "color", "baselineOffset", "link", "highlight" };

        private static readonly HashSet<string> KnownParagraphStyleFields = new HashSet<string>
        { "align", "lineSpacing", "indentLeft", "indentRight", "indentFirstLine", "spaceBefore", "spaceAfter", "pageBreakBefore", "shadingFill", "borders" };

        // Shared by UpdateTextStyle (parses these from JSON) and CopyFormatCmd
        // (reads these live from a source paragraph's Font, no JSON round
        // trip). Native-enum params, not the JSON-facing bool/string shapes,
        // so a copy never loses fidelity a JSON-driven caller doesn't need
        // anyway - e.g. `underline` can be any real WdUnderline style here,
        // not just the true/false-&gt;Single/None the schema knows.
        private static void ApplyTextStyle(
            Word.Range range, bool? bold, bool? italic, Word.WdUnderline? underline, bool? strike,
            float? sizePoints, string font, Word.WdColor? color, bool? superscript, bool? subscript,
            Word.WdColorIndex? highlight)
        {
            if (bold.HasValue) range.Font.Bold = bold.Value ? 1 : 0;
            if (italic.HasValue) range.Font.Italic = italic.Value ? 1 : 0;
            if (underline.HasValue) range.Font.Underline = underline.Value;
            if (strike.HasValue) range.Font.StrikeThrough = strike.Value ? 1 : 0;
            if (sizePoints.HasValue) range.Font.Size = sizePoints.Value;
            if (font != null) range.Font.Name = font;
            if (color.HasValue) range.Font.Color = color.Value;
            if (superscript.HasValue) range.Font.Superscript = superscript.Value ? 1 : 0;
            if (subscript.HasValue) range.Font.Subscript = subscript.Value ? 1 : 0;
            if (highlight.HasValue) range.HighlightColorIndex = highlight.Value;
        }

        // Shared by UpdateParagraphStyle and CopyFormatCmd - see ApplyTextStyle's
        // comment above. Borders are deliberately NOT here - CopyFormatCmd
        // copies them per-side (Top/Left/Bottom/Right) directly, since this
        // method's all-on/all-off boolean shape is a disclosed simplification
        // for WRITING, not something a READ of mixed border sides should be
        // forced through.
        private static void ApplyParagraphStyle(
            Word.Paragraph p, Word.WdParagraphAlignment? align, float? lineSpacing, float? indentLeft,
            float? indentRight, float? indentFirstLine, float? spaceBefore, float? spaceAfter,
            bool? pageBreakBefore, Word.WdColor? shadingFill)
        {
            Word.ParagraphFormat fmt = p.Format;
            if (align.HasValue) fmt.Alignment = align.Value;
            if (lineSpacing.HasValue) fmt.LineSpacing = lineSpacing.Value;
            if (indentLeft.HasValue) fmt.LeftIndent = indentLeft.Value;
            if (indentRight.HasValue) fmt.RightIndent = indentRight.Value;
            if (indentFirstLine.HasValue) fmt.FirstLineIndent = indentFirstLine.Value;
            if (spaceBefore.HasValue) fmt.SpaceBefore = spaceBefore.Value;
            if (spaceAfter.HasValue) fmt.SpaceAfter = spaceAfter.Value;
            if (pageBreakBefore.HasValue) fmt.PageBreakBefore = pageBreakBefore.Value ? 1 : 0;
            if (shadingFill.HasValue) p.Shading.BackgroundPatternColor = shadingFill.Value;
        }

        private static void UpdateTextStyle(JsonElement cmd)
        {
            var matches = ResolveTargetParagraphs(cmd.GetProperty("target"));
            JsonElement style = cmd.GetProperty("style");
            HashSet<string> fields = new HashSet<string>();
            foreach (JsonElement f in cmd.GetProperty("fields").EnumerateArray()) fields.Add(f.GetString());
            ToolArgs.ValidateKnownFields(fields, KnownTextStyleFields, "updateTextStyle");

            if (matches.Count == 0)
            {
                throw new InvalidOperationException("updateTextStyle: no paragraphs matched target.");
            }

            foreach (var (_, p) in matches)
            {
                Word.Range range = p.Range;
                bool? bold = fields.Contains("bold") && style.TryGetProperty("bold", out var boldEl)
                    ? boldEl.ValueKind == JsonValueKind.True : (bool?)null;
                bool? italic = fields.Contains("italic") && style.TryGetProperty("italic", out var italicEl)
                    ? italicEl.ValueKind == JsonValueKind.True : (bool?)null;
                Word.WdUnderline? underline = fields.Contains("underline") && style.TryGetProperty("underline", out var underlineEl)
                    ? (underlineEl.ValueKind == JsonValueKind.True ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone)
                    : (Word.WdUnderline?)null;
                bool? strike = fields.Contains("strike") && style.TryGetProperty("strike", out var strikeEl)
                    ? strikeEl.ValueKind == JsonValueKind.True : (bool?)null;
                float? sizePoints = fields.Contains("sizeHalfPoints") && style.TryGetProperty("sizeHalfPoints", out var sizeEl) && sizeEl.ValueKind == JsonValueKind.Number
                    ? (float)(sizeEl.GetDouble() / 2.0) : (float?)null;
                string font = fields.Contains("font") && style.TryGetProperty("font", out var fontEl) && fontEl.ValueKind == JsonValueKind.String
                    ? fontEl.GetString() : null;
                Word.WdColor? color = fields.Contains("color") && style.TryGetProperty("color", out var colorEl) && colorEl.ValueKind == JsonValueKind.String
                    ? (Word.WdColor)ColorUtil.HexToOle(colorEl.GetString()) : (Word.WdColor?)null;
                bool? superscript = null, subscript = null;
                if (fields.Contains("baselineOffset") && style.TryGetProperty("baselineOffset", out var baseline) && baseline.ValueKind == JsonValueKind.String)
                {
                    string b = baseline.GetString();
                    superscript = b == "SUPERSCRIPT";
                    subscript = b == "SUBSCRIPT";
                }
                if (fields.Contains("link") && style.TryGetProperty("link", out var link) && link.ValueKind == JsonValueKind.Object)
                {
                    string url = link.GetProperty("url").GetString();
                    ActiveDoc.Hyperlinks.Add(range, url);
                }
                Word.WdColorIndex? highlight = null;
                if (fields.Contains("highlight") && style.TryGetProperty("highlight", out var highlightEl) && highlightEl.ValueKind == JsonValueKind.String)
                {
                    Word.WdColorIndex idx;
                    if (!HighlightColors.TryGetValue(highlightEl.GetString(), out idx))
                        throw new ArgumentException("updateTextStyle: unknown highlight color '" + highlightEl.GetString() +
                                                    "'. Valid: " + string.Join(", ", HighlightColors.Keys) + ".");
                    highlight = idx;
                }
                ApplyTextStyle(range, bold, italic, underline, strike, sizePoints, font, color, superscript, subscript, highlight);
            }
        }

        private static void UpdateParagraphStyle(JsonElement cmd)
        {
            var matches = ResolveTargetParagraphs(cmd.GetProperty("target"));
            JsonElement style = cmd.GetProperty("style");
            HashSet<string> fields = new HashSet<string>();
            foreach (JsonElement f in cmd.GetProperty("fields").EnumerateArray()) fields.Add(f.GetString());
            ToolArgs.ValidateKnownFields(fields, KnownParagraphStyleFields, "updateParagraphStyle");

            if (matches.Count == 0)
            {
                throw new InvalidOperationException("updateParagraphStyle: no paragraphs matched target.");
            }

            foreach (var (_, p) in matches)
            {
                Word.WdParagraphAlignment? align = null;
                if (fields.Contains("align") && style.TryGetProperty("align", out var alignEl) && alignEl.ValueKind == JsonValueKind.String)
                {
                    switch (alignEl.GetString())
                    {
                        case "left": align = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;
                        case "center": align = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;
                        case "right": align = Word.WdParagraphAlignment.wdAlignParagraphRight; break;
                        case "justify": align = Word.WdParagraphAlignment.wdAlignParagraphJustify; break;
                    }
                }
                float? lineSpacing = fields.Contains("lineSpacing") && style.TryGetProperty("lineSpacing", out var ls) && ls.ValueKind == JsonValueKind.Number
                    ? (float)ls.GetDouble() : (float?)null;
                float? indentLeft = fields.Contains("indentLeft") && style.TryGetProperty("indentLeft", out var il) && il.ValueKind == JsonValueKind.Number
                    ? (float)il.GetDouble() : (float?)null;
                float? indentRight = fields.Contains("indentRight") && style.TryGetProperty("indentRight", out var ir) && ir.ValueKind == JsonValueKind.Number
                    ? (float)ir.GetDouble() : (float?)null;
                float? indentFirstLine = fields.Contains("indentFirstLine") && style.TryGetProperty("indentFirstLine", out var ifl) && ifl.ValueKind == JsonValueKind.Number
                    ? (float)ifl.GetDouble() : (float?)null;
                float? spaceBefore = fields.Contains("spaceBefore") && style.TryGetProperty("spaceBefore", out var sb) && sb.ValueKind == JsonValueKind.Number
                    ? (float)sb.GetDouble() : (float?)null;
                float? spaceAfter = fields.Contains("spaceAfter") && style.TryGetProperty("spaceAfter", out var sa) && sa.ValueKind == JsonValueKind.Number
                    ? (float)sa.GetDouble() : (float?)null;
                bool? pageBreakBefore = fields.Contains("pageBreakBefore") && style.TryGetProperty("pageBreakBefore", out var pbb)
                    ? pbb.ValueKind == JsonValueKind.True : (bool?)null;
                Word.WdColor? shadingFill = fields.Contains("shadingFill") && style.TryGetProperty("shadingFill", out var shading) && shading.ValueKind == JsonValueKind.String
                    ? (Word.WdColor)ColorUtil.HexToOle(shading.GetString()) : (Word.WdColor?)null;
                ApplyParagraphStyle(p, align, lineSpacing, indentLeft, indentRight, indentFirstLine, spaceBefore, spaceAfter, pageBreakBefore, shadingFill);

                if (fields.Contains("borders") && style.TryGetProperty("borders", out var borders))
                {
                    bool on = borders.ValueKind == JsonValueKind.True;
                    foreach (Word.Border border in p.Borders)
                    {
                        border.LineStyle = on ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleNone;
                    }
                }
            }
        }

        // Format painter: copies ALL of sourceBlockIndex's character (Font)
        // and paragraph (ParagraphFormat/Shading/4-side Borders) formatting
        // onto every paragraph matched by `target`, atomically. Whole-
        // paragraph granularity only - no sub-paragraph text-run targeting.
        // Hyperlinks are deliberately never copied (real Word Format Painter
        // doesn't carry them either - a hyperlink is a document part, not a
        // font/paragraph attribute).
        private static void CopyFormatCmd(JsonElement cmd)
        {
            int sourceBlockIndex = cmd.GetProperty("sourceBlockIndex").GetInt32();
            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            int count = paragraphs.Count;
            if (sourceBlockIndex < 0 || sourceBlockIndex >= count)
                throw new ArgumentOutOfRangeException("sourceBlockIndex",
                    "copyFormat: sourceBlockIndex must be between 0 and " + (count - 1) + " (" + count + " paragraph(s) in the document).");

            Word.Paragraph source = paragraphs[sourceBlockIndex + 1];
            Word.Range srcRange = source.Range;
            Word.ParagraphFormat srcFmt = source.Format;

            // One read, one consistent snapshot, applied identically to every
            // target below - the atomicity guarantee. NOTE: a source range
            // spanning non-uniform character formatting (e.g. half-bold) can
            // return Word's "mixed value" sentinel for some of these
            // properties rather than a real value - unverified against real
            // Word from this codebase (see the plan's Risks section); this
            // reads the raw live values without a mixed-value guard for now.
            bool bold = srcRange.Font.Bold == -1;
            bool italic = srcRange.Font.Italic == -1;
            Word.WdUnderline underline = srcRange.Font.Underline;
            bool strike = srcRange.Font.StrikeThrough == -1;
            float sizePoints = srcRange.Font.Size;
            string font = srcRange.Font.Name;
            Word.WdColor color = srcRange.Font.Color;
            bool superscript = srcRange.Font.Superscript == -1;
            bool subscript = srcRange.Font.Subscript == -1;
            Word.WdColorIndex highlight = srcRange.HighlightColorIndex;

            Word.WdParagraphAlignment align = srcFmt.Alignment;
            float lineSpacing = srcFmt.LineSpacing;
            float indentLeft = srcFmt.LeftIndent;
            float indentRight = srcFmt.RightIndent;
            float indentFirstLine = srcFmt.FirstLineIndent;
            float spaceBefore = srcFmt.SpaceBefore;
            float spaceAfter = srcFmt.SpaceAfter;
            bool pageBreakBefore = srcFmt.PageBreakBefore == -1;
            Word.WdColor shadingFill = source.Shading.BackgroundPatternColor;

            // Explicit 4-named-side copy - NOT a foreach over p.Borders (that
            // collection's all-on/all-off write-path shape, used by
            // UpdateParagraphStyle above, would misrepresent a paragraph with
            // mixed border sides if reused for a read). Mirrors PP-23's table-
            // border fix: only Top/Left/Bottom/Right, never the diagonal
            // entries.
            Word.WdLineStyle topStyle = source.Borders[Word.WdBorderType.wdBorderTop].LineStyle;
            Word.WdColor topColor = source.Borders[Word.WdBorderType.wdBorderTop].Color;
            Word.WdLineStyle leftStyle = source.Borders[Word.WdBorderType.wdBorderLeft].LineStyle;
            Word.WdColor leftColor = source.Borders[Word.WdBorderType.wdBorderLeft].Color;
            Word.WdLineStyle bottomStyle = source.Borders[Word.WdBorderType.wdBorderBottom].LineStyle;
            Word.WdColor bottomColor = source.Borders[Word.WdBorderType.wdBorderBottom].Color;
            Word.WdLineStyle rightStyle = source.Borders[Word.WdBorderType.wdBorderRight].LineStyle;
            Word.WdColor rightColor = source.Borders[Word.WdBorderType.wdBorderRight].Color;

            var targets = ResolveTargetParagraphs(cmd.GetProperty("target"));
            if (targets.Count == 0)
            {
                throw new InvalidOperationException("copyFormat: no paragraphs matched target.");
            }

            foreach (var (_, p) in targets)
            {
                ApplyTextStyle(p.Range, bold, italic, underline, strike, sizePoints, font, color, superscript, subscript, highlight);
                ApplyParagraphStyle(p, align, lineSpacing, indentLeft, indentRight, indentFirstLine, spaceBefore, spaceAfter, pageBreakBefore, shadingFill);

                p.Borders[Word.WdBorderType.wdBorderTop].LineStyle = topStyle;
                if (topStyle != Word.WdLineStyle.wdLineStyleNone) p.Borders[Word.WdBorderType.wdBorderTop].Color = topColor;
                p.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = leftStyle;
                if (leftStyle != Word.WdLineStyle.wdLineStyleNone) p.Borders[Word.WdBorderType.wdBorderLeft].Color = leftColor;
                p.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = bottomStyle;
                if (bottomStyle != Word.WdLineStyle.wdLineStyleNone) p.Borders[Word.WdBorderType.wdBorderBottom].Color = bottomColor;
                p.Borders[Word.WdBorderType.wdBorderRight].LineStyle = rightStyle;
                if (rightStyle != Word.WdLineStyle.wdLineStyleNone) p.Borders[Word.WdBorderType.wdBorderRight].Color = rightColor;
            }
        }

        // PP-12 Task 2: fixed, explicit preset set - each implemented by
        // applying Word's own proven default bullet/number list (rather than
        // constructing a ListTemplate from a gallery index, which the plan
        // itself flags as unstable across Office versions/locales) and then,
        // where the preset needs more than the default, overriding the
        // resulting level's NumberStyle/NumberFormat explicitly. The two
        // Wingdings-glyph variants (diamond/checkbox) are the least certain
        // of the seven without an interactive Word session to verify against -
        // flagged in this plan's verification file; narrow the enum to drop
        // them if they don't render correctly (Step 7's sanctioned fallback).
        private static readonly HashSet<string> BulletPresets = new HashSet<string>
        {
            "BULLET_DISC_CIRCLE_SQUARE", "BULLET_DIAMOND_X", "BULLET_CHECKBOX",
            "NUMBERED_DECIMAL", "NUMBERED_DECIMAL_ALPHA_ROMAN", "NUMBERED_UPPERALPHA", "NUMBERED_UPPERROMAN",
        };

        private static void ApplyBulletPreset(Word.Range range, string preset)
        {
            switch (preset)
            {
                case "BULLET_DISC_CIRCLE_SQUARE":
                    range.ListFormat.ApplyBulletDefault();
                    break;
                case "BULLET_DIAMOND_X":
                    range.ListFormat.ApplyBulletDefault();
                    range.ListFormat.ListTemplate.ListLevels[1].NumberFormat = "¨"; // Wingdings diamond-ish glyph
                    range.ListFormat.ListTemplate.ListLevels[1].Font.Name = "Wingdings";
                    break;
                case "BULLET_CHECKBOX":
                    range.ListFormat.ApplyBulletDefault();
                    range.ListFormat.ListTemplate.ListLevels[1].NumberFormat = "£"; // Wingdings empty-box glyph
                    range.ListFormat.ListTemplate.ListLevels[1].Font.Name = "Wingdings";
                    break;
                case "NUMBERED_DECIMAL":
                    range.ListFormat.ApplyNumberDefault();
                    break;
                case "NUMBERED_DECIMAL_ALPHA_ROMAN":
                    // Word's per-level glyph sequence needs real multi-level
                    // nesting to show the alpha/roman sub-levels; this file's
                    // flat per-paragraph model has no such nesting, so level 1
                    // stays plain decimal - narrower than genoffice's version,
                    // but honestly so (documented in the schema description).
                    range.ListFormat.ApplyNumberDefault();
                    break;
                case "NUMBERED_UPPERALPHA":
                    range.ListFormat.ApplyNumberDefault();
                    range.ListFormat.ListTemplate.ListLevels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleUppercaseLetter;
                    break;
                case "NUMBERED_UPPERROMAN":
                    range.ListFormat.ApplyNumberDefault();
                    range.ListFormat.ListTemplate.ListLevels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleUppercaseRoman;
                    break;
                default:
                    throw new ArgumentException("createParagraphBullets: unknown bulletPreset '" + preset +
                                                "'. Valid: " + string.Join(", ", BulletPresets) + ".");
            }
        }

        // Returns a report string (PP-12 Task 2 Step 5 / Task 4) instead of
        // void + a bare "ok" - the caller (ApplyCommands) uses this text
        // directly so a skipped-heading count is visible, not silently lost.
        private static string CreateParagraphBullets(JsonElement cmd)
        {
            var matches = ResolveTargetParagraphs(cmd.GetProperty("target"));
            string preset = cmd.TryGetProperty("bulletPreset", out var bp) && bp.ValueKind == JsonValueKind.String ? bp.GetString() : null;

            if (matches.Count == 0)
            {
                throw new InvalidOperationException("createParagraphBullets: no paragraphs matched target.");
            }

            int applied = 0, skippedHeadings = 0;
            foreach (var (_, p) in matches)
            {
                Word.Range range = p.Range;
                string styleName = range.get_Style().NameLocal;
                if (styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) { skippedHeadings++; continue; } // headings are matched but left unchanged, mirrors genoffice
                if (preset != null) ApplyBulletPreset(range, preset);
                else range.ListFormat.ApplyBulletDefault(); // absent bulletPreset keeps the pre-existing default behavior
                applied++;
            }

            return $"createParagraphBullets: {applied} applied, {skippedHeadings} heading(s) skipped.";
        }

        // Returns a report string (PP-12 Task 4) instead of void + a bare
        // "ok" - a target matching only non-list paragraphs previously
        // reported success while changing nothing.
        private static string DeleteParagraphBullets(JsonElement cmd)
        {
            var matches = ResolveTargetParagraphs(cmd.GetProperty("target"));
            if (matches.Count == 0)
            {
                throw new InvalidOperationException("deleteParagraphBullets: no paragraphs matched target.");
            }
            int removed = 0, skippedNonList = 0;
            foreach (var (_, p) in matches)
            {
                Word.Range range = p.Range;
                if (range.ListFormat.ListType == Word.WdListType.wdListNoNumbering) { skippedNonList++; continue; } // non-list-item matches silently skipped, mirrors genoffice
                range.ListFormat.RemoveNumbers();
                removed++;
            }
            return $"deleteParagraphBullets: {removed} removed, {skippedNonList} non-list paragraph(s) skipped.";
        }

    }
}

