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
                if (fields.Contains("bold") && style.TryGetProperty("bold", out var bold))
                    range.Font.Bold = bold.ValueKind == JsonValueKind.True ? 1 : 0;
                if (fields.Contains("italic") && style.TryGetProperty("italic", out var italic))
                    range.Font.Italic = italic.ValueKind == JsonValueKind.True ? 1 : 0;
                if (fields.Contains("underline") && style.TryGetProperty("underline", out var underline))
                    range.Font.Underline = underline.ValueKind == JsonValueKind.True ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone;
                if (fields.Contains("strike") && style.TryGetProperty("strike", out var strike))
                    range.Font.StrikeThrough = strike.ValueKind == JsonValueKind.True ? 1 : 0;
                if (fields.Contains("sizeHalfPoints") && style.TryGetProperty("sizeHalfPoints", out var size) && size.ValueKind == JsonValueKind.Number)
                    range.Font.Size = (float)(size.GetDouble() / 2.0);
                if (fields.Contains("font") && style.TryGetProperty("font", out var font) && font.ValueKind == JsonValueKind.String)
                    range.Font.Name = font.GetString();
                if (fields.Contains("color") && style.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
                    range.Font.Color = (Word.WdColor)ColorUtil.HexToOle(color.GetString());
                if (fields.Contains("baselineOffset") && style.TryGetProperty("baselineOffset", out var baseline) && baseline.ValueKind == JsonValueKind.String)
                {
                    string b = baseline.GetString();
                    range.Font.Superscript = b == "SUPERSCRIPT" ? 1 : 0;
                    range.Font.Subscript = b == "SUBSCRIPT" ? 1 : 0;
                }
                if (fields.Contains("link") && style.TryGetProperty("link", out var link) && link.ValueKind == JsonValueKind.Object)
                {
                    string url = link.GetProperty("url").GetString();
                    ActiveDoc.Hyperlinks.Add(range, url);
                }
                if (fields.Contains("highlight") && style.TryGetProperty("highlight", out var highlight) && highlight.ValueKind == JsonValueKind.String)
                {
                    Word.WdColorIndex idx;
                    if (!HighlightColors.TryGetValue(highlight.GetString(), out idx))
                        throw new ArgumentException("updateTextStyle: unknown highlight color '" + highlight.GetString() +
                                                    "'. Valid: " + string.Join(", ", HighlightColors.Keys) + ".");
                    range.HighlightColorIndex = idx;
                }
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
                Word.ParagraphFormat fmt = p.Format;
                if (fields.Contains("align") && style.TryGetProperty("align", out var align) && align.ValueKind == JsonValueKind.String)
                {
                    switch (align.GetString())
                    {
                        case "left": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;
                        case "center": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;
                        case "right": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; break;
                        case "justify": fmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify; break;
                    }
                }
                if (fields.Contains("lineSpacing") && style.TryGetProperty("lineSpacing", out var ls) && ls.ValueKind == JsonValueKind.Number)
                    fmt.LineSpacing = (float)ls.GetDouble();
                if (fields.Contains("indentLeft") && style.TryGetProperty("indentLeft", out var il) && il.ValueKind == JsonValueKind.Number)
                    fmt.LeftIndent = (float)il.GetDouble();
                if (fields.Contains("indentRight") && style.TryGetProperty("indentRight", out var ir) && ir.ValueKind == JsonValueKind.Number)
                    fmt.RightIndent = (float)ir.GetDouble();
                if (fields.Contains("indentFirstLine") && style.TryGetProperty("indentFirstLine", out var ifl) && ifl.ValueKind == JsonValueKind.Number)
                    fmt.FirstLineIndent = (float)ifl.GetDouble();
                if (fields.Contains("spaceBefore") && style.TryGetProperty("spaceBefore", out var sb) && sb.ValueKind == JsonValueKind.Number)
                    fmt.SpaceBefore = (float)sb.GetDouble();
                if (fields.Contains("spaceAfter") && style.TryGetProperty("spaceAfter", out var sa) && sa.ValueKind == JsonValueKind.Number)
                    fmt.SpaceAfter = (float)sa.GetDouble();
                if (fields.Contains("pageBreakBefore") && style.TryGetProperty("pageBreakBefore", out var pbb))
                    fmt.PageBreakBefore = pbb.ValueKind == JsonValueKind.True ? 1 : 0;
                if (fields.Contains("shadingFill") && style.TryGetProperty("shadingFill", out var shading) && shading.ValueKind == JsonValueKind.String)
                    p.Shading.BackgroundPatternColor = (Word.WdColor)ColorUtil.HexToOle(shading.GetString());
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

