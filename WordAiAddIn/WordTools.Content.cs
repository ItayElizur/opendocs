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
        // Resolves a Range's 0-based paragraph index (matching
        // ActiveDoc.Paragraphs' own indexing, which read_blocks/
        // apply_commands/find_text/get_headings all address by) without
        // Word's slow positional Paragraphs[i] lookup - indexing the
        // Paragraphs collection by position has to re-walk the document
        // from the start on EVERY single access, which turned a scan of N
        // positions into roughly O(N^2) internally (confirmed root cause of
        // a real freeze report). This instead marches forward once via the
        // cheap Paragraph.Next() chain, and only as far as needed since the
        // last call - callers must request positions in non-decreasing
        // document order (true for both find_text and get_headings, which
        // each only ever move forward through the document), so the total
        // marching work across a whole call is O(N), not O(N) per lookup.
        private sealed class ParagraphIndexResolver
        {
            private Word.Paragraph _current;
            private int _index;

            public ParagraphIndexResolver(Word.Document doc)
            {
                _current = doc.Paragraphs.First;
                _index = 0;
            }

            public int IndexAt(int rangeStart)
            {
                while (_current != null && _current.Range.End <= rangeStart)
                {
                    _index++;
                    try { _current = _current.Next(); }
                    catch (System.Runtime.InteropServices.COMException) { _current = null; }
                }
                return _index;
            }
        }

        // Read-only search - unlike apply_commands' find_replace, this never
        // touches the document. Added because there was previously no way to
        // locate text without either mutating it (find_replace) or reading
        // the whole document paragraph-by-paragraph via read_blocks.
        //
        // Plain-substring queries use Word's own native Find engine - the
        // same one behind Ctrl+F - which does a single optimized traversal
        // and only costs work proportional to the number of MATCHES, not the
        // number of paragraphs in the document (the original implementation
        // scanned every paragraph via positional Paragraphs[i] indexing
        // regardless of match count, which is what caused a real reported
        // freeze on a large document). Word's Find has no regex mode (only
        // its own more limited wildcard syntax), so a regex:true query still
        // needs a per-paragraph scan - but via the cheap forward
        // Paragraph.Next() chain, not positional indexing.
        private static ToolResult FindText(JsonElement input)
        {
            string query = input.GetProperty("query").GetString();
            bool useRegex = input.TryGetProperty("regex", out var rx) && rx.ValueKind == JsonValueKind.True;
            bool matchCase = input.TryGetProperty("matchCase", out var mc) && mc.ValueKind == JsonValueKind.True;
            int maxResults = input.GetProperty("max_results").GetInt32();

            Word.Document doc = ActiveDoc;
            var sb = new System.Text.StringBuilder();
            int found = 0;
            var resolver = new ParagraphIndexResolver(doc);

            if (useRegex)
            {
                var regex = new System.Text.RegularExpressions.Regex(query, matchCase
                    ? System.Text.RegularExpressions.RegexOptions.None
                    : System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                Word.Paragraph p = doc.Paragraphs.First;
                while (p != null && found < maxResults)
                {
                    string text = p.Range.Text.TrimEnd('\r', '\a', '\n');
                    if (regex.IsMatch(text))
                    {
                        int index = resolver.IndexAt(p.Range.Start);
                        sb.AppendLine($"[{index}] {text}");
                        found++;
                    }
                    try { p = p.Next(); }
                    catch (System.Runtime.InteropServices.COMException) { p = null; }
                }
            }
            else
            {
                Word.Range searchRange = doc.Content;
                Word.Find findObj = searchRange.Find;
                findObj.ClearFormatting();
                findObj.Text = query;
                findObj.MatchCase = matchCase;
                findObj.MatchWildcards = false;
                findObj.Forward = true;
                findObj.Wrap = Word.WdFindWrap.wdFindStop;
                while (found < maxResults && findObj.Execute())
                {
                    int index = resolver.IndexAt(searchRange.Start);
                    Word.Paragraph containing = searchRange.Paragraphs[1];
                    string text = containing.Range.Text.TrimEnd('\r', '\a', '\n');
                    sb.AppendLine($"[{index}] {text}");
                    found++;
                }
            }

            return new ToolResult { Output = found > 0 ? sb.ToString() : "No matches.", Summary = "find_text" };
        }

        // Navigation-Pane-style outline: every Heading-styled paragraph with
        // its index and level, so the model can see document structure
        // without reading every paragraph via read_blocks.
        //
        // Uses Word's own wdGoToHeading jump - the same internal heading
        // index that powers the Navigation Pane and "Browse by Heading" -
        // which lands directly on each heading without ever touching a
        // non-heading paragraph, instead of scanning every paragraph's style
        // name to find the ones that are headings.
        private static ToolResult GetHeadings()
        {
            Word.Document doc = ActiveDoc;
            var sb = new System.Text.StringBuilder();
            int found = 0;
            var resolver = new ParagraphIndexResolver(doc);

            Word.Range cursor = doc.Content;
            cursor.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            // wdGoToHeading wraps back to the first heading once it runs out
            // of headings ahead (same as the "Browse by Heading" scrollbar
            // control) rather than signaling "no more" - detected below via
            // next.Start <= cursor.Start. This safety cap is a defensive
            // backstop in case that wrap ever isn't caught (e.g. an
            // off-by-one on a document with an unusual heading at the very
            // end), so a real bug there degrades to a merely-incomplete
            // result instead of an infinite loop.
            int safetyLimit = doc.Paragraphs.Count + 1;
            for (int i = 0; i < safetyLimit; i++)
            {
                Word.Range next;
                try
                {
                    next = cursor.GoTo(What: Word.WdGoToItem.wdGoToHeading, Which: Word.WdGoToDirection.wdGoToNext);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    break; // no headings in the document at all
                }
                if (next.Start <= cursor.Start) break; // wrapped back to the start - no more headings ahead

                Word.Paragraph headingPara = next.Paragraphs[1];
                int index = resolver.IndexAt(headingPara.Range.Start);
                string styleName = headingPara.Range.get_Style().NameLocal;
                string digits = new string(styleName.Where(char.IsDigit).ToArray());
                int level = int.TryParse(digits, out int lvl) ? lvl : 1;
                string text = headingPara.Range.Text.TrimEnd('\r', '\a', '\n');
                sb.AppendLine($"[{index}] H{level}: {text}");
                found++;

                cursor = next.Duplicate;
                cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                cursor.Move(Word.WdUnits.wdCharacter, 1);
            }

            return new ToolResult { Output = found > 0 ? sb.ToString() : "No headings in this document.", Summary = "get_headings" };
        }

        private static ToolResult GetDocumentContext()
        {
            Word.Document doc = ActiveDoc;
            int paraCount = doc.Paragraphs.Count;
            int wordCount = doc.Words.Count;
            string preview = doc.Content.Text;
            if (preview.Length > 300) preview = preview.Substring(0, 300) + "...";
            string output = $"Paragraphs: {paraCount}, Words: {wordCount}\nPreview: {preview}";
            return new ToolResult { Output = output, Summary = "read document context" };
        }

        // PP-10 Task 1 + Task 3: positional insert, plain-text or restricted
        // HTML. Backward compatible: {text} with no afterBlockIndex/html
        // keeps the exact original end-of-document, single-paragraph
        // behavior (Global Constraint - existing prompts/history depend on it).
        private static ToolResult InsertContent(JsonElement input)
        {
            bool hasText = input.TryGetProperty("text", out var textEl) && textEl.ValueKind == JsonValueKind.String;
            bool hasHtml = input.TryGetProperty("html", out var htmlEl) && htmlEl.ValueKind == JsonValueKind.String;
            if (hasText && hasHtml)
                throw new ArgumentException("insert_content: pass exactly one of text or html, not both.");
            if (!hasText && !hasHtml)
                throw new ArgumentException("insert_content: text or html is required.");

            int? afterBlockIndex = input.TryGetProperty("afterBlockIndex", out var abEl) && abEl.ValueKind == JsonValueKind.Number
                ? abEl.GetInt32() : (int?)null;

            int paragraphsBefore = ActiveDoc.Paragraphs.Count;
            Word.Range at;
            if (afterBlockIndex.HasValue)
            {
                at = RangeAfterBlock(afterBlockIndex.Value);
            }
            else
            {
                // No position given: exact original behavior - append at the
                // very end of the document.
                at = ActiveDoc.Content;
                at.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

            string reportedText;
            if (hasHtml)
            {
                InsertHtmlFragment(at, htmlEl.GetString());
                reportedText = "(HTML fragment)";
            }
            else
            {
                string text = textEl.GetString();
                // Split on newlines into separate paragraphs - a multi-line
                // text previously produced one paragraph with literal line
                // breaks, mangling the model's own formatting intent.
                string[] lines = text.Split('\n');
                foreach (string line in lines)
                {
                    at.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    at.InsertParagraphAfter();
                    at.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    at.Text = line.TrimEnd('\r');
                }
                reportedText = text;
            }

            int paragraphsAfter = ActiveDoc.Paragraphs.Count;
            int insertedCount = paragraphsAfter - paragraphsBefore;
            int firstInserted = afterBlockIndex.HasValue ? afterBlockIndex.Value + 1 : paragraphsBefore;
            int lastInserted = firstInserted + insertedCount - 1;

            return new ToolResult
            {
                Output = $"Inserted paragraphs [{firstInserted}-{lastInserted}]: {reportedText}",
                Mutated = true,
                Summary = "insert_content",
            };
        }

        // PP-10 Task 4: 'html' mode emits the same restricted subset
        // InsertHtmlFragment accepts, so read_blocks -> replace_blocks/
        // insert_content round-trips (headings/bold/italic/underline/list
        // membership survive; anything outside the subset, e.g. font color,
        // is documented as not surviving). Capped well below text mode - a
        // per-word COM property read for every paragraph is far slower than
        // a flat Range.Text read.
        private const int HtmlReadBlocksMaxParagraphs = 100;

        // Post-hoc addition (2026-08-27, user-reported): 'text' mode (the
        // default) previously had NO cap at all - only one Range.Text read
        // per paragraph, so it was never capped the way 'html' mode was, but
        // an unbounded range on a very large document still means an
        // unbounded amount of walking and an unbounded output string (a
        // separate, non-perf concern - context/token budget). Not
        // independently benchmarked against a measured time budget the way
        // read_formats' 200-cell cap or html mode's 100-paragraph cap were -
        // chosen conservatively; raise it if real usage shows it's too tight.
        private const int TextReadBlocksMaxParagraphs = 1000;

        private static string ReadBlockAsHtml(Word.Paragraph p)
        {
            var sb = new System.Text.StringBuilder();
            bool bold = false, italic = false, underline = false;
            foreach (Word.Range word in p.Range.Words)
            {
                bool wBold = (int)word.Font.Bold == 1;
                bool wItalic = (int)word.Font.Italic == 1;
                bool wUnderline = word.Font.Underline != Word.WdUnderline.wdUnderlineNone;
                // Close/open tags in a fixed order (u, i, b) so nesting is
                // always well-formed regardless of which properties changed.
                if (underline && !wUnderline) { sb.Append("</u>"); underline = false; }
                if (italic && !wItalic) { sb.Append("</i>"); italic = false; }
                if (bold && !wBold) { sb.Append("</b>"); bold = false; }
                if (wBold && !bold) { sb.Append("<b>"); bold = true; }
                if (wItalic && !italic) { sb.Append("<i>"); italic = true; }
                if (wUnderline && !underline) { sb.Append("<u>"); underline = true; }
                sb.Append(TextUtil.HtmlEscape(word.Text));
            }
            if (underline) sb.Append("</u>");
            if (italic) sb.Append("</i>");
            if (bold) sb.Append("</b>");
            return sb.ToString().TrimEnd('\r', '\a', '\n');
        }

        private static ToolResult ReadBlocks(JsonElement input)
        {
            int startIndex = input.GetProperty("startIndex").GetInt32();
            int endIndex = input.GetProperty("endIndex").GetInt32();
            string format = input.TryGetProperty("format", out var fmtEl) && fmtEl.ValueKind == JsonValueKind.String ? fmtEl.GetString() : "text";
            Word.Document doc = ActiveDoc;
            Word.Paragraphs paragraphs = doc.Paragraphs;
            int count = paragraphs.Count;
            endIndex = Math.Min(endIndex, count - 1);
            if (startIndex < 0 || startIndex > endIndex)
            {
                return new ToolResult { Output = "Invalid range.", IsError = true, Summary = "read_blocks" };
            }

            int maxParagraphs = format == "html" ? HtmlReadBlocksMaxParagraphs : TextReadBlocksMaxParagraphs;
            if ((endIndex - startIndex + 1) > maxParagraphs)
            {
                throw new ArgumentException("read_blocks: format:'" + format + "' is capped at " + maxParagraphs +
                    " paragraphs per call - page the request into smaller ranges.");
            }

            var sb = new System.Text.StringBuilder();
            string openList = null; // null | "ul" | "ol"

            // Walks forward via Paragraph.Next() instead of positional
            // paragraphs[i + 1] indexing - Paragraphs is not a real array in
            // Word's COM object model, so indexing it by position re-walks
            // the document from the start on EVERY single access (confirmed
            // root cause of a real reported freeze elsewhere in this file -
            // find_text/get_headings/ResolveTargetParagraphs all hit the
            // same trap and were fixed the same way).
            Word.Paragraph p = paragraphs.First;
            for (int skip = 0; skip < startIndex && p != null; skip++) p = p.Next();

            for (int i = startIndex; i <= endIndex && p != null; i++)
            {
                // p must advance exactly once per iteration regardless of
                // which branch below runs - finally guarantees that even
                // though several branches `continue` early.
                try
                {
                    if (format != "html")
                    {
                        string plain = p.Range.Text.TrimEnd('\r', '\a', '\n');
                        sb.AppendLine($"[{i}] {plain}");
                        continue;
                    }

                    string styleName = p.Range.get_Style().NameLocal;
                    bool isListItem = p.Range.ListFormat.ListType != Word.WdListType.wdListNoNumbering;
                    bool isNumbered = isListItem && p.Range.ListFormat.ListType != Word.WdListType.wdListBullet;
                    string content = ReadBlockAsHtml(p);

                    if (isListItem)
                    {
                        string wantList = isNumbered ? "ol" : "ul";
                        if (openList != wantList)
                        {
                            if (openList != null) sb.AppendLine($"</{openList}>");
                            sb.AppendLine($"<{wantList}>");
                            openList = wantList;
                        }
                        sb.AppendLine($"[{i}] <li>{content}</li>");
                        continue;
                    }
                    if (openList != null) { sb.AppendLine($"</{openList}>"); openList = null; }

                    int headingLevel = 0;
                    if (styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
                    {
                        string digits = new string(styleName.Where(char.IsDigit).ToArray());
                        if (int.TryParse(digits, out int lvl) && lvl >= 1 && lvl <= 3) headingLevel = lvl;
                    }
                    string tag = headingLevel > 0 ? "h" + headingLevel : "p";
                    sb.AppendLine($"[{i}] <{tag}>{content}</{tag}>");
                }
                finally
                {
                    p = p.Next();
                }
            }
            if (format == "html" && openList != null) sb.AppendLine($"</{openList}>");

            return new ToolResult { Output = sb.ToString(), Summary = "read_blocks" };
        }

        // PP-10 Task 2 + Task 3 Step 5: non-destructive replace (preserves
        // the first replaced paragraph's style by default, instead of
        // silently stripping heading/list identity), plus an html
        // alternative to text.
        private static ToolResult ReplaceBlocks(JsonElement input)
        {
            int startIndex = input.GetProperty("startIndex").GetInt32();
            int endIndex = input.GetProperty("endIndex").GetInt32();
            bool hasText = input.TryGetProperty("text", out var textEl) && textEl.ValueKind == JsonValueKind.String;
            bool hasHtml = input.TryGetProperty("html", out var htmlEl) && htmlEl.ValueKind == JsonValueKind.String;
            if (hasText && hasHtml)
                throw new ArgumentException("replace_blocks: pass exactly one of text or html, not both.");
            if (!hasText && !hasHtml)
                throw new ArgumentException("replace_blocks: text or html is required (pass an empty text to delete the range).");
            // Default true: an absent preserveFormatting field means preserve.
            bool preserveFormatting = !input.TryGetProperty("preserveFormatting", out var pf) || pf.ValueKind != JsonValueKind.False;

            Word.Document doc = ActiveDoc;
            Word.Paragraphs paragraphs = doc.Paragraphs;
            endIndex = Math.Min(endIndex, paragraphs.Count - 1);
            if (startIndex < 0 || startIndex > endIndex)
            {
                return new ToolResult { Output = "Invalid range.", IsError = true, Summary = "replace_blocks" };
            }

            // preserveFormatting is meaningless for the html path - the
            // fragment's own tags (h1-h3, etc.) dictate paragraph style, and
            // reapplying the old style afterward would overwrite that intent.
            string capturedStyle = (preserveFormatting && !hasHtml) ? paragraphs[startIndex + 1].Range.get_Style().NameLocal : null;

            Word.Range range = doc.Range(paragraphs[startIndex + 1].Range.Start, paragraphs[endIndex + 1].Range.End);

            if (hasHtml)
            {
                range.Delete();
                InsertHtmlFragment(range, htmlEl.GetString());
                return new ToolResult { Output = $"Replaced paragraphs {startIndex}-{endIndex} with HTML fragment.", Mutated = true, Summary = "replace_blocks" };
            }

            string text = textEl.GetString() ?? "";
            range.Text = text;

            string preservedNote = "";
            if (capturedStyle != null && text.Length > 0)
            {
                // Reapply the FIRST replaced paragraph's style to every
                // resulting paragraph. Multi-paragraph replacements where the
                // source paragraphs had differing styles are genuinely
                // ambiguous - deliberately resolved this way rather than
                // attempting a per-paragraph mapping, since the counts can differ.
                Word.Range newRange = doc.Range(paragraphs[startIndex + 1].Range.Start, range.End);
                foreach (Word.Paragraph p in newRange.Paragraphs)
                {
                    p.Range.set_Style(capturedStyle);
                }
                preservedNote = $" (preserved '{capturedStyle}' style)";
            }
            else if (!preserveFormatting)
            {
                preservedNote = " (formatting stripped, preserveFormatting:false)";
            }

            return new ToolResult { Output = $"Replaced paragraphs {startIndex}-{endIndex} with: {text}{preservedNote}", Mutated = true, Summary = "replace_blocks" };
        }

    }
}

