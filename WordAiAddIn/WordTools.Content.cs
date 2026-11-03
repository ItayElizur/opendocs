using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    // Real COM tool execution against the live Word document, called from the
    // WebView2-hosted AgentLoop via the JSON WebMessage bridge.
    public static partial class WordTools
    {
        // Resolves a Range's 0-based paragraph index via a forward-marching
        // Paragraph.Next() chain instead of Word's slow positional Paragraphs[i]
        // lookup; callers must request positions in non-decreasing document order.
        // See WordTools.Content.cs.md.
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

        // Read-only search (never mutates, unlike find_replace). Plain-substring
        // queries use Word's native Find engine (cost proportional to match count,
        // not paragraph count); a regex:true query falls back to a per-paragraph
        // scan via the cheap Paragraph.Next() chain, since Word's Find has no regex
        // mode. See WordTools.Content.cs.md.
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

        private sealed class HeadingInfo
        {
            public int Index;   // 0-based paragraph index - the same one read_blocks uses
            public int Level;   // 1-9
            public string Text;
        }

        // Navigation-Pane-style outline of the Heading-styled paragraphs, found via Word's own wdGoToHeading jump (the
        // index behind the Navigation Pane) rather than by scanning every paragraph's style. Returns at most `cap`
        // headings (0 = all); `more` says the document has further ones. Used by get_headings and by the outline in
        // get_document_context. See WordTools.Content.cs.md.
        private static List<HeadingInfo> ReadHeadings(Word.Document doc, int cap, out bool more)
        {
            var list = new List<HeadingInfo>();
            more = false;

            // The jump below only finds headings AFTER the cursor, so a heading in the very first paragraph has to be
            // picked up separately (a document that opens with its title heading is the usual case).
            Word.Paragraph first = doc.Paragraphs.First;
            if (first.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText) list.Add(HeadingOf(first, 0));

            // A heading's paragraph index is the previous heading's index plus the paragraphs in between. Counting a
            // range's paragraphs is one cheap call; the earlier approach marched Paragraph.Next() from the top of the
            // document to each heading (about 3-4 ms per paragraph passed).
            int prevStart = 0, prevIndex = 0;
            Word.Range cursor = doc.Content;
            cursor.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            // wdGoToHeading wraps back to the first heading once it runs out, rather than signaling "no more" (detected
            // via next.Start <= cursor.Start below). safetyLimit is a defensive backstop against an uncaught wrap.
            int safetyLimit = doc.Paragraphs.Count + 1;
            for (int i = 0; i < safetyLimit; i++)
            {
                if (cap > 0 && list.Count > cap) break; // one past the cap is enough to know there is more
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

                int start = next.Start;
                int index = prevIndex + doc.Range(prevStart, start).Paragraphs.Count;
                list.Add(HeadingOf(next.Paragraphs[1], index));
                prevStart = start;
                prevIndex = index;

                cursor = next.Duplicate;
                cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                cursor.Move(Word.WdUnits.wdCharacter, 1);
            }

            if (cap > 0 && list.Count > cap)
            {
                more = true;
                list.RemoveRange(cap, list.Count - cap);
            }
            return list;
        }

        private static HeadingInfo HeadingOf(Word.Paragraph para, int index)
        {
            int level = (int)para.OutlineLevel;
            if (level < 1 || level > 9)
            {
                // Not an outline-level paragraph after all: fall back to the digits in the style name ("Heading 3").
                string styleName = para.Range.get_Style().NameLocal;
                string digits = new string(styleName.Where(char.IsDigit).ToArray());
                level = int.TryParse(digits, out int lvl) && lvl >= 1 && lvl <= 9 ? lvl : 1;
            }
            return new HeadingInfo { Index = index, Level = level, Text = para.Range.Text.TrimEnd('\r', '\a', '\n') };
        }

        private static ToolResult GetHeadings()
        {
            bool more;
            List<HeadingInfo> headings = ReadHeadings(ActiveDoc, 0, out more);
            var sb = new System.Text.StringBuilder();
            foreach (HeadingInfo h in headings) sb.AppendLine($"[{h.Index}] H{h.Level}: {h.Text}");
            return new ToolResult { Output = headings.Count > 0 ? sb.ToString() : "No headings in this document.", Summary = "get_headings" };
        }

        // Shorter than it used to be (the opening text is the least useful part); the outline below it is the map.
        private const int ContextPreviewChars = 150;
        private const int ContextOutlineHeadings = 25;
        private const int ContextOutlineTextChars = 80;

        private static ToolResult GetDocumentContext()
        {
            Word.Document doc = ActiveDoc;
            int paraCount = doc.Paragraphs.Count;
            int wordCount = doc.Words.Count;
            // Only the opening characters are needed - not the whole document's text.
            string preview = doc.Range(0, Math.Min(doc.Content.End, ContextPreviewChars + 1)).Text ?? "";
            if (preview.Length > ContextPreviewChars) preview = preview.Substring(0, ContextPreviewChars) + "...";
            string output = $"Paragraphs: {paraCount}, Words: {wordCount}\nPreview: {preview}";

            bool more;
            List<HeadingInfo> headings = ReadHeadings(doc, ContextOutlineHeadings, out more);
            if (headings.Count > 0)
            {
                var outline = new System.Text.StringBuilder();
                outline.Append("Outline (" + (more ? "first " + headings.Count + " headings; get_headings lists all" : "all headings") + "):");
                foreach (HeadingInfo h in headings)
                {
                    string text = h.Text.Length > ContextOutlineTextChars ? h.Text.Substring(0, ContextOutlineTextChars) + "..." : h.Text;
                    outline.Append("\n[" + h.Index + "] H" + h.Level + ": " + text);
                }
                output += "\n" + outline;
            }

            string objects = ObjectCountsLine(doc);
            if (objects.Length > 0) output += "\n" + objects;
            return new ToolResult { Output = output, Summary = "read document context" };
        }

        // Positional insert, plain-text or restricted HTML. Backward compatible:
        // {text} with no afterBlockIndex/html keeps the original end-of-document
        // behavior, since existing prompts/history depend on it.
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

        // 'html' mode emits the same restricted subset InsertHtmlFragment accepts, so
        // read_blocks -> replace_blocks/insert_content round-trips (anything outside
        // the subset, e.g. font color, doesn't survive). Capped well below text mode
        // since a per-word COM property read per paragraph is far slower than a flat
        // Range.Text read.
        private const int HtmlReadBlocksMaxParagraphs = 100;

        // Chosen conservatively (not benchmarked like the html-mode cap above).
        // See WordTools.Content.cs.md.
        private const int TextReadBlocksMaxParagraphs = 1000;

        private const int MaxObjectLabelsPerBlock = 8;

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

            // Walks forward via Paragraph.Next(), not positional indexing - same
            // O(n^2) trap as ParagraphIndexResolver above (and find_text/get_headings/
            // ResolveTargetParagraphs).
            Word.Paragraph p = paragraphs.First;
            for (int skip = 0; skip < startIndex && p != null; skip++) p = p.Next();

            // Objects (tables, shapes, footnotes, ...) anchored in a block are listed after it as
            // "{table#0, shape "Text Box 3" textbox hasText}". Only the objects inside the requested paragraphs are
            // collected (Word's range-scoped collections, a few ms; a document-wide pass cost seconds on shape-heavy
            // documents) and merged during this forward walk; the per-paragraph Range reads are skipped entirely when
            // the range holds none. The end of the range is one positional lookup, the walk itself stays forward-only.
            List<ObjectAnchor> anchors = p == null
                ? new List<ObjectAnchor>()
                : CollectObjectAnchorsInRange(doc, p.Range.Start, paragraphs[endIndex + 1].Range.End);
            int nextAnchor = 0;

            for (int i = startIndex; i <= endIndex && p != null; i++)
            {
                // finally guarantees p advances exactly once per iteration even
                // though several branches `continue` early.
                try
                {
                    string objs = "";
                    if (anchors.Count > 0)
                    {
                        Word.Range pr = p.Range;
                        objs = ObjectAnchors.LabelsFor(anchors, ref nextAnchor, pr.Start, pr.End, MaxObjectLabelsPerBlock);
                    }

                    if (format != "html")
                    {
                        string plain = p.Range.Text.TrimEnd('\r', '\a', '\n');
                        sb.AppendLine($"[{i}] {plain}{objs}");
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
                        sb.AppendLine($"[{i}] <li>{content}</li>{objs}");
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
                    sb.AppendLine($"[{i}] <{tag}>{content}</{tag}>{objs}");
                }
                finally
                {
                    p = p.Next();
                }
            }
            if (format == "html" && openList != null) sb.AppendLine($"</{openList}>");

            return new ToolResult { Output = sb.ToString(), Summary = "read_blocks" };
        }

        // Non-destructive replace: preserves the first replaced paragraph's style by
        // default (instead of silently stripping heading/list identity), plus an
        // html alternative to text.
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
                // Reapply the FIRST replaced paragraph's style to every resulting
                // paragraph - multi-paragraph replacements with differing source
                // styles are genuinely ambiguous, so a per-paragraph mapping isn't attempted.
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

