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
    public static class WordTools
    {
        // Task 11 (per-document since PP-1): server-side editing-mode gate,
        // keyed by the same per-document id TaskPaneHost.GetChatId() produces,
        // so a mode change in one window's pane never affects another
        // window's. Absent key defaults to FullAutonomy so existing behavior
        // is unchanged for a document that has never had its mode set.
        private static readonly Dictionary<string, EditingMode> ModeByDoc = new Dictionary<string, EditingMode>();

        public static void SetMode(string docKey, EditingMode mode)
        {
            ModeByDoc[docKey] = mode;
        }

        private static EditingMode ModeFor(string docKey)
        {
            EditingMode m;
            return ModeByDoc.TryGetValue(docKey, out m) ? m : EditingMode.FullAutonomy;
        }

        // Tools that are always safe to run regardless of editing mode - they
        // never touch document content. Everything else is gated below.
        private static readonly HashSet<string> AlwaysAllowedTools = new HashSet<string>
        {
            "get_document_context", "read_blocks", "read_chart", "read_table", "read_smartart",
        };

        public static ToolResult Execute(string docKey, string name, JsonElement input)
        {
            // Post-hoc diagnostic addition (2026-08-24): every tool call and
            // every top-level failure is logged, regardless of which tool -
            // this is the single catch-all that guarantees a repro always
            // shows up in the log even if a specific method's own logging
            // (WriteChartData/ReadChart/etc.) missed the actual failure point.
            DebugLog.Write("Execute: " + name + " input=" + input.GetRawText());
            try
            {
                EditingMode mode = ModeFor(docKey);
                bool isAlwaysAllowed = AlwaysAllowedTools.Contains(name);
                bool isAddComment = name == "add_comment";
                // "Mutating" here means "changes document content/structure" -
                // used only to decide whether TrackRevisions should be toggled.
                // add_comment does not mutate document content (it adds a
                // comment annotation), so it's excluded from this set even
                // though it is gated like a mutating tool below.
                bool isContentMutating = !isAlwaysAllowed && !isAddComment;

                if (mode == EditingMode.ReadOnly && !isAlwaysAllowed)
                {
                    return new ToolResult { Output = "Blocked: editing mode is Read Only.", IsError = true, Summary = name };
                }
                if (mode == EditingMode.CommentOnly && !isAlwaysAllowed && !isAddComment)
                {
                    return new ToolResult { Output = "Blocked: editing mode is Comment Only - use add_comment instead of editing content directly.", IsError = true, Summary = name };
                }

                if (isContentMutating)
                {
                    ActiveDoc.TrackRevisions = (mode == EditingMode.TrackChanges);
                }

                switch (name)
                {
                    case "get_document_context":
                        return GetDocumentContext();
                    case "insert_content":
                        return InsertContent(input);
                    case "edit_chart":
                        return EditChart(input);
                    case "read_chart":
                        return ReadChart(input);
                    case "add_table":
                        return AddTable(input);
                    case "edit_table":
                        return EditTable(input);
                    case "read_table":
                        return ReadTable(input);
                    case "add_smartart":
                        return AddSmartArt(input);
                    case "edit_smartart":
                        return EditSmartArt(input);
                    case "read_smartart":
                        return ReadSmartArt(input);
                    case "read_blocks":
                        return ReadBlocks(input);
                    case "find_text":
                        return FindText(input);
                    case "get_headings":
                        return GetHeadings();
                    case "replace_blocks":
                        return ReplaceBlocks(input);
                    case "apply_commands":
                        return ApplyCommands(input);
                    case "add_comment":
                        return AddComment(input);
                    case "add_image":
                        return AddImage(input);
                    default:
                        return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("Execute: " + name, ex);
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        private static ToolResult AddComment(JsonElement input)
        {
            string anchorText = input.GetProperty("anchorText").GetString();
            string commentText = input.GetProperty("commentText").GetString();
            Word.Document doc = ActiveDoc;
            Word.Range range = doc.Content;
            range.Find.ClearFormatting();
            range.Find.Text = anchorText;
            bool found = range.Find.Execute();
            if (!found)
            {
                return new ToolResult { Output = $"Could not find text to anchor comment: '{anchorText}'", IsError = true, Summary = "add_comment" };
            }
            doc.Comments.Add(range, commentText);
            return new ToolResult { Output = "Comment added.", Mutated = true, Summary = "add_comment" };
        }

        // Known limitation (PP-1 Task 5 Step 5): resolves whichever document is
        // ACTIVE right now, not necessarily the one whose pane initiated this
        // tool call. A tool call is always initiated by a user in the
        // currently-focused window, so this is normally correct - but a
        // long-running run whose user switches windows mid-run would write
        // into the newly-active document instead of the one the run started
        // against. Fixing this needs per-document COM target resolution
        // across every executor method - out of scope here; left as a known
        // issue for a follow-up item.
        private static Word.Document ActiveDoc => Globals.ThisAddIn.Application.ActiveDocument;

        // PP-10 Task 1: shared insertion-point resolver. afterBlockIndex is
        // 0-based over ActiveDoc.Paragraphs, matching every other
        // block-addressed tool in this file. -1 means "before the first
        // paragraph", matching insertToc/moveBlocks' existing convention.
        // Consumed by insert_content, apply_commands' chart anchoring (PP-9),
        // and add_image (PP-11) - one helper, not three copies.
        private static Word.Range RangeAfterBlock(int afterBlockIndex)
        {
            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            if (afterBlockIndex < -1 || afterBlockIndex > paragraphs.Count - 1)
                throw new ArgumentOutOfRangeException("afterBlockIndex",
                    "afterBlockIndex must be between -1 and " + (paragraphs.Count - 1) + ".");
            if (afterBlockIndex == -1)
            {
                Word.Range start = paragraphs[1].Range;
                start.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                return start;
            }
            Word.Range r = paragraphs[afterBlockIndex + 1].Range;
            r.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            return r;
        }

        // PP-23: shared end-of-document insertion point, matching InsertContent's
        // and AddImage's existing inline "collapse doc.Content to the end" idiom
        // (extracted here rather than duplicated a third time).
        private static Word.Range EndOfDocumentRange()
        {
            Word.Range end = ActiveDoc.Content;
            end.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            return end;
        }

        // Chart-type vocabulary now lives in OfficeAi.Shared.ChartTypes -
        // one table for all three add-ins (PP-9's "one chart vocabulary"
        // intent, now enforced by construction rather than by comment).

        // PP-9: ported from PowerPointTools.AddChartPpt's data-writing block -
        // the embedded chart workbook MUST be closed and released in a
        // finally, or a leaked hidden Excel process stays alive for the rest
        // of the Word session. seriesArray items are {name?, values}.
        // Post-hoc fix (2026-08-24, user-reported): the embedded workbook's
        // OLE server occasionally still throws "The remote procedure call
        // failed" (HRESULT 0x800706BE) even after the Clear()+batched-write
        // fix above - a known, documented transient failure mode for rapid
        // COM calls against Office's embedded chart-data Excel object, not
        // something a single call can eliminate. Retrying after a short
        // delay is the standard mitigation; only the specific known
        // transient RPC HRESULTs are retried, so a genuine logic error
        // (bad range, etc.) still fails immediately rather than being
        // masked for 3 attempts.
        private static readonly int[] TransientComHResults =
        {
            unchecked((int)0x800706BE), // RPC_S_CALL_FAILED - "The remote procedure call failed."
            unchecked((int)0x8001010A), // RPC_E_SERVERCALL_RETRYLATER - "The message filter indicated that the application is busy."
            unchecked((int)0x800706BA), // RPC_S_SERVER_UNAVAILABLE
        };

        // Post-hoc diagnostic addition (2026-08-24): every attempt (success,
        // retried failure, AND a non-retried failure) is logged via
        // DebugLog - this is the thing that finally shows the REAL exception
        // detail (type/HResult/message/stack) from a live repro, instead of
        // guessing further. `label` identifies which call site this is, for
        // when multiple charts/reads happen in one session.
        private static void RetryTransientCom(Action action, string label = "RetryTransientCom")
        {
            const int maxAttempts = 3;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    DebugLog.Write(label + ": attempt " + attempt + " starting");
                    action();
                    DebugLog.Write(label + ": attempt " + attempt + " SUCCEEDED");
                    return;
                }
                catch (Exception ex) when (attempt < maxAttempts && Array.IndexOf(TransientComHResults, ex.HResult) >= 0)
                {
                    DebugLog.WriteException(label + ": attempt " + attempt + " (transient, retrying)", ex);
                    System.Threading.Thread.Sleep(200 * attempt);
                }
                catch (Exception ex)
                {
                    // Either the last attempt, or an HResult not in the
                    // transient list - logged before rethrowing so the real
                    // failure is captured even when no more retries happen.
                    DebugLog.WriteException(label + ": attempt " + attempt + " (NOT retried - rethrowing)", ex);
                    throw;
                }
            }
        }

        private static void WriteChartData(dynamic chart, List<string> categories, JsonElement seriesArray)
        {
            var seriesList = seriesArray.EnumerateArray().ToList();
            DebugLog.Write("WriteChartData: ENTER, categories=" + categories.Count + " series=" + seriesList.Count);
            int expectedLen = categories.Count > 0 ? categories.Count : (seriesList.Count > 0 ? seriesList[0].GetProperty("values").GetArrayLength() : 0);
            if (categories.Count == 0)
            {
                for (int i = 1; i <= expectedLen; i++) categories.Add(i.ToString());
            }
            foreach (JsonElement s in seriesList)
            {
                int len = s.GetProperty("values").GetArrayLength();
                if (len != categories.Count)
                    throw new ArgumentException("edit_chart: series '" + (s.TryGetProperty("name", out var nm) ? nm.GetString() : "") +
                                                "' has " + len + " value(s) but there are " + categories.Count + " categor" + (categories.Count == 1 ? "y" : "ies") + " - every series must match the category count.");
            }

            // Post-hoc fix (2026-08-24, code-review finding while adding
            // diagnostics): chart.ChartData.Workbook was fetched OUTSIDE the
            // RetryTransientCom-protected block, so if THIS specific call is
            // the flaky one (plausible under the "OLE server not fully live
            // yet" hypothesis - it is the very first COM call that opens the
            // embedded object), the retry wrapper never got a chance to help
            // at all. Moved inside the lambda so every attempt re-opens it
            // fresh; declared here (nullable) so `finally` can still clean up
            // whichever attempt actually succeeded.
            dynamic dataWorkbook = null;
            try
            {
                // Build the whole grid in memory up front (pure C#, no COM) -
                // only the write itself needs to go through RetryTransientCom.
                int rowCount = categories.Count + 1; // +1 header row
                int colCount = seriesList.Count + 1; // +1 category column
                object[,] grid = new object[rowCount, colCount];
                grid[0, 0] = "";
                int colIdx = 0;
                foreach (JsonElement s in seriesList)
                {
                    string name = s.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String ? nameEl.GetString() : "Series " + (colIdx + 1);
                    grid[0, colIdx + 1] = name;
                    colIdx++;
                }
                for (int r = 0; r < categories.Count; r++)
                {
                    grid[r + 1, 0] = categories[r];
                }
                colIdx = 0;
                foreach (JsonElement s in seriesList)
                {
                    int r = 0;
                    foreach (JsonElement v in s.GetProperty("values").EnumerateArray())
                    {
                        grid[r + 1, colIdx + 1] = v.GetDouble();
                        r++;
                    }
                    colIdx++;
                }

                RetryTransientCom(() =>
                {
                    // Post-hoc fix (2026-08-24, user-reported the RPC failure
                    // recurring even after the first fix): a brief settle
                    // delay immediately after the embedded OLE workbook is
                    // opened, before the first COM call against it. This is a
                    // documented mitigation for this exact class of embedded-
                    // chart-data-workbook flakiness - the automation surface
                    // is not always fully live the instant ChartData.Workbook
                    // returns. Cheap (one UI-thread sleep) relative to the
                    // cost of a failed/retried chart creation.
                    System.Threading.Thread.Sleep(120);

                    DebugLog.Write("WriteChartData: getting chart.ChartData.Workbook");
                    dataWorkbook = chart.ChartData.Workbook;

                    DebugLog.Write("WriteChartData: getting Worksheets[1]");
                    dynamic sheet = dataWorkbook.Worksheets[1];

                    // Confirmed repro: a brand-new chart's embedded workbook comes
                    // pre-seeded by Word/Office with placeholder sample data (a
                    // default chart template, commonly 4 categories x 3 series).
                    // Without clearing it first, only the cells the NEW data
                    // actually occupies get overwritten - any leftover placeholder
                    // cells beyond that extent stay in the sheet and get plotted
                    // alongside the real data, producing phantom extra
                    // categories/series the user never asked for.
                    DebugLog.Write("WriteChartData: Cells.Clear()");
                    sheet.Cells.Clear();

                    DebugLog.Write("WriteChartData: writing " + rowCount + "x" + colCount + " grid via Resize+Value2");
                    dynamic topLeft = sheet.Cells[1, 1];
                    dynamic writeRange = topLeft.Resize[rowCount, colCount];
                    writeRange.Value2 = grid;

                    // ACTUAL ROOT CAUSE (2026-08-24, confirmed via .NET
                    // reflection against the real referenced
                    // Microsoft.Office.Interop.Word.dll, not a guess):
                    // Word.Chart.SetSourceData's real signature is
                    // SetSourceData(String Source, Object PlotBy) - the
                    // first parameter is a STRING, not a Range at all. Every
                    // prior attempt (round 1's sheet.Range(topLeft,
                    // bottomRight), round 2's reused writeRange, and this
                    // round's sheet.Range[a1Range]) was passing a Range COM
                    // object where the method actually expects a string -
                    // "Could not convert argument 0" was ALWAYS this type
                    // mismatch, not a marshaling-path quirk. The correct
                    // call passes a plain "SheetName!A1:B4"-style reference
                    // string - no Range object needed at all.
                    string a1Range = "A1:" + TextUtil.ColumnLetter(colCount) + rowCount;
                    string sourceRef = (string)sheet.Name + "!" + a1Range;
                    DebugLog.Write("WriteChartData: SetSourceData(\"" + sourceRef + "\")");
                    chart.SetSourceData(sourceRef);
                    DebugLog.Write("WriteChartData: SetSourceData returned OK");
                }, "WriteChartData");
            }
            finally
            {
                // ROOT CAUSE FOUND (2026-08-24, via DebugLog): this cleanup
                // previously had no catch of its own - when SetSourceData
                // failed above (see the real bug this block is next to), the
                // chart/embedded-workbook was left in a state where
                // dataWorkbook.Close() ALSO threw (a real, observed
                // RPC_E_DISCONNECTED). In C#, an exception thrown from a
                // `finally` block while another exception is already
                // propagating from the `try` block REPLACES it - so the
                // user only ever saw this cleanup-time exception
                // ("The object invoked has disconnected from its clients"),
                // never the real SetSourceData ArgumentException that caused
                // it. This is exactly why two prior rounds of fixes,
                // diagnosing from the user's reported error text alone,
                // chased the wrong theory. Cleanup failures are now caught
                // and logged here instead of being allowed to propagate and
                // mask whatever real exception is already in flight.
                if (dataWorkbook != null)
                {
                    try
                    {
                        dataWorkbook.Close(SaveChanges: true);
                    }
                    catch (Exception closeEx)
                    {
                        DebugLog.WriteException("WriteChartData: cleanup Close() failed (secondary - not masking the real exception)", closeEx);
                    }
                    finally
                    {
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(dataWorkbook); }
                        catch (Exception releaseEx) { DebugLog.WriteException("WriteChartData: cleanup ReleaseComObject() failed (secondary)", releaseEx); }
                    }
                }
            }
        }

        // PP-10 Task 3: restricted-HTML insertion. Supported set, fixed and
        // small: block <p> <h1>-<h3> <ul>/<ol> with <li>; inline <b>/<strong>
        // <i>/<em> <u> <br>. Nothing else - no tables, no images (PP-11), no
        // attributes, no nested lists. Parsed via XElement.Parse (built on
        // XmlReader) rather than a regex/hand scanner - gives well-formedness
        // checking for free, so a malformed fragment throws before anything
        // is written. The whole fragment is validated against the supported
        // tag set BEFORE any Word write happens, so an unsupported tag
        // halfway through cannot leave a partial insert.
        private static readonly HashSet<string> HtmlBlockTags = new HashSet<string> { "p", "h1", "h2", "h3", "ul", "ol" };
        private static readonly HashSet<string> HtmlInlineTags = new HashSet<string> { "b", "strong", "i", "em", "u", "br" };

        private static System.Xml.Linq.XElement ParseHtmlFragment(string html)
        {
            // Normalize the most likely void-element mistake before parsing,
            // rather than failing the call over it.
            string normalized = System.Text.RegularExpressions.Regex.Replace(
                html, "<br\\s*>", "<br/>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            System.Xml.Linq.XElement root;
            try
            {
                root = System.Xml.Linq.XElement.Parse("<root>" + normalized + "</root>");
            }
            catch (System.Xml.XmlException ex)
            {
                throw new ArgumentException("Malformed HTML fragment (must be well-formed XHTML - closed tags, <br/> not <br>): " + ex.Message);
            }
            ValidateHtmlTags(root, true);
            return root;
        }

        private static void ValidateHtmlTags(System.Xml.Linq.XElement el, bool isBlockContext)
        {
            foreach (System.Xml.Linq.XElement child in el.Elements())
            {
                string tag = child.Name.LocalName.ToLowerInvariant();
                if (isBlockContext)
                {
                    if (tag == "li") continue; // only valid directly inside ul/ol, checked below
                    if (!HtmlBlockTags.Contains(tag))
                        throw new ArgumentException("Unsupported HTML tag '<" + tag + ">'. Supported: " +
                            string.Join(", ", HtmlBlockTags) + ", li (inside ul/ol), " + string.Join(", ", HtmlInlineTags) + ".");
                    if (tag == "ul" || tag == "ol")
                    {
                        foreach (System.Xml.Linq.XElement liOrOther in child.Elements())
                        {
                            if (liOrOther.Name.LocalName.ToLowerInvariant() != "li")
                                throw new ArgumentException("<" + tag + "> may only contain <li> children, found <" + liOrOther.Name.LocalName + ">.");
                            ValidateHtmlTags(liOrOther, false);
                        }
                    }
                    else
                    {
                        ValidateHtmlTags(child, false);
                    }
                }
                else
                {
                    if (!HtmlInlineTags.Contains(tag))
                        throw new ArgumentException("Unsupported HTML tag '<" + tag + ">' in inline content. Supported inline: " +
                            string.Join(", ", HtmlInlineTags) + ".");
                    if (tag != "br") ValidateHtmlTags(child, false);
                }
            }
        }

        // Writes one paragraph's inline content (text + b/strong/i/em/u/br)
        // into `cursor`, which must be collapsed at the start of an empty
        // paragraph. Word.Range is a COM reference type - Collapse/Text
        // mutate the same underlying range the caller holds, so no ref
        // parameter is needed for the recursion to see the cursor advance.
        private static void WriteInlineNodes(Word.Range cursor, IEnumerable<System.Xml.Linq.XNode> nodes, bool bold, bool italic, bool underline)
        {
            foreach (System.Xml.Linq.XNode node in nodes)
            {
                System.Xml.Linq.XText textNode = node as System.Xml.Linq.XText;
                if (textNode != null)
                {
                    string text = textNode.Value;
                    if (text.Length == 0) continue;
                    cursor.Text = text;
                    cursor.Font.Bold = bold ? 1 : 0;
                    cursor.Font.Italic = italic ? 1 : 0;
                    cursor.Font.Underline = underline ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone;
                    cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    continue;
                }
                System.Xml.Linq.XElement el = node as System.Xml.Linq.XElement;
                if (el == null) continue;
                string tag = el.Name.LocalName.ToLowerInvariant();
                if (tag == "br")
                {
                    // A soft line break within the same paragraph, not a new paragraph.
                    cursor.InsertBreak(Word.WdBreakType.wdLineBreak);
                    cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    continue;
                }
                WriteInlineNodes(cursor, el.Nodes(),
                    bold || tag == "b" || tag == "strong",
                    italic || tag == "i" || tag == "em",
                    underline || tag == "u");
            }
        }

        // Inserts a validated HTML fragment (see ParseHtmlFragment) starting
        // at `at`. Each block element becomes its own new paragraph, using
        // the same InsertParagraphAfter+collapse idiom InsertContent already
        // uses for plain text - the paragraph `at` itself pointed into is
        // never merged into, only new paragraphs after it are created.
        private static void InsertHtmlFragment(Word.Range at, string html)
        {
            System.Xml.Linq.XElement root = ParseHtmlFragment(html);
            Word.Range cursor = at;
            foreach (System.Xml.Linq.XElement block in root.Elements())
            {
                string tag = block.Name.LocalName.ToLowerInvariant();
                if (tag == "ul" || tag == "ol")
                {
                    foreach (System.Xml.Linq.XElement li in block.Elements())
                    {
                        cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        cursor.InsertParagraphAfter();
                        cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        Word.Paragraph para = cursor.Paragraphs[1];
                        WriteInlineNodes(cursor, li.Nodes(), false, false, false);
                        if (tag == "ul") para.Range.ListFormat.ApplyBulletDefault();
                        else para.Range.ListFormat.ApplyNumberDefault();
                    }
                }
                else
                {
                    cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    cursor.InsertParagraphAfter();
                    cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    Word.Paragraph para = cursor.Paragraphs[1];
                    WriteInlineNodes(cursor, block.Nodes(), false, false, false);
                    if (tag.Length == 2 && tag[0] == 'h')
                    {
                        para.Range.set_Style("Heading " + tag[1]);
                    }
                }
            }
        }

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

        // dynamic: Word's chart object model (Shapes.AddChart2 / Chart / SeriesCollection) mirrors
        // Excel/PowerPoint's shared chart engine; using dynamic avoids pinning down the exact
        // Interop type names for this spike and lets any signature mismatch surface immediately at
        // runtime instead of guessing overloads at compile time.
        //
        // PP-9: create-or-edit against an explicit list of ALL charts (inline
        // first, then floating - see Task 4 Step 4), addressed by chartIndex,
        // with real categories/named multi-series/chart-type support ported
        // from PowerPointTools.AddChartPpt.
        // Every chart shape, inline first then floating, in that fixed order
        // so chartIndex is predictable across calls (PP-9 Task 4 Step 4).
        // Shared by EditChart and ReadChart so both address charts identically.
        // internal (not private): post-hoc fix (2026-08-24, user-reported)
        // needs this same addressing from TaskPaneHost.OnSelectionChanged, so
        // a selected chart shape can be reported with the SAME chartIndex
        // edit_chart/read_chart would use, rather than a second, possibly
        // drifting copy of this resolution logic.
        internal static List<dynamic> ListChartShapes(dynamic doc)
        {
            var chartShapes = new List<dynamic>();
            foreach (dynamic shp in doc.InlineShapes)
            {
                try { if ((int)shp.HasChart == -1 /* msoTrue */) chartShapes.Add(shp); } catch { }
            }
            foreach (dynamic shp in doc.Shapes)
            {
                if ((int)shp.HasChart == -1 /* msoTrue */) chartShapes.Add(shp);
            }
            return chartShapes;
        }

        // Lets the model inspect an existing chart's current title/type/
        // categories/series before deciding what to change via edit_chart -
        // without this, an incremental edit (e.g. "remove one category") has
        // no way to know what the other categories/series currently are,
        // since edit_chart REPLACES the whole dataset rather than patching
        // it. Reads from the chart's embedded workbook (the same object
        // WriteChartData writes to) via the same Cells/UsedRange/.Value2
        // pattern already proven working by the write side, rather than the
        // Series.Values/.XValues COM properties directly (whose exact
        // marshaled array shape in this dynamic context is not something
        // this environment can verify without a live Word session).
        private static ToolResult ReadChart(JsonElement input)
        {
            dynamic doc = ActiveDoc;
            var chartShapes = ListChartShapes(doc);
            if (chartShapes.Count == 0)
                return new ToolResult { Output = "No charts in this document.", Summary = "read_chart" };

            int chartIndex = input.TryGetProperty("chartIndex", out var ci) && ci.ValueKind == JsonValueKind.Number ? ci.GetInt32() : 0;
            if (chartIndex < 0 || chartIndex >= chartShapes.Count)
                throw new ArgumentOutOfRangeException("chartIndex",
                    "chartIndex must be between 0 and " + (chartShapes.Count - 1) + " (" + chartShapes.Count + " chart(s) in the document).");

            dynamic chart = chartShapes[chartIndex].Chart;

            string title = (bool)chart.HasTitle ? (string)chart.ChartTitle.Text : null;
            int typeCode = (int)chart.ChartType;
            string typeName = null;
            foreach (var kv in ChartTypes.ByName) { if (kv.Value == typeCode) { typeName = kv.Key; break; } }

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Chart " + chartIndex + " of " + chartShapes.Count + " (pass this index to edit_chart to target it):");
            sb.AppendLine("Title: " + (title ?? "(none)"));
            sb.AppendLine("Type: " + (typeName ?? ("unrecognized chart type code " + typeCode)));

            DebugLog.Write("ReadChart: ENTER, chartIndex=" + chartIndex);
            // Post-hoc fix (2026-08-24, same code-review finding as
            // WriteChartData): ChartData.Workbook moved inside the retry
            // lambda so a flaky OPEN, not just a flaky subsequent call, also
            // gets retried.
            dynamic dataWorkbook = null;
            try
            {
                // Post-hoc fix (2026-08-24, user-reported "read chart still
                // doesn't work"): same settle-delay + retry protection as the
                // write path (WriteChartData) - opening the embedded OLE
                // workbook via ChartData.Workbook is not guaranteed to be
                // immediately ready for automation calls.
                RetryTransientCom(() =>
                {
                    System.Threading.Thread.Sleep(120);
                    DebugLog.Write("ReadChart: getting chart.ChartData.Workbook");
                    dataWorkbook = chart.ChartData.Workbook;
                    DebugLog.Write("ReadChart: getting Worksheets[1]/UsedRange");
                    dynamic sheet = dataWorkbook.Worksheets[1];
                    dynamic usedRange = sheet.UsedRange;
                    int rowCount = (int)usedRange.Rows.Count;
                    int colCount = (int)usedRange.Columns.Count;

                    if (rowCount < 2 || colCount < 2)
                    {
                        sb.AppendLine("No data (empty chart).");
                        return;
                    }

                    // Excel COM Range.Value2 returns a 1-based 2D array for a
                    // multi-cell range (well-established Excel Interop
                    // behavior) - read the actual bounds rather than assume
                    // 0 or 1, so this is correct either way.
                    object[,] grid = (object[,])usedRange.Value2;
                    int rowLb = grid.GetLowerBound(0), rowUb = grid.GetUpperBound(0);
                    int colLb = grid.GetLowerBound(1), colUb = grid.GetUpperBound(1);

                    var categories = new List<string>();
                    for (int r = rowLb + 1; r <= rowUb; r++) categories.Add(Convert.ToString(grid[r, colLb]));
                    sb.AppendLine("Categories (" + categories.Count + "): " + string.Join(", ", categories));

                    for (int c = colLb + 1; c <= colUb; c++)
                    {
                        string seriesName = Convert.ToString(grid[rowLb, c]);
                        var values = new List<string>();
                        for (int r = rowLb + 1; r <= rowUb; r++) values.Add(Convert.ToString(grid[r, c]));
                        sb.AppendLine("Series " + (c - colLb - 1) + " \"" + seriesName + "\": " + string.Join(", ", values));
                    }
                }, "ReadChart");
            }
            finally
            {
                // Same exception-masking fix as WriteChartData's finally block.
                if (dataWorkbook != null)
                {
                    try { dataWorkbook.Close(SaveChanges: false); }
                    catch (Exception closeEx) { DebugLog.WriteException("ReadChart: cleanup Close() failed (secondary - not masking the real exception)", closeEx); }
                    finally
                    {
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(dataWorkbook); }
                        catch (Exception releaseEx) { DebugLog.WriteException("ReadChart: cleanup ReleaseComObject() failed (secondary)", releaseEx); }
                    }
                }
            }

            DebugLog.Write("ReadChart: EXIT ok");
            return new ToolResult { Output = sb.ToString().TrimEnd(), Summary = "read_chart" };
        }

        private static ToolResult EditChart(JsonElement input)
        {
            DebugLog.Write("EditChart: ENTER input=" + input.GetRawText());
            dynamic doc = ActiveDoc;
            var chartShapes = ListChartShapes(doc);

            bool createRequested = input.TryGetProperty("create", out var cr) && cr.ValueKind == JsonValueKind.True;
            int chartIndex = input.TryGetProperty("chartIndex", out var ci) && ci.ValueKind == JsonValueKind.Number ? ci.GetInt32() : 0;

            dynamic chartShape;
            bool created;
            if (createRequested || chartShapes.Count == 0)
            {
                int typeCode = 51; // xlColumnClustered default
                if (input.TryGetProperty("chartType", out var ctEl) && ctEl.ValueKind == JsonValueKind.String)
                {
                    if (!ChartTypes.ByName.TryGetValue(ctEl.GetString(), out typeCode))
                        throw new ArgumentException("edit_chart: unknown chartType '" + ctEl.GetString() +
                                                    "'. Valid: " + string.Join(", ", ChartTypes.ByName.Keys) + ".");
                }

                if (input.TryGetProperty("afterBlockIndex", out var abEl) && abEl.ValueKind == JsonValueKind.Number)
                {
                    // Inline: flows with the text, which is what "add a chart
                    // after this paragraph" means. A floating shape at a fixed
                    // origin (the no-position path below) would overlap prose.
                    DebugLog.Write("EditChart: AddChart2 (anchored, afterBlockIndex=" + abEl.GetInt32() + ")");
                    Word.Range at = RangeAfterBlock(abEl.GetInt32());
                    dynamic floatingAtAnchor = doc.Shapes.AddChart2(-1, (Microsoft.Office.Core.XlChartType)typeCode, 0, 0, 300, 200, Anchor: at);
                    chartShape = floatingAtAnchor.ConvertToInlineShape();
                    DebugLog.Write("EditChart: AddChart2 (anchored) OK");
                }
                else
                {
                    // No position given: keep today's behavior exactly (floating
                    // shape at document origin) so existing calls do not move.
                    DebugLog.Write("EditChart: AddChart2 (floating, no position)");
                    chartShape = doc.Shapes.AddChart2(-1, (Microsoft.Office.Core.XlChartType)typeCode, 0, 0, 300, 200);
                    DebugLog.Write("EditChart: AddChart2 (floating) OK");
                }
                created = true;
            }
            else
            {
                if (chartIndex < 0 || chartIndex >= chartShapes.Count)
                    throw new ArgumentOutOfRangeException("chartIndex",
                        "chartIndex must be between 0 and " + (chartShapes.Count - 1) + " (" + chartShapes.Count + " chart(s) in the document).");
                chartShape = chartShapes[chartIndex];
                created = false;
            }

            dynamic chart = chartShape.Chart;

            // Type change before writing data - some type changes reset series formatting.
            if (input.TryGetProperty("chartType", out var chartTypeEl) && chartTypeEl.ValueKind == JsonValueKind.String)
            {
                int typeCode;
                if (!ChartTypes.ByName.TryGetValue(chartTypeEl.GetString(), out typeCode))
                    throw new ArgumentException("edit_chart: unknown chartType '" + chartTypeEl.GetString() +
                                                "'. Valid: " + string.Join(", ", ChartTypes.ByName.Keys) + ".");
                DebugLog.Write("EditChart: chart.ChartType = " + typeCode);
                chart.ChartType = (Microsoft.Office.Core.XlChartType)typeCode;
            }

            // Normalize the legacy single-series shorthand into `series` up
            // front, so WriteChartData only ever handles one shape.
            JsonElement seriesArray;
            bool hasSeries = input.TryGetProperty("series", out seriesArray) && seriesArray.ValueKind == JsonValueKind.Array;
            bool hasLegacyValues = input.TryGetProperty("values", out var legacyValues) && legacyValues.ValueKind == JsonValueKind.Array;
            var categories = new List<string>();
            if (input.TryGetProperty("categories", out var catsEl) && catsEl.ValueKind == JsonValueKind.Array)
                foreach (JsonElement c in catsEl.EnumerateArray()) categories.Add(c.GetString());

            int seriesCount = 0, categoryCount = categories.Count;
            if (hasSeries)
            {
                WriteChartData(chart, categories, seriesArray);
                seriesCount = seriesArray.GetArrayLength();
                categoryCount = categories.Count;
            }
            else if (hasLegacyValues)
            {
                // Build a synthetic one-element series array with no name,
                // matching legacy edit_chart({title, values}) behavior.
                using (JsonDocument synthetic = JsonDocument.Parse(
                    "[{\"values\":" + legacyValues.GetRawText() + "}]"))
                {
                    WriteChartData(chart, categories, synthetic.RootElement);
                }
                seriesCount = 1;
                categoryCount = categories.Count;
            }
            else if (created)
            {
                // A brand-new chart with no data at all would be created
                // blank/broken - seed a minimal default series, matching the
                // old hardcoded {1,2,3} fallback's intent of never leaving a
                // newly-created chart truly dataless.
                using (JsonDocument synthetic = JsonDocument.Parse("[{\"values\":[1,2,3]}]"))
                {
                    WriteChartData(chart, categories, synthetic.RootElement);
                }
                seriesCount = 1;
                categoryCount = categories.Count > 0 ? categories.Count : 3;
                hasLegacyValues = true; // so the result text reports the seeded data, not "data unchanged"
            }

            string title = null;
            if (input.TryGetProperty("title", out var titleEl) && titleEl.ValueKind == JsonValueKind.String)
            {
                title = titleEl.GetString();
                chart.HasTitle = true;
                chart.ChartTitle.Text = title;
            }

            // Recompute the resolved index by position rather than trusting
            // dynamic/COM RCW reference equality against the pre-creation
            // list - cheap, since a document's chart count is always small.
            int resolvedIndex = chartIndex;
            if (created)
            {
                int freshIdx = 0;
                foreach (dynamic shp in doc.InlineShapes)
                {
                    try { if ((int)shp.HasChart == -1) { if (shp == chartShape) resolvedIndex = freshIdx; freshIdx++; } } catch { }
                }
                foreach (dynamic shp in doc.Shapes)
                {
                    if ((int)shp.HasChart == -1) { if (shp == chartShape) resolvedIndex = freshIdx; freshIdx++; }
                }
            }

            string titlePart = title != null ? $"title='{title}'" : "title unchanged";
            // Only report series/category counts when data was actually
            // written this call - reporting "0 series" on a call that only
            // changed the title/type would misleadingly imply the existing
            // data was cleared.
            string dataPart = (hasSeries || hasLegacyValues) ? $", {seriesCount} series, {categoryCount} categories" : ", data unchanged";
            return new ToolResult
            {
                Output = $"Chart {(created ? "created" : "updated")} at chartIndex {resolvedIndex}: {titlePart}{dataPart}.",
                Mutated = true,
                Summary = "edit_chart",
            };
        }

        // PP-23 Task 1: 0-based at the tool boundary, matching every other
        // index in this file; Document.Tables is 1-based in COM. Document
        // order (no inline-vs-floating split needed - Word tables are always
        // flow content, unlike charts/SmartArt which can float).
        private static Word.Table ResolveTable(JsonElement input)
        {
            // Bug found via DebugLog from a real repro (2026-08-24): this
            // used GetProperty (required), throwing KeyNotFoundException
            // whenever tableIndex was omitted - directly contradicting this
            // tool's own documented "omit to target the first table"
            // behavior (and ReadTable's own correct TryGetProperty pattern,
            // right below in this same file). Fixed to match.
            int tableIndex = input.TryGetProperty("tableIndex", out var ti) && ti.ValueKind == JsonValueKind.Number ? ti.GetInt32() : 0;
            Word.Tables tables = ActiveDoc.Tables;
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentOutOfRangeException("tableIndex",
                    "tableIndex must be between 0 and " + (tables.Count - 1) + " (" + tables.Count + " table(s) in the document).");
            return tables[tableIndex + 1];
        }

        private static ToolResult AddTable(JsonElement input)
        {
            int rows = input.GetProperty("rows").GetInt32();
            int cols = input.GetProperty("cols").GetInt32();
            if (rows < 1 || cols < 1)
                throw new ArgumentException("add_table: rows and cols must each be at least 1.");

            int? afterBlockIndex = input.TryGetProperty("afterBlockIndex", out var abEl) && abEl.ValueKind == JsonValueKind.Number
                ? abEl.GetInt32() : (int?)null;
            Word.Range at = afterBlockIndex.HasValue ? RangeAfterBlock(afterBlockIndex.Value) : EndOfDocumentRange();

            Word.Table table = ActiveDoc.Tables.Add(at, rows, cols);

            if (input.TryGetProperty("cells", out var cells) && cells.ValueKind == JsonValueKind.Array)
            {
                int r = 0;
                foreach (JsonElement rowEl in cells.EnumerateArray())
                {
                    if (r >= rows) break; // ignore extra rows beyond the declared size rather than throwing mid-write
                    int c = 0;
                    foreach (JsonElement cellEl in rowEl.EnumerateArray())
                    {
                        if (c >= cols) break;
                        table.Cell(r + 1, c + 1).Range.Text = cellEl.GetString();
                        c++;
                    }
                    r++;
                }
            }

            int newIndex = ActiveDoc.Tables.Count - 1; // Tables.Add appends; stable immediately after the call
            return new ToolResult
            {
                Output = "Table added at index " + newIndex + " (" + rows + " rows x " + cols + " cols).",
                Mutated = true,
                Summary = "add_table",
            };
        }

        private static ToolResult EditTable(JsonElement input)
        {
            Word.Table table = ResolveTable(input);
            string kind = input.GetProperty("kind").GetString();
            switch (kind)
            {
                case "set_cell":
                {
                    int row = input.GetProperty("row").GetInt32();
                    int col = input.GetProperty("col").GetInt32();
                    if (row < 0 || row >= table.Rows.Count)
                        throw new ArgumentOutOfRangeException("row", "row must be between 0 and " + (table.Rows.Count - 1) + ".");
                    if (col < 0 || col >= table.Columns.Count)
                        throw new ArgumentOutOfRangeException("col", "col must be between 0 and " + (table.Columns.Count - 1) + ".");
                    table.Cell(row + 1, col + 1).Range.Text = input.GetProperty("text").GetString();
                    return new ToolResult { Output = "Cell [" + row + "," + col + "] updated.", Mutated = true, Summary = "edit_table" };
                }
                case "insert_row":
                case "delete_row":
                case "insert_col":
                case "delete_col":
                {
                    // Same index-always-existing, before/after-picks-side
                    // convention as PowerPoint's edit_table_structure - index
                    // always addresses an EXISTING row/column (0-based).
                    int index = input.GetProperty("index").GetInt32();
                    bool before = input.TryGetProperty("before", out var b) && b.ValueKind == JsonValueKind.True;
                    if (kind == "insert_row" || kind == "delete_row")
                    {
                        if (index < 0 || index >= table.Rows.Count)
                            throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Rows.Count - 1) + " for " + kind + ".");
                        if (kind == "insert_row") table.Rows.Add(table.Rows[before ? index + 1 : Math.Min(index + 2, table.Rows.Count + 1)]);
                        else table.Rows[index + 1].Delete();
                    }
                    else
                    {
                        if (index < 0 || index >= table.Columns.Count)
                            throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (table.Columns.Count - 1) + " for " + kind + ".");
                        if (kind == "insert_col") table.Columns.Add(table.Columns[before ? index + 1 : Math.Min(index + 2, table.Columns.Count + 1)]);
                        else table.Columns[index + 1].Delete();
                    }
                    return new ToolResult
                    {
                        Output = kind + " applied at index " + index + ". Row/column indices after this point have shifted - re-read the table before another structural edit in the same run.",
                        Mutated = true,
                        Summary = "edit_table",
                    };
                }
                case "set_style":
                {
                    if (input.TryGetProperty("styleName", out var styleEl) && styleEl.ValueKind == JsonValueKind.String)
                    {
                        try { table.set_Style(styleEl.GetString()); }
                        catch (Exception ex) { throw new ArgumentException("edit_table: '" + styleEl.GetString() + "' is not a valid table style name in this document/template. " + ex.Message); }
                    }
                    if (input.TryGetProperty("headerRow", out var hdr))
                        table.ApplyStyleHeadingRows = hdr.ValueKind == JsonValueKind.True;
                    if (input.TryGetProperty("bandedRows", out var band))
                        table.ApplyStyleRowBands = band.ValueKind == JsonValueKind.True;
                    // Post-hoc fix (2026-08-24, user-reported): this branch had
                    // no border support at all - "borders" is a real field on
                    // updateParagraphStyle elsewhere in this file, and the model
                    // (reasonably, given that precedent) called edit_table with
                    // the same field name expecting the same effect. Since
                    // set_style never checked for it, the call silently did
                    // nothing - a real gap, not a user error. table.Borders
                    // mirrors the Word.Border collection updateParagraphStyle
                    // already uses for paragraph borders, applied here at the
                    // whole-table level (outside + inside edges).
                    if (input.TryGetProperty("borders", out var bordersEl))
                    {
                        bool on = bordersEl.ValueKind == JsonValueKind.True;
                        Word.WdColor color = input.TryGetProperty("borderColor", out var bc) && bc.ValueKind == JsonValueKind.String
                            ? (Word.WdColor)ColorUtil.HexToOle(bc.GetString())
                            : (Word.WdColor)ColorUtil.HexToOle("#000000");
                        // Post-hoc fix (2026-08-24, user-reported): table.Borders
                        // is not just the 6 grid sides - it also includes
                        // wdBorderDiagonalDown/wdBorderDiagonalUp (the rare
                        // cell-split diagonal lines), so the blind foreach over
                        // the whole collection turned those on too, producing
                        // crisscrossing diagonals across every cell. Enumerate
                        // only the real table grid sides explicitly.
                        Word.WdBorderType[] sides =
                        {
                            Word.WdBorderType.wdBorderTop, Word.WdBorderType.wdBorderLeft,
                            Word.WdBorderType.wdBorderBottom, Word.WdBorderType.wdBorderRight,
                            Word.WdBorderType.wdBorderHorizontal, Word.WdBorderType.wdBorderVertical,
                        };
                        foreach (Word.WdBorderType side in sides)
                        {
                            Word.Border border = table.Borders[side];
                            border.LineStyle = on ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleNone;
                            if (on) border.Color = color;
                        }
                    }
                    return new ToolResult { Output = "Table style updated.", Mutated = true, Summary = "edit_table" };
                }
                case "set_shading":
                {
                    // Post-hoc addition (2026-08-24, user-requested): fills
                    // cell background color at cell/row/col/whole-table
                    // scope. Word.Cell.Shading.BackgroundPatternColor is the
                    // same property/pattern updateParagraphStyle's
                    // shadingFill already uses on paragraphs elsewhere in
                    // this file - applied per-cell here since Word tables
                    // have no single "shade this row" API, only per-cell
                    // shading (matches PowerPoint's own EditTableStyle,
                    // which does the identical per-cell loop for its
                    // shadingColor field).
                    string scope = input.GetProperty("scope").GetString();
                    Word.WdColor color = (Word.WdColor)ColorUtil.HexToOle(input.GetProperty("color").GetString());
                    int rowCount = table.Rows.Count, colCount = table.Columns.Count;
                    switch (scope)
                    {
                        case "cell":
                        {
                            int row = input.GetProperty("row").GetInt32();
                            int col = input.GetProperty("col").GetInt32();
                            if (row < 0 || row >= rowCount)
                                throw new ArgumentOutOfRangeException("row", "row must be between 0 and " + (rowCount - 1) + ".");
                            if (col < 0 || col >= colCount)
                                throw new ArgumentOutOfRangeException("col", "col must be between 0 and " + (colCount - 1) + ".");
                            table.Cell(row + 1, col + 1).Shading.BackgroundPatternColor = color;
                            break;
                        }
                        case "row":
                        {
                            int row = input.GetProperty("row").GetInt32();
                            if (row < 0 || row >= rowCount)
                                throw new ArgumentOutOfRangeException("row", "row must be between 0 and " + (rowCount - 1) + ".");
                            for (int c = 1; c <= colCount; c++)
                            {
                                // A merged cell's non-anchor positions throw on
                                // direct Cell(r,c) access - skip those rather
                                // than failing the whole row, same tolerance
                                // ReadTable already applies.
                                try { table.Cell(row + 1, c).Shading.BackgroundPatternColor = color; } catch { }
                            }
                            break;
                        }
                        case "col":
                        {
                            int col = input.GetProperty("col").GetInt32();
                            if (col < 0 || col >= colCount)
                                throw new ArgumentOutOfRangeException("col", "col must be between 0 and " + (colCount - 1) + ".");
                            for (int r = 1; r <= rowCount; r++)
                            {
                                try { table.Cell(r, col + 1).Shading.BackgroundPatternColor = color; } catch { }
                            }
                            break;
                        }
                        case "table":
                        {
                            for (int r = 1; r <= rowCount; r++)
                                for (int c = 1; c <= colCount; c++)
                                {
                                    try { table.Cell(r, c).Shading.BackgroundPatternColor = color; } catch { }
                                }
                            break;
                        }
                        default:
                            throw new ArgumentException("edit_table: unknown scope '" + scope + "' for set_shading. Valid: cell, row, col, table.");
                    }
                    return new ToolResult { Output = "Table shading applied (" + scope + ").", Mutated = true, Summary = "edit_table" };
                }
                default:
                    throw new ArgumentException("edit_table: unknown kind '" + kind + "'. Valid: set_cell, insert_row, delete_row, insert_col, delete_col, set_style, set_shading.");
            }
        }

        private static ToolResult ReadTable(JsonElement input)
        {
            Word.Tables tables = ActiveDoc.Tables;
            if (tables.Count == 0)
                return new ToolResult { Output = "No tables in this document.", Summary = "read_table" };

            int tableIndex = input.TryGetProperty("tableIndex", out var ti) && ti.ValueKind == JsonValueKind.Number ? ti.GetInt32() : 0;
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentOutOfRangeException("tableIndex", "tableIndex must be between 0 and " + (tables.Count - 1) + " (" + tables.Count + " table(s) in the document).");
            Word.Table table = tables[tableIndex + 1];

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Table " + tableIndex + " of " + tables.Count + " (" + table.Rows.Count + " rows x " + table.Columns.Count + " cols):");
            for (int r = 0; r < table.Rows.Count; r++)
            {
                var cellsOut = new List<string>();
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    // A merged cell can make Cell(r,c) throw for the cells it no
                    // longer owns - report a placeholder rather than failing the
                    // whole read over one merged region.
                    try { cellsOut.Add(table.Cell(r + 1, c + 1).Range.Text.TrimEnd('\r', '\a')); }
                    catch { cellsOut.Add("(merged)"); }
                }
                sb.AppendLine("[" + r + "] " + string.Join(" | ", cellsOut));
            }
            return new ToolResult { Output = sb.ToString().TrimEnd(), Summary = "read_table" };
        }

        // PP-23 Task 4: ported from PowerPointTools.SmartArtLayoutNames /
        // ResolveSmartArtLayout verbatim - same seven keys, same
        // two-distinct-errors design (unknown key vs. valid-key-but-not-in-
        // this-install's-gallery). SmartArt is the Office-shared object
        // model, not PowerPoint-specific - Application.SmartArtLayouts
        // resolves identically against this add-in's own ThisAddIn.
        private static readonly Dictionary<string, string> SmartArtLayoutNames = new Dictionary<string, string>
        {
            ["list"] = "Basic Block List",
            ["process"] = "Basic Process",
            ["cycle"] = "Basic Cycle",
            ["hierarchy"] = "Organization Chart",
            ["pyramid"] = "Basic Pyramid",
            ["matrix"] = "Basic Matrix",
            ["venn"] = "Basic Venn",
        };

        private static dynamic ResolveSmartArtLayout(string layoutKey)
        {
            string targetName;
            if (!SmartArtLayoutNames.TryGetValue(layoutKey, out targetName))
                throw new ArgumentException("add_smartart: unknown layout '" + layoutKey + "'. Valid: " +
                                            string.Join(", ", SmartArtLayoutNames.Keys) + ".");
            dynamic layouts = Globals.ThisAddIn.Application.SmartArtLayouts;
            foreach (dynamic layout in layouts)
            {
                if (string.Equals((string)layout.Name, targetName, StringComparison.OrdinalIgnoreCase))
                {
                    return layout;
                }
            }
            throw new InvalidOperationException("add_smartart: no SmartArt layout named '" + targetName +
                                                "' was found in this Office install's gallery - this install may be " +
                                                "non-English, where the built-in gallery's display names differ from " +
                                                "the standard English ones this tool assumes.");
        }

        // Post-hoc addition (2026-08-24, user-reported: "smart art has no
        // change style/color for an existing element"). Unlike layouts,
        // SmartArt color schemes and quick styles are NOT a fixed enum in
        // this Office object model (confirmed via reflection against the
        // referenced Office 15 PIA - Microsoft.Office.Core has no
        // MsoSmartArtColorType/MsoSmartArtQuickStyleType at all) - they are
        // live COM collections (Application.SmartArtColors /
        // .SmartArtQuickStyles) of SmartArtColor/SmartArtQuickStyle objects,
        // each with a .Name populated at runtime by this install's own
        // gallery. Rather than guess a curated list of exact display-name
        // strings (unverifiable without a live session, and a wrong guess
        // would either fail or - worse - silently match nothing), this
        // resolves by case-insensitive SUBSTRING match against whatever
        // names this install actually has, and a miss lists the real
        // available names so the caller can retry correctly instead of
        // guessing blind a second time.
        private static dynamic ResolveSmartArtGalleryItem(dynamic collection, string query, string toolName, string whatKind)
        {
            dynamic firstMatch = null;
            var namesSeen = new List<string>();
            foreach (dynamic item in collection)
            {
                string name = (string)item.Name;
                namesSeen.Add(name);
                if (firstMatch == null && name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) firstMatch = item;
            }
            if (firstMatch != null) return firstMatch;
            string available = namesSeen.Count > 20
                ? string.Join(", ", namesSeen.GetRange(0, 20)) + ", ... (" + namesSeen.Count + " total)"
                : string.Join(", ", namesSeen);
            throw new ArgumentException(toolName + ": no " + whatKind + " matching '" + query + "' found in this Office install's gallery. Available: " + available + ".");
        }

        private static ToolResult AddSmartArt(JsonElement input)
        {
            string layoutKey = input.GetProperty("layout").GetString();
            dynamic layout = ResolveSmartArtLayout(layoutKey);

            int? afterBlockIndex = input.TryGetProperty("afterBlockIndex", out var abEl) && abEl.ValueKind == JsonValueKind.Number
                ? abEl.GetInt32() : (int?)null;
            float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
            float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 300f;

            dynamic doc = ActiveDoc;
            dynamic shape;
            if (afterBlockIndex.HasValue)
            {
                // Mirrors PP-9's anchored-chart-creation path exactly, including
                // its caveat: whether Shapes.AddSmartArt truly accepts a named
                // Anchor parameter in this PIA is UNVERIFIED - flagged as
                // elevated risk in the plan/verification file.
                Word.Range at = RangeAfterBlock(afterBlockIndex.Value);
                dynamic floatingAtAnchor = doc.Shapes.AddSmartArt(layout, 0, 0, width, height, Anchor: at);
                shape = floatingAtAnchor.ConvertToInlineShape();
            }
            else
            {
                float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
                float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
                shape = doc.Shapes.AddSmartArt(layout, left, top, width, height);
            }

            dynamic smartArt = shape.SmartArt;

            // Post-hoc fix (2026-08-24, user-reported): AddSmartArt seeds the
            // new diagram with the layout's own default placeholder nodes
            // (the same "[Text]" prompts the ribbon's SmartArt gallery shows) -
            // same bug shape as the chart-data fix above (pre-seeded content
            // never cleared before writing). Without clearing them first, the
            // requested items were APPENDED after the placeholders instead of
            // replacing them, leaving visible "[Text]" nodes above the real
            // ones. Delete every existing node before adding the real ones,
            // same idea as the chart fix's sheet.Cells.Clear().
            dynamic existingNodes = smartArt.Nodes;
            for (int i = (int)existingNodes.Count; i >= 1; i--)
            {
                existingNodes.Item(i).Delete();
            }

            foreach (JsonElement item in input.GetProperty("items").EnumerateArray())
            {
                dynamic node = smartArt.Nodes.Add();
                node.TextFrame2.TextRange.Text = item.GetString();
            }
            return new ToolResult { Output = "SmartArt added (" + input.GetProperty("items").GetArrayLength() + " node(s)).", Mutated = true, Summary = "add_smartart" };
        }

        // PP-23 Task 5: SmartArt shapes are not chart shapes and are not
        // tables - a small, separate list-and-resolve helper, mirroring
        // ListChartShapes'/ResolveTable's shape but for shape.HasSmartArt
        // instead of shape.HasChart.
        //
        // Post-hoc fix (2026-08-24, user-reported): HasSmartArt returns an
        // MsoTriState, not a real bool, exactly like HasChart elsewhere in
        // this file - a plain (bool) cast either throws (silently swallowed
        // by the try/catch below) or never matches, so no shape was ever
        // recognized as SmartArt and read_smartart/edit_smartart always
        // reported "no SmartArt diagrams" even right after add_smartart had
        // just created one. Fixed with the same (int)x == -1 comparison
        // ListChartShapes already uses for HasChart.
        internal static List<dynamic> ListSmartArtShapes(dynamic doc)
        {
            var shapes = new List<dynamic>();
            foreach (dynamic shp in doc.InlineShapes)
            {
                try { if ((int)shp.HasSmartArt == -1 /* msoTrue */) shapes.Add(shp); } catch { }
            }
            foreach (dynamic shp in doc.Shapes)
            {
                try { if ((int)shp.HasSmartArt == -1 /* msoTrue */) shapes.Add(shp); } catch { }
            }
            return shapes;
        }

        // Post-hoc fix (2026-08-24, user-reported): reading N diagrams
        // previously needed N separate read_smartart calls (one per index) -
        // extracted so ReadSmartArt can read every diagram in one call when
        // smartArtIndex is omitted, matching what the user actually wanted
        // ("a way to read the entire smartart text at once").
        private static string ReadOneSmartArt(dynamic shape, int index, int total)
        {
            dynamic smartArt = shape.SmartArt;
            dynamic nodes = smartArt.Nodes;
            int count = (int)nodes.Count;
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("SmartArt " + index + " of " + total + " (" + count + " node(s)):");
            for (int i = 1; i <= count; i++)
            {
                dynamic node = nodes.Item(i);
                string text = "";
                try { text = (string)node.TextFrame2.TextRange.Text; } catch { }
                sb.AppendLine("[" + (i - 1) + "] " + text);
            }
            return sb.ToString().TrimEnd();
        }

        private static ToolResult ReadSmartArt(JsonElement input)
        {
            dynamic doc = ActiveDoc;
            var shapes = ListSmartArtShapes(doc);
            if (shapes.Count == 0)
                return new ToolResult { Output = "No SmartArt diagrams in this document.", Summary = "read_smartart" };

            bool hasIndex = input.TryGetProperty("smartArtIndex", out var si) && si.ValueKind == JsonValueKind.Number;
            if (!hasIndex)
            {
                // No index given: read every diagram in one call rather than
                // forcing one call per diagram.
                var all = new List<string>();
                for (int i = 0; i < shapes.Count; i++) all.Add(ReadOneSmartArt(shapes[i], i, shapes.Count));
                return new ToolResult { Output = string.Join("\n\n", all), Summary = "read_smartart" };
            }

            int index = si.GetInt32();
            if (index < 0 || index >= shapes.Count)
                throw new ArgumentOutOfRangeException("smartArtIndex", "smartArtIndex must be between 0 and " + (shapes.Count - 1) + " (" + shapes.Count + " diagram(s) in the document).");
            return new ToolResult { Output = ReadOneSmartArt(shapes[index], index, shapes.Count), Summary = "read_smartart" };
        }

        private static ToolResult EditSmartArt(JsonElement input)
        {
            dynamic doc = ActiveDoc;
            var shapes = ListSmartArtShapes(doc);
            if (shapes.Count == 0)
                throw new InvalidOperationException("edit_smartart: no SmartArt diagrams in this document.");

            int index = input.TryGetProperty("smartArtIndex", out var si) && si.ValueKind == JsonValueKind.Number ? si.GetInt32() : 0;
            if (index < 0 || index >= shapes.Count)
                throw new ArgumentOutOfRangeException("smartArtIndex", "smartArtIndex must be between 0 and " + (shapes.Count - 1) + " (" + shapes.Count + " diagram(s) in the document).");

            dynamic smartArt = shapes[index].SmartArt;
            dynamic nodes = smartArt.Nodes;
            string kind = input.GetProperty("kind").GetString();
            switch (kind)
            {
                case "set_text":
                {
                    int nodeIndex = input.GetProperty("nodeIndex").GetInt32();
                    int count = (int)nodes.Count;
                    if (nodeIndex < 0 || nodeIndex >= count)
                        throw new ArgumentOutOfRangeException("nodeIndex", "nodeIndex must be between 0 and " + (count - 1) + " (" + count + " node(s)).");
                    nodes.Item(nodeIndex + 1).TextFrame2.TextRange.Text = input.GetProperty("text").GetString();
                    return new ToolResult { Output = "Node " + nodeIndex + " updated.", Mutated = true, Summary = "edit_smartart" };
                }
                case "add_node":
                {
                    dynamic newNode = nodes.Add();
                    if (input.TryGetProperty("text", out var textEl) && textEl.ValueKind == JsonValueKind.String)
                        newNode.TextFrame2.TextRange.Text = textEl.GetString();
                    return new ToolResult { Output = "Node added at index " + ((int)nodes.Count - 1) + ".", Mutated = true, Summary = "edit_smartart" };
                }
                case "delete_node":
                {
                    int nodeIndex = input.GetProperty("nodeIndex").GetInt32();
                    int count = (int)nodes.Count;
                    if (nodeIndex < 0 || nodeIndex >= count)
                        throw new ArgumentOutOfRangeException("nodeIndex", "nodeIndex must be between 0 and " + (count - 1) + " (" + count + " node(s)).");
                    nodes.Item(nodeIndex + 1).Delete();
                    return new ToolResult { Output = "Node " + nodeIndex + " deleted. Later node indices have shifted - re-read (read_smartart) before another node edit in the same run.", Mutated = true, Summary = "edit_smartart" };
                }
                case "set_style":
                {
                    bool changed = false;
                    if (input.TryGetProperty("colorName", out var cnEl) && cnEl.ValueKind == JsonValueKind.String)
                    {
                        dynamic colors = Globals.ThisAddIn.Application.SmartArtColors;
                        smartArt.Color = ResolveSmartArtGalleryItem(colors, cnEl.GetString(), "edit_smartart", "color scheme");
                        changed = true;
                    }
                    if (input.TryGetProperty("quickStyleName", out var qsEl) && qsEl.ValueKind == JsonValueKind.String)
                    {
                        dynamic quickStyles = Globals.ThisAddIn.Application.SmartArtQuickStyles;
                        smartArt.QuickStyle = ResolveSmartArtGalleryItem(quickStyles, qsEl.GetString(), "edit_smartart", "quick style");
                        changed = true;
                    }
                    if (!changed)
                        throw new ArgumentException("edit_smartart: set_style requires at least one of colorName or quickStyleName.");
                    return new ToolResult { Output = "SmartArt style updated.", Mutated = true, Summary = "edit_smartart" };
                }
                case "set_layout":
                {
                    // Post-hoc addition (2026-08-24, user-reported: "smart art
                    // cant change layout"). Reuses ResolveSmartArtLayout
                    // verbatim (same curated 7-key map + gallery lookup
                    // add_smartart already uses) - SmartArt.Layout is
                    // settable (confirmed via the same reflection pass that
                    // found .Color/.QuickStyle), so changing an EXISTING
                    // diagram's layout is the same resolve-then-assign shape
                    // as creating one, just against smartArt.Layout instead
                    // of the AddSmartArt call.
                    string layoutKey = input.GetProperty("layout").GetString();
                    smartArt.Layout = ResolveSmartArtLayout(layoutKey);
                    return new ToolResult { Output = "SmartArt layout changed to '" + layoutKey + "'.", Mutated = true, Summary = "edit_smartart" };
                }
                default:
                    throw new ArgumentException("edit_smartart: unknown kind '" + kind + "'. Valid: set_text, add_node, delete_node, set_style, set_layout.");
            }
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

        // PP-5: mirrors WORD_COMMAND_SCHEMAS's `required` arrays in
        // WordAiAddIn/web-src/entry.ts exactly (minus "kind" itself, which is
        // validated separately in ApplyCommands before this table is
        // consulted) - the two must be edited together. This is the actual
        // guarantee: the TS schema is documentation the model reads, not a
        // validator that runs (not every provider enforces oneOf/const), so
        // this precheck is what turns a missing field into a specific,
        // per-command error instead of a raw COM/NullReference exception.
        private static readonly Dictionary<string, string[]> RequiredFields = new Dictionary<string, string[]>
        {
            ["set_bold"] = new[] { "startIndex", "endIndex", "value" },
            ["set_italic"] = new[] { "startIndex", "endIndex", "value" },
            ["set_heading"] = new[] { "index", "level" },
            ["find_replace"] = new[] { "find", "replace" },
            ["updateTextStyle"] = new[] { "target", "style", "fields" },
            ["updateParagraphStyle"] = new[] { "target", "style", "fields" },
            ["deleteBlocks"] = new[] { "target" },
            ["moveBlocks"] = new[] { "blockIndexes", "afterBlockIndex" },
            ["createParagraphBullets"] = new[] { "target" },
            ["deleteParagraphBullets"] = new[] { "target" },
            ["updateImageProperties"] = new[] { "imageIndex", "properties", "fields" },
            ["insertToc"] = new[] { "afterBlockIndex" },
        };

        // Fields where an explicit JSON null is a caller error rather than a
        // value - e.g. set_bold's "value": null previously reached
        // GetBoolean() and threw an opaque InvalidOperationException where
        // this clean "missing required field" message belongs. Deliberately
        // narrow: only add a field here after confirming null has no
        // legitimate meaning for it (Excel's set_cell "value" is the
        // counter-example - null there means "clear the cell").
        private static readonly Dictionary<string, string[]> NonNullFields = new Dictionary<string, string[]>
        {
            ["set_bold"] = new[] { "value" },
            ["set_italic"] = new[] { "value" },
            ["set_heading"] = new[] { "level" },
        };

        // PP-12 Task 3 (the half PP-5 Task 4 Step 1 did not cover): each
        // result line is prefixed with the command's 0-based position in the
        // batch, and a summary header states how many succeeded/failed - with
        // partial batches now the norm (no rollback - Word COM offers no
        // batch transaction, and a hand-rolled undo would be less reliable
        // than this honest report; the user retains Word's own Ctrl+Z), the
        // model needs to know WHICH command in a batch of several identical
        // kinds failed, not just that "one of them" did.
        private static ToolResult ApplyCommands(JsonElement input)
        {
            var lines = new System.Text.StringBuilder();
            bool anyMutated = false;
            int failedCount = 0;
            int commandIndex = 0;
            int totalCount = 0;
            foreach (JsonElement cmd in input.GetProperty("commands").EnumerateArray())
            {
                totalCount++;
                string kind = null;
                try
                {
                    JsonElement kindEl;
                    if (!cmd.TryGetProperty("kind", out kindEl) || kindEl.ValueKind != JsonValueKind.String)
                        throw new ArgumentException("Command is missing a string \"kind\" field.");
                    kind = kindEl.GetString();
                    ToolArgs.ValidateRequired(kind, cmd, RequiredFields, "Command", NonNullFields);
                    switch (kind)
                    {
                        case "set_bold":
                            SetRunProperty(cmd, (range, value) => range.Bold = value ? 1 : 0);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "set_italic":
                            SetRunProperty(cmd, (range, value) => range.Italic = value ? 1 : 0);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "set_heading":
                            SetHeading(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "find_replace":
                            int replacements = FindReplace(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: {replacements} replacement(s)");
                            if (replacements > 0) anyMutated = true;
                            break;
                        case "updateTextStyle":
                            UpdateTextStyle(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "updateParagraphStyle":
                            UpdateParagraphStyle(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "deleteBlocks":
                            DeleteBlocksCmd(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "moveBlocks":
                            MoveBlocksCmd(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "createParagraphBullets":
                        {
                            string report = CreateParagraphBullets(cmd);
                            lines.AppendLine($"[{commandIndex}] {report}");
                            if (report.StartsWith("createParagraphBullets: 0 applied")) { /* nothing changed */ } else anyMutated = true;
                            break;
                        }
                        case "deleteParagraphBullets":
                        {
                            string report = DeleteParagraphBullets(cmd);
                            lines.AppendLine($"[{commandIndex}] {report}");
                            if (report.StartsWith("deleteParagraphBullets: 0 removed")) { /* nothing changed */ } else anyMutated = true;
                            break;
                        }
                        case "updateImageProperties":
                            UpdateImageProperties(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "insertToc":
                            InsertTocCmd(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        default:
                            lines.AppendLine($"[{commandIndex}] {kind}: unknown command kind"); failedCount++; break;
                    }
                }
                catch (Exception ex)
                {
                    lines.AppendLine($"[{commandIndex}] " + (kind ?? "(unknown kind)") + ": ERROR - " + ex.Message);
                    failedCount++;
                }
                commandIndex++;
            }
            string summary = $"Applied {totalCount - failedCount} of {totalCount} command(s)" + (failedCount > 0 ? $" ({failedCount} failed)." : ".");
            return new ToolResult { Output = summary + "\n" + lines, Mutated = anyMutated, IsError = failedCount > 0, Summary = "apply_commands" };
        }

        private static void SetRunProperty(JsonElement cmd, Action<Word.Range, bool> apply)
        {
            int startIndex = cmd.GetProperty("startIndex").GetInt32();
            int endIndex = cmd.GetProperty("endIndex").GetInt32();
            bool value = cmd.GetProperty("value").GetBoolean();
            Word.Document doc = ActiveDoc;
            Word.Paragraphs paragraphs = doc.Paragraphs;
            endIndex = Math.Min(endIndex, paragraphs.Count - 1);
            Word.Range range = doc.Range(paragraphs[startIndex + 1].Range.Start, paragraphs[endIndex + 1].Range.End);
            apply(range, value);
        }

        private static void SetHeading(JsonElement cmd)
        {
            int index = cmd.GetProperty("index").GetInt32();
            int level = cmd.GetProperty("level").GetInt32();
            Word.Paragraph p = ActiveDoc.Paragraphs[index + 1];
            p.Range.set_Style(level == 0 ? "Normal" : "Heading " + level);
        }

        private static int FindReplace(JsonElement cmd)
        {
            string find = cmd.GetProperty("find").GetString();
            string replace = cmd.GetProperty("replace").GetString();
            bool matchCase = cmd.TryGetProperty("matchCase", out var mc) && mc.GetBoolean();
            Word.Find findObj = ActiveDoc.Content.Find;
            findObj.ClearFormatting();
            findObj.Text = find;
            findObj.Replacement.ClearFormatting();
            findObj.Replacement.Text = replace;
            findObj.MatchCase = matchCase;
            // wdReplaceAll only reports whether ANYTHING matched (true/false),
            // not how many - the caller's "N replacement(s)" message used to
            // always say 0 or 1 regardless of the real count. Loop
            // wdReplaceOne instead (the same one-at-a-time advance Word's own
            // "Replace All" button does internally) so the count is accurate.
            int count = 0;
            while (findObj.Execute(Replace: Word.WdReplace.wdReplaceOne))
            {
                count++;
                if (count > 10000) break; // safety net, not expected to ever trigger
            }
            return count;
        }

        // Returns each matched paragraph's 0-based index AND its already-
        // resolved Paragraph object (not just the index) - every caller used
        // to turn around and re-look up paragraphs[i + 1] per match, paying
        // Word's slow positional-indexing cost a second time. Returning the
        // object we already have in hand during the walk removes that
        // second lookup entirely.
        private static List<(int Index, Word.Paragraph Paragraph)> ResolveTargetParagraphs(JsonElement target)
        {
            string nodeType = target.TryGetProperty("nodeType", out var nt) && nt.ValueKind == JsonValueKind.String ? nt.GetString() : null;
            int? headingLevel = target.TryGetProperty("headingLevel", out var hl) && hl.ValueKind == JsonValueKind.Number ? hl.GetInt32() : (int?)null;
            string containsText = target.TryGetProperty("containsText", out var ct) && ct.ValueKind == JsonValueKind.String ? ct.GetString() : null;
            bool matchCase = target.TryGetProperty("matchCase", out var mc) && mc.ValueKind == JsonValueKind.True;
            HashSet<int> blockIndexes = null;
            if (target.TryGetProperty("blockIndexes", out var bi) && bi.ValueKind == JsonValueKind.Array)
            {
                blockIndexes = new HashSet<int>();
                foreach (JsonElement e in bi.EnumerateArray()) blockIndexes.Add(e.GetInt32());
            }
            string scope = target.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String ? sc.GetString() : "document";

            if (nodeType == null && containsText == null && blockIndexes == null)
            {
                throw new ArgumentException("Target must specify at least one of nodeType, containsText, or blockIndexes.");
            }

            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            int selStart = -1, selEnd = -1;
            if (scope == "selection")
            {
                Word.Selection sel = Globals.ThisAddIn.Application.Selection;
                if (sel.Type != Word.WdSelectionType.wdNoSelection)
                {
                    selStart = sel.Range.Start;
                    selEnd = sel.Range.End;
                }
            }

            // Walks forward via the collection's own enumerator instead of
            // positional paragraphs[i + 1] indexing - Paragraphs is not a
            // real array in Word's COM object model, so indexing it by
            // position re-walks the document from the start on EVERY single
            // access, turning a full scan into roughly O(n^2) internally
            // (confirmed root cause of a real reported freeze in find_text/
            // get_headings, fixed there the same way). Every command that
            // funnels through this one function - updateTextStyle,
            // updateParagraphStyle, deleteBlocks, createParagraphBullets,
            // deleteParagraphBullets - inherits the fix.
            var result = new List<(int, Word.Paragraph)>();
            int i = 0;
            foreach (Word.Paragraph p in paragraphs)
            {
                // The body has several early `continue`s (skip this
                // paragraph, keep walking) - wrapping it in try/finally
                // guarantees i still advances exactly once per paragraph on
                // every path, since `continue` inside a try still runs its
                // finally before moving to the next iteration.
                try
                {
                    if (scope == "selection")
                    {
                        if (selStart == -1) continue;
                        if (p.Range.Start > selEnd || p.Range.End < selStart) continue;
                    }

                    if (blockIndexes != null && !blockIndexes.Contains(i)) continue;

                    string styleName = p.Range.get_Style().NameLocal;
                    bool isHeading = styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
                    bool isListItem = p.Range.ListFormat.ListType != Word.WdListType.wdListNoNumbering;

                    if (nodeType == "heading" && !isHeading) continue;
                    if (nodeType == "paragraph" && (isHeading || isListItem)) continue;
                    if (nodeType == "listItem" && !isListItem) continue;

                    if (nodeType == "heading" && headingLevel.HasValue)
                    {
                        string levelDigits = new string(styleName.Where(char.IsDigit).ToArray());
                        if (!int.TryParse(levelDigits, out int actualLevel) || actualLevel != headingLevel.Value) continue;
                    }

                    if (containsText != null)
                    {
                        string text = p.Range.Text ?? "";
                        bool found = matchCase
                            ? text.Contains(containsText)
                            : text.IndexOf(containsText, StringComparison.OrdinalIgnoreCase) >= 0;
                        if (!found) continue;
                    }

                    result.Add((i, p));
                }
                finally
                {
                    i++;
                }
            }
            return result;
        }

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

        private static void DeleteBlocksCmd(JsonElement cmd)
        {
            var matches = ResolveTargetParagraphs(cmd.GetProperty("target"));
            if (matches.Count == 0)
            {
                throw new InvalidOperationException("deleteBlocks: no paragraphs matched target.");
            }
            if (matches.Count >= ActiveDoc.Paragraphs.Count)
            {
                // Deleting every paragraph would leave zero - clear content instead,
                // leaving one empty paragraph (mirrors genoffice's own guard).
                ActiveDoc.Content.Text = "";
                return;
            }
            // Delete in descending order so an earlier Paragraph object's Range
            // is never asked to reason about content after it disappearing.
            matches.Sort((a, b) => b.Index.CompareTo(a.Index));
            foreach (var (_, p) in matches)
            {
                p.Range.Delete();
            }
        }

        private static void MoveBlocksCmd(JsonElement cmd)
        {
            var blockIndexes = new List<int>();
            foreach (JsonElement e in cmd.GetProperty("blockIndexes").EnumerateArray()) blockIndexes.Add(e.GetInt32());
            if (blockIndexes.Count == 0)
            {
                throw new ArgumentException("moveBlocks: blockIndexes must contain at least one index.");
            }
            int afterBlockIndex = cmd.GetProperty("afterBlockIndex").GetInt32();

            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            int count = paragraphs.Count;
            if (blockIndexes.Any(i => i < 0 || i >= count) || afterBlockIndex < -1 || afterBlockIndex >= count)
            {
                throw new ArgumentException("moveBlocks: index out of range.");
            }
            if (blockIndexes.Contains(afterBlockIndex))
            {
                throw new ArgumentException("moveBlocks: afterBlockIndex cannot be one of the moved blocks.");
            }

            blockIndexes.Sort();
            // Capture each moved paragraph's content as an OOXML string - a true,
            // detached snapshot (plain text, not a live COM Range reference) -
            // before any deletion shifts indices, so formatting survives the
            // move without depending on FormattedText's live-vs-copy semantics.
            var captured = blockIndexes.Select(i => paragraphs[i + 1].Range.WordOpenXML).ToList();

            // Delete moved paragraphs in descending order.
            var deleteOrder = new List<int>(blockIndexes);
            deleteOrder.Reverse();
            foreach (int i in deleteOrder)
            {
                paragraphs[i + 1].Range.Delete();
            }

            // Recompute the insertion point: afterBlockIndex shifts down by however
            // many moved blocks were originally BEFORE it.
            int shift = blockIndexes.Count(i => i < afterBlockIndex);
            int adjustedAfter = afterBlockIndex - shift;

            Word.Range insertionPoint = adjustedAfter == -1
                ? ActiveDoc.Range(0, 0)
                : ActiveDoc.Paragraphs[adjustedAfter + 1].Range;
            insertionPoint.Collapse(adjustedAfter == -1 ? Word.WdCollapseDirection.wdCollapseStart : Word.WdCollapseDirection.wdCollapseEnd);

            foreach (string xml in captured)
            {
                insertionPoint.InsertXML(xml);
                insertionPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
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

        private static void UpdateImageProperties(JsonElement cmd)
        {
            int imageIndex = cmd.GetProperty("imageIndex").GetInt32(); // 0-based index into doc.InlineShapes
            Word.InlineShapes shapes = ActiveDoc.InlineShapes;
            if (imageIndex < 0 || imageIndex >= shapes.Count)
            {
                throw new ArgumentException("updateImageProperties: imageIndex out of range.");
            }
            Word.InlineShape shape = shapes[imageIndex + 1];
            JsonElement properties = cmd.GetProperty("properties");
            HashSet<string> fields = new HashSet<string>();
            foreach (JsonElement f in cmd.GetProperty("fields").EnumerateArray()) fields.Add(f.GetString());

            const float pxToPoints = 0.75f; // 96dpi px -> points, matches genoffice's own pixel model
            float? newWidth = null, newHeight = null;
            if (fields.Contains("widthPx") && properties.TryGetProperty("widthPx", out var w) && w.ValueKind == JsonValueKind.Number)
                newWidth = (float)w.GetDouble() * pxToPoints;
            if (fields.Contains("heightPx") && properties.TryGetProperty("heightPx", out var h) && h.ValueKind == JsonValueKind.Number)
                newHeight = (float)h.GetDouble() * pxToPoints;

            if (newWidth.HasValue && !newHeight.HasValue)
            {
                newHeight = shape.Height * (newWidth.Value / shape.Width); // proportional scale from current size
            }
            else if (newHeight.HasValue && !newWidth.HasValue)
            {
                newWidth = shape.Width * (newHeight.Value / shape.Height);
            }
            if (newWidth.HasValue) shape.Width = newWidth.Value;
            if (newHeight.HasValue) shape.Height = newHeight.Value;

            if (fields.Contains("align") && properties.TryGetProperty("align", out var align) && align.ValueKind == JsonValueKind.String)
            {
                switch (align.GetString())
                {
                    case "left": shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;
                    case "center": shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;
                    case "right": shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; break;
                }
            }
        }

        // PP-11: same air-gapped local-file-only rule as Excel's
        // AddImageExcel/PowerPoint's replace_image, worded consistently. The
        // File.Exists check is the one addition over Excel's version -
        // AddPicture on a missing file throws a bare COMException with a
        // useless message; this lets the model correct the path next turn.
        private static string ValidateLocalImagePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("add_image: path is required.");
            if (path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException(
                    "add_image: remote URLs are not supported in this air-gapped deployment - use a local file path.");
            if (!System.IO.File.Exists(path))
                throw new System.IO.FileNotFoundException("add_image: no file at '" + path + "'.");
            return path;
        }

        private static ToolResult AddImage(JsonElement input)
        {
            string path = ValidateLocalImagePath(input.GetProperty("path").GetString());
            bool floating = input.TryGetProperty("floating", out var flEl) && flEl.ValueKind == JsonValueKind.True;
            int afterBlockIndex = input.TryGetProperty("afterBlockIndex", out var abEl) && abEl.ValueKind == JsonValueKind.Number
                ? abEl.GetInt32() : int.MinValue; // sentinel: append at end (-1 already means "start of document" in this file's convention)

            Word.Document doc = ActiveDoc;
            Word.Range at;
            if (afterBlockIndex == int.MinValue)
            {
                at = doc.Content;
                at.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
            else
            {
                at = RangeAfterBlock(afterBlockIndex);
            }

            float? widthPoints = input.TryGetProperty("widthPoints", out var wEl) && wEl.ValueKind == JsonValueKind.Number ? (float?)wEl.GetDouble() : null;
            float? heightPoints = input.TryGetProperty("heightPoints", out var hEl) && hEl.ValueKind == JsonValueKind.Number ? (float?)hEl.GetDouble() : null;
            string altText = input.TryGetProperty("altText", out var altEl) && altEl.ValueKind == JsonValueKind.String ? altEl.GetString() : null;

            float finalWidth, finalHeight;
            string addressability;

            if (floating)
            {
                double left = at.get_Information(Word.WdInformation.wdHorizontalPositionRelativeToPage);
                double top = at.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage);
                Word.Shape shape = doc.Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                    (float)left, (float)top, -1, -1);
                float naturalW = shape.Width, naturalH = shape.Height;
                GeometryUtil.ResolveImageSize(naturalW, naturalH, widthPoints, heightPoints, out finalWidth, out finalHeight);
                shape.Width = finalWidth;
                shape.Height = finalHeight;
                if (altText != null) shape.AlternativeText = altText;
                addressability = "not addressable by apply_commands/updateImageProperties (floating)";
            }
            else
            {
                Word.InlineShape shape = doc.InlineShapes.AddPicture(path, LinkToFile: false, SaveWithDocument: true, Range: at);
                float naturalW = shape.Width, naturalH = shape.Height;
                GeometryUtil.ResolveImageSize(naturalW, naturalH, widthPoints, heightPoints, out finalWidth, out finalHeight);
                shape.Width = finalWidth;
                shape.Height = finalHeight;
                if (altText != null) shape.AlternativeText = altText;
                // InlineShapes is ordered by document position, not insertion
                // time - the new shape is only the LAST entry if it was
                // appended at the document's end. Find its real index by
                // position instead of assuming Count-1.
                int newIndex = -1;
                int shapeStart = shape.Range.Start;
                for (int idx = 0; idx < doc.InlineShapes.Count; idx++)
                {
                    if (doc.InlineShapes[idx + 1].Range.Start == shapeStart) { newIndex = idx; break; }
                }
                addressability = newIndex >= 0
                    ? $"addressable via apply_commands/updateImageProperties at imageIndex {newIndex}"
                    : "inserted, but its index could not be resolved";
            }

            return new ToolResult
            {
                Output = $"Image inserted from '{path}' ({finalWidth:0}x{finalHeight:0}pt) - {addressability}.",
                Mutated = true,
                Summary = "add_image",
            };
        }

        private static void InsertTocCmd(JsonElement cmd)
        {
            int afterBlockIndex = cmd.GetProperty("afterBlockIndex").GetInt32();
            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            bool hasHeadings = false;
            foreach (Word.Paragraph p in paragraphs)
            {
                if (p.Range.get_Style().NameLocal.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) { hasHeadings = true; break; }
            }
            if (!hasHeadings)
            {
                throw new InvalidOperationException("insertToc: document has no heading-styled paragraphs to build a table of contents from.");
            }
            if (afterBlockIndex < -1 || afterBlockIndex >= paragraphs.Count)
            {
                throw new ArgumentException("insertToc: afterBlockIndex out of range.");
            }

            Word.Range insertionPoint = afterBlockIndex == -1
                ? ActiveDoc.Range(0, 0)
                : paragraphs[afterBlockIndex + 1].Range;
            insertionPoint.Collapse(afterBlockIndex == -1 ? Word.WdCollapseDirection.wdCollapseStart : Word.WdCollapseDirection.wdCollapseEnd);
            insertionPoint.InsertParagraphAfter();
            insertionPoint.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            // Word's own native TOC field - auto-scans heading-styled paragraphs and
            // produces real, page-numbered entries directly. This is a more direct,
            // simpler native equivalent than genoffice's own hand-built TOC field-XML
            // workaround (real Word already paginates; genoffice's web renderer doesn't).
            ActiveDoc.TablesOfContents.Add(insertionPoint, UseHeadingStyles: true);
        }
    }
}
