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

    }
}

