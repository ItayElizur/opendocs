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
        // Server-side editing-mode gate, keyed by the same per-document id
        // TaskPaneHost.GetChatId() produces, so a mode change in one window never
        // affects another. Absent key defaults to FullAutonomy.
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
            // find_text/get_headings are read-only too - were previously missing here,
            // which both blocked them in Read Only mode and wrapped them in a
            // pointless undo custom record.
            "find_text", "get_headings",
        };

        public static ToolResult Execute(string docKey, string name, JsonElement input)
        {
            // Logs every tool call and failure regardless of which tool, so a repro
            // always shows up even if a specific method's own logging missed the
            // failure point.
            DebugLog.Write("Execute: " + name + " input=" + input.GetRawText());
            try
            {
                EditingMode mode = ModeFor(docKey);
                bool isAlwaysAllowed = AlwaysAllowedTools.Contains(name);
                bool isAddComment = name == "add_comment";
                // "Mutating" means "changes document content/structure", used only to
                // decide whether TrackRevisions toggles. add_comment doesn't mutate
                // content, so it's excluded here even though it's gated like a
                // mutating tool below.
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

                // Word's native undo stack registers one entry per COM write, not per
                // tool call; StartCustomRecord/EndCustomRecord groups all writes in a
                // tool call into one undo step. Skipped for always-allowed/undo/redo
                // tools. Must run in try/finally - an uncaught throw mid-mutation would
                // leave Word recording forever. See WordTools.cs.md.
                bool shouldRecordUndo = !isAlwaysAllowed && name != "undo_last_action" && name != "redo_last_action";
                if (shouldRecordUndo)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord(name);
                }
                try
                {
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
                        case "undo_last_action":
                            return UndoLastAction();
                        case "redo_last_action":
                            return RedoLastAction();
                        default:
                            return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                    }
                }
                finally
                {
                    // Logged, not rethrown - an uncaught throw from EndCustomRecord()
                    // here would discard an already-successful ToolResult (C#'s
                    // finally-after-return semantics) and report a real mutation as a
                    // generic failure. See WordTools.cs.md.
                    if (shouldRecordUndo)
                    {
                        try
                        {
                            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                        }
                        catch (Exception endEx)
                        {
                            DebugLog.WriteException("Execute: EndCustomRecord for " + name, endEx);
                        }
                    }
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

        // Resolves whichever document is ACTIVE right now, not necessarily the one
        // whose pane initiated this call - a long-running run whose user switches
        // windows mid-run would write into the wrong document. Known limitation,
        // not fixed here. See WordTools.cs.md.
        private static Word.Document ActiveDoc => Globals.ThisAddIn.Application.ActiveDocument;

        // Shared insertion-point resolver. afterBlockIndex is 0-based over
        // ActiveDoc.Paragraphs, matching every other block-addressed tool; -1 means
        // "before the first paragraph". Consumed by insert_content, chart anchoring,
        // and add_image - one helper, not three copies.
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

        // Shared end-of-document insertion point, matching InsertContent's/AddImage's
        // inline idiom (extracted rather than duplicated a third time).
        private static Word.Range EndOfDocumentRange()
        {
            Word.Range end = ActiveDoc.Content;
            end.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            return end;
        }

        // Chart-type vocabulary now lives in OfficeAi.Shared.ChartTypes - one table
        // shared by all three add-ins.

        // See WordTools.cs.md for a historical note on the chart workbook retry
        // (ComRetry.Run usage in WordTools.Charts.cs) left here.

    }
}

