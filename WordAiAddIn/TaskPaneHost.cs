using System.Diagnostics;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public partial class TaskPaneHost : PaneHostBase
    {
        private readonly Word.Document _document;
        private readonly int _hwnd;
        private string _chatId;

        // Deliberately does NOT dereference _document here (no .Path/.FullName
        // read). Accessing Application.ActiveDocument-equivalent state eagerly,
        // at the exact moment this constructor runs inside CustomTaskPanes.Add
        // (called from Application_WindowActivate), hits a COM timing issue in
        // Word's own startup/activation sequence and silently kills the whole
        // add-in connection (VSTO never connects it - no exception, no
        // resiliency-disabled entry, just Connect=False forever). Confirmed by
        // direct repro. GetChatId() computes lazily on first actual use, by
        // which point the pane is visible and the user has triggered a
        // message - so _document's state is guaranteed settled.
        public TaskPaneHost(Word.Document document, int hwnd) : base("WordAiAddIn")
        {
            _document = document;
            _hwnd = hwnd;
        }

        protected override ToolResult ExecuteTool(string name, JsonElement input)
        {
            return WordTools.Execute(GetChatId(), name, input);
        }

        protected override string GetChatId()
        {
            // A saved id is final - never re-checked again. An "unsaved-" id
            // is provisional: re-check the document's Path on every call, so
            // the first use after the user saves migrates chat history and
            // doc settings onto the real per-file id (FT-1 Task 7b). The Path
            // read is one cheap COM property on operations (load-history,
            // append-message, etc.) that are already doing file I/O.
            if (_chatId != null && !_chatId.StartsWith("unsaved-")) return _chatId;

            if (string.IsNullOrEmpty(_document.Path))
            {
                // An unsaved document has no on-disk Path; Document.FullName
                // falls back to its temp Name (e.g. "Document1") in that case,
                // which is not a stable key across sessions - and with
                // multiple panes now possible in one process, "unsaved-<pid>"
                // alone would collide across two different unsaved documents,
                // so the window handle is folded in too.
                return _chatId ?? (_chatId = "unsaved-" + Process.GetCurrentProcess().Id + "-" + _hwnd);
            }

            string saved = ChatStore.ChatIdForFile(_document.FullName);
            if (_chatId != null)
            {
                ChatStore.Migrate("WordAiAddIn", _chatId, saved);
                DocSettingsStore.Migrate("WordAiAddIn", _chatId, saved);
            }
            // Once _chatId is set to a saved (non-"unsaved-") id, the guard at
            // the top of this method returns it immediately on every later
            // call - so a subsequent Save As (which changes _document.FullName
            // again) does NOT re-key. The conversation and guidelines stay
            // with the id first saved to, not the copy. Whether they should
            // follow the copy instead has no obviously right answer; changing
            // this silently would be worse than this documented quirk.
            return _chatId = saved;
        }

        protected override void SetEditingMode(EditingMode mode)
        {
            WordTools.SetMode(GetChatId(), mode);
        }

        public void OnSelectionChanged(Word.Selection selection)
        {
            // ROOT CAUSE FOUND (2026-08-24, via DebugLog from a real repro):
            // a shape selection (clicking a chart/SmartArt) DOES set
            // Start != End (hasSelection=true) as normal, but
            // selection.Text returns NULL rather than "" for that selection
            // type - fullText.Length below threw NullReferenceException on
            // every single chart/SmartArt click, aborting OnSelectionChanged
            // before any of the object-detection code even ran. This was a
            // pre-existing gap (predates this session's object-detection
            // work) that only became visible once the previously-silent
            // outer catch in ThisAddIn.cs's Application_WindowSelectionChange
            // started logging instead of swallowing.
            bool hasSelection = selection.Start != selection.End;
            string fullText = hasSelection ? (selection.Text ?? "") : "";
            if (fullText.Length > 24000) fullText = fullText.Substring(0, 24000);
            string preview = fullText.Length > 40 ? fullText.Substring(0, 40) : fullText;

            // Post-hoc fix (2026-08-24, user-reported): the selection payload
            // previously carried only text, with no addressability at all -
            // the one selection kind FT-2 left un-addressable (Excel's range
            // gets an A1 address, PowerPoint's gets slideIndex/shapeIndex).
            // Without a paragraph index, the model has no way to call
            // replace_blocks on exactly what the user selected, and fell back
            // to insert_content instead - "translate this" appended a new
            // paragraph rather than replacing the selected one. 0-based,
            // matching every other paragraph index in this tool surface.
            int startBlockIndex = -1, endBlockIndex = -1;
            if (hasSelection)
            {
                Word.Paragraphs paragraphs = selection.Document.Paragraphs;
                int count = paragraphs.Count;
                for (int i = 0; i < count; i++)
                {
                    Word.Range r = paragraphs[i + 1].Range;
                    if (startBlockIndex == -1 && selection.Start < r.End) startBlockIndex = i;
                    if (selection.End <= r.End) { endBlockIndex = i; break; }
                }
                if (startBlockIndex == -1) startBlockIndex = count - 1;
                if (endBlockIndex == -1) endBlockIndex = count - 1;
            }

            // Post-hoc addition (2026-08-24, user-reported: selecting a
            // table/chart/SmartArt "doesn't appear under selection" -
            // selection.Text for a shape selection is empty or a placeholder
            // character, not a useful pointer. Detect these three object
            // kinds and report the SAME 0-based index the read_table/
            // read_chart/read_smartart tools use, so the UI and the model
            // both get an actionable pointer instead of nothing.
            string objectKind = null;
            int objectIndex = -1;
            try
            {
                // Diagnostic addition (2026-08-24): logs Word's own raw
                // Selection.Type - lets us tell definitively whether Word
                // itself considers a shape "selected" for a given click
                // (wdSelectionShape/wdSelectionInlineShape) versus the click
                // just moving the text cursor near the shape without
                // selecting it (which would still report Start==End,
                // withinTable=false, and an empty ShapeRange - all "normal"
                // Word behavior, not a bug in this detection code).
                DebugLog.Write("OnSelectionChanged: Selection.Type=" + selection.Type + " Start=" + selection.Start + " End=" + selection.End);
                bool withinTable = (bool)selection.get_Information(Word.WdInformation.wdWithInTable);
                int inlineCount = selection.InlineShapes.Count;
                DebugLog.Write("OnSelectionChanged: withinTable=" + withinTable + " inlineShapes.Count=" + inlineCount);
                if (withinTable)
                {
                    Word.Table selTable = selection.Tables[1];
                    Word.Tables allTables = selection.Document.Tables;
                    for (int i = 0; i < allTables.Count; i++)
                    {
                        if (allTables[i + 1].Range.Start == selTable.Range.Start) { objectKind = "table"; objectIndex = i; break; }
                    }
                }
                else if (inlineCount > 0)
                {
                    objectKind = ClassifySelectedShape(selection.Document, selection.InlineShapes[1], out objectIndex);
                }
                else
                {
                    dynamic shapeRange = selection.ShapeRange;
                    int shapeRangeCount = (int)shapeRange.Count;
                    DebugLog.Write("OnSelectionChanged: shapeRange.Count=" + shapeRangeCount);
                    if (shapeRangeCount > 0)
                    {
                        objectKind = ClassifySelectedShape(selection.Document, shapeRange[1], out objectIndex);
                    }
                }
                DebugLog.Write("OnSelectionChanged: resolved objectKind=" + (objectKind ?? "(null)") + " objectIndex=" + objectIndex);
            }
            catch (System.Exception ex)
            {
                // Post-hoc diagnostic addition (2026-08-24): was a silent
                // catch-all before - now logged, since "selection doesn't
                // show a pointer" could mean this is throwing every time
                // rather than just finding nothing.
                DebugLog.WriteException("OnSelectionChanged: object detection", ex);
            }

            // ROOT CAUSE FOUND (2026-08-24, via DebugLog from a real repro):
            // objectKind resolved correctly to "table" (confirmed in the
            // log) even though hasSelection (Start != End) was FALSE - a
            // plain cursor click/placement inside a table cell, or a click
            // directly on a chart/SmartArt shape, does NOT set Start != End
            // in Word's object model the way dragging across text does.
            // bootstrap.ts's toSelectionScopeUpdate/defaultDescribeSelection
            // both bail out immediately on hasSelection:false (line 1,
            // matching every other app's same-shaped payload), so a
            // correctly-detected object was being discarded before it ever
            // reached the UI or the model - this is what made table/chart/
            // SmartArt selection never show a pointer, regardless of how
            // correct the detection itself was. The EFFECTIVE hasSelection
            // sent downstream now also counts "an object was detected" as a
            // selection, since that is exactly as actionable as a text
            // range or an Excel/PowerPoint selection.
            bool effectiveHasSelection = hasSelection || objectKind != null;
            DebugLog.Write("OnSelectionChanged: hasSelection(raw)=" + hasSelection + " effectiveHasSelection=" + effectiveHasSelection);

            // FT-2 Task 1: routed through the shared debounce - WindowSelectionChange
            // fires on every caret move, same as Excel/PowerPoint's selection events.
            string signature = "word:" + selection.Start + "-" + selection.End;
            PostSelection(new
            {
                kind = "selection-changed",
                hasSelection = effectiveHasSelection,
                preview,
                fullText,
                startBlockIndex,
                endBlockIndex,
                objectKind,
                objectIndex,
            }, signature);
        }

        // Matches a selected shape (inline or floating) against
        // WordTools.ListChartShapes/ListSmartArtShapes by .Name, reusing
        // those tools' own addressing rather than a second copy of the
        // HasChart/HasSmartArt MsoTriState-comparison logic that could drift
        // out of sync with it.
        private static string ClassifySelectedShape(dynamic doc, dynamic shape, out int index)
        {
            index = -1;
            string name;
            try { name = (string)shape.Name; }
            catch (System.Exception ex) { DebugLog.WriteException("ClassifySelectedShape: shape.Name", ex); return null; }

            bool hasChart = false, hasSmartArt = false;
            try { hasChart = (int)shape.HasChart == -1; } catch (System.Exception ex) { DebugLog.WriteException("ClassifySelectedShape: shape.HasChart", ex); }
            try { hasSmartArt = (int)shape.HasSmartArt == -1; } catch (System.Exception ex) { DebugLog.WriteException("ClassifySelectedShape: shape.HasSmartArt", ex); }
            DebugLog.Write("ClassifySelectedShape: name=" + name + " hasChart=" + hasChart + " hasSmartArt=" + hasSmartArt);

            if (hasChart)
            {
                var charts = WordTools.ListChartShapes(doc);
                DebugLog.Write("ClassifySelectedShape: ListChartShapes returned " + charts.Count + " chart(s)");
                for (int i = 0; i < charts.Count; i++)
                {
                    string n = null; try { n = (string)charts[i].Name; } catch { }
                    DebugLog.Write("ClassifySelectedShape: chart[" + i + "].Name=" + n);
                    if (n == name) { index = i; return "chart"; }
                }
            }
            if (hasSmartArt)
            {
                var arts = WordTools.ListSmartArtShapes(doc);
                DebugLog.Write("ClassifySelectedShape: ListSmartArtShapes returned " + arts.Count + " diagram(s)");
                for (int i = 0; i < arts.Count; i++)
                {
                    string n = null; try { n = (string)arts[i].Name; } catch { }
                    DebugLog.Write("ClassifySelectedShape: smartart[" + i + "].Name=" + n);
                    if (n == name) { index = i; return "smartart"; }
                }
            }
            return null;
        }
    }
}
