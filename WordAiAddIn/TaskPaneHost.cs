using System.Diagnostics;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public partial class TaskPaneHost : PaneHostBase
    {
        private readonly Word.Document _document;
        private readonly int _hwnd;
        private string _chatId;

        // Deliberately does not dereference _document here - doing so at this exact
        // COM timing silently kills the add-in connection. See TaskPaneHost.cs.md.
        public TaskPaneHost(Word.Document document, int hwnd) : base("WordAiAddIn")
        {
            _document = document;
            _hwnd = hwnd;
        }

        protected override Task<ToolResult> ExecuteTool(string name, JsonElement input)
        {
            return Task.FromResult(WordTools.Execute(GetChatId(), name, input));
        }

        // A saved id is cached permanently; an "unsaved-" id is re-checked against
        // the document's Path on each call, so the first use after save migrates
        // chat history/doc settings onto the real per-file id. See TaskPaneHost.cs.md.
        protected override string GetChatId()
        {
            if (_chatId != null && !_chatId.StartsWith("unsaved-")) return _chatId;

            if (string.IsNullOrEmpty(_document.Path))
            {
                // Unsaved document: Document.FullName falls back to a temp Name
                // (e.g. "Document1"), not stable/unique across panes - pid+hwnd
                // avoids collisions between two different unsaved documents.
                return _chatId ?? (_chatId = "unsaved-" + Process.GetCurrentProcess().Id + "-" + _hwnd);
            }

            string saved = ChatStore.ChatIdForFile(_document.FullName);
            if (_chatId != null)
            {
                ChatStore.Migrate("WordAiAddIn", _chatId, saved);
                DocSettingsStore.Migrate("WordAiAddIn", _chatId, saved);
            }
            // Sticky: once set to a saved id, a later Save As does NOT re-key -
            // the conversation stays with the id first saved to, not the copy.
            // See TaskPaneHost.cs.md.
            return _chatId = saved;
        }

        protected override void SetEditingMode(EditingMode mode)
        {
            WordTools.SetMode(GetChatId(), mode);
        }

        protected override int GetOfficeUiLanguageId()
        {
            return Globals.ThisAddIn.GetOfficeUiLanguageId();
        }

        public void OnSelectionChanged(Word.Selection selection)
        {
            // selection.Text is null (not "") for a shape selection even though
            // Start != End - the null-coalesce below avoids an NRE on chart/SmartArt
            // clicks. See TaskPaneHost.cs.md.
            bool hasSelection = selection.Start != selection.End;
            string fullText = hasSelection ? (selection.Text ?? "") : "";
            if (fullText.Length > 24000) fullText = fullText.Substring(0, 24000);
            string preview = fullText.Length > 40 ? fullText.Substring(0, 40) : fullText;

            // 0-based start/end paragraph indices covering the selection, so the
            // model can target replace_blocks on exactly what's selected instead
            // of falling back to insert_content. See TaskPaneHost.cs.md.
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

            // Detects a selected table/chart/SmartArt and reports the same 0-based
            // index used by read_table/read_chart/read_smartart. See TaskPaneHost.cs.md.
            string objectKind = null;
            int objectIndex = -1;
            try
            {
                // Logs Word's raw Selection.Type to distinguish an actual shape
                // selection from the cursor merely being near a shape (both can
                // report Start==End). See TaskPaneHost.cs.md.
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
                // Logged rather than silently swallowed - "no pointer shown" could
                // mean this throws every time, not just finds nothing. See TaskPaneHost.cs.md.
                DebugLog.WriteException("OnSelectionChanged: object detection", ex);
            }

            // A plain click inside a table cell or on a chart/SmartArt shape does
            // NOT set Start != End the way dragging across text does, so a detected
            // object also counts as "has selection" downstream. See TaskPaneHost.cs.md.
            bool effectiveHasSelection = hasSelection || objectKind != null;
            DebugLog.Write("OnSelectionChanged: hasSelection(raw)=" + hasSelection + " effectiveHasSelection=" + effectiveHasSelection);

            // Routed through the shared debounce - WindowSelectionChange fires on
            // every caret move, same as Excel/PowerPoint's selection events.
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

        // Matches a selected shape against WordTools.ListChartShapes/ListSmartArtShapes
        // by .Name, reusing those tools' own addressing instead of a second copy of
        // the HasChart/HasSmartArt comparison logic.
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
