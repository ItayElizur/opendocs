using System.Collections.Generic;
using System.Diagnostics;
using System.Text.Json;
using OfficeAi.Shared;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAiAddIn
{
    public partial class TaskPaneHost : PaneHostBase
    {
        private readonly PowerPoint.Presentation _presentation;
        private readonly int _hwnd;
        private string _chatId;

        // Deliberately does NOT dereference _presentation here (no
        // .Path/.FullName read) - see WordAiAddIn/TaskPaneHost.cs's identical
        // comment for the confirmed repro of why an eager read at
        // construction time silently kills the whole add-in connection.
        public TaskPaneHost(PowerPoint.Presentation presentation, int hwnd) : base("PowerPointAiAddIn")
        {
            _presentation = presentation;
            _hwnd = hwnd;
        }

        protected override ToolResult ExecuteTool(string name, JsonElement input)
        {
            return PowerPointTools.Execute(GetChatId(), name, input);
        }

        protected override string GetChatId()
        {
            // A saved id is final - never re-checked again. An "unsaved-" id
            // is provisional: re-check the presentation's Path on every call,
            // so the first use after the user saves migrates chat history and
            // doc settings onto the real per-file id (FT-1 Task 7b). The Path
            // read is one cheap COM property on operations (load-history,
            // append-message, etc.) that are already doing file I/O.
            if (_chatId != null && !_chatId.StartsWith("unsaved-")) return _chatId;

            if (string.IsNullOrEmpty(_presentation.Path))
            {
                // An unsaved presentation has no on-disk Path; Presentation.FullName
                // falls back to its temp Name (e.g. "Presentation1") in that case,
                // which is not a stable key across sessions - and with multiple
                // panes now possible in one process, "unsaved-<pid>" alone would
                // collide across two different unsaved presentations, so the
                // window handle is folded in too.
                return _chatId ?? (_chatId = "unsaved-" + Process.GetCurrentProcess().Id + "-" + _hwnd);
            }

            string saved = ChatStore.ChatIdForFile(_presentation.FullName);
            if (_chatId != null)
            {
                ChatStore.Migrate("PowerPointAiAddIn", _chatId, saved);
                DocSettingsStore.Migrate("PowerPointAiAddIn", _chatId, saved);
            }
            // Save As after this point does NOT re-key - see WordAiAddIn/
            // TaskPaneHost.cs's identical comment for the rationale.
            return _chatId = saved;
        }

        protected override void SetEditingMode(EditingMode mode)
        {
            PowerPointTools.SetMode(GetChatId(), mode);
        }

        // FT-2 Task 3: called from ThisAddIn's WindowSelectionChange handler.
        // Debounced through PaneHostBase.PostSelection (Task 1) -
        // WindowSelectionChange fires on every shape click during ordinary
        // editing and would flood the WebView2 bridge otherwise.
        public void OnSelectionChanged(PowerPoint.Selection sel)
        {
            switch (sel.Type)
            {
                case PowerPoint.PpSelectionType.ppSelectionSlides:
                    OnSlidesSelected(sel);
                    break;
                case PowerPoint.PpSelectionType.ppSelectionShapes:
                    OnShapesSelected(sel);
                    break;
                case PowerPoint.PpSelectionType.ppSelectionText:
                    OnTextSelected(sel);
                    break;
                default:
                    PostSelection(new { kind = "selection-changed", app = "powerpoint", hasSelection = false }, "ppt:none");
                    break;
            }
        }

        // Task 3 Step 7: the slide sorter's multi-select reports the whole list.
        private void OnSlidesSelected(PowerPoint.Selection sel)
        {
            var indexes = new List<int>();
            // Task 3 Step 3: Slides is 1-based in COM, ResolveShape's tools
            // are 0-based - convert once, here, and never again downstream.
            foreach (PowerPoint.Slide s in sel.SlideRange) indexes.Add(s.SlideIndex - 1);
            string signature = "ppt:slides:" + string.Join(",", indexes);
            PostSelection(new
            {
                kind = "selection-changed",
                app = "powerpoint",
                selKind = "slides",
                hasSelection = true,
                slideIndexes = indexes,
            }, signature);
        }

        private void OnShapesSelected(PowerPoint.Selection sel)
        {
            PowerPoint.Slide slide = sel.SlideRange[1];
            int slideIndex = slide.SlideIndex - 1;
            var shapeIndexes = new List<int>();
            var names = new List<string>();
            var previews = new List<string>();
            int count = 0;
            foreach (PowerPoint.Shape shape in sel.ShapeRange)
            {
                shapeIndexes.Add(ShapeIndexInSlide(slide, shape));
                // Task 3 Step 4: names are far more meaningful to a model than
                // a bare index and cost nothing to include.
                names.Add(shape.Name);
                // Task 3 Step 5: capped preview, at most 5 shapes - HasTextFrame
                // must be checked first, reading TextRange.Text on a shape with
                // no text frame throws.
                if (count < 5)
                {
                    string preview = "";
                    try
                    {
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue &&
                            shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            preview = shape.TextFrame.TextRange.Text;
                            if (preview.Length > 80) preview = preview.Substring(0, 80);
                        }
                    }
                    catch { /* best-effort preview only */ }
                    previews.Add(preview);
                }
                count++;
            }
            string signature = "ppt:shapes:" + slideIndex + ":" + string.Join(",", shapeIndexes);
            PostSelection(new
            {
                kind = "selection-changed",
                app = "powerpoint",
                selKind = "shapes",
                hasSelection = true,
                slideIndex,
                shapeIndexes,
                names,
                textPreview = previews,
            }, signature);
        }

        // Task 3 Step 6: a text cursor/selection inside one shape - a
        // sub-selection of the shape, not the shape itself.
        private void OnTextSelected(PowerPoint.Selection sel)
        {
            PowerPoint.Slide slide = sel.SlideRange[1];
            int slideIndex = slide.SlideIndex - 1;
            PowerPoint.Shape shape = sel.ShapeRange[1];
            int shapeIndex = ShapeIndexInSlide(slide, shape);
            string text = "";
            try { text = sel.TextRange.Text ?? ""; }
            catch { /* best-effort */ }
            if (text.Length > 2000) text = text.Substring(0, 2000);
            string signature = "ppt:text:" + slideIndex + ":" + shapeIndex + ":" + text.GetHashCode();
            PostSelection(new
            {
                kind = "selection-changed",
                app = "powerpoint",
                selKind = "shapeText",
                hasSelection = true,
                slideIndex,
                shapeIndex,
                text,
            }, signature);
        }

        // Shapes selected via Sel.ShapeRange don't carry their own 0-based
        // position within the slide's Shapes collection (the form
        // ResolveShape/the tools take) - Id is unique per slide and stable,
        // unlike Name, so this searches for it once per selected shape.
        private static int ShapeIndexInSlide(PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                if (slide.Shapes[i].Id == shape.Id) return i - 1;
            }
            return -1;
        }
    }
}
