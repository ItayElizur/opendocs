using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    // Mirrors the Word/Excel Task 11/16 editing-mode enum pattern.
    public enum EditingMode { ReadOnly, CommentOnly, TrackChanges, FullAutonomy }

    public static class PowerPointTools
    {
        // Editing-mode gating (mirrors the Word/Excel Task 11/16 pattern this plan establishes
        // elsewhere): the tool list offered to the model is filtered client-side per mode
        // (web-src/entry.ts, first line of defense - smaller prompts, fewer wasted turns), but
        // Execute() independently re-checks mode here as defense-in-depth, since nothing stops a
        // misbehaving or malicious model response from calling a tool that wasn't offered.
        public static EditingMode Mode = EditingMode.FullAutonomy;

        // Tools always allowed regardless of editing mode (read-only, no document mutation).
        private static readonly System.Collections.Generic.HashSet<string> AlwaysAllowedTools =
            new System.Collections.Generic.HashSet<string> { "get_deck_context", "read_slide" };

        public static ToolResult Execute(string name, JsonElement input)
        {
            try
            {
                if (!AlwaysAllowedTools.Contains(name) && !IsMutationAllowed(Mode))
                {
                    return new ToolResult
                    {
                        Output = "Blocked: editing mode is " + ModeLabel(Mode) + ".",
                        IsError = true,
                        Summary = name,
                    };
                }

                switch (name)
                {
                    case "get_deck_context": return GetDeckContext();
                    case "read_slide": return ReadSlide(input);
                    case "set_element_text": return SetElementText(input);
                    case "set_element_style": return SetElementStyle(input);
                    case "set_element_transform": return SetElementTransform(input);
                    case "add_text_box": return AddTextBox(input);
                    case "add_shape": return AddShape(input);
                    case "delete_element": return DeleteElement(input);
                    case "add_slide": return AddSlide(input);
                    case "set_element_fill": return SetElementFill(input);
                    case "set_element_stroke": return SetElementStroke(input);
                    case "set_slide_background": return SetSlideBackground(input);
                    case "ungroup_element": return UngroupElement(input);
                    case "add_table": return AddTable(input);
                    case "edit_table_cell": return EditTableCell(input);
                    case "edit_table_structure": return EditTableStructure(input);
                    case "edit_table_style": return EditTableStyle(input);
                    case "add_chart": return AddChartPpt(input);
                    case "edit_chart": return EditChartPpt(input);
                    case "add_smartart": return AddSmartArt(input);
                    case "crop_image": return CropImage(input);
                    case "replace_image": return ReplaceImagePpt(input);
                    case "set_picture_opacity": return SetPictureOpacity(input);
                    default: return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        // Read Only and Comment Only modes block all mutating tools (PowerPoint has no
        // comment-equivalent tool in this pass - see plan backlog - so Comment Only currently
        // behaves identically to Read Only: no mutating tools available). Track Changes is scoped
        // to simple allow/block gating for now (same as Excel's Task 16 scoping note) rather than a
        // native PowerPoint revision-tracking UI. Full Autonomy allows everything.
        private static bool IsMutationAllowed(EditingMode mode)
        {
            return mode == EditingMode.TrackChanges || mode == EditingMode.FullAutonomy;
        }

        private static string ModeLabel(EditingMode mode)
        {
            switch (mode)
            {
                case EditingMode.ReadOnly: return "Read Only";
                case EditingMode.CommentOnly: return "Comment Only";
                case EditingMode.TrackChanges: return "Track Changes";
                default: return "Full Autonomy";
            }
        }

        private static PowerPoint.Presentation ActivePresentation => Globals.ThisAddIn.Application.ActivePresentation;

        private static string ShapeText(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                return shape.TextFrame.TextRange.Text;
            }
            return "";
        }

        private static ToolResult GetDeckContext()
        {
            var sb = new StringBuilder();
            int i = 0;
            foreach (PowerPoint.Slide slide in ActivePresentation.Slides)
            {
                var texts = new System.Collections.Generic.List<string>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    string t = ShapeText(shape).Replace("\r", " ").Trim();
                    if (t.Length > 0) texts.Add(t);
                }
                string preview = string.Join(" | ", texts);
                if (preview.Length > 120) preview = preview.Substring(0, 120) + "...";
                sb.AppendLine($"[{i}] {preview}");
                i++;
            }
            return new ToolResult { Output = sb.ToString(), Summary = "get_deck_context" };
        }

        private static ToolResult ReadSlide(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (slideIndex < 0 || slideIndex >= slides.Count)
            {
                return new ToolResult { Output = "Invalid slide index.", IsError = true, Summary = "read_slide" };
            }
            PowerPoint.Slide slide = slides[slideIndex + 1];
            var sb = new StringBuilder();
            int shapeIndex = 0;
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                sb.AppendLine($"[{shapeIndex}] {shape.Name}: {ShapeText(shape)}");
                shapeIndex++;
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_slide" };
        }

        private static PowerPoint.Shape ResolveShape(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int shapeIndex = input.GetProperty("shapeIndex").GetInt32();
            return ActivePresentation.Slides[slideIndex + 1].Shapes[shapeIndex + 1];
        }

        private static ToolResult SetElementText(JsonElement input)
        {
            ResolveShape(input).TextFrame.TextRange.Text = input.GetProperty("text").GetString();
            return new ToolResult { Output = "Text updated.", Mutated = true, Summary = "set_element_text" };
        }

        private static ToolResult SetElementStyle(JsonElement input)
        {
            PowerPoint.TextRange range = ResolveShape(input).TextFrame.TextRange;
            if (input.TryGetProperty("bold", out var bold)) range.Font.Bold = bold.GetBoolean() ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            if (input.TryGetProperty("italic", out var italic)) range.Font.Italic = italic.GetBoolean() ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            if (input.TryGetProperty("fontSize", out var fontSize)) range.Font.Size = (float)fontSize.GetDouble();
            if (input.TryGetProperty("color", out var color))
            {
                string hex = color.GetString().TrimStart('#');
                int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                range.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
            }
            return new ToolResult { Output = "Style updated.", Mutated = true, Summary = "set_element_style" };
        }

        private static ToolResult SetElementTransform(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            if (input.TryGetProperty("left", out var left)) shape.Left = (float)left.GetDouble();
            if (input.TryGetProperty("top", out var top)) shape.Top = (float)top.GetDouble();
            if (input.TryGetProperty("width", out var width)) shape.Width = (float)width.GetDouble();
            if (input.TryGetProperty("height", out var height)) shape.Height = (float)height.GetDouble();
            if (input.TryGetProperty("rotation", out var rotation)) shape.Rotation = (float)rotation.GetDouble();
            return new ToolResult { Output = "Transform updated.", Mutated = true, Summary = "set_element_transform" };
        }

        private static ToolResult AddTextBox(JsonElement input)
        {
            PowerPoint.Slide slide = ActivePresentation.Slides[input.GetProperty("slideIndex").GetInt32() + 1];
            float left = (float)input.GetProperty("left").GetDouble();
            float top = (float)input.GetProperty("top").GetDouble();
            float width = (float)input.GetProperty("width").GetDouble();
            float height = (float)input.GetProperty("height").GetDouble();
            PowerPoint.Shape shape = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            shape.TextFrame.TextRange.Text = input.GetProperty("text").GetString();
            return new ToolResult { Output = "Text box added.", Mutated = true, Summary = "add_text_box" };
        }

        private static ToolResult AddShape(JsonElement input)
        {
            PowerPoint.Slide slide = ActivePresentation.Slides[input.GetProperty("slideIndex").GetInt32() + 1];
            string shapeType = input.GetProperty("shapeType").GetString();
            Microsoft.Office.Core.MsoAutoShapeType autoShapeType =
                shapeType == "oval" ? Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval :
                shapeType == "roundRect" ? Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle :
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle;
            float left = (float)input.GetProperty("left").GetDouble();
            float top = (float)input.GetProperty("top").GetDouble();
            float width = (float)input.GetProperty("width").GetDouble();
            float height = (float)input.GetProperty("height").GetDouble();
            PowerPoint.Shape shape = slide.Shapes.AddShape(autoShapeType, left, top, width, height);
            if (input.TryGetProperty("text", out var text)) shape.TextFrame.TextRange.Text = text.GetString();
            return new ToolResult { Output = "Shape added.", Mutated = true, Summary = "add_shape" };
        }

        private static ToolResult DeleteElement(JsonElement input)
        {
            ResolveShape(input).Delete();
            return new ToolResult { Output = "Shape deleted.", Mutated = true, Summary = "delete_element" };
        }

        private static ToolResult AddSlide(JsonElement input)
        {
            int sourceIndex = input.GetProperty("sourceIndex").GetInt32();
            bool clearText = !input.TryGetProperty("clearText", out var ct) || ct.ValueKind != JsonValueKind.False;
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (sourceIndex < 0 || sourceIndex >= slides.Count)
            {
                return new ToolResult { Output = "Invalid sourceIndex.", IsError = true, Summary = "add_slide" };
            }
            dynamic source = slides[sourceIndex + 1];
            dynamic dupRange = source.Duplicate(); // returns a SlideRange containing exactly the new slide
            dynamic newSlide = dupRange[1];
            newSlide.MoveTo(sourceIndex + 2);
            if (clearText)
            {
                foreach (PowerPoint.Shape shape in newSlide.Shapes)
                {
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        shape.TextFrame.TextRange.Text = "";
                    }
                }
            }
            return new ToolResult { Output = "Slide added after index " + sourceIndex + ".", Mutated = true, Summary = "add_slide" };
        }

        private static int HexToOle(string hex)
        {
            hex = hex.TrimStart('#');
            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = Convert.ToInt32(hex.Substring(4, 2), 16);
            return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
        }

        private static ToolResult SetElementFill(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            string fill = input.GetProperty("fill").GetString();
            if (fill == "none")
            {
                shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            else
            {
                shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                shape.Fill.ForeColor.RGB = HexToOle(fill);
            }
            return new ToolResult { Output = "Fill updated.", Mutated = true, Summary = "set_element_fill" };
        }

        private static ToolResult SetElementStroke(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            bool remove = input.TryGetProperty("remove", out var r) && r.ValueKind == JsonValueKind.True;
            if (remove)
            {
                shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            else
            {
                shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                if (input.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
                {
                    shape.Line.ForeColor.RGB = HexToOle(color.GetString());
                }
                shape.Line.Weight = input.TryGetProperty("widthPt", out var width) && width.ValueKind == JsonValueKind.Number ? (float)width.GetDouble() : 1f;
            }
            return new ToolResult { Output = "Stroke updated.", Mutated = true, Summary = "set_element_stroke" };
        }

        private static ToolResult SetSlideBackground(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int oleColor = HexToOle(input.GetProperty("color").GetString());
            PowerPoint.Slides slides = ActivePresentation.Slides;

            void Apply(PowerPoint.Slide s)
            {
                s.Background.Fill.ForeColor.RGB = oleColor;
                s.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
            }

            if (slideIndex == -1)
            {
                foreach (PowerPoint.Slide s in slides) Apply(s);
            }
            else
            {
                Apply(slides[slideIndex + 1]);
            }
            return new ToolResult { Output = "Background updated.", Mutated = true, Summary = "set_slide_background" };
        }

        private static ToolResult UngroupElement(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            shape.Ungroup();
            return new ToolResult { Output = "Shape ungrouped - re-read the slide (read_slide) to get updated shape indices before addressing the promoted children.", Mutated = true, Summary = "ungroup_element" };
        }

        private static ToolResult AddTable(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int rows = input.GetProperty("rows").GetInt32();
            int cols = input.GetProperty("cols").GetInt32();
            float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
            float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
            float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
            float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 200f;

            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            PowerPoint.Shape tableShape = slide.Shapes.AddTable(rows, cols, left, top, width, height);
            if (input.TryGetProperty("cells", out var cells) && cells.ValueKind == JsonValueKind.Array)
            {
                int r = 0;
                foreach (JsonElement rowEl in cells.EnumerateArray())
                {
                    int c = 0;
                    foreach (JsonElement cellEl in rowEl.EnumerateArray())
                    {
                        tableShape.Table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange.Text = cellEl.GetString();
                        c++;
                    }
                    r++;
                }
            }
            return new ToolResult { Output = "Table added.", Mutated = true, Summary = "add_table" };
        }

        private static PowerPoint.Table ResolveTable(JsonElement input)
        {
            return ResolveShape(input).Table;
        }

        private static ToolResult EditTableCell(JsonElement input)
        {
            PowerPoint.Table table = ResolveTable(input);
            int row = input.GetProperty("row").GetInt32();
            int col = input.GetProperty("col").GetInt32();
            string text = input.GetProperty("paragraphs").GetString();
            table.Cell(row + 1, col + 1).Shape.TextFrame.TextRange.Text = text;
            return new ToolResult { Output = "Cell updated.", Mutated = true, Summary = "edit_table_cell" };
        }

        private static ToolResult EditTableStructure(JsonElement input)
        {
            PowerPoint.Table table = ResolveTable(input);
            string kind = input.GetProperty("kind").GetString();
            int index = input.GetProperty("index").GetInt32();
            bool before = input.TryGetProperty("before", out var b) && b.ValueKind == JsonValueKind.True;
            switch (kind)
            {
                case "insert-row": table.Rows.Add(before ? index + 1 : index + 2); break;
                case "delete-row": table.Rows[index + 1].Delete(); break;
                case "insert-col": table.Columns.Add(before ? index + 1 : index + 2); break;
                case "delete-col": table.Columns[index + 1].Delete(); break;
                default: return new ToolResult { Output = "Unknown structure kind: " + kind, IsError = true, Summary = "edit_table_structure" };
            }
            return new ToolResult { Output = "Table structure updated.", Mutated = true, Summary = "edit_table_structure" };
        }

        private static ToolResult EditTableStyle(JsonElement input)
        {
            PowerPoint.Table table = ResolveTable(input);
            if (input.TryGetProperty("firstRow", out var firstRow))
            {
                table.FirstRow = firstRow.ValueKind == JsonValueKind.True;
            }
            if (input.TryGetProperty("bandRow", out var bandRow))
            {
                table.HorizBanding = bandRow.ValueKind == JsonValueKind.True;
            }
            if (input.TryGetProperty("shadingColor", out var shading) && shading.ValueKind == JsonValueKind.String)
            {
                int color = HexToOle(shading.GetString());
                foreach (PowerPoint.Row row in table.Rows)
                {
                    foreach (PowerPoint.Cell cell in row.Cells)
                    {
                        cell.Shape.Fill.ForeColor.RGB = color;
                    }
                }
            }
            if (input.TryGetProperty("borderColor", out _) || input.TryGetProperty("borderWidthPt", out _) || input.TryGetProperty("borderPreset", out _))
            {
                bool visible = !(input.TryGetProperty("borderPreset", out var bp) && bp.ValueKind == JsonValueKind.String && bp.GetString() == "none");
                float weight = input.TryGetProperty("borderWidthPt", out var bw) && bw.ValueKind == JsonValueKind.Number ? (float)bw.GetDouble() : 1f;
                int color = input.TryGetProperty("borderColor", out var bc) && bc.ValueKind == JsonValueKind.String ? HexToOle(bc.GetString()) : HexToOle("#000000");
                PowerPoint.PpBorderType[] sides = { PowerPoint.PpBorderType.ppBorderTop, PowerPoint.PpBorderType.ppBorderBottom, PowerPoint.PpBorderType.ppBorderLeft, PowerPoint.PpBorderType.ppBorderRight };
                foreach (PowerPoint.Row row in table.Rows)
                {
                    foreach (PowerPoint.Cell cell in row.Cells)
                    {
                        foreach (PowerPoint.PpBorderType side in sides)
                        {
                            PowerPoint.LineFormat border = cell.Borders[side];
                            border.Visible = visible ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
                            border.Weight = weight;
                            border.ForeColor.RGB = color;
                        }
                    }
                }
            }
            return new ToolResult { Output = "Table style updated.", Mutated = true, Summary = "edit_table_style" };
        }

        private static readonly Dictionary<string, int> PptChartTypeMap = new Dictionary<string, int>
        {
            ["bar"] = 51,          // xlColumnClustered
            ["barStacked"] = 52,   // xlColumnStacked
            ["line"] = 4,          // xlLine
            ["area"] = 1,          // xlArea
            ["pie"] = 5,           // xlPie
            ["doughnut"] = -4120,  // xlDoughnut
        };

        private static ToolResult AddChartPpt(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            string kindStr = input.GetProperty("kind").GetString();
            int typeCode = PptChartTypeMap.TryGetValue(kindStr, out var t) ? t : 51;
            var categories = new List<string>();
            foreach (JsonElement c in input.GetProperty("categories").EnumerateArray()) categories.Add(c.GetString());
            float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
            float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
            float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
            float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 300f;

            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            dynamic chartShape = slide.Shapes.AddChart2(-1, (Microsoft.Office.Core.XlChartType)typeCode, left, top, width, height);
            dynamic chart = chartShape.Chart;

            // Chart data lives in an embedded Excel workbook - open, write the grid,
            // close, and RELEASE explicitly so no hidden Excel host process leaks.
            dynamic dataWorkbook = chart.ChartData.Workbook;
            try
            {
                dynamic sheet = dataWorkbook.Worksheets[1];
                JsonElement seriesArray = input.GetProperty("series");
                int colIdx = 0;
                foreach (JsonElement s in seriesArray.EnumerateArray())
                {
                    sheet.Cells[1, colIdx + 2].Value = s.GetProperty("name").GetString();
                    colIdx++;
                }
                for (int r = 0; r < categories.Count; r++)
                {
                    sheet.Cells[r + 2, 1].Value = categories[r];
                }
                colIdx = 0;
                foreach (JsonElement s in seriesArray.EnumerateArray())
                {
                    int r = 0;
                    foreach (JsonElement v in s.GetProperty("values").EnumerateArray())
                    {
                        sheet.Cells[r + 2, colIdx + 2].Value = v.GetDouble();
                        r++;
                    }
                    colIdx++;
                }
                dynamic usedRange = sheet.UsedRange;
                chart.SetSourceData(usedRange);
            }
            finally
            {
                try
                {
                    dataWorkbook.Close(SaveChanges: true);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dataWorkbook);
                }
            }

            if (input.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
            }
            return new ToolResult { Output = "Chart added.", Mutated = true, Summary = "add_chart" };
        }

        private static ToolResult EditChartPpt(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            dynamic chart = shape.Chart;

            if (input.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String && PptChartTypeMap.TryGetValue(ct.GetString(), out var typeCode))
            {
                chart.ChartType = typeCode;
            }
            if (input.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
            }
            if (input.TryGetProperty("legendPos", out var legendPos) && legendPos.ValueKind == JsonValueKind.String)
            {
                string pos = legendPos.GetString();
                if (pos == "none")
                {
                    chart.HasLegend = false;
                }
                else
                {
                    chart.HasLegend = true;
                    chart.Legend.Position = pos == "r" ? -4152 : pos == "t" ? -4160 : pos == "l" ? -4131 : -4107;
                }
            }
            if (input.TryGetProperty("dataLabels", out var dl))
            {
                bool show = dl.ValueKind == JsonValueKind.True;
                foreach (dynamic series in chart.SeriesCollection())
                {
                    series.HasDataLabels = show;
                }
            }
            if (input.TryGetProperty("gridlines", out var gl))
            {
                chart.Axes(2 /* xlValue */).HasMajorGridlines = gl.ValueKind == JsonValueKind.True;
            }
            return new ToolResult { Output = "Chart updated.", Mutated = true, Summary = "edit_chart" };
        }

        // Verified against the standard English display names for PowerPoint's built-in SmartArt
        // layout gallery. Live cross-check against Application.SmartArtLayouts on this machine's
        // Office install (plan Task 6 Step 1) requires interactive Office GUI access that was not
        // available in this environment - remains a manual follow-up for a human with GUI access.
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
            string targetName = SmartArtLayoutNames.TryGetValue(layoutKey, out var name) ? name : "Basic Block List";
            dynamic layouts = Globals.ThisAddIn.Application.SmartArtLayouts;
            foreach (dynamic layout in layouts)
            {
                if (string.Equals((string)layout.Name, targetName, StringComparison.OrdinalIgnoreCase))
                {
                    return layout;
                }
            }
            throw new InvalidOperationException("add_smartart: no SmartArt layout named '" + targetName + "' found - see plan Task 6 Step 1.");
        }

        private static ToolResult AddSmartArt(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            string layoutKey = input.GetProperty("layout").GetString();
            float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
            float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
            float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
            float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 300f;

            dynamic layout = ResolveSmartArtLayout(layoutKey);
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            dynamic shape = slide.Shapes.AddSmartArt(layout, left, top, width, height);
            dynamic smartArt = shape.SmartArt;

            // genoffice's own version only ever produces a flat item list - maps
            // 1:1 to sequential top-level nodes, no nested tree-building needed.
            foreach (JsonElement item in input.GetProperty("items").EnumerateArray())
            {
                dynamic node = smartArt.Nodes.Add();
                node.TextFrame2.TextRange.Text = item.GetString();
            }
            return new ToolResult { Output = "SmartArt added.", Mutated = true, Summary = "add_smartart" };
        }

        private static ToolResult CropImage(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            float l = (float)input.GetProperty("l").GetDouble();
            float t = (float)input.GetProperty("t").GetDouble();
            float r = (float)input.GetProperty("r").GetDouble();
            float b = (float)input.GetProperty("b").GetDouble();
            // Approximation, documented deliberately: fractions are applied against
            // the shape's CURRENT on-slide size, not the original uncropped source
            // image - classic Interop has no reliable "natural size" property once a
            // picture has already been resized/cropped on the slide. Correct for a
            // freshly-inserted, never-before-cropped picture; imprecise under
            // repeated crop calls on the same shape.
            shape.PictureFormat.CropLeft = l * shape.Width;
            shape.PictureFormat.CropTop = t * shape.Height;
            shape.PictureFormat.CropRight = r * shape.Width;
            shape.PictureFormat.CropBottom = b * shape.Height;
            return new ToolResult { Output = "Image cropped.", Mutated = true, Summary = "crop_image" };
        }

        private static ToolResult ReplaceImagePpt(JsonElement input)
        {
            string localPath = input.GetProperty("localPath").GetString();
            if (localPath.StartsWith("http://") || localPath.StartsWith("https://"))
            {
                return new ToolResult { Output = "replace_image: remote URLs are not supported in this air-gapped deployment - use a local file path.", IsError = true, Summary = "replace_image" };
            }
            if (!System.IO.File.Exists(localPath))
            {
                return new ToolResult { Output = "replace_image: file not found: " + localPath, IsError = true, Summary = "replace_image" };
            }
            PowerPoint.Shape oldShape = ResolveShape(input);
            bool keepCrop = input.TryGetProperty("keepCrop", out var kc) && kc.ValueKind == JsonValueKind.True;

            float left = oldShape.Left, top = oldShape.Top, width = oldShape.Width, height = oldShape.Height, rotation = oldShape.Rotation;
            int zPos = oldShape.ZOrderPosition;
            float cropLeft = 0, cropTop = 0, cropRight = 0, cropBottom = 0;
            if (keepCrop)
            {
                cropLeft = oldShape.PictureFormat.CropLeft;
                cropTop = oldShape.PictureFormat.CropTop;
                cropRight = oldShape.PictureFormat.CropRight;
                cropBottom = oldShape.PictureFormat.CropBottom;
            }
            PowerPoint.Slide slide = (PowerPoint.Slide)oldShape.Parent;
            oldShape.Delete();

            PowerPoint.Shape newShape = slide.Shapes.AddPicture(localPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, left, top, width, height);
            newShape.Rotation = rotation;
            if (keepCrop)
            {
                newShape.PictureFormat.CropLeft = cropLeft;
                newShape.PictureFormat.CropTop = cropTop;
                newShape.PictureFormat.CropRight = cropRight;
                newShape.PictureFormat.CropBottom = cropBottom;
            }
            // Restore approximate z-order: send to back, then bring forward to the
            // original stack position.
            newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
            for (int i = 1; i < zPos; i++)
            {
                newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringForward);
            }
            return new ToolResult { Output = "Image replaced.", Mutated = true, Summary = "replace_image" };
        }

        private static ToolResult SetPictureOpacity(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            float opacity = (float)input.GetProperty("opacity").GetDouble();
            dynamic dShape = shape;
            dShape.Fill.Transparency = 1f - opacity;
            return new ToolResult { Output = "Opacity updated.", Mutated = true, Summary = "set_picture_opacity" };
        }
    }
}
