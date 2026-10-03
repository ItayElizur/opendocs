using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public static partial class ExcelTools
    {
        // Forces every series' XValues to an explicit range, since Excel's category-label auto-detection
        // fails for a numeric leftmost column/top row - see ExcelTools.Charts.cs.md.
        private static void ApplyChartCategoryRange(dynamic chart, Excel.Range categories)
        {
            dynamic seriesCollection = chart.SeriesCollection();
            int count = seriesCollection.Count;
            for (int i = 1; i <= count; i++)
            {
                seriesCollection.Item(i).XValues = categories;
            }
        }

        private static string AddChart(JsonElement op)
        {
            Excel.Worksheet sheet = Sheet(op);
            string dataRange = op.GetProperty("dataRange").GetString();
            dynamic chartObjects = sheet.ChartObjects();
            dynamic chartObj = chartObjects.Add(100, 20, 400, 250);
            dynamic chart = chartObj.Chart;
            chart.SetSourceData(sheet.Range[dataRange]);
            if (op.TryGetProperty("categoryRange", out var catEl) && catEl.ValueKind == JsonValueKind.String)
            {
                ApplyChartCategoryRange(chart, sheet.Range[catEl.GetString()]);
            }
            int chartTypeCode = 51; // xlColumnClustered
            if (op.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String)
            {
                if (!ChartTypes.ByName.TryGetValue(ct.GetString(), out chartTypeCode))
                    throw new ArgumentException("add_chart: unknown chartType '" + ct.GetString() +
                                                "'. Valid: " + string.Join(", ", ChartTypes.ByName.Keys) + ".");
            }
            chart.ChartType = chartTypeCode;
            if (op.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
            }
            if (op.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String)
            {
                chartObj.Name = nameEl.GetString();
            }
            // Include the name so a follow-up edit_chart in the same batch can address it without guessing
            // Excel's auto-assigned "Chart 1"-style name (PP-15 Task 4).
            return "name=" + (string)chartObj.Name;
        }

        private static void EditChartExcel(JsonElement op)
        {
            string chartName = op.GetProperty("chartPath").GetString(); // this project's visualId-equivalent for charts
            dynamic chartObjects = Sheet(op).ChartObjects();
            dynamic chartObj = chartObjects.Item(chartName);
            dynamic chart = chartObj.Chart;

            // Rebind before ChartType below: some chart-type changes reset the plot (PP-15 Task 2).
            // dataSheet lets the chart live on a different sheet than its data.
            if (op.TryGetProperty("dataRange", out var dr) && dr.ValueKind == JsonValueKind.String)
            {
                Excel.Worksheet dataSheet = op.TryGetProperty("dataSheet", out var dsEl) && dsEl.ValueKind == JsonValueKind.String
                    ? (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[dsEl.GetString()]
                    : Sheet(op);
                Excel.Range source = dataSheet.Range[dr.GetString()];
                if (op.TryGetProperty("plotBy", out var pb) && pb.ValueKind == JsonValueKind.String)
                {
                    Excel.XlRowCol plotBy;
                    if (pb.GetString() == "rows") plotBy = Excel.XlRowCol.xlRows;
                    else if (pb.GetString() == "columns") plotBy = Excel.XlRowCol.xlColumns;
                    else throw new ArgumentException("edit_chart: unknown plotBy '" + pb.GetString() + "'. Valid: rows, columns.");
                    chart.SetSourceData(source, plotBy);
                }
                else
                {
                    chart.SetSourceData(source); // Excel infers orientation from the range shape when omitted
                }
            }

            if (op.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String)
            {
                int typeCode;
                if (!ChartTypes.ByName.TryGetValue(ct.GetString(), out typeCode))
                    throw new ArgumentException("edit_chart: unknown chartType '" + ct.GetString() +
                                                "'. Valid: " + string.Join(", ", ChartTypes.ByName.Keys) + ".");
                chart.ChartType = typeCode;
            }

            // Independent of dataRange/plotBy above - fixes the x-axis to an
            // explicit column/row even when the chart isn't otherwise being
            // rebound (see ApplyChartCategoryRange).
            if (op.TryGetProperty("categoryRange", out var catEl) && catEl.ValueKind == JsonValueKind.String)
            {
                Excel.Worksheet categorySheet = op.TryGetProperty("dataSheet", out var catDsEl) && catDsEl.ValueKind == JsonValueKind.String
                    ? (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[catDsEl.GetString()]
                    : Sheet(op);
                ApplyChartCategoryRange(chart, categorySheet.Range[catEl.GetString()]);
            }
            if (op.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
            }
            if (op.TryGetProperty("legend", out var legend) && legend.ValueKind == JsonValueKind.String)
            {
                // Errors on an unmatched value instead of silently defaulting to bottom (PP-21 Task 2 Step 5) -
                // see ExcelTools.Charts.cs.md.
                string pos = legend.GetString();
                if (pos == "none") { chart.HasLegend = false; }
                else
                {
                    int position;
                    switch (pos)
                    {
                        case "right": position = -4152; break; // xlLegendPositionRight
                        case "top": position = -4160; break;   // xlLegendPositionTop
                        case "left": position = -4131; break;  // xlLegendPositionLeft
                        case "bottom": position = -4107; break; // xlLegendPositionBottom
                        default:
                            throw new ArgumentException("edit_chart: unknown legend '" + pos + "'. Valid: none, right, top, left, bottom.");
                    }
                    chart.HasLegend = true;
                    chart.Legend.Position = position;
                }
            }
            if (op.TryGetProperty("dataLabels", out var dl) && dl.ValueKind == JsonValueKind.String)
            {
                bool show = dl.GetString() != "none";
                foreach (dynamic series in chart.SeriesCollection())
                {
                    series.HasDataLabels = show;
                    if (show && dl.GetString() == "percent") series.DataLabels().ShowPercentage = true;
                }
            }
            if (op.TryGetProperty("seriesColors", out var colors) && colors.ValueKind == JsonValueKind.Object)
            {
                foreach (JsonProperty prop in colors.EnumerateObject())
                {
                    int seriesIndex = int.Parse(prop.Name);
                    dynamic series = chart.SeriesCollection(seriesIndex + 1);
                    series.Format.Fill.ForeColor.RGB = ColorUtil.HexToOle(prop.Value.GetString());
                }
            }

            if (op.TryGetProperty("seriesData", out var seriesData) && seriesData.ValueKind == JsonValueKind.Array)
            {
                int seriesIdx = 0;
                foreach (JsonElement sd in seriesData.EnumerateArray())
                {
                    dynamic series = chart.SeriesCollection(seriesIdx + 1);
                    if (sd.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String)
                    {
                        series.Name = nameEl.GetString();
                    }
                    seriesIdx++;
                }
            }
        }

        // Returns the target address as the addressing handle for a later edit - SparklineGroup exposes no
        // separate stable id (PP-18 Task 3 Step 5) - see ExcelTools.Charts.cs.md.
        private static string AddSparkline(JsonElement op)
        {
            string dataRange = op.GetProperty("dataRange").GetString();
            // Required, not defaulted to dataRange - that default silently drew the sparkline over its own
            // source data (prior bug).
            if (!op.TryGetProperty("targetCell", out var tc) || tc.ValueKind != JsonValueKind.String)
                throw new ArgumentException("add_sparkline: 'targetCell' is required (the cell immediately right of, or below, dataRange is the usual choice - it must not overlap dataRange).");
            string targetCell = tc.GetString();

            string type = "line";
            if (op.TryGetProperty("type", out var t) && t.ValueKind == JsonValueKind.String)
            {
                type = t.GetString();
                if (type != "line" && type != "column" && type != "stacked")
                    throw new ArgumentException("add_sparkline: unknown type '" + type + "'. Valid: line, column, stacked.");
            }

            Excel.Worksheet sheet = Sheet(op);
            Excel.Range dataRangeObj = sheet.Range[dataRange];
            Excel.Range targetRangeObj = sheet.Range[targetCell];

            // One sparkline per row (multi-cell target matching row count) or one for the whole range
            // (single-cell target) - anything else fails opaquely in the COM call below.
            int dataRows = dataRangeObj.Rows.Count;
            int targetCells = targetRangeObj.Cells.Count;
            bool overlaps = string.Equals(
                dataRangeObj.Address[true, true, Excel.XlReferenceStyle.xlA1, false],
                targetRangeObj.Address[true, true, Excel.XlReferenceStyle.xlA1, false],
                StringComparison.OrdinalIgnoreCase);
            if (overlaps)
                throw new ArgumentException("add_sparkline: targetCell must not be the same as dataRange - the sparkline would draw over its own source data.");
            if (targetCells != 1 && targetCells != dataRows)
                throw new ArgumentException("add_sparkline: targetCell has " + targetCells + " cell(s) but dataRange has " +
                    dataRows + " row(s) - targetCell must be a single cell (one sparkline for the whole range) or match the row count (one sparkline per row).");

            dynamic groups = targetRangeObj.SparklineGroups;
            Excel.XlSparkType sparkType = type == "column" ? Excel.XlSparkType.xlSparkColumn
                : type == "stacked" ? Excel.XlSparkType.xlSparkColumnStacked100
                : Excel.XlSparkType.xlSparkLine;
            dynamic group = groups.Add(sparkType, dataRangeObj.Address[true, true, Excel.XlReferenceStyle.xlA1, true]);
            if (op.TryGetProperty("color", out var color) && color.ValueKind == JsonValueKind.String)
            {
                group.SeriesColor.Color = ColorUtil.HexToOle(color.GetString());
            }
            return "target=" + targetCell;
        }

        // Returns the shape's name so a follow-up edit_shape/delete_visual can address it (PP-16 Task 3).
        private static string AddShapeExcel(JsonElement op)
        {
            string shapeType = op.GetProperty("shapeType").GetString();
            string anchorCell = op.GetProperty("anchorCell").GetString();
            Excel.Range anchor = Sheet(op).Range[anchorCell];
            float left = (float)(double)anchor.Left;
            float top = (float)(double)anchor.Top;
            float width = 100f, height = 60f;

            Excel.Shape shape;
            if (string.Equals(shapeType, "textbox", StringComparison.OrdinalIgnoreCase))
            {
                shape = Sheet(op).Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            }
            else
            {
                int msoTypeInt;
                if (!ShapeTypes.ByName.TryGetValue(shapeType, out msoTypeInt))
                    throw new ArgumentException("add_shape: unknown shapeType '" + shapeType + "'. Valid: textbox, " +
                                                string.Join(", ", ShapeTypes.ByName.Keys) + ".");
                shape = Sheet(op).Shapes.AddShape((Microsoft.Office.Core.MsoAutoShapeType)msoTypeInt, left, top, width, height);
            }
            if (op.TryGetProperty("fillColor", out var fill) && fill.ValueKind == JsonValueKind.String)
            {
                shape.Fill.ForeColor.RGB = ColorUtil.HexToOle(fill.GetString());
            }
            if (op.TryGetProperty("text", out var text) && text.ValueKind == JsonValueKind.String)
            {
                shape.TextFrame.Characters().Text = text.GetString();
            }
            if (op.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String)
            {
                shape.Name = nameEl.GetString();
            }
            return shape.Name;
        }

        private static Excel.Shape ResolveShapeByName(JsonElement op, string idField)
        {
            string visualId = op.GetProperty(idField).GetString();
            return Sheet(op).Shapes.Item(visualId);
        }

        private static void EditShapeExcel(JsonElement op)
        {
            Excel.Shape shape = ResolveShapeByName(op, "visualId");
            try
            {
                if (op.TryGetProperty("text", out var text) && text.ValueKind == JsonValueKind.String)
                {
                    shape.TextFrame.Characters().Text = text.GetString();
                }
            }
            catch (System.Runtime.InteropServices.COMException) { /* shape doesn't support a text frame */ }
            if (op.TryGetProperty("fillColor", out var fill) && fill.ValueKind == JsonValueKind.String)
            {
                shape.Fill.ForeColor.RGB = ColorUtil.HexToOle(fill.GetString());
            }
            if (op.TryGetProperty("anchorCell", out var anchorCell) && anchorCell.ValueKind == JsonValueKind.String)
            {
                Excel.Range anchor = Sheet(op).Range[anchorCell.GetString()];
                shape.Left = (float)(double)anchor.Left;
                shape.Top = (float)(double)anchor.Top;
            }
        }

        private static void DeleteVisual(JsonElement op)
        {
            ResolveShapeByName(op, "visualId").Delete();
        }

        private static void AddImageExcel(JsonElement op)
        {
            string path = op.GetProperty("path").GetString();
            if (path.StartsWith("http://") || path.StartsWith("https://"))
            {
                throw new NotSupportedException("add_image: remote URLs are not supported in this air-gapped deployment - use a local file path.");
            }
            string anchorCell = op.GetProperty("anchorCell").GetString();
            Excel.Range anchor = Sheet(op).Range[anchorCell];
            Sheet(op).Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                (float)(double)anchor.Left, (float)(double)anchor.Top, -1, -1);
        }

    }
}

