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

            // Fetched inside the ComRetry.Run lambda (not before it) so a flaky
            // open gets retried too, not just later calls. See WordTools.Charts.cs.md.
            dynamic dataWorkbook = null;
            try
            {
                // Build the whole grid in memory up front (pure C#, no COM) -
                // only the write itself needs to go through ComRetry.Run.
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

                ComRetry.Run(() =>
                {
                    // 120ms settle delay: the embedded chart-data workbook's automation
                    // surface isn't always live the instant ChartData.Workbook returns.
                    // See WordTools.Charts.cs.md.
                    System.Threading.Thread.Sleep(120);

                    DebugLog.Write("WriteChartData: getting chart.ChartData.Workbook");
                    dataWorkbook = chart.ChartData.Workbook;

                    DebugLog.Write("WriteChartData: getting Worksheets[1]");
                    dynamic sheet = dataWorkbook.Worksheets[1];

                    // A new chart's embedded workbook is pre-seeded with placeholder
                    // sample data; without clearing it, leftover cells beyond the new
                    // data's extent get plotted as phantom categories/series. See WordTools.Charts.cs.md.
                    DebugLog.Write("WriteChartData: Cells.Clear()");
                    sheet.Cells.Clear();

                    DebugLog.Write("WriteChartData: writing " + rowCount + "x" + colCount + " grid via Resize+Value2");
                    dynamic topLeft = sheet.Cells[1, 1];
                    dynamic writeRange = topLeft.Resize[rowCount, colCount];
                    writeRange.Value2 = grid;

                    // SetSourceData(String Source, Object PlotBy) takes a STRING, not a
                    // Range - pass a plain "SheetName!A1:B4" reference string. See WordTools.Charts.cs.md.
                    string a1Range = "A1:" + TextUtil.ColumnLetter(colCount) + rowCount;
                    string sourceRef = (string)sheet.Name + "!" + a1Range;
                    DebugLog.Write("WriteChartData: SetSourceData(\"" + sourceRef + "\")");
                    chart.SetSourceData(sourceRef);
                    DebugLog.Write("WriteChartData: SetSourceData returned OK");
                }, "WriteChartData");
            }
            finally
            {
                // Cleanup failures are caught and logged, not allowed to propagate -
                // an exception from `finally` while another is already propagating from
                // `try` REPLACES it, masking the real exception. See WordTools.Charts.cs.md.
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

        // Every chart shape (inline first, then floating) in that fixed order, so
        // chartIndex is predictable across calls. Shared by EditChart, ReadChart, and
        // TaskPaneHost.OnSelectionChanged so all three address charts identically.
        // internal (not private) for that last caller. See WordTools.Charts.cs.md.
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

        // Lets the model inspect a chart's current title/type/categories/series before
        // an incremental edit_chart call, since edit_chart replaces the whole dataset
        // rather than patching it. Reads via the embedded workbook's Cells/UsedRange/
        // .Value2 (same path WriteChartData writes through). See WordTools.Charts.cs.md.
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
            // Same fix as WriteChartData: ChartData.Workbook fetched inside the retry
            // lambda so a flaky open, not just a flaky subsequent call, gets retried.
            dynamic dataWorkbook = null;
            try
            {
                // Same settle-delay + retry protection as WriteChartData's write path.
                ComRetry.Run(() =>
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

                    // Range.Value2 returns a 1-based 2D array for a multi-cell range -
                    // read the actual bounds rather than assume 0 or 1.
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
                // A brand-new chart with no data would be blank/broken - seed a
                // minimal default series (matches the old hardcoded {1,2,3} fallback).
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

    }
}

