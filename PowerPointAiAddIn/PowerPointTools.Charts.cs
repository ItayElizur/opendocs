using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        private static ToolResult AddChartPpt(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            string kindStr = input.GetProperty("kind").GetString();
            int typeCode;
            if (!ChartTypes.ByName.TryGetValue(kindStr, out typeCode))
                throw new ArgumentException("add_chart: unknown kind '" + kindStr + "'. Valid: " +
                                            string.Join(", ", ChartTypes.ByName.Keys) + ".");
            var categories = new List<string>();
            foreach (JsonElement c in input.GetProperty("categories").EnumerateArray()) categories.Add(c.GetString());

            // PP-22: a series/categories length mismatch would otherwise write a
            // ragged grid and produce a silently wrong chart - reject it up front.
            JsonElement seriesForValidation = input.GetProperty("series");
            foreach (JsonElement s in seriesForValidation.EnumerateArray())
            {
                int valueCount = 0;
                foreach (JsonElement v in s.GetProperty("values").EnumerateArray()) valueCount++;
                if (valueCount != categories.Count)
                {
                    string seriesName = s.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String ? nameEl.GetString() : "(unnamed)";
                    throw new ArgumentException("add_chart: series '" + seriesName + "' has " + valueCount +
                                                " value(s) but there are " + categories.Count +
                                                " categories - every series must have exactly one value per category.");
                }
            }

            float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
            float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
            float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
            float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 300f;

            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            dynamic chartShape = slide.Shapes.AddChart2(-1, (Microsoft.Office.Core.XlChartType)typeCode, left, top, width, height);
            // 0-based index into slide.Shapes - AddChart2 appends the new shape at
            // the end of the collection, so this is stable immediately after the
            // call. Returned below so the model can edit_chart it without a
            // separate read_slide round trip (PP-22 Task 1 Step 5).
            int newShapeIndex = slide.Shapes.Count - 1;
            dynamic chart = chartShape.Chart;

            // Chart data lives in an embedded Excel workbook - open, write the grid,
            // close, and RELEASE explicitly so no hidden Excel host process leaks.
            // Post-hoc fix (2026-08-24, ported from Word's identical fix,
            // found via a real repro's DebugLog): moved inside the retry
            // lambda below so a flaky OPEN, not just a flaky subsequent
            // call, also gets retried.
            dynamic dataWorkbook = null;
            try
            {
                var seriesList = input.GetProperty("series").EnumerateArray().ToList();

                // Build the whole grid in memory up front (pure C#, no COM) -
                // only the write itself needs to go through ComRetry.Run.
                int rowCount = categories.Count + 1; // +1 header row
                int colCount = seriesList.Count + 1; // +1 category column
                object[,] grid = new object[rowCount, colCount];
                grid[0, 0] = "";
                int colIdx = 0;
                foreach (JsonElement s in seriesList)
                {
                    grid[0, colIdx + 1] = s.GetProperty("name").GetString();
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
                    // Post-hoc fix (2026-08-24, user-reported the RPC failure
                    // recurring even after the first fix - ported from
                    // Word's identical fix): a brief settle delay right after
                    // the embedded OLE workbook opens, before the first COM
                    // call against it - the automation surface is not always
                    // fully live the instant ChartData.Workbook returns.
                    System.Threading.Thread.Sleep(120);

                    dataWorkbook = chart.ChartData.Workbook;
                    dynamic sheet = dataWorkbook.Worksheets[1];

                    // Confirmed repro (Word's identical port of this same pattern,
                    // PP-9): a brand-new chart's embedded workbook comes pre-seeded
                    // by Office with placeholder sample data (a default chart
                    // template, commonly 4 categories x 3 series). Without
                    // clearing it first, only the cells the NEW data actually
                    // occupies get overwritten - any leftover placeholder cells
                    // stay in the sheet and get plotted alongside the real data.
                    sheet.Cells.Clear();

                    dynamic topLeft = sheet.Cells[1, 1];
                    dynamic writeRange = topLeft.Resize[rowCount, colCount];
                    writeRange.Value2 = grid;

                    // ACTUAL ROOT CAUSE (2026-08-24, confirmed via .NET
                    // reflection against the real referenced
                    // Microsoft.Office.Interop.PowerPoint.dll, not a guess -
                    // same finding as Word's identical code):
                    // PowerPoint.Chart.SetSourceData's real signature is
                    // SetSourceData(String Source, Object PlotBy) - the
                    // first parameter is a STRING, not a Range. Every prior
                    // attempt (a reused writeRange, then a Range built from
                    // an A1 string) was passing a Range COM object where the
                    // method actually expects a plain "SheetName!A1:B4"
                    // reference string - no Range object needed at all.
                    string a1Range = "A1:" + TextUtil.ColumnLetter(colCount) + rowCount;
                    string sourceRef = (string)sheet.Name + "!" + a1Range;
                    chart.SetSourceData(sourceRef);
                });
            }
            finally
            {
                // ROOT CAUSE FOUND (2026-08-24): this cleanup had no catch of
                // its own - when SetSourceData failed above, the chart/
                // embedded-workbook was left in a state where
                // dataWorkbook.Close() ALSO threw a second, unrelated
                // exception, which (per C# finally semantics) REPLACED the
                // real SetSourceData exception before it ever reached the
                // caller. Cleanup failures are now caught and swallowed
                // (not re-thrown) so they can never mask a real exception
                // already in flight.
                if (dataWorkbook != null)
                {
                    try
                    {
                        dataWorkbook.Close(SaveChanges: true);
                    }
                    catch { /* secondary failure - do not mask the real exception */ }
                    finally
                    {
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(dataWorkbook); } catch { }
                    }
                }
            }

            if (input.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
            }
            string named = ApplyOptionalName(slide.Shapes[newShapeIndex + 1], input);
            return new ToolResult { Output = "Chart added at shapeIndex " + newShapeIndex + (named != null ? " (\"" + named + "\")" : "") + ".", Mutated = true, Summary = "add_chart" };
        }

        // PP-21: legendPos's natural names (a model will say "right", not the
        // genoffice-ism "r") plus the original short aliases for back-compat.
        // xlLegendPositionCorner was considered and dropped - its code could not
        // be verified against a live Office install (no interactive GUI access
        // in this environment), and guessing it wrong would just replace one
        // silent-wrong-result bug with another.
        private static readonly Dictionary<string, int> PptLegendPositions =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["right"] = -4152, ["r"] = -4152,   // xlLegendPositionRight
            ["top"] = -4160, ["t"] = -4160,     // xlLegendPositionTop
            ["left"] = -4131, ["l"] = -4131,    // xlLegendPositionLeft
            ["bottom"] = -4107, ["b"] = -4107,  // xlLegendPositionBottom
        };

        private static ToolResult EditChartPpt(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            if (shape.HasChart != Microsoft.Office.Core.MsoTriState.msoTrue)
                throw new InvalidOperationException("edit_chart: shape " + input.GetProperty("shapeIndex").GetInt32() +
                                                    " on slide " + input.GetProperty("slideIndex").GetInt32() + " is not a chart.");
            dynamic chart = shape.Chart;
            var applied = new List<string>();

            if (input.TryGetProperty("chartType", out var ct) && ct.ValueKind == JsonValueKind.String)
            {
                int typeCode;
                if (!ChartTypes.ByName.TryGetValue(ct.GetString(), out typeCode))
                    throw new ArgumentException("edit_chart: unknown chartType '" + ct.GetString() + "'. Valid: " +
                                                string.Join(", ", ChartTypes.ByName.Keys) + ".");
                chart.ChartType = typeCode;
                applied.Add("chartType=" + ct.GetString());
            }
            if (input.TryGetProperty("title", out var title) && title.ValueKind == JsonValueKind.String)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title.GetString();
                applied.Add("title");
            }
            if (input.TryGetProperty("legendPos", out var legendPos) && legendPos.ValueKind == JsonValueKind.String)
            {
                string pos = legendPos.GetString();
                if (string.Equals(pos, "none", StringComparison.OrdinalIgnoreCase))
                {
                    chart.HasLegend = false;
                }
                else
                {
                    int posCode;
                    if (!PptLegendPositions.TryGetValue(pos, out posCode))
                        throw new ArgumentException("edit_chart: unknown legendPos '" + pos + "'. Valid: none, " +
                                                    string.Join(", ", PptLegendPositions.Keys) + ".");
                    chart.HasLegend = true;
                    chart.Legend.Position = posCode;
                }
                applied.Add("legendPos=" + pos);
            }
            if (input.TryGetProperty("dataLabels", out var dl) && dl.ValueKind == JsonValueKind.String)
            {
                // Matches Excel's edit_chart vocabulary (none|value|percent) -
                // previously a boolean, which silently turned labels OFF for any
                // non-true value, including the string "value".
                string mode = dl.GetString();
                if (mode != "none" && mode != "value" && mode != "percent")
                    throw new ArgumentException("edit_chart: unknown dataLabels '" + mode + "'. Valid: none, value, percent.");
                bool show = mode != "none";
                foreach (dynamic series in chart.SeriesCollection())
                {
                    series.HasDataLabels = show;
                    if (show && mode == "percent") series.DataLabels().ShowPercentage = true;
                }
                applied.Add("dataLabels=" + mode);
            }
            if (input.TryGetProperty("gridlines", out var gl) && (gl.ValueKind == JsonValueKind.True || gl.ValueKind == JsonValueKind.False))
            {
                bool show = gl.ValueKind == JsonValueKind.True;
                try
                {
                    chart.Axes(2 /* xlValue */).HasMajorGridlines = show;
                    applied.Add("gridlines=" + show);
                }
                catch (Exception)
                {
                    throw new InvalidOperationException(
                        "edit_chart: this chart's type has no value axis (e.g. pie/doughnut) - gridlines do not apply to it.");
                }
            }
            return new ToolResult
            {
                Output = applied.Count > 0
                    ? "Chart updated: " + string.Join(", ", applied) + "."
                    : "No recognized chart properties were provided - nothing changed.",
                Mutated = applied.Count > 0,
                Summary = "edit_chart",
            };
        }

        // Verified against the standard English display names for PowerPoint's built-in SmartArt
        // layout gallery. Live cross-check against Application.SmartArtLayouts on this machine's
        // Office install (plan Task 6 Step 1) requires interactive Office GUI access that was not
        // available in this environment - remains a manual follow-up for a human with GUI access.

    }
}

