using System;
using System.Linq;
using System.Text.Json;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public struct ToolResult
    {
        public string Output;
        public bool IsError;
        public bool Mutated;
        public string Summary;
    }

    // Spike 3: real COM tool execution against the live Word document, called
    // from the WebView2-hosted AgentLoop via the JSON WebMessage bridge.
    public static class WordTools
    {
        public static ToolResult Execute(string name, JsonElement input)
        {
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
                    default:
                        return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        private static Word.Document ActiveDoc => Globals.ThisAddIn.Application.ActiveDocument;

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

        private static ToolResult InsertContent(JsonElement input)
        {
            string text = input.TryGetProperty("text", out var t) ? t.GetString() : "(no text provided)";
            Word.Document doc = ActiveDoc;
            Word.Range range = doc.Content;
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.Text = text;
            return new ToolResult
            {
                Output = "Inserted text at end of document: " + text,
                Mutated = true,
                Summary = "insert_content",
            };
        }

        // dynamic: Word's chart object model (Shapes.AddChart2 / Chart / SeriesCollection) mirrors
        // Excel/PowerPoint's shared chart engine; using dynamic avoids pinning down the exact
        // Interop type names for this spike and lets any signature mismatch surface immediately at
        // runtime instead of guessing overloads at compile time.
        private static ToolResult EditChart(JsonElement input)
        {
            string title = input.TryGetProperty("title", out var t) ? t.GetString() : "Chart";
            double[] values = input.TryGetProperty("values", out var v) && v.ValueKind == JsonValueKind.Array
                ? v.EnumerateArray().Select(x => x.GetDouble()).ToArray()
                : new double[] { 1, 2, 3 };

            dynamic doc = ActiveDoc;
            dynamic chartShape = null;
            foreach (dynamic shp in doc.Shapes)
            {
                if ((int)shp.HasChart == -1 /* msoTrue */)
                {
                    chartShape = shp;
                    break;
                }
            }
            bool created = false;
            if (chartShape == null)
            {
                // 51 = xlColumnClustered
                chartShape = doc.Shapes.AddChart2(-1, 51, 0, 0, 300, 200);
                created = true;
            }

            dynamic chart = chartShape.Chart;
            chart.HasTitle = true;
            chart.ChartTitle.Text = title;
            dynamic series = chart.SeriesCollection(1);
            series.Values = values;

            return new ToolResult
            {
                Output = $"Chart {(created ? "created" : "updated")}: title='{title}', values=[{string.Join(", ", values)}]",
                Mutated = true,
                Summary = "edit_chart",
            };
        }
    }
}
