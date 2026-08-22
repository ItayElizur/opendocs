using System;
using System.Linq;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{

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
                    case "read_blocks":
                        return ReadBlocks(input);
                    case "replace_blocks":
                        return ReplaceBlocks(input);
                    case "apply_commands":
                        return ApplyCommands(input);
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

        private static ToolResult ReadBlocks(JsonElement input)
        {
            int startIndex = input.GetProperty("startIndex").GetInt32();
            int endIndex = input.GetProperty("endIndex").GetInt32();
            Word.Document doc = ActiveDoc;
            Word.Paragraphs paragraphs = doc.Paragraphs;
            int count = paragraphs.Count;
            endIndex = Math.Min(endIndex, count - 1);
            if (startIndex < 0 || startIndex > endIndex)
            {
                return new ToolResult { Output = "Invalid range.", IsError = true, Summary = "read_blocks" };
            }
            var sb = new System.Text.StringBuilder();
            for (int i = startIndex; i <= endIndex; i++)
            {
                Word.Paragraph p = paragraphs[i + 1];
                string text = p.Range.Text.TrimEnd('\r', '\a', '\n');
                sb.AppendLine($"[{i}] {text}");
            }
            return new ToolResult { Output = sb.ToString(), Summary = "read_blocks" };
        }

        private static ToolResult ReplaceBlocks(JsonElement input)
        {
            int startIndex = input.GetProperty("startIndex").GetInt32();
            int endIndex = input.GetProperty("endIndex").GetInt32();
            string text = input.GetProperty("text").GetString() ?? "";
            Word.Document doc = ActiveDoc;
            Word.Paragraphs paragraphs = doc.Paragraphs;
            endIndex = Math.Min(endIndex, paragraphs.Count - 1);
            if (startIndex < 0 || startIndex > endIndex)
            {
                return new ToolResult { Output = "Invalid range.", IsError = true, Summary = "replace_blocks" };
            }
            Word.Range range = doc.Range(paragraphs[startIndex + 1].Range.Start, paragraphs[endIndex + 1].Range.End);
            range.Text = text;
            return new ToolResult { Output = $"Replaced paragraphs {startIndex}-{endIndex} with: {text}", Mutated = true, Summary = "replace_blocks" };
        }

        private static ToolResult ApplyCommands(JsonElement input)
        {
            var lines = new System.Text.StringBuilder();
            bool anyMutated = false;
            bool anyError = false;
            foreach (JsonElement cmd in input.GetProperty("commands").EnumerateArray())
            {
                string kind = cmd.GetProperty("kind").GetString();
                try
                {
                    switch (kind)
                    {
                        case "set_bold":
                            SetRunProperty(cmd, (range, value) => range.Bold = value ? 1 : 0);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_italic":
                            SetRunProperty(cmd, (range, value) => range.Italic = value ? 1 : 0);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "set_heading":
                            SetHeading(cmd);
                            lines.AppendLine(kind + ": ok"); anyMutated = true; break;
                        case "find_replace":
                            int replacements = FindReplace(cmd);
                            lines.AppendLine($"{kind}: {replacements} replacement(s)");
                            if (replacements > 0) anyMutated = true;
                            break;
                        default:
                            lines.AppendLine(kind + ": unknown command kind"); anyError = true; break;
                    }
                }
                catch (Exception ex)
                {
                    lines.AppendLine(kind + ": ERROR - " + ex.Message); anyError = true;
                }
            }
            return new ToolResult { Output = lines.ToString(), Mutated = anyMutated, IsError = anyError, Summary = "apply_commands" };
        }

        private static void SetRunProperty(JsonElement cmd, Action<Word.Range, bool> apply)
        {
            int startIndex = cmd.GetProperty("startIndex").GetInt32();
            int endIndex = cmd.GetProperty("endIndex").GetInt32();
            bool value = cmd.GetProperty("value").GetBoolean();
            Word.Document doc = ActiveDoc;
            Word.Paragraphs paragraphs = doc.Paragraphs;
            endIndex = Math.Min(endIndex, paragraphs.Count - 1);
            Word.Range range = doc.Range(paragraphs[startIndex + 1].Range.Start, paragraphs[endIndex + 1].Range.End);
            apply(range, value);
        }

        private static void SetHeading(JsonElement cmd)
        {
            int index = cmd.GetProperty("index").GetInt32();
            int level = cmd.GetProperty("level").GetInt32();
            Word.Paragraph p = ActiveDoc.Paragraphs[index + 1];
            p.Range.set_Style(level == 0 ? "Normal" : "Heading " + level);
        }

        private static int FindReplace(JsonElement cmd)
        {
            string find = cmd.GetProperty("find").GetString();
            string replace = cmd.GetProperty("replace").GetString();
            bool matchCase = cmd.TryGetProperty("matchCase", out var mc) && mc.GetBoolean();
            Word.Find findObj = ActiveDoc.Content.Find;
            findObj.ClearFormatting();
            findObj.Text = find;
            findObj.Replacement.ClearFormatting();
            findObj.Replacement.Text = replace;
            findObj.MatchCase = matchCase;
            bool found = findObj.Execute(Replace: Word.WdReplace.wdReplaceAll);
            return found ? 1 : 0;
        }
    }
}
