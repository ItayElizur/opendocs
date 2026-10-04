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
        private static dynamic ResolveSmartArtLayout(string layoutKey)
        {
            string targetName;
            if (!SmartArtLayouts.ByName.TryGetValue(layoutKey, out targetName))
                throw new ArgumentException("add_smartart: unknown layout '" + layoutKey + "'. Valid: " +
                                            string.Join(", ", SmartArtLayouts.ByName.Keys) + ".");
            dynamic layouts = Globals.ThisAddIn.Application.SmartArtLayouts;
            foreach (dynamic layout in layouts)
            {
                if (string.Equals((string)layout.Name, targetName, StringComparison.OrdinalIgnoreCase))
                {
                    return layout;
                }
            }
            throw new InvalidOperationException("add_smartart: no SmartArt layout named '" + targetName +
                                                "' was found in this Office install's gallery - this install may be " +
                                                "non-English, where the built-in gallery's display names differ from " +
                                                "the standard English ones this tool assumes.");
        }

        // Unlike layouts, SmartArt color schemes/quick styles aren't a fixed enum in
        // this object model - they're live COM collections populated at runtime by
        // this install's gallery. Resolves by case-insensitive substring match against
        // the real names; a miss lists what's actually available. See WordTools.SmartArt.cs.md.
        private static dynamic ResolveSmartArtGalleryItem(dynamic collection, string query, string toolName, string whatKind)
        {
            dynamic firstMatch = null;
            var namesSeen = new List<string>();
            foreach (dynamic item in collection)
            {
                string name = (string)item.Name;
                namesSeen.Add(name);
                if (firstMatch == null && name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) firstMatch = item;
            }
            if (firstMatch != null) return firstMatch;
            string available = namesSeen.Count > 20
                ? string.Join(", ", namesSeen.GetRange(0, 20)) + ", ... (" + namesSeen.Count + " total)"
                : string.Join(", ", namesSeen);
            throw new ArgumentException(toolName + ": no " + whatKind + " matching '" + query + "' found in this Office install's gallery. Available: " + available + ".");
        }

        private static ToolResult AddSmartArt(JsonElement input)
        {
            string layoutKey = input.GetProperty("layout").GetString();
            dynamic layout = ResolveSmartArtLayout(layoutKey);

            int? afterBlockIndex = input.TryGetProperty("afterBlockIndex", out var abEl) && abEl.ValueKind == JsonValueKind.Number
                ? abEl.GetInt32() : (int?)null;
            float width = input.TryGetProperty("w", out var w) ? (float)w.GetDouble() : 400f;
            float height = input.TryGetProperty("h", out var h) ? (float)h.GetDouble() : 300f;

            dynamic doc = ActiveDoc;
            dynamic shape;
            if (afterBlockIndex.HasValue)
            {
                // Mirrors the chart's anchored-creation path; whether AddSmartArt truly
                // accepts a named Anchor parameter in this PIA is unverified.
                Word.Range at = RangeAfterBlock(afterBlockIndex.Value);
                dynamic floatingAtAnchor = doc.Shapes.AddSmartArt(layout, 0, 0, width, height, Anchor: at);
                shape = floatingAtAnchor.ConvertToInlineShape();
            }
            else
            {
                float left = input.TryGetProperty("x", out var x) ? (float)x.GetDouble() : 100f;
                float top = input.TryGetProperty("y", out var y) ? (float)y.GetDouble() : 100f;
                shape = doc.Shapes.AddSmartArt(layout, left, top, width, height);
            }

            dynamic smartArt = shape.SmartArt;

            // AddSmartArt seeds the diagram with the layout's default placeholder
            // "[Text]" nodes; delete them first or requested items get appended after
            // them instead of replacing them (same idea as the chart fix's
            // Cells.Clear()). See WordTools.SmartArt.cs.md.
            dynamic existingNodes = smartArt.Nodes;
            for (int i = (int)existingNodes.Count; i >= 1; i--)
            {
                existingNodes.Item(i).Delete();
            }

            foreach (JsonElement item in input.GetProperty("items").EnumerateArray())
            {
                dynamic node = smartArt.Nodes.Add();
                node.TextFrame2.TextRange.Text = item.GetString();
            }
            return new ToolResult { Output = "SmartArt added (" + input.GetProperty("items").GetArrayLength() + " node(s)).", Mutated = true, Summary = "add_smartart" };
        }

        // Mirrors ListChartShapes for shape.HasSmartArt. HasSmartArt returns an
        // MsoTriState, not a bool - compared via (int)x == -1, same fix as HasChart.
        // See WordTools.SmartArt.cs.md.
        internal static List<dynamic> ListSmartArtShapes(dynamic doc)
        {
            var shapes = new List<dynamic>();
            foreach (dynamic shp in doc.InlineShapes)
            {
                try { if ((int)shp.HasSmartArt == -1 /* msoTrue */) shapes.Add(shp); } catch { }
            }
            foreach (dynamic shp in doc.Shapes)
            {
                try { if ((int)shp.HasSmartArt == -1 /* msoTrue */) shapes.Add(shp); } catch { }
            }
            return shapes;
        }

        // Extracted so ReadSmartArt can read every diagram in one call when
        // smartArtIndex is omitted, instead of one call per diagram.
        private static string ReadOneSmartArt(dynamic shape, int index, int total)
        {
            dynamic smartArt = shape.SmartArt;
            dynamic nodes = smartArt.Nodes;
            int count = (int)nodes.Count;
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("SmartArt " + index + " of " + total + " (" + count + " node(s)):");
            for (int i = 1; i <= count; i++)
            {
                dynamic node = nodes.Item(i);
                string text = "";
                try { text = (string)node.TextFrame2.TextRange.Text; } catch { }
                sb.AppendLine("[" + (i - 1) + "] " + text);
            }
            return sb.ToString().TrimEnd();
        }

        private static ToolResult ReadSmartArt(JsonElement input)
        {
            dynamic doc = ActiveDoc;
            var shapes = ListSmartArtShapes(doc);
            if (shapes.Count == 0)
                return new ToolResult { Output = "No SmartArt diagrams in this document.", Summary = "read_smartart" };

            bool hasIndex = input.TryGetProperty("smartArtIndex", out var si) && si.ValueKind == JsonValueKind.Number;
            if (!hasIndex)
            {
                // No index given: read every diagram in one call rather than
                // forcing one call per diagram.
                var all = new List<string>();
                for (int i = 0; i < shapes.Count; i++) all.Add(ReadOneSmartArt(shapes[i], i, shapes.Count));
                return new ToolResult { Output = string.Join("\n\n", all), Summary = "read_smartart" };
            }

            int index = si.GetInt32();
            if (index < 0 || index >= shapes.Count)
                throw new ArgumentOutOfRangeException("smartArtIndex", "smartArtIndex must be between 0 and " + (shapes.Count - 1) + " (" + shapes.Count + " diagram(s) in the document).");
            return new ToolResult { Output = ReadOneSmartArt(shapes[index], index, shapes.Count), Summary = "read_smartart" };
        }

        private static ToolResult EditSmartArt(JsonElement input)
        {
            dynamic doc = ActiveDoc;
            var shapes = ListSmartArtShapes(doc);
            if (shapes.Count == 0)
                throw new InvalidOperationException("edit_smartart: no SmartArt diagrams in this document.");

            int index = input.TryGetProperty("smartArtIndex", out var si) && si.ValueKind == JsonValueKind.Number ? si.GetInt32() : 0;
            if (index < 0 || index >= shapes.Count)
                throw new ArgumentOutOfRangeException("smartArtIndex", "smartArtIndex must be between 0 and " + (shapes.Count - 1) + " (" + shapes.Count + " diagram(s) in the document).");

            dynamic smartArt = shapes[index].SmartArt;
            dynamic nodes = smartArt.Nodes;
            string kind = input.GetProperty("kind").GetString();
            switch (kind)
            {
                case "set_text":
                {
                    int nodeIndex = input.GetProperty("nodeIndex").GetInt32();
                    int count = (int)nodes.Count;
                    if (nodeIndex < 0 || nodeIndex >= count)
                        throw new ArgumentOutOfRangeException("nodeIndex", "nodeIndex must be between 0 and " + (count - 1) + " (" + count + " node(s)).");
                    nodes.Item(nodeIndex + 1).TextFrame2.TextRange.Text = input.GetProperty("text").GetString();
                    return new ToolResult { Output = "Node " + nodeIndex + " updated.", Mutated = true, Summary = "edit_smartart" };
                }
                case "add_node":
                {
                    dynamic newNode = nodes.Add();
                    if (input.TryGetProperty("text", out var textEl) && textEl.ValueKind == JsonValueKind.String)
                        newNode.TextFrame2.TextRange.Text = textEl.GetString();
                    return new ToolResult { Output = "Node added at index " + ((int)nodes.Count - 1) + ".", Mutated = true, Summary = "edit_smartart" };
                }
                case "delete_node":
                {
                    int nodeIndex = input.GetProperty("nodeIndex").GetInt32();
                    int count = (int)nodes.Count;
                    if (nodeIndex < 0 || nodeIndex >= count)
                        throw new ArgumentOutOfRangeException("nodeIndex", "nodeIndex must be between 0 and " + (count - 1) + " (" + count + " node(s)).");
                    nodes.Item(nodeIndex + 1).Delete();
                    return new ToolResult { Output = "Node " + nodeIndex + " deleted. Later node indices have shifted - re-read (read_smartart) before another node edit in the same run.", Mutated = true, Summary = "edit_smartart" };
                }
                case "set_style":
                {
                    bool changed = false;
                    if (input.TryGetProperty("colorName", out var cnEl) && cnEl.ValueKind == JsonValueKind.String)
                    {
                        dynamic colors = Globals.ThisAddIn.Application.SmartArtColors;
                        smartArt.Color = ResolveSmartArtGalleryItem(colors, cnEl.GetString(), "edit_smartart", "color scheme");
                        changed = true;
                    }
                    if (input.TryGetProperty("quickStyleName", out var qsEl) && qsEl.ValueKind == JsonValueKind.String)
                    {
                        dynamic quickStyles = Globals.ThisAddIn.Application.SmartArtQuickStyles;
                        smartArt.QuickStyle = ResolveSmartArtGalleryItem(quickStyles, qsEl.GetString(), "edit_smartart", "quick style");
                        changed = true;
                    }
                    if (!changed)
                        throw new ArgumentException("edit_smartart: set_style requires at least one of colorName or quickStyleName.");
                    return new ToolResult { Output = "SmartArt style updated.", Mutated = true, Summary = "edit_smartart" };
                }
                case "set_layout":
                {
                    // Reuses ResolveSmartArtLayout - SmartArt.Layout is settable
                    // (confirmed via reflection), same resolve-then-assign shape as
                    // creating one. See WordTools.SmartArt.cs.md.
                    string layoutKey = input.GetProperty("layout").GetString();
                    smartArt.Layout = ResolveSmartArtLayout(layoutKey);
                    return new ToolResult { Output = "SmartArt layout changed to '" + layoutKey + "'.", Mutated = true, Summary = "edit_smartart" };
                }
                default:
                    throw new ArgumentException("edit_smartart: unknown kind '" + kind + "'. Valid: set_text, add_node, delete_node, set_style, set_layout.");
            }
        }

    }
}

