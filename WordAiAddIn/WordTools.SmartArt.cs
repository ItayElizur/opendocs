using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    // Spike 3: real COM tool execution against the live Word document, called
    // from the WebView2-hosted AgentLoop via the JSON WebMessage bridge.
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

        // Post-hoc addition (2026-08-24, user-reported: "smart art has no
        // change style/color for an existing element"). Unlike layouts,
        // SmartArt color schemes and quick styles are NOT a fixed enum in
        // this Office object model (confirmed via reflection against the
        // referenced Office 15 PIA - Microsoft.Office.Core has no
        // MsoSmartArtColorType/MsoSmartArtQuickStyleType at all) - they are
        // live COM collections (Application.SmartArtColors /
        // .SmartArtQuickStyles) of SmartArtColor/SmartArtQuickStyle objects,
        // each with a .Name populated at runtime by this install's own
        // gallery. Rather than guess a curated list of exact display-name
        // strings (unverifiable without a live session, and a wrong guess
        // would either fail or - worse - silently match nothing), this
        // resolves by case-insensitive SUBSTRING match against whatever
        // names this install actually has, and a miss lists the real
        // available names so the caller can retry correctly instead of
        // guessing blind a second time.
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
                // Mirrors PP-9's anchored-chart-creation path exactly, including
                // its caveat: whether Shapes.AddSmartArt truly accepts a named
                // Anchor parameter in this PIA is UNVERIFIED - flagged as
                // elevated risk in the plan/verification file.
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

            // Post-hoc fix (2026-08-24, user-reported): AddSmartArt seeds the
            // new diagram with the layout's own default placeholder nodes
            // (the same "[Text]" prompts the ribbon's SmartArt gallery shows) -
            // same bug shape as the chart-data fix above (pre-seeded content
            // never cleared before writing). Without clearing them first, the
            // requested items were APPENDED after the placeholders instead of
            // replacing them, leaving visible "[Text]" nodes above the real
            // ones. Delete every existing node before adding the real ones,
            // same idea as the chart fix's sheet.Cells.Clear().
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

        // PP-23 Task 5: SmartArt shapes are not chart shapes and are not
        // tables - a small, separate list-and-resolve helper, mirroring
        // ListChartShapes'/ResolveTable's shape but for shape.HasSmartArt
        // instead of shape.HasChart.
        //
        // Post-hoc fix (2026-08-24, user-reported): HasSmartArt returns an
        // MsoTriState, not a real bool, exactly like HasChart elsewhere in
        // this file - a plain (bool) cast either throws (silently swallowed
        // by the try/catch below) or never matches, so no shape was ever
        // recognized as SmartArt and read_smartart/edit_smartart always
        // reported "no SmartArt diagrams" even right after add_smartart had
        // just created one. Fixed with the same (int)x == -1 comparison
        // ListChartShapes already uses for HasChart.
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

        // Post-hoc fix (2026-08-24, user-reported): reading N diagrams
        // previously needed N separate read_smartart calls (one per index) -
        // extracted so ReadSmartArt can read every diagram in one call when
        // smartArtIndex is omitted, matching what the user actually wanted
        // ("a way to read the entire smartart text at once").
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
                    // Post-hoc addition (2026-08-24, user-reported: "smart art
                    // cant change layout"). Reuses ResolveSmartArtLayout
                    // verbatim (same curated 7-key map + gallery lookup
                    // add_smartart already uses) - SmartArt.Layout is
                    // settable (confirmed via the same reflection pass that
                    // found .Color/.QuickStyle), so changing an EXISTING
                    // diagram's layout is the same resolve-then-assign shape
                    // as creating one, just against smartArt.Layout instead
                    // of the AddSmartArt call.
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

