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
        // PP-22 Task 2 Step 4: index-based lookup (SmartArtLayouts is
        // index-addressable, and the built-in gallery order is stable across
        // installs) was considered as a locale-independent alternative to
        // name-matching. Not switched - index stability across Office versions
        // is an assumption no better founded than the display-name assumption,
        // and names at least fail loudly with a diagnostic message below,
        // whereas a wrong index would silently insert the wrong diagram.
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
            // Distinct from the unknown-key case above: the key was valid, but
            // this Office install's gallery has no layout under that display
            // name - the one diagnosis that actually points at a non-English
            // install, which a silent fallback could never surface.
            throw new InvalidOperationException("add_smartart: no SmartArt layout named '" + targetName +
                                                "' was found in this Office install's gallery - this install may be " +
                                                "non-English, where the built-in gallery's display names differ from " +
                                                "the standard English ones this tool assumes (see plan Task 6 Step 1).");
        }

        // Ported from WordTools.cs (2026-08-27) to bring PowerPoint's SmartArt
        // surface to parity with Word's, which already had edit/read.
        //
        // The one real difference from Word's version: Word finds SmartArt in a
        // flat document (doc.InlineShapes + doc.Shapes), while PowerPoint's
        // shapes live per-slide. So a diagram is addressed by (slideIndex,
        // smartArtIndex-within-that-slide) rather than a single document-wide
        // index - slide-scoped indices match how every other PowerPoint tool
        // here addresses shapes.
        internal static List<dynamic> ListSmartArtShapesOnSlide(PowerPoint.Slide slide)
        {
            var shapes = new List<dynamic>();
            foreach (dynamic shp in slide.Shapes)
            {
                // HasSmartArt is an MsoTriState (-1 == msoTrue), not a bool -
                // the same trap PP-23 hit on the Word side. The try/catch
                // covers shape types that do not expose the property at all.
                try { if ((int)shp.HasSmartArt == -1) shapes.Add(shp); } catch { }
            }
            return shapes;
        }

        // Resolves (slideIndex, smartArtIndex) to a diagram's SmartArt object,
        // with errors that say how many are actually present.
        private static dynamic ResolveSmartArtOnSlide(JsonElement input, string toolName)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (slideIndex < 0 || slideIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("slideIndex",
                    "slideIndex must be between 0 and " + (slides.Count - 1) + ".");

            var shapes = ListSmartArtShapesOnSlide(slides[slideIndex + 1]);
            if (shapes.Count == 0)
                throw new InvalidOperationException(toolName + ": no SmartArt diagrams on slide " + slideIndex + ".");

            int index = input.TryGetProperty("smartArtIndex", out var si) && si.ValueKind == JsonValueKind.Number
                ? si.GetInt32() : 0;
            if (index < 0 || index >= shapes.Count)
                throw new ArgumentOutOfRangeException("smartArtIndex",
                    "smartArtIndex must be between 0 and " + (shapes.Count - 1) +
                    " (" + shapes.Count + " diagram(s) on slide " + slideIndex + ").");
            return shapes[index].SmartArt;
        }

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
            throw new ArgumentException(toolName + ": no " + whatKind + " matching '" + query +
                                        "' found in this Office install's gallery. Available: " + available + ".");
        }

        private static ToolResult ReadSmartArt(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (slideIndex < 0 || slideIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("slideIndex",
                    "slideIndex must be between 0 and " + (slides.Count - 1) + ".");

            var shapes = ListSmartArtShapesOnSlide(slides[slideIndex + 1]);
            if (shapes.Count == 0)
                return new ToolResult { Output = "No SmartArt diagrams on slide " + slideIndex + ".", Summary = "read_smartart" };

            // No smartArtIndex given: read every diagram on the slide in one
            // call rather than forcing a round trip per diagram (matches
            // Word's read_smartart).
            bool hasIndex = input.TryGetProperty("smartArtIndex", out var si) && si.ValueKind == JsonValueKind.Number;
            if (!hasIndex)
            {
                var all = new List<string>();
                for (int i = 0; i < shapes.Count; i++) all.Add(ReadOneSmartArt(shapes[i], i, shapes.Count));
                return new ToolResult { Output = string.Join("\n\n", all), Summary = "read_smartart" };
            }

            int index = si.GetInt32();
            if (index < 0 || index >= shapes.Count)
                throw new ArgumentOutOfRangeException("smartArtIndex",
                    "smartArtIndex must be between 0 and " + (shapes.Count - 1) +
                    " (" + shapes.Count + " diagram(s) on slide " + slideIndex + ").");
            return new ToolResult { Output = ReadOneSmartArt(shapes[index], index, shapes.Count), Summary = "read_smartart" };
        }

        private static string ReadOneSmartArt(dynamic shape, int index, int total)
        {
            dynamic smartArt = shape.SmartArt;
            dynamic nodes = smartArt.Nodes;
            int count = (int)nodes.Count;
            var sb = new StringBuilder();
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

        private static ToolResult EditSmartArt(JsonElement input)
        {
            dynamic smartArt = ResolveSmartArtOnSlide(input, "edit_smartart");
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
                    string layoutKey = input.GetProperty("layout").GetString();
                    smartArt.Layout = ResolveSmartArtLayout(layoutKey);
                    return new ToolResult { Output = "SmartArt layout changed to '" + layoutKey + "'.", Mutated = true, Summary = "edit_smartart" };
                }
                default:
                    throw new ArgumentException("edit_smartart: unknown kind '" + kind + "'. Valid: set_text, add_node, delete_node, set_style, set_layout.");
            }
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

            // Post-hoc fix (2026-08-24, found via Word's identical port of
            // this code, PP-23): AddSmartArt seeds the diagram with the
            // layout's own default "[Text]" placeholder nodes. Without
            // clearing them first, the requested items were appended after
            // the placeholders instead of replacing them.
            dynamic existingNodes = smartArt.Nodes;
            for (int i = (int)existingNodes.Count; i >= 1; i--)
            {
                existingNodes.Item(i).Delete();
            }

            // genoffice's own version only ever produces a flat item list - maps
            // 1:1 to sequential top-level nodes, no nested tree-building needed.
            foreach (JsonElement item in input.GetProperty("items").EnumerateArray())
            {
                dynamic node = smartArt.Nodes.Add();
                node.TextFrame2.TextRange.Text = item.GetString();
            }
            return new ToolResult { Output = "SmartArt added.", Mutated = true, Summary = "add_smartart" };
        }

    }
}

