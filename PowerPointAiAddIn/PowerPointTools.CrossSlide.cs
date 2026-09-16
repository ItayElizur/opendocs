using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // copy_element/move_element: cross-slide relocation, via property-based
        // reconstruction on the destination slide - NEVER the clipboard
        // (Shape.Copy()/Shapes.Paste() would work for every shape kind, but
        // touches the real OS clipboard, the same hazard already avoided for
        // Word). PowerPoint's object model has no clipboard-free way to
        // reparent a Shape to a different slide's Shapes collection, so this
        // instead reads the source shape's properties and calls the matching
        // add_text_box/add_shape/add_table/add_chart/add_smartart tool on the
        // destination slide - only for the 5 kinds those tools can build from
        // scratch. Groups, pictures, OLE objects, and media have no such
        // primitive and are refused with a specific named error rather than a
        // lossy approximation. Same-slide duplication (duplicate_element,
        // PowerPointTools.Elements.cs) is unrestricted, since native
        // Shape.Duplicate() needs no reconstruction.

        private enum ReconstructionKind { TextBox, AutoShape, Table, Chart, SmartArt }

        // Built once from OfficeAi.Shared's existing maps - kept local to this
        // app rather than added to OfficeAi.Shared, so this feature touches no
        // shared file.
        private static readonly Dictionary<int, string> AutoShapeCodeToName = BuildAutoShapeCodeToName();
        private static readonly Dictionary<int, string> ChartTypeCodeToName = ChartTypes.ByName.ToDictionary(kv => kv.Value, kv => kv.Key);
        private static readonly Dictionary<string, string> SmartArtNameToKey =
            SmartArtLayouts.ByName.ToDictionary(kv => kv.Value, kv => kv.Key, StringComparer.OrdinalIgnoreCase);

        private static Dictionary<int, string> BuildAutoShapeCodeToName()
        {
            // ShapeTypes.ByName is not 1:1 ("rectangle"/"oval" alias "rect"/
            // "ellipse") - keep the first-seen name per code, which is the
            // canonical non-alias name since aliases are declared after their
            // canonical entry.
            var map = new Dictionary<int, string>();
            foreach (var kv in ShapeTypes.ByName)
            {
                if (!map.ContainsKey(kv.Value)) map[kv.Value] = kv.Key;
            }
            return map;
        }

        private static JsonElement BuildJson(Dictionary<string, object> data)
        {
            byte[] bytes = JsonSerializer.SerializeToUtf8Bytes(data);
            using (JsonDocument doc = JsonDocument.Parse(bytes))
            {
                return doc.RootElement.Clone();
            }
        }

        private static ReconstructionKind ResolveReconstructionKind(PowerPoint.Shape shape, string toolName)
        {
            if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue) return ReconstructionKind.Table;
            if (shape.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue) return ReconstructionKind.Chart;

            dynamic dshape = shape;
            bool hasSmartArt = false;
            try { hasSmartArt = (int)dshape.HasSmartArt == -1; } catch { }
            if (hasSmartArt) return ReconstructionKind.SmartArt;

            Microsoft.Office.Core.MsoShapeType type = shape.Type;
            bool noPrimitive =
                type == Microsoft.Office.Core.MsoShapeType.msoGroup ||
                type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture ||
                type == Microsoft.Office.Core.MsoShapeType.msoEmbeddedOLEObject ||
                type == Microsoft.Office.Core.MsoShapeType.msoLinkedOLEObject ||
                type == Microsoft.Office.Core.MsoShapeType.msoOLEControlObject ||
                type == Microsoft.Office.Core.MsoShapeType.msoMedia;
            if (noPrimitive) throw UnsupportedKindError(shape, toolName);

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                return type == Microsoft.Office.Core.MsoShapeType.msoAutoShape ? ReconstructionKind.AutoShape : ReconstructionKind.TextBox;
            }
            throw UnsupportedKindError(shape, toolName);
        }

        private static ArgumentException UnsupportedKindError(PowerPoint.Shape shape, string toolName)
        {
            return new ArgumentException(toolName + ": shape kind '" + ShapeKindLabel(shape) +
                "' is not supported for cross-slide copy/move - no property-based reconstruction exists for it in this tool. " +
                "Supported kinds: text box, basic autoshape, table, chart, SmartArt.");
        }

        private static PowerPoint.Shape CallAddAndGetNewShape(PowerPoint.Slide destSlide, Action addCall)
        {
            int countBefore = destSlide.Shapes.Count;
            addCall();
            if (destSlide.Shapes.Count != countBefore + 1)
                throw new InvalidOperationException("Reconstruction did not produce exactly one new shape.");
            return destSlide.Shapes[destSlide.Shapes.Count];
        }

        private static PowerPoint.Shape ReconstructTextBox(PowerPoint.Shape source, PowerPoint.Slide destSlide, int destSlideIndex, float left, float top)
        {
            string text = source.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue ? source.TextFrame.TextRange.Text : "";
            var json = new Dictionary<string, object>
            {
                ["slideIndex"] = destSlideIndex,
                ["left"] = left, ["top"] = top,
                ["width"] = source.Width, ["height"] = source.Height,
                ["text"] = text,
            };
            return CallAddAndGetNewShape(destSlide, () => AddTextBox(BuildJson(json)));
        }

        private static PowerPoint.Shape ReconstructAutoShape(PowerPoint.Shape source, PowerPoint.Slide destSlide, int destSlideIndex, float left, float top, string toolName)
        {
            int code = (int)source.AutoShapeType;
            string shapeTypeName;
            if (!AutoShapeCodeToName.TryGetValue(code, out shapeTypeName))
                throw new ArgumentException(toolName + ": this shape's autoshape preset has no name in add_shape's 26-preset vocabulary - not supported for cross-slide copy/move.");

            var json = new Dictionary<string, object>
            {
                ["slideIndex"] = destSlideIndex,
                ["shapeType"] = shapeTypeName,
                ["left"] = left, ["top"] = top,
                ["width"] = source.Width, ["height"] = source.Height,
            };
            if (source.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && source.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                json["text"] = source.TextFrame.TextRange.Text;

            return CallAddAndGetNewShape(destSlide, () => AddShape(BuildJson(json)));
        }

        // Deliberate non-goal: per-cell shading/border formatting is not
        // copied - no existing per-cell-formatting reader exists in this
        // codebase, and building one is its own feature-sized addition.
        private static PowerPoint.Shape ReconstructTable(PowerPoint.Shape source, PowerPoint.Slide destSlide, int destSlideIndex, float left, float top)
        {
            PowerPoint.Table srcTable = source.Table;
            int rows = srcTable.Rows.Count;
            int cols = srcTable.Columns.Count;
            var cellRows = new List<List<string>>();
            for (int r = 1; r <= rows; r++)
            {
                var rowCells = new List<string>();
                for (int c = 1; c <= cols; c++)
                    rowCells.Add(srcTable.Cell(r, c).Shape.TextFrame.TextRange.Text);
                cellRows.Add(rowCells);
            }

            var json = new Dictionary<string, object>
            {
                ["slideIndex"] = destSlideIndex,
                ["rows"] = rows, ["cols"] = cols,
                ["x"] = left, ["y"] = top, ["w"] = source.Width, ["h"] = source.Height,
                ["cells"] = cellRows,
            };
            PowerPoint.Shape dest = CallAddAndGetNewShape(destSlide, () => AddTable(BuildJson(json)));
            dest.Table.FirstRow = srcTable.FirstRow;
            dest.Table.HorizBanding = srcTable.HorizBanding;
            return dest;
        }

        // Deliberately reads live off SeriesCollection()/XValues/Values/Name
        // rather than reopening ChartData.Workbook, which this codebase's own
        // debugging history (STATUS.md's "Live debugging session") documents
        // as the flakiest COM path here. XValues/Values marshaling for a
        // chart built by this same tool's SetSourceData call is unverified
        // without a live session - see the plan's Risks.
        private static PowerPoint.Shape ReconstructChart(PowerPoint.Shape source, PowerPoint.Slide destSlide, int destSlideIndex, float left, float top, string toolName)
        {
            dynamic chart = source.Chart;
            int typeCode = (int)chart.ChartType;
            string kindName;
            if (!ChartTypeCodeToName.TryGetValue(typeCode, out kindName))
                throw new ArgumentException(toolName + ": this chart's type is not one of the named kinds add_chart supports - not supported for cross-slide copy/move.");

            dynamic seriesCollection = chart.SeriesCollection();
            int seriesCount = (int)seriesCollection.Count;
            if (seriesCount == 0)
                throw new InvalidOperationException(toolName + ": chart has no series - nothing to reconstruct.");

            object[] catsRaw = (object[])seriesCollection.Item(1).XValues;
            List<string> categories = catsRaw.Select(o => Convert.ToString(o)).ToList();

            var seriesList = new List<Dictionary<string, object>>();
            for (int i = 1; i <= seriesCount; i++)
            {
                dynamic s = seriesCollection.Item(i);
                string name = (string)s.Name;
                object[] valsRaw = (object[])s.Values;
                List<double> values = valsRaw.Select(o => Convert.ToDouble(o)).ToList();
                seriesList.Add(new Dictionary<string, object> { ["name"] = name, ["values"] = values });
            }

            var json = new Dictionary<string, object>
            {
                ["slideIndex"] = destSlideIndex, ["kind"] = kindName,
                ["categories"] = categories, ["series"] = seriesList,
                ["x"] = left, ["y"] = top, ["w"] = source.Width, ["h"] = source.Height,
            };
            try
            {
                if ((bool)chart.HasTitle) json["title"] = (string)chart.ChartTitle.Text;
            }
            catch { }

            return CallAddAndGetNewShape(destSlide, () => AddChartPpt(BuildJson(json)));
        }

        private static PowerPoint.Shape ReconstructSmartArt(PowerPoint.Shape source, PowerPoint.Slide destSlide, int destSlideIndex, float left, float top, string toolName)
        {
            dynamic smartArt = source.SmartArt;
            string liveLayoutName = (string)smartArt.Layout.Name;
            string layoutKey;
            if (!SmartArtNameToKey.TryGetValue(liveLayoutName, out layoutKey))
                throw new ArgumentException(toolName + ": this SmartArt's layout '" + liveLayoutName +
                    "' is not one of the 7 reconstructable layouts (list/process/cycle/hierarchy/pyramid/matrix/venn) - not supported for cross-slide copy/move.");

            dynamic nodes = smartArt.Nodes;
            int count = (int)nodes.Count;
            var items = new List<string>();
            for (int i = 1; i <= count; i++)
            {
                string text = "";
                try { text = (string)nodes.Item(i).TextFrame2.TextRange.Text; } catch { }
                items.Add(text);
            }

            var json = new Dictionary<string, object>
            {
                ["slideIndex"] = destSlideIndex, ["layout"] = layoutKey,
                ["x"] = left, ["y"] = top, ["w"] = source.Width, ["h"] = source.Height,
                ["items"] = items,
            };
            return CallAddAndGetNewShape(destSlide, () => AddSmartArt(BuildJson(json)));
        }

        // Resolves the source's kind, reconstructs it on destSlide, and
        // copies fill/stroke/text formatting for fidelity. Guarantees: a
        // failed reconstruction never leaves an orphan partial shape on the
        // destination slide (caught and deleted before rethrowing) - which in
        // turn guarantees a failed move_element never deletes the source,
        // since that only happens after this returns successfully.
        private static PowerPoint.Shape ReconstructShapeOnSlide(PowerPoint.Shape source, PowerPoint.Slide destSlide, int destSlideIndex, JsonElement input, string toolName)
        {
            ReconstructionKind kind = ResolveReconstructionKind(source, toolName);
            float destLeft = input.TryGetProperty("left", out var l) ? (float)l.GetDouble() : source.Left;
            float destTop = input.TryGetProperty("top", out var t) ? (float)t.GetDouble() : source.Top;

            int countBefore = destSlide.Shapes.Count;
            try
            {
                PowerPoint.Shape dest;
                switch (kind)
                {
                    case ReconstructionKind.TextBox: dest = ReconstructTextBox(source, destSlide, destSlideIndex, destLeft, destTop); break;
                    case ReconstructionKind.AutoShape: dest = ReconstructAutoShape(source, destSlide, destSlideIndex, destLeft, destTop, toolName); break;
                    case ReconstructionKind.Table: dest = ReconstructTable(source, destSlide, destSlideIndex, destLeft, destTop); break;
                    case ReconstructionKind.Chart: dest = ReconstructChart(source, destSlide, destSlideIndex, destLeft, destTop, toolName); break;
                    case ReconstructionKind.SmartArt: dest = ReconstructSmartArt(source, destSlide, destSlideIndex, destLeft, destTop, toolName); break;
                    default: throw new InvalidOperationException(toolName + ": unhandled reconstruction kind.");
                }
                CopyTextFormatting(source, dest);
                CopyFillFormatting(source, dest);
                CopyStrokeFormatting(source, dest);
                ApplyOptionalName(dest, input);
                return dest;
            }
            catch
            {
                if (destSlide.Shapes.Count > countBefore)
                {
                    try { destSlide.Shapes[destSlide.Shapes.Count].Delete(); } catch { }
                }
                throw;
            }
        }

        private static ToolResult CopyOrMoveElement(JsonElement input, bool cut)
        {
            string toolName = cut ? "move_element" : "copy_element";
            PowerPoint.Shape source = ResolveTopLevelShape(input, toolName);

            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int targetSlideIndex = input.GetProperty("targetSlideIndex").GetInt32();
            PowerPoint.Slides slides = ActivePresentation.Slides;
            if (targetSlideIndex < 0 || targetSlideIndex >= slides.Count)
                throw new ArgumentOutOfRangeException("targetSlideIndex", toolName + ": targetSlideIndex must be between 0 and " + (slides.Count - 1) + ".");
            if (targetSlideIndex == slideIndex)
                throw new ArgumentException(toolName + ": source and target are the same slide - use duplicate_element for a same-slide copy, or set_element_transform to reposition within a slide.");

            PowerPoint.Slide destSlide = slides[targetSlideIndex + 1];
            PowerPoint.Shape dest = ReconstructShapeOnSlide(source, destSlide, targetSlideIndex, input, toolName);
            int newShapeIndex = dest.ZOrderPosition - 1;

            if (cut) source.Delete();

            string action = cut ? "moved" : "copied";
            string extra = cut ? " The shape has been removed from its original slide - other shapes' indices there may have shifted too." : "";
            return new ToolResult
            {
                Output = "Shape " + action + " to slide " + targetSlideIndex + " - new shape at shapeIndex " + newShapeIndex +
                         ". Other shapes' indices on the destination slide may have shifted." + extra +
                         " Re-read the affected slide(s) (read_slide) before addressing another shape by index in the same run.",
                Mutated = true,
                Summary = toolName,
            };
        }

        private static ToolResult CopyElement(JsonElement input) { return CopyOrMoveElement(input, false); }
        private static ToolResult MoveElement(JsonElement input) { return CopyOrMoveElement(input, true); }
    }
}
