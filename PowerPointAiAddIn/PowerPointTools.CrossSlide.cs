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
        // add_text_box/add_shape/add_table/add_chart/add_smartart tool (or, for
        // lines/groups, a dedicated reconstruction path) on the destination
        // slide. Pictures, linked pictures, media, and embedded/linked OLE
        // objects have no property-based reconstruction primitive at all
        // (would need Shape.Export() to a temp file and re-import - a real,
        // confirmed COM method, just not implemented here yet) and are refused
        // with a specific named error rather than a lossy approximation; same
        // for freeform/custom-geometry shapes (no primitive for arbitrary
        // vertex data). Same-slide duplication (duplicate_element,
        // PowerPointTools.Elements.cs) is unrestricted, since native
        // Shape.Duplicate() needs no reconstruction.

        private enum ReconstructionKind { TextBox, AutoShape, Line, Table, Chart, SmartArt, Group }

        // Chart reconstruction still goes through add_chart's own JSON tool
        // call (it needs the whole series/categories pipeline, not a single
        // direct COM call), so this reverse lookup is still needed there.
        // AutoShape and SmartArt used to have their own equivalent reverse-
        // lookup maps here too, but don't any more (user-requested,
        // 2026-09-23): both now call Shapes.AddShape/AddSmartArt DIRECTLY
        // with the source's own raw MsoAutoShapeType/SmartArtLayout instead
        // of routing through add_shape's/add_smartart's curated, MODEL-
        // facing name vocabularies (26 presets / 7 layouts) - see
        // ReconstructAutoShape/ReconstructSmartArt below. That vocabulary
        // exists so a model has a manageable set of names to pick from when
        // creating something from scratch; it was never a real limit on
        // what PowerPoint itself can reconstruct, only on this file's
        // previous implementation choice to route through those tools'
        // JSON interface instead of calling the underlying COM method
        // directly with a value already in hand.
        private static readonly Dictionary<int, string> ChartTypeCodeToName = ChartTypes.ByName.ToDictionary(kv => kv.Value, kv => kv.Key);

        private static JsonElement BuildJson(Dictionary<string, object> data)
        {
            byte[] bytes = JsonSerializer.SerializeToUtf8Bytes(data);
            using (JsonDocument doc = JsonDocument.Parse(bytes))
            {
                return doc.RootElement.Clone();
            }
        }

        // Kind names confirmed via .NET reflection against the real referenced
        // PIA (Microsoft.Office.Core.MsoShapeType's full member list) before
        // writing this, not recalled - same discipline as this file's other
        // reflection-confirmed calls.
        private static ReconstructionKind ResolveReconstructionKind(PowerPoint.Shape shape, string toolName)
        {
            if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue) return ReconstructionKind.Table;
            if (shape.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue) return ReconstructionKind.Chart;

            dynamic dshape = shape;
            bool hasSmartArt = false;
            try { hasSmartArt = (int)dshape.HasSmartArt == -1; } catch { }
            if (hasSmartArt) return ReconstructionKind.SmartArt;

            Microsoft.Office.Core.MsoShapeType type = shape.Type;
            if (type == Microsoft.Office.Core.MsoShapeType.msoGroup) return ReconstructionKind.Group;
            if (type == Microsoft.Office.Core.MsoShapeType.msoLine) return ReconstructionKind.Line;

            if (type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture ||
                type == Microsoft.Office.Core.MsoShapeType.msoMedia)
            {
                throw new ArgumentException(toolName + ": shape kind '" + ShapeKindLabel(shape) +
                    "' is not supported for cross-slide copy/move yet - reconstructing a picture/video/audio needs an export step " +
                    "(Shape.Export() to a temp file, then re-import) this tool doesn't do. Not supported: pictures, linked pictures, video/audio.");
            }
            if (type == Microsoft.Office.Core.MsoShapeType.msoEmbeddedOLEObject ||
                type == Microsoft.Office.Core.MsoShapeType.msoLinkedOLEObject ||
                type == Microsoft.Office.Core.MsoShapeType.msoOLEControlObject)
            {
                throw new ArgumentException(toolName + ": shape kind '" + ShapeKindLabel(shape) +
                    "' (embedded/linked OLE object) is not supported for cross-slide copy/move - no property-based way to extract or reconstruct arbitrary embedded object content.");
            }
            if (type == Microsoft.Office.Core.MsoShapeType.msoFreeform)
            {
                throw new ArgumentException(toolName + ": shape kind 'freeform' is not supported for cross-slide copy/move - " +
                    "no reconstruction primitive exists for arbitrary custom vertex geometry.");
            }

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
                "Supported kinds: text box, ANY autoshape preset, line, table, ANY SmartArt layout, chart, and groups of any of these (recursively).");
        }

        // Real-user-confirmed (2026-09-23, live testing): this threw on a
        // table specifically, with no further detail - kindLabel + the exact
        // before/after counts are now included so the NEXT occurrence
        // actually says whether 0, 2, or more shapes appeared (a genuine
        // PowerPoint quirk - e.g. AddTable interacting with an existing
        // content placeholder on the destination slide - is one live
        // hypothesis, not yet confirmed) rather than just "not exactly one".
        private static PowerPoint.Shape CallAddAndGetNewShape(PowerPoint.Slide destSlide, Action addCall, string kindLabel)
        {
            int countBefore = destSlide.Shapes.Count;
            addCall();
            int countAfter = destSlide.Shapes.Count;
            if (countAfter != countBefore + 1)
                throw new InvalidOperationException(kindLabel + " reconstruction added " + (countAfter - countBefore) +
                    " shape(s) instead of exactly 1 (slide had " + countBefore + ", now has " + countAfter +
                    ") - possibly interacting with an existing placeholder on the destination slide.");
            return destSlide.Shapes[countAfter];
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
            return CallAddAndGetNewShape(destSlide, () => AddTextBox(BuildJson(json)), "Text box");
        }

        // Calls Shapes.AddShape directly with the source's own raw
        // MsoAutoShapeType (confirmed via reflection to be exactly AddShape's
        // parameter type) - generically supports every one of PowerPoint's
        // ~180 autoshape presets, not just add_shape's curated 26-name
        // vocabulary. Gets an immediate shape reference back from AddShape
        // itself, so (unlike Table/Chart/TextBox, which still route through
        // a JSON tool call) there's no shape-count check that could fail.
        private static PowerPoint.Shape ReconstructAutoShape(PowerPoint.Shape source, PowerPoint.Slide destSlide, float left, float top)
        {
            PowerPoint.Shape dest = destSlide.Shapes.AddShape(source.AutoShapeType, left, top, source.Width, source.Height);
            if (source.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && source.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                string text = source.TextFrame.TextRange.Text;
                PowerPoint.TextRange range = dest.TextFrame.TextRange;
                range.Text = text;
                ApplyAutoDirection(range, text);
            }
            return dest;
        }

        // A line's begin/end direction isn't separately readable on an
        // EXISTING Shape - Shapes.AddLine's BeginX/BeginY/EndX/EndY are
        // constructor arguments, not properties you can read back. PowerPoint
        // represents an existing line as a bounding box (Left/Top/Width/
        // Height) plus HorizontalFlip/VerticalFlip - the exact same mechanism
        // CopyRotationAndFlip already replays for every other shape kind, so
        // this creates the line in a default top-left-to-bottom-right
        // direction and lets the caller's CopyRotationAndFlip call flip it to
        // match, reproducing the original diagonal. UNVERIFIED without a live
        // test - if wrong, a line copied/moved cross-slide would come out
        // mirrored on one axis; the fallback would be reading
        // ConnectorFormat/computing direction some other way.
        private static PowerPoint.Shape ReconstructLine(PowerPoint.Shape source, PowerPoint.Slide destSlide, float left, float top)
        {
            return destSlide.Shapes.AddLine(left, top, left + source.Width, top + source.Height);
        }

        // Deliberate non-goal: per-cell shading/border formatting and merged-
        // cell STRUCTURE are not reproduced - no existing per-cell-formatting
        // reader exists in this codebase, and re-merging cells on the
        // destination is its own feature-sized addition. But a merged cell on
        // the SOURCE table is not just an unreproduced nicety: real-user-
        // confirmed (2026-09-23) this crashed the entire copy with "The
        // specified value is out of range" - PowerPoint's Table.Cell(r,c)
        // throws for grid positions that fall inside a merged region but
        // aren't its anchor cell (a documented COM quirk, not a coding
        // mistake in the original loop), so a source table with ANY merged
        // cells made every single copy_element/move_element on that table
        // fail outright. Each cell read is now independently caught - a
        // merged-away position just comes back blank instead of aborting the
        // whole table.
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
                {
                    string text = "";
                    try { text = srcTable.Cell(r, c).Shape.TextFrame.TextRange.Text; } catch { }
                    rowCells.Add(text);
                }
                cellRows.Add(rowCells);
            }

            var json = new Dictionary<string, object>
            {
                ["slideIndex"] = destSlideIndex,
                ["rows"] = rows, ["cols"] = cols,
                ["x"] = left, ["y"] = top, ["w"] = source.Width, ["h"] = source.Height,
                ["cells"] = cellRows,
            };
            PowerPoint.Shape dest = CallAddAndGetNewShape(destSlide, () => AddTable(BuildJson(json)), "Table");
            try { dest.Table.FirstRow = srcTable.FirstRow; } catch { }
            try { dest.Table.HorizBanding = srcTable.HorizBanding; } catch { }
            // Table "style" (banding colors, header emphasis, border theme) -
            // confirmed via reflection: Table.Style is read-only (a
            // TableStyle with an Id/Name), applied via the separate
            // ApplyStyle(styleId, saveFormatting) method rather than a
            // settable Style property.
            try { dest.Table.ApplyStyle(srcTable.Style.Id, true); } catch { }

            // Real-user-confirmed gap (2026-09-23): only cell TEXT was
            // copied, not each cell's own formatting (bold/color/size/etc
            // within a cell, or that cell's own fill). Each table cell has
            // its own Shape (Cell.Shape, confirmed via reflection) - a real
            // Shape with its own TextFrame/Fill/Line - so it gets the SAME
            // treatment as any other shape: a native format-painter pass
            // first, then the targeted per-run text copy for exact
            // fidelity. UNVERIFIED without a live test: whether PickUp/Apply
            // behaves the same on a cell's Shape (reached via Table.Cell(r,c)
            // .Shape, not the slide's own Shapes collection) as on a normal
            // top-level shape - if it doesn't, the per-run CopyTextFormatting/
            // CopyTextEffects calls right after it still cover the main
            // complaint (text formatting) independently. A merged cell
            // throws reading Cell(r,c) on either side (the same PowerPoint
            // quirk documented above for the text read) - caught per-cell so
            // one merged cell doesn't skip formatting for the rest of the
            // table. Reproducing the MERGE itself (not just formatting)
            // remains a non-goal: this PIA exposes no way to read back which
            // cells are already merged.
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    try
                    {
                        PowerPoint.Shape srcCellShape = srcTable.Cell(r, c).Shape;
                        PowerPoint.Shape destCellShape = dest.Table.Cell(r, c).Shape;
                        CopyViaPickUpApply(srcCellShape, destCellShape);
                        CopyTextFormatting(srcCellShape, destCellShape);
                        CopyTextEffects(srcCellShape, destCellShape);
                    }
                    catch { }
                }
            }
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

            return CallAddAndGetNewShape(destSlide, () => AddChartPpt(BuildJson(json)), "Chart");
        }

        // Reuses the source's own live SmartArtLayout object directly
        // (Shape.SmartArt.Layout, confirmed via reflection to be EXACTLY
        // Shapes.AddSmartArt's parameter type) instead of name-matching
        // against add_smartart's curated 7-layout vocabulary - generically
        // supports every layout in the presentation's SmartArt gallery.
        // Clears the layout's default placeholder nodes and re-adds the
        // source's real node text, the same sequence add_smartart's own
        // handler uses (AddSmartArt seeds "[Text]" placeholders that need
        // clearing first, or the real content would be appended after them
        // instead of replacing them).
        private static PowerPoint.Shape ReconstructSmartArt(PowerPoint.Shape source, PowerPoint.Slide destSlide, float left, float top)
        {
            dynamic srcSmartArt = source.SmartArt;
            dynamic layout = srcSmartArt.Layout;
            dynamic destShapeDyn = destSlide.Shapes.AddSmartArt(layout, left, top, source.Width, source.Height);
            dynamic destSmartArt = destShapeDyn.SmartArt;

            // Real-user-confirmed gap (2026-09-23): SmartArt "style" (the
            // gallery's color variation AND the 3-D/bevel look) wasn't copied
            // at all - only node text was. Both QuickStyle (governs the 3-D/
            // flat "look", e.g. Polished/Inset/Cartoon) and Color (the
            // color-variation gallery, e.g. "Colorful - Accent Colors") are
            // confirmed via reflection to be plain read/write properties on
            // Microsoft.Office.Core.SmartArt taking the source's own live
            // SmartArtQuickStyle/SmartArtColor object directly - same
            // "reuse the source's own reference" pattern as Layout above,
            // not a name-matched vocabulary.
            try { destSmartArt.QuickStyle = srcSmartArt.QuickStyle; } catch { }
            try { destSmartArt.Color = srcSmartArt.Color; } catch { }

            dynamic existingNodes = destSmartArt.Nodes;
            for (int i = (int)existingNodes.Count; i >= 1; i--)
            {
                existingNodes.Item(i).Delete();
            }

            dynamic srcNodes = srcSmartArt.Nodes;
            int count = (int)srcNodes.Count;
            for (int i = 1; i <= count; i++)
            {
                string text = "";
                try { text = (string)srcNodes.Item(i).TextFrame2.TextRange.Text; } catch { }
                try
                {
                    dynamic node = destSmartArt.Nodes.Add();
                    node.TextFrame2.TextRange.Text = text;
                }
                catch { /* one bad node shouldn't lose every other node's text */ }
            }
            return (PowerPoint.Shape)destShapeDyn;
        }

        // Recursively reconstructs every child of a group at its own
        // translated position, then re-groups them (Shapes.Range(...).Group(),
        // the same COM call group_element already uses). Each child gets its
        // own full formatting + name-inheritance via ReconstructShapeCore,
        // exactly as if it had been copied individually. If any descendant
        // (at any nesting depth) is an unsupported kind, the exception
        // propagates all the way up - the caller (ReconstructShapeOnSlide)
        // deletes every shape reconstructed so far in this whole top-level
        // call, so a group copy/move never leaves a partial group behind.
        //
        // UNVERIFIED without a live test: GroupItems children are assumed to
        // report slide-absolute Left/Top (not group-relative) - if wrong,
        // every child would land shifted incorrectly relative to each other
        // (likely still roughly in the right area, just not reproducing the
        // group's exact internal layout).
        //
        // Real-user-confirmed (2026-09-23, live testing, THIRD round): a
        // group containing a SmartArt failed outright with PowerPoint's own
        // "Grouping is disabled for the selected shapes." - a genuine,
        // documented PowerPoint limitation: a SmartArt graphic cannot be
        // grouped together with ANY other shape via Group(), at any nesting
        // level, full stop. (An earlier theory here - that this was an
        // index/ordering bug in identifying which shapes to group - was
        // wrong; the name-based Shapes.Range(names) fix below is still
        // correct and kept, but it doesn't change WHETHER Group() accepts a
        // SmartArt in its range, which PowerPoint refuses unconditionally.)
        // Given that hard constraint, a SmartArt child is pulled OUT of the
        // Group() call entirely and returned as an extra, ungrouped, top-
        // level sibling shape instead (via the `extraStandalone` list,
        // threaded through the whole recursion) - still positioned exactly
        // where it belonged inside the original group's layout, just not
        // literally part of any Group() PowerPoint object, because no COM
        // call can make that true. The caller (ReconstructShapeOnSlide/
        // CopyOrMoveElement) reports every such extra shape's index clearly
        // in the tool result rather than leaving the user to discover a
        // stray shape on their own.
        private static PowerPoint.Shape ReconstructGroup(PowerPoint.Shape source, PowerPoint.Slide destSlide, int destSlideIndex, float left, float top, string toolName, List<PowerPoint.Shape> extraStandalone)
        {
            float offsetX = left - source.Left;
            float offsetY = top - source.Top;
            var childNames = new List<string>();

            foreach (PowerPoint.Shape child in source.GroupItems)
            {
                float childLeft = child.Left + offsetX;
                float childTop = child.Top + offsetY;
                PowerPoint.Shape newChild;
                try
                {
                    newChild = ReconstructShapeCore(child, destSlide, destSlideIndex, childLeft, childTop, toolName, extraStandalone);
                }
                catch (Exception ex)
                {
                    // Names each failing child, and the nesting depth it
                    // failed at, instead of a bare rethrow - so a future
                    // failure here is a clear, loud message naming exactly
                    // what broke, not a silent bad result.
                    throw new InvalidOperationException(toolName + ": failed reconstructing group child '" + child.Name +
                        "' (" + ShapeKindLabel(child) + ") inside group '" + source.Name + "': " + ex.Message, ex);
                }
                string uniqueChildName = MakeUniqueNameOnSlide(newChild, child.Name);
                if (uniqueChildName != newChild.Name) newChild.Name = uniqueChildName;

                bool childIsSmartArt;
                try { childIsSmartArt = ResolveReconstructionKind(child, toolName) == ReconstructionKind.SmartArt; }
                catch { childIsSmartArt = false; }
                if (childIsSmartArt) extraStandalone.Add(newChild);
                else childNames.Add(newChild.Name);
            }

            if (childNames.Count == 0)
            {
                // Every child was SmartArt (pulled into extraStandalone
                // above) - nothing left that could ever be "this group".
                // Use the last one as the primary return value instead of
                // throwing; it's still on the destination slide, correctly
                // positioned, and the caller reports it as an extra anyway.
                if (extraStandalone.Count > 0)
                {
                    PowerPoint.Shape primary = extraStandalone[extraStandalone.Count - 1];
                    extraStandalone.RemoveAt(extraStandalone.Count - 1);
                    return primary;
                }
                throw new InvalidOperationException(toolName + ": group '" + source.Name + "' has no reconstructable children.");
            }
            if (childNames.Count == 1)
            {
                // PowerPoint's ShapeRange.Group() requires at least 2 shapes
                // - a group whose only reconstructable, groupable child is
                // itself (e.g. a nested group with a single non-SmartArt
                // member) has nothing to literally re-group, so return that
                // one shape directly rather than attempting an invalid
                // single-shape Group() call, which would throw and abort the
                // WHOLE top-level copy/move over one single-child nested
                // group.
                return destSlide.Shapes[childNames[0]];
            }

            PowerPoint.ShapeRange range = destSlide.Shapes.Range(childNames.ToArray());
            return range.Group();
        }

        // Dispatches on kind and applies the formatting appropriate to it -
        // separated from ReconstructShapeOnSlide so ReconstructGroup can call
        // it recursively per child without re-entering the top-level
        // orphan-cleanup try/catch (one cleanup, at the true top level, covers
        // every shape any depth of recursion created).
        private static PowerPoint.Shape ReconstructShapeCore(PowerPoint.Shape source, PowerPoint.Slide destSlide, int destSlideIndex, float left, float top, string toolName, List<PowerPoint.Shape> extraStandalone)
        {
            ReconstructionKind kind = ResolveReconstructionKind(source, toolName);
            PowerPoint.Shape dest;
            switch (kind)
            {
                case ReconstructionKind.TextBox: dest = ReconstructTextBox(source, destSlide, destSlideIndex, left, top); break;
                case ReconstructionKind.AutoShape: dest = ReconstructAutoShape(source, destSlide, left, top); break;
                case ReconstructionKind.Line: dest = ReconstructLine(source, destSlide, left, top); break;
                case ReconstructionKind.Table: dest = ReconstructTable(source, destSlide, destSlideIndex, left, top); break;
                case ReconstructionKind.Chart: dest = ReconstructChart(source, destSlide, destSlideIndex, left, top, toolName); break;
                case ReconstructionKind.SmartArt: dest = ReconstructSmartArt(source, destSlide, left, top); break;
                case ReconstructionKind.Group: dest = ReconstructGroup(source, destSlide, destSlideIndex, left, top, toolName, extraStandalone); break;
                default: throw new InvalidOperationException(toolName + ": unhandled reconstruction kind.");
            }

            if (kind == ReconstructionKind.Group)
            {
                // Every child already carries its own full formatting from
                // its own recursive call above - only rotation is meaningful
                // on the group shape itself (Fill/Line/Text don't apply to a
                // group the way they do a leaf shape).
                CopyRotationAndFlip(source, dest);
            }
            else if (kind == ReconstructionKind.Line)
            {
                // Lines have no fill/text - stroke + geometry only. PickUp/
                // Apply first as a generic broad pass (picks up e.g. shadow/
                // 3-D on a line too), then the targeted stroke copy on top.
                CopyViaPickUpApply(source, dest);
                CopyStrokeFormatting(source, dest);
                CopyRotationAndFlip(source, dest);
            }
            else if (kind == ReconstructionKind.Table || kind == ReconstructionKind.Chart || kind == ReconstructionKind.SmartArt)
            {
                // Real-user-confirmed (2026-09-23): applying the leaf-shape
                // Text/Fill/Stroke formatting pipeline to a table crashed the
                // whole copy ("The specified value is out of range") - these
                // three kinds are graphic frames, not plain shapes: text
                // lives inside a table's individual cells (not the outer
                // frame), and Fill/Line on the outer frame either don't
                // apply the same way or aren't what a user means by "copy
                // this table/chart/SmartArt's formatting" anyway. Each of
                // these already gets its OWN dedicated style copy where one
                // exists (ReconstructTable's ApplyStyle call,
                // ReconstructSmartArt's QuickStyle/Color); only rotation and
                // (harmless, no-op where absent) adjustment handles apply
                // uniformly here.
                CopyRotationAndFlip(source, dest);
                CopyShapeAdjustments(source, dest);
            }
            else
            {
                // User-asked (2026-09-23): a generic format-painter pass
                // FIRST (see CopyViaPickUpApply's comment) - this is meant
                // to catch bevel/3-D and anything else not individually
                // hardcoded below, natively, via PowerPoint's own format
                // painter rather than another hand-enumerated property list.
                // The targeted calls that follow then guarantee per-run text
                // fidelity and the other specifics this tool documents.
                CopyViaPickUpApply(source, dest);
                CopyTextFormatting(source, dest);
                CopyTextEffects(source, dest);
                CopyFillFormatting(source, dest);
                CopyStrokeFormatting(source, dest);
                CopyRotationAndFlip(source, dest);
                CopyShapeAdjustments(source, dest);

                // Real-user-confirmed (2026-09-23, live testing): a shape's
                // Height came out shrunk after being reconstructed, even
                // though it was created at source.Height explicitly - a later
                // copy of that already-reconstructed shape did NOT shrink
                // further, consistent with PowerPoint's own text autofit
                // (TextFrame.AutoSize) reacting once to the mismatch between
                // the shape's initial default-theme text (from AddTextBox/
                // AddShape's own handler) and the real per-run fonts
                // CopyTextFormatting just applied, and settling once font
                // sizes actually match the source's. Re-asserting the size
                // after every formatting step guarantees the final result
                // matches the source regardless of what autofit did in
                // between - width/height were never meant to be affected by
                // any of the formatting calls above.
                try { dest.Width = source.Width; dest.Height = source.Height; } catch { }
            }
            return dest;
        }

        // Resolves the source's kind, reconstructs it on destSlide (recursing
        // through every level of a group), and copies full formatting.
        // Guarantees: a failed reconstruction never leaves ANY orphan shape on
        // the destination slide - every shape added since this call started
        // (one for a leaf kind, several for a group) is deleted before
        // rethrowing - which in turn guarantees a failed move_element never
        // deletes the source, since that only happens after this returns
        // successfully. Inherits the source's own Name (deduped on the
        // destination slide) unless the caller passed an explicit `name`.
        private static PowerPoint.Shape ReconstructShapeOnSlide(PowerPoint.Shape source, PowerPoint.Slide destSlide, int destSlideIndex, JsonElement input, string toolName, List<PowerPoint.Shape> extraStandalone)
        {
            float destLeft = input.TryGetProperty("left", out var l) ? (float)l.GetDouble() : source.Left;
            float destTop = input.TryGetProperty("top", out var t) ? (float)t.GetDouble() : source.Top;

            int countBefore = destSlide.Shapes.Count;
            try
            {
                PowerPoint.Shape dest = ReconstructShapeCore(source, destSlide, destSlideIndex, destLeft, destTop, toolName, extraStandalone);
                string named = ApplyOptionalName(dest, input);
                if (named == null)
                {
                    string unique = MakeUniqueNameOnSlide(dest, source.Name);
                    if (unique != dest.Name) dest.Name = unique;
                }
                return dest;
            }
            catch
            {
                while (destSlide.Shapes.Count > countBefore)
                {
                    try { destSlide.Shapes[destSlide.Shapes.Count].Delete(); }
                    catch { break; }
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
            var extraStandalone = new List<PowerPoint.Shape>();
            PowerPoint.Shape dest = ReconstructShapeOnSlide(source, destSlide, targetSlideIndex, input, toolName, extraStandalone);
            int newShapeIndex = dest.ZOrderPosition - 1;

            if (cut) source.Delete();

            string action = cut ? "moved" : "copied";
            string extra = cut ? " The shape has been removed from its original slide - other shapes' indices there may have shifted too." : "";
            string smartArtNote = "";
            if (extraStandalone.Count > 0)
            {
                string idxList = string.Join(", ", extraStandalone.Select(s => (s.ZOrderPosition - 1).ToString()));
                smartArtNote = " Note: this group contained SmartArt, which PowerPoint does not allow grouping together with any other shape (a real PowerPoint limitation, not a bug in this tool) - it was reconstructed as a separate, ungrouped shape at shapeIndex(es) " +
                    idxList + ", positioned to match its place in the original group.";
            }
            return new ToolResult
            {
                Output = "Shape " + action + " to slide " + targetSlideIndex + " - new shape at shapeIndex " + newShapeIndex +
                         ". Other shapes' indices on the destination slide may have shifted." + extra + smartArtNote +
                         " Re-read the affected slide(s) (read_slide) before addressing another shape by index in the same run.",
                Mutated = true,
                Summary = toolName,
            };
        }

        private static ToolResult CopyElement(JsonElement input) { return CopyOrMoveElement(input, false); }
        private static ToolResult MoveElement(JsonElement input) { return CopyOrMoveElement(input, true); }
    }
}
