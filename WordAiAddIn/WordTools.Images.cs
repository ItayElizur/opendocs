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
        // PP-11: same air-gapped local-file-only rule as Excel's
        // AddImageExcel/PowerPoint's replace_image, worded consistently. The
        // File.Exists check is the one addition over Excel's version -
        // AddPicture on a missing file throws a bare COMException with a
        // useless message; this lets the model correct the path next turn.
        private static string ValidateLocalImagePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("add_image: path is required.");
            if (path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException(
                    "add_image: remote URLs are not supported in this air-gapped deployment - use a local file path.");
            if (!System.IO.File.Exists(path))
                throw new System.IO.FileNotFoundException("add_image: no file at '" + path + "'.");
            return path;
        }

        private static ToolResult AddImage(JsonElement input)
        {
            string path = ValidateLocalImagePath(input.GetProperty("path").GetString());
            bool floating = input.TryGetProperty("floating", out var flEl) && flEl.ValueKind == JsonValueKind.True;
            int afterBlockIndex = input.TryGetProperty("afterBlockIndex", out var abEl) && abEl.ValueKind == JsonValueKind.Number
                ? abEl.GetInt32() : int.MinValue; // sentinel: append at end (-1 already means "start of document" in this file's convention)

            Word.Document doc = ActiveDoc;
            Word.Range at;
            if (afterBlockIndex == int.MinValue)
            {
                at = doc.Content;
                at.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
            else
            {
                at = RangeAfterBlock(afterBlockIndex);
            }

            float? widthPoints = input.TryGetProperty("widthPoints", out var wEl) && wEl.ValueKind == JsonValueKind.Number ? (float?)wEl.GetDouble() : null;
            float? heightPoints = input.TryGetProperty("heightPoints", out var hEl) && hEl.ValueKind == JsonValueKind.Number ? (float?)hEl.GetDouble() : null;
            string altText = input.TryGetProperty("altText", out var altEl) && altEl.ValueKind == JsonValueKind.String ? altEl.GetString() : null;

            float finalWidth, finalHeight;
            string addressability;

            if (floating)
            {
                double left = at.get_Information(Word.WdInformation.wdHorizontalPositionRelativeToPage);
                double top = at.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage);
                Word.Shape shape = doc.Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                    (float)left, (float)top, -1, -1);
                float naturalW = shape.Width, naturalH = shape.Height;
                GeometryUtil.ResolveImageSize(naturalW, naturalH, widthPoints, heightPoints, out finalWidth, out finalHeight);
                shape.Width = finalWidth;
                shape.Height = finalHeight;
                if (altText != null) shape.AlternativeText = altText;
                addressability = "not addressable by apply_commands/updateImageProperties (floating)";
            }
            else
            {
                Word.InlineShape shape = doc.InlineShapes.AddPicture(path, LinkToFile: false, SaveWithDocument: true, Range: at);
                float naturalW = shape.Width, naturalH = shape.Height;
                GeometryUtil.ResolveImageSize(naturalW, naturalH, widthPoints, heightPoints, out finalWidth, out finalHeight);
                shape.Width = finalWidth;
                shape.Height = finalHeight;
                if (altText != null) shape.AlternativeText = altText;
                // InlineShapes is ordered by document position, not insertion
                // time - the new shape is only the LAST entry if it was
                // appended at the document's end. Find its real index by
                // position instead of assuming Count-1.
                int newIndex = -1;
                int shapeStart = shape.Range.Start;
                for (int idx = 0; idx < doc.InlineShapes.Count; idx++)
                {
                    if (doc.InlineShapes[idx + 1].Range.Start == shapeStart) { newIndex = idx; break; }
                }
                addressability = newIndex >= 0
                    ? $"addressable via apply_commands/updateImageProperties at imageIndex {newIndex}"
                    : "inserted, but its index could not be resolved";
            }

            return new ToolResult
            {
                Output = $"Image inserted from '{path}' ({finalWidth:0}x{finalHeight:0}pt) - {addressability}.",
                Mutated = true,
                Summary = "add_image",
            };
        }

    }
}

