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
        // kind selects the operation: add (localPath), crop (l/t/r/b fractions) or replace (localPath, optional keepCrop).
        private static ToolResult EditPicture(JsonElement input)
        {
            string kind = input.GetProperty("kind").GetString();
            switch (kind)
            {
                case "add": return AddPictureToSlide(input);
                case "crop": return CropImage(input);
                case "replace": return ReplaceImagePpt(input);
                default: throw new ArgumentException("edit_picture: unknown kind '" + kind + "'. Valid: add, crop, replace.");
            }
        }

        // Inserts a picture from a local file. Only one of width/height scales proportionally; both stretch to
        // exactly that size; neither keeps the natural size, shrunk to fit if it would overflow the slide.
        // An omitted left/top centers the picture on that axis.
        private static ToolResult AddPictureToSlide(JsonElement input)
        {
            string localPath = input.GetProperty("localPath").GetString();
            string pathError = PathGuard.LocalFileError(localPath);
            if (pathError != null)
                return new ToolResult { Output = "edit_picture: " + pathError, IsError = true };
            localPath = localPath.Trim();

            float? width = input.TryGetProperty("width", out var wEl) && wEl.ValueKind == JsonValueKind.Number ? (float?)wEl.GetDouble() : null;
            float? height = input.TryGetProperty("height", out var hEl) && hEl.ValueKind == JsonValueKind.Number ? (float?)hEl.GetDouble() : null;
            if ((width.HasValue && width.Value <= 0) || (height.HasValue && height.Value <= 0))
                throw new ArgumentException("edit_picture: width and height must be greater than 0.");

            PowerPoint.Slide slide = ResolveSlide(input);
            float slideW = ActivePresentation.PageSetup.SlideWidth, slideH = ActivePresentation.PageSetup.SlideHeight;

            PowerPoint.Shape shape = slide.Shapes.AddPicture(localPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, -1, -1);
            // The shape now exists - delete it again if anything below throws, rather than leave it unplaced.
            try
            {
                float w = shape.Width, h = shape.Height;
                if (width.HasValue && height.HasValue) { w = width.Value; h = height.Value; }
                else if (width.HasValue) { h = h * width.Value / w; w = width.Value; }
                else if (height.HasValue) { w = w * height.Value / h; h = height.Value; }
                else
                {
                    float scale = Math.Min(1f, Math.Min(slideW / w, slideH / h));
                    w *= scale;
                    h *= scale;
                }

                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                shape.Width = w;
                shape.Height = h;
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                shape.Left = input.TryGetProperty("left", out var lEl) && lEl.ValueKind == JsonValueKind.Number ? (float)lEl.GetDouble() : (slideW - w) / 2;
                shape.Top = input.TryGetProperty("top", out var tEl) && tEl.ValueKind == JsonValueKind.Number ? (float)tEl.GetDouble() : (slideH - h) / 2;

                string named = ApplyOptionalName(shape, input);
                int newShapeIndex = shape.ZOrderPosition - 1;
                return new ToolResult
                {
                    Output = "Picture added at shapeIndex " + newShapeIndex + (named != null ? " (\"" + named + "\")" : "") + ".",
                    Mutated = true,
                };
            }
            catch
            {
                try { shape.Delete(); } catch { }
                throw;
            }
        }

        // Crop works on a plain/linked picture, or a placeholder that holds one. Replace does not accept a
        // placeholder: it deletes the old shape, and PowerPoint then recreates an empty placeholder in its place
        // (seen live on a title: its text was lost and an empty title appeared).
        private static void RequirePicture(PowerPoint.Shape shape, JsonElement input, string verb, bool allowPlaceholder)
        {
            var type = shape.Type;
            bool isPicture = type == Microsoft.Office.Core.MsoShapeType.msoPicture || type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture;
            if (!isPicture && allowPlaceholder && type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                isPicture = shape.PlaceholderFormat.ContainedType == Microsoft.Office.Core.MsoShapeType.msoPicture;
            if (isPicture) return;

            string where = "shape " + input.GetProperty("shapeIndex") + " (\"" + shape.Name + "\") on slide " + input.GetProperty("slideIndex");
            throw new ArgumentException("edit_picture: " + where + " is not a picture, so it cannot be " + verb + " - nothing was changed. " +
                (type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder && !allowPlaceholder
                    ? "A placeholder would be deleted and leave an empty one behind. "
                    : "") +
                "Call read_slide to find the picture's shapeIndex.");
        }

        private static ToolResult CropImage(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            RequirePicture(shape, input, "cropped", true);
            float l = (float)input.GetProperty("l").GetDouble();
            float t = (float)input.GetProperty("t").GetDouble();
            float r = (float)input.GetProperty("r").GetDouble();
            float b = (float)input.GetProperty("b").GetDouble();
            // Approximation: fractions apply to the shape's CURRENT size, not the
            // original uncropped image (Interop exposes no natural-size property
            // once resized/cropped). See PowerPointTools.Images.cs.md for why.
            shape.PictureFormat.CropLeft = l * shape.Width;
            shape.PictureFormat.CropTop = t * shape.Height;
            shape.PictureFormat.CropRight = r * shape.Width;
            shape.PictureFormat.CropBottom = b * shape.Height;
            return new ToolResult { Output = "Image cropped.", Mutated = true };
        }

        private static ToolResult ReplaceImagePpt(JsonElement input)
        {
            string localPath = input.GetProperty("localPath").GetString();
            string pathError = PathGuard.LocalFileError(localPath);
            if (pathError != null)
                return new ToolResult { Output = "edit_picture: " + pathError, IsError = true };
            localPath = localPath.Trim();
            PowerPoint.Shape oldShape = ResolveShape(input);
            RequirePicture(oldShape, input, "replaced", false);
            bool keepCrop = ToolArgs.OptionalFlag(input, "keepCrop");

            float left = oldShape.Left, top = oldShape.Top, width = oldShape.Width, height = oldShape.Height, rotation = oldShape.Rotation;
            int zPos = oldShape.ZOrderPosition;
            float cropLeft = 0, cropTop = 0, cropRight = 0, cropBottom = 0;
            if (keepCrop)
            {
                cropLeft = oldShape.PictureFormat.CropLeft;
                cropTop = oldShape.PictureFormat.CropTop;
                cropRight = oldShape.PictureFormat.CropRight;
                cropBottom = oldShape.PictureFormat.CropBottom;
            }
            PowerPoint.Slide slide = (PowerPoint.Slide)oldShape.Parent;
            oldShape.Delete();

            PowerPoint.Shape newShape = slide.Shapes.AddPicture(localPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, left, top, width, height);
            newShape.Rotation = rotation;
            if (keepCrop)
            {
                newShape.PictureFormat.CropLeft = cropLeft;
                newShape.PictureFormat.CropTop = cropTop;
                newShape.PictureFormat.CropRight = cropRight;
                newShape.PictureFormat.CropBottom = cropBottom;
            }
            // Restore approximate z-order: send to back, then bring forward to the
            // original stack position.
            newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
            for (int i = 1; i < zPos; i++)
            {
                newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringForward);
            }
            return new ToolResult { Output = "Image replaced.", Mutated = true };
        }

        private static ToolResult SetPictureOpacity(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            float opacity = (float)input.GetProperty("opacity").GetDouble();
            dynamic dShape = shape;
            dShape.Fill.Transparency = 1f - opacity;
            return new ToolResult { Output = "Opacity updated.", Mutated = true };
        }
    }
}

