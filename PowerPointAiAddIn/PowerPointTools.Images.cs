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
        private static ToolResult CropImage(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            float l = (float)input.GetProperty("l").GetDouble();
            float t = (float)input.GetProperty("t").GetDouble();
            float r = (float)input.GetProperty("r").GetDouble();
            float b = (float)input.GetProperty("b").GetDouble();
            // Approximation, documented deliberately: fractions are applied against
            // the shape's CURRENT on-slide size, not the original uncropped source
            // image - classic Interop has no reliable "natural size" property once a
            // picture has already been resized/cropped on the slide. Correct for a
            // freshly-inserted, never-before-cropped picture; imprecise under
            // repeated crop calls on the same shape.
            shape.PictureFormat.CropLeft = l * shape.Width;
            shape.PictureFormat.CropTop = t * shape.Height;
            shape.PictureFormat.CropRight = r * shape.Width;
            shape.PictureFormat.CropBottom = b * shape.Height;
            return new ToolResult { Output = "Image cropped.", Mutated = true, Summary = "crop_image" };
        }

        private static ToolResult ReplaceImagePpt(JsonElement input)
        {
            string localPath = input.GetProperty("localPath").GetString();
            if (localPath.StartsWith("http://") || localPath.StartsWith("https://"))
            {
                return new ToolResult { Output = "replace_image: remote URLs are not supported in this air-gapped deployment - use a local file path.", IsError = true, Summary = "replace_image" };
            }
            if (!System.IO.File.Exists(localPath))
            {
                return new ToolResult { Output = "replace_image: file not found: " + localPath, IsError = true, Summary = "replace_image" };
            }
            PowerPoint.Shape oldShape = ResolveShape(input);
            bool keepCrop = input.TryGetProperty("keepCrop", out var kc) && kc.ValueKind == JsonValueKind.True;

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
            return new ToolResult { Output = "Image replaced.", Mutated = true, Summary = "replace_image" };
        }

        private static ToolResult SetPictureOpacity(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveShape(input);
            float opacity = (float)input.GetProperty("opacity").GetDouble();
            dynamic dShape = shape;
            dShape.Fill.Transparency = 1f - opacity;
            return new ToolResult { Output = "Opacity updated.", Mutated = true, Summary = "set_picture_opacity" };
        }
    }
}

