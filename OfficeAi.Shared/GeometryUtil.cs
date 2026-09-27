using System;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Pure geometry helpers shared by the Word/Excel/PowerPoint tool layers.
    /// </summary>
    public static class GeometryUtil
    {
        // Scales the missing dimension proportionally from the natural size -
        // never distorting an image by defaulting the missing dimension to a
        // constant.
        public static void ResolveImageSize(float naturalWidth, float naturalHeight, float? widthPoints, float? heightPoints, out float finalWidth, out float finalHeight)
        {
            if (widthPoints.HasValue && !heightPoints.HasValue)
                heightPoints = naturalHeight * (widthPoints.Value / naturalWidth);
            else if (heightPoints.HasValue && !widthPoints.HasValue)
                widthPoints = naturalWidth * (heightPoints.Value / naturalHeight);
            finalWidth = widthPoints ?? naturalWidth;
            finalHeight = heightPoints ?? naturalHeight;
        }

        // Corner shorthand for slide-master-pinned elements (logo/watermark/label) -
        // computes a left/top so a shape of the given size sits marginPt inside the
        // named corner of a slideWidth x slideHeight slide.
        public static void ResolveCornerPosition(float slideWidth, float slideHeight, float width, float height, float marginPt, string corner, out float left, out float top)
        {
            switch (corner)
            {
                case "topLeft":
                    left = marginPt; top = marginPt; break;
                case "topRight":
                    left = slideWidth - width - marginPt; top = marginPt; break;
                case "bottomLeft":
                    left = marginPt; top = slideHeight - height - marginPt; break;
                case "bottomRight":
                    left = slideWidth - width - marginPt; top = slideHeight - height - marginPt; break;
                default:
                    throw new ArgumentException("Unknown corner '" + corner + "'. Valid: topLeft, topRight, bottomLeft, bottomRight.");
            }
        }
    }
}
