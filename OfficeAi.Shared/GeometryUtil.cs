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
    }
}
