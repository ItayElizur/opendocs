namespace OfficeAi.Shared
{
    /// <summary>
    /// "#RRGGBB" hex color to the OLE/BGR integer Office's COM APIs want.
    /// Shared by Word/Excel/PowerPoint (each casts the int to its own color
    /// type at the call site - see the seam rule: this assembly never
    /// exposes an Office interop type, only plain int/string, since an
    /// embedded interop type cannot cross an assembly boundary as a generic
    /// type argument - confirmed via CS1769 while building Phase 0's
    /// ShapeTypes.cs).
    /// </summary>
    public static class ColorUtil
    {
        /// <summary>
        /// Note ColorTranslator.ToOle byte-swaps to BGR - do NOT "simplify"
        /// this to a plain (r &lt;&lt; 16) | (g &lt;&lt; 8) | b, which is the
        /// opposite byte order and would silently swap red and blue.
        /// </summary>
        public static int HexToOle(string hex)
        {
            hex = hex.TrimStart('#');
            int r = System.Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = System.Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = System.Convert.ToInt32(hex.Substring(4, 2), 16);
            return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
        }
    }
}
