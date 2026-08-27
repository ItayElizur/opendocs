using System;

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
            if (hex == null) throw new ArgumentException("Color is required, e.g. \"#RRGGBB\".", nameof(hex));
            string h = hex.Trim().TrimStart('#');

            // "abc" is the widely-used CSS shorthand for "aabbcc" - accepted
            // because a model asked for "a light grey" will often produce it,
            // and the old code failed it with an opaque Substring error
            // rather than a usable message.
            if (h.Length == 3)
                h = new string(new[] { h[0], h[0], h[1], h[1], h[2], h[2] });

            if (h.Length != 6 || !IsHexDigits(h))
                throw new ArgumentException(
                    "Invalid color \"" + hex + "\". Expected 6-digit hex \"#RRGGBB\" (or 3-digit \"#RGB\"), e.g. \"#1A73E8\".",
                    nameof(hex));

            int r = Convert.ToInt32(h.Substring(0, 2), 16);
            int g = Convert.ToInt32(h.Substring(2, 2), 16);
            int b = Convert.ToInt32(h.Substring(4, 2), 16);
            return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
        }

        private static bool IsHexDigits(string s)
        {
            foreach (char c in s)
            {
                bool ok = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
                if (!ok) return false;
            }
            return true;
        }
    }
}
