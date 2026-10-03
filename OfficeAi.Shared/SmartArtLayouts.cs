using System.Collections.Generic;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Friendly SmartArt layout key to the layout's display name in Office's
    /// built-in gallery, shared by Word and PowerPoint. Display names are
    /// English-only and matched case-insensitively - a non-English Office
    /// install's gallery won't match, surfaced as a distinct error rather
    /// than a silent fallback. See SmartArtLayouts.cs.md.
    /// </summary>
    public static class SmartArtLayouts
    {
        public static readonly Dictionary<string, string> ByName = new Dictionary<string, string>
        {
            ["list"] = "Basic Block List",
            ["process"] = "Basic Process",
            ["cycle"] = "Basic Cycle",
            ["hierarchy"] = "Organization Chart",
            ["pyramid"] = "Basic Pyramid",
            ["matrix"] = "Basic Matrix",
            ["venn"] = "Basic Venn",
        };

        /// <summary>
        /// Maps a layout key to its gallery display name, or throws with the
        /// valid keys listed. Shared so both apps produce the identical error
        /// text for the identical mistake.
        /// </summary>
        public static string DisplayNameFor(string layoutKey, string toolName)
        {
            string targetName;
            if (!ByName.TryGetValue(layoutKey, out targetName))
                throw new System.ArgumentException(
                    toolName + ": unknown layout '" + layoutKey + "'. Valid: " +
                    string.Join(", ", ByName.Keys) + ".");
            return targetName;
        }

        /// <summary>
        /// The "key was valid but this install's gallery has no layout under
        /// that display name" case - see the localisation note above.
        /// </summary>
        public static System.InvalidOperationException NotInGallery(string targetName, string toolName)
        {
            return new System.InvalidOperationException(
                toolName + ": no SmartArt layout named '" + targetName +
                "' was found in this Office install's gallery - this install may be " +
                "non-English, where the built-in gallery's display names differ from " +
                "the standard English ones this tool assumes.");
        }
    }
}
