using System.Collections.Generic;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Friendly SmartArt layout key to the layout's display name in Office's
    /// built-in gallery. Shared by Word and PowerPoint, which previously held
    /// byte-identical copies.
    ///
    /// The values are English display names, matched case-insensitively
    /// against `Application.SmartArtLayouts`. That is a real constraint worth
    /// knowing: on a non-English Office install the gallery's display names
    /// differ, and the lookup finds nothing. Both apps surface that as a
    /// distinct error rather than silently falling back to some other layout,
    /// because "no layout by that name in this install" is the one diagnosis
    /// that actually points at the localisation problem.
    ///
    /// Resolving a key to a live layout object stays app-side: it walks
    /// `Globals.ThisAddIn.Application.SmartArtLayouts`, and `Globals` is
    /// generated per add-in project, so it cannot move here.
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
