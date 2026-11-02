using System;
using System.Text.Json;

namespace OfficeAi.Shared
{
    /// <summary>
    /// The flow of the shared <c>save</c> tool, identical in Word, Excel and PowerPoint: save in place, or - when
    /// <c>saveAs</c> is true or the file has never been saved - run the app's Save As dialog. Each app supplies only the
    /// two COM steps; the flag parsing and the result messages live here.
    /// </summary>
    public static class SaveTool
    {
        public const string CancelledMessage = "The user cancelled the Save As dialog; nothing was saved.";

        /// <param name="input">The tool input; <c>saveAs</c> is optional and must be true or false.</param>
        /// <param name="currentPath">The open file's folder path; empty when it has never been saved.</param>
        /// <param name="saveInPlace">Saves the file and returns its full name.</param>
        /// <param name="saveAs">Runs the Save As dialog; returns the saved file's full name, or null when cancelled.</param>
        public static ToolResult Run(JsonElement input, string currentPath, Func<string> saveInPlace, Func<string> saveAs)
        {
            bool wantSaveAs = ToolArgs.OptionalFlag(input, "saveAs");
            if (!wantSaveAs && !string.IsNullOrEmpty(currentPath))
            {
                return new ToolResult { Output = "Saved " + saveInPlace() + ".", Summary = "save" };
            }

            string savedAs = saveAs();
            if (savedAs == null)
            {
                return new ToolResult { Output = CancelledMessage, Summary = "save" };
            }
            return new ToolResult { Output = "Saved as " + savedAs + ".", Summary = "save" };
        }
    }
}
