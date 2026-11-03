using System;
using System.Text.Json;

namespace OfficeAi.Shared
{
    /// <summary>
    /// The shared <c>save</c> tool's flow, identical in Word, Excel and PowerPoint: save in place, or - when
    /// <c>saveAs</c> is true or the file has never been saved - run the app's Save As dialog. Each app supplies only the
    /// three COM steps; the flag parsing and the result messages live here.
    /// </summary>
    public static class SaveTool
    {
        public const string CancelledMessage = "The user cancelled the Save As dialog; nothing was saved.";

        /// <param name="input">The tool input; <c>saveAs</c> is optional and must be true or false.</param>
        /// <param name="currentPath">The open file's folder path; empty when it has never been saved.</param>
        /// <param name="saveInPlace">Saves the file in place.</param>
        /// <param name="runSaveAsDialog">Runs the Save As dialog and saves; returns false when the user cancelled.</param>
        /// <param name="fullName">The open file's full name, read after saving.</param>
        public static ToolResult Run(JsonElement input, string currentPath, Action saveInPlace, Func<bool> runSaveAsDialog, Func<string> fullName)
        {
            bool wantSaveAs = ToolArgs.OptionalFlag(input, "saveAs");
            if (!wantSaveAs && !string.IsNullOrEmpty(currentPath))
            {
                saveInPlace();
                return new ToolResult { Output = "Saved " + fullName() + ".", Summary = "save" };
            }

            if (!runSaveAsDialog())
            {
                return new ToolResult { Output = CancelledMessage, Summary = "save" };
            }
            return new ToolResult { Output = "Saved as " + fullName() + ".", Summary = "save" };
        }
    }
}
