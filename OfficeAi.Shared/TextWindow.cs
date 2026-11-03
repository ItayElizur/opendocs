using System;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Returns one window of a long text with a note saying what was left out, so a caller can read the rest by
    /// passing the offset the note names. Used by Word's read_object for text boxes, notes, headers/footers and TOCs.
    /// </summary>
    public static class TextWindow
    {
        public const int DefaultCap = 4000;

        /// <param name="text">The full text.</param>
        /// <param name="offset">Character offset to start at; values below 0 mean 0.</param>
        /// <param name="cap">Maximum characters to return.</param>
        public static string Slice(string text, int offset, int cap)
        {
            text = text ?? "";
            if (offset < 0) offset = 0;
            if (cap < 1) cap = DefaultCap;
            if (offset == 0 && text.Length <= cap) return text;
            if (offset >= text.Length)
                return offset == 0 ? "" : "(offset " + offset + " is past the end: the text has " + text.Length + " characters)";

            int end = Math.Min(text.Length, offset + cap);
            // do not cut a surrogate pair in half
            if (end < text.Length && end > offset + 1 && char.IsHighSurrogate(text[end - 1])) end--;
            string slice = text.Substring(offset, end - offset);
            return end < text.Length
                ? slice + "\n... (characters " + offset + "-" + end + " of " + text.Length + "; call read_object again with offset: " + end + " for the rest)"
                : slice + "\n(end of text: characters " + offset + "-" + text.Length + " of " + text.Length + ")";
        }
    }
}
