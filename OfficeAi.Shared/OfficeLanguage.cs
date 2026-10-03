namespace OfficeAi.Shared
{
    /// <summary>
    /// Resolves an Office LCID to the panel's own supported UI language. Only
    /// "he"/"en" are supported; any LCID that isn't Hebrew (1037) resolves to
    /// "en". See OfficeLanguage.cs.md for the LCID property-access details.
    /// </summary>
    public static class OfficeLanguage
    {
        private const int HebrewLcid = 1037;

        public static string ResolveUiLanguage(int lcid)
        {
            return lcid == HebrewLcid ? "he" : "en";
        }

        // The one place the brand name's per-language rendering is decided.
        // "אופן דוקס" is a phonetic transliteration, not a translation -
        // matches chat-ui.ts's panelTitle Hebrew string. See OfficeLanguage.cs.md.
        public static string ResolveBrandName(int lcid)
        {
            return ResolveUiLanguage(lcid) == "he" ? "אופן דוקס" : "OpenDocs";
        }
    }
}
