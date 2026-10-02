using System;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Resolves Office's "Office Theme" setting (File &gt; Account) to a plain
    /// "dark"/"light" verdict via an undocumented registry value (no
    /// object-model property exists). Every failure mode degrades to "light"
    /// rather than throwing. See OfficeTheme.cs.md.
    /// </summary>
    public static class OfficeTheme
    {
        private const string OfficeThemeKey = @"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common";
        private const string OfficeThemeValue = "UI Theme";
        private const string PersonalizeKey = @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
        private const string PersonalizeValue = "AppsUseLightTheme";

        /// <summary>
        /// Test seam matching Microsoft.Win32.Registry.GetValue's signature -
        /// the real read is `registryGetValue ?? Registry.GetValue`. See
        /// OfficeTheme.cs.md.
        /// </summary>
        public static string ReadEffectiveTheme(Func<string, string, object, object> registryGetValue = null)
        {
            var getValue = registryGetValue ?? Microsoft.Win32.Registry.GetValue;

            try
            {
                object raw = getValue(OfficeThemeKey, OfficeThemeValue, null);
                if (!TryToInt(raw, out int uiTheme)) return "light";

                switch (uiTheme)
                {
                    case 3: // Dark Gray
                    case 4: // Black
                        return "dark";
                    case 5: // White
                    case 7: // Colorful
                        return "light";
                    case 6: // Use System Setting
                        return ResolveSystemTheme(getValue);
                    default:
                        return "light";
                }
            }
            catch
            {
                return "light";
            }
        }

        private static string ResolveSystemTheme(Func<string, string, object, object> getValue)
        {
            object raw = getValue(PersonalizeKey, PersonalizeValue, null);
            if (!TryToInt(raw, out int appsUseLightTheme)) return "light";
            return appsUseLightTheme == 0 ? "dark" : "light";
        }

        private static bool TryToInt(object raw, out int value)
        {
            if (raw == null)
            {
                value = 0;
                return false;
            }

            try
            {
                value = Convert.ToInt32(raw);
                return true;
            }
            catch
            {
                value = 0;
                return false;
            }
        }
    }
}
