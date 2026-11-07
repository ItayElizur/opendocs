using System;
using System.IO;

namespace OfficeAi.Shared
{
    // Checks a file path the model supplied before it reaches an Office "insert picture" call.
    public static class PathGuard
    {
        // Returns null when the path is a full local file path to an existing file, otherwise a short reason.
        // A network (UNC) path is refused because merely checking or opening it makes Windows connect to that server
        // and send the user's login hash; a path inside untrusted document text could point it anywhere.
        public static string LocalFileError(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return "a file path is required.";
            string p = path.Trim();
            if (p.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                p.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                p.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                return "remote URLs are not supported in this air-gapped deployment - use a local file path.";

            bool network = (p[0] == '\\' || p[0] == '/') && p.Length > 1 && (p[1] == '\\' || p[1] == '/');
            bool onDrive = p.Length >= 3 && char.IsLetter(p[0]) && p[1] == ':' && (p[2] == '\\' || p[2] == '/');
            if (network || !onDrive)
                return "only a full local file path on a drive (such as C:\\folder\\picture.png) is allowed - network and relative paths are not.";

            if (!File.Exists(p)) return "no file at '" + p + "'.";
            return null;
        }
    }
}
