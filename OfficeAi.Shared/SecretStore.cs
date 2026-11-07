using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeAi.Shared
{
    // Keeps the AI provider API keys out of the web page's plain-text localStorage. They are encrypted with the
    // Windows user's own key (DPAPI, DataProtectionScope.CurrentUser), so another user or a copied profile folder
    // cannot read them. The host injects them into the page at start (see WebViewBridgeHost) and the page sends
    // changed keys back with a "save-secrets" message.
    public static class SecretStore
    {
        private static string PathFor(string appDataFolderName)
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                appDataFolderName, "secrets.bin");
        }

        // provider id -> API key. Empty when nothing is stored or the file cannot be read.
        public static Dictionary<string, string> Load(string appDataFolderName)
        {
            var result = new Dictionary<string, string>();
            try
            {
                string path = PathFor(appDataFolderName);
                if (!File.Exists(path)) return result;
                byte[] plain = ProtectedData.Unprotect(File.ReadAllBytes(path), null, DataProtectionScope.CurrentUser);
                using (JsonDocument doc = JsonDocument.Parse(Encoding.UTF8.GetString(plain)))
                {
                    if (doc.RootElement.ValueKind != JsonValueKind.Object) return result;
                    foreach (JsonProperty p in doc.RootElement.EnumerateObject())
                        if (p.Value.ValueKind == JsonValueKind.String) result[p.Name] = p.Value.GetString();
                }
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("SecretStore.Load", ex);
            }
            return result;
        }

        public static void Save(string appDataFolderName, Dictionary<string, string> secrets)
        {
            string path = PathFor(appDataFolderName);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            byte[] plain = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(secrets));
            byte[] cipher = ProtectedData.Protect(plain, null, DataProtectionScope.CurrentUser);
            string temp = path + ".tmp";
            File.WriteAllBytes(temp, cipher);
            if (File.Exists(path)) File.Delete(path);
            File.Move(temp, path);
        }
    }
}
