using System;
using System.IO;
using System.Text.Json;

namespace OfficeAi.Shared
{
    public struct DocSettings
    {
        public string SystemMessage;
    }

    // Per-document settings (currently just the free-text system message),
    // modeled on ChatStore but a single JSON document per file rather than an
    // append-only log - kept in its own DocSettings\ folder, sibling to
    // ChatHistory\, so the chat-log loader never sees a non-.jsonl file.
    public static class DocSettingsStore
    {
        // The system message is prepended to every turn's system prompt (via
        // AgentLoop's systemSuffix), so an unbounded value is a per-turn token
        // cost the user cannot see. Mirrored as a maxlength on the UI textarea.
        public const int MaxSystemMessageChars = 8192;

        private static string StoreDir(string appDataFolderName)
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                appDataFolderName, "DocSettings");
            Directory.CreateDirectory(dir);
            return dir;
        }

        private static string SettingsPath(string appDataFolderName, string chatId)
        {
            return Path.Combine(StoreDir(appDataFolderName), chatId + ".json");
        }

        public static DocSettings Load(string appDataFolderName, string chatId)
        {
            string path = SettingsPath(appDataFolderName, chatId);
            if (!File.Exists(path)) return new DocSettings { SystemMessage = "" };

            try
            {
                string json = File.ReadAllText(path);
                using (JsonDocument doc = JsonDocument.Parse(json))
                {
                    JsonElement root = doc.RootElement;
                    string message = root.TryGetProperty("systemMessage", out var m) && m.ValueKind == JsonValueKind.String
                        ? m.GetString()
                        : "";
                    return new DocSettings { SystemMessage = message ?? "" };
                }
            }
            catch (Exception)
            {
                // Malformed or unreadable file: return empty rather than
                // throwing - this runs during pane startup, and an exception
                // there is a dead add-in (matches ChatStore.LoadSinceLastDivider's
                // "skip a malformed line rather than losing the whole file" spirit).
                return new DocSettings { SystemMessage = "" };
            }
        }

        public static void Save(string appDataFolderName, string chatId, DocSettings settings)
        {
            string message = settings.SystemMessage ?? "";
            if (message.Length > MaxSystemMessageChars) message = message.Substring(0, MaxSystemMessageChars);

            string path = SettingsPath(appDataFolderName, chatId);
            string json = JsonSerializer.Serialize(new { systemMessage = message });

            // Atomic write: write to a temp file in the same directory, then
            // move over the target, so a crash mid-write cannot leave a
            // truncated file that silently discards the user's guidelines.
            string tempPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            File.WriteAllText(tempPath, json);
            if (File.Exists(path)) File.Delete(path);
            File.Move(tempPath, path);
        }

        // FT-1 Task 7b: called once a provisional ("unsaved-...") chat id has
        // just resolved to a real, path-derived one. Non-empty wins - a single
        // JSON document cannot be concatenated the way ChatStore's JSONL can,
        // so real guidelines already on the target must never be overwritten
        // by an empty provisional value.
        public static void Migrate(string appDataFolderName, string oldChatId, string newChatId)
        {
            string oldPath = SettingsPath(appDataFolderName, oldChatId);
            if (!File.Exists(oldPath)) return;

            try
            {
                DocSettings provisional = Load(appDataFolderName, oldChatId);
                if (!string.IsNullOrEmpty(provisional.SystemMessage))
                {
                    Save(appDataFolderName, newChatId, provisional);
                }
                File.Delete(oldPath);
            }
            catch (Exception)
            {
                // Non-fatal: this runs inside pane operations, and an
                // exception here (locked file, permissions) must not kill the
                // add-in. Worst case the provisional file is left in place and
                // retried on the next save.
            }
        }
    }
}
