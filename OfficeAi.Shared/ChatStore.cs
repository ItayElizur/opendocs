using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeAi.Shared
{
    public struct ChatRecord
    {
        public string Role;
        public string Text;
        public long Ts;
    }

    public static class ChatStore
    {
        // Derives the 8-hex chat id from an arbitrary stable key. Word/Excel/
        // PowerPoint pass a document file path (see ChatIdForFile); Outlook has
        // no file, so it passes the mailbox's primary SMTP address instead.
        public static string ChatIdForKey(string key)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(key ?? ""));
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 8; i++) sb.Append(hash[i].ToString("x2"));
                return sb.ToString();
            }
        }

        public static string ChatIdForFile(string filePath)
        {
            return ChatIdForKey(filePath);
        }

        private static string ChatPath(string appDataFolderName, string chatId)
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                appDataFolderName, "ChatHistory");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, chatId + ".jsonl");
        }

        private static void AppendRecord(string appDataFolderName, string chatId, string role, string text)
        {
            string json = JsonSerializer.Serialize(new
            {
                role,
                text,
                ts = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
            });
            File.AppendAllText(ChatPath(appDataFolderName, chatId), json + "\n");
        }

        public static void AppendMessage(string appDataFolderName, string chatId, string role, string text)
        {
            AppendRecord(appDataFolderName, chatId, role, text);
        }

        public static void AppendDivider(string appDataFolderName, string chatId)
        {
            AppendRecord(appDataFolderName, chatId, "divider", "");
        }

        public static List<ChatRecord> LoadSinceLastDivider(string appDataFolderName, string chatId)
        {
            string path = ChatPath(appDataFolderName, chatId);
            var all = new List<ChatRecord>();
            if (!File.Exists(path)) return all;

            foreach (string line in File.ReadAllLines(path))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                try
                {
                    using (JsonDocument doc = JsonDocument.Parse(line))
                    {
                        JsonElement root = doc.RootElement;
                        all.Add(new ChatRecord
                        {
                            Role = root.GetProperty("role").GetString(),
                            Text = root.GetProperty("text").GetString(),
                            Ts = root.GetProperty("ts").GetInt64(),
                        });
                    }
                }
                catch (JsonException)
                {
                    // skip a malformed line rather than losing the whole file
                }
            }

            int lastDivider = -1;
            for (int i = 0; i < all.Count; i++)
            {
                if (all[i].Role == "divider") lastDivider = i;
            }
            return all.Skip(lastDivider + 1).Where(r => r.Role != "divider").ToList();
        }

        // FT-1 Task 7b: called once a provisional ("unsaved-...") chat id has
        // just resolved to a real, path-derived one. This store is append-only
        // JSONL, so concatenating the provisional file's lines onto whatever
        // the target already has (the user may have saved over a path they'd
        // chatted about before) is trivially valid and chronologically
        // correct - no merge logic needed beyond "append, then remove the
        // source". A missing source is a silent no-op (nothing to migrate).
        public static void Migrate(string appDataFolderName, string oldChatId, string newChatId)
        {
            string oldPath = ChatPath(appDataFolderName, oldChatId);
            if (!File.Exists(oldPath)) return;

            try
            {
                string newPath = ChatPath(appDataFolderName, newChatId);
                string oldContent = File.ReadAllText(oldPath);
                File.AppendAllText(newPath, oldContent);
                File.Delete(oldPath);
            }
            catch (Exception)
            {
                // Non-fatal: this runs inside pane operations (see
                // DocSettingsStore.Migrate's identical rationale) - an
                // exception here must not kill the add-in.
            }
        }
    }
}
