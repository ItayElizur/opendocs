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
        public static string ChatIdForFile(string filePath)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(filePath));
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 8; i++) sb.Append(hash[i].ToString("x2"));
                return sb.ToString();
            }
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
    }
}
