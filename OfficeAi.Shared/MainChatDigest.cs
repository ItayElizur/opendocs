using System.Collections.Generic;
using System.Text;

namespace OfficeAi.Shared
{
    // read_main_chat's answer: the main Outlook window's chat since its last New chat, as an item window's model reads
    // it. Only the latest messages, each cut short, so a long main chat cannot flood the draft's context.
    public static class MainChatDigest
    {
        public const int MaxMessages = 20;
        public const int MaxChars = 2000;

        public static string Format(IList<ChatRecord> sinceLastDivider)
        {
            int count = sinceLastDivider?.Count ?? 0;
            if (count == 0) return "The main Outlook window's chat is empty (nothing since its last New chat).";
            int from = count > MaxMessages ? count - MaxMessages : 0;
            var sb = new StringBuilder();
            sb.AppendLine(from > 0
                ? "The last " + MaxMessages + " of " + count + " messages in the main Outlook window's chat:"
                : "The main Outlook window's chat (" + count + " messages):");
            for (int i = from; i < count; i++)
            {
                string text = sinceLastDivider[i].Text ?? "";
                if (text.Length > MaxChars) text = text.Substring(0, MaxChars) + " ...[cut]";
                sb.AppendLine();
                sb.AppendLine(sinceLastDivider[i].Role + ": " + text);
            }
            return sb.ToString().TrimEnd();
        }
    }
}
