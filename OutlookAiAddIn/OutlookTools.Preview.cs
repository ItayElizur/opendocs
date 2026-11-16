using System;
using System.Collections.Generic;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        // Backs the panel's Approve/Deny card (preview_action in OutlookTools.cs). Given the call about to run, returns one JSON
        // object: "blocked" (the exact refusal the call would get, so the panel can skip the card) or the item's type, subject,
        // sender and date. Never the body, and never changes anything.
        private static ToolResult PreviewAction(string mbxKey, JsonElement input)
        {
            string tool = Str(input, "tool", null);
            JsonElement call = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("input", out JsonElement inner) ? inner : default(JsonElement);
            var fields = new Dictionary<string, string>();
            try
            {
                bool isDelete = tool == "delete_email";
                // The item the call acts on (or replies to); a new email or event has none.
                string id = Str(call, "message_id", null) ?? Str(call, "event_id", null);
                if (!string.IsNullOrEmpty(id))
                {
                    dynamic d = ItemById(id, Str(call, "store_id", null) ?? StoreOf(call));
                    fields["type"] = ItemTypeKey(d);
                    try { fields["subject"] = d.Subject ?? ""; } catch { }
                    try { fields["from"] = ItemSender(d); } catch { }
                    try { fields["date"] = ItemDate(d); } catch { }

                    if (isDelete && Bool(call, "permanent", false))
                    {
                        string blocked = PermanentModeRefusal(ModeFor(mbxKey));
                        if (blocked == null)
                        {
                            Outlook.Folder source = null;
                            try { source = (Outlook.Folder)d.Parent; } catch { }
                            Outlook.Folder deleted = source != null
                                ? (Outlook.Folder)source.Store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)
                                : (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                            blocked = PermanentLocationRefusal(fields.ContainsKey("subject") ? fields["subject"] : "", source, deleted);
                        }
                        if (blocked != null) fields["blocked"] = blocked;
                    }
                }
            }
            catch (Exception ex)
            {
                // The id does not resolve: say so now, instead of after the user approves.
                DebugLog.WriteException("PreviewAction " + tool, ex);
                fields.Clear();
                fields["blocked"] = "The item could not be found: " + ex.Message;
            }
            return new ToolResult { Output = JsonSerializer.Serialize(fields), Summary = "preview_action" };
        }

        private static string ItemTypeKey(dynamic item)
        {
            switch ((int)item.Class)
            {
                case 43: return "mail";                  // olMail
                case 53: return "meetingRequest";        // olMeetingRequest
                case 54: case 55: case 56: return "meetingResponse"; // olMeetingResponseNegative / Positive / Tentative
                case 57: return "meetingCancellation";   // olMeetingCancellation
                case 26: return "event";                 // olAppointment
                case 48: return "task";                  // olTask
                default: return "other";
            }
        }

        // An appointment has an organizer; every mail-like item (mail, invitations, responses) has a sender.
        private static string ItemSender(dynamic item)
        {
            return (int)item.Class == 26 ? (string)item.Organizer : (string)item.SenderName;
        }

        private static string ItemDate(dynamic item)
        {
            DateTime when = (int)item.Class == 26 ? (DateTime)item.Start : (DateTime)item.ReceivedTime;
            return when.ToString("yyyy-MM-dd HH:mm");
        }
    }
}
