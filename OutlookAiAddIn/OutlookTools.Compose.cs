using System;
using System.Net;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        // Draft-and-display only. These NEVER call .Send() - a native Outlook
        // compose/appointment window is opened for the user to review and send.
        private const string Signature = "\n\n— Written with Airchat";

        private static string SeedSignature(string body)
        {
            if (string.IsNullOrEmpty(body)) return body ?? "";
            return body + Signature;
        }

        private static string PrependHtml(string body, string existingHtml)
        {
            if (string.IsNullOrEmpty(body)) return existingHtml;
            string html = WebUtility.HtmlEncode(body).Replace("\n", "<br>");
            return html + "<br><br>" + (existingHtml ?? "");
        }

        private static ToolResult DraftEmail(JsonElement input)
        {
            string to = Str(input, "to", "");
            string subject = Str(input, "subject", "");
            string body = Str(input, "body", "");

            Outlook.MailItem m = (Outlook.MailItem)App.CreateItem(Outlook.OlItemType.olMailItem);
            if (!string.IsNullOrEmpty(to)) m.To = to;
            m.Subject = subject;
            m.Body = SeedSignature(body);
            m.Display(false);
            return new ToolResult { Output = "Opened a draft in Outlook for the user to review and send.", Summary = "draft_email" };
        }

        private static ToolResult ReplyEmail(JsonElement input, bool all)
        {
            string id = ReqStr(input, "message_id");
            string body = Str(input, "body", "");
            string tool = all ? "reply_all_email" : "reply_email";

            Outlook.MailItem orig = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (orig == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = tool };

            Outlook.MailItem reply = all ? orig.ReplyAll() : orig.Reply();
            if (!string.IsNullOrEmpty(body)) reply.HTMLBody = PrependHtml(body, reply.HTMLBody);
            reply.Display(false);
            return new ToolResult { Output = "Opened a " + (all ? "reply-all" : "reply") + " draft in Outlook for the user to review and send.", Summary = tool };
        }

        private static ToolResult ForwardEmail(JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            string to = Str(input, "to", "");
            string body = Str(input, "body", "");

            Outlook.MailItem orig = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (orig == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = "forward_email" };

            Outlook.MailItem fwd = orig.Forward();
            if (!string.IsNullOrEmpty(to)) fwd.To = to;
            if (!string.IsNullOrEmpty(body)) fwd.HTMLBody = PrependHtml(body, fwd.HTMLBody);
            fwd.Display(false);
            return new ToolResult { Output = "Opened a forward draft in Outlook for the user to review and send.", Summary = "forward_email" };
        }

        private static ToolResult DraftEvent(JsonElement input)
        {
            Outlook.AppointmentItem a = (Outlook.AppointmentItem)App.CreateItem(Outlook.OlItemType.olAppointmentItem);
            a.Subject = Str(input, "subject", "");
            a.Location = Str(input, "location", "");
            a.Body = Str(input, "body", "");

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            if (start.HasValue) a.Start = start.Value;
            if (end.HasValue) a.End = end.Value;

            string req = Str(input, "required_attendees", "");
            string opt = Str(input, "optional_attendees", "");
            if (!string.IsNullOrEmpty(req) || !string.IsNullOrEmpty(opt))
            {
                a.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                AddAttendees(a, req, Outlook.OlMeetingRecipientType.olRequired);
                AddAttendees(a, opt, Outlook.OlMeetingRecipientType.olOptional);
                try { a.Recipients.ResolveAll(); } catch { }
            }
            a.Display(false);
            return new ToolResult { Output = "Opened an appointment draft in Outlook for the user to review and send.", Summary = "draft_event" };
        }

        private static void AddAttendees(Outlook.AppointmentItem a, string csv, Outlook.OlMeetingRecipientType type)
        {
            if (string.IsNullOrEmpty(csv)) return;
            foreach (string part in csv.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string addr = part.Trim();
                int lt = addr.LastIndexOf('<');
                if (lt >= 0 && addr.EndsWith(">")) addr = addr.Substring(lt + 1, addr.Length - lt - 2).Trim();
                if (addr.Length == 0) continue;
                Outlook.Recipient r = a.Recipients.Add(addr);
                r.Type = (int)type;
            }
        }
    }
}
