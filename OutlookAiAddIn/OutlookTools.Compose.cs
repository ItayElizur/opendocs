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

        // Send-and-dispatch, Full-autonomy-only counterparts to the
        // draft/reply/forward/event tools above. Same bodies, but .Send()/
        // .Save() instead of .Display(false) - no native window, no review
        // step. Gated in OutlookTools.cs's ExecuteAsync (SendTierTools),
        // never reachable below Full Autonomy.
        private static ToolResult SendEmail(JsonElement input)
        {
            // to is required, unlike draft_email's - draft_email opens a
            // compose window where a human can add a missing recipient
            // before anything goes out; send_email has no such window, so
            // Send()-ing with no recipient set would either throw or, worse,
            // block on a native Outlook resolution prompt this add-in isn't
            // expecting a response to.
            string to = ReqStr(input, "to");
            string subject = Str(input, "subject", "");
            string body = Str(input, "body", "");

            Outlook.MailItem m = (Outlook.MailItem)App.CreateItem(Outlook.OlItemType.olMailItem);
            m.To = to;
            m.Subject = subject;
            m.Body = SeedSignature(body);
            m.Send();
            return new ToolResult { Output = "Sent to " + to + ": \"" + subject + "\".", Mutated = true, Summary = "send_email" };
        }

        private static ToolResult SendReply(JsonElement input, bool all)
        {
            string id = ReqStr(input, "message_id");
            string body = Str(input, "body", "");
            string tool = all ? "send_reply_all" : "send_reply";

            Outlook.MailItem orig = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (orig == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = tool };

            Outlook.MailItem reply = all ? orig.ReplyAll() : orig.Reply();
            if (!string.IsNullOrEmpty(body)) reply.HTMLBody = PrependHtml(body, reply.HTMLBody);
            string to = reply.To;
            reply.Send();
            return new ToolResult { Output = "Sent " + (all ? "reply-all" : "reply") + " to " + to + ": \"" + (orig.Subject ?? "") + "\".", Mutated = true, Summary = tool };
        }

        private static ToolResult SendForward(JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            string to = ReqStr(input, "to");
            string body = Str(input, "body", "");

            Outlook.MailItem orig = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (orig == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = "send_forward" };

            Outlook.MailItem fwd = orig.Forward();
            fwd.To = to;
            if (!string.IsNullOrEmpty(body)) fwd.HTMLBody = PrependHtml(body, fwd.HTMLBody);
            fwd.Send();
            return new ToolResult { Output = "Forwarded to " + to + ": \"" + (orig.Subject ?? "") + "\".", Mutated = true, Summary = "send_forward" };
        }

        // Non-meeting appointments only .Save() - Outlook has nobody to send
        // them to. A meeting (attendees present) must .Send() instead:
        // .Save() alone would leave it sitting unset on the organizer's
        // calendar with invitees never notified. Both members confirmed via
        // reflection against the referenced Outlook PIA (_AppointmentItem
        // exposes both Send() and Save()) before writing this branch.
        private static ToolResult CreateEvent(JsonElement input)
        {
            Outlook.AppointmentItem a = (Outlook.AppointmentItem)App.CreateItem(Outlook.OlItemType.olAppointmentItem);
            a.Subject = Str(input, "subject", "");
            a.Location = Str(input, "location", "");
            a.Body = SeedSignature(Str(input, "body", ""));

            DateTime? start = DateArg(input, "start");
            DateTime? end = DateArg(input, "end");
            if (start.HasValue) a.Start = start.Value;
            if (end.HasValue) a.End = end.Value;

            string req = Str(input, "required_attendees", "");
            string opt = Str(input, "optional_attendees", "");
            bool isMeeting = !string.IsNullOrEmpty(req) || !string.IsNullOrEmpty(opt);
            if (isMeeting)
            {
                a.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                AddAttendees(a, req, Outlook.OlMeetingRecipientType.olRequired);
                AddAttendees(a, opt, Outlook.OlMeetingRecipientType.olOptional);
                try { a.Recipients.ResolveAll(); } catch { }
                a.Send();
                // This confirmation line is the only place the user sees who
                // an irreversible, unreviewed invite went to - do not let an
                // empty req (optional_attendees-only) produce a malformed
                // "to ; alice@example.com." leading separator.
                string attendeeList = string.IsNullOrEmpty(req) ? opt : string.IsNullOrEmpty(opt) ? req : req + "; " + opt;
                return new ToolResult { Output = "Created and sent invite: \"" + (a.Subject ?? "") + "\" to " + attendeeList + ".", Mutated = true, Summary = "create_event" };
            }

            a.Save();
            return new ToolResult { Output = "Created event: \"" + (a.Subject ?? "") + "\".", Mutated = true, Summary = "create_event" };
        }
    }
}
