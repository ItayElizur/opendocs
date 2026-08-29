using System;
using System.Collections.Generic;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public partial class TaskPaneHost : PaneHostBase
    {
        private string _chatId;

        // Same "never dereference COM in the constructor" rule the other hosts
        // document: do NOT touch Application.Session here - it hits the same
        // VSTO-connect timing hazard. PrimarySmtpAddress() runs lazily on the
        // first GetChatId(), by which point the session is settled.
        public TaskPaneHost() : base("OutlookAiAddIn")
        {
        }

        protected override ToolResult ExecuteTool(string name, JsonElement input)
        {
            return OutlookTools.Execute(GetChatId(), name, input);
        }

        protected override string GetChatId()
        {
            // Per-mailbox: one rolling chat keyed by the primary SMTP address.
            // No file-path / provisional-id / Migrate lifecycle - a mailbox has
            // a stable identity from the first call, so the id is final once set.
            if (_chatId != null) return _chatId;
            _chatId = "mbx-" + ChatStore.ChatIdForKey(PrimarySmtpAddress() ?? "");
            return _chatId;
        }

        protected override void SetEditingMode(EditingMode mode)
        {
            OutlookTools.SetMode(GetChatId(), mode);
        }

        private string _smtp;

        private string PrimarySmtpAddress()
        {
            if (_smtp != null) return _smtp;
            try
            {
                Outlook.NameSpace ns = Globals.ThisAddIn.Application.Session;
                Outlook.AddressEntry ae = ns.CurrentUser.AddressEntry;

                if (ae != null && ae.Type == "EX")
                {
                    Outlook.ExchangeUser eu = ae.GetExchangeUser();
                    if (eu != null && !string.IsNullOrEmpty(eu.PrimarySmtpAddress))
                        return _smtp = eu.PrimarySmtpAddress;
                }
                if (ae != null)
                {
                    try
                    {
                        object v = ae.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"); // PR_SMTP_ADDRESS
                        string s = v as string;
                        if (!string.IsNullOrEmpty(s)) return _smtp = s;
                    }
                    catch { }
                    if (!string.IsNullOrEmpty(ae.Address) && ae.Address.Contains("@"))
                        return _smtp = ae.Address;
                }
                if (ns.Accounts != null && ns.Accounts.Count > 0)
                {
                    string acc = ns.Accounts[1].SmtpAddress;
                    if (!string.IsNullOrEmpty(acc)) return _smtp = acc;
                }
                return _smtp = ns.CurrentUser.Name;
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("PrimarySmtpAddress", ex);
                return null;
            }
        }

        // Wired from ThisAddIn.cs's per-Explorer SelectionChange sink. Builds the
        // 'selection-changed' bridge payload for the Explorer's currently-
        // selected mail item(s) / conversation - see bootstrap.ts's 'mail'
        // SelectionContext variant.
        public void OnSelectionChanged(Outlook.Selection selection)
        {
            var ids = new List<string>();
            string subject = "";
            string senderName = "";
            string conversationTopic = null;
            string folderName = "";

            try
            {
                Outlook.Explorer active = Globals.ThisAddIn.Application.ActiveExplorer();
                if (active != null && active.CurrentFolder != null) folderName = active.CurrentFolder.Name;
            }
            catch { }

            try
            {
                int count = selection != null ? selection.Count : 0;
                if (count > 50) count = 50;
                for (int i = 1; i <= count; i++)
                {
                    object item = selection[i];
                    Outlook.MailItem mail = item as Outlook.MailItem;
                    Outlook.MeetingItem meeting = item as Outlook.MeetingItem;
                    if (mail != null)
                    {
                        ids.Add(mail.EntryID);
                        if (ids.Count == 1)
                        {
                            subject = mail.Subject ?? "";
                            senderName = mail.SenderName ?? "";
                            conversationTopic = mail.ConversationTopic;
                        }
                    }
                    else if (meeting != null)
                    {
                        ids.Add(meeting.EntryID);
                        if (ids.Count == 1)
                        {
                            subject = meeting.Subject ?? "";
                            senderName = meeting.SenderName ?? "";
                            conversationTopic = meeting.ConversationTopic;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLog.WriteException("OnSelectionChanged", ex);
            }

            string signature = ids.Count > 0 ? "outlook:" + string.Join(",", ids) : "outlook:none";
            PostSelection(new
            {
                kind = "selection-changed",
                app = "outlook",
                hasSelection = ids.Count > 0,
                count = ids.Count,
                entryIds = ids.ToArray(),
                subject,
                senderName,
                folderName,
                conversationTopic,
            }, signature);
        }
    }
}
