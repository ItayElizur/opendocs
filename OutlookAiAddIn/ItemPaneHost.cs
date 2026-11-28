using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    // The chat pane inside one item window: a mail draft, or an appointment or meeting. Its chat belongs to that item
    // (ItemChatId), with its own editing mode and its own undo track. See ItemPaneHost.cs.md.
    public class ItemPaneHost : PaneHostBase
    {
        private const string AppDataFolder = "OutlookAiAddIn";
        private const int BodyPreviewChars = 500;

        private object _item;
        private string _chatId;
        private readonly Outlook.ItemEvents_10_PropertyChangeEventHandler _onPropertyChange;

        public ItemPaneHost(object item) : base(AppDataFolder)
        {
            _item = item;
            _onPropertyChange = name =>
            {
                try { PostItem(); }
                catch (Exception ex) { DebugLog.WriteException("ItemPaneHost.PropertyChange", ex); }
            };
            if (item is Outlook.ItemEvents_10_Event events) events.PropertyChange += _onPropertyChange;
        }

        // The item window's own gate, the second copy of entry.ts's item-window lists. Only these mailbox reads pass on
        // to OutlookTools; every mailbox action stays in the main window. Must stay in sync with entry.ts's
        // itemReadOnlyTools.
        private static readonly HashSet<string> ItemReadTools = new HashSet<string>
        {
            "search_emails", "get_email", "open_email", "get_attachment", "list_folders", "search_contacts",
            "list_events", "get_event", "find_meeting_slots", "list_color_categories", "list_tasks",
        };

        protected override Task<ToolResult> ExecuteTool(string name, JsonElement input)
        {
            if (!ItemReadTools.Contains(name))
            {
                return Task.FromResult(new ToolResult
                {
                    Output = "Blocked: " + name + " is done from the main Outlook window, not from this item's window.",
                    IsError = true,
                    Summary = name,
                });
            }
            return OutlookTools.ExecuteAsync(GetChatId(), name, input);
        }

        // A new item has no EntryID until its first save (Outlook's AutoSave does one after a few minutes), so the chat
        // starts on a provisional id and is moved onto the saved one by the first call after that save.
        protected override string GetChatId()
        {
            string entryId = null, storeId = null;
            try
            {
                dynamic d = _item;
                entryId = d.EntryID;
                if (!string.IsNullOrEmpty(entryId)) storeId = ((Outlook.MAPIFolder)d.Parent).StoreID;
            }
            catch { /* the window is closing, or the item is not saved yet */ }

            string id = ItemChatId.Resolve(_chatId, storeId, entryId, out string migrateFrom);
            if (migrateFrom != null)
            {
                ChatStore.Migrate(AppDataFolder, migrateFrom, id);
                DocSettingsStore.Migrate(AppDataFolder, migrateFrom, id);
                OutlookTools.MoveChat(migrateFrom, id);
            }
            return _chatId = id;
        }

        protected override void SetEditingMode(EditingMode mode)
        {
            OutlookTools.SetMode(GetChatId(), mode);
        }

        protected override int GetOfficeUiLanguageId()
        {
            return Globals.ThisAddIn.GetOfficeUiLanguageId();
        }

        // Who the user is and the mailbox's work week, as the main window's pane gives them.
        protected override Task<string> GetPromptContextAsync()
        {
            return TaskPaneHost.MailboxPromptContextAsync();
        }

        protected override void OnPageStarted()
        {
            try { PostItem(); }
            catch (Exception ex) { DebugLog.WriteException("ItemPaneHost.OnPageStarted", ex); }
        }

        // Sent as the pane's selection, so the model is told what is in this window. Refreshed on every property change;
        // typing in the body does not raise one, so the body preview can lag until the next change.
        private void PostItem()
        {
            if (_item is Outlook.MailItem mail)
            {
                string subject = mail.Subject ?? "", to = mail.To ?? "", cc = mail.CC ?? "", bcc = mail.BCC ?? "";
                string body = Preview(mail.Body);
                string entryId = mail.EntryID ?? "";
                PostItemPayload(new
                {
                    kind = "selection-changed",
                    app = "outlook",
                    hasSelection = true,
                    itemKind = "mail",
                    entryId,
                    subject,
                    to,
                    cc,
                    bcc,
                    bodyPreview = body,
                });
            }
            else if (_item is Outlook.AppointmentItem appt)
            {
                string subject = appt.Subject ?? "", location = appt.Location ?? "";
                string start = Iso(appt.Start), end = Iso(appt.End);
                string required = appt.RequiredAttendees ?? "", optional = appt.OptionalAttendees ?? "", resources = appt.Resources ?? "";
                bool allDay = appt.AllDayEvent;
                bool isMeeting = appt.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting;
                bool isOrganizer = appt.MeetingStatus != Outlook.OlMeetingStatus.olMeetingReceived &&
                                   appt.MeetingStatus != Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled;
                string body = Preview(appt.Body);
                string entryId = appt.EntryID ?? "";
                PostItemPayload(new
                {
                    kind = "selection-changed",
                    app = "outlook",
                    hasSelection = true,
                    itemKind = "appointment",
                    entryId,
                    subject,
                    location,
                    start,
                    end,
                    allDay,
                    isMeeting,
                    isOrganizer,
                    requiredAttendees = required,
                    optionalAttendees = optional,
                    resources,
                    bodyPreview = body,
                });
            }
        }

        // The payload itself is the "did it change?" key, so a field added to it can never be missed by the check.
        private void PostItemPayload(object payload)
        {
            PostSelection(payload, JsonSerializer.Serialize(payload));
        }

        private static string Iso(DateTime value)
        {
            return value.ToString("yyyy-MM-dd'T'HH:mm", CultureInfo.InvariantCulture);
        }

        private static string Preview(string body)
        {
            string text = (body ?? "").Trim();
            return text.Length <= BodyPreviewChars ? text : text.Substring(0, BodyPreviewChars) + "...";
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _item != null)
            {
                try { if (_item is Outlook.ItemEvents_10_Event events) events.PropertyChange -= _onPropertyChange; }
                catch { }
                _item = null;
                if (_chatId != null) OutlookTools.ForgetChat(_chatId);
            }
            base.Dispose(disposing);
        }
    }
}
