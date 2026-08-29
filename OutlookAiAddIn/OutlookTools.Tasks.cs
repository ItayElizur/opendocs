using System;
using System.Globalization;
using System.Text;
using System.Text.Json;
using OfficeAi.Shared;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn
{
    public static partial class OutlookTools
    {
        private static ToolResult ListTasks(JsonElement input)
        {
            int limit = Math.Max(1, Int(input, "limit", 50));
            bool includeCompleted = Bool(input, "include_completed", false);

            Outlook.Folder tasksFolder = (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks);
            Outlook.Table table = tasksFolder.GetTable(Type.Missing, Outlook.OlTableContents.olUserItems);
            table.Columns.RemoveAll();
            table.Columns.Add("EntryID");
            table.Columns.Add("Subject");
            table.Columns.Add("DueDate");
            table.Columns.Add("StartDate");
            table.Columns.Add("Status");
            table.Columns.Add("PercentComplete");
            table.Columns.Add("Complete");
            table.Columns.Add("ReminderTime");

            var sb = new StringBuilder();
            int n = 0;
            while (!table.EndOfTable && n < limit)
            {
                Outlook.Row row = table.GetNextRow();
                bool complete = false;
                try { complete = Convert.ToBoolean(row["Complete"]); } catch { }
                if (complete && !includeCompleted) continue;
                n++;
                sb.AppendLine("- task_id: " + Convert.ToString(row["EntryID"], CultureInfo.InvariantCulture));
                sb.AppendLine("  subject: " + (Convert.ToString(row["Subject"], CultureInfo.InvariantCulture) ?? ""));
                sb.AppendLine("  due: " + DateCell(row["DueDate"]) + "  start: " + DateCell(row["StartDate"]));
                sb.AppendLine("  status: " + Convert.ToString(row["Status"], CultureInfo.InvariantCulture) +
                              "  percent: " + Convert.ToString(row["PercentComplete"], CultureInfo.InvariantCulture) +
                              "  complete: " + complete);
            }
            if (n == 0) return new ToolResult { Output = includeCompleted ? "No tasks." : "No open tasks.", Summary = "list_tasks" };
            return new ToolResult { Output = sb.ToString(), Summary = "list_tasks" };
        }

        private static string DateCell(object v)
        {
            try
            {
                DateTime d = Convert.ToDateTime(v, CultureInfo.InvariantCulture);
                if (d.Year < 1900 || d.Year > 4000) return "(none)";
                return Iso(d);
            }
            catch { return "(none)"; }
        }

        private static ToolResult CreateTask(JsonElement input)
        {
            string subject = ReqStr(input, "subject");
            Outlook.TaskItem t = (Outlook.TaskItem)App.CreateItem(Outlook.OlItemType.olTaskItem);
            t.Subject = subject;

            string body = Str(input, "body", null);
            if (body != null) t.Body = body;
            DateTime? due = DateArg(input, "due_date");
            if (due.HasValue) t.DueDate = due.Value;
            DateTime? sd = DateArg(input, "start_date");
            if (sd.HasValue) t.StartDate = sd.Value;
            DateTime? rem = DateArg(input, "reminder_time");
            if (rem.HasValue) { t.ReminderSet = true; t.ReminderTime = rem.Value; }
            string imp = Str(input, "importance", null);
            if (imp != null) t.Importance = ParseImportance(imp);

            t.Save();
            return new ToolResult { Output = "Task created: " + subject + "\ntask_id: " + t.EntryID, Mutated = true, Summary = "create_task" };
        }

        private static ToolResult UpdateTask(JsonElement input)
        {
            string id = ReqStr(input, "task_id");
            Outlook.TaskItem t = ItemById(id, null) as Outlook.TaskItem;
            if (t == null) return new ToolResult { Output = "task_id does not resolve to a task.", IsError = true, Summary = "update_task" };

            string subject = Str(input, "subject", null);
            if (subject != null) t.Subject = subject;
            DateTime? due = DateArg(input, "due_date");
            if (due.HasValue) t.DueDate = due.Value;
            DateTime? sd = DateArg(input, "start_date");
            if (sd.HasValue) t.StartDate = sd.Value;
            int pct = Int(input, "percent_complete", -1);
            if (pct >= 0 && pct <= 100) t.PercentComplete = pct;
            string status = Str(input, "status", null);
            if (status != null) t.Status = ParseTaskStatus(status);
            if (Bool(input, "mark_complete", false)) { t.Complete = true; t.PercentComplete = 100; }

            t.Save();
            return new ToolResult { Output = "Task updated: " + (t.Subject ?? ""), Mutated = true, Summary = "update_task" };
        }

        private static ToolResult SetReminder(JsonElement input)
        {
            string id = ReqStr(input, "item_id");
            bool clear = Bool(input, "clear", false);
            object item = ItemById(id, null);
            dynamic d = item;

            if (clear)
            {
                d.ReminderSet = false;
            }
            else
            {
                DateTime? rem = DateArg(input, "reminder_time");
                if (!rem.HasValue) return new ToolResult { Output = "reminder_time is required unless clear=true.", IsError = true, Summary = "set_reminder" };
                d.ReminderSet = true;
                d.ReminderTime = rem.Value;
            }
            d.Save();
            return new ToolResult { Output = clear ? "Reminder cleared." : "Reminder set.", Mutated = true, Summary = "set_reminder" };
        }

        private static ToolResult SetEmailReminder(JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            Outlook.MailItem mail = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (mail == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = "set_email_reminder" };

            mail.MarkAsTask(ParseMarkInterval(Str(input, "mark_interval", null)));
            DateTime? due = DateArg(input, "due_date");
            if (due.HasValue) { mail.TaskDueDate = due.Value; mail.TaskStartDate = due.Value; }
            DateTime? rem = DateArg(input, "reminder_time");
            if (rem.HasValue) { mail.ReminderSet = true; mail.ReminderTime = rem.Value; }
            mail.Save();

            return new ToolResult
            {
                Output = "Follow-up flag set on \"" + (mail.Subject ?? "") + "\"" + (rem.HasValue ? " with reminder " + Iso(rem.Value) : "") + ".",
                Mutated = true,
                Summary = "set_email_reminder",
            };
        }

        private static Outlook.OlImportance ParseImportance(string s)
        {
            switch ((s ?? "").Trim().ToLowerInvariant())
            {
                case "high": return Outlook.OlImportance.olImportanceHigh;
                case "low": return Outlook.OlImportance.olImportanceLow;
                default: return Outlook.OlImportance.olImportanceNormal;
            }
        }

        private static Outlook.OlTaskStatus ParseTaskStatus(string s)
        {
            switch ((s ?? "").Trim().ToLowerInvariant().Replace(" ", ""))
            {
                case "inprogress": return Outlook.OlTaskStatus.olTaskInProgress;
                case "complete":
                case "completed": return Outlook.OlTaskStatus.olTaskComplete;
                case "waiting": return Outlook.OlTaskStatus.olTaskWaiting;
                case "deferred": return Outlook.OlTaskStatus.olTaskDeferred;
                default: return Outlook.OlTaskStatus.olTaskNotStarted;
            }
        }

        private static Outlook.OlMarkInterval ParseMarkInterval(string s)
        {
            switch ((s ?? "").Trim().ToLowerInvariant().Replace(" ", ""))
            {
                case "today": return Outlook.OlMarkInterval.olMarkToday;
                case "tomorrow": return Outlook.OlMarkInterval.olMarkTomorrow;
                case "nextweek": return Outlook.OlMarkInterval.olMarkNextWeek;
                case "nodate": return Outlook.OlMarkInterval.olMarkNoDate;
                case "complete": return Outlook.OlMarkInterval.olMarkComplete;
                default: return Outlook.OlMarkInterval.olMarkThisWeek;
            }
        }
    }
}
