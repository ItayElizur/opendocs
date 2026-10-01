using System;
using System.Collections.Generic;
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
            bool includeFlaggedEmails = Bool(input, "include_flagged_emails", true);

            var sb = new StringBuilder();
            int n = AppendTasks(sb, limit, includeCompleted);
            if (includeFlaggedEmails && n < limit) n += AppendFlaggedEmails(sb, limit - n, includeCompleted);

            if (n == 0) return new ToolResult { Output = includeCompleted ? "No tasks." : "No open tasks.", Summary = "list_tasks" };
            return new ToolResult { Output = sb.ToString(), Summary = "list_tasks" };
        }

        // Shared by AppendTasks/AppendFlaggedEmails - the task_id/kind/subject/
        // due/start block is identical for both. Takes plain values rather
        // than an Outlook.Row: AppendTasks' values come straight off the
        // table row, AppendFlaggedEmails' come off a resolved MailItem
        // instead (see the comment in AppendFlaggedEmails for why).
        private static void AppendTaskHeader(StringBuilder sb, string kind, string entryId, string subject, object due, object start)
        {
            sb.AppendLine("- task_id: " + entryId);
            sb.AppendLine("  kind: " + kind);
            sb.AppendLine("  subject: " + (subject ?? ""));
            sb.AppendLine("  due: " + DateCell(due) + "  start: " + DateCell(start));
        }

        private static int AppendTasks(StringBuilder sb, int limit, bool includeCompleted)
        {
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

            int n = 0;
            while (!table.EndOfTable && n < limit)
            {
                Outlook.Row row = table.GetNextRow();
                bool complete = false;
                try { complete = Convert.ToBoolean(row["Complete"]); } catch { }
                if (complete && !includeCompleted) continue;
                n++;
                AppendTaskHeader(sb, "task",
                    Convert.ToString(row["EntryID"], CultureInfo.InvariantCulture),
                    Convert.ToString(row["Subject"], CultureInfo.InvariantCulture),
                    row["DueDate"], row["StartDate"]);
                sb.AppendLine("  status: " + Convert.ToString(row["Status"], CultureInfo.InvariantCulture) +
                              "  percent: " + Convert.ToString(row["PercentComplete"], CultureInfo.InvariantCulture) +
                              "  complete: " + complete);
            }
            return n;
        }

        // Outlook's "Flag for follow up" on a mail item never creates a
        // TaskItem in the Tasks folder - it just sets flag/date properties on
        // the mail in place, wherever it lives. The To-Do List is Outlook's
        // own aggregation of every flagged item across the mailbox, so it's
        // the one place that surfaces those without walking every folder.
        // Real tasks show up in there too; skip them since AppendTasks
        // already listed those from the Tasks folder directly.
        //
        // Only EntryID/Subject/MessageClass are pulled via the Table - those
        // are confirmed-valid Table column names (used elsewhere already).
        // FlagStatus is NOT (Table.Columns.Add("FlagStatus") throws "the
        // property is unknown" at runtime, despite FlagStatus being a real
        // MailItem property) - Table's recognized column-name set and the
        // object model's property names are two separate, only partially
        // overlapping things, and there's no guarantee TaskDueDate/
        // TaskStartDate would have fared any better. Resolving the actual
        // item and reading its properties directly sidesteps that guessing
        // game entirely, at the cost of one COM call per flagged row (this
        // list is small, unlike bulk mail listing).
        private static int AppendFlaggedEmails(StringBuilder sb, int limit, bool includeCompleted)
        {
            Outlook.Folder toDo = (Outlook.Folder)Ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderToDo);
            Outlook.Table table = toDo.GetTable(Type.Missing, Outlook.OlTableContents.olUserItems);
            table.Columns.RemoveAll();
            table.Columns.Add("EntryID");
            table.Columns.Add("MessageClass");

            int n = 0;
            while (!table.EndOfTable && n < limit)
            {
                Outlook.Row row = table.GetNextRow();
                string cls = Convert.ToString(row["MessageClass"], CultureInfo.InvariantCulture) ?? "";
                if (!cls.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase)) continue;

                string entryId = Convert.ToString(row["EntryID"], CultureInfo.InvariantCulture);
                // Also needed for update_task to resolve the item outside the
                // default store (see ItemById) - the To-Do List can surface
                // items from anywhere in the profile, GetItemFromID by EntryID
                // alone can't.
                Outlook.MailItem mail = null;
                try { mail = ItemById(entryId, null) as Outlook.MailItem; } catch { }
                if (mail == null) continue;

                bool complete = mail.FlagStatus == Outlook.OlFlagStatus.olFlagComplete;
                if (complete && !includeCompleted) continue;

                string folderName = "(unknown)";
                try
                {
                    Outlook.Folder parent = mail.Parent as Outlook.Folder;
                    if (parent != null) folderName = parent.Name;
                }
                catch { }

                n++;
                AppendTaskHeader(sb, "flagged_email", entryId, mail.Subject, mail.TaskDueDate, mail.TaskStartDate);
                sb.AppendLine("  folder: " + folderName);
                sb.AppendLine("  complete: " + complete);
            }
            return n;
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

        private static ToolResult CreateTask(string mbxKey, JsonElement input)
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
            RecordCreated(mbxKey, "create_task", "task_id", t, subject);
            return new ToolResult { Output = "Task created: " + subject + "\ntask_id: " + t.EntryID, Mutated = true, Summary = "create_task" };
        }

        // Snapshotted as a group: Complete/PercentComplete/Status drive each
        // other in Outlook, so undo restores all of them together. Complete
        // comes before PercentComplete/Status so restoring Complete=false
        // doesn't reset the other two after they're written.
        private static readonly string[] TaskUndoProps = { "Subject", "DueDate", "StartDate", "Complete", "PercentComplete", "Status" };
        private static readonly string[] FlaggedMailUndoProps = { "TaskDueDate", "TaskStartDate", "FlagStatus" };

        private static ToolResult UpdateTask(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "task_id");
            // folder is only needed when task_id is a flagged email from a
            // non-default store (see AppendFlaggedEmails' "folder" field) -
            // real tasks are always in the default store's Tasks folder, so
            // StoreOf(input) returning null here (no folder passed) still
            // resolves them exactly as before.
            object item = ItemById(id, StoreOf(input));

            Outlook.TaskItem t = item as Outlook.TaskItem;
            if (t != null)
            {
                object[] before = ReadProps(t, TaskUndoProps);
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
                RecordSnapshot(mbxKey, "update_task", t, t.Subject ?? "", TaskUndoProps, before);
                return new ToolResult { Output = "Task updated: " + (t.Subject ?? ""), Mutated = true, Summary = "update_task" };
            }

            // task_id from list_tasks' "kind: flagged_email" rows is the
            // mail's own EntryID - there's no TaskItem to cast to, just flag
            // properties on the mail itself.
            Outlook.MailItem mail = item as Outlook.MailItem;
            if (mail != null)
            {
                object[] before = ReadProps(mail, FlaggedMailUndoProps);
                DateTime? due = DateArg(input, "due_date");
                if (due.HasValue) mail.TaskDueDate = due.Value;
                DateTime? sd = DateArg(input, "start_date");
                if (sd.HasValue) mail.TaskStartDate = sd.Value;
                if (Bool(input, "mark_complete", false)) mail.FlagStatus = Outlook.OlFlagStatus.olFlagComplete;

                // subject/status/percent_complete don't map to a flag on a
                // mail item - say so explicitly rather than silently no-op'ing
                // fields the caller asked to change.
                var ignored = new List<string>();
                if (Str(input, "subject", null) != null) ignored.Add("subject");
                if (Str(input, "status", null) != null) ignored.Add("status");
                if (Int(input, "percent_complete", -1) >= 0) ignored.Add("percent_complete");

                mail.Save();
                RecordSnapshot(mbxKey, "update_task", mail, mail.Subject ?? "", FlaggedMailUndoProps, before);
                string output = "Flagged email updated: " + (mail.Subject ?? "");
                if (ignored.Count > 0) output += "\n(ignored - only apply to real tasks: " + string.Join(", ", ignored) + ")";
                return new ToolResult { Output = output, Mutated = true, Summary = "update_task" };
            }

            return new ToolResult { Output = "task_id does not resolve to a task or a flagged email.", IsError = true, Summary = "update_task" };
        }

        private static ToolResult SetReminder(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "item_id");
            bool clear = Bool(input, "clear", false);
            object item = ItemById(id, null);
            dynamic d = item;
            string[] props = { "ReminderTime", "ReminderSet" };
            object[] before = ReadProps(item, props);

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
            string subject = "";
            try { subject = d.Subject; } catch { }
            RecordSnapshot(mbxKey, "set_reminder", item, subject, props, before);
            return new ToolResult { Output = clear ? "Reminder cleared." : "Reminder set.", Mutated = true, Summary = "set_reminder" };
        }

        private static ToolResult SetEmailReminder(string mbxKey, JsonElement input)
        {
            string id = ReqStr(input, "message_id");
            Outlook.MailItem mail = ItemById(id, StoreOf(input)) as Outlook.MailItem;
            if (mail == null) return new ToolResult { Output = "message_id does not resolve to a mail item.", IsError = true, Summary = "set_email_reminder" };

            bool wasMarkedAsTask = mail.IsMarkedAsTask;
            object[] before = ReadProps(mail, EmailFlagEntry.FlagProps);
            Outlook.OlMarkInterval interval = ParseMarkInterval(Str(input, "mark_interval", null));
            mail.MarkAsTask(interval);
            DateTime? due = DateArg(input, "due_date");
            if (due.HasValue) { mail.TaskDueDate = due.Value; mail.TaskStartDate = due.Value; }
            DateTime? rem = DateArg(input, "reminder_time");
            if (rem.HasValue) { mail.ReminderSet = true; mail.ReminderTime = rem.Value; }
            mail.Save();
            RecordEmailFlag(mbxKey, mail, wasMarkedAsTask, interval, before);

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
