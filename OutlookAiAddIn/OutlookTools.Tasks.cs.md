# OutlookTools.Tasks.cs

## `AppendFlaggedEmails`

Outlook's "Flag for follow up" on a mail item never creates a
TaskItem in the Tasks folder - it just sets flag/date properties on
the mail in place, wherever it lives. The To-Do List is Outlook's
own aggregation of every flagged item across the mailbox, so it's
the one place that surfaces those without walking every folder.
Real tasks show up in there too; skip them since AppendTasks
already listed those from the Tasks folder directly.

Only EntryID/Subject/MessageClass are pulled via the Table - those
are confirmed-valid Table column names (used elsewhere already).
FlagStatus is NOT (Table.Columns.Add("FlagStatus") throws "the
property is unknown" at runtime, despite FlagStatus being a real
MailItem property) - Table's recognized column-name set and the
object model's property names are two separate, only partially
overlapping things, and there's no guarantee TaskDueDate/
TaskStartDate would have fared any better. Resolving the actual
item and reading its properties directly sidesteps that guessing
game entirely, at the cost of one COM call per flagged row (this
list is small, unlike bulk mail listing).
