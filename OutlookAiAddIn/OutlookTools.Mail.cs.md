# OutlookTools.Mail.cs

## `OpenEmail`

Display() on a MailItem is well-documented Outlook automation
behavior to mark the item read as a side effect (distinct from
mark_email_read's explicit UnRead assignment above) - this tool
is classified read-only (AlwaysAllowedTools, entry.ts's
readOnlyTools) and must not silently flip UnRead underneath
that guarantee. Restore the original state immediately after
Display() rather than reclassifying the tool to a higher tier,
since the actual intent here is "never mutates the item", not
"mutates, but only with permission" - see the code review this
fixed for the two ways to close this gap.
