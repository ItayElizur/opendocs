import type { ConfirmRule } from '@officeai/app-shell'

// How the Approve/Deny card looks for Outlook's calls. Whether a card is shown at all is the administrator's call
// (policy.json requireApproval). The first matching rule wins, so the specific delete_email rule comes before the
// general one. send_email/send_event have no card: they exist only in Full autonomy, which never asks. `preview` asks the add-in for the item
// behind the call (type, subject, sender, date) and skips the card when the call would be refused; `fields` adds rows from
// the call's own arguments when there is no item yet. The body is never shown.
const text = (v: unknown): string => (Array.isArray(v) ? v.join(', ') : typeof v === 'string' ? v : '')
export const OUTLOOK_CONFIRM_ACTIONS: ConfirmRule[] = [
  { name: 'delete_email', when: (i) => i.permanent === true, danger: true, preview: true,
    title: { en: 'Delete this message permanently? This cannot be undone.', he: 'למחוק את ההודעה לצמיתות? אי אפשר לבטל.' } },
  { name: 'delete_email', preview: true,
    title: { en: 'Move this message to Deleted Items?', he: 'להעביר את ההודעה לפריטים שנמחקו?' } },
  { name: 'respond_meeting', danger: true, preview: true,
    fields: (i) => [{ key: 'response', value: text(i.response) }],
    title: { en: 'Send this response to the organizer?', he: 'לשלוח את התגובה למארגן?' } },
]
